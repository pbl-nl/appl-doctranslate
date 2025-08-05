"""
TXT translation with Azure OpenAI and Formatting Preservation

This script reads a .txt file, translates the text content using Azure OpenAI
while preserving text structure, spacing, indentation, and layout, then outputs 
a new translated .txt file.

"""
import os
import re
import time
from typing import List, Tuple, Dict
from openai import AzureOpenAI
# local imports
import utils


def detect_encoding(file_path: str) -> str:
    """
    Detect the encoding of a text file.
    
    Args:
        file_path: Path to the text file
        
    Returns:
        Detected encoding
    """
    encodings = ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252', 'iso-8859-1']
    
    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                f.read()
            return encoding
        except UnicodeDecodeError:
            continue
    
    # Fallback to utf-8 with error handling
    return 'utf-8'


def parse_text_structure(text: str) -> List[Dict]:
    """
    Parse text into structured segments preserving formatting.
    
    Args:
        text: Input text content
        
    Returns:
        List of dictionaries with structure information
    """
    lines = text.split('\n')
    structures = []
    
    for line_num, line in enumerate(lines):
        if not line.strip():
            # Empty line or whitespace-only line
            structures.append({
                'content': "",
                'leading_whitespace': line,
                'trailing_whitespace': "",
                'is_empty': True,
                'line_number': line_num
            })
        else:
            # Line with content - extract leading and trailing whitespace
            leading_match = re.match(r'^(\s*)', line)
            trailing_match = re.search(r'(\s*)$', line)
            
            leading_ws = leading_match.group(1) if leading_match else ""
            trailing_ws = trailing_match.group(1) if trailing_match else ""
            content = line.strip()
            
            structures.append({
                'content': content,
                'leading_whitespace': leading_ws,
                'trailing_whitespace': trailing_ws,
                'is_empty': False,
                'line_number': line_num
            })
    
    return structures


def group_structures_for_translation(structures: List[Dict]) -> List[Tuple[List[int], List[str]]]:
    """
    Group text structures into translation batches while preserving context.
    
    Args:
        structures: List of structure dictionaries
        
    Returns:
        List of tuples containing (indices, texts) for batch translation
    """
    batches = []
    current_batch_indices = []
    current_batch_texts = []
    current_length = 0
    max_batch_length = 3000  # Conservative limit for token count
    
    for i, struct in enumerate(structures):
        if struct['is_empty'] or not struct['content'].strip():
            # If we have accumulated content, finish the current batch
            if current_batch_texts:
                batches.append((current_batch_indices.copy(), current_batch_texts.copy()))
                current_batch_indices.clear()
                current_batch_texts.clear()
                current_length = 0
            continue
        
        text_length = len(struct['content'])
        
        # Start new batch if current would be too long
        if current_length + text_length > max_batch_length and current_batch_texts:
            batches.append((current_batch_indices.copy(), current_batch_texts.copy()))
            current_batch_indices.clear()
            current_batch_texts.clear()
            current_length = 0
        
        current_batch_indices.append(i)
        current_batch_texts.append(struct['content'])
        current_length += text_length
    
    # Add final batch if it has content
    if current_batch_texts:
        batches.append((current_batch_indices, current_batch_texts))
    
    return batches


def translate_text_batch(client: AzureOpenAI, deployment_name: str, texts: List[str], target_language: str) -> List[str]:
    """
    Translate multiple text strings in a single API call.
    
    Args:
        client: Azure OpenAI client
        deployment_name: Azure OpenAI deployment name
        texts: List of texts to translate
        target_language: Target language
        
    Returns:
        List of translated texts
    """
    if not texts or all(not text.strip() for text in texts):
        return texts
    
    prompt = f"""Translate the following texts to {target_language}.
    IMPORTANT INSTRUCTIONS:
    - Preserve the exact meaning and tone of each segment
    - Maintain any special formatting, punctuation, or symbols
    - Keep the same structure and style as the original
    - Each segment should be translated independently but consider context
    - Return ONLY the translations, one per line, in the exact same order
    - Do not add explanations, numbers, or extra text

    Text segments to translate:
    """
    for text in texts:
        prompt += f"{text}\n"
    
    try:
        response = client.chat.completions.create(
            model=deployment_name,
            messages=[
                {
                    "role": "system", 
                    "content": "You are a professional translator. Translate accurately while preserving formatting, structure, and meaning. Return only the translated text, one translation per line, in the same order as provided."
                },
                {"role": "user", "content": prompt}
            ],
            max_tokens=4000,
            temperature=0
        )
        
        translated_content = response.choices[0].message.content.strip()
        translated_lines = translated_content.split('\n')
        
        # Clean up translations and ensure we have the right number
        cleaned_translations = []
        for line in translated_lines:
            line = line.strip()
            if line:  # Only add non-empty lines
                # Remove numbering if the model added it
                line = re.sub(r'^\d+\.\s*', '', line)
                cleaned_translations.append(line)
        
        # Ensure we have the same number of translations as inputs
        while len(cleaned_translations) < len(texts):
            cleaned_translations.append(texts[len(cleaned_translations)])
        
        return cleaned_translations[:len(texts)]
        
    except Exception as e:
        print(f"Translation error: {e}")
        return texts  # Return original texts if translation fails


def reconstruct_text(structures: List[Dict]) -> str:
    """
    Reconstruct the text from structure dictionaries.
    
    Args:
        structures: List of structure dictionaries
        
    Returns:
        Reconstructed text with preserved formatting
    """
    lines = []
    for struct in structures:
        if struct['is_empty']:
            lines.append(struct['leading_whitespace'])
        else:
            reconstructed_line = struct['leading_whitespace'] + struct['content'] + struct['trailing_whitespace']
            lines.append(reconstructed_line)
    
    return '\n'.join(lines)


def translate_txt_document(client: AzureOpenAI, model: str, input_path: str, target_language: str, output_folder: str, save_as_pdf: bool) -> bool:
    """
    Translate a text file while preserving formatting.
    
    Args:
        input_path: Path to input .txt file
        output_path: Path to output .txt file
        target_lang: Target language
        source_lang: Source language
        
    Returns:
        True if successful, False otherwise
    """
    try:
        file_name = os.path.basename(input_path)
        output_file_path = os.path.join(output_folder, target_language + "_" + file_name)
        
        # Detect encoding and read file
        encoding = detect_encoding(input_path)
        print(f"Detected encoding: {encoding}")
        
        with open(input_path, 'r', encoding=encoding, errors='replace') as f:
            original_text = f.read()
        
        print(f"Original text length: {len(original_text)} characters")
        
        # Parse text structure
        print("Analyzing text structure...")
        structures = parse_text_structure(original_text)
        
        # Group structures for efficient translation
        translation_batches = group_structures_for_translation(structures)
        print(f"Created {len(translation_batches)} translation batches")
        
        # Translate each batch
        for batch_num, (indices, texts) in enumerate(translation_batches):
            print(f"Translating batch {batch_num + 1}/{len(translation_batches)} ({len(texts)} segments)")
            
            translated_texts = translate_text_batch(client, model, texts, target_language)
            
            # Apply translations back to structures
            for idx, translated_text in zip(indices, translated_texts):
                if idx < len(structures):
                    structures[idx]['content'] = translated_text
            
            # Small delay between batches to avoid rate limiting
            if batch_num < len(translation_batches) - 1:
                time.sleep(0.5)
        
        # Reconstruct the translated text
        print("Reconstructing translated text...")
        translated_text = reconstruct_text(structures)
        
        # Save translated file
        output_encoding = 'utf-8'  # Always save as UTF-8 for best compatibility
        with open(output_file_path, 'w', encoding=output_encoding, newline='') as f:
            f.write(translated_text)
        
        # if indicated, save as pdf file
        if save_as_pdf:
            pdf_file_name = os.path.splitext(file_name)[0] + ".pdf"
            pdf_file_path = os.path.join(output_folder, target_language + "_" + pdf_file_name)
            utils.convert_txt_to_pdf(output_file_path, pdf_file_path)
            # add watermark to created pdf file
            utils.add_watermark(pdf_file_path, pdf_file_path, "X:/User/troosts/projects/translator/watermark.pdf")
            # remove converted .docx file
            os.remove(output_file_path)

        print(f"Translated text length: {len(translated_text)} characters")
        return True
        
    except Exception as e:
        print(f"Error processing file: {e}")
        import traceback
        traceback.print_exc()
        return False


def preview_structure(input_path: str, max_lines: int = 20):
    """
    Preview the structure analysis of a text file.
    
    Args:
        input_path: Path to input .txt file
        max_lines: Maximum number of lines to preview
    """
    try:
        encoding = detect_encoding(input_path)
        
        with open(input_path, 'r', encoding=encoding, errors='replace') as f:
            text = f.read()
        
        structures = parse_text_structure(text)
        
        print(f"\nStructure Preview (first {max_lines} lines):")
        print("-" * 60)
        
        for i, struct in enumerate(structures[:max_lines]):
            if struct['is_empty']:
                print(f"Line {i+1:3d}: [EMPTY] '{struct['leading_whitespace']}'")
            else:
                leading_repr = repr(struct['leading_whitespace']) if struct['leading_whitespace'] else "''"
                trailing_repr = repr(struct['trailing_whitespace']) if struct['trailing_whitespace'] else "''"
                content_preview = struct['content'][:50] + "..." if len(struct['content']) > 50 else struct['content']
                print(f"Line {i+1:3d}: Leading:{leading_repr} Content:'{content_preview}' Trailing:{trailing_repr}")
        
        if len(structures) > max_lines:
            print(f"... and {len(structures) - max_lines} more lines")
            
    except Exception as e:
        print(f"Error previewing structure: {e}")


# def main():
#     parser = argparse.ArgumentParser(
#         description="Translate TXT files using Azure OpenAI while preserving formatting"
#     )
#     parser.add_argument(
#         "input_file", 
#         help="Input .txt file path"
#     )
#     parser.add_argument(
#         "output_file", 
#         nargs='?',
#         help="Output .txt file path (optional for preview mode)"
#     )
#     parser.add_argument(
#         "--source-lang", 
#         default="auto",
#         help="Source language (default: auto-detect)"
#     )
#     parser.add_argument(
#         "--target-lang", 
#         default="Spanish",
#         help="Target language (default: Spanish). Use full language names."
#     )
#     parser.add_argument(
#         "--preview", 
#         action="store_true",
#         help="Preview the text structure analysis without translating"
#     )
#     parser.add_argument(
#         "--preview-lines", 
#         type=int, 
#         default=20,
#         help="Number of lines to show in preview mode (default: 20)"
#     )
    
#     args = parser.parse_args()
    
#     # Validate input file
#     if not Path(args.input_file).exists():
#         print(f"Error: Input file '{args.input_file}' not found.")
#         sys.exit(1)
    
#     # Preview mode
#     if args.preview:
#         print(f"Previewing structure of: {args.input_file}")
#         preview_structure(args.input_file, args.preview_lines)
#         return
    
#     # Translation mode
#     if not args.output_file:
#         print("Error: Output file is required for translation mode.")
#         print("Use --preview flag to preview structure without translating.")
#         sys.exit(1)
    
#     # Create output directory if it doesn't exist
#     output_path = Path(args.output_file)
#     output_path.parent.mkdir(parents=True, exist_ok=True)
    
#     print(f"Translating from '{args.source_lang}' to '{args.target_lang}' using Azure OpenAI")
#     print(f"Input: {args.input_file}")
#     print(f"Output: {args.output_file}")
    
#     success = translate_file(args.input_file, args.output_file, args.target_lang, args.source_lang)
    
#     if success:
#         print(f"✅ Translation completed successfully!")
#         print(f"📄 Output saved to: {args.output_file}")
#     else:
#         print("❌ Translation failed. Please check the error messages above.")
#         sys.exit(1)


# if __name__ == "__main__":
#     main()
