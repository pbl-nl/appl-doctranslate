"""
DOCX Translation with Azure OpenAI and Formatting Preservation

This script reads a .docx file, translates the text content using Azure OpenAI
while preserving all formatting, fonts, and layout, then outputs a new translated .docx file.

"""
import os
import time
from typing import List
from docx import Document
from openai import AzureOpenAI
# local imports
import utils


def translate_text_batch(client: AzureOpenAI, deployment_name: str, texts: List[str], target_language: str) -> List[str]:
    """
    Translate multiple text strings in a single API call for efficiency.
    
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
    
    # Filter out empty texts but keep track of their positions
    text_map = {}
    non_empty_texts = []
    for i, text in enumerate(texts):
        if text.strip():
            text_map[len(non_empty_texts)] = i
            non_empty_texts.append(text)
    
    if not non_empty_texts:
        return texts
    
    prompt = f"""Translate the following texts to {target_language}. 
    Preserve the exact meaning and tone. Maintain any formatting markers or special characters.
    Return only the translations, separated by "---TRANSLATION_SEPARATOR---", in the same order as provided.

    Texts to translate:
    """
    for i, text in enumerate(non_empty_texts):
        prompt += f"\n{i+1}. {text}"
    
    try:
        response = client.chat.completions.create(
            model=deployment_name,
            messages=[
                {"role": "system", "content": "You are a professional translator. Translate accurately while preserving formatting and meaning."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=4000,
            temperature=0
        )
        
        translated_content = response.choices[0].message.content.strip()
        
        # Split the response by separator
        if "---TRANSLATION_SEPARATOR---" in translated_content:
            translated_parts = translated_content.split("---TRANSLATION_SEPARATOR---")
        else:
            # Fallback: split by numbered list if separator not used
            translated_parts = []
            lines = translated_content.split('\n')
            current_translation = ""
            
            for line in lines:
                line = line.strip()
                if line and (line.startswith(f"{len(translated_parts)+1}.") or 
                            line.startswith(f"{len(translated_parts)+1} ")):
                    if current_translation:
                        translated_parts.append(current_translation.strip())
                    current_translation = line.split('.', 1)[1].strip() if '.' in line else line
                else:
                    current_translation += " " + line if current_translation and line else line
            
            if current_translation:
                translated_parts.append(current_translation.strip())
        
        # Clean up translations and map back to original positions
        result = texts.copy()
        for i, translation in enumerate(translated_parts[:len(non_empty_texts)]):
            original_index = text_map[i]
            result[original_index] = translation.strip()
        
        return result
        
    except Exception as e:
        print(f"Translation error: {e}")
        # Return original texts if translation fails
        return texts


def collect_paragraph_texts(paragraphs) -> List[str]:
    """
    Collect all text content from paragraphs.
    
    Args:
        paragraphs: List of paragraph objects
        
    Returns:
        List of text strings
    """
    texts = []
    for paragraph in paragraphs:
        if paragraph.text.strip():
            texts.append(paragraph.text)
        else:
            texts.append("")
    return texts


def apply_translations_to_paragraphs(paragraphs, translations: List[str]):
    """
    Apply translations to paragraphs while preserving formatting.
    
    Args:
        paragraphs: List of paragraph objects
        translations: List of translated texts
    """
    for paragraph, translation in zip(paragraphs, translations):
        if not translation.strip():
            continue
            
        # Store original formatting from first meaningful run
        original_formatting = None
        for run in paragraph.runs:
            if run.text.strip():
                original_formatting = {
                    'bold': run.bold,
                    'italic': run.italic,
                    'underline': run.underline,
                    'font_name': run.font.name,
                    'font_size': run.font.size,
                    'font_color': run.font.color.rgb if run.font.color.rgb else None,
                    'highlight_color': run.font.highlight_color,
                }
                break
        
        # Clear existing runs
        for run in paragraph.runs:
            run.clear()
        
        # Remove all runs except the first one
        while len(paragraph.runs) > 1:
            paragraph._element.remove(paragraph.runs[-1]._element)
        
        # Apply translation to first run
        if paragraph.runs:
            first_run = paragraph.runs[0]
            first_run.text = translation
            
            # Apply original formatting
            if original_formatting:
                if original_formatting['bold'] is not None:
                    first_run.bold = original_formatting['bold']
                if original_formatting['italic'] is not None:
                    first_run.italic = original_formatting['italic']
                if original_formatting['underline'] is not None:
                    first_run.underline = original_formatting['underline']
                if original_formatting['font_name']:
                    first_run.font.name = original_formatting['font_name']
                if original_formatting['font_size']:
                    first_run.font.size = original_formatting['font_size']
                if original_formatting['font_color']:
                    first_run.font.color.rgb = original_formatting['font_color']
                if original_formatting['highlight_color']:
                    first_run.font.highlight_color = original_formatting['highlight_color']


def translate_table_cells(client: AzureOpenAI, model: str, table, target_language: str):
    """
    Translate all text in table cells using batch processing.
    
    Args:
        table: python-docx table object
    """
    all_paragraphs = []
    
    # Collect all paragraphs from all cells
    for row in table.rows:
        for cell in row.cells:
            all_paragraphs.extend(cell.paragraphs)
    
    if not all_paragraphs:
        return
    
    # Collect texts and translate in batches
    texts = collect_paragraph_texts(all_paragraphs)
    
    # Process in batches of 20 to avoid token limits
    batch_size = 20
    translated_texts = []
    
    for i in range(0, len(texts), batch_size):
        batch = texts[i:i + batch_size]
        batch_translations = translate_text_batch(client=client,
                                                  deployment_name=model,
                                                  texts=batch,
                                                  target_language=target_language)
        translated_texts.extend(batch_translations)
        
        # Small delay between batches to avoid rate limiting
        if i + batch_size < len(texts):
            time.sleep(0.5)
    
    # Apply translations
    apply_translations_to_paragraphs(all_paragraphs, translated_texts)


def translate_docx_document(client: AzureOpenAI, model: str, input_path: str, target_language: str, output_folder: str, save_as_pdf: bool) -> bool:
    """
    Translate an entire DOCX document while preserving formatting.
    
    Args:
        client: Azure OpenAI client
        model: chosen Azure OpenAI model deployment
        input_path: path to input .docx file
        target_language: language to translate to
        output_folder: output folder name
        save_as_pdf: indicator to save as pdf

    Returns:
        True if successful, False otherwise
    """
    try:
        file_name = os.path.basename(input_path)
        output_file_path = os.path.join(output_folder, target_language + "_" + file_name)
        # Load the document
        doc = Document(input_path)
        
        # Translate main document paragraphs
        if doc.paragraphs:
            texts = collect_paragraph_texts(doc.paragraphs)
            
            # Process in batches
            batch_size = 20
            translated_texts = []
            
            for i in range(0, len(texts), batch_size):
                batch = texts[i:i + batch_size]
                print(f"Processing batch {i//batch_size + 1}/{(len(texts)-1)//batch_size + 1}")
                batch_translations = translate_text_batch(client=client,
                                                          deployment_name=model,
                                                          texts=batch,
                                                          target_language=target_language)
                translated_texts.extend(batch_translations)
                
                # Small delay between batches
                if i + batch_size < len(texts):
                    time.sleep(0.5)
            
            apply_translations_to_paragraphs(doc.paragraphs, translated_texts)
        
        # Translate tables
        if doc.tables:
            print("Translating tables...")
            for table_idx, table in enumerate(doc.tables):
                print(f"Processing table {table_idx+1}/{len(doc.tables)}")
                translate_table_cells(client=client,
                                      model=model,
                                      table=table,
                                      target_language=target_language)
        
        # Translate headers and footers
        print("Translating headers and footers...")
        for section_idx, section in enumerate(doc.sections):
            print(f"  Processing section {section_idx+1}/{len(doc.sections)}")
            
            # Translate header
            if section.header and section.header.paragraphs:
                header_texts = collect_paragraph_texts(section.header.paragraphs)
                header_translations = translate_text_batch(client=client,
                                                           deployment_name=model,
                                                           texts=header_texts,
                                                           target_language=target_language)
                apply_translations_to_paragraphs(section.header.paragraphs, header_translations)
                
                # Translate header tables
                for table in section.header.tables:
                    translate_table_cells(client=client,
                                          model=model,
                                          table=table,
                                          target_language=target_language)
            
            # Translate footer
            if section.footer and section.footer.paragraphs:
                footer_texts = collect_paragraph_texts(section.footer.paragraphs)
                footer_translations = translate_text_batch(client=client,
                                                           deployment_name=model,
                                                           texts=footer_texts,
                                                           target_language=target_language)
                apply_translations_to_paragraphs(section.footer.paragraphs, footer_translations)
                
                # Translate footer tables
                for table in section.footer.tables:
                    translate_table_cells(client=client,
                                          model=model,
                                          table=table,
                                          target_language=target_language)
        
        # Save the translated document
        print(f"Saving translated document: {output_file_path}")
        doc.save(output_file_path)

        # if indicated, save as pdf file
        if save_as_pdf:
            pdf_file_name = os.path.splitext(file_name)[0] + ".pdf"
            pdf_file_path = os.path.join(output_folder, target_language + "_" + pdf_file_name)
            utils.convert_docx_to_pdf(output_file_path, pdf_file_path)
            # add watermark to created pdf file
            watermark_file_path = os.path.abspath(os.path.join(os.getcwd(), "watermark.pdf"))
            utils.add_watermark(pdf_file_path, pdf_file_path, watermark_file_path)
            # remove converted .docx file
            os.remove(output_file_path)

        return True
        
    except Exception as e:
        print(f"Error processing document: {e}")
        import traceback
        traceback.print_exc()
        return False
