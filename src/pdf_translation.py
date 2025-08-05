"""
PDF Translation with Azure OpenAI and Formatting Preservation

This script reads a .pdf file, translates the text content using Azure OpenAI
while preserving formatting, fonts, and layout as much as possible, then outputs a new translated .pdf file.

"""
import os
import pymupdf
from openai import AzureOpenAI
import utils
from pathlib import Path


def translate_text(client: AzureOpenAI, model: str, text: str, target_language: str) -> str:
    """Translate text using Azure OpenAI GPT-4o"""
    try:
        # Split text into chunks if it's too long (to handle token limits)
        max_chunk_size = 3000  # Conservative chunk size
        chunks = [text[i:i+max_chunk_size] for i in range(0, len(text), max_chunk_size)]
        
        translated_chunks = []
        
        for chunk in chunks:
            if not chunk.strip():
                continue
                
            prompt = f"""Translate the following text to {target_language}. 
            Maintain the original formatting, structure, and meaning as much as possible.
            Only return the translated text, no additional comments or explanations.
            
            Text to translate:
            {chunk}
            """
            
            response = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": "You are a professional translator. Translate accurately while preserving formatting and context."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=4000,
                temperature=0
            )
            
            translated_chunks.append(response.choices[0].message.content)
        
        return '\n'.join(translated_chunks)
        
    except Exception as e:
        return f"Translation error: {str(e)}"


def translate_pdf_document(client: AzureOpenAI, model: str, input_path: str, target_language: str, output_path: str) -> bool:
    """
    Translate an entire pdf document while preserving formatting.
    
    Args:
        input_path: Path to input .pdf file
        output_path: Path to output .pdf file
        
    Returns:
        True if successful, False otherwise
    """
    try:
        file_name = os.path.basename(input_path)
        output_file_path = os.path.join(output_path, target_language + "_" + file_name)
        # utils.add_watermark(file_name, output_pdf, watermark_pdf)
        # Define color "white"
        WHITE = pymupdf.pdfcolor["white"]
        # This flag ensures that text will be dehyphenated after extraction.
        textflags = pymupdf.TEXT_DEHYPHENATE
        # Open the document
        doc = pymupdf.open(input_path)
        # Define an Optional Content layer in the document named "translation", and activate it by default.
        ocg_xref = doc.add_ocg("translation", on=True)
        # Iterate over all pages
        for i, page in enumerate(doc.pages()):
            status_text = f"Translating page {i} of file {file_name}..."
            # yield status_text
            # Extract text grouped like lines in a paragraph.
            blocks = page.get_text("blocks", flags=textflags)
            # Every block of text is contained in a rectangle ("bbox")
            for block in blocks:
                # area containing the text
                bbox = block[:4]
                # the text of this block
                block_text = block[4]
                # Invoke the actual translation
                translated_text = translate_text(client=client,
                                                 model=model,
                                                 text=block_text,
                                                 target_language=target_language)
                # Cover the original text with a white rectangle.
                page.draw_rect(bbox, color=None, fill=WHITE, oc=ocg_xref)
                # Write the translated text into the rectangle
                page.insert_htmlbox(bbox, translated_text, oc=ocg_xref)
        # save file to output folder and add watermark
        doc.ez_save(output_file_path)
        doc.close()
        utils.add_watermark(output_file_path, output_file_path, "X:/User/troosts/projects/translator/watermark.pdf")

        return True
        
    except Exception as e:
        print(f"Error processing document: {e}")
        import traceback
        traceback.print_exc()
        
        return False

