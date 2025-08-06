"""
inspired by:
- https://medium.com/@pymupdf/translating-pdfs-a-practical-pymupdf-guide-c1c54b024042
- https://www.gradio.app/

Code generation supported by Claude
"""
import gradio as gr
import os
from openai import AzureOpenAI
from dotenv import load_dotenv
# local imports
import utils
import settings
import txt_translation
import docx_translation
import pdf_translation


def list_files_in_directory(directory_path):
    try:
        # List all files in the given directory
        files = os.listdir(directory_path)
        # Filter to include only files (not directories)
        files = [file for file in files if os.path.isfile(os.path.join(directory_path, file))]
        return gr.Dropdown(choices=files)
    except FileNotFoundError:
        return gr.Dropdown(choices=["Folder not found. Please enter a valid path."])
    except Exception as e:
        return gr.Dropdown(choices=[str(e)])


def process_translation(file_list, input_folder, target_language, save_as_pdf, progress=gr.Progress()):
    """
    Main processing function
    """
    status_text = ""
    
    if file_list is None:
        return "Please select one or more files first.", None
    
    if not target_language:
        return "Please select a target language.", None
    
    try:
        # if watermark file doesn't exist yet, create it
        if not os.path.exists("watermark.pdf"):
            utils.create_watermark("generated with PBL translator", "watermark.pdf")
        
        # if output folder doesn't exist yet, create it
        output_folder = os.path.join(input_folder, "translations")
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        # translation loop over all selected files
        for i, file_name in enumerate(file_list):
            print(i, file_name)
            file_path = os.path.join(input_folder, file_name)
            file_extension = os.path.splitext(file_path)[1].lower()
            status_message = f"Translating file {i+1}/{len(file_list)}: {file_name}"
            print(status_message)
            progress(i/len(file_list), desc=status_message)
            yield status_message

            if file_extension == ".txt":
                txt_translation.translate_txt_document(client=client,
                                                       model=AZURE_DEPLOYMENT_NAME,
                                                       input_path=file_path,
                                                       target_language=target_language,
                                                       output_folder=output_folder,
                                                       save_as_pdf=save_as_pdf)
            elif file_extension == ".docx":
                # print(f"Translating Word document: {file_name} to {target_language}")
                docx_translation.translate_docx_document(client=client,
                                                         model=AZURE_DEPLOYMENT_NAME,
                                                         input_path=file_path,
                                                         target_language=target_language,
                                                         output_folder=output_folder,
                                                         save_as_pdf=save_as_pdf)
            elif file_extension == ".pdf":
                # print(f"Translating PDF document: {file_name} to {target_language}")
                pdf_translation.translate_pdf_document(client=client,
                                                       model=AZURE_DEPLOYMENT_NAME,
                                                       input_path=file_path,
                                                       target_language=target_language,
                                                       output_path=output_folder)
            completed_message = f"Translation completed for {file_name}.\n"
            yield completed_message

        final_message = f"Done!, all files successfully translated to {target_language}."
        yield final_message
    
    except Exception as e:
        return f"Processing error: {str(e)}", None


# load Azure OpenAI api key
load_dotenv(dotenv_path="Y:/Kennisbasis/Datascience/ChatPBLenv/.env")

# Configure Azure OpenAI
AZURE_OPENAI_ENDPOINT = settings.AZURE_OPENAI_ENDPOINT
AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_VERSION = settings.AZURE_OPENAI_API_VERSION
AZURE_DEPLOYMENT_NAME = settings.AZURE_DEPLOYMENT_NAME

# Initialize Azure OpenAI client
client = AzureOpenAI(
    azure_endpoint=AZURE_OPENAI_ENDPOINT,
    api_key=AZURE_OPENAI_API_KEY,
    api_version=AZURE_OPENAI_VERSION,
    azure_deployment=AZURE_DEPLOYMENT_NAME
)

# Language options
LANGUAGES = [
    "Dutch", "German", "English", "Spanish", "French", "Italian", "Portuguese", 
    "Russian", "Chinese (Simplified)", "Chinese (Traditional)", "Japanese",
    "Korean", "Arabic", "Hindi", "Bengali", "Urdu", "Turkish",
    "Polish", "Czech", "Hungarian", "Romanian", "Bulgarian",
    "Greek", "Hebrew", "Thai", "Vietnamese", "Indonesian",
    "Malay", "Tagalog", "Swahili", "Amharic", "Yoruba"
]

# Create Gradio interface
with gr.Blocks(title="File Translation Tool", theme="soft") as demo:
    gr.Markdown("# PBL File Translation Tool")
    gr.Markdown("Upload a document and translate it to your chosen language using PBL's Azure OpenAI GPT-4o")
    # Instructions
    gr.Markdown("""
    ## Instructions:
    1. **Enter a folder path**. Copy-paste a file path from Windows File Explorer
    2. **Select one or more files to translate**: Supported file extensions are: .pdf, .docx, and .txt
    2. **Select target language** from the drop-down list of languages
    3. **Choose output format**: Check "Save as PDF" for PDF output (with watermark), otherwise saves in original format
    4. **Click "Translate Document"**: Translated files are written to subfolder "translations"
    NB: Formatting is preserved as much as possible, but not always possible. Text in images will not be translated.
    """)
    
    with gr.Row():
        with gr.Column():
            status_state = gr.State("")

            # Folder path entry point
            input_folder = gr.Textbox(label="Enter folder path")

            # Shows list of files to select when folder path is entered
            file_list = gr.Dropdown(label="Select file(s)", multiselect=True)

            # Language selection
            language_dropdown = gr.Dropdown(
                choices=LANGUAGES,
                label="Target Language",
                value="Dutch",
                interactive=True
            )
            
            # PDF save option
            save_pdf_checkbox = gr.Checkbox(
                label="Save as PDF",
                value=False
            )
            
            # Process button
            process_button = gr.Button(
                value="Translate Document",
                variant="primary",
                size="lg"
            )

            # Output status
            status_output = gr.Textbox(
                label="Status",
                interactive=True,
                lines=3
            )
            
    # when input is a folder path, a dropdown list file_list will be shown
    input_folder.change(fn=list_files_in_directory, inputs=input_folder, outputs=file_list)

    # Setup event handler: on click, selected files are being translated
    process_button.click(
        fn=process_translation,
        inputs=[file_list, input_folder, language_dropdown, save_pdf_checkbox],
        outputs=status_output
    )


# Launch the interface
if __name__ == "__main__":
    demo.launch(
        share=False,
        debug=True,
        server_name="127.0.0.1",
        server_port=7860
    )
