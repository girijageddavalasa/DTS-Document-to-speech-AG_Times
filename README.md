# DTS-Document-to-speech-AG_Times

# Word Document + PDF Reader

This Python script provides a simple graphical user interface (GUI) to read text from both Word documents (.docx) and PDF files. It utilizes the `tkinter` library for the GUI, `pyperclip` for copying text to the clipboard, `docx` for handling Word documents, `PyPDF2` for handling PDF files, and `gtts` for converting text to speech.

## How to Use

1. **Upload Word Document**: Click on the "Upload Word Document" button to select a Word document (.docx) from your file system. The path of the selected file will be displayed in the adjacent entry field.

2. **Read docx**: Click on the "Read docx" button to extract text from the uploaded Word document. The extracted text will be copied to the clipboard, and an audio file named `output.mp3` will be generated with the text-to-speech conversion.

3. **Upload PDF**: Click on the "Upload PDF" button to select a PDF file from your file system. The path of the selected file will be displayed in the adjacent entry field.

4. **Read PDF**: Click on the "Read PDF" button to extract text from the uploaded PDF document. Similar to reading from a Word document, the extracted text will be copied to the clipboard, and an audio file named `output.mp3` will be generated with the text-to-speech conversion.

## Requirements

- Python 3.x
- tkinter
- docx
- pyperclip
- gtts
- yPDF2

## Running the Script

Ensure you have Python installed on your system along with the required libraries. Then, simply execute the script. It will launch a GUI window where you can perform the aforementioned actions.

```
python your_script_name.py
```

## Note

- The script assumes valid Word documents (.docx) and PDF files. If the files are corrupted or incomplete, errors may occur.
- Make sure to have a default audio player set up on your system to play the generated `output.mp3` files.
