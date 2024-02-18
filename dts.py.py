import tkinter as tk
from tkinter import filedialog
from docx import Document
import os
import pyperclip
from gtts import gTTS
from PyPDF2 import PdfReader

global address
def upload_word_document():
    file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    if file_path:
        entry.delete(0, tk.END)
        entry.insert(0, file_path)
def upload_pdf_document():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        entry2.delete(0, tk.END)
        entry2.insert(0, file_path)


def copy_text_word_document():
    global address
    address = entry.get()
    extracted_text = extract_text_from_word_document(address)
    if extracted_text:
        pyperclip.copy(extracted_text)  # Copy extracted text to clipboard
        status_label.config(text="Text copied to clipboard!")
        text_to_speech(extracted_text, language='en', file_name='output.mp3')  # Convert text to speech
    else:
        status_label.config(text="Failed to extract text from Word document!")

def extract_text_from_word_document(address):
    try:
        doc = Document(address)
        
        text = ""
        
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"  # Add each paragraph's text to the string
        
        return text.strip()  # Strip leading/trailing whitespace and return
    except Exception as e:
        print("Error:", e)
        return None
def copy_text_from_pdf():
    global address
    address = entry2.get()
    extracted_text = extract_text_from_pdf(address)
    if extracted_text:
        pyperclip.copy(extracted_text)  # Copy extracted text to clipboard
        status_label.config(text="Text copied to clipboard!")
        text_to_speech(extracted_text, language='en', file_name='output.mp3')  # Convert text to speech
    else:
        status_label.config(text="Failed to extract text from PDF document!")





def extract_text_from_pdf(address):
    try:
        with open(address, 'rb') as file:
            # Create a PdfReader object
            reader = PdfReader(file)
            # Initialize an empty string to store the text
            text = ""
            # Iterate through all pages in the PDF document
            for page in reader.pages:
                # Extract text from the page and append it to the string
                text += page.extract_text()
            # Return the extracted text after stripping leading/trailing whitespace
            return text.strip()
    except Exception as e:
        print("Error:", e)
        return None
    except PyPDF2.PdfReadError as e:
        print("Error: PDF is corrupted or incomplete -", e)
        return None
    except FileNotFoundError:
        print("Error: File not found")
        return None
    






def text_to_speech(text, language='en', file_name='output.mp3'):
    tts = gTTS(text=text, lang=language, slow=False)
    tts.save(file_name)
    os.system(f'start {file_name}')  # Opens the file with the default audio player


if __name__ == "__main__":
    
    

    root = tk.Tk()
    root.title("Word Document + PDF Reader")

    frame = tk.Frame(root)
    frame.pack(padx=30, pady=30)
    frame.configure(bg='blue')

    upload_button = tk.Button(frame, text="Upload Word Document", command=upload_word_document)
    upload_button.grid(row=0, column=0, padx=5, pady=5)

    entry = tk.Entry(frame, width=50)
    entry.grid(row=0, column=1, padx=5, pady=5)

    copy_button = tk.Button(frame, text="Read docx", command=copy_text_word_document)
    copy_button.grid(row=1, column=0, columnspan=2, pady=5)
#2
    upload_button2 = tk.Button(frame, text="Upload PDF", command=upload_pdf_document)
    upload_button2.grid(row=3, column=0, padx=5, pady=5)

    entry2 = tk.Entry(frame, width=50)
    entry2.grid(row=3, column=1, padx=5, pady=5)

    copy_button2= tk.Button(frame, text="Read PDF", command=copy_text_from_pdf)
    copy_button2.grid(row=4, column=0, columnspan=2, pady=5)

    

    status_label = tk.Label(root, text="")
    status_label.pack()

    root.mainloop()
