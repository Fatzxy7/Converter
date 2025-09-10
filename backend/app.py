from flask import Flask, request, send_file, render_template
from werkzeug.utils import secure_filename
import os
from pdf2docx import Converter
from PIL import Image
import docx2txt
import pandas as pd
from fpdf import FPDF
from moviepy.editor import AudioFileClip, VideoFileClip

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
CONVERTED_FOLDER = 'converted'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    file = request.files['file']
    action = request.form.get('action')
    if not file or not action:
        return "No file or action selected", 400

    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    name, ext = os.path.splitext(filename)
    output_file = os.path.join(CONVERTED_FOLDER, f"{name}_{action}")

    try:
        # PDF → DOCX
        if ext.lower() == ".pdf" and action == "pdf_to_docx":
            output_file += ".docx"
            cv = Converter(filepath)
            cv.convert(output_file, start=0, end=None)
            cv.close()

        # PDF → TXT
        elif ext.lower() == ".pdf" and action == "pdf_to_txt":
            output_file += ".txt"
            # Simple PDF → TXT, bisa pakai PyPDF2
            import PyPDF2
            reader = PyPDF2.PdfReader(filepath)
            text = ""
            for page in reader.pages:
                text += page.extract_text() + "\n"
            with open(output_file, "w", encoding="utf-8") as f:
                f.write(text)

        # TXT → PDF
        elif ext.lower() == ".txt" and action == "txt_to_pdf":
            output_file += ".pdf"
            pdf = FPDF()
            pdf.add_page()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.set_font("Arial", size=12)
            with open(filepath, "r", encoding="utf-8") as f:
                for line in f:
                    pdf.cell(0, 10, line.strip(), ln=True)
            pdf.output(output_file)

        # DOCX → TXT
        elif ext.lower() == ".docx" and action == "docx_to_txt":
            output_file += ".txt"
            text = docx2txt.process(filepath)
            with open(output_file, "w", encoding="utf-8") as f:
                f.write(text)

        # JPG/PNG → PDF
        elif ext.lower() in [".jpg", ".jpeg", ".png"] and action == "img_to_pdf":
            output_file += ".pdf"
            image = Image.open(filepath)
            image.convert("RGB").save(output_file)

        # XLSX → CSV
        elif ext.lower() == ".xlsx" and action == "xlsx_to_csv":
            output_file += ".csv"
            df = pd.read_excel(filepath)
            df.to_csv(output_file, index=False)

        # MP4 → MP3
        elif ext.lower() == ".mp4" and action == "mp4_to_mp3":
            output_file += ".mp3"
            clip = VideoFileClip(filepath)
            clip.audio.write_audiofile(output_file)
            clip.close()

        # MP3 → WAV
        elif ext.lower() == ".mp3" and action == "mp3_to_wav":
            output_file += ".wav"
            clip = AudioFileClip(filepath)
            clip.write_audiofile(output_file)
            clip.close()

        else:
            return "Conversion type not supported", 400

        return send_file(output_file, as_attachment=True)

    except Exception as e:
        return str(e), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
