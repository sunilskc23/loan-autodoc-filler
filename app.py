from flask import Flask, render_template, request, jsonify, send_file
import os
import io
import re
import shutil
import tempfile
import zipfile
import subprocess
from docxtpl import DocxTemplate
from pypdf import PdfMerger

# ReportLab Fallback Engine for 100% Guaranteed Cloud PDF rendering
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from docx import Document

app = Flask(__name__)

DOCS_FOLDER = os.path.join(os.path.dirname(__file__), 'Docs')

def extract_placeholders(docx_path):
    try:
        with zipfile.ZipFile(docx_path, 'r') as z:
            xml_content = z.read('word/document.xml').decode('utf-8')
            clean_text = re.sub(r'<[^>]+>', '', xml_content)
            matches = re.findall(r'\{\{\s*([A-Za-z0-9_]+)\s*\}\}', clean_text)
            return list(set(matches))
    except Exception:
        return []

def docx_to_pdf_reportlab(docx_path, pdf_path):
    try:
        doc = Document(docx_path)
        pdf_doc = SimpleDocTemplate(
            pdf_path,
            pagesize=letter,
            rightMargin=36,
            leftMargin=36,
            topMargin=36,
            bottomMargin=36
        )
        
        styles = getSampleStyleSheet()
        normal_style = styles['Normal']
        normal_style.fontSize = 10
        normal_style.leading = 14

        story = []
        for p in doc.paragraphs:
            text = p.text.strip()
            if text:
                # Basic bold rendering for generated text
                safe_text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                story.append(Paragraph(safe_text, normal_style))
                story.append(Spacer(1, 6))
            else:
                story.append(Spacer(1, 8))

        pdf_doc.build(story)
        return True
    except Exception as err:
        print("Reportlab error:", err)
        return False

def convert_to_pdf(doc_path, output_dir):
    pdf_name = os.path.basename(doc_path).replace('.docx', '.pdf')
    pdf_path = os.path.join(output_dir, pdf_name)

    # 1. Try LibreOffice if available in server image
    try:
        res = subprocess.run([
            'libreoffice', '--headless', '--convert-to', 'pdf',
            doc_path, '--outdir', output_dir
        ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=15)
        if os.path.exists(pdf_path):
            return pdf_path
    except Exception:
        pass

    # 2. Try Windows MS Word COM API (if running on local machine)
    try:
        from docx2pdf import convert
        convert(doc_path, pdf_path)
        if os.path.exists(pdf_path):
            return pdf_path
    except Exception:
        pass

    # 3. Pure Python Cloud Engine (Always succeeds on Render Free tier)
    if docx_to_pdf_reportlab(doc_path, pdf_path) and os.path.exists(pdf_path):
        return pdf_path

    return None

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/api/get-docs', methods=['GET'])
def get_docs():
    files_list = []
    if os.path.exists(DOCS_FOLDER):
        for root, _, files in os.walk(DOCS_FOLDER):
            for file in sorted(files):
                if file.endswith('.docx') and not file.startswith('~$'):
                    full_path = os.path.join(root, file)
                    rel_dir = os.path.relpath(root, DOCS_FOLDER)
                    category = 'OtherDocs' if 'OtherDocs' in rel_dir else 'SecurityDocs'
                    placeholders = extract_placeholders(full_path)
                    
                    files_list.append({
                        "name": file,
                        "categoryFolder": category,
                        "placeholders": placeholders,
                        "checked": False
                    })
    return jsonify(files_list)

@app.route('/generate-complete-zip', methods=['POST'])
def generate_complete_zip():
    data = request.json
    form_data = data.get('formData', {})
    selected_docs = data.get('selectedDocs', [])
    borrower_name = form_data.get('BORROWER_NAME', 'Customer').strip().replace(' ', '_')
    if not borrower_name:
        borrower_name = 'Customer'

    temp_dir = tempfile.mkdtemp()
    zip_buffer = io.BytesIO()
    pdf_merger = PdfMerger()
    generated_pdf_paths = []

    try:
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as master_zip:
            for root, _, files in os.walk(DOCS_FOLDER):
                for file in sorted(files):
                    if file in selected_docs:
                        src_path = os.path.join(root, file)
                        doc = DocxTemplate(src_path)
                        
                        # Empty fields replace with blank lines
                        render_context = {k: (v if v else "______________________") for k, v in form_data.items()}
                        doc.render(render_context)
                        
                        # 1. Save & Zip DOCX
                        filled_docx_path = os.path.join(temp_dir, file)
                        doc.save(filled_docx_path)
                        
                        with open(filled_docx_path, 'rb') as f:
                            master_zip.writestr(f"Word_Files/{file}", f.read())
                        
                        # 2. Convert to Individual PDF & Zip
                        pdf_path = convert_to_pdf(filled_docx_path, temp_dir)
                        if pdf_path and os.path.exists(pdf_path):
                            pdf_file_name = os.path.basename(pdf_path)
                            with open(pdf_path, 'rb') as pf:
                                master_zip.writestr(f"PDF_Files/{pdf_file_name}", pf.read())
                            
                            generated_pdf_paths.append(pdf_path)

            # 3. Create ALL-IN-ONE Merged PDF booklet at Root of ZIP
            if generated_pdf_paths:
                for p_path in generated_pdf_paths:
                    pdf_merger.append(p_path)
                
                merged_pdf_path = os.path.join(temp_dir, f"All_In_One_Loan_Booklet_{borrower_name}.pdf")
                pdf_merger.write(merged_pdf_path)
                pdf_merger.close()

                if os.path.exists(merged_pdf_path):
                    with open(merged_pdf_path, 'rb') as mf:
                        master_zip.writestr(f"All_In_One_Loan_Booklet_{borrower_name}.pdf", mf.read())

        zip_buffer.seek(0)
        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name=f'Loan_Package_{borrower_name}.zip'
        )

    except Exception as e:
        return {"error": str(e)}, 500
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

if __name__ == '__main__':
    app.run(debug=True)
