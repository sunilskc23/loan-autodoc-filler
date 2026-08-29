from flask import Flask, render_template, request, send_file
import os
import io
import shutil
import tempfile
from docxtpl import DocxTemplate
from pypdf import PdfMerger
# Windows local & cloud converter support
try:
    from docx2pdf import convert
except ImportError:
    convert = None

app = Flask(__name__)

DOCS_FOLDER = os.path.join(os.path.dirname(__file__), 'Docs')

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/generate-merged-pdf', methods=['POST'])
def generate_merged_pdf():
    data = request.json
    form_data = data.get('formData', {})
    selected_docs = data.get('selectedDocs', [])
    borrower_name = form_data.get('BORROWER_NAME', 'Customer').strip().replace(' ', '_')
    if not borrower_name:
        borrower_name = 'Customer'

    temp_dir = tempfile.mkdtemp()
    merger = PdfMerger()

    try:
        pdf_paths = []
        
        # 1. Scan and fill matched documents
        for root, _, files in os.walk(DOCS_FOLDER):
            for file in sorted(files):
                if file.endswith('.docx') and not file.startswith('~$'):
                    file_upper = file.upper()
                    # Check if document code is matched in selectedDocs
                    if any(doc_code.upper() in file_upper for doc_code in selected_docs):
                        src_path = os.path.join(root, file)
                        doc = DocxTemplate(src_path)
                        
                        # Replace placeholders (Blank lines for empty fields)
                        render_context = {k: (v if v else "______________________") for k, v in form_data.items()}
                        doc.render(render_context)
                        
                        filled_docx_path = os.path.join(temp_dir, file)
                        doc.save(filled_docx_path)
                        
                        # Convert to PDF
                        pdf_file_name = file.replace('.docx', '.pdf')
                        pdf_path = os.path.join(temp_dir, pdf_file_name)
                        
                        # Convert using docx2pdf
                        convert(filled_docx_path, pdf_path)
                        
                        if os.path.exists(pdf_path):
                            pdf_paths.append(pdf_path)

        # 2. Merge all individual PDFs into Single PDF
        for p in pdf_paths:
            merger.append(p)

        merged_pdf_io = io.BytesIO()
        merger.write(merged_pdf_io)
        merger.close()
        merged_pdf_io.seek(0)

        return send_file(
            merged_pdf_io,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f'Loan_Set_{borrower_name}.pdf'
        )

    except Exception as e:
        return {"error": str(e)}, 500
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

if __name__ == '__main__':
    app.run(debug=True, port=5000)