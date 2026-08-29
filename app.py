from flask import Flask, render_template, request, jsonify, send_file
import os
import io
import re
import shutil
import tempfile
import zipfile
from docxtpl import DocxTemplate, RichText

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

    try:
        # Prepare Context: Only User Filled Data is BOLD via RichText
        render_context = {}
        for key, val in form_data.items():
            if val and str(val).strip():
                # User filled value will be strictly BOLD
                render_context[key] = RichText(f" {str(val).strip()} ", bold=True)
            else:
                # Unfilled field remains normal blank line
                render_context[key] = "______________________"

        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as master_zip:
            for root, _, files in os.walk(DOCS_FOLDER):
                for file in files:
                    if file in selected_docs:
                        src_path = os.path.join(root, file)
                        doc = DocxTemplate(src_path)
                        
                        # Render with Bold RichText
                        doc.render(render_context)
                        
                        # Save filled DOCX
                        filled_docx_path = os.path.join(temp_dir, file)
                        doc.save(filled_docx_path)
                        
                        # Add filled DOCX to ZIP
                        with open(filled_docx_path, 'rb') as f:
                            master_zip.writestr(f"Word_Files/{file}", f.read())

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
