from flask import Flask, render_template, request, jsonify, send_file
import os
import io
import re
import json
import zipfile
from docxtpl import DocxTemplate

app = Flask(__name__)

DOCS_FOLDER = os.path.join(os.path.dirname(__file__), 'Docs')
RATES_FILE = os.path.join(os.path.dirname(__file__), 'rates.json')

def extract_placeholders(docx_path):
    """
    Scans entire DOCX package: Body, Headers, Footers & Footnotes.
    Returns a unified, unique list of placeholder tags.
    """
    found_placeholders = set()
    try:
        with zipfile.ZipFile(docx_path, 'r') as z:
            for filename in z.namelist():
                # Scan document body, all headers, footers and footnotes
                if filename.startswith('word/') and (
                    filename == 'word/document.xml' or 
                    'header' in filename or 
                    'footer' in filename or 
                    'footnotes' in filename
                ):
                    try:
                        xml_content = z.read(filename).decode('utf-8', errors='ignore')
                        # Strip XML tags to merge split runs
                        clean_text = re.sub(r'<[^>]+>', '', xml_content)
                        matches = re.findall(r'\{\{\s*([A-Za-z0-9_]+)\s*\}\}', clean_text)
                        found_placeholders.update(matches)
                    except Exception:
                        continue
        return sorted(list(found_placeholders))
    except Exception:
        return []

@app.route('/')
def home():
    return render_template('index.html')

# Appraisal Page Route
@app.route('/appraisal')
def appraisal_page():
    return render_template('appraisal.html')

# API: Templates aur Placeholders auto-load karne ke liye
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

# API: Benchmark Rates Sync (Server Side)
@app.route('/api/get-rates', methods=['GET'])
def get_rates():
    if os.path.exists(RATES_FILE):
        try:
            with open(RATES_FILE, 'r') as f:
                return jsonify(json.load(f))
        except Exception:
            pass
    return jsonify({"eblr": "9.25", "mclr": "8.60"})

@app.route('/api/save-rates', methods=['POST'])
def save_rates():
    try:
        data = request.json or {}
        eblr = str(data.get('eblr', '9.25')).strip()
        mclr = str(data.get('mclr', '8.60')).strip()
        with open(RATES_FILE, 'w') as f:
            json.dump({"eblr": eblr, "mclr": mclr}, f)
        return jsonify({"success": True, "eblr": eblr, "mclr": mclr})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

# API: Stable ZIP Generation (Headers, Footers & Body In-Memory Rendering)
@app.route('/generate-complete-zip', methods=['POST'])
def generate_complete_zip():
    data = request.json or {}
    form_data = data.get('formData', {})
    selected_docs = data.get('selectedDocs', [])
    
    borrower_name = form_data.get('BORROWER_NAME', 'Customer').strip().replace(' ', '_')
    borrower_name = re.sub(r'[^A-Za-z0-9_]', '', borrower_name) or 'Customer'

    zip_buffer = io.BytesIO()

    try:
        # Unfilled placeholders ko clean line se replace karna
        render_context = {
            k: (str(v).strip() if v is not None and str(v).strip() != "" else "______________________") 
            for k, v in form_data.items()
        }

        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as master_zip:
            for root, _, files in os.walk(DOCS_FOLDER):
                for file in files:
                    if file in selected_docs:
                        src_path = os.path.join(root, file)
                        
                        doc = DocxTemplate(src_path)
                        # docxtpl automatically renders body, headers, and footers
                        doc.render(render_context)
                        
                        # In-memory save (disk I/O bypassed for Render performance)
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        doc_io.seek(0)
                        
                        master_zip.writestr(f"Word_Files/{file}", doc_io.read())

        zip_buffer.seek(0)
        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name=f'Loan_Package_{borrower_name}.zip'
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)