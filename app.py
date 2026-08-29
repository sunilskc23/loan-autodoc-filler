from flask import Flask, render_template, request, jsonify, send_file
import os
import io
import re
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET

app = Flask(__name__)

DOCS_FOLDER = os.path.join(os.path.dirname(__file__), 'Docs')
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
XML_NS = "{http://www.w3.org/XML/1998/namespace}"

def extract_placeholders(docx_path):
    try:
        with zipfile.ZipFile(docx_path, 'r') as z:
            xml_content = z.read('word/document.xml').decode('utf-8')
            clean_text = re.sub(r'<[^>]+>', '', xml_content)
            matches = re.findall(r'\{\{\s*([A-Za-z0-9_]+)\s*\}\}', clean_text)
            return list(set(matches))
    except Exception:
        return []

def safe_replace_xml_with_bold(xml_bytes, form_data):
    try:
        # Register namespaces to prevent Word XML syntax damage
        ET.register_namespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        ET.register_namespace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math')
        ET.register_namespace('v', 'urn:schemas-microsoft-com:vml')
        ET.register_namespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
        ET.register_namespace('w14', 'http://schemas.microsoft.com/office/word/2010/wordml')

        tree = ET.fromstring(xml_bytes)

        # Loop through all paragraphs
        for p in tree.iter(f"{W_NS}p"):
            p_children = list(p)
            for idx, r in enumerate(p_children):
                if r.tag != f"{W_NS}r":
                    continue

                t_nodes = r.findall(f"{W_NS}t")
                if not t_nodes:
                    continue

                full_run_text = "".join([t.text or "" for t in t_nodes])
                if "{{" not in full_run_text:
                    continue

                # Check which key is in this run
                matched_key = None
                for key in form_data:
                    placeholder = "{{" + key + "}}"
                    if placeholder in full_run_text:
                        matched_key = key
                        break

                if matched_key:
                    placeholder = "{{" + matched_key + "}}"
                    user_val = form_data.get(matched_key, "").strip()
                    is_filled = bool(user_val)
                    insert_text = f" {user_val} " if is_filled else "______________________"

                    parts = full_run_text.split(placeholder)
                    parent_idx = list(p).index(r)
                    p.remove(r)

                    existing_rPr = r.find(f"{W_NS}rPr")

                    insert_offset = 0
                    for p_i, part in enumerate(parts):
                        if part:
                            # Before/After normal text
                            new_r = ET.Element(f"{W_NS}r")
                            if existing_rPr is not None:
                                new_r.append(ET.fromstring(ET.tostring(existing_rPr)))
                            new_t = ET.SubElement(new_r, f"{W_NS}t")
                            new_t.set(f"{XML_NS}space", "preserve")
                            new_t.text = part
                            p.insert(parent_idx + insert_offset, new_r)
                            insert_offset += 1

                        # Insert filled value
                        if p_i < len(parts) - 1:
                            val_r = ET.Element(f"{W_NS}r")
                            val_rPr = ET.fromstring(ET.tostring(existing_rPr)) if existing_rPr is not None else ET.Element(f"{W_NS}rPr")
                            
                            # Apply BOLD only if filled by user
                            if is_filled:
                                if val_rPr.find(f"{W_NS}b") is None:
                                    ET.SubElement(val_rPr, f"{W_NS}b")
                            
                            if len(val_rPr) > 0 or is_filled:
                                val_r.append(val_rPr)

                            val_t = ET.SubElement(val_r, f"{W_NS}t")
                            val_t.set(f"{XML_NS}space", "preserve")
                            val_t.text = insert_text
                            p.insert(parent_idx + insert_offset, val_r)
                            insert_offset += 1

        return ET.tostring(tree, encoding='utf-8', xml_declaration=True)
    except Exception:
        return xml_bytes

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

    zip_buffer = io.BytesIO()

    try:
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as master_zip:
            for root, _, files in os.walk(DOCS_FOLDER):
                for file in files:
                    if file in selected_docs:
                        src_path = os.path.join(root, file)
                        
                        # Process DOCX cleanly in-memory
                        with zipfile.ZipFile(src_path, 'r') as in_docx:
                            out_docx_io = io.BytesIO()
                            with zipfile.ZipFile(out_docx_io, 'w', zipfile.ZIP_DEFLATED) as out_docx:
                                for item in in_docx.infolist():
                                    file_content = in_docx.read(item.filename)
                                    if item.filename in ["word/document.xml", "word/header1.xml", "word/header2.xml", "word/footer1.xml", "word/footer2.xml"]:
                                        file_content = safe_replace_xml_with_bold(file_content, form_data)
                                    out_docx.writestr(item, file_content)

                            out_docx_io.seek(0)
                            master_zip.writestr(f"Word_Files/{file}", out_docx_io.getvalue())

        zip_buffer.seek(0)
        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name=f'Loan_Package_{borrower_name}.zip'
        )

    except Exception as e:
        return {"error": str(e)}, 500

if __name__ == '__main__':
    app.run(debug=True)
