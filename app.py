import os
import re
import base64
import requests
import tempfile
import shutil
from flask import Flask, request, jsonify, send_file
from pptx import Presentation
import copy

app = Flask(__name__)

# ── Helpers ───────────────────────────────────────────────────────────────────

def sanitize_filename(name):
    """Remove characters invalid in OneDrive filenames and Graph API URLs."""
    name = re.sub(r'[<>:"/\\|?*&]', '-', name)
    name = re.sub(r'-+', '-', name)
    return name.strip('-')


# ── OneDrive helpers ──────────────────────────────────────────────────────────

def get_onedrive_access_token():
    client_id     = os.environ['ONEDRIVE_CLIENT_ID']
    client_secret = os.environ['ONEDRIVE_CLIENT_SECRET']
    refresh_token = os.environ['ONEDRIVE_REFRESH_TOKEN']

    response = requests.post(
        'https://login.microsoftonline.com/common/oauth2/v2.0/token',
        data={
            'client_id':     client_id,
            'client_secret': client_secret,
            'refresh_token': refresh_token,
            'grant_type':    'refresh_token',
            'scope':         'offline_access Files.ReadWrite.All',
        },
        timeout=30
    )
    response.raise_for_status()
    token_data = response.json()
    if 'refresh_token' in token_data:
        os.environ['ONEDRIVE_REFRESH_TOKEN'] = token_data['refresh_token']
    return token_data['access_token']


def upload_to_onedrive(file_path, filename, folder_id, access_token):
    CHUNK_SIZE = 10 * 1024 * 1024

    create_session_url = (
        f'https://graph.microsoft.com/v1.0/me/drive/items/{folder_id}:/{filename}:/createUploadSession'
    )
    session_response = requests.post(
        create_session_url,
        headers={
            'Authorization': f'Bearer {access_token}',
            'Content-Type':  'application/json',
        },
        json={
            'item': {
                '@microsoft.graph.conflictBehavior': 'rename',
                'name': filename,
            }
        },
        timeout=30
    )
    session_response.raise_for_status()
    upload_url = session_response.json()['uploadUrl']

    file_size = os.path.getsize(file_path)
    uploaded  = 0

    with open(file_path, 'rb') as f:
        while uploaded < file_size:
            chunk         = f.read(CHUNK_SIZE)
            chunk_len     = len(chunk)
            range_end     = uploaded + chunk_len - 1
            content_range = f'bytes {uploaded}-{range_end}/{file_size}'

            chunk_response = requests.put(
                upload_url,
                headers={
                    'Content-Length': str(chunk_len),
                    'Content-Range':  content_range,
                },
                data=chunk,
                timeout=120
            )
            if chunk_response.status_code not in (200, 201, 202):
                chunk_response.raise_for_status()
            uploaded += chunk_len

    result = chunk_response.json()
    return result.get('webUrl', result.get('id', 'uploaded'))


# ── PPTX helpers ──────────────────────────────────────────────────────────────

def download_file(url, dest_path):
    response = requests.get(url, allow_redirects=True, timeout=60)
    response.raise_for_status()
    with open(dest_path, 'wb') as f:
        f.write(response.content)


def _rewrite_rids(element, rId_map):
    """Recursively rewrite r:id, r:embed, r:link attributes using rId_map."""
    rId_attrs = [
        '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id',
        '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed',
        '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}link',
    ]
    for attr in rId_attrs:
        if attr in element.attrib:
            old = element.attrib[attr]
            if old in rId_map:
                element.attrib[attr] = rId_map[old]
    for child in element:
        _rewrite_rids(child, rId_map)


def merge_pptx_files(pptx_paths, output_path):
    """
    Merge multiple PPTX files preserving images.
    Loads base from pptx_paths[0], appends all slides from subsequent files,
    saves once at the end. No intermediate reload to avoid corruption.
    """
    if len(pptx_paths) == 1:
        shutil.copy(pptx_paths[0], output_path)
        return

    base_prs = Presentation(pptx_paths[0])

    for pptx_path in pptx_paths[1:]:
        src_prs = Presentation(pptx_path)

        for src_slide in src_prs.slides:
            # Add blank slide to base
            blank_layout = (
                base_prs.slide_layouts[6]
                if len(base_prs.slide_layouts) > 6
                else base_prs.slide_layouts[0]
            )
            new_slide = base_prs.slides.add_slide(blank_layout)

            # Remove placeholder shapes from blank layout
            for shape in list(new_slide.placeholders):
                sp = shape._element
                sp.getparent().remove(sp)

            # Copy all relationships from src slide → new slide, build rId map
            rId_map = {}
            for rId, rel in src_slide.part.rels.items():
                try:
                    if rel.is_external:
                        new_rId = new_slide.part.relate_to(
                            rel.target_ref, rel.reltype, is_external=True
                        )
                    else:
                        new_rId = new_slide.part.relate_to(
                            rel.target_part, rel.reltype
                        )
                    rId_map[rId] = new_rId
                except Exception:
                    pass

            # Deep-copy shapes and rewrite rIds in XML
            spTree = new_slide.shapes._spTree
            for shape in src_slide.shapes:
                el_copy = copy.deepcopy(shape._element)
                _rewrite_rids(el_copy, rId_map)
                spTree.append(el_copy)

        del src_prs

    # Save once — no intermediate reloads
    base_prs.save(output_path)


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'service': 'gamma-merger'})


@app.route('/merge', methods=['POST'])
def merge():
    data = request.get_json()
    if not data:
        return jsonify({'error': 'No JSON body provided'}), 400

    urls  = data.get('urls',  [])
    files = data.get('files', [])
    if not urls and not files:
        return jsonify({'error': 'Provide urls or files array'}), 400

    work_dir   = tempfile.mkdtemp()
    pptx_paths = []

    try:
        if urls:
            for i, url in enumerate(urls):
                dest = os.path.join(work_dir, f'part{i+1}.pptx')
                try:
                    download_file(url, dest)
                    pptx_paths.append(dest)
                except Exception as e:
                    return jsonify({'error': f'Failed to download part {i+1}', 'details': str(e)}), 500
        elif files:
            for i, b64 in enumerate(files):
                dest = os.path.join(work_dir, f'part{i+1}.pptx')
                try:
                    with open(dest, 'wb') as f:
                        f.write(base64.b64decode(b64))
                    pptx_paths.append(dest)
                except Exception as e:
                    return jsonify({'error': f'Failed to decode part {i+1}', 'details': str(e)}), 500

        if not pptx_paths:
            return jsonify({'error': 'No valid PPTX files to merge'}), 400

        output_path = os.path.join(work_dir, 'merged.pptx')
        merge_pptx_files(pptx_paths, output_path)

        return send_file(
            output_path,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            as_attachment=True,
            download_name='presentation_merged.pptx'
        )
    except Exception as e:
        return jsonify({'error': 'Merge failed', 'details': str(e)}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.route('/merge-and-upload', methods=['POST'])
def merge_and_upload():
    data = request.get_json()
    if not data:
        return jsonify({'error': 'No JSON body provided'}), 400

    urls      = data.get('urls', [])
    raw_name  = data.get('filename', 'merged_presentation.pptx')
    folder_id = os.environ.get('ONEDRIVE_FOLDER_ID', '')

    if not urls:
        return jsonify({'error': 'Provide urls array'}), 400
    if not folder_id:
        return jsonify({'error': 'ONEDRIVE_FOLDER_ID environment variable not set'}), 500

    filename = sanitize_filename(raw_name)
    if not filename.lower().endswith('.pptx'):
        filename += '.pptx'

    work_dir   = tempfile.mkdtemp()
    pptx_paths = []

    try:
        for i, url in enumerate(urls):
            dest = os.path.join(work_dir, f'part{i+1}.pptx')
            try:
                download_file(url, dest)
                pptx_paths.append(dest)
            except Exception as e:
                return jsonify({'error': f'Failed to download part {i+1}', 'details': str(e)}), 500

        if not pptx_paths:
            return jsonify({'error': 'No valid PPTX files to merge'}), 400

        output_path = os.path.join(work_dir, filename)
        merge_pptx_files(pptx_paths, output_path)

        try:
            access_token = get_onedrive_access_token()
        except Exception as e:
            return jsonify({'error': 'Failed to get OneDrive access token', 'details': str(e)}), 500

        try:
            web_url = upload_to_onedrive(output_path, filename, folder_id, access_token)
        except Exception as e:
            return jsonify({'error': 'Failed to upload to OneDrive', 'details': str(e)}), 500

        return jsonify({
            'status':   'uploaded',
            'filename': filename,
            'web_url':  web_url,
        })

    except Exception as e:
        return jsonify({'error': 'merge-and-upload failed', 'details': str(e)}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
