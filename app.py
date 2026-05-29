import os
import json
import base64
import requests
import tempfile
import shutil
from flask import Flask, request, jsonify
from pptx import Presentation
import copy

app = Flask(__name__)

def download_file(url, dest_path):
    response = requests.get(url, allow_redirects=True, timeout=60)
    response.raise_for_status()
    with open(dest_path, 'wb') as f:
        f.write(response.content)

def merge_pptx_files(pptx_paths, output_path):
    if len(pptx_paths) == 1:
        shutil.copy(pptx_paths[0], output_path)
        return

    base_prs = Presentation(pptx_paths[0])

    for pptx_path in pptx_paths[1:]:
        src_prs = Presentation(pptx_path)

        for slide in src_prs.slides:
            blank_layout = base_prs.slide_layouts[6] if len(base_prs.slide_layouts) > 6 else base_prs.slide_layouts[0]
            new_slide = base_prs.slides.add_slide(blank_layout)

            for shape in list(new_slide.placeholders):
                sp = shape._element
                sp.getparent().remove(sp)

            for shape in slide.shapes:
                el_copy = copy.deepcopy(shape._element)
                new_slide.shapes._spTree.append(el_copy)

        base_prs.save(output_path)
        del src_prs
        base_prs = Presentation(output_path)

    base_prs.save(output_path)

def get_onedrive_token():
    client_id     = os.environ['ONEDRIVE_CLIENT_ID']
    client_secret = os.environ['ONEDRIVE_CLIENT_SECRET']
    refresh_token = os.environ['ONEDRIVE_REFRESH_TOKEN']

    resp = requests.post(
        'https://login.microsoftonline.com/common/oauth2/v2.0/token',
        data={
            'grant_type':    'refresh_token',
            'client_id':     client_id,
            'client_secret': client_secret,
            'refresh_token': refresh_token,
            'scope':         'https://graph.microsoft.com/.default'
        }
    )
    resp.raise_for_status()
    return resp.json()['access_token']

def upload_to_onedrive(file_path, filename, folder_id, token):
    file_size = os.path.getsize(file_path)
    headers = {'Authorization': f'Bearer {token}'}

    # Create upload session for large files
    session_url = f'https://graph.microsoft.com/v1.0/me/drive/items/{folder_id}:/{filename}:/createUploadSession'
    session_resp = requests.post(session_url, headers=headers, json={
        'item': {'@microsoft.graph.conflictBehavior': 'rename'}
    })
    session_resp.raise_for_status()
    upload_url = session_resp.json()['uploadUrl']

    # Upload in 10MB chunks
    chunk_size = 10 * 1024 * 1024
    uploaded = 0

    with open(file_path, 'rb') as f:
        while uploaded < file_size:
            chunk = f.read(chunk_size)
            chunk_len = len(chunk)
            headers_chunk = {
                'Content-Length': str(chunk_len),
                'Content-Range': f'bytes {uploaded}-{uploaded + chunk_len - 1}/{file_size}'
            }
            resp = requests.put(upload_url, headers=headers_chunk, data=chunk)
            uploaded += chunk_len

    # Get the file URL
    item_resp = requests.get(
        f'https://graph.microsoft.com/v1.0/me/drive/items/{folder_id}:/{filename}',
        headers={'Authorization': f'Bearer {token}'}
    )
    item_resp.raise_for_status()
    return item_resp.json().get('webUrl')

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'service': 'gamma-merger'})

@app.route('/merge', methods=['POST'])
def merge():
    data = request.get_json()
    if not data:
        return jsonify({'error': 'No JSON body provided'}), 400

    urls      = data.get('urls', [])
    files     = data.get('files', [])
    filename  = data.get('filename', 'presentation_merged.pptx')
    folder_id = data.get('folder_id') or os.environ.get('ONEDRIVE_FOLDER_ID')

    if not urls and not files:
        return jsonify({'error': 'Provide urls or files array'}), 400

    work_dir   = tempfile.mkdtemp()
    pptx_paths = []

    try:
        if urls:
            for i, url in enumerate(urls):
                dest = os.path.join(work_dir, f'part{i+1}.pptx')
                download_file(url, dest)
                pptx_paths.append(dest)
        elif files:
            for i, b64 in enumerate(files):
                dest = os.path.join(work_dir, f'part{i+1}.pptx')
                with open(dest, 'wb') as f:
                    f.write(base64.b64decode(b64))
                pptx_paths.append(dest)

        output_path = os.path.join(work_dir, 'merged.pptx')
        merge_pptx_files(pptx_paths, output_path)

        token    = get_onedrive_token()
        file_url = upload_to_onedrive(output_path, filename, folder_id, token)

        return jsonify({'url': file_url, 'filename': filename})

    except Exception as e:
        return jsonify({'error': 'Merge failed', 'details': str(e)}), 500

    finally:
        shutil.rmtree(work_dir, ignore_errors=True)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
