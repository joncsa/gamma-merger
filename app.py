import os
import base64
import requests
import tempfile
import shutil
from flask import Flask, request, jsonify, send_file
from pptx import Presentation
import copy

app = Flask(__name__)

# ── OneDrive helpers ──────────────────────────────────────────────────────────

def get_onedrive_access_token():
    """Get a fresh OneDrive access token using the refresh token."""
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

    # Persist the new refresh token back to the environment so it stays fresh
    if 'refresh_token' in token_data:
        os.environ['ONEDRIVE_REFRESH_TOKEN'] = token_data['refresh_token']

    return token_data['access_token']


def upload_to_onedrive(file_path, filename, folder_id, access_token):
    """
    Upload a file to OneDrive using the resumable (chunked) upload session.
    Handles files of any size — no memory bottleneck.
    Returns the OneDrive web URL of the uploaded file.
    """
    CHUNK_SIZE = 10 * 1024 * 1024  # 10 MB chunks

    # Create an upload session
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

    # Upload in chunks
    file_size = os.path.getsize(file_path)
    uploaded  = 0

    with open(file_path, 'rb') as f:
        while uploaded < file_size:
            chunk = f.read(CHUNK_SIZE)
            chunk_len   = len(chunk)
            range_end   = uploaded + chunk_len - 1
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

            # 202 = still uploading, 200/201 = complete
            if chunk_response.status_code not in (200, 201, 202):
                chunk_response.raise_for_status()

            uploaded += chunk_len

    result = chunk_response.json()
    return result.get('webUrl', result.get('id', 'uploaded'))


# ── PPTX helpers ──────────────────────────────────────────────────────────────

def download_file(url, dest_path):
    """Download a file from URL, following redirects."""
    response = requests.get(url, allow_redirects=True, timeout=60)
    response.raise_for_status()
    with open(dest_path, 'wb') as f:
        f.write(response.content)


def merge_pptx_files(pptx_paths, output_path):
    """Merge multiple PPTX files into one using python-pptx, sequentially."""
    if len(pptx_paths) == 1:
        shutil.copy(pptx_paths[0], output_path)
        return

    base_prs = Presentation(pptx_paths[0])

    for pptx_path in pptx_paths[1:]:
        src_prs = Presentation(pptx_path)
        for slide in src_prs.slides:
            blank_layout = (
                base_prs.slide_layouts[6]
                if len(base_prs.slide_layouts) > 6
                else base_prs.slide_layouts[0]
            )
            new_slide = base_prs.slides.add_slide(blank_layout)
            for shape in list(new_slide.placeholders):
                sp = shape._element
                sp.getparent().remove(sp)
            for shape in slide.shapes:
                el_copy = copy.deepcopy(shape._element)
                new_slide.shapes._spTree.append(el_copy)
        # Save after each merge and reload to free memory
        base_prs.save(output_path)
        del src_prs
        base_prs = Presentation(output_path)

    base_prs.save(output_path)


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'service': 'gamma-merger'})


@app.route('/merge', methods=['POST'])
def merge():
    """
    Merge multiple PPTX files into one and return as binary download.
    Body: { "urls": [...] }  or  { "files": ["base64", ...] }
    """
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
    """
    Merge multiple PPTX files and upload directly to OneDrive.
    No large binary is ever passed through n8n memory.

    Body:
    {
      "urls":     ["https://...", ...],   // Gamma export URLs
      "filename": "My Course.pptx"        // Desired filename in OneDrive
    }

    Returns:
    {
      "status":   "uploaded",
      "filename": "My Course.pptx",
      "web_url":  "https://onedrive.live.com/..."
    }
    """
    data = request.get_json()
    if not data:
        return jsonify({'error': 'No JSON body provided'}), 400

    urls     = data.get('urls', [])
    filename = data.get('filename', 'merged_presentation.pptx')
    folder_id = os.environ.get('ONEDRIVE_FOLDER_ID', '')

    if not urls:
        return jsonify({'error': 'Provide urls array'}), 400
    if not folder_id:
        return jsonify({'error': 'ONEDRIVE_FOLDER_ID environment variable not set'}), 500

    # Ensure filename ends with .pptx
    if not filename.lower().endswith('.pptx'):
        filename += '.pptx'

    work_dir   = tempfile.mkdtemp()
    pptx_paths = []

    try:
        # Step 1 — Download all PPTX parts from Gamma
        for i, url in enumerate(urls):
            dest = os.path.join(work_dir, f'part{i+1}.pptx')
            try:
                download_file(url, dest)
                pptx_paths.append(dest)
            except Exception as e:
                return jsonify({'error': f'Failed to download part {i+1}', 'details': str(e)}), 500

        if not pptx_paths:
            return jsonify({'error': 'No valid PPTX files to merge'}), 400

        # Step 2 — Merge into one file on disk
        output_path = os.path.join(work_dir, filename)
        merge_pptx_files(pptx_paths, output_path)

        # Step 3 — Get fresh OneDrive access token
        try:
            access_token = get_onedrive_access_token()
        except Exception as e:
            return jsonify({'error': 'Failed to get OneDrive access token', 'details': str(e)}), 500

        # Step 4 — Upload directly to OneDrive via chunked upload session
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
