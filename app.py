import os
import base64
import requests
import tempfile
import shutil
from flask import Flask, request, jsonify, send_file
from pptx import Presentation
import copy

app = Flask(__name__)

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'service': 'gamma-merger'})

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
            blank_layout = base_prs.slide_layouts[6] if len(base_prs.slide_layouts) > 6 else base_prs.slide_layouts[0]
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

@app.route('/merge', methods=['POST'])
def merge():
    """
    Merge multiple PPTX files into one.

    Accepts JSON body with either:
      { "urls": ["https://...", "https://..."] }
      { "files": ["base64string", "base64string"] }

    Returns the merged PPTX as a binary download.
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


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
