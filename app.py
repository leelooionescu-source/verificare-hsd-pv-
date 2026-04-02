import os
import uuid
import shutil
import time
import threading
from flask import Flask, render_template, request, redirect, url_for, session, send_from_directory, flash, jsonify
from process_verify import read_master, process_all, generate_report

app = Flask(__name__)
app.secret_key = 'verificare-hsd-pv-secret-2026'
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200MB

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOADS_DIR = os.path.join(BASE_DIR, 'uploads')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
os.makedirs(UPLOADS_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Store processing progress and results
progress_store = {}
results_store = {}


def get_session_dir(subdir):
    if 'session_id' not in session:
        session['session_id'] = str(uuid.uuid4())
    path = os.path.join(BASE_DIR, subdir, session['session_id'])
    os.makedirs(path, exist_ok=True)
    return path


def cleanup_old_sessions():
    cutoff = time.time() - 86400
    for base in [UPLOADS_DIR, OUTPUT_DIR]:
        if not os.path.exists(base):
            continue
        for d in os.listdir(base):
            path = os.path.join(base, d)
            if os.path.isdir(path) and os.path.getmtime(path) < cutoff:
                shutil.rmtree(path, ignore_errors=True)


@app.route('/')
def index():
    cleanup_old_sessions()
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload():
    master_file = request.files.get('master')
    if not master_file or master_file.filename == '':
        flash('Fișierul MASTER este obligatoriu!', 'error')
        return redirect(url_for('index'))

    doc_files = request.files.getlist('docs')
    if not doc_files or all(f.filename == '' for f in doc_files):
        flash('Încărcați cel puțin un fișier Word (.docx)!', 'error')
        return redirect(url_for('index'))

    upload_dir = get_session_dir('uploads')
    # Clear previous
    shutil.rmtree(upload_dir, ignore_errors=True)
    os.makedirs(upload_dir, exist_ok=True)

    docs_dir = os.path.join(upload_dir, 'docs')
    os.makedirs(docs_dir, exist_ok=True)

    master_path = os.path.join(upload_dir, 'master.xlsx')
    master_file.save(master_path)

    doc_count = 0
    doc_paths = []
    for f in doc_files:
        if f.filename and f.filename.lower().endswith('.docx'):
            fpath = os.path.join(docs_dir, f.filename)
            f.save(fpath)
            doc_paths.append(fpath)
            doc_count += 1

    session['master_path'] = master_path
    session['doc_paths'] = doc_paths
    session['doc_count'] = doc_count

    return redirect(url_for('process'))


@app.route('/process')
def process():
    if 'master_path' not in session:
        flash('Încărcați fișierele mai întâi.', 'error')
        return redirect(url_for('index'))
    return render_template('processing.html', doc_count=session.get('doc_count', 0))


@app.route('/run-process', methods=['POST'])
def run_process():
    if 'master_path' not in session:
        return jsonify({'error': 'No files'}), 400

    master_path = session['master_path']
    doc_paths = session['doc_paths']
    sid = session.get('session_id', '')
    output_dir_path = get_session_dir('output')

    progress_store[sid] = {'current': 0, 'total': 0, 'file': 'Pornire...', 'done': False}

    results_store[sid] = None

    def background_process():
        def progress_cb(current, total, filename):
            progress_store[sid] = {'current': current, 'total': total, 'file': filename, 'done': False}

        try:
            results, master_entries = process_all(master_path, doc_paths, progress_callback=progress_cb)

            report_path = os.path.join(output_dir_path, 'Raport neconcordante.xlsx')
            generate_report(results, report_path)

            results_store[sid] = {
                'results': results,
                'report_path': report_path,
                'total_ok': sum(1 for r in results if not r['issues']),
                'total_issues': sum(1 for r in results if r['issues']),
                'total_neconcordante': sum(len(r['issues']) for r in results),
            }
            progress_store[sid] = {'current': len(results), 'total': len(results), 'file': 'Finalizat', 'done': True}
        except Exception as e:
            progress_store[sid] = {'current': 0, 'total': 0, 'file': str(e), 'done': True, 'error': str(e)}

    thread = threading.Thread(target=background_process)
    thread.start()

    return jsonify({'ok': True, 'message': 'Processing started'})


@app.route('/progress')
def progress():
    sid = session.get('session_id', '')
    p = progress_store.get(sid, {'current': 0, 'total': 0, 'file': '', 'done': False})
    return jsonify(p)


@app.route('/results')
def results():
    sid = session.get('session_id', '')
    data = results_store.get(sid)
    if not data:
        flash('Nu există rezultate.', 'error')
        return redirect(url_for('index'))

    # Store report path in session for download
    session['report_path'] = data['report_path']

    return render_template('results.html',
        results=data['results'],
        total_ok=data['total_ok'],
        total_issues=data['total_issues'],
        total_neconcordante=data['total_neconcordante'])


@app.route('/download-report')
def download_report():
    if 'report_path' not in session:
        flash('Nu există raport.', 'error')
        return redirect(url_for('index'))
    path = session['report_path']
    return send_from_directory(os.path.dirname(path), os.path.basename(path), as_attachment=True)


if __name__ == '__main__':
    print('\n' + '=' * 50)
    print('  Verificare H si PV vs MASTER')
    print('=' * 50)
    print(f'  Local:  http://localhost:5060')
    print('=' * 50 + '\n')
    app.run(host='0.0.0.0', port=5060, debug=False)
