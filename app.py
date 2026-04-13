import os, io, json, tempfile, time, pickle, logging
from pathlib import Path
from flask import Flask, request, jsonify, send_file, render_template
from werkzeug.utils import secure_filename
import rekap_rek as rr

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200 MB

# Session dir — di dalam folder app agar persistent di Railway
SESSION_DIR = Path(__file__).parent / 'sessions'
SESSION_DIR.mkdir(exist_ok=True)

def _session_path(session_id):
    safe = ''.join(c for c in session_id if c.isalnum() or c in '_-')[:80]
    return SESSION_DIR / f"{safe}.pkl"

def _save_session(session_id, meta, transactions):
    path = _session_path(session_id)
    try:
        with open(path, 'wb') as f:
            pickle.dump({'meta': meta, 'transactions': transactions, 'ts': time.time()}, f)
        log.info(f"Session saved: {session_id} → {path} ({path.stat().st_size} bytes)")
    except Exception as e:
        log.error(f"Failed to save session: {e}")

def _load_session(session_id):
    path = _session_path(session_id)
    log.info(f"Loading session: {session_id} → {path} exists={path.exists()}")
    if not path.exists():
        # List semua session yang ada untuk debug
        existing = list(SESSION_DIR.glob('*.pkl'))
        log.warning(f"Session not found. Available sessions: {[p.name for p in existing]}")
        return None
    try:
        with open(path, 'rb') as f:
            data = pickle.load(f)
        log.info(f"Session loaded OK: {len(data.get('transactions',[]))} transactions")
        return data
    except Exception as e:
        log.error(f"Failed to load session: {e}")
        return None

def _cleanup_sessions():
    now = time.time()
    for p in SESSION_DIR.glob('*.pkl'):
        try:
            if now - p.stat().st_mtime > 7200:
                p.unlink()
        except Exception:
            pass

def allowed(filename):
    return filename.lower().endswith('.pdf')

def parse_pdfs(files):
    all_transactions = []
    meta_combined = {
        "accountNo": "", "companyName": "", "period": "",
        "opening": 0, "totalDebet": 0, "totalKredit": 0, "closing": 0,
        "currency": "IDR"
    }
    periods = []
    file_results = []

    with tempfile.TemporaryDirectory() as tmpdir:
        for f in files:
            if not allowed(f.filename):
                continue
            fname = secure_filename(f.filename)
            fpath = os.path.join(tmpdir, fname)
            f.save(fpath)
            try:
                meta, txs = rr.parse_pdf(fpath)
                all_transactions.extend(txs)
                file_results.append({
                    'file': fname,
                    'rekening': meta['accountNo'],
                    'nama': meta['companyName'],
                    'periode': meta['period'],
                    'jumlah': len(txs),
                    'penjualan': sum(1 for t in txs if t['kategori'] == 'Penjualan'),
                })
                if not meta_combined['accountNo'] and meta['accountNo']:
                    meta_combined['accountNo']   = meta['accountNo']
                    meta_combined['companyName'] = meta['companyName']
                meta_combined['totalDebet']  += meta.get('totalDebet', 0)
                meta_combined['totalKredit'] += meta.get('totalKredit', 0)
                if meta.get('period'):
                    periods.append(meta['period'])
                if not meta_combined['opening'] and meta.get('opening'):
                    meta_combined['opening'] = meta['opening']
                if meta.get('closing'):
                    meta_combined['closing'] = meta['closing']
                if meta.get('currency', 'IDR') != 'IDR':
                    meta_combined['currency'] = meta['currency']
            except Exception as e:
                import traceback
                log.error(f"Parse error {fname}: {traceback.format_exc()}")
                file_results.append({'file': fname, 'error': str(e)})

    if periods:
        meta_combined['period'] = (
            f"{periods[0].split(' - ')[0]} - {periods[-1].split(' - ')[-1]}"
        )
    return all_transactions, meta_combined, file_results


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/proses', methods=['POST'])
def proses():
    _cleanup_sessions()
    files = request.files.getlist('pdfs')
    if not files or all(f.filename == '' for f in files):
        return jsonify({'error': 'Tidak ada file yang diupload'}), 400

    all_transactions, meta, file_results = parse_pdfs(files)
    if not all_transactions:
        return jsonify({'error': 'Tidak ada transaksi berhasil dibaca'}), 400

    session_id = (meta.get('accountNo') or 'sesi') + '_' + str(int(time.time()))
    _save_session(session_id, meta, all_transactions)

    return jsonify({
        'session_id': session_id,
        'meta': meta,
        'files': file_results,
        'total_tx': len(all_transactions),
        'total_penj': sum(1 for t in all_transactions if t['kategori'] == 'Penjualan'),
        'transactions': [
            {
                'no': i + 1,
                'month': t['month'],
                'date': t['date'],
                'desc': t['desc'],
                'debet': t['debet'],
                'kredit': t['kredit'],
                'balance': t['balance'],
                'kategori': t['kategori'],
                'customer': t.get('customer') or (rr._extract_customer_name(t['desc']) if t['kategori'] == 'Penjualan' else ''),
                'customerAuto': rr._extract_customer_name(t['desc']) if t['kategori'] == 'Penjualan' else '',
            }
            for i, t in enumerate(all_transactions)
        ]
    })


@app.route('/download', methods=['POST'])
def download():
    data       = request.get_json(silent=True) or {}
    session_id = data.get('session_id', '')
    overrides  = data.get('overrides', [])

    log.info(f"Download request: session_id={session_id!r}")

    if not session_id:
        return jsonify({'error': 'session_id kosong. Silakan proses ulang PDF.'}), 400

    session = _load_session(session_id)
    if not session:
        return jsonify({'error': 'Session tidak ditemukan. Silakan proses ulang PDF.'}), 400

    meta = session['meta']
    txs  = [dict(t) for t in session['transactions']]

    # Terapkan override kategori dan customer dari user
    kat_map  = {o['no']: o['kategori'] for o in overrides if 'no' in o and 'kategori' in o}
    cust_map = {o['no']: o['customer'] for o in overrides if 'no' in o and 'customer' in o}
    for i, tx in enumerate(txs):
        no = tx.get('no', i + 1)
        if no in kat_map:
            tx['kategori'] = kat_map[no]
        if no in cust_map:
            tx['customer'] = cust_map[no]

    acc = meta.get('accountNo', 'rekening')
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp_path = tmp.name
    try:
        rr.build_excel(txs, meta, tmp_path)
        buf = io.BytesIO()
        with open(tmp_path, 'rb') as fh:
            buf.write(fh.read())
        buf.seek(0)
    finally:
        try:
            os.unlink(tmp_path)
        except:
            pass

    return send_file(
        buf,
        as_attachment=True,
        download_name=f"rekap_{acc}.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
