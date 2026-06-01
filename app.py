"""
Trading Screener Web App — File-based job store (fixes worker restart issues)
"""

from flask import Flask, request, jsonify, render_template, send_file
import pandas as pd
import numpy as np
import io, os, json, tempfile, time, threading, uuid
import warnings
warnings.filterwarnings('ignore')

from screener import analyze_stock, create_excel_report

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

JOBS_DIR = '/tmp/screener_jobs'
os.makedirs(JOBS_DIR, exist_ok=True)


def job_path(jid):
    return os.path.join(JOBS_DIR, f'{jid}.json')


def read_job(jid):
    try:
        with open(job_path(jid), 'r') as f:
            return json.load(f)
    except:
        return None


def write_job(jid, data):
    tmp = job_path(jid) + '.tmp'
    with open(tmp, 'w') as f:
        json.dump(data, f)
    os.replace(tmp, job_path(jid))


def new_job(jid):
    write_job(jid, {
        'logs': [],
        'progress': {'current': 0, 'total': 0},
        'done': False,
        'result': None,
        'error': None,
    })


def push_log(jid, msg, kind='log'):
    """Append log line — read/modify/write with file lock."""
    lock_path = job_path(jid) + '.lock'
    import fcntl
    with open(lock_path, 'w') as lf:
        fcntl.flock(lf, fcntl.LOCK_EX)
        data = read_job(jid) or {}
        data.setdefault('logs', []).append({'kind': kind, 'msg': msg})
        write_job(jid, data)


def update_progress(jid, current, total):
    lock_path = job_path(jid) + '.lock'
    import fcntl
    with open(lock_path, 'w') as lf:
        fcntl.flock(lf, fcntl.LOCK_EX)
        data = read_job(jid) or {}
        data['progress'] = {'current': current, 'total': total}
        write_job(jid, data)


def finish_job(jid, result=None, error=None):
    lock_path = job_path(jid) + '.lock'
    import fcntl
    with open(lock_path, 'w') as lf:
        fcntl.flock(lf, fcntl.LOCK_EX)
        data = read_job(jid) or {}
        data['done']   = True
        data['result'] = result
        data['error']  = error
        write_job(jid, data)


def run_scan(jid, symbols, exchange, period, min_score):
    try:
        suffix = {'NSE': '.NS', 'BSE': '.BO', 'US': ''}[exchange]
        total  = len(symbols)

        update_progress(jid, 0, total)
        push_log(jid, f'📋 {total} stocks found — starting scan…', 'info')

        sell_stocks, hold_stocks, errors = [], [], {}

        for i, sym in enumerate(symbols, 1):
            ticker_sym = sym + suffix if not sym.endswith(suffix) else sym
            push_log(jid, f'[{i}/{total}]  Fetching  {sym}…', 'log')

            result, err = analyze_stock(ticker_sym, period)

            if result:
                result['symbol_clean'] = sym
                score = result['score']
                bar   = '█' * int(score) + '░' * (10 - int(score))
                icon  = '🔴' if score >= 7 else '🟠' if score >= 5 else '🟡' if score >= 3.5 else '⚪'
                push_log(jid,
                    f'[{i}/{total}]  {icon} {sym:<14} Score: {score:.1f}/10  [{bar}]  {result["signal"]}',
                    'result')
                if score >= min_score:
                    sell_stocks.append(result)
                else:
                    hold_stocks.append(result)
            else:
                push_log(jid, f'[{i}/{total}]  ⚠️  {sym:<14} Error: {str(err)[:60]}', 'error')
                errors[sym] = str(err)

            update_progress(jid, i, total)

        sell_stocks.sort(key=lambda x: x['score'], reverse=True)
        hold_stocks.sort(key=lambda x: x['score'], reverse=True)

        push_log(jid, '─' * 52, 'log')
        push_log(jid,
            f'✅ Done — 🔴 {len(sell_stocks)} sell  ⚪ {len(hold_stocks)} hold  ⚠️ {len(errors)} errors',
            'info')

        def clean(r):
            return {
                'symbol':    r.get('symbol_clean', r['symbol']),
                'signal':    r['signal'],
                'score':     round(float(r['score']), 1),
                'entry':     round(float(r['entry']), 2),
                'stop_loss': round(float(r['stop_loss']), 2),
                'target1':   round(float(r['target1']), 2),
                'target2':   round(float(r['target2']), 2),
                'rr_ratio':  str(r['rr_ratio']),
                'signals':   r.get('signals', {}),
            }

        finish_job(jid, result={
            'total':       total,
            'sell_count':  len(sell_stocks),
            'hold_count':  len(hold_stocks),
            'error_count': len(errors),
            'sell':        [clean(r) for r in sell_stocks],
            'hold':        [clean(r) for r in hold_stocks],
            'errors':      errors,
        })

        # cleanup after 15 min
        def cleanup():
            time.sleep(900)
            for ext in ('', '.tmp', '.lock'):
                try: os.remove(job_path(jid) + ext)
                except: pass
        threading.Thread(target=cleanup, daemon=True).start()

    except Exception as e:
        push_log(jid, f'❌ Fatal: {e}', 'error')
        finish_job(jid, error=str(e))


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/start_scan', methods=['POST'])
def start_scan():
    try:
        exchange  = request.form.get('exchange', 'NSE')
        period    = request.form.get('period', '3mo')
        min_score = float(request.form.get('min_score', 3.5))

        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400

        f        = request.files['file']
        df_input = pd.read_excel(io.BytesIO(f.read()))

        symbols = []
        for col in ['Symbol', 'symbol', 'Stock', 'Ticker', 'SYMBOL', 'Name']:
            if col in df_input.columns:
                symbols = df_input[col].dropna().astype(str).str.strip().str.upper().tolist()
                symbols = [s for s in symbols if s and s != 'NAN']
                break
        if not symbols:
            symbols = df_input.iloc[:, 0].dropna().astype(str).str.strip().str.upper().tolist()
            symbols = [s for s in symbols if s and s != 'NAN']

        if not symbols:
            return jsonify({'error': 'No symbols found in Excel'}), 400

        jid = str(uuid.uuid4())
        new_job(jid)

        t = threading.Thread(target=run_scan,
                             args=(jid, symbols, exchange, period, min_score),
                             daemon=True)
        t.start()
        return jsonify({'job_id': jid, 'total': len(symbols)})

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/status/<job_id>')
def status(job_id):
    try:
        last_seen = int(request.args.get('last_seen', 0))
        data = read_job(job_id)
        if not data:
            return jsonify({'error': 'Job not found'}), 404

        new_logs   = data['logs'][last_seen:]
        total_logs = len(data['logs'])

        return jsonify({
            'logs':       new_logs,
            'progress':   data.get('progress', {}),
            'done':       data.get('done', False),
            'result':     data.get('result'),
            'error':      data.get('error'),
            'total_logs': total_logs,
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/download', methods=['POST'])
def download():
    try:
        data = request.get_json()
        sell = data.get('sell', [])
        hold = data.get('hold', [])

        def prep(r):
            try:    rr_float = float(str(r.get('rr_ratio', '0')).replace('x', ''))
            except: rr_float = 0.0
            return {**r, 'rr_ratio': rr_float, 'signals': r.get('signals', {})}

        tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        tmp.close()
        create_excel_report([prep(r) for r in sell], [prep(r) for r in hold], tmp.name)

        return send_file(tmp.name, as_attachment=True, download_name='sell_signals.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False, threaded=True)
