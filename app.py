"""
Trading Screener Web App — Polling based (SSE timeout fix)
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

# jobs store: { job_id: { logs:[], progress:{}, done:bool, result:None, error:None } }
jobs = {}
jobs_lock = threading.Lock()


def new_job():
    jid = str(uuid.uuid4())
    with jobs_lock:
        jobs[jid] = {
            'logs': [],
            'progress': {'current': 0, 'total': 0},
            'done': False,
            'result': None,
            'error': None,
            'last_sent': 0   # index of last log already fetched by client
        }
    return jid


def push_log(jid, msg, kind='log'):
    with jobs_lock:
        jobs[jid]['logs'].append({'kind': kind, 'msg': msg})


def run_scan(jid, symbols, exchange, period, min_score):
    try:
        suffix = {'NSE': '.NS', 'BSE': '.BO', 'US': ''}[exchange]
        total = len(symbols)

        with jobs_lock:
            jobs[jid]['progress'] = {'current': 0, 'total': total}

        push_log(jid, f'📋 {total} stocks found — starting scan…', 'info')

        sell_stocks, hold_stocks, errors = [], [], {}

        for i, sym in enumerate(symbols, 1):
            ticker_sym = sym + suffix if not sym.endswith(suffix) else sym
            push_log(jid, f'[{i}/{total}]  Fetching  {sym}…', 'log')

            result, err = analyze_stock(ticker_sym, period)

            if result:
                result['symbol_clean'] = sym
                score = result['score']
                bar = '█' * int(score) + '░' * (10 - int(score))
                icon = '🔴' if score >= 7 else '🟠' if score >= 5 else '🟡' if score >= 3.5 else '⚪'
                push_log(jid, f'[{i}/{total}]  {icon} {sym:<14} Score: {score:.1f}/10  [{bar}]  {result["signal"]}', 'result')
                if score >= min_score:
                    sell_stocks.append(result)
                else:
                    hold_stocks.append(result)
            else:
                push_log(jid, f'[{i}/{total}]  ⚠️  {sym:<14} Error: {str(err)[:60]}', 'error')
                errors[sym] = str(err)

            with jobs_lock:
                jobs[jid]['progress'] = {'current': i, 'total': total}

        sell_stocks.sort(key=lambda x: x['score'], reverse=True)
        hold_stocks.sort(key=lambda x: x['score'], reverse=True)

        push_log(jid, '─' * 52, 'log')
        push_log(jid, f'✅ Done — 🔴 {len(sell_stocks)} sell  ⚪ {len(hold_stocks)} hold  ⚠️ {len(errors)} errors', 'info')

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

        with jobs_lock:
            jobs[jid]['result'] = {
                'total':       total,
                'sell_count':  len(sell_stocks),
                'hold_count':  len(hold_stocks),
                'error_count': len(errors),
                'sell':  [clean(r) for r in sell_stocks],
                'hold':  [clean(r) for r in hold_stocks],
                'errors': errors,
            }
            jobs[jid]['done'] = True

        # cleanup after 10 min
        def cleanup():
            time.sleep(600)
            with jobs_lock:
                jobs.pop(jid, None)
        threading.Thread(target=cleanup, daemon=True).start()

    except Exception as e:
        push_log(jid, f'❌ Fatal: {e}', 'error')
        with jobs_lock:
            jobs[jid]['error'] = str(e)
            jobs[jid]['done'] = True


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

        f = request.files['file']
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

        jid = new_job()
        t = threading.Thread(target=run_scan,
                             args=(jid, symbols, exchange, period, min_score),
                             daemon=True)
        t.start()
        return jsonify({'job_id': jid, 'total': len(symbols)})

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/status/<job_id>')
def status(job_id):
    """Polling endpoint — returns new logs since last_seen index + progress + done flag."""
    try:
        last_seen = int(request.args.get('last_seen', 0))
        with jobs_lock:
            job = jobs.get(job_id)
            if not job:
                return jsonify({'error': 'Job not found'}), 404

            new_logs  = job['logs'][last_seen:]
            progress  = job['progress'].copy()
            done      = job['done']
            result    = job['result'] if done else None
            error     = job['error']
            total_logs = len(job['logs'])

        return jsonify({
            'logs':       new_logs,
            'progress':   progress,
            'done':       done,
            'result':     result,
            'error':      error,
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
            try:
                rr_float = float(str(r.get('rr_ratio', '0')).replace('x', ''))
            except:
                rr_float = 0.0
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
