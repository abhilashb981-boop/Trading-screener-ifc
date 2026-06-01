"""
Trading Screener Web App — with SSE real-time logs
"""

from flask import Flask, request, jsonify, render_template, send_file, Response, stream_with_context
import pandas as pd
import numpy as np
import io, os, json, tempfile, time, threading, uuid
import warnings
warnings.filterwarnings('ignore')

from screener import analyze_stock, create_excel_report

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024

# In-memory job store  { job_id: { 'events': [], 'done': False, 'result': None } }
jobs = {}
jobs_lock = threading.Lock()


def new_job():
    jid = str(uuid.uuid4())
    with jobs_lock:
        jobs[jid] = {'events': [], 'done': False, 'result': None, 'error': None}
    return jid


def push(jid, msg, kind='log'):
    """Append an SSE event to the job queue."""
    with jobs_lock:
        jobs[jid]['events'].append({'kind': kind, 'msg': msg})


def run_scan(jid, symbols, exchange, period, min_score):
    """Background thread — scans stocks and pushes SSE events."""
    try:
        suffix = {'NSE': '.NS', 'BSE': '.BO', 'US': ''}[exchange]
        total = len(symbols)
        sell_stocks, hold_stocks, errors = [], [], {}

        push(jid, f'📋 {total} stocks found — starting scan…', 'info')

        for i, sym in enumerate(symbols, 1):
            ticker_sym = sym + suffix if not sym.endswith(suffix) else sym
            push(jid, f'[{i}/{total}]  Fetching  {sym}…', 'log')

            result, err = analyze_stock(ticker_sym, period)

            if result:
                result['symbol_clean'] = sym
                score = result['score']
                bar = '█' * int(score) + '░' * (10 - int(score))
                signal = result['signal']

                if score >= 7:
                    icon = '🔴'
                elif score >= 5:
                    icon = '🟠'
                elif score >= 3.5:
                    icon = '🟡'
                else:
                    icon = '⚪'

                push(jid, f'[{i}/{total}]  {icon} {sym:<14} Score: {score:.1f}/10  [{bar}]  {signal}', 'result')

                if score >= min_score:
                    sell_stocks.append(result)
                else:
                    hold_stocks.append(result)
            else:
                push(jid, f'[{i}/{total}]  ⚠️  {sym:<14} Error: {str(err)[:60]}', 'error')
                errors[sym] = err

            # progress event
            push(jid, json.dumps({'current': i, 'total': total}), 'progress')

        sell_stocks.sort(key=lambda x: x['score'], reverse=True)
        hold_stocks.sort(key=lambda x: x['score'], reverse=True)

        push(jid, '─' * 52, 'log')
        push(jid, f'✅ Scan complete — 🔴 {len(sell_stocks)} sell  ⚪ {len(hold_stocks)} hold  ⚠️ {len(errors)} errors', 'info')

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
                'sell':        [clean(r) for r in sell_stocks],
                'hold':        [clean(r) for r in hold_stocks],
                'errors':      errors,
            }
            jobs[jid]['done'] = True

    except Exception as e:
        push(jid, f'❌ Fatal error: {e}', 'error')
        with jobs_lock:
            jobs[jid]['error'] = str(e)
            jobs[jid]['done'] = True


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/start_scan', methods=['POST'])
def start_scan():
    """Parse Excel, create job, start background thread, return job_id."""
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
        t = threading.Thread(target=run_scan, args=(jid, symbols, exchange, period, min_score), daemon=True)
        t.start()

        return jsonify({'job_id': jid})

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/stream/<job_id>')
def stream(job_id):
    """SSE endpoint — streams events for a job."""
    def generate():
        sent = 0
        while True:
            with jobs_lock:
                job = jobs.get(job_id)
                if not job:
                    yield 'event: error\ndata: Job not found\n\n'
                    return

                events = job['events'][sent:]
                done   = job['done']
                result = job['result']

            for ev in events:
                yield f"event: {ev['kind']}\ndata: {ev['msg']}\n\n"
                sent += 1

            if done:
                if result:
                    yield f"event: done\ndata: {json.dumps(result)}\n\n"
                else:
                    yield f"event: error\ndata: {job.get('error','Unknown error')}\n\n"
                # cleanup after 5 min
                def cleanup():
                    time.sleep(300)
                    with jobs_lock:
                        jobs.pop(job_id, None)
                threading.Thread(target=cleanup, daemon=True).start()
                return

            time.sleep(0.3)

    return Response(stream_with_context(generate()),
                    mimetype='text/event-stream',
                    headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'})


@app.route('/download', methods=['POST'])
def download():
    try:
        data = request.get_json()
        sell = data.get('sell', [])
        hold = data.get('hold', [])

        def prep(r):
            rr = r.get('rr_ratio', '0')
            try:
                rr_float = float(str(rr).replace('x', ''))
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
