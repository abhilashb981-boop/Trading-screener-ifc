"""
Trading Screener Web App
Flask backend — wraps the existing screener logic
"""

from flask import Flask, request, jsonify, render_template, send_file
import pandas as pd
import numpy as np
import io
import os
import warnings
warnings.filterwarnings('ignore')

# Import all indicator functions from existing screener
from screener import (
    calc_rsi, calc_macd, calc_bollinger, calc_supertrend, calc_atr,
    calc_stochastic, calc_adx, detect_bearish_candle, detect_cisd,
    detect_institutional_funding_candle, analyze_stock, create_excel_report
)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10MB max upload


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/scan', methods=['POST'])
def scan():
    """Main scan endpoint — accepts Excel file + params, returns JSON results"""
    try:
        # Get params
        exchange = request.form.get('exchange', 'NSE')
        period   = request.form.get('period', '3mo')
        min_score = float(request.form.get('min_score', 3.5))

        # Read symbols from uploaded Excel
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400

        f = request.files['file']
        if f.filename == '':
            return jsonify({'error': 'Empty filename'}), 400

        df_input = pd.read_excel(io.BytesIO(f.read()))

        # Find symbol column
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

        suffix = {'NSE': '.NS', 'BSE': '.BO', 'US': ''}[exchange]

        sell_stocks = []
        hold_stocks = []
        errors = {}
        results_all = []

        for sym in symbols:
            ticker_sym = sym + suffix if not sym.endswith(suffix) else sym
            result, err = analyze_stock(ticker_sym, period)

            if result:
                result['symbol_clean'] = sym  # without suffix for display
                if result['score'] >= min_score:
                    sell_stocks.append(result)
                else:
                    hold_stocks.append(result)
                results_all.append(result)
            else:
                errors[sym] = err

        sell_stocks.sort(key=lambda x: x['score'], reverse=True)
        hold_stocks.sort(key=lambda x: x['score'], reverse=True)

        # Serialize for JSON (numpy floats → python floats)
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

        return jsonify({
            'total':      len(symbols),
            'sell_count': len(sell_stocks),
            'hold_count': len(hold_stocks),
            'error_count':len(errors),
            'sell':  [clean(r) for r in sell_stocks],
            'hold':  [clean(r) for r in hold_stocks],
            'errors': errors,
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/download', methods=['POST'])
def download():
    """Generate and return the Excel report"""
    try:
        import json, tempfile
        data = request.get_json()
        sell = data.get('sell', [])
        hold = data.get('hold', [])

        # Re-add numpy-style fields expected by create_excel_report
        def prep(r):
            rr = r.get('rr_ratio', '0')
            try:
                rr_float = float(str(rr).replace('x',''))
            except:
                rr_float = 0.0
            return {**r, 'rr_ratio': rr_float, 'signals': r.get('signals', {})}

        tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        tmp.close()
        create_excel_report([prep(r) for r in sell], [prep(r) for r in hold], tmp.name)

        return send_file(
            tmp.name,
            as_attachment=True,
            download_name='sell_signals.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
