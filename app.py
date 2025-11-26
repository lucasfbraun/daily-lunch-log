from flask import Flask, request, render_template, send_file, redirect, url_for
import pandas as pd
import io
import re

app = Flask(__name__)


def find_column(columns, keywords):
    cols = list(columns)
    for kw in keywords:
        for c in cols:
            if kw in c.lower():
                return c
    return None


def only_digits(s):
    return re.sub(r"\D", "", str(s))


@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert():
    file = request.files.get('file')
    if not file:
        return redirect(url_for('index'))

    try:
        df = pd.read_excel(file, engine='openpyxl')
    except Exception as e:
        return f"Erro ao ler o arquivo Excel: {e}", 400

    cols = [c.lower() for c in df.columns]
    date_col = find_column(df.columns, ['data', 'date'])
    mat_col = find_column(df.columns, ['matric', 'matrícula', 'matricula', 'codigo', 'cod'])
    val_col = find_column(df.columns, ['valor', 'value', 'refei'])

    if not date_col or not mat_col or not val_col:
        return (
            "Colunas não encontradas. O arquivo Excel precisa ter colunas 'data', 'matricula' e 'valor'.",
            400,
        )

    # Normalize and parse (try day-first parsing for Brazilian dates first)
    try:
        df[date_col] = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
        # if all parsed as NaT, try without dayfirst
        if df[date_col].isna().all():
            df[date_col] = pd.to_datetime(df[date_col], dayfirst=False, errors='coerce')
    except Exception:
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')

    df = df.dropna(subset=[date_col, mat_col, val_col])

    # prepare grouping keys
    df['date_str'] = df[date_col].dt.strftime('%Y%m%d')
    df['mat_digits'] = df[mat_col].apply(only_digits)

    # ensure valor numeric
    df[val_col] = pd.to_numeric(df[val_col], errors='coerce').fillna(0)

    # Sum all meals per matricula and take the latest date for each matricula
    summed = df.groupby('mat_digits', as_index=False)[val_col].sum()
    latest_dates = df.groupby('mat_digits', as_index=False)[date_col].max()
    merged = pd.merge(summed, latest_dates, on='mat_digits', how='left')

    # Format lines (one line per matricula: latest date + total)
    lines = []
    for _, row in merged.iterrows():
        latest_dt = row[date_col]
        date_str_latest = latest_dt.strftime('%Y%m%d') if not pd.isna(latest_dt) else ''
        date_part = f"{date_str_latest}000"
        mat = str(row['mat_digits']).zfill(9)
        cents = int(round(row[val_col] * 100))
        val_str = str(cents).zfill(7)
        lines.append(f"{date_part} {mat} {val_str}")

    txt = "\n".join(lines)

    buf = io.BytesIO()
    buf.write(txt.encode('utf-8'))
    buf.seek(0)

    return send_file(buf, as_attachment=True, download_name='output.txt', mimetype='text/plain')


if __name__ == '__main__':
    app.run(debug=True)
