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
    # If numeric, avoid decimal point creating extra zeros (e.g., 125.0 -> '1250')
    try:
        # handle pandas/numpy numeric types too
        import numbers
        if isinstance(s, numbers.Integral):
            return str(s)
        if isinstance(s, numbers.Real):
            # if it's an integer value like 125.0, convert to int
            if float(s).is_integer():
                return str(int(s))
    except Exception:
        pass
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
    # Accept both accented and unaccented forms for 'codigo'
    mat_col = find_column(df.columns, ['matric', 'matrícula', 'matricula', 'codigo', 'código', 'cod'])
    # Prefer explicit header 'valor refeição' (case-insensitive), but accept common alternatives
    val_col = find_column(df.columns, ['valor refeição', 'valor', 'value', 'refei'])

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

    # We want one line per matrícula: sum all values per matricula and use the last date
    agg = df.groupby('mat_digits', as_index=False).agg({val_col: 'sum', date_col: 'max'})

    # Format lines
    lines = []
    for _, row in agg.iterrows():
        # date_col contains a Timestamp (max per matricula)
        date_part = row[date_col].strftime('%Y%m%d') + '000'
        mat = str(row['mat_digits']).zfill(10)
        cents = int(round(row[val_col] * 100))
        val_str = str(cents).zfill(7)
        lines.append(f"{date_part} {mat} {val_str}")

    txt = "\n".join(lines)

    buf = io.BytesIO()
    buf.write(txt.encode('utf-8'))
    buf.seek(0)

    return send_file(buf, as_attachment=True, download_name='output.txt', mimetype='text/plain')


@app.route('/convert_pj', methods=['POST'])
def convert_pj():
    file = request.files.get('file')
    if not file:
        return redirect(url_for('index'))

    try:
        df = pd.read_excel(file, engine='openpyxl')
    except Exception as e:
        return f"Erro ao ler o arquivo Excel: {e}", 400

    # Find columns
    mat_col = find_column(df.columns, ['matric', 'matrícula', 'matricula', 'codigo', 'código', 'cod'])
    val_col = find_column(df.columns, ['valor refeição', 'valor', 'value', 'refei'])
    name_col = find_column(df.columns, ['nome', 'name', 'funcionario', 'funcionário'])

    if not mat_col or not val_col:
        return (
            "Colunas não encontradas. O arquivo Excel precisa ter colunas 'codigo' e 'valor refeição'.",
            400,
        )

    # Clean data
    df = df.dropna(subset=[mat_col, val_col])

    # Prepare matricula
    df['mat_digits'] = df[mat_col].apply(only_digits)
    df[val_col] = pd.to_numeric(df[val_col], errors='coerce').fillna(0)

    # Group by matricula and sum values, keep the first name
    agg_dict = {val_col: 'sum'}
    if name_col:
        agg_dict[name_col] = 'first'
    
    agg = df.groupby('mat_digits', as_index=False).agg(agg_dict)
    
    # Rename columns
    if name_col:
        agg.columns = ['Código', 'Valor Total Refeição', 'Nome']
        agg = agg[['Código', 'Nome', 'Valor Total Refeição']]  # Reorder columns
    else:
        agg.columns = ['Código', 'Valor Total Refeição']

    # Create Excel output
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        agg.to_excel(writer, index=False, sheet_name='Totalizacao')
    output.seek(0)

    return send_file(output, as_attachment=True, download_name='almoco_pj_totalizado.xlsx', 
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5002, debug=True)
