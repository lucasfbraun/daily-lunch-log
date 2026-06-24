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

    # Join with CRLF (Windows line ending) and encode as UTF-8
    txt = "\r\n".join(lines) + "\r\n"

    buf = io.BytesIO()
    # Keep UTF-8 encoding but use CRLF line endings
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


@app.route('/convert_pdf', methods=['POST'])
def convert_pdf():
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from datetime import datetime

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
            "Colunas não encontradas. O arquivo Excel precisa ter colunas 'codigo' e 'valor'.",
            400,
        )

    # Clean data
    df = df.dropna(subset=[mat_col, val_col])
    df['mat_digits'] = df[mat_col].apply(only_digits)
    df[val_col] = pd.to_numeric(df[val_col], errors='coerce').fillna(0)

    # Group by matricula
    agg_dict = {val_col: 'sum'}
    if name_col:
        agg_dict[name_col] = 'first'

    agg = df.groupby('mat_digits', as_index=False).agg(agg_dict)
    if name_col:
        agg = agg.sort_values(name_col)

    def fmt_brl(v):
        """Format value as Brazilian currency: R$ 1.234,56"""
        s = f"{v:,.2f}"
        s = s.replace(',', 'X').replace('.', ',').replace('X', '.')
        return f"R$ {s}"

    # Build PDF
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        rightMargin=2 * cm, leftMargin=2 * cm,
        topMargin=2 * cm, bottomMargin=2 * cm,
        title='Relatório de Almoços por Funcionário',
        author='Grupo Flexível',
    )

    CIANO = colors.HexColor('#0097A7')
    CIANO_DARK = colors.HexColor('#007c8a')
    DARK = colors.HexColor('#212121')
    GRAY = colors.HexColor('#595959')
    LIGHT_GRAY = colors.HexColor('#EEEEEE')
    WHITE = colors.white

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'ReportTitle', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=20,
        textColor=DARK, alignment=TA_CENTER, spaceAfter=4,
    )
    subtitle_style = ParagraphStyle(
        'ReportSubtitle', parent=styles['Normal'],
        fontName='Helvetica', fontSize=12,
        textColor=CIANO, alignment=TA_CENTER, spaceAfter=2,
    )
    date_style = ParagraphStyle(
        'ReportDate', parent=styles['Normal'],
        fontName='Helvetica', fontSize=9,
        textColor=GRAY, alignment=TA_RIGHT,
    )

    elements = []

    # Header
    elements.append(Paragraph('Grupo Flexível', title_style))
    elements.append(Spacer(1, 0.4 * cm))
    elements.append(Paragraph('Relatório de Almoços por Funcionário', subtitle_style))
    elements.append(Spacer(1, 0.2 * cm))
    elements.append(Paragraph(f'Gerado em: {datetime.now().strftime("%d/%m/%Y às %H:%M")}', date_style))
    elements.append(Spacer(1, 0.6 * cm))

    # Table rows
    if name_col:
        headers = ['Nome', 'Código', 'Total (R$)']
        col_widths = [9 * cm, 4 * cm, 4 * cm]
        data_rows = [
            [str(row[name_col]), str(row['mat_digits']), fmt_brl(row[val_col])]
            for _, row in agg.iterrows()
        ]
    else:
        headers = ['Código', 'Total (R$)']
        col_widths = [9 * cm, 8 * cm]
        data_rows = [
            [str(row['mat_digits']), fmt_brl(row[val_col])]
            for _, row in agg.iterrows()
        ]

    total = agg[val_col].sum()
    if name_col:
        total_row = ['TOTAL GERAL', '', fmt_brl(total)]
    else:
        total_row = ['TOTAL GERAL', fmt_brl(total)]

    rows = [headers] + data_rows + [total_row]
    n = len(rows)

    table = Table(rows, colWidths=col_widths, repeatRows=1)

    style_cmds = [
        # Header row
        ('BACKGROUND', (0, 0), (-1, 0), CIANO),
        ('TEXTCOLOR', (0, 0), (-1, 0), WHITE),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        # Body rows
        ('FONTNAME', (0, 1), (-1, n - 2), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, n - 2), 10),
        ('ALIGN', (0, 1), (0, n - 2), 'LEFT'),
        ('ALIGN', (1, 1), (-1, n - 2), 'CENTER'),
        # Total row
        ('BACKGROUND', (0, -1), (-1, -1), CIANO_DARK),
        ('TEXTCOLOR', (0, -1), (-1, -1), WHITE),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, -1), (-1, -1), 11),
        ('ALIGN', (0, -1), (-1, -1), 'CENTER'),
        # Grid and padding
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
        ('TOPPADDING', (0, 0), (-1, -1), 7),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 7),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('ROWBACKGROUND', (0, 0), (-1, 0), CIANO),
    ]

    # Alternating row background (skip header row 0 and total row -1)
    for i in range(2, n - 1, 2):
        style_cmds.append(('BACKGROUND', (0, i), (-1, i), LIGHT_GRAY))

    table.setStyle(TableStyle(style_cmds))
    elements.append(table)

    # Footer spacer + count
    elements.append(Spacer(1, 0.4 * cm))
    count_style = ParagraphStyle(
        'FooterCount', parent=styles['Normal'],
        fontName='Helvetica', fontSize=9,
        textColor=GRAY, alignment=TA_LEFT,
    )
    elements.append(Paragraph(f'Total de funcionários: {len(data_rows)}', count_style))

    doc.build(elements)
    buf.seek(0)

    return send_file(
        buf, as_attachment=True,
        download_name='almoco_totalizado.pdf',
        mimetype='application/pdf',
    )


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001, debug=True)
