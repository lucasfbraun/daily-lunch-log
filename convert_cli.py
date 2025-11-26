"""CLI converter: lê um arquivo Excel e gera `output.txt` no mesmo diretório.

Uso: python convert_cli.py caminho/para/arquivo.xlsx
"""
import sys
import os
import re
import pandas as pd


def find_column(columns, keywords):
    for kw in keywords:
        for c in columns:
            if kw in c.lower():
                return c
    return None


def only_digits(s):
    return re.sub(r"\D", "", str(s))


def convert(path, output_path=None):
    if output_path is None:
        base = os.path.dirname(path)
        output_path = os.path.join(base, 'output.txt')

    df = pd.read_excel(path, engine='openpyxl')

    date_col = find_column(df.columns, ['data', 'date'])
    mat_col = find_column(df.columns, ['matric', 'matrícula', 'matricula', 'codigo', 'cod'])
    val_col = find_column(df.columns, ['valor', 'value', 'refei'])

    if not date_col or not mat_col or not val_col:
        raise ValueError("Colunas necessárias não encontradas. Procure por 'data', 'codigo/matricula' e 'valor'.")

    # parse dates (try dayfirst)
    df[date_col] = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
    if df[date_col].isna().all():
        df[date_col] = pd.to_datetime(df[date_col], dayfirst=False, errors='coerce')

    df = df.dropna(subset=[date_col, mat_col, val_col])
    df['date_str'] = df[date_col].dt.strftime('%Y%m%d')
    df['mat_digits'] = df[mat_col].apply(only_digits)
    df[val_col] = pd.to_numeric(df[val_col], errors='coerce').fillna(0)

    # Sum all meals per matricula and take the latest date for each matricula
    summed = df.groupby('mat_digits', as_index=False)[val_col].sum()
    latest_dates = df.groupby('mat_digits', as_index=False)[date_col].max()
    merged = pd.merge(summed, latest_dates, on='mat_digits', how='left')

    lines = []
    for _, row in merged.iterrows():
        latest_dt = row[date_col]
        date_str_latest = latest_dt.strftime('%Y%m%d') if not pd.isna(latest_dt) else ''
        date_part = f"{date_str_latest}000"
        mat = str(row['mat_digits']).zfill(9)
        cents = int(round(row[val_col] * 100))
        val_str = str(cents).zfill(7)
        lines.append(f"{date_part} {mat} {val_str}")

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))

    return output_path


def main():
    if len(sys.argv) < 2:
        print('Uso: python convert_cli.py arquivo.xlsx [output.txt]')
        sys.exit(1)

    path = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) >= 3 else None

    if not os.path.exists(path):
        print('Arquivo não encontrado:', path)
        sys.exit(1)

    try:
        result = convert(path, out)
        print('Gerado:', result)
    except Exception as e:
        print('Erro:', e)
        sys.exit(2)


if __name__ == '__main__':
    main()
