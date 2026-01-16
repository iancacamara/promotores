import pandas as pd

CSV_URL = "https://docs.google.com/spreadsheets/d/1xINj1dg33ynLut2f4hg366qAk2mwO9yr/export?format=csv&gid=1784284167"

def main():
    df = pd.read_csv(CSV_URL)

    # filtro ORIGEM
    if "ORIGEM" in df.columns:
        df = df[df["ORIGEM"].isin(["CAMIL", "SPOT"])]

    total_linhas = len(df)
    colunas = list(df.columns)

    html = f"""
    <!DOCTYPE html>
    <html lang="pt-br">
    <head>
      <meta charset="UTF-8" />
      <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
      <title>Dashboard Atendimento - Janeiro</title>
    </head>
    <body style="font-family: Arial; padding: 20px;">
      <h1>Dashboard Atendimento - Janeiro</h1>
      <p><b>Linhas válidas após filtro ORIGEM:</b> {total_linhas}</p>
      <h2>Colunas encontradas</h2>
      <ul>
        {''.join([f"<li>{c}</li>" for c in colunas])}
      </ul>
      <p>Próximo passo: montar os KPIs e gráficos.</p>
    </body>
    </html>
    """

    with open("docs/index.html", "w", encoding="utf-8") as f:
        f.write(html)

if __name__ == "__main__":
    main()
