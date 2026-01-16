import pandas as pd

CSV_URL = "https://docs.google.com/spreadsheets/d/1xINj1dg33ynLut2f4hg366qAk2mwO9yr/export?format=csv&gid=1784284167"
TETO_CLT = 176

def safe_nunique(series: pd.Series) -> int:
    if series is None:
        return 0
    s = series.dropna().astype(str).str.strip()
    s = s[s != ""]
    return int(s.nunique())

def safe_sum(series: pd.Series) -> float:
    if series is None:
        return 0.0
    return float(series.fillna(0).sum())

def main():
    df = pd.read_csv(CSV_URL)

    # filtro ORIGEM
    if "ORIGEM" in df.columns:
        df["ORIGEM"] = df["ORIGEM"].astype(str).str.strip()
        df = df[df["ORIGEM"].isin(["CAMIL", "SPOT"])]

    # padronizar colunas principais (sem renomear; só limpando espaços)
    for col in ["PROMOTOR", "SUPERVISOR FINAL", "NOME FANTASIA", "REGIONAL VOLUME", "CIDADE", "BANDEIRA"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # garantir numéricos
    for col in ["FREQ.SEMANA", "TEMPO DE ATENDIMENTO POR VISITA"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # =========================
    # CÁLCULOS (Horas / Teto)
    # =========================
    df["HORAS_SEMANA_LOJA"] = df.get("FREQ.SEMANA", 0) * df.get("TEMPO DE ATENDIMENTO POR VISITA", 0)

    # horas por promotor
    horas_semana_promotor = (
        df.groupby("PROMOTOR", dropna=False)["HORAS_SEMANA_LOJA"]
        .sum()
        .reset_index()
        .rename(columns={"HORAS_SEMANA_LOJA": "HORAS_SEMANA_PROMOTOR"})
    )
    horas_semana_promotor["HORAS_MES_PREVISTA"] = horas_semana_promotor["HORAS_SEMANA_PROMOTOR"] * 4
    horas_semana_promotor["SALDO_HORAS"] = TETO_CLT - horas_semana_promotor["HORAS_MES_PREVISTA"]
    horas_semana_promotor["ULTRAPASSOU_TETO"] = horas_semana_promotor["HORAS_MES_PREVISTA"] > TETO_CLT
    horas_semana_promotor["HORAS_EXCEDENTES"] = (horas_semana_promotor["HORAS_MES_PREVISTA"] - TETO_CLT).clip(lower=0)

    # =========================
    # KPIs
    # =========================
    total_linhas = len(df)
    qtd_promotores = safe_nunique(df["PROMOTOR"]) if "PROMOTOR" in df.columns else 0
    qtd_supervisores = safe_nunique(df["SUPERVISOR FINAL"]) if "SUPERVISOR FINAL" in df.columns else 0
    qtd_lojas = safe_nunique(df["NOME FANTASIA"]) if "NOME FANTASIA" in df.columns else 0
    qtd_bandeiras = safe_nunique(df["BANDEIRA"]) if "BANDEIRA" in df.columns else 0
    qtd_cidades = safe_nunique(df["CIDADE"]) if "CIDADE" in df.columns else 0

    total_horas_mes_prevista = safe_sum(horas_semana_promotor["HORAS_MES_PREVISTA"])
    qtd_promotores_acima_teto = int(horas_semana_promotor["ULTRAPASSOU_TETO"].sum())
    total_horas_excedentes = safe_sum(horas_semana_promotor["HORAS_EXCEDENTES"])

    # =========================
    # MÉDIAS (gestão)
    # =========================
    # promotores por supervisor
    if "SUPERVISOR FINAL" in df.columns and "PROMOTOR" in df.columns:
        prom_por_sup = (
            df.dropna(subset=["SUPERVISOR FINAL", "PROMOTOR"])
              .groupby("SUPERVISOR FINAL")["PROMOTOR"]
              .nunique()
        )
        media_prom_por_sup = float(prom_por_sup.mean()) if len(prom_por_sup) else 0.0
    else:
        media_prom_por_sup = 0.0

    # lojas por supervisor
    if "SUPERVISOR FINAL" in df.columns and "NOME FANTASIA" in df.columns:
        lojas_por_sup = (
            df.dropna(subset=["SUPERVISOR FINAL", "NOME FANTASIA"])
              .groupby("SUPERVISOR FINAL")["NOME FANTASIA"]
              .nunique()
        )
        media_lojas_por_sup = float(lojas_por_sup.mean()) if len(lojas_por_sup) else 0.0
    else:
        media_lojas_por_sup = 0.0

    # lojas por promotor
    if "PROMOTOR" in df.columns and "NOME FANTASIA" in df.columns:
        lojas_por_prom = (
            df.dropna(subset=["PROMOTOR", "NOME FANTASIA"])
              .groupby("PROMOTOR")["NOME FANTASIA"]
              .nunique()
        )
        media_lojas_por_prom = float(lojas_por_prom.mean()) if len(lojas_por_prom) else 0.0
    else:
        media_lojas_por_prom = 0.0

    # =========================
    # TABELA RESUMO PROMOTORES (top 20 por horas)
    # =========================
    top_promotores = horas_semana_promotor.sort_values("HORAS_MES_PREVISTA", ascending=False).head(20)

    def fmt_num(x: float) -> str:
        return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    cards_html = f"""
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px;margin:16px 0;">
      <div style="padding:12px;border:1px solid #ddd;border-radius:10px;"><b>Linhas válidas (CAMIL/SPOT)</b><div style="font-size:26px;">{total_linhas}</div></div>
      <div style="padding:12px;border:1px solid #ddd;border-radius:10px;"><b>Promotores</b><div style="font-size:26px;">{qtd_promotores}</div></div>
      <div style="padding:12px;border:1px solid #ddd;border-radius:10px;"><b>Supervisores</b><div style="font-size:26px;">{qtd_supervisores}</div></div>
      <div style="padding:12px;border:1px solid #ddd;border-radius:10px;"><b>Lojas (NOME FANTASIA)</b><div style="font-size:26px;">{qtd_lojas}</div></div>
      <div style="padding:12px;border:1px solid #ddd;border-radius:10px;"><b>Bandeiras</b><div style="font-size:26px;">{qtd_bandeiras}</div></div>
      <div style="padding:12px;border:1px solid #ddd;border-radius:10px;"><b>Cidades</b><div style="font-size:26px;">{qtd_cidades}</div></div>
      <div style="padding:12px;border:1px solid #ddd;border-radius:10px;"><b>Total horas mês (previsto)</b><div style="font-size:26px;">{fmt_num(total_horas_mes_prevista)}</div></div>
      <div style="padding:12px;border:1px solid #ddd;border-radius:10px;"><b>Promotores &gt; {TETO_CLT}h</b><div style="font-size:26px;">{qtd_promotores_acima_teto}</div></div>
      <div style="padding:12px;border:1px solid #ddd;border-radius:10px;"><b>Horas excedentes (total)</b><div style="font-size:26px;">{fmt_num(total_horas_excedentes)}</div></div>
      <div style="padding:12px;border:1px solid #ddd;border-radius:10px;"><b>Média promotores / supervisor</b><div style="font-size:26px;">{fmt_num(media_prom_por_sup)}</div></div>
      <div style="padding:12px;border:1px solid #ddd;border-radius:10px;"><b>Média lojas / supervisor</b><div style="font-size:26px;">{fmt_num(media_lojas_por_sup)}</div></div>
      <div style="padding:12px;border:1px solid #ddd;border-radius:10px;"><b>Média lojas / promotor</b><div style="font-size:26px;">{fmt_num(media_lojas_por_prom)}</div></div>
    </div>
    """

    # tabela top promotores
    rows = []
    for _, r in top_promotores.iterrows():
        prom = str(r.get("PROMOTOR", "")).strip()
        hm = float(r.get("HORAS_MES_PREVISTA", 0))
        hs = float(r.get("HORAS_SEMANA_PROMOTOR", 0))
        saldo = float(r.get("SALDO_HORAS", 0))
        exced = float(r.get("HORAS_EXCEDENTES", 0))
        flag = "SIM" if bool(r.get("ULTRAPASSOU_TETO", False)) else "NÃO"
        rows.append(f"""
          <tr>
            <td>{prom}</td>
            <td style="text-align:right;">{fmt_num(hs)}</td>
            <td style="text-align:right;">{fmt_num(hm)}</td>
            <td style="text-align:right;">{fmt_num(saldo)}</td>
            <td style="text-align:center;">{flag}</td>
            <td style="text-align:right;">{fmt_num(exced)}</td>
          </tr>
        """)

    tabela_html = f"""
    <h2 style="margin-top:22px;">Top 20 promotores por horas mês (previsto)</h2>
    <div style="overflow:auto;border:1px solid #eee;border-radius:10px;">
      <table style="width:100%;border-collapse:collapse;">
        <thead>
          <tr style="background:#f7f7f7;">
            <th style="text-align:left;padding:10px;border-bottom:1px solid #eee;">Promotor</th>
            <th style="text-align:right;padding:10px;border-bottom:1px solid #eee;">Horas semana</th>
            <th style="text-align:right;padding:10px;border-bottom:1px solid #eee;">Horas mês (x4)</th>
            <th style="text-align:right;padding:10px;border-bottom:1px solid #eee;">Saldo até 176</th>
            <th style="text-align:center;padding:10px;border-bottom:1px solid #eee;">Ultrapassou?</th>
            <th style="text-align:right;padding:10px;border-bottom:1px solid #eee;">Excedentes</th>
          </tr>
        </thead>
        <tbody>
          {''.join(rows)}
        </tbody>
      </table>
    </div>
    """

    html = f"""
<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>Dashboard Atendimento - Janeiro</title>
</head>
<body style="font-family: Arial; padding: 20px; max-width: 1200px; margin: 0 auto;">
  <h1>Dashboard Atendimento - Janeiro</h1>
  <p style="color:#555;">
    Fonte: Google Sheets (aba JANEIRO_2026) | Atualização automática via GitHub Actions (a cada 1h)
  </p>

  {cards_html}

  {tabela_html}

  <details style="margin-top:20px;">
    <summary style="cursor:pointer;font-weight:bold;">Ver colunas encontradas (debug)</summary>
    <ul>
      {''.join([f"<li>{c}</li>" for c in list(df.columns)])}
    </ul>
  </details>
</body>
</html>
"""

    with open("docs/index.html", "w", encoding="utf-8") as f:
        f.write(html)

if __name__ == "__main__":
    main()
