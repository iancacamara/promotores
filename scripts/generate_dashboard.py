import pandas as pd

CSV_URL = "https://docs.google.com/spreadsheets/d/1xINj1dg33ynLut2f4hg366qAk2mwO9yr/export?format=csv&gid=1784284167"
TETO_CLT = 176


def fmt_num(x: float) -> str:
    try:
        x = float(x)
    except Exception:
        x = 0.0
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def safe_nunique(series: pd.Series) -> int:
    if series is None:
        return 0
    s = series.dropna().astype(str).str.strip()
    s = s[s != ""]
    return int(s.nunique())


def safe_sum(series: pd.Series) -> float:
    if series is None:
        return 0.0
    return float(pd.to_numeric(series, errors="coerce").fillna(0).sum())


def pick_col(df: pd.DataFrame, possiveis: list[str]) -> str | None:
    # match direto
    for p in possiveis:
        if p in df.columns:
            return p
    # match "normalizado"
    norm = {str(c).lower().replace("\n", " ").strip().replace("  ", " "): c for c in df.columns}
    for p in possiveis:
        key = p.lower().replace("\n", " ").strip().replace("  ", " ")
        if key in norm:
            return norm[key]
    return None


def write_debug_html(df: pd.DataFrame, msg: str) -> None:
    debug_cols = "<br>".join([str(c) for c in df.columns])
    html = f"""
<!DOCTYPE html>
<html lang="pt-br">
<head><meta charset="UTF-8"><title>Dashboard Atendimento - Janeiro (Debug)</title></head>
<body style="font-family:Arial;padding:20px;max-width:1200px;margin:0 auto;">
  <h1>Dashboard Atendimento - Janeiro</h1>
  <h2 style="color:#b00020;">Erro de configuração</h2>
  <p>{msg}</p>
  <h3>Colunas lidas do Google Sheets (CSV)</h3>
  <div style="padding:12px;border:1px solid #ddd;border-radius:10px;line-height:1.6;">
    {debug_cols}
  </div>
</body>
</html>
"""
    with open("docs/index.html", "w", encoding="utf-8") as f:
        f.write(html)


def main():
    df = pd.read_csv(CSV_URL)

    # Normaliza nomes de colunas
    df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]

    # Mapeia colunas (aceitando variações)
    COL_PROMOTOR = pick_col(df, ["PROMOTOR"])
    COL_SUPERVISOR = pick_col(df, ["SUPERVISOR FINAL"])
    COL_LOJA = pick_col(df, ["NOME FANTASIA"])
    COL_REGIONAL = pick_col(df, ["REGIONAL VOLUME"])
    COL_CIDADE = pick_col(df, ["CIDADE"])
    COL_BANDEIRA = pick_col(df, ["BANDEIRA"])
    COL_ORIGEM = pick_col(df, ["ORIGEM"])
    COL_FREQ = pick_col(df, ["FREQ.SEMANA", "FREQ SEMANA", "FREQ"])
    COL_TEMPO = pick_col(df, ["TEMPO DE ATENDIMENTO POR VISITA", "TEMPO ATENDIMENTO POR VISITA", "TEMPO"])

    # Validações mínimas
    if COL_PROMOTOR is None:
        write_debug_html(df, "A coluna PROMOTOR não foi encontrada no CSV. Verifique o nome exato na aba JANEIRO_2026.")
        return

    if COL_ORIGEM is None:
        write_debug_html(df, "A coluna ORIGEM não foi encontrada no CSV. Verifique se ela existe e está no cabeçalho.")
        return

    if COL_FREQ is None or COL_TEMPO is None:
        write_debug_html(df, "Não encontrei FREQ.SEMANA e/ou TEMPO DE ATENDIMENTO POR VISITA no CSV. Verifique os nomes no cabeçalho.")
        return

    # Filtro ORIGEM
    df[COL_ORIGEM] = df[COL_ORIGEM].astype(str).str.strip()
    df = df[df[COL_ORIGEM].isin(["CAMIL", "SPOT"])].copy()

    # Limpa texto das colunas principais (se existirem)
    for col in [COL_PROMOTOR, COL_SUPERVISOR, COL_LOJA, COL_REGIONAL, COL_CIDADE, COL_BANDEIRA]:
        if col is not None and col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # Converte numéricos
    df[COL_FREQ] = pd.to_numeric(df[COL_FREQ], errors="coerce").fillna(0)
    df[COL_TEMPO] = pd.to_numeric(df[COL_TEMPO], errors="coerce").fillna(0)

    # Cálculo horas semana loja
    df["HORAS_SEMANA_LOJA"] = df[COL_FREQ] * df[COL_TEMPO]

    # Horas por promotor
    horas_prom = (
        df.groupby(COL_PROMOTOR, dropna=False)["HORAS_SEMANA_LOJA"]
        .sum()
        .reset_index()
        .rename(columns={"HORAS_SEMANA_LOJA": "HORAS_SEMANA_PROMOTOR"})
    )

    horas_prom["HORAS_MES_PREVISTA"] = horas_prom["HORAS_SEMANA_PROMOTOR"] * 4
    horas_prom["SALDO_HORAS"] = TETO_CLT - horas_prom["HORAS_MES_PREVISTA"]
    horas_prom["ULTRAPASSOU_TETO"] = horas_prom["HORAS_MES_PREVISTA"] > TETO_CLT
    horas_prom["HORAS_EXCEDENTES"] = (horas_prom["HORAS_MES_PREVISTA"] - TETO_CLT).clip(lower=0)

    # KPIs
    total_linhas = int(len(df))
    qtd_promotores = safe_nunique(df[COL_PROMOTOR])
    qtd_supervisores = safe_nunique(df[COL_SUPERVISOR]) if (COL_SUPERVISOR in df.columns) else 0
    qtd_lojas = safe_nunique(df[COL_LOJA]) if (COL_LOJA in df.columns) else 0
    qtd_bandeiras = safe_nunique(df[COL_BANDEIRA]) if (COL_BANDEIRA in df.columns) else 0
    qtd_cidades = safe_nunique(df[COL_CIDADE]) if (COL_CIDADE in df.columns) else 0

    total_horas_mes_prevista = safe_sum(horas_prom["HORAS_MES_PREVISTA"])
    qtd_promotores_acima_teto = int(horas_prom["ULTRAPASSOU_TETO"].sum())
    total_horas_excedentes = safe_sum(horas_prom["HORAS_EXCEDENTES"])

    # Médias
    if COL_SUPERVISOR in df.columns:
        prom_por_sup = df.dropna(subset=[COL_SUPERVISOR, COL_PROMOTOR]).groupby(COL_SUPERVISOR)[COL_PROMOTOR].nunique()
        media_prom_por_sup = float(prom_por_sup.mean()) if len(prom_por_sup) else 0.0
    else:
        media_prom_por_sup = 0.0

    if (COL_SUPERVISOR in df.columns) and (COL_LOJA in df.columns):
        lojas_por_sup = df.dropna(subset=[COL_SUPERVISOR, COL_LOJA]).groupby(COL_SUPERVISOR)[COL_LOJA].nunique()
        media_lojas_por_sup = float(lojas_por_sup.mean()) if len(lojas_por_sup) else 0.0
    else:
        media_lojas_por_sup = 0.0

    if COL_LOJA in df.columns:
        lojas_por_prom = df.dropna(subset=[COL_PROMOTOR, COL_LOJA]).groupby(COL_PROMOTOR)[COL_LOJA].nunique()
        media_lojas_por_prom = float(lojas_por_prom.mean()) if len(lojas_por_prom) else 0.0
    else:
        media_lojas_por_prom = 0.0

    # Top 20 promotores
    top = horas_prom.sort_values("HORAS_MES_PREVISTA", ascending=False).head(20)

    # =========================
    # RESUMO POR SUPERVISOR FINAL
    # =========================
    if COL_SUPERVISOR in df.columns:
        # base: promotores, lojas, cidades por supervisor
        agg_dict = {
            "PROMOTORES_SUPERVISOR": (COL_PROMOTOR, "nunique"),
        }
        if (COL_LOJA is not None) and (COL_LOJA in df.columns):
            agg_dict["LOJAS_SUPERVISOR"] = (COL_LOJA, "nunique")
        else:
            agg_dict["LOJAS_SUPERVISOR"] = ("HORAS_SEMANA_LOJA", "count")

        if (COL_CIDADE is not None) and (COL_CIDADE in df.columns):
            agg_dict["CIDADES_SUPERVISOR"] = (COL_CIDADE, "nunique")
        else:
            agg_dict["CIDADES_SUPERVISOR"] = ("HORAS_SEMANA_LOJA", "count")

        sup_base = df.groupby(COL_SUPERVISOR).agg(**agg_dict).reset_index()

        # horas do time: precisa ligar PROMOTOR -> SUPERVISOR
        promotor_sup = (
            df[[COL_PROMOTOR, COL_SUPERVISOR]]
            .dropna()
            .drop_duplicates()
        )

        horas_prom_sup = horas_prom.merge(promotor_sup, on=COL_PROMOTOR, how="left")

        sup_horas = horas_prom_sup.groupby(COL_SUPERVISOR).agg(
            HORAS_MES_TIME=("HORAS_MES_PREVISTA", "sum"),
            PROMOTORES_ACIMA_TETO=("ULTRAPASSOU_TETO", "sum"),
            HORAS_EXCEDENTES_TIME=("HORAS_EXCEDENTES", "sum"),
        ).reset_index()

        sup_resumo = sup_base.merge(sup_horas, on=COL_SUPERVISOR, how="left").fillna(0)

        # média lojas/promotor dentro do supervisor
        sup_resumo["MEDIA_LOJAS_POR_PROMOTOR_NO_SUP"] = (
            sup_resumo["LOJAS_SUPERVISOR"] / sup_resumo["PROMOTORES_SUPERVISOR"].replace({0: 1})
        )

        sup_resumo = sup_resumo.sort_values("HORAS_MES_TIME", ascending=False)
    else:
        sup_resumo = None

    # =========================
    # RESUMO POR REGIONAL VOLUME
    # =========================
    if (COL_REGIONAL is not None) and (COL_REGIONAL in df.columns):
        agg_reg = {
            "PROMOTORES_REGIONAL": (COL_PROMOTOR, "nunique"),
        }

        if (COL_SUPERVISOR is not None) and (COL_SUPERVISOR in df.columns):
            agg_reg["SUPERVISORES_REGIONAL"] = (COL_SUPERVISOR, "nunique")
        else:
            agg_reg["SUPERVISORES_REGIONAL"] = ("HORAS_SEMANA_LOJA", "count")

        if (COL_LOJA is not None) and (COL_LOJA in df.columns):
            agg_reg["LOJAS_REGIONAL"] = (COL_LOJA, "nunique")
        else:
            agg_reg["LOJAS_REGIONAL"] = ("HORAS_SEMANA_LOJA", "count")

        if (COL_CIDADE is not None) and (COL_CIDADE in df.columns):
            agg_reg["CIDADES_REGIONAL"] = (COL_CIDADE, "nunique")
        else:
            agg_reg["CIDADES_REGIONAL"] = ("HORAS_SEMANA_LOJA", "count")

        reg_base = df.groupby(COL_REGIONAL).agg(**agg_reg).reset_index()

        # horas por regional: ligar promotor -> regional e somar horas promotor
        promotor_reg = (
            df[[COL_PROMOTOR, COL_REGIONAL]]
            .dropna()
            .drop_duplicates()
        )
        horas_prom_reg = horas_prom.merge(promotor_reg, on=COL_PROMOTOR, how="left")

        reg_horas = horas_prom_reg.groupby(COL_REGIONAL).agg(
            HORAS_MES_REGIONAL=("HORAS_MES_PREVISTA", "sum"),
            PROMOTORES_ACIMA_TETO=("ULTRAPASSOU_TETO", "sum"),
            HORAS_EXCEDENTES_REGIONAL=("HORAS_EXCEDENTES", "sum"),
        ).reset_index()

        reg_resumo = reg_base.merge(reg_horas, on=COL_REGIONAL, how="left").fillna(0)

        reg_resumo["MEDIA_LOJAS_POR_PROMOTOR_NO_REG"] = (
            reg_resumo["LOJAS_REGIONAL"] / reg_resumo["PROMOTORES_REGIONAL"].replace({0: 1})
        )

        reg_resumo = reg_resumo.sort_values("HORAS_MES_REGIONAL", ascending=False)
    else:
        reg_resumo = None

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

    # =========================
    # HTML - TABELA REGIONAIS
    # =========================
    if reg_resumo is not None:
        reg_rows = []
        for _, r in reg_resumo.iterrows():
            reg = str(r.get(COL_REGIONAL, "")).strip()
            reg_rows.append(f"""
              <tr>
                <td>{reg}</td>
                <td style="text-align:right;">{int(r.get("SUPERVISORES_REGIONAL", 0))}</td>
                <td style="text-align:right;">{int(r.get("PROMOTORES_REGIONAL", 0))}</td>
                <td style="text-align:right;">{int(r.get("LOJAS_REGIONAL", 0))}</td>
                <td style="text-align:right;">{int(r.get("CIDADES_REGIONAL", 0))}</td>
                <td style="text-align:right;">{fmt_num(r.get("MEDIA_LOJAS_POR_PROMOTOR_NO_REG", 0))}</td>
                <td style="text-align:right;">{fmt_num(r.get("HORAS_MES_REGIONAL", 0))}</td>
                <td style="text-align:right;">{int(r.get("PROMOTORES_ACIMA_TETO", 0))}</td>
                <td style="text-align:right;">{fmt_num(r.get("HORAS_EXCEDENTES_REGIONAL", 0))}</td>
              </tr>
            """)

        tabela_regionais_html = f"""
        <h2 style="margin-top:22px;">Resumo por Regional Volume</h2>
        <div style="overflow:auto;border:1px solid #eee;border-radius:10px;">
          <table style="width:100%;border-collapse:collapse;">
            <thead>
              <tr style="background:#f7f7f7;">
                <th style="text-align:left;padding:10px;border-bottom:1px solid #eee;">Regional</th>
                <th style="text-align:right;padding:10px;border-bottom:1px solid #eee;">Supervisores</th>
                <th style="text-align:right;padding:10px;border-bottom:1px solid #eee;">Promotores</th>
                <th style="text-align:right;padding:10px;border-bottom:1px solid #eee;">Lojas</th>
                <th style="text-align:right;padding:10px;border-bottom:1px solid #eee;">Cidades</th>
                <th style="text-align:right;padding:10px;border-bottom:1px solid #eee;">Média lojas/promotor</th>
                <th style="text-align:right;padding:10px;border-bottom:1px solid #eee;">Horas mês (regional)</th>
                <th style="text-align:right;padding:10px;border-bottom:1px solid #eee;">Prom. &gt; {TETO_CLT}h</th>
                <th style="text-align:right;padding:10px;border-bottom:1px solid #eee;">Horas excedentes</th>
              </tr>
            </thead>
            <tbody>
              {''.join(reg_rows)}
            </tbody>
          </table>
        </div>
        """
    else:
        tabela_regionais_html = ""

    # =========================
    # HTML - TABELA SUPERVISORES
    # =========================
    if sup_resumo is not None:
        sup_rows = []
        for _, r in sup_resumo.iterrows():
            sup = str(r.get(COL_SUPERVISOR, "")).strip()
            sup_rows.append(f"""
              <tr>
                <td>{sup}</td>
                <td style="text-align:right;">{int(r.get("PROMOTORES_SUPERVISOR", 0))}</td>
                <td style="text-align:right;">{int(r.get("LOJAS_SUPERVISOR", 0))}</td>
                <td style="text-align:right;">{int(r.get("CIDADES_SUPERVISOR", 0))}</td>
                <td style="text-align:right;">{fmt_num(r.get("MEDIA_LOJAS_POR_PROMOTOR_NO_SUP", 0))}</td>
                <td style="text-align:right;">{fmt_num(r.get("HORAS_MES_TIME", 0))}</td>
                <td style="text-align:right;">{int(r.get("PROMOTORES_ACIMA_TETO", 0))}</td>
                <td style="text-align:right;">{fmt_num(r.get("HORAS_EXCEDENTES_TIME", 0))}</td>
              </tr>
            """)

        tabela_supervisores_html = f"""
        <h2 style="margin-top:22px;">Resumo por Supervisor Final</h2>
        <div style="overflow:auto;border:1px solid #eee;border-radius:10px;">
          <table style="width:100%;border-collapse:collapse;">
            <thead>
              <tr style="background:#f7f7f7;">
                <th style="text-align:left;padding:10px;border-bottom:1px solid #eee;">Supervisor</th>
                <th style="text-align:right;padding:10px;border-bottom:1px solid #eee;">Promotores</th>
                <th style="text-align:right;padding:10px;border-bottom:1px solid #eee;">Lojas</th>
                <th style="text-align:right;padding:10px;border-bottom:1px solid #eee;">Cidades</th>
                <th style="text-align:right;padding:10px;border-bottom:1px solid #eee;">Média lojas/promotor</th>
                <th style="text-align:right;padding:10px;border-bottom:1px solid #eee;">Horas mês (time)</th>
                <th style="text-align:right;padding:10px;border-bottom:1px solid #eee;">Prom. &gt; {TETO_CLT}h</th>
                <th style="text-align:right;padding:10px;border-bottom:1px solid #eee;">Horas excedentes</th>
              </tr>
            </thead>
            <tbody>
              {''.join(sup_rows)}
            </tbody>
          </table>
        </div>
        """
    else:
        tabela_supervisores_html = ""

    # =========================
    # HTML - TOP 20 PROMOTORES
    # =========================
    rows = []
    for _, r in top.iterrows():
        prom = str(r.get(COL_PROMOTOR, "")).strip()
        hs = float(r.get("HORAS_SEMANA_PROMOTOR", 0))
        hm = float(r.get("HORAS_MES_PREVISTA", 0))
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

  {tabela_regionais_html}

  {tabela_supervisores_html}

  {tabela_html}

  <details style="margin-top:20px;">
    <summary style="cursor:pointer;font-weight:bold;">Ver colunas lidas (debug)</summary>
    <div style="padding:12px;border:1px solid #ddd;border-radius:10px;line-height:1.6;">
      {"<br>".join([str(c) for c in df.columns])}
    </div>
  </details>
</body>
</html>
"""

    with open("docs/index.html", "w", encoding="utf-8") as f:
        f.write(html)


if __name__ == "__main__":
    main()
