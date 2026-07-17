"""
Coletor de Dados de Mercado - Dashboard Inteligência de Mercado
Fontes:
  - World Bank Pink Sheet  → preços mensais em US$/mt (referência institucional)
  - yfinance               → preços semanais CBOT + câmbio BRL=X
  - Planilha interna       → preços de compra anuais por cultivar

Saída: Excel com 4 abas
  1. Commodities_Mensal   (World Bank, US$/mt)
  2. Commodities_Semanal  (yfinance, cents/bushel + câmbio + convertido R$)
  3. Cambio_Mensal        (yfinance BRL=X agregado por mês)
  4. Sementes             (planilha interna, formato vertical)
"""

import yfinance as yf
import pandas as pd
import requests
from io import BytesIO
from datetime import date, timedelta

# CONFIGURAÇÕES

DATA_INICIO      = date(2007, 1, 1)
CAMINHO_SAIDA    = "commodities_cambio_tratado.xlsx"
CAMINHO_SEMENTES = "preco_maximo_safra.xlsx"  

# URL permanente do World Bank Pink Sheet (arquivo mensal histórico)
URL_WORLD_BANK = (
    "https://thedocs.worldbank.org/en/doc/"
    "74e8be41ceb20fa0da750cda2f6b9e4e-0050012026/related/"
    "CMO-Historical-Data-Monthly.xlsx"
)

# Nomes EXATOS das colunas no arquivo do World Bank (aba "Monthly Prices")
# confirmados via inspeção manual do arquivo (colunas 24, 30, 37)
WB_COLUNAS = {
    "Milho":  "Maize",          # US$/mt — coluna 30
    "Soja":   "Soybeans",       # US$/mt — coluna 24
    "Trigo":  "Wheat, US HRW"   # US$/mt — coluna 37
}

TICKERS_COMMODITIES = {
    "Milho": "ZC=F",   # cents/bushel
    "Soja":  "ZS=F",   # cents/bushel
    "Trigo": "ZW=F",   # cents/bushel
}

# ─────────────────────────────────────────────
# BLOCO 1: Data de corte dinâmica
# ─────────────────────────────────────────────

def get_data_corte():
    """Último dia do mês fechado anterior ao mês atual."""
    hoje = date.today()
    return hoje.replace(day=1) - timedelta(days=1)

DATA_CORTE = get_data_corte()
print(f"Coletando dados de {DATA_INICIO} até {DATA_CORTE}\n")

# ─────────────────────────────────────────────
# BLOCO 2: World Bank Pink Sheet — MENSAL
# ─────────────────────────────────────────────

def buscar_worldbank_mensal():
    """
    Baixa o CMO-Historical-Data-Monthly.xlsx do World Bank e extrai
    Milho, Soja e Trigo em US$/mt como série mensal desde 2007.

    Particularidades da estrutura do arquivo (confirmadas por inspeção manual):
      - Linha 4 (0-indexed): nomes das commodities
      - Linha 5 (0-indexed): unidades, ex: "($/mt)" — precisa ser descartada
      - Linha 6 em diante: dados, com índice no formato "1960M01" (ano+M+mês)
      - Valores ausentes aparecem como "…" (reticências), não célula vazia
    """
    print("=== WORLD BANK PINK SHEET (MENSAL) ===")
    try:
        resp = requests.get(URL_WORLD_BANK, timeout=60)
        resp.raise_for_status()
        arquivo = BytesIO(resp.content)
    except Exception as e:
        print(f"  ERRO ao baixar World Bank: {e}")
        return pd.DataFrame()

    try:
        # header=4 usa a linha de nomes de commodity como cabeçalho
        # skiprows=[5] descarta a linha de unidades logo em seguida
        df_raw = pd.read_excel(
            arquivo,
            sheet_name="Monthly Prices",
            header=4,
            skiprows=[5],
            index_col=0,
            na_values=["…", "..", "...", "…\u200b"]
        )
    except Exception as e:
        print(f"  ERRO ao ler aba 'Monthly Prices': {e}")
        return pd.DataFrame()

    # seleciona só as colunas de interesse pelo nome exato
    colunas_usar = {}
    for nome, col_wb in WB_COLUNAS.items():
        if col_wb in df_raw.columns:
            colunas_usar[nome] = col_wb
        else:
            print(f"  AVISO: coluna '{col_wb}' não encontrada para {nome}")

    if not colunas_usar:
        print("  ERRO: nenhuma coluna de commodity encontrada — confira WB_COLUNAS")
        return pd.DataFrame()

    df = df_raw[list(colunas_usar.values())].copy()
    df.columns = list(colunas_usar.keys())

    # converte índice "1960M01" -> Timestamp usando parsing manual
    # formato: AAAA + "M" + MM
    datas_str = df.index.astype(str).str.replace("M", "-", regex=False)
    df.index = pd.to_datetime(datas_str, format="%Y-%m", errors="coerce")
    df = df[df.index.notna()]

    # filtra período de interesse
    df = df[(df.index >= pd.Timestamp(DATA_INICIO)) &
            (df.index <= pd.Timestamp(DATA_CORTE))]
    df = df.apply(pd.to_numeric, errors="coerce")
    df = df.dropna(how="all")

    # formato final: Data | Ano | Mes | Commodity | Preco_USD_mt
    registros = []
    for commodity in df.columns:
        serie = df[[commodity]].copy()
        serie.columns = ["Preco_USD_mt"]
        serie["Commodity"] = commodity
        serie["Ano"]  = serie.index.year
        serie["Mes"]  = serie.index.month
        serie["Data"] = serie.index.strftime("%Y-%m")
        registros.append(serie.reset_index(drop=True))

    resultado = pd.concat(registros, ignore_index=True)
    resultado = resultado[["Data", "Ano", "Mes", "Commodity", "Preco_USD_mt"]].sort_values(
        ["Commodity", "Data"]
    )

    print(f"  OK: {resultado['Commodity'].nunique()} commodities × "
          f"{resultado['Data'].nunique()} meses coletados")
    return resultado


# ─────────────────────────────────────────────
# BLOCO 3: yfinance — SEMANAL (CBOT + câmbio)
# ─────────────────────────────────────────────

def buscar_yfinance_semanal():
    """
    Baixa dados diários de CBOT e BRL=X via yfinance,
    agrega para frequência SEMANAL (última cotação da semana).
    """
    print("\n=== YFINANCE — SEMANAL (CBOT + CÂMBIO) ===")

    todos = list(TICKERS_COMMODITIES.values()) + ["BRL=X"]
    dados = yf.download(
        todos,
        start=DATA_INICIO,
        end=DATA_CORTE,
        progress=False,
        auto_adjust=True
    )
    fechamento = dados["Close"]

    # agrega diário → semanal (última cotação de cada semana)
    # W-FRI: semana encerra na sexta-feira
    fechamento_semanal = fechamento.resample("W-FRI").last()
    fechamento_semanal = fechamento_semanal.dropna(how="all")

    registros = []
    for nome, ticker in TICKERS_COMMODITIES.items():
        if ticker not in fechamento_semanal.columns:
            print(f"  AVISO: {ticker} não encontrado")
            continue

        serie = fechamento_semanal[[ticker, "BRL=X"]].copy()
        serie.columns = ["Preco_USD_CBOT", "Cambio_USD_BRL"]
        serie = serie.dropna()

        # cents/bushel → US$/mt (fatores de conversão padrão CBOT)
        # milho/trigo: 1 bushel = 25.4012 kg → 1 mt = 39.368 bushels
        # soja:        1 bushel = 27.2155 kg → 1 mt = 36.744 bushels
        fatores = {"Milho": 39.368, "Soja": 36.744, "Trigo": 39.368}
        fator = fatores.get(nome, 39.368)
        serie["Preco_USD_mt"]        = (serie["Preco_USD_CBOT"] / 100 * fator).round(2)
        serie["Preco_BRL_Convertido"] = (serie["Preco_USD_mt"] * serie["Cambio_USD_BRL"]).round(2)

        serie["Commodity"]   = nome
        serie["Data_Semana"] = serie.index.strftime("%Y-%m-%d")
        serie["Ano"]         = serie.index.year
        serie["Mes"]         = serie.index.month
        serie["Semana"]      = serie.index.isocalendar().week.astype(int)

        registros.append(serie.reset_index(drop=True))
        print(f"  OK: {nome} → {len(serie)} semanas coletadas")

    resultado = pd.concat(registros, ignore_index=True)
    resultado = resultado[[
        "Data_Semana", "Ano", "Mes", "Semana", "Commodity",
        "Preco_USD_CBOT", "Preco_USD_mt", "Cambio_USD_BRL", "Preco_BRL_Convertido"
    ]].sort_values(["Commodity", "Data_Semana"])

    return resultado


# ─────────────────────────────────────────────
# BLOCO 4: Câmbio mensal (para referência do WB)
# ─────────────────────────────────────────────

def buscar_cambio_mensal():
    """
    Extrai BRL=X do yfinance e agrega para média mensal.
    Útil para converter os dados do World Bank para R$.
    """
    print("\n=== CÂMBIO MENSAL (BRL=X via yfinance) ===")

    dados = yf.download("BRL=X", start=DATA_INICIO, end=DATA_CORTE,
                        progress=False, auto_adjust=True)

    if isinstance(dados.columns, pd.MultiIndex):
        dados.columns = dados.columns.get_level_values(0)

    cambio = dados[["Close"]].copy()
    cambio.columns = ["Cambio_USD_BRL"]
    cambio_mensal = cambio.resample("MS").mean().round(4)  # MS = início do mês
    cambio_mensal["Ano"] = cambio_mensal.index.year
    cambio_mensal["Mes"] = cambio_mensal.index.month
    cambio_mensal["Data"] = cambio_mensal.index.strftime("%Y-%m")
    cambio_mensal = cambio_mensal.reset_index(drop=True)
    cambio_mensal = cambio_mensal[["Data", "Ano", "Mes", "Cambio_USD_BRL"]]

    print(f"  OK: {len(cambio_mensal)} meses de câmbio coletados")
    return cambio_mensal


# ─────────────────────────────────────────────
# BLOCO 5: Transformação da planilha de sementes
# ─────────────────────────────────────────────

def limpar_preco(v):
    """Converte valor de preço para float, respeitando se já é numérico."""
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("R$", "").replace(" ", "")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except:
        return None

def transformar_sementes():
    print("\n=== SEMENTES (PLANILHA INTERNA) ===")
    df = pd.read_excel(
        CAMINHO_SEMENTES,
        sheet_name="Preço Histórico",
        header=2
    )

    df = df.rename(columns={df.columns[0]: "Cultivar"})
    df = df.dropna(subset=["Cultivar"])

    colunas_ano = [
        c for c in df.columns
        if isinstance(c, (int, float)) and 2000 <= int(c) <= 2100
    ]
    df = df[["Cultivar"] + colunas_ano]

    df_vertical = df.melt(
        id_vars="Cultivar",
        value_vars=colunas_ano,
        var_name="Ano",
        value_name="Preco_kg"
    )

    df_vertical["Ano"]      = df_vertical["Ano"].astype(int)
    df_vertical["Preco_kg"] = df_vertical["Preco_kg"].apply(limpar_preco)
    df_vertical = df_vertical.sort_values(["Cultivar", "Ano"]).reset_index(drop=True)

    n_cult    = df_vertical["Cultivar"].nunique()
    n_anos    = df_vertical["Ano"].nunique()
    n_validos = df_vertical["Preco_kg"].notna().sum()
    print(f"  OK: {n_cult} cultivares × {n_anos} anos → "
          f"{len(df_vertical)} linhas ({n_validos} valores preenchidos)")
    return df_vertical


# ─────────────────────────────────────────────
# BLOCO 6: Execução e salvamento
# ─────────────────────────────────────────────

if __name__ == "__main__":
    df_mensal   = buscar_worldbank_mensal()
    df_semanal  = buscar_yfinance_semanal()
    df_cambio   = buscar_cambio_mensal()
    df_sementes = transformar_sementes()

    print(f"\n=== SALVANDO {CAMINHO_SAIDA} ===")
    with pd.ExcelWriter(CAMINHO_SAIDA, engine="openpyxl") as writer:
        if not df_mensal.empty:
            df_mensal.to_excel(writer,  sheet_name="Commodities_Mensal",  index=False)
        df_semanal.to_excel(writer,     sheet_name="Commodities_Semanal", index=False)
        df_cambio.to_excel(writer,      sheet_name="Cambio_Mensal",       index=False)
        df_sementes.to_excel(writer,    sheet_name="Sementes",            index=False)

    print(f"✓ Arquivo salvo: {CAMINHO_SAIDA}")
    if not df_mensal.empty:
        print(f"  Aba 'Commodities_Mensal':  {len(df_mensal)} linhas")
    print(f"  Aba 'Commodities_Semanal': {len(df_semanal)} linhas")
    print(f"  Aba 'Cambio_Mensal':       {len(df_cambio)} linhas")
    print(f"  Aba 'Sementes':            {len(df_sementes)} linhas")