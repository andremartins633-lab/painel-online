import re
import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

# =========================
# Configurações do App
# =========================
st.set_page_config(page_title="Painel Online (Google Sheets)", layout="wide")
st.title("Painel Online — baseado no seu Excel (rodando no Google Sheets)")
st.caption("A1:E31 = entrada | F:J = fórmulas (no Sheets) | L1:N14 = resultados")

# Lê Secrets
SHEET_URL = st.secrets["sheet_url"]
SHEET_NAME = st.secrets.get("sheet_name", "PAINEL")
SERVICE_EMAIL = st.secrets["gcp_service_account"]["client_email"]
st.caption(f"Conectando como: {SERVICE_EMAIL}")

# Faixas fixas
INPUT_RANGE = "A1:E31"     # área de edição (com cabeçalho em A1:E1)
RESULT_RANGE = "L1:N14"    # resultados
PAISES = ["Bolivia", "Paraguai", "Argentina"]  # dropdown da coluna D

# =========================
# Conexão com Google APIs
# =========================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# 🔧 CORREÇÃO: normaliza a chave privada (corrige \\n -> \n)
sa_info = dict(st.secrets["gcp_service_account"])
if isinstance(sa_info.get("private_key"), str):
    sa_info["private_key"] = sa_info["private_key"].replace("\\n", "\n")

creds = Credentials.from_service_account_info(sa_info, scopes=SCOPES)
gc = gspread.authorize(creds)

# extrai ID da planilha para fallback
m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", SHEET_URL)
SHEET_ID = m.group(1) if m else None

def open_sheet():
    try:
        # preferimos por ID, é mais robusto (URL pode ter querystring ?gid=)
        if SHEET_ID:
            return gc.open_by_key(SHEET_ID)
        return gc.open_by_url(SHEET_URL)
    except gspread.exceptions.SpreadsheetNotFound:
        st.error("Planilha não encontrada. Confira `sheet_url` nos Secrets e se a Service Account tem acesso (Editor).")
        st.caption(f"Service account: {SERVICE_EMAIL}")
        st.stop()
    except gspread.exceptions.APIError:
        st.error("Falha de API ao abrir a planilha. Habilite as APIs **Google Sheets** e **Google Drive** no projeto do seu JSON.")
        st.stop()
    except Exception as e:
        st.error(f"Erro ao abrir a planilha: {e}")
        st.caption(f"Service account: {SERVICE_EMAIL}")
        st.stop()

try:
    sh = open_sheet()
    ws = sh.worksheet(SHEET_NAME)
except gspread.exceptions.WorksheetNotFound:
    st.error(f"Aba '{SHEET_NAME}' não encontrada. Verifique `sheet_name` nos Secrets.")
    st.stop()

# =========================
# Utilitários de Range
# =========================
def read_range_as_df(worksheet, cell_range: str, headers=True, width=None, height=None):
    vals = worksheet.get(cell_range, value_render_option="UNFORMATTED_VALUE") or []
    # Normaliza quantidade de linhas/colunas para manter grade estável
    if height is not None:
        while len(vals) < height:
            vals.append([])
    if width is not None:
        vals = [row + [""] * (width - len(row)) for row in vals]

    df = pd.DataFrame(vals)
    if headers and len(df) > 0:
        df.columns = df.iloc[0].fillna("").astype(str)
        df = df[1:].reset_index(drop=True)
    else:
        df.columns = [f"Col_{i+1}" for i in range(df.shape[1])]
    return df.fillna("")

def write_df_to_range(worksheet, cell_range: str, df: pd.DataFrame, include_headers=True):
    data = []
    if include_headers:
        data.append(list(df.columns))
    # substitui NaN por string vazia para evitar "nan" no Sheets
    rows = df.astype(object).where(pd.notnull(df), "").values.tolist()
    data.extend(rows)
    worksheet.update(cell_range, data, value_input_option="USER_ENTERED")

# =========================
# ENTRADA: A1:E31 (com dropdown em D)
# =========================
# Vamos ler A1:E31 assumindo que a primeira linha (A1:E1) é cabeçalho.
# Altura total incluindo cabeçalho = 32 linhas; largura = 5 colunas.
df_inputs = read_range_as_df(ws, "A1:E32", headers=True, width=5, height=32)

# Garante exatamente 31 linhas de dados (sem contar o cabeçalho)
if df_inputs.shape[0] < 31:
    add = 31 - df_inputs.shape[0]
    df_inputs = pd.concat(
        [df_inputs, pd.DataFrame([[""] * df_inputs.shape[1]] * add, columns=df_inputs.columns)],
        ignore_index=True
    )
elif df_inputs.shape[0] > 31:
    df_inputs = df_inputs.iloc[:31].copy()

st.subheader("Entrada (A1:E31)")

col_configs = {}
if df_inputs.shape[1] >= 4:
    d_col_name = df_inputs.columns[3]  # 4ª coluna = D
    col_configs[d_col_name] = st.column_config.SelectboxColumn(
        label=d_col_name,
        options=PAISES,
        help="Escolha: Bolivia, Paraguai ou Argentina",
        required=False
    )

edited = st.data_editor(
    df_inputs,
    num_rows=31,  # fixa total de linhas
    hide_index=True,
    use_container_width=True,
    column_config=col_configs,
    key="editor_inputs"
)

c1, c2 = st.columns(2)
with c1:
    if st.button("Salvar no Sheets", type="primary"):
        try:
            write_df_to_range(ws, INPUT_RANGE, edited, include_headers=True)
            st.success("Dados salvos! O Google Sheets recalcula as fórmulas automaticamente.")
            st.toast("Google Sheets atualizado ✅", icon="✅")
        except Exception as e:
            st.error(f"Erro ao salvar: {e}")

with c2:
    if st.button("Recarregar da planilha"):
        st.rerun()

st.divider()

# =========================
# RESULTADOS: L1:N14
# =========================
st.subheader("Resultados (L1:N14)")
try:
    # Lê 14 linhas (inclui o cabeçalho) e 3 colunas
    df_result = read_range_as_df(ws, "L1:N14", headers=True, width=3, height=14)
    st.dataframe(df_result, hide_index=True, use_container_width=True)
except Exception as e:
    st.error(f"Erro ao ler L1:N14: {e}")
    st.caption("Confirme se as fórmulas e referências estão corretas na planilha.")

