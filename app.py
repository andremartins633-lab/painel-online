import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# ========= CONFIG =========
# Defina em st.secrets:
# st.secrets["gcp_service_account"] = {... json do service account ...}
# st.secrets["sheet_url"] = "https://docs.google.com/spreadsheets/d/SEU_ID/edit"
# st.secrets["sheet_name"] = "PAINEL"   # nome da aba com seu layout
SHEET_URL = st.secrets["sheet_url"]
SHEET_NAME = st.secrets.get("sheet_name", "PAINEL")

# Faixas fixas
INPUT_RANGE = "A1:E31"    # área de digitação
RESULT_RANGE = "L1:N14"   # resultados
# F:J ficam no Sheets como fórmulas, você não precisa mexer aqui.

# Opções da lista suspensa (coluna D)
PAISES = ["Bolivia", "Paraguai", "Argentina"]

# ========= GSPREAD CLIENT =========
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
creds = Credentials.from_service_account_info(
    st.secrets["gcp_service_account"], scopes=SCOPES
)
gc = gspread.authorize(creds)
sh = gc.open_by_url(SHEET_URL)
ws = sh.worksheet(SHEET_NAME)

st.set_page_config(page_title="Painel Online (Google Sheets)", layout="wide")
st.title("Painel Online — Modelo baseado no Excel (agora no Google Sheets)")
st.caption("A1:E31 = entrada | F:J = fórmulas | L1:N14 = resultados")

# ========= FUNÇÕES AUX =========
def read_range_as_df(worksheet, rng, headers=True, width=None, height=None):
    vals = worksheet.get(rng, value_render_option="UNFORMATTED_VALUE") or []
    # Pad right/bottom para manter grade estável
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

def write_df_to_range(worksheet, rng, df, include_headers=True):
    data = []
    if include_headers:
        data.append(list(df.columns))
    data.extend(df.astype(object).where(pd.notnull(df), "").values.tolist())
    worksheet.update(rng, data, value_input_option="USER_ENTERED")

# ========= CARREGAR ENTRADAS (A1:E31) =========
# Se sua primeira linha são cabeçalhos, headers=True. Se não, defina headers=False e forneça nomes.
df_inputs = read_range_as_df(ws, INPUT_RANGE, headers=True, width=5, height=32)

# Garante 31 linhas de dados (sem a linha de cabeçalho)
if df_inputs.shape[0] < 31:
    # anexa linhas em branco até 31
    add = 31 - df_inputs.shape[0]
    df_inputs = pd.concat([df_inputs, pd.DataFrame([[""]*df_inputs.shape[1]]*add, columns=df_inputs.columns)], ignore_index=True)

st.subheader("Entrada (A1:E31)")
# Config da coluna D como Selectbox (lista suspensa)
col_configs = {}
if df_inputs.shape[1] >= 4:
    d_col_name = df_inputs.columns[3]
    col_configs[d_col_name] = st.column_config.SelectboxColumn(
        label=d_col_name,
        options=PAISES,
        help="Escolha: Bolivia, Paraguai ou Argentina",
        required=False
    )

edited = st.data_editor(
    df_inputs,
    num_rows=31,  # fixa 31 linhas
    use_container_width=True,
    column_config=col_configs,
    key="editor_inputs"
)

c1, c2 = st.columns(2)
with c1:
    if st.button("Salvar no Sheets"):
        # grava de volta em A1:E31
        try:
            # Mantém o cabeçalho original
            write_df_to_range(ws, INPUT_RANGE, edited, include_headers=True)
            st.success("Dados salvos com sucesso!")
            st.toast("Google Sheets atualizado", icon="✅")
        except Exception as e:
            st.error(f"Erro ao salvar: {e}")

with c2:
    if st.button("Recarregar"):
        st.rerun()

st.divider()

# ========= LER RESULTADOS (L1:N14) =========
st.subheader("Resultados (L1:N14)")
try:
    df_result = read_range_as_df(ws, RESULT_RANGE, headers=True, width=3, height=14)
    st.dataframe(df_result, use_container_width=True)
except Exception as e:
    st.error(f"Erro ao ler L1:N14: {e}")

st.caption("Obs.: Cálculos (F:J e demais abas) rodam no Google Sheets.")