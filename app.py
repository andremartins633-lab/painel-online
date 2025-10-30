import re
import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Painel Online (Google Sheets)", layout="wide")
st.title("Painel Online — baseado no seu Excel (rodando no Google Sheets)")
st.caption("A1:E31 = entrada | F:J = fórmulas (no Sheets) | L1:N14 = resultados")

# --------- Config dos Secrets ---------
SHEET_URL = st.secrets["sheet_url"]
SHEET_NAME = st.secrets.get("sheet_name", "PAINEL")
SA = dict(st.secrets["gcp_service_account"])  # cópia mutável
SERVICE_EMAIL = SA.get("client_email", "sem-email")
st.caption(f"Conectando como: {SERVICE_EMAIL}")

# --------- Normalização forte da private_key ---------
def normalize_private_key(pk: str) -> str:
    if not isinstance(pk, str):
        return ""
    pk = pk.strip()

    # Caso tenha sido colada como uma única linha com \\n:
    if "\\n" in pk and "\n" not in pk:
        pk = pk.replace("\\n", "\n")

    # Remove espaços em branco extras nas bordas das linhas
    pk = "\n".join([line.strip() for line in pk.splitlines()])

    # Garante cabeçalho/rodapé
    head = "-----BEGIN PRIVATE KEY-----"
    tail = "-----END PRIVATE KEY-----"
    if head not in pk:
        pk = f"{head}\n{pk}"
    if tail not in pk:
        pk = f"{pk}\n{tail}"

    # Garante \n final
    if not pk.endswith("\n"):
        pk += "\n"
    return pk

try:
    pk = SA.get("private_key", "")
    SA["private_key"] = normalize_private_key(pk)
    # sanity check mínima
    if "BEGIN PRIVATE KEY" not in SA["private_key"] or "END PRIVATE KEY" not in SA["private_key"]:
        raise ValueError("private_key sem BEGIN/END após normalização.")
except Exception as e:
    st.error("Sua private_key ainda está com formatação inválida no Secrets. "
             "Abra Settings → Secrets e cole a chave exatamente como no JSON, "
             "ou em uma única linha com \\n entre as quebras. ")
    st.caption(f"Detalhe técnico: {e}")
    st.stop()

# --------- Conexão Google ---------
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

try:
    creds = Credentials.from_service_account_info(SA, scopes=SCOPES)
    gc = gspread.authorize(creds)
except Exception as e:
    st.error("Falha ao criar credenciais com a private_key informada. "
             "Revise o campo `private_key` nos Secrets (ver dicas abaixo).")
    st.caption(f"Detalhe técnico: {e}")
    st.stop()

# --------- Abrir planilha ---------
m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", SHEET_URL)
SHEET_ID = m.group(1) if m else None

def open_sheet():
    try:
        if SHEET_ID:
            return gc.open_by_key(SHEET_ID)
        return gc.open_by_url(SHEET_URL)
    except gspread.exceptions.SpreadsheetNotFound:
        st.error("Planilha não encontrada. Confira `sheet_url` e se a service account tem acesso (Editor).")
        st.caption(f"Service account: {SERVICE_EMAIL}")
        st.stop()
    except gspread.exceptions.APIError:
        st.error("Falha de API. Habilite **Google Sheets API** e **Google Drive API** no projeto do JSON.")
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

# --------- Utilitários ---------
def read_range_as_df(worksheet, cell_range: str, headers=True, width=None, height=None):
    vals = worksheet.get(cell_range, value_render_option="UNFORMATTED_VALUE") or []
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
    rows = df.astype(object).where(pd.notnull(df), "").values.tolist()
    data.extend(rows)
    worksheet.update(cell_range, data, value_input_option="USER_ENTERED")

# --------- Entrada A1:E31 (dropdown em D) ---------
INPUT_RANGE = "A1:E31"
PAISES = ["Bolivia", "Paraguai", "Argentina"]

df_inputs = read_range_as_df(ws, "A1:E32", headers=True, width=5, height=32)
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
    d_col = df_inputs.columns[3]
    col_configs[d_col] = st.column_config.SelectboxColumn(
        label=d_col,
        options=PAISES,
        help="Escolha: Bolivia, Paraguai ou Argentina",
        required=False
    )

edited = st.data_editor(
    df_inputs,
    num_rows=31,
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

# --------- Resultados L1:N14 ---------
st.subheader("Resultados (L1:N14)")
try:
    df_result = read_range_as_df(ws, "L1:N14", headers=True, width=3, height=14)
    st.dataframe(df_result, hide_index=True, use_container_width=True)
except Exception as e:
    st.error(f"Erro ao ler L1:N14: {e}")
    st.caption("Confirme se as fórmulas e referências estão corretas na planilha.")
