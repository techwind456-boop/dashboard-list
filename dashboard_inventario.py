import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import json

# --- Configuração Google Sheets usando secrets.toml ---
secrets_json = st.secrets["google_service_account"]["json"]
creds_dict = json.loads(secrets_json)
SCOPE = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPE)
client = gspread.authorize(creds)

# --- ID da planilha ---
SPREADSHEET_ID = "1dthZulTwmj_80LGk4hSaH58CoLAKmw5ypLlUUgsK9hY"
sheet = client.open_by_key(SPREADSHEET_ID)

# --- Funções ---
def load_data(tab_name):
    """Carrega a aba como DataFrame e corrige problemas de cabeçalho"""
    worksheet = sheet.worksheet(tab_name)
    data = worksheet.get_all_records(empty2zero=False)
    df = pd.DataFrame(data)
    
    # Limpa espaços nos nomes das colunas
    df.columns = [str(col).strip() for col in df.columns]
    
    # Checa se Item_EN e Item_PT existem
    if "Item_EN" not in df.columns or "Item_PT" not in df.columns:
        st.error(f"A aba '{tab_name}' precisa ter as colunas 'Item_EN' e 'Item_PT'.")
        return pd.DataFrame()
    
    # Converte todas as colunas de WOs para string
    for col in df.columns[2:]:
        df[col] = df[col].astype(str)
    
    return df

def save_quantities(tab_name, df, wo_column):
    """Salva toda a coluna da WO selecionada no Google Sheets"""
    worksheet = sheet.worksheet(tab_name)

    try:
        col_index = df.columns.get_loc(wo_column) + 1  # 1-based index para Google Sheets
    except KeyError:
        st.error(f"Coluna {wo_column} não encontrada no DataFrame.")
        return

    values = df[wo_column].tolist()

    # Determina letra da coluna
    if col_index <= 26:
        col_letter = chr(64 + col_index)
    else:
        div, mod = divmod(col_index - 1, 26)
        col_letter = chr(64 + div) + chr(65 + mod)

    cell_range = f"{col_letter}2:{col_letter}{len(values)+1}"
    worksheet.update(cell_range, [[v] for v in values], value_input_option='USER_ENTERED')

def load_active_wos():
    """Carrega as WOs ativas da aba Config"""
    try:
        config_df = pd.DataFrame(sheet.worksheet("Config").get_all_records())
        wo_list = [str(wo).strip() for wo in config_df["WO_ativas"].dropna().tolist()]
        return wo_list
    except Exception:
        st.error("Erro ao carregar WOs ativas da aba Config.")
        return []

# --- Streamlit ---
st.set_page_config(page_title="Inventário & Consumíveis", layout="wide")
st.title("Dashboard de Inventário & Consumíveis")

# Dropdown de WOs
wo_list = load_active_wos()
if not wo_list:
    st.warning("Não há WOs ativas na aba Config.")
    st.stop()

selected_wo = st.selectbox("Selecione a WO", wo_list)

# Abas para Inventário e Consumíveis
tabs = st.tabs(["Inventário", "Consumíveis"])
for tab_name, tab in zip(["Inventario", "Consumiveis"], tabs):
    with tab:
        df = load_data(tab_name)
        if df.empty:
            st.warning(f"Nenhum dado encontrado na aba '{tab_name}'.")
            continue
        
        if selected_wo not in df.columns:
            st.error(f"A coluna da WO '{selected_wo}' não existe na aba '{tab_name}'.")
        else:
            st.success(f"Dados carregados do Google Sheets para WO {selected_wo}")
            
            # Tabela editável da WO
            edited_df = st.data_editor(
                df[["Item_EN", "Item_PT", selected_wo]],
                disabled=["Item_EN", "Item_PT"],
                num_rows="fixed"
            )
            
            if st.button(f"Salvar {tab_name} - WO {selected_wo}"):
                save_quantities(tab_name, edited_df, selected_wo)
                st.success(f"Quantidades de '{tab_name}' atualizadas no Google Sheets para WO {selected_wo}")
