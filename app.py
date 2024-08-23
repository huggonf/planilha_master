import streamlit as st
import pandas as pd
from io import BytesIO

def converter_input_para_float(user_input):
    """Converte a entrada do usuário em uma lista de números float."""
    try:
        return [float(num.strip()) for num in user_input.split(',')]
    except ValueError:
        st.error("Entrada inválida! Certifique-se de digitar apenas números separados por vírgula.")
        return []

def processar_dados(xl, numeros_list):
    """Processa todas as abas e retorna um dicionário com dados filtrados por código Master."""
    dados_por_master = {}
    
    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name)
        
        if 'Master' not in df.columns or 'codcor' not in df.columns:
            st.error(f"Colunas 'Master' ou 'codcor' não encontradas na aba '{sheet_name}'.")
            continue

        df['Master'] = df.apply(lambda row: row['codcor'] if pd.isna(row['Master']) else row['Master'], axis=1)
        df['Master'] = df['Master'].fillna(0).astype(float)
        df_filtrado = df[df['Master'].isin(numeros_list)]
        
        for master in df_filtrado['Master'].unique():
            master = int(master)
            if master not in dados_por_master:
                dados_por_master[master] = pd.DataFrame()
            dados_por_master[master] = pd.concat([dados_por_master[master], df_filtrado[df_filtrado['Master'] == master]])
    
    return dados_por_master

def salvar_em_arquivos(dados_por_master):
    """Salva os dados filtrados em arquivos Excel separados, um para cada código Master."""
    arquivos = []
    for master, df in dados_por_master.items():
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Dados Filtrados', index=False)
        output.seek(0)
        arquivos.append((f"Master_{master}.xlsx", output))
    return arquivos

def main():
    st.title("Processador de Arquivos Excel")
    
    uploaded_file = st.file_uploader("Escolha um arquivo Excel", type="xlsx")
    if uploaded_file:
        xl = pd.ExcelFile(uploaded_file)
        
        user_input = st.text_input("Digite uma lista de números separados por vírgula (ex: 1,2,3):")
        numeros_list = converter_input_para_float(user_input)
        
        if numeros_list:
            dados_por_master = processar_dados(xl, numeros_list)
            
            if dados_por_master:
                arquivos = salvar_em_arquivos(dados_por_master)
                
                for nome_arquivo, buffer in arquivos:
                    st.download_button(
                        label=f"Baixar {nome_arquivo}",
                        data=buffer,
                        file_name=nome_arquivo,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.write("Nenhum dado encontrado para os códigos fornecidos.")

if __name__ == "__main__":
    main()
