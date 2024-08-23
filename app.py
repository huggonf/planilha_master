import streamlit as st
import pandas as pd
from io import BytesIO

# Funções para o processo "SUB"
def consolidar_dados(planilha, codigos):
    """Consolida dados de todas as abas de uma planilha .xlsx, aplicando filtros."""
    dados_consolidados = pd.DataFrame()
    xl = pd.ExcelFile(planilha)

    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name)
        df['Master'] = df.apply(lambda row: row['codcor'] if pd.isna(row['Master']) else row['Master'], axis=1)
        df_filtrado = df[df['codcor'].isin(codigos)]
        dados_consolidados = pd.concat([dados_consolidados, df_filtrado])

    return dados_consolidados

def salvar_planilhas_por_valor(df, nome_planilha):
    """Cria arquivos Excel separados em memória para cada valor único de 'codcor'."""
    arquivos = {}
    
    for valor in df['codcor'].dropna().unique():
        subconjunto = df[df['codcor'] == valor]
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            subconjunto.to_excel(writer, index=False)
        buffer.seek(0)
        arquivos[f'{int(valor)}_{nome_planilha}'] = buffer

    # Tratamento para valores NaN
    subconjunto_nan = df[df['codcor'].isna()]
    if not subconjunto_nan.empty:
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            subconjunto_nan.to_excel(writer, index=False)
        buffer.seek(0)
        arquivos[f'NaN_{nome_planilha}'] = buffer
    
    return arquivos

# Funções para o processo "MASTER"
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

def salvar_em_arquivos(dados_por_master, nome_planilha):
    """Salva os dados filtrados em arquivos Excel separados, um para cada código Master."""
    arquivos = []
    for master, df in dados_por_master.items():
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Dados Filtrados', index=False)
        output.seek(0)
        arquivos.append((f"Master_{master}_{nome_planilha}", output))
    return arquivos

# Função principal para a aplicação Streamlit
def main():
    st.title("Extrator Excel Bônus Semanal")
    st.markdown("Antes de fazer o upload do arquivo, certifique-se que a coluna do sub está como codcor e a do master como Master")
    
    # Escolha do tipo de processamento
    opcao = st.selectbox("Escolha o tipo de processamento", ["SUB", "MASTER"])

    # Seção de upload de arquivo
    uploaded_file = st.file_uploader("Anexe a Planilha", type="xlsx")
    
    if uploaded_file:
        nome_planilha = uploaded_file.name
        
        if opcao == "SUB":
            # Processo SUB
            codigos_input = st.text_input("Digite os códigos para filtrar, separados por vírgula (exemplo: 93857,92007,9323)")
            
            if codigos_input:
                codigos = [int(codigo.strip()) for codigo in codigos_input.split(',') if codigo.strip().isdigit()]
                
                if codigos:
                    st.write(f"Filtrando com os códigos: {codigos}")
                    dados_consolidados = consolidar_dados(uploaded_file, codigos)
                    
                    # Salvar e disponibilizar arquivos para download
                    arquivos_excel = salvar_planilhas_por_valor(dados_consolidados, nome_planilha)
                    
                    for nome_arquivo, arquivo_buffer in arquivos_excel.items():
                        st.download_button(
                            label=f'Baixar {nome_arquivo}',
                            data=arquivo_buffer,
                            file_name=nome_arquivo,
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )
                else:
                    st.error("Por favor, insira códigos válidos.")
            else:
                st.warning("Digite os códigos para filtrar.")
        
        elif opcao == "MASTER":
            # Processo MASTER
            user_input = st.text_input("Digite os códigos para filtrar, separados por vírgula (exemplo: 93857,92007,9323)")
            numeros_list = converter_input_para_float(user_input)
            
            if numeros_list:
                dados_por_master = processar_dados(pd.ExcelFile(uploaded_file), numeros_list)
                
                if dados_por_master:
                    arquivos = salvar_em_arquivos(dados_por_master, nome_planilha)
                    
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
