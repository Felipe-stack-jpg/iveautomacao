import pandas as pd
import streamlit as st
import io
import unicodedata
from fpdf import FPDF

def clean_text(col):
    return (
        col.fillna('')
           .astype(str)
           .str.lower()
           .str.strip()
           .apply(lambda x: ''.join(
               c for c in unicodedata.normalize('NFKD', x) if not unicodedata.combining(c)
           ))
           .str.replace(r'\s+', ' ', regex=True)
    )

# Função para gerar PDF
def gerar_pdf_relatorio(df):
    buffer = io.BytesIO()

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.set_title("Relatório de Vendas")

    def escrever_linha(titulo, valor):
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 10, txt=titulo, ln=True)
        pdf.set_font("Arial", '', 12)
        pdf.multi_cell(0, 10, txt=valor)
        pdf.ln(5)

    total_vendas = df['Valor Venda'].sum()
    ticket_medio = df['Valor Venda'].mean()

    vendas_por_dep = df.groupby('Departamento')['Valor Venda'].sum()
    num_vendas_por_dep = df['Departamento'].value_counts()
    vendas_por_agente = df.groupby('Agente')['Valor Venda'].sum()
    ticket_por_agente = df.groupby('Agente')['Valor Venda'].mean()

    escrever_linha(" Total de Vendas", f"R$ {total_vendas:,.2f}")
    escrever_linha(" Ticket Médio de Vendas", f"R$ {ticket_medio:,.2f}")

    escrever_linha(" Total de Vendas por Departamento",
        '\n'.join(f"{dep}: R$ {valor:,.2f}" for dep, valor in vendas_por_dep.items()))

    escrever_linha(" Número de Vendas por Departamento",
        '\n'.join(f"{dep}: {qtd}" for dep, qtd in num_vendas_por_dep.items()))

    escrever_linha(" Total de Vendas por Atendente",
        '\n'.join(f"{agente}: R$ {valor:,.2f}" for agente, valor in vendas_por_agente.items()))

    escrever_linha(" Ticket Médio por Atendente",
        '\n'.join(f"{agente}: R$ {valor:,.2f}" for agente, valor in ticket_por_agente.items()))

    pdf_bytes = pdf.output(dest='S').encode('latin1')  # gera bytes do PDF
    buffer = io.BytesIO(pdf_bytes)
    return buffer



st.title("Relatório Distribuidora de Autopeças")

uploaded_file = st.file_uploader('Escolha o arquivo Excel', type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df['Data Abertura'] = pd.to_datetime(df['Data Abertura'], dayfirst=True)

    # Limpeza para criar chave sem data
    df['chave_sem_data'] = (
        clean_text(df['Nome']) + '_' +
        clean_text(df['Contato']) + '_' +
        clean_text(df['Canal']) + '_' +
        clean_text(df['Agente'])
    )

    # Ordenar por chave e data
    df = df.sort_values(by=['chave_sem_data', 'Data Abertura']).reset_index(drop=True)

    indices_para_manter = []
    grupo_atual = None
    ultimo_tempo = None

    for idx, row in df.iterrows():
        chave = row['chave_sem_data']
        tempo = row['Data Abertura']

        if grupo_atual != chave:
            # Novo grupo
            indices_para_manter.append(idx)
            grupo_atual = chave
            ultimo_tempo = tempo
        else:
            # Mesmo grupo: verifica diferença de tempo com ultimo mantido
            diff = (tempo - ultimo_tempo).total_seconds() / 60  # em minutos
            if diff >= 10:
                indices_para_manter.append(idx)
                ultimo_tempo = tempo
            # Se diff < 10, ignora (considera duplicado)

    df_limpo = df.loc[indices_para_manter]

    # Checkbox para filtrar por data
    filtrar_data = st.checkbox("Filtrar por data específica?")

    if filtrar_data:
        data_minima = df_limpo['Data Abertura'].min().date()
        data_maxima = df_limpo['Data Abertura'].max().date()

        intervalo = st.date_input(
            "Escolha o período",
            value=(data_minima, data_maxima)
        )

        if isinstance(intervalo, tuple) and len(intervalo) == 2:
            data_inicio, data_fim = intervalo
            df_filtrado = df_limpo[
                (df_limpo['Data Abertura'] >= pd.to_datetime(data_inicio)) &
                (df_limpo['Data Abertura'] <= pd.to_datetime(data_fim))
            ]
        else:
            st.warning("Por favor, selecione um **período completo** com data inicial e final.")
            df_filtrado = df_limpo.copy()
    else:
        df_filtrado = df_limpo.copy()

    # Garante que a coluna Nome apareça como texto no Excel
    df_filtrado['Nome'] = df_filtrado['Nome'].astype(str)
    df_filtrado['Contato'] = df_filtrado['Contato'].astype(str)

    buffer = io.BytesIO()
    df_filtrado.to_excel(buffer, index=False, engine='openpyxl')
    buffer.seek(0)

    st.download_button(
        label="Baixar planilha limpa (Excel)",
        data=buffer,
        file_name="planilha_limpa.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    if not df_filtrado.empty:
        pdf_buffer = gerar_pdf_relatorio(df_filtrado)
        st.download_button(
            label=" Baixar Relatório em PDF",
            data=pdf_buffer,
            file_name="relatorio_vendas.pdf",
            mime="application/pdf"
        )
    else:
        st.warning("Nenhum dado encontrado no período selecionado para gerar o relatório.")

st.markdown(
    """
    <div style='text-align: center; color: gray; margin-top: 30px; font-size: 12px;'>
        © 2025 Dev Felipe
    </div>
    """,
    unsafe_allow_html=True
)