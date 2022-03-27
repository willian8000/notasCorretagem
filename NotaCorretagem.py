import os
import re
import sys
import pandas
import tabula
import tkinter as tk
import streamlit as st
from streamlit import cli as stcli
from tkinter import filedialog

def ler_xls(file, movimentacao):
    data = tabula.read_pdf(file, multiple_tables=True, pages='all', stream=True, guess=False)
    df = data[0]

    # REDEFINE O NOME DAS COLUNAS
    if len(df.columns) == 3:
        df.columns = ['A', 'C', 'D']
    else:
        df.columns = ['A', 'B', 'C', 'D']

    # print(df)

    # # PEGA A DATA E NUMERO DA NOTA
    if len(df.columns) == 3:
        idx_data_nota = df[df['C'].str.contains('Data pregão:', na=False)].index
        v = str(df['C'].iloc[idx_data_nota]).strip()
    else:
        idx_data_nota = df[df['D'].str.contains('Data pregão:', na=False)].index
        v = str(df['D'].iloc[idx_data_nota]).strip()

    data_nota = re.search(r'[0-9]{2}.[0-9]{2}.[0-9]{2}', v).group()
    nr_nota = re.search(r'Nota: ([0-9]*)', v).group(True)

    idx_movimentacao = list(df[df['A'].str.contains('1-Bovespa', na=False)].index)

    for idx in idx_movimentacao:
        compra_venda = re.search(r'1-Bovespa\s(\w)+\s', df['A'].iloc[idx]).group(True)
        sigla = re.search(r'^([A-Z0-9]+)', df['C'].iloc[idx]).group(True)

        if len(df.columns) == 3:
            tratamento = df['C'].iloc[idx]
            tratamento = re.sub(sigla, '', tratamento)
            tratamento = re.sub('[A-Z]', '', tratamento)
        else:
            tratamento = df['D'].iloc[idx]
            tratamento = re.sub(sigla, '', tratamento)
            tratamento = re.sub('[A-Z]', '', tratamento)

        # quantidade, valor, valor_total, debito_credito = str(df['C'].iloc[idx]).split()
        quantidade, valor, valor_total = tratamento.split()
        percent =  round((float(valor_total.replace(",", ".").strip())) / 8306, 3)
        corretagem =  round(percent * 30, 2)
        emolumentos =  round(percent * 2.11, 2)
        tx_registro =  round(percent * 3.55, 2)
        tx_liquidacao =  round(percent * 1.9, 2)
        iss =  round(percent * 2.89, 2)

        movimentacao.append({
                "Empresa": '',
                "SIGLA": sigla,
                "NOTA": nr_nota,
                "data compra/venda": data_nota,
                "Quantidade": quantidade,
                "$ compra": valor if compra_venda == 'C' else '',
                "Total compra": valor_total if compra_venda == 'C' else '',
                "$ venda": valor if compra_venda == 'V' else '',
                "Total venda": valor_total if compra_venda == 'V' else '',
                "%": '',
                "corretagem": '',
                "emolumentos": '',
                "tx_registro": '',
                "Taxa de liquidação": '',
                "ISS": '',
                "Total da nota": '',
                "PM Compra": '',
                "PM Venda": '',
                "Total": ''
        })

def web():
    movimentacao = []
    files = []
    root_excel = ""

    # Set up tkinter
    root = tk.Tk()
    root.withdraw()

    # Make folder picker dialog appear on top of other windows
    root.wm_attributes('-topmost', 1)

    clicked = st.button('Selecionar pasta e gerar planilha')
    if clicked:
        dirname = st.text_input('Selected folder:',
                                filedialog.askdirectory(master=root),
                                disabled=True)
        for root, dirs, f in os.walk(dirname):
            root_excel = root
            for name in f:
                print(root + "/" + name)
                files.append(root + "/" + name)
                ler_xls(root + "/" + name, movimentacao)

        movimentacao_df = pandas.DataFrame(data=movimentacao)
        movimentacao_df.sort_values(["data compra/venda", "SIGLA"])
        movimentacao_df.to_excel(root_excel + "/" + 'notas_corretagem.xlsx')
        st.success('Arquivo gerado com sucesso')

if __name__ == '__main__':
    if st._is_running_with_streamlit:
        web()
    else:
        sys.argv = ["streamlit", "run", sys.argv[0]]
        sys.exit(stcli.main())