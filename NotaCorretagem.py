import os
import re
import sys
import pandas
import tabula
import tkinter as tk
import streamlit as st
from streamlit import cli as stcli
from tkinter import filedialog
from unidecode import unidecode
import warnings

warnings.filterwarnings("ignore")
pandas.set_option('display.max_columns', None)

def get_all_PDF(dirname):
    list_pdf = []
    for root, directories, files in os.walk(dirname):
        root_excel = root
        for name in [name for name in files if os.path.splitext(name)[1] == '.pdf']:
            list_pdf.append(root + "/" + name)
    return list_pdf, root_excel

def get_notas_by_page(data):
    nota_paginas = {}
    for num_pagina in range(0, len(data)):
        df = data[num_pagina]

        print(df)

        if not df[df['A'].str.contains('(nota folha).*(data pregao)', case=True, na=False, flags=re.IGNORECASE,
                                       regex=True)].index.empty:
            idx = df[df['A'].str.contains('(nota folha).*(data pregao)', case=True, na=False, flags=re.IGNORECASE,
                                          regex=True)].index[0] + 1
        elif not df[df['A'].str.contains('(data pregao).*(nota)', case=True, na=False, flags=re.IGNORECASE,
                                         regex=True)].index.empty:
            idx = df[df['A'].str.contains('(data pregao).*(nota)', case=True, na=False, flags=re.IGNORECASE,
                                          regex=True)].index[0]
        elif not df[df['A'].str.contains('(numero da nota).*(data pregao)', case=True, na=False, flags=re.IGNORECASE,
                                         regex=True)].index.empty:
            idx = df[df['A'].str.contains('(numero da nota).*(data pregao)', case=True, na=False, flags=re.IGNORECASE,
                                          regex=True)].index[0] + 1
        elif not df[df['A'].str.contains('([0-9]*\.?[0-9]*)\s\d\s([0-9]+\/[0-9]+\/[0-9]+)', case=True, na=False, flags=re.IGNORECASE,
                                         regex=True)].index.empty:
            idx = df[df['A'].str.contains('([0-9]*\.?[0-9]*)\s\d\s([0-9]+\/[0-9]+\/[0-9]+)', case=True, na=False, flags=re.IGNORECASE,
                                          regex=True)].index[0]

        print(idx)
        data_nota, nr_nota = get_nota_data(str(df.iloc[idx]['A']))

        if nr_nota not in nota_paginas:
            nota_paginas[nr_nota] = {"data_nota": data_nota, "paginas": [num_pagina]}
        else:
            nota_paginas[nr_nota]["paginas"].append(num_pagina)

    return nota_paginas

def normalizar_dataframe(data, root, debug=False):
    data_out = []
    for num_pagina in range(0, len(data)):
        df = data[num_pagina]

        df.fillna(' ', inplace=True)

        for i, column in enumerate(df.columns):
            if i == 0:
                anchor = column
            else:
                df[anchor] += ' ' + df[column]
                df.drop(columns=column, inplace=True)

        df.columns = ['A']
        df['A'] = df['A'].apply(unidecode)

        data_out.append(df)

        if debug:
            print(df)
            df.to_csv(rf'{root}\pandas.txt', header=None, index=None, sep=' ', mode='a')

    return data_out

def reading_pdf(list_pdf, root):
    movimetacoes = []
    for file in list_pdf:
        print(file)
        data = tabula.read_pdf(file, multiple_tables=True, pages='all', stream=True, guess=False)

        data = normalizar_dataframe(data, root)
        notas_por_pagina = get_notas_by_page(data)


        for nota in notas_por_pagina:
            nr_nota = nota
            data_nota = notas_por_pagina[nota]["data_nota"]

            print(nr_nota)
            idx_movimentacao = []
            df_movimentacao = []

            for pagina in notas_por_pagina[nota]["paginas"]:
                df = data[pagina]

                # EMOLULMENTOS
                if not df[df['A'].str.contains('(diversas emolumentos)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index.empty:
                    idx_emolumentos = df[df['A'].str.contains('(diversas emolumentos)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index[0]
                elif not df[df['A'].str.contains('(emolumentos\s*\-*[0-9]+,[0-9]+\s+D$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index.empty:
                    idx_emolumentos = df[df['A'].str.contains('(emolumentos\s*\-*[0-9]+,[0-9]+\s+D$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index[0]

                try:
                    emolumentos = get_emolumentos(str(df.iloc[idx_emolumentos]['A']))
                except:
                    emolumentos = None

                # TAXA DE REGISTRO
                if not df[df['A'].str.contains('(registro\s*\-*[0-9]+,[0-9]+\s+D$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index.empty:
                    idx_taxa_registro = df[df['A'].str.contains('(registro\s*\-*[0-9]+,[0-9]+\s+D$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index[0]
                elif not df[df['A'].str.contains('(registro\(3\)\s*\-*[0-9]+,[0-9]+\s+D$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index.empty:
                    idx_taxa_registro = df[df['A'].str.contains('(registro\(3\)\s*\-*[0-9]+,[0-9]+\s+D$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index[0]
                elif not df[df['A'].str.contains('(registro\s*\-*[0-9]+,[0-9]+$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index.empty:
                    idx_taxa_registro = df[df['A'].str.contains('(registro\s*\-*[0-9]+,[0-9]+$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index[0]

                try:
                    taxa_registro = get_taxa_registro(str(df.iloc[idx_taxa_registro]['A']))
                except:
                    taxa_registro = None

                # TAXA DE LIQUIDACAO
                if not df[df['A'].str.contains('(taxa de liquidacao\s*\-*[0-9]+,[0-9]+\s+D$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index.empty:
                    idx_taxa_liquidacao = df[df['A'].str.contains('(taxa de liquidacao\s*\-*[0-9]+,[0-9]+\s+D$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index[0]
                elif not df[df['A'].str.contains('(taxa de liquidacao\(2\)\s*\-*[0-9]+,[0-9]+\s+D$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index.empty:
                    idx_taxa_liquidacao = df[df['A'].str.contains('(taxa de liquidacao\(2\)\s*\-*[0-9]+,[0-9]+\s+D$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index[0]
                elif not df[df['A'].str.contains('(taxa de liquidacao\s*\-*[0-9]+,[0-9]+$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index.empty:
                    idx_taxa_liquidacao = df[df['A'].str.contains('(taxa de liquidacao\s*\-*[0-9]+,[0-9]+$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index[0]

                try:
                    taxa_liquidacao = get_taxa_liquidacao(str(df.iloc[idx_taxa_liquidacao]['A']))
                except:
                    taxa_liquidacao = None

                # ISS
                if not df[df['A'].str.contains('(ISS\s*\(SAO\s*PAULO\))', case=True, na=False, flags=re.IGNORECASE, regex=True)].index.empty:
                    idx_iss = df[df['A'].str.contains('(ISS\s*\(SAO\s*PAULO\))', case=True, na=False, flags=re.IGNORECASE, regex=True)].index[0]
                elif not df[df['A'].str.contains('(impostos\s*\-*[0-9]+,[0-9]+\s*$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index.empty:
                    idx_iss = df[df['A'].str.contains('(impostos\s*\-*[0-9]+,[0-9]+\s*$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index[0]
                elif not df[df['A'].str.contains('(impostos\s*\-*[0-9]+,[0-9]+\s+D$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index.empty:
                    idx_iss = df[df['A'].str.contains('(impostos\s*\-*[0-9]+,[0-9]+\s+D$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index[0]
                elif not df[df['A'].str.contains('(ISS\s*\-*[0-9]+,[0-9]+\s+D$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index.empty:
                    idx_iss = df[df['A'].str.contains('(ISS\s*\-*[0-9]+,[0-9]+\s+D$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index[0]

                try:
                    iss = get_iss(str(df.iloc[idx_iss]['A']))
                except:
                    iss = None

                # CORRETAGEM
                if not df[df['A'].str.contains('(corretagem\s*\-*[0-9]+,[0-9]+$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index.empty:
                    idx_corretagem = df[df['A'].str.contains('(corretagem\s*\-*[0-9]+,[0-9]+$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index[0]
                elif not df[df['A'].str.contains('(taxa\soperacional\s*\-*[0-9]+,[0-9]+\s+D$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index.empty:
                    idx_corretagem = df[df['A'].str.contains('(taxa\soperacional\s*\-*[0-9]+,[0-9]+\s+D$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index[0]
                elif not df[df['A'].str.contains('(corretagem\s*\-*[0-9]+,[0-9]+\s+D$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index.empty:
                    idx_corretagem = df[df['A'].str.contains('(corretagem\s*\-*[0-9]+,[0-9]+\s+D$)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index[0]

                try:
                    corretagem = get_corretagem(str(df.iloc[idx_corretagem]['A']))
                except:
                    corretagem = None

                idx_movimentacao = list(df[df['A'].str.contains('(BOVESPA) [C|V]+', case=False, na=False, flags=re.IGNORECASE, regex=True)].index)
                df_movimentacao += list(df.iloc[idx_movimentacao]['A'])

            movimetacoes += get_movimentacao(df_movimentacao, data_nota, nr_nota, emolumentos, taxa_registro, taxa_liquidacao, iss, corretagem)

    movimetacoes = pandas.DataFrame(movimetacoes).sort_values(["data compra/venda", "SIGLA"])
    print('   Gerando arquivo..')
    movimetacoes.to_excel(root + "/" + 'notas_corretagem.xlsx', index=False, sheet_name="NOTAS")

def get_movimentacao(movimentacoes, data_nota, nr_nota, emolumentos, taxa_registro, taxa_liquidacao, iss, corretagem):
    valores_movimentacoes = []
    valor_operacaoes = 0

    for i, movimentacao in enumerate(movimentacoes):
        # print(i, movimentacao)
        for word in ['pregao', 'data', 'nota', 'no', '\:', 'folha', '\|', '1-BOVESPA', 'BOVESPA', '(NM)', 'VISTA',
                     '[C|D]$', 'opcao de [a-z]+', '[a-z]*#', '(VVAR)(?![A-Z0-9])', 'ON/s+[0-9]+,[0-9]+', '/sD/s*',
                     'ON/s', '[0-9]{2}/[0-9]{2}', '\.', 'TERMO', 'FRACIONARIO']:
            movimentacao = re.sub(word, '', movimentacao, flags=re.IGNORECASE).strip()

        # print(movimentacao)

        # COMPRA OU VENDA
        if match := re.search('^([C|V])', movimentacao, re.IGNORECASE):
            compra_venda = match.group(1).strip()
            movimentacao = re.sub('^' + compra_venda, '', movimentacao, flags=re.IGNORECASE).strip()

        # VALOR TOTAL
        if match := re.search('([0-9]+,[0-9]+)$', movimentacao, re.IGNORECASE):
            valor_total_mov = match.group(1).strip()
            movimentacao = re.sub(valor_total_mov + '$', '', movimentacao, flags=re.IGNORECASE).strip()

        # VALOR UNIDADE
        if match := re.search('([0-9]+,[0-9]+)$', movimentacao, re.IGNORECASE):
            valor_unidade = match.group(1).strip()
            movimentacao = re.sub(valor_unidade + '$', '', movimentacao, flags=re.IGNORECASE).strip()

        # QUANTIDADE
        if match := re.search('([0-9]+)$', movimentacao, re.IGNORECASE):
            quantidade = match.group(1).strip()
            movimentacao = re.sub(quantidade, '', movimentacao, flags=re.IGNORECASE).strip()

        # SIGLA
        sigla = movimentacao
        for word in ['[0-9]+,[0-9]+', 'ON/s+', '/sD$']:
            sigla = re.sub(word, '', sigla, flags=re.IGNORECASE).strip()

        quantidade = float(str(quantidade).replace(",", ".").strip())
        valor_unidade = float(str(valor_unidade).replace(",", ".").strip())
        valor_total_mov = float(str(valor_total_mov).replace(",", ".").strip())
        emolumentos = float(str(emolumentos).replace(",", ".").strip()) if emolumentos is not None else None
        taxa_registro = float(str(taxa_registro).replace(",", ".").strip()) if taxa_registro is not None else None
        taxa_liquidacao = float(str(taxa_liquidacao).replace(",", ".").strip()) if taxa_liquidacao is not None else None
        iss = float(str(iss).replace(",", ".").strip()) if iss is not None else None
        corretagem = float(str(corretagem).replace(",", ".").strip()) if corretagem is not None else None
        valor_operacaoes += valor_total_mov

        print(data_nota, nr_nota, compra_venda, sigla, quantidade, valor_unidade, valor_total_mov, valor_operacaoes,
              emolumentos, taxa_registro, taxa_liquidacao, iss, corretagem)

        resultado = {
            "Empresa": None,
            "SIGLA": sigla,
            "NOTA": nr_nota,
            "data compra/venda": data_nota,
            "Quantidade": quantidade,
            "$ compra": valor_unidade if compra_venda == 'C' else None,
            "Total compra": valor_total_mov if compra_venda == 'C' else None,
            "$ venda": valor_unidade if compra_venda == 'V' else None,
            "Total venda": valor_total_mov if compra_venda == 'V' else None,
            "%": 0,
            "corretagem": corretagem,
            "emolumentos": emolumentos,
            "tx_registro": taxa_registro,
            "Taxa de liquidação": taxa_liquidacao,
            "ISS": iss,
            "Total da nota": 0,
            "PM Compra": valor_total_mov if compra_venda == 'C' else None,
            "PM Venda": valor_total_mov if compra_venda == 'V' else None,
            "Total": valor_total_mov if compra_venda == 'V' else None
        }

        valores_movimentacoes.append(resultado)

    # ATUALIZA A PORCENTAGEM E OS IMPOSTOS
    for i, mov in enumerate(valores_movimentacoes):
        valores_movimentacoes[i]["%"] = 100 * (float(valores_movimentacoes[i]['Total compra'] or valores_movimentacoes[i]['Total venda']) / valor_operacaoes)
        valores_movimentacoes[i]["data compra/venda"] = pandas.to_datetime(valores_movimentacoes[i]["data compra/venda"], format='%d/%m/%Y')
        valores_movimentacoes[i]["emolumentos"] = valores_movimentacoes[i]["emolumentos"] * (valores_movimentacoes[i]["%"] / 100)
        valores_movimentacoes[i]["tx_registro"] = valores_movimentacoes[i]["tx_registro"] * (valores_movimentacoes[i]["%"] / 100)
        valores_movimentacoes[i]["Taxa de liquidação"] = valores_movimentacoes[i]["Taxa de liquidação"] * (valores_movimentacoes[i]["%"] / 100)
        valores_movimentacoes[i]["ISS"] = valores_movimentacoes[i]["ISS"] * (valores_movimentacoes[i]["%"] / 100)
        valores_movimentacoes[i]["corretagem"] = valores_movimentacoes[i]["corretagem"] * (valores_movimentacoes[i]["%"] / 100)
        valores_movimentacoes[i]["Total da nota"] = (
                float(valores_movimentacoes[i]['Total compra'] or valores_movimentacoes[i]['Total venda']) +
                valores_movimentacoes[i]["corretagem"] +
                valores_movimentacoes[i]["emolumentos"] +
                valores_movimentacoes[i]["tx_registro"] +
                valores_movimentacoes[i]["Taxa de liquidação"] +
                valores_movimentacoes[i]["ISS"]
        )
        valores_movimentacoes[i]["PM Compra"] = valores_movimentacoes[i]["Total da nota"] / valores_movimentacoes[i]["Quantidade"] if compra_venda == 'C' else None
        valores_movimentacoes[i]["PM Venda"] = valores_movimentacoes[i]["Total da nota"] / valores_movimentacoes[i]["Quantidade"] if compra_venda == 'V' else None
    return valores_movimentacoes

def remove_words(v):
    for word in ['pregao', 'data', 'nota', 'no', '\:', 'folha', '\|', '1-BOVESPA', 'BOVESPA', '(NM)', 'VISTA',
                 '[C|D]$', 'opcao de [a-z]+', '[a-z]*#', '(VVAR)(?![A-Z0-9])', 'ON/s*[0-9]+,[0-9]+', '/sD/s',
                 'ON/s', '[0-9]{2}/[0-9]{2}']:
        v = re.sub(word, '', v, flags=re.IGNORECASE).strip()
    return v

def get_emolumentos(v):
    if match := re.search('emolumentos\s*\-*([0-9]+,[0-9]+)$', v, flags=re.IGNORECASE):
        emolumentos = match.group(1)
    if match := re.search('emolumentos\s*\-*([0-9]+,[0-9]+)\s+D$', v, flags=re.IGNORECASE):
        emolumentos = match.group(1)
    return emolumentos

def get_taxa_registro(v):
    if match := re.search('registro\(3\)\s*\-*([0-9]+,[0-9]+)\s+D$', v, flags=re.IGNORECASE):
        taxa_registro = match.group(1)
    if match := re.search('registro\s*\-*([0-9]+,[0-9]+)\s+D$', v, flags=re.IGNORECASE):
        taxa_registro = match.group(1)
    if match := re.search('registro\s*\-*([0-9]+,[0-9]+)$', v, flags=re.IGNORECASE):
        taxa_registro = match.group(1)
    return taxa_registro

def get_taxa_liquidacao(v):
    if match := re.search('taxa de liquidacao\s*\-*([0-9]+,[0-9]+)\s+D$', v, flags=re.IGNORECASE):
        taxa_liquidacao = match.group(1)
    if match := re.search('taxa de liquidacao\(2\)\s*\-*([0-9]+,[0-9]+)\s+D$', v, flags=re.IGNORECASE):
        taxa_liquidacao = match.group(1)
    if match := re.search('taxa de liquidacao\s*\-*([0-9]+,[0-9]+)$', v, flags=re.IGNORECASE):
        taxa_liquidacao = match.group(1)
    return taxa_liquidacao

def get_iss(v):
    if match := re.search('([0-9]+,[0-9]+)$', v, flags=re.IGNORECASE):
        iss = match.group(1)
    if match := re.search('impostos\s*\-*([0-9]+,[0-9]+)\s*$', v, flags=re.IGNORECASE):
        iss = match.group(1)
    if match := re.search('impostos\s*\-*([0-9]+,[0-9]+)\s+D$', v, flags=re.IGNORECASE):
        iss = match.group(1)
    if match := re.search('ISS\s*\-*([0-9]+,[0-9]+)\s+D$', v, flags=re.IGNORECASE):
        iss = match.group(1)

    return iss

def get_corretagem(v):
    if match := re.search('([0-9]+,[0-9]+)$', v, flags=re.IGNORECASE):
        corretagem = match.group(1)
    if match := re.search('([0-9]+,[0-9]+)\s+D$', v, flags=re.IGNORECASE):
        corretagem = match.group(1)

    return corretagem

def get_nota_data(v):
    print(v)
    for word in ['pregao', 'data', 'nota', 'no', '\:', 'folha', '\|', '1-BOVESPA', 'BOVESPA', '(NM)', 'VISTA',
                 '[C|D]$', 'opcao de [a-z]+', '[a-z]*#', '(VVAR)(?![A-Z0-9])', 'ON/s*[0-9]+,[0-9]+', '/sD/s']:
        v = re.sub(word, '', v, flags=re.IGNORECASE).strip()

    if match := re.search('([0-9]+/[0-9]+/[0-9]+)', v, flags=re.IGNORECASE):
        data_nota = match.group(1)

    if match := re.search('BR ([0-9]+)', v, re.IGNORECASE):
        nr_nota = match.group(1).strip().split()[0]
    elif match := re.sub('([0-9]+/[0-9]+/[0-9]+)', '', v, re.IGNORECASE):
        nr_nota = match.strip().split()[0]
    return data_nota, nr_nota


def web():
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
        mylist, root = get_all_PDF(dirname)
        reading_pdf(mylist, root)
        st.success('Arquivo gerado com sucesso')


if __name__ == '__main__':
    if st._is_running_with_streamlit:
        web()
    else:
        sys.argv = ["streamlit", "run", sys.argv[0]]
        sys.exit(stcli.main())