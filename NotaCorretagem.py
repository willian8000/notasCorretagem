import os
import re
import sys
import pandas
import tabula
import tkinter as tk
import streamlit as st
from streamlit import cli as stcli
from tkinter import filedialog

pandas.set_option('display.max_columns', None)

def ler_xls_xp(file, movimentacao):
    data = tabula.read_pdf(file, multiple_tables=True, pages='all', stream=True, guess=False)
    df = data[0]
    error_nota = False

    if len(df.columns) == 8:
        df['Unnamed: 0'] += ' ' + df['Unnamed: 1']
    if len(df.columns) == 6:
        df['Unnamed: 0'] += ' ' + df['Unnamed: 1']
    if len(df.columns) == 7:
        df['Unnamed: 0'] += ' ' + df['Unnamed: 1']
        # df.drop(columns=['Unnamed: 1', 'Unnamed: 5'], inplace=True)

    # REDEFINE O NOME DAS COLUNAS
    if len(df.columns) == 3:
        df.columns = ['A', 'C', 'D']
    elif len(df.columns) == 4:
        df.columns = ['A', 'B', 'C', 'D']
    elif len(df.columns) == 5:
        df.columns = ['A', 'B', 'C', 'D', 'E']
    elif len(df.columns) == 6:
        df.columns = ['A', 'B', 'C', 'D', 'E', 'F']
    elif len(df.columns) == 7:
        df.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G']
    elif len(df.columns) == 8:
        df.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']

    df.fillna(' ', inplace=True)

    # # PEGA A DATA E NUMERO DA NOTA
    try:
        if len(df.columns) == 3:
            idx_data_nota = df[df['C'].str.contains('Data pregão:', case=False, na=False)].index
            v = str(df['C'].iloc[idx_data_nota]).strip()
        elif len(df.columns) == 5:
            idx_data_nota = df[df['D'].str.contains('Data pregão', case=False, na=False)].index
            v_nota = str(df['C'].iloc[idx_data_nota + 1]).strip()
            v_data = str(df['D'].iloc[idx_data_nota + 1]).strip()
        elif len(df.columns) == 4:
            idx_data_nota = df[df['D'].str.contains('Data pregão:', case=False, na=False)].index
            v = str(df['D'].iloc[idx_data_nota]).strip()
        elif len(df.columns) == 6:
            idx_data_nota = df[df['E'].str.contains('Data pregão', case=False, na=False)].index
            v_nota = str(df['D'].iloc[idx_data_nota + 1]).strip()
            v_data = str(df['E'].iloc[idx_data_nota + 1]).strip()
        elif len(df.columns) == 7:
            idx_data_nota = df[df['F'].str.contains('Data pregão', case=False, na=False)].index
            v_nota = str(df['E'].iloc[idx_data_nota + 1]).strip()
            v_data = str(df['F'].iloc[idx_data_nota + 1]).strip()
        elif len(df.columns) == 8:
            idx_data_nota = df[df['H'].str.contains('Data pregão', case=False, na=False)].index
            v_nota = str(df['F'].iloc[idx_data_nota + 1]).strip()
            v_data = str(df['H'].iloc[idx_data_nota + 1]).strip()
    except AssertionError as error:
        error_nota = True

    # VALORES DE IMPOSTOS
    if not error_nota:
        if len(df.columns) in (3, 4):
            #  PADRAO NOTA INTER 3 E 4 COLUNAS
            data_nota = re.search(r'[0-9]{2}/[0-9]{2}/[0-9]*', v).group()
            nr_nota = re.search(r'Nota: ([0-9]+)', v).group(True)
        elif len(df.columns) == 5:
            # PADRAO NOTA XP
            data_nota = re.search(r'[0-9]{2}/[0-9]{2}/[0-9]*', v_data).group()
            nr_nota = v_nota.split()[1]
            # VALOR TOTAL DA OPERACAO
            idx_vl_operacao = df[df['A'].str.contains('Valor das operações', case=False, na=False)].index
            vl_operacao = str(df['B'].iloc[idx_vl_operacao]).strip()
            vl_operacao = vl_operacao.replace(".", "")
            vl_operacao =  re.search(r'([0-9]+,[0-9]+)', vl_operacao).group(True)
            # EMOLUMENTOS
            idx_emolumentos = df[df['B'].str.contains('Emolumentos', case=False, na=False)].index
            vl_emolumentos = str(df['D'].iloc[idx_emolumentos]).strip()
            vl_emolumentos = vl_emolumentos.replace(".", "")
            vl_emolumentos =  re.search(r'([0-9]+,[0-9]+)', vl_emolumentos).group(True)
            # TX REGISTRO
            idx_tx_registro = df[df['B'].str.contains('Taxa de Registro', case=False, na=False)].index
            vl_tx_registro = str(df['D'].iloc[idx_tx_registro]).strip()
            vl_tx_registro = vl_tx_registro.replace(".", "")
            # print(vl_tx_registro)
            vl_tx_registro =  re.search(r'([0-9]+,[0-9]+)', vl_tx_registro).group(True)
            # TX LIQUIDACAO
            idx_tx_liquidacao = df[df['B'].str.contains('Taxa de liquidação', case=False, na=False)].index
            vl_tx_liquidacao = str(df['D'].iloc[idx_tx_liquidacao]).strip()
            vl_tx_liquidacao = vl_tx_liquidacao.replace(".", "")
            vl_tx_liquidacao =  re.search(r'([0-9]+,[0-9]+)', vl_tx_liquidacao).group(True)
            # CORRETAGEM
            idx_corretagem = df[df['B'].str.contains('Taxa Operacional', case=False, na=False)].index
            vl_corretagem = str(df['D'].iloc[idx_corretagem]).strip()
            vl_corretagem = vl_corretagem.replace(".", "")
            vl_corretagem =  re.search(r'([0-9]+,[0-9]+)', vl_corretagem).group(True)
            # ISS
            idx_iss = df[df['B'].str.contains('Impostos', case=False, na=False)].index
            vl_iss = str(df['D'].iloc[idx_iss]).strip()
            vl_iss = vl_iss.replace(".", "")
            vl_iss =  re.search(r'([0-9]+,[0-9]+)', vl_iss).group(True)
        elif len(df.columns) == 6:
            # PADRAO NOTA XP
            data_nota = re.search(r'[0-9]{2}/[0-9]{2}/[0-9]*', v_data)
            nr_nota = v_nota.split()[1]
            # VALOR TOTAL DA OPERACAO
            idx_vl_operacao = df[df['A'].str.contains('Valor das operações', case=False, na=False)].index
            vl_operacao = str(df['B'].iloc[idx_vl_operacao]).strip()
            vl_operacao = vl_operacao.replace(".", "")
            vl_operacao =  re.search(r'([0-9]+,[0-9]+)', vl_operacao)
            # EMOLUMENTOS
            idx_emolumentos = df[df['B'].str.contains('Emolumentos', case=False, na=False)].index
            vl_emolumentos = str(df['E'].iloc[idx_emolumentos]).strip()
            vl_emolumentos = vl_emolumentos.replace(".", "")
            vl_emolumentos =  re.search(r'([0-9]+,[0-9]+)', vl_emolumentos)
            # TX REGISTRO
            idx_tx_registro = df[df['B'].str.contains('Taxa de Registro', case=False, na=False)].index
            vl_tx_registro = str(df['E'].iloc[idx_tx_registro]).strip()
            vl_tx_registro = vl_tx_registro.replace(".", "")
            vl_tx_registro =  re.search(r'([0-9]+,[0-9]+)', vl_tx_registro)
            # TX LIQUIDACAO
            idx_tx_liquidacao = df[df['B'].str.contains('Taxa de liquidação', case=False, na=False)].index
            vl_tx_liquidacao = str(df['E'].iloc[idx_tx_liquidacao]).strip()
            vl_tx_liquidacao = vl_tx_liquidacao.replace(".", "")
            vl_tx_liquidacao =  re.search(r'([0-9]+,[0-9]+)', vl_tx_liquidacao)
            # CORRETAGEM
            idx_corretagem = df[df['B'].str.contains('Taxa Operacional', case=False, na=False)].index
            vl_corretagem = str(df['E'].iloc[idx_corretagem]).strip()
            vl_corretagem = vl_corretagem.replace(".", "")
            vl_corretagem =  re.search(r'([0-9]+,[0-9]+)', vl_corretagem)
            # ISS
            idx_iss = df[df['B'].str.contains('Impostos', case=False, na=False)].index
            vl_iss = str(df['E'].iloc[idx_iss]).strip()
            vl_iss = vl_iss.replace(".", "")
            vl_iss =  re.search(r'([0-9]+,[0-9]+)', vl_iss)
        elif len(df.columns) == 7:
            # PADRAO NOTA XP
            data_nota = re.search(r'[0-9]{2}/[0-9]{2}/[0-9]*', v_data).group()
            nr_nota = v_nota.split()[1]
            # VALOR TOTAL DA OPERACAO
            vl_operacao = str(df['C'].iloc[30]).strip()
            vl_operacao = vl_operacao.replace(".", "")
            vl_operacao =  re.search(r'^([0-9]+,[0-9]+)', vl_operacao).group(True)
            # print(vl_operacao)
            # EMOLUMENTOS
            idx_emolumentos = df[df['B'].str.contains('Emolumentos', case=False, na=False)].index
            vl_emolumentos = str(df['F'].iloc[idx_emolumentos]).strip()
            vl_emolumentos = vl_emolumentos.replace(".", "")
            vl_emolumentos =  re.search(r'([0-9]+,[0-9]+)', vl_emolumentos).group(True)
            # TX REGISTRO
            idx_tx_registro = df[df['B'].str.contains('Taxa de Registro', case=False, na=False)].index
            vl_tx_registro = str(df['F'].iloc[idx_tx_registro]).strip()
            vl_tx_registro = vl_tx_registro.replace(".", "")
            vl_tx_registro =  re.search(r'([0-9]+,[0-9]+)', vl_tx_registro).group(True)
            # TX LIQUIDACAO
            idx_tx_liquidacao = df[df['B'].str.contains('Taxa de liquidação', case=False, na=False)].index
            vl_tx_liquidacao = str(df['F'].iloc[idx_tx_liquidacao]).strip()
            vl_tx_liquidacao = vl_tx_liquidacao.replace(".", "")
            vl_tx_liquidacao =  re.search(r'([0-9]+,[0-9]+)', vl_tx_liquidacao).group(True)
            # CORRETAGEM
            idx_corretagem = df[df['B'].str.contains('Taxa Operacional', case=False, na=False)].index
            vl_corretagem = str(df['F'].iloc[idx_corretagem]).strip()
            vl_corretagem = vl_corretagem.replace(".", "")
            vl_corretagem =  re.search(r'([0-9]+,[0-9]+)', vl_corretagem).group(True)
            # ISS
            idx_iss = df[df['B'].str.contains('Impostos', case=False, na=False)].index
            vl_iss = str(df['F'].iloc[idx_iss]).strip()
            vl_iss = vl_iss.replace(".", "")
            vl_iss =  re.search(r'([0-9]+,[0-9]+)', vl_iss).group(True)
        elif len(df.columns) == 8:
            # PADRAO NOTA XP
            data_nota = re.search(r'[0-9]{2}/[0-9]{2}/[0-9]*', v_data).group()
            nr_nota = v_nota.split()[1]
            # VALOR TOTAL DA OPERACAO
            vl_operacao = str(df['C'].iloc[27]).strip()
            vl_operacao = vl_operacao.replace(".", "")
            vl_operacao =  re.search(r'^([0-9]+,[0-9]+)', vl_operacao).group(True)
            # print(vl_operacao)
            # EMOLUMENTOS
            # idx_emolumentos = df[df['B'].str.contains('Emolumentos', na=False)].index
            vl_emolumentos = str(df['H'].iloc[28]).strip()
            vl_emolumentos = vl_emolumentos.replace(".", "")
            vl_emolumentos =  re.search(r'([0-9]+,[0-9]+)', vl_emolumentos).group(True)
            # TX REGISTRO
            # idx_tx_registro = df[df['B'].str.contains('Taxa de Registro', na=False)].index
            vl_tx_registro = str(df['H'].iloc[23]).strip()
            vl_tx_registro = vl_tx_registro.replace(".", "")
            vl_tx_registro =  re.search(r'([0-9]+,[0-9]+)', vl_tx_registro).group(True)
            # TX LIQUIDACAO
            # idx_tx_liquidacao = df[df['B'].str.contains('Taxa de liquidação', na=False)].index
            vl_tx_liquidacao = str(df['H'].iloc[22]).strip()
            vl_tx_liquidacao = vl_tx_liquidacao.replace(".", "")
            vl_tx_liquidacao =  re.search(r'([0-9]+,[0-9]+)', vl_tx_liquidacao).group(True)
            # CORRETAGEM
            # idx_corretagem = df[df['B'].str.contains('Taxa Operacional', na=False)].index
            vl_corretagem = str(df['H'].iloc[32]).strip()
            vl_corretagem = vl_corretagem.replace(".", "")
            vl_corretagem =  re.search(r'([0-9]+,[0-9]+)', vl_corretagem).group(True)
            # ISS
            # idx_iss = df[df['B'].str.contains('Impostos', na=False)].index
            vl_iss = str(df['H'].iloc[35]).strip()
            vl_iss = vl_iss.replace(".", "")
            vl_iss =  re.search(r'([0-9]+,[0-9]+)', vl_iss).group(True)


    # MOVIMENTACAO
    idx_movimentacao = list(df[df['A'].str.upper().str.contains('(BOVESPA)', case=False, na=False)].index)

    # SIGLA E VALORES DE COMPRA E VENDA
    for idx in idx_movimentacao:
        if len(df.columns) in (3, 4):
            # PADRAO NOTA INTER 3 E 4 COLUNAS
            compra_venda = re.search(r'BOVESPA\s(\w)+\s', df['A'].iloc[idx], re.IGNORECASE).group(True)
            sigla = re.search(r'^([A-Z0-9]+)', df['C'].iloc[idx]).group(True)
        elif len(df.columns) == 5:
            # PADRAO NOTA XP
            compra_venda = re.search(r'BOVESPA\s(\w)+\s*', df['A'].iloc[idx], re.IGNORECASE).group(True)
            # print(str(df['B'].iloc[idx]))
            sigla = str(df['B'].iloc[idx])
            sigla = re.sub('VISTA', '', sigla, flags=re.IGNORECASE).strip()

            if 'OPCAO DE' in (sigla):
                sigla = re.sub('OPCAO DE VENDA', '', sigla, flags=re.IGNORECASE).strip()
                sigla = re.sub('OPCAO DE COMPRA', '', sigla, flags=re.IGNORECASE).strip()
                sigla = re.sub('^[0-9]+/[0-9]+', '', sigla, flags=re.IGNORECASE).strip()
                sigla = sigla.split()[0]
        elif len(df.columns) == 6:
            # PADRAO NOTA XP
            compra_venda = re.search(r'BOVESPA\s(\w)+\s*', df['A'].iloc[idx], re.IGNORECASE).group(True)
            # print(str(df['B'].iloc[idx]))
            sigla = str(df['B'].iloc[idx])
            sigla = re.sub('VISTA', '', sigla, flags=re.IGNORECASE).strip()
            sigla = re.sub('( NM)', '', sigla, flags=re.IGNORECASE).strip()

            if 'OPCAO DE' in (sigla):
                sigla = re.sub('OPCAO DE VENDA', '', sigla, flags=re.IGNORECASE).strip()
                sigla = re.sub('OPCAO DE COMPRA', '', sigla, flags=re.IGNORECASE).strip()
                sigla = re.sub('^[0-9]+/[0-9]+', '', sigla, flags=re.IGNORECASE).strip()
                sigla = sigla.split()[0]

            sigla = sigla.split()[0]
        elif len(df.columns) in (7, 8):
            # PADRAO NOTA XP
            compra_venda = re.search(r'BOVESPA\s(\w)+\s*', df['A'].iloc[idx], re.IGNORECASE).group(True)
            # print(str(df['C'].iloc[idx]))
            sigla = str(df['C'].iloc[idx])
            sigla = re.sub('VISTA', '', sigla, flags=re.IGNORECASE).strip()

            if 'OPCAO DE' in (sigla):
                sigla = re.sub('OPCAO DE VENDA', '', sigla, flags=re.IGNORECASE).strip()
                sigla = re.sub('OPCAO DE COMPRA', '', sigla, flags=re.IGNORECASE).strip()
                sigla = re.sub('^[0-9]+/[0-9]+', '', sigla, flags=re.IGNORECASE).strip()
                sigla = sigla.split()[0]

        # TRATAMENTO DOS VALORES ENCONTRADOS
        if len(df.columns) in (3, 4):
            tratamento = df['C'].iloc[idx]
            tratamento = re.sub(sigla, '', tratamento)
            tratamento = re.sub('[A-Z]', '', tratamento)
            quantidade, valor, valor_total = tratamento.split()
        elif len(df.columns) == 5:
            quantidade = str(df['B'].iloc[idx]).split()[-1]
            valor = str(df['C'].iloc[idx]).split()[0]
            valor_total = str(df['D'].iloc[idx]).split()[0]
        elif len(df.columns) == 6:
            quantidade = str(df['C'].iloc[idx]).split()[0]
            valor = str(df['D'].iloc[idx]).split()[0]
            print(df['E'])
            valor_total = str(df['E'].iloc[idx]).split()
        elif len(df.columns) == 7:
            quantidade = str(df['D'].iloc[idx]).split()[0]
            valor = str(df['D'].iloc[idx]).split()[1]
            valor_total = str(df['F'].iloc[idx]).split()[0]
        elif len(df.columns) == 8:
            quantidade = str(df['E'].iloc[idx]).split()[0]
            valor = str(df['F'].iloc[idx]).split()[0]
            valor_total = str(df['H'].iloc[idx]).split()[0]

        # SE NAO HOUVE ERRO NA NOTA TENTA CALCULAR OS IMPOSTOS
        if not error_nota:
            print(valor_total)
            percent =  float(valor_total
                             .replace(".", "")
                             .replace(",", ".").strip()) / float(vl_operacao.replace(",", ".").strip())

            corretagem = percent * float(vl_corretagem.replace(",", ".").strip())
            emolumentos = percent * float(vl_emolumentos.replace(",", ".").strip())
            tx_registro = percent * float(vl_tx_registro.replace(",", ".").strip())
            tx_liquidacao = percent * float(vl_tx_liquidacao.replace(",", ".").strip())
            iss = percent * float(vl_iss.replace(",", ".").strip())

        # MONTA DICIONARIO COM OS DADOS EXTRAIDOS DO PDF
        resultado = {
                "Empresa": '',
                "SIGLA": sigla,
                "NOTA": nr_nota,
                "data compra/venda": data_nota,
                "Quantidade": quantidade,
                "$ compra": valor if compra_venda == 'C' else '',
                "Total compra": valor_total if compra_venda == 'C' else '',
                "$ venda": valor if compra_venda == 'V' else '',
                "Total venda": valor_total if compra_venda == 'V' else '',
                "%": percent * 100,
                "corretagem": corretagem,
                "emolumentos": emolumentos,
                "tx_registro": tx_registro,
                "Taxa de liquidação": tx_liquidacao,
                "ISS": iss,
                "Total da nota":  (float(valor_total
                                         .replace(".", "")
                                         .replace(",", ".").strip()) +
                                    corretagem + emolumentos + tx_registro + tx_liquidacao + iss) / int(quantidade
                                                                                                        .replace(".", "")),
                "PM Compra": valor_total if compra_venda == 'C' else '',
                "PM Venda": valor_total if compra_venda == 'V' else '',
                "Total": valor_total if compra_venda == 'V' else ''
        }
        # print(resultado)
        movimentacao.append(resultado)


def ler_xls_easy(file, movimentacao):
    data = tabula.read_pdf(file, multiple_tables=True, pages='all', stream=True, guess=False)
    df = data[0]

    # print("Qt Colunas: " + str(len(df.columns)))

    try:
        # REDEFINE O NOME DAS COLUNAS
        if len(df.columns) == 3:
            df.columns = ['A', 'C', 'D']
        elif len(df.columns) == 4:
            df.columns = ['A', 'B', 'C', 'D']
        elif len(df.columns) == 5:
            df.columns = ['A', 'B', 'C', 'D', 'E']
        elif len(df.columns) == 6:
            df['Unnamed: 0'] += ' ' + df['Unnamed: 1']
            df.columns = ['A', 'B', 'C', 'D', 'E', 'F']

        # print(df)

        # # PEGA A DATA E NUMERO DA NOTA
        if len(df.columns) == 5:
            idx_data_nota = df[df['D'].str.contains('Data pregão', case=False, na=False)].index
            v_nota = str(df['C'].iloc[idx_data_nota + 1]).strip()
            v_data = str(df['D'].iloc[idx_data_nota + 1]).strip()
        elif len(df.columns) == 6:
            idx_data_nota = df[df['F'].str.contains('Data pregão', case=False, na=False)].index
            v_nota = str(df['D'].iloc[idx_data_nota + 1]).strip()
            v_data = str(df['F'].iloc[idx_data_nota + 1]).strip()

        # print(df)
        if len(df.columns) in (3, 4):
            #  PADRAO NOTA INTER 3 E 4 COLUNAS
            data_nota = re.search(r'[0-9]{2}/[0-9]{2}/[0-9]*', v).group()
            nr_nota = re.search(r'Nota: ([0-9]+)', v).group(True)
        elif len(df.columns) == 5:
            data_nota = re.search(r'[0-9]{2}/[0-9]{2}/[0-9]*', v_data).group()
            nr_nota = v_nota.split()[1]
            # VALOR TOTAL DA OPERACAO
            idx_vl_operacao = df[df['A'].str.contains('Valor das operações', case=False, na=False)].index
            vl_operacao = str(df['B'].iloc[idx_vl_operacao]).strip()
            vl_operacao = vl_operacao.replace(".", "")
            vl_operacao =  re.search(r'([0-9]+,[0-9]+)', vl_operacao).group(True)
            # EMOLUMENTOS
            idx_emolumentos = df[df['B'].str.contains('Emolumentos', case=False, na=False)].index
            vl_emolumentos = str(df['D'].iloc[idx_emolumentos]).strip()
            vl_emolumentos = vl_emolumentos.replace(".", "")
            vl_emolumentos =  re.search(r'([0-9]+,[0-9]+)', vl_emolumentos).group(True)
            # TX REGISTRO
            idx_tx_registro = df[df['B'].str.contains('Taxa de Registro', case=False, na=False)].index
            vl_tx_registro = str(df['D'].iloc[idx_tx_registro]).strip()
            vl_tx_registro = vl_tx_registro.replace(".", "")
            # print(vl_tx_registro)
            vl_tx_registro =  re.search(r'([0-9]+,[0-9]+)', vl_tx_registro).group(True)
            # TX LIQUIDACAO
            idx_tx_liquidacao = df[df['B'].str.contains('Taxa de liquidação', case=False, na=False)].index
            vl_tx_liquidacao = str(df['D'].iloc[idx_tx_liquidacao]).strip()
            vl_tx_liquidacao = vl_tx_liquidacao.replace(".", "")
            vl_tx_liquidacao =  re.search(r'([0-9]+,[0-9]+)', vl_tx_liquidacao).group(True)
            # CORRETAGEM
            idx_corretagem = df[df['B'].str.contains('Taxa Operacional', case=False, na=False)].index
            vl_corretagem = str(df['D'].iloc[idx_corretagem]).strip()
            vl_corretagem = vl_corretagem.replace(".", "")
            vl_corretagem =  re.search(r'([0-9]+,[0-9]+)', vl_corretagem).group(True)
            # ISS
            idx_iss = df[df['B'].str.contains('Impostos', case=False, na=False)].index
            vl_iss = str(df['D'].iloc[idx_iss]).strip()
            vl_iss = vl_iss.replace(".", "")
            vl_iss =  re.search(r'([0-9]+,[0-9]+)', vl_iss).group(True)
        elif len(df.columns) == 6:
            # PADRAO NOTA XP
            data_nota = re.search(r'[0-9]{2}/[0-9]{2}/[0-9]*', v_data).group()
            nr_nota = v_nota.split()[1]
            # VALOR TOTAL DA OPERACAO
            idx_vl_operacao = df[df['A'].str.contains('Valor das operações', case=False, na=False)].index
            vl_operacao = str(df['C'].iloc[idx_vl_operacao]).strip()
            vl_operacao = vl_operacao.replace(".", "")
            # print(vl_operacao)

            try:
                vl_operacao =  re.search(r'([0-9]+,[0-9]+)', vl_operacao).group(True)
            except:
                vl_operacao = None

            # EMOLUMENTOS
            idx_emolumentos = df[df['C'].str.contains('Emolumentos', case=False, na=False)].index
            vl_emolumentos = str(df['F'].iloc[idx_emolumentos]).strip()
            vl_emolumentos = vl_emolumentos.replace(".", "")
            vl_emolumentos =  re.search(r'([0-9]+,[0-9]+)', vl_emolumentos).group(True)
            # TX REGISTRO
            idx_tx_registro = df[df['C'].str.contains('Taxa de Registro', case=False, na=False)].index
            vl_tx_registro = str(df['F'].iloc[idx_tx_registro]).strip()
            vl_tx_registro = vl_tx_registro.replace(".", "")
            # print(vl_tx_registro)
            vl_tx_registro =  re.search(r'([0-9]+,[0-9]+)', vl_tx_registro).group(True)
            # TX LIQUIDACAO
            idx_tx_liquidacao = df[df['C'].str.contains('Taxa de liquidação', case=False, na=False)].index
            vl_tx_liquidacao = str(df['F'].iloc[idx_tx_liquidacao]).strip()
            vl_tx_liquidacao = vl_tx_liquidacao.replace(".", "")
            vl_tx_liquidacao =  re.search(r'([0-9]+,[0-9]+)', vl_tx_liquidacao).group(True)
            # CORRETAGEM
            idx_corretagem = df[df['C'].str.contains('Taxa Operacional', case=False, na=False)].index
            vl_corretagem = str(df['F'].iloc[idx_corretagem]).strip()
            vl_corretagem = vl_corretagem.replace(".", "")
            vl_corretagem =  re.search(r'([0-9]+,[0-9]+)', vl_corretagem).group(True)
            # ISS
            idx_iss = df[df['C'].str.contains('Impostos', case=False, na=False)].index
            vl_iss = str(df['F'].iloc[idx_iss]).strip()
            vl_iss = vl_iss.replace(".", "")
            vl_iss =  re.search(r'([0-9]+,[0-9]+)', vl_iss).group(True)

        idx_movimentacao = list(df[df['A'].str.upper().str.contains('(BOVESPA)', case=False, na=False)].index)

        # print(df[df['A'].str.contains('Valor das operações', na=False)].index)
        # valor_operacao =

        # print(data_nota, nr_nota, vl_operacao, vl_emolumentos, vl_tx_registro, vl_tx_liquidacao, vl_corretagem, vl_iss)


        for idx in idx_movimentacao:
            # print('INDEX: ' + str(idx))

            # print(df['A'].iloc[idx])
            # print(df['B'].iloc[idx])
            # print(str(df['C'].iloc[idx]))
            # print(str(df['D'].iloc[idx]).split())

            if len(df.columns) in (3, 4):
                # PADRAO NOTA INTER 3 E 4 COLUNAS
                compra_venda = re.search(r'BOVESPA\s(\w)+\s', df['A'].iloc[idx], re.IGNORECASE).group(True)
                sigla = re.search(r'^([A-Z0-9]+)', df['C'].iloc[idx]).group(True)
            elif len(df.columns) == 5:
                # PADRAO NOTA XP
                compra_venda = re.search(r'BOVESPA\s(\w)+\s*', df['A'].iloc[idx], re.IGNORECASE).group(True)
                # print(str(df['B'].iloc[idx]))
                sigla = str(df['B'].iloc[idx])
                sigla = re.sub('VISTA', '', sigla, flags=re.IGNORECASE).strip()

                if 'OPCAO DE' in (sigla):
                    sigla = re.sub('OPCAO DE VENDA', '', sigla, flags=re.IGNORECASE).strip()
                    sigla = re.sub('OPCAO DE COMPRA', '', sigla, flags=re.IGNORECASE).strip()
                    sigla = re.sub('^[0-9]+/[0-9]+', '', sigla, flags=re.IGNORECASE).strip()
                    sigla = sigla.split()[0]
            elif len(df.columns) == 6:
                compra_venda = re.search(r'(BOVESPA)\s(\w)+\s*', df['A'].iloc[idx], re.IGNORECASE).group(True)
                # print(str(df['B'].iloc[idx]))
                sigla = str(df['C'].iloc[idx])
                sigla = re.sub('VISTA', '', sigla, flags=re.IGNORECASE).strip()
                sigla = re.sub('( NM)', '', sigla, flags=re.IGNORECASE).strip()

                if 'OPCAO DE' in (sigla):
                    sigla = re.sub('OPCAO DE VENDA', '', sigla, flags=re.IGNORECASE).strip()
                    sigla = re.sub('OPCAO DE COMPRA', '', sigla, flags=re.IGNORECASE).strip()
                    sigla = re.sub('^[0-9]+/[0-9]+', '', sigla, flags=re.IGNORECASE).strip()
                    sigla = sigla.split()[0]

                sigla = sigla.split()[0]

            # print(sigla)

            if len(df.columns) in (3, 4):
                tratamento = df['C'].iloc[idx]
                tratamento = re.sub(sigla, '', tratamento)
                tratamento = re.sub('[A-Z]', '', tratamento)
                quantidade, valor, valor_total = tratamento.split()
            elif len(df.columns) == 5:
                quantidade = str(df['B'].iloc[idx]).split()[-1]
                valor = str(df['C'].iloc[idx]).split()[0]
                valor_total = str(df['D'].iloc[idx]).split()[0]
            elif len(df.columns) == 6:
                quantidade = str(df['C'].iloc[idx]).split()[-1]
                valor = str(df['D'].iloc[idx]).split()[0]
                valor_total = str(df['F'].iloc[idx]).split()[0]

            if vl_operacao is not None:
                percent =  float(valor_total
                                 .replace(".", "")
                                 .replace(",", ".").strip()) / float(vl_operacao.replace(",", ".").strip())

                corretagem = percent * float(vl_corretagem.replace(",", ".").strip())
                emolumentos = percent * float(vl_emolumentos.replace(",", ".").strip())
                tx_registro = percent * float(vl_tx_registro.replace(",", ".").strip())
                tx_liquidacao = percent * float(vl_tx_liquidacao.replace(",", ".").strip())
                iss = percent * float(vl_iss.replace(",", ".").strip())
            else:
                percent = ''
                corretagem = ''
                emolumentos = ''
                tx_registro = ''
                tx_liquidacao = ''
                iss = ''

            # print (data_nota, nr_nota, compra_venda, sigla, quantidade, valor, valor_total, vl_operacao,
            #        percent, emolumentos, tx_registro)

            resultado = {
                    "Empresa": '',
                    "SIGLA": sigla,
                    "NOTA": nr_nota,
                    "data compra/venda": data_nota,
                    "Quantidade": quantidade,
                    "$ compra": valor if compra_venda == 'C' else '',
                    "Total compra": valor_total if compra_venda == 'C' else '',
                    "$ venda": valor if compra_venda == 'V' else '',
                    "Total venda": valor_total if compra_venda == 'V' else '',
                    "%": percent * 100,
                    "corretagem": corretagem,
                    "emolumentos": emolumentos,
                    "tx_registro": tx_registro,
                    "Taxa de liquidação": tx_liquidacao,
                    "ISS": iss,
                    "Total da nota":  (float(valor_total
                                             .replace(".", "")
                                             .replace(",", ".").strip()) +
                                        corretagem + emolumentos + tx_registro + tx_liquidacao + iss) / int(quantidade
                                                                                                            .replace(".", "")) if vl_operacao is not None else '',
                    "PM Compra": valor_total if compra_venda == 'C' else '',
                    "PM Venda": valor_total if compra_venda == 'V' else '',
                    "Total": valor_total if compra_venda == 'V' else ''
            }
            # print(resultado)
            movimentacao.append(resultado)
    except:
        print("Erro ao importar!")

def ler_xls_invoice(file, movimentacao):
    data = tabula.read_pdf(file, multiple_tables=True, pages='all', stream=True, guess=False)
    df = data[0]

    # print("Qt Colunas: " + str(len(df.columns)))

    # REDEFINE O NOME DAS COLUNAS
    if len(df.columns) == 3:
        df.columns = ['A', 'C', 'D']
    elif len(df.columns) == 4:
        df.columns = ['A', 'B', 'C', 'D']
    elif len(df.columns) == 6:
        df.columns = ['A', 'B', 'C', 'D', 'E', 'F']
    elif len(df.columns) == 7:
        df.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G']
    elif len(df.columns) == 8:
        df.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']

    # print(df)

    # # PEGA A DATA E NUMERO DA NOTA
    if len(df.columns) == 4:
        idx_data_nota = df[df['C'].str.contains('Data pregão', case=False, na=False)].index
        v_data = str(df['C'].iloc[idx_data_nota + 1]).strip()
        v_nota = str(df['B'].iloc[idx_data_nota + 1]).split('\n')[0]

    # print(df)
    if len(df.columns) == 4:
        #  PADRAO NOTA INTER 3 E 4 COLUNAS
        data_nota = re.search(r'[0-9]{2}/[0-9]{2}/[0-9]*', v_data).group()
        nr_nota = re.search(r'([0-9]+)$', v_nota).group(True)
        # VALOR TOTAL DA OPERACAO

        idx_operacao = df[df['A'].str.contains('^Valor das Operações$', case=False, na=False)].index
        vl_operacao = str(df['B'].iloc[idx_operacao]).strip()
        vl_operacao = vl_operacao.replace(".", "")
        vl_operacao =  re.search(r'([0-9]+,[0-9]+)', vl_operacao).group(True)
        # print(vl_operacao)
        # EMOLUMENTOS
        idx_emolumentos = df[df['B'].str.contains('Emolumentos', case=False, na=False)].index
        vl_emolumentos = str(df['D'].iloc[idx_emolumentos]).strip()
        vl_emolumentos = vl_emolumentos.replace(".", "")
        vl_emolumentos =  re.search(r'([0-9]+,[0-9]+)', vl_emolumentos).group(True)
        # TX REGISTRO
        idx_tx_registro = df[df['B'].str.contains('Taxa de Registro', case=False, na=False)].index
        vl_tx_registro = str(df['D'].iloc[idx_tx_registro]).strip()
        vl_tx_registro = vl_tx_registro.replace(".", "")
        vl_tx_registro =  re.search(r'([0-9]+,[0-9]+)', vl_tx_registro).group(True)
        # TX LIQUIDACAO
        idx_tx_liquidacao = df[df['B'].str.contains('(Taxa de liquidação)$', case=False, na=False)].index
        vl_tx_liquidacao = str(df['D'].iloc[idx_tx_liquidacao]).strip()
        vl_tx_liquidacao = vl_tx_liquidacao.replace(".", "")
        # print(idx_tx_liquidacao, vl_tx_liquidacao)
        vl_tx_liquidacao =  re.search(r'([0-9]+,[0-9]+)', vl_tx_liquidacao).group(True)
        # CORRETAGEM
        idx_corretagem = df[df['B'].str.contains('# - Negócio direto Corretagem', case=False, na=False)].index
        vl_corretagem = str(df['D'].iloc[idx_corretagem]).strip()
        vl_corretagem = vl_corretagem.replace(".", "")
        vl_corretagem =  re.search(r'([0-9]+,[0-9]+)', vl_corretagem).group(True)
        # ISS
        idx_iss = df[df['B'].str.contains('ISS', na=False)].index
        vl_iss = str(df['D'].iloc[idx_iss]).strip()
        vl_iss = vl_iss.replace(".", "")
        vl_iss =  re.search(r'([0-9]+,[0-9]+)', vl_iss).group(True)

    # MOVIMENTACAO
    idx_movimentacao = list(df[df['A'].str.upper().str.contains('(BOVESPA)', case=False, na=False)].index)

    # print(df[df['A'].str.contains('Valor das operações', na=False)].index)
    # valor_operacao =

    print(data_nota, nr_nota, vl_operacao, vl_emolumentos, vl_tx_registro, vl_tx_liquidacao, vl_corretagem, vl_iss)

    try:
        for idx in idx_movimentacao:
            # print('INDEX: ' + str(idx))

            # SIGLA DA MOVIMENTACAO
            if len(df.columns) == 4:
                compra_venda = re.search(r'BOVESPA\s(\w)+\s*', df['A'].iloc[idx], re.IGNORECASE).group(True)
                # print(str(df['B'].iloc[idx]))
                sigla = str(df['B'].iloc[idx])
                sigla = re.sub('VISTA', '', sigla, flags=re.IGNORECASE).strip()
                sigla = re.sub(' CI', '', sigla, flags=re.IGNORECASE).strip()
                sigla = re.sub(' ER', '', sigla, flags=re.IGNORECASE).strip()

                if 'OPCAO DE' in (sigla):
                    sigla = re.sub('OPCAO DE VENDA', '', sigla, flags=re.IGNORECASE).strip()
                    sigla = re.sub('OPCAO DE COMPRA', '', sigla, flags=re.IGNORECASE).strip()
                    sigla = re.sub('^[0-9]+/[0-9]+', '', sigla, flags=re.IGNORECASE).strip()
                    sigla = sigla.split()[0]
                sigla = sigla.split()[0]

            # QUANTIDADE E VALOR
            if len(df.columns) == 4:
                quantidade = str(df['B'].iloc[idx]).split()[-1]
                valor = str(df['C'].iloc[idx]).split()[0]
                valor_total = str(df['D'].iloc[idx]).split()[0]

            percent =  float(valor_total
                             .replace(".", "")
                             .replace(",", ".").strip()) / float(vl_operacao.replace(",", ".").strip())

            corretagem = percent * float(vl_corretagem.replace(",", ".").strip())
            emolumentos = percent * float(vl_emolumentos.replace(",", ".").strip())
            tx_registro = percent * float(vl_tx_registro.replace(",", ".").strip())
            tx_liquidacao = percent * float(vl_tx_liquidacao.replace(",", ".").strip())
            iss = percent * float(vl_iss.replace(",", ".").strip())

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
                    "%": percent * 100,
                    "corretagem": corretagem,
                    "emolumentos": emolumentos,
                    "tx_registro": tx_registro,
                    "Taxa de liquidação": tx_liquidacao,
                    "ISS": iss,
                    "Total da nota":  (float(valor_total
                                             .replace(".", "")
                                             .replace(",", ".").strip()) +
                                        corretagem + emolumentos + tx_registro + tx_liquidacao + iss) / int(quantidade
                                                                                                            .replace(".", "")),
                    "PM Compra": valor_total if compra_venda == 'C' else '',
                    "PM Venda": valor_total if compra_venda == 'V' else '',
                    "Total": valor_total if compra_venda == 'V' else ''
            })
    except:
        print("Erro ao importar!")

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
        for root, directories, files in os.walk(dirname):
            root_excel = root
            for name in (files if os.path.splitext(name)[1] == '.sql' else None):
                print(root + "/" + name)
                files.append(root + "/" + name)
                ler_xls_xp(root + "/" + name, movimentacao)
                # ler_xls_invoice(root + "/" + name, movimentacao)
                # ler_xls_easy(root + "/" + name, movimentacao)

        movimentacao_df = pandas.DataFrame(data=movimentacao)
        movimentacao_df.sort_values(["data compra/venda", "SIGLA"])
        movimentacao_df.to_excel(root_excel + "/" + 'notas_corretagem.xlsx')
        st.success('Arquivo gerado com sucesso')


        for root, directories, files in os.walk(path):
            if os.path.split(root)[1] == 'Types':
                order_list.append((1, os.path.split(root)[1],
                                   [os.path.join(root,  name) for name in files
                                    if os.path.splitext(name)[1] == '.sql']))

if __name__ == '__main__':
    # if st._is_running_with_streamlit:
    #     web()
    # else:
    #     sys.argv = ["streamlit", "run", sys.argv[0]]
    #     sys.exit(stcli.main())

    from unidecode import unidecode
    import warnings
    warnings.filterwarnings("ignore")

    def get_PDF(dirname):
        list_pdf = []
        for root, directories, files in os.walk(dirname):
            root_excel = root
            for name in [name for name in files if os.path.splitext(name)[1] == '.pdf']:
                list_pdf.append(root + "/" + name)
        return list_pdf, root_excel

    def reading_pdf(list_pdf, root):
        movimetacoes = []
        for file in list_pdf:
            print(file)
            data = tabula.read_pdf(file, multiple_tables=True, pages='all', stream=True, guess=False)
            df = data[0]
            df.fillna(' ', inplace=True)

            for i, column in enumerate(df.columns):
                if i == 0:
                    anchor = column
                else:
                    df[anchor] += ' ' + df[column]
                    df.drop(columns=column, inplace=True)

            df.columns = ['A']
            df['A'] = df['A'].apply(unidecode)

            idx_movimentacao = list(df[df['A'].str.upper().str.contains('(BOVESPA) [C|V]+', case=False, na=False, flags=re.IGNORECASE, regex=True)].index)

            if not df[df['A'].str.upper().str.contains('(nota folha).*(data pregao)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index.empty:
                idx = df[df['A'].str.upper().str.contains('(nota folha).*(data pregao)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index[0] + 1
            elif not df[df['A'].str.upper().str.contains('(data pregao).*(nota)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index.empty:
                idx = df[df['A'].str.upper().str.contains('(data pregao).*(nota)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index[0]
            elif not df[df['A'].str.upper().str.contains('(numero da nota).*(data pregao)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index.empty:
                idx = df[df['A'].str.upper().str.contains('(numero da nota).*(data pregao)', case=True, na=False, flags=re.IGNORECASE, regex=True)].index[0] + 1

            data_nota, nr_nota = get_nota_data(str(df.iloc[idx]['A']))
            movimetacoes += get_movimentacao(df.iloc[idx_movimentacao]['A'], data_nota, nr_nota)

        movimetacoes = pandas.DataFrame(movimetacoes).sort_values(["data compra/venda", "SIGLA"])
        print('   Gerando arquivo..')
        movimetacoes.to_excel(root + "/" + 'notas_corretagem.xlsx', index=False)


    def get_movimentacao(movimentacoes, data_nota, nr_nota):
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


            quantidade = float(quantidade.replace(",", ".").strip())
            valor_unidade = float(valor_unidade.replace(",", ".").strip())
            valor_total_mov = float(valor_total_mov.replace(",", ".").strip())
            valor_operacaoes += valor_total_mov

            # print(data_nota, nr_nota, compra_venda, sigla, quantidade, valor_unidade, valor_total_mov)
            # valores_movimentacoes.append((data_nota, nr_nota, compra_venda, sigla, quantidade, valor_unidade, valor_total_mov))

            resultado = {
                "Empresa": '',
                "SIGLA": sigla,
                "NOTA": nr_nota,
                "data compra/venda": data_nota,
                "Quantidade": quantidade,
                "$ compra": valor_unidade if compra_venda == 'C' else '',
                "Total compra": valor_total_mov if compra_venda == 'C' else '',
                "$ venda": valor_unidade if compra_venda == 'V' else '',
                "Total venda": valor_total_mov if compra_venda == 'V' else '',
                "%": 0,
                "corretagem": '',
                "emolumentos": '',
                "tx_registro": '',
                "Taxa de liquidação": '',
                "ISS": '',
                "Total da nota": '',
                "PM Compra": valor_total_mov if compra_venda == 'C' else '',
                "PM Venda": valor_total_mov if compra_venda == 'V' else '',
                "Total": valor_total_mov if compra_venda == 'V' else ''
            }

            valores_movimentacoes.append(resultado)

        # ATUALIZA A PORCENTAGEM
        for i, mov in enumerate(valores_movimentacoes):
            valores_movimentacoes[i]["%"] = 100 * (float(valores_movimentacoes[i]['Total compra'] or valores_movimentacoes[i]['Total venda']) / valor_operacaoes)
            valores_movimentacoes[i]["data compra/venda"] = pandas.to_datetime(valores_movimentacoes[i]["data compra/venda"], format='%d/%m/%Y')
        return valores_movimentacoes


    def remove_words(v):
        for word in ['pregao', 'data', 'nota', 'no', '\:', 'folha', '\|', '1-BOVESPA', 'BOVESPA', '(NM)', 'VISTA',
                     '[C|D]$', 'opcao de [a-z]+', '[a-z]*#', '(VVAR)(?![A-Z0-9])', 'ON/s*[0-9]+,[0-9]+', '/sD/s',
                     'ON/s', '[0-9]{2}/[0-9]{2}']:
            v = re.sub(word, '', v, flags=re.IGNORECASE).strip()
        return v

    def get_nota_data(v):
        # print(v)
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

    minhas_pastas = [
        r"C:\Users\w.a.alves.da.silva\OneDrive - Accenture\Desktop\NotaCorretagem\Notas\notasdaclear",
        r"C:\Users\w.a.alves.da.silva\OneDrive - Accenture\Desktop\NotaCorretagem\Notas\notasinter",
        r"C:\Users\w.a.alves.da.silva\OneDrive - Accenture\Desktop\NotaCorretagem\Notas\notasxp",
        r"C:\Users\w.a.alves.da.silva\OneDrive - Accenture\Desktop\NotaCorretagem\Notas\notaseasyinvest"
    ]

    for pasta in minhas_pastas:
        movimetacoes_pdf = []
        mylist, root = get_PDF(pasta)
        reading_pdf(mylist, root)