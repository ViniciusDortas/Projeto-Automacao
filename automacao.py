import pandas as pd
import numpy as np
import os

# ============
# CONSTANTES
# ============

ARQUIVO_VENDAS = r"Bases de Dados/Vendas.xlsx"
ARQUIVO_EMAILS = r"Bases de Dados/Emails.csv"
ARQUIVO_LOJAS = r"Bases de Dados/Lojas.csv"

FAT_META_DIA = 1000
FAT_META_ANO = 1_650_000
VAR_META_DIA = 4
VAR_META_ANO = 120
TICKET_META = 500

# ==========
# FUNÇÕES
# ==========

def carregar_dados(vendas, emails, lojas):
    df_vendas = pd.read_excel(vendas)
    df_vendas["Data"] = pd.to_datetime(df_vendas["Data"])
    df_vendas["Ano"] = df_vendas["Data"].dt.year

    df_emails = pd.read_csv(emails)
    df_lojas = pd.read_csv(lojas, sep=";")

    df_emails_lojas = df_emails.merge(
        df_lojas,
        on="Loja",
        how="left"
    )
    df_emails_lojas["ID Loja"] = df_emails_lojas["ID Loja"].astype("Int64")

    return df_vendas, df_lojas, df_emails_lojas


def calcular_faturamento(df):
    faturamento_dia = (
        df.groupby(["ID Loja", "Data"], as_index=False)["Valor Final"]
        .sum()
        .rename(columns={"Valor Final":"Faturamento Diario"})
    )

    faturamento_ano = (
        df.groupby(["ID Loja", "Ano"], as_index=False)["Valor Final"]
        .sum()
        .rename(columns={"Valor Final":"Faturamento Anual"})
    )

    return faturamento_dia, faturamento_ano
    

def calcular_variedade(df):
    variedade_itens_dia = (
        df.groupby(["ID Loja", "Data"], as_index=False)["Produto"]
        .nunique()
        .rename(columns={"Produto":"Variedade Diaria"})
    )

    variedade_itens_ano = (
        df.groupby(["ID Loja", "Ano"], as_index=False)["Produto"]
        .nunique()
        .rename(columns={"Produto":"Variedade Anual"})
    )

    return variedade_itens_dia, variedade_itens_ano


def calcular_ticket_medio(df):
    ticket_medio_dia = (
        df.groupby(["ID Loja", "Data"], as_index=False)[["Quantidade", "Valor Final"]]
        .sum()
        .assign(
            Ticket_Medio_Dia=lambda x: (
                x["Valor Final"] / x["Quantidade"].replace(0, pd.NA)
            ).round(2)
        )
        .drop(columns=["Valor Final", "Quantidade"])
        .rename(columns=lambda x: x.replace("_", " "))
    )   

    ticket_medio_ano = (    
        df.groupby(["ID Loja", "Ano"], as_index=False)[["Quantidade", "Valor Final"]]
        .sum()
        .assign(
            Ticket_Medio_Ano=lambda x: (
                x["Valor Final"] / x["Quantidade"].replace(0, pd.NA)
            ).round(2)
        )
        .drop(columns=["Valor Final", "Quantidade"])
        .rename(columns=lambda x: x.replace("_", " "))
    )

    return ticket_medio_dia, ticket_medio_ano


def verificar_meta(df, valor, meta):
    df = df.copy()
    df["Bateu Meta"] = np.where(df[valor] >= meta, "Sim", "Não")
    return df


def aplicar_metas(lista_dfs):
    dfs_novos = [
        verificar_meta(df, valor, meta)
        for df, valor, meta in lista_dfs
    ]

    return dfs_novos


def data_atual(df):

    total_lojas = df["ID Loja"].nunique()

    data_mais_recente = (
        df
        .groupby("Data")["ID Loja"]
        .nunique()
        .loc[lambda x: x == total_lojas]
        .index
        .max()
    )
    ano_atual = data_mais_recente.year

    return data_mais_recente, ano_atual


def filtro_data(df, coluna, data):
    return df[df[coluna] == data]


def filtrar_data_dfs(df_fat, df_var, df_ticket, coluna, data):

    lista_dfs_dia = (
        df_fat,
        df_var,
        df_ticket,
    )

    return [
        filtro_data(df, coluna, data).set_index("ID Loja")
        for df in lista_dfs_dia
    ]


def loc_onepage(df, id_loja, coluna):
    if id_loja in df.index:
        return df.at[id_loja, coluna]
    return 0


def gerar_onepage_loja(
    id_loja,
    faturamento_dia_fil,
    variedade_dia_fil,
    ticket_dia_fil,
    faturamento_ano_fil,
    variedade_ano_fil,
    ticket_ano_fil,
):
    
    onepage_dia = pd.DataFrame({
    "Indicador": [
        "Faturamento",
        "Variedade",
        "Ticket Medio"
    ],
    "Valor Dia": [
        loc_onepage(faturamento_dia_fil, id_loja, "Faturamento Diario"),
        loc_onepage(variedade_dia_fil, id_loja, "Variedade Diaria"),
        loc_onepage(ticket_dia_fil, id_loja, "Ticket Medio Dia")
    ],
    "Meta Dia": [
        FAT_META_DIA,
        VAR_META_DIA,
        TICKET_META
    ],
    "Cenário Dia": [
        loc_onepage(faturamento_dia_fil, id_loja, "Bateu Meta"),
        loc_onepage(variedade_dia_fil, id_loja, "Bateu Meta"),
        loc_onepage(ticket_dia_fil, id_loja, "Bateu Meta")
    ]       
    }).set_index("Indicador")

    onepage_ano = pd.DataFrame({
    "Indicador": [
        "Faturamento",
        "Variedade",
        "Ticket Medio"
    ],
    "Valor Ano": [
        loc_onepage(faturamento_ano_fil, id_loja, "Faturamento Anual"),
        loc_onepage(variedade_ano_fil, id_loja, "Variedade Anual"),
        loc_onepage(ticket_ano_fil, id_loja, "Ticket Medio Ano")
    ],
    "Meta Ano": [
        FAT_META_ANO,
        VAR_META_ANO,
        TICKET_META
    ],
    "Cenário Ano": [
        loc_onepage(faturamento_ano_fil, id_loja, "Bateu Meta"),
        loc_onepage(variedade_ano_fil, id_loja, "Bateu Meta"),
        loc_onepage(ticket_ano_fil, id_loja, "Bateu Meta")
    ]       
    }).set_index("Indicador")


    def bolinha(valor):
        if valor == "Sim":
            return '<span class="verde"></span>'
        else:
            return '<span class="vermelha"></span>'
        
    onepage_dia["Cenário Dia"] = onepage_dia["Cenário Dia"].apply(bolinha)
    onepage_ano["Cenário Ano"] = onepage_ano["Cenário Ano"].apply(bolinha)

 
    html_dia = onepage_dia.to_html(escape=False)
    html_ano = onepage_ano.to_html(escape=False)

    corpo_html = f"""
    <html>
        <head>
            <style>

                table {{
                    border-collapse: collapse;
                    width: 60%;
                }}

                th, td {{
                    border: 1px solid #ddd;
                    padding: 8px;
                    text-align: center;
                }}

                th {{
                    background-color: #1f4e79;
                    color: white;
                }}

                .verde {{
                    height: 12px;
                    width: 12px;
                    background-color: green;
                    border-radius: 50%;
                    display: inline-block;
                }}

                .vermelha {{
                    height: 12px;
                    width: 12px;
                    background-color: red;
                    border-radius: 50%;
                    display: inline-block;
                }}

            </style>
        </head>
        <body>

            <h1>OnePage - Loja {id_loja}</h1>

            <p>Resumo de desempenho da loja no dia e no acumulo anual.</p>

            <br><br>

            <h2>Resultado do Dia</h2>
            {html_dia}

            <br><br>

            <h2>Resultado do Ano</h2>
            {html_ano}

        </body>
    </html>
    """

    return corpo_html


def gerar_ranking(df, coluna, df_lojas):
    df = df.copy()

    df = df.sort_values(by=coluna, ascending=False)

    mapa_lojas = df_lojas.set_index("ID Loja")["Loja"]

    df["Loja"] = df.index.map(mapa_lojas)

    df["Ranking"] = range(1, len(df) + 1)

    df = df[["Ranking", "Loja", coluna]]

    return df


def gerar_relatorios(metricas, df_lojas):

    tabelas_html = ""

    for coluna, df in metricas.items():

        ranking = gerar_ranking(df, coluna, df_lojas)  

        ranking_html = ranking.to_html(index=False, escape=False)

        tabelas_html += f"""
        <h2>{coluna}</h2>
        {ranking_html}
        <br><br>
        """
        
    corpo_html = f"""
    <html>
        <head>
            <style>

                body {{
                    font-family: Arial, sans-serif;
                }}

                table {{
                    border-collapse: collapse;
                    width: 80%;
                }}

                th, td {{
                    border: 1px solid #ddd;
                    padding: 8px;
                    text-align: center;
                }}

                th {{
                    background-color: #1f4e79;
                    color: white;
                }}

                h1 {{
                    color: #1f4e79;
                }}

                h2 {{
                    margin-top: 30px;
                    color: #333;
                }}

            </style> 
        </head>
        <body>

            <h1>Relatorio - Diretoria</h1>

            <p>Resumo geral de desempenho das lojas.</p>

            <br><br>

            {tabelas_html}

        </body>
    </html>
    """
    
    return corpo_html
    

# ========
# MAIN
# ========

def main():
    df_vendas, df_lojas, df_emails_lojas = carregar_dados(
        ARQUIVO_VENDAS,
        ARQUIVO_EMAILS,
        ARQUIVO_LOJAS
    )

    # ===== CALCULANDO INDICADORES =====

    faturamento_dia, faturamento_ano = calcular_faturamento(df_vendas)
    variedade_itens_dia, variedade_itens_ano = calcular_variedade(df_vendas)
    ticket_medio_dia, ticket_medio_ano = calcular_ticket_medio(df_vendas)
    
    # ===== APLICANDO AS METAS =====

    lista_dfs = [
        (faturamento_dia, "Faturamento Diario", FAT_META_DIA),
        (faturamento_ano, "Faturamento Anual", FAT_META_ANO),
        (variedade_itens_dia, "Variedade Diaria", VAR_META_DIA),
        (variedade_itens_ano, "Variedade Anual", VAR_META_ANO),
        (ticket_medio_dia, "Ticket Medio Dia", TICKET_META),
        (ticket_medio_ano, "Ticket Medio Ano", TICKET_META)
    ]

    (
        faturamento_dia,
        faturamento_ano,
        variedade_itens_dia,
        variedade_itens_ano,
        ticket_medio_dia,
        ticket_medio_ano
    ) = aplicar_metas(lista_dfs)

    # ===== FILTRANDOS DFS PELA DATA =====

    data_mais_recente, ano_atual = data_atual(df_vendas)

    (
        faturamento_dia_fil,
        variedade_dia_fil,
        ticket_dia_fil,
    ) = filtrar_data_dfs(
        faturamento_dia,
        variedade_itens_dia,
        ticket_medio_dia,
        "Data",
        data_mais_recente
        )

    (
        faturamento_ano_fil,
        variedade_ano_fil,
        ticket_ano_fil,
    ) = filtrar_data_dfs(
        faturamento_ano,
        variedade_itens_ano,
        ticket_medio_ano,
        "Ano",
        ano_atual
    )

    # ===== CRIANDO OS ONEPAGES =====

    onepages = {}

    for id_loja in df_emails_lojas["ID Loja"].dropna().unique():
        
        onepages[id_loja] = gerar_onepage_loja(
            id_loja,
            faturamento_dia_fil,
            variedade_dia_fil,
            ticket_dia_fil,
            faturamento_ano_fil,
            variedade_ano_fil,
            ticket_ano_fil
        )

    # ===== CRIANDO OS RELATORIOS =====

    metricas = {
        "Faturamento Diario": faturamento_dia_fil,
        "Variedade Diaria": variedade_dia_fil,
        "Ticket Medio Dia": ticket_dia_fil,
        "Faturamento Anual": faturamento_ano_fil,
        "Variedade Anual": variedade_ano_fil,
        "Ticket Medio Ano": ticket_ano_fil
    }

    relatorios = gerar_relatorios(metricas, df_lojas)


    print("Processamento concluído")


if __name__ == "__main__":
    main()