import pandas as pd
import numpy as np

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

    return df_vendas, df_emails_lojas


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


def filtrar_data_dfs(df, coluna, data):
    return df[df[coluna] == data]


def loc_onepage(df, id_loja, coluna):
    if id_loja in df.index:
        return df.at[id_loja, coluna]
    return 0
    
# ========
# MAIN
# ========

def main():
    df_vendas, df_emails_lojas = carregar_dados(
        ARQUIVO_VENDAS,
        ARQUIVO_EMAILS,
        ARQUIVO_LOJAS
    )

    faturamento_dia, faturamento_ano = calcular_faturamento(df_vendas)
    variedade_itens_dia, variedade_itens_ano = calcular_variedade(df_vendas)
    ticket_medio_dia, ticket_medio_ano = calcular_ticket_medio(df_vendas)
    
    lista_dfs = [
        (faturamento_dia, "Faturamento Diario", FAT_META_DIA),
        (faturamento_ano, "Faturamento Anual", FAT_META_ANO),
        (variedade_itens_dia, "Variedade Diaria", VAR_META_DIA),
        (variedade_itens_ano, "Variedade Anual", VAR_META_ANO),
        (ticket_medio_dia, "Ticket Medio Dia", TICKET_META),
        (ticket_medio_ano, "Ticket Medio Ano", TICKET_META)
    ]

    dfs_novos = [
        verificar_meta(df, valor, meta)
        for df, valor, meta in lista_dfs
    ]

    (
        faturamento_dia,
        faturamento_ano,
        variedade_itens_dia,
        variedade_itens_ano,
        ticket_medio_dia,
        ticket_medio_ano
    ) = dfs_novos

    data_mais_recente, ano_atual = data_atual(df_vendas)

    lista_dfs_dia = (
        faturamento_dia,
        variedade_itens_dia,
        ticket_medio_dia,
    )

    dfs_filtrados_dia = [
        filtrar_data_dfs(df, "Data", data_mais_recente).set_index("ID Loja")
        for df in lista_dfs_dia
    ]

    (
        faturamento_dia_fil,
        variedade_dia_fil,
        ticket_dia_fil,
    ) = dfs_filtrados_dia


    lista_dfs_ano = (
        faturamento_ano,
        variedade_itens_ano,
        ticket_medio_ano
    )

    dfs_filtrados_ano = [
        filtrar_data_dfs(df, "Ano", ano_atual).set_index("ID Loja")
        for df in lista_dfs_ano
    ]

    (
        faturamento_ano_fil,
        variedade_ano_fil,
        ticket_ano_fil,
    ) = dfs_filtrados_ano


    for id_loja in df_emails_lojas["ID Loja"].dropna().unique():
        onepage_dia = pd.DataFrame({
            "Indicador": [
                "Faturamento",
                "Variedade",
                "Ticket Medio"
            ],
            "Valor Dia": [
                loc_onepage(
                    faturamento_dia_fil,
                    id_loja,
                    "Faturamento Diario",
                ),
                loc_onepage(
                    variedade_dia_fil,
                    id_loja,
                    "Variedade Diaria",
                ),
                loc_onepage(
                    ticket_dia_fil,
                    id_loja,
                    "Ticket Medio Dia",
                )
            ],
            "Meta Dia": [
                FAT_META_DIA,
                VAR_META_DIA,
                TICKET_META
            ],
            "Cenário Dia": [
                loc_onepage(
                    faturamento_dia_fil,
                    id_loja,
                    "Bateu Meta",
                ),
                loc_onepage(
                    variedade_dia_fil,
                    id_loja,
                    "Bateu Meta",
                ),
                loc_onepage(
                    ticket_dia_fil,
                    id_loja,
                    "Bateu Meta",
                )
            ]       
        }).set_index("Indicador")

        onepage_ano = pd.DataFrame({
            "Indicador": [
                "Faturamento",
                "Variedade",
                "Ticket Medio"
            ],
            "Valor Ano": [
                loc_onepage(
                    faturamento_ano_fil,
                    id_loja,
                    "Faturamento Anual",
                ),
                loc_onepage(
                    variedade_ano_fil,
                    id_loja,
                    "Variedade Anual",
                ),
                loc_onepage(
                    ticket_ano_fil,
                    id_loja,
                    "Ticket Medio Ano",
                )
            ],
            "Meta Ano": [
                FAT_META_ANO,
                VAR_META_ANO,
                TICKET_META
            ],
            "Cenário Ano": [
                loc_onepage(
                    faturamento_ano_fil,
                    id_loja,
                    "Bateu Meta",
                ),
                loc_onepage(
                    variedade_ano_fil,
                    id_loja,
                    "Bateu Meta",
                ),
                loc_onepage(
                    ticket_ano_fil,
                    id_loja,
                    "Bateu Meta",
                )
            ]       
        }).set_index("Indicador")

    print("Processamento concluído")


if __name__ == "__main__":
    main()