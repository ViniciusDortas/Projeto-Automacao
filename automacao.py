import pandas as pd
import numpy as np

# ============
# CONSTANTES
# ============

ARQUIVO_DADOS = r"Bases de Dados/Vendas.xlsx"

FAT_META_DIA = 1000
FAT_META_ANO = 1_650_000
VAR_META_DIA = 4
VAR_META_ANO = 120
TICKET_META = 500

# ===========
# FUNÇÕES
# ==========

def carregar_dados(caminho):
    df = pd.read_excel(caminho)
    df["Data"] = pd.to_datetime(df["Data"])
    df["Ano"] = df["Data"].dt.year
    return df


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

# ========
# MAIN
# ========

def main():
    df = carregar_dados(ARQUIVO_DADOS)

    faturamento_dia, faturamento_ano = calcular_faturamento(df)
    variedade_itens_dia, variedade_itens_ano = calcular_variedade(df)
    ticket_medio_dia, ticket_medio_ano = calcular_ticket_medio(df)
    
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

    print("Processamento concluído")


if __name__ == "__main__":
    main()