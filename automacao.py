import pandas as pd


def carregar_dados(caminho):
    df = pd.read_excel(caminho)
    df["Ano"] = df["Data"].dt.year
    return df


def calcular_faturamento(df):
    faturamento_dia = (
        df.groupby(["ID Loja", "Data"])["Valor Final"]
        .sum()
        .reset_index()
        .rename(columns={"Valor Final":"Faturamento Diario"})
    )

    faturamento_ano = (
        df.groupby(["ID Loja", "Ano"])["Valor Final"]
        .sum()
        .reset_index()
        .rename(columns={"Valor Final":"Faturamento Anual"})
    )

    return faturamento_dia, faturamento_ano
    

def calcular_variedade(df):
    variedade_itens_dia = (
        df.groupby(["ID Loja", "Data"])["Produto"]
        .nunique()
        .reset_index()
        .rename(columns={"Produto":"Variedade Diaria"})
    )

    variedade_itens_ano = (
        df.groupby(["ID Loja", "Ano"])["Produto"]
        .nunique()
        .reset_index()
        .rename(columns={"Produto":"Variedade Anual"})
    )

    return variedade_itens_dia, variedade_itens_ano


def main():
    df = carregar_dados(r"Bases de Dados/Vendas.xlsx")

    faturamento_dia, faturamento_ano = calcular_faturamento(df)
    variedade_itens_dia, variedade_itens_ano = calcular_variedade(df)

    print("Processamento concluído")

if __name__ == "__main__":
    main()