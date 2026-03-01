import pandas as pd
import numpy as np
import os
import smtplib
from email.message import EmailMessage

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

    df_emails = pd.read_csv(emails)
    df_lojas = pd.read_csv(lojas, sep=";")

    df_emails_lojas = df_emails.merge(
        df_lojas,
        on="Loja",
        how="left"
    )
    df_emails_lojas["ID Loja"] = df_emails_lojas["ID Loja"].astype("Int64")

    return df_vendas, df_lojas, df_emails_lojas


def calcular_faturamento(df, nome_coluna):

    resultado = (
        df.groupby("ID Loja")["Valor Final"]
        .sum()
        .rename( nome_coluna)
        .to_frame()
    )

    return resultado
    

def calcular_variedade(df, nome_coluna):

    resultado = (
        df.groupby("ID Loja")["Produto"]
        .nunique()
        .rename(nome_coluna)
        .to_frame()
    )

    return resultado


def calcular_ticket_medio(df, nome_coluna):

    resultado = (
        df.groupby("ID Loja")[["Quantidade", "Valor Final"]]
        .sum()
    )

    resultado[nome_coluna] = (
        resultado["Valor Final"] /
        resultado["Quantidade"].replace(0, pd.NA)
    ).round(2)

    resultado = resultado.drop(columns=["Valor Final", "Quantidade"])

    return resultado


def verificar_meta(df, valor, meta):

    df = df.copy()
    df["Bateu Meta"] = np.where(df[valor] >= meta, "Sim", "Não")
    return df


def aplicar_metas(dict_dfs):

    novos_dfs = {}

    for nome, info in dict_dfs.items():

        df_atualizado = verificar_meta(
           info["df"],
           info["coluna"],
           info["meta"] 
        )

        novos_dfs[nome] = {
            "df": df_atualizado,
            "coluna": info["coluna"],
            "meta": info["meta"]
        }

    return novos_dfs


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


def gerar_onepage_loja(
    id_loja,
    faturamento_dia,
    variedade_dia,
    ticket_dia,
    faturamento_ano,
    variedade_ano,
    ticket_ano,
):
    
    indicadores = [
        ("Faturamento", faturamento_dia, "Faturamento Diario", FAT_META_DIA,
                          faturamento_ano, "Faturamento Anual", FAT_META_ANO),

        ("Variedade", variedade_dia, "Variedade Diaria", VAR_META_DIA,
                        variedade_ano, "Variedade Anual", VAR_META_ANO),

        ("Ticket Medio", ticket_dia, "Ticket Medio Dia", TICKET_META,
                           ticket_ano, "Ticket Medio Ano", TICKET_META)
    ]

    dados_dia = []
    dados_ano = []

    for nome, df_dia, col_dia, meta_dia, df_ano, col_ano, meta_ano in indicadores:
    
        # ==== DIA ====
        if id_loja in df_dia.index:
            valor_dia = df_dia.at[id_loja, col_dia]
            cenario_dia = df_dia.at[id_loja, "Bateu Meta"]
        else:
            valor_dia = 0
            cenario_dia = "Não"

        # ==== ANO ====
        if id_loja in df_ano.index:
            valor_ano = df_ano.at[id_loja, col_ano]
            cenario_ano = df_ano.at[id_loja, "Bateu Meta"]
        else:
            valor_ano = 0
            cenario_ano = "Não"

        dados_dia.append([nome, valor_dia, meta_dia, cenario_dia])
        dados_ano.append([nome, valor_ano, meta_ano, cenario_ano])

    onepage_dia = pd.DataFrame(
        dados_dia,
        columns=["Indicador", "Valor Dia", "Meta Dia", "Cenário Dia"]
    ).set_index("Indicador")

    onepage_ano = pd.DataFrame(
        dados_ano,
        columns=["Indicador", "Valor Ano", "Meta Ano", "Cenário Ano"]
    ).set_index("Indicador")

    return onepage_dia, onepage_ano


def gerar_html_onepage(onepage_dia, onepage_ano, id_loja, data):

    def bolinha(valor):
        if valor == "Sim":
            return '<span class="verde"></span>'
        else:
            return '<span class="vermelha"></span>'
        
    onepage_dia["Cenário Dia"] = onepage_dia["Cenário Dia"].apply(bolinha)
    onepage_ano["Cenário Ano"] = onepage_ano["Cenário Ano"].apply(bolinha)

 
    html_dia = onepage_dia.to_html(escape=False)
    html_ano = onepage_ano.to_html(escape=False)

    data = data.strftime("%d/%m/%Y")

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

            <h1>OnePage - Loja {id_loja} - {data}</h1>

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


def gerar_relatorios(dict_dfs, df_lojas):

    relatorios = {}

    for nome, info in dict_dfs.items():

        relatorios[info["coluna"]] = gerar_ranking(info["df"], info["coluna"], df_lojas)  

    return relatorios


def gerar_html_relatorios(relatorios, data):

    tabelas_html = ""

    for coluna, ranking in relatorios.items():

        ranking_html = ranking.to_html(index=False, escape=False)

        tabelas_html += f"""
        <h2>{coluna}</h2>
        {ranking_html}
        <br><br>
        """

    data = data.strftime("%d/%m/%Y")
        
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

            <h1>Relatorio - Diretoria - {data}</h1>

            <p>Resumo geral de desempenho das lojas.</p>

            <br><br>

            {tabelas_html}

        </body>
    </html>
    """

    return corpo_html

    
def salvar_dados(onepages, relatorio, data):

    os.makedirs("Arquivos/OnePages", exist_ok=True)

    data = data.strftime("%d-%m-%Y")

    for id_loja, html in onepages.items():
        caminho = os.path.join("Arquivos/OnePages", f"OnePage_Loja{id_loja}_{data}.html")

        with open(caminho, "w", encoding="utf-8") as f:
            f.write(html)

    caminho = os.path.join("Arquivos", f"Relatorio_{data}.html")
    with open(caminho, "w", encoding="utf-8") as f:
        f.write(relatorio)


def enviar_email(destinario, assunto, corpo_html):

    remetente = os.getenv("EMAIL_USER")
    senha = os.getenv("EMAIL_SENHA")

    if not remetente or not senha:
        raise EnvironmentError("Variáveis EMAIL_USER ou EMAIL_SENHA não definidos")

    msg = EmailMessage()
    msg["From"] = remetente
    msg["To"] = destinario
    msg["Subject"] = assunto
    msg.add_alternative(corpo_html, subtype="html")

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(remetente, senha)
        smtp.send_message(msg)


def enviar_onepages(df_email_lojas, df_lojas, onepages, data):

    data = data.strftime("%d/%m/%Y")
    
    mapa_lojas = df_lojas.set_index("ID Loja")["Loja"]

    for _, linha in df_email_lojas.iterrows():
        id_loja = linha["ID Loja"]
        gerente = linha["Gerente"]
        email = linha["Email"]

        if pd.isna(id_loja):
            continue

        nome_loja = mapa_lojas[id_loja]

        assunto = f"OnePage - {nome_loja} - {data}"

        corpo = f"""
        <p>Olá {gerente},</p>

        <p>Segue o relátorio de desempenho da loja <strong>{nome_loja}</strong>
        referente ao dia {data}</p>

        <br>

        {onepages[id_loja]}

        <br>

        <p>Atenciosamente<br>
        Equipe Comercial</p>
        """

        enviar_email(email, assunto, corpo)

        print(f"Email enviado para a Loja {id_loja}")


def enviar_relatorio_diretoria(df_email_lojas, relatorio, data):

    data = data.strftime("%d/%m/%Y")

    df_diretoria = df_email_lojas[df_email_lojas["ID Loja"].isna()]

    for _, linha in df_diretoria.iterrows():
    
        email = linha["Email"]

        assunto = f"Relatório Geral - {data}"

        corpo = f"""
        <p>Olá Diretoria,</p>

        <p>Segue o relatório geral do dia {data}</p>

        <br>

        {relatorio}

        <br>

        <p>Atenciosamente,<br>
        Equipe Comercial</p>
        """

        enviar_email(email, assunto, corpo)

        print("Email enviado para a Diretoria")


# ========
# MAIN
# ========

def main():

    df_vendas, df_lojas, df_emails_lojas = carregar_dados(
        ARQUIVO_VENDAS,
        ARQUIVO_EMAILS,
        ARQUIVO_LOJAS
    )

    # ===== CRIANDO OS DFS =====

    data_mais_recente, ano_atual = data_atual(df_vendas)

    df_dia = df_vendas[df_vendas["Data"] == data_mais_recente]
    df_ano = df_vendas[df_vendas["Data"].dt.year == ano_atual]


    faturamento_dia = calcular_faturamento(df_dia, "Faturamento Diario")
    faturamento_ano = calcular_faturamento(df_ano, "Faturamento Anual")
    variedade_dia   = calcular_variedade(df_dia, "Variedade Diaria")
    variedade_ano   = calcular_variedade(df_ano, "Variedade Anual")
    ticket_dia      = calcular_ticket_medio(df_dia, "Ticket Medio Dia")
    ticket_ano      = calcular_ticket_medio(df_ano, "Ticket Medio Ano")


    # ===== APLICANDO AS METAS =====

    dict_dfs = {
        "faturamento_dia": {
            "df": faturamento_dia,
            "coluna": "Faturamento Diario",
            "meta": FAT_META_DIA
        },
        "faturamento_ano": {
            "df": faturamento_ano,
            "coluna": "Faturamento Anual",
            "meta": FAT_META_ANO
        },
        "variedade_dia": {
            "df": variedade_dia,
            "coluna": "Variedade Diaria",
            "meta": VAR_META_DIA
        },
        "variedade_ano": {
            "df": variedade_ano,
            "coluna": "Variedade Anual",
            "meta": VAR_META_ANO
        },
        "ticket_dia": {
            "df": ticket_dia,
            "coluna": "Ticket Medio Dia",
            "meta": TICKET_META
        },
        "ticket_ano": {
            "df": ticket_ano,
            "coluna": "Ticket Medio Ano",
            "meta": TICKET_META
        }
    }

    dict_dfs = aplicar_metas(dict_dfs)

    faturamento_dia = dict_dfs["faturamento_dia"]["df"]
    faturamento_ano = dict_dfs["faturamento_ano"]["df"]
    variedade_dia   = dict_dfs["variedade_dia"]["df"]
    variedade_ano   = dict_dfs["variedade_ano"]["df"]
    ticket_dia      = dict_dfs["ticket_dia"]["df"]
    ticket_ano      = dict_dfs["ticket_ano"]["df"]

    # ===== GERANDO OS ONEPAGES =====

    onepages = {}

    for id_loja in df_emails_lojas["ID Loja"].dropna().unique():
        
        onepage_dia, onepage_ano = gerar_onepage_loja(
            id_loja,
            faturamento_dia,
            variedade_dia,
            ticket_dia,
            faturamento_ano,
            variedade_ano,
            ticket_ano
        )

        onepages[id_loja] = gerar_html_onepage(
            onepage_dia, onepage_ano, id_loja, data_mais_recente
        )

    # ===== GERANDO OS RELATORIOS =====

    rankings = gerar_relatorios(dict_dfs, df_lojas)

    relatorios = gerar_html_relatorios(rankings, data_mais_recente)

    # ===== SALVANDO E ENVIANDO OS DADOS =====

    salvar_dados(onepages, relatorios, data_mais_recente)

    enviar_onepages(df_emails_lojas, df_lojas, onepages, data_mais_recente)
    enviar_relatorio_diretoria(df_emails_lojas, relatorios, data_mais_recente)

    print("Processamento concluído")


if __name__ == "__main__":
    main()