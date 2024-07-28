def extrai():
    """
    Objetivo da Função: Extrair os compromissos do outlook e exportá-los em um excel

    """

    ## Importa as bibliotecas e funções
    import win32com.client
    from datetime import datetime, timedelta
    import pandas as pd

    ## Conecta no Outlook
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    ## Acessa o calendário
    calendar_folder = namespace.GetDefaultFolder(9)

    ## Pega a data de hoje, para filtrar
    start_date = datetime.now()
    end_date = start_date + timedelta(days=30)

    ## Formata a data no formato esperado pelo Outlook
    start_date_str = start_date.strftime("%m/%d/%Y %H:%M %p")
    end_date_str = end_date.strftime("%m/%d/%Y %H:%M %p")

    ## Cria o filtro
    filter = f"[Start] >= '{start_date_str}' AND [End] <= '{end_date_str}'"

    ## Extrai os itens aplicando o filtro
    items = calendar_folder.Items

    ## É necessário também parametrizar os compromissos recorrentes
    items.IncludeRecurrences = True

    ## Ordena (Apenas para o print)
    items.Sort("[Start]")

    ## Aplica o filtro
    restricted_items = items.Restrict(filter)

    ## Iterar sobre os compromissos e extrai informações
    ## Cria uma lista vazia, para armazenar os dados e posteriormente transforma-los em um dataframe

    lista_vazia = []
    for compromisso in restricted_items:
        # print("Assunto:", compromisso.Subject)
        # print("Início:", compromisso.Start)
        # print("Término:", compromisso.End)
        # print("Local:", compromisso.Location)
        # print("-------------------------------------------------")

        ## Cria um dataframe para armazenar esta linha
        df_compilar = pd.DataFrame(
            [
                {
                    "Assunto": str(compromisso.Subject),
                    "Início": str(compromisso.Start),
                    "Término": str(compromisso.End),
                }
            ]
        )

        ## Armazena esta linha na lista vazia
        lista_vazia.append(df_compilar)

    ## Concatena os resultados em um dataframe final
    df_final = pd.concat(lista_vazia)

    ## Reseta o index
    df_final = df_final.reset_index(drop=True)

    ## Filtra a data mais recente
    df_final["Filtro"] = pd.to_datetime(df_final["Início"])
    df_final = df_final[
        df_final["Filtro"].dt.date == df_final["Filtro"].dt.date.min()
    ].reset_index(drop=True)

    ## Deleta a coluna usada apenas para o filtra
    del df_final["Filtro"]

    ## Exporta os resultados
    df_final.to_excel("output/extraidos.xlsx", index=False)
