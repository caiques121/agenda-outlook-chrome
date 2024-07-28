def insere():
    """
    Objetivo da Função: Carregar os compromissos extraídos do outlook e inseri-los no Google Calendar

    """
    ## Carrega libs
    import os
    import pandas as pd
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

    ## Carrega os compromissos extraídos
    df = pd.read_excel("output/extraidos.xlsx")

    ## Configurar o escopo da API
    SCOPES = ["https://www.googleapis.com/auth/calendar"]

    ## Cria a função que autentica
    def authenticate_google_api():
        """
        Esta função:
        - Checa se existe o arquivo token.json
            - O arquivo token.json armazena o token de acesso do usuário e é criado automaticamente quando o fluxo de autorização é concluído pela primeira vez.
        - Se não existe ou expirou, faz o login usando as credenciais obtidas no Google Cloud Console
        - Retorna as credenciais, no fim da função
        """

        ## Checa se o aquivo token.json existe
        creds = None
        if os.path.exists("token.json"):
            creds = Credentials.from_authorized_user_file("token.json", SCOPES)

        ## Faz login, se necessário
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    "input/credentials.json", SCOPES
                )
                creds = flow.run_local_server(port=0)
            with open("token.json", "w") as token:
                token.write(creds.to_json())
        return creds

    ## Checa se o compromisso já existe, para não inseri-lo duas vezes
    def event_exists(service, calendar_id, subject, start_time, end_time):
        events_result = (
            service.events()
            .list(
                calendarId=calendar_id,
                q=subject,
                singleEvents=True,
            )
            .execute()
        )
        events = events_result.get("items", [])
        return len(events) > 0

    ## Cria o compromisso no Google Calendar
    def create_google_calendar_event(service, subject, start_time, end_time):
        event = {
            "summary": subject,
            "start": {
                "dateTime": start_time,
                "timeZone": "America/Sao_Paulo",  # Ajuste para seu fuso horário
            },
            "end": {
                "dateTime": end_time,
                "timeZone": "America/Sao_Paulo",
            },
        }
        event = service.events().insert(calendarId="primary", body=event).execute()
        print(f"Evento Criado: {event.get('htmlLink')}")

    ## Autentica e cria o serviço da API do Google Calendar
    creds = authenticate_google_api()
    service = build("calendar", "v3", credentials=creds)
    calendar_id = "primary"

    ## Formata as datas
    df["Início"] = pd.to_datetime(df["Início"])
    df["Início"] = df["Início"] + pd.Timedelta(hours=3)
    df["Início"] = df["Início"].dt.strftime("%Y-%m-%dT%H:%M:%S%z")

    df["Término"] = pd.to_datetime(df["Término"])
    df["Término"] = df["Término"] + pd.Timedelta(hours=3)
    df["Término"] = df["Término"].dt.strftime("%Y-%m-%dT%H:%M:%S%z")

    ## Itera pelos compromissos no DataFrame e cria os eventos no Google Calendar
    for index, row in df.iterrows():
        subject = row["Assunto"]
        start_time = str(row["Início"])
        end_time = str(row["Término"])

        if not event_exists(service, calendar_id, subject, start_time, end_time):
            create_google_calendar_event(service, subject, start_time, end_time)
        else:
            print(f"Evento já existente: {subject} from {start_time} to {end_time}")
