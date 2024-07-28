# Migrador de Compromissos do Outlook para o Google Calendar

Este script facilita a migração dos compromissos da sua conta Outlook para o Google Calendar utilizando `win32com` e a API do Google Calendar.

## Instalação

Para instalar todas as dependências necessárias, utilize o comando abaixo:

```bash
pip install -r requirements.txt
```

## Uso

Abaixo está um exemplo de como usar o script para migrar compromissos:

```python
# Carrega funções e bibliotecas
import os
from functions.extrai import extrai
from functions.insere import insere

# 1. Extrai os compromissos do Outlook
extrai()

# 2. Insere os compromissos no Google Calendar
insere()
```

### Função `extrai()`

A função `extrai()` utiliza `win32com` para acessar os eventos no Outlook instalado no Windows, extrair os compromissos, armazená-los em um DataFrame, e finalmente exportar este DataFrame para um arquivo Excel.

### Função `insere()`

A função `insere()` carrega o arquivo Excel exportado e utiliza a API do Google Calendar para inserir os compromissos na conta do Google Calendar.

## Contribuições

Pull requests são bem-vindos. Para mudanças significativas, por favor, abra uma issue primeiro para discutir o que você gostaria de modificar.

Por favor, assegure-se de atualizar os testes conforme necessário.

## Licença

Este projeto está licenciado sob a licença MIT. Veja o arquivo [MIT](https://choosealicense.com/licenses/mit/) para mais detalhes.
