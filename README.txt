Feito por Gesley Rosa

Teste Automação de Dados - Site Magazine Luiza

Pacotes: Selenium, Pandas, Openpyxl, Yagmail, dotenv

Instalação de pacotes: pip install selenium pandas openpyxl yagmail python-dotenv

Coloque o ChromeDriver compatível com sua versão do navegador no PATH.

## Como configurar as credenciais

1. Ative a verificação em duas etapas no Gmail
2. Gere uma senha de aplicativo: https://myaccount.google.com/apppasswords
3. Crie o arquivo `config/credentials.env` com:

EMAIL_REMETENTE=seuemail@gmail.com  
SENHA_APP=sua-senha-do-app

Esse arquivo é ignorado pelo Git para proteger suas informações.

## Executar

python main.py

O resultado estará em `Output/Notebooks.xlsx` e será enviado por e-mail automaticamente.

