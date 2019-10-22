# **Gerador de planilhas no Google Drive utilizando indicadores do SonarQube**

**Primeiro passo**: Criar um projeto [Google API](https://console.developers.google.com/)

Necessário ativar as bibliotecas Google Sheets API e Google Drive  
[Tutorial mais completo](https://medium.com/@CROSP/manage-google-spreadsheets-with-python-and-gspread-6530cc9f15d1)

**Segundo passo**: Criar uma planilha para servir de template

Necessário criar a planilha e compartilhá-la com o email do cliente Google API (informado no arquivo json).

**Terceiro passo**: Adaptar as informações de consulta aos dados do sonar e de edição da planilha

*Arquivo:*

 

       {
        	"template_sheet_key":  "id da planilha de template",
        	"sheet_title":  "Título fixo da planilha ",
        	"variable_sheet_title_complement":  "Complemento do título (variável)",  
            "spreadsheet_users":  [
                "email_usuario_planilha@email.com"   
            ],
            "url_sonar_server":  "https://sonar.dtidigital.com.br/",
            "url_sonar_api":  "api/measures/component?component=",
            "sonar_token":  "token-sonar",
            "metrics_to_analyze":  [
        	    "bugs",
        	    "code_smells",
        	    "ncloc",
        	    "coverage" 
            ],
            "keys_projects_to_analyze":  [
    	        "id do projeto 1 a analisar",
    	        "id do projeto 2 a analisar"
            ],
            "number_first_sprint": 1,
            "number_current_sprint": 1,
            "template_sheet_interval_between_sprints": 2,
            "columns":  [1,  2,  3,  4],
            "initial_line":  2
       }

## **Exemplo de Uso**

Utilizando uma [planilha de exemplo](https://docs.google.com/spreadsheets/d/1LmbLPtU1RY2oOhp8-vZEpyV5qMQW3cWetej0aKeotzg/edit#gid=0) adaptaremos o arquivo de dados da consulta e o script para conseguirmos preenchê-la automaticamente.

#### Alterações no arquivo json
    {
	    "template_sheet_key": "1LmbLPtU1RY2oOhp8-vZEpyV5qMQW3cWetej0aKeotzg",
	    "sheet_title": "Teste ",
	    "variable_sheet_title_complement": "1",
	    "spreadsheet_users": [
	        "email@email.com"
	    ],

	    "url_sonar_server": "https://sonar.dtidigital.com.br/",
	    "url_sonar_api": "api/measures/component?component=",
	    "sonar_token": "",
	    "metrics_to_analyze": [
	        "ncloc",
	        "code_smells",
	        "bugs"
	    ],
	    "keys_projects_to_analyze": [
	        ""
	    ],

	    "number_first_sprint": 1,
	    "number_current_sprint": 1,
	    "template_sheet_interval_between_sprints": 1,

	    "columns": [2, 3, 4],
    	    "initial_line": 2
    }

#### Alterações no script

    googleCredentialsFilePath =  'caminho-para-arquivo-cliente-google-api.json'
    
    configFilePath =  './caminho-para-arquivo-de-configuracao.json'
