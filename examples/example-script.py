import gspread
from gspread_formatting import *
from oauth2client.service_account import ServiceAccountCredentials
import requests, json

googleCredentialsFilePath = 'client_secret.json'
configFilePath = './config.json'

#
# GOOGLE SHEETS API ACCESS
#
scope = ['https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name(googleCredentialsFilePath, scope)
client = gspread.authorize(credentials)

# READ FROM JSON CONFIG FILE
with open(configFilePath, 'r', encoding='utf-8') as file:
    config = json.load(file)

#
# CONFIG VARIABLES
#
templateSheetKey = config['template_sheet_key']
sheetTitle = config['sheet_title']
sheetTitleComplement = config['variable_sheet_title_complement']
projectsToAnalyzeTogether = config['keys_projects_to_analyze']
sonarToken = config['sonar_token']
sonarServerUrl = config['url_sonar_server']
sonarApiUrl = config['url_sonar_api']
metricsToAnalyze = config['metrics_to_analyze']
spreadsheetUsers = config['spreadsheet_users']
firstSprintNumber = config['number_first_sprint']
currentSprintNumber = config['number_current_sprint']
intervalsBetweenSprints = config['template_sheet_interval_between_sprints']
columns = config['columns']
initialLine = config['initial_line']

#
# FUNÇÕES
#
def obterDadosSonar():
  session = requests.Session()
  session.auth = sonarToken, ''
  call = getattr(session, 'get')

  # PROJECTS DATA
  for metric in metricsToAnalyze:
    vars()[metric] = 0

  for project in projectsToAnalyzeTogether:
    callUrl = sonarServerUrl+sonarApiUrl+project+'&metricKeys='

    for metric in metricsToAnalyze:
      callUrl += metric + ','

    analysisData = call(callUrl)
    analysisDataJson = analysisData.json()
    component = analysisDataJson['component']

    for metric in metricsToAnalyze:
      for measure in component['measures']:
        value = float(measure['value'])

        if (measure['metric'] == metric):
          vars()[metric] += value
  
  dictionaryDomain = {}
  for metric in metricsToAnalyze:
    dictionaryDomain.update({metric: vars()[metric]})

  return dictionaryDomain


def fillWorksheetData(**dados):
  spreadsheet = client.open(sheetTitle + sheetTitleComplement)
  sheet = spreadsheet.sheet1

  cont = 0
  linesToAdd = (((currentSprintNumber-firstSprintNumber) + 1) * intervalsBetweenSprints)
  for metric in metricsToAnalyze:
      sheet.update_cell(initialLine + linesToAdd + cont, columns[0], dados[metric])
      cont += 1

def fillWorksheet():
  fillWorksheetData(**(obterDadosSonar()))

##
##
try:
  print("Abrindo planilha...")
  spreadsheet = client.open(sheetTitle + sheetTitleComplement)
  print("Planilha encontrada. Preenchendo dados...")
  fillWorksheet()
  print("Concluído. Planilha preenchida.")
except gspread.exceptions.SpreadsheetNotFound:
  print("Planilha não encontrada")
  spreadsheet = client.copy(templateSheetKey, sheetTitle + sheetTitleComplement, True)

  print("Planilha criada")
  for u in spreadsheetUsers:
      spreadsheet.share(u, perm_type='user', role='writer')
  print("Planilha compartilhada com usuários")

  print("Preenchendo dados anteriores ao início da migração...")
  fillWorksheet(True)
  print("Concluído. Planilha preenchida.")
####
####