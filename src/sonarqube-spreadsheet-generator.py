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
with open(configFilePath , 'r') as file:
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
cellsIntervalBetweenSprints = config['template_sheet_interval_between_sprints']
columns = config['columns']
initialLine = config['initial_line']
#
#
def fillWorksheetData():
  session = requests.Session()
  session.auth = sonarToken, ''
  call = getattr(session, 'get')

  for project in projectsToAnalyzeTogether:
    callUrl = sonarServerUrl+sonarApiUrl+project+'&metricKeys='

    for metric in metricsToAnalyze:
      callUrl += metric + ','

    analysisData = call(callUrl)
    analysisDataJson = analysisData.json()
    component = analysisDataJson['component']

    for metric in metricsToAnalyze:
      vars()[metric] = 0
      
      for measure in component['measures']:
        value = float(measure['value'])

        if (measure['metric'] == metric):
            vars()[metric] += value
  print("Dados obtidos do Sonar")

###
  print("Montando planilha...")

  spreadsheet = client.open(sheetTitle + sheetTitleComplement)
  sheet = spreadsheet.sheet1

  linesToAdd = (((currentSprintNumber-firstSprintNumber) + 1) * cellsIntervalBetweenSprints)

  i = 0
  for metric in metricsToAnalyze:
      sheet.update_cell(initialLine + linesToAdd, columns[i], vars()[metric])
      i += 1

##
##
try:
  print("Abrindo planilha...")
  spreadsheet = client.open(sheetTitle + sheetTitleComplement)
  
  print("Planilha encontrada. Preenchendo dados...")
  
  fillWorksheetData()
  
  print("Concluído. Planilha preenchida.")

except gspread.exceptions.SpreadsheetNotFound:
  print("Planilha não encontrada! Criando planilha...")
  spreadsheet = client.copy(templateSheetKey, sheetTitle + sheetTitleComplement, True)

  print("Planilha criada")
  
  for u in spreadsheetUsers:
      spreadsheet.share(u, perm_type='user', role='writer')
  
  print("Planilha compartilhada com usuários")

  print("Preenchendo planilha...")
  
  fillWorksheetData()
  
  print("Concluído. Planilha preenchida.")

####
####