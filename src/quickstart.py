#!/usr/bin/python
import sys

# Check for minimum Python version
MIN_PYTHON = (3, 4)
if sys.version_info < MIN_PYTHON:
    sys.exit("Python %s.%s or later is required.\n" % MIN_PYTHON)



import os, re, datetime, json

import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
import google

import mimetypes

import xlsxwriter

# Necessario per la conversione dei mesi
import locale
locale.setlocale(locale.LC_TIME, "it_IT")

import pandas as pd

import logging
logger = logging.getLogger()
logger.setLevel(logging.INFO)

# import httplib2
# httplib2.debuglevel = 4

from pathlib import Path

import script

# If modifying these scopes, delete the file token.pickle.
SCOPES = [
	'openid', 
	'https://www.googleapis.com/auth/drive.readonly',
	'https://www.googleapis.com/auth/userinfo.email',
	'https://www.googleapis.com/auth/userinfo.profile',
	'https://www.googleapis.com/auth/spreadsheets.readonly', 
	'https://www.googleapis.com/auth/calendar'
]



# The ID and range of a sample spreadsheet.
SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1ut9jSOMG6YInD8_qhE_jZb1qhgxYeeKYZrOhf-IxdNU/edit#gid=990516356'
SHEET_NAME = 'TURNI 2021 '
EXPORTED_FILENAME = ".xlsx"



# Calendar Share Options (https://developers.google.com/calendar/concepts/sharing#sharing_calendars)
CALENDAR_DEFAULT_NAME = "ICTSM H24 Calendar"
CALENDAR_DEFAULT_NAME = "H24 Calendar - Test"
CALENDAR_CAN_READ 	= [
	"veronicabracco2001@gmail.com",
]
CALENDAR_CAN_WRITE 	= []









VALUES_DIRECTION = "ROWS"


def load_credentials ():
	"""Shows basic usage of the Sheets API.
	Prints values from a sample spreadsheet.
	"""
	creds = None

	token_file = Path(script.getParent(), 'token.pickle').absolute()
	client_secrets_file = Path(script.getParent(), 'credentials.json').absolute()
	print(token_file, client_secrets_file)

	# The file token.pickle stores the user's access and refresh tokens, and is
	# created automatically when the authorization flow completes for the first
	# time.
	if token_file.exists():
		with open(token_file, 'rb') as token:
			creds = pickle.load(token)
		# creds = pickle.load(token_file.read_file(mode='rb'))

	# If there are no (valid) credentials available, let the user log in.
	if not creds or not creds.valid:
		if creds and creds.expired and creds.refresh_token:
			try: 
				creds.refresh(google.auth.transport.requests.Request())
			except google.auth.exceptions.RefreshError:
				print("Refresh token scaduto o revocato")
				token_file.unlink()

				return None
		else:
			flow = InstalledAppFlow.from_client_secrets_file(
				client_secrets_file=client_secrets_file,
				scopes=SCOPES
			)
			
			creds = flow.run_local_server(
				port=0, 
				success_message="Il flusso di autenticazione è terminato. Ora puoi chiudere questa finestra"
			)

		# Save the credentials for the next run
		with open(token_file, 'wb') as token:
			pickle.dump(creds, token)
			
	return creds


def download_from_drive(fileId:str, credentials, extension:str="xlsx"):
	downloaded_name = f"{fileId}.{extension}"
	correct_mimeType = mimetypes.types_map[f".{extension}"]

	print(f"Download started... [To: {extension}]")
	with build('drive', 'v3', credentials=credentials) as drive_service:
		data = drive_service.files() \
							.export(
								fileId=fileId, 
								mimeType=correct_mimeType
							) \
							.execute()

		if not data: 
			print(f"Nessun foglio trovato con l'id '{fileId}' ({SPREADSHEET_URL})")
			return None

		with open(downloaded_name, 'wb') as f:
			f.write(data)

		if not os.path.exists(downloaded_name):
			print(f"Non sono riuscito a scaricare il foglio con id '{fileId}' ({SPREADSHEET_URL})")
			return None
		
		return downloaded_name


def get_user_info (credentials):
	user_info = {}
	with build('oauth2', 'v2', credentials=credentials) as user_info_service:
		user_info = user_info_service.userinfo().get().execute()

	return user_info


def convert_month_to_number(month_name):
	datetime_object = datetime.datetime.strptime(month_name[0:3], "%b")
	return datetime_object.month


# url = "https://www.googleapis.com/oauth2/v2/userinfo?alt=json"
# headers = {
# 	'Content-Type': "application/json",
# 	'Authorization': f"Bearer {access_token}",
# }

# response = requests.request("GET", url, headers=headers)
# print(response.status_code)
# print(response.text)


class GoogleCalendar(object):
	@staticmethod
	def getAllCalendars(service):
		calendars = []
		page_token = None
		while True:
			calendar_list = service.calendarList().list(pageToken=page_token).execute()
			calendars.extend(calendar_list['items'])

			page_token = calendar_list.get('nextPageToken')
			if not page_token:
				break
		
		return calendars

	@staticmethod
	def importEvents(events: list, credentials: google.oauth2.credentials.Credentials) -> bool:
		with build('calendar', 'v3', credentials=credentials) as service:
			all_calendars = GoogleCalendar.getAllCalendars(service=service)
			calendar_id = ""

			print("User's Calendars:")
			print("\n\n".join([f'\tID: {calendar["id"]}\n\tSummary: {calendar["summary"]}' for calendar in all_calendars]))
			print("\n\n")

			# If calendar exists, then destroy it
			if CALENDAR_DEFAULT_NAME in [calendar["summary"] for calendar in all_calendars]:
				calendar_id = [calendar["id"] for calendar in all_calendars if calendar["summary"] == CALENDAR_DEFAULT_NAME][0]
				print(f"Calendario '{CALENDAR_DEFAULT_NAME}' già presente con ID '{calendar_id}'. Lo cancello...")
				service.calendars().delete(calendarId=calendar_id).execute()
				

			
				
			# Then create it once again
			print(f"Calendario '{CALENDAR_DEFAULT_NAME}' non presente. Lo creo...")
			
			calendar = {
				'summary': CALENDAR_DEFAULT_NAME,
				'description': 'Questo è un test per il sync dei turni H24 del Team ICTSM Trenitalia',
				'timeZone': 'Europe/Rome'
			}
			created_calendar = service.calendars().insert(body=calendar).execute()
			print(f"Calendar created: {json.dumps(created_calendar, indent=4)}")

			calendar_id = created_calendar['id']



			# Add the created calendar to the User's list
			calendar_list_entry = {
				'id': calendar_id,
				'backgroundColor': '#c71912',
				'foregroundColor': '#ffffff',
				'hidden': False
			}
			created_calendar_list_entry = service.calendarList().insert(body=calendar_list_entry, colorRgbFormat=True).execute()
			print(f"Calendar inserted: {json.dumps(created_calendar_list_entry, indent=4)}")
			
			

			
			# Set the correct permissions
			for reader_user in CALENDAR_CAN_READ:
				rule = {
					'scope': {
						'type': 'user',
						'value': reader_user,
					},
					'role': 'reader'
				}

				created_rule = service.acl().insert(calendarId=calendar_id, body=rule).execute()
				print(f"--> Rule inserted: {json.dumps(created_rule, indent=4)}")
			else:
				print("Nessun altro utente specificato")

			return False

			event = {
				'summary': 'TEST - Google I/O 2015',
				'location': '800 Howard St., San Francisco, CA 94103',
				'description': 'A chance to hear more about Google\'s developer products.',
				'start': {
					'dateTime': '2021-02-08T15:00:00Z',
					# 'timeZone': 'America/Los_Angeles',
				},
				'end': {
					'dateTime': '2021-02-08T17:00:00Z',
					# 'timeZone': 'America/Los_Angeles',
				},
				'recurrence': [
					# 'RRULE:FREQ=DAILY;COUNT=2'
				],
				'attendees': [
					# {'email': 'lpage@example.com'},
					# {'email': 'sbrin@example.com'},
				],
				'reminders': {
					'useDefault': False,
					'overrides': [
						{'method': 'email', 'minutes': 24 * 60},
						{'method': 'popup', 'minutes': 10},
					],
				},
			}

			event = service.events().insert(calendarId=calendar_id, body=event).execute()
			print(f"Event added: {json.dumps(event, indent=4)}")



def main():
	# First load credentials from file or show consent screen
	credentials = load_credentials()
	if not credentials:
		print("Errore durante la ricezione delle credenziali o dell'autenticazione")
		return False


	# Find Google Drive file ID from Spreadsheet URL 
	spreadsheet_id = re.match("https://docs.google.com/spreadsheets/d/(.*?)/", SPREADSHEET_URL).groups(0)[0]
	

	# Download file from Google Drive
	# downloaded_filename = ""
	downloaded_filename = download_from_drive(credentials=credentials, fileId=spreadsheet_id, extension="xlsx")
	if not downloaded_filename:
		print("Impossibile procedere: File non scaricato")
		return False

	print(f"File scaricato: {downloaded_filename}\n")


	# Main program
	try:
		# Get user info (email, name...)
		user_info = get_user_info(credentials=credentials)
		print(f"User Info: {json.dumps(user_info, indent=4)}")
		user_lastname = user_info["family_name"]

		events = []

		# Call the Calendar API
		# GoogleCalendar.importEvents(events)



	finally:
		print()
		if downloaded_filename:
			os.remove(downloaded_filename)
			print(f"File rimosso: '{downloaded_filename}'")



	return



	# How values should be represented in the output.
	# The default render option is ValueRenderOption.FORMATTED_VALUE.
	value_render_option = 'FORMULA'  # TODO: Update placeholder value.

	# How dates, times, and durations should be represented in the output.
	# This is ignored if value_render_option is
	# FORMATTED_VALUE.
	# The default dateTime render option is [DateTimeRenderOption.SERIAL_NUMBER].
	date_time_render_option = 'FORMATTED_STRING'  # TODO: Update placeholder value.
	
	with build('sheets', 'v4', credentials=credentials) as service:
		# Call the Sheets API
		sheet = service.spreadsheets()

		workbook = sheet.get(spreadsheetId=SPREADSHEET_ID, includeGridData=False).execute()
		available_sheets = [ sheet["properties"]["title"] for sheet in workbook["sheets"]]

		print("Scrivo i dati su file...")
		with open(file="workbook2.json", mode="w") as fd:
			fd.write(json.dumps(workbook))

		return

		if not SHEET_NAME in available_sheets: 
			print(f"Sheet '{SHEET_NAME}' does not exist. [Valid values: {available_sheets}]")
			return

		print(f"Found '{SHEET_NAME}'")

		sheet_info = [ sheet for sheet in workbook["sheets"] if sheet["properties"]["title"] == SHEET_NAME ]
		if len(sheet_info) > 1: 
			print("ERRORE - Trovato più di un foglio con lo stesso nome!!")
			return

		sheet_info = sheet_info[0]
		highest_row = sheet_info["properties"]["gridProperties"]["rowCount"]
		highest_col = sheet_info["properties"]["gridProperties"]["columnCount"]
		
		print(f"[Rows] MAX: '{highest_row}'")
		print(f"[Cols] MAX: '{highest_col}'")
		
		custom_range = f"A1:{xlsxwriter.utility.xl_col_to_name(highest_col-1)}{highest_row}"
		print(f"Fetching custom range: {custom_range}")
		# cells = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=custom_range, majorDimension=VALUES_DIRECTION, valueRenderOption=value_render_option).execute()
		cells = sheet.values().batchGetByDataFilter(
			spreadsheetId=SPREADSHEET_ID, 
			dataFilters={
				"developerMetadataLookup": {
					"locationMatchingStrategy": "INTERSECTING_LOCATION",
					"visibility": "true"
				},
				"a1Range": custom_range, 
			},
			majorDimension=VALUES_DIRECTION, 
			valueRenderOption=value_render_option
		).execute()
		
		# print("Scrivo i dati su file...")
		# with open(file="cells.json", mode="w") as fd:
		# 	fd.write(json.dumps(cells))

		# cells[riga][colonna]

		intervallo = [
			# f"{cells['values'][3][col]} {cells['values'][2][col]} {cells['values'][1][col]}" for col in range(4, highest_col-1) if len(cells["values"][1]) > col
			f"{xlsxwriter.utility.xl_col_to_name(col)}2: {cells['values'][2][col]}/{(cells['values'][1][col])}" for col in range(4, highest_col-1) if len(cells["values"][1]) > col
		]

		print(intervallo)
		intervallo = [
			# f"{cells['values'][3][col]} {cells['values'][2][col]} {cells['values'][1][col]}" for col in range(4, highest_col-1) if len(cells["values"][1]) > col
			f"{cells['values'][2][col]}/{convert_month_to_number(cells['values'][1][col])}" for col in range(4, highest_col-1) if len(cells["values"][1]) > col
		]
		
		current_year = (datetime.datetime.now()).year

		def format_date(date):
			curr_day = int(date.split("/")[0])
			curr_month = int(date.split("/")[1])

			prev_datetime = datetime.datetime(day=curr_day, month=curr_month, year=(current_year-1))
			curr_datetime = datetime.datetime(day=curr_day, month=curr_month, year=current_year)
			result = prev_datetime if curr_month > int(intervallo[-1].split("/")[1]) else curr_datetime
			
			return result
		
		for value in intervallo:
			print(f"Converting {value} => ", end="", flush=True)
			new_date = format_date(value)
			print(new_date)

		return

		# intervallo2 = map(
		# 	format_date,
		# 	intervallo 
		# )
		# print(intervallo)
		# print(intervallo2)
		# print(list(intervallo2))
		# return
		users = []
		frames = []

		# Loop through all cells
		# Inizio dalla seconda colonna perchè nella prima è presente unicamente il titolo
		for row in range(1, highest_row-1):
			# print(f"Riga {row}: ", end="", flush=True)

			# Se non c'è la persona
			if not cells["values"][row] or not cells["values"][row][2]:
				# print("Salto riga senza orari...")
				continue
			elif str(cells["values"][row][4]).startswith("=") and str(cells["values"][row][5]).startswith("="):
				print("Raggiunta la fine dei turni (iniziati calcoli)")
				break

			cognome = str(cells["values"][row][2]).strip().upper()
			turni = [
				f"{cells['values'][row][col]}" for col in range(4, highest_col-1) if col-3 <= len(intervallo)
			]

			d = {
				"giorni": intervallo,
				"turni": turni
			}
			users.append(cognome)
			frames.append(pd.DataFrame.from_dict(d, orient='index'))
				
			
		turni_orari = pd.concat(frames, keys=users)

		print(turni_orari)

		NEEDED_USER = "Salvarani"
		wanted_data = turni_orari.loc[NEEDED_USER.upper()]

		for (columnName, columnData) in turni_orari.loc['SALVARANI'].iteritems():
			giorno = columnData.values[0]
			turno = columnData.values[1]

		return
		values = workbook.get('values', [])

		if not values:
			print('No data found.')
		else:
			print('Name, Major:')
			for row in values:
				# Print columns A and E, which correspond to indices 0 and 4.
				print('%s, %s' % (row[0], row[4]))



if __name__ == '__main__':
	main()
