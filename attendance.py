#!/usr/bin/python3 

# This file is available from GitHub at:
# https://github.com/maciej-radziejewski/get-attendance

import os, sys, subprocess
import csv
import datetime
from datetime import datetime, date, timedelta, time

import locale
locale.setlocale(locale.LC_ALL, '') # Init the locale from the system settings

# These parameters need to be adjusted to your system.

dir_downloads = os.path.expanduser('~') + '/Downloads' # Set to the location where the Teams lists and reports are stored.

# Column headers will probably be different in your language.
# Column numbers, i.e. what appears in which column, may change between Teams versions.

column_header_full_name = 'Imię i nazwisko' # In attendance lists and reports
column_number_full_name_in_lists = 0
column_number_full_name_in_reports = 0

column_header_time_mark = 'Znacznik czasu' # In attendance lists
column_number_time_mark_in_lists = 2

column_header_join_time = 'Godzina dołączenia' # In attendance reports
column_number_join_time_in_reports = 1

column_header_leave_time = 'Godzina opuszczenia' # In attendance reports
column_number_leave_time_in_reports = 2

column_header_role = 'Rola' # In attendance reports
column_number_role_in_reports = 5

organizer_role_in_reports = 'Organizator'

number_of_columns_in_reports = 7 # Number of columns in the main part of the attendance report, after the irregular header.

time_format = '%d.%m.%Y, %H:%M:%S' # The way dates and times appear in Teams lists and reports.


irregular_classes = False


class UserError(Exception):
	def __init__ (self, message):
		self.message = message
	def __str__ (self):
		return self.message



def get_schedule (fname):
	'''
	Returns the schedule as a list of:
	day of week,
	meeting time (mod 7 days),
	group name,
	suffix to add to the date.

	The suffix is used if a group has multiple meetings on one day.
	'''

	days_of_week = {} # Map weekday name to a date near the present.
	d = date.today()
	for i in range(7):
		d2 = d+timedelta(i)
		days_of_week[d2.strftime('%A')] = d2

	schedule = []
	try:
		with open(fname, newline='') as csvfile:
			reader = csv.reader(csvfile)
			for row in reader:
				for i in range(len(row)):
					row[i]=row[i].strip()
				schedule.append([row[0], datetime.strptime(row[1], '%H:%M').time(), row[2]])
				if row[2] == 'schedule':
					print('No class/group can be named "schedule".')
					exit()
	except Exception:
		print ('I was not able to read the schedule.')
		return []

	multiple_meetings = {} # Check which lists need multiple entries per day.
	for entry in schedule:
		day_group = str((entry[0], entry[2]))
		if day_group not in multiple_meetings:
			multiple_meetings[day_group] = []
		multiple_meetings[day_group].append(entry)

	for dg in multiple_meetings:
		if len(multiple_meetings[dg]) > 1:
			for entry in multiple_meetings[dg]:
				entry.append(entry[1].strftime(' %H:%M'))
		else:
			for entry in multiple_meetings[dg]:
				entry.append('')

	for entry in schedule:
		entry[1] = datetime.combine(days_of_week[entry[0]], entry[1])

	return schedule

def create_placeholder_schedule (fname, meetings):
	times = set()
	for t, duration, participants in meetings:
		# Get hour and minutes as the number of minutes.
		basedate = datetime(2021,5,3) # Monday, to get weekdays ordered.
		m = (t - basedate)/timedelta(minutes=1)
		# Round to the nearest multiple of 15.
		m += 7
		m -= (m % 15)
		t = basedate + timedelta(minutes = m)
		tmod7 = basedate + timedelta(minutes = m % (7*24*60))
		times.add ((tmod7, t.strftime('%A'), t.strftime('%H:%M')))
	with open(fname, 'w', newline='') as csvfile:
		writer = csv.writer(csvfile)
		for i, (t7, d, t) in enumerate(sorted(times)):
			writer.writerow([d, t, 'class' + str(i+1)])

def get_attendance_list (source):
	'''Parses the output of Ms Teams attendance list
	that you can download during the meeting after clicking the list of participants.
	Returns a tuple:
	earliest time recorded, duration, the list of participants, column header for names.
	The meeting organizer should be the first participant listed in the file and
	should appear an odd number of times. The organizer is excluded from the list.
	'''
	participants = set()
	times = set()
	organizer_count = 0
	awaiting_headers = True
	try:
		with open(source, newline='', encoding='utf-16') as csvfile:
			reader = csv.reader(csvfile, dialect='excel-tab')
			for row in reader:
				if awaiting_headers:
					awaiting_headers = False
					if row[column_number_full_name_in_lists] != column_header_full_name:
						raise UserError ('Column header ' + column_header_full_name + ' expected and ' + row[column_number_full_name_in_lists] + ' encountered.')
					if row[column_number_time_mark_in_lists] != column_header_time_mark:
						raise UserError ('Column header ' + column_header_time_mark + ' expected and ' + row[column_number_time_mark_in_lists] + ' encountered.')
				else:
					name = row[column_number_full_name_in_lists]
					if organizer_count == 0:
						organizer = name
					try:
						times.add(datetime.strptime(row[column_number_time_mark_in_lists], time_format))
					except ValueError:
						# Apparently Teams sometimes randomly switches to the US time format even on non-US systems
						locale.setlocale(locale.LC_TIME, 'en_US') # The locale change is local - will be reversed in a moment
						times.add(datetime.strptime(row[column_number_time_mark_in_lists], '%m/%d/%Y, %I:%M:%S %p'))
					if name == organizer:
						organizer_count += 1
					else:
						participants.add(name)
		if organizer_count%2 != 1:
			raise UserError ('The meeting organizer should be the first participant listed in the file and should appear an odd number of times.')
		if len(participants) == 0:
			raise UserError ('The file is empty.')
		locale.setlocale(locale.LC_ALL, '')
		return (min(times), (max(times) - min(times)).seconds, participants)
	except UserError as exc:
		print ('Error reading ' + source + '.')
		print (exc)
		locale.setlocale(locale.LC_ALL, '')
		return None

def get_attendance_report (source):
	'''Parses the output of Ms Teams attendance report
	that you can download after the meeting from the meeting chat.
	Returns a tuple:
	earliest time recorded, duration, the list of participants, column header for names.
	The meeting organizer should be the first participant listed in the file.
	On the basis of this assumption the role of the first participant
	is taken to mean "Organizer" and any people with that role are
	excluded from the list.
	'''
	participants = set()
	times = set()
	organizer_count = 0
	awaiting_headers = True
	try:
		with open(source, newline='', encoding='utf-16') as csvfile:
			reader = csv.reader(csvfile, dialect='excel-tab')
			for row in reader:
				if len(row) == number_of_columns_in_reports:
					if awaiting_headers:
						awaiting_headers = False
						if row[column_number_full_name_in_reports] != column_header_full_name:
							raise UserError ('Column header ' + column_header_full_name + ' expected and ' + row[column_number_full_name_in_reports] + ' encountered.')
						if row[column_number_join_time_in_reports] != column_header_join_time:
							raise UserError ('Column header ' + column_header_join_time + ' expected and ' + row[column_number_join_time_in_reports] + ' encountered.')
						if row[column_number_leave_time_in_reports] != column_header_leave_time:
							raise UserError ('Column header ' + column_header_leave_time + ' expected and ' + row[column_number_leave_time_in_reports] + ' encountered.')
						if row[column_number_role_in_reports] != column_header_role:
							raise UserError ('Column header ' + column_header_role + ' expected and ' + row[column_number_role_in_reports] + ' encountered.')
					else:
						name = row[column_number_full_name_in_reports]
						try:
							times.add(datetime.strptime(row[column_number_join_time_in_reports], time_format))
							times.add(datetime.strptime(row[column_number_leave_time_in_reports], time_format))
						except ValueError:
							# Apparently Teams sometimes randomly switches to the US time format even on non-US systems
							locale.setlocale(locale.LC_TIME, 'en_US') # The locale change is local - will be reversed in a moment
							times.add(datetime.strptime(row[column_number_join_time_in_reports], '%m/%d/%Y, %I:%M:%S %p'))
							times.add(datetime.strptime(row[column_number_leave_time_in_reports], '%m/%d/%Y, %I:%M:%S %p'))
						if row[column_number_role_in_reports] != organizer_role_in_reports:
							participants.add(name)
		if awaiting_headers:
			raise UserError ('The number of data columns is apparently ' + str(len(row)) + '. I expected ' + str(number_of_columns_in_reports) + '.')
		if len(participants) == 0:
			raise UserError ('No entries found in the report.')
		locale.setlocale(locale.LC_ALL, '')
		return (min(times), (max(times) - min(times)).seconds, participants)
	except UserError as exc:
		print ('Error reading ' + source + '.')
		print (exc)
		locale.setlocale(locale.LC_ALL, '')
		return None

def get_all_attendance():
	meetings = []
	files_defined = False
	for root, dirs, files in os.walk(dir_downloads):
		files_defined = True
		break # Just get the files in the top dir.
	if files_defined:
		print('Reading the lists/reports:')
		for fname in files:
			if fname.endswith ('.csv') and 'meetingAttendanceList' in fname:
				print (fname)
				m = get_attendance_list (dir_downloads + '/' + fname)
			elif fname.endswith ('.csv') and 'meetingAttendanceReport' in fname:
				print (fname)
				m = get_attendance_report (dir_downloads + '/' + fname)
			else:
				m = None
			if m:
				meetings.append (m)
		return sorted(meetings)
	else:
		print ('Please, correct the downloads dir setting at the beginning of the script file.')
		print ('Present setting: ' + dir_downloads)
		exit()

def mark_attendance_on_list (fname, column_name, participants):
	'''Marks the attendance of a given set of participants on a list in a CSV file.
	If the file does not exist yet, an empty file is created.
	Participants not in the set are marked as absent.
	Participants not previously on the list are added to the list,
	marked as absent on previous meetings.
	Participation is marked under a given column heading (date or date+time).
	If such a column already exists in the list file, the set of present
	participants is merged with those previously marked as present.
	'''
	rows = []
	name_ind = {}
	col_ind = {}
	print ('Updating the file ' + fname + ', column ' + column_name + '.')
	try:
		with open(fname, newline='') as csvfile:
			reader = csv.reader(csvfile)
			for n,row in enumerate(reader):
				if n == 0:
					for k,entry in enumerate(row):
						col_ind[entry] = k
					rows.append(row)
				elif row[0] in name_ind:
					# Check for double entries in the existing list.
					m = name_ind[row[0]]
					for k in range(1, len(row)):
						rows[m][k] = max(rows[m][k], row[k])
				else:
					name_ind[row[0]] = len(rows)
					rows.append(row)
	except FileNotFoundError:
		rows = [[column_header_full_name]]
		name_ind = {}
		col_ind = {}
		print ('Creating a new, empty list: ' + fname)

	try:
		need_to_sort = False # Only sort the rows if new participants were added.
		for name in participants:
			if name not in name_ind:
				need_to_sort = True
				row = [name] + [0 for i in range(1, len(rows[0]))]
				name_ind[name] = len(rows)
				rows.append(row)

		if column_name not in col_ind:
			for n,row in enumerate(rows):
				if n == 0:
					col_ind[column_name] = len(row)
					row.append(column_name)
				else:
					row.append(0)

		for name in participants:
			rows[name_ind[name]][col_ind[column_name]] = 1

		if need_to_sort:
			rows[1:len(rows)] = sorted(rows[1:len(rows)], key=lambda row: locale.strxfrm(row[0]))

		with open(fname, 'w', newline='') as csvfile:
			writer = csv.writer(csvfile)
			for row in rows:
				writer.writerow(row)
	except Exception:
		print ('Error encountered when updating the file: ' + fname)


def distance_mod_week (delta):
	minutes = delta/timedelta(minutes=1)
	return abs((minutes+7*12*60) % (7*24*60) - 7*12*60)

if not irregular_classes:
	schedule = get_schedule('schedule.csv')
meetings = get_all_attendance()

if irregular_classes:
	for t, duration, participants in meetings:
		mark_attendance_on_list ('attendance.csv', str(t.date()) + t.strftime(' %H:%M'), participants)
elif schedule:
	for t, duration, participants in meetings:
		slot = len(schedule)
		distance = 24*60
		for (s, (day_of_week, meeting_time, group_name, suffix)) in enumerate(schedule):
			dist = distance_mod_week(t - meeting_time)
			if dist < distance:
				slot = s
				distance = dist
		if slot == len(schedule) or distance > max(duration, 30):
			print ('Classes not found on ' + str(t) + ' (' + t.strftime('%A') + ').')
			if slot != len(schedule):
				print ('The closest entry in the schedule is ' + str(distance) + ' minutes away: ' + schedule[slot][0] + ' ' + schedule[slot][2] + schedule[slot][3])
			print ('You may need to edit the schedule to add a missing entry.')
			print ('')
		else:
			mark_attendance_on_list (schedule[slot][2] + '.csv', str(t.date()) + schedule[slot][3], participants)
else:
	try:
		# If the schedule file is found, we do not do anything, just report and exit.
		with open('schedule.csv') as f:
			pass
		print ('The schedule file exists, but does not appear to contain valid data.')
		exit()
	except FileNotFoundError:
		pass
	# If there is no schedule file, create a new one.
	print('The schedule file does not exist.')
	if meetings:
		print('I am going to create one for you, on the basis of attendance lists found.')
		print('Please, correct and complete the schedule, supply proper group names, and')
		print('re-run to get attendance marked.')
		create_placeholder_schedule ('schedule.csv', meetings)
		exit()
	else:
		print('I could create a sample schedule for you, but I need some attendance lists')
		print('or reports as a starting point. If you do have some in your Downloads')
		print('folder, I could not read them properly. Please update the configuration')
		print('settings at the beginning of this script.')
		exit()
