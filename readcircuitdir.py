"""Reads the cicuit directory files exported as CSV and put them in
congregation and people address book csv file."""
from __future__ import print_function, unicode_literals
import glob
import logging
import os
import csv
import re

import openpyxl

from nameparser import HumanName
from address import AddressParser
import phonenumbers

CONGREGATION_CSV = "ca95_output-congregations.csv"
CONTACTS_CSV = "ca95_output-contacts.csv"
PTC_CSV = 'ca95_output-ptc.csv'


def striptelnum(tel):
	"""Strips the "C-", "H-" prefix or " H", " C" from the phone number"""
	if tel[:2] in ["C-", "H-"]:
		tel = tel[2:]
	if tel[-2:] in [" C", " H"]:
		tel = tel[:-2]
	return tel


def createcircuitcsv(inputfile, congregationsfile, contactsfile):
	logger = logging.getLogger('readcsvfiles')
	inputfilename = os.path.join('data', inputfile)
	logger.info('Reading congregation CSV file {}...'.format(inputfilename))
	with open(inputfilename, "rb") as csvfile:
		reader = csv.reader(csvfile)
		congregation = {}
		congregation["name_simple"] = reader.next()[0].strip()
		congregation["name"] = "\"Kingdom Hall - {}\"".format(congregation["name_simple"])
		congregation["street"] = reader.next()[0].strip()
		city, state_postal = reader.next()[0].split(",", 1)
		congregation["city"] = city.strip()
		state, postal = state_postal.rsplit(" ", 1)
		congregation["state"] = state.strip()
		congregation["postal"] = postal.strip()
		congregation["complete_addr"] = "\"{}, {}, {}\"".format(congregation["street"], city,
			state_postal)
		idx = 1
		# at least make #1 exist
		congregation["Phone 1 - Type"] = "Main"
		congregation["Phone 1 - Value"] = ""
		ismeetingsection = False
		for row in reader:
			if row[0].startswith("ELDERS"):
				break
			if row[0].strip().upper().startswith("MEETING"):
				ismeetingsection = True
				congregation["notes"] = "\"" + row[1].strip()
				continue
			if ismeetingsection:
				congregation["notes"] += ", {}".format(row[1].strip())
			else:
				congregation["Phone {} - Type".format(idx)] = "Main"
				congregation["Phone {} - Value".format(idx)] = row[0].strip()
				idx += 1
		# close the quote for the notes
		congregation["notes"] += "\""

		logger.info('Grabbing congregation elder information...')
		elders = []
		reader.next()
		try:
			for row in reader:
				if sum([len(each) for each in row]) == 0:
					continue
				row = [each.decode(errors="ignore") for each in row]
				if row and row[0].upper().startswith("MINISTERIAL"):
					break
				row.extend([each.decode(errors="ignore") for each in reader.next()])
				elders.append(row)
		except StopIteration:
			pass

		logger.info('Grabbing congregation servants information...')
		servants = []
		try:
			reader.next()
			for row in reader:
				if sum([len(each) for each in row]) == 0:
					continue
				row = [each.decode(errors="ignore") for each in row]
				if row[0].startswith("Rev ") or row[0].startswith("Revis"):
					break
				row.extend([each.decode(errors="ignore") for each in reader.next()])
				servants.append(row)
		except StopIteration:
			pass

		# merge contacts together
		ap = AddressParser()
		for role, contacts in zip(["Elder", "MS"], [elders, servants]):
			for contact in contacts:
				fullname = HumanName(contact[0])
				if fullname.last == "":
					continue
				fulladdr = ap.parse_address(contact[2])
				eachcontact = {}
				eachcontact["lastname"] = fullname.last
				eachcontact["firstname"] = fullname.first
				eachcontact["middlename"] = fullname.middle
				eachcontact["suffixname"] = fullname.suffix
				eachcontact["name"] = "\"" + "{}, {} {}".format(fullname.last, fullname.first,
					fullname.middle).strip() + "\""
				eachcontact["name_suffix"] = fullname.suffix
				eachcontact["complete_addr"] = "\"{}\"".format(contact[2])
				eachcontact["street"] = " ".join([each for each in [fulladdr.house_number,
					fulladdr.street_prefix, fulladdr.street, fulladdr.street_suffix] if each])
				eachcontact["city"] = fulladdr.city
				eachcontact["state"] = fulladdr.state
				eachcontact["postal"] = fulladdr.zip
				eachcontact["role"] = "\"{}\"".format(contact[3].strip())
				eachcontact["email1"] = contact[6].strip()
				eachcontact["home#"] = striptelnum(contact[1].strip())
				eachcontact["cell#"] = striptelnum(contact[5].strip())
				eachcontact["keywords"] = "\"JW,{}\"".format(role)
				notes = [role, congregation["name_simple"].split(",")[0].strip()]
				other_role = contact[4].strip()
				if other_role:
					notes.append(other_role)
				eachcontact["notes"] = "\"{}\"".format(",".join(notes))
				try:
					contactsfile.write("{firstname},{middlename},{lastname},{suffixname},,,"
						"{email1},{home#},{cell#},{complete_addr},US,,,,,,,{keywords},{notes}"
						"\n".format(**eachcontact))
				except UnicodeEncodeError:
					print("DEBUG: {} / {}".format(fullname.last, eachcontact))
					raise
	logger.info('Writing to congregation CSV file...')
	congregationsfile.write(",,,,,,,,,,,{name},{Phone 1 - Value},,,{complete_addr},US,\"KH,JW\","
		"{notes}\n".format(**congregation))
	logger.info('... Done.')


def extractcsv(xlsx, prefix):
	"""Extract all the congregation worksheet as CSV file.

	Args:
	- `xlsx` (`str`): Excel spreadsheet.  See Notes.

	Returns:
	- `list(str)`: list of filename

	Notes:
	- Workbook must not be password protected.  That means, the original
	workbook must be opened prior and save to another file without the password.
	The new workbook must have `-nopasswd` before the extension.  The workbook
	without password will be deleted after a successful call.  This will make
	sure that the workbook contents are kept secured in the original file only.
	"""
	logger = logging.getLogger('extractcsv')
	EXCLUDE_WORKSHEET = ['memo', 'cover', 'congs', 'revision log']
	logger.info('Opening workbook {}...'.format(xlsx))
	if '-nopasswd.' not in xlsx:
		errmsg = 'Filename should be in *-nopassword.xlsx format.'
		logger.error(errmsg)
		raise ValueError(errmsg)

	workbook = openpyxl.load_workbook(xlsx)
	worksheets = [each for each in workbook.sheetnames if each.lower() not in EXCLUDE_WORKSHEET]
	ret = []
	for sheetname in worksheets:
		fname = '{}{}.csv'.format(prefix, filter(lambda x: x in [chr(each)
			for each in range(ord('a'), ord('z') + 1)], sheetname.lower()))
		ret.append(fname)
		logger.info('Opening sheetname {}...'.format(sheetname))
		sheet = workbook[sheetname]
		logger.info('... creating csv file {}...'.format(fname))
		with open(os.path.join('data', fname), 'wb') as csvhandle:
			csvobj = csv.writer(csvhandle)
			for row in sheet.rows:
				newrow = []
				for cell in row:
					if cell.value is None:
						newrow.append('')
					elif isinstance(cell.value, (unicode, str)):
						newrow.append(cell.value.encode('utf-8').strip())
					else:
						newrow.append(cell.value)
				csvobj.writerow(newrow)
				# csvobj.writerow(['' if cell.value is None else cell.value.encode('utf-8').strip()
				# 	for cell in row])
	workbook.close()
	logger.info('Done.')
	return ret


def is_eldersrow(row):
	"""Check if CSV row is an elder's row."""
	return row[0].startswith('ELDERS')


def is_msrow(row):
	"""Check if CSV row is an MS row."""
	return row[0].startswith('MINISTERIAL SERVANTS')


def is_newsection(row):
	"""Check if CSV in new section"""
	cell = row[0].lower()
	return cell.startswith('elders') or cell.startswith('ministerial ') or \
		cell.startswith('revised') or cell.startswith('rev ')


def decoderow(row):
	"""Decode all rows string data."""
	return [each.decode(errors='ignore').strip() for each in row]


def is_empty(row):
	"""Checks if row is empty"""
	return sum([len(each) for each in row]) == 0


def get_contacttype(key, value):
	"""Check the cell data to see type of contact information.

	Args:
	- `key` (`str`): self-explained
	- `value` (`str`): self-explained

	Returns:
	- `str`: key for the dictionary
	"""
	# logger = logging.getLogger('get_contacttype')
	key = key.lower()
	if key in ['cell', 'home', 'mobile'] and value is not None:
		# logger.debug('Phone ... {!r} / {!r}'.format(key, value))
		return key + '#', striptelnum(phonenumbers.format_number(phonenumbers.parse(value, 'US'),
			phonenumbers.PhoneNumberFormat.NATIONAL))
	return key, value


def get_ptcdata(row, reader):
	"""Retrieves the PTC data.

	Args:
	- `row` (`csv row object`): current row object
	- `reader` (`csv object`): self-explained

	Returns:
	- `tuple` (`?`)
	"""
	row = decoderow(row)
	name = HumanName(filter(lambda x: x != '*', row[0]))
	data = {'fullname': '"{}"'.format(name.full_name), 'lastname': name.last,
		'firstname': name.first, 'middlename': name.middle, 'suffixname': name.suffix,
		'complete_addr': '', 'role': '', 'home#': '',
		'name': '\"{}, {} {}\"'.format(name.last, name.first, name.middle),
		'talks': filter(lambda x: x != '"', row[3]).strip(), 'cell#': ''}
	if row[2] not in [None, 'None', '']:
		key, value = get_contacttype(*row[1:3])
		data[key] = value
	for row in reader:
		row = decoderow(row)
		if is_empty(row):
			break
		if row[0]:
			# this will overwrite role data if it has one already
			data['role'] = '"{}"'.format(filter(lambda x: x not in ['(', ')'], row[0]).upper())
		if row[1]:
			if row[2] not in [None, 'None', '']:
				key, value = get_contacttype(*row[1:3])
				data[key] = value
		if row[3]:
			if len(data['talks']) > 0 and data['talks'][-1] != ',':
				data['talks'] += ', '
			data['talks'] += filter(lambda x: x != '"', row[3])
	return data


def createptccsv(inputfile, ptcfile):
	"""Creates a PTC CSV file.

	Args:
	- `inputfile` (`str`): input file name
	- `ptcfile` (`file handle`): handle of the output PTC file
	"""
	logger = logging.getLogger('createptccsv')
	inputfilename = os.path.join('data', inputfile)
	logger.info('Reading PTC CSV file {}...'.format(inputfilename))
	congregation = re.match('ca95talks-(.+).csv', inputfile).group(1)
	with open(inputfilename, 'rb') as csvfile:
		reader = csv.reader(csvfile)
		for row in reader:
			if is_eldersrow(row):
				break

		# read the elders data
		for privilege in ['Elder', 'MS']:
			for row in reader:
				row = decoderow(row)
				if is_empty(row):
					continue
				elif is_newsection(row):
					break
				data = get_ptcdata(row, reader)
				data['keywords'] = '"JW,{}"'.format(privilege)
				# if data['role']:
				# 	data['notes'] = '"{}"'.format(','.join([data['role'], 'Congregation: ' + congregation]))
				# else:
				# 	data['notes'] = 'Congregation: ' + congregation
				if 'email1' not in data:
					data['email1'] = '' if 'email' not in data else data['email']
				ptcfile.write('{fullname},{firstname},{middlename},{lastname},{suffixname},'
					'{role},{congregation},{email1},{home#},{cell#},"{talks}",'
					'{keywords},\n'.format(congregation=congregation, **data))
	logger.info('... Done')


def main():
	"""Main program"""
	# outlook CSV fields:
	# "First Name, Middle Name, Last Name, Title, Suffix, Location, E-mail Address, Home Phone,"
	# 	"Mobile Phone, Home Street, Home City, Home State, Home Postal Code, Home Country,"
	# 	"Company, Business Phone, Job Title, Department, Business Street, Business City,"
	# 	"Business State, Business Postal Code, Business Country, Keywords, Notes"

	# First Name	Middle Name	Last Name	Title	Suffix	Initials	Web Page	Gender
	# Birthday	Anniversary	Location	Language	Internet Free Busy	Notes	E-mail Address
	# E-mail 2 Address	E-mail 3 Address	Primary Phone	Home Phone	Home Phone 2
	# Mobile Phone	Pager	Home Fax	Home Address	Home Street	Home Street 2
	# Home Street 3	Home Address PO Box	Home City	Home State	Home Postal Code
	# Home Country	Spouse	Children	Manager's Name	Assistant's Name	Referred By
	# Company Main Phone	Business Phone	Business Phone 2	Business Fax
	# Assistant's Phone	Company	Job Title	Department	Office Location	Organizational ID Number
	# Profession	Account	Business Address	Business Street	Business Street 2
	# Business Street 3	Business Address PO Box	Business City	Business State
	# Business Postal Code	Business Country	Other Phone	Other Fax	Other Address
	# Other Street	Other Street 2	Other Street 3	Other Address PO Box	Other City
	# Other State	Other Postal Code	Other Country	Callback	Car Phone	ISDN
	# Radio Phone	TTY/TDD Phone	Telex	User 1	User 2	User 3	User 4	Keywords
	# Mileage	Hobby	Billing Information	Directory Server	Sensitivity	Priority
	# Private	Categories

	# google CSV fields
	# "Name,Name Suffix,Location,Occupation,E-mail 1 - Type,"
	# 				"E-mail 1 - Value,E-mail 2 - Type,E-mail 2 - Value,E-mail 3 - Type,"
	# 				"E-mail 3 - Value,Phone 1 - Type,Phone 1 - Value,Phone 2 - Type,"
	# 				"Phone 2 - Value,Phone 3 - Type,Phone 3 - Value,Phone 4 - Type,"
	# 				"Phone 4 - Value,Address 1 - Type,Address 1 - Formatted,Address 1 - Street,"
	# 				"Address 1 - City,Address 1 - Region,"
	# 				"Address 1 - Postal Code,Address 1 - Country,Organization 1 - Type,"
	# 				"Organization 1 - Name,Organization 1 - Title,Organization 1 - Department,"
	# 				"Organization 1 - Location,Organization 1 - Job Description\n"
	logger = logging.getLogger()
	logger.setLevel(logging.DEBUG)
	handler = logging.StreamHandler()
	handler.setLevel(logging.DEBUG)
	logger.addHandler(handler)

	xlsxs = [glob.glob(os.path.join('data', 'CA-95-*-nopasswd.xlsx'))[0],
		glob.glob(os.path.join('data', 'Approved*-nopasswd.xlsx'))[0]]
	congregations = extractcsv(xlsxs[0], 'ca95-')
	ptcs = extractcsv(xlsxs[1], 'ca95talks-')

	if False:
		congregation_csv = os.path.join('data', CONGREGATION_CSV)
		contacts_csv = os.path.join('data', CONTACTS_CSV)
		logger.info('Creating {}...'.format(congregation_csv))
		with open(congregation_csv, "w") as congregationsfile:
			logger.info('Creating {}...'.format(contacts_csv))
			with open(contacts_csv, "w") as contactsfile:
				for outfile in [congregationsfile, contactsfile]:
					outfile.write("First Name,Middle Name,Last Name,Suffix,Title,Location,"
						"E-mail Address,Home Phone,Mobile Phone,Home Address,"
						"Home Country,Company,Business Phone,Job Title,Department,"
						"Business Address,Business Country,Keywords,Notes\n")
				for eachfile in congregations:
					createcircuitcsv(eachfile, congregationsfile, contactsfile)

	ptc_csv = os.path.join('data', PTC_CSV)
	logger.info('Creating {}...'.format(ptc_csv))
	with open(ptc_csv, 'w') as ptcfile:
		ptcfile.write('Full Name,First Name,Middle Name,Last Name,Suffix,Title,Location,'
			'E-mail address,Home Phone,Mobile Phone,Talks,Notes\n')
		for eachfile in ptcs:
			createptccsv(eachfile, ptcfile)

	for xlsx in xlsxs:
		logger.info('Deleting {}...'.format(xlsx))
		os.unlink(xlsx)

if __name__ == "__main__":
	main()
