"""Reads the cicuit directory files exported as CSV and put them in
congregation and people address book csv file."""
from __future__ import print_function, unicode_literals
import csv

from nameparser import HumanName
from address import AddressParser

CONGREGATIONS = ["ca95-appianway.csv", "ca95-bailey.csv", "ca95-berryessa.csv",
	"ca95-crystalsprings.csv", "ca95-folsomlake.csv", "ca95-goldengate.csv", "ca95-hallave.csv",
	"ca95-lakeelizabeth.csv", "ca95-lawlerranchpkwy.csv", "ca95-marconi.csv", "ca95-north.csv",
	"ca95-northbay.csv", "ca95-paradisevalleydr.csv", "ca95-pioneerway.csv", "ca95-russpark.csv",
	"ca95-sierra.csv", "ca95-southshore.csv", "ca95-stonelake.csv", "ca95-story.csv",
	"ca95-westcloverrd.csv", "ca95-westlake.csv"]
CONGREGATION_CSV = "ca95-congregations.csv"
CONTACTS_CSV = "ca95-contacts.csv"


def striptelnum(tel):
	"""Strips the "C-", "H-" prefix or " H", " C" from the phone number"""
	if tel[:2] in ["C-", "H-"]:
		tel = tel[2:]
	if tel[-2:] in [" C", " H"]:
		tel = tel[:-2]
	return tel


def readcsvfiles(inputfile, congregationsfile, contactsfile):
	with open(inputfile, "rb") as csvfile:
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
	congregationsfile.write(",,,,,,,,,,,{name},{Phone 1 - Value},,,{complete_addr},US,\"KH,JW\","
		"{notes}\n".format(**congregation))


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

	with open(CONGREGATION_CSV, "w") as congregationsfile:
		with open(CONTACTS_CSV, "w") as contactsfile:
			for outfile in [congregationsfile, contactsfile]:
				outfile.write("First Name,Middle Name,Last Name,Suffix,Title,Location,"
					"E-mail Address,Home Phone,Mobile Phone,Home Address,"
					"Home Country,Company,Business Phone,Job Title,Department,"
					"Business Address,Business Country,Keywords,Notes\n")

			for eachfile in CONGREGATIONS:
				readcsvfiles(eachfile, congregationsfile, contactsfile)

if __name__ == "__main__":
	main()
