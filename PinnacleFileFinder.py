"""
	
	Name	: PinnacleFileFinder.py

	Usage	: PinnacleFileFinder.py -h

	Author	: Bar Harel

	Description:
		- Takes a .AXP file and creates a list of all the used files in that pinnacle project with their order and time of appearance.
		- The list can be output as a text file or .csv for use with programs like Excel

	Todo:
		- Add all possible file formats

	Changelog:
		- 06/03/15 - GitHub :-)
		- 21/02/15 - Creation
"""

import re, argparse, os, csv

# The encoding pinnacle studio uses
PINNACLE_ENCODING = "utf_16_le"

# File formats
FILE_FORMATS = "jpg|JPG|MOV|mov|png|PNG|avi|AVI"

# Unicode RE format for getting the time and name
TIME_NAME_RE = ur"RecIn=\".*?\(([^>]*?)\).*?<Name>([^\n]+?\.(?:%s))</Name>" % (FILE_FORMATS)

# Default output paths
CSV_DEFAULT = r".\PinnacleFiles.csv"
TXT_DEFAULT = r".\PinnacleFiles.txt"

# Max name for file
MAX_FILE_NAME = 100

def convert_time(seconds_as_float):
	"""
		Function 	: convert_time(seconds_as_float) --> hours, minutes, seconds, ms
		
		Purpose:
			- Convert the time from seconds to an hour, minute, second, ms tupple
	"""
	# Conversion
	seconds, ms = divmod(seconds_as_float,1)
	minutes, seconds = divmod(seconds, 60)
	hours, minutes = divmod(minutes, 60)

	return hours, minutes, seconds, ms

def output_file(findings_list, output_format, output_path):
	"""
		Function 	: output_txt(findings_list, output_format, output_path) --> NoneType
		
		Purpose:
			- Output the file in the specified format
	"""
	# Txt output
	if output_format == "txt":
		
		# The final string used to store the formatted file
		final_str = u"Pinnacle studio file list:\n"

		# Set a counter for the files
		counter = 1

		# Go over the findings
		for appearance_time, file_name in findings_list:
			
			# Failsafe in case of false positive matching
			if len(file_name) > MAX_FILE_NAME:
				continue
			
			# Convert time to hours, mintes, seconds, ms
			try:
				hours, minutes, seconds, ms = convert_time(float(appearance_time))

			# In case of conversion errors
			except ValueError as err:
				continue

			# The time string
			time_str = "%02d:%02d:%02d.%s" % (hours, minutes, seconds, str(ms)[2:])

			# Format the output
			final_str += u"%d: %-25s \tat %02d:%02d:%02d.%s\n" % (counter, file_name, hours, minutes, seconds, str(ms)[2:])

			# Increase counter
			counter += 1

		# Write the result to the output file
		try:
			with open(output_path,"w") as my_file:
				my_file.write(final_str)
		except IOError as err:
			print "Error opening or writing to the output file."

	# CSV output
	elif output_format == "csv":
		try:
			with open(output_path,"wb") as my_file:
				
				# Generate the csv file writer
				file_writer = csv.writer(my_file)

				# Go over the findings
				for appearance_time, file_name in findings_list:
					
					# Failsafe in case of false positive matching
					if len(file_name) > MAX_FILE_NAME:
						continue
					
					# Convert time to hours, mintes, seconds, ms
					try:
						hours, minutes, seconds, ms = convert_time(float(appearance_time))

					# In case of conversion errors
					except ValueError as err:
						continue

					# The time string
					time_str = "%02d:%02d:%02d.%s" % (hours, minutes, seconds, str(ms)[2:])

					# Output the row
					file_writer.writerow([file_name, time_str])

		except IOError, csv.Error:
			print "Error opening or writing to the output file."
	
	else:
		print "ERROR: Invalid output format"

def main():
	"""
		Function: main() --> NoneType
		
		Purpose:
			- Control the flow of the program
	"""
	# Parse arguments
	parser = argparse.ArgumentParser(description="Find file names and time of appearance from a pinnacle studio AXP file.")
	parser.add_argument("axp_file", help="Path to the .axp file")
	parser.add_argument("-o", "--output_file", help=("Output file, defaults to '%s' in case of txt and '%s' in case of csv" % (TXT_DEFAULT, CSV_DEFAULT)))
	parser.add_argument("-csv", help="Output the file in csv format.", action="store_true")
	args = parser.parse_args()

	# Check if input file exists
	if not os.path.exists(args.axp_file):
		print "ERROR: Invalid input path."
		return

	# Check the extension
	if args.axp_file[-4:].lower() != ".axp":
		print "Error: Not a .axp file"
		return

	# Unicode RE for getting the time and name
	try:
		time_name_re = re.compile(TIME_NAME_RE, re.S|re.U)
	except re.error as err:
		print "ERROR: Bad input RE."
		return

	# Open and read from the file
	try:
		with open(args.axp_file, "r") as input_file:
			input_str = input_file.read()
	except IOError as err:
		print "Error opening or reading from input file."
		return

	# Decode using the pinnacle studio encoding
	input_str = input_str.decode(PINNACLE_ENCODING)

	# Find the re matches in the string
	findings = time_name_re.findall(input_str)

	# Check the specified output format
	output_format = "csv" if args.csv else "txt"

	# Check output file path
	if args.output_file is None:
		output_path = CSV_DEFAULT if args.csv else TXT_DEFAULT
	else:
		output_path = args.output_file

	output_file(findings, output_format, output_path)

if __name__ == "__main__":
	main()