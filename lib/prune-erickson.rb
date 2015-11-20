#!/usr/bin/env ruby
require 'roo'
require 'spreadsheet'
require "csv"
require 'colorize'
# require 'fileUtils'


# ------------------
# FUNCTIONS
# ------------------


			
def labelChecker (row, testvar)

	row.each_with_index{|val, index| 

		unless val != "OppCommunity" and val !="FirstName" and val != "LastName" and val != "Address" and val!="City" and val != "State" and val != "Zipcode" and val != "Phone" and val != "Email" and val != "ILorCC"
			testvar.push(index)
		end
	}
end

def fieldChecker (row)
	unless row.include? "OppCommunity"	
		print "you are missing the OppCommunity field\n".yellow.blink
	else
	end
	unless row.include? "FirstName"	
		print "you are missing the FirstName field\n".yellow.blink
	end
	unless row.include? "LastName"	
		print "you are missing the LastName field\n".yellow.blink
	end
	unless row.include? "Phone"	
		print "you are missing the phone field\n".yellow.blink
	end
	unless row.include? "Address"	
		print "you are missing the Address field\n".yellow.blink
	end
	unless row.include? "City"	
		print "you are missing the City field\n".yellow.blink
	end
	unless row.include? "State"	
		print "you are missing the State field\n".yellow.blink
	end
	unless row.include? "Zipcode"	
		print "you are missing the Zipcode field\n".yellow.blink
	end
	unless row.include? "Address"	
		print "you are missing the Address field\n".yellow.blink
	end
	unless row.include? "Address"	
		print "you are missing the Address field\n".yellow.blink
	end
	unless row.include? "Email"	
		print "you are missing the Email field\n".yellow.blink
	end
end

def processData (sheet, newvar, testvar)
	
	newvar.push "zero"
	sheet.each do |row|
		row.each_with_index {|val, index| 
			if testvar.include? index
				
				newvar.push val
			end
		}
	end
	print "~~>writing new file \n\n"
	
	newvar.map! { |x| 

		if x == "OppCommunity"
			print "renaming field #{x} \n".colorize(:light_blue)
			x = "Community of Interest"
		elsif x == "FirstName"
			print "renaming field #{x} \n".colorize(:light_blue)
			x = "First Name"	
		elsif x == "LastName"
			print "renaming field #{x} \n".colorize(:light_blue)
			x = "Last Name"
		elsif x == "Address"
			print "renaming field #{x} \n".colorize(:light_blue)
			x = "Address 1"
		elsif x == "city"
			print "renaming field #{x} \n".colorize(:light_blue)
			x = "city"
		elsif x == "State"
			print "renaming field #{x} \n".colorize(:light_blue)
			x = "State or Province"
		elsif x == "Zipcode"
			print "renaming field #{x} \n".colorize(:light_blue)
			x = "Zip or Postal Code"
		elsif x == "Email"
			print "renaming field #{x} \n".colorize(:light_blue)
			x = "Email Address"
		elsif x == "Phone"
			print "renaming field #{x} \n".colorize(:light_blue)
			x = "Home Phone"
		else
			x=x
		end
	}
end

def printFile (newvar, testvar, output, il_or_cc, k=0, l=1,offset=0)
	# print newvar
	newBook = Spreadsheet::Workbook.new
	newSheet = newBook.create_worksheet
	if il_or_cc == ""
		testvar.each do |x|
			newSheet.row(0).push x
		end
	end
	for i in k..newvar.length
		for j in l..testvar.length-1
				newSheet.row(i+offset).push newvar[(i*testvar.length)+j]
		end
		if il_or_cc != ""
			if i > 0
				newSheet.row(i).push "#{il_or_cc}"

			else
				newSheet.row(i).push "ILorCC"
			end
		end
	end
	print "output file is in same file under out.xls to open, use \n \n open #{output}"
	print "\n"
	newBook.write output
end

def com_separated_files (output)
	puts output
	labels = Array.new
	apl_array = Array.new
	ach_array = Array.new
	bbv_array = Array.new
	ccv_array = Array.new
	cci_array = Array.new
	dvf_array = Array.new
	eth_array = Array.new
	frv_array = Array.new
	gsv_array = Array.new
	hsd_array = Array.new
	lhn_array = Array.new
	lph_array = Array.new
	mgc_array = Array.new
	ocv_array = Array.new
	rwv_array = Array.new
	sbv_array = Array.new
	tck_array = Array.new
	wcd_array = Array.new
	
	original = Spreadsheet.open(output)
	original_sheet = original.worksheet(0)
	
	
	rowlength = original_sheet.row(0)
	puts rowlength.length
	# make a directory for the new files
	unless File.directory?("com_separated_files")
		%x(mkdir com_separated_files)
	end
	original_sheet.each do |row|
		if row[0] == "apl".upcase
			
			# print row to other array and print label to a community array
			row.each do |x|
				apl_array.push x
			end
			unless labels.include?("apl".upcase)
				labels.push "apl".upcase
			end
		end
		if row[0] == "ach".upcase
			
			# print row to other array and print label to a community array
			row.each do |x|
				ach_array.push x
			end
			unless labels.include?("ach".upcase)
				labels.push "ach".upcase
			end
		end
		if row[0] == "bbv".upcase
			
			# print row to other array and print label to a community array
			row.each do |x|
				bbv_array.push x
			end
			unless labels.include?("bbv".upcase)
				labels.push "bbv".upcase
			end
		end
		if row[0] == "ccv".upcase
			
			# print row to other array and print label to a community array
			row.each do |x|
				ccv_array.push x
			end
			unless labels.include?("ccv".upcase)
				labels.push "ccv".upcase
			end
		end
		if row[0] == "cci".upcase
			
			# print row to other array and print label to a community array
			row.each do |x|
				cci_array.push x
			end
			unless labels.include?("cci".upcase)
				labels.push "cci".upcase
			end
		end
		if row[0] == "dvf".upcase
			
			# print row to other array and print label to a community array
			row.each do |x|
				dvf_array.push x
			end
			unless labels.include?("dvf".upcase)
				labels.push "dvf".upcase
			end
		end
		if row[0] == "eth".upcase
			
			# print row to other array and print label to a community array
			row.each do |x|
				eth_array.push x
			end
			unless labels.include?("eth".upcase)
				labels.push "eth".upcase
			end
		end
		if row[0] == "frv".upcase
			
			# print row to other array and print label to a community array
			row.each do |x|
				frv_array.push x
			end
			unless labels.include?("frv".upcase)
				labels.push "frv".upcase
			end
		end
		if row[0] == "gsv".upcase
			
			# print row to other array and print label to a community array
			row.each do |x|
				gsv_array.push x
			end
			unless labels.include?("gsv".upcase)
				labels.push "gsv".upcase
			end
		end
		if row[0] == "hsd".upcase
			
			# print row to other array and print label to a community array
			row.each do |x|
				hsd_array.push x
			end
			unless labels.include?("hsd".upcase)
				labels.push "hsd".upcase
			end
		end
		if row[0] == "lhn".upcase
			
			# print row to other array and print label to a community array
			row.each do |x|
				lhn_array.push x
			end
			unless labels.include?("lhn".upcase)
				labels.push "lhn".upcase
			end
		end
		if row[0] == "lph".upcase
			
			# print row to other array and print label to a community array
			row.each do |x|
				lph_array.push x
			end
			unless labels.include?("lph".upcase)
				labels.push "lph".upcase
			end
		end
		if row[0] == "mgc".upcase
			
			# print row to other array and print label to a community array
			row.each do |x|
				mgc_array.push x
			end
			unless labels.include?("mgc".upcase)
				labels.push "mgc".upcase
			end
		end
		if row[0] == "ocv".upcase
			
			# print row to other array and print label to a community array
			row.each do |x|
				ocv_array.push x
			end
			unless labels.include?("ocv".upcase)
				labels.push "ocv".upcase
			end
		end
		if row[0] == "rwv".upcase
			
			# print row to other array and print label to a community array
			row.each do |x|
				rwv_array.push x
			end
			unless labels.include?("rwv".upcase)
				labels.push "rwv".upcase
			end
		end
		if row[0] == "sbv".upcase
			
			# print row to other array and print label to a community array
			row.each do |x|
				sbv_array.push x
			end
			unless labels.include?("sbv".upcase)
				labels.push "sbv".upcase
			end
		end
		if row[0] == "tck".upcase
			
			# print row to other array and print label to a community array
			row.each do |x|
				tck_array.push x
			end
			unless labels.include?("tck".upcase)
				labels.push "tck".upcase
			end
		end
		if row[0] == "wcd".upcase
			
			# print row to other array and print label to a community array
			row.each do |x|
				wcd_array.push x
			end
			unless labels.include?("wcd".upcase)
				labels.push "wcd".upcase
			end
		end
	end
	print "creating files for these communities:\n".blue
	puts "#{labels}".blue
	print rowlength
	if apl_array.length > 0
		printFile(apl_array, rowlength, "com_separated_files/apl.xls", "", 0, 0,1)
	end
	if ach_array.length > 0
		printFile(ach_array, rowlength, "com_separated_files/ach.xls", "", 0, 0,1)
	end
	if bbv_array.length > 0
		printFile(bbv_array, rowlength, "com_separated_files/bbv.xls", "", 0, 0,1)
	end
	if ccv_array.length > 0
		printFile(ccv_array, rowlength, "com_separated_files/ccv.xls", "", 0, 0,1)
	end
	if cci_array.length > 0
		printFile(cci_array, rowlength, "com_separated_files/cci.xls", "", 0, 0,1)
	end
	if dvf_array.length > 0
		printFile(dvf_array, rowlength, "com_separated_files/dvf.xls", "", 0, 0,1)
	end
	if eth_array.length > 0
		printFile(eth_array, rowlength, "com_separated_files/eth.xls", "", 0, 0,1)
	end
	if frv_array.length > 0
		printFile(frv_array, rowlength, "com_separated_files/frv.xls", "", 0, 0,1)
	end
	if gsv_array.length > 0
		printFile(gsv_array, rowlength, "com_separated_files/gsv.xls", "", 0, 0,1)
	end
	if hsd_array.length > 0
		printFile(hsd_array, rowlength, "com_separated_files/hsd.xls", "", 0, 0,1)
	end
	if lhn_array.length > 0
		printFile(lhn_array, rowlength, "com_separated_files/lhn.xls", "", 0, 0,1)
	end
	if lph_array.length > 0
		printFile(lph_array, rowlength, "com_separated_files/lph.xls", "", 0, 0,1)
	end
	if mgc_array.length > 0
		printFile(mgc_array, rowlength, "com_separated_files/mgc.xls", "", 0, 0,1)
	end
	if ocv_array.length > 0
		printFile(ocv_array, rowlength, "com_separated_files/ocv.xls", "", 0, 0,1)
	end
	if rwv_array.length > 0
		printFile(rwv_array, rowlength, "com_separated_files/rwv.xls", "", 0, 0,1)
	end
	if sbv_array.length > 0
		printFile(sbv_array, rowlength, "com_separated_files/sbv.xls", "", 0, 0,1)
	end
	if tck_array.length > 0
		printFile(tck_array, rowlength, "com_separated_files/tck.xls", "", 0, 0,1)
	end
	if wcd_array.length > 0
		printFile(wcd_array, rowlength, "com_separated_files/wcd.xls", "", 0, 0,1)
	end


end

# --------------------------------------------
# GRABS INFORMATION FROM COMMAND LINE AND MAKES DATA AVAILABLE FOR PROGRAM
fileName = ARGV[0]
sheetNum = ARGV[1].to_i-1
output = ARGV[2]
print "\n\n"
print "Is this an email campaign for Continuing Care? (CC) [y/n]".yellow

res = STDIN.gets.strip
print "\n\n"
if res == 'y' || res == 'Y'
	il_or_cc = 'CC'
else
	il_or_cc = "IL"
end
print "Do you wish to create separate files based on community names?[y/n]".yellow
comres= STDIN.gets.strip

# --------------------------------------------
# ASKING FOR HELP
if ARGV.length == 1 and ARGV[0] == "help"
	print "
	THANK YOU FOR USING PRUNE!! /_(O__O)_/"
	print "
   \n
             ..     ...
	    |  |   /  /
	    |  |  /  /
	    |  | /  /
	    |  |/  ;-._ 
	    }  ` _/  / ;
	    |  /` ) /  /
	    | /  /_/_ /| \s
	    |/  /      |
	    (  ' \ '-  |
	     \    `.  /
	      |      |
	      |      | \n\n\n".colorize(:light_blue)
	puts "to edit a file use this format:"
	puts "FOR .xls \n".yellow

	puts "prune-erickson [filename.xls (string)] [sheet-number (number)] [output.xls (string)]\n\n\n"
	puts "FOR .csv \n".yellow
	puts "prune-erickson [filename.csv (string)] [output.xls (string)]\n\n\n"
	puts "Make sure to cd into your correct directory where the file is located.
	the output file will be in the same directory \n\n\n".yellow
elsif ARGV.length == 1 and (ARGV[0] == "-v" or ARGV[0] == "-version" or ARGV[0] == "version")
	print "\n
				-------------------
				| VERSION | 0.2.1 |
				-------------------
"
print "
		~~~~~~~THIS SOFTWARE IS IN ACTIVE DEVELOPMENT~~~~~
		Keep in mind that this project is 0.0.* and is still
		continuously being updated and user tested. 
		documentation, options and features are subject to change. 

		Stay tuned for the official release 1.0.0 which could be any
		day now.
	".yellow
# --------------------------------------------
# if the user input has to do with the help or version options this code is run
elsif ARGV.length == 1 and ARGV[0] != "version" or ARGV[0] != "help" or ARGV[0] != '-version' or ARGV[0] != "-v"  
	# line 68 catches any user input that relates to .xls files in the form of 
	# prune-erickson yourfile.xls [digit] outputfile.xls
	if ARGV[0] =~ /\S+.xls\b/ and ARGV[1] =~ /\d/ and ARGV[2] =~ /\S+.xls\b/
		
		Spreadsheet.client_encoding = 'UTF-8'
		print "~>reading file you have requested"
		print "\n"
		# --------------------------------------------
		# reads file
		#checking if the file is not nil and then creating a new workbook to copy necessary info into.
		if File.exist?(fileName.to_s)
			book = Spreadsheet.open fileName
			sheet1 = book.worksheet sheetNum
			newBook = Spreadsheet::Workbook.new
			newSheet = newBook.create_worksheet

			newSheet.name = "output"
			testvar = Array.new
			newvar = Array.new
			print "~>looking for unnecessary information"
			print "\n"
			if sheet1 == nil
				print "\n
				(_(T__T)_/ \n\n
				this sheet seems to be empty...
				are you selecting the wrong one? \n".red
			else
				labelChecker(sheet1.row(0), testvar)
				fieldChecker(sheet1.row(0))
				processData(sheet1, newvar, testvar)
				printFile(newvar, testvar, output, il_or_cc)
				if comres == 'y' || comres == 'Y'
					com_separated_files(output)
				end
			end
		else
			print "\n\n
			WHOOPS!!!

			It looks like this file is in the wrong folder or does not exsist. 
			try locating the folder and inputing the correct path\n\n".red
		end	
	elsif ARGV[0] =~ /\S+[.csv*]/ and ARGV[1] =~ /\S+.xls\b/
		output = ARGV[1]
		if File.exist?(fileName)
			file_data = CSV.read(fileName)
			newBook = Spreadsheet::Workbook.new
			newSheet = newBook.create_worksheet

			newSheet.name = "output"

			testvar = Array.new
			newvar = Array.new
			print "~>looking for unnecessary information"
			print "\n"

			labelChecker(file_data[0], testvar)
			fieldChecker(file_data[0])
			processData(file_data, newvar, testvar)
			printFile(newvar, testvar, output, il_or_cc)
			if comres == 'y' || comres == 'Y'
				com_separated_files(output)
			end
		else
			print "\n\n
			WHOOPS!!!

			It looks like this file is in the wrong folder or does not exsist. 
			try locating the folder and inputing the correct path\n\n".red
		end	
	elsif fileName =~ /\S+.xlsx/ and ARGV[1] =~ /\d/ and ARGV[2] =~ /\S+.xls\b/
		
		unless File.exist?(fileName)
			print "\n\n
		WHOOPS!!!

		It looks like this file is in the wrong folder or does not exsist. 
		try locating the folder and inputing the correct path\n\n".red
		else

			xlsx = Roo::Spreadsheet.open(fileName)
			newBook = Spreadsheet::Workbook.new
			newSheet = newBook.create_worksheet
			xlsxSheet = xlsx.sheet(sheetNum)
			if xlsxSheet == nil
				print "\n
				(_(T__T)_/ \n\n
				this sheet seems to be empty...
				are you selecting the wrong one? \n".red
			else
				newSheet.name = "output"
				header = xlsx.sheet(sheetNum).row(1)
				# take first row, and loop through the items in the array
				# sort indexes that are not necessary in the array
				testvar = Array.new
				newvar = Array.new				
				labelChecker(header, testvar)
				fieldChecker(header)
				processData(xlsxSheet, newvar, testvar)
				printFile(newvar, testvar, output, il_or_cc)
				if comres == 'y' || comres == 'Y'
					com_separated_files(output)
				end
			end
		end
	else
		print "\n 
		Are you trying to use PRUNE? 
		why dont you try this command:\n\n
		prune-erickson help \n\n\n
		"
	end
end