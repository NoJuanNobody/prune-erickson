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

def printFile (newvar, testvar, newSheet, newBook, output)
	# print newvar
	for i in 0..newvar.length
		for j in 1..testvar.length
				
		
				newSheet.row(i).push newvar[(i*testvar.length)+j]
		end
		if i > 0
			newSheet.row(i).push "IL"
		else
			newSheet.row(i).push "ILorCC"
		end
	end
	print "output file is in same file under out.xls to open, use \n \n open #{output}"
	print "\n"
	newBook.write output
end

# --------------------------------------------
# GRABS INFORMATION FROM COMMAND LINE AND MAKES DATA AVAILABLE FOR PROGRAM
fileName = ARGV[0]
sheetNum = ARGV[1].to_i-1
output = ARGV[2]
if sheetNum > 0	
	print fileName
	print "\n"

	print sheetNum
end	
print "\n\n"
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
				printFile(newvar, testvar, newSheet, newBook, output)
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
			printFile(newvar, testvar, newSheet, newBook, output)
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
				printFile(newvar, testvar, newSheet, newBook, output)
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