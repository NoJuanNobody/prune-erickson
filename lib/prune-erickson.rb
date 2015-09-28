require 'spreadsheet'


class Prune
	def self.run(fileName,sheetNum,output)


		# --------------------------------------------
		# GRABS INFORMATION FROM COMMAND LINE AND MAKES DATA AVAILABLE FOR PROGRAM
		# fileName = ARGV[0]
		# sheetNum = ARGV[1].to_i-1
		# output = ARGV[2]
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
			THANK YOU FOR USING PRUNE!! /_(O__O)_/ 
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
			      |      | \n\n\n
			to edit a file use this format:

			prune-erickson [filename.xls (string)] [sheet-number (number)] [output.xls (string)]\n\n\n\n
			"
		else
		# --------------------------------------------
		Spreadsheet.client_encoding = 'UTF-8'
		print "~>reading file you have requested"
		print "\n"
		# reads file
		book = Spreadsheet.open fileName
		sheet1 = book.worksheet sheetNum

		print '\n'
		newBook = Spreadsheet::Workbook.new
		newSheet = newBook.create_worksheet

		newSheet.name = "output"

		testvar = Array.new

		print "~>looking for unnecessary information"
		print "\n"
		sheet1.row(0).each_with_index{|val, index| 

			unless val != "OppCommunity" and val !="FirstName" and val != "LastName" and val != "Address" and val!="City" and val != "State" and val != "Zipcode" and val != "Phone" and val != "Email" and val != "ILorCC"
				testvar.push(index)
			end
		}

		unless sheet1.row(0).include? "OppCommunity"	
			print "you are missing the OppCommunity field\n".upcase
		else

		end
		unless sheet1.row(0).include? "FirstName"	
			print "you are missing the FirstName field\n".upcase
		end
		unless sheet1.row(0).include? "LastName"	
			print "you are missing the LastName field\n".upcase
		end
		unless sheet1.row(0).include? "Phone"	
			print "you are missing the phone field\n".upcase
		end
		unless sheet1.row(0).include? "Address"	
			print "you are missing the Address field\n".upcase
		end
		unless sheet1.row(0).include? "City"	
			print "you are missing the City field\n".upcase
		end
		unless sheet1.row(0).include? "State"	
			print "you are missing the State field\n".upcase
		end
		unless sheet1.row(0).include? "Zipcode"	
			print "you are missing the Zipcode field\n".upcase
		end
		unless sheet1.row(0).include? "Address"	
			print "you are missing the Address field\n".upcase
		end
		unless sheet1.row(0).include? "Address"	
			print "you are missing the Address field\n".upcase
		end
		unless sheet1.row(0).include? "Email"	
			print "you are missing the Email field\n".upcase
		end
		# puts testvar
		newvar = Array.new
		newvar.push "zero"
		sheet1.each do |row|
			row.each_with_index {|val, index| 
				if testvar.include? index
					
					newvar.push val
				end
			}
		end
		print "~~>writing new file \n\n"
		# print newvar
		newvar.map! { |x| 

			if x == "OppCommunity"
				print "renaming field #{x} \n"
				x = "Community"
			elsif x == "FirstName"
				print "renaming field #{x} \n"
				x = "First Name"	
			elsif x == "LastName"
				print "renaming field #{x} \n"
				x = "Last Name"
			elsif x == "Address"
				print "renaming field #{x} \n"
				x = "Address"
			elsif x == "city"
				print "renaming field #{x} \n"
				x = "city"
			elsif x == "State"
				print "renaming field #{x} \n"
				x = "State"
			elsif x == "Zipcode"
				print "renaming field #{x} \n"
				x = "Zipcode"
			elsif x == "Email"
				print "renaming field #{x} \n"
				x = "Email"
			else
				x=x
			end
		}
		# print newvar
		for i in 0..newvar.length
			for j in 1..testvar.length
					
			
					newSheet.row(i).push newvar[(i*testvar.length)+j]
			end
		end
		print "output file is in same file under out.xls to open, use \n \n open #{output}"
		print "\n"
		newBook.write output
		end
	end
end