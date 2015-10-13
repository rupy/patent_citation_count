require 'spreadsheet'
require 'csv'

CITE_ROW_NUM = 9

class NcRc
	attr_accessor :nc, :rc, :name
	def initialize()
		@nc = @rc = 0
	end
end

def ncrc_count(patent_num_counter, ncrc, w)
	patent_num_counter[w] += 1
	if patent_num_counter[w] == 1
		ncrc.nc += 1
	else
		ncrc.rc += 1
	end
end

Spreadsheet.client_encoding = 'UTF-8'

book = Spreadsheet.open './masaki_register.xls'
sheet1 = book.worksheet 0

patent_num_counter = Hash.new(0)

ncrc_arr = []

sheet1.reverse_each.with_index do |row, i|

	break if i == sheet1.last_row_index

	# puts "<#{row[0]}>"
	nc = 0
	rc = 0
	ncrc = NcRc.new
	ncrc.name = row[0]

	str = row[CITE_ROW_NUM]
	# puts "> #{str}"
	word_store = []
	while idx = str.index(/([[:space:]]+)/)
		matched_space = $1
		target_word = str[0..idx-1]
		if target_word =~ /^(特|GB|CN|FR|DE|USP|実|GP|EPA|USA|WO)/
			if word_store.size != 0
				w = word_store.join(" ")
				# puts w
				ncrc_count(patent_num_counter, ncrc, w)
	  			word_store = []
			end
			w = target_word
			# puts w
			ncrc_count(patent_num_counter, ncrc, w)
		else
			word_store.push target_word
		end
		str = str[idx+matched_space.size..-1] # 分割
	end

	if word_store.size != 0
		if str =~ /^(特|GB|CN|FR|DE|USP|実|GP|EPA|USA|WO)/
			w = word_store.join(" ")
			# puts w
			ncrc_count(patent_num_counter, ncrc, w)
			# puts str
			ncrc_count(patent_num_counter, ncrc, str)
		else
			word_store.push str
			w = word_store.join(" ")
			# puts w
			ncrc_count(patent_num_counter, ncrc, w)
		end
	else
		# puts str
		ncrc_count(patent_num_counter, ncrc, str)
	end
	
	ncrc_arr.push ncrc
end

CSV.open("result.csv", "wb", encoding: 'Windows-31J') do |csv|
	ncrc_arr.reverse_each do |ncrc|
		puts "#{ncrc.name}, #{ncrc.nc}, #{ncrc.rc}"
		csv << [ncrc.name, ncrc.nc, ncrc.rc]
	end
end
