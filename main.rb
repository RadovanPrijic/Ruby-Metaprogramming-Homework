# D:\sj_ruby_domacizadatak\sj_ruby_domacizadatak/main.rb

require_relative 'xlsx_parser'
require_relative 'xls_parser'

# x = XlsxFile.new('./sample.xlsx')
# x = XlsFile.new('./sample2.xls')

puts '-------------------------------'
puts 'ISPIS TABELE'

p x.t

puts '-------------------------------'
puts 'SINTAKSA NIZA TEST'

p x.t[0][0]

puts '-------------------------------'
puts 'ROW TEST'

p x.row(0)
p x.row(0)[0]

puts '-------------------------------'
puts 'EACH TEST'

x.each do |cell|
    p cell
end

puts '-------------------------------'
puts 'CUSTOM NIZ SINTAKSA TEST'

p x.table["header1"]
p x.table["header1"][0]

puts '-------------------------------'
puts 'DINAMICKA METODA TEST'

p x.header1
p x.header1[0]

puts '-------------------------------'
puts 'SUM TEST'

p x.header1.sum