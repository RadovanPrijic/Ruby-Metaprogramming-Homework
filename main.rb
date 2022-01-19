# ruby D:\sj_ruby_domacizadatak\sj_ruby_domacizadatak/main.rb

require_relative 'xlsx_parser'
require_relative 'xls_parser'

# x = XlsxFile.new('./sample_1.xlsx')
# y = XlsxFile.new('./sample_2.xlsx')
x = XlsFile.new('./sample_3.xls')
y = XlsFile.new('./sample_4.xls')

puts '-------------------------------'
puts 'ISPIS TABELE REDOVA'

p x.t

puts '-------------------------------'
puts 'ISPIS TABELE KOLONA'

p x.table

puts '-------------------------------'
puts 'BRISANJE PRAZNIH REDOVA'

x.remove_empty_rows
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

x.add_column_methods
p x.header1
p x.header1[0]

puts '-------------------------------'
puts 'SUM TEST'

p x.header1.sum

puts '-------------------------------'
puts 'SABIRANJE TABELA TEST'

p x + y

puts '-------------------------------'
puts 'ODUZIMANJE TABELA TEST'

p x - y