# D:\sj_ruby_domacizadatak\sj_ruby_domacizadatak/xlsx_parser.rb

require 'roo'

class XlsxFile
    attr_accessor :path, :file, :table, :t, :row

    def initialize(path)
        @path = path
        @file = Roo::Spreadsheet.open("#@path")
        @file = Roo::Excelx.new("#@path", {:expand_merged_ranges => false})
        @table =  nil
        @row = nil
        self.initialize_table
        self.initialize_second_table
    end

    def initialize_table
        @file.each_with_pagename do |name, sh|
            if sh.first_column != nil then
                @table = Hash[]
                @row = Array.new(sh.last_row - sh.first_row + 1)

                rowCnt = 0
                flag = 0
                col_name = ""

                sh.first_column.upto(sh.last_column) do |column|
                    col_to_add = Column.new

                    sh.first_row.upto(sh.last_row) do |row|

                        if rowCnt == 0 then
                            col_name = sh.cell(row, column)
                            table[col_name] = nil
                        else
                            col_to_add << sh.cell(row, column)
                        end

                        if (sh.formula(row, column).to_s.include? "SUBTOTAL") || (sh.formula(row, column).to_s.include? "TOTAL")
                            flag = 1
                        end

                        rowCnt += 1
                    end

                    table[col_name] = col_to_add
                    rowCnt = 0
                end

                if flag == 1 then
                    table.each_value do |array|
                        array.pop
                    end
                end
            end
        end
    end
    
    def initialize_second_table
        @file.each_with_pagename do |name, sh|
            if sh.first_row != nil then
                @t =  Array.new(sh.last_row - sh.first_row + 1)
                @row = Array.new(sh.last_row - sh.first_row + 1)

                rowCnt = 0
                flag = 0
                row_to_remove = -1

                sh.first_row.upto(sh.last_row) do |row|
                    arr = []

                    sh.first_column.upto(sh.last_column) do |column|
                        arr << sh.cell(row, column)

                        if (sh.formula(row, column).to_s.include? "SUBTOTAL") || (sh.formula(row, column).to_s.include? "TOTAL")
                            flag = 1
                        end
                    end

                    if flag == 1 then
                        row_to_remove = rowCnt
                    end

                    t[rowCnt] = *arr
                    arr.clear

                    rowCnt += 1
                end

                @t.delete_at(row_to_remove)
            end
        end
    end

    def row(nr)
        @row = t[nr]
    end

    def each(&block)
        @t.each(&block)
    end

end

class Column < Array

    def sum
        sum = 0

        self.each do |el|
            if el != nil then
                sum += el.to_i
            end
        end

        sum
    end

end

def add_method(c, m, &b)
    c.class_eval {
      define_method(m, &b)
    }
end

x = XlsxFile.new('./sample.xlsx')

x.table.each do |key, value|
    add_method(XlsxFile, key) do
        value
    end
end

# p x.t

# p x.t[0][1]

# p x.row(0)[0]

# x.each do |cell|
#     p cell
# end

# p x.table["header1"]
# p x.table["header1"][0]

# p x.header1
# p x.header1[0]
# p x.header1.sum






