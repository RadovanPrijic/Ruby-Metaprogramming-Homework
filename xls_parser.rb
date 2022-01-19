require 'roo-xls'

class XlsFile
    attr_accessor :path, :file, :table, :t, :row # 'table' je hash sa kolonama tabele, a 't' predstavlja niz redova tabele.

    def initialize(path)
        @path = path
        @file = Roo::Spreadsheet.open("#@path")
        @table =  nil
        @row = nil
        self.initialize_table # Inicijalizacija tabele (hash-a) sa kolonama.
        self.initialize_second_table # Inicijalizacija tabele (niza) sa redovima.
    end

    def initialize_table # Funkcija za inicijalizaciju tabele sa kolonama.
        @file.each_with_pagename do |name, sh|
            if sh.first_column != nil then
                @table = Hash[] # Pravim hash gde ce kljucevi biti imena header-a, a vrednosti odgovarajuce kolone.
                @row = Array.new(sh.last_row - sh.first_row + 1) # Inicijalizujem niz za redove (sluzi za 3. zadatak).

                # Pomocne varijable za pravljenje hash-a.
                row_cnt = 0
                col_name = ""
                rows_to_remove = []

                sh.first_column.upto(sh.last_column) do |column| # Iteriram kroz kolone.
                    col_to_add = Column.new # Za svaku kolonu pravim objekat klase Column.

                    sh.first_row.upto(sh.last_row) do |row| # Iteriram kroz redove.

                        if row_cnt == 0 then
                            col_name = sh.cell(row, column) # U prvom redu uzimam ime header-a i smestam ga u pomocnu varijablu.
                            table[col_name] = nil # Stavljam ime header-a kao kljuc u mom hash-u.
                        else
                            col_to_add << sh.cell(row, column) # Smestam svaku vrednost date kolone u objekat Column.
                        end

                        # Ako dati red sadrzi kljucnu rec SUBTOTAL ili kljucnu rec TOTAL
                        if (sh.cell(row, column).to_s.include? "SUBTOTAL") || (sh.cell(row, column).to_s.include? "TOTAL")
                            if !(rows_to_remove.include? row_cnt) then # i nije vec "oznacen" za brisanje,
                                rows_to_remove << row_cnt # onda ga dodajem u niz redova koje treba ukloniti.
                            end
                        end

                        row_cnt += 1
                    end

                    table[col_name] = col_to_add # Povezem napravljenu kolonu sa odgovarajucim kljucem (header-om).
                    row_cnt = 0
                end

                rows_already_deleted = 1 # Varijabla koja mi sluzi za pracenje vec izbrisanog broja redova.

                rows_to_remove.each do |row_nr| # Svaki red koji sam "oznacio" za brisanje uklanjam iz tabele tabele kolona.

                    table.each_value do |col|
                        col.delete_at(row_nr - rows_already_deleted)
                    end

                    rows_already_deleted += 1
                end

            end
        end
    end
    
    def initialize_second_table # Funkcija za inicijalizaciju tabele sa redovima.
        @file.each_with_pagename do |name, sh|
            if sh.first_row != nil then
                @t =  Array.new(sh.last_row - sh.first_row + 1) # Inicijalizujem niz u kog cu smestiti redove tabele.

                # Pomocne varijable za pravljenje niza.
                row_cnt = 0
                rows_to_remove = []

                sh.first_row.upto(sh.last_row) do |row| # Iteriram kroz redove.
                    arr = [] # Inicijalizujem pomocni niz u kog ce stavljati vrednosti reda.

                    sh.first_column.upto(sh.last_column) do |column| # Iteriram kroz kolone.
                        arr << sh.cell(row, column) # Smestam svaku vrednost datog reda u pomocni niz.

                        # Ako dati red sadrzi kljucnu rec SUBTOTAL ili kljucnu rec TOTAL
                        if (sh.cell(row, column).to_s.include? "SUBTOTAL") || (sh.cell(row, column).to_s.include? "TOTAL")
                            if !(rows_to_remove.include? row_cnt) then # i nije vec "oznacen" za brisanje,
                                rows_to_remove << row_cnt # onda ga dodajem u niz redova koje treba ukloniti.
                            end
                        end
                    end

                    t[row_cnt] = *arr # U pocetni niz dodajem pomocni niz sa svim vrednostima reda.
                    row_cnt += 1
                end

                rows_already_deleted = 0 # Varijabla koja mi sluzi za pracenje vec izbrisanog broja redova.

                rows_to_remove.each do |row_nr| # Svaki red koji sam "oznacio" za brisanje uklanjam iz tabele redova.
                    t.delete_at(row_nr - rows_already_deleted)
                    rows_already_deleted += 1
                end

            end
        end
    end

    def row(nr) # Funkcija s kojom vracam odredjeni red.
        @row = t[nr]
    end

    def each(&block) # Modifikovani each.
        @t.each(&block)
    end

    def remove_empty_rows # Funkcija za uklanjanje praznih redova.
        # Pomocne varijable.
        empty_cells = 0
        row_cnt = 0
        rows_to_remove = []

        t.each do |row| # Iteriram kroz redove.
            row.each do |cell| # Iteriram kroz celije u redu.
                if cell == nil then
                    empty_cells += 1 # Za svaku vrednost koja je jednaka nil povecavam broj praznih celija.
                end
            end

            if row.length == empty_cells then
                rows_to_remove << row_cnt # Ako mi je duzina reda jednaka broju praznih celija, onda moram taj red obrisati.
            end

            empty_cells = 0
            row_cnt += 1
        end

        rows_already_deleted = 0 # Varijabla koja mi sluzi za pracenje vec izbrisanog broja redova.
        rows_already_deleted_cols = 1 # Varijabla koja mi sluzi za pracenje vec izbrisanog broja redova, ali za kolone mora biti za 1 veca (zbog toga sto se header stavlja kao kljuc).

        rows_to_remove.each do |row_nr| # Svaki red koji sam "oznacio" za brisanje uklanjam iz tabele redova i tabele kolona.

            t.delete_at(row_nr - rows_already_deleted)

            table.each_value do |col|
                col.delete_at(row_nr - rows_already_deleted_cols)
            end

            rows_already_deleted += 1
            rows_already_deleted_cols += 1
        end
    end

    def add_method(c, m, &b) # Pomocna funkcija koja sluzi za dinamicko dodavanje metoda u klasu.
        c.class_eval {
          define_method(m, &b)
        }
    end
    
    def add_column_methods # Funkcija za dinamicko dodavanje metoda za kolone (omogucavaju direktan pristup kolonama).
        self.table.each do |key, value|
            add_method(XlsFile, key) do
                value
            end
        end    
    end

    def +(second_file) # Pregazena genericka funkcija za sabiranje.
        if self.t[0].eql?(second_file.t[0]) then
            return self.t + second_file.t[1..] # Ako tabele imaju identicne header-e, sabiram (spajam) njihove redove.
        end
    end

    def -(second_file) # Pregazena genericka funkcija za oduzimanje.
        if self.t[0].eql?(second_file.t[0]) then
            return self.t - second_file.t[1..] # Ako tabele imaju identicne header-e, oduzimam iz prve sve redove koji su i u drugoj.
        end
    end

end

class Column < Array # Pomocna klasa koju koristim za reprezentaciju kolona.

    def sum # Funkcija za sabiranje vrednosti kolone.
        sum = 0

        self.each do |el| # Iteriram kroz vrednosti kolona i, ako nisu nil, sabiram ih.
            if el != nil then
                sum += el.to_i
            end
        end

        sum
    end

end