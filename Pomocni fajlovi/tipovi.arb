# Ruby ima veliki broj tipova podataka koje je moguće koristiti

#Brojevi
#Ruby podržava, Integer, Float, Complex, BigDecimal, Rational
#Sve ove klase nasleđuju Numeric klasu

#Integer je specifičan po tome što ima dve "interpretacije"
#Ukoliko broj koji sadrži staje u 31bit, onda je Fixnum klasa, ukoliko ne onda je Bignum klasa 
#Sve metode unutar Numeric objekata su immutable, dakle ne menjaju samu vrednost objekta

#Ukoliko integer pocinje sa 0, a ima više od jedne cifre, interpretira se da nije u bazi 10
# 0xFF su heksadecimalni, 0b11111 su binarni, 0377 je oktalni
# Ukoliko nam treba negativna vvrednost samo stavimo - ispred

a= 0xFF #postavice vrednost na 255

#Da bi nesto proglasili floatom koristimo .
# 0.0, -3.14, 6.02e23

#Idalje treba imati u vidu da postoji standardna float greska, gde nije moguće predstaviti brojeve kao 1/10
#naprimer ovo bi bilo false 0.4-0.3 == 0.1
#Da bi ovo prevazišli koristimo BigDecimal 

#Postoje standardni operatori +,-,*,/
#Automatski se vrse konverzije fixnum <-> Bigint pri operacijama

#deljenje zavisi od delioca
x = 5/2     # result is 2
y = 5.0/2   # result is 2.5
z = 5/2.0   # result is 2.5

#Deljenjem sa nulom, dobijamo ZeroDivisionError
#Ukoliko delimo float sa nulom, dobijamo infinity
0.0/0.0 #Vraća NaN

#Ima standardni modulo operator

x = 5%2 # Vraća 1

#Treba napomenuti da Ruby drugačije zaokruživanje pri deljenju i modulo
#Ova razlika se ogleda samo u negativnim brojevima, gde zaokružuje ka minus beskonacno
#Dakle -a/b je isto što i a/-b
#Ali ne mora biti jednako sa -(a/b)

#Postoji i ** operator
x**4          # This is the same thing as x*x*x*x
x**-1         # The same thing as 1/x
x**(1/3.0)    # The cube root of x
x**(1/4)      # Oops! Integer division means this is x**0, which is always 1
x**(1.0/4.0)  # This is the fourth-root of x

#Takodje su implementirani standardne bitovne operacije << , >>, &, |
#Možemo takođe pristupati bitovima nekog broja, naprimer

5[0] # Vratiće nam 1, jer je poslednji biz zapravo 1, 101

#Neke osnovne funkcije koje nude brojevi

0.zero?        # => true (is this number zero?)
1.0.zero?      # => false
0.0.nonzero?   # => nil (works like false)
1.nonzero?     # => 1 (works like true)
1.integer?     # => true
1.0.integer?   # => false
1.scalar?      # => false: not a complex number. 
0.even?        # => true (Ruby 1.9)
0.odd?         # => false#

ZERO, INF, NAN = 0.0, 1.0/0.0, 0.0/0.0 
ZERO.finite?   # => true: is this number finite?
INF.finite?    # => false
NAN.finite?    # => false
ZERO.infinite? # => nil: is this number infinite? 
INF.infinite?  # => 1
-INF.infinite? # => -1
NAN.infinite?  # => nil
ZERO.nan?      # => false: is this number not-a-number?
INF.nan?       # => false
NAN.nan?       # => true


#Roundovanje
1.1.ceil     # =>  2: ceiling: smallest integer >= its argument
-1.1.ceil    # => -1: ceiling: smallest integer >= its argument
1.9.floor    # =>  1: floor: largest integer <= its argument
-1.9.floor   # => -2: floor: largest integer <= its argument
1.1.round    # =>  1: round to nearest integer
0.5.round    # =>  1: round toward infinity when halfway between integers 
-0.5.round   # => -1: or round toward negative infinity
1.1.truncate # =>  1: chop off fractional part: round toward zero
-1.1.to_i    # => -1: synonym for truncate

# Absolute value and sign
-2.0.abs     # => 2.0: absolute value
-2.0<=>0.0   # => -1: use <=> operator to compute sign of a number


#Random brojevi
rand       # => 0.964395196505186
rand       # => 0.390523655919935
rand(100)  # => 81
rand(100)  # => 32

#Stringovi

#Stringovi su u Rubiju, mutable, dakle mogu se menjati, i biti menjani u metodama
#String generalno ima veoma mocnu biblioteku sam po sebi
#Medjutim u vecini slucajeva se koriste literali

'This is a simple Ruby string literal'
'Won\'t you eat that pie?' #eskejpujemo sa \
#Ukoliko želimo backslash možemo samo napisati \\
#Generalno  '' stringovi su slabiji u smislu da ne podržavaju previše \ funkcionalnosti 
#naprimer "" podržavaju \n, \t, \" 
"\t\"This quote begins with a tab and ends with a newline\"\n"
"\\"  # A single backslash


#Možemo i emberdovati unicode karaktere preko \u
"\u20ac" # => "€"

# "" takođe podržava i #{} notaciju koja nam dozvoljava da embedujemo expression u string
"360 degrees=#{2*Math::PI} radians" # "360 degrees=6.28318530717959 radians"

#Ukoliko želimo da embedujemo samo promenjivu možemo koristiti drugu notaciju
#Ovaj način možemo koristit samo ukoliko promenjive počinju sa @,@@,$
$salutation = 'hello'     # pravimo globalnu varijablu
"#$salutation world"      # korisimo je u stringu


#String sa "" možemo pisati u više linija, ukoliko ne želimo da nam new line uđe u string, koristimo /
"This string literal
has two lines \
but is written on three"

#U slucajevima kada imamo neki SQL ili neke veće stringove koji moraju imati '' ili "" koristimo %q i %Q
%q(Don't worry about escaping ' characters!) #Pravi '' string
%Q|"How are you?", he said| #Pravi "" string

#Stringove možemo koristiti i kao poziv nekih funkcija same konzole

%x[dir] #će pozvati dir komandu, i vratiti njen output
s
if windows
  listcmd = 'dir'
else  listcmd = 'ls'
end
listing = `#{listcmd}`

#Svaki put kada napravimo literal, ruby će napraviti novu instancu ovog stringa
#generalno iz ovog razloga ih treba izbegavati u loopovima
10.times { puts "test".object_id }

#String operatori

# + operator vrši konkatenaciju
planet = "Earth"
"Hello" + " " + planet    # Produces "Hello Earth"


#Treba držati u vidu da + ne konvertuje ono što je sa desne strane u string, kao u Javi
"Hello planet #" + planet_number.to_s  # to_s converts to a string

#Za ovako nešto koristimo string interpolaciju, koja automatski konveruje u string
"Hello planet ##{planet_number}"

#Operator << dodaje desni argument na levi argument
#Ovaj operator će promeniti levi objekat
greeting = "Hello"
greeting << " " << "World"
puts greeting   # Outputs "Hello World

#Možemo množiti da ponovimo string
ellipsis = '.'*3    # Evaluates to '...'

#String podržava standardne == i != komparacije
#Dva stringa su jednaka sam i samo ako im je dužina ista i svi karakteri su isti
# <, <=, >, >= porede karakter kodove da utvrde šta je veće
#Komparacija je case sensitive (U ASCI, počinje se sa velikim slovima)


#Stringovima se može pristupati i preko [] notacije.
#Ova notacija se obično koristi za ekstrakciju karaktera/podstringova kao i ya promenu samog stringa

s = 'hello';   
s[0]           # 'h': the first character of the string, as a string
s[s.length-1]  # 'o': the last character 'o'
s[-1]          # 'o': another way of accessing the last character
s[-2]          # 'l': the second-to-last character
s[-s.length]   # 'h': another way of accessing the first character
s[s.length]    # nil: there is no character at that index

#Promenu individualnih karaktera možemo uraditi preko ove notacije
s = "hello"      # Begin with a greeting
s[-1] = ""       # Delete the last character; s is now "hell"
s[-1] = "p!"     # Change new last character and add one; s is now "help!"

#Primenom ove notacije možemo izvući i substring iz nekog stringa
#Koristimo [X,Y] notaciju, to jest odakle, dokle
 = "hello"
 s[0,2]          # "he"
 s[-1,1]         # "o": returns a string
 s[0,0]          # "": a zero-length substring is always empty
 s[0,10]         # "hello": returns all the characters that are available
 s[s.length,1]   # "": there is an empty string immediately beyond the end
 s[s.length+1,1] # nil: it is an error to read past that
 s[0,-1]         # nil: negative lengths don't make any sense

s = "hello"
s[0,1] = "H"              # Replace first letter with a capital letter
s[s.length,0] = " world"  # Append by assigning beyond the end of the string
s[5,0] = ","              # Insert a comma, without deleting anything
s[5,6] = ""               # Delete with no insertion; 
s == "Hellod"

#Kao parametar unutar [] notacije možemo koristiti i range
s = "hello"
s[2..3]           # "ll": characters 2 and 3
s[-3..-1]         # "llo": negative indexes work, too
s[0..0]           # "h": this Range includes one character index
s[0...0]          # "": this Range is empty
s[2..1]           # "": this Range is also empty
s[7..10]          # nil: this Range is outside the string bounds
s[-2..-1] = "p!"     # Replacement: s becomes "help!"
s[0...0] = "Please " # Insertion: s becomes "Please help!"
s[6..10] = ""        # Deletion: s becomes "Please!

#String implementira iterator klasu, koja prolazi kroz svaki char
# each_byte, each_char, each_line

s = "¥1000"
s.each_char {|x| print "#{x} " }         # Prints "¥ 1 0 0 0". 
0.upto(s.size-1) {|i| print "#{s[i]} "}  # Nije efikasno kada se koristi UTF8


#Neke od korisnih funkcija kod Ruby stringova
s = "hello"
s.insert(5, " there") # Same as s[5] = " there". Alters s. Returns new s.
s.slice(0,5)          # Same as s[0,5]. Returns a substring.
s.slice!(5,6)         # Deletion. Same as s[5,6]="". Returns deleted substring.
s.eql?("hello world") # True. Same as ==.
s.empty?         # => false
s.index('l')         # => 2: index of first l in string
s.index(?l)          # => 2: works with character codes as well
s.index(/l+/)        # => 2: works with regular expressions, too
s.rindex('l')        # => 3: index of rightmost l in string

s.start_with? "hell" # => true.  Note singular "start" not "starts"
s.end_with? "bells"  # => false
s.include?("ll")     # => true: "hello" includes "ll"
s.include?(?H)       # => false: "hello" does not include character H
s.match(/[aeiou]/) {|m| m.to_s} # => "e": return first vowel

"this is it".split     # => ["this", "is", "it"]: split on spaces by default

#Search and replace za string
s.sub("l", "L")            # => "heLlo": Just replace first occurrence
s.gsub("l", "L")           # => "heLLo": Replace all occurrences
s.gsub!("l", "L")           # => "heLLo": Replace all occurrences, destruktivan


s.upcase      # => "WORLD"
s.upcase!     # => "WORLD"; alter s in place
s.downcase    # => "world"
s.capitalize  # => "World": first letter upper, rest lower
s.capitalize! # => "World": alter s in place
s.swapcase    # => "wORLD": alter case of each letter


s = "hello\r\n"      # A string with a line terminator
s.chomp!             # => "hello": remove one line terminator from end
s.chomp              # => "hello": no line terminator so no change
s.chomp!             # => nil: return of nil indicates no change made
s.chomp("o")         # => "hell": remove "o" from end

s = "\t hello \n"   # Whitespace at beginning and end
s.strip             # => "hello"




#Nizovi
#Skup vrednosti organizovan, po poziciji, tako da se svakoj moze pristupiti preko svog indeksa
#Nizovi pocinju od 0
# Size/length metode vraćaju dužinu niza
#Poslednji element je size-1
#Pristup indeksu koji ne postoji ce vratiti nil, i neće vratiti exception
#Ruby nizovi nisu tipizirani, dakle mogu sadržati bilo koje elemente, bilo kog tipa, naprimer String i int se mogu čuvati unutar niza u isto vreme
#Mutable su, dakle mogu biti menjani tokom poziva metoda
#Dinamicke su velicine, kako dodajemo elemente, niz se sam proširuje

[1, 2, 3]         # Niz koji sadrži tri Fixnum
[-10...0, 0..10,] # Niz koji sadrži dva range-a
[[1,2],[3,4],[5]] # Ugnježdeni niz
[x+y, x-y, x*y]   # Elementi niza mogu biti i expressioni

#generalno često se koristi notacija inicijalizacije niza na neke stringove
words = %w[this is a test]  # Same as: ['this', 'is', 'a', 'test']
white = %W(\s \t \r \n)     # Same as: ["\s", "\t", "\r", "\n"]

#Postoji i konstruktor, kao i za većinu klasa
#Proseledjivanjem parametara kosntruktoru možemo da napravimo specifične nizove
#TAkodje, možemo i jednostavno instancirati novi niz sa [] 
empty = Array.new       # []: returns a new empty array
nils = Array.new(3)     # [nil, nil, nil]: new array with 3 nil elements
count = Array.new(3) {|i| i+1}  # [1,2,3]: 3 elements computed from index

#Da bi pristupili nekom elementu koristimo [] notaciju
a = [0, 1, 4, 9, 16]   # Array holds the squares of the indexes
a[0]        # First element is 0
a[-1]       # Last element is 16
a[-2]       # Second to last element is 9
a[a.size-1] # Another way to query the last element
a[-a.size]  # Another way to query the first element
a[8]        # Querying beyond the end returns nil
a[-8]       # Querying before the start returns nil, too

#Preko []= notacije možemo i postavljati element
[0] = "zero"   # a is ["zero", 1, 4, 9, 16]
a[-1] = 1..16   # a is ["zero", 1, 4, 9, 1..16]
a[8] = 64       # a is ["zero", 1, 4, 9, 1..16, nil, nil, nil, 64]
a[-9] = 81      # Error: can't assign before the start of an array


#kao i stringovi možemo izvlačiti podskupove, to jest, neki deo niza
a = ('a'..'e').to_a   # Range converted to ['a', 'b', 'c', 'd', 'e']
a[0,0]                # []: this subarray has zero elements
a[1,1]                # ['b']: a one-element array
a[-2,2]               # ['d','e']: the last two elements of the array
a[0..2]               # ['a', 'b', 'c']: the first three elements
a[-2..-1]             # ['d','e']: the last two elements of the array
a[0...-1]             # ['a', 'b', 'c', 'd']: all but the last element

#Ovu notaciju možemo koristiti i za promenu niza
a[0,2] = ['A', 'B']      # a becomes ['A', 'B', 'c', 'd', 'e']
a[2...5]=['C', 'D', 'E'] # a becomes ['A', 'B', 'C', 'D', 'E']
a[0,0] = [1,2,3]         # Insert elements at the beginning of a
a[0..2] = []             # Delete those elements
a[-1,1] = ['Z']          # Replace last element with another
a[-1,1] = 'Z'            # For single elements, the array is optional

#Nizove možemo i sabirati, ovo konkatenira nizove

a = [1, 2, 3] + [4, 5]    # [1, 2, 3, 4, 5]
a = a + [[6, 7, 8]]       # [1, 2, 3, 4, 5, [6, 7, 8]]
a = a + 9                 # Error: righthand side must be an array

#možemo oduzimati, gde uklanjamo elemente koji se nalaze na desnoj strani operatora
['a', 'b', 'c', 'd', 'a'] - ['b', 'c', 'd']    # ['a', 'a']

# << operator je generalno najkorisniji jer na omogućava da sa lakoćom dodajemo stvari na kraj niza
a = []        # Start with an empty array
a << 1        # a is [1]
a << 2 << 3   # a is [1, 2, 3]
a << [4,5,6]  # a is [1, 2, 3, [4, 5, 6]]


#Niz podržava & i | operatore koji služe za uniju i presek
# | spaja dva niza i uklanja sve duplikate
# & vraća samo elemente koji se nalaze u oba niza
a = [1, 1, 2, 2, 3, 3, 4]
b = [5, 5, 4, 4, 3, 3, 2]
a | b    # [1, 2, 3, 4, 5]: duplicates are removed
b | a    # [5, 4, 3, 2, 1]: elements are the same, but order is different
a & b    # [2, 3, 4]
b & a    # [4, 3, 2]

#Neke od korisnih funkcija u nizovima
a = [1,2,3,4,5,6]
a.delete_at(4)     # => 5: a is now [1,2,3,4,6]
a.delete_at(-1)    # => 6: a is now [1,2,3,4]
a.delete_at(4)     # => nil: a is unchanged

#možemo brisati i po vrednosti
a.delete(4)        # => 4: a is [1,2,3]
a[1] = 1           # a is now [1,1,3]
a.delete(1)        # => 1: a is now [3]: both 1s removed

#možemo brisati ukoliko se ispuni uslov
a = [1,2,3]
a.delete_if {|x| x.odd?} # Remove odd values: a is now [2]

#Možemo brisati i delove niza preko slice
a = [1,2,3,4,5,6,7,8]
a.slice!(0)        # => 1: remove element 0: a is [2,3,4,5,6,7,8]
a.slice!(-1,1)     # => [8]: remove subarray at end: a is [2,3,4,5,6,7]
a.slice!(2..3)     # => [4,5]: works with ranges: a is [2,3,6,7]
a.slice!(4,2)      # => []: empty array just past end: a unchanged
a.slice!(5,2)      # => nil: a now holds [2,3,6,7,nil]!

#Imamo više načina da iteriramo kroz njih
a = ['a','b','c']
a.each {| elt| print elt }         # The basic each iterator prints "abc"
a.reverse_each {|e| print e}       # Array-specific: prints "cba" 
a.each_index {|i| print i}         # Array-specific: prints "012"
a.each_with_index{|e,i| print e,i} # Enumerable: prints "a0b1c2"
a.map {|x| x.upcase}               # Enumerable: returns ['A','B','C'] #Menjamo svako x na tu vrednost
a.map! {|x| x.upcase}              # Array-specific: alters a in place
a.collect! {|x| x.downcase!}       # collect! is synonym for map!ž


#Pretrage
a = %w[h e l l o]
a.include?('e')                    # => true
a.index('l')                       # => 2: index of first match
a.rindex('l')                      # => 3: search backwards
a.index {|c| c =~ /[aeiou]/}       # => 1: index of first vowel. Ruby 1.9.

#Sortiranja
a.sort     # => %w[e h l l o]: copy a and sort the copy
a.sort!    # Sort in place: a now holds ['e','h','l','l','o']
a = [1,2,3,4,5]               # A new array to sort into evens and odds
a.sort! {|a,b| a%2 <=> b%2}   # Compare elements modulo 2

a = [1,2,3]     # Start ordered
puts a.shuffle  # Shuffle randomly. E.g.: [3,1,2]. Also shuffle!

#Možemo array da koristimo kao stack preko push i pop
a = []a.push(1)     # => [1]: a is now [1]
a.push(2,3)   # => [1,2,3]: a is now [1,2,3]
a.pop         # => 3: a is now [1,2]
a.pop         # => 2: a is now [1]
a.pop         # => 1: a is now []
a.pop         # => nil: a is still []

#Shift skida prvi element sa niza
a = []a.push(1)     # => [1]: a is [1]
a.push(2)     # => [1,2]: a is [1,2]
a.shift       # => 1: a is [2]
a.push(3)     # => [2,3]: a is [2,3]
a.shift       # => 2: a is [3]
a.shift       # => 3: a is []
a.shift       # => nil: a is []

#neke druge stvari
[0,1]*3                   # => [0,1,0,1,0,1]: * operator repeats
[1, [2, [3]]].flatten     # => [1,2,3]: recursively flatten; also flatten!
[1, [2, [3]]].flatten(1)  # => [1,2,[3]]: specify # of levels; 
[1,2,3].reverse           # => [3,2,1]: also reverse!
a.transpose               # => [[1,2,3],[:a,:b,:c]]: swap rows/cols



#Hashevi
#Hash je struktura podataka koja čuva objekte koji su ključevi i vrednosti
#U drugim jezicima se zovu mape/asocijativni nizovi
#U pozacini se nalazi Hash table
#Svaki od ključeva mora imati hash metodu, koja vraća njegov hash za ključ
#Dakle ukoliko menjamo .equals? trebalo bi da promenimo i hash ukoliko želimo tu klasu da koristimo kao ključ
#mdjutim, generalno uvek se kao kljuc koriste simboli
#Ukoliko koristimo string, hash struktura ce da napravi kopiju našeg stringa, i nju ce koristiti, ovo se radi zato sto je string sam po sebi mutable


#Za dodavanje elemenata u hash koristimo []= notaciju
numbers = Hash.new     # Create a new, empty, hash object
numbers["one"] = 1     # Map the String "one" to the Fixnum 1
numbers["two"] = 2     # Note that we are using array notation here
numbers[:number] = 3


#mogli smo da koristimo i numbers = {}
#Literal hasha se pise po sledecoj notaciji
numbers = { "one" => 1, "two" => 2, "three" => 3 }
#Da bi cela stvar bila elegantnija možemo koristiti simbole
numbers = { :one => 1, :two => 2, :three => 3 }
#A da bi cela stvar bila jos elegantnija mozemo koristi drugu notaciju
numbers = { one: 1, two: 2, three: 3 }

h = { :one => 1, :two => 2 }
h[:one]       # => 1: find value associated with a key
h[:three]     # => nil: the key does not exist in the hash
h.assoc :one  # => [:one, 1]: find key/value pair. 
h.index 1     # => :one: search for key associated with a value

#Provera postojanja je generalno brza posto hash tabele su izrazito brza struktura za "pretraživanje"
h = { :a => 1, :b => 2 }
h.key?(:a)       # true: :a is a key in h
h.has_key?(:b)   # true: has_key? is a synonym for key?
h.include?(:c)   # false: include? is another synonym
h.member?(:d)    # false: member? is yet another synonym

#Kao alternativu za [] notaciju možemo koristit fetch, kako bi obradili slučaj da nešto ne postoji
h.fetch(:a)      # => 1: works like [] for existing keys
h.fetch(:c)      # Raises IndexError for nonexistent key
h.fetch(:c, 33)  # => 33: uses specified value if key is not found
h.fetch(:c) {|k| k.to_s } # => "c": calls block if key not found

#Ukoliko želimo sve key/value elemente koji ispunjavaju neki uslov koristimo select
h = { :a => 1, :b => 2, :c => 3 }
h.select {|k,v| v % 2 == 0 } # => {:b=>2} 

#Da bi spojili dva hasha korsitimo merge komandu
k = h.merge(j) # If h and j share keys, use values from j

#brisanje elemenata mora se vršiti preko funkcije, ne možemo jednostavno staviti null
h = {:a=>1, :b=>2}
h[:a] = nil      # h now holds {:a=> nil, :b=>2 }
h.include? :a    # => true
h.delete :b      # => 2: returns deleted value: h now holds {:a=>nil}
h.include? :b    # => false
h.delete :b      # => nil: key not found
{|k| raise IndexError, k.to_s } # Invoke block if key not foundh.delete(:b)

#Možemo i brisati više elemenata pod uslovom
h.delete_if {|k,v| k.to_s < 'b' }  # => {:b=>2, :c=>3 }

#Мožemo očistiti ceo hash 
h.clear

#Hash ima iterator
h.each do |key, value|                
  print "#{key}:#{value} "     # Prints "a:1 b:2 c:3" 
end


#Simboli

#Ruby interpreter čuva tabelu simbola, gde čuva imena svih klasa, metoda i varijabli kao simbole.
#Na ovaj način izbegava poredjenje stringova
#Metod je naprimer identifikovan po njegovoj poziciji u tabeli simbola
#Generalno simboli se mogu koristiti i za druge svrhe, obično kao neki ID za nape klase
#Dva stringa mogu imati isti sadržaj ali biti dva različita objekta, ovo nije nikada sličaj sa simbolima


:symbol                   # A Symbol literal
:"symbol"                 # The same literal
:'another long symbol'    # Quotes are useful for symbols with spaces
s = "string"
sym = :"#{s}"             # The Symbol :string

#Generalno se koriste kod refleksije 
o.respond_to? :each

#Ovaj kod će proveriti da li objekat sadrži metodu size, i pozvati size preko .send()
name = :size
if o.respond_to? name  
	o.send(name)
end

#Stringovi se lako mogu prebacivati u simbole
str = "string"     # Begin with a string
sym = str.to_sym   # Another way to do the same thing
str = sym.to_s     # Convert back to a string
str = sym.id2name  # Another way to do it


# Ruby Objekti
#Ruby je veoma objektno orijentisan jezik, sve je objekat, n postoje primitivni tipovi
#Svi obekti nasledjuju Objekat, i nasleđuju sve unutar te klase

#Kada radimo sa objektima u rubiju, radimo sa referencama nad objektima. (pointeri=
#Pri pozivu metora, uvek se prosledjuje referenca
#Dakle metode mogu menjati objekat koji je prosledjen

#Da bi napravili instancu neke klase, koristimmo metodu .new
#Ova metoda alocira memoriju za ovaj objekat, i poziva initialize metodu
myObject = MyClass.new

#Nema potrebe da dealociramo neki objekat, već postoji garbage kolektor, koji to automatski radi
#Objekat postaje kandidat za garbage collector kada u memoriji ne postoje više reference na njega
#Ovime se osigurava da su memory leakovi retkost

#U rubiju svaki objekat ima svoj object_id kome se pristupa preko istoimene metode
#Ovaj ID će biti jedinstven za ovaj objekat, dokle god ovaj objekat postoji
o.object_id

#Da bi videli koje je klase neki objekat možemo pozvati određene metode
o = "test"  # This is a value
o.class     # Returns an object representing the String class
o.class.superclass            # Object: superclass of String is Object
o.class.superclass.superclass # BasicObject

#Da bi proverili da li je nas objekat neka konkretna klasa koristimo sledeći kod
o.class == String       # true if is o a String
o.instance_of? String   # true if o is a String

#da bi proverili da li je nešto podklasa nečega koristimo is_a?
x = 1                    # This is the value we're working with
x.instance_of? Fixnum    # true: is an instance of Fixnum
x.instance_of? Numeric   # false: instance_of? doesn't check inheritance
x.is_a? Fixnum           # true: x is a Fixnum
x.is_a? Integer          # true: x is an Integer
x.is_a? Numeric          # true: x is a Numeric
x.is_a? Comparable       # true: works with mixin modules, too
x.is_a? Object           # true for any value of x

#Možemo takođe koristiti i === da postignemo isto, međutim nije toliko elegantno

#Generalno se klasa nekog objekta nikada ne menja. međutim samo ponašanje tog objekta se može menjati
#Neki objekat u realnom vremenu može dodati sebi neke metode, ili tako dalje
#Generalno, ne interesuje nas previše sama klasa, već metode na koje se sam objekat odaziva
#ovo  ponasanje testiramo preko respond_to?

o.respond_to? :"<<"  # true if o has an << operator

#equal? poredi dva objekta
#za bilo koja dva objekta, uvek se po defaultu vraća false
a = "Ruby"       # One reference to one String object
b = c = "Ruby"   # Two references to another String object
a.equal?(b)      # false: a and b are different objects
b.equal?(c)      # true: b and c refer to the same object

#Po konvenciji podklase nikada ne overrajduju equal? metod

#Drugi način za proveru je da jednsotavno proverimo njihov object_id
a.object_id == b.object_id   # Works like a.equal?(b)

#Operator == se korsiti za kontekstno poređenje, to jest, kako bi rekli zdravorazumsko
#Generalno logika je suprotna od Jave, == za kontekst, .equals za reference

a = "Ruby"    # One String object
b = "Ruby"    # A different String object with the same content
a.equal?(b)   # false: a and b do not refer to the same object
a == b        # true: but these two distinct objects have equal values

#Klase mogu imati i === operator koji se korsiti za testiranje pripadnosti, obično nekom regular expressionu
# po defaultu u većini klasa je isti kao ==

#U rubiju da bi uradili sort, koristimo <=> operator , koji naravno možemo overrajdovati
1 <=> 5     # -1
5 <=> 5     # 0
9 <=> 5     # 1
"1" <=> 5   # nil: integers and strings are not comparable

#Kastovanje objekata
#Većina klasa ima već implementirane standarne kastove
# to_s, to_i, to_f, to_a
#Može se korisiti i try_convert metoda, u slučaju kada nismo sigurno da je moguće uraditi konverziju
Array.try_convert(o) #returns o.to_ary if o defines that method

#True i False takođe imaju to_s koji se može pozivati i vraća string "true" ili "false"

#Objekte je moguće kopirati reko .clone da bi napravili novi objekat
#dup takođe pravi kopiju, ali to čini i kod freeze-ovanih objekata

#Objekte je moguće zalediti, to jest načiniti ih da budu immutable
#Dakle nije moguće promenti njegove promenjive
#Čak i poziv setera neće uspeti
s = "ice"      # Strings are mutable objects
s.freeze       # Make this string immutable
s.frozen?      # true: it has been frozen
s.upcase!      # TypeError: can't modify frozen string
s[0] = "ni"    # TypeError: can't modify frozen string

#Treba držati u vidu da ne postoji način da neki objekat un-freezujemo nakon freezovanja
#Jedino rešenje je da napravimo duplikat preko dup funkcije, koji nije freezovan


#Možemo i taintovati neki objekat
#Naprimer ukoliko imamo neki user input, ili neki SQL koji je user mogao da menja
#Svaki objekat koji je na neki način nastao iz ovog objekta je takođe taitnted
#Ruby limitira neke metode

s = "untrusted"   # Objects are normally untainted
s.taint           # Mark this untrusted object as tainted
s.tainted?        # true: it is tainted
s.upcase.tainted? # true: derived objects are tainted
s[3,4].tainted?   # true: substrings are tainted

#Generalno ovo treba koristit sa globalnom varijablom $SAFE da uključujemo i isključujemo razne napredne ruby komande koje mogu biti pozvane

#Nesto zanimljivo 

#Možemo podesiti nesto ukoliko nije null

#eeww... we don't like this one.
if(x.nil?)
  x = get_some_object
end

#možemo uraditi oneliner preko if
x = get_some_object if x.nil?

#ili zapravo iskoristiti ||= operator
x ||= get_some_object


