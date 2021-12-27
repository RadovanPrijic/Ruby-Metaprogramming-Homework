#Komentari
# This entire line is a comment
x = "#This is a string"              # And this is a comment 
y = /#This is a regular expression/  # Here's another comment

#Ruby nema multiline komentare kao C/Java to jest /* */ sintaksu

#
# Ovako radimo multiline comment
#


# Postoje i embedovani dokumenti, npr ukoliko hocemo da cuvamo HTML/SQL itd
# Moze se koristiti za brzo komentarisanje koda itd...

 =begin This used to begin a comment. 
       Now it is itself commented out!
	   The code that goes here is no longer commented out
 =end
 
# Literali, vrednosti koje se pojavljuju direktno u kodu, takođe su expressioni

1
1.0
'one'
"two"
/three/

#Ruby identifier, je ime neke promenjive/metode/klase itd.
#mogu sadržati brojeve, slova, underscore
#ne smeju početi brojem, sadržati whitespace, i specijalne znake

#Ruby podrazumeva da ukoliko je neka promenjiva napisana velikim slovom, da je konstanta
#Ruby interpreter će izdati upozorenje ukoliko želimo da menjamo ove konstante tokom izvršenja koda
# Klase i moduli počinju sa Velikim slovom
# Promenjive sa više reči pišemo sa _ , dakle na_primer 
#Ova pravila se važe samo pri unicode encodingu našeg koda
#ukoliko bi koristili UTF8, mogli bi stavljati i specijalne karaktere na početak naprimer čćšđž

i
x2
old_value
_internal #mozemo korostiti underscore za private stvari, međutim interpreteru ništa ne znači
PI        #konstanta


#Ruby je case sensitive jezik, naprimer end i END su totalno drugačije stvari

#Varijable generalno mogu imati prefikse i "sufikse"

$files 		#globalna varijabla
@data		#instance varijabla
@@counter	#Klasna varijabla
empty?		#Metoda koja vraća bool
sort!		#"Destruktivna" metoda
timeout=	#metoda koja podešava neku promenjivu

#Rubz ima rezervisane reči, kao i većina jezika
__LINE__
__Encoding__
__FILE__
BEGIN 
END
alias
and
begin
break
case
class
def
defined?
do
else
elsif

#I tako dalje... https://www.geeksforgeeks.org/ruby-keywords/

#Whitespace se koristi u rubiju da se delimitiraju identifieri generalno, većina whitespaca se ignorise od strane interpretera
#Newline igra bitnu ulogu u Ruby jeziku, služi kao statement terminator
#Dakle newline igra istu ulogu kao ; u drugim programskim jezicima
#Takođe se može koristiti i ; kao statement terminator, ali generalno se ommituje kao deo konvencije


 funkcija( "aaa", "bbb", "ccc"
		   "ddd", "eee", "fff")
#Generalno u programiranju se javljaju slucajevi da hocemo da koristimo new line, kao nacin da nam kod bude pregledniji
#Naprimer pisanje puno parametara neke funkcije
#Potrebno je znati gde se stvalja new line, ne bi li se zbunio interpreter
#Generalno interpreter iz konteksta donosi odluku da li bi trebalo da termiise statement 

total = x +  #U ovom slučaju bi interpreter nastavio parsiranje
y

total = x	# ovo je potpuni expression, i rubi ce ga tako interpretirati
+y			# ovo se interpretira kao zasebni expression, i generalno je beskorisan

#Od Ruby 1.9 dodata je funkcionalnost da ukoliko nova linija počinje sa tačkom, da se interpretira kao deo prethodnog statementa
#Generalno ovo se dosta koristi kod API-ja

animals = Array.new
.push("dog")
.push("cow")
.push("cat")
.sort

a=[1,5,2,3]
b= a.sort
p a
p b

#Ovo je moguće jer svaka od ovih metoda vraća niz, nad kojim je izvršena operacija

#Ruby dozvoljava da ne koristimo zagrade pri pozivu metoda, međutim ovo uvodi problematiku whitespace-a
f(3+2)+1 #Radi kao očekivano prosledjuje expression 3+2 kao parametar, i nakon toga dodaje 1
f (3+2)+1 # (3+2)+1 će biti shvaćeno kao jedan expression i prosleđeno kao parametar u f
		  # Generalno ruby bi trebalo da izda upozorenje pri drugom slučaju

#pravila
#Nemojte stavljati space pre zagrada
#Ukoliko jedan od expressiona koji je parametar, ima zagradu, uvek treba koristiti zagrade pri pozivu metode
#Možete kod pokretati sa ruby -w da bi dobili warninge za ovakve slučajeve


#Do sada smo spominjali expressione, kao osnovnu funkcionlanost ruby jezika, i njegovu specifičnost
#Generalno bilo koji kod koji se izvršava i vraća neku vrednost
#Literali su takođe expressioni, jer vraćaju neku vrednost

[1,2,3]   			   # An Array literal
{1=>"one", 2=>"two"}   # A Hash literal
1..3                   # A Range literal

#Kompleksniji expressioni mogu biti napravljeni korišćenjem operatora
1         # A primary expression
x         # Another primary expression
x = 1     # An assignment expression
x = x + 1 # An expression with two operators

#Primenom expressiona sa određenim keywordovima dobijamo statemente, npr, if while itd

if x < 10 then   # If this expression is true
    x = x + 1      # izvrši ovaj statement
  end              
  
while x < 10 do  # While this expression is true...  
  print x        #  izvrši ovaj statement
  x = x + 1      #  izvrši ovaj statement
end              

#Generalno statement je nešto što se izvršava usled nekog expressiona
#Više expressiona i statementa -> funkcije, metode, lambde itd...

#Ruby programi generalno imaju blokovnu strukturu, kao i većina drugih jezika
#Po ruby konvenciji umesto taba, generalno koristimo dva spejsa
#Generalno postoje dva tipa bloka
#Jedan tip su blokovi koji imaju parametre, i generalno su iteratorskog tipa

3.times { print "Ruby!!"}

#Block je u ovom slučaju sve između {} 
#Praksa je da se za multiline blokove koriste DO i END keywordovi, dok za single line {}

1..10 do |x|
  print x
end

#Drugi tip bloka je generalno deklaracionog karaktera, i služi da opiše početak i kraj nečega, npr klasa itd..
#Ovaj tip obično nema zagrade već sadrži keywordove

module Stats                          # A module  
  class Dataset                       # A class in the module   
    def initialize(filename)          # A method in the class    
      IO.foreach(filename) do |line|  # A block in the method        
        if line[0,1] == "#"           # An if statement in the block          
          next                        # A simple statement in the if        
        end                           # End the if body      
      end                             # End the block    
    end                               # End the method body  
   end                                # End the class body
end                                   # End the module body

#Struktura ruby fajla

#Ruby fajl može početi sa shebang komentarom, koji kaže unixolikim OSovima kako treba da izvrše fajl
#sledeća linija može biti coding komentar, kojim se generalno podešava encoding
# coding: utf-8
#Ruby fajl može imati __END__ keyword, do koga će se program izvršavati, sve ispod se ignoriše
#Sve ispod __END__ se može tretirati kao program data, i pristupati mu preko DATA objekta, npr DATA.read

#Generalno retko kada se sve stavlja u jedan ruby fajl, obično se pišu mali ruby fajlovi, i koristi veliki broj biblioteka i importa


#Izvršenje Ruby koda
	
#Ruby je skriptni jezik
#Ruby programi su ništa više nego lista statementa koje treba izvršiti
#Ovi statementi su generalno sekvencijalne prirode, dakle po redosledu u kom se pojavljuju
#Kao i svi drugi jezici izvršenje koda se modifikuje preko kontrolnih struktura

#Ne postoji main metod, od koga se počinje, već se interpreteri jednostavno prosledi fajl, koji on od početka počinje da izvršava
#Za svaki modul,klasu, i metod. Ruby izvršava kao bilo koji drugi statement kada stigne do njega.
#"Izvršavanjem" ove metode, Ruby za nas pravi "objekat" to jest definiciju te metode kojem možemo pristupiti
#Pozivom ove metode, ruby će ući u blok te metode, i izvršiti ga

#Ruby interpreter se poziva preko komandne linije, i prosleđuje mu se skripta za izvršenje
#Obično se ova skripta prosledjuje u obliku putanje do skript fajla
#Nakon ovoga Ruby interpreter ulazi u fajl i počinje da ga obrađuje
#Prvo executuje sve Begin blokove #Mala optimizacija
#Vraća se na početak i počinje sa izvrešenjem svake linije, sve dok
	#Ne dodje do kraja fajla
	#Ne dodje do komande kojom se prekida izvršenje fajla
	#Pročita logički kraj fajla označen sa __END__