#Svaki objekat je instanca klase, svaka klasa ima metode na koje se odaziva
#Klase mogu da nasledjuju druge klase, i da overrajdruju metode superklase
#Klase takodje mogu da nasledjuju metode modula

#Ruby objekti su strogo enkapsulirani. Stanju objekta se može pristupiti samo preko metoda koje definiše.
#Varijable neke instance, nisu dostupne van te instance. (Java private variables)
#Da bi pristupili ovim objektima pravimo getter-setter metode. 
#Metode unutar klase mogu imati "Public" "protected" i "private" visibility

#Za razliku od standardnih jezika Rubijeve klase su otvorene, dakle svaki rubz program može dodati metode u već postojeću klasi, kao i singleton metode na određeni objekat
#Ruby klase se deklarisu sa kljucnom reci class

#class Point 
#end

#self kljucna rec se odnosi na samu klasu, kada se poziva van metoda, kada se poziva unutar metoda onda je objekat
#kljucna rec class je takodje expression, i vraća poslednji expression unutar klase, obicno poslednju metodu
#kljucna rec def uvek vraća nil

#Klase instanciramo sa new metodom
p = Point.new

#Sve klase imaju metodu new, koja vraća novu instancu
p.class # => Point
p.is_a? Point # => true

#Konsturktur metoda u rubiju se zove initialize
#Dakle klasa poziva new metodu, i onda za tu instancu poziva initialize metodu, sa svim parametrima new
#Treba napomenuti da je initialize private funkcija, po defaultu
def initialize(x,y)
  @x, @y = x, y
end

#Zbog razlike u self kljucnoj reci bitno je gde definišemo naše instancne varijable
#Ovaj kod u rubiju ne radi ono što mislimo da radi
class Point
  @x = 0 # Create instance variable @x and assign a default. WRONG!
  @y = 0 # Create instance variable @y and assign a default. WRONG!
  def initialize(x,y)
    @x, @y = x, y # Now initialize previously created @x and @y.
  end
end

#Varijable u rubiju se uvek dodeljuju na self scope, dakle, ukoliko su van funkcije bice bindovane na klasu, a ne na instancu
#Dakle ukoliko varijable inicijalizujemo unutar metode, onda postaju instance varijable


# .to_s
class Point
  def initialize(x,y)
    @x, @y = x, y
  end
  def to_s # Return a String that represents this point
    "(#@x,#@y)" # Just interpolate the instance variables into a string
  end
end

#Podesavanje gettera i settera
class Point
  def initialize(x,y)
    @x, @y = x, y
  end
  
  def x # The accessor (or getter) method for @x
    @x
  end
  
  def y # The accessor method for @y
    @y
  end
end

#Postoji i kraca notacija za definiciju metoda 
def x; @x; end # The getter method for @x

#Ovako bi onda koristili gettere
p = Point.new(1,2)
q = Point.new(p.x*2, p.y*3)


#Setteri
class MutablePoint
def initialize(x,y); @x, @y = x, y; end
  def x=(value) # The setter method for @x
    @x = value
  end
  
  def y=(value) # The setter method for @y
    @y = value
  end
end

class Klasa 

	@x

	
	def func()
		@x=2
	end


end



#Kada deklarišemo ovako promenjive nemožemo im pristupati preko gettera unutar metode klase
#Dakle ne bi bilo tačno napisati x=2, ovo će samo napraviti novu promenjivu x
#Ukoliko baš želimo da koristimo setter, onda radimo ovako self.x=2 ili jednostavnije @x=2


#Česta je praksa da su getteri setteri proste standardne klase, bez dodatne logike
#Za generisanje ovih stanradnih gettera settera imamo kracu nostaciju
#Dakle koristimo attr_accessor i attr_reader kome prosledimo simbole naših instancnih varijabli
class Point
  attr_accessor :x, :y # Define accessor methods for our instance variables
  attr_reader :x, :y # Define reader methods for our instance variables
end 

#Ima i jos kraci nacin...
attr :x #definiše getter za @x
attr :x, true #definiše i getter i setter za @x

#Ova funkcionalnost je omogucena preko meta programiranja, gde pozivom attr funkcije, ona našoj klasi dodaje metode, i sama piše kod

#Ukoliko hoćemo da iteriramo kroz našu klasu Point, moramo da definišemo each funkciju 
def each
  yield @x
  yield @y
end

#Iteriramo ovako
p = Point.new(1,2)
p.each {|x| print x } # Prints "12"

#Inkludovanjem Enumerable modula, dobijamo pristup preko 20 iteratora
p.all? {|x| x == 0 } # True if the block is true for all elements

#Equality proveravamo preko ==

def ==(o) # Is self == o?
  if o.is_a? Point # If o is a Point object
    @x==o.x && @y==o.y # then compare the fields.
  elsif # If o is not a Point
    false # then, by definition, self != o.
  end
end


#Ukoliko želimo da poredimo objekte treba da inkludujemo comparable
#Domparable nam omogućava da nam funkcionišu > < == operatori, takođe  dozvoljava sort, min, max itd..
#Da bi on radio moramo implementirati <=> funkciju

def <=>(other)
return nil unless other.instance_of? Point
@x**2 + @y**2<=>other.x**2 + other.y**2
end

#Ruby takođe podržava neki vid "structova" ...ovo je način da brzo generišemo anonimne klase
#Struct ce napraviti klasu, sa promenjivama i getterima, sa nekim imenom
Struct.new("Point", :x, :y) # Creates new class Struct::Point
Point = Struct.new(:x, :y) # Creates new class, assigns to Point
p = Point.new(1,2)
p.x # => 1
p.y # => 2
p.x = 3 # => 3
p.x # => 3

#Structovi takođe automatski podržavaju i [] i []= notaciju. Takođe imaju i iterator
q = Point.new(4,2)
q == p # => true
h = {q => 1} # Create a hash using q as a key
h[p] # => 1: extract value using p as key
q.to_s # => "#<struct Point x=4, y=2>"

#Možemo čak i prošitiri naše Structove
Point = Struct.new(:x, :y) # Create new class, assign to Point
class Point # Open Point class for new methods
  def add!(other) # Define an add! method
	self.x += other.x
    self.y += other.y
    self
  end
  
  include Comparable # Include a module for the class
  def <=>(other) # Define the <=> operator
     return nil unless other.instance_of? Point
     self.x**2 + self.y**2 <=> other.x**2 + other.y**2
  end
end

#Kreiramo Singleton metode preko ImeKlase.metoda

def Point.sum(points) # Return the sum of an arbitrary number of points
  x = y = 0
  points.each {|p| x += p.x; y += p.y }
  Point.new(x,y)
end


#takođe mogu biti definisane i preko self
def self.sum(*points) # Return the sum of an arbitrary number of points
  x = y = 0
  points.each {|p| x += p.x; y += p.y }
  Point.new(x,y)
end

#Ima i treći način...
class << Point # Syntax for adding methods to a single object
  def sum(*points) # This is the class method Point.sum
    x = y = 0
    points.each {|p| x += p.x; y += p.y }
    Point.new(x,y)
  end

end

#Klasne varijable, to jest u drugim jezicima bi ovo bilo static
class Point
  # Initialize our class variables in the class definition itself
  @@n = 0 # How many points have been created
  @@totalX = 0 # The sum of all X coordinates
  @@totalY = 0 # The sum of all Y coordinates
  def initialize(x,y) # Initialize method
    @x,@y = x, y # Sets initial values for instance variables
    # Use the class variables in this instance method to collect data
    @@n += 1 # Keep track of how many Points have been created
    @@totalX += x # Add these coordinates to the totals
    @@totalY += y
  end
end

#U rubiju metode mogu biti public,private i protected
#Metode su po defaultu public

#Private može pozvati samo ta klasa i njene podklase, nemože se bindovati na self (nemoze biti static)
#Protected je isto sto i private ,ali može se bindovati na self, dakle može biti statička kojoj pristupaju sve instance klase i podklasa

class Point
  # public methods go here
  
  # The following methods are protected
  protected
  
  def protected_method()
  end
  # protected methods go here
  
  # The following methods are private
  private
  
  def private_method()
  end
  
  # private methods go here
end

#Postoji i drugi način da eksplicitno navedemo
class Widget
  def x # Accessor method for @x
    @x
  end

  protected :x # Make it protected

  def utility_method # Define a method
    nil
  end
  
  private :utility_method # And make it private
  
end

#Međutim preko metaprogramiranje i refleksije svakako je moguce pristupiti ovim metodama
w = Widget.new # Create a Widget
  w.instance_eval { utility_method } # Another way to invoke it
w.instance_eval { @x } # Read instance variable of w

#Ukoliko želimo da izbegnemo pozivanje private funkcija iz nekog razloga možemo pozvati w.public_send


#Ruby ima nasledjivanje, samo jednu klasu je moguce naslediti
#Object/BasicObject je osnovna klasa koju svi nasledjuju

#Nasledjivanje u rubiju se radi preko < karaktera

class Point3D < Point
end

#Nasledjivanje funkcionise standardno, nasledjuje se initialize metoda, i sve metode koje vec poseduje
#Ima i super funkciju
#Ukoliko pozovemo super bez argumenata, automatski se svi argumenti naseg initialize prosledjuju superu
class Point3D < Point
  def initialize(x,y,z)
    # Pass our first two arguments along to the superclass initialize method
    super(x,y)
    # And deal with the third argument ourself
    @z = z;
  end
end

#Promenjivama iz nadklase, i metoama iz nad klase pristupamo kao da su unutar nase funkcije
#Treba držati na umu da ukoliko iskoristimo promenjivu sa istom nazivom, pregazićemo i varijablu od ancestora
#Dakle treba paziti pri nazivu instnace varijabli, inace ce verovatno biti bugova
class Point3D < Point
  def initialize(x,y,z)
    super(x,y)
    @z = z;
  end

  def to_s
    "(#@x, #@y, #@z)" # Variables @x and @y inherited?
  end
end


#Ukoliko želimo da napravimo "factory" metode, onda jednostavno proglasimo new za private
#Tako da se objekat može instancirati samo preko tih metoda
class Point
  # Define an initialize method as usual...
  def initialize(x,y) # Expects Cartesian coordinates
    @x,@y = x,y
  end

  # But make the factory method new private
  private_class_method :new
  
  def Point.cartesian(x,y) # Factory method for Cartesian coordinates
    new(x,y) # We can still call new from other class methods
  end
  
  def Point.polar(r, theta) # Factory method for polar coordinates
    new(r*Math.cos(theta), r*Math.sin(theta))
  end
end


#Drugi način da napravimo objekat je preko dup ili clone
#Ove dve metode prave shallow kopije, to jest za sve varijable unutar nekog objekta će biti uzete samo reference a ne vrednosti
#Da promenimo ovo ponašanje koristimo initialize_copy metodu

class Point # A point in n-space
  def initialize(coords) # Accept an arbitrary # of coordinates
    @coords = coords # Store the coordinates in an array
  end

  def initialize_copy(orig) # If someone copies this Point object
    @coords = orig.coords.dup # Make a copy of the coordinates array, too
  end
end

#U rubiju da definišemo singleton, nasledjujemo modul Singleton
#Ovo nam dodaje klasnu metodu .instance, međutim ova klasa ne prima promenjive
#Ova nova metoda automatski za nas kreira tu jednu instancu, koja je dostupna
#Takodje automatski cini new da bude private

Point.instance.print(self)

#Moduli su skup varijabli i metoda, slicno kao klase
#Moduli ne mogu da budu instancirani niti da budu nasledjeni
#Dakle ne postoji hijerarhija modula
#Modul je instanca Modul klase, a Class je podklasa Modula. Dakle, klasa je modul sa dodatnim funkcionalnostima

#Module je moguće koristiti kao namespace, i kao mixin

#Naprimer lepo je grupisati metode u odredjene grupe, u slucajevima kada je objektno progrmairnaje nepotrebno
#Ovako izbegavamo globalni namespace, i funkcije koje su vec implementirane. Naprimer mozemo napraviti encode i decode funkciju
#Ove metode ispod su deklarisane kao klasne varijable i pristupamo im preko imena namespejsa
#Taodje smo umesto self mogli lako da napisemo i Base64, to jest ime klase
module Base64
  def self.encode
  end
  
  def self.decode
  end
end

text = Base64.encode(data)
data = Base64.decode(text)

#Mixin nacin je drugaciji, i ponasa sse kao infterfejs u javi odprilike
#Za ovakvku funkcionalnost koristimo include ImeModula

#Zanimljivo je recimo da je include zapravo funkcija
class Point
  include Comparable
  
end

#Treba drzati na umu i klase kojima gledamo sta je neka klasa
#instance_of gleda samo klase dok is_a gleda i module
"text".instance_of? Comparable # => false
"text".is_a? Comparable # => true


#Generalno moduli/klase se obicno nalaze u različitim fajlovima, da bi povezali ove fajlove korismo require i load
#require je generalno pametniji, on trazi po instalacionom folderu, pokusava da ne loaduje dva puta istu stvar, itd...
#Load/Require takodje primaju putanje do ovih fajlova

#Ukoliko želimo da olakšamo celu stvar možemo podesiti rubz load_path preko $LOAD_PATH varijable
#Takođe može se koristit i $:
#Po defaultu je ovo mesto gde se instalira ruby


#Ruby specifično hendla klasne metode, on to čini tako što svaki objekat ima svoju eigenklasu, koja zapravo sadrži ove statične metode
#Takođe se može zvati singleton klasa ili metaklasa

#Da bi pristupili eigenklasi i dodavali joj metode možemo koristit << notaciju

class << Point
  def class_method1 # This is an instance method of the eigenclass.
  end # It is also a class method of Point.
  def class_method2
  end
end

#Unutar klase možemo pristupiti eigenklasi ovako

class Point
  # instance methods go here
  class << self
    # class methods go here as instance methods of the eigenclass
  end
end

class Point # Create or open the class Point
class Point3D < Point # Create a subclass of Point
class << Point # Open the eigenclass of the object Point

#Lookup metoda je mehanizam koji ruby koristi da nadje određenu metodu po imenu unutar neke klase, ili njenih nadklasa/modula
#1. Prvo se proverava Eigenklasa za metodu koju tražimo
#2. Ukoliko nije tamo, traži se kao instance method
#3. Ukoliko ga nema u klasi, Traži se u svim modulima, Pretražuje se hrnološki, prvo one koji su prvi includovani
#4. Ukoliko ga nema u modulima, traži se u svim nadklasama, i za svaku od njih se ponavlja korak 1-3
#5. Ukoliko ne postoji ovaj metod onda se traži metod method_missing(), to jest ponavlja korak 1-4 za method_missing. Ukoliko ga nema, poziva se default koji nudi kernel

#Generalno ovaj mehanizam se dosta koristi u metaprogramiranju, i method_missing() možemo koristiti kao fallback
#To jest ukoliko nema ta metoda kreiraj je itd...










