#Ruby je dinamicki jezik, moguće je ubaciti metode u klase pri runtajmu, kreiranje aliasa za metode, čak i definicija metoda nad objektom.
#Biblioteka koja je osnov ove funkcionalnost je biblioteka koja se bavi refleksijom
#Refleksija je mogućnost programa da "inspektuje" svoje trenutno stanje i svoju strukturu.
#Dakle veoma lako možemo recimo videti sve metode na koje se odaziva Hash klasa, ili da dobijemo vrednost neke promenjive od specificne instance

#Preko reflection API-ja takođe možemo i da menjamo stanje i strukturu programa. Možemo dodavati varijable, pozivati metode itd...
#Generalno ovaj API je kljucan za bilo koji vid metaprogramiranja, to jest programa koji pisu same sebe

#Najprostije refleksivne metode
o.class
c.superclass
o.instance_of? c #Da li je određene klase
o.is_a? c #Da li je određena klasa, ili jedna od njenih podklasa. Ukoliko je C modul, testira se da li sadrži modul
o.kind_of? #Sinonim sa is_a?
c === o #Po defaultu poziva se is_a?
o.respond_to? name #Da li objekat o ima puiblci ili protected metod sa odre]enim imenom. 



module A; end # Empty module
module B; include A; end; # Module B includes A
class C; include B; end; # Class C includes module B
C < B # => true: C includes B
B < A # => true: B includes A
C < A # => true
Fixnum < Integer # => true: all fixnums are integers
Integer < Fixnum # => false: not all integers are fixnums
String < Numeric # => nil: strings are not numbers, napomena da je null u bilo kom expressionu zapravo false

#Alternativna notacja
C.include?(B) # => true
C.included_modules # => [B, A, Kernel]
#Treba držati u vidu da je include kojim mixinujemo modul zapravo obična privatna funkcija

A.ancestors # => [A]
B.ancestors # => [B, A]
C.ancestors # => [C, B, A, Object, Kernel]
String.ancestors # => [String, Enumerable, Comparable, Object, Kernel]

#Moguće je promeniti ili dodati nasledjivanje nekom objektu
module Greeter; def hi; "hello"; end; end # A silly module
s = "string object"
s.extend(Greeter) # Add hi as a singleton method to s
s.hi # => "hello"
String.extend(Greeter) # Add hi as a class method of String
String.hi # => "hello"

#Class i Module su klase kao i svake druge, i moguce ih je dinamicki kreirati, i dodavati im funkcionalnosti
M = Module.new # Define a new module M
C = Class.new # Define a new class C
D = Class.new(C) { # Define a subclass of C
  include M # that includes module M
}
D.to_s # => "D": class gets constant name by magic


#Zavisno od konteksta u kome se nalazimo možemo pozivati funkcije local_variables, class_variables i global_variables
global_variables # => ["$DEBUG", "$SAFE", ...] Ovo dobijamo gde god da pozovemo ovu funkciju
x = 1 # Define a local variable
local_variables # => ["x"]

# Define a simple class
class Point
  def initialize(x,y); @x,@y = x,y; end # Define instance variables
  
  @@classvar = 1 # Define a class variable
  ORIGIN = Point.new(0,0) # Define a constant
end

Point::ORIGIN.instance_variables # => ["@y", "@x"]
Point.class_variables # => ["@@classvar"]
Point.constants # => ["ORIGIN"]


#Eval funkcija je jedna od stvari koja se najviše koristi pri testiranju, i proverama
#Dakle Eval funkcija se ponaša tako što uzima bilo koji string, i izvršava ga kao da je kod
x = 1
eval "x + 1" # => 2

#Primetite da eval uzima kontekst u kome se nalazi, i svestan je svega što postoji
#Moguće je čak i podesiti kontekst u kome će se kod izvršiti.
#Kernel nam nudi binding promenjivu, koju kada prosledimo dobijamo tačno kontekst u kojem je pozvan
#Dakle u primeru ispod, binding vraćamo dok smo unutar Object klase(Koju smo proširili), koju sve klase nasledjuju
#Kao parametar eval prosledjujemo binding, od određenog objekta, koji u ovom slučaju predstavlja kontest klase

class Object # Open Object to add a new method
  def bindings # Note plural on this method
    binding # This is the predefined Kernel method
  end
end

class Test # 
  def initialize(x); @x = x; end
end

t = Test.new(10) # Create a test object
eval("@x", t.bindings) # => 10: We've peeked inside t

#Ukoliko nemamo neki egzotični slučaj, možemo pojednostaviti ovaj gore navedeni slučaj
#U većini slučajeva želimo da eval izvrši neki objekat ili klasa, to jest da podesimo sta mu tačno znači self

o.instance_eval("@x") # Return the value of o's instance variable @x

# Define an instance method len of String to return string length
String.class_eval("def len; size; end")
# Use instance_eval to define class method String.empty
String.instance_eval("def empty; ''; end")

#Malo je kontraintuitivno, class_eval kreira nove metode, dok instance_eval kerira nove static metode

#Ove dve metode takođe mogu da prime i blok, dakle nema potrebe da pišemo string
o.instance_eval { @x }
String.class_eval {
  def len
    size
  end
}
String.class_eval { alias len size }
String.instance_eval { def empty; ""; end }

#Ruby nam takođe nudi metode za pristup promenjivama, menjanje tih promenjivih, cak i uklanjanje promenjivih
o = Object.new
o.instance_variable_set(:@x, 0) # Note required @ prefix
o.instance_variable_get(:@x) # => 0
o.instance_variable_defined?(:@x) # => true

Object.class_variable_set(:@@x, 1) 
Object.class_variable_get(:@@x) 
Object.class_variable_defined?(:@@x) # => true; Ruby 1.9 and later

Math.const_set(:EPI, Math::E*Math::PI)
Math.const_get(:EPI) # => 8.53973422267357
Math.const_defined? :EPI # => true

#Za svrhu uklanjanja moramo koristit send ili eval metode
o.instance_eval { remove_instance_variable :@x }
String.class_eval { remove_class_variable(:@@x) }
Math.send :remove_const, :EPI # Use send to invoke private method


#Možemo gledati i menjati sve metode unutar neke klase/objekta
o = "a string"
o.methods # => [ names of all public methods ]
o.public_methods # => the same thing
o.public_methods(false) # Exclude inherited methods
o.protected_methods # => []: there aren't any
o.private_methods # => array of all private methods
o.private_methods(false) # Exclude inherited private methods
def o.single; 1; end # Define a singleton method
o.singleton_methods # => ["single"] (or [:single] in 1.9)

String.instance_methods == "s".public_methods # => true
String.instance_methods(false) == "s".public_methods(false) # => true

String.public_method_defined? :reverse # => true
String.protected_method_defined? :reverse # => false
String.private_method_defined? :initialize # => true
String.method_defined? :upcase! # => true

#Možemo takođe dobiti i metod objekat koji možemo zvati i dalje inspektovati
metoda= "s".method(:reverse) # => Method object
metoda.call
 
#Naravno postoji i lakši način, koristimo send
#Send izvršava ime metode koja je prvi argument, sve ostalo su argumenti
#Preko send takođe možemo lako pristupati i private i protected metodama
"hello".send :upcase # => "HELLO": invoke an instance method
Math.send(:sin, Math::PI/2) # => 1.0: invoke a class method

#Da bi dodali neku metodu u samu klasu, moramo biti unutar same klase
#Za ovo se koristi define_method koji je private, dakle mora se biti unutar klase da bi se pozvao
# Add an instance method named m to class c with body b
#Za uklanjanje metoda možemo jednstavno pozvati samo remove_method
def add_method(c, m, &b)
  c.class_eval {
    define_method(m, &b)
  }
end
add_method(String, :greet) { "Hello, " + self }
"world".greet # => "Hello, world"

#Možemo dodavati i aliase, međutim ovo je takođe private metoda
def backup(c, m, prefix="original")
  n = :"#{prefix}_#{m}" # Compute the alias
  c.class_eval { # Because alias_method is private
    alias_method n, m # Make n an alias for m
  }
end
backup(String, :reverse)
"test".original_reverse # => "tset"

#Možemo method_missing da koristimo da "dinamički" kreiramo metode, kako se pozivaju
#Naprimer, poziv unutar Hash-a, da koristi ključ, kao metodu
class Hash

  def method_missing(key, *args)
    text = key.to_s
    if text[-1,1] == "=" # If key ends with = set a value
	  self[text.chop.to_sym] = args[0] # Strip = from key
	  else
	    self[key] # ...just return the key value
	  end
	end
	
end

h = {} # Create an empty hash object
h.one = 1 # Same as h[:one] = 1
puts h.one # Prints 1. Same as puts h[:one]

#Ruby nam takođe nudi osnovne Hook-ove. U suštini call-back funkcije koje se izvršavaju kada se ispuni neki uslov.
#Osnovni hook je za nasledjivanje, kada se neka naša klasa nasledi, pozvaće se .inherited
#Na ovaj način možemo dodati dodatnu funkcionalnost ili ispitati da li ta klasa ispunjava određene uslove
#Naprimer da printamo kada god se nasledi naša klasa
def Object.inherited(c)
  puts "class #{c} < #{self}"
end

#Kod inkludovanja se dobija dodatna funkcionalnost, pošto inkludovanje modula može menjati klasu
#Dakle svaki modul koji inkludujemo može ne samo da doda metode, već može i da menja naše metode
#Naprimer Final modul, koji zabranjuje nasledjivanje

module Final # A class that includes Final can't be subclassed
  def self.included(c) # When included in class c
    c.instance_eval do # Define a class method of c
      def inherited(sub) # To detect subclasses
        raise Exception, # And abort with an exception
        "Attempt to create subclass #{sub} of Final class #{self}"
      end
    end
  end
end

#Postoji i hook koji se poziva svaki put kada dodajemo metodu
#Takođe postoji i sličan hook samo za singleton metode
def String.method_added(name)
  puts "New instance method #{name} added to String"
end

# Including this module in a class prevents instances of that class
# from having singleton methods added to them. Any singleton methods added
# are immediately removed again.
module Strict
  def singleton_method_added(name)
    STDERR.puts "Warning: singleton #{name} added to a Strict object"
    eigenclass = class << self; self; end
    eigenclass.class_eval { remove_method name }
  end
end

#Nije bas metaprogramiranje, ali može biti korisno, kod debagovanja i genralno eksperimentacije
#ObjectSpace modul definiše par korisnih metoda kao each_object
#Ova funkcija vraća sve objekte neke klase, za koje zna interpreter

# Print out a list of all known classes
ObjectSpace.each_object(Class) {|c| puts c }

#ObjectSpace._id2ref vraća objekat za neki ID koji smo mu prosledili

#Dodatne funkcionalnosti se tiču garbage collerctora
#Nudi nam metodu za finalizaciju, pred garbage collection tog objekta
#ObjectSpace.define_finalizer
#Međutim ne smemo da koristimo objekat koji je odabran za garbage collection već samo druge objekte
#Takođe postoji funkcija direktnog poziva nad garbage collectorom ObjectSpace.garbage_collect



#Primeri
#Generisanje metoda

class CarModel
  FEATURES = ["engine", "wheel", "airbag", "alarm", "stereo"]

  FEATURES.each do |feature|
    define_method("#{feature}_info=") do |info|
      instance_variable_set("@#{feature}_info", info)
    end

    define_method("#{feature}_info") do
      instance_variable_get("@#{feature}_info")
    end

    define_method "feature_price=" do |price|
      instance_variable_set("@#{feature}_price", price)
    end

    define_method("#{feature}_price") do
      instance_variable_get("@#{feature}_price")
    end
  end
end


#Takođe smo mogli i lakše

class CarModel
  # define a class macro for setting features
  def self.features(*args)
    args.each do |feature|
      attr_accessor "#{feature}_price", "#{feature}_info"
    end
  end

  # set _info and _price methods for each of these features
  features :engine, :wheel, :airbag, :alarm, :stereo
end

#Preko method_missing možemo dinamički praviti nove promenjive
class CarModel
  def method_missing(name, *args)
    name = name.to_s
    super unless name =~ /(_info|_price)=?$/
    if name =~ (/=$/)
      instance_variable_set("@#{name.chop}", args.first)
    else
      instance_variable_get("@#{name}")
    end
  end
end

@car_model = CarModel.new

@car_model.stereo_info    = "CD/MP3 Player"
@car_model.stereo_price   = "£79.99"

@car_model.stereo_info  # => "CD/MP3 Player"
@car_model.stereo_price # => "£79.99"

#Dinamičko dodavanje metoda zahteva korišćenje principa Alias Chaininga
#1. Kreiramo Alijas a metodu koju želimo da menjamo. Ovaj alias nam nudi način da pristupimo nepromenjenoj metodi
#2. Kreiramo novu verziju metode. Ova verzija može pozvati orginalnu verziju pre ili posle promena koje želi da učini

#Prvi primer je tracing za svako korišćenje reči require i load. Koristi Object.inherited hook.
#Nakon pamćenja svog nasledjivanja, poziva Kernel.at_exit i printa rezultat

# We define this module to hold the global state we require, so that
# we don't alter the global namespace any more than necessary.
module ClassTrace
  # This array holds our list of files loaded and classes defined.
  # Each element is a subarray holding the class defined or the
  # file loaded and the stack frame where it was defined or loaded.
  T = [] # Array to hold the files loaded
  
  # Now define the constant OUT to specify where tracing output goes.
  # This defaults to STDERR, but can also come from command-line arguments
  if x = ARGV.index("--traceout") # If argument exists
    OUT = File.open(ARGV[x+1], "w") # Open the specified file
    ARGV[x,2] = nil # And remove the arguments
  else
    OUT = STDERR # Otherwise default to STDERR
  end
end

# Alias chaining step 1: define aliases for the original methods
alias original_require require
alias original_load load

# Alias chaining step 2: define new versions of the methods
def require(file)
  ClassTrace::T << [file,caller[0]] # Remember what was loaded where, caller[0] vraća ime metode koja je pozvala ovu metodu
  original_require(file) # Invoke the original method
end

def load(*args)
  ClassTrace::T << [args[0],caller[0]] # Remember what was loaded where
  original_load(*args) # Invoke the original method
end

# This hook method is invoked each time a new class is defined
def Object.inherited(c)
  ClassTrace::T << [c,caller[0]] # Remember what was defined where
end

# Kernel.at_exit registers a block to be run when the program exits
# We use it to report the file and class data we collected
Kernel.at_exit {
  o = ClassTrace::OUT
  o.puts "="*60
  o.puts "Files Loaded and Classes Defined:"
  o.puts "="*60
  ClassTrace::T.each do |what,where|
    if what.is_a? Class # Report class (with hierarchy) defined
      o.puts "Defined: #{what.ancestors.join('<-')} at #{where}"
    else # Report file loaded
      o.puts "Loaded: #{what} at #{where}"
    end
  end
}


#U dosta slucajeva u Ruby jeziku se praktikuje "izmisljanje" pomcnih jezika (Domain specific language)
#Rubijeva fleksibilna sintaksa nam dozvoljava da pisemo kod koji uopšte ne izgleda kao ruby, a ipak ima mapirane mocne funkcije iza sebe

#U ovom primeru hocemo da napravimo DSL za generisanje XML-a, to jest HTML-a
#Sintaksa jezika bi izgledala ovako
pagetitle = "Test Page for XML.generate"
XML.generate(STDOUT) do
  html do
    head do
      title { pagetitle }
      comment "This is a test"
    end
    body do
      h1(:style => "font-family:sans-serif") { pagetitle }
      ul :type=>"square" do
        li { Time.now }
        li { RUBY_VERSION }
      end
    end
  end
end

#Ovaj kod bi trebalo da generiše sledeći HTML
<html>
  <head>
    <title>Test Page for XML.generate</title>
    <!-- This is a test -->
    </head>
  <body>
    <h1 style='font-family:sans-serif'>Test Page for XML.generate</h1>
    <ul type='square'>
      <li>2007-08-19 16:19:58 -0700</li>
      <li>1.9.0</li>
    </ul>
  </body>
</html>



class XML

  # Create an instance of this class, specifying a stream or object to
  # hold the output. This can be any object that responds to <<(String).
  def initialize(out)
    @out = out # Remember where to send our output
  end
  
  # Output the specified object as a comment, return nil.
  def comment(text)
    @out << "<!-- #{text} -->"
    nil
  end

  # Output a tag with the specified name and attributes.
  # If there is a block invoke it to output or return content.
  # Return nil.
  def tag(tagname, attributes={})
    # Output the tag name
    @out << "<#{tagname}"
	
    # Output the attributes
    attributes.each {|attr,value| @out << " #{attr}='#{value}'" }
	
	if block_given?
	  # This block has content
	  @out << '>' # End the opening tag
      content = yield # Invoke the block to output or return content
      if content # If any content returned
        @out << content.to_s # Output it as a string
      end
      @out << "</#{tagname}>" # Close the tag
	else
	  # Otherwise, this is an empty tag, so just close it.
	  @out << '/>'
	end
    nil # Tags output themselves, so they don't return any content
  end
  
  # The code below is what changes this from an ordinary class into a DSL.
  # First: any unknown method is treated as the name of a tag.
  alias method_missing tag
  
  # Second: run a block in a new instance of the class.
  def self.generate(out, &block)
    XML.new(out).instance_eval(&block)
  end
end

#Možemo učiniti celu stvar pametnijom tako što ćemo napraviti bazičnu validaciju
#Deklarišemo HTMLForm koji sadrži sve što je dozvoljeno
#Dakle biće deklarisana funkcija element koja prima hash, sa dozvoljenim elementima
class HTMLForm < XMLGrammar
element :form,  :action => REQ,
                :method => "GET",
			    :enctype => "application/x-www-form-urlencoded",
			    :name => OPT
				
element :input, :type => "text", :name => OPT, :value => OPT,
	            :maxlength => OPT, :size => OPT, :src => OPT,
				:checked => BOOL, :disabled => BOOL, :readonly => BOOL
				
element :textarea, :rows => REQ, :cols => REQ, :name => OPT,
				   :disabled => BOOL, :readonly => BOOL
				   
element :button, :name => OPT, :value => OPT,
				 :type => "submit", :disabled => OPT
end

#PRimenom ovih pravila ovakav kod će biti obrađen
HTMLForm.generate(STDOUT) do
  comment "This is a simple HTML form"
  form :name => "registration",
       :action => "http://www.example.com/register.cgi" do
        content "Name:"
        input :name => "name"
        content "Address:"
        textarea :name => "address", :rows=>6, :cols=>40 do
          "Please enter your mailing address here"
        end
       button { "Submit" }
    end
end

class XMLGrammar
  # Create an instance of this class, specifying a stream or object to
  # hold the output. This can be any object that responds to <<(String).
  def initialize(out)
    @out = out # Remember where to send our output
  end
  
  # Invoke the block in an instance that outputs to the specified stream.
  def self.generate(out, &block)
    new(out).instance_eval(&block)
  end
  
  # Define an allowed element (or tag) in the grammar.
  # This class method is the grammar-specification DSL
  # and defines the methods that constitute the XML-output DSL.
  
  def self.element(tagname, attributes={})
    @allowed_attributes ||= {}
    @allowed_attributes[tagname] = attributes
    class_eval %Q{
      def #{tagname}(attributes={}, &block)
        tag(:#{tagname},attributes,&block)
      end
    }

  # These are constants used when defining attribute values.
  OPT = :opt # for optional attributes
  REQ = :req # for required attributes
  BOOL = :bool # for attributes whose value is their own name
  
  def self.allowed_attributes
    @allowed_attributes
  end
  
  # Output the specified object as CDATA, return nil.
  def content(text)
    @out << text.to_s
    nil
  end
  
  # Output the specified object as a comment, return nil.
  def comment(text)
    @out << "<!-- #{text} -->"
    nil
  end
  
  # Output a tag with the specified name and attribute.
  # If there is a block, invoke it to output or return content.
  # Return nil.
  def tag(tagname, attributes={})
    # Output the tag name
	@out << "<#{tagname}"
	
	# Get the allowed attributes for this tag.
	allowed = self.class.allowed_attributes[tagname]
	
	# First, make sure that each of the attributes is allowed.
	# Assuming they are allowed, output all of the specified ones.
	attributes.each_pair do |key,value|
	  raise "unknown attribute: #{key}" unless allowed.include?(key)
	  @out << " #{key}='#{value}'"
	end
	
	# Now look through the allowed attributes, checking for
    # required attributes that were omitted and for attributes with
    # default values that we can output.

    allowed.each_pair do |key,value|
    # If this attribute was already output, do nothing.
      next if attributes.has_key? key
    
	  if (value == REQ)
        raise "required attribute '#{key}' missing in <#{tagname}>"
      elsif value.is_a? String
        @out << " #{key}='#{value}'"
      end
    end
	
	if block_given?
      # This block has content
      @out << '>' # End the opening tag
      content = yield # Invoke the block to output or return content
      if content # If any content returned
        @out << content.to_s # Output it as a string
      end
      @out << "</#{tagname}>" # Close the tag
    else
      # Otherwise, this is an empty tag, so just close it.
      @out << '/>'
	end
	nil # Tags output themselves, so they don't return any content.
  end
end
	
	
	









