import random

# variables
print("\n__--== VARIABLES ==--__")
x = 5
Y = 7.26
A = 1j
X = "Radu"
a, b, c = ["Ion", 8.44, 5], (1, 2, 3), range(7)
z = Z = "banana"
y = False
B = {"name": "Radu", "age": 25}
C = {1, 2, 3}
print(x, X, y, Y)
print(A, B, C)
print(a, b, c)
print(z)

# variable types
print("\n__--== VARIABLE TYPES ==--__")
print(type(x))
print(type(Y))
print(type(A))
print(type(a))
print(type(b))
print(type(c))
print(type(X))
print(type(y))
print(type(B))

# function & global variable
print("\n__--== FUNCTION & GLOBAL VARIABLE ==--__")
print("x: " + str(x))


def my_function():
    global x
    x = "Radu"
    print("Hi, " + x)


my_function()
print("x: " + str(x))

# random number
print("\n__--== RANDOM NUMBER ==--__")
random_number = random.randrange(1, 8)
print("Random in range 1-8: " + str(random_number))
random_number = random.random()
print("Random number: " + str(random_number))

# for loop
print("\n__--== FOR LOOP ==--__")
for w in a:
    print(w)

# string
print("\n__--== STRING ==--__")
print("Radu has " + str(len("Radu")) + " letters")
j = "Longer sentence, for TESTING purposes."
if "sent" in j:
    print("\"sent\" is in \"" + j + "\"")
if "Radu" not in j:
    print("\"Radu\" is not in \"" + j + "\"")
print("0123456789ABCDEF")
print(j)
print(j[2:8])
print(j[:8])
print(j[5:])
print(j[-8:-2:2])
print(j.upper())
print(j.lower())
print(j.casefold())
print(j.replace("TES", "BA"))
print(j.strip())
print(j.split())

# boolean
print("\n__--== BOOLEAN ==--__")
print(1 > 5)

full_name = input("What's your name? ")
age = int(input("What's your age? "))
height = float(input("What's your height? "))

if age > 60:
    print("Hi {0}! You are {1} years old and {2}cm tall. You can retire!".format(full_name, age, height))
else:
    print("Hi {0}! You are {1} years old and {2}cm tall. You can not retire yet!".format(full_name, age, height))
