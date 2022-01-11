class Person:
    def __init__(self, first_name, last_name):
        self.first_name = first_name
        self.last_name = last_name

    def print_name(self):
        print("\nHi! My name is {} {}.".format(self.first_name, self.last_name))


class Student(Person):
    def __init__(self, first_name, last_name, grade):
        Person.__init__(self, first_name, last_name)
        self.grade = grade

    def print_name(self):
        super().print_name()
        print("My grade is {}".format(self.grade))


p1 = Person("Radu", "Ionescu")
p1.print_name()

s1 = Student("Gigi", "Popescu", 8.38)
s1.print_name()
