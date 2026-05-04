class Student:
    def __init__(self, name, email, phone, course, date, paid=False):
        self.name = name
        self.email = email
        self.phone = phone
        self.course = course
        self.date = date
        self.paid = paid

    def __str__(self):
        return (
            f"Student(name={self.name}, "
            f"email={self.email}, "
            f"phone={self.phone}, "
            f"course={self.course}, "
            f"date={self.date}, "
            f"paid={self.paid})"
        )