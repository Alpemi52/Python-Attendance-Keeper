class Student:
    def __init__(self, data) -> None:
        self.Id = data[0]
        self.Name = data[1]
        self.Department = data[2]
        self.Section = data[3]
    def toString(self):
        splitted = str(self.Name).split(" ")
        lastname = splitted[-1]
        fs_name = "" 
        for name in splitted[0:-1]:
            fs_name+= name + " "

        return "{}, {}, {}".format(lastname, fs_name ,self.Id)