import win32com.server.util

debugging = 1

if debugging:
    from win32com.server.dispatcher import DefaultDebugDispatcher
    my_dispatcher = DefaultDebugDispatcher
else:
    my_dispatcher = None

# Define a class that implements the custom COM interface
class Person:
    _public_methods_ = []
    _public_attrs_ = ["name", "age"]

    def __init__(self, name, age):
        self.name = name
        self.age = age


# Define a Python class that will be registered as a COM object
class MyCOMObject:
    _public_methods_ = ["hello", "add", "get_person", "substract"]
    _reg_clsid_ = '{44ee76c7-1290-4ea6-8189-00d5d7cd712a}'
    _reg_desc_ = "My Python COM server"
    _reg_progid_ = "Python.MyCOMObject"

    def hello(self, name):
        return f"Hello, {name}!"

    def add(self, a, b):
        return a + b

    def get_person(self):
        person = Person("wxinix", 35)
        wrapped = win32com.server.util.wrap(person, useDispatcher = my_dispatcher)
        return wrapped

    def substract(self, a, b):
        return a - b

# Register the class as a COM object
if __name__=='__main__':
    import win32com.server.register
    win32com.server.register.UseCommandLine(MyCOMObject, debug = debugging)