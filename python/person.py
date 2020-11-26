

class Person:
    # constructor
    def __init__(self, name, age, height, weight):
        self.name = name
        self.age = age
        self.height = height
        self.weight = weight


John = Person("John", 20, 178, 69)
Smith = Person("Smith", 18, 168, 60)


print(John.name, "の属性")
print("年齢：", John.age)
print("身長：", John.height)
print("体重：", John.weight)
print()
print(Smith.name, "の属性")
print("年齢：", Smith.age)
print("身長：", Smith.height)
print("体重：", Smith.weight)