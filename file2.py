class foo:

    def __init__(self):
        j=0

    def func(self):
        from file1 import ArithmeticOperations
        ArithmeticOperations = ArithmeticOperations()

        obj = self.file1.ArithmeticOperations()
        c = obj.get_sum(3, 8)
        print(c)


