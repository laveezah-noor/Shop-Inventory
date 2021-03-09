class Customer:

    def _init_(self):
        self.Customers={}


    def AddCustomer(self):
        self.name = input("Name: ")
        self.email = input("Email: ")
        self.telephone = input("Telephone: ")
        self.address = input("Address: ")
        print("New Customer is added")


    def DisplayCustomer(self):
        print("Customer Name :", self.name)
        print("Customer Email:", self.email)
        print("Customer Address:", self.address)
        print("Customer Telephone:" ,self.telephone)

a = Customer()
print(a.AddCustomer())
print(a.DisplayCustomer())

class Product:
    def _init_(self, productid, productprice, producttype):
        self.productid = productid
        self.productprice = productprice
        self.producttype = producttype

    def addProduct(self):
        pass

    def selectProduct(self, productid):
        self.productid = productid
        return"selected"

class Order(Product):
    def _init_(self, orderid, productid, customerid, customername, orderdate, amount, cart):
        self.orderid = orderid
        self.productid = productid
        self.customerid = customerid
        self.customername = customername
        self.orderdate = orderdate
        self.amount = amount

    def CreateOrder(self):
        pass
    def EditOrder(self):
        pass