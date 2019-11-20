''' in a module

Sub abc()

    Dim order As New salesOrder
    
    order.number = "abcde012345"
    order.customer.name = "companyXyz"
    order.customer.number = "501001"
    
    order.PrintName
    
End Sub

''' in a class module named 'salesOrder'

'properties
Public number As String
Public customer As New salesOrderCustomer 'this class contains a variable which will store an object from another class

'method
Public Sub PrintName()
    Debug.Print number
    Debug.Print customer.name
    Debug.Print customer.number
End Sub

''' in a class module named 'salesOrderCustomer'

'properties
Public name As String
Public number As String

''' this is how I would do the above in javascript

function salesOrder(number, customerName, customerNumber) {
	this.number = number;
  this.customer = {
  	name: customerName,
    number: customerNumber
  }
}
salesOrder.prototype.display = function() {
	console.log(this.number);
  console.log(this.customer.name);
  console.log(this.customer.number);
}
var order = new salesOrder("abcde012345","companyXyz","501001")
order.display();

