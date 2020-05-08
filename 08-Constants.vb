' Constant is a named memory location used to hold a value that CANNOT be changed during
'   the script execution. If a user tries to change a Constant value, the script execution
'   ends up with an error. Constants are declared the same way the variables are declared.

' Following are the rules for naming a constant.
    ' You must use a letter as the first character.
    ' You can't use a space, period (.), exclamation mark (!), or the characters @, &, $, # in the name.
    ' Name can't exceed 255 characters in length.
    ' You cannot use Visual Basic reserved keywords as variable name.

' In VBA, we need to assign a value to the declared Constants.
'   An error is thrown, if we try to change the value of the constant.    

' Syntax
' Const <<constant_name>> As <<constant_type>> = <<constant_value>>

' Example
' Let us create a button "Constant_demo" to demonstrate how to work with constants.

Private Sub Constant_demo_Click() 
   Const MyInteger As Integer = 42 
   Const myDate As Date = #2/2/2020# 
   Const myDay As String = "Sunday" 
   
   MsgBox "Integer is " & MyInteger & Chr(10) & "myDate is " 
      & myDate & Chr(10) & "myDay is " & myDay  
End Sub