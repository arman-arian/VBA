' Variable is a named memory location used to hold a value that can be changed during
'   the script execution. 

' Following are the basic rules for naming a variable.
    ' You must use a letter as the first character.
    ' You can't use a space, period (.), exclamation mark (!),
    '   or the characters @, &, $, # in the name.
    ' Name can't exceed 255 characters in length.
    ' You cannot use Visual Basic reserved keywords as variable name.

' In VBA, you need to declare the variables before using them.

' Syntax
' Dim <<variable_name>> As <<variable_type>>

' Data Types
' There are many VBA data types, which can be divided into two main categories,
'   namely numeric and non-numeric data types.

' Numeric Data Types
' Following table displays the numeric data types and the allowed range of values.
    ' Type	    Range of Values
    ' Byte	    0 to 255
    ' Integer	-32,768 to 32,767
    ' Long      -2,147,483,648 to 2,147,483,648
    ' Single	-3.402823E+38 to -1.401298E-45 for negative values
    '           1.401298E-45 to 3.402823E+38 for positive values.
    ' Double	-1.79769313486232e+308 to -4.94065645841247E-324 for negative values
    '           4.94065645841247E-324 to 1.79769313486232e+308 for positive values.
    ' Currency	-922,337,203,685,477.5808 to 922,337,203,685,477.5807
    ' Decimal	+/- 79,228,162,514,264,337,593,543,950,335 if no decimal is use
    '           +/- 7.9228162514264337593543950335 (28 decimal places).

' Non-Numeric Data Types
' Following table displays the non-numeric data types and the allowed range of values.
    ' Type	                    Range of Values
    ' String (fixed length)	    1 to 65,400 characters
    ' String (variable length)	0 to 2 billion characters
    ' Date	                    January 1, 100 to December 31, 9999
    ' Boolean	                True or False
    ' Object	                Any embedded object
    ' Variant (numeric)	        Any value as large as double
    ' Variant (text)	        Same as variable-length string

' Example
Private Sub say_helloworld_Click()
   Dim password As String
   password = "Admin#1"

   Dim num As Integer
   num = 1234

   Dim BirthDay As Date
   BirthDay = DateValue("30 / 10 / 2020")

   MsgBox "Passowrd is " & password & Chr(10) & "Value of num is " &
      num & Chr(10) & "Value of Birthday is " & BirthDay
End Sub

' Chr(10) - NewLine