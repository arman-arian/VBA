' The InputBox function prompts the users to enter values. After entering the values,
'   if the user clicks the OK button or presses ENTER on the keyboard, the InputBox function
'   will return the text in the text box. If the user clicks the Cancel button, the function
'   will return an empty string ("").

' Syntax
' InputBox(prompt[,title][,default][,xpos][,ypos][,helpfile,context])

' Parameter Description
' Prompt − A required parameter. A String that is displayed as a message in the dialog box.
'   The maximum length of prompt is approximately 1024 characters. If the message extends to more
'   than a line, then the lines can be separated using a carriage return character (Chr(13)) or
'   a linefeed character (Chr(10)) between each line.

' Title − An optional parameter. A String expression displayed in the title bar of the dialog box.
'   If the title is left blank, the application name is placed in the title bar.

' Default − An optional parameter. A default text in the text box that the user would like
'   to be displayed.

' XPos − An optional parameter. The position of X axis represents the prompt distance from
'   the left side of the screen horizontally. If left blank, the input box is horizontally centered.

' YPos − An optional parameter. The position of Y axis represents the prompt distance from
'   the left side of the screen vertically. If left blank, the input box is vertically centered.

' Helpfile − An optional parameter. A String expression that identifies the helpfile to be used to
'   provide context-sensitive Help for the dialog box.

' context − An optional parameter. A Numeric expression that identifies the Help context number
'   assigned by the Help author to the appropriate Help topic. If context is provided,
'   helpfile must also be provided.

' Example
' Let us calculate the area of a rectangle by getting values from the user at run time
'   with the help of two input boxes (one for length and one for width).

Sub Button1_Click()
   MsgBox findArea()
End Sub

Function findArea() 
   Dim Length As Double 
   Dim Width As Double 
   
   Length = InputBox("Enter Length ", "Enter a Number") 
   Width = InputBox("Enter Width", "Enter a Number") 
   findArea = Length * Width 
End Function