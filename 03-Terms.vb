' You will acquaint yourself with the commonly used excel VBA terminologies.
' These terminologies will be used in further modules, hence understanding
'   each one of these is important.

' Dim
' Dim stands for Dimension and is a statement used to declare a variable name and
'   type that you want to create.

' Modules
' Modules is the area where the code is written. 
' This is also where any macros that you record are stored.
' To insert a Module, navigate to Insert → Module. Once a module is inserted 'module1' is created.
' Within the modules, we can write VBA code and the code is written within a Procedure.
' A Procedure/Sub Procedure is a series of VBA statements instructing what to do.


' Procedure
' Procedures are a group of statements executed as a whole, which instructs Excel how to perform a 
'   specific task. The task performed can be a very simple or a very complicated task.
'   However, it is a good practice to break down complicated procedures into smaller ones.
' The two main types of Procedures are Sub and Function.


' Function
' A function is a group of reusable code, which can be called anywhere in your program.
' This eliminates the need of writing the same code over and over again.
' This helps the programmers to divide a big program into a number of small and manageable functions.
' This can either be used by your macros to obtain a certain output or they can be used in
'   the Excel Formula Bar to perform calculations on your cell's values.
' Apart from inbuilt Functions, VBA allows to write user-defined functions as well and statements 
'   are written between Function and End Function.

Function Test() 
' Your Code
End Function

' Sub
' Sub-procedures work similar to functions. While sub procedures DO NOT Return a value,
'   functions may or may not return a value. Sub procedures CAN be called without call keyword.
'Sub procedures are always enclosed within Sub and End Sub statements.

Sub Test() 
' Your Code
End Sub

' Userforms
' Userforms are pop-up boxes that allow users to enter inputs or choose options.
' Microsoft uses these all the time in their applications.
' Some examples of these are Error Message Boxes, Dialog Boxes, and the Macro Recorder.
' The cool thing is that VBA gives you the ability to create your own custom Userforms! 
