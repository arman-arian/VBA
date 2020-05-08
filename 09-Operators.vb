' An Operator can be defined using a simple expression - 4 + 5 is equal to 9.
' Here, 4 and 5 are called operands and + is called operator.
' VBA supports following types of operators −
    ' Arithmetic Operators
    ' Comparison Operators
    ' Logical (or Relational) Operators
    ' Concatenation Operators

' The Arithmatic Operators
' Following arithmetic operators are supported by VBA.
' Assume variable A holds 5 and variable B holds 10, then −

'   Operator	    Description	                                               Example
'   --------------------------------------------------------------------------------------------------
'   +	            Adds the two operands	                                   A + B will give 15
'   -	            Subtracts the second operand from the first	               A - B will give -5
'   *	            Multiplies both the operands	                           A * B will give 50
'   /	            Divides the numerator by the denominator	               B / A will give 2
'   %	            Modulus operator and the remainder after                   B % A will give 0
'                        an integer division	                               
'   ^	            Exponentiation operator	                                   B ^ A will give 100000


' The Comparison Operators
' There are following comparison operators supported by VBA.
' Assume variable A holds 10 and variable B holds 20, then −

'   Operator	    Description	                                              Example
'   --------------------------------------------------------------------------------------------------
'   =	            Checks if the value of the two operands                   (A = B) is False.
'                    are equal or not. If yes, then the condition is true.	

'   <>	            Checks if the value of the two operands                   (A <> B) is True.
'                    are equal or not. If the values are not equal,
'                    then the condition is true.

'   >	            Checks if the value of the left operand                   (A > B) is False.
'                    is greater than the value of the right operand
'                     If yes, then the condition is true.

'   <	            Checks if the value of the left operand                   (A < B) is True.
'                    is less than the value of the right operand.
'                    If yes, then the condition is true.	

'   >=	            Checks if the value of the left operand	                  (A >= B) is False.
'                    is greater than or equal to the value of
'                    the right operand. If yes, then the condition is true.

'   <=	            Checks if the value of the left operand is less           (A <= B) is True.
'                    than or equal to the value of the right operand.
'                    If yes, then the condition is true.


' The Logical Operators
' Following logical operators are supported by VBA.
' Assume variable A holds 10 and variable B holds 0, then −

'   Operator    	Description	                                        Example
'   --------------------------------------------------------------------------------------------------
'   AND	            Called Logical AND operator.                        a<>0 AND b<>0 is False.
'                   If both the conditions are True,
'                    then the Expression is true.

'   OR	            Called Logical OR Operator.                         a<>0 OR b<>0 is true.
'                    If any of the two conditions are True,
'                    then the condition is true.

'   NOT	            Called Logical NOT Operator. Used to reverse        NOT(a<>0 OR b<>0) is false.
'                    the logical state of its operand. If a condition
'                    is true, then Logical NOT operator will make false.	

'   XOR	            Called Logical Exclusion. It is the combination     (a<>0 XOR b<>0) is true.
'                    of NOT and OR Operator. If one, and only one,
'                    of the expressions evaluates to be True, the result is True.	


' The Concatenation Operators
' Following Concatenation operators are supported by VBA.
' Assume variable A holds 5 and variable B holds 10 then −

'   Operator	        Description                             	        Example
'   --------------------------------------------------------------------------------------------------
'   +	                Adds two Values as Variable. Values are Numeric	    A + B will give 15
'   &	                Concatenates two Values	                            A & B will give 510

' Assume variable A = "Microsoft" and variable B = "VBScript", then −

' Operator	            Description	                           Example
'   --------------------------------------------------------------------------------------------------
'   +	                Concatenates two Values	               A + B will give MicrosoftVBScript
'   &	                Concatenates two Values	               A & B will give MicrosoftVBScript

' Note − Concatenation Operators can be used for both numbers and strings.
' The output depends on the context, if the variables hold numeric value or string value.