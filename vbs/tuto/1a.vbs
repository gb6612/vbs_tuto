' Nombres

number1 = 15
number2 = 6
WScript.StdOut.WriteLine number1 & " op " & number2
WScript.StdOut.WriteLine "* : " & (number1 * number2)
WScript.StdOut.WriteLine "/ : " & (number1 / number2)  
'WScript.StdOut.WriteLine "div : " & Int(number1 div number2)  
WScript.StdOut.WriteLine "mod(reste) : " & (number1 mod number2)  
WScript.StdOut.WriteLine "^(exposant) : " & (number1 ^ number2)  

MyNumber = Abs(50.3)    ' Returns 50.3.
MyNumber = Abs(-50.3)   ' Returns 50.3.


MyNumber = Int(99.8)    ' Returns 99.
MyNumber = Fix(99.2)    ' Returns 99.

MyNumber = Int(-99.8)    ' Returns -100.
MyNumber = Fix(-99.8)    ' Returns -99.

MyNumber = Int(-99.2)    ' Returns -100.
MyNumber = Fix(-99.2)    ' Returns -99.


MySign = Sgn(12)   ' Returns 1.
MySign = Sgn(-2.4) ' Returns -1.
MySign = Sgn(0)    ' Returns 0.


