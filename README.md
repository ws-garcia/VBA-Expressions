# ![VBA Expressions](/docs/assets/img/VBA%20Expressions.png)
[![GitHub](https://img.shields.io/github/license/ws-garcia/VBA-Expressions?style=plastic)](https://github.com/ws-garcia/VBA-Expressions/blob/master/LICENSE) [![GitHub release (latest by date)](https://img.shields.io/github/v/release/ws-garcia/VBA-Expressions?style=plastic)](https://github.com/ws-garcia/VBA-Expressions/releases/latest)

## Introductory words
VBA Expressions is a powerful string expression evaluator for VBA, which puts more than 60 mathematical, financial, date-time, logic and text manipulation functions at the user's fingertips. The `VBAexpressions.cls` class mediates almost all VBA functions as well as custom functions exposed through it. 

Although the main development goal of the class was the integration with [CSV Interface](https://github.com/ws-garcia/VBA-CSV-interface), VBA Expressions has evolved to become a support tool for students and teachers of science, accounting and engineering; this due to the added capability to solve systems of equations and non-linear equations in one variable.

## Advantages
* __Easy to use and integrate__.
* __Basic math operators__: `+` `-` `*` `/` `\` `^` `!`
* __Logical expressions__: `&` (AND), `|` (OR), `||` (XOR)
* __Binary relations__: `<`, `<=`, `<>`, `>=`, `=`, `>`, `$` (LIKE)
* __More than 60 built-in functions__: `Max`, `Sin`, `IRR`, `Switch`, `Iff`, `DateDiff`, `Solve`, `fZero`, `Format`...
* __Very flexible and powerful__: variables, constants and user-defined functions (UDFs) support.
* __Implied multiplication for variables, constants and functions__: `5avg(2;abs(-3-7tan(5));9)`, `5(x)` and `x(2)(3)` are valid expressions.
* __Evaluation of arrays of expressions given as text strings, as in Java__: curly brackets must be used to define arrays`{{...};{...}}`
* __Floating point notation input support__: `-5E-5`, `(1.434E3+1000)*2/3.235E-5` are valid inputs.
* __Free of external VBA dependencies__: does not use dll.

## Supported expressions

The evaluation approach used is similar to the one we humans use: divide the function into sub-expressions, create a symbolic string to build an expression evaluation flow, split the sub-expressions into chunks of operations (tokens) by tokenization, evaluate all the tokens. 

The module can evaluate mathematical expressions such as:

+ `5*avg(2;abs(-3-7*tan(5));9)-12*pi-e+(7/sin(30)-4!)*min(cos(30);cos(150))`
+ `min(cos(sin(30))+2^2;1)`
+ \* `GCD(1280;240;100;30*cos(0);10*DET({{sin(atn(1)*2); 0; 0}; {0; 2; 0}; {0; 0; 3}}))`

\*`GCD` is an user-defined function (UDF).

Allowed expressions must follow the following grammar:

```
Expression    =     ([{"("}]  SubExpr [{Operator [{"("}] SubExpr [{")"}]}] [{")"}] | {["("] ["{"] List [{";" List}] ["}"] [")"]}) 
SubExpr       =     Token [{Operator Token}]
Token         =     [{Unary}] Argument [(Operator | Function) ["("] [{Unary}] [Argument] [")"]]
Argument      =     (List | Variable | Operand)
List          =     "{" ["{"] SubExpr [{";" SubExpr}] ["}"] "}"
Unary         =     "-" | "+" | ~
Operand       =     ({Digit} ["."] [{Digit}] ["E"("-" | "+"){Digit}] | (True | False) | "'"Alphabet"'")
Variable      =     Alphabet [{Decimal}] [{(Digit | Alphabet)}]
Alphabet      =     "A-Z" | "a-z"
Decimal       =     "."
Digit         =     "0-9"
Operator      =     "+" | "-" | "*" | "/" | "\" | "^" | "%" | "!" | "<" | "<=" | "<>" | ">" | ">=" | "=" | "$" | "&" | "|" | "||"
Function      =     "abs" | "sin" | "cos" | "min" |...|[UDF]
```

## Operators precedence
VBA expressions uses the following precedence rules to evaluate mathematical expressions:

1. `()`               Grouping: evaluates functions arguments as well.
2. `! - +`            Unary operators: exponentiation is the only operation that violates this. Ex.: `-2 ^ 2 = -4 | (-2) ^ 2 = 4`.
3. `^`                Exponentiation: Although Excel and Matlab evaluate nested exponentiations from left to right, Google, mathematicians and several modern programming languages, such as Perl, Python and Ruby, evaluate this operation from right to left. VBA expressions also evals in Python way: a^b^c = a^(b^c).
4. `\* / % `          Multiplication, division, modulo: from left to right.
5. `+ -`              Addition and subtraction: from left to right.
6. `< <= <> >= = > $` Binary relations.
7. `~`                Logical negation.
8. `&`                Logical AND.
9. `||`               Logical XOR.
10. `|`               Logical OR.

## Variables
Users can enter variables and set/assign their values for the calculations. Variable names must meet the following requirements:
1. Start with a letter.
2. End in a letter or number. `"x.1"`, `"number1"`, `"value.a"` are valid variable names.
3. A variable named `"A"` is distinct from another variable named `"a"`, since variables are case-sensitive. This rule is broken by the constant `PI`, since `PI=Pi=pi=pI`.
4. The token `"E"` cannot be used as variable due this token is reserved for floating point computation. For example, the expression `"2.5pi+3.5e"` will be evaluated to `~17.3679680`, but a expression like `"2.5pi+3.5E"` will return an error.

## User-defined functions (UDF)
Users can register custom modules to expose and use their functions through the VBAcallBack.cls module. All UDFs must have a single Variant argument that will receive an one-dimensional array of strings (one element for each function argument).

Here is a working example of UDF function creation
```
Sub AddingNewFunctions()
    Dim Evaluator As VBAexpressions
    Dim UDFnames() As Variant
    Dim Result As String
    
    Set Evaluator = New VBAexpressions
    UDFnames() = Array("GCD")
    With Evaluator
        .DeclareUDF UDFnames, "UserDefFunctions"                        'Declare the Greatest Common Divisor function
                                                                        'defined in the UDfunctions class module. This need
                                                                        'an instance in the VBAcallBack class module.
        ' The determinant of a diagonal matrix. It is defined
        ' as the product of the elements in its diagonal.
        ' For our case: 1*2*3=6. (Note that sin(atn(1)*2)=sin(pi/2)=1)
        .Create "GCD(1280;240;100;30*cos(0);10*DET({{sin(atn(1)*2); 0; 0}; {0; 2; 0}; {0; 0; 3}}))"
        Result = .Eval
    End With
End Sub
```
## Working with arrays
VBA expressions can evaluate matrix functions whose arguments are given as arrays/vectors, using a syntax like [Java](https://www.w3schools.com/java/java_arrays.asp). The following expression will calculate, and format to percentage, the internal rate of return (`IRR`) of a cash flow described using a one dimensional array with 5 entries:

`FORMAT(IRR({{-70000;12000;15000;18000;21000}});'Percent')`

However, user-defined array functions need to take care of creating arrays from a string, the `ArrayFromString` method can be used for this purpose.

## Using the code
VBA Expressions is an easy-to-use library, this section shows some examples of how to use the most common properties and methods

```
Sub SimpleMathEval()
    Dim Evaluator As VBAexpressions
    Set Evaluator = New VBAexpressions
    With Evaluator
        .Create "(((((((((((-123.456-654.321)*1.1)*2.2)*3.3)+4.4)+5.5)+6.6)*7.7)*8.8)+9.9)+10.10)"
        If .ReadyToEval Then    'Evaluates only if the expression was successfully parsed.
            .Eval
        End If
    End With
End Sub
Sub LateVariableAssignment()
    Dim Evaluator As VBAexpressions
    Set Evaluator = New VBAexpressions
    With Evaluator
        .Create "Pi.e * 5.2Pie.1 + 3.1Pie"
        If .ReadyToEval Then
            Debug.Print "Variables: "; .CurrentVariables    'Print the list of parsed variables
            .Eval ("Pi.e=1; Pie.1=2; Pie=3")                'Late variable assignment
            Debug.Print .Expression; " = "; .Result; _
                        "; for: "; .CurrentVarValues        'Print stored result, expression and values used in evaluation
        End If
    End With
End Sub
Sub EarlyVariableAssignment()
    Dim Evaluator As VBAexpressions
    Set Evaluator = New VBAexpressions
    With Evaluator
        .Create "Pi.e * 5.2Pie.1 + 3.1Pie"
        If .ReadyToEval Then
            Debug.Print "Variables: "; .CurrentVariables
            .VarValue("Pi.e") = 1
            .ImplicitVarValue("Pie.1") = "2*Pi.e"
            .ImplicitVarValue("Pie") = "Pie.1/3"
            .Eval
            Debug.Print .Expression; " = "; .Result; _
                        "; for: "; .CurrentVarValues
        End If
    End With
End Sub
Sub TrigFunctions()
    Dim Evaluator As VBAexpressions
    Set Evaluator = New VBAexpressions
    With Evaluator
        .Create "asin(sin(30))"
        If .ReadyToEval Then
            .Degrees = True               'Eval in degrees
            .Eval
        End If
    End With
End Sub
Sub StringFunctions()
    Dim Evaluator As VBAexpressions
    Set Evaluator = New VBAexpressions
    
    With Evaluator
        .Create "CONCAT(CHOOSE(1;x;'2nd';'3th';'4th';'5th');'Element';'selected';'/')"
        .Eval ("x='1st'")
    End With
End Sub
Sub LogicalFunctions()
    Dim Evaluator As VBAexpressions
    Set Evaluator = New VBAexpressions
    
    With Evaluator
        .Create "IFF(x > y & x > 0; x; y)"                   
        .Eval("x=70;y=15")                 'This will be evaluated to 70
    End With
End Sub
Sub SolveSystemOfLinearEquations()
    Dim Evaluator As VBAexpressions
    
    Set Evaluator = New VBAexpressions
    With Evaluator
        'Create an array of vectors a, b, c, d and	e
        .Create "SOLVE(ARRAY(a;b;c;d;e);{{'x1';'x2';'x3';'x4';'x5'}};{{100;100;100;100;100}};True)"
		  
        'Define coeficient vectors
        .VarValue("a") = "{4;-1;0;1;0}"
        .VarValue("b") = "{-1;4;-1;0;1}"
        .VarValue("c") = "{0;-1;4;-1;0}"
        .VarValue("d") = "{1;0;-1;4;-1}"
        .VarValue("e") = "{0;1;0;-1;4}"
		  .Eval										'Evaluates to [x1 = 25; x2 = 35.7142857143; x3 = 42.8571428571; x4 = 35.7142857143; x5 = 25]
    End With
End Sub

'@------------------------------------------------------
' Here a list of the new functions and its results
'***********************************ADVANCED MATH FUNCTIONS***************************************************************
''' FZERO('2x^2+x-12';-2;3) : x = 2.21221445045296       (find a zero for the given function in the interval -2<=x<=3)
'''
''' a = {1;0;4}; b = {1;1;6}; c = {-3;0;-10}; d = {2;3;4}
''' 			LUDECOMP(ARRAY(a;b;c)) : {{-3;0;-10};{-0.333333333333333;1;2.66666666666667};{-0.333333333333333;0;0.666666666666667}}  
''' 			LUSOLVE(ARRAY(a;b;c);{{'x';'y';'z'}};{{2;3;4}};True) : x = -18; y = -9; z = 5 
''' 			MMULT(INVERSE(ARRAY(a;b;c));ARRAY(d)) : {{-18};{-9};{5}} 
''' 			MMULT(ARRAY(a;b;c);INVERSE(ARRAY(a;b;c))) : {{1;0;0};{0;1;0};{0;0;1}}
'***********************************FINANTIAL FUNCTIONS***************************************************************
''' FORMAT(SYD(10000;5000;5;2);'Currency') = '$1,333.33'
''' FORMAT(SLN(10000;0;5);'Currency') = '$2,000.00'
''' FORMAT(RATE(2*12; -250; 5000; 0; 1);'Percent') = '1.66%'
''' FORMAT(PV(0.075/12; 2*12; 250; 0; 0);'Currency') = '($5,555.61)'
''' FORMAT(PPMT(0.06/52; 20; 4*52; 8000; 0; 0);'Currency') = '($34.81)'
''' FORMAT(PMT(0.075/12; 2*12; 5000; 0; 1);'Currency') = '($223.60)'
''' FORMAT(NPV(0.1;{{-10000;3000;4200;6800}});'Currency') = '$1,188.44'
''' FORMAT(NPER(0.0525/1; -200; 1500);'0.00') = '9.78'
''' FORMAT(MIRR({{-7500;3000;5000;1200;4000}};0.05;0.08);'Percent') = '18.74%'
'''
''' Microsoft example. VBA Expressions use a custom solver to compute IRR
''' FORMAT(IRR({{-70000;12000;15000;18000;21000}});'Percent') = '-2.12%'
''' FORMAT(IRR({{-70000;12000;15000;18000;21000;26000}});'Percent') = '8.66%'
''' FORMAT(IRR({{-70000;12000;15000}};true);'Percent') = '-44.35%'					'Find a solution with negative values for IRR
'''
''' FORMAT(IPMT(0.0525/1; 4; 10*1; 6500);'Currency') = '($256.50)'
''' FORMAT(FV(0.0525/1; 10*1; -100; -6500; 0);'Currency') = '$12,115.19'
''' FORMAT(DDB(10000; 5000; 5; 2);'Currency') = '$1,000.00'
'***********************************DATE AND TIME FUNCTIONS***************************************************************
''' YEAR(NOW()) = 2022
''' WEEKDAYNAME(1;true;2) = 'lun.'
''' WEEKDAY(NOW()) = 2
''' TIMEVALUE(NOW()) = '7:53:00 a. m.'
''' TIMESERIAL(x;y;z) = '6:45:50 a. m.' for x = 7; y = -15; z = 50
''' MONTH(NOW()) = 10
''' MONTHNAME(x;y) = 'marzo' for x = 3; y = false
''' MONTH(x) = 10 for x = '10/10/2022 7:53:10 a. m.'
''' MINUTE(x) = 53 for x = '10/10/2022 7:53:12 a. m.'
''' HOUR(x) = 7 for x = '10/10/2022 7:53:15 a. m.'
''' DAY(DATE()) = 10
''' DATEVALUE(DATE()) = '10/10/2022'
''' DATESERIAL(2022;x+2;3y) = '21/12/2022' for x = 10; y = 7
''' DATEPART(x;DATE()) = 2022 for x = 'yyyy'
''' DATEPART(x;DATE()) = 10 for x = 'm'
''' DATEPART(x;DATE()) = 10 for x = 'd'
''' DATEPART(x;DATE()) = 4 for x = 'q'
''' DATEDIFF(x;DATE();DATEADD(x;y;DATE())) = 3 for x = 'yyyy'; y = 3
''' DATEDIFF(x;DATE();DATEADD(x;y;DATE())) = 2 for x = 'q'; y = 2
''' DATEDIFF(x;DATE();DATEADD(x;y;DATE())) = 5 for x = 'm'; y = 5
''' DATEDIFF(x;DATE();DATEADD(x;y;DATE())) = 10 for x = 'd'; y = 10
''' DATEADD(x;y;DATE()) = 10/10/2025 for x = 'yyyy'; y = 3
''' DATEADD(x;y;DATE()) = 10/4/2023 for x = 'q'; y = 2
''' DATEADD(x;y;DATE()) = 10/3/2023 for x = 'm'; y = 5
''' DATEADD(x;y;DATE()) = 20/10/2022 for x = 'd'; y = 10
''' DATE() = '10/10/2022'
'***********************************STRING FUNCTIONS***************************************************************
''' UCASE(x) = ' THIS STRING ' for x = ' This String '
''' TRIM(x) = 'Capi tal' for x = '  Capi tal '
''' RIGHT(2x+20-5+x;2) = '90' for x = 25
''' RIGHT(2x+20-5+x;2) = '09' for x = 98
''' REPLACE(x;'a';'A';1;2) = 'CApitAl' for x = 'Capital'
''' MID(x;4;3) = 'ion' for x = 'Region'
''' LEN(x) = 6 for x = 'Region'
''' LEFT(2x+20-5+x;2) = '90' for x = 25
''' LEFT(2x+20-5+x;2) = '30' for x = 98
''' LCASE(x) = 'this string' for x = 'This String'
''' LCASE(x) = '98' for x = 98
''' FORMAT(x;'Percent') = '98.10%' for x = 0.981
''' CHR(x) = ';' for x = 59
''' ASC(x) = 82 for x = 'Region'
''' SWITCH(x='Asia';1;x='Africa';2;x='Oceania';3) = 1 for x = 'Asia'
'@------------------------------------------------------
```
## Tested 
[![Rubberduck](https://user-images.githubusercontent.com/5751684/48656196-a507af80-e9ef-11e8-9c09-1ce3c619c019.png)](https://github.com/rubberduck-vba/Rubberduck/) 
## Licence

Copyright (C) 2022  [W. García](https://github.com/ws-garcia/).

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program.  If not, see <https://www.gnu.org/licenses/gpl-3.0.html>.

