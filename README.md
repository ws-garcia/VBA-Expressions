# ![VBA Expressions](/docs/assets/img/VBA%20Expressions.png)
[![GitHub](https://img.shields.io/github/license/ws-garcia/VBA-Expressions?style=plastic)](https://github.com/ws-garcia/VBA-Expressions/blob/master/LICENSE) [![GitHub release (latest by date)](https://img.shields.io/github/v/release/ws-garcia/VBA-Expressions?style=plastic)](https://github.com/ws-garcia/VBA-Expressions/releases/latest)

## Introductory words
VBA Expressions is a powerful mathematical expressions evaluator for VBA strings. The `VBAexpressions.cls` class serves as an intermediary between user interfaces and the main VBA/custom functions exposed through it. The main development goal of the class is to integrate it with [CSV Interface](https://github.com/ws-garcia/VBA-CSV-interface), with as minimal programming effort as possible, and to allow users to perform complex queries from CSV files using built-in and custom functions.

## Advantages
* __Easy to use and integrate__.
* __Basic math operators__: `+` `-` `*` `/` `\` `^` `!`
* __Logical expressions__: `& (AND)` `| (OR)` `|| (XOR)`
* __Binary relations__: `< <= <> >= = >`
* __More than 20 built-in functions__: `Max`, `Min`, `Avg`, `Sin`, `Ceil`, `Floor`...
* __Very flexible__: variables, constants and user-defined functions (UDFs) support.
* __Implied multiplication for variables, constants and functions__: `5avg(2;abs(-3-7tan(5));9)` is valid expression; `5(2)` is not.
* __Evaluation of arrays of expressions given as text strings, as in Java__: curly brackets must be used to define arrays`{{...};{...}}`
* __Floating point notation input support__: `-5E-5`, `(1.434E3+1000)*2/3.235E-5` are valid inputs.
* __Free of external VBA dependencies__: does not use dll.

## Supported expressions

The evaluation approach used is similar to the one we humans use: divide the function into sub-expressions, create a symbolic string to build an expression evaluation flow, split the sub-expressions into chunks of operations (tokens) by tokenization, evaluate all the tokens. 

The module can evaluate mathematical expressions such as:

+ `5*avg(2;abs(-3-7*tan(5));9)-12*pi-e+(7/sin(30)-4!)*min(cos(30);cos(150))`
+ `min(cos(sin(30))+2^2;1)`
+ \* `GCD(1280;240;100;30*cos(0);10*DET({{sin(atn(1)*2); 0; 0}; {0; 2; 0}; {0; 0; 3}}))`

\*`GCD` and `DET` are user-defined functions (UDF).

Allowed expressions must follow the following grammar:

```
Expression    =     ([{"("}]  SubExpr [{Operator [{"("}] SubExpr [{")"}]}] [{")"}] | {["("] ["{"] List [{";" List}] ["}"] [")"]}
SubExpr       =     Token [{Operator Token}]
Token         =     [{Unary}] Argument [(Operator | Function) ["("] [{Unary}] [Argument] [")"]]
Argument      =     (List | Variable | Operand)
List          =     ["{"] SubExpr [{";" SubExpr}] ["}"]
Unary         =     "-" | "+" | ~
Operand       =     ({Digit} ["."] [{Digit}] ["E"("-" | "+"){Digit}] | (True | False))
Variable      =     Alphabet [{Decimal}] [{(Digit | Alphabet)}]
Alphabet      =     "A-Z" | "a-z"
Decimal       =     "."
Digit         =     "0-9"
Operator      =     "+" | "-" | "*" | "/" | "\" | "^" | "%" | "!" | "<" | "<=" | "<>" | ">" | ">=" | "=" | "&" | "|" | "||"
Function      =     "abs" | "sin" | "cos" | "min" |...|[UDF]
```

## Operators precedence
VBA expressions uses the following precedence rules to evaluate mathematical expressions:

1. `()`             Grouping: evaluates functions arguments as well.
2. `! - +`          Unary operators: exponentiation is the only operation that violates this. Ex.: `-2 ^ 2 = -4 | (-2) ^ 2 = 4`.
3. `^`              Exponentiation: Although Excel and Matlab evaluate nested exponentiations from left to right, Google, mathematicians and several modern programming languages, such as Perl, Python and Ruby, evaluate this operation from right to left. VBA expressions also evals in Python way: a^b^c = a^(b^c).
4. `\* / % `        Multiplication, division, modulo: from left to right.
5. `+ -`            Addition and subtraction: from left to right.
6. `< <= <> >= = >` Binary relations.
7. `~`              Logical negation.
8. `&`              Logical AND.
9. `||`             Logical XOR.
10. `|`             Logical OR.

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
    UDFnames() = Array("GCD", "DET")
    With Evaluator
        .DeclareUDF UDFnames, "UserDefFunctions"                        'Declare the Greatest Common Divisor and Determinant functions
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
VBA expressions can evaluate matrix functions whose arguments are given as arrays/vectors, using a syntax like [Java](https://www.w3schools.com/java/java_arrays.asp). The following expression will calculate the determinant (`DET`) of a matrix composed of 3 vectors with 3 elements each:

`DET({{sin(atn(1)*2); 0; 0}; {0; 2; 0}; {0; 0; 3}})`

If the user needs to evaluate a function that accepts more than one argument, including more than one array, all arrays arguments must be passed surrounded by parentheses "({...})". For example, a function call that emulates the SQL IN statement using an array argument and a reference value can be written as follows.

`IN_(({{sin(atn(1)*2); 2; 3; 4; 5}});1)`

The above will pass this array of strings to the `IN_` function:

`[{{1;2;3;4;5}}] [1]`

However, matrix functions need to take care of creating arrays from a string, the ArrayFromString method can be used for this purpose.

As an illustration, the `UDFunctions.cls` module has an implementation of the `DET` function with an example of using the array handle function. In addition, the `GCD` function is implemented as a demo.

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
            .VarValue("Pie.1") = 2
            .VarValue("Pie") = 3
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
```

## Licence

Copyright (C) 2022  [W. García](https://github.com/ws-garcia/).

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program.  If not, see <https://www.gnu.org/licenses/gpl-3.0.html>.

