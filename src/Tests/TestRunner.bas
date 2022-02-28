Attribute VB_Name = "TestRunner"
Option Explicit
Option Private Module
Private Evaluator As VBAexpressions
Private expected As String
Private actual As String

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

Private Function GetResult(Expression As String _
                        , Optional VariablesValues As String = vbNullString) As String
    On Error Resume Next
    Set Evaluator = New VBAexpressions
    
    With Evaluator
        .Create Expression
        GetResult = .Eval(VariablesValues)
    End With
End Function

'@TestMethod("VBA Expressions")
Private Sub Parentheses()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "(((((((((((-123.456-654.321)*1.1)*2.2)*3.3)+4.4)+5.5)+6.6)*7.7)*8.8)+9.9)+10.10)" _
                        )
    expected = "-419741.48578672"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("VBA Expressions")
Private Sub ParenthesesAndSingleFunction()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "(1+(2-5)*3+8/(5+3)^2)/sqr(4^2+3^2)" _
                        )
    expected = "-1.575"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("VBA Expressions")
Private Sub FunctionsWithMoreThanOneArgument()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "min(5;6;max(-0.6;-3))" _
                        )
    expected = "-0.6"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("VBA Expressions")
Private Sub NestedFunctions()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "tan(sqr(abs(ln(x))))" _
                        , "x = " & Exp(1) _
                        )
    expected = "1.5574077246549"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("VBA Expressions")
Private Sub FloatingPointArithmetic()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "(1.434E3+1000)*2/3.235E-5" _
                        )
    expected = "150479134.46677"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("VBA Expressions")
Private Sub ExponentiationPrecedence()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "4^3^2" _
                        )
    expected = "262144"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("VBA Expressions")
Private Sub Factorials()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "25!/(24!)" _
                        )
    expected = "25"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("VBA Expressions")
Private Sub Precedence()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "5avg(2;abs(-3-7tan(5));9)-12pi-e+(7/sin(30)-4!)*min(cos(30);cos(150))" _
                        )
    expected = "7.56040693890688"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("VBA Expressions")
Private Sub Variables()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "Pi.e * 5.2Pie.1 + 3.1Pie" _
                        , "Pi.e = 1; Pie.1 = 2; Pie = 3" _
                        )
    expected = "19.7"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("VBA Expressions")
Private Sub UDFsAndArrays()
    On Error GoTo TestFail
    '///////////////////////////////////////////////////////////////////////////////////
    ' For illustrative purposes only. These UDFs are already implemented.
    '
    ' Dim UDFnames() As Variant
    ' UDFnames() = Array("GCD", "DET")
    '
    ' Evaluator.DeclareUDF UDFnames, "UserDefFunctions"    'Declaring the UDFs. This need
                                                           'an instance in the VBAcallBack
                                                           'class module.
    '
    '///////////////////////////////////////////////////////////////////////////////////
    actual = GetResult( _
                        "GCD(1280;240;100;30*cos(0);10*DET({{sin(atn(1)*2); 0; 0}; {0; 2; 0}; {0; 0; 3}}))" _
                        )
    expected = "10"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("VBA Expressions")
Private Sub LogicalOperatorsNumericOutput()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "(x<=0)* x^2 + (x>0 & x<=1)* Ln(x+1) + (x>1)* Sqr(x-Ln(2))" _
                        , "x = 6" _
                        )
    expected = "2.30366074313039"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("VBA Expressions")
Private Sub TestLogicalOperatorsBooleanOutput()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "x>0 & Sqr(x-Ln(2))>=3 | tan(x)<0" _
                        , "x = 6" _
                        )
    expected = "True"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("VBA Expressions")
Private Sub TestTrigFunctions()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "tan(pi/4)^3-((3*sin(pi/4)-sin(3*pi/4))/(3*cos(pi/4)+cos(3*pi/4)))" _
                        )
    expected = "0"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("VBA Expressions")
Private Sub TestModFunction()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "- (-1) + (+1) + 1.000 / 1.000 + 1 * (1) * (0.2) * (5) * (-1) * (--1) + 4 % 5 % 45   % 1 " _
                        )
    expected = "2"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("VBA Expressions")
Private Sub testStringComp()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "Region = 'Central America'" _
                        , "Region = 'Asia'" _
                        )
    expected = "False"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

