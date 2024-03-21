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

Public Sub ShowEvalForm(control As IRibbonControl)
    'EvalForm_frm.Show False
End Sub

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

Private Function GetResult(expression As String _
                        , Optional VariablesValues As String = vbNullString) As String
    On Error Resume Next
    Set Evaluator = New VBAexpressions
    
    With Evaluator
        .Create expression
        GetResult = .Eval(VariablesValues)
    End With
End Function

'@TestMethod("Parentheses")
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
'@TestMethod("Parentheses and Single Function")
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
'@TestMethod("Functions with More than One Argument")
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
'@TestMethod("Nested Functions")
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
'@TestMethod("Floating Point Arithmetic")
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
'@TestMethod("Exponentiation Precedence")
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
'@TestMethod("Factorials")
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
'@TestMethod("Precedence")
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
'@TestMethod("Variables")
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
'@TestMethod("UDFs and Basic Array Functions")
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
'@TestMethod("Logical Operators with Numeric Output")
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
'@TestMethod("Logical Operators with Boolean Output")
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
'@TestMethod("Trig Functions")
Private Sub TestTrigFunctions()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "ROUND(tan(pi/4)^3-((3*sin(pi/4)-sin(3*pi/4))/(3*cos(pi/4)+cos(3*pi/4)));14)" _
                        )
    expected = "0"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("Mod Function")
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
'@TestMethod("String Arguments and Parameters")
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
'@TestMethod("String Arguments and Parameters")
Private Sub testStringArgumentes()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "REPLACE(x;'a';'A';1;2)" _
                        , "x = 'Capital'" _
                        )
    expected = "'CApitAl'"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("Implied Multiplication")
Private Sub ImpliedMultiplication()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "5(2)(3)(4)" _
                        )
    expected = "120"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("Array Arguments And Variables")
Private Sub DirectArrayVariableAssignment()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "MSUM(f;g)" _
                        , "f = {{1;0;4};{1;1;6}}; g = {{-3;0;-10};{2;3;4}}")
    expected = "{{-2;0;-6};{3;4;10}}"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("Assign Value To 2nd Order Variable")
Private Sub AssignValueTo2ndOrderVariable()
    Dim expr As VBAexpressions
    
    Set expr = New VBAexpressions
    With expr
        .Create "MMULT(f;g)"
        .VarValue("a") = "{1;0;4}"
        .VarValue("b") = "{1;1;6}"
        .VarValue("c") = "{-3;0;-10}"
        .ImplicitVarValue("f") = "ARRAY(a;b;c)"
        .ImplicitVarValue("g") = "INVERSE(f)"
        actual = .Eval
    End With
    expected = "{{1;0;0};{0;1;0};{0;0;1}}"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("Array Constructors And Parsers")
Private Sub ArrayConstructorsAndParsers()
    Dim expr As VBAexpressions
    Dim jaggedArr() As Variant
    
    Set expr = New VBAexpressions
    With expr
        jaggedArr() = .ArrayFromString2("{{1;0;4};{{1;1;6};{2;3}};{3};{2;5}}")
        actual = .ArrayToString(jaggedArr)
    End With
    expected = "{{1;0;4};{{1;1;6};{2;3}};{3};{2;5}}"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("Array Arguments And Variables")
Private Sub DirectArrayArguments()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "MSUM({{1;0;4};{1;1;6}};{{-3;0;-10};{2;3;4}})" _
                        )
    expected = "{{-2;0;-6};{3;4;10}}"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("Array Arguments And Variables")
Private Sub MatrixNegation()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "MNEG({{1;0;4};{1;1;6}})" _
                        )
    expected = "{{-1;0;-4};{-1;-1;-6}}"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("Array Arguments And Variables")
Private Sub VectorsMultiplication()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "MMULT({{1;0;4}};{{1;1;6}})" _
                        )
    expected = "25"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("Statistical Functions")
Private Sub NORM()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "ROUND(NORM(0.05);8)" _
                        )
    expected = "0.96012239"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("Statistical Functions")
Private Sub CHISQ()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "ROUND(CHISQ(4;15);8)" _
                        )
    expected = "0.99773734"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("Statistical Functions")
Private Sub GAUSS()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "ROUND(GAUSS(0.05);8)" _
                        )
    expected = "0.01993881"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("Statistical Functions")
Private Sub ERF()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "ROUND(ERF(0.05);8)" _
                        )
    expected = "0.05637198"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("Statistical Functions")
Private Sub STUDT()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "ROUND(STUDT(0.8;15);8)" _
                        )
    expected = "0.43619794"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("Statistical Functions")
Private Sub anorm()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "ROUND(ANORM(0.75);8)" _
                        )
    expected = "0.31863936"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("Statistical Functions")
Private Sub AGAUSS()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "ROUND(AGAUSS(0.75);8)" _
                        )
    expected = "0.67448975"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("Statistical Functions")
Private Sub AERF()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "ROUND(AERF(0.95);8)" _
                        )
    expected = "1.38590382"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("Statistical Functions")
Private Sub ACHISQ()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "ROUND(ACHISQ(0.75;15);8)" _
                        )
    expected = "11.03653766"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("Statistical Functions")
Private Sub FISHF()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "ROUND(FISHF(5.5;1.5;3);8)" _
                        )
    expected = "0.21407698"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("Statistical Functions")
Private Sub ASTUDT()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "ROUND(ASTUDT(0.05;15);8)" _
                        )
    expected = "2.13144955"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("Statistical Functions")
Private Sub AFISHF()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "ROUND(AFISHF(0.05;1.5;3);8)" _
                        )
    expected = "18.55325631"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("Statistical Functions")
Private Sub iBETA()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "ROUND(iBETA(0.5;1;3);8)" _
                        )
    expected = "0.875"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("Statistical Functions")
Private Sub BETAINV()
    On Error GoTo TestFail
    
    actual = GetResult( _
                        "ROUND(BETAINV(0.5;1;3);8)" _
                        )
    expected = "0.20629947"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("Linear Algebra Functions")
Private Sub QR()
    On Error GoTo TestFail
    Dim QRstr As String
    Dim QRarr As Variant
    
    Dim oHelper As VBAexpressions
    
    QRstr = GetResult("QR({{12;-51;4};{6;167;-68};{-4;24;-41}})")
    Set oHelper = New VBAexpressions
    With oHelper
        QRarr = .ArrayFromString2(QRstr)
        actual = GetResult("MROUND(MMULT(A;B);0)", "A=" & .ArrayToString(QRarr(0)) & ";" & "B=" & .ArrayToString(QRarr(1)))
    End With
    Set oHelper = Nothing
    expected = "{{12;-51;4};{6;167;-68};{-4;24;-41}}"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("Linear Algebra Functions")
Private Sub CholeskyDecomposition()
    On Error GoTo TestFail

    actual = GetResult("MROUND(MTRANSPOSE(CHOLESKY({{2.5;1.1;0.3};{2.2;1.9;0.4};{1.8;0.1;0.3}}));4)")
    expected = "{{1.5811;0.6957;0.1897};{0;1.19;0.2252};{0;0;0.4618}}" 'Octave output
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("Linear Algebra Functions")
Private Sub CholeskySolve()
    On Error GoTo TestFail

    actual = GetResult( _
                        "MROUND(CHOLSOLVE(ARRAY(a;b;c);{{'x';'y';'z'}};{{76;295;1259}};False);4)" _
                        , "a={6;15;55};b={15;55;225};c={55;225;979}")
    expected = "{{1;1;1}}"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("Linear Algebra Functions")
Private Sub CholeskyInverse()
    On Error GoTo TestFail

    actual = GetResult( _
                        "MROUND(CHOLINVERSE(ARRAY(a;b;c));4)" _
                        , "a={6;15;55};b={15;55;225};c={55;225;979}")
    expected = "{{0.8214;-0.5893;0.0893};{-0.5893;0.7268;-0.1339};{0.0893;-0.1339;0.0268}}" 'Octave output
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("Linear Algebra Functions")
Private Sub LSQRsolve()
    On Error GoTo TestFail

    actual = GetResult( _
                        "MROUND(LSQRSOLVE(A;b);4)" _
                        , "A={{2;4};{-5;1};{3;-8}};b={{10;-9.5;12}}")
    expected = "{{2.6576;-0.1196}}"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("Linear Algebra Functions")
Private Sub LUsolve()
    On Error GoTo TestFail

    actual = GetResult( _
                        "LUSOLVE(ARRAY(a;b;c);{{'x';'y';'z'}};{{2;3;4}};True)" _
                        , "a={1;0;4};b={1;1;6};c={-3;0;-10}")
    expected = "x = -18; y = -9; z = 5"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("Linear Algebra/Stats Functions")
Private Sub StraightLineFit()
    On Error GoTo TestFail

    actual = GetResult( _
                        "FIT(A;1)" _
                        , "A={{-2;40};{-1;50};{0;62};{1;58};{2;60}}")
    expected = "{{54 + 4.8*x};{0.7024}}"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("Linear Algebra/Stats Functions")
Private Sub PolynomialFit()
    On Error GoTo TestFail

    actual = GetResult( _
                        "FIT(A;1;4)" _
                        , "A={{-2;40};{-1;50};{0;62};{1;58};{2;60}}")
    expected = "{{62 + 3.6667*x -9.6667*x^2 + 0.3333*x^3 + 1.6667*x^4};{1}}"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("Linear Algebra/Stats Functions")
Private Sub ExponentialFit()
    On Error GoTo TestFail

    actual = GetResult( _
                        "FIT(A;2)" _
                        , "A={{0;0.1};{0.5;0.45};{1;2.15};{1.5;9.15};{2;40.35};{2.5;180.75}}")
    expected = "{{0.102*e^(2.9963*x)};{0.9998}}"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("Linear Algebra/Stats Functions")
Private Sub ExponentialFit2()
    On Error GoTo TestFail

    actual = GetResult( _
                        "FIT(A;3)" _
                        , "A={{0;10};{1;21};{2;35};{3;59};{4;92};{5;200};{6;400};{7;610}}")
    expected = "{{10.4992*(1.7959^x)};{0.9906}}"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("Linear Algebra/Stats Functions")
Private Sub PowerFit()
    On Error GoTo TestFail

    actual = GetResult( _
                        "FIT(A;4)" _
                        , "A={{2;27.8};{3;62.1};{4;110};{5;161}}")
    expected = "{{7.3799*x^1.9302};{0.9977}}"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("Linear Algebra/Stats Functions")
Private Sub LogarithmicFit()
    On Error GoTo TestFail

    actual = GetResult( _
                        "FIT(A;5)" _
                        , "A={{1;0.01};{2;1};{3;1.15};{4;1.3};{5;1.52};{6;1.84};{7;2.01};{8;2.05};{9;2.3};{10;2.25}}")
    expected = "{{0.9521*ln(x)+0.1049};{0.9752}}"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
'@TestMethod("Linear Algebra/Stats Functions")
Private Sub testMLR_FNominal()
    On Error GoTo TestFail

    actual = GetResult( _
                        "MLR(X;Y;True)" _
                        , "X={{1;1};{2;2};{3;3};{4;4};{5;1};{6;2};{7;3};{8;4}};Y={{2;4.1;5.8;7.8;5.5;5.2;8.2;11.1}}")
    expected = "{{0.0625 + 0.6438*X1 + 1.3013*X2};{0.9415;0.9181}}"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("Linear Algebra/Stats Functions")
Private Sub testMLR_CFNominal()
    On Error GoTo TestFail

    actual = GetResult( _
                        "MLR(X;Y;False)" _
                        , "X={{1;1};{2;2};{3;3};{4;4};{5;1};{6;2};{7;3};{8;4}};Y={{2;4.1;5.8;7.8;5.5;5.2;8.2;11.1}}")
    expected = "{{{{0.0625;0.6438;1.3013}}};{0.9415;0.9181}}"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("Linear Algebra/Stats Functions")
Private Sub testMLR_FNominal_Interactions()
    On Error GoTo TestFail

    actual = GetResult( _
                        "MLR(X;Y;True;'X1:X2')" _
                        , "X={{1;1};{2;2};{3;3};{4;4};{5;1};{6;2};{7;3};{8;4}};Y={{2;4.1;5.8;7.8;5.5;5.2;8.2;11.1}}")
    expected = "{{0.8542 + 0.4458*X1 + 0.945*X2 + 0.0792*X1*X2};{0.947;0.9072}}"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("Linear Algebra/Stats Functions")
Private Sub testMLR_CFNominal_Interactions()
    On Error GoTo TestFail

    actual = GetResult( _
                        "MLR(X;Y;False;'X1:X2')" _
                        , "X={{1;1};{2;2};{3;3};{4;4};{5;1};{6;2};{7;3};{8;4}};Y={{2;4.1;5.8;7.8;5.5;5.2;8.2;11.1}}")
    expected = "{{{{0.8542;0.4458;0.945;0.0792}}};{0.947;0.9072}}"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("Linear Algebra/Stats Functions")
Private Sub testMLR_FNamed_Interactions()
    On Error GoTo TestFail

    actual = GetResult( _
                        "MLR(X;Y;True;'Height:Width';'Height;Width')" _
                        , "X={{1;1};{2;2};{3;3};{4;4};{5;1};{6;2};{7;3};{8;4}};Y={{2;4.1;5.8;7.8;5.5;5.2;8.2;11.1}}")
    expected = "{{0.8542 + 0.4458*Height + 0.945*Width + 0.0792*Height*Width};{0.947;0.9072}}"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("Linear Algebra/Stats Functions")
Private Sub testMLR_CFNamed_Interactions()
    On Error GoTo TestFail

    actual = GetResult( _
                        "MLR(X;Y;False;'Height:Width';'Height;Width')" _
                        , "X={{1;1};{2;2};{3;3};{4;4};{5;1};{6;2};{7;3};{8;4}};Y={{2;4.1;5.8;7.8;5.5;5.2;8.2;11.1}}")
    expected = "{{{{0.8542;0.4458;0.945;0.0792}}};{0.947;0.9072}}"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("Linear Algebra/Stats Functions")
Private Sub testMLR_FNamed_MultiInteractions()
    On Error GoTo TestFail

    actual = GetResult( _
                        "MLR(X;Y;True;'Height:Width;Height:Height';'Height;Width')" _
                        , "X={{1;1};{2;2};{3;3};{4;4};{5;1};{6;2};{7;3};{8;4}};Y={{2;4.1;5.8;7.8;5.5;5.2;8.2;11.1}}")
    expected = "{{2.0875 + 2.08*Height -2.1075*Width -0.37*Height*Height + 0.7575*Height*Width};{0.9638;0.9155}}"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("Linear Algebra/Stats Functions")
Private Sub testMLR_CFname_MultiInteractions()
    On Error GoTo TestFail

    actual = GetResult( _
                        "MLR(X;Y;False;'Height:Width;Height:Height';'Height;Width')" _
                        , "X={{1;1};{2;2};{3;3};{4;4};{5;1};{6;2};{7;3};{8;4}};Y={{2;4.1;5.8;7.8;5.5;5.2;8.2;11.1}}")
    expected = "{{{{2.0875;2.08;-2.1075;-0.37;0.7575}}};{0.9638;0.9155}}"
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
