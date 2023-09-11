Attribute VB_Name = "Test_Strings"
'@TestModule
'@Folder("Tests")
'@IgnoreModule UseMeaningfulName
Option Explicit
Option Private Module

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup() 'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize() 'This method runs before every test in the module..
    
End Sub

'@TestCleanup
Private Sub TestCleanup() 'this method runs after every test in the module.
    On Error GoTo 0
End Sub



'@TestMethod
Public Sub Test_AscW_()
    On Error GoTo TestFail
    Dim Chr As Long
    Assert.IsTrue 72 = Strings.AscW2("H")
    Chr = 257
    Assert.IsTrue VBA.AscW(ChrW(Chr)) = Strings.AscW2(ChrW(Chr))
    Chr = 32767
    Assert.IsTrue 32767 = Strings.AscW2(ChrW(Chr))
    Chr = 32768
    Assert.IsTrue 32768 = Strings.AscW2(ChrW(Chr))
    Chr = 37769
    Assert.IsTrue 37769 = Strings.AscW2(ChrW(Chr))
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod
Public Sub Test_AscW_ParamValidation()
    On Error GoTo TestFail
    Dim Chr As Long
    Assert.IsTrue 72 = Strings.AscW2("Help me")
    Chr = 257
    Assert.AreEqual 0, Strings.AscW2(vbNullString) 'Throws error
    
TestExit:
    Exit Sub
TestFail:
    If Err.Number = 9 Then
        Resume TestExit
    Else
        Resume Next
    End If
End Sub



'@TestMethod
Public Sub Test_IsNullOrEmpty_()
    On Error GoTo TestFail
    Dim s As String
    Assert.IsTrue Strings.IsNullOrEmpty(s)
    s = vbNullString
    Assert.IsTrue Strings.IsNullOrEmpty(s)
    s = ""
    Assert.IsTrue Strings.IsNullOrEmpty(s)
    s = ";"
    Assert.IsFalse Strings.IsNullOrEmpty(s)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
