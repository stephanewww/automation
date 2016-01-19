'---------------------------------------------------------------------------------------
' Module    : VUnit
' Author    : swojewoda
' Date      : 10/02/2014
' Purpose   : VBA UnitTest Class
'             UnitTest XUnit like, similar to VBLiteUnit, except smaller, and you can choose the way you want to present your results
' Version   : 0.02
' Changelog : Instance
' Requierements : Webshell, Microsoft scripting runtime, Microsoft script control, Microsoft Scriptlet librairy, Microsoft shell control and Automation
' Microsoft Regular Expression, Microsoft WMI Scripting, Microsoft_JScript, Microsoft Script Host Model Object, WSHControllerLibrairy
'---------------------------------------------------------------------------------------

Private pUnitTests As New Collection
Private pFS As New FileSystemObject
Private pTestStream As TextStream
Private pWSH As New WshShell
Private pTestCounter As Long
Private pOKTestCounter As Long
Private pTestContent As String
Private pTestFileName As String
Private pTestPath As String

Property Get TestFullPathName() As String
TestFullPathName = pTestPath & "\" & pTestFileName

End Property

Property Let testFileName(ByVal fn As String)
pTestFileName = fn
End Property

Property Get testFileName() As String
testFileName = pTestFileName
End Property

Property Let testPathName(ByVal pn As String)
pTestPath = pn
End Property
Property Get testPathName() As String
testPathName = pTestPath
End Property

Property Get OKTestCounter() As Long
OKTestCounter = pOKTestCounter

End Property

'---------------------------------------------------------------------------------------
' Procedure : addUTest
' Author    : swojewoda
' Date      : 10/02/2014
' Purpose   : ajout et réalisation d'un test unitaire
'             testTitle permet est la chaine de presentation du test
'             testValue est le test par lui meme => sum(1,1)=2
'             expectedValue est la valeur attendue
'---------------------------------------------------------------------------------------
'
Private Sub addUTest2(ByVal testTitle As String, ByVal testValue As Variant, ByVal expectedValue As Variant)
Dim testResult As Boolean: testResult = False
Dim testString As Boolean

On Error GoTo err_
pTestCounter = pTestCounter + 1
testString = IIf(testValue = expectedValue, True, False)
If testString = True Then pOKTestCounter = pOKTestCounter + 1

err_:
pUnitTests.Add (VBA.Format(pTestCounter, "0000") & " " & testTitle & " " & testString)

End Sub

Sub assertTrue(ByVal testTitle As String, ByVal testValue As Variant)

addUTest testTitle, testValue = True

End Sub

Sub assertFalse(ByVal testTitle As String, ByVal testValue As Variant)

addUTest testTitle, testValue = False

End Sub


Sub assertEquals(ByVal testTitle As String, ByVal testValue As Variant, ByVal expectedValue As Variant)


addUTest testTitle, IIf(testValue = expectedValue, True, False)

End Sub

Sub assertNotEquals(ByVal testTitle As String, ByVal testValue As Variant, ByVal expectedValue As Variant)

addUTest testTitle, IIf(testValue <> expectedValue, True, False)

End Sub

Sub IsError(ByVal testTitle As String, ByVal testValue As Variant)

addUTest testTitle, IIf(IsError(testValue) = True, True, False)


End Sub

Sub addUTest(ByVal testTitle As String, ByVal condExpression As Variant)
On Error GoTo err_
pTestCounter = pTestCounter + 1
testString = condExpression
If testString = True Then pOKTestCounter = pOKTestCounter + 1

err_:
pUnitTests.Add (VBA.Format(pTestCounter, "00000") & VBA.Chr(9) & testString & VBA.Chr(9) & testTitle)


End Sub


Private Sub Class_Initialize()

End Sub

'---------------------------------------------------------------------------------------
' Procedure : init
' Author    : swojewoda
' Date      : 10/02/2014
' Purpose   : initialisation de la class
' todo      : preparer une orientation de presentation soit texte soit cmd
'---------------------------------------------------------------------------------------
'
Function init(Optional ByVal testFileName As String = "", Optional ByVal testPathName As String = "") As Boolean


   On Error GoTo init_Error

pTestFileName = IIf(testFileName = "", "utests.bat", testFileName)
pTestPath = IIf(testPathName = "", ActivePresentation.Path, testPathName)
If (pFS.FolderExists(pTestPath) = True) Then Set pTestStream = pFS.CreateTextFile(Me.TestFullPathName)

init = True
init_Error:


End Function

'---------------------------------------------------------------------------------------
' Procedure : run
' Author    : swojewoda
' Date      : 10/02/2014
' Purpose   : modalite pour presenter les resultats
'---------------------------------------------------------------------------------------
'
Function run(Optional ByVal logAsCMD As Boolean = True) As Boolean

Dim t As String, delimiter As String, intro As String, endfile As String

delimiter = IIf(logAsCMD = True, "& echo ", vbCrLf)
intro = IIf(logAsCMD = True, "echo ", "")
endfile = "Total " & pOKTestCounter & "/" & pTestCounter & _
        IIf(logAsCMD = True, "& pause 10", "")
    For Each Item In pUnitTests
        t = t & Item & delimiter
    Next

pTestStream.Write intro & t & endfile
pTestStream.Close
If (logAsCMD = True) Then pWSH.run Me.TestFullPathName, vbNormalFocus, 2

End Function
