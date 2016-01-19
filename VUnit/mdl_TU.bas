'how to use the VUnit.cls

Sub atestClassTest()

Dim cTest As New VUnit
Dim myString As String, myRange As String

cTest.init
myRange = "Feuil1!A2"
myString = "moi je " & _
    VBA.Chr(10) & "hello {{" & myRange & "}} ok"



cTest.assertTrue "True : 1<2", 1<2
cTest.assertFalse "False 1>2", 1>2
cTest.assertEquals "Equals 1 = 1", 1,1
cTest.assertNotEquals "A<>B", A,B

End Sub
