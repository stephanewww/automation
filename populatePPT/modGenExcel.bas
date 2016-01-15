'---------------------------------------------------------------------------------------
' Module    : ModGen
' Author    : swojewoda
' Date      : 06/10/2011
' Purpose   : Main functions to hack XL in an easyway
' reminder : To use filesystem, add "filesystemobject" dll with microsoft script control & microsoft scripting runtime
'---------------------------------------------------------------------------------------

Private Declare Function CloseClipboard Lib "user32" () As Long
 Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long

Public Sub EffacerPressePapier()

     OpenClipboard (0)
     EmptyClipboard
     CloseClipboard

End Sub

Public Function todayDate() As String
Dim str As String
str = VBA.Year(Now)
If Len(VBA.Month(Now)) = 1 Then
    str = str & "0" & VBA.Month(Now)
Else
    str = str & VBA.Month(Now)
End If

If Len(VBA.Day(Now)) = 1 Then
    todayDate = str & "0" & VBA.Day(Now)
Else
    todayDate = str & VBA.Day(Now)
End If

End Function

Public Sub saveMeToday()
On Error GoTo quitWithoutSaving
Dim myName As String, myfd As FileDialog

myName = ThisWorkbook.Path & "\CahierDesNouveautes" & todayDate & ".xlsm"

Set myfd = Application.FileDialog(msoFileDialogSaveAs)

With myfd
    .AllowMultiSelect = False
    
    .Title = "Enregistrer la fiche projet à la date du jour"
    .FilterIndex = 2
    .InitialFileName = myName
    .Show
    ThisWorkbook.SaveAs .SelectedItems(1), FileFormat:=52
    
End With

Exit Sub
quitWithoutSaving:


End Sub

Public Function doesWSexist(ByVal wsName As String, ByVal wb As Workbook) As Boolean
   On Error GoTo doesWSexist_Error
   doesWSexist = True
If Len(wb.Sheets(wsName).Name) > 0 Then Exit Function
    
doesWSexist_Error:
doesWSexist = False
End Function

Public Function TrouverColonne(ByVal NumCol As Long) As String

    Dim sResult As String

    sResult = Replace(VBA.Chr(Int(((NumCol - 1) / 26) + 64)) & VBA.Chr(((NumCol - 1) Mod 26) + 65), "@", "")
    TrouverColonne = sResult

End Function

Public Function lastLineOfWorksheet(projectSheet As Worksheet, Optional selectedRow As String = "A") As Long


lastLineOfWorksheet = projectSheet.Range(selectedRow & "65536").End(xlUp).Offset(0, 0).Row

End Function

Function rangeOfRange(xlWS As Worksheet, Optional topWS As Long = 2) As String
'renvoie une plage allant du bas à gauche en haut à droite d'une feuille
Dim basGauche As String, hautDroite As String

basGauche = xlWS.Range("A65536").End(xlUp).Address
hautDroite = xlWS.Range("IV" & topWS).End(xlToLeft).Offset(-1, 0).Address

rangeOfRange = basGauche & ":" & hautDroite

End Function

Function rangeOfRange2(xlWS As Worksheet, Optional ByVal decalage As Long = 0) As String
'renvoie une plage allant du bas à gauche en haut à droite d'une feuille
Dim basGauche As String, hautDroite As String

basGauche = xlWS.Range("A65536").End(xlUp).Address
hautDroite = xlWS.Range("IV1").End(xlToLeft).Offset(decalage, 0).Address

rangeOfRange2 = basGauche & ":" & hautDroite

End Function

Public Sub fileAbout()
Call MsgBox("Fiche Projet ITM-Alimentaire" _
            & vbCrLf & "Développement STIME - DOSI" _
            & vbCrLf & "(CC) 2010" _
            , vbInformation Or vbSystemModal, "A propos")


End Sub

Function SelectionFichier(Optional descr As String = "choisir le fichier", _
    Optional typeFichier As String = "*.xls, *.xlsx, *.xlsm", _
    Optional fdType As MsoFileDialogType = msoFileDialogFilePicker) As String
    
    Dim fd As FileDialog

    SelectionFichier = ""
    Set fd = Application.FileDialog(fdType)
    
    With fd
    
        .AllowMultiSelect = False
        .Filters.Clear
        .InitialFileName = ""
        .Filters.Add Description:=descr, Extensions:=typeFichier
        .Show

        If .SelectedItems.Count = 0 Then Exit Function
        SelectionFichier = .SelectedItems(1)
        
    End With
    
    Set fd = Nothing

End Function


Private Function joinObjectName(myObj As Object, Optional delimiter As String = ";") As String

Dim i, nbIter As Long

nbIter = myObj.Count
joinObjectName = myObj(1)
For i = 2 To nbIter
    joinObjectName = joinObjectName + delimiter + myObj(i)
    
Next

End Function
