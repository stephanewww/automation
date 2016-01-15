
'---------------------------------------------------------------------------------------
' Module    : mdl_populatePPT
' Author    : swojewoda
' Date      : 13/01/2016
' Purpose   : main module to populate a PPT with a Excel source
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Procedure : populateMyPPT
' Author    : swojewoda
' Date      : 13/01/2016
' Purpose   : main function to populate a PPT through excel
' @todo     : use object format
'---------------------------------------------------------------------------------------
'
Sub populateMyPPT()
Dim filePath As String
Dim myXL As New Excel.Application
Dim myWS As Excel.Workbook
Dim cSlide As Slide
Dim nPPT As Presentation, nPPTName As String
Dim rapportErreur As String
   On Error GoTo populateMyPPT_Error
filePath = ModGenExcel.SelectionFichier

If Len(filePath) < 1 Then Exit Sub

Set myWS = myXL.Workbooks.Open(filePath)

nPPTName = ActivePresentation.Name & "_populated.pptm"
ActivePresentation.SaveCopyAs nPPTName
Set nPPT = ActivePresentation.Parent.Open(nPPTName)

For Each cSlide In ActivePresentation.Slides
    populateActiveSlide cSlide, myWS
    
Next
    
nPPT.Save

myWS.Close
myXL.Quit
Exit Sub

populateMyPPT_Error:

myWS.Close
myXL.Quit

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure populateMyPPT of Module mdl_populatePPT"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : populateActiveSlide
' Author    : swojewoda
' Date      : 14/01/2016
' Purpose   : go through a slide to complete it with the given ws
'---------------------------------------------------------------------------------------
'
Private Function populateActiveSlide(ByVal cSlide As Slide, ByVal xlWb As Workbook) As String

Dim cShape As Shape

   On Error GoTo populateActiveSlide_Error

For Each cShape In cSlide.Shapes
If cShape.HasTextFrame = msoTrue Then
Debug.Print cSlide.SlideIndex & " | " & cShape.TextFrame.TextRange.Text

    If hasToPopulateTheShape(cShape) = True Then _
        populateActiveSlide = populateActiveSlide & populateTheShape(cShape, xlWb)
End If
Next
    

   On Error GoTo 0
   Exit Function

populateActiveSlide_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure populateActiveSlide of Module mdl_populatePPT"
End Function

Private Function hasToPopulateTheShape(ByVal aShape As Shape) As Boolean

On Error GoTo endF

If isFieldContentToChange(aShape.TextFrame.TextRange.Text) = True Then _
    hasToPopulateTheShape = True

endF:

End Function

Private Function populateTheShape(ByVal aShape As Shape, ByVal xlWb As Workbook) As String
Dim strToChange, strToReplace, wsName, rangeString, tmpShapeTxtContent As String


tmpShapeTxtContent = aShape.TextFrame.TextRange.Text

While isFieldContentToChange(tmpShapeTxtContent) = True
    strToChange = findStringToChange(tmpShapeTxtContent)
    wsName = findSheetInString(strToChange)
    rangeString = findRangeInString(strToChange)
    
    
    
    If ModGenExcel.doesWSexist(wsName, xlWb) = True And Len(rangeString) > 0 Then
        strToReplace = xlWb.Sheets(wsName).Range(rangeString).Value
        
        aShape.TextFrame.TextRange.Replace "{{" & strToChange & "}}", strToReplace
        tmpShapeTxtContent = aShape.TextFrame.TextRange.Text
    End If
    
    Debug.Print " |" & wsName & " | " & rangeString & " | " & tmpShapeTxtContent
    
Wend

err_:
'aShape.TextFrame.TextRange.Text = tmpShapeTxtContent
 


End Function

Private Sub mockVariable()


End Sub


Function isFieldContentToChange(ByVal checkedString As String) As Boolean

isFieldContentToChange = checkedString Like "*{{*}}*"

End Function

Function findStringToChange(ByVal validString As String) As String
Dim startS, lengthS  As Long


If isFieldContentToChange(validString) = True Then

startS = InStr(1, validString, "{{") + 2
lengthS = InStr(1, validString, "}}") - startS
    findStringToChange = Mid(validString, _
                        startS, lengthS)
                        
End If
End Function


Function findSheetInString(ByVal validString As String) As String

Dim nCutSheet As Long
nCutSheet = InStr(1, validString, "!") - 1

If nCutSheet > 0 Then findSheetInString = Left(validString, nCutSheet)


End Function

Function findRangeInString(ByVal validString As String) As String
Dim nCutSheet As Long
nCutSheet = InStr(1, validString, "!")

If nCutSheet > 0 Then findRangeInString = Right(validString, Len(validString) - nCutSheet)


End Function
