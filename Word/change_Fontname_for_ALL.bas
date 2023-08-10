Attribute VB_Name = "Module9"
Sub Change_FontName_for_all()
    Dim newFontName As String
    
    
    Dim message As String, title As String, default As String
    
    
    ' Prompt for font name
    message = "Which font name would you like to change to?"
    title = "InputBox"
    default = "Noto Sans S Chinese Bold"
    newFontName = InputBox(message, title, default)
    
    If newFontName = "" Then
        MsgBox "User cancelled", vbInformation
        Exit Sub
    End If
    

    For Each para In ActiveDocument.Paragraphs
        para.Range.Font.Name = newFontName
        
    Next
    MsgBox "Font name changed to " & newFontName, vbInformation
        
End Sub





