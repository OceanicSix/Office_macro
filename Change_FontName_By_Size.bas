Attribute VB_Name = "Module4"
Sub Change_FontName_By_Size()
    Dim newFontName
    Dim targetFontSize As Double
    
    
    Dim Message As String, Title As String, Default As String
    
    ' Prompt for target font size
    Message = "Which font size you need to modify?"
    Title = "InputBox"
    Default = "10.5"
    targetFontSize = InputBox(Message, Title, Default)
    

    ' Prompt for new font name
    Message = "Which font name would you like to change to?"
    Title = "InputBox"
    Default = "Noto Sans S Chinese Bold"
    newFontName = InputBox(Message, Title, Default)
    
    If newFontName = "" Or targetFontSize = 0 Then
        MsgBox "User cancelled", vbInformation
        Exit Sub
    End If
    
    For Each para In ActiveDocument.Paragraphs
        For Each Char In para.Range.Characters
            If Char.Font.Size = targetFontSize Then
                Char.Font.Name = newFontName
            End If
        Next Char
    Next para
    MsgBox "Font size of " & targetFontSize & " changed to " & newFontName, vbInformation
        
End Sub




