Attribute VB_Name = "Module2"
Sub Change_FontSize_By_Size()
    Dim newFontSize As Double
    
    Dim targetFontSize As Double
    
    
    Dim Message As String, Title As String, Default As String, targetStyle As String
    
    ' Prompt for target font size
    Message = "Which font size you need to modify?"
    Title = "InputBox"
    Default = "10.5"
    targetFontSize = InputBox(Message, Title, Default)
    

    ' Prompt for new fontsize
    Message = "Which fontsize would you like to change to?"
    Title = "InputBox"
    Default = "12"
    newFontSize = InputBox(Message, Title, Default)
    
    If newFontSize = 0 Or targetFontSize = 0 Then
        MsgBox "User cancelled", vbInformation
        Exit Sub
    End If
    
    For Each para In ActiveDocument.Paragraphs
        For Each Char In para.Range.Characters
            If Char.Font.Size = targetFontSize Then
                Char.Font.Size = newFontSize
            End If
        Next Char
    Next para
    MsgBox "Font size of " & targetFontSize & " changed to " & newFontSize & " points.", vbInformation
        
End Sub



