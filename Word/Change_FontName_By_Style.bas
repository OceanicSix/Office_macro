Attribute VB_Name = "Module3"
Sub Change_FontName_By_Style()
    Dim myStyle As Style
    Dim newFontName As String
    
    
    Dim Message As String, Title As String, Default As String, targetStyle As String
    
    ' Prompt for style
    Message = "Which style do you need to modify?"
    Title = "InputBox"
    Default = "Ax 6ÕýÎÄ"
    targetStyle = InputBox(Message, Title, Default)
    
    ' Check if the style exists in the active document
    Dim foundStyle As Boolean
    foundStyle = False
    
    For Each myStyle In ActiveDocument.Styles
        If myStyle.NameLocal = targetStyle Then
            foundStyle = True
            Exit For
        End If
    Next myStyle
    
    If Not foundStyle Then
        MsgBox "The style '" & targetStyle & "' was not found in the document.", vbExclamation
        Exit Sub
    End If
    
    
    ' Prompt for font name
    Message = "Which font name would you like to change to?"
    Title = "InputBox"
    Default = "Arial"
    newFontName = InputBox(Message, Title, Default)
    
    If newFontName = "" Then
        MsgBox "User cancelled", vbInformation
        Exit Sub
    End If
    
    If foundStyle Then
        For Each para In ActiveDocument.Paragraphs
            If para.Style.NameLocal = targetStyle Then
                para.Range.Font.Name = newFontName
            End If
        Next para
        MsgBox "Font name for style '" & targetStyle & "' changed to " & newFontName, vbInformation

    End If
        
End Sub




