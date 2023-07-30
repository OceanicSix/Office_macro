Attribute VB_Name = "Module5"
Sub Reorder_Headings()
    Dim targetStyle As String
    
    'Defind style used in Table of content which need to be skipped later
    Dim ToCStyle As String
    ToC_Style = "TOC 3"
    
    
    Dim Message As String, Title As String, Default As String
    
    ' Set style name that will change heading order
    Message = "Which heading style you need to modify?"
    Title = "InputBox"
    Default = "Ax 3级标题"
    targetStyle = InputBox(Message, Title, Default)
    
    Dim counter As String
    Dim mainCounter As String
    Dim subCounter As Integer
    
    'Set heading Number
    
    Message = "Which heading you need to modify?"
    Title = "InputBox"
    Default = "2.1.1"
    counter = InputBox(Message, Title, Default)
    
    If targetStyle = "" Or counter = "" Then
        MsgBox "User cancelled", vbInformation
        Exit Sub
    End If
    
    
    mainCounter = Split(counter, ".")(0) & "." & Split(counter, ".")(1) 'return 2.1
    
    subCounter = 1
    
    pattern = mainCounter & "*【*】*"
    
    For Each para In ActiveDocument.Paragraphs
        ' Check for style name, contain String "2.1"
        If para.Style.NameLocal = targetStyle And para.Range.Text Like pattern Then
            ' Update the heading text with the new numbering format
            
            Dim headingText As String
            headingText = para.Range.Text
            
            ' Extract the text part of the heading (e.g 2.1.2【中危】SP登录界面漏洞)
            Dim textPart As String
            textPart = "【" + Split(headingText, "【")(1)
            
            ' Update the heading text with the new numbering format
            para.Range.ParagraphFormat.Style = targetStyle
            para.Range.Text = mainCounter & "." & subCounter & textPart
            
            ' Increment sub-counter for the next heading
            subCounter = subCounter + 1
        End If
    Next para
    
    ' Change the heading format back
    
    For Each para In ActiveDocument.Paragraphs
        ' Find para match heading pattern and not in ToC
        If para.Range.Text Like pattern And para.Style.NameLocal <> ToC_Style Then
            para.Range.ParagraphFormat.Style = targetStyle
        End If
    Next para
    
    MsgBox "Headings reordered successfully!", vbInformation
End Sub

