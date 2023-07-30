Attribute VB_Name = "Module7"
Sub EN_update_vulnlist_table()
    Dim tb As Table
    Dim cell As cell
    
    Dim ToCStyle As String
    ToC_Style = "TOC 3"
    
    Dim tb_index As Integer
    tb_index = 3 'locate the target table (2nd table)
    Set tb = ActiveDocument.Tables(tb_index)
    
    ' Set prompt for vuln category
    Dim Message, Title, Default, category, counter
    Message = "Enter the vuln category"
    Title = "InputBox"    ' Set title.
    Default = "Web"    ' Set default.
    category = InputBox(Message, Title, Default)
    
    'set prompt for heading number
    Message = "Enter the heading number"
    Title = "InputBox"    ' Set title.
    Default = "2.1.1"    ' Set default.
    counter = InputBox(Message, Title, Default)
    
    If category = "" Or counter = "" Then
        MsgBox "User cancelled", vbInformation
        Exit Sub
    End If
    
    
    'calculate number of vuln
    Dim count As Integer
    count = 0
    
    Dim headText As New Collection
    search_pattern = Split(counter, ".")(0) & "." & Split(counter, ".")(1) & "*¡¾*¡¿*"
    For Each para In ActiveDocument.Paragraphs
        If para.Range.Text Like search_pattern And para.Style.NameLocal <> ToC_Style Then
            headText.Add para.Range.Text  'store heading text for later use
            count = count + 1
        End If
    Next para
    
    
    'loop through table
    For rowNum = 2 To count + 1 'skip the first row as it is table header
    
        'Add row if not exist
        If rowNum > tb.Rows.count Then
            Dim newRow
            Set newRow = tb.Rows.Add
            With newRow.Range.Font
            .Name = "Noto Sans S Chinese" ' Replace with the desired font name
            .Size = 10.5 ' Replace with the desired font size
            .Bold = False ' Set to True for bold, False for regular
            ' Add more font properties as needed
    End With
            
        End If
        
        tb.cell(rowNum, 1).Range.Text = (rowNum - 1)
        tb.cell(rowNum, 2).Range.Text = category
        tb.cell(rowNum, 3).Range.Text = Replace(Split(headText.Item(rowNum - 1), "¡¿")(1), vbCr, "") 'remove trailing new line character
        tb.cell(rowNum, 4).Range.Text = Split(Split(headText.Item(rowNum - 1), "¡¾")(1), "¡¿")(0) ' Retrieve text betweeen []
        tb.cell(rowNum, 5).Range.Text = "5.0"
           
    Next rowNum
    
    'Delete extra row
    If count + 1 < tb.Rows.count Then
        Dim extra_row
        For extra_row = count + 2 To tb.Rows.count
            tb.Rows(extra_row).Delete
        Next extra_row
    End If
    
    MsgBox "Update table successfully!", vbInformation
    
End Sub

