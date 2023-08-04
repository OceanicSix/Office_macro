Attribute VB_Name = "Module8"

Function get_csvv_score(vul_type As String) As String
    Dim result As String
    
    If InStr(vul_type, "高") > 0 Or InStr(vul_type, "High") > 0 Or InStr(vul_type, "high") > 0 Then
        result = "8.0"
    ElseIf InStr(vul_type, "中") > 0 Or InStr(vul_type, "Medium") > 0 Or InStr(vul_type, "medium") > 0 Then
        result = "6.0"
    ElseIf InStr(vul_type, "低") > 0 Or InStr(vul_type, "Low") > 0 Or InStr(vul_type, "low") > 0 Then
        result = "4.0"
    Else
        result = "0"
    End If
    
    get_csvv_score = result
End Function

' give a vuln heading ---2.1.2, find its parent heading ---2.1 web
Function get_heading(heading_no As String, vuln_heading As Collection) As String
    For Each Item In vuln_heading
        If InStr(Item, Split(heading_no, ".")(0) & "." & Split(heading_no, ".")(1)) Then
            result = Item
        End If
    Next Item
    get_heading = result
End Function
' give a parent heading ----2.1 web, find vuln category
Function get_category(vul_type As String) As String
    Dim result As String
    
    If InStr(vul_type, "web") > 0 Or InStr(vul_type, "Web") > 0 Then
        result = "Web"
    ElseIf InStr(vul_type, "android") > 0 Or InStr(vul_type, "Android") > 0 Then
        result = "Android"
    ElseIf InStr(vul_type, "iOS") > 0 Or InStr(vul_type, "IOS") > 0 Or InStr(vul_type, "ios") > 0 Then
        result = "iOS"
    Else
        result = "something wrong with the category"
    End If
    
    get_category = result
End Function


'first go over all paragraph, and store the heading text for 2.* and 2.*.*.[]

' Then add table rows and enter content;

' for category, will compare heading_no 2.1.1 with an array ( 2.1 web, 2.2 android, 2.3 iOS) and find the right category

' for csvv score, will compare level of vuln with array ( high, medium, low) and return the score

Sub Update_vulnlist_table()
    Dim tb As Table
    Dim cell As cell
    
    Dim ToCStyle As String
    ToC_Style = "TOC 3"
    
    Dim tb_index As Integer
    tb_index = 3 'locate the target table (2nd table)
    Set tb = ActiveDocument.Tables(tb_index)
    
    Dim message, title, default, counter
    
    'set prompt for heading number
    message = "Enter the heading number"
    title = "InputBox"    ' Set title.
    default = "2"    ' Set default.
    counter = InputBox(message, title, default)
    
    If counter = "" Then
        MsgBox "User cancelled", vbInformation
        Exit Sub
    End If
    
    
    'calculate number of vuln
    Dim count As Integer
    count = 0
    
    Dim headText As New Collection
    Dim vuln_heading As New Collection 'store all parent heading for vuln like "2.1 Web", "2.2 Android"
    
    search_pattern = counter & "." & "*.*【*】*"
    vuln_heading_pattern = counter & ".*"
    
    
    For Each para In ActiveDocument.Paragraphs
        If para.Range.Text Like search_pattern And para.Style.NameLocal <> ToC_Style Then
            headText.Add para.Range.Text  'store heading text for later use
            count = count + 1
        End If
        
        If para.Range.Text Like vuln_heading_pattern And para.Style.NameLocal = "Ax 2级标题" Then
            vuln_heading.Add para.Range.Text 'Store parent vuln heading
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
        
        'set font style for each cell in a row
        Dim col_num
        col_num = tb.Columns.count
        
        For col = 1 To col_num
            tb.cell(rowNum, col).Range.Font.Name = "Noto Sans S Chinese"
            tb.cell(rowNum, col).Range.Font.Size = 10.5
            tb.cell(rowNum, col).Range.Font.Bold = False
        Next col
        'add content
        
        tb.cell(rowNum, 1).Range.Text = Split(headText.Item(rowNum - 1), "【")(0) 'return heading no, like 2.1.2
        tb.cell(rowNum, 2).Range.Text = get_category(get_heading(tb.cell(rowNum, 1).Range.Text, vuln_heading))
        tb.cell(rowNum, 3).Range.Text = Replace(Split(headText.Item(rowNum - 1), "】")(1), vbCr, "") 'remove trailing new line character
        tb.cell(rowNum, 4).Range.Text = Split(Split(headText.Item(rowNum - 1), "【")(1), "】")(0) ' Retrieve text betweeen []
        tb.cell(rowNum, 5).Range.Text = get_csvv_score(tb.cell(rowNum, 4).Range.Text) ' get csvv socre based on vuln level
           
    Next rowNum
    
    'Delete extra row
    If count + 1 < tb.Rows.count Then
        Dim extra_row
        Dim table_num
        table_num = tb.Rows.count
       For extra_row = count + 2 To table_num
            tb.Rows(count + 2).Delete
        Next extra_row
    End If
    
    MsgBox "Update table successfully!", vbInformation
    
End Sub




