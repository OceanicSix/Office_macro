Attribute VB_Name = "Module4"
Sub 更改文字格式()
    Dim ppt As Slide
    Dim sh As Shape
    
    'Get input
    Dim Message1, Title1, Default1, Message2, Title2, Default2, message3, Default3, Message4, default4
    Dim myShape, Size, font_type, font_bold, font_color
    
    Message1 = "请输入要修改的文本框名称"
    Title = "输入界面"
    Default1 = "如文本框 1"
    myShape = InputBox(Message1, Title, Default1)
    
    Message2 = "请输入修改后的字体大小"
    Title = "输入界面"
    Default2 = "28"
    Size = InputBox(Message2, Title, Default2)
    
    message3 = "请输入修改后的字体"
    Title = "输入界面"
    Default3 = "微软雅黑"
    font_type = InputBox(message3, Title, Default3)
    
    message3 = "是否设置为粗体,设置是或否"
    Title = "输入界面"
    Default3 = "是"
    font_style = InputBox(message3, Title, Default3)
    
    If font_style = "是" Then
        font_bold = True
    Else
        font_bold = False
    End If
    
    message3 = "设置字体颜色，以RGB格式设置。通过逗号分开"
    Title = "输入界面"
    Default3 = "0,50,150"
    font_color = InputBox(message3, Title, Default3)
    red = Split(font_color, ",")(0)
    green = Split(font_color, ",")(1)
    blue = Split(font_color, ",")(2)
    
    For Each ppt In ActivePresentation.Slides
        On Error GoTo errhand
            Set sh = ppt.Shapes(myShape)
            sh.TextFrame.TextRange.Font.Size = Size
            sh.TextFrame.TextRange.Font.Name = font_type
            sh.TextFrame.TextRange.Font.Bold = font_bold
            sh.TextFrame.TextRange.Font.Color = RGB(red, green, blue)
        
errhand:
            Resume Next
    Next ppt
    
    
    
End Sub
