Attribute VB_Name = "Module4"
Sub �������ָ�ʽ()
    Dim ppt As Slide
    Dim sh As Shape
    
    'Get input
    Dim Message1, Title1, Default1, Message2, Title2, Default2, message3, Default3, Message4, default4
    Dim myShape, Size, font_type, font_bold, font_color
    
    Message1 = "������Ҫ�޸ĵ��ı�������"
    Title = "�������"
    Default1 = "���ı��� 1"
    myShape = InputBox(Message1, Title, Default1)
    
    Message2 = "�������޸ĺ�������С"
    Title = "�������"
    Default2 = "28"
    Size = InputBox(Message2, Title, Default2)
    
    message3 = "�������޸ĺ������"
    Title = "�������"
    Default3 = "΢���ź�"
    font_type = InputBox(message3, Title, Default3)
    
    message3 = "�Ƿ�����Ϊ����,�����ǻ��"
    Title = "�������"
    Default3 = "��"
    font_style = InputBox(message3, Title, Default3)
    
    If font_style = "��" Then
        font_bold = True
    Else
        font_bold = False
    End If
    
    message3 = "����������ɫ����RGB��ʽ���á�ͨ�����ŷֿ�"
    Title = "�������"
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
