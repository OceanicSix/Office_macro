Attribute VB_Name = "Module1"
Sub ɾ��ͼ��()
    Dim sh As Shape
    Dim ppt As Slide
    Dim myShape As String
  
    'Get input
    Dim Message1, Title1, Default1
    Message1 = "������Ҫɾ����ͼ������"
    Title1 = "�������"
    Default1 = "��ͼƬ 1"
    myShape = InputBox(Message1, Title1, Default1)
    ' Delete item
    For Each ppt In ActivePresentation.Slides
        On Error GoTo errhand
            Set sh = ppt.Shapes(myShape)
            sh.Delete
            
errhand:
            Resume Next
    Next ppt

End Sub





