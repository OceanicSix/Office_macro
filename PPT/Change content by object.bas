Attribute VB_Name = "Module3"

Sub ������������()
Dim sh As Shape
Dim ppt As Slide
Dim myShape As String
Dim myText As String


'Get input
Dim Message1, Title1, Default1, Message2, Title2, Default2
Message1 = "������Ҫ�޸ĵ��ı�������"
Title1 = "�������"
Default1 = "���ı��� 1"
myShape = InputBox(Message1, Title1, Default1)

Message2 = "������Ҫ�ĳɵ���������"
Title2 = "�������"
Default2 = "����д������"
myText = InputBox(Message2, Title2, Default2)



'Change content
For Each ppt In ActivePresentation.Slides
    On Error GoTo errhand
        Set sh = ppt.Shapes(myShape)
        sh.TextFrame.TextRange.Text = myText
        
errhand:
        Resume Next
Next ppt

End Sub

