Attribute VB_Name = "Module3"

Sub 更改文字内容()
Dim sh As Shape
Dim ppt As Slide
Dim myShape As String
Dim myText As String


'Get input
Dim Message1, Title1, Default1, Message2, Title2, Default2
Message1 = "请输入要修改的文本框名称"
Title1 = "输入界面"
Default1 = "如文本框 1"
myShape = InputBox(Message1, Title1, Default1)

Message2 = "请输入要改成的文字内容"
Title2 = "输入界面"
Default2 = "这里写新内容"
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

