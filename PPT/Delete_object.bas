Attribute VB_Name = "Module1"
Sub 删除图形()
    Dim sh As Shape
    Dim ppt As Slide
    Dim myShape As String
  
    'Get input
    Dim Message1, Title1, Default1
    Message1 = "请输入要删除的图形名称"
    Title1 = "输入界面"
    Default1 = "如图片 1"
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





