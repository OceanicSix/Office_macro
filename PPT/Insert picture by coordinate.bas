Attribute VB_Name = "Module2"
Sub 添加图片()
    Dim sh As Shape
    Dim ppt As Slide
    
    Dim mypic, location, left, right, Size, height, width
    
    'Get input
    Dim Message1, Title, Default1, Message2, Default2, message3, Default3
    Message1 = "请输入要修改的文本框名称"
    Title = "输入界面"
    Default1 = "图片路径(路径不要有引号” )"
    mypic = InputBox(Message1, Title, Default1)
    
    Message2 = "请输入图片位置(距离左边 和 上边 的位置,用逗号分开。PPT默认宽720）"
    Title = "输入界面"
    Default2 = "0,0"
    location = InputBox(Message2, Title, Default2)
    left = Split(location, ",")(0)
    right = Split(location, ",")(1)
    
     message3 = "请输入图片大小( 宽度和高度，用逗号分开）"
    Title = "输入界面"
    Default3 = "100,100"
    Size = InputBox(message3, Title, Default3)
    width = Split(Size, ",")(0)
    height = Split(Size, ",")(1)
    
    
    
    'Add pic
    For Each ppt In ActivePresentation.Slides
        On Error GoTo errhand
           ppt.Shapes.AddPicture(mypic, msoFalse, msoTrue, left, right, width, height).Select
errhand:
            Resume Next
    Next ppt
End Sub
