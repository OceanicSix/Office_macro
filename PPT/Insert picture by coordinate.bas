Attribute VB_Name = "Module2"
Sub ���ͼƬ()
    Dim sh As Shape
    Dim ppt As Slide
    
    Dim mypic, location, left, right, Size, height, width
    
    'Get input
    Dim Message1, Title, Default1, Message2, Default2, message3, Default3
    Message1 = "������Ҫ�޸ĵ��ı�������"
    Title = "�������"
    Default1 = "ͼƬ·��(·����Ҫ�����š� )"
    mypic = InputBox(Message1, Title, Default1)
    
    Message2 = "������ͼƬλ��(������� �� �ϱ� ��λ��,�ö��ŷֿ���PPTĬ�Ͽ�720��"
    Title = "�������"
    Default2 = "0,0"
    location = InputBox(Message2, Title, Default2)
    left = Split(location, ",")(0)
    right = Split(location, ",")(1)
    
     message3 = "������ͼƬ��С( ��Ⱥ͸߶ȣ��ö��ŷֿ���"
    Title = "�������"
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
