Class CObjectDocument
  Private mDict, mBody

  Public Property Get Body()
    Set Body = mBody
  End Property

  Public Function iTem(ByVal name)
    If Not Exists(name) Then 
      Err.Number = vbObjectError + 304
      Err.Description = "不存在该对象" & name
      Exit Function
    End If
    Set iTem = mDict.Item(lCase(name))
  End Function

  Public Function Exists(ByVal name)
    Exists = mDict.Exists(LCase(name))
  End Function

  Public Function Remove(ByVal name)
    If Exists(name) Then Dict.Remove(name)
  End Function

  Public Function Add(ByVal Obj)
    mDict.Add Obj.name, Obj
    Set Add = Obj
  End Function

  Private Sub Class_Initialize()
    Set mDict = CreateObject("Scripting.Dictionary")
    Set mBody = (New CObjectNode).Create(Me,0)
    mBody.name = "body"
    mBody.x = 0: mBody.y = 0
    Add mBody
  End Sub

  Private Sub Class_Terminate()
    mDict.RemoveAll
    Set mDict = Nothing
    If isObject(mBody) Then Set mBody = Nothing
  End Sub
End Class

Class CObjectNode
  Public X, Y, Angle, Alpha
  Private mParent, mDocument, mIndex
  Private mName, mWidth, mHeight, mImage, mVisual, mUsable
  Private mChildNode, mChildCount

  Public Property Get Usable()
    Usable = mUsable
  End Property

  Public Property Let Usable(ByVal value)
    mUsable = value
  End Property

  Public Property Get Visual()
    Visual = mVisual
  End Property

  Public Property Let Visual(ByVal value)
    mVisual = value
    mUsable = value
    Dim i:for i = 0 to mChildCount - 1
      mChildNode(i).Visual = mVisual
      mChildNode(i).Usable = mUsable
    next
  End Property

  Public Property Get Width()
    If mWidth Then
      Width = mWidth
    else
      Width = mImage.Width
    End If
  End Property

  Public Property Let Width(ByVal value)
    If Width < 0 Then 
      Err.Number = vbObjectError + 304
      Err.Description = "节点的宽度不能为负数"
      Exit Property
    End If
    mWidth = value
  End Property

  Public Property Get Height()
    If mHeight Then
      Height = mHeight
    else
      Height = mImage.Height
    End If
  End Property

  Public Property Let Height(ByVal value)
    If Height < 0 Then 
      Err.Number = vbObjectError + 305
      Err.Description = "节点的高度不能为负数"
      Exit Property
    End If
    mHeight = value
  End Property

  Public Property Get Name()
    Name = mName
  End Property

  Public Property Let Name(ByVal value)
    If len(value) = 0 Then
      Err.Number = vbObjectError + 302
      Err.Description = "子节点名称不能为空"
      Exit Property
    End If
    mName = lcase(value)
  End Property

  Public Property Get Parent()
    If isObject(mParent) Then Set Parent = mParent
  End Property

  Public Property Let Parent(ByVal obj)
    If Not isObject(obj) Then
      Err.Number = vbObjectError + 301
      Err.Description = "Object节点对象父级必须是对象"
      Exit Property
    End If
    Set mParent = obj
  End Property

  Public Property Get ChildNode(ByVal index)
    If index > mChildCount Or index < 0 Then
      Err.Number = vbObjectError + 303
      Err.Description = "索引超出范围"
      Exit Property
    End If
    Set ChildNode = mChildNode(index)
  End Property

  Public Property Get ChildCount()
    ChildCount = mChildCount
  End Property

  Public Property Get Image()
    If IsObject(mImage) Then Set Image = mImage 
  End Property

  Public Property Let Image(ByVal obj)
    If isObject(obj) Then Set mImage = obj
  End Property

  Public Sub SetImage(ByVal obj)
    If isObject(obj) Then Set mImage = obj
  End Sub
  
  Public Function GetX()
    if isObject(mParent) Then 
      GetX = mParent.GetX + X
      Exit Function
    End If
    GetX = 0
  End Function

  Public Function GetY()
    if isObject(mParent) Then 
      GetY = mParent.GetY + Y
      Exit Function
    End If
    GetY = 0
  End Function

  ' 注意这个是删除子节点中的对象
  Public Function Remove(ByVal index)
    aList.RemoveAt index
  End Function

  ' 这个是自我删除
  ' 这机制有个问题: 外部使用时记得手动释放引用的对象
  Public Function Del()
    If isObject(mParent) Then mParent.Remove mIndex
    If isObject(mDocument) Then mDocument.Remove mName
  End Function

  Public Function Add(ByVal namestr)
    ' 新建节点
    Dim node: Set node = (New CObjectNode).Create(mDocument, mChildCount)
    node.Parent = Me
    node.name = namestr
    mDocument.Add node
    mChildNode.Add node
    mChildCount = mChildCount + 1
    Set Add = node
  End Function

  Public Function Create(ByVal document,ByVal index)
    If index >= 0 Then mIndex = index 
    If isObject(document) Then Set mDocument = document
    Set mChildNode = CreateObject("System.Collections.ArrayList")
    X = 0: Y = 0: Visual = True: Alpha = 1: Set Create = Me
  End Function

  Private Sub Class_Terminate()
    If IsObject(mChildNode) Then 
      mChildNode.Clear
      Set mChildNode = Nothing
    End If
    If IsObject(mParent) Then Set mParent = Nothing
  End Sub
End Class

' 简化接口
Public Function Object(ByVal name)
  Set Object = Document.iTem(name)
End Function