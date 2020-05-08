Class CEvent
  Private mParent
  Private mArrayList, mArrayCount, mDict

  Public Property Get iTem(ByVal key)
    Set iTem = mDict.Item(key)
  End Property

  Public Property Get Object(ByVal idx)
    Set Object = mArrayList(idx)
  End Property

  Public Property Get Count()
    Count = mArrayCount
  End Property

  ' ÉèÖÃ
  Public Function SetMousedown(ByVal Obj,ByVal funcname)
    SetMousedown = Add(Obj, funcname, 1)
  End Function

  Public Function SetMouseup(ByVal Obj,ByVal funcname)
    SetMouseup = Add(Obj, funcname, 2)
  End Function

  Public Function SetClick(ByVal Obj,ByVal funcname)
    SetClick = Add(Obj, funcname, 3)
  End Function

  Private function Add(ByVal Obj,ByVal lfuncname, ByVal mode)
    Dim eventobj: Set eventobj = New CEventNode
    With eventobj 
      Set .func = GetRef(lfuncname)
      Set .Parent = Me
      Set .Node = Obj
      .funcname = lfuncname
      .Mode = mode
      .Index = mArrayCount
    End With
    mArrayList.Add eventobj
    Add = mArrayCount
    mArrayCount = mArrayCount + 1
  End Function

  Private Sub Move(ByVal idx, ByVal key)
    mArrayList.RemoveAt idx
    mDict.Remove key
    mArrayCount = mArrayCount - 1
  End Sub

  Public Function Create(ByVal parent)
    Set Create = Me: Set mParent = parent: mArrayCount = 0
    Set mDict = CreateObject("Scripting.Dictionary")
    Set mArrayList = CreateObject("System.Collections.ArrayList")
    mParent.parentwindow.onbeforeunload = GetRef("BaiscEvent_onbeforeunload")
    mParent.parentwindow.onmousedown = GetRef("BaiscEvent_onmousedown")
    mParent.parentwindow.onmouseup = GetRef("BaiscEvent_onmouseup")
    mParent.parentwindow.onclick = GetRef("BaiscEvent_onclick")
  End Function

  Private Sub Class_Terminate()
    If IsObject(mDict) Then Set mDict = Nothing
    If IsObject(mArrayList) Then Set mArrayList = Nothing
    If IsObject(mParent) Then Set mParent = Nothing
  End Sub
End Class

Class CEventNode
  Public func, funcname, Node, Mode, parent, Index

  Public Sub Del()
    parent.Move Index, funcname
  End Sub

  Private Sub Class_Terminate()
    If IsObject(func) Then Set func = Nothing
    If IsObject(Node) Then Set Node = Nothing
    If IsObject(parent) Then Set parent = Nothing
  End Sub
End Class

Public Function SetMousedown(ByVal Obj,ByVal funcname)
  SetMousedown = EventHandle.SetMousedown(obj, funcname)
End Function

Public Function SetMouseup(ByVal Obj,ByVal funcname)
  SetMouseup = EventHandle.SetMouseup(obj, funcname)
End Function

Public Function SetClick(ByVal Obj,ByVal funcname)
  SetClick = EventHandle.SetClick(obj, funcname)
End Function

' ÊÂ¼þ
Public Sub BaiscEvent_onBeforeUnload()
  Call Game.Quit()
End Sub

Public Sub BaiscEvent_onmousedown(ByVal e)
  mouseEvent e.offsetX, e.offsetY, 1
End Sub

Public Sub BaiscEvent_onmouseup(ByVal e)
  mouseEvent e.offsetX, e.offsetY, 2
End Sub

Public Sub BaiscEvent_onclick(ByVal e)
  mouseEvent e.offsetX, e.offsetY, 3
End Sub

Public Sub mouseEvent(ByVal x, ByVal y, ByVal Mode)
  Dim num, node, obj
  num = EventHandle.Count
  if num = 0 then Exit Sub
  Dim i:for i = num - 1 to 0 step -1
    Set obj = EventHandle.Object(i)
    Set node = obj.node

    If node.Usable And _ 
      x >= node.x And x <= node.x + node.width And _ 
      y >= node.y And y <= node.y + node.Height And _ 
      obj.mode = Mode _
    Then
    obj.func node
    End If
  next
End Sub