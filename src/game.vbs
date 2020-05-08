Class CGame
  Private mState

  Public Property Get State()
    State = mState
  End Property

  Private Sub Render(ByVal obj)
    ' 对象可视
    If Not obj.Visual Then Exit Sub
    ' 绘图
    If IsObject(Obj.Image) Then
      Dim x, y, w, h
      w = Obj.Width: h = Obj.height
      x = - w / 2: y = - h / 2
      Display.save
      Display.translate Obj.GetX() - x , Obj.GetY() - y
      Display.globalAlpha Obj.Alpha
      Display.rotate Obj.angle
      Display.drawImage Obj.Image.Parent , _
        Obj.Image.X, Obj.Image.Y, Obj.Image.Width, Obj.Image.height, _ 
        x, y, w, h
      Display.restore
    End If
    ' 迭代
    if Obj.ChildCount then
      Dim i:for i = 0 to Obj.ChildCount - 1
        Render Obj.ChildNode(i)
      next
    end If
  End Sub

  Public Sub UpDate(ByVal SpendTime)
    If mState = 0 Then Exit Sub
    Call Clock.UpDate(SpendTime)
    Call Render(Document.Body)
    Call Display.UpDate()
  End Sub

  Public Sub Quit()
    mState = 0
    If isObject(Display) Then
      Display.Close()
      Wscript.Quit()
    End If
  End Sub

  Public Function Start(ByVal funcname)
    mState = 1
    Dim func: Set func = GetRef(funcname): Call func
  End Function

  Private Sub Class_Initialize()
    Set Display = (New CDisplay).Create(GAME_TITLE,DISPLAY_WIDTH,DISPLAY_HEIGHT)
    Set Assets = Display.Assets
    Set EventHandle = Display.EventHandle
    Set Document = New CObjectDocument
    Set Clock = New CClock
  End Sub

  Private Sub Class_Terminate()
    If IsObject(Clock) Then Set Clock = Nothing
    If IsObject(Document) Then Set Document = Nothing
    If IsObject(Assets) Then Set Assets = Nothing
    If IsObject(Display) Then Set Display = Nothing
  End Sub
End Class