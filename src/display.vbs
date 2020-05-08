Class CDisplay
  Private mIE, mDOM, mAssets, mEventHandle
  Private mDisplay, mDisContext, mBuffer, mBufContext
  Private mWidth, mHeight

  Public Property Get Assets()
    Set Assets = mAssets
  End Property

  Public Property Get EventHandle()
    Set EventHandle = mEventHandle
  End Property

  ' 提供具体的绘图API
  Public Sub save()
    mBufContext.save
  End Sub

  Public Sub restore()
    mBufContext.restore
  End Sub

  Public Sub translate(ByVal x, ByVal y)
    mBufContext.translate x, y
  End Sub

  Public Sub rotate(ByVal angle)
    mBufContext.rotate angle * 0.0174533 '3.1415926535 / 180
  End Sub

  Public Sub globalAlpha(ByVal alpha)
    mBufContext.globalAlpha = alpha
  End Sub
  
  Public Sub drawImage(ByVal img, _
    ByVal sx,ByVal sy,ByVal swidth,ByVal sheight, _
    ByVal x,ByVal y,ByVal width,ByVal height)
    mBufContext.drawImage img, _
      sx, sy, swidth, sheight, _
      x, y, width, height 
  End Sub

  Public Function CreateImage()
    Set CreateImage = mDOM.createElement("img")
  End Function

  Public Property Get Visible()
    If isObject(mIE) Then Visible = mIE.visible
  End Property

  Public Sub Close()
    If Visible Then mIE.Quit()
  End Sub

  Public Sub Update()
    Call mDisContext.drawImage(mBuffer, 0, 0)
  End Sub

  Public Function Create(ByVal title, ByVal width, ByVal height)
    Set Create = Me
    ' 关掉所有IE, 避免影响程序
    Dim process: For Each process In GetObject("WinMgmts:").InstancesOf("Win32_Process")
      If LCase(process.name) = "iexplore.exe" then process.terminate()
    Next

    ' 新建ie窗口
    Set mIE = wscript.CreateObject("internetexplorer.application","BaiscEvent_")
    If Not IsObject(mIE) Then
      Err.Number = vbObjectError + 101
      Err.Description = "不能获取IE浏览器对象"
      Exit Function
    end if
  
    ' 设置ie对象
    With mIE
      .visible = False: .navigate "about:blank"
      .MenuBar = False: .AddressBar = False
      .ToolBar = False: .StatusBar  = False
      .Resizable = False ' 禁止修改窗口大小(pass - 下次做自动缩放)
      .Width = width +  (.Width - .Document.body.clientWidth)
      .Height = height + (.Height - .Document.body.clientHeight)
      .Left = fix((.document.parentwindow.screen.availwidth  - .width)/2)
      .Top = fix((.document.parentwindow.screen.availheight - .height)/2)
    End With
    Set mDOM = mIE.Document

    ' 判断是否为支持H5的浏览器。
    Dim ieUser,ieVer
    ieUser = mDOM.parentwindow.navigator.userAgent
    If InStrRev(ieUser,"Trident") = 0 Then
      Err.Number = vbObjectError + 102
      Err.Description = "启动的浏览器不是Trident内核"
      Exit Function
    End If

    If InStrRev(ieUser,"rv:11.0") = 0 Then
      ' 理论上ie9就支持H5。这里为避免兼容性的问题，直接就写死11。
      ' pass - 可以加入低版本检查，11以下的版本标记是MSIE。
      Err.Number = vbObjectError + 103
      Err.Description = "IE版本低于11"
      Exit Function
    End If

    ' 设置基础结构
    mDOM.body.innerhtml = "<!DOCTYPE html><html lang=zh-CH>" & _ 
      "<head><meta charset=UTF-8/><title>" & title & "</title>" & _
      "<style type=text/css>body,canvas{margin: 0px;padding: 0px;overflow:hidden;}" & _
      "canvas{background-color:#000;}</style></head><body scroll=no></body></html>"
    set mDisplay = mDOM.createElement("Canvas")
    set mDisContext = mDisplay.GetContext("2d")
    mDisplay.Width = width: mDisplay.height = height
    mDOM.body.appendChild(mDisplay)
    set mBuffer = mDOM.createElement("Canvas")
    set mBufContext = mBuffer.GetContext("2d")
    mBuffer.Width = width: mBuffer.height = height

    ' 创建附属类
    Set mAssets = (New CAssets).Create(Me)
    Set mEventHandle = (New CEvent).Create(mDOM)

    ' 加载完成显示
    mIE.visible = True
  End Function

  Private Sub Class_Terminate()
    Call Close()
    If IsObject(mDisContext) Then Set mDisContext = Nothing
    If IsObject(mDisplay) Then Set mDisplay = Nothing
    If IsObject(mBufContext) Then Set mBufContext = Nothing
    If IsObject(mBuffer) Then Set mBuffer = Nothing
    If IsObject(mAssets) Then Set mAssets = Nothing
    If IsObject(mEventHandle) Then Set mEventHandle = Nothing
    If IsObject(mDOM) Then Set mDOM = Nothing
    If IsObject(mIE) Then Set mIE = Nothing
  End Sub
End Class