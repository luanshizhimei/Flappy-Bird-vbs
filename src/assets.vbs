Class CAssets
  Private mParent
  Private mDict
  Private mFso
  Private mPath
  Public ScriptName

  Public Property Get File(ByVal key)
    If Not mDict.Exists(key) then 
      Err.Number = vbObjectError + 206
      Err.Description = "该文件尚未载入: " & key
      Exit Property
    End If
    Set File = mDict.Item(key)
  End Property

  Private Function LoadImger(ByVal key,ByVal obj)
    If mDict.Exists(key) Then
      Err.Number = vbObjectError + 203
      Err.Description = "已载入该文件: " & key
      Exit Function
    End If

    Dim imger: Set imger = New CAssetsImage
    If IsObject(obj) Then ' imger对象
      Set imger.Parent = obj
    Elseif len(obj) Then ' 路径
      Set imger.Parent = mParent.CreateImage()
      imger.Parent.src = obj
    End If

    mDict.Add key, imger
    Set LoadImger = imger
  End Function

  Private Sub LoadXML(ByVal xml, ByVal obj)
    Dim Node: For Each Node In xml.childNodes
      If LCase(Node.nodeName) = "picture" Then
        Dim name, src, Image
        name = Node.getAttribute("name")
        src = Node.getAttribute("src")
        If Len(src) Then
          Set Image = LoadImger(name, GetPath(src))
        ElseIf IsObject(obj) Then
          Set Image = LoadImger(name, obj)
        End If
        ' 设置
        With Image
          .x = int(Node.getAttribute("x"))
          .y = int(Node.getAttribute("y"))
          .width = int(Node.getAttribute("width"))
          .height = int(Node.getAttribute("height"))
        End with
        ' 迭代
        Call LoadXML(Node,Image.Parent)
      End If
    Next
  End Sub

  Private Function GetPath(path)
    Dim agreemark, srt
    agreemark = InStr(path,":/")
    If lCase(Left(path,5)) = "data:" Then 
      GetPath = path
      Exit Function
    End If 
    If agreemark Then
      ' 网页地址解析（我当时怎么想的？）
      Dim httpmark: httpmark = LCase(Left(path,agreemark - 1))
      If httpmark = "http" Or httpmark = "https" Then
        GetPath = path: Exit Function 
      End If
      GetPath = path
      Exit Function
    Else
      ' 相对路径转换为绝对路径
      Dim pathidx, rootidx, length, char, regstr
      length = Len(path)
      rootidx = Len(mPath)
      Do While (char = "." Or char = "/" Or char = Empty) And pathidx < length
        pathidx = pathidx + 1
        char = Mid(path,pathidx,1)
        regstr = regstr & char
        If regstr = "../" Then
          rootidx = InStrRev(mPath,"/",rootidx - 1)
          regstr = ""
        End If
      Loop
      srt = Left(mPath, rootidx) & Right(path, length - pathidx + 1)
    End If
    
    ' 判断文件是否存在（仅本地文件）
    If Not mFso.FileExists(path) Then
      Err.Number = vbObjectError + 201
      Err.Description = "文件不存在: " & path
      Exit Function
    End If

    GetPath = srt
  End Function

  Public Sub Load(ByVal path)
    Dim lenght, sufidx, suff, nameidx, name
    path = getPath(path)
    lenght = len(path)
    sufidx = InStrRev(path,".")
    suff = lcase(Right(path, lenght - sufidx))
    nameidx = InStrRev(path,"/") + 1
    name = Mid(path, nameidx,sufidx - nameidx)

    Select Case suff
      Case "jpg","jpeg","png","bmp"
        Dim imger, w, h
        Set imger = LoadImger(name, path)
        imger.X = 0: imger.Y = 0
        w = .Parent.Width
        h = .Parent.Height
        ' 如果浏览器没有反应过来就用wia来获得分辨率
        If Not(w and h) Then
          With CreateObject("WIA.ImageFile")
            .LoadFile path
            w = .Width: h = .Height
          End With
        End If
        imger.Width = w: imger.Height = h

      Case "xml"
        With CreateObject("MSXML2.DOMDocument")
          .async = False
          .load path
          If .parseError.errorCode Then 
            Err.Number = vbObjectError + 203
            Err.Description = "xml配置错误(" & .parseError.reason & ")"
            Exit Sub
          End If
          call LoadXML(.documentElement,vbNull)
        End With

      Case "vbs"
        Dim file, line
        set file=mFso.opentextfile(path,1,false)
        Do While file.AtEndOfStream = False
          line = file.Readline
          If LCase(Left(line, 11)) = Chr(39) & "#file:xml;" then
            With CreateObject("MSXML2.DOMDocument")
              .async = False
              .loadXML Mid(line, 12, Len(line) - 11)
              If .parseError.errorCode Then 
                Err.Number = vbObjectError + 203
                Err.Description = "xml配置错误(" & .parseError.reason & ")"
                Exit Sub
              End If
              call LoadXML(.documentElement,vbNull)
            End With
          End If
        loop
      Case path = false
        Exit Sub

      Case Else
        Err.Number = vbObjectError + 202
        Err.Description = "导入未知资源文件"
    End Select
  End Sub
  
  Public Function Create(ByVal parent)
    Set Create = Me: Set mParent = parent
    Set mDict = CreateObject("Scripting.Dictionary")
    Set mFso = CreateObject("Scripting.FileSystemObject")
    ScriptName = Wscript.ScriptFullName
    mPath = Replace(mFso.GetFile(ScriptName).ParentFolder.Path,"\","/")
    ScriptName = Replace(ScriptName,"\","/")
    If Right(mPath,1) <> "/" Then mPath = mPath & "/"
  End Function

  Private Sub Class_Terminate()
    If IsObject(mFso) Then Set mFso = Nothing
    If IsObject(mDict) Then 
      Call mDict.RemoveAll()
      Set mDict = Nothing
    End If
    If IsObject(mParent) Then Set mParent = Nothing
  End Sub
End Class

Class CAssetsImage
  Public X, Y, Width, Height
  Public Parent
End Class

' 简化接口
Public Function File(ByVal key)
  Set File = Assets.File(key)
End Function