'合成脚本

'pass - 这个可以做成命令行
'pass - 实现图片资源的预编译
' 考虑还是不在这个脚本里实现这个功能，考虑将资源assets.xml独立出来。然后通过这个脚本将资源xml数据直接编译到生产的脚本文件中。
' 关于那个assets.xml大概主要还是指向程序所需的资源比如：图片，图片切片，文本之类的。然后再通过这个合成脚本直接输出到成品脚本的常量中。(之后想了一下要不还是做成内嵌文件的形式？)
' assets.xml 
'  - picture 图片，可以嵌套
'    - name : 资源名称，程序中的引用标识
'    - src : 指向图片资源。如果没有就默认指向上一级的图片资源，如果上一级没有就往更上一级，指导顶级picture；若顶级都没有，就为空资源。
'    - x : 定义截图位置x,默认0
'    - y : 定义截图位置y,默认0
'    - width : 定义宽，默认有指向图片资源就是图片宽 - x，没有就是上一级图片宽 - x。
'    - height : 定义高，同上
'  - text 文本 - pass 以后有空在做吧
'    - gren : 分组 
' 关于那个编译的办法，建议还是将模块分开吧。
' game - 控制器主体
' assets - 资源管理器
' object - 游戏对象
' audio - 音频 - pass
' animation - 动画效果
' 新增语法
'  require

Option Explicit

Const PATH_SRC = "src/"
Const PATH_BUILD = "/"

Const PATH_VBSEDIT = "D:\Vbsedit\vbsedit.exe"

Const ANNOTATION = True

class CCode
  Public m_fso
  Public m_dic
  Public m_path
  Public m_script
  
  Public Sub exec(byval file)
    file = m_path & file
    ExecuteGlobal m_fso.OpenTextFile(file,1).ReadAll()
  End Sub
  
  Private Function loadFile(ByVal file)
    '空文件不载入
    If m_fso.GetFile(file).size Then
      file = Replace(file,"\","/")
      Select Case LCase(Right(file,Len(file) - InStrRev(file,".")))
        Case "vbs" ' - 载入脚本
          If Not m_dic.Exists(file) Then
            '插入文件标记
            m_script = m_script & _
              "' ==============================" & vbNewLine & _
              "'  From " & Right(file,Len(file) - InStrRev(file,"/")) & vbNewLine & _
              "' ==============================" & vbNewLine
              
            Dim script,line
            Set script = m_fso.OpenTextFile(file,1)

            Do While script.AtEndOfStream = False
              line = script.Readline
              
              '简单实现去掉注释
              If ANNOTATION Then
                Dim Ms: Ms = InStr(line,"'")
                If Ms = 1 Or LCase(Left(line,3)) = "rem" Then 
                  line = ""
                ElseIf Ms > 1 Then
                  line = Left(line,Ms - 1)
                End If
                'pass - 实现避免去掉字符串中的'
              End If
              
              '去掉仅包含空格的一行
              If Len(Trim(line)) = 0 Then line = ""
              
              '去掉多余空行
              If Len(line) Then line = RTrim(line) & vbNewLine

              m_script = m_script & line
            Loop
            
            script.Close
            m_dic.Add file, True '登记已载入的脚本
          End If
        Case "xml" ' - 载入数据
          m_script = "'#file:xml;"
          Dim xml: Set xml = m_fso.OpenTextFile(file,1)
          m_script = m_script + xml.readall() + vbNewLine
      End Select
    End If
  End Function
  
  Public Function load(ByVal path)
    path = m_path & path
    '文件
    If m_fso.FileExists(path) Then
      Call loadFile(path)
    End If
    '文件夹
    If m_fso.FolderExists(path) Then
      Dim file: For Each file In m_fso.GetFolder(path).files
        Call loadFile(file.path)
      Next
    End If
  End Function
  
  Public Function Build(ByVal path)
    '默认情况
    If Right(path,1) = "/" Then 
      path = m_path & path
      If Len(GAME_TITLE) Then 
        path = path & GAME_TITLE & ".vbs"
      Else
        path = path & "NewProject.vbs"
      End If
    End If
    '写文件
    With m_fso.CreateTextFile(path,True)
      .Write m_script
      .close
    End With
    Build = Replace(path,"/","\")
  End Function
  
  private sub Class_Initialize
    Set m_dic = CreateObject("Scripting.Dictionary")
    Set m_fso = CreateObject("Scripting.FileSystemObject")
    m_path = Replace(m_fso.GetFile(Wscript.ScriptFullName).ParentFolder.Path,"\","/")
    If Right(m_path,1) <> "/" Then m_path = m_path & "/"
  end sub

  private sub Class_Terminate
    Set m_dic = Nothing
    Set m_fso = Nothing    
  end Sub
  
end Class


Dim Code: Set Code = New CCode
With Code
  '获得配置信息
  .exec PATH_SRC & "head.vbs" 
  .load "assets/assets.xml"
  .load PATH_SRC & "head.vbs"
  .load PATH_SRC & "app.vbs"
  .load PATH_SRC & "linkedList.vbs"
  .load PATH_SRC & "clock.vbs"
  .load PATH_SRC & "object.vbs"
  .load PATH_SRC & "assets.vbs"
  .load PATH_SRC & "event.vbs"
  .load PATH_SRC & "display.vbs"
  .load PATH_SRC & "game.vbs"
  .load PATH_SRC & "main.vbs"
  
  '生成脚本, 并打开文件
  CreateObject("WScript.Shell").Run PATH_VBSEDIT & " """ & .Build(PATH_BUILD) & """"
End With
