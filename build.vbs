'�ϳɽű�

'pass - �����������������
'pass - ʵ��ͼƬ��Դ��Ԥ����
' ���ǻ��ǲ�������ű���ʵ��������ܣ����ǽ���Դassets.xml����������Ȼ��ͨ������ű�����Դxml����ֱ�ӱ��뵽�����Ľű��ļ��С�
' �����Ǹ�assets.xml�����Ҫ����ָ������������Դ���磺ͼƬ��ͼƬ��Ƭ���ı�֮��ġ�Ȼ����ͨ������ϳɽű�ֱ���������Ʒ�ű��ĳ����С�(֮������һ��Ҫ������������Ƕ�ļ�����ʽ��)
' assets.xml 
'  - picture ͼƬ������Ƕ��
'    - name : ��Դ���ƣ������е����ñ�ʶ
'    - src : ָ��ͼƬ��Դ�����û�о�Ĭ��ָ����һ����ͼƬ��Դ�������һ��û�о�������һ����ָ������picture����������û�У���Ϊ����Դ��
'    - x : �����ͼλ��x,Ĭ��0
'    - y : �����ͼλ��y,Ĭ��0
'    - width : �����Ĭ����ָ��ͼƬ��Դ����ͼƬ�� - x��û�о�����һ��ͼƬ�� - x��
'    - height : ����ߣ�ͬ��
'  - text �ı� - pass �Ժ��п�������
'    - gren : ���� 
' �����Ǹ�����İ취�����黹�ǽ�ģ��ֿ��ɡ�
' game - ����������
' assets - ��Դ������
' object - ��Ϸ����
' audio - ��Ƶ - pass
' animation - ����Ч��
' �����﷨
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
    '���ļ�������
    If m_fso.GetFile(file).size Then
      file = Replace(file,"\","/")
      Select Case LCase(Right(file,Len(file) - InStrRev(file,".")))
        Case "vbs" ' - ����ű�
          If Not m_dic.Exists(file) Then
            '�����ļ����
            m_script = m_script & _
              "' ==============================" & vbNewLine & _
              "'  From " & Right(file,Len(file) - InStrRev(file,"/")) & vbNewLine & _
              "' ==============================" & vbNewLine
              
            Dim script,line
            Set script = m_fso.OpenTextFile(file,1)

            Do While script.AtEndOfStream = False
              line = script.Readline
              
              '��ʵ��ȥ��ע��
              If ANNOTATION Then
                Dim Ms: Ms = InStr(line,"'")
                If Ms = 1 Or LCase(Left(line,3)) = "rem" Then 
                  line = ""
                ElseIf Ms > 1 Then
                  line = Left(line,Ms - 1)
                End If
                'pass - ʵ�ֱ���ȥ���ַ����е�'
              End If
              
              'ȥ���������ո��һ��
              If Len(Trim(line)) = 0 Then line = ""
              
              'ȥ���������
              If Len(line) Then line = RTrim(line) & vbNewLine

              m_script = m_script & line
            Loop
            
            script.Close
            m_dic.Add file, True '�Ǽ�������Ľű�
          End If
        Case "xml" ' - ��������
          m_script = "'#file:xml;"
          Dim xml: Set xml = m_fso.OpenTextFile(file,1)
          m_script = m_script + xml.readall() + vbNewLine
      End Select
    End If
  End Function
  
  Public Function load(ByVal path)
    path = m_path & path
    '�ļ�
    If m_fso.FileExists(path) Then
      Call loadFile(path)
    End If
    '�ļ���
    If m_fso.FolderExists(path) Then
      Dim file: For Each file In m_fso.GetFolder(path).files
        Call loadFile(file.path)
      Next
    End If
  End Function
  
  Public Function Build(ByVal path)
    'Ĭ�����
    If Right(path,1) = "/" Then 
      path = m_path & path
      If Len(GAME_TITLE) Then 
        path = path & GAME_TITLE & ".vbs"
      Else
        path = path & "NewProject.vbs"
      End If
    End If
    'д�ļ�
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
  '���������Ϣ
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
  
  '���ɽű�, �����ļ�
  CreateObject("WScript.Shell").Run PATH_VBSEDIT & " """ & .Build(PATH_BUILD) & """"
End With
