Private Sub main()
  Set Game = new CGame
  Game.Start "GameStart"

  Dim Interval , LastTime , SpendTime
  If GAME_FPS Then Interval = Int(1000 / GAME_FPS)

  Do While Game.State
    ' 遇到错误就退出
    If Err.Number Then
      MsgBox "Err code: " & Err.Number & vbNewLine & "Description: " & Err.Description, _ 
        vbCritical + vbSystemModal, GAME_TITLE
      Call Game.Quit()
    End If

    ' 游戏循环
    LastTime = Timer
    Game.UpDate(SpendTime)
    SpendTime = Int((Timer-LastTime)*1000)

    ' 锁帧
    If SpendTime < Interval Then 
      WScript.Sleep Interval - SpendTime
      SpendTime = Interval
    ElseIf SpendTime = 0 Then
      SpendTime = 1
    End If

    ' 帧数计算(调试用)
    Debug.WriteLine "Fps: ",Fix(1000 / SpendTime)
  loop

  Set Game = Nothing
End Sub

Call main()