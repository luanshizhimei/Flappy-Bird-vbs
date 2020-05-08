Private Sub main()
  Set Game = new CGame
  Game.Start "GameStart"

  Dim Interval , LastTime , SpendTime
  If GAME_FPS Then Interval = Int(1000 / GAME_FPS)

  Do While Game.State
    ' ����������˳�
    If Err.Number Then
      MsgBox "Err code: " & Err.Number & vbNewLine & "Description: " & Err.Description, _ 
        vbCritical + vbSystemModal, GAME_TITLE
      Call Game.Quit()
    End If

    ' ��Ϸѭ��
    LastTime = Timer
    Game.UpDate(SpendTime)
    SpendTime = Int((Timer-LastTime)*1000)

    ' ��֡
    If SpendTime < Interval Then 
      WScript.Sleep Interval - SpendTime
      SpendTime = Interval
    ElseIf SpendTime = 0 Then
      SpendTime = 1
    End If

    ' ֡������(������)
    Debug.WriteLine "Fps: ",Fix(1000 / SpendTime)
  loop

  Set Game = Nothing
End Sub

Call main()