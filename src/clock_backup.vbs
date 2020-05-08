Class CClock
  Private mTasklist

  Public Function setInterval(ByVal obj, ByVal funcname, ByVal stime)
    With Add(Obj, funcname)
      .Mode = 1
      .StartTime = stime
      setInterval = .Index
    End With
  End Function

  Public Function setTimeout(ByVal obj, ByVal funcname, ByVal stime)
    With Add(Obj, funcname)
      .Mode = 2
      .StartTime = stime
      setTimeout = .Index
    End With
  End Function

  Public Function setUpdate(ByVal obj, ByVal funcname)
    With Add(Obj, funcname)
      .Mode = 3
      .StartTime = 0
      setUpdate = .Index
    End With
  End Function

  Public Function setTime(ByVal obj, ByVal funcname, ByVal starttime,ByVal endtime)
    With Add(Obj, funcname)
      .Mode = 4
      .StartTime = starttime
      .EndTime = endtime
      setTime = .Index
    End With
  End Function

  Private Function Add(ByVal obj, ByVal funcname)
    Dim task: Set task = New CClockTask
    With task
      Set .Obj = Obj
      Set .Func = GetRef(funcname)
      Set .Parent = Me
      .Index = mTasklist.Add(task)
    End With
    
    Set Add = task
  End Function

  Public Sub clearClock(ByVal index)
    mTasklist.Remove index
  End Sub

  Public Sub Update(ByVal spendtime)
    If mTasklist.Count = 0 Then Exit Sub
    mTasklist.StartTraversal
    Do
      mTasklist.Prev.Update spendtime
    Loop Until mTasklist.EOF
  End Sub

  Private Sub Class_Initialize()
    Set mTasklist = New ClinkedList
  End Sub

  Private Sub Class_Terminate()
    If isObject(mTasklist) Then Set mTasklist = Nothing
  End Sub
End Class

Class CClockTask
  Public Parent, Index, Mode, Obj, Func
  Private mStartTime, mEndTime
  Private msRest, meRest

  Public Property Let EndTime(ByVal value)
    If value < 0 Then 
      Err.Number = vbObjectError + 502
      Err.Description = "定时器时间不能为负数"
      Exit Property
    End If
    mEndTime = value
    meRest = mEndTime
  End Property

  Public Property Let StartTime(ByVal value)
    If value < 0 Then 
      Err.Number = vbObjectError + 501
      Err.Description = "定时器时间不能为负数"
      Exit Property
    End If
    mStartTime = value
    msRest = mStartTime
  End Property

  Public Sub Update(spendtime)
    msRest = msRest - spendtime
    If msRest > 0 Then Exit Sub

    Select Case Mode
      Case 1 ' setInterval
        Func Obj
        msRest = mStartTime
      Case 2 ' setTimeout
        Func Obj
        Parent.clearClock Index
      Case 3 ' setUpdate
        Func Obj, spendtime
      Case 4 ' setTime
        meRest = meRest - spendtime
        If meRest < 0 Then 
          Parent.clearClock Index
          Exit Sub
        End If
        Func Obj, meRest / mEndTime
    End Select
  End Sub

  Private Sub Class_Terminate()
    If isObject(Parent) Then Set Parent = Nothing
    If isObject(Obj) Then Set Obj = Nothing
    If isObject(Func) Then Set Func = Nothing
  End Sub
End Class

' 简化接口
Public Function setInterval(ByVal obj, ByVal funcname, ByVal stime)
  setInterval = Clock.setInterval(obj, funcname, stime)
End Function

Public Function setTimeout(ByVal obj, ByVal funcname, ByVal stime)
  setTimeout = Clock.setTimeout(obj, funcname, stime)
End Function

Public Function setUpdate(ByVal obj, ByVal funcname)
  setUpdate = Clock.setUpdate(obj, funcname)
End Function

Public Function setTime(ByVal obj, ByVal funcname, ByVal starttime,ByVal endtime)
  setTime = Clock.setTime(obj, funcname, starttime, endtime)
End Function

Public Sub clearClock(ByVal index)
  Clock.clearClock index
End Sub