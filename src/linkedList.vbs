Class ClinkedList
  Private mMem(), mSize, mUsedCount
  Private mQueue, mPointer, mTop
  Private mcThreshold_One
  Private mcThreshold_Two
  Private mcExpandInit
  Private mcExpendBigPer
  Private mcExpandMaxPort

  Public Property Get Count()
    Count = mUsedCount
  End Property

  Public Property Get EOF()
    If mPointer = 0 Then 
      EOF = True: Exit Property
    End If
    EOF = False
  End Property

  Public Property Get Pointer()
    Pointer = mPointer
  End Property

  ' 初始化指针
  Public Sub StartTraversal()
    mPointer = mTop
  End Sub

  Public Function Prev()
    If mPointer = 0 Then
      Call StartTraversal
    End If
    Set Prev = mMem(0, mPointer)
    mPointer = mMem(1, mPointer)
  End Function

  Private Sub Expand()
    Dim lngCount, oldCount
    oldCount = mSize
    ' 扩容
    If mUsedCount < mcThreshold_One Then
      lngCount = mSize * 2
    Elseif mUsedCount < mcThreshold_Two Then
      lngCount = mSize * 3 / 2
    Else
      lngCount = mSize + mcExpendBigPer
    End If
    mSize = lngCount
    ReDim Preserve mMem(2,mSize)
    ' 登记可用空间
    Dim I: For I = oldCount + 1 to mSize
      mQueue.EnQueue I
    Next
  End Sub

  Private Function GetIndex()
    If mUsedCount > mSize * mcExpandMaxPort Then
      Call Expand
    End If
    GetIndex = mQueue.DeQueue
  End Function

  Public Function Add(ByVal Obj)
    mUsedCount = mUsedCount + 1
    Dim idx: idx = GetIndex()
    mMem(2,mTop) = idx ' Next
    mMem(1,idx) = mTop ' Prev
    Set mMem(0,idx) = Obj ' value
    mTop = idx
    Add = mTop ' return index
  End Function

  Public Sub Remove(ByVal idx)
    If Idx <= 0 Or Idx > mSize Then
      Err.Number = vbObjectError + 11
      Err.Description = "索引超出方位：" & idx
      Exit Sub
    End If
    ' 交换
    Dim up, np
    up = Int(mMem(1,idx))
    np = Int(mMem(2,idx))
    If up Then mMem(2,up) = np
    If np Then mMem(1,np) = up
    ' 更新顶部
    If idx = mTop Then mTop = up
    ' 清空
    mMem(1,idx) = 0
    mMem(2,idx) = 0
    Set mMem(0,idx) = Nothing
    mUsedCount = mUsedCount - 1
    ' 登记可用空间
    mQueue.EnQueue idx
  End Sub

  Private Sub Class_Initialize()
    mcThreshold_One = 10000
    mcThreshold_Two = 10000000
    mcExpendBigPer = 10000000
    mcExpandInit = 10
    mcExpandMaxPort = 0.8
    mSize = mcExpandInit
    ReDim mMem(2,mSize)
    Set mQueue = CreateObject("System.Collections.Queue")
    Dim I: For I = 1 to mSize
      mQueue.EnQueue I
    Next
  End Sub
  Private sub Class_Terminate()
    Erase mMem
    If IsObject(mQueue) Then Set mQueue = Nothing
  end sub
End Class


' Class test
' Public id
' Public Function create(num)
' id = num
' Set create = Me
' End Function
' End Class

' Set linklist = New ClinkedList

' For i = 0 To 20
'   Call linklist.Add((New test).create(i))
'   If Random(0,1) Then linklist.Remove(i)
' Next

' Debug.WriteLine "写入长度: " & linklist.Count
' i = 0
' Do
'   i = i + 1
'   Debug.WriteLine i & ", " & linklist.Pointer & ": " & linklist.Prev.id
' Loop Until linklist.EOF

' Public Function Random(ByVal n,ByVal m)
'   Randomize(): Random =  Int((m - n + 1) * Rnd + n)
' End Function