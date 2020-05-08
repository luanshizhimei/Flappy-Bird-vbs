Dim Speed, GameState
Dim Body, Bird, Pipe_one, Pipe_two, score, score_panel
Dim PassState

Const JUMPHEIGHT = 80
Const PIPESPACE = 100
Const EASYSPACE = 12

Const MEDALSCORE_ONE = 100
Const MEDALSCORE_TWO = 500
Const MEDALSCORE_THREE = 999

' 初始化
Public Sub GameStart()
  GameState = 0
  PassState = False
  Speed = 72 ' px / 1s(1000ms)
  Assets.load Assets.ScriptName
  Dim bg, h: h = Hour(Now)
  If h >= 8 And h < 20 Then 
    bg = "bg_day"
  else
    bg = "bg_night"
  End If
  Dim playbutton, scorebutton
  Set Body = Object("Body")
  Body.Image = File(bg)
  SetClick Body, "Body_Click"

  ' 水管组
  Set Pipe_one = (New CPipe).Create("Pipe_one")
  Pipe_one.Visual = False
  Set Pipe_two = (New CPipe).Create("Pipe_two")
  Pipe_two.Visual = False

  Set Pipe_one.Another = Pipe_two
  Set Pipe_two.Another = Pipe_one

  ' 地面
  Dim land: Set land = Body.Add("land")
  With land
    .Image = File("land")
    .X = 0
    .Y = DISPLAY_HEIGHT - land.Image.height
  End With
  setUpdate land, "landAnimation"
  Set land = Nothing

  ' 小鸟
  Set Bird = (New CBird).Create(Body.Add("bird"))
  bird.x = 120: bird.y = 200: Bird.Dire = 1
  setInterval bird, "birdFlyAnimation", 100
  Bird.SetAnimation setUpdate(bird, "birdStartAnimation")

  ' 标题
  With Body.Add("title")
    .Image = File("title")
    .x = 55: .y = 142
  End With

  ' 开始按键
  Set playbutton = Body.Add("playbutton")
  With playbutton
    .Image = File("button_play")
    .x = 20: .y = 340
  End With
  SetDefaultButton playbutton
  SetClick playbutton, "playbutton_Click"

  ' 排名按键
  Set scorebutton = Body.Add("scorebutton")
  With scorebutton
    .Image = File("button_score")
    .x = 154: .y = 340
  End With
  SetDefaultButton scorebutton
  SetClick scorebutton, "scorebutton_Click"

  ' rate按钮
  With Body.Add("ratebutton")
    .Image = File("button_rate")
    .x = 107: .y = 270
  End With
  SetDefaultButton Object("ratebutton")

  ' 分数
  Set Score = (New CNumber).Create(Body, "Score", "font_")
  Score.value = 0
  Score.y = 80
  Score.x = Fix((DISPLAY_WIDTH - Score.width) / 2)
  Score.Margin = 5
  Score.Visual = False

  ' 得分牌
  Set score_panel = (New CScore_panel).Create("score_panel")
  score_panel.y = 512
  score_panel.Visual = False

  ' 提示
  With Body.Add("tutorial")
    .Image = File("tutorial")
    .x = 87: .y = 240
    .Visual = False
  End With

  ' 遮罩（黑）
  With Body.Add("Shade")
    .Image = File("black")
    .x = -10: .y = -10
    .Width = DISPLAY_WIDTH + 20
    .Height = DISPLAY_HEIGHT + 20
    .Alpha = 0
  End With
End Sub

' 事件
Public Sub SetDefaultButton(ByVal node)
  SetMousedown node, "button_mousedown"
  SetMouseup node, "button_mouseup"
End Sub

Public Sub button_mousedown(ByVal node)
  Node.x = Node.x + 2
  Node.y = Node.y + 2
End Sub

Public Sub button_mouseup(ByVal node)
  Node.x = Node.x - 2
  Node.y = Node.y - 2
End Sub

Public Sub Body_Click(ByVal node)
  Select Case GameState
    Case 1
      ' 淡出动画
      Dim title, tutorial
      Set title = Object("title")
      Set tutorial = Object("tutorial")
      setTime title, "FadeOutAnimation", 0 , 500
      setTime tutorial, "FadeOutAnimation", 0 , 500
      setTimeout title, "Hide", 500
      setTimeout tutorial, "Hide", 500
      ' 设置
      Pipe_two.x = 178
      Pipe_one.Update
      Pipe_one.Visual = True
      Pipe_two.Update
      Pipe_two.Visual = True
      Score.Visual = True
      ' 设置逻辑动画
      Bird.ClearAnimation
      Bird.SetAnimation setUpdate(bird, "birdJumpAnimation")
      Pipe_one.SetUpdate setUpdate(Pipe_one, "PipeUpdate")
      Pipe_two.SetUpdate setUpdate(Pipe_two, "PipeUpdate")
      bird.Jump
      GameState = 2
    Case 2
      bird.Jump
  End Select
End Sub

Public Sub playbutton_Click(ByVal node)
  node.Usable = False
  ' 遮罩动画
  Dim Shade: Set Shade = Object("Shade")
  Shade.Visual = True
  Shade.Alpha = 0
  setTime Shade, "FadeInAnimation", 0 , 500
  setTime Shade, "FadeOutAnimation", 500 , 500
  setTimeout Shade, "clearView", 1000
  ' 设置
  setTimeout Object("bird"), "toWelcomeView", 500
End Sub

Public Sub toWelcomeView(ByVal node)
  With Object("title")
    .X = 46: .Y = 138
    .Image = File("text_ready")
  End With 
  Dim tutorial
  Set tutorial = Object("tutorial")
  tutorial.Alpha = 1
  tutorial.Visual = True
  Score.value = 0
  Bird.X = 50
  bird.y = 200
  Bird.angle = 0
  Bird.ClearAnimation
  Bird.SetAnimation setUpdate(bird, "birdStartAnimation")
  PassState = False

  Hide Object("playbutton")
  Hide Object("scorebutton")
  Hide Object("ratebutton")
  Hide Object("score_panel")
  Hide Object("Pipe_one")
  Hide Object("Pipe_two")
End Sub

Public Sub clearView(ByVal node)
  Node.Visual = False
  GameState = 1
End Sub

Public Sub Hide(ByVal Node)
  Node.Visual = False
End Sub

Public Sub Show(ByVal Node)
  Node.Visual = True
End Sub

Public Sub scorebutton_Click(ByVal node)
  debug.writeline "原版那个要游戏中心才可以使用，就不想做了"
End Sub

' 动画

Public Sub FadeInAnimation(ByVal node, ByVal rate)
  Dim Alpha: Alpha = 1 - round(rate, 5)
  If node.Alpha = 100 Then Alpha = 0
  node.Alpha = Alpha
End Sub

Public Sub FadeOutAnimation(ByVal node, ByVal rate)
  Dim Alpha: Alpha = round(rate, 5)
  If node.Alpha = 0 Then Alpha = 100
  node.Alpha = Alpha
End Sub

Public Sub landAnimation(ByVal Obj, ByVal spendtime)
  Dim x: x = Obj.x
  ' 重置状态
  If x = -48 Then Obj.x = 0: Exit Sub
  ' 计算下一个位置
  Dim dx: dx = Ceil(Speed * spendtime / 1000)
  If x - dx < -48 Then dx = 48 + x
  Obj.x = x - dx
End Sub

Public Sub birdFlyAnimation(ByVal Obj)
  Dim idx: idx = Obj.Index
  If idx < 2 Then 
    Obj.Index = idx + 1
  Else
    Obj.Index = 0
  End If
End Sub

Public Sub birdStartAnimation(ByVal Obj, ByVal spendtime)
  If Bird.Dire = 1 And Bird.y >= 210 Then
    Bird.Dire = -1
  ElseIf Bird.Dire = -1 And Bird.y <= 200 Then
    Bird.Dire = 1
  End If
  Dim dy: dy = Ceil(Speed * spendtime / 3000)
  Bird.y = Bird.Dire * dy + Bird.y
End Sub

Public Sub birdJumpAnimation(ByVal Obj, ByVal spendtime)
  Dim dy, da
  dy = Ceil(Speed * spendtime / 400) * Bird.Dire
  da = dy * 2 + Bird.angle
  dy = dy + Bird.y
  ' 下降
  If Bird.Dire = -1 And dy <= Bird.Dsy Then 
    Bird.y = Bird.Dsy
    Bird.Dsy = 512
    Bird.Dire = 1
  Else
    Bird.y = dy
  End If
  ' 抬头
  If Bird.Dire = -1 And da >= -45 Then 
    Bird.angle = da
  ElseIf Bird.Dire = 1 And da <= 90 Then
    Bird.angle = da
  End If
End Sub

Public Sub scorePanelAnimation(ByVal Obj, ByVal spendtime)
  If Obj.y <= 200 Then
    Obj.y = 200: Obj.ClearAnimation
    Exit Sub
  End If
  Dim dy: dy = Ceil(Speed * spendtime / 100) 
  Obj.y = Obj.y - dy
End Sub

Public Sub PipeUpdate(ByVal Obj, ByVal spendtime)
  ' 碰撞检测
  If PassState = False And Obj.x >= 0 And Obj.x <= 50 Then
    Dim y: y = bird.y
    If y + EASYSPACE > Obj.SpaceY And _ 
      y + bird.height - EASYSPACE < Obj.SpaceY + PIPESPACE Then
      Score.value = Score.value + 1
      Score.x = Fix((DISPLAY_WIDTH - Score.width) / 2)
    Else
      GameState = 3
      ' 清除动画
      Bird.ClearAnimation
      Pipe_one.ClearUpdate
      Pipe_two.ClearUpdate

      ' 设置
      score_panel.y = 512
      score_panel.Visual = True
      Dim title, playbutton, scorebutton
      Set title = Object("title")
      Set playbutton = Object("playbutton")
      Set scorebutton = Object("scorebutton")
      title.Image = File("text_game_over")
      title.Visual = True
      playbutton.Visual = True
      scorebutton.Visual = True
      score_panel.Show(Score.value)
      ' 动画
      score_panel.SetAnimation setUpdate(score_panel, "scorePanelAnimation")
      setTime title, "FadeInAnimation", 0 , 500
      setTime playbutton, "FadeInAnimation", 0 , 500
      setTime scorebutton, "FadeInAnimation", 0 , 500
    End If
    PassState = True
  End If
  ' 移动水管
  If Obj.x <= -52 Then
    PassState = False
    Obj.Update: Exit Sub
  End If
  Dim dx: dx = Ceil(Speed * spendtime / 1000)
  Obj.x = Obj.x - dx
End Sub

' 单例（要是支持继承多态就好了

Class CBird
  Private mBird, mStyle, mIndex, mheight
  Public Dire, Dsy
  private mAnimation, mAnimationIndex

  Public Property Get Height()
    Height = mheight
  End Property

  Public Sub Jump()
    Dsy = mBird.y - JUMPHEIGHT
    Dire = -1
  End Sub

  Public Sub ClearAnimation()
    If mAnimationIndex = 0 Then Exit Sub 
    clearClock mAnimationIndex
    mAnimationIndex = 0
  End Sub

  Public Sub SetAnimation(ByVal idx)
    mAnimationIndex = idx
  End Sub

  Public Property Get angle()
    angle = mBird.angle
  End Property

  Public Property Let angle(ByVal value)
    mBird.angle = value
  End Property

  Public Property Get Y()
    y = mBird.y
  End Property

  Public Property Let Y(ByVal value)
    mBird.y = value
  End Property

  Public Property Get X()
    x = mBird.x
  End Property

  Public Property Let X(ByVal value)
    mBird.x = value
  End Property

  Public Property Get Index()
    Index = mIndex
  End Property

  Public Property Let Index(ByVal value)
    mIndex = value
    mBird.Image = File("bird" & mStyle & "_" & mIndex)
  End Property

  Public Sub Change()
    mStyle = Random(0, 2)
  End Sub

  Public Function Create(ByVal parent)
    Set mBird = parent
    mIndex = 0
    mheight = 48'mBird.height
    Change()
    Set Create = Me
  End Function
End Class

Class CPipe
  Private mPipe, mPipe_up, mPipe_down
  Private mSpaceY, mUpdateIndex
  Public Another

  Public Sub ClearUpdate()
    If mUpdateIndex = 0 Then Exit Sub 
    clearClock mUpdateIndex
    mUpdateIndex = 0
  End Sub

  Public Sub SetUpdate(ByVal idx)
    mUpdateIndex = idx
  End Sub

  Public Property Get Visual()
    Visual = mPipe.Visual
  End Property

  Public Property Let Visual(ByVal value)
    mPipe.Visual = value
  End Property

  Public Property Get SpaceY()
    SpaceY = mSpaceY
  End Property

  Public Property Get x()
    x = mPipe.x
  End Property

  Public Property Let x(ByVal value)
    mPipe.x = value
  End Property

  Public Sub Update()
    mSpaceY = Random(40, 280)
    mPipe_down.y = mSpaceY - 320
    mPipe_up.y = mSpaceY + PIPESPACE
    mPipe.x = Another.x + 200
  End Sub

  Public Function Create(ByVal Name)
    Set mPipe = Object("Body").Add(Name)
    Set mPipe_up = mPipe.Add(Name & "_up")
    Set mPipe_down = mPipe.Add(Name & "_down")
    Set Create = Me
    mPipe_up.Image = File("pipe_up")
    mPipe_down.Image = File("pipe_down")
    mUpdateIndex = 0
  End Function
End Class

Class CNumber
  Private mNumber, mName
  Private mValue
  Private mBitlist(), mBitCount
  Private mStyle, mWidth
  Private mX, mY
  Public Margin

  Public Property Get Visual()
    Visual = mNumber.Visual
  End Property

  Public Property let Visual(ByVal value)
    mNumber.Visual = value
  End Property

  Public Property Get Width()
    Width = mWidth
  End Property

  Public Property Get Y()
    Y = mNumber.y
  End Property

  Public Property let Y(ByVal num)
    mNumber.y = num
  End Property

  Public Property Get X()
    X = mNumber.X
  End Property
    
  Public Property let X(ByVal num)
    mNumber.X = num
  End Property
  
  Public Property Get value()
    value = mValue
  End Property

  Public Property let value(ByVal num)
    mValue = num
    Dim mBit, I, mOutdigit, bit
    If mValue Then mBit = Ceil(Log(mValue) / log(10))
    If mValue Mod 10 = 0 Then mBit = mBit + 1
    ' 新增
    If mBit > mBitCount Then
      ReDim Preserve mBitlist(mBit)
      For I = mBitCount + 1 To mBit
        Set mBitlist(I) = mNumber.Add(mName & "_bit" & I)
      Next
      mBitCount = mBit
    End If
    ' 显示
    mWidth = 0
    mOutdigit = 0
    bit = 0
    For I = mBitCount - 1 to 0 Step -1
      mOutdigit = mOutdigit * 10
      bit = Fix(mValue / 10 ^ I) - mOutdigit
      mOutdigit = mOutdigit + bit
      Show I, bit
    Next
  End Property

  Private Sub Show(ByVal bit, ByVal num)
    bit = bit + 1
    If mWidth = 0 And num = 0 And mBitCount > 1 Then
      mBitlist(bit).Visual = False
      Exit Sub
    End If
    mWidth = mWidth + Margin
    mBitlist(bit).X = mWidth
    mBitlist(bit).Image = File(mStyle & num)
    mWidth = mWidth + mBitlist(bit).Width
  End Sub

  Public Function Create(ByVal parent, ByVal name,ByVal Style)
    Margin = 0
    Set mNumber = parent.Add(name)
    Set Create = Me
    mName = name
    mBitCount = 0
    mStyle = Style
  End Function
End Class

Class CScore_panel
  Private mPanel, mMedal, mNew
  Private mScoreNum, mBestNum, mBestScore
  Private mAnimationIndex

  Public Sub ClearAnimation()
    If mAnimationIndex = 0 Then Exit Sub 
    clearClock mAnimationIndex
    mAnimationIndex = 0
  End Sub

  Public Sub SetAnimation(ByVal idx)
    mAnimationIndex = idx
  End Sub

  Public Property Get y()
    y = mPanel.y
  End Property

  Public Property Let y(ByVal value)
    mPanel.y = value
  End Property

  Public Property Get Visual()
    Visual = mPanel.Visual
  End Property

  Public Property Let Visual(ByVal value)
    mPanel.Visual = value
  End Property

  Public Sub Show(ByVal value)
    mScoreNum.value = value
    mScoreNum.x = 210 - mScoreNum.width
    mBestNum.value = mBestScore
    mBestNum.x = 210 - mBestNum.width
    If value >= MEDALSCORE_THREE Then
      mMedal.Image = File("medals_3")
    ElseIf value >= MEDALSCORE_TWO Then
      mMedal.Image = File("medals_2")
    ElseIf value >= MEDALSCORE_ONE Then
      mMedal.Image = File("medals_1")
    ElseIf value < MEDALSCORE_ONE Then
      mMedal.Image = File("medals_0")
    End If
    mNew.Visual = False
    ' 显示新纪录
    If value > mBestScore Then
      mBestScore = value
      mNew.x = mScoreNum.x - 35
      mNew.y = 37
      mNew.Visual = True
    End If
  End Sub

  Public Function Create(ByVal Name)
    mBestScore = 0
    Set mPanel = Body.Add(name)
    mPanel.Image = File("score_panel")
    mPanel.x = Fix((DISPLAY_WIDTH - mPanel.width) / 2)
    Set mMedal = mPanel.Add("Medal")
    mMedal.Image = File("medals_0")
    mMedal.x = 32: mMedal.y = 45
    Set mScoreNum = (New CNumber).Create(mPanel, "scoreNum", "number_score_")
    mScoreNum.value = 0
    mScoreNum.x = 210 - mScoreNum.width
    mScoreNum.y = 35
    mScoreNum.Margin = 2
    Set mBestNum = (New CNumber).Create(mPanel, "BestNum", "number_score_")
    mBestNum.value = 0
    mBestNum.x = 210 - mBestNum.width
    mBestNum.y = 78
    mBestNum.Margin = 2
    Set mNew = mPanel.Add("New")
    mNew.Image = File("new")
    mNew.Visual = False
    Set Create = Me
  End Function
End Class

' 一些常用的函数
Public Function Ceil(ByVal num)
  Ceil = -int(-num)
End Function

Public Function Random(ByVal n,ByVal m)
  Randomize(): Random =  Int((m - n + 1) * Rnd + n)
End Function