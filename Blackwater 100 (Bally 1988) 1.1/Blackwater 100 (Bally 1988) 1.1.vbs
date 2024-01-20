Option Explicit 
Randomize
'Dim OptionReset
'OptionReset = 1  'Uncomment to reset to default options in case of error OR keep all changes temporary

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="black100"
const UseSolenoids	= 1
const UseLamps		= 0
const UseGI			= 0 								'Only WPC games have special GI circuit.
Const SCoin        = "BW Coin"


LoadVPM "01210000","6803.vbs",3.10

'***********************************************************************************
'****            		 Constants and global variables						****
'***********************************************************************************

'Constants for sound functions
'Const SoundPanning = 1										'1 enables sound panning for VPinball v9.16 and above
'Const TableWidth = 1100                         	        'used for sound panning, enter the value specified in table options "Table Width" 
'Const PanningFactor = 0.7									'panning factor 0..1, 1 = left speaker off, if ball is rightmost and vice versa, 0 = no panning
'Const SoundVel = 10											'sound is played if ball velocity is above this threshold

'********************
'Solenoids Call backs
'********************

SolCallback(1) = "bsRight.SolOut"
SolCallback(2) = "bsTop.SolOut"
SolCallback(3) = "SolGateDown"
SolCallback(4) = "SolGateRaise"
'SolCallback(5) = "SolLeftSling" 
'SolCallback(6) = "SolRightSling" 
SolCallback(8) =  "dtLL.SolDropUp" 'Sol 7
SolCallback(14) = "bsTrough.SolIn"   'Sol 8
SolCallback(12) = "bsTrough.SolOut"  'Sol 9
SolCallback(15) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"   'Sol 10
SolCallback(18) = "PFGI"        'GI 'SOl 11
'SolCallback() = ""           'Flippers Sol 12
'SolCallback() = ""           'Backglass GI Sol13

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFContactors):LeftFlipper.RotateToEnd:FlipperLB.RotateToEnd:FlipperLT.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):LeftFlipper.RotateToStart:FlipperLB.RotateToStart:FlipperLT.RotateToStart
     End If
  End Sub
  
Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFContactors):RightFlipper.RotateToEnd:FlipperRT.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):RightFlipper.RotateToStart:FlipperRT.RotateToStart
     End If
End Sub 



'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

Class DropTarget
  Private m_primary, m_secondary, m_prim, m_sw, m_animate, m_isDropped

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(input): Set m_primary = input: End Property

  Public Property Get Secondary(): Set Secondary = m_secondary: End Property
  Public Property Let Secondary(input): Set m_secondary = input: End Property

  Public Property Get Prim(): Set Prim = m_prim: End Property
  Public Property Let Prim(input): Set m_prim = input: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(input): m_sw = input: End Property

  Public Property Get Animate(): Animate = m_animate: End Property
  Public Property Let Animate(input): m_animate = input: End Property

  Public Property Get IsDropped(): IsDropped = m_isDropped: End Property
  Public Property Let IsDropped(input): m_isDropped = input: End Property

  Public default Function init(primary, secondary, prim, sw, animate, isDropped)
    Set m_primary = primary
    Set m_secondary = secondary
    Set m_prim = prim
    m_sw = sw
    m_animate = animate
    m_isDropped = isDropped

    Set Init = Me
  End Function
End Class

'Start Gate Prim
Dim DropTargetPositions: DropTargetPositions=Array(-50, -40, -30, -20, -10, 0, 10, 10)

Dim DT0, DT1, DT2, DT3, DT4

Set DT0 = (new DropTarget)(Nothing,Gate_DT, Gate_P1,200.1,0,false)
Set DT1 = (new DropTarget)(Nothing,Nothing, Gate_P2,200.1,0,false)
Set DT2 = (new DropTarget)(Nothing,Nothing, Gate_P3,200.1,0,false)
Set DT3 = (new DropTarget)(Nothing,Nothing, Gate_P4,175.2,0,false)
Set DT4 = (new DropTarget)(Nothing,Nothing, Gate_P5,175.2,0,false)


Dim DropTargets: DropTargets=Array(DT0, DT1, DT2, DT3, DT4)

Sub SolGateDown(enabled)
	If enabled then 
    DropTargets(0).animate = -1
    DropTargets(1).animate = -1
    DropTargets(2).animate = -1
    DropTargets(3).animate = -1
    DropTargets(4).animate = -1
'		PlaySound SoundFX("Popper",DOFContactors)
		Gate_DT.isdropped = 1
		Gate_Wall.isdropped = 1

	End If
End Sub

Sub SolGateRaise(enabled)
	If enabled Then 
    DropTargets(0).animate = 1
    DropTargets(1).animate = 1
    DropTargets(2).animate = 1
    DropTargets(3).animate = 1
    DropTargets(4).animate = 1
		PlaySound SoundFX("Popper",DOFContactors)
		Gate_DT.isdropped = 0
		Gate_Wall.isdropped = 0

	End If
End Sub

'Playfield GI
Sub PFGI(Enabled)
	If Enabled Then
		dim xx
		For each xx in GI:xx.State = 0: Next
        PlaySound "fx_relay"
	Else
		For each xx in GI:xx.State = 1: Next
        PlaySound "fx_relay"
	End If
End Sub


'***********************************************************************************
'****               	  	Table Event Handling								****
'***********************************************************************************
Dim bsTrough, bsTop, bsRight, dtLL, SMagnet

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = ""&chr(13)&"You Suck"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
		.hidden = 0
        .Games(cGameName).Settings.Value("sound")=1
		'.PuPHide = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

	PinMAMETimer.Interval = PinMAMEInterval
	PinMAMETimer.Enabled = 1

    vpmNudge.TiltSwitch=50
    vpmNudge.Sensitivity=1
	vpmNudge.TiltObj = Array(LeftSlingshot,RightSlingshot)

	' Trough handler
	Set bsTrough = New cvpmBallStack 
		bsTrough.InitSw 8,42,43,41,0,0,0,0
		bsTrough.InitKick BallRelease,60,8
		bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
		bsTrough.Balls = 3

	'Kicker Top in the Service Manual mapped sw17
	Set bsTop = New cvpmBallStack
		bsTop.InitSaucer sw18,18,160,7
		bsTop.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
		bsTop.KickForceVar = 2

	'Kicker Right in the Service Manual mapped sw18
     Set bsRight = New cvpmBallStack
        bsRight.InitSaucer sw17,17,230,20
		bsRight.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
		bsRight.KickForceVar = 2

 	Set dtLL=New cvpmDropTarget
		dtLL.InitDrop Array(sw25,sw26,sw27,sw28,sw29),Array(25,26,27,28,29)
		dtLL.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

     ' Saucer Magnet (Using low powered Magnet to simulate drop in playfield surface around the saucer -> Idea "stolen" from JP)
    Set SMagnet = New cvpmMagnet
        SMagnet.InitMagnet SaucerMagnet, 3
        SMagnet.GrabCenter = False
        SMagnet.MagnetOn = 1
        SMagnet.CreateEvents "SMagnet"

End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyCode = LeftMagnaSave Then Controller.Switch(5) = 1
	If KeyCode = RightMagnaSave Then Controller.Switch(7) = 1
	If keycode = PlungerKey Then Plunger.Pullback:playsound"plungerpull"
	If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)

	If KeyCode = LeftMagnaSave Then Controller.Switch(5) = 0
	If KeyCode = RightMagnaSave Then Controller.Switch(7) = 0
	If keycode = PlungerKey Then Plunger.Fire:PlaySound"plunger"
	If KeyUpHandler(keycode) Then Exit Sub
End Sub

'***********************************************************************************
'****						  Drains and Kickers           						****
'***********************************************************************************

Sub Drain_Hit:bsTrough.addball me : playsound"BW Drain" : End Sub

'Service manual has sw2 and sw3 mapped sw17 and sw18
Sub sw17_Hit: bsRight.AddBall 0 :PlaySound "BW SaucerHit" : End Sub
Sub sw18_Hit: bsTop.AddBall 0 :PlaySound "BW SaucerHit" : End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 21
    PlaySound SoundFX("right_slingshot",DOFContactors), 0,1, 0.05,0.05 '0,1, AudioPan(RightSlingShot), 0.05,0,0,1,AudioFade(RightSlingShot)
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
'	gi1.State = 0:Gi2.State = 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	vpmTimer.PulseSw 22
    PlaySound SoundFX("left_slingshot",DOFContactors), 0,1, -0.05,0.05 '0,1, AudioPan(LeftSlingShot), 0.05,0,0,1,AudioFade(LeftSlingShot)
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
'	gi3.State = 0:Gi4.State = 0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub


'***********************************************************************************
'****						  Drop Targets and Start gate  						****
'*********************************************************************************** 
'Service Manual sw31,32,33,34,35

 Sub Sw25_Dropped:dtLL.Hit 1 :End Sub  
 Sub Sw26_Dropped:dtLL.Hit 2 :End Sub  
 Sub Sw27_Dropped:dtLL.Hit 3 :End Sub
 Sub Sw28_Dropped:dtLL.Hit 4 :End Sub  
 Sub Sw29_Dropped:dtLL.Hit 5 :End Sub

'***********************************************************************************
'****						      Targets                 						****
'*********************************************************************************** 

 Sub sw38_Hit:vpmTimer.PulseSw 38:PlaySound "BW Target":End Sub
 Sub sw14_Hit:vpmTimer.PulseSw 14:PlaySound "BW Target":End Sub
 Sub sw40_Hit:vpmTimer.PulseSw 40:PlaySound "BW Target":End Sub

 Sub sw30_Hit:vpmTimer.PulseSw 30:PlaySound "BW Target":End Sub
 Sub sw31_Hit:vpmTimer.PulseSw 31:PlaySound "BW Target":End Sub
 Sub sw32_Hit:vpmTimer.PulseSw 32:PlaySound "BW Target":End Sub

 Sub sw1_Hit:vpmTimer.PulseSw 1:PlaySound "BW Target":End Sub
 Sub sw2_Hit:vpmTimer.PulseSw 2:PlaySound "BW Target":End Sub
 Sub sw3_Hit:vpmTimer.PulseSw 3:PlaySound "BW Target":End Sub
 Sub sw4_Hit:vpmTimer.PulseSw 4:PlaySound "BW Target":End Sub

 Sub sw36_Hit:vpmTimer.PulseSw 36:PlaySound "BW Target":End Sub
 Sub sw35_Hit:vpmTimer.PulseSw 35:PlaySound "BW Target":End Sub
 Sub sw34_Hit:vpmTimer.PulseSw 34:PlaySound "BW Target":End Sub
 Sub sw33_Hit:vpmTimer.PulseSw 33:PlaySound "BW Target":End Sub

'***********************************************************************************
'****						   		Gates			      						****
'***********************************************************************************

Sub sw19_Hit:vpmTimer.PulseSw 19:PlaySound "BW Gate":End Sub
Sub sw37_Hit:vpmTimer.PulseSw 37:PlaySound "BW Gate":End Sub
Sub sw12_Hit:vpmTimer.PulseSw 12:PlaySound "BW Gate":End Sub


Sub sw39_Hit:Controller.Switch(39) = 1:End Sub
Sub sw39_UnHit:Controller.Switch(39) = 0:End Sub
Sub sw20_Hit:Controller.Switch(20) = 1::End Sub
Sub sw20_UnHit:Controller.Switch(20) = 0::End Sub
Sub sw44_Hit:Controller.Switch(40) = 1::End Sub
Sub sw44_UnHit:Controller.Switch(40) = 0:End Sub

'***********************************************************************************
'****						   Rollovers and triggers      						****
'*********************************************************************************** 

  Sub sw24_Hit:Controller.Switch(24) = 1:PlaySound "BW Sensor":End Sub
  Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub
  Sub sw13_Hit:Controller.Switch(13) = 1:PlaySound "BW Sensor":End Sub
  Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub
  Sub sw13a_Hit:Controller.Switch(13) = 1:PlaySound "BW Sensor":End Sub
  Sub sw13a_UnHit:Controller.Switch(13) = 0:End Sub
  Sub sw23_Hit:Controller.Switch(23) = 1:PlaySound "BW Sensor":End Sub
  Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub
  Sub sw47_Hit:Controller.Switch(47) = 1:End Sub
  Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub
  Sub sw48_Hit:Controller.Switch(48) = 1:End Sub
  Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub
  Sub sw46_Hit:Controller.Switch(46) = 1:End Sub
  Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub
  Sub sw45_Hit:Controller.Switch(45) = 1:End Sub
  Sub sw45_UnHit:Controller.Switch(45) = 0:End Sub
  Sub sw16_Hit:Controller.Switch(16) = 1:End Sub
  Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub


'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 5 'lamp fading speed
LampTimer.Enabled = 1

' Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step

	   'Special Handling
	   'If chgLamp(ii,0) = 2 Then solTrough chgLamp(ii,1)
	   'If chgLamp(ii,0) = 4 Then PFGI chgLamp(ii,1)

        Next
    End If
    UpdateLamps
End Sub

Sub UpdateLamps
	NFadeL 1, l1
	NFadeL 2, l2
	NFadeL 3, l3
	NFadeL 4, l4
	NFadeL 5, l5
	NFadeL 6, l6
	NFadeL 7, l7
'	NFadeL 8, l8
	NFadeL 9, l9
	NFadeL 10, l10
	NFadeL 11, l11
'	NFadeL 12, l12
'	NFadeL 13, l13
'	NFadeL 14, l14
'	NFadeL 15, l15
'	NFadeL 16, l16
	NFadeL 17, l17
	NFadeL 18, l18
	NFadeL 19, l19
	NFadeL 20, l20
	NFadeL 21, l21
	NFadeL 22, l22 'Backwall
'	NFadeL 23, l23
	NFadeL 24, l24
	NFadeL 25, l25
	NFadeL 26, l26
	NFadeL 27, l27
'	NFadeL 28, l28
'	NFadeL 29, l29
'	NFadeL 30, l30
'	NFadeL 31, l31
'	NFadeL 32, l32
	NFadeL 33, l33
	NFadeL 34, l34 'Backbox Left
	NFadeL 35, l35 'Left Dome
	NFadeL 36, l36 'Plunger 
	NFadeL 37, l37'Backbox Right
	NFadeL 38, l38 'Right Apron Dome
'	NFadeL 39, l39 'Left Apron Dome
	NFadeL 40, l40
	NFadeL 41, l41
	NFadeL 42, l42
'	NFadeL 43, l43 'Backwall
'	NFadeL 44, l44 'Backwall
'	NFadeL 45, l45
'	NFadeL 46, l46
'	NFadeL 47, l47
'	NFadeL 48, l48
    NFadeL 49, l49
    NFadeL 50, l50
'   NFadeL 51, l51
	NFadeL 52, l52
    NFadeL 53, l53
	NFadeL 54, l54
    NFadeL 55, l55
	NFadeL 56, l56
	NFadeL 57, l57
	NFadeL 58, l58
Flashm 53, Flasher53
	NFadeL 59, l59
Flashm 69, Flasher69
'	NFadeL 60, l60
Flashm 85, Flasher85
'	NFadeL 61, l61
'	NFadeL 62, l62
	NFadeL 63, l63
'	NFadeL 64, l64
	NFadeL 65, l65
	NFadeL 66, l66
'	NFadeL 67, l67
	NFadeL 68, l68
	NFadeL 69, l69
	NFadeL 70, l70
	NFadeL 71, l71
	NFadeL 72, l72
	NFadeL 73, l73
	NFadeL 74, l74
	NFadeL 75, l75
'	NFadeL 76, l76
'	NFadeL 77, l77
	NFadeL 78, l78
	NFadeL 79, l79
'	NFadeL 80, l80 'Right Dome
	NFadeL 81, l81
	NFadeL 83, l83
	NFadeL 84, l84
	NFadeL 85, l85
	NFadeL 86, l86
	NFadeL 87, l87
	NFadeL 88, l88
	NFadeL 89, l89
	NFadeL 90, l90
	NFadeL 94, l94
	NFadeL 95, l95



End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.4   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.2 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub AllLampsOff
    Dim x
    For x = 0 to 200
        SetLamp x, 0
    Next
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
        LampState(nr) = abs(value)
        FadingLevel(nr) = abs(value) + 4
    End If
End Sub

' Lights: used for VP10 standard lights, the fading is handled by VP itself

Sub NFadeL(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.state = 0:FadingLevel(nr) = 0
        Case 5:object.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeLm(nr, object) ' used for multiple lights
    Select Case FadingLevel(nr)
        Case 4:object.state = 0
        Case 5:object.state = 1
    End Select
End Sub

'Lights, Ramps & Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.image = a:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
        Case 9:object.image = c:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1         'wait
        Case 13:object.image = d:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
        Case 9:object.image = c
        Case 13:object.image = d
    End Select
End Sub

Sub NFadeObj(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 0 'off
        Case 5:object.image = a:FadingLevel(nr) = 1 'on
    End Select
End Sub

Sub NFadeObjm(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
    End Select
End Sub

' Flasher objects

Sub Flash(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = FlashLevel(nr)
End Sub

'*********************************************************************
'                 Primitive Animation Subroutine
'*********************************************************************

Sub DropTargetBankUp(first, last)
	Dim i
	For i = first to last: DropTargets(i)(4) = 1: Next
	PlaySound "BW BankReset", 1
End Sub

Sub DropTargetHit(nr)
 	PlaySound "BW DropTargetHit", 0
	DropTargets(nr)(4) = -1
End Sub

Sub DropTargetsTimer_Timer
	Dim i
	For i = 0 to uBound(DropTargets)
		If DropTargets(i)(4) < 0 Then DropTargetDown i										' Drop Targets that have to go down
		If DropTargets(i)(4) > 0 Then DropTargetUp i										' Drop Targets that have to go up
	Next
End Sub

Sub DropTargetDown(i)
	Dim pos
	If DropTargets(i)(4) = -6 Then															' Bottom position reached?
		DropTargets(i)(4) = 0																' If yes, mark status as parked
		If NOT DropTargets(i)(1) is Nothing Then DropTargets(i)(1).IsDropped = 1			' Mark "down" postion in DropTarget
		If DropTargets(i)(0) > 0 Then Controller.Switch(DropTargets(i)(0)) = 1				' Turn ROM switch on 
		If NOT DropTargets(i)(5) is Nothing Then DropTargets(i)(5).Image = "BW GI off"
	Else
		pos = DropTargetPositions(DropTargets(i)(4) + 5) + DropTargets(i)(3)				' calculate new position
		If DropTargets(i)(2).z > pos Then 													' is current position higher?
			DropTargets(i)(2).z = pos														' if yes, step down
'			DropTargets(i)(2).TriggerSingleUpdate											' Tell VP to update the primitive
		End If																				' Drop Target moved!
		DropTargets(i)(4) = DropTargets(i)(4) - 1											' Decrement position variable
	End If
End Sub

Sub DropTargetUp(i)
	Dim pos
	If DropTargets(i)(4) = 8 Then															' Top position reached?
		DropTargets(i)(4) = 0																' If yes, mark status as parked
		If NOT DropTargets(i)(1) is Nothing Then DropTargets(i)(1).IsDropped = 0			' Mark "up" postion in DropTarget
		DropTargets(i)(2).z = DropTargets(i)(3) + DropTargetPositions(5)					' Drive the primitive in "Up" position
'		DropTargets(i)(2).TriggerSingleUpdate												' Tell VP to update the primitive
		If DropTargets(i)(0) > 0 Then Controller.Switch(DropTargets(i)(0)) = 0				' Turn ROM switch off 
		If NOT DropTargets(i)(5) is Nothing Then DropTargets(i)(5).Image = "BW_PF_off"
	Else
		pos = DropTargetPositions(DropTargets(i)(4)) + DropTargets(i)(3)					' calculate new position
		If DropTargets(i)(2).z < pos Then 													' is current position lower?
			DropTargets(i)(2).z = pos														' if yes, step up
'			DropTargets(i)(2).TriggerSingleUpdate											' Tell VP to update the primitive
		End If																				' Drop Target moved!
	DropTargets(i)(4) = DropTargets(i)(4) + 1												' Decrement position in variable
	End If
End Sub

Sub DropTargets_Init: Gate_DT.IsDropped = 1:End Sub


' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX, Rothbauerw, Thalamus and Herweh
' PlaySound sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
' *******************************************************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position

Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Set position as table object (Use object or light but NOT wall) and Vol to 1

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

' set position as table object and Vol + RndPitch manually

Sub PlaySoundAtVolPitch(sound, tableobj, Vol, RndPitch)
  PlaySound sound, 1, Vol, AudioPan(tableobj), RndPitch, 0, 0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed.

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

'Set position as table object and Vol manually.

Sub PlaySoundAtVol(sound, tableobj, Volume)
  PlaySound sound, 1, Volume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallAbsVol(sound, VolMult)
  PlaySound sound, 0, VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

' requires rampbump1 to 7 in Sound Manager

Sub RandomBump(voladj, freq)
  Dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
  PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, AudioPan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' set position as bumperX and Vol manually. Allows rapid repetition/overlaying sound

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
  PlaySound sound, 0, ABS(BOT.velz)/17, Pan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

' play a looping sound at a location with volume
Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
  PlaySound sound, -1, Vol, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
  RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = ball.x * 2 / table1.width-1
  If tmp > 0 Then
    Pan = Csng(tmp ^10)
  Else
    Pan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function VolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  VolMulti = Csng(BallVel(ball) ^2 / 150 ) * Multiplier
End Function

Function DVolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  DVolMulti = Csng(BallVel(ball) ^2 / 150 ) * Multiplier
  debug.print DVolMulti
End Function

Function BallRollVol(ball) ' Calculates the Volume of the sound based on the ball speed
  BallRollVol = Csng(BallVel(ball) ^2 / (80000 - (79900 * Log(RollVol) / Log(100))))
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function BallVelZ(ball) 'Calculates the ball speed in the -Z
  BallVelZ = INT((ball.VelZ) * -1 )
End Function

Function VolZ(ball) ' Calculates the Volume of the sound based on the ball speed in the Z
  VolZ = Csng(BallVelZ(ball) ^2 / 200)*1.2
End Function

'*** Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order

Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy)
  Dim AB, BC, CD, DA
  AB = (bx*py) - (by*px) - (ax*py) + (ay*px) + (ax*by) - (ay*bx)
  BC = (cx*py) - (cy*px) - (bx*py) + (by*px) + (bx*cy) - (by*cx)
  CD = (dx*py) - (dy*px) - (cx*py) + (cy*px) + (cx*dy) - (cy*dx)
  DA = (ax*py) - (ay*px) - (dx*py) + (dy*px) + (dx*ay) - (dy*ax)

  If (AB <= 0 AND BC <=0 AND CD <= 0 AND DA <= 0) Or (AB >= 0 AND BC >=0 AND CD >= 0 AND DA >= 0) Then
    InRect = True
  Else
    InRect = False
  End If
End Function


'********************************************************************
'      JP's VP10 Rolling Sounds (+rothbauerw's Dropping Sounds)
'********************************************************************

Const tnob = 5 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingTimer_Timer()
    Dim BOT, b
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

    For b = 0 to UBound(BOT)
        ' play the rolling sound for each ball
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' play ball drop sounds
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_ball_drop" & b, 0, ABS(BOT(b).velz)/17, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub



'*****************************************
'	ninuzzu's	FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()

	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle

	FlipperLP.RotY = LeftFlipper.CurrentAngle
	FlipperRP.RotY = RightFlipper.CurrentAngle

	'Apron Flipper
	FlipperLBP.RotY = FlipperLB.CurrentAngle

	'Mini PF
	FlipperLTP.RotY = FlipperLT.CurrentAngle
	FlipperRTP.RotY = FlipperRT.CurrentAngle

End Sub

'*****************************************
'	ninuzzu's	BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

'Sub BallShadowUpdate_timer()
    Dim BOT, b
    BOT = GetBalls
    ' hide shadow of deleted balls
    If UBound(BOT)<(tnob-1) Then
        For b = (UBound(BOT) + 1) to (tnob-1)
            BallShadow(b).visible = 0
        Next
    End If
    ' exit the Sub if no balls on the table
'    If UBound(BOT) = -1 Then Exit Sub
    ' render the shadow for each ball
    For b = 0 to UBound(BOT)
        If BOT(b).X < Table1.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 12
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
'End Sub



'************************************
' What you need to add to your table
'************************************

' a timer called RollingTimer. With a fast interval, like 10
' one collision sound, in this script is called fx_collide
' as many sound files as max number of balls, with names ending with 0, 1, 2, 3, etc
' for ex. as used in this script: fx_ballrolling0, fx_ballrolling1, fx_ballrolling2, fx_ballrolling3, etc


'******************************************
' Explanation of the rolling sound routine
'******************************************

' sounds are played based on the ball speed and position

' the routine checks first for deleted balls and stops the rolling sound.

' The For loop goes through all the balls on the table and checks for the ball speed and
' if the ball is on the table (height lower than 30) then then it plays the sound
' otherwise the sound is stopped, like when the ball has stopped or is on a ramp or flying.

' The sound is played using the VOL, AUDIOPAN, AUDIOFADE and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the AUDIOPAN & AUDIOFADE functions will change the stereo position
' according to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.


Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub LeftFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

' Wire ramp sounds

Sub RampSound1_Hit: PlaySoundAtVol"fx_metalrolling", ActiveBall, 1: End Sub
Sub RampSound2_Hit: PlaySoundAtVol"fx_metalrolling", ActiveBall, 1: End Sub
Sub RampSound3_Hit: PlaySoundAtVol"fx_metalrolling", ActiveBall, 1: End Sub
