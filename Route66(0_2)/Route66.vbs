' ***********************************************************************
'               VISUAL PINBALL X EM Script
'                  Basic EM script, 1 player only 
'		          uses core.vbs for extra functions
'
'  Route 66, an original VPX table in the Electro-Mechanical style of the 1970's
'
'  Initial starting source table was Big Deal VPX7 - version by Klodo81 2022, Big Deal version 2.1
'
'  I watched an interesting video on youtube recently that helped me better understand old EM pinball:
'	check out "Technology Connections" "Wires do the math" as he explains what's going on inside the AZTEK machine.
'
'  Thanks again to the Visual Pinball community - for making the software, making tutorials, and letting people reuse objects and code.
'
'  Version 0.0 (initial release - likely full of bugs)
'  Version 0.1 (quick fix because I left the 1 second counter as 0.1 second from testing)
'  Version 0.2 (update B2S to use illumination snippits instead of player 2 reels for the odometer so that it stops that annoying clicking)
'				Mostly visual changes. A few minor scoring changes.
'
'	For what its worth, you have my permission to use or modify anything here.  
'	I don't need any credit and don't blame me if it doesn't work.
' ***********************************************************************


Option Explicit
Randomize

' core.vbs constants
Const BallSize = 50 ' 50 is the normal size
Const BallMass = 1  ' 1 is the normal ball mass.

' load extra vbs files
LoadCoreFiles

Sub LoadCoreFiles
    On Error Resume Next
    ExecuteGlobal GetTextFile("core.vbs")
    If Err Then MsgBox "Can't open core.vbs"
    On Error Resume Next
    ExecuteGlobal GetTextFile("controller.vbs")
    If Err Then MsgBox "Can't open controller.vbs"
End Sub


'* * * * * * * * * * * * * * * * * * * * *

' Constants
Const TableName = "Route66" ' file name to save config
Const cGameName = "Route66" ' B2S name

Const MaxBallsOnPlayfield = 3
Const BallsPerGame = 5          
Const TiltSensitivity = 6


' Global variables

Dim BallsRemaining
Dim PlayerScore
Dim MileCount
Dim OdometerAddAmount
Dim BumperHitCount
Dim TiltMeasure

Dim Add1
Dim Add10
Dim Add100
Dim Add1000

' slingshot displays
Dim Lr01flex
Dim Rr01flex

' Control variables
Dim BallsOnPlayfield

' Boolean variables
Dim isGameInPlay
Dim isTilted
Dim isProjectorEnabled
       
Dim BackglassStatus 'for B2s attract mode
' core.vbs variables


'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)

'If Keycode = AddCreditKey Then
'	if GI006.state = 0 then 
'		GiOn
'	Else	
'		GiOff
'	end if
'End If
	If isGameInPlay Then
		If keycode = PlungerKey Then
			Plunger.Pullback
			PlaySoundAt "fx_plungerpull", plunger
		End If

		If NOT isTilted Then
			If keycode = LeftTiltKey Then Nudge 90, 8:PlaySound "fx_nudge", 0, 1, -0.1, 0.25:CheckTilt
			If keycode = RightTiltKey Then Nudge 270, 8:PlaySound "fx_nudge", 0, 1, 0.1, 0.25:CheckTilt
			If keycode = CenterTiltKey Then Nudge 0, 9:PlaySound "fx_nudge", 0, 1, 1, 0.25:CheckTilt

			If keycode = LeftFlipperKey  Then FlippersUpLeft
			If keycode = RightFlipperKey Then FlippersUpRight
		End If

	Else
		If Keycode = AddCreditKey Then
			PlaySound "fx_coin"
		End If

		'kecode 3 is the "2" key on the keyboard
		If keycode = StartGameKey or keycode = 3 or keycode = 4 or keycode = 5 Then
			If keycode = 3 then
				MileCount = 0
			End If
			If keycode = 4 then
				isProjectorEnabled = False
			Else
				isProjectorEnabled = True
			End If
			If keycode = 5 Then
				OdometerAddAmount = -1
			Else
				OdometerAddAmount = 1
			End If

			cleanupTimer.Enabled=True
			'when the cleanupTimer finishes, it will go to Game_Init to start the game
        End If
    End If

End Sub


Sub Table1_KeyUp(ByVal keycode)

    'plunger release
    If keycode = PlungerKey Then
        Plunger.Fire
        If PlungerLaneBallTrigger.UserValue="Hit" Then
            PlaySoundAt "fx_plunger", plunger
        Else
            PlaySoundAt "fx_plunger_empty", plunger
        End If
    End If

	If isGameInPlay Then
		If keycode = LeftFlipperKey Then FlippersDownLeft
		If keycode = RightFlipperKey Then FlippersDownRight
	End If

End Sub


Sub AddABall
	PlaySound "fx_bell_low"
    PlaySound "fx_bell_high"
    BallsRemaining = BallsRemaining + 1
End Sub

'******************
' Table stop/pause
'******************

Sub table1_Paused
End Sub

Sub table1_unPaused
End Sub


'*******************
'  Flipper Subs
'*******************

Sub FlippersUpLeft()
	If isTilted Then Exit Sub
	If Not isGameInPlay Then Exit Sub

    PlaySoundAt "LeftflipperupH-2dB", LeftFlipper
    LeftFlipper.RotateToEnd
	LeftFlipperTop.RotateToEnd
End Sub

Sub FlippersDownLeft()
    PlaySoundAt "LeftflipperdownH", LeftFlipper
	LeftFlipper.RotateToStart
	LeftFlipperTop.RotateToStart
End Sub

Sub FlippersUpRight()
	If isTilted Then Exit Sub
	If Not isGameInPlay Then Exit Sub

    PlaySoundAt "RightflipperupH-2dB", RightFlipper
    RightFlipper.RotateToEnd
	RightFlipperTop.RotateToEnd
End Sub

Sub FlippersDownRight()
    PlaySoundAt "RightflipperdownH", RightFlipper
    RightFlipper.RotateToStart
	RightFlipperTop.RotateToStart
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub LeftFlipperTop_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipperTop_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

'************************
'       Targets
'************************

'************************
' Rollover Targets
'************************
Sub RolloverTopLeft_Hit
	PlaySoundAt "fx_Sensor", RolloverTopLeft
	PlLeftEnter.State=1
	If PlRightEnter.State>0 Then
		AddScore 10
	Else
		AddScore 1
	End If
	OneSecondTimer.Enabled = True
End Sub

Sub RolloverTopRight_Hit
	PlaySoundAt "fx_Sensor", RolloverTopRight
	PlRightEnter.State=1
	If PlLeftEnter.State>0 Then
		AddScore 10
	Else
		AddScore 1
	End If
	OneSecondTimer.Enabled = True
End Sub

Sub RolloverOutlaneLeft_Hit
	PlaySoundAtBall "fx_Sensor"
	AddScore 100
	If PlOutlaneLeft.state = 1 Then
		AddABall()
		If PlOutlaneRight.state = 0 Then
			PlOutlaneLeft.state = 0
		End If
	Else		
		PlOutlaneLeft.State = 1
	End If
End Sub

Sub RolloverOutlaneRight_Hit
	PlaySoundAtBall "fx_Sensor"
	AddScore 100
	If PlOutlaneRight.state = 1 Then
' if we have already hit the right outlane, award an extra ball
		AddABall()
		If PlOutlaneLeft.state = 0 Then
' if we have not hit the left outlane, turn of the right extra ball light
' so: if you hit both the left and right outlanes, you'll always get an extra ball when you hit again
			PlOutlaneRight.state = 0
		End If
	Else		
		PlOutlaneRight.State = 1
	End If
End Sub

Sub RolloverStarInlaneLeft_Hit
	PlaySoundAtBall "fx_Sensor"      
	If PlLeftEnter.State=1 Then
		AddScore 200
	Else
		AddScore 10 + 90 * PlStarInlaneLeft.state + (100 * PlStarInlaneRight.state)
	End If
	PlStarInlaneLeft.state = 1
End Sub

Sub RolloverStarInlaneRight_Hit
	PlaySoundAtBall "fx_Sensor"
	If PlRightEnter.State=1 Then
		AddScore 200
	Else
		AddScore 10 + 90 * PlStarInlaneRight.state + (100 * PlStarInlaneLeft.state)
	End If
	PlStarInlaneRight.state = 1
End Sub

Sub RolloverStarTopLeft_hit  
	PlaySoundAtBall "fx_Sensor"                                                       
	AddScore ((10 + 90 * PlRolloverStarTopLeft.state) * (1 + PlRolloverStarTopCenter.state) * (1 + PlRolloverStarTopRight.state))
	PlRolloverStarTopLeft.state = 1
	If PlArrowLoopLeft.state = 2 Then 
		PlArrowLoopLeft.state = 1
		AddScore 500
	End If
End Sub

Sub RolloverStarTopCenter_hit
	PlaySoundAtBall "fx_Sensor"                                                       
	AddScore ((10 + 90 * PlRolloverStarTopCenter.state) * (1 + PlRolloverStarTopLeft.state) * (1 + PlRolloverStarTopRight.state))
	PlRolloverStarTopCenter.state = 1
End Sub

Sub RolloverStarTopRight_hit
	PlaySoundAtBall "fx_Sensor"                                                       
	AddScore ((10 + 90 * PlRolloverStarTopRight.state) * (1 + PlRolloverStarTopCenter.state) * (1 + PlRolloverStarTopLeft.state))
	PlRolloverStarTopRight.state = 1
	If PlArrowLoopRight.state = 2 Then 
		PlArrowLoopRight.state = 1
		AddScore 500
	End If
End Sub

Sub RolloverStarCenterEnter_hit
	PlaySoundAtBall "fx_Sensor"                                                       
	AddScore 10 
	PlRolloverStarCenterEnter.state = 1
	OneSecondTimer.Enabled = True
End Sub

'************************
' Hit Targets
'************************

Sub TargetTopLoop_hit
	PlaySoundAtBall "fx_TargetHit"
    AddScore 10 + 90 * PlLoop1.state
	If PlLoop4.state = 1 Then PlLoop5.state = 1
	If PlLoop3.state = 1 Then PlLoop4.state = 1
	If PlLoop2.state = 1 Then PlLoop3.state = 1
	If PlLoop1.state = 1 Then PlLoop2.state = 1
	PlLoop1.state = 1
	FlashTargetTopLoop.duration 1, 100, 0
End Sub

Sub TargetSmallUpLeft_Hit
	PlaySoundAtBall "fx_TargetHit"
	AddScore 1
End Sub

Sub TargetSmallUpRight_Hit
	PlaySoundAtBall "fx_TargetHit"
	AddScore 1
End Sub

Sub TargetTopLeft_Hit
	PlaySoundAtBall "fx_TargetHit"
	AddScore 10 + 90 * PlTargetTopLeft.state  * (1 + PlTargetTopRight.state)
	FlashKicker.duration 1, 100, 0
	If PlTargetTopLeft.state = 1 Then PlTargetTopLeft2.state = 1
	PlTargetTopLeft.state = 1
End Sub

Sub TargetTopRight_Hit
	PlaySoundAtBall "fx_TargetHit"
	AddScore ((10 + 90 * PlTargetTopRight.state) * (1 + PlTargetTopLeft.state))
	FlashKicker.duration 1, 100, 0
	If PlTargetTopRight.state = 1 Then PlTargetTopRight2.state = 1
	PlTargetTopRight.state = 1
End Sub

Sub TargetMidLeft_Hit
	PlaySoundAtBall "fx_TargetHit"
    AddScore 10 + 90 * PlArrowTargetMidLeft.state
	FlashTargetMidLeft.duration 1, 100, 0
	If PlArrowTargetMidLeft.state = 2 Then
		PlArrowTargetMidLeft.state = 1
		Addscore 500
	End If
End Sub

Sub TargetMidLeft2_Hit
	PlaySoundAtBall "fx_TargetHit"
    AddScore 300
	PlTargetMidLeft2.state = 1
End Sub

Sub TargetBumperLeft_hit
	PlaySoundAt "fx_TargetHit", TargetBumperLeft
    AddScore 10 + 90 * PlTargetBumperLeft.state
	PlTargetBumperLeft.state = 1

	If LightBumperUpLeft.State = 0 Then
		LightBumperUpLeft.State = 1
	ElseIf LightBumperUpLeft.State = 1 Then
		If LightBumperUpRight.State > 0 Then
			LightBumperUpLeft.State = 2
		Else
			LightBumperUpLeft.State = 0
		End If
	End If

	If LightBumperLowLeft.State = 0 Then
		LightBumperLowLeft.State = 1
	ElseIf LightBumperLowLeft.State = 1 Then
		If LightBumperLowRight.State > 0 Then
			LightBumperLowLeft.State = 2
		Else
			LightBumperLowLeft.State = 0
		End If
	End If

End Sub

Sub TargetBumperRight_hit
	PlaySoundAt "fx_TargetHit", TargetBumperRight
    AddScore 10 + 90 * PlTargetBumperRight.state
	PlTargetBumperRight.state = 1

	If LightBumperUpRight.State = 0 Then
		LightBumperUpRight.State = 1
	ElseIf LightBumperUpRight.State = 1 Then
		If LightBumperUpLeft.State > 0 Then
			LightBumperUpRight.State = 2
		Else
			LightBumperUpRight.State = 0
		End If
	End If

	If LightBumperLowRight.State = 0 Then
		LightBumperLowRight.State = 1
	ElseIf LightBumperLowRight.State = 1 Then
		If LightBumperLowLeft.State > 0 Then
			LightBumperLowRight.State = 2
		Else
			LightBumperLowRight.State = 0
		End If
	End If
End Sub


'************************
'       Drop Targets
'************************

Sub TargetDropLeft1_Hit
	PlaySoundAt "fx_TargetDrop1", TargetDropLeft1
    AddScoreNoSound 10
End Sub
Sub TargetDropLeft2_Hit
	PlaySoundAt "fx_TargetDrop1", TargetDropLeft2
    AddScoreNoSound 10
End Sub
Sub TargetDropLeft3_Hit
	PlaySoundAt "fx_TargetDrop1", TargetDropLeft3
    AddScoreNoSound 10
End Sub
Sub TargetDropLeft4_Hit
	PlaySoundAt "fx_TargetDrop1", TargetDropLeft4
    AddScoreNoSound 10
End Sub
Sub TargetDropLeft5_Hit
	PlaySoundAt "fx_TargetDrop1", TargetDropLeft5
    AddScoreNoSound 10
End Sub

Sub TargetDropLeft1_Dropped()
	CheckLeftDropsCompleted()
End Sub
Sub TargetDropLeft2_Dropped()
	CheckLeftDropsCompleted()
End Sub
Sub TargetDropLeft3_Dropped()
	CheckLeftDropsCompleted()
End Sub
Sub TargetDropLeft4_Dropped()
	CheckLeftDropsCompleted()
End Sub
Sub TargetDropLeft5_Dropped()
	CheckLeftDropsCompleted()
End Sub

Sub TargetDropRight1_Hit
	PlaySoundAt "fx_TargetDrop1", TargetDropRight1
    AddScoreNoSound 10
End Sub
Sub TargetDropRight2_Hit
	PlaySoundAt "fx_TargetDrop1", TargetDropRight2
    AddScoreNoSound 10
End Sub
Sub TargetDropRight3_Hit
	PlaySoundAt "fx_TargetDrop1", TargetDropRight3
    AddScoreNoSound 10
End Sub
Sub TargetDropRight4_Hit
	PlaySoundAt "fx_TargetDrop1", TargetDropRight4
    AddScoreNoSound 10
End Sub

Sub TargetDropRight1_Dropped()
	CheckRightDropsCompleted()
End Sub
Sub TargetDropRight2_Dropped()
	CheckRightDropsCompleted()
End Sub
Sub TargetDropRight3_Dropped()
	CheckRightDropsCompleted()
End Sub
Sub TargetDropRight4_Dropped()
	CheckRightDropsCompleted()
End Sub


Sub RaiseDropTargetsLeft
	If TargetDropLeft1.IsDropped = 1 Then
		TargetDropLeft1.IsDropped = 0
		PlaySoundAt "fx_TargetUp", TargetDropLeft1
	End If
	If TargetDropLeft2.IsDropped = 1 Then
		TargetDropLeft2.IsDropped = 0
		PlaySoundAt "fx_TargetUp", TargetDropLeft2
	End If
	If TargetDropLeft3.IsDropped = 1 Then
		TargetDropLeft3.IsDropped = 0
		PlaySoundAt "fx_TargetUp", TargetDropLeft3
	End If
	If TargetDropLeft4.IsDropped = 1 Then
		TargetDropLeft4.IsDropped = 0
		PlaySoundAt "fx_TargetUp", TargetDropLeft4
	End If
	If TargetDropLeft5.IsDropped = 1 Then
		TargetDropLeft5.IsDropped = 0
		PlaySoundAt "fx_TargetUp", TargetDropLeft5
	End If
End Sub

Sub RaiseDropTargetsRight
	If TargetDropRight1.IsDropped = 1 Then
		TargetDropRight1.IsDropped = 0
		PlaySoundAt "fx_TargetUp", TargetDropRight1
	End If
	If TargetDropRight2.IsDropped = 1 Then
		TargetDropRight2.IsDropped = 0
		PlaySoundAt "fx_TargetUp", TargetDropRight2
	End If
	If TargetDropRight3.IsDropped = 1 Then
		TargetDropRight3.IsDropped = 0
		PlaySoundAt "fx_TargetUp", TargetDropRight3
	End If
	If TargetDropRight4.IsDropped = 1 Then
		TargetDropRight4.IsDropped = 0
		PlaySoundAt "fx_TargetUp", TargetDropRight4
	End If
End Sub


Sub CheckLeftDropsCompleted
	If TargetDropLeft1.IsDropped = 1 and TargetDropLeft2.IsDropped = 1 and TargetDropLeft3.IsDropped = 1 _
       and TargetDropLeft4.IsDropped = 1 and TargetDropLeft5.IsDropped = 1 Then
        AddScore 1000
		LightBonus()
		RaiseDropTargetsLeft
	End If

End Sub

Sub CheckRightDropsCompleted
	If TargetDropRight1.IsDropped = 1 and TargetDropRight2.IsDropped = 1 and TargetDropRight3.IsDropped = 1 and TargetDropRight4.IsDropped = 1 Then
        AddScore 1000
		LightBonus()
		RaiseDropTargetsRight
	End If
End Sub

Sub LightBonus
	If PLBonus5.state = 1 Then
		FlasherProjector.ImageA="special"
		AddScore 10000
	End If
	If PlBonus4.state = 1 Then PlBonus5.state = 1
	If PlBonus3.state = 1 Then PlBonus4.state = 1
	If PlBonus2.state = 1 Then PlBonus3.state = 1
	If PlBonus1.state = 1 Then PlBonus2.state = 1
	PlBonus1.state = 1
End Sub

'************************
'       Spinners
'************************

Sub SpinnerTopCenter_Spin()
    PlaySoundAt "fx_spinner", SpinnerTopCenter
	If PlArrowCenterSpinner.state = 2 Then
		PlArrowCenterSpinner.state = 1
		AddScore 100
	Elseif PlArrowCenterSpinner.state = 1 Then
		AddScore 10
	Else
		AddScore 1
	End If
	OneSecondTimer.Enabled = True
End Sub

Sub SpinnerToShooter_Spin()
    PlaySoundAt "fx_spinner", SpinnerToShooter
End Sub


'*****************
'       Kickers
'*****************

Sub KickerTop_Hit
    PlaySoundAt "fx_kicker_enter", kickerTop
    If isTilted Then KickerTop.kick 180, 5: Exit Sub
	KickerTop.UserValue = True
	'KickerTop.TimerEnabled=True
	If PlArrowKicker.state = 2 Then
		PlArrowKicker.state = 1	
	End If
	BonusCountTimer.Enabled = True
End Sub

Sub KickerTop_Release
	if KickerTop.UserValue = True Then
		KickerTop.kick 170, 5 
		PlaySoundAt "fx_kicker", KickerTop
        'KickerTop.TimerEnabled=True
		KickerTop.UserValue = False
		PlTargetTopCenter.state = 1
    End If
End Sub

'*********
' Bumpers
'*********

Sub BumperUpLeft_Hit 
    If isTilted Then Exit Sub
    PlaySoundAt "fx_Bumper", BumperUpLeft
	FlashBumperUpLeft.State=1
	FlashBumperUpLeft2.State=1
	FlashBumperUpLeft.TimerEnabled = True
    AddScore 10 + 90 * LightBumperUpLeft.State
	BumperHitCount = BumperHitCount + 1
	'debug.print BumperHitCount
	If BumperHitCount >= 10 Then
		LightArrow()
		BumperHitCount = 0
	End If
End Sub

Sub BumperUpRight_Hit 
    If isTilted Then Exit Sub
    PlaySoundAt "fx_Bumper", BumperUpRight
	FlashBumperUpRight.State=1
	FlashBumperUpRight2.State=1
	FlashBumperUpRight.TimerEnabled = True
    AddScore 10 + 90 * LightBumperUpRight.State
	BumperHitCount = BumperHitCount + 1
	If BumperHitCount >= 10 Then
		LightArrow()
		BumperHitCount = 0
	End If
End Sub

Sub BumperLowLeft_Hit 
    If isTilted Then Exit Sub
    PlaySoundAt "fx_Bumper", BumperLowLeft
	FlashBumperLowLeft.State=1
	FlashBumperLowLeft2.State=1
	FlashBumperLowLeft.TimerEnabled = True
    AddScore 10 + 90 * LightBumperLowLeft.State
	BumperHitCount = BumperHitCount + 1
	If BumperHitCount >= 10 Then
		LightArrow()
		BumperHitCount = 0
	End If
End Sub

Sub BumperLowRight_Hit 
    If isTilted Then Exit Sub
    PlaySoundAt "fx_Bumper", BumperLowRight
	FlashBumperLowRight.State=1
	FlashBumperLowRight2.State=1
	FlashBumperLowRight.TimerEnabled = True
    AddScore 10 + 90 * LightBumperLowRight.State
	BumperHitCount = BumperHitCount + 1
	If BumperHitCount >= 10 Then
		LightArrow()
		BumperHitCount = 0
	End If
End Sub

Sub FlashBumperUpLeft_Timer
    FlashBumperUpLeft.State=0
	FlashBumperUpLeft2.state=0
	FlashBumperUpLeft.TimerEnabled = False
End Sub

Sub FlashBumperUpRight_Timer
    FlashBumperUpRight.State=0
	FlashBumperUpRight2.state=0
	FlashBumperUpRight.TimerEnabled = False
End Sub

Sub FlashBumperLowLeft_Timer
    FlashBumperLowLeft.State=0
	FlashBumperLowLeft2.state=0
	FlashBumperLowLeft.TimerEnabled = False
End Sub

Sub FlashBumperLowRight_Timer
    FlashBumperLowRight.State=0
	FlashBumperLowRight2.state=0
	FlashBumperLowRight.TimerEnabled = False
End Sub



'* * * * * * * *
'  table_init = turn on machine
'    game_init = start button - start new game  
'      
'        turn_init = start a new turn (the next ball)
'          ball_init = create a new ball in the plunger lane
'	       ball_end = destroy a ball
'        turn_end = end a turn

'    game_end = game over 
'  table_exit = turn off machine   
'**********



'******************
' Table power on
'******************
Sub Table1_Init()

    'VPObjects_Init
    LoadEM

    isGameInPlay = False
    
    UpdateBackglassInfoDisplay()

    Add1 = 0    
    Add10 = 0
    Add100 = 0
    Add1000 = 0

    'turn on GI lights
    vpmtimer.addtimer 1000, "GiOn '"

    ' start the attract mode
    StartAttractMode()

    ' Remove desktop items in FS mode
    Dim x
    
    For each x in cReels
		If Table1.ShowDT then
            x.Visible = 1
		Else
            x.Visible = 0
        End If
    Next

	FlasherProjector.ImageA = "milestart"

	'wait a bit so that the B2S is loaded first
	vpmtimer.addtimer 1000, "LoadConfig() '"	

End Sub


'****************************************
' Init for a new game - start button pressed
'****************************************
Sub Game_Init()

    isGameInPlay = True
    isTilted = False

    StopAttractMode
    GiOn
	TurnOffPlayfieldLights()

	FlasherProjector.ImageA = "milestart"
    PlayerScore = 0
	ResetScores()
	ResetMiles()
    BallsRemaining = BallsPerGame

    ' first ball.  delay just a bit to let the score reset sounds finish
    vpmtimer.addtimer 1000, "Turn_Init '"
    
End Sub

'****************************************
'  start a new turn (the next ball)
'****************************************
Sub Turn_Init()
    TiltMeasure = 0

    TurnInitResetLights
    TurnInitResetVariables
    TurnInitResetPlayfield

    UpdateBackglassInfoDisplay()

    'using the timer to pause a bit before adding the ball when the turn starts, it just feels more natural
    vpmtimer.addtimer 1000, "Ball_Init() '"	
End Sub

Sub TurnInitResetLights() 'init lights for new ball/player
	'most lights are reset by the turn end bonus count process
	LightBumperUpLeft.State = 0 
	LightBumperLowLeft.state = 0 
	LightBumperUpRight.State = 0
	LightBumperLowRight.State = 0
	PlRolloverStarTopLeft.state = 0
	PlRolloverStarTopCenter.state = 0
	PlRolloverStarTopRight.state = 0
	PlLeftEnter.state = 0
	PlRolloverStarCenterEnter.state = 0
	PlRightEnter.state = 0
	PlStarInlaneLeft.state = 0
	PlStarInlaneRight.state = 0
End Sub

Sub TurnInitResetVariables() 'init variables new ball/player
    '
End Sub

Sub TurnInitResetPlayfield()
    RaiseDropTargetsLeft()
    RaiseDropTargetsRight()
	LightRandomBumper()
End Sub

Sub LightRandomBumper
	Randomize timer
	Dim aRnd : aRnd = Int(4 * Rnd)
	Select Case(aRnd)
		Case 0:LightBumperLowLeft.state = 1
		Case 1:LightBumperLowRight.state = 1
		Case 2:LightBumperUpLeft.state = 1
		Case 3:LightBumperUpRight.state = 1
	End Select

End Sub

'****************************************
'  Create a new ball in the plunger lane
'****************************************
Sub Ball_Init()
    'debug.print "Ball_Init"
    
    If BallsOnPlayfield>=MaxBallsOnPlayfield Then Exit Sub
    
    BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass
    
    PlaySoundAt "fx_Ballrel", BallRelease
    BallRelease.Kick 90, 4

    BallsOnPlayfield = BallsOnPlayfield + 1
End Sub

    
Sub PlungerLaneBallTrigger_Hit()                      
	PlungerLaneBallTrigger.UserValue="Hit" 'use like a switch to know if there is a ball for the plunger to hit
End Sub

Sub PlungerLaneBallTrigger_UnHit()
	PlungerLaneBallTrigger.UserValue="UnHit"
End Sub


'****************************************
'  A ball is no longer in play, but there may still be other balls on the playfield
'****************************************
Sub Ball_End()
    'debug.print "Ball_End"


	If(BallsOnPlayfield = 0) Then
       'no balls left on playfield, the turn is done.
		OneSecondTimer.Enabled = False
		vpmtimer.addtimer 0010, "Turn_End() '"
    End If

End Sub


'****************************************
'end the turn (no balls left on playfield)
'****************************************
Sub Turn_End()

    BallsRemaining = BallsRemaining - 1

    If isTilted Then
'un-Tilt
        isTilted = False
		TurnOffPlayfieldLights()
		EnableTable()
        UpdateBackglassInfoDisplay()
        TiltMeasure = 0
	End If

	BonusCountTimer.Enabled = True 'this will count up the bonus and go to the next turn

End Sub

Sub Game_End()

    isGameInPlay = False

	FlasherProjector.ImageA = "mileend"
    PlaySound "fx_match"

    UpdateBackglassInfoDisplay()

    FlippersDownLeft
    FlippersDownRight

    'turn off gi Lights
    GiOff

    ' start the attract mode
    vpmtimer.addtimer 1000, "StartAttractMode '"

End Sub


'******************
' Table power off
'******************
Sub table1_Exit
    SaveConfig 
    If B2SOn Then Controller.Stop 'stop the B2S controller
End Sub

'**************
'     TILT
'**************

Sub CheckTilt
    TiltMeasure = TiltMeasure + TiltSensitivity
    TiltDecreaseTimer.Enabled = True
    If TiltMeasure> 15 Then
		PlaySound "fx_knocker2"
        isTilted = True
		DisableTable()
        ReelTilt.SetValue 1
        If B2SOn then
            Controller.B2SSetTilt 1
        End If
		If isProjectorEnabled Then 
			Dim ballinplay:ballinplay = BallsPerGame - BallsRemaining + 1
			Select Case(ballinplay)
				Case 1:FlasherProjector.ImageA = "Tilt1"
				Case 2:FlasherProjector.ImageA = "Tilt2"
				Case 3:FlasherProjector.ImageA = "Tilt3"
				Case 4:FlasherProjector.ImageA = "Tilt4"
				Case 5:FlasherProjector.ImageA = "Tilt5"
			End Select
		End If
    End If
End Sub

Sub TiltDecreaseTimer_Timer
    ' DecreaseTilt
    If TiltMeasure> 0 Then
        TiltMeasure = TiltMeasure - 0.1
    Else
        TiltDecreaseTimer.Enabled = False
    End If
End Sub

Sub DisableTable
    GiOff
    FlippersDownLeft()
    FlippersDownRight()
    BumperUpLeft.Threshold = 100
    BumperUpRight.Threshold = 100
    BumperLowLeft.Threshold = 100
    BumperLowRight.Threshold = 100
	TurnOffPlayfieldLights()
'    LeftSlingshot.Disabled = 1
'    RightSlingshot.Disabled = 1
End Sub

Sub EnableTable
    GiOn
    BumperUpLeft.Threshold = 1
    BumperUpRight.Threshold = 1
    BumperLowLeft.Threshold = 1
    BumperLowRight.Threshold = 1
End Sub


'***************************************************************
'             Supporting Ball & Sound Functions v4.0
'  includes random pitch in PlaySoundAt and PlaySoundAtBall
'***************************************************************

Dim TableWidth, TableHeight

TableWidth = Table1.width
TableHeight = Table1.height

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / TableWidth-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10))
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = (SQR((ball.VelX ^2) + (ball.VelY ^2)))
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp
    tmp = ball.y * 2 / TableHeight-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10))
    End If
End Function

Sub PlaySoundAt(soundname, tableobj) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
    PlaySound soundname, 0, 1, Pan(tableobj), 0.2, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics
    PlaySound soundname, 0, Vol(ActiveBall), pan(ActiveBall), 0.2, Pitch(ActiveBall) * 10, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallNoSpeed(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics
    PlaySound soundname, 0, 1, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function

'***********************************
'   JP's VP10 Rolling Sounds v4.0
'   JP's Ball Shadows
'   JP's Ball Speed Control
'   Rothbauer's dropping sounds
'***********************************

Const tnob = 19   'total number of balls
Const lob = 0     'number of locked balls
Const maxvel = 60 'max ball velocity
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate() 'call this routine from any realtime timer you may have, running at an interval of 10 is good.

    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
    BOT = GetBalls

    ' stop the sound of deleted balls and hide the shadow
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
        cBallShadow(b).Y = 2100 'under the apron 'may differ from table to table
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' draw the ball shadow
    For b = lob to UBound(BOT)
        cBallShadow(b).X = BOT(b).X
        cBallShadow(b).Y = BOT(b).Y
        cBallShadow(b).Height = BOT(b).Z - Ballsize / 2 + 1

    'play the rolling sound for each ball
        If BallVel(BOT(b)) > 1 Then
            If BOT(b).z < 30 Then
                ballpitch = Pitch(BOT(b))
                ballvol = Vol(BOT(b))
            Else
                ballpitch = Pitch(BOT(b)) + 50000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b)) * 5
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, ballvol, Pan(BOT(b)), 0, ballpitch, 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' rothbauerw's Dropping Sounds
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        End If

        ' jps ball speed control
        If BOT(b).VelX AND BOT(b).VelY <> 0 Then
            speedfactorx = ABS(maxvel / BOT(b).VelX)
            speedfactory = ABS(maxvel / BOT(b).VelY)
            If speedfactorx < 1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactorx
                BOT(b).VelY = BOT(b).VelY * speedfactorx
            End If
            If speedfactory < 1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactory
                BOT(b).VelY = BOT(b).VelY * speedfactory
            End If
        End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'***************************************
'     Collection collision sounds
'***************************************

'Sub ShooterLaneGate_Hit
'	debug.print"shooterlanegatehit hit"
'End Sub
Sub cMetals_Hit(idx):PlaySoundAtBall "fx_MetalHit":End Sub
Sub cMetalWires_Hit(idx):PlaySoundAtBall "fx_MetalWire":End Sub
Sub cRubber_Bands_Hit(idx):PlaySoundAtBall "fx_rubber_band":End Sub
Sub cRubber_Posts_Hit(idx):PlaySoundAtBall "fx_rubber_post":End Sub
Sub cRubber_Pins_Hit(idx):PlaySoundAtBall "fx_rubber_pin":End Sub
Sub cPlastics_Hit(idx):PlaySoundAtBall "fx_PlasticHit":End Sub
Sub cGates_Hit(idx):PlaySoundAtBall "fx_Gate":End Sub
Sub cWoods_Hit(idx):PlaySoundAtBall "fx_Woodhit":End Sub

' *********************************************************************
'  Drain
' *********************************************************************

Sub Drain_Hit()
    Drain.DestroyBall
    PlaySoundAt "fx_drain", Drain

    'there is no active game
    If NOT isGameInPlay Then Exit Sub

    If BallsOnPlayfield > 0 Then
        BallsOnPlayfield = BallsOnPlayfield - 1
    End If

    Ball_End

End Sub


' ****************************************
'             Score functions
' ****************************************

Sub AddScore(Points)
    If isTilted Then Exit Sub

    If Points = 0 Then
		'nothing
	ElseIf Points < 100 Then
		PlaySound "fx_bigbell10"
	Elseif Points < 1000 Then
		PlaySound "fx_bigbell100"
	Else
        PlaySound "fx_bigbell1000"
    End If
    AddScoreNoSound Points
End Sub

Sub AddScoreNoSound(Points)
    If isTilted Then Exit Sub

    If Points = 1 OR Points = 10 OR Points = 100 Or Points = 1000 Then
        PlayerScore = PlayerScore + points
        UpdateScore points
    End If

    If Points > 1 And Points < 10 Then
        Add1 = Add1 + Points \ 1
        AddScore1Timer.Enabled = TRUE
    End If

    If Points > 10 And Points < 100 Then
        Add10 = Add10 + Points \ 10
        AddScore10Timer.Enabled = TRUE
    End If
    If Points > 100 AND Points < 1000 Then
        Add100 = Add100 + Points \ 100
        AddScore100Timer.Enabled = TRUE
    End If
    If Points > 1000 Then
        Add1000 = Add1000 + Points \ 1000
        AddScore1000Timer.Enabled = TRUE
    End If

End Sub


'************************************
'       Score sound Timers
'************************************

Sub AddScore1Timer_Timer()
    if Add1 > 0 then
        AddScore 1
        Add1 = Add1 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore10Timer_Timer()
    if Add10 > 0 then
        AddScore 10
        Add10 = Add10 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore100Timer_Timer()
    if Add100 > 0 then
        AddScore 100
        Add100 = Add100 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore1000Timer_Timer()
    if Add1000 > 0 then
        AddScore 1000
        Add1000 = Add1000 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub UpdateScore(playerpoints)
    ScoreReel1.Addvalue playerpoints
    If B2SOn then
        Controller.B2SSetScorePlayer1 PlayerScore
    end if
End Sub

Sub ResetScores
	'ScoreReel1.setvalue 11111
    ScoreReel1.ResetToZero
    If B2SOn then
        Controller.B2SSetScorePlayer1 0
		'Controller.B2SSetData 81,0
		'Controller.B2SSetData 82,0
		'Controller.B2SSetData 83,0
		'Controller.B2SSetData 84,0
    End If
End Sub

Sub ResetMiles
    ScoreReel2.SetValue MileCount
    If B2SOn then
        Controller.B2SSetScorePlayer2 MileCount
	End If
End Sub

'*******************
'     Bonus
'*******************

Sub BonusCountTimer_Timer
	
	If PlLeftEnter.state > 0 AND PlRightEnter.state > 0 Then
		PlLeftEnter.state = 0 
		PlRightEnter.state = 0
		If PlRolloverStarCenterEnter.state > 0 Then
			AddScore 500
			PlRolloverStarCenterEnter.state = 0
		Else
			AddScore 100
		End If
		Exit Sub
	End If

	If PlLoop5.state > 0 Then
		AddScore 500
		PlLoop5.state = 0
		Exit Sub
	End If
	If PlLoop4.state > 0 Then
		AddScore 400
		PlLoop4.state = 0
		Exit Sub
	End If
	If PlLoop3.state > 0 Then
		AddScore 300
		PlLoop3.state = 0
		Exit Sub
	End If
	If PlLoop2.state > 0 Then
		AddScore 200
		PlLoop2.state = 0
		Exit Sub
	End If
	If PlLoop1.state > 0 Then
		AddScore 100
		PlLoop1.state = 0
		Exit Sub
	End If
	If PlTargetTopLeft2.state > 0 Then
		AddScore 100
		PlTargetTopLeft2.state = 0
		Exit Sub
	End If
	If PlTargetTopLeft.state > 0 Then
		AddScore 10
		PlTargetTopLeft.state = 0
		Exit Sub
	End If
	If PlTargetTopRight2.state > 0 Then
		AddScore 100
		PlTargetTopRight2.state = 0
		Exit Sub
	End If
	If PlTargetTopRight.state > 0 Then
		AddScore 10
		PlTargetTopRight.state = 0
		Exit Sub
	End If
	If PlTargetTopCenter.state > 0 Then
		AddScore 100
		PlTargetTopCenter.state = 0
		Exit Sub
	End If
	If PlTargetBumperLeft.state > 0 Then
		AddScore 10
		PlTargetBumperLeft.state = 0
		Exit Sub
	End If
	If PlTargetBumperRight.state > 0 Then
		AddScore 10
		PlTargetBumperRight.state = 0
		Exit Sub
	End If
	If PlTargetMidLeft2.state > 0 Then
		AddScore 10
		PlTargetMidLeft2.state = 0
		Exit Sub
	End If

	If PlBonus5.state > 0 Then
		AddScore 5000
		PlBonus5.state = 0
		Exit Sub
	End If
	If PlBonus4.state > 0 Then
		AddScore 4000
		PlBonus4.state = 0
		Exit Sub
	End If
	If PlBonus3.state > 0 Then
		AddScore 3000
		PlBonus3.state = 0
		Exit Sub
	End If
	If PlBonus2.state > 0 Then
		AddScore 2000
		PlBonus2.state = 0
		Exit Sub
	End If
	If PlBonus1.state > 0 Then
		AddScore 1000
		PlBonus1.state = 0
		Exit Sub
	End If

	If PlArrowLoopLeft.state = 1 Then
		AddScore 5000
		PlArrowLoopLeft.state = 0
		Exit Sub
	End If
	If PlArrowLoopRight.state = 1 Then
		AddScore 5000
		PlArrowLoopRight.state = 0
		Exit Sub
	End If
	If PlArrowCenterSpinner.state = 1 Then 
		AddScore 5000
		PlArrowCenterSpinner.state = 0
		Exit Sub
	End If
	If PlArrowKicker.state = 1 Then 
		AddScore 5000
		PlArrowKicker.state = 0
		Exit Sub
	End If
	If PlArrowTargetMidLeft.state = 1 Then 
		AddScore 5000
		PlArrowTargetMidLeft.state = 0
		Exit Sub
	End If

	BonusCountTimer.Enabled = False

	If KickerTop.UserValue = True Then
		'triggered from Kicker hit
		vpmtimer.addtimer 0100, "KickerTop_Release '"
	Else
		'triggered from turn_end
		If BallsRemaining > 0 Then
			vpmtimer.addtimer 0100, "Turn_Init '"
		Else
			vpmtimer.addtimer 0100, "Game_End '"
		End If
	End If

End Sub

'***********
' Playfield lights
'***********

Sub TurnOnPlayfieldLights()
    Dim bulb
    For each bulb in cGameLights
        bulb.State = 1
    Next
End Sub

Sub TurnOffPlayfieldLights()
    Dim bulb
    For each bulb in cGameLights
        bulb.State = 0
    Next
End Sub

'***********
' GI lights
'***********

Sub GiOn 
	'debug.print "GiOn"
	PlaySound"fx_gion"
	Dim bulb
    For each bulb in cGiLights
        bulb.State = 1
        'debug.print bulb.name
    Next
End Sub

Sub GiOff 
	'debug.print "GiOff"
	PlaySound"fx_gioff"    
    Dim bulb
    For each bulb in cGiLights
        bulb.State = 0
		'debug.print bulb.name
    Next
End Sub

' ********************************
'        Attract Mode
' ********************************
' use the"Blink Pattern" of each light

Sub StartAttractMode()
	'this is an EM machine, so pretty dull attract mode....
	TurnOnPlayfieldLights()
	BackglassAttractTimer.Enabled=True
End Sub

Sub BackglassAttractTimer_Timer
	'not used
     If Not B2SOn Then Exit Sub
End Sub

Sub StopAttractMode()
    BackglassAttractTimer.Enabled=False
    BackglassAttractTitleOn()
    TurnOffPlayfieldLights()
End Sub

Sub BackglassAttractTitleOn
	'not used
     If Not B2SOn Then Exit Sub
End Sub

Sub BackglassAttractTitleOff
	'not used
     If Not B2SOn Then Exit Sub
End Sub


Sub LoadConfig

    Dim FileObj
    Dim ConfigFile, TextStr
    Dim fileline(20)

    Set FileObj = CreateObject("Scripting.FileSystemObject")
    If Not FileObj.FolderExists(UserDirectory) Then Exit Sub
    If Not FileObj.FileExists(UserDirectory & TableName & ".txt") Then Exit Sub
    Set ConfigFile = FileObj.GetFile(UserDirectory & TableName& ".txt")
    Set TextStr = ConfigFile.OpenAsTextStream(1, 0)

    Dim aLine : aLine = 1
    do until aLine > 20
        fileline(aLine) = TextStr.ReadLine
        If TextStr.AtEndOfStream Then Exit Do
		aLine = aLine + 1
    loop

    TextStr.Close
    Set ConfigFile = Nothing
    Set FileObj = Nothing

    On Error Resume Next
    MileCount = CInt(Mid(fileline(1),7)) 

    On Error Resume Next
    PlayerScore = CInt(Mid(fileline(2),7)) 
	If PlayerScore = 0 Then PlayerScore = 1234

	'UpdateBackglassInfoDisplay()
	ScoreReel1.setvalue PlayerScore
	ScoreReel2.setvalue MileCount
    If B2SOn then
        Controller.B2SSetScorePlayer 1, PlayerScore
		B2sOdometer()
		'Controller.B2SSetScorePlayer 2, MileCount
    end if

End Sub

Sub SaveConfig
    Dim FileObj
    Set FileObj = CreateObject("Scripting.FileSystemObject")
    If Not FileObj.FolderExists(UserDirectory)then
        Exit Sub
    End If
	Dim ConfigFile
    Set ConfigFile = FileObj.CreateTextFile(UserDirectory & TableName & ".txt", True)
	ConfigFile.WriteLine "Miles=" & MileCount
	ConfigFile.WriteLine "Score=" & PlayerScore
    ConfigFile.Close
    Set ConfigFile = Nothing
    Set FileObj = Nothing
End Sub


'****************************************
' Realtime updates
'****************************************

Sub GameTimer_Timer
	Ballcheck
    RollingUpdate
End Sub

Sub Ballcheck
	'if any ball is off of the playfield, set the location right in front of the drain so it gets destroyed
    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
    BOT = GetBalls

    For b = 0 to UBound(BOT)
		if BOT(b).X < 0 or BOT(b).X > Table1.Width or BOT(b).Y < 0 or BOT(b).X > Table1.Height Then
			'debug.print b & " " & BOT(b).X & " " & BOT(b).Y
			BOT(b).x = Drain.x
			BOT(b).y = Drain.y - BOT(b).Radius
		end if
		
	Next
End Sub

'Sub Ballcheck_game_init
	
 '   Dim BOT, b
 '   BOT = GetBalls

  '  For b = 0 to UBound(BOT)
'	Next
'End Sub

Sub cleanupTimer_Timer
	'starting new game. if any balls left from testing, set the location right in front of the drain so it gets destroyed
	'BUT, use a timer to spread it out because if there are a bunch of them they pile up
    Dim BOT, b
    BOT = GetBalls

    'debug.print UBound(BOT)
	If Ubound(BOT) >= 0 Then
		BOT(b).x = Drain.x
		BOT(b).y = Drain.y - BOT(b).Radius
	Else
		cleanupTimer.Enabled = false
		Game_init
	End If
End Sub

' ***********************************************************************
'  *********************************************************************
'   *********     G A M E  C O D E  S T A R T S  H E R E      *********
'  *********************************************************************
' ***********************************************************************
Sub LightArrow

	'all arrows lit
	If PlArrowLoopLeft.state = 1 And PlArrowLoopRight.state = 1 And PlArrowCenterSpinner.state = 1 And PlArrowKicker.state = 1 And PlArrowTargetMidLeft.state = 1 Then Exit Sub

	If PlArrowLoopLeft.state = 2 Then PlArrowLoopLeft.state = 0
	If PlArrowLoopRight.state = 2 Then PlArrowLoopRight.state = 0
	If PlArrowCenterSpinner.state = 2 Then PlArrowCenterSpinner.state = 0
	If PlArrowKicker.state = 2 Then PlArrowKicker.state = 0
	If PlArrowTargetMidLeft.state = 2 Then PlArrowTargetMidLeft.state = 0

    Dim isDone: isDone=False
    do until isDone

		Randomize timer
		Dim aRnd : aRnd = Int((5 * Rnd) + 1)

		If aRnd = 1 and PlArrowLoopLeft.state = 0 Then 
			PlArrowLoopLeft.state = 2
			isDone = True
		End If
		If aRnd = 2 and PlArrowLoopRight.state = 0 Then 
			PlArrowLoopRight.state = 2
			isDone = True
		End If
		If aRnd = 3 and PlArrowCenterSpinner.state = 0 Then 
			PlArrowCenterSpinner.state = 2
			isDone = True
		End If
		If aRnd = 4 and PlArrowKicker.state = 0 Then 
			PlArrowKicker.state = 2
			isDone = True
		End If
		If aRnd = 5 and PlArrowTargetMidLeft.state = 0 Then 
			PlArrowTargetMidLeft.state = 2
			isDone = True
		End If
    loop

End Sub


'**********************
'       Slingshots
'**********************

Dim LStep, RStep

Sub SlingshotLeftWall_Slingshot   'left slingshot
	'debug.print "slingshot"
    If isTilted Then Exit Sub
    PlaySoundAt "fx_slingshot", Lemk
	
	FlashSlingLowLeft.state = 1
    LStep = 0
    SlingshotLeftWall.TimerEnabled = True
    AddScore 10

End Sub

Sub SlingShotLeftWall_Timer    
    Select Case LStep
        Case 0:LeftSling2.Visible = 1:Lemk.RotX = 2
        Case 1:LeftSling3.Visible = 1:LeftSling2.Visible = 0:Lemk.RotX = 14
        Case 2:LeftSling4.Visible = 1:LeftSling3.Visible = 0:Lemk.RotX = 26
        Case 3:LeftSling3.Visible = 1:LeftSling4.Visible = 0:Lemk.RotX = 14
        Case 4:LeftSling2.Visible = 1:LeftSling3.Visible = 0:Lemk.RotX = 2
        Case 5:LeftSling2.Visible = 0:Lemk.RotX = -20:FlashSlingLowLeft.state = 0:SlingShotLeftWall.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub SlingShotRightWall_Slingshot    'right slingshot
    If isTilted Then Exit Sub
    PlaySoundAt "fx_slingshot", Remk

	FlashSlingLowRight.state = 1
    RStep = 0
    SlingShotRightWall.TimerEnabled = True
    AddScore 10

End Sub

Sub SlingShotRightWall_Timer
    Select Case RStep
        Case 0:RightSling2.Visible = 1:Remk.RotX = 2
        Case 1:RightSling3.Visible = 1:RightSling2.Visible = 0:Remk.RotX = 14
        Case 2:RightSling4.Visible = 1:RightSling3.Visible = 0:Remk.RotX = 26
        Case 3:RightSling3.Visible = 1:RightSling4.Visible = 0:Remk.RotX = 14
        Case 4:RightSling2.Visible = 1:RightSling3.Visible = 0:Remk.RotX = 2
        Case 5:RightSling2.Visible = 0:Remk.RotX = -20:FlashSlingLowRight.state = 0:SlingShotRightWall.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'***********************
'       Rubbers
'***********************

Sub Rubber010_Hit
	'upper right slingshot behind some drop targets
	FlashSlingUpRight.duration 1, 100, 0
	'FlashSlingUpRight.state = 1
	'FlashSlingUpRight.TimerEnabled = True
	AddScore 1
End Sub

Sub FlashSlingUpRight_Timer
	FlashSlingUpRight.state = 0
	FlashSlingUpRight.TimerEnabled = False
End Sub

Sub LeftUpperSlingWall_Hit
	FlashSlingUpLeft.duration 1, 100, 0
	'FlashSlingUpLeft.state = 1
	'FlashSlingUpLeft.TimerEnabled = True
	AddScore 1
End Sub

Sub FlashSlingUpLeft_Timer
	FlashSlingUpLeft.state = 0
	FlashSlingUpLeft.TimerEnabled = False
End Sub


'***************************************************************
' Backglass display
'***************************************************************

Sub UpdateBackglassInfoDisplay()

	'handles backglass and backdrop display for things other than scores
	' shoot again || game over || credits || active player || number of players || player turn (aka ball in play) || tilt

	If isGameInPlay Then
		ReelGameOver.SetValue 0
		If B2SOn Then Controller.B2SSetGameOver 0
	Else
        ReelGameOver.SetValue 1
		If B2SOn Then Controller.B2SSetGameOver 1
    End If

    Dim ballinplay
    ballinplay = BallsPerGame - BallsRemaining + 1
	If ballinplay < 0 OR NOT isGameInPlay Then
		ballinplay = 0
	End If

    ReelBallInPlay.SetValue ballinplay
    If B2SOn Then Controller.B2SSetBallInPlay ballinplay


    If isTilted Then
        ReelTilt.SetValue 1
        If B2SOn Then Controller.B2SSetTilt 1
	Else
        ReelTilt.SetValue 0
        If B2SOn Then Controller.B2SSetTilt 0
	End If
            
End Sub

Sub OneSecondTimer_Timer
	'we are traveling at the incredible speed of 1 mile per second = 3600 miles per hour - about 10 times the speed of sound.
	'so...uh...better wear a seat belt or something...
	If isGameInPlay=False OR isTilted OR isProjectorEnabled = False Then 
		OneSecondTimer.Enabled=False
		Exit Sub
	End If

	If MileCount > 2144 Then 
		MileCount = 0
	ElseIf (MileCount < 0) OR (MileCount = 0 AND OdometerAddAmount < 0) Then 
		MileCount = 2144
	Else
		MileCount = MileCount + OdometerAddAmount
	End If

	ScoreReel2.Setvalue MileCount
	B2sOdometer()
	'If B2SOn then Controller.B2SSetScorePlayer2 MileCount


	'dig this: all of our pictures are in the image manager and named "mile####"
	'we have them all listed in this array.  Each time we hit this routine, we take the current mile we are on and generate an image name
	'then we look in the array.  If we find a match, we change the picture showing in the "projector" flasher object.

	Dim PicId, PicArray
	PicArray = Array( "mile0001","mile0004","mile0010","mile0025","mile0035","mile0040","mile0046","mile0050","mile0057", _
"mile0060","mile0066","mile0070","mile0078","mile0080","mile0085","mile0090","mile0105","mile0110","mile0120","mile0130", _
"mile0135","mile0142","mile0145","mile0151","mile0154","mile0157","mile0160","mile0163","mile0166","mile0169","mile0172", _
"mile0175","mile0180","mile0182","mile0184","mile0190","mile0193","mile0196","mile0200","mile0206","mile0209","mile0212", _
"mile0215","mile0218","mile0221","mile0224","mile0227","mile0230","mile0233","mile0237","mile0245","mile0250","mile0255", _
"mile0260","mile0264","mile0268","mile0272","mile0275","mile0277","mile0280","mile0283","mile0286","mile0289","mile0293", _
"mile0296","mile0305","mile0310","mile0320","mile0325","mile0340","mile0350","mile0355","mile0360","mile0365","mile0380", _
"mile0390","mile0400","mile0405","mile0408","mile0420","mile0425","mile0430","mile0460","mile0465","mile0470","mile0480", _
"mile0500","mile0510","mile0515","mile0520","mile0525","mile0530","mile0535","mile0540","mile0550","mile0560","mile0570", _
"mile0575","mile0580","mile0585","mile0590","mile0595","mile0598","mile0601","mile0604","mile0607","mile0610","mile0615", _
"mile0620","mile0625","mile0630","mile0635","mile0640","mile0645","mile0650","mile0655","mile0663","mile0670","mile0675", _
"mile0685","mile0690","mile0695","mile0700","mile0705","mile0710","mile0715","mile0735","mile0740","mile0745","mile0750", _
"mile0755","mile0762","mile0767","mile0771","mile0775","mile0800","mile0805","mile0810","mile0815","mile0850","mile0870", _
"mile0875","mile0900","mile0905","mile0945","mile0956","mile0960","mile0965","mile0970","mile0983","mile0991","mile1015", _
"mile1020","mile1030","mile1045","mile1050","mile1062","mile1066","mile1090","mile1110","mile1139","mile1141","mile1146", _
"mile1150","mile1160","mile1170","mile1200","mile1220","mile1280","mile1285","mile1340","mile1344","mile1389","mile1415", _
"mile1420","mile1430","mile1474","mile1500","mile1540","mile1545","mile1550","mile1565","mile1570","mile1575","mile1585", _
"mile1600","mile1605","mile1616","mile1620","mile1625","mile1635","mile1655","mile1660","mile1670","mile1685","mile1691", _
"mile1720","mile1725","mile1733","mile1741","mile1750","mile1755","mile1760","mile1765","mile1794","mile1803","mile1808", _
"mile1823","mile1840","mile1857","mile1860","mile1868","mile1880","mile1890","mile1900","mile1925","mile1928","mile1934", _
"mile1940","mile1945","mile1950","mile1955","mile1960","mile1980","mile1990","mile2000","mile2045","mile2050","mile2060", _
"mile2080","mile2085","mile2091","mile2107","mile2115","mile2118","mile2124","mile2127","mile2130","mile2138","mile2140")

	'gotta add some leading zeros to make the name format "mile####"
	if milecount < 9 Then
		PicId = "mile000"& milecount
	else 
		if milecount < 99 Then
			PicId = "mile00" & milecount
		else 
			if milecount < 999 Then
				PicId = "mile0" & milecount
			else 
				if milecount < 9999 Then
					PicId = "mile" & milecount
				end If
			end if
		end if
	end If

	'debug.print PicId
	'checking the array for a match
	dim z
    For each z in PicArray
        if z = PicId then
			'debug.print z
			FlasherProjector.ImageA = PicId
		end if
    Next


End Sub

Sub B2sOdometer()

	If NOT B2SOn Then Exit Sub

'ones
	Dim x:	x = MileCount mod 10
    Select Case x
        Case 0:Controller.B2SSetData 1, 10
		Case 1:Controller.B2SSetData 1, 1
		Case 2:Controller.B2SSetData 1, 2
		Case 3:Controller.B2SSetData 1, 3
		Case 4:Controller.B2SSetData 1, 4
		Case 5:Controller.B2SSetData 1, 5
		Case 6:Controller.B2SSetData 1, 6
		Case 7:Controller.B2SSetData 1, 7
		Case 8:Controller.B2SSetData 1, 8
		Case 9:Controller.B2SSetData 1, 9
    End Select

'tens
	x = INT(MileCount / 10) mod 10
    Select Case x
        Case 0:Controller.B2SSetData 2, 10
		Case 1:Controller.B2SSetData 2, 1
		Case 2:Controller.B2SSetData 2, 2
		Case 3:Controller.B2SSetData 2, 3
		Case 4:Controller.B2SSetData 2, 4
		Case 5:Controller.B2SSetData 2, 5
		Case 6:Controller.B2SSetData 2, 6
		Case 7:Controller.B2SSetData 2, 7
		Case 8:Controller.B2SSetData 2, 8
		Case 9:Controller.B2SSetData 2, 9
    End Select

'hundreds
	x = INT(MileCount / 100) mod 10
    Select Case x
        Case 0:Controller.B2SSetData 3, 10
		Case 1:Controller.B2SSetData 3, 1
		Case 2:Controller.B2SSetData 3, 2
		Case 3:Controller.B2SSetData 3, 3
		Case 4:Controller.B2SSetData 3, 4
		Case 5:Controller.B2SSetData 3, 5
		Case 6:Controller.B2SSetData 3, 6
		Case 7:Controller.B2SSetData 3, 7
		Case 8:Controller.B2SSetData 3, 8
		Case 9:Controller.B2SSetData 3, 9
    End Select

'thousands
	x = INT(MileCount / 1000) mod 10
	debug.print x
    Select Case x
        Case 0:Controller.B2SSetData 4, 10
		Case 1:Controller.B2SSetData 4, 1
		Case 2:Controller.B2SSetData 4, 2
		Case 3:Controller.B2SSetData 4, 3
		Case 4:Controller.B2SSetData 4, 4
		Case 5:Controller.B2SSetData 4, 5
		Case 6:Controller.B2SSetData 4, 6
		Case 7:Controller.B2SSetData 4, 7
		Case 8:Controller.B2SSetData 4, 8
		Case 9:Controller.B2SSetData 4, 9
    End Select
End Sub
