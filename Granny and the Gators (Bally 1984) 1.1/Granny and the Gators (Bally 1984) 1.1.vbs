Option Explicit

On Error Resume Next
	ExecuteGlobal GetTextFile("controller.vbs")
	If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01530000","Granny.vbs",3.1


' Sub LoadVPM(VPMver,VBSfile,VBSver)
' 	On Error Resume Next
' 		If ScriptEngineMajorVersion<5 Then MsgBox"VB Script Engine 5.0 or higher required"
' 		ExecuteGlobal GetTextFile(VBSfile)
' 		If Err Then MsgBox"Unable to open "&VBSfile&". Ensure that it is in the same folder as this table."&vbNewLine&Err.Description:Err.Clear

' '**************************************************************************************************
' '********************************ACTIVATE BACKGLASS************************************************
' If ShowDT=true then
' 		Set Controller=CreateObject("VPinMAME.Controller")
' 	else
' 		Set Controller=CreateObject("B2S.server") 
' End if
' '**************************************************************************************************

' 		If Err Then MsgBox"Can't Load VPinMAME."&vbNewLine&Err.Description
' 		If VPMver>"" Then If Controller.Version<VPMver Or Err Then MsgBox"VPinMAME ver "&VPMver&" required.":Err.Clear
' 		If VPinMAMEDriverVer<VBSver Or Err Then MsgBox VBSFile&" ver "&VBSver&" or higher required."
' 	On Error Goto 0
' End Sub



Const UseSolenoids=1,UseLamps=1,UseSync=1,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown",SCoin="Coin3"

SolCallback(1)="bsTrough.SolOut"
SolCallback(2)="dtDrop.SolDropUp"
SolCallback(3)="bsSaucer2.SolOut"
SolCallback(4)="bsSaucer.SolOut"
SolCallback(8)="Flippers"

Sub Flippers(Enabled)
	If Enabled Then
		FlipperOn=1
	Else
		FlipperOn=0
		If LS=1 Then LeftFlipper.RotateToStart
		If RS=1 Then RightFlipper.RotateToStart
		If LS=1 Or RS=1 Then
			PlaySound"FlipperDown"
			LS=0
			RS=0
		End If
	End If
End Sub

Sub SetDisplayToElement(Element)
 		Exit Sub
  	Dim playerRect
 	playerRect=Controller.GetClientRect(GetPlayerHwnd)
 	Dim playerWidth, playerHeight
 	playerWidth=playerRect(2)-playerRect(0)
 	playerHeight=playerRect(3)-playerRect(1)
 	Dim Game
 	Set Game=Controller.Game
 	Dim x,y
  	x=Element.x*playerWidth/1000.0-1
 	y=Element.y*playerHeight/750.0-1
 	Game.Settings.SetDisplayPosition x,y,GetPlayerHwnd
 	Set Game=nothing
End Sub

Sub Table1_Paused:Controller.Pause=True:End Sub
Sub Table1_unPaused:Controller.Pause=False:End Sub

Dim bsTrough,bsSaucer,bsSaucer2,dtDrop,FlipperOn,LS,RS
FlipperOn=0:LS=0:RS=0


Sub Table1_Exit			' STAT add
	Controller.Stop
End Sub


Sub Table1_Init
	On Error Resume Next
		Controller.GameName="granny":If Err Then MsgBox"Can't start Game"&cGameName&vbNewLine&Err.Description:Exit Sub
		Controller.SplashInfoLine="Granny And The Gators"
		Controller.HandleMechanics=0
		Controller.HandleKeyboard=0
		Controller.ShowDMDOnly=1
		Controller.ShowFrame=0
		Controller.ShowTitle=0

'**************COMMENT OUT THE FOLLOWING LINE IF PLAYING IN CABINET MODE**********
If Table1.ShowDT then
			SetDisplayToElement TextBox1
		Else
			TextBox1.visible=False
End If

		Controller.Run GetPlayerHwnd:If Err Then MsgBox Err.Description:Exit Sub
	On Error Goto 0
	vpmNudge.TiltSwitch=15:vpmNudge.Sensitivity=5:PinMAMETimer.Interval=PinMAMEInterval:PinMAMETimer.Enabled=1
	
	Set bsTrough=New cvpmBallStack
	bsTrough.InitSaucer Drain,18,345,12
    bsTrough.InitExitSnd"Popper","SolOn"
	bsTrough.AddBall 0
	Drain.CreateBall
	
	
	Set bsSaucer=New cvpmBallStack
	bsSaucer.InitSaucer Kicker1,25,167,8
    bsSaucer.InitExitSnd"Popper","SolOn"
	
	Set bsSaucer2=New cvpmBallStack
	bsSaucer2.InitSaucer Kicker2,26,0,14
    bsSaucer2.InitExitSnd"Popper","SolOn"
	
	Set dtDrop=New cvpmDropTarget
	dtDrop.InitDrop Array(Drop1,Drop2,Drop3,Drop4,Drop5),Nothing
	dtDrop.InitSnd"FlapOpen","FlapClos"
	dtDrop.CreateEvents"dtDrop"
		
	vpmCreateEvents AllSwitches
	vpmMapLights AllLights
End Sub

Sub Drain_Hit:bsTrough.AddBall 0:LPad.state=1:End Sub
Sub Kicker1_Hit:bsSaucer.AddBall 0:End Sub
Sub Kicker2_Hit:bsSaucer2.AddBall 0:End Sub

ExtraKeyHelp=KeyName(StartGameKey)&vbTab&"1 Player Start+Fire"&vbNewLine&_
KeyName(KeyFront)&vbTab&"2 Player Start+Fire"&vbNewLine&_
KeyName(KeyJoyLeft)&vbTab&"Left Paddle"&vbNewLine&_
KeyName(KeyJoyRight)&vbTab&"Right Paddle"&vbNewLine&_
KeyName(KeyJoyUp)&vbTab&"Power Paddle"&vbNewLine&_
KeyName(KeyJoyDown)&vbTab&"Fire"



Sub Table1_KeyDown(ByVal KeyCode)
	If KeyCode=StartGameKey Then Controller.Switch(-1)=1
	If KeyCode=KeyFront Then Controller.Switch(0)=1
	If KeyCode=KeyJoyLeft Then Controller.Switch(-2)=1
	If KeyCode=KeyJoyRight Then Controller.Switch(-3)=1
	If KeyCode=KeyJoyUp Then Controller.Switch(41)=1
	If KeyCode=KeyJoyDown Then Controller.Switch(-1)=1
	If KeyCode=LeftFlipperKey Then
		If FlipperOn=1 And LS=0 Then
			LeftFlipper.RotateToEnd
			PlaySound"FlipperUp"
			LS=1
		End If
	End If
	If KeyCode=RightFlipperKey Then
		If FlipperOn=1 And RS=0 Then
			RightFlipper.RotateToEnd
			PlaySound"FlipperUp"
			RS=1
		End If
	End If
	If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyCode=StartGameKey Then Controller.Switch(-1)=0
	If KeyCode=KeyFront Then Controller.Switch(0)=0
	If KeyCode=KeyJoyLeft Then Controller.Switch(-2)=0
	If KeyCode=KeyJoyRight Then Controller.Switch(-3)=0
	If KeyCode=KeyJoyUp Then Controller.Switch(41)=0
	If KeyCode=KeyJoyDown Then Controller.Switch(-1)=0
	If KeyCode=LeftFlipperKey And LS=1 Then
		LeftFlipper.RotateToStart
		PlaySound"FlipperDown"
		LS=0
	End If
	If KeyCode=RightFlipperKey And RS=1 Then
		RightFlipper.RotateToStart
		PlaySound"FlipperDown"
		RS=0
	End If
	If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

sub S29_hit
Playsound "Wallhit"
vpmTimer.PulseSw 29
End Sub
sub S30_hit
Playsound "Wallhit"
vpmTimer.PulseSw 30
End Sub
sub S31_hit
Playsound "Wallhit"
vpmTimer.PulseSw 31
End Sub
sub S32_hit
Playsound "Wallhit"
vpmTimer.PulseSw 32
End Sub
sub S33_hit
Playsound "Wallhit"
vpmTimer.PulseSw 33
End Sub
sub S34_hit
Playsound "Wallhit"
vpmTimer.PulseSw 34
End Sub
sub S35_hit
Playsound "Wallhit"
vpmTimer.PulseSw 35
End Sub
sub S36_hit
Playsound "Wallhit"
vpmTimer.PulseSw 36
End Sub

Sub Trigger1_Hit
	LPad.state=0
End Sub

Sub ButtonTimer_timer
If LPad.State=0 then bigbutton1.z=80: bigbutton4.z=-80 End If
If LPad.State=1 then bigbutton1.z=-80: bigbutton4.z=80 End If
If LPad.State=0 then bigbutton3.z=80: bigbutton6.z=-80 End If
If LPad.State=1 then bigbutton3.z=-80: bigbutton6.z=80 End If
If L7.State=0 then bigbutton2.z=80: bigbutton5.z=-80 End If
If L7.State=1 then bigbutton2.z=-80: bigbutton5.z=80 End If
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 1 ' total number of balls
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

	' play the rolling sound for each ball
    For b = 0 to UBound(BOT)
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*200, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next
End Sub


Sub Gates_Hit(idx)
	Playsound "gate"
End Sub

Sub Targets_Hit(idx)
	Playsound "droptarget2"
End Sub

Sub Plastics_Hit(idx)
	Playsound "flip_hit_2"
End Sub

Sub Rubbers_Hit(idx)
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub

Sub Metals_Hit(idx)
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "metal_hit1", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "metal_hit2", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "metal_hit3", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub

Sub LeftFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub


Sub Gate_Collide(parm)
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "metal_hit1", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "metal_hit2", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "metal_hit3", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub

Sub Rubbers_Hit(idx)
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub

Sub Metals_Hit(idx)
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "metal_hit1", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "metal_hit2", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "metal_hit3", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub

Sub LeftFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub Gate_Collide(parm)
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "metal_hit1", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "metal_hit2", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "metal_hit3", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub


