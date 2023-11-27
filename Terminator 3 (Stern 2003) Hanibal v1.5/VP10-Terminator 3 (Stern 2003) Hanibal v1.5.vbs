
'**********************************************************************************************************************************************************
'******************************************** Terminator 3 Hanibal 4k VPX Version *************************************************************************
'**********************************************************************************************************************************************************
'
' directb2s for this Table by STAT, here: http://www.vpforums.org/index.php?app=downloads&showfile=13315
' Optimizations by Darthmarino
'

 Option Explicit
 Randomize
 
 Dim fade, x, bstrough, cbcaptive, bstxhole, hole31, bsvuk, dtdroptarget, Bump1, Bump2, Bump3, launch, GunMech, plungerIM,RPGSpezial

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

 LoadVPM "01560000", "SEGA.VBS", 3.26
 
 Const UseSolenoids = 1
 Const UseLamps = 1
 Const UseGI = 0
 Const UseSync = 0
 Const HandleMech = 1

 Const SCoin = "Coin"

 BSize = 25.5
 BMass = ((BSize*2)^3)/100000
 
'Table InitAddSnd

 Sub t3_Init

'*****************************************************************************
'************************** GAME VARIATION HERE! ******************************
'***     To use original Stern instruction cards, set the ICards to =0     ***
'***   To use custom cards from pinballcards.com, set the ICards to =1     ***
'*****************************************************************************

Const Icards = 1

 CapKicker.CreateBall
 CapKicker.Kick 180, 1
 K1.createball
 GW.isdropped=1
 If Icards=0 Then CardLCust.Z=-1:CardLorig.Z=70:CardRCust.Z=-1:CardRorig.Z=70:Else CardLCust.Z=70:CardLorig.Z=-1:CardRCust.Z=70:CardRorig.Z=-1:End If
 Dim cGameName
 vpmInit Me
 cGameName = "term3"
 With Controller
  .GameName = cGameName
 .SplashInfoLine = "Terminator 3 - Rise of the Machines - Stern 2003" & vbNewLine & "by Hanibal" 
 .HandleKeyboard = 0
 .ShowTitle = 0
 .ShowDMDOnly = 1
 .ShowFrame = 0
 .HandleMechanics = 0
 .Hidden = 0
 .Games(cGameName).Settings.Value("rol")=0
 .Games(cGameName).Settings.Value("ror")=0
 On Error Resume Next
 .Run GetPlayerHWnd
 If Err Then MsgBox Err.Description
 On Error Goto 0
 End With
 
'TROUGH INIT

 Set bsTrough = New cvpmBallStack
 With bsTrough
 .InitSw 0, 14, 13, 12, 11, 0, 0, 0
 .InitKick BallRelease, 45, 9
 .InitEntrySnd "Solenoid", "Solenoid"
 .InitExitSnd "ballrel", "Solenoid"
 .Balls = 4
 End With
 
'RIGHT VUK INIT
 
  Set bsvuk=new cvpmballstack
 With bsvuk
 .InitSw 0,36,0,0,0,0,0,0
 .InitKick vukout,180,8
 .InitExitSnd "Holeexit", "Solenoid"
 End With
 
'LEFT VUK INIT

 Set bstxhole=new cvpmballstack
 With bstxhole
 .InitSw 0,33,34,35,0,0,0,0
 .InitKick tvukout,180,12
 .InitExitSnd "Holeexit", "Solenoid" 
 End With
 
'DROPTARGET INIT

 Set dtdroptarget=new cvpmdroptarget
 dtdroptarget.initdrop sw25,25

'HIDE DESKTOP MODE BACKGROUND LIGHTS

 If T3.ShowDT = False Then
    Apron2.Z=-1
    Apron1.Z=70
	RPGFlasheye.visible=False
	RPGFLASH28a.visible=false
	Light1.visible=False
	Light2.visible=False
	Light3.visible=False
	Light4.visible=False
	Light5.visible=False
	l65.visible=False
	l66.visible=False
	l67.visible=False
	l68.visible=False
	l69.visible=False
	RPGFLASH27a.visible=false
	RPGFLASH27a1.visible=false
	RPGFLASH27a2.visible=false
	RPGFLASH27a3.visible=false
	RPGFLASH27a4.visible=false
	RPGFLASH4.visible=false
	RPGFLASH4a.visible=false
	RPGFLASH4a1.visible=false
	RPGFLASH4a2.visible=false
	RPGFLASH4a3.visible=false
Light1a.visible=False
Light2a.visible=False
Light3a.visible=False
Light4a.visible=False
Light5a.visible=False

Else

Apron1.Z=-1
Apron2.Z=70
Wall76.isdropped = 1
rpgwalls.isdropped = 1
Wall476.isdropped = 1
Wall77.isdropped = 1
bbknothome.isdropped = 1
Wall78.isdropped = 1

sw41a.visible=False
sw42a.visible=False
sw43a.visible=False
sw44a.visible=False
sw45a.visible=False
RPGFLASH28a1.visible=false
RPGFLASH27a5.visible=false
RPGFLASH27a6.visible=false
RPGFLASH1.visible=false
RPGrichtung.visible=false

l69a.visible=false
l68a.visible=false
l67a.visible=false
l66a.visible=false
l65a.visible=false
Light1a.visible=False
Light2a.visible=False
Light3a.visible=False
Light4a.visible=False
Light5a.visible=False



End If
 
'PINMAME TIMER INIT

 PinMAMETimer.Interval = PinMAMEInterval
 PinMAMETimer.Enabled = 1
 'T3.maxballspeed = 45
 'StartShake
 
 'NUDGE INIT
 
 vpmNudge.TiltSwitch = 56
 vpmNudge.Sensitivity = 0.5
 vpmNudge.TiltObj = Array(lbumper, bbumper, rbumper, LeftSlingshot, RightSlingshot)
 
'MISC INIT
 
 CenterPost.isdropped = 1
 LeftPost.isdropped = 1
 Bwdiverter1a.isdropped = 1
 ballinbbk=true
 kickback.pullback
 kickramp.collidable = 0
  LS.isdropped = 1
  RS.isdropped = 1
RPGStopp.isdropped = 1

	If B2SOn Then Controller.B2SSetData 129,1
 
 End Sub
 

 Sub t3_Exit: Controller.Stop : End Sub
 Sub t3_Paused:Controller.Pause = 1:End Sub
 Sub t3_unPaused:Controller.Pause = 0:End Sub
 
'KEYBOARD HANDLE ******************************************************************************************************************************************

 Sub T3_KeyDown(ByVal Keycode)
 
 If keycode=plungerkey then controller.switch(55)=true
 If keycode=keyfront then controller.switch(53)=true:end if
 'If keycode = LeftTiltKey Then vpmNudge.DoNudge 80, 1.6 
 'If keycode = RightTiltKey Then vpmNudge.DoNudge 280, 1.6
 'If keycode = CenterTiltKey Then vpmNudge.DoNudge 0, 1.6

 If Keycode = LeftTiltKey Then Nudge 90, 8: PlaySound "fx_nudge_left"
 If Keycode = RightTiltKey Then Nudge 270, 8: PlaySound "fx_nudge_right"
 If Keycode = CenterTiltKey Then Nudge 0, 8: PlaySound "fx_nudge_forward"

 'If keycode=plungerkey then
 'If RPGrichtung.visible = True And RPGReset = False Then 
 'Set rpgball= RPGreal.Createsizedball(9):rpgball.image = "rpgball":rpgball.id = 666: RPGreal.kick RPGfeuer,16: RPGrichtung.visible = 0: RPGrichtung1.visible = 1: RPGReset = True
 'playsound "rpgshot"
 'End If
 'End If

 If keycode = KeyRules Then Rules
 If KeyDownHandler(KeyCode) Then Exit Sub
 End Sub
 
 Sub T3_KeyUp(ByVal Keycode)
 If keycode=plungerkey then controller.switch(55)=false
 If keycode=keyfront then controller.switch(53)=false:end if
 If KeyUpHandler(KeyCode) Then Exit Sub
 End Sub
 
'FLIPPERS *************************************************************************************************************************************************
 
 SolCallback(sLLFlipper) = "SolLeftFlippers"
 SolCallback(sLRFlipper) = "SolRightFlippers"
 
 Sub SolLeftFlippers(Enabled)
 SolFlipL LeftFlipper, Enabled
 End Sub
 
 Sub SolRightFlippers(Enabled)
 SolFlipR RightFlipper, Enabled
 End Sub
 
 Sub SolFlipL(Flip, Enabled)
 If Enabled Then
 PlaySound "FlipperUpLeft"
 Flip.RotateToEnd
 Else
 PlaySound "FlipperDownLeft"
 Flip.RotateToStart
 End If
 End Sub
 
 Sub SolFlipR(Flip, Enabled)
 If Enabled Then
 PlaySound "FlipperUpRight"
 Flip.RotateToEnd
 Else
 PlaySound "FlipperDownRight"
 Flip.RotateToStart
 End If
 End Sub
 
'SLINGSHOTS ***********************************************************************************************************************************************
 
 Sub LeftSlingShot_Slingshot:LS.IsDropped = 0:PlaySound "slingshotleft":vpmTimer.PulseSw 59:Me.TimerEnabled = 1:End Sub
 Sub LeftSlingShot_Timer:Me.TimerEnabled = 0:LS.IsDropped = 1:End Sub
 
 Sub RightSlingShot_Slingshot:RS.IsDropped = 0:PlaySound "slingshotright":vpmTimer.PulseSw 62:Me.TimerEnabled = 1:End Sub
 Sub RightSlingShot_Timer:Me.TimerEnabled = 0:RS.IsDropped = 1:End Sub
 
'DRAIN HOLES **********************************************************************************************************************************************
 
 Sub Drain_Hit:Playsound "drain":bsTrough.AddBall Me:End Sub
 Sub Drain2_Hit:Playsound "drain":bsTrough.AddBall Me:End Sub
 Sub Drain1_Hit:Playsound "drain":bsTrough.AddBall Me:End Sub
 Sub Drain4_Hit:Playsound "drain":bsTrough.AddBall Me:End Sub
 Sub Drain3_Hit:Playsound "drain":bsTrough.AddBall Me:End Sub
 
 'AUTOLAUNCH HANDLER ***************************************************************************************************************************************

 Sub shooterlane_Hit():Controller.Switch(16) = 1:End Sub:Sub shooterlane_Unhit():Controller.Switch(16) = 0:End Sub
 
 Const IMPowerSetting = 48
 Const IMTime = 0.6  
 Set plungerIM = New cvpmImpulseP
 With plungerIM
 .InitImpulseP swplunger, IMPowerSetting, IMTime
 '.Random 0.3
 .InitExitSnd "plunger2", "plunger"
 .CreateEvents "plungerIM"
 End With
 
 Sub Auto_Plunger(Enabled)
 If Enabled Then
 PlungerIM.AutoFire
 End If
 End Sub
 
'SOLENOIDS*************************************************************************************************************************************************
 
 Solcallback(1)  = "bstrough.solout"
 Solcallback(2)  = "Auto_plunger"
 Solcallback(3)  = "dtdroptarget.soldropup"
 Solcallback(4)  = "Flash4"
 Solcallback(5)  = "sol5" '5 backbox kicker
 Solcallback(8)  = "dtdroptarget.solhit 1," 
 Solcallback(12) = "sol12"
 Solcallback(13) = "bsvuk.solout"
 Solcallback(14) = "bstxhole.solout"
 Solcallback(20) = "Sol20"
 Solcallback(21) = "sol21"
 Solcallback(22) = "sol22"
 Solcallback(23) = "sol23"
 Solcallback(26) = "Flash26"
 Solcallback(27) = "Flash27"
 Solcallback(28) = "Flash28"


 Solcallback(29) = ""'"SetLamp 129,"
 Solcallback(30) = "Flash30"
 Solcallback(31) = "Flash31"
 Solcallback(32) = "Flash32"
 
'RPG GUN HANDLER (BASED ON JAMINS CODE) *******************************************************************************************************************

Dim rpgpos,rpgindex,ballinbbk, rpgball, b2si, b2skicknr
 b2skicknr = 0

 'Sub sol5(enabled)
 'If enabled Then
 'If ballinbbk=true Then kickit

 'End If
 'End Sub

 

 Sub ResetB2SGun
	For b2si = 120 to 129
	Controller.B2SSetData b2si, 0
	Next
	For b2skicknr = 151 to 155
    Controller.B2SSetData b2skicknr, 0
    Next
    Controller.B2SSetData 129, 1
 End Sub

 rpgpos=0
 rpgindex=1

 Sub sol20timer_timer()
 rpgpos=rpgpos+rpgindex
 If rpgpos = 1 then Light1.state = 1:Light2.state = 0:Light3.state = 0:Light4.state = 0:Light5.state = 0:Light1a.state = 1:Light2a.state = 0:Light3a.state = 0:Light4a.state = 0:Light5a.state = 0
 If rpgpos = 2 then Light1.state = 0:Light2.state = 1:Light3.state = 0:Light4.state = 0:Light5.state = 0:Light1a.state = 0:Light2a.state = 1:Light3a.state = 0:Light4a.state = 0:Light5a.state = 0
 If rpgpos = 3 then Light1.state = 0:Light2.state = 0:Light3.state = 1:Light4.state = 0:Light5.state = 0:Light1a.state = 0:Light2a.state = 0:Light3a.state = 1:Light4a.state = 0:Light5a.state = 0
 If rpgpos = 4 then Light1.state = 0:Light2.state = 0:Light3.state = 0:Light4.state = 1:Light5.state = 0:Light1a.state = 0:Light2a.state = 0:Light3a.state = 0:Light4a.state = 1:Light5a.state = 0
 If rpgpos = 5 then Light1.state = 0:Light2.state = 0:Light3.state = 0:Light4.state = 0:Light5.state = 1:Light1a.state = 0:Light2a.state = 0:Light3a.state = 0:Light4a.state = 0:Light5a.state = 1
 If rpgpos = 6 then Light1.state = 0:Light2.state = 0:Light3.state = 0:Light4.state = 0:Light5.state = 1:Light1a.state = 0:Light2a.state = 0:Light3a.state = 0:Light4a.state = 0:Light5a.state = 1

	If B2SOn Then
		ResetB2SGun:
		'If rpgpos > 0 and b2skicknr > 0 Then Controller.B2SSetData 125-rpgpos, 1
If rpgpos = 0 and b2skicknr > 0 Then Controller.B2SSetData 124, 1
If rpgpos = 1 and b2skicknr > 0 Then Controller.B2SSetData 124, 1
If rpgpos = 2 and b2skicknr > 0 Then Controller.B2SSetData 123, 1
If rpgpos = 3 and b2skicknr > 0 Then Controller.B2SSetData 122, 1
If rpgpos = 4 and b2skicknr > 0 Then Controller.B2SSetData 121, 1
If rpgpos = 5 and b2skicknr > 0 Then Controller.B2SSetData 120, 1
If rpgpos = 6 and b2skicknr > 0 Then Controller.B2SSetData 120, 1
		'If rpgpos > 0 and b2skicknr = 0 Then Controller.B2SSetData 129-rpgpos, 1
If rpgpos =0 and b2skicknr = 0 Then Controller.B2SSetData 129, 1
If rpgpos =1 and b2skicknr = 0 Then Controller.B2SSetData 129, 1
If rpgpos =2 and b2skicknr = 0 Then Controller.B2SSetData 128, 1
If rpgpos =3 and b2skicknr = 0 Then Controller.B2SSetData 127, 1
If rpgpos =4 and b2skicknr = 0 Then Controller.B2SSetData 126, 1
If rpgpos =5 and b2skicknr = 0 Then Controller.B2SSetData 125, 1
If rpgpos =6 and b2skicknr = 0 Then Controller.B2SSetData 125, 1
		'If rpgpos = 0 and b2skicknr = 0 Then Controller.B2SSetData 129,1
		'If rpgpos = 0 and b2skicknr > 0 Then Controller.B2SSetData 124,1
	End If

  If T3.ShowDT = True Then
 If rpgpos>5 then rpgindex=-1
 If rpgpos<1 then rpgindex=1
 If ballinbbk=true then
 playsound "motor"
 rpgreel.setvalue rpgpos
 else
 playsound "motor"
 rpgreel.setvalue rpgpos
 End If
 Else

 If rpgpos>4 then rpgindex=-1
 If rpgpos<1 then rpgindex=1
 If rpgpos=0 then
 bbknothome.isdropped=true
 else
 bbknothome.isdropped=false
 End If
 If ballinbbk=true then
 playsound "motor"
 rpgreel.setvalue rpgpos
 else
 playsound "motor"
 rpgreel.setvalue rpgpos
 End If
 End If



 End Sub

 Sub kickit()
 Select Case rpgpos
 Case 0
 Light1.state = 1
 Light2.state = 0
 Light3.state = 0
 Light4.state = 0
 Light5.state = 0
 Light1a.state = 1
 Light2a.state = 0
 Light3a.state = 0
 Light4a.state = 0
 Light5a.state = 0
 Kickto1
 Case 1
 Light1.state = 1
 Light2.state = 0
 Light3.state = 0
 Light4.state = 0
 Light5.state = 0
 Light1a.state = 1
 Light2a.state = 0
 Light3a.state = 0
 Light4a.state = 0
 Light5a.state = 0
 Kickto1
 Case 2
 Light1.state = 0
 Light2.state = 1
 Light3.state = 0
 Light4.state = 0
 Light5.state = 0
 Light1a.state = 0
 Light2a.state = 1
 Light3a.state = 0
 Light4a.state = 0
 Light5a.state = 0
 Kickto2
 Case 3
 Light1.state = 0
 Light2.state = 0
 Light3.state = 1
 Light4.state = 0
 Light5.state = 0
 Light1a.state = 0
 Light2a.state = 0
 Light3a.state = 1
 Light4a.state = 0
 Light5a.state = 0
 Kickto3
 Case 4
 Light1.state = 0
 Light2.state = 0
 Light3.state = 0
 Light4.state = 1
 Light5.state = 0
 Light1a.state = 0
 Light2a.state = 0
 Light3a.state = 0
 Light4a.state = 1
 Light5a.state = 0
 Kickto4
 Case 5
 Light1.state = 0
 Light2.state = 0
 Light3.state = 0
 Light4.state = 0
 Light5.state = 1
 Light1a.state = 0
 Light2a.state = 0
 Light3a.state = 0
 Light4a.state = 0
 Light5a.state = 1
 Kickto5
 Case 6
 Light1.state = 0
 Light2.state = 0
 Light3.state = 0
 Light4.state = 0
 Light5.state = 1
 Light1a.state = 0
 Light2a.state = 0
 Light3a.state = 0
 Light4a.state = 0
 Light5a.state = 1
 Kickto5
 End select
 End sub


Sub B2SRPGKick(b2si)
	If B2SOn Then ResetB2SGun: Controller.B2SSetData b2si-31, 1: Controller.B2SSetData b2si, 1
End Sub

 Sub Kickto1()

 If T3.ShowDT = True Then
 Set rpgball= bb1.Createsizedball(10):rpgball.image = "rpgball":rpgball.id = 666:bb1.kick 270,9
K2.enabled=false
K1.kick 270, 50, 27
GW.isdropped=0
 playsound "rpgshot"
 ballinbbk=false
	b2skicknr = 155: B2SRPGKick(b2skicknr)
Else
 Set rpgball= bb1a.Createsizedball(12):rpgball.image = "rpgball":rpgball.id = 666:bb1a.kick 275,16
 playsound "rpgshot"
 ballinbbk=false
	b2skicknr = 155: B2SRPGKick(b2skicknr)
END If
 End Sub




 Sub Kickto2()
 If T3.ShowDT = True Then
 Set rpgball= bb2.Createsizedball(10):rpgball.image = "rpgball":rpgball.id = 666: bb2.kick 270,8
K2.enabled=false
K1.kick 270, 45, 42
GW.isdropped=0
 playsound "rpgshot"
 ballinbbk=false
	b2skicknr = 154: B2SRPGKick(b2skicknr)
Else
 Set rpgball= bb2a.Createsizedball(12):rpgball.image = "rpgball":rpgball.id = 666: bb2a.kick 295,16
 playsound "rpgshot"
 ballinbbk=false
	b2skicknr = 154: B2SRPGKick(b2skicknr)
END If
 End Sub
 
 Sub Kickto3()
 If T3.ShowDT = True Then
 Set rpgball= bb3.Createsizedball(10):rpgball.image = "rpgball":rpgball.id = 666: bb3.kick 270,7
K2.enabled=false
K1.kick 270, 50, 47
GW.isdropped=0
 playsound "rpgshot"
 ballinbbk=false
	b2skicknr = 153: B2SRPGKick(b2skicknr)
Else
 Set rpgball= bb3a.Createsizedball(12):rpgball.image = "rpgball":rpgball.id = 666: bb3a.kick 300,16
 playsound "rpgshot"
 ballinbbk=false
	b2skicknr = 153: B2SRPGKick(b2skicknr)
END If
 End Sub
 
 Sub Kickto4()
 If T3.ShowDT = True Then
 Set rpgball= bb4.Createsizedball(10):rpgball.image = "rpgball":rpgball.id = 666: bb4.kick 270,6
K2.enabled=false
K1.kick 270, 48, 59
GW.isdropped=0
 playsound "rpgshot"
 ballinbbk=false
	b2skicknr = 152: B2SRPGKick(b2skicknr)
Else
 Set rpgball= bb4a.Createsizedball(12):rpgball.image = "rpgball":rpgball.id = 666: bb4a.kick 305,16
 playsound "rpgshot"
 ballinbbk=false
	b2skicknr = 152: B2SRPGKick(b2skicknr)
END If
 End Sub
 
 Sub Kickto5()
 If T3.ShowDT = True Then
 Set rpgball= bb5.Createsizedball(10):rpgball.image = "rpgball":rpgball.id = 666: bb5.kick 270,5
K2.enabled=false
K1.kick 270, 45, 66
GW.isdropped=0
 playsound "rpgshot"
 ballinbbk=false
	b2skicknr = 151: B2SRPGKick(b2skicknr)
Else
 Set rpgball= bb5a.Createsizedball(12):rpgball.image = "rpgball":rpgball.id = 666: bb5a.kick 325,18
 playsound "rpgshot"
 ballinbbk=false
	b2skicknr = 151: B2SRPGKick(b2skicknr)
END If
 End Sub


Sub Trigger1_Hit
K2.enabled=True
End Sub
 
 Sub rpgtrig_hit()
 If ActiveBall.Velx >-1 Then 
 ActiveBall.VelX = 0
 End If
 End Sub

 Sub sw41_Hit:vpmTimer.PulseSw 41:PlaySound "target":End Sub
 
 Sub sw42_Hit:vpmTimer.PulseSw 42:PlaySound "target":End Sub
 
 Sub sw43_Hit:vpmTimer.PulseSw 43:PlaySound "target":End Sub
 
 Sub sw44_Hit:vpmTimer.PulseSw 44:PlaySound "target":End Sub
  
 Sub sw45_Hit:vpmTimer.PulseSw 45:PlaySound "target":End Sub

 Sub sw41a_Hit:vpmTimer.PulseSw 41:PlaySound "target":End Sub
 
 Sub sw42a_Hit:vpmTimer.PulseSw 42:PlaySound "target":End Sub
 
 Sub sw43a_Hit:vpmTimer.PulseSw 43:PlaySound "target":End Sub
 
 Sub sw44a_Hit:vpmTimer.PulseSw 44:PlaySound "target":End Sub
  
 Sub sw45a_Hit:vpmTimer.PulseSw 45:PlaySound "target":End Sub

 Sub bb0a_hit():me.destroyball:ballinbbk=true:rpgreel.setvalue 0:RPGReset = False: End Sub
 
 Sub bb0_hit():me.destroyball:rpgreel.setvalue 0:b2skicknr = 0: K2.destroyball:K1.createball:ballinbbk=true:GW.isdropped=1:playsound "target":End Sub
 Sub bb6_hit():me.destroyball:rpgreel.setvalue 0:b2skicknr = 0: K2.destroyball:K1.createball:ballinbbk=true:GW.isdropped=1:playsound "target":End Sub
 Sub bb7_hit():me.destroyball:rpgreel.setvalue 0:b2skicknr = 0: K2.destroyball:K1.createball:ballinbbk=true:GW.isdropped=1:playsound "target":End Sub
 Sub bb8_hit():me.destroyball:rpgreel.setvalue 0:b2skicknr = 0: K2.destroyball:K1.createball:ballinbbk=true:GW.isdropped=1:playsound "target":End Sub
 Sub bb9_hit():me.destroyball:rpgreel.setvalue 0:b2skicknr = 0: K2.destroyball:K1.createball:ballinbbk=true:GW.isdropped=1:playsound "target":End Sub
 Sub bb10_hit():me.destroyball:rpgreel.setvalue 0:b2skicknr = 0: K2.destroyball:K1.createball:ballinbbk=true:GW.isdropped=1:playsound "target":End Sub


'******************************* Real RPG Kanone by Hanibal *************************************************************************

DIM RPGschritt,RPGfeuer, RPGReady,RPGReset


Sub RPGKanone_Timer ()

If RPGrichtung.objrotz > 45 Then
RPGschritt = -1.5
End If

If RPGrichtung.objrotz < 1 and RPGrichtung.visible = 0 Then
RPGschritt = 0
StopSound "motor"
End If


If RPGrichtung.objrotz < 1  Then
IF RPGSpezial = False Then
StopSound "motor"
RPGKanone.enabled=False
End If
End If

If RPGrichtung.objrotz < 1 And RPGReset = False Then
RPGschritt = 1
End If

If RPGrichtung.objrotz < 6 Then
If RPGReady = True Then



RPGStopp.isdropped = 1
End If

End If



RPGrichtung.objrotz = RPGrichtung.objrotz + RPGschritt
RPGrichtung1.objrotz = RPGrichtung1.objrotz + RPGschritt
RPGfeuer = RPGrichtung.objrotz +270

If RPGschritt < 0 Then
 playsound "motor", 0, 10 / (5*Rnd), -0.15, 0.15
End If

If RPGschritt > 0 Then
 playsound "motor", 0, 10 / (5*Rnd), -0.15, 0.15
End If


'****************************** Module for Backglass


	If B2SOn Then
		ResetB2SGun:

IF RPGrichtung.visible = 0 Then

If RPGrichtung.objrotz = 0 Then Controller.B2SSetData 124, 1
If RPGrichtung.objrotz > 1 AND RPGrichtung.objrotz < 9 Then Controller.B2SSetData 124, 1
If RPGrichtung.objrotz > 8 AND RPGrichtung.objrotz < 17 Then Controller.B2SSetData 123, 1
If RPGrichtung.objrotz > 16 AND RPGrichtung.objrotz < 25 Then Controller.B2SSetData 122, 1
If RPGrichtung.objrotz > 24 AND RPGrichtung.objrotz < 37 Then Controller.B2SSetData 121, 1
If RPGrichtung.objrotz > 36 AND RPGrichtung.objrotz < 43 Then Controller.B2SSetData 120, 1
If RPGrichtung.objrotz > 42 Then Controller.B2SSetData 120, 1

Else

If RPGrichtung.objrotz = 0 Then Controller.B2SSetData 129, 1
If RPGrichtung.objrotz > 1 AND RPGrichtung.objrotz < 9 Then Controller.B2SSetData 129, 1
If RPGrichtung.objrotz > 8 AND RPGrichtung.objrotz < 17 Then Controller.B2SSetData 128, 1
If RPGrichtung.objrotz > 16 AND RPGrichtung.objrotz < 25 Then Controller.B2SSetData 127, 1
If RPGrichtung.objrotz > 24 AND RPGrichtung.objrotz < 37 Then Controller.B2SSetData 126, 1
If RPGrichtung.objrotz > 36 AND RPGrichtung.objrotz < 43 Then Controller.B2SSetData 125, 1
If RPGrichtung.objrotz > 42 Then Controller.B2SSetData 125, 1

End If


	End If






End Sub



Sub RPGStopp_hit: RPGReady = True: End Sub



Sub RPGTigger_Hit: RPGStopp.isdropped = 0: End Sub

Sub RPGTigger1_Hit
If RPGReady = True Then
RPGReady = False
RPGrichtung.visible = 1
RPGrichtung1.visible = 0 
End If
End Sub


'RPG BALL FOR DESKTOP BACKGROUND*************************************************************************
Sub RPGBallS_Timer()
If ballinbbk=true then 
K1.destroyball
K1.createball
End If
RPGBallS.Enabled=True
End Sub
 
 
'KICKBACK HANDLER *****************************************************************************************************************************************

 Sub Sol12(enabled)
 If enabled Then
 Kickramp.collidable = 1
 Kickback.fire
 Playsound "holeexit"
 Else
 Kickback.pullback
 Kickramp.collidable = 0
 End If
 End Sub
 
'RIGHT VUK HANDLER*****************************************************************************************************************************************

 Sub Tophole_hit():bsvuk.addball me:end sub
 
'LEFT VUK HANDLER******************************************************************************************************************************************

 Sub txshot_hit():vpmtimer.pulseswitch 31,0,0:bstxhole.addball me:Playsound "VUKEnter":End Sub
 Sub txshota_hit():bstxhole.addball me:Playsound "VUKEnter":End Sub
 Sub txshotb_hit():bstxhole.addball me:Playsound "VUKEnter":End Sub
 Sub txshotc_hit():bstxhole.addball me:Playsound "VUKEnter":End Sub
 Sub txshotd_hit():bstxhole.addball me:Playsound "VUKEnter":End Sub
 
'LEFT UPPER POST HANDLER **********************************************************************************************************************************


Sub KickHelpOn_hit
Kicker1.enabled=True
End Sub

Sub KickHelpOff_hit
Kicker1.enabled=False
End Sub

Sub Kicker1_hit
kicker1.kick 100,activeball.velx*2.5
End Sub

 Sub sol22(enabled)
 If enabled Then
 LeftPost.isdropped = 0
 LeftPost2.Z=0
 PlaySound "PlastikHit"
 Else
 LeftPost.isdropped = 1
 LeftPost2.Z=-30
 PlaySound "PlastikHit"
 End If
 End Sub
 
'CENTER UPPER POST HANDLER ********************************************************************************************************************************

 Sub sol23(enabled)
 If enabled then
 CenterPost.isdropped = 0
 CenterPost2.Z=0
 PlaySound "PlastikHit"
 Else
 CenterPost.isdropped = 1
 CenterPost2.Z=-30
 PlaySound "PlastikHit"
 End If
 End Sub

'BACKBOX MOTOR ********************************************************************************************************************************************

 Sub sol20(enabled)
 If enabled Then
 If T3.ShowDT = True Then
 sol20timer.enabled=true
 End If

 If T3.ShowDT = False Then
 RPGSpezial = True
 RPGKanone.enabled=true

 End If

 Else

 sol20timer.enabled=false
 RPGSpezial = False

 Light1.state = 0
 Light2.state = 0
 Light3.state = 0 		
 Light4.state = 0
 Light5.state = 0
 Light1a.state = 0
 Light2a.state = 0
 Light3a.state = 0 		
 Light4a.state = 0
 Light5a.state = 0

 If B2SOn Then 
 ResetB2SGun: 
 End If

 End If
 End Sub
 
'BACKPANEL DIVERTER

 Sub sol21(enabled)
 If enabled Then 
 bwdiverter1.isdropped = 1
 bwdiverter1a.isdropped = 0
 PlaySound "PlastikHit"
 Else
 bwdiverter1.isdropped = 0
 bwdiverter1a.isdropped = 1
 PlaySound "PlastikHit"
 End If
 End Sub
 
'BACKBOX KICKER

 Sub sol5(enabled)
 If enabled Then
  If T3.ShowDT = True Then
  If ballinbbk=true Then kickit
  Else

If RPGrichtung.visible = True And RPGReset = False Then 


Set rpgball= RPGreal.Createsizedball(9):rpgball.image = "rpgball":rpgball.id = 666: RPGreal.kick RPGfeuer,16: RPGrichtung.visible = 0: RPGrichtung1.visible = 1: RPGReset = True: RPGschritt = -1.5
 playsound "rpgshot"
If RPGrichtung.objrotz > 0 and RPGrichtung.objrotz < 16 Then b2skicknr = 155: B2SRPGKick(b2skicknr)
If RPGrichtung.objrotz > 15 and RPGrichtung.objrotz < 25 Then b2skicknr = 154: B2SRPGKick(b2skicknr)
If RPGrichtung.objrotz > 24 and RPGrichtung.objrotz < 37 Then b2skicknr = 153: B2SRPGKick(b2skicknr)
If RPGrichtung.objrotz > 36 and RPGrichtung.objrotz < 43 Then b2skicknr = 152: B2SRPGKick(b2skicknr)
If RPGrichtung.objrotz > 42 Then b2skicknr = 151: B2SRPGKick(b2skicknr)

End If

 End If

 End If
 End sub
 

 
'BUMPERS
 
 Sub LBumper_Hit:vpmTimer.PulseSw 49:PlaySound "BumperMiddle":Bumperlicht1.state = 1:Me.TimerEnabled = 1:End Sub
 Sub LBumper_Timer():Bumperlicht1.state = 0:Me.TimerEnabled = 0:End Sub


 
 Sub RBumper_Hit:vpmTimer.PulseSw 50:PlaySound "BumperMiddle":Bumperlicht2.state = 1:Me.TimerEnabled = 1:End Sub
 Sub RBumper_Timer():Bumperlicht2.state = 0:Me.TimerEnabled = 0:End Sub

 
 Sub BBumper_Hit:vpmTimer.PulseSw 51:PlaySound "BumperMiddle":Bumperlicht3.state = 1:Me.TimerEnabled = 1:End Sub
 Sub BBumper_Timer():Bumperlicht3.state = 0:Me.TimerEnabled = 0:End Sub

 

 'Switches

 Sub sw24_Hit:Controller.Switch(24) = 1:PlaySound "sensor":End Sub
 Sub sw24_Unhit:Controller.Switch(24) = 0:End Sub
 
 Sub sw28_Hit:Controller.Switch(28) = 1:PlaySound "sensor":End Sub
 Sub sw28_Unhit:Controller.Switch(28) = 0:End Sub
 
 Sub sw29_Hit:Controller.Switch(29) = 1:PlaySound "sensor":End Sub
 Sub sw29_Unhit:Controller.Switch(29) = 0:End Sub
 
 Sub sw30_Hit:Controller.Switch(30) = 1:PlaySound "sensor":End Sub
 Sub sw30_Unhit:Controller.Switch(30) = 0:End Sub
 
 Sub sw32_Hit:Controller.Switch(32) = 1:PlaySound "sensor":End Sub
 Sub sw32_Unhit:Controller.Switch(32) = 0:End Sub
 
 Sub sw37_Hit:Controller.Switch(37) = 1:PlaySound "sensor":End Sub
 Sub sw37_Unhit:Controller.Switch(37) = 0:End Sub
 
 Sub sw38_Hit:Controller.Switch(38) = 1:PlaySound "sensor":End Sub
 Sub sw38_Unhit:Controller.Switch(38) = 0:End Sub
 
 Sub sw39_Hit:Controller.Switch(39) = 1:PlaySound "sensor":End Sub
 Sub sw39_Unhit:Controller.Switch(39) = 0:End Sub
 
 Sub sw40_Hit:Controller.Switch(40) = 1:PlaySound "sensor":End Sub
 Sub sw40_Unhit:Controller.Switch(40) = 0:End Sub
 
 Sub sw46_Hit:Controller.Switch(46) = 1:PlaySound "sensor":End Sub
 Sub sw46_Unhit:Controller.Switch(46) = 0:End Sub
 
 Sub sw52_Hit:Controller.Switch(52) = 1:PlaySound "sensor":End Sub
 Sub sw52_Unhit:Controller.Switch(52) = 0:End Sub 
 
 Sub sw57_Hit:Controller.Switch(57) = 1:PlaySound "sensor":End Sub
 Sub sw57_Unhit:Controller.Switch(57) = 0:End Sub
 
 Sub sw58_Hit:Controller.Switch(58) = 1:PlaySound "sensor":End Sub
 Sub sw58_Unhit:Controller.Switch(58) = 0:End Sub
 
 Sub sw60_Hit:Controller.Switch(60) = 1:PlaySound "sensor":End Sub
 Sub sw60_Unhit:Controller.Switch(60) = 0:End Sub
 
 Sub sw61_Hit:Controller.Switch(61) = 1:PlaySound "sensor":End Sub
 Sub sw61_Unhit:Controller.Switch(61) = 0:End Sub
 
'*************************************Ziele*************************************************************

 Sub sw10_Hit:vpmTimer.PulseSw 10:PlaySound "target":End Sub

 
 Sub sw17_Hit:vpmTimer.PulseSw 17:PlaySound "target":Switchlicht6.state = 1 :Me.TimerEnabled = 1:End Sub
 Sub sw17_Timer:Switchlicht6.state = 0 :Me.TimerEnabled = 0:End Sub
  
 Sub sw18_Hit:vpmTimer.PulseSw 18:PlaySound "target":Switchlicht5.state = 1 :Me.TimerEnabled = 1:End Sub
 Sub sw18_Timer:Switchlicht5.state = 0 :Me.TimerEnabled = 0:End Sub
  
 Sub sw19_Hit:vpmTimer.PulseSw 19:PlaySound "target":Switchlicht4.state = 1 :Me.TimerEnabled = 1:End Sub
 Sub sw19_Timer:Switchlicht4.state = 0 :Me.TimerEnabled = 0:End Sub

 
 Sub sw20_Hit:vpmTimer.PulseSw 20:PlaySound "target":Switchlicht1.state = 1 :Me.TimerEnabled = 1:End Sub
 Sub sw20_Timer:Switchlicht1.state = 0 :Me.TimerEnabled = 0:End Sub
 
 Sub sw21_Hit:vpmTimer.PulseSw 21:PlaySound "target":Switchlicht2.state = 1 :Me.TimerEnabled = 1:End Sub
 Sub sw21_Timer:Switchlicht2.state = 0 :Me.TimerEnabled = 0:End Sub
 
 Sub sw22_Hit:vpmTimer.PulseSw 22:PlaySound "target":Switchlicht3.state = 1 :Me.TimerEnabled = 1:End Sub
 Sub sw22_Timer:Switchlicht3.state = 0 :Me.TimerEnabled = 0:End Sub
 
 Sub sw23_Hit:vpmTimer.PulseSw 23:PlaySound "target":End Sub
 
 Sub sw25_hit():dtdroptarget.hit 1:PlaySound "target":End sub
 
 'RAMP HELPERS *********************************************************************************************************************************************

  Sub Launch1_Hit()
 Launch1.DestroyBall
 Launch1a.CreateBall
 Launch1a.kick 270,10
 End Sub 
 
  Sub RHelp1_Hit()
 ActiveBall.VelZ = -2
 ActiveBall.VelY = 0
 ActiveBall.VelX = 0
 StopSound "metalrolling"
 PlaySound "ballhit"
 End Sub
 
 Sub RHelp2_Hit()
 ActiveBall.VelZ = -2
 ActiveBall.VelY = 0
 ActiveBall.VelX = 0
 StopSound "metalrolling"
 PlaySound "ballhit"
 End Sub
 
 Sub RHelp3_Hit()
 Playsound "metalhit2", 1, 5, 0.15
 ActiveBall.VelZ = -2
 ActiveBall.VelY = 0
 ActiveBall.VelX = 0
 End Sub
 
 Sub RHelp4_Hit()
 Playsound "metalhit2", 1, 5, -0.15
 ActiveBall.VelZ = -2
 ActiveBall.VelY = 0
 ActiveBall.VelX = 0
 End Sub
 
 Sub RHelp5_Hit()
 ActiveBall.VelZ = -2
 ActiveBall.VelY = 0
 ActiveBall.VelX = 0
 End Sub
 
 Sub RHelp6_Hit()
 ActiveBall.VelX = 20
 StopSound "metalrolling"
 PlaySound ""
 End Sub
 
 'SOUND EFFECTS ********************************************************************************************************************************************

 Sub Strig_Hit()
 Playsound "metalrolling"
 End Sub
 
 Sub Strig2_Hit()
 Playsound "metalrolling"
 End Sub
 
 Sub Strig3_Hit()
 Playsound "metalrolling"
 End Sub
 
 Sub Strig4_Hit()
 Playsound "metalrolling"
 End Sub
 
  Sub Post1_Hit()
  Playsound "rubber"
  End Sub
 
 'RULES ***************************************************************************************************************************************************

 Dim RuleWindow
 RuleWindow = 0
 Sub Rules
 If RuleWindow = 0 Then
 Dim objShell:Set objShell = CreateObject("Wscript.Shell")
 objShell.Run "Rules-T3.exe"
 RuleWindow = 1
 Else
 RuleWindow = 0
 End If
 End Sub
 


'**************		Flasher	*****************

Sub Flash4(enabled)
	If Enabled Then	
		RPGFLASH4.State = 1
		RPGFLASH4a1.State = 1
		RPGFLASH4a2.State = 1
		RPGFLASH4a3.State = 1
		RPGFLASH4a.State = 1
		RPGFLASHeye.State = 1
	Else
		RPGFLASH4.State = 0
		RPGFLASH4a1.State = 0
		RPGFLASH4a2.State = 0
		RPGFLASH4a3.State = 0
		RPGFLASH4a.State = 0
		RPGFLASHeye.State = 0

	End If
End Sub

Sub Flash26(enabled)
	If Enabled Then	
		Flasher26a.State = 1
		Flasher26b.State = 1
		Flasher26c.State = 1

	Else
		Flasher26a.State = 0
		Flasher26b.State = 0
		Flasher26c.State = 0


	End If
End Sub

Sub Flash27(enabled)
	If Enabled Then	
		RPGFLASH27a.State = 1
		RPGFLASH27a1.State = 1
RPGFLASH27a2.State = 1
RPGFLASH27a3.State = 1
RPGFLASH27a4.State = 1
RPGFLASH27a5.State = 1
RPGFLASH27a6.State = 1
	Else
		RPGFLASH27a.State = 0
		RPGFLASH27a1.State = 0
RPGFLASH27a2.State = 0
RPGFLASH27a3.State = 0
RPGFLASH27a4.State = 0
RPGFLASH27a5.State = 0
RPGFLASH27a6.State = 0
	End If
End Sub

Sub Flash28(enabled)
	If Enabled Then	
		RPGFLASH28a.State = 1
		RPGFLASH28a1.State = 1

	Else
		RPGFLASH28a.State = 0
		RPGFLASH28a1.State = 0

	End If
End Sub



Sub Flash30(enabled)
	If Enabled Then	
		Flasher30a.State = 1
		Flasher30a1.State = 1
		Flasher30b.State = 1
		Flasher30b1.State = 1
		Flasher30c.State = 1
		Flasher30c1.State = 1
		Flasher30d.State = 1
		Flasher30d1.State = 1

		Flasher30p1.Image = "DomeOn"
		Flasher30p2.Image = "DomeOn"
		Flasher30p3.Image = "DomeOn"
		Flasher30p4.Image = "DomeOn"
	Else
		Flasher30a.State = 0
		Flasher30a1.State = 0
		Flasher30b.State = 0
		Flasher30b1.State = 0
		Flasher30c.State = 0
		Flasher30c1.State = 0
		Flasher30d.State = 0
		Flasher30d1.State = 0

		Flasher30p1.Image = "DomeOff"
		Flasher30p2.Image = "DomeOff"
		Flasher30p3.Image = "DomeOff"
		Flasher30p4.Image = "DomeOff"
	End If
End Sub



Sub Flash31(enabled)
	If Enabled Then	
		Flasher31a.State = 1
		Flasher31a1.State = 1
		Flasher31b.State = 1
		Flasher31b1.State = 1
		Flasher31p1.Image = "DomeOn"
		Flasher31p2.Image = "DomeOn"
	Else
		Flasher31a.State = 0
		Flasher31a1.State = 0
		Flasher31b.state = 0
		Flasher31b1.state = 0
		Flasher31p1.Image = "DomeOff"
		Flasher31p2.Image = "DomeOff"
	End If
End Sub 



Sub Flash32(enabled)
	If Enabled Then	
		Flasher32a.State = 1
		Flasher32a1.State = 1
		Flasher32b.State = 1
		Flasher32b1.State = 1
		Flasher32p1.Image = "DomeOn"
		Flasher32p2.Image = "DomeOn"
	Else
		Flasher32a.State = 0
		Flasher32a1.State = 0
		Flasher32b.state = 0
		Flasher32b1.state = 0
		Flasher32p1.Image = "DomeOff"
		Flasher32p2.Image = "DomeOff"
	End If
End Sub
 

 
'**************		Light Inserts	*****************

On Error Resume Next
Dim i

For i=0 To 127
	If IsObject(eval("L" & i)) Then
    	Execute "Set Lights(" & i & ")  = L" & i
	End If
Next





'*************************
' GI - needs new vpinmame
'*************************

Set GICallback = GetRef("GIUpdate")

Sub GIUpdate(no, Enabled)
Dim x
    For each x in GIGruppe
        x.State = ABS(Enabled)
    Next


End Sub


'******  Hanibal's Special Flashers


Dim FlashState(200), FlashLevel(200)

Sub Kombiflasher(nr, a, object)
DIM Helfer
    IF a.state = 0 Then 
	Helfer = 0
    Else
    Helfer = 1
    End If

    Select Case Helfer
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - (20)
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
            Object.opacity = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + (30)
            If FlashLevel(nr) > 1200 Then
                FlashLevel(nr) = 1200
                FlashState(nr) = -2 'completely on
            End if
            Object.opacity = FlashLevel(nr)
    End Select
End Sub


Sub Kombiflasher2(nr, a, object,b)
DIM Helfer
    IF a.state = 0 Then 
	Helfer = 0
    Else
    Helfer = 1
    End If
	b.state = Helfer
    Select Case Helfer
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - (20)
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
            Object.opacity = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + (30)
            If FlashLevel(nr) > 1200 Then
                FlashLevel(nr) = 1200
                FlashState(nr) = -2 'completely on
            End if
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

Sub Kombilicht(nr, a, b)
DIM Helfer
    IF a.state = 0 Then 
	Helfer = 0
    Else
    Helfer = 1
    End If
	b.state = Helfer
End Sub

Sub Kombilichtb(nr, a, b)
DIM Helfer
    IF a.state = 0 Then 
	Helfer = 1
    Else
    Helfer = 0
    End If
	b.state = Helfer
End Sub


'Flasher Updates


Sub Flasherupdate_timer

Kombiflasher 119, Flasher30a, Flasherr30a
Kombiflasher 120, Flasher30b, Flasherr30b
Kombiflasher 121, Flasher30c, Flasherr30c
Kombiflasher 122, Flasher30d, Flasherr30d
Kombiflasher 123, Flasher31a, Flasherr31a
Kombiflasher 124, Flasher31b, Flasherr31b
Kombiflasher 125, Flasher32a, Flasherr32a
Kombiflasher 126, Flasher32b, Flasherr32b
Kombiflasher 127, Flasher26a, Flasherr26a
Kombiflasher 137, Flasher26b, Flasherr26b


Kombiflasher 128, GL5, FlasherGLLED1
Kombiflasher 129, GL4, FlasherGLLED2
Kombiflasher 130, GL3, FlasherGLLED3
Kombiflasher 131, GL2, FlasherGLLED4
Kombiflasher 132, GL1, FlasherGLLED5
Kombiflasher 133, GL60, FlasherGLLED6

Kombiflasher 134, l71, FlasherTLicht
Kombiflasher 135, l64, FlasherTx1
Kombiflasher2 136, l64, FlasherTx2,l64a


Kombilicht 1, l1,l1a
Kombilicht 2, l2,l2a
Kombilicht 3, l3,l3a
Kombilicht 4, l4,l4a
Kombilicht 5, l5,l5a
Kombilicht 6, l6,l6a
Kombilicht 7, l7,l7a
Kombilicht 8, l8,l8a
Kombilicht 9, l9,l9a
Kombilicht 10, l10,l10a
Kombilicht 11, l11,l11a
Kombilicht 12, l12,l12a
Kombilicht 13, l13,l13a
Kombilicht 14, l14,l14a
Kombilicht 15, l15,l15a
Kombilicht 16, l16,l16a
Kombilicht 17, l17,l17a
Kombilicht 18, l18,l18a
Kombilicht 19, l19,l19a
Kombilicht 20, l20,l20a
Kombilicht 21, l21,l21a
Kombilicht 22, l22,l22a
Kombilicht 23, l23,l23a
Kombilicht 24, l24,l24a
Kombilicht 25, l25,l25a
Kombilicht 26, l26,l26a
Kombilicht 27, l27,l27a
Kombilicht 28, l28,l28a
Kombilicht 29, l29,l29a
Kombilicht 30, l30,l30a
Kombilicht 31, l31,l31a
Kombilicht 32, l32,l32a
Kombilicht 33, l33,l33a
Kombilicht 34, l34,l34a
Kombilicht 35, l35,l35a
Kombilicht 36, l36,l36a
Kombilicht 37, l37,l37a
Kombilicht 38, l38,l38a
Kombilicht 39, l39,l39a
Kombilicht 40, l40,l40a
Kombilicht 41, l41,l41a
Kombilicht 42, l42,l42a
Kombilicht 43, l43,l43a
Kombilicht 44, l44,l44a
Kombilicht 45, l45,l45a
Kombilicht 46, l46,l46a
Kombilicht 47, l47,l47a
Kombilicht 48, l48,l48a
Kombilicht 49, l49,l49a
Kombilicht 50, l50,l50a
Kombilicht 51, l51,l51a
Kombilicht 52, l52,l52a
Kombilicht 55, l55,l55a
Kombilicht 56, l56,l56a
Kombilicht 57, l57,l57a
Kombilicht 58, l58,l58a
Kombilicht 59, l59,l59a
Kombilicht 60, l60,l60a
Kombilicht 61, l61,l61a
Kombilicht 62, l62,l62a

Kombilicht 65, l65,l65a
Kombilicht 66, l66,l66a
Kombilicht 67, l67,l67a
Kombilicht 68, l68,l68a
Kombilicht 69, l69,l69a


Kombilicht 72, l72,l72a








Kombilicht 38, l38,l38a
Kombilicht 39, l39,l39a
Kombilicht 40, l40,l40a



Kombilicht 80, l80,l80a

End Sub 
 



 Sub Augen_Timer

' ******Hanibals Random Lights Script
l71.Intensity = (50+(20*Rnd))
FlasherTLicht.amount = l71.Intensity *1.5

l64.Intensity = (50+(20*Rnd))
l64a.Intensity = l64.Intensity
FlasherTx1.amount = l64.Intensity /2
FlasherTx2.amount = l64.Intensity /2

Bumperlicht1.Intensity = (20+(5*Rnd))
Bumperlicht2.Intensity = Bumperlicht1.Intensity
Bumperlicht3.Intensity = Bumperlicht1.Intensity
Cardlight1.Intensity = (20+(2*Rnd))
Cardlight2.Intensity = Cardlight1.Intensity
Flasher30a1.Intensity = (10+(3*Rnd))
Flasher30b1.Intensity = Flasher30a1.Intensity
Flasher30c1.Intensity = Flasher30a1.Intensity
Flasher30d1.Intensity = Flasher30a1.Intensity
Flasherr30a.amount = Flasher30a1.Intensity *4
Flasherr30b.amount = Flasher30a1.Intensity *4
Flasherr30c.amount = Flasher30a1.Intensity *4
Flasherr30d.amount = Flasher30a1.Intensity *4


Flasher26a.Intensity = (30+(10*Rnd))
Flasher26b.Intensity = Flasher26a.Intensity
Flasher26c.Intensity = Flasher26a.Intensity
Flasherr26b.amount = Flasher26a.Intensity *4

GLLampe1.Intensity = (20+(2*Rnd))
GLLampe2.Intensity = GLLampe1.Intensity
GLLampe3.Intensity = GLLampe1.Intensity
GLLampe4.Intensity = GLLampe1.Intensity

GLLampe35.Intensity = (60+(5*Rnd))
GLLampe5.Intensity = GLLampe35.Intensity
GLLampe6.Intensity = GLLampe35.Intensity
GLLampe7.Intensity = GLLampe35.Intensity


 End Sub



'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub


'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 10 ' total number of balls
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z > 0 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next
End Sub




' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub


Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub


Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub


'Sub LRRail_Hit:PlaySound "fx_metalrolling", 0, 150, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0:End Sub
'
'Sub RLRail_Hit:PlaySound "fx_metalrolling", 0, 150, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0:End Sub

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 1)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "T3" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / T3.width-1
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



 
 