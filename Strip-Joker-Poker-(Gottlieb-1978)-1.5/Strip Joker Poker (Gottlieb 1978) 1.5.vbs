'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                    +
'+             STRIP JOKER POKER - EM (Gottlieb 1978)                 +
'+           (C) 2022 NH 80's Retro Entertainment Group               +
'+                                                                    +
'+                                                                    +
'+	         Original Joker Poker VPX by BorgDog, 2015                +
'+                                                                    +
'+  Graphics+ Update and Strip Joker Poker Pup-Pack by HiRez00 - 2022 +
'+                 Original Backglass by: HiRez00                     +
'+      Additional Pup-Pack Support and Testing by PinballFan2018     +
'+          Physics, Sounds, Shadows Update by Apophis - 2022         +
'+                                                                    +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


Option Explicit
Randomize

'+++++  PLEASE NOTE: THIS IS A 1 (SINNGLE) PLAYER GAME  +++++
'
'+++++++++++++++++++++++
'+     TABLE NOTES     +
'+++++++++++++++++++++++
'
' Special Pup-Pack "Strip Videos" will play for every 10,000 point awarded. All Characters have 10 items of clothing to remove.
'
' There and 13 LUT brightness levels built into this table.
' Hold down LEFT MAGNASAVE and then press RIGHT MAGNASAVE to adjust / cycle through the different LUT brightness levels.
' The LUT you choose will be automatically saved when you exit the table.
'
'
' OPTIONS MENU: Hold doown the LEFT FLIPPER for 2-3 seconds BEFORE you start a new game to access the table's OPTIONS MENU to
' select Apron Color, 3 or 5 Ball Play, Free Play and Replay Options. 
'
' SPECIAL WHEN LIT IS ONLY ACTIVE IF YOU ARE PLAYING THE (3) BALL VERSION - NOT THE (5) BALL.
'
'+++++++++++++++++++++++
'+    TABLE OPTIONS    +
'+++++++++++++++++++++++
'--------------------------------------------------------------------------------------------------------------------
BallRollingVolume = .8 	'Ball rolling sound volume number range from 0 to 1 where: 1 = Loudest  0 = Silent
'--------------------------------------------------------------------------------------------------------------------
VolumeDial = .6			'Mechanical sound effects volume. Recommended values should be no greater than 1.
'--------------------------------------------------------------------------------------------------------------------
AddFlipperShadows = 1    	'1 = Flipper Shadows ON  0 = Flipper Shawdows OFF - Turn OFF if using Ambient Occlusion
'--------------------------------------------------------------------------------------------------------------------
DynamicBallShadowsOn = 0	'0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
AmbientBallShadowOn = 0		'0 = Static shadow under ball ("flasher" image, like JP's)
							'1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
							'2 = flasher image shadow, but it moves like ninuzzu's



'****** PuP Variables ******

Dim usePUP: Dim cPuPPack: Dim PuPlayer: Dim PUPStatus: PUPStatus=false ' dont edit this line!!!

'*************************** PuP Settings for this table ********************************

usePUP   = true               ' enable Pinup Player functions for this table
cPuPPack = "st-jp-em"    ' name of the PuP-Pack / PuPVideos folder for this table

Const SongVolume = 0.1 ' 1 is full volume.

'--------------------> ' Check bonus level
Const Video1 = 10000
Const Video2 = 20000
Const Video3 = 30000
Const Video4 = 40000
Const Video5 = 50000
Const Video6 = 60000
Const Video7 = 70000
Const Video8 = 80000
Const Video9 = 90000
Const Video10 = 100000

Const cGameName = "jokerpoker_1978"
Const Ballsize = 50
Const BallMass = 1
Const HSFileName="Strip_JP_EM_1978.txt"

Dim gilvl : gilvl=0
Dim operatormenu, Options
Dim bonuscount
Dim balls, PlungeBall, rotatortemp, tempplayerup
Dim replays 
Dim Replay1Table(3), Replay2Table(3), Replay3Table(3)
dim replay1, replay2, replay3
Dim Add10, Add100, Add1000
Dim hisc
Dim maxplayers, players, player
Dim credit, freeplay
Dim score(4)
Dim sreels(4)
Dim state
Dim tilt
Dim tiltsens
Dim ballinplay
Dim matchnumb
Dim ballrenabled
Dim rstep, lstep, dtstep
Dim rep(4)
Dim rst
Dim eg
Dim scn
Dim scn1
Dim bell
Dim AltApron
Dim Strip1
Dim Strip2
Dim Strip3
Dim Strip4
Dim Strip5
Dim Strip6
Dim Strip7
Dim Strip8
Dim Strip9
Dim Strip10
Dim PlayerList
Dim VidTime
Dim IntroStart
Dim CharSelectActive
Dim CharSelectOption
Dim CharSelectHold
Dim ActiveChar
Dim PupCall
Dim CoinLockOut
Dim BallRollingVolume
Dim VolumeDial
Dim AddFlipperShadows
Dim DynamicBallShadowsOn
Dim AmbientBallShadowOn
Dim FileObj
Dim ScoreFile
Dim TextStr,TextStr2
Dim i,j, ii, objekt, light

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

sub Table1_init
	LoadEM
    CharSelectActive=False
	maxplayers=4
	Replay1Table(1)=90000
	Replay2Table(1)=130000
	Replay3Table(1)=160000
	Replay1Table(2)=110000
	Replay2Table(2)=130000
	Replay3Table(2)=190000
	Replay1Table(3)=120000
	Replay2Table(3)=150000
	Replay3Table(3)=190000
	hideoptions
    VidTime=0
    AltApron=1
	player=1
	freeplay=0
	balls=3
	replays=1
	HSA1=8
	HSA2=18
	HSA3=26
    matchnumb=100
    score(1)=000000
    score(2)=000000
    score(3)=000000
    score(4)=000000
    credit=0
	hisc=90000
    IntroStart=False
    PlayerList=0
	loadhs
    CharSelectOption=0
    CharSelectHold=1
    CoinLockOut=0
    UpdateApron
	UpdatePostIt
	gamov.text=""
	credittxt.setvalue(credit)
	Replay1=Replay1Table(Replays)
	Replay2=Replay2Table(Replays)
	Replay3=Replay3Table(Replays)
	OptionBalls.image="OptionsBalls"&Balls
	OptionReplays.image="OptionsReplays"&replays
	OptionFreeplay.image="OptionsFreeplay"&freeplay
	OptionApron.image="OptionsApron"&altapron
	RepCard.image = "ReplayCard"&replays
	if balls=3 then	
		InstCard.image="InstCard3balls"
	  else
		InstCard.image="InstCard5balls"
	end if

    If Table1.ShowDT = True Then
        for each objekt in backdropstuff : objekt.visible = 1 : next
    End If

    If Table1.ShowDT = False Then
        for each objekt in backdropstuff : objekt.visible = 0 : next
    End If

	startGame.enabled=true
	if matchnumb="" then matchnumb=100	
	if matchnumb=100 then
		matchtxt.text="00"
	  else
		matchtxt.text=matchnumb
	end if
	for i = 1 to maxplayers
		EVAL("ScoreReel"&i).setvalue(score(i))
		if score(i)>99999 and score(i)<>"" then
			EVAL("p100k"&i).text="100,000"
		  else
			EVAL("p100k"&i).text=" "
		end if
	next
	tilt=false
	If credit > 0 or freeplay = 1 Then DOF 113, DOFOn
	Drain.CreateBall
    LoadLUT
End sub

Sub GameTimer_Timer
	Cor.Update 						'update ball tracking
	RollingUpdate					'update rolling sounds
	DoDTAnim 						'handle drop target animations
End Sub

Sub FrameTimer_Timer
	If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
    FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle
End Sub


sub startGame_timer
	playsound "poweron"
	lightdelay.enabled=true
	gamov.timerenabled=1
	tilttxt.timerenabled=1
	If B2SOn Then
		if freeplay=0 then Controller.B2ssetCredits Credit
		Controller.B2ssetMatch 34, Matchnumb
		Controller.B2SSetGameOver 35,1
		Controller.B2SSetScorePlayer 5, hisc
		for i = 1 to maxplayers
			Controller.B2SSetScorePlayer i, Score(i) MOD 100000
		next
	End If
	me.enabled=false
end sub

sub lightdelay_timer
	If (Credit > 0 or freeplay=1) Then DOF 113, 1
	For each light in GIlights:light.state=1:Next
	For each light in BumperLights:light.state=1:Next
	If B2SOn then Controller.B2SSetData 99,1
	me.enabled=false
	gilvl=1
end sub


sub gamov_timer
	if state=false then
		If B2SOn then Controller.B2SSetGameOver 35,0			
		gamov.text=""
		gtimer.enabled=true
	end if
	gamov.timerenabled=0
end sub

sub gtimer_timer
    if state=false then
		gamov.text="Game Over"
		If B2SOn then Controller.B2SSetGameOver 35,1
		gamov.timerenabled=1
		gamov.timerinterval= (INT (RND*10)+5)*100
    end if
	me.enabled=0
end sub

sub tilttxt_timer
	if state=false then
		tilttxt.text=""
		If B2SOn then Controller.B2SSetTilt 33,0	
		ttimer.enabled=true
	end if
	tilttxt.timerenabled=0
end sub

sub ttimer_timer
	if state=false then
		tilttxt.text="TILT"
		If B2SOn then Controller.B2SSetTilt 33,1
		tilttxt.timerenabled=1
		tilttxt.timerinterval= (INT (RND*10)+5)*100
	end if
	me.enabled=0
end sub

Sub Table1_KeyDown(ByVal keycode)
   
	if keycode=AddCreditKey then
		Select Case Int(rnd*3)
			Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
			Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
			Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
		End Select
		coindelay.enabled=true 
    end if


'******LUT Switches ******

	If keycode = LeftMagnaSave Then bLutActive = True
	If keycode = RightMagnaSave Then 
		If bLutActive Then NextLUT: End If
    End If

'******LUT Switches End ******

    if keycode=StartGameKey and (Credit>0 or freeplay=1) and CharSelectHold=1 and CoinLockOut=0 and state=false And Not HSEnterMode=true then
        tilttxt.timerenabled=0
        gamov.timerenabled=0
        ttimer.enabled=false
        gtimer.enabled=false
        gamov.text=""
        tilttxt.text=""
        pupevent 20
        CharSelectHold=0
'		playsound "cluper"
		if freeplay=0 then
			credit=credit-1
            CoinLockOut=1
			If credit < 1 Then DOF 113, DOFOff
			credittxt.setvalue(credit)
		end if
        eg=1
        state=false
        CharSelectDelay
    End if

	If HSEnterMode Then HighScoreProcessKey(keycode)

	If keycode = PlungerKey Then
		Plunger.PullBack
		SoundPlungerPull
	End If

	If keycode=LeftFlipperKey and State = false and HSEnterMode=False and OperatorMenu=0 then
		OperatorMenuTimer.Enabled = true
	end if

	If keycode=LeftFlipperKey and State = false and HSEnterMode=False and OperatorMenu=1 then
		Options=Options+1
		If Options=6 then Options=1
		playsound "drop1"
		Select Case (Options)
			Case 1:
				Option1.visible=true
				Option5.visible=False
			Case 2:
				Option2.visible=true
				Option1.visible=False
			Case 3:
				Option3.visible=true
				Option2.visible=False
			Case 4:
				Option4.visible=true
				Option3.visible=False
			Case 5:
				Option5.visible=true
				Option4.visible=False
		End Select
	end if


'************ Select Character *************

	If keycode=RightFlipperKey and eg=1 and state=false and CharSelectHold=0 and HSEnterMode=False and OperatorMenu=0 then
		CharSelectOption=CharSelectOption+1
		If CharSelectOption>16 then CharSelectOption=1
		playsound "drop1"
		Select Case (CharSelectOption)
			Case 1:
				pupevent 651
			Case 2:
				pupevent 652
			Case 3:
				pupevent 653
			Case 4:
				pupevent 654
			Case 5:
				pupevent 655
			Case 6:
				pupevent 656
			Case 7:
				pupevent 657
			Case 8:
				pupevent 658
			Case 9:
				pupevent 659
			Case 10:
				pupevent 660
			Case 11:
				pupevent 661
			Case 12:
				pupevent 662
			Case 13:
				pupevent 663
			Case 14:
				pupevent 664
			Case 15:
				pupevent 665
			Case 16:
				pupevent 666
		End Select
	end if

If keycode=LeftFlipperKey and eg=1 and state=false and CharSelectHold=0 and HSEnterMode=False and OperatorMenu=0 then
		CharSelectOption=CharSelectOption-1
		If CharSelectOption<1 then CharSelectOption=16
		playsound "drop1"
		Select Case (CharSelectOption)
			Case 1:
				pupevent 651
			Case 2:
				pupevent 652
			Case 3:
				pupevent 653
			Case 4:
				pupevent 654
			Case 5:
				pupevent 655
			Case 6:
				pupevent 656
			Case 7:
				pupevent 657
			Case 8:
				pupevent 658
			Case 9:
				pupevent 659
			Case 10:
				pupevent 660
			Case 11:
				pupevent 661
			Case 12:
				pupevent 662
			Case 13:
				pupevent 663
			Case 14:
				pupevent 664
			Case 15:
				pupevent 665
			Case 16:
				pupevent 666
		End Select
	end if

	If keycode=StartGameKey and eg=1 and state=false and HSEnterMode=False and CharSelectHold=0 and OperatorMenu=0 then
              if CharSelectOption = 0 Then Exit Sub
			  if CharSelectOption = 1 Then ActiveChar=1: PlaySound "cluper": CharSelectOption=0: state=true: eg=0: pupevent 21: CharacterCalls: DelayedStart: Exit Sub
              if CharSelectOption = 2 Then ActiveChar=2: PlaySound "cluper": CharSelectOption=0: state=true: eg=0: pupevent 21: CharacterCalls: DelayedStart: Exit Sub
              if CharSelectOption = 3 Then ActiveChar=3: PlaySound "cluper": CharSelectOption=0: state=true: eg=0: pupevent 21: CharacterCalls: DelayedStart: Exit Sub
              if CharSelectOption = 4 Then ActiveChar=4: PlaySound "cluper": CharSelectOption=0: state=true: eg=0: pupevent 21: CharacterCalls: DelayedStart: Exit Sub
			  if CharSelectOption = 5 Then ActiveChar=5: PlaySound "cluper": CharSelectOption=0: state=true: eg=0: pupevent 21: CharacterCalls: DelayedStart: Exit Sub
              if CharSelectOption = 6 Then ActiveChar=6: PlaySound "cluper": CharSelectOption=0: state=true: eg=0: pupevent 21: CharacterCalls: DelayedStart: Exit Sub
              if CharSelectOption = 7 Then ActiveChar=7: PlaySound "cluper": CharSelectOption=0: state=true: eg=0: pupevent 21: CharacterCalls: DelayedStart: Exit Sub
              if CharSelectOption = 8 Then ActiveChar=8: PlaySound "cluper": CharSelectOption=0: state=true: eg=0: pupevent 21: CharacterCalls: DelayedStart: Exit Sub
			  if CharSelectOption = 9 Then ActiveChar=9: PlaySound "cluper": CharSelectOption=0: state=true: eg=0: pupevent 21: CharacterCalls: DelayedStart: Exit Sub
              if CharSelectOption = 10 Then ActiveChar=10: PlaySound "cluper": CharSelectOption=0: state=true: eg=0: pupevent 21: CharacterCalls: DelayedStart: Exit Sub
              if CharSelectOption = 11 Then ActiveChar=11: PlaySound "cluper": CharSelectOption=0: state=true: eg=0: pupevent 21: CharacterCalls: DelayedStart: Exit Sub
              if CharSelectOption = 12 Then ActiveChar=12: PlaySound "cluper": CharSelectOption=0: state=true: eg=0: pupevent 21: CharacterCalls: DelayedStart: Exit Sub
			  if CharSelectOption = 13 Then ActiveChar=13: PlaySound "cluper": CharSelectOption=0: state=true: eg=0: pupevent 21: CharacterCalls: DelayedStart: Exit Sub
              if CharSelectOption = 14 Then ActiveChar=14: PlaySound "cluper": CharSelectOption=0: state=true: eg=0: pupevent 21: CharacterCalls: DelayedStart: Exit Sub
              if CharSelectOption = 15 Then ActiveChar=15: PlaySound "cluper": CharSelectOption=0: state=true: eg=0: pupevent 21: CharacterCalls: DelayedStart: Exit Sub
              if CharSelectOption = 16 Then ActiveChar=16: PlaySound "cluper": CharSelectOption=0: state=true: eg=0: pupevent 21: CharacterCalls: DelayedStart: Exit Sub
	End If


	If keycode=RightFlipperKey and State = false and HSEnterMode=False and OperatorMenu=1 then
	  PlaySound "cluper"
	  Select Case (Options)
		Case 1:
			if altapron=0 Then
                altapron=1
				Wapron.Image="jp-apron2"
                Wall6.Image="jp-apron2"
		        Pcover.Image="plunger2"
                PrimWalls.Image="primwalls-black"
			 Else
                altapron=0
			    Wapron.Image="jp-apron1"
                Wall6.Image="jp-apron1"
		        Pcover.Image="plunger1"
                PrimWalls.Image="primwalls-white"
			end if
			OptionApron.image="OptionsApron"&AltApron
		Case 2:
			if Balls=3 then
				Balls=5
				InstCard.image="InstCard5balls"
			  else
				Balls=3
				InstCard.image="InstCard3balls"
			end if
			OptionBalls.image = "OptionsBalls"&Balls 
		Case 3:
			if freeplay=0 Then
				freeplay=1
			  Else
				freeplay=0
			end if
			OptionFreeplay.image="OptionsFreeplay"&freeplay
		Case 4:
			Replays=Replays+1
			if Replays>3 then
				Replays=1
			end if
			Replay1=Replay1Table(Replays)
			Replay2=Replay2Table(Replays)
			Replay3=Replay3Table(Replays)
			OptionReplays.image = "OptionsReplays"&replays
			repcard.image = "ReplayCard"&replays
		Case 5:
			OperatorMenu=0
			savehs
			HideOptions
	  End Select
	End If


  if tilt=false and state=true then
	If keycode = LeftFlipperKey Then
		FlipperActivate LeftFlipper, LFPress
		SolLFlipper True
		DOF 101,DOFOn
		PlaySound "Buzz",-1,.05,-0.05, 0.05
	End If
    
	If keycode = RightFlipperKey Then
		FlipperActivate RightFlipper, RFPress
		SolRFlipper True
		Lupflip.state=1
		DOF 102,DOFOn
		PlaySound "Buzz1",-1,.05,0.05,0.05
	End If
    
	If keycode = LeftTiltKey Then
		Nudge 90, 3
		SoundNudgeLeft
		checktilt
	End If
    
	If keycode = RightTiltKey Then
		Nudge 270, 3
		SoundNudgeRight
		checktilt
	End If
    
	If keycode = CenterTiltKey Then
		Nudge 0, 3
		SoundNudgeCenter
		checktilt
	End If

	If keycode = MechanicalTilt Then
		gametilted
	End If

  end if  
End Sub

Sub OperatorMenuTimer_Timer
	OperatorMenu=1
	Displayoptions
	Options=1
End Sub

Sub DisplayOptions
	OptionsBack.visible = true
	Option1.visible = True
    OptionApron.visible = True
	OptionBalls.visible = True
    OptionReplays.visible = True
	OptionFreeplay.visible = True
End Sub

Sub HideOptions
	for each objekt In OptionMenu
		objekt.visible = false
	next
End Sub


Sub Table1_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.Fire
		if PlungeBall=1 then
			SoundPlungerReleaseBall
		  else
			SoundPlungerReleaseNoBall
		end if
	End If
    If keycode = LeftMagnaSave Then bLutActive = False

	if keycode = LeftFlipperKey then
		OperatorMenuTimer.Enabled = false
	end if

   If tilt=false and state=true then
	If keycode = LeftFlipperKey Then
		FlipperDeActivate LeftFlipper, LFPress
		SolLFlipper False
		DOF 101,DOFOff
		StopSound "Buzz"
	End If
    
	If keycode = RightFlipperKey Then
		FlipperDeActivate RightFlipper, RFPress
		SolRFlipper False
		DOF 102,DOFOff
		Lupflip.state=0
		StopSound "Buzz1"
	End If
   End if
End Sub

sub coindelay_timer
	addcredit
    coindelay.enabled=false
end sub

sub resettimer_timer
    rst=rst+1
	for i = 1 to maxplayers
		EVAL("ScoreReel"&i).resettozero
    next 
    If B2SOn then
		for i = 1 to maxplayers
		  Controller.B2SSetScorePlayer i, score(i)
		  Controller.B2SSetScoreRollover 24 + i, 0
		next
	End If
    if rst=18 then
		newgame
		resettimer.enabled=false
    end if
end sub

Sub addcredit
	if freeplay=0 then
      credit=credit+1
	  DOF 113, DOFOn
      if credit>15 then credit=15
	  credittxt.setvalue(credit)
	  If B2SOn Then Controller.B2ssetCredits Credit
	end if
End sub

Sub Drain_Hit()
	RandomSoundDrain Drain
	DOF 128, DOFPulse
	me.timerenabled=1
End Sub

Sub Drain_timer
'	for each light in BonusLights:light.state=1: next    '***** for testing bonus scoring routine
	scorebonus.enabled=true
	me.timerenabled=0
End Sub	

sub ballhome_hit
	ballrenabled=1
	plungeball=1
end sub

sub ballhome_unhit
	DOF 129, DOFPulse
	plungeball=0
end sub

sub ballrel_hit
	if ballrenabled=1 then
		shootagain.state=lightstateoff
		If B2SOn then Controller.B2SSetShootAgain 36,0
		ballrenabled=0
	end if
end sub

sub scorebonus_timer
	if tilt=true Then
		LDoubleBonus.timerenabled=true
	  else
		bonuscount=-1
		Lbonus1.timerenabled=true
	end if
	scorebonus.enabled=false
End sub

sub Lbonus1_timer
	bonuscount=bonuscount+1
	if bonuscount>LDoubleBonus.state then
		bonuscount=-1
		Lbonus2.timerenabled=true
		Lbonus1.timerenabled=false
		Lbonus1.timerinterval=300
	 else
		if AddScore1000Timer.enabled = false then
			playsound "scorebonus"
			Lbonus1A.state=1
			if Lbonus1.state=1 then addscore 1000
		else
			bonuscount=bonuscount-1
			Lbonus1.timerinterval=100		
		end if
	end if
end sub

sub Lbonus2_timer
	bonuscount=bonuscount+1
	if bonuscount>LDoubleBonus.state then
		bonuscount=-1
		Lbonus3.timerenabled=true
		Lbonus2.timerenabled=false
		Lbonus2.timerinterval=300	
	 else
		  if AddScore1000Timer.enabled = false then
			  playsound "scorebonus"
			  Lbonus1A.state=0
			  Lbonus2A.state=1
			  if Lbonus2.state=1 then addscore 2000
		  else
			bonuscount=bonuscount-1
			Lbonus2.timerinterval=100	
		  end if
	end if
end sub

sub Lbonus3_timer
	bonuscount=bonuscount+1
	if bonuscount>LDoubleBonus.state then
		bonuscount=-1
		Lbonus4.timerenabled=true
		Lbonus3.timerenabled=false
		Lbonus3.timerinterval=300	
	 else
		  if AddScore1000Timer.enabled = false then
			  playsound "scorebonus"
			  Lbonus2A.state=0
			  Lbonus3A.state=1
			  if Lbonus3.state=1 then addscore 3000
		  else
			bonuscount=bonuscount-1
			Lbonus3.timerinterval=100	
		  end if
	end if
end sub

sub Lbonus4_timer
	bonuscount=bonuscount+1
	if bonuscount>LDoubleBonus.state then
		bonuscount=-1
		Lbonus5.timerenabled=true
		Lbonus4.timerenabled=false
		Lbonus4.timerinterval=300	
	 else
		  if AddScore1000Timer.enabled = false then
			  playsound "scorebonus"
			  Lbonus3A.state=0
			  Lbonus4A.state=1
			  if Lbonus4.state=1 then addscore 4000
		  else
			bonuscount=bonuscount-1
			Lbonus4.timerinterval=100	
		  end if
	end if
end sub

sub Lbonus5_timer
	bonuscount=bonuscount+1
	if bonuscount>LDoubleBonus.state then
		LDoubleBonus.timerenabled=true
		Lbonus5.timerenabled=false
		Lbonus5.timerinterval=300	
	 else
		  if AddScore1000Timer.enabled = false then
			  playsound "scorebonus"
			  Lbonus4A.state=0
			  Lbonus5A.state=1
			  if Lbonus5.state=1 then addscore 5000
		  else
			bonuscount=bonuscount-1
			Lbonus5.timerinterval=100	
		  end if
	end if
end sub

sub LDoubleBonus_timer
		  if AddScore1000Timer.enabled = false then
			  Lbonus5a.state=0
			  if shootagain.state=lightstateon and tilt=false then
				newball
				ballreltimer.enabled=true
			  else
			   if players=1 or player=players then 
				 player=1
			   else
				 player=player+1
			   end if
				If B2SOn then Controller.B2ssetplayerup player
                Controller.B2SSetData 91,0
'				for i = 1 to maxplayers	
'					EVAL("pup"&i).state=0
'				Next
'				EVAL("pup"&player).state=1
                Controller.B2SSetData 91,1
				nextball
			 end if
			 LDoubleBonus.timerenabled=false
		  end if
end sub


sub newgame
	for i=1 to 2
		EVAL("Bumper"&i).hashitevent = 1
	Next
	player=1
	for i = 1 to maxplayers
		EVAL("p100k"&i).text=""
		EVAL("pup"&i).state=0
	    score(i)=0
		rep(i)=0
	next
    pup1.state=1
	If B2SOn then
	  for i = 1 to maxplayers
		Controller.B2SSetScorePlayer i, score(i)
	  next 
	End If
    eg=0
	shootagain.state=lightstateoff
	shootagain.state=0
    tilttxt.text=" "
	gamov.text=" "
    If B2SOn then
		Controller.B2SSetGameOver 35,0
		Controller.B2SSetTilt 33,0
		Controller.B2SSetMatch 34,0
	End If
	bipreel.setvalue 1
	matchtxt.text=" "
	newball
	Drain.kick 60,45,0
	RandomSoundBallRelease Drain
	DOF 127,DOFPulse
    Strip1=False
    Strip2=False
    Strip3=False
    Strip4=False
    Strip5=False
    Strip6=False
    Strip7=False
    Strip8=False
    Strip9=False
    Strip10=False
end sub


sub newball
	DTstep=1
	resetDT.enabled=true
	bumperlight1.state=1
	for each light in GIlights:light.state=1:next
	For each light in ABClights:light.State = 1: Next
	for each light in BonusLights:light.state=0: next
	LExtraBall.state=0
	Lspecial.state=0
	gilvl=1
End Sub


sub nextball
    if tilt=true then
	  for i=1 to 2
		EVAL("Bumper"&i).hashitevent = 1
	  Next
      tilt=false
      tilttxt.text=" "
		If B2SOn then
			Controller.B2SSetTilt 33,0
'			Controller.B2ssetdata 1, 1
		End If
    end if
	if player=1 then ballinplay=ballinplay+1
	if ballinplay>balls then     
        If Strip10=False Then
           EndVideoTimer.Enabled = True
           pupevent 800+ActiveChar
        End If
        state=false
		playsound "GameOver"
		eg=1
        CharSelectHold=1
		ballreltimer.enabled=true
	  else
			PlaySound("RotateThruPlayers"),0,.05,0,0.25
			TempPlayerUp=Player
			PlayerUpRotator.enabled=true
'		biptext.text=ballinplay
		bipreel.setvalue ballinplay
		If B2SOn then 
			Controller.B2ssetballinplay 32, Ballinplay
            Controller.B2SSetData 91,1
			Controller.B2ssetplayerUp Player
		end if
		if state=true then
		  newball
		  ballreltimer.enabled=true
		end if
	end if
End Sub

sub ballreltimer_timer
  if eg=1 then
	  turnoff
      matchnum
	  state=false
	  bipreel.setvalue 0
	  gamov.text="Game Over"
	  gamov.timerenabled=1
	  tilttxt.timerenabled=1
	  for i = 1 to maxplayers
		EVAL("Canplay"&i).state=0 'Modded
		EVAL("pup"&i).state = 0
	  next
	  checkhighscore
	  If B2SOn then 
        Controller.B2SSetGameOver 35,1
        Controller.B2ssetballinplay 32, 0
	    Controller.B2sStartAnimation "EOGame"
	    Controller.B2SSetScorePlayer 5, hisc
	    Controller.B2ssetcanplay 31, 0
	    Controller.B2ssetplayerup 0
        Controller.B2SSetData 91,0
	  End If
	  ballreltimer.enabled=false
  else
	Drain.kick 60,45,0
	RandomSoundBallRelease Drain
	DOF 127,DOFPulse
    ballreltimer.enabled=false
  end if
end sub

Sub PlayerUpRotator_timer()
		If RotatorTemp<5 then
			TempPlayerUp=TempPlayerUp+1
			If TempPlayerUp>4 then
				TempPlayerUp=1
			end if
			For each objekt in PlayerHuds
				objekt.state = 0
			next
			For each objekt in PlayerHUDScores
				objekt.state=0
			next
			PlayerHuds(TempPlayerUp-1).State = 1
			PlayerHUDScores(TempPlayerUp-1).state=1
			If B2SOn Then
				Controller.B2SSetPlayerUp TempPlayerUp
				Controller.B2SSetData 81,0
				Controller.B2SSetData 82,0
				Controller.B2SSetData 83,0
				Controller.B2SSetData 84,0
				Controller.B2SSetData 80+TempPlayerUp,1
                Controller.B2SSetData 91,1
			End If

		else
			if B2SOn then
				Controller.B2SSetPlayerUp Player
				Controller.B2SSetData 81,0
				Controller.B2SSetData 82,0
				Controller.B2SSetData 83,0
				Controller.B2SSetData 84,0
				Controller.B2SSetData 80+Player,1
                Controller.B2SSetData 91,1
			end if
			PlayerUpRotator.enabled=false
			RotatorTemp=1
			For each objekt in PlayerHuds
				objekt.State = 0
			next
			For each objekt in PlayerHUDScores
				objekt.state=0
			next
			PlayerHuds(Player-1).State = 1
			PlayerHUDScores(Player-1).state=1
		end if
		RotatorTemp=RotatorTemp+1
end sub

sub CheckHighScore
	dim hiscstate
	hiscstate=0
	for i=1 to players
		if score(i)>hisc then 
			hisc=score(i)
			hiscstate=1
		end if
	next
	if hiscstate=1 then HighScoreEntryInit()
	UpdatePostIt 
	savehs
end sub

sub matchnum
	matchnumb=(INT (RND*10))*10
	if matchnumb=0 then
		matchtxt.text="00"
		matchnumb=100
	  else
		matchtxt.text=matchnumb
	end if
	If B2SOn then Controller.B2SSetMatch 34,Matchnumb
	For i=1 to players
		if (matchnumb)=(score(i) mod 100) then 
		  addcredit
		  PlaySound SoundFXDOF("knocker_1",111,DOFPulse,DOFKnocker)
		  DOF 112, DOFPulse
	    end if
    next
end sub

'********** Bumpers

Sub Bumper1_Hit     'left bumper
   if tilt=false then
	RandomSoundBumperTop Bumper1
	DOF 107, DOFPulse
	DOF 108, DOFPulse
	FlashBumpers
	if balls = 3 then
		addscore 1000
	  else
		addscore 100
	end if
   end if
End Sub


Sub Bumper2_Hit    'right bumper
   if tilt=false then
    RandomSoundBumperMiddle Bumper2
	DOF 109, DOFPulse
	DOF 110, DOFPulse
	FlashBumpers
	if balls = 3 then
		addscore 1000
	  else
		addscore 100
	end if
   end if
End Sub



Sub PairedlampTimer_timer
	bumper1light1.state=BumperLight1.state
	bumper2light1.state=bumperlight2.state
	LleftA.state=LtopA.state
	LleftB.state=LtopB.state
	LrightB.state=LtopB.state
	LrightC.state=LtopC.state
	PupA1.state = Pup1.state
	PupA2.state = pup2.state
	PupA3.state = PUP3.state
	PupA4.state = PUP4.state
	LFlip.RotY = LeftFlipper.CurrentAngle
	RFlip.RotY = RightFlipper.CurrentAngle
	RFlip1.RotY = RightFlip1.CurrentAngle-90	
	Pgate.rotz = (Gate.currentangle*.75)+25
end sub

sub FlashBumpers
	if bumperlight1.state = 1 then
		for each light in BumperLights: light.duration 0, 200, 1:Next
	end If
end sub



'************** Slings


Sub LeftSlingShot_Slingshot
	RandomSoundSlingshotLeft slingL
	DOF 103, DOFPulse
	DOF 105, DOFPulse
	addscore 10
    LSling.Visible = 0
    LSling1.Visible = 1
	slingL.objroty = 15	
    LStep = 1
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:slingL.objroty = 7
        Case 4:slingL.objroty = 0:LSLing2.Visible = 0:LSLing.Visible = 1:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub Dingwalls_Hit(idx)
	addscore 10
End Sub	

Sub DingwallA_Hit()
	SlingA.visible=0
	SlingA1.visible=1
	dingwalla.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwalla_timer
	select case dingwalla.uservalue
		Case 1: SlingA1.visible=0: SlingA.visible=1
		case 2:	SlingA.visible=0: SlingA2.visible=1
		Case 3: SlingA2.visible=0: SlingA.visible=1: Me.timerenabled=0
	end Select
	dingwalla.uservalue=dingwalla.uservalue+1
end sub

Sub DingwallB_Hit()
	Slingb.visible=0
	Slingb1.visible=1
	dingwallb.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwallb_timer
	select case DingwallB.uservalue
		Case 1: Slingb1.visible=0: SlingB.visible=1
		case 2:	SlingB.visible=0: Slingb2.visible=1
		Case 3: Slingb2.visible=0: SlingB.visible=1: Me.timerenabled=0
	end Select
	dingwallb.uservalue=DingwallB.uservalue+1
end sub

Sub DingwallC_Hit()
	Slingc.visible=0
	Slingc1.visible=1
	dingwallc.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwallc_timer
	select case Dingwallc.uservalue
		Case 1: Slingc1.visible=0: Slingc.visible=1
		case 2:	Slingc.visible=0: Slingc2.visible=1
		Case 3: Slingc2.visible=0: Slingc.visible=1: Me.timerenabled=0
	end Select
	dingwallc.uservalue=Dingwallc.uservalue+1
end sub

Sub DingwallD_Hit()
	Slingd.visible=0
	Slingd1.visible=1
	dingwalld.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwalld_timer
	select case Dingwalld.uservalue
		Case 1: Slingd1.visible=0: Slingd.visible=1
		case 2:	Slingd.visible=0: Slingd2.visible=1
		Case 3: Slingd2.visible=0: Slingd.visible=1: Me.timerenabled=0
	end Select
	dingwalld.uservalue=Dingwalld.uservalue+1
end sub

Sub DingwallE_Hit()
	Slinge.visible=0
	Slinge1.visible=1
	dingwalle.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwalle_timer
	select case Dingwalle.uservalue
		Case 1: Slinge1.visible=0: Slinge.visible=1
		case 2:	Slinge.visible=0: Slinge2.visible=1
		Case 3: Slinge2.visible=0: Slinge.visible=1: Me.timerenabled=0
	end Select
	dingwalle.uservalue=Dingwalle.uservalue+1
end sub

Sub DingwallF_Hit()
	Slingf.visible=0
	Slingf1.visible=1
	dingwallf.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwallf_timer
	select case Dingwallf.uservalue
		Case 1: Slingf1.visible=0: Slingf.visible=1
		case 2:	Slingf.visible=0: Slingf2.visible=1
		Case 3: Slingf2.visible=0: Slingf.visible=1: Me.timerenabled=0
	end Select
	dingwallf.uservalue=Dingwallf.uservalue+1
end sub

Sub DingwallG_Hit()
	Slingg.visible=0
	Slingg1.visible=1
	dingwallg.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwallg_timer
	select case Dingwallg.uservalue
		Case 1: Slingg1.visible=0: Slingg.visible=1
		case 2:	Slingg.visible=0: Slingg2.visible=1
		Case 3: Slingg2.visible=0: Slingg.visible=1: Me.timerenabled=0
	end Select
	dingwallg.uservalue=Dingwallg.uservalue+1
end sub

Sub DingwallH_Hit()
	SlingH.visible=0
	Slingh1.visible=1
	dingwallh.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwallh_timer
	select case Dingwallh.uservalue
		Case 1: Slingh1.visible=0: Slingh.visible=1
		case 2:	Slingh.visible=0: Slingh2.visible=1
		Case 3: Slingh2.visible=0: SlingH.visible=1: Me.timerenabled=0
	end Select
	DingwallH.uservalue=Dingwallh.uservalue+1
end sub

Sub DingwallI_Hit()
	SlingI.visible=0
	Slingi1.visible=1
	DingwallI.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwalli_timer
	select case DingwallI.uservalue
		Case 1: Slingi1.visible=0: SlingI.visible=1
		case 2:	SlingI.visible=0: Slingi2.visible=1
		Case 3: Slingi2.visible=0: SlingI.visible=1: Me.timerenabled=0
	end Select
	DingwallI.uservalue=DingwallI.uservalue+1
end sub

Sub DingwallJ_Hit()
	SlingJ.visible=0
	Slingj1.visible=1
	DingwallJ.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwallj_timer
	select case DingwallJ.uservalue
		Case 1: Slingj1.visible=0: Slingj.visible=1
		case 2:	SlingJ.visible=0: Slingj2.visible=1
		Case 3: Slingj2.visible=0: SlingJ.visible=1: Me.timerenabled=0
	end Select
	DingwallJ.uservalue=DingwallJ.uservalue+1
end sub

'********** Triggers     

sub TGspecial_hit
	DOF 130, DOFPulse
	addscore 500
	if Lspecial.state = 1 then
		addcredit
		PlaySound SoundFXDOF("knocker_1",111,DOFPulse,DOFKnocker)
		DOF 112, DOFPulse
	end if
end sub 

sub TGleftA_hit
	DOF 123, DOFPulse
	if LtopA.state=1 then 
		addscore 5000
	  else
		addscore 500
	end if
	LtopA.state=0
	CheckAwards
end sub    

sub TGleftB_hit
	DOF 124, DOFPulse
	if LtopB.state=1 then 
		addscore 5000
	  else
		addscore 500
	end if
	LtopB.state=0
	CheckAwards
end sub

sub TGrightB_hit
	DOF 125, DOFPulse
	if LtopB.state=1 then 
		addscore 5000
	  else
		addscore 500
	end if
	LtopB.state=0
	CheckAwards
end sub

sub TGrightC_hit
	DOF 126, DOFPulse
	if LtopC.state=1 then 
		addscore 5000
	  else
		addscore 500
	end if
	LtopC.state=0
	CheckAwards
end sub

sub TGtopA_hit
	DOF 120, DOFPulse
	if LtopA.state=1 then 
		addscore 5000
	  else
		addscore 500
	end if
	LtopA.state=0
	CheckAwards
end sub    

sub TGtopB_hit  
	DOF 121, DOFPulse
	if LtopB.state=1 then 
		addscore 5000
	  else
		addscore 500
	end if
	LtopB.state=0
	CheckAwards
end sub 

sub TGtopC_hit  
	DOF 122, DOFPulse
	if LtopC.state=1 then 
		addscore 5000
	  else
		addscore 500
	end if
	LtopC.state=0
	CheckAwards
end sub 


'********** Drop targets


Sub DTA1_Hit : DTHit 1 : TargetBouncer Activeball, 1 : End Sub
Sub DTA2_Hit : DTHit 2 : TargetBouncer Activeball, 1 : End Sub
Sub DTJK_Hit : DTHit 3 : TargetBouncer Activeball, 1 : End Sub
Sub DTA3_Hit : DTHit 4 : TargetBouncer Activeball, 1 : End Sub
Sub DTA4_Hit : DTHit 5 : TargetBouncer Activeball, 1 : End Sub
Sub DT10_Hit : DTHit 6 : TargetBouncer Activeball, 1 : End Sub
Sub DTJ1_Hit : DTHit 7 : TargetBouncer Activeball, 1 : End Sub
Sub DTJ2_Hit : DTHit 8 : TargetBouncer Activeball, 1 : End Sub
Sub DTQ1_Hit : DTHit 9 : TargetBouncer Activeball, 1 : End Sub
Sub DTQ2_Hit : DTHit 10 : TargetBouncer Activeball, 1 : End Sub
Sub DTQ3_Hit : DTHit 11 : TargetBouncer Activeball, 1 : End Sub
Sub DTK1_Hit : DTHit 12 : TargetBouncer Activeball, 1 : End Sub
Sub DTK2_Hit : DTHit 13 : TargetBouncer Activeball, 1 : End Sub
Sub DTK4_Hit : DTHit 15 : End Sub
Sub DTK3_Hit : DTHit 14 : TargetBouncer Activeball, 1 : End Sub



Sub resetDT_timer

	Select Case DTStep
		Case 1:							'TEN AND JACKS
			DOF 134,2
			PlaySound SoundFX("BankReset",DOFContactors), 0, .7, -.5, 0.05
			DTRaise 6
			DTRaise 7
			DTRaise 8
		Case 2:							'QUEENS
			DOF 109,2
			PlaySound SoundFX("BankReset",DOFContactors), 0, .7, 0, 0.05
			DTRaise 9
			DTRaise 10
			DTRaise 11
			for i = 0 to 1
				DTlights(i).state=0
			Next
		Case 3:							'KINGS
			DOF 135,2
			PlaySound SoundFX("BankReset",DOFContactors), 0, .7, .5, 0.05
			DTRaise 12
			DTRaise 13
			DTRaise 14
			DTRaise 15
			for i = 2 to 7
				DTlights(i).state=0
			Next
		Case 4:							'ACES AND JOKER
			DOF 107,2
			PlaySound SoundFX("BankReset", DOFContactors), 0, .7, -.5, 0.05
			DTRaise 1
			DTRaise 2
			DTRaise 3
			DTRaise 4
			DTRaise 5
			for i = 8 to 13
				DTlights(i).state=0
			Next
		Case 5:
			resetDT.enabled=False
	end Select
	DTStep=DTStep+1
end sub



'************ Target


Sub Textraball_Hit()
	DOF 115, DOFPulse
	addscore 500
	if LExtraBall.state=1 then
		ShootAgain.state=1
		If B2SOn then Controller.B2SSetShootAgain 36,1
	end if
End Sub

Sub CheckAwards
	if DTIsDropped(7)+DTIsDropped(8)=2 then Lbonus2.state=1		'JACKS
	if DTIsDropped(9)+DTIsDropped(10)+DTIsDropped(11)=3 then Lbonus3.state=1	'QUEENS
	if DTIsDropped(12)+DTIsDropped(13)+DTIsDropped(14)+DTIsDropped(15)=4 then Lbonus4.state=1	'KINGS
	if DTIsDropped(1)+DTIsDropped(2)+DTIsDropped(3)+DTIsDropped(4)+DTIsDropped(5)=5 then 	'ACES AND JOKER
		LExtraBall.state=1
		if balls=3 then Lspecial.state=1
		Lbonus5.state=1
	end if
	If LtopA.state + LtopB.state + LtopC.state= 0 Then
		LExtraBall.state= 1
		Ldoublebonus.state= 1
		if balls=3 then Lspecial.state=1
	End If
End sub


sub addscore(points)
  if tilt=false and state=true then
    If Points < 100 and AddScore10Timer.enabled = false Then
        Add10 = Points \ 10
        AddScore10Timer.Enabled = TRUE
      ElseIf Points < 1000 and AddScore100Timer.enabled = false Then
        Add100 = Points \ 100
        AddScore100Timer.Enabled = TRUE
      ElseIf AddScore1000Timer.enabled = false Then
        Add1000 = Points \ 1000
        AddScore1000Timer.Enabled = TRUE
    End If
  end if
  PupScoreMonitor
End Sub

Sub AddScore10Timer_Timer()
    if Add10 > 0 then
        AddPoints 10
        Add10 = Add10 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore100Timer_Timer()
    if Add100 > 0 then
        AddPoints 100
        Add100 = Add100 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore1000Timer_Timer()
    if Add1000 > 0 then
        AddPoints 1000
        Add1000 = Add1000 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddPoints(Points)
    score(player)=score(player)+points
	EVAL("ScoreReel"&player).addvalue(points)
	If B2SOn Then Controller.B2SSetScorePlayer player, score(player) MOD 100000

    ' Sounds: there are 3 sounds: tens, hundreds and thousands
    If Points = 100 AND(Score(player) MOD 1000) \ 100 = 0 Then  'New 1000 reel
        PlaySound SoundFXDOF("bell1000",133,DOFPulse,DOFChimes)
      ElseIf Points = 10 AND(Score(player) MOD 100) \ 10 = 0 Then 'New 100 reel
        PlaySound SoundFXDOF("bell100",132,DOFPulse,DOFChimes)
      ElseIf points = 1000 Then
        PlaySound SoundFXDOF("bell1000",133,DOFPulse,DOFChimes)
	  elseif Points = 100 Then
        PlaySound SoundFXDOF("bell100",132,DOFPulse,DOFChimes)
      Else
        PlaySound SoundFXDOF("bell10",131,DOFPulse,DOFChimes)
    End If
    ' check replays and rollover
	if score(player)=>99999 then
		If B2SOn then Controller.B2SSetScoreRollover 24 + player, 1
		EVAL("p100k"&player).text="100,000"
	End if
    if score(player)=>replay1 and rep(player)=0 then
		addcredit
		rep(player)=1
		PlaySound SoundFXDOF("knocker_1",111,DOFPulse,DOFKnocker)
		DOF 112, DOFPulse
    end if
    if score(player)=>replay2 and rep(player)=1 then
		addcredit
		rep(player)=2
		PlaySound SoundFXDOF("knocker_1",111,DOFPulse,DOFKnocker)
		DOF 112, DOFPulse
    end if
    if score(player)=>replay3 and rep(player)=2 then
		addcredit
		rep(player)=3
		PlaySound SoundFXDOF("knocker_1",111,DOFPulse,DOFKnocker)
		DOF 112, DOFPulse
    end if
    PupScoreMonitor
end sub 

Sub CheckTilt
	If Tilttimer.Enabled = True Then 
	 TiltSens = TiltSens + 1
	 if TiltSens = 3 Then GameTilted
	Else
	 TiltSens = 0
	 Tilttimer.Enabled = True
	End If
End Sub

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

Sub GameTilted
	Tilt = True
	tilttxt.text="TILT"
    If B2SOn Then Controller.B2SSetTilt 33,1
    If B2SOn Then Controller.B2ssetdata 1, 0
	playsound "tilt"
	turnoff
End Sub

sub turnoff
    bumper1.hashitevent=0
    bumper2.hashitevent=0
	LeftFlipper.RotateToStart
	DOF 101,DOFOff
	StopSound "Buzz"
	RightFlipper.RotateToStart
	Rightflip1.RotateToStart
	Lupflip.state=0
	DOF 102,DOFOff
	StopSound "Buzz1"
end sub    




'******************************************************
'****  BALL ROLLING AND DROP SOUNDS
'******************************************************
Const tnob = 2 ' total number of balls
Const lob = 0	'locked balls on start; might need some fiddling depending on how your locked balls are done

ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

Sub InitRolling
	Dim i
	For i = 0 to tnob
		rolling(i) = False
	Next
End Sub

Sub RollingUpdate()
	Dim BOT, b
	BOT = GetBalls

	' stop the sound of deleted balls
	For b = UBound(BOT) + 1 to tnob
		If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
		rolling(b) = False
		StopSound("BallRoll_" & b)
	Next

	' exit the sub if no balls on the table
	If UBound(BOT) = -1 Then Exit Sub

	' play the rolling sound for each ball

	For b = 0 to UBound(BOT)
		If BallVel(BOT(b)) > 1 AND BOT(b).z < 30 Then
			rolling(b) = True
			PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(BOT(b)) * BallRollingVolume * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))

		Else
			If rolling(b) = True Then
				StopSound("BallRoll_" & b)
				rolling(b) = False
			End If
		End If

		' Ball Drop Sounds
		If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
			If DropCount(b) >= 5 Then
				DropCount(b) = 0
				If BOT(b).velz > -7 Then
					RandomSoundBallBouncePlayfieldSoft BOT(b)
				Else
					RandomSoundBallBouncePlayfieldHard BOT(b)
				End If				
			End If
		End If
		If DropCount(b) < 5 Then
			DropCount(b) = DropCount(b) + 1
		End If

		' "Static" Ball Shadows
		If AmbientBallShadowOn = 0 Then
			If BOT(b).Z > 30 Then
				BallShadowA(b).height=BOT(b).z - BallSize/4		'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
			Else
				BallShadowA(b).height=BOT(b).z - BallSize/2 + 5
			End If
			BallShadowA(b).Y = BOT(b).Y + Ballsize/5 + fovY
			BallShadowA(b).X = BOT(b).X
			BallShadowA(b).visible = 1
		End If
	Next
End Sub


'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************

'
sub savehs
	Dim FileObj
	Dim ScoreFile
	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then 
		Exit Sub
	End if
	Set ScoreFile=FileObj.CreateTextFile(UserDirectory & HSFileName,True)
    ScoreFile.WriteLine credit
    ScoreFile.WriteLine hisc
    ScoreFile.WriteLine matchnumb
    ScoreFile.WriteLine score(1)
    ScoreFile.WriteLine score(2)
    ScoreFile.WriteLine score(3)
    ScoreFile.WriteLine score(4)
	ScoreFile.WriteLine balls
	ScoreFile.WriteLine HSA1
	ScoreFile.WriteLine HSA2
	ScoreFile.WriteLine HSA3
	ScoreFile.WriteLine freeplay
	ScoreFile.WriteLine altapron
    ScoreFile.Close
	Set ScoreFile=Nothing
	Set FileObj=Nothing
end sub

sub loadhs
    Dim FileObj
	Dim ScoreFile
    dim temp1
    dim temp2
    dim temp3
    dim temp4
    dim temp5
    dim temp6
    dim temp7
    dim temp8
    dim temp9
    dim temp10
    dim temp11
    dim temp12
    dim temp13
    Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then 
		Exit Sub
	End if
	If Not FileObj.FileExists(UserDirectory & HSFileName) then
		Exit Sub
	End if
	Set ScoreFile=FileObj.GetFile(UserDirectory & HSFileName)
	Set TextStr=ScoreFile.OpenAsTextStream(1,0)
		If (TextStr.AtEndOfStream=True) then
			Exit Sub
        End if
		temp1=textStr.readLine
		temp2=textstr.readline
		temp3=textstr.readline
		temp4=textstr.readline
		temp5=textstr.readline
		temp6=textstr.readline
		temp7=textstr.readline
		temp8=textStr.readLine
		temp9=textstr.readline
		temp10=textstr.readline
		temp11=textstr.readline
		temp12=textstr.readline
		temp13=textstr.readline
	TextStr.Close
     credit = CDbl(temp1)
     hisc = CDbl(temp2) 
     matchnumb = CDbl(temp3)
     score(1) = CDbl(temp4)
     score(2) = CDbl(temp5)
     score(3) = CDbl(temp6)
     score(4) = CDbl(temp7)
     balls = CDbl(temp8)
     HSA1 = CDbl(temp9)
     HSA2 = CDbl(temp10)
     HSA3 = CDbl(temp11)
     freeplay = CDbl(temp12)
     altapron = CDbl(temp13)
     Set ScoreFile=Nothing
	 Set FileObj=Nothing
end sub



Sub Table1_Exit()
	Savehs
    pupflasher.VideoCapUpdate=""
	turnoff
	If B2SOn Then Controller.stop
End Sub

'==========================================================================================================================================
'============================================================= START OF HIGH SCORES ROUTINES =============================================================
'==========================================================================================================================================
'
'ADD LINE TO TABLE_KEYDOWN SUB WITH THE FOLLOWING:    If HSEnterMode Then HighScoreProcessKey(keycode) AFTER THE STARTGAME ENTRY
'ADD And Not HSEnterMode=true TO IF KEYCODE=STARTGAMEKEY
'TO SHOW THE SCORE ON POST-IT ADD LINE AT RELEVENT LOCATION THAT HAS:  UpdatePostIt
'TO INITIATE ADDING INITIALS ADD LINE AT RELEVENT LOCATION THAT HAS:  HighScoreEntryInit()
'ADD THE FOLLOWING LINES TO TABLE_INIT TO SETUP POSTIT
'	if HSA1="" then HSA1=25
'	if HSA2="" then HSA2=25
'	if HSA3="" then HSA3=25
'	UpdatePostIt
'ADD HSA1, HSA2 AND HSA3 TO SAVE AND LOAD VALUES FOR TABLE
'ADD A TIMER NAMED HighScoreFlashTimer WITH INTERVAL 100 TO TABLE
'SET HSSSCOREX BELOW TO WHATEVER VARIABLE YOU USE FOR HIGH SCORE.
'IMPORT POST-IT IMAGES


Dim HSA1, HSA2, HSA3
Dim HSEnterMode, hsLetterFlash, hsEnteredDigits(3), hsCurrentDigit, hsCurrentLetter
Dim HSArray  
Dim HSScoreM,HSScore100k, HSScore10k, HSScoreK, HSScore100, HSScore10, HSScore1, HSScorex	'Define 6 different score values for each reel to use
HSArray = Array("Postit0","postit1","postit2","postit3","postit4","postit5","postit6","postit7","postit8","postit9","postitBL","postitCM","Tape")
Const hsFlashDelay = 4

' ***********************************************************
'  HiScore DISPLAY 
' ***********************************************************

Sub UpdatePostIt
	dim tempscore
	HSScorex = hisc
	TempScore = HSScorex
	HSScore1 = 0
	HSScore10 = 0
	HSScore100 = 0
	HSScoreK = 0
	HSScore10k = 0
	HSScore100k = 0
	HSScoreM = 0
	if len(TempScore) > 0 Then
		HSScore1 = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore10 = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore100 = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScoreK = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore10k = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore100k = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScoreM = cint(right(Tempscore,1))
	end If
	Pscore6.image = HSArray(HSScoreM):If HSScorex<1000000 Then PScore6.image = HSArray(10)
	Pscore5.image = HSArray(HSScore100K):If HSScorex<100000 Then PScore5.image = HSArray(10)
	PScore4.image = HSArray(HSScore10K):If HSScorex<10000 Then PScore4.image = HSArray(10)
	PScore3.image = HSArray(HSScoreK):If HSScorex<1000 Then PScore3.image = HSArray(10)
	PScore2.image = HSArray(HSScore100):If HSScorex<100 Then PScore2.image = HSArray(10)
	PScore1.image = HSArray(HSScore10):If HSScorex<10 Then PScore1.image = HSArray(10)
	PScore0.image = HSArray(HSScore1):If HSScorex<1 Then PScore0.image = HSArray(10)
	if HSScorex<1000 then
		PComma.image = HSArray(10)
	else
		PComma.image = HSArray(11)
	end if
	if HSScorex<1000000 then
		PComma2.image = HSArray(10)
	else
		PComma2.image = HSArray(11)
	end if
'	if showhisc=1 and showhiscnames=1 then
'		for each objekt in hiscname:objekt.visible=1:next
		HSName1.image = ImgFromCode(HSA1, 1)
		HSName2.image = ImgFromCode(HSA2, 2)
		HSName3.image = ImgFromCode(HSA3, 3)
'	  else
'		for each objekt in hiscname:objekt.visible=0:next
'	end if
   CoinLockout=0
End Sub

Function ImgFromCode(code, digit)
	Dim Image
	if (HighScoreFlashTimer.Enabled = True and hsLetterFlash = 1 and digit = hsCurrentLetter) then
		Image = "postitBL"
	elseif (code + ASC("A") - 1) >= ASC("A") and (code + ASC("A") - 1) <= ASC("Z") then
		Image = "postit" & chr(code + ASC("A") - 1)
	elseif code = 27 Then
		Image = "PostitLT"
    elseif code = 0 Then
		image = "PostitSP"
    Else
      msgbox("Unknown display code: " & code)
	end if
	ImgFromCode = Image
End Function

Sub HighScoreEntryInit()
	HSA1=0:HSA2=0:HSA3=0
	HSEnterMode = True
	hsCurrentDigit = 0
	hsCurrentLetter = 1:HSA1=1
	HighScoreFlashTimer.Interval = 250
	HighScoreFlashTimer.Enabled = True
	hsLetterFlash = hsFlashDelay
End Sub

Sub HighScoreFlashTimer_Timer()
	hsLetterFlash = hsLetterFlash-1
	UpdatePostIt
	If hsLetterFlash=0 then 'switch back
		hsLetterFlash = hsFlashDelay
	end if
End Sub


' ***********************************************************
'  HiScore ENTER INITIALS 
' ***********************************************************

Sub HighScoreProcessKey(keycode)
    If keycode = LeftFlipperKey Then
        PlaySound "cluper"
		hsLetterFlash = hsFlashDelay
		Select Case hsCurrentLetter
			Case 1:
				HSA1=HSA1-1:If HSA1=-1 Then HSA1=26 'no backspace on 1st digit
				UpdatePostIt
			Case 2:
				HSA2=HSA2-1:If HSA2=-1 Then HSA2=27
				UpdatePostIt
			Case 3:
				HSA3=HSA3-1:If HSA3=-1 Then HSA3=27
				UpdatePostIt
		 End Select
    End If

	If keycode = RightFlipperKey Then
		hsLetterFlash = hsFlashDelay
        PlaySound "cluper"
		Select Case hsCurrentLetter
			Case 1:
				HSA1=HSA1+1:If HSA1>26 Then HSA1=0
				UpdatePostIt
			Case 2:
				HSA2=HSA2+1:If HSA2>27 Then HSA2=0
				UpdatePostIt
			Case 3:
				HSA3=HSA3+1:If HSA3>27 Then HSA3=0
				UpdatePostIt
		 End Select
	End If
	
    If keycode = StartGameKey Then
        PlaySound "drop1"
        CoinLockOut=1
		Select Case hsCurrentLetter
			Case 1:
				hsCurrentLetter=2 'ok to advance
				HSA2=HSA1 'start at same alphabet spot
'				EMReelHSName1.SetValue HSA1:EMReelHSName2.SetValue HSA2
			Case 2:
				If HSA2=27 Then 'bksp
					HSA2=0
					hsCurrentLetter=1
				Else
					hsCurrentLetter=3 'enter it
					HSA3=HSA2 'start at same alphabet spot
				End If
			Case 3:
				If HSA3=27 Then 'bksp
					HSA3=0
					hsCurrentLetter=2
				Else
					savehs 'enter it
					HighScoreFlashTimer.Enabled = False
					HSEnterMode = False
                    pupevent 700
                    CharSelectHold=1
				End If
		End Select
		UpdatePostIt
    End If
End Sub


'*************
'   JP'S LUT
'*************

Dim bLutActive, LUTImage
Dim x
Sub LoadLUT
	bLutActive = False
    x = LoadValue(cGameName, "LUTImage")
    If(x <> "") Then LUTImage = x Else LUTImage = 0
	UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT: LUTImage = (LUTImage +1 ) MOD 13: UpdateLUT: SaveLUT: End Sub

Sub UpdateLUT
Select Case LutImage
Case 0: Table1.ColorGradeImage = "LUT0":ScreenLight.intensity=0
Case 1: Table1.ColorGradeImage = "LUT1":ScreenLight.intensity=0.3
Case 2: Table1.ColorGradeImage = "LUT2":ScreenLight.intensity=0.4
Case 3: Table1.ColorGradeImage = "LUT3":ScreenLight.intensity=0.5
Case 4: Table1.ColorGradeImage = "LUT4":ScreenLight.intensity=0.6
Case 5: Table1.ColorGradeImage = "LUT5":ScreenLight.intensity=0.8
Case 6: Table1.ColorGradeImage = "LUT6":ScreenLight.intensity=1
Case 7: Table1.ColorGradeImage = "LUT7":ScreenLight.intensity=1.1
Case 8: Table1.ColorGradeImage = "LUT8":ScreenLight.intensity=1.4
Case 9: Table1.ColorGradeImage = "LUT9":ScreenLight.intensity=1.7
Case 10: Table1.ColorGradeImage = "LUT10":ScreenLight.intensity=2
Case 11: Table1.ColorGradeImage = "LUT11":ScreenLight.intensity=2.3
Case 12: Table1.ColorGradeImage = "LUT12":ScreenLight.intensity=2.7
End Select
End Sub



'*********** FLIPPER SHADOWS ************

EnableFlipperShadows
Sub EnableFlipperShadows
       if AddFlipperShadows = 1 Then
        FlipperLSh.visible=true
		FLipperRSh.visible=true
       Else
        FlipperLSh.visible=false
		FLipperRSh.visible=false
End If
End Sub



'***************************************************************
'****  VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'***************************************************************

Const fovY					= 0		'Offset y position under ball to account for layback or inclination (more pronounced need further back)
Const DynamicBSFactor 		= 0.95	'0 to 1, higher is darker
Const AmbientBSFactor 		= 0.7	'0 to 1, higher is darker
Const AmbientMovement		= 2		'1 to 4, higher means more movement as the ball moves left and right
Const Wideness				= 20	'Sets how wide the dynamic ball shadows can get (20 +5 thinness should be most realistic for a 50 unit ball)
Const Thinness				= 5		'Sets minimum as ball moves away from source

Dim sourcenames, currentShadowCount, DSSources(30), numberofsources, numberofsources_hold
sourcenames = Array ("","","")
currentShadowCount = Array (0,0,0)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!
dim objrtx1(2), objrtx2(2)
dim objBallShadow(2)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2)

DynamicBSInit

sub DynamicBSInit()
	Dim iii, source

	for iii = 0 to tnob									'Prepares the shadow objects before play begins
		Set objrtx1(iii) = Eval("RtxBallShadow" & iii)
		objrtx1(iii).material = "RtxBallShadow" & iii
		objrtx1(iii).z = iii/1000 + 0.01
		objrtx1(iii).visible = 0

		Set objrtx2(iii) = Eval("RtxBall2Shadow" & iii)
		objrtx2(iii).material = "RtxBallShadow2_" & iii
		objrtx2(iii).z = (iii)/1000 + 0.02
		objrtx2(iii).visible = 0

		currentShadowCount(iii) = 0

		Set objBallShadow(iii) = Eval("BallShadow" & iii)
		objBallShadow(iii).material = "BallShadow" & iii
		UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
		objBallShadow(iii).Z = iii/1000 + 0.04
		objBallShadow(iii).visible = 0

		BallShadowA(iii).Opacity = 100*AmbientBSFactor
		BallShadowA(iii).visible = 0
	Next

	iii = 0

	For Each Source in DynamicSources
		DSSources(iii) = Array(Source.x, Source.y)
		iii = iii + 1
	Next
	numberofsources = iii
	numberofsources_hold = iii
end sub


Sub DynamicBSUpdate
	Dim falloff:	falloff = 150			'Max distance to light sources, can be changed if you have a reason
	Dim ShadowOpacity, ShadowOpacity2 
	Dim s, Source, LSd, currentMat, AnotherSource, BOT, iii
	BOT = GetBalls

	'Hide shadow of deleted balls
	For s = UBound(BOT) + 1 to tnob
		objrtx1(s).visible = 0
		objrtx2(s).visible = 0
		objBallShadow(s).visible = 0
		BallShadowA(s).visible = 0
	Next

	If UBound(BOT) < lob Then Exit Sub		'No balls in play, exit

	'The Magic happens now
	For s = lob to UBound(BOT)

	' *** Normal "ambient light" ball shadow
	'Layered from top to bottom. If you had an upper pf at for example 80 and ramps even above that, your segments would be z>110; z<=110 And z>100; z<=100 And z>30; z<=30 And z>20; Else invisible

		If AmbientBallShadowOn = 1 Then			'Primitive shadow on playfield, flasher shadow in ramps
			If BOT(s).Z > 30 Then							'The flasher follows the ball up ramps while the primitive is on the pf
				If BOT(s).X < tablewidth/2 Then
					objBallShadow(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
				Else
					objBallShadow(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
				End If
				objBallShadow(s).Y = BOT(s).Y + BallSize/10 + fovY
				objBallShadow(s).visible = 1

				BallShadowA(s).X = BOT(s).X
				BallShadowA(s).Y = BOT(s).Y + BallSize/5 + fovY
				BallShadowA(s).height=BOT(s).z - BallSize/4		'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
				BallShadowA(s).visible = 1
			Elseif BOT(s).Z <= 30 And BOT(s).Z > 20 Then	'On pf, primitive only
				objBallShadow(s).visible = 1
				If BOT(s).X < tablewidth/2 Then
					objBallShadow(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
				Else
					objBallShadow(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
				End If
				objBallShadow(s).Y = BOT(s).Y + fovY
				BallShadowA(s).visible = 0
			Else											'Under pf, no shadows
				objBallShadow(s).visible = 0
				BallShadowA(s).visible = 0
			end if

		Elseif AmbientBallShadowOn = 2 Then		'Flasher shadow everywhere
			If BOT(s).Z > 30 Then							'In a ramp
				BallShadowA(s).X = BOT(s).X
				BallShadowA(s).Y = BOT(s).Y + BallSize/5 + fovY
				BallShadowA(s).height=BOT(s).z - BallSize/4		'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
				BallShadowA(s).visible = 1
			Elseif BOT(s).Z <= 30 And BOT(s).Z > 20 Then	'On pf
				BallShadowA(s).visible = 1
				If BOT(s).X < tablewidth/2 Then
					BallShadowA(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
				Else
					BallShadowA(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
				End If
				BallShadowA(s).Y = BOT(s).Y + Ballsize/10 + fovY
				BallShadowA(s).height=BOT(s).z - BallSize/2 + 5
			Else											'Under pf
				BallShadowA(s).visible = 0
			End If
		End If

		' *** Dynamic shadows
		If DynamicBallShadowsOn Then
			If BOT(s).Z < 30 Then 'And BOT(s).Y < (TableHeight - 200) Then 'Or BOT(s).Z > 105 Then		'Defining when and where (on the table) you can have dynamic shadows
				For iii = 0 to numberofsources - 1 
					LSd=DistanceFast((BOT(s).x-DSSources(iii)(0)),(BOT(s).y-DSSources(iii)(1)))	'Calculating the Linear distance to the Source
					If LSd < falloff And gilvl > 0 Then						    'If the ball is within the falloff range of a light and light is on (we will set numberofsources to 0 when GI is off)
						currentShadowCount(s) = currentShadowCount(s) + 1		'Within range of 1 or 2
						if currentShadowCount(s) = 1 Then						'1 dynamic shadow source
							sourcenames(s) = iii
							currentMat = objrtx1(s).material
							objrtx2(s).visible = 0 : objrtx1(s).visible = 1 : objrtx1(s).X = BOT(s).X : objrtx1(s).Y = BOT(s).Y + fovY
	'						objrtx1(s).Z = BOT(s).Z - 25 + s/1000 + 0.01						'Uncomment if you want to add shadows to an upper/lower pf
							objrtx1(s).rotz = AnglePP(DSSources(iii)(0), DSSources(iii)(1), BOT(s).X, BOT(s).Y) + 90
							ShadowOpacity = (falloff-LSd)/falloff									'Sets opacity/darkness of shadow by distance to light
							objrtx1(s).size_y = Wideness*ShadowOpacity+Thinness						'Scales shape of shadow with distance/opacity
							UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity*DynamicBSFactor^2,RGB(0,0,0),0,0,False,True,0,0,0,0
							If AmbientBallShadowOn = 1 Then
								currentMat = objBallShadow(s).material									'Brightens the ambient primitive when it's close to a light
								UpdateMaterial currentMat,1,0,0,0,0,0,AmbientBSFactor*(1-ShadowOpacity),RGB(0,0,0),0,0,False,True,0,0,0,0
							Else
								BallShadowA(s).Opacity = 100*AmbientBSFactor*(1-ShadowOpacity)
							End If

						Elseif currentShadowCount(s) = 2 Then
																	'Same logic as 1 shadow, but twice
							currentMat = objrtx1(s).material
							AnotherSource = sourcenames(s)
							objrtx1(s).visible = 1 : objrtx1(s).X = BOT(s).X : objrtx1(s).Y = BOT(s).Y + fovY
	'						objrtx1(s).Z = BOT(s).Z - 25 + s/1000 + 0.01							'Uncomment if you want to add shadows to an upper/lower pf
							objrtx1(s).rotz = AnglePP(DSSources(AnotherSource)(0),DSSources(AnotherSource)(1), BOT(s).X, BOT(s).Y) + 90
							ShadowOpacity = (falloff-DistanceFast((BOT(s).x-DSSources(AnotherSource)(0)),(BOT(s).y-DSSources(AnotherSource)(1))))/falloff
							objrtx1(s).size_y = Wideness*ShadowOpacity+Thinness
							UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0

							currentMat = objrtx2(s).material
							objrtx2(s).visible = 1 : objrtx2(s).X = BOT(s).X : objrtx2(s).Y = BOT(s).Y + fovY
	'						objrtx2(s).Z = BOT(s).Z - 25 + s/1000 + 0.02							'Uncomment if you want to add shadows to an upper/lower pf
							objrtx2(s).rotz = AnglePP(DSSources(iii)(0), DSSources(iii)(1), BOT(s).X, BOT(s).Y) + 90
							ShadowOpacity2 = (falloff-LSd)/falloff
							objrtx2(s).size_y = Wideness*ShadowOpacity2+Thinness
							UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity2*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0
							If AmbientBallShadowOn = 1 Then
								currentMat = objBallShadow(s).material									'Brightens the ambient primitive when it's close to a light
								UpdateMaterial currentMat,1,0,0,0,0,0,AmbientBSFactor*(1-max(ShadowOpacity,ShadowOpacity2)),RGB(0,0,0),0,0,False,True,0,0,0,0
							Else
								BallShadowA(s).Opacity = 100*AmbientBSFactor*(1-max(ShadowOpacity,ShadowOpacity2))
							End If
						end if
					Else
						currentShadowCount(s) = 0
						BallShadowA(s).Opacity = 100*AmbientBSFactor
					End If
				Next
			Else									'Hide dynamic shadows everywhere else
				objrtx2(s).visible = 0 : objrtx1(s).visible = 0
			End If
		End If
	Next
End Sub
'****************************************************************
'****  END VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'****************************************************************


'******************************************************
' FLIPPERS
'******************************************************

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled)
	If Enabled Then
		LF.Fire  'leftflipper.rotatetoend

		If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then 
			RandomSoundReflipUpLeft LeftFlipper
		Else 
			SoundFlipperUpAttackLeft LeftFlipper
			RandomSoundFlipperUpLeft LeftFlipper
		End If		
	Else
		LeftFlipper.RotateToStart
		If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
			RandomSoundFlipperDownLeft LeftFlipper
		End If
		FlipperLeftHitParm = FlipperUpSoundLevel
	End If
End Sub

Sub SolRFlipper(Enabled)
	If Enabled Then
		RF.Fire 'rightflipper.rotatetoend
		Rightflip1.RotateToEnd
		If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
			RandomSoundReflipUpRight RightFlipper
		Else 
			SoundFlipperUpAttackRight RightFlipper
			RandomSoundFlipperUpRight RightFlipper
		End If
	Else
		RightFlipper.RotateToStart
		Rightflip1.RotateToStart
		If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
			RandomSoundFlipperDownRight RightFlipper
		End If	
		FlipperRightHitParm = FlipperUpSoundLevel
	End If
End Sub

' Flipper collide subs
Sub LeftFlipper_Collide(parm)
	CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
	LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
	CheckLiveCatch Activeball, RightFlipper, RFCount, parm
	RightFlipperCollide parm
End Sub



'******************************************************
' Flippers Polarity 
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity


'*******************************************
' Late 70's to early 80's
Sub InitPolarity()
        dim x, a : a = Array(LF, RF)
        for each x in a
                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
                x.enabled = True
                x.TimeDelay = 80
        Next

        AddPt "Polarity", 0, 0, 0
        AddPt "Polarity", 1, 0.05, -2.7        
        AddPt "Polarity", 2, 0.33, -2.7
        AddPt "Polarity", 3, 0.37, -2.7        
        AddPt "Polarity", 4, 0.41, -2.7
        AddPt "Polarity", 5, 0.45, -2.7
        AddPt "Polarity", 6, 0.576,-2.7
        AddPt "Polarity", 7, 0.66, -1.8
        AddPt "Polarity", 8, 0.743, -0.5
        AddPt "Polarity", 9, 0.81, -0.5
        AddPt "Polarity", 10, 0.88, 0

        addpt "Velocity", 0, 0,         1
        addpt "Velocity", 1, 0.16, 1.06
        addpt "Velocity", 2, 0.41,         1.05
        addpt "Velocity", 3, 0.53,         1'0.982
        addpt "Velocity", 4, 0.702, 0.968
        addpt "Velocity", 5, 0.95,  0.968
        addpt "Velocity", 6, 1.03,         0.945

        LF.Object = LeftFlipper        
        LF.EndPoint = EndPointLp
        RF.Object = RightFlipper
        RF.EndPoint = EndPointRp
End Sub


' Flipper trigger hit subs
Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub




'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
	dim a : a = Array(LF, RF)
	dim x : for each x in a
		x.addpoint aStr, idx, aX, aY
	Next
End Sub

Class FlipperPolarity
	Public DebugOn, Enabled
	Private FlipAt        'Timer variable (IE 'flip at 723,530ms...)
	Public TimeDelay        'delay before trigger turns off and polarity is disabled TODO set time!
	private Flipper, FlipperStart,FlipperEnd, FlipperEndY, LR, PartialFlipCoef
	Private Balls(20), balldata(20)

	dim PolarityIn, PolarityOut
	dim VelocityIn, VelocityOut
	dim YcoefIn, YcoefOut
	Public Sub Class_Initialize 
		redim PolarityIn(0) : redim PolarityOut(0) : redim VelocityIn(0) : redim VelocityOut(0) : redim YcoefIn(0) : redim YcoefOut(0)
		Enabled = True : TimeDelay = 50 : LR = 1:  dim x : for x = 0 to uBound(balls) : balls(x) = Empty : set Balldata(x) = new SpoofBall : next 
	End Sub

	Public Property let Object(aInput) : Set Flipper = aInput : StartPoint = Flipper.x : End Property
	Public Property Let StartPoint(aInput) : if IsObject(aInput) then FlipperStart = aInput.x else FlipperStart = aInput : end if : End Property
	Public Property Get StartPoint : StartPoint = FlipperStart : End Property
	Public Property Let EndPoint(aInput) : FlipperEnd = aInput.x: FlipperEndY = aInput.y: End Property
	Public Property Get EndPoint : EndPoint = FlipperEnd : End Property        
	Public Property Get EndPointY: EndPointY = FlipperEndY : End Property

	Public Sub AddPoint(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out) 
		Select Case aChooseArray
			case "Polarity" : ShuffleArrays PolarityIn, PolarityOut, 1 : PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY : ShuffleArrays PolarityIn, PolarityOut, 0
			Case "Velocity" : ShuffleArrays VelocityIn, VelocityOut, 1 :VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY : ShuffleArrays VelocityIn, VelocityOut, 0
			Case "Ycoef" : ShuffleArrays YcoefIn, YcoefOut, 1 :YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY : ShuffleArrays YcoefIn, YcoefOut, 0
		End Select
		if gametime > 100 then Report aChooseArray
	End Sub 

	Public Sub Report(aChooseArray)         'debug, reports all coords in tbPL.text
		if not DebugOn then exit sub
		dim a1, a2 : Select Case aChooseArray
			case "Polarity" : a1 = PolarityIn : a2 = PolarityOut
			Case "Velocity" : a1 = VelocityIn : a2 = VelocityOut
			Case "Ycoef" : a1 = YcoefIn : a2 = YcoefOut 
				case else :tbpl.text = "wrong string" : exit sub
		End Select
		dim str, x : for x = 0 to uBound(a1) : str = str & aChooseArray & " x: " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
		tbpl.text = str
	End Sub

	Public Sub AddBall(aBall) : dim x : for x = 0 to uBound(balls) : if IsEmpty(balls(x)) then set balls(x) = aBall : exit sub :end if : Next  : End Sub

	Private Sub RemoveBall(aBall)
		dim x : for x = 0 to uBound(balls)
			if TypeName(balls(x) ) = "IBall" then 
				if aBall.ID = Balls(x).ID Then
					balls(x) = Empty
					Balldata(x).Reset
				End If
			End If
		Next
	End Sub

	Public Sub Fire() 
		Flipper.RotateToEnd
		processballs
	End Sub

	Public Property Get Pos 'returns % position a ball. For debug stuff.
		dim x : for x = 0 to uBound(balls)
			if not IsEmpty(balls(x) ) then
				pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
			End If
		Next                
	End Property

	Public Sub ProcessBalls() 'save data of balls in flipper range
		FlipAt = GameTime
		dim x : for x = 0 to uBound(balls)
			if not IsEmpty(balls(x) ) then
				balldata(x).Data = balls(x)
			End If
		Next
		PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
		PartialFlipCoef = abs(PartialFlipCoef-1)
	End Sub
	Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function        'Timer shutoff for polaritycorrect

	Public Sub PolarityCorrect(aBall)
		if FlipperOn() then 
			dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1

			'y safety Exit
			if aBall.VelY > -8 then 'ball going down
				RemoveBall aBall
				exit Sub
			end if

			'Find balldata. BallPos = % on Flipper
			for x = 0 to uBound(Balls)
				if aBall.id = BallData(x).id AND not isempty(BallData(x).id) then 
					idx = x
					BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
					if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                                'find safety coefficient 'ycoef' data
				end if
			Next

			If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
				BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
				if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                                                'find safety coefficient 'ycoef' data
			End If

			'Velocity correction
			if not IsEmpty(VelocityIn(0) ) then
				Dim VelCoef
				VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

				if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

				if Enabled then aBall.Velx = aBall.Velx*VelCoef
				if Enabled then aBall.Vely = aBall.Vely*VelCoef
			End If

			'Polarity Correction (optional now)
			if not IsEmpty(PolarityIn(0) ) then
				If StartPoint > EndPoint then LR = -1        'Reverse polarity if left flipper
				dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

				if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
			End If
		End If
		RemoveBall aBall
	End Sub
End Class

'******************************************************
'  FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS 
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
	dim x, aCount : aCount = 0
	redim a(uBound(aArray) )
	for x = 0 to uBound(aArray)        'Shuffle objects in a temp array
		if not IsEmpty(aArray(x) ) Then
			if IsObject(aArray(x)) then 
				Set a(aCount) = aArray(x)
			Else
				a(aCount) = aArray(x)
			End If
			aCount = aCount + 1
		End If
	Next
	if offset < 0 then offset = 0
	redim aArray(aCount-1+offset)        'Resize original array
	for x = 0 to aCount-1                'set objects back into original array
		if IsObject(a(x)) then 
			Set aArray(x) = a(x)
		Else
			aArray(x) = a(x)
		End If
	Next
End Sub

' Used for flipper correction and rubber dampeners
Sub ShuffleArrays(aArray1, aArray2, offset)
	ShuffleArray aArray1, offset
	ShuffleArray aArray2, offset
End Sub

' Used for flipper correction, rubber dampeners, and drop targets
Function BallSpeed(ball) 'Calculates the ball speed
	BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)        'Set up line via two points, no clamping. Input X, output Y
	dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
	Y = M*x+b
	PSlope = Y
End Function

' Used for flipper correction
Class spoofball 
	Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius 
	Public Property Let Data(aBall)
		With aBall
			x = .x : y = .y : z = .z : velx = .velx : vely = .vely : velz = .velz
			id = .ID : mass = .mass : radius = .radius
		end with
	End Property
	Public Sub Reset()
		x = Empty : y = Empty : z = Empty  : velx = Empty : vely = Empty : velz = Empty 
		id = Empty : mass = Empty : radius = Empty
	End Sub
End Class

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
	dim y 'Y output
	dim L 'Line
	dim ii : for ii = 1 to uBound(xKeyFrame)        'find active line
		if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
	Next
	if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)        'catch line overrun
	Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

	if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )         'Clamp lower
	if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )        'Clamp upper

	LinearEnvelope = Y
End Function


'******************************************************
'  FLIPPER TRICKS 
'******************************************************

RightFlipper.timerinterval=1
Rightflipper.timerenabled=True

sub RightFlipper_timer()
	FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
	FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
	FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
	FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
end sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
	Dim b, BOT
	BOT = GetBalls

	If Flipper1.currentangle = Endangle1 and EOSNudge1 <> 1 Then
		EOSNudge1 = 1
		'debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
		If Flipper2.currentangle = EndAngle2 Then 
			For b = 0 to Ubound(BOT)
				If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
					'Debug.Print "ball in flip1. exit"
					exit Sub
				end If
			Next
			For b = 0 to Ubound(BOT)
				If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
					BOT(b).velx = BOT(b).velx / 1.3
					BOT(b).vely = BOT(b).vely - 0.5
				end If
			Next
		End If
	Else 
		If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 then EOSNudge1 = 0
	End If
End Sub

'*****************
' Maths
'*****************
Dim PI: PI = 4*Atn(1)

Function dSin(degrees)
	dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
	dcos = cos(degrees * Pi/180)
End Function

Function Atn2(dy, dx)
	If dx > 0 Then
		Atn2 = Atn(dy / dx)
	ElseIf dx < 0 Then
		If dy = 0 Then 
			Atn2 = pi
		Else
			Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
		end if
	ElseIf dx = 0 Then
		if dy = 0 Then
			Atn2 = 0
		else
			Atn2 = Sgn(dy) * pi / 2
		end if
	End If
End Function

'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
	Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) ' Distance between a point and a line where point is px,py
	DistancePL = ABS((by - ay)*px - (bx - ax) * py + bx*ay - by*ax)/Distance(ax,ay,bx,by)
End Function

Function DistanceFast(x, y)
	dim ratio, ax, ay
	ax = abs(x)					'Get absolute value of each vector
	ay = abs(y)
	ratio = 1 / max(ax, ay)		'Create a ratio
	ratio = ratio * (1.29289 - (ax + ay) * ratio * 0.29289)
	if ratio > 0 then			'Quickly determine if it's worth using
		DistanceFast = 1/ratio
	Else
		DistanceFast = 0
	End if
end Function

Function max(a,b)
	if a > b then 
		max = a
	Else
		max = b
	end if
end Function

Function Radians(Degrees)
	Radians = Degrees * PI /180
End Function

Function AnglePP(ax,ay,bx,by)
	AnglePP = Atn2((by - ay),(bx - ax))*180/PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
	DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle+90))+Flipper.x, Sin(Radians(Flipper.currentangle+90))+Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
	Dim DiffAngle
	DiffAngle  = ABS(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
	If DiffAngle > 180 Then DiffAngle = DiffAngle - 360

	If DistanceFromFlipper(ballx,bally,Flipper) < 48 and DiffAngle <= 90 and Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
		FlipperTrigger = True
	Else
		FlipperTrigger = False
	End If        
End Function


'*************************************************
'  End - Check ball distance from Flipper for Rem
'*************************************************

dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 1 'EM's to late 80's
Const EOSAnew = 1
Const EOSRampup = 0
Const SOSRampup = 3

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.055  'EM's

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
	FlipperPress = 1
	Flipper.Elasticity = FElasticity

	Flipper.eostorque = EOST         
	Flipper.eostorqueangle = EOSA         
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
	FlipperPress = 0
	Flipper.eostorqueangle = EOSA
	Flipper.eostorque = EOST*EOSReturn/FReturn


	If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
		Dim b, BOT
		BOT = GetBalls

		For b = 0 to UBound(BOT)
			If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
				If BOT(b).vely >= -0.4 Then BOT(b).vely = -0.4
			End If
		Next
	End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState) 
	Dim Dir
	Dir = Flipper.startangle/Abs(Flipper.startangle)        '-1 for Right Flipper

	If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
		If FState <> 1 Then
			Flipper.rampup = SOSRampup 
			Flipper.endangle = FEndAngle - 3*Dir
			Flipper.Elasticity = FElasticity * SOSEM
			FCount = 0 
			FState = 1
		End If
	ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) and FlipperPress = 1 then
		if FCount = 0 Then FCount = GameTime

		If FState <> 2 Then
			Flipper.eostorqueangle = EOSAnew
			Flipper.eostorque = EOSTnew
			Flipper.rampup = EOSRampup                        
			Flipper.endangle = FEndAngle
			FState = 2
		End If
	Elseif Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 and FlipperPress = 1 Then 
		If FState <> 3 Then
			Flipper.eostorque = EOST        
			Flipper.eostorqueangle = EOSA
			Flipper.rampup = Frampup
			Flipper.Elasticity = FElasticity
			FState = 3
		End If

	End If
End Sub

Const LiveDistanceMin = 30  'minimum distance in vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114  'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
	Dim Dir
	Dir = Flipper.startangle/Abs(Flipper.startangle)    '-1 for Right Flipper
	Dim LiveCatchBounce                                                                                                                        'If live catch is not perfect, it won't freeze ball totally
	Dim CatchTime : CatchTime = GameTime - FCount

	if CatchTime <= LiveCatch and parm > 6 and ABS(Flipper.x - ball.x) > LiveDistanceMin and ABS(Flipper.x - ball.x) < LiveDistanceMax Then
		if CatchTime <= LiveCatch*0.5 Then                                                'Perfect catch only when catch time happens in the beginning of the window
			LiveCatchBounce = 0
		else
			LiveCatchBounce = Abs((LiveCatch/2) - CatchTime)        'Partial catch when catch happens a bit late
		end If

		If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
		ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
		ball.angmomx= 0
		ball.angmomy= 0
		ball.angmomz= 0
    Else
        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf Activeball, parm
	End If
End Sub


'******************************************************
'****  END FLIPPER CORRECTIONS
'******************************************************








'******************************************************
'****  PHYSICS DAMPENERS
'******************************************************
'
' These are data mined bounce curves, 
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR



Sub dPosts_Hit(idx) 
	RubbersD.dampen Activeball
	TargetBouncer Activeball, 1
End Sub

Sub dSleeves_Hit(idx) 
	SleevesD.Dampen Activeball
	TargetBouncer Activeball, 0.7
End Sub


dim RubbersD : Set RubbersD = new Dampener        'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False        'shows info in textbox "TBPout"
RubbersD.Print = False        'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.1        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967        'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64        'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener        'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False        'shows info in textbox "TBPout"
SleevesD.Print = False        'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'#########################    Adjust these values to increase or lessen the elasticity

dim FlippersD : Set FlippersD = new Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False	
FlippersD.addpoint 0, 0, 1.1	
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
	Public Print, debugOn 'tbpOut.text
	public name, Threshold         'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
	Public ModIn, ModOut
	Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): End Sub 

	Public Sub AddPoint(aIdx, aX, aY) 
		ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
		if gametime > 100 then Report
	End Sub

	public sub Dampen(aBall)
		if threshold then if BallSpeed(aBall) < threshold then exit sub end if end if
		dim RealCOR, DesiredCOR, str, coef
		DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
		RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
		coef = desiredcor / realcor 
		if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
		"actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline 
		if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)

		aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
		if debugOn then TBPout.text = str
	End Sub

	public sub Dampenf(aBall, parm) 'Rubberizer is handle here
		dim RealCOR, DesiredCOR, str, coef
		DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
		RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
		coef = desiredcor / realcor 
		If abs(aball.velx) < 2 and aball.vely < 0 and aball.vely > -3.75 then 
			aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
		End If
	End Sub

	Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
		dim x : for x = 0 to uBound(aObj.ModIn)
			addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
		Next
	End Sub


	Public Sub Report()         'debug, reports all coords in tbPL.text
		if not debugOn then exit sub
		dim a1, a2 : a1 = ModIn : a2 = ModOut
		dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
		TBPout.text = str
	End Sub

End Class



'******************************************************
'  TRACK ALL BALL VELOCITIES
'  FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

dim cor : set cor = New CoRTracker

Class CoRTracker
	public ballvel, ballvelx, ballvely

	Private Sub Class_Initialize : redim ballvel(0) : redim ballvelx(0): redim ballvely(0) : End Sub 

	Public Sub Update()	'tracks in-ball-velocity
		dim str, b, AllBalls, highestID : allBalls = getballs

		for each b in allballs
			if b.id >= HighestID then highestID = b.id
		Next

		if uBound(ballvel) < highestID then redim ballvel(highestID)	'set bounds
		if uBound(ballvelx) < highestID then redim ballvelx(highestID)	'set bounds
		if uBound(ballvely) < highestID then redim ballvely(highestID)	'set bounds

		for each b in allballs
			ballvel(b.id) = BallSpeed(b)
			ballvelx(b.id) = b.velx
			ballvely(b.id) = b.vely
		Next
	End Sub
End Class



'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************




'******************************************************
' VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1 		'0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7 	'Level of bounces. Recommmended value of 0.7 when TargetBouncerEnabled=1

sub TargetBouncer(aBall,defvalue)
    dim zMultiplier, vel, vratio
    if TargetBouncerEnabled = 1 and aball.z < 30 then
        'debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
        vel = BallSpeed(aBall)
        if aBall.velx = 0 then vratio = 1 else vratio = aBall.vely/aBall.velx
        Select Case Int(Rnd * 6) + 1
            Case 1: zMultiplier = 0.2*defvalue
			Case 2: zMultiplier = 0.25*defvalue
            Case 3: zMultiplier = 0.3*defvalue
			Case 4: zMultiplier = 0.4*defvalue
            Case 5: zMultiplier = 0.45*defvalue
            Case 6: zMultiplier = 0.5*defvalue
        End Select
        aBall.velz = abs(vel * zMultiplier * TargetBouncerFactor)
        aBall.velx = sgn(aBall.velx) * sqr(abs((vel^2 - aBall.velz^2)/(1+vratio^2)))
        aBall.vely = aBall.velx * vratio
        'debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
        'debug.print "conservation check: " & BallSpeed(aBall)/vel
	end if
end sub



'******************************************************
'  DROP TARGETS INITIALIZATION
'******************************************************

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

'Define a variable for each drop target
Dim DTA1v,DTA2v,DTJKv,DTA3v,DTA4v,DT10v,DTJ1v,DTJ2v,DTQ1v,DTQ2v,DTQ3v,DTK1v,DTK2v,DTK3v,DTK4v

'Set array with drop target objects
'
'DropTargetvar = Array(primary, secondary, prim, swtich, animate)
' 	primary: 			primary target wall to determine drop
'	secondary:			wall used to simulate the ball striking a bent or offset target after the initial Hit
'	prim:				primitive target used for visuals and animation
'							IMPORTANT!!! 
'							rotz must be used for orientation
'							rotx to bend the target back
'							transz to move it up and down
'							the pivot point should be in the center of the target on the x, y and at or below the playfield (0) on z
'	switch:				ROM switch number
'	animate:			Arrary slot for handling the animation instrucitons, set to 0
'
'	Values for annimate: 1 - bend target (hit to primary), 2 - drop target (hit to secondary), 3 - brick target (high velocity hit to secondary), -1 - raise target 

Set DTA1v = (new DropTarget)(DTA1, DTA1a, DTA1p, 1, 0, false)
Set DTA2v = (new DropTarget)(DTA2, DTA2a, DTA2p, 2, 0, false)
Set DTJKv = (new DropTarget)(DTJK, DTJKa, DTJKp, 3, 0, false)
Set DTA3v = (new DropTarget)(DTA3, DTA3a, DTA3p, 4, 0, false)
Set DTA4v = (new DropTarget)(DTA4, DTA4a, DTA4p, 5, 0, false)
Set DT10v = (new DropTarget)(DT10, DT10a, DT10p, 6, 0, false)
Set DTJ1v = (new DropTarget)(DTJ1, DTJ1a, DTJ1p, 7, 0, false)
Set DTJ2v = (new DropTarget)(DTJ2, DTJ2a, DTJ2p, 8, 0, false)
Set DTQ1v = (new DropTarget)(DTQ1, DTQ1a, DTQ1p, 9, 0, false)
Set DTQ2v = (new DropTarget)(DTQ2, DTQ2a, DTQ2p, 10, 0, false)
Set DTQ3v = (new DropTarget)(DTQ3, DTQ3a, DTQ3p, 11, 0, false)
Set DTK1v = (new DropTarget)(DTK1, DTK1a, DTK1p, 12, 0, false)
Set DTK2v = (new DropTarget)(DTK2, DTK2a, DTK2p, 13, 0, false)
Set DTK3v = (new DropTarget)(DTK3, DTK3a, DTK3p, 14, 0, false)
Set DTK4v = (new DropTarget)(DTK4, DTK4a, DTK4p, 15, 0, false)


Dim DTArray
DTArray = Array(DTA1v,DTA2v,DTJKv,DTA3v,DTA4v,DT10v,DTJ1v,DTJ2v,DTQ1v,DTQ2v,DTQ3v,DTK1v,DTK2v,DTK3v,DTK4v)
Dim DTIsDropped(15)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 90 'in milliseconds
Const DTDropUpSpeed = 50 'in milliseconds
Const DTDropUnits = 44 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick

Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
Const DTHitSound = "" 'Drop Target Hit sound
Const DTDropSound = "DropTarget_Down" 'Drop Target Drop sound
Const DTResetSound = "DropTarget_Up" 'Drop Target reset sound

Const DTMass = 0.2 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance


'******************************************************
'  DROP TARGETS FUNCTIONS
'******************************************************

Sub DTHit(switch)
	Dim i
	i = DTArrayID(switch)

	PlayTargetSound
	DTArray(i).animate =  DTCheckBrick(Activeball,DTArray(i).prim)
	If DTArray(i).animate = 1 or DTArray(i).animate = 3 or DTArray(i).animate = 4 Then
		DTBallPhysics Activeball, DTArray(i).prim.rotz, DTMass
	End If
	DoDTAnim
End Sub

Sub DTRaise(switch)
	Dim i
	i = DTArrayID(switch)
	DTIsDropped(switch) = 0

	DTArray(i).animate = -1
	DoDTAnim
End Sub

Sub DTDrop(switch)
	Dim i
	i = DTArrayID(switch)
	
	DTArray(i).animate = 1
	DoDTAnim
End Sub

Function DTArrayID(switch)
	Dim i
	For i = 0 to uBound(DTArray) 
		If DTArray(i).sw = switch Then DTArrayID = i:Exit Function 
	Next
End Function


sub DTBallPhysics(aBall, angle, mass)
	dim rangle,bangle,calc1, calc2, calc3
	rangle = (angle - 90) * 3.1416 / 180
	bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))

	calc1 = cor.BallVel(aball.id) * cos(bangle - rangle) * (aball.mass - mass) / (aball.mass + mass)
	calc2 = cor.BallVel(aball.id) * sin(bangle - rangle) * cos(rangle + 4*Atn(1)/2)
	calc3 = cor.BallVel(aball.id) * sin(bangle - rangle) * sin(rangle + 4*Atn(1)/2)

	aBall.velx = calc1 * cos(rangle) + calc2
	aBall.vely = calc1 * sin(rangle) + calc3
End Sub


'Check if target is hit on it's face or sides and whether a 'brick' occurred
Function DTCheckBrick(aBall, dtprim) 
	dim bangle, bangleafter, rangle, rangle2, Xintersect, Yintersect, cdist, perpvel, perpvelafter, paravel, paravelafter
	rangle = (dtprim.rotz - 90) * 3.1416 / 180
	rangle2 = dtprim.rotz * 3.1416 / 180
	bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
	bangleafter = Atn2(aBall.vely,aball.velx)

	Xintersect = (aBall.y - dtprim.y - tan(bangle) * aball.x + tan(rangle2) * dtprim.x) / (tan(rangle2) - tan(bangle))
	Yintersect = tan(rangle2) * Xintersect + (dtprim.y - tan(rangle2) * dtprim.x)

	cdist = Distance(dtprim.x, dtprim.y, Xintersect, Yintersect)

	perpvel = cor.BallVel(aball.id) * cos(bangle-rangle)
	paravel = cor.BallVel(aball.id) * sin(bangle-rangle)

	perpvelafter = BallSpeed(aBall) * cos(bangleafter - rangle) 
	paravelafter = BallSpeed(aBall) * sin(bangleafter - rangle)

	If perpvel > 0 and  perpvelafter <= 0 Then
		If DTEnableBrick = 1 and  perpvel > DTBrickVel and DTBrickVel <> 0 and cdist < 8 Then
			DTCheckBrick = 3
		Else
			DTCheckBrick = 1
		End If
	ElseIf perpvel > 0 and ((paravel > 0 and paravelafter > 0) or (paravel < 0 and paravelafter < 0)) Then
		DTCheckBrick = 4
	Else 
		DTCheckBrick = 0
	End If
End Function


Sub DoDTAnim()
	Dim i
	For i=0 to Ubound(DTArray)
		DTArray(i).animate = DTAnimate(DTArray(i).primary,DTArray(i).secondary,DTArray(i).prim,DTArray(i).sw,DTArray(i).animate)
	Next
End Sub

Function DTAnimate(primary, secondary, prim, switch,  animate)
	dim transz, switchid
	Dim animtime, rangle

	switchid = switch

	rangle = prim.rotz * PI / 180

	DTAnimate = animate

	if animate = 0  Then
		primary.uservalue = 0
		DTAnimate = 0
		Exit Function
	Elseif primary.uservalue = 0 then 
		primary.uservalue = gametime
	end if

	animtime = gametime - primary.uservalue

	If (animate = 1 or animate = 4) and animtime < DTDropDelay Then
		primary.collidable = 0
	If animate = 1 then secondary.collidable = 1 else secondary.collidable= 0
		prim.rotx = DTMaxBend * cos(rangle)
		prim.roty = DTMaxBend * sin(rangle)
		DTAnimate = animate
		Exit Function
		elseif (animate = 1 or animate = 4) and animtime > DTDropDelay Then
		primary.collidable = 0
		If animate = 1 then secondary.collidable = 1 else secondary.collidable= 0
		prim.rotx = DTMaxBend * cos(rangle)
		prim.roty = DTMaxBend * sin(rangle)
		animate = 2
		SoundDropTargetDrop prim
	End If

	if animate = 2 Then
		transz = (animtime - DTDropDelay)/DTDropSpeed *  DTDropUnits * -1
		if prim.transz > -DTDropUnits  Then
			prim.transz = transz
		end if

		prim.rotx = DTMaxBend * cos(rangle)/2
		prim.roty = DTMaxBend * sin(rangle)/2

		if prim.transz <= -DTDropUnits Then 
			prim.transz = -DTDropUnits
			secondary.collidable = 0
			DTAction switchid
			primary.uservalue = 0
			DTAnimate = 0
			Exit Function
		Else
			DTAnimate = 2
			Exit Function
		end If 
	End If

	If animate = 3 and animtime < DTDropDelay Then
		primary.collidable = 0
		secondary.collidable = 1
		prim.rotx = DTMaxBend * cos(rangle)
		prim.roty = DTMaxBend * sin(rangle)
	elseif animate = 3 and animtime > DTDropDelay Then
		primary.collidable = 1
		secondary.collidable = 0
		prim.rotx = 0
		prim.roty = 0
		primary.uservalue = 0
		DTAnimate = 0
		Exit Function
	End If

	if animate = -1 Then
		transz = (1 - (animtime)/DTDropUpSpeed) *  DTDropUnits * -1

		If prim.transz = -DTDropUnits Then
			Dim BOT, b
			BOT = GetBalls

			For b = 0 to UBound(BOT)
				If InRotRect(BOT(b).x,BOT(b).y,prim.x, prim.y, prim.rotz, -25,-10,25,-10,25,25,-25,25) and BOT(b).z < prim.z+DTDropUnits+25 Then
					BOT(b).velz = 20
				End If
			Next
		End If

		if prim.transz < 0 Then
			prim.transz = transz
		elseif transz > 0 then
			prim.transz = transz
		end if

		if prim.transz > DTDropUpUnits then 
			DTAnimate = -2
			prim.transz = DTDropUpUnits
			prim.rotx = 0
			prim.roty = 0
			primary.uservalue = gametime
		end if
		primary.collidable = 0
		secondary.collidable = 1

	End If

	if animate = -2 and animtime > DTRaiseDelay Then
		prim.transz = (animtime - DTRaiseDelay)/DTDropSpeed *  DTDropUnits * -1 + DTDropUpUnits 
		if prim.transz < 0 then
			prim.transz = 0
			primary.uservalue = 0
			DTAnimate = 0

			primary.collidable = 1
			secondary.collidable = 0
		end If 
	End If
End Function

Sub DTAction(switchid)
	DTIsDropped(switchid) = 1
	Select Case switchid
		Case 1:
			DOF 114,DOFPulse
			Ldta1.state=1
			addscore 1000
			CheckAwards
		Case 2:
			DOF 114,DOFPulse
			LdtA2.state=1
			Ldta2b.state=1
			addscore 1000
			CheckAwards
		Case 3:
			DOF 114,DOFPulse
			Ldtjk.state=1
			Ldtjkb.state=1
			addscore 1000	
			CheckAwards
		Case 4:
			DOF 114,DOFPulse
			LdtA3.state=1
			LdtA3b.state=1
			addscore 1000
			CheckAwards
		Case 5:
			DOF 114,DOFPulse
			LdtA4.state=1
			addscore 1000	
			CheckAwards
		Case 6:
			DOF 118,DOFPulse
			addscore 500
			lbonus1.state=1
		Case 7:
			DOF 116,DOFPulse
			addscore 500
			CheckAwards
		Case 8:
			DOF 116,DOFPulse
			addscore 500
			CheckAwards
		Case 9:
			DOF 116,DOFPulse
			addscore 500	
			CheckAwards
		Case 10:
			DOF 116,DOFPulse
			Ldtq2.state=1
			addscore 500
			CheckAwards
		Case 11:
			DOF 116,DOFPulse
			Ldtq3.state=1	
			addscore 500
			CheckAwards
		Case 12:
			DOF 117,DOFPulse
			LdtK1.state=1
			addscore 1000
			CheckAwards
		Case 13:
			DOF 117,DOFPulse
			LdtK2.state=1	
			addscore 1000
			CheckAwards
		Case 14:
			DOF 117,DOFPulse
			LdtK3.state=1
			addscore 1000
			CheckAwards
		Case 15:
			DOF 117,DOFPulse
			LdtK4.state=1
			addscore 1000	
			CheckAwards

	End Select
End Sub




'******************************************************
'  DROP TARGET
'  SUPPORTING FUNCTIONS 
'******************************************************


' Used for drop targets
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

Function InRotRect(ballx,bally,px,py,angle,ax,ay,bx,by,cx,cy,dx,dy)
    Dim rax,ray,rbx,rby,rcx,rcy,rdx,rdy
    Dim rotxy
    rotxy = RotPoint(ax,ay,angle)
    rax = rotxy(0)+px : ray = rotxy(1)+py
    rotxy = RotPoint(bx,by,angle)
    rbx = rotxy(0)+px : rby = rotxy(1)+py
    rotxy = RotPoint(cx,cy,angle)
    rcx = rotxy(0)+px : rcy = rotxy(1)+py
    rotxy = RotPoint(dx,dy,angle)
    rdx = rotxy(0)+px : rdy = rotxy(1)+py

    InRotRect = InRect(ballx,bally,rax,ray,rbx,rby,rcx,rcy,rdx,rdy)
End Function

Function RotPoint(x,y,angle)
    dim rx, ry
    rx = x*dCos(angle) - y*dSin(angle)
    ry = x*dSin(angle) + y*dCos(angle)
    RotPoint = Array(rx,ry)
End Function


'******************************************************
'****  END DROP TARGETS
'******************************************************





'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************


'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1														'volume level; range [0, 1]
NudgeLeftSoundLevel = 1													'volume level; range [0, 1]
NudgeRightSoundLevel = 1												'volume level; range [0, 1]
NudgeCenterSoundLevel = 1												'volume level; range [0, 1]
StartButtonSoundLevel = 0.1												'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr											'volume level; range [0, 1]
PlungerPullSoundLevel = 1												'volume level; range [0, 1]
RollingSoundFactor = 1.1/5		

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010           						'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635								'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0                        						'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45                      						'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel								'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel								'sound helper; not configurable
SlingshotSoundLevel = 0.95												'volume level; range [0, 1]
BumperSoundFactor = 4.25												'volume multiplier; must not be zero
KnockerSoundLevel = 1 													'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2									'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055/5											'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075/5											'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075/5										'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025									'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025									'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8									'volume level; range [0, 1]
WallImpactSoundFactor = 0.075											'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075/3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5/5													'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10											'volume multiplier; must not be zero
DTSoundLevel = 0.25														'volume multiplier; must not be zero
RolloverSoundLevel = 0.25                              					'volume level; range [0, 1]
SpinnerSoundLevel = 0.5                              					'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor 

DrainSoundLevel = 0.8														'volume level; range [0, 1]
BallReleaseSoundLevel = 1												'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2									'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015										'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025/5													'volume multiplier; must not be zero


'/////////////////////////////  SOUND PLAYBACK FUNCTIONS  ////////////////////////////
'/////////////////////////////  POSITIONAL SOUND PLAYBACK METHODS  ////////////////////////////
' Positional sound playback methods will play a sound, depending on the X,Y position of the table element or depending on ActiveBall object position
' These are similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
' For surround setup - positional sound playback functions will fade between front and rear surround channels and pan between left and right channels
' For stereo setup - positional sound playback functions will only pan between left and right channels
' For mono setup - positional sound playback functions will not pan between left and right channels and will not fade between front and rear channels

' PlaySound full syntax - PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
' Note - These functions will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlaySoundAtLevelStatic(playsoundparams, aVol, tableobj)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
	PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
	PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
	PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
	PlaySound soundname, 1, aVol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
	PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
	Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
	Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
	PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


'******************************************************
'  Fleep  Supporting Ball & Sound Functions
'******************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
    tmp = tableobj.y * 2 / tableheight-1

	if tmp > 7000 Then
		tmp = 7000
	elseif tmp < -7000 Then
		tmp = -7000
	end if

    If tmp > 0 Then
		AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / tablewidth-1

	if tmp > 7000 Then
		tmp = 7000
	elseif tmp < -7000 Then
		tmp = -7000
	end if

    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
	Vol = Csng(BallVel(ball) ^2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
	Volz = Csng((ball.velz) ^2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
	Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
	BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
	VolPlayfieldRoll = RollingSoundFactor * 0.0005 * Csng(BallVel(ball) ^2.5)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
	PitchPlayfieldRoll = BallVel(ball) ^2 * 3
End Function

Function RndInt(min, max)
	RndInt = Int(Rnd() * (max-min + 1) + min)' Sets a random number integer between min and max
End Function

Function RndNum(min, max)
	RndNum = Rnd() * (max-min) + min' Sets a random number between min and max
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////
Sub SoundStartButton()
	PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
	PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
End Sub

Sub SoundNudgeRight()
	PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
	PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
End Sub


Sub SoundPlungerPull()
	PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseBall()
	PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, Plunger	
End Sub

Sub SoundPlungerReleaseNoBall()
	PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub


'/////////////////////////////  KNOCKER SOLENOID  ////////////////////////////
Sub KnockerSolenoid()
	PlaySoundAtLevelStatic SoundFX("Knocker_1",DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////
Sub RandomSoundDrain(drainswitch)
	PlaySoundAtLevelStatic ("Drain_" & Int(Rnd*11)+1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
	PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd*7)+1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundSlingshotLeft(sling)
	PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd*10)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
	PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd*8)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundBumperTop(Bump)
	PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
	PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
	PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  SPINNER SOUNDS  ////////////////////////////
Sub SoundSpinner(spinnerswitch)
	PlaySoundAtLevelStatic ("Spinner"), SpinnerSoundLevel, spinnerswitch
End Sub


'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////
Sub SoundFlipperUpAttackLeft(flipper)
	FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
	PlaySoundAtLevelStatic ("Flipper_Attack-L01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
	FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
	PlaySoundAtLevelStatic ("Flipper_Attack-R01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////
Sub RandomSoundFlipperUpLeft(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd*9)+1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd*9)+1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd*7)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd*8)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
	FlipperLeftHitParm = parm/10
	If FlipperLeftHitParm > 1 Then
		FlipperLeftHitParm = 1
	End If
	FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
	RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
	FlipperRightHitParm = parm/10
	If FlipperRightHitParm > 1 Then
		FlipperRightHitParm = 1
	End If
	FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
	RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
	PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd*7)+1), parm  * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////
Sub RandomSoundRollover()
	PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd*4)+1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
	RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////
Sub Rubbers_Hit(idx)
	dim finalspeed
	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
	If finalspeed > 5 then		
		RandomSoundRubberStrong 1
	End if
	If finalspeed <= 5 then
		RandomSoundRubberWeak()
	End If	
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////
Sub RandomSoundRubberStrong(voladj)
	Select Case Int(Rnd*10)+1
		Case 1 : PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
		Case 2 : PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
		Case 3 : PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
		Case 4 : PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
		Case 5 : PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
		Case 6 : PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
		Case 7 : PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
		Case 8 : PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
		Case 9 : PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
		Case 10 : PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6*voladj
	End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////
Sub RandomSoundRubberWeak()
	PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd*9)+1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////
Sub Walls_Hit(idx)
	RandomSoundWall()      
End Sub

Sub RandomSoundWall()
	dim finalspeed
	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
	If finalspeed > 16 then 
		Select Case Int(Rnd*5)+1
			Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 5 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
		End Select
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
		Select Case Int(Rnd*4)+1
			Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
		End Select
	End If
	If finalspeed < 6 Then
		Select Case Int(Rnd*3)+1
			Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
		End Select
	End if
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////
Sub RandomSoundMetal()
	PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd*13)+1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

'/////////////////////////////  METAL - EVENTS  ////////////////////////////

Sub Metals_Hit (idx)
	RandomSoundMetal
End Sub

Sub ShooterDiverter_collide(idx)
	RandomSoundMetal
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE  ////////////////////////////
'/////////////////////////////  BOTTOM ARCH BALL GUIDE - SOFT BOUNCES  ////////////////////////////
Sub RandomSoundBottomArchBallGuide()
	dim finalspeed
	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
	If finalspeed > 16 then 
		PlaySoundAtLevelActiveBall ("Apron_Bounce_"& Int(Rnd*2)+1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
		Select Case Int(Rnd*2)+1
			Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
			Case 2 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
		End Select
	End If
	If finalspeed < 6 Then
		Select Case Int(Rnd*2)+1
			Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
			Case 2 : PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
		End Select
	End if
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////
Sub RandomSoundBottomArchBallGuideHardHit()
	PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd*3)+1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
	If Abs(cor.ballvelx(activeball.id) < 4) and cor.ballvely(activeball.id) > 7 then
		RandomSoundBottomArchBallGuideHardHit()
	Else
		RandomSoundBottomArchBallGuide
	End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////
Sub RandomSoundFlipperBallGuide()
	dim finalspeed
	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
	If finalspeed > 16 then 
		Select Case Int(Rnd*2)+1
			Case 1 : PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
			Case 2 : PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
		End Select
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
		PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd*3)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
	End If
	If finalspeed < 6 Then
		PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd*7)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
	End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////
Sub RandomSoundTargetHitStrong()
	PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()		
	PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
	dim finalspeed
	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
	If finalspeed > 10 then
		RandomSoundTargetHitStrong()
		RandomSoundBallBouncePlayfieldSoft Activeball
	Else 
		RandomSoundTargetHitWeak()
	End If	
End Sub

Sub Targets_Hit (idx)
	PlayTargetSound	
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////
Sub RandomSoundBallBouncePlayfieldSoft(aBall)
	Select Case Int(Rnd*9)+1
		Case 1 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
		Case 2 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
		Case 3 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
		Case 4 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
		Case 5 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
		Case 6 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
		Case 7 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
		Case 8 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
		Case 9 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
	End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
	PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd*7)+1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////
Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
	Select Case Int(Rnd*5)+1
		Case 1 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 2 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 3 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 4 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 5 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
	End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()			
	PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd*2)+1), GateSoundLevel, Activeball
End Sub

Sub SoundHeavyGate()
	PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, Activeball
End Sub

Sub Gates_hit(idx)
	SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)	
	SoundPlayfieldGate	
End Sub	

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
	PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
	PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub


Sub Arch1_hit()
	If Activeball.velx > 1 Then SoundPlayfieldGate
	StopSound "Arch_L1"
	StopSound "Arch_L2"
	StopSound "Arch_L3"
	StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
	If activeball.velx < -8 Then
		RandomSoundRightArch
	End If
End Sub

Sub Arch2_hit()
	If Activeball.velx < 1 Then SoundPlayfieldGate
	StopSound "Arch_R1"
	StopSound "Arch_R2"
	StopSound "Arch_R3"
	StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
	If activeball.velx > 10 Then
		RandomSoundLeftArch
	End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
	PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd*2)+1), SaucerLockSoundLevel, Activeball
End Sub

Sub SoundSaucerKick(scenario, saucer)
	Select Case scenario
		Case 0: PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
		Case 1: PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
	End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////
Sub OnBallBallCollision(ball1, ball2, velocity)
	Dim snd
	Select Case Int(Rnd*7)+1
		Case 1 : snd = "Ball_Collide_1"
		Case 2 : snd = "Ball_Collide_2"
		Case 3 : snd = "Ball_Collide_3"
		Case 4 : snd = "Ball_Collide_4"
		Case 5 : snd = "Ball_Collide_5"
		Case 6 : snd = "Ball_Collide_6"
		Case 7 : snd = "Ball_Collide_7"
	End Select

	PlaySound (snd), 0, Csng(velocity) ^2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
	PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd*6)+1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
	PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd*6)+1), 200, obj
End Sub

'/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

Const RelayFlashSoundLevel = 0.315									'volume level; range [0, 1];
Const RelayGISoundLevel = 1.05									'volume level; range [0, 1];

Sub Sound_GI_Relay(toggle, obj)
	Select Case toggle
		Case 1
			PlaySoundAtLevelStatic ("Relay_GI_On"), 0.025*RelayGISoundLevel, obj
		Case 0
			PlaySoundAtLevelStatic ("Relay_GI_Off"), 0.025*RelayGISoundLevel, obj
	End Select
End Sub

Sub Sound_Flash_Relay(toggle, obj)
	Select Case toggle
		Case 1
			PlaySoundAtLevelStatic ("Relay_Flash_On"), 0.025*RelayFlashSoundLevel, obj			
		Case 0
			PlaySoundAtLevelStatic ("Relay_Flash_Off"), 0.025*RelayFlashSoundLevel, obj		
	End Select
End Sub

'/////////////////////////////////////////////////////////////////
'					End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************



'//////////////////// PINUP PLAYER: STARTUP & CONTROL SECTION //////////////////////////

' This is used for the startup and control of Pinup Player

Sub PuPStart(cPuPPack)
    If PUPStatus=true then Exit Sub
    If usePUP=true then
        Set PuPlayer = CreateObject("PinUpPlayer.PinDisplay")
        If PuPlayer is Nothing Then
            usePUP=false
            PUPStatus=false
        Else
            PuPlayer.B2SInit "",cPuPPack 'start the Pup-Pack
            PUPStatus=true
        End If
    End If
End Sub

Sub pupevent(EventNum)
    if (usePUP=false or PUPStatus=false) then Exit Sub
    PuPlayer.B2SData "D"&EventNum,1  'send event to Pup-Pack
End Sub

' ******* How to use PUPEvent to trigger / control a PuP-Pack *******

' Usage: pupevent(EventNum)

' EventNum = PuP Exxx trigger from the PuP-Pack

' Example: pupevent 102

' This will trigger E102 from the table's PuP-Pack

' DO NOT use any Exxx triggers already used for DOF (if used) to avoid any possible confusion

'************ PuP-Pack Startup **************

PuPStart(cPuPPack) 'Check for PuP - If found, then start Pinup Player / PuP-Pack

pupflasher.VideoCapWidth= 300	
pupflasher.VideoCapHeight=300

Sub pupflasher_Timer()
	pupflasher.VideoCapUpdate="PUPSCREEN11"
End Sub

'************ Select Character *************

Sub CharSelectDelay
       playsound "cluper"
       pupevent 650
End Sub


Sub DelayedStart
      SoundStartButton
	  if state=true then
		player=1
		ballinplay=1
		If B2SOn Then 
			if freeplay=0 then Controller.B2ssetCredits Credit
			Controller.B2ssetballinplay 32, Ballinplay
            Controller.B2SSetData 91,1
			Controller.B2ssetplayerup 1
			Controller.B2ssetcanplay 31, 1
			Controller.B2SSetGameOver 0
			Controller.B2SSetScorePlayer 5, hisc
		End If
	    pup1.state=1
		tilt=false
		state=true
		playsound "initialize"
        EndVideoTimer.Enabled = False
        Strip1=False
        Strip2=False
        Strip3=False
        Strip4=False
        Strip5=False
        Strip6=False
        Strip7=False
        Strip8=False
        Strip9=False
        Strip10=False
        PlayerList=0
		PlaySound("RotateThruPlayers"),0,.05,0,0.25
 		TempPlayerUp=Player
		PlayerUpRotator.enabled=true
		players=1
        CoinLockOut=0
		CANPLAY1.state=1
		rst=0
		resettimer.enabled=true
	  else if state=true and (Credit>0 or freeplay=1) and players < maxplayers and Ballinplay=1 then
		if freeplay=0 then
			credit=credit-1
			If credit < 1 Then DOF 113, DOFOff
			credittxt.setvalue(credit)
		end if
		EVAL("Canplay"&players).state=0
		'players=players+1
		'EVAL("Canplay"&players).state=1
		If B2SOn then
			Controller.B2ssetCredits Credit
			'Controller.B2ssetcanplay 31, players
		End If
		playsound "cluper" 
	   end if 
	  end if
End Sub


Sub PupScoreMonitor
     If Score(player) >= Video1 AND Strip1 = False Then
        PlayerList=1
        PlayerLineUp
     End If

     If Score(player) >= Video2 AND Strip2 = False Then
        PlayerList=2
        PlayerLineUp
     End If

     If Score(player) >= Video3 AND Strip3 = False Then
        PlayerList=3
        PlayerLineUp
     End If

     If Score(player) >= Video4 AND Strip4 = False Then
        PlayerList=4
        PlayerLineUp
     End If

     If Score(player) >= Video5 AND Strip5 = False Then
        PlayerList=5
        PlayerLineUp
     End If

     If Score(player) >= Video6 AND Strip6 = False Then
        PlayerList=6
        PlayerLineUp
     End If

     If Score(player) >= Video7 AND Strip7 = False Then
        PlayerList=7
        PlayerLineUp
     End If

     If Score(player) >= Video8 AND Strip8 = False Then
        PlayerList=8
        PlayerLineUp
     End If

     If Score(player) >= Video9 AND Strip9 = False Then
        PlayerList=9
        PlayerLineUp
     End If

     If Score(player) >= Video10 AND Strip10 = False Then
        PlayerList=10
        PlayerLineUp
     End If
End Sub

Sub CharacterCalls
     If ActiveChar=1 then PupCall=50: pupevent 450: End If
     If ActiveChar=2 then PupCall=60: pupevent 451: End If
     If ActiveChar=3 then PupCall=70: pupevent 452: End If
     If ActiveChar=4 then PupCall=80: pupevent 453: End If
     If ActiveChar=5 then PupCall=90: pupevent 454: End If
     If ActiveChar=6 then PupCall=100: pupevent 455: End If
     If ActiveChar=7 then PupCall=110: pupevent 456: End If
     If ActiveChar=8 then PupCall=120: pupevent 457: End If
     If ActiveChar=9 then PupCall=130: pupevent 458: End If
     If ActiveChar=10 then PupCall=140: pupevent 459: End If
     If ActiveChar=11 then PupCall=150: pupevent 460: End If
     If ActiveChar=12 then PupCall=160: pupevent 461: End If
     If ActiveChar=13 then PupCall=170: pupevent 462: End If
     If ActiveChar=14 then PupCall=180: pupevent 463: End If
     If ActiveChar=15 then PupCall=190: pupevent 464: End If
     If ActiveChar=16 then PupCall=200: pupevent 465: End If
     pupevent 220+ActiveChar                                          'for testing
End Sub

Sub PlayerLineUp
       If Strip1 = False AND Playerlist=1 Then
            pupevent PupCall
            Strip1 = True
            TimingChart
       End If
       If Strip2 = False AND Playerlist=2 Then
            pupevent PupCall+1
            Strip2 = True
            TimingChart
       End If
       If Strip3 = False AND Playerlist=3 Then
            pupevent PupCall+2
            Strip3 = True
            TimingChart
       End If
       If Strip4 = False AND Playerlist=4 Then
            pupevent PupCall+3
            Strip4 = True
            TimingChart
       End If
       If Strip5 = False AND Playerlist=5 Then
            pupevent PupCall+4
            Strip5 = True
            TimingChart
       End If
       If Strip6 = False AND Playerlist=6 Then
            pupevent PupCall+5
            Strip6 = True
            TimingChart
       End If
       If Strip7 = False AND Playerlist=7 Then
            pupevent PupCall+6
            Strip7 = True
            TimingChart
       End If
       If Strip8 = False AND Playerlist=8 Then
            pupevent PupCall+7
            Strip8 = True
            TimingChart
       End If
       If Strip9 = False AND Playerlist=9 Then
            pupevent PupCall+8
            Strip9 = True
            TimingChart
       End If
       If Strip10 = False AND Playerlist=10 Then
            pupevent PupCall+9
            Strip10 = True
            TimingChart
       End If
End Sub

Dim CharNumber 

Sub TimingChart
    CharNumber=ActiveChar
    If ActiveChar=CharNumber AND PlayerList=1 Then pupevent PupCall+200: End If
    If ActiveChar=CharNumber AND PlayerList=2 Then pupevent PupCall+201: End If
    If ActiveChar=CharNumber AND PlayerList=3 Then pupevent PupCall+202: End If
    If ActiveChar=CharNumber AND PlayerList=4 Then pupevent PupCall+203: End If
    If ActiveChar=CharNumber AND PlayerList=5 Then pupevent PupCall+204: End If
    If ActiveChar=CharNumber AND PlayerList=6 Then pupevent PupCall+205: End If
    If ActiveChar=CharNumber AND PlayerList=7 Then pupevent PupCall+206: End If
    If ActiveChar=CharNumber AND PlayerList=8 Then pupevent PupCall+207: End If
    If ActiveChar=CharNumber AND PlayerList=9 Then pupevent PupCall+208: End If
    If ActiveChar=CharNumber AND PlayerList=10 Then pupevent PupCall+209: End If
End Sub


Dim EndVideoTimerCheck
EndVideoTimerCheck=0

Sub EndVideoTimer_Timer
    CoinLockOut=1
    If EndVideoTimer.Enabled = True Then
       If ActiveChar=1 then
          If EndVideoTimerCheck >= 849 Then
             Pupevent 700
             CoinLockOut=0
             EndVideoTimer.Enabled = False
             EndVideoTimerCheck=0
          End If
       End if
       If ActiveChar=2 then
          If EndVideoTimerCheck >= 807 Then
             Pupevent 700
             CoinLockOut=0
             EndVideoTimer.Enabled = False
             EndVideoTimerCheck=0
          End If
       End if
       If ActiveChar=3 then
          If EndVideoTimerCheck >= 944 Then
             Pupevent 700
             CoinLockOut=0
             EndVideoTimer.Enabled = False
             EndVideoTimerCheck=0
          End If
       End if
       If ActiveChar=4 AND Strip9=False then
          If EndVideoTimerCheck >= 828 Then
             Pupevent 700
             CoinLockOut=0
             EndVideoTimer.Enabled = False
             EndVideoTimerCheck=0
          End If
       End if
       If ActiveChar=5 then
          If EndVideoTimerCheck >= 941 Then
             Pupevent 700
             CoinLockOut=0
             EndVideoTimer.Enabled = False
             EndVideoTimerCheck=0
          End If
       End if
       If ActiveChar=6 then
          If EndVideoTimerCheck >= 927 Then
             Pupevent 700
             CoinLockOut=0
             EndVideoTimer.Enabled = False
             EndVideoTimerCheck=0
          End If
       End if
       If ActiveChar=7 then
          If EndVideoTimerCheck >= 900 Then
             Pupevent 700
             CoinLockOut=0
             EndVideoTimer.Enabled = False
             EndVideoTimerCheck=0
          End If
       End if
       If ActiveChar=8 then
          If EndVideoTimerCheck >= 930 Then
             Pupevent 700
             CoinLockOut=0
             EndVideoTimer.Enabled = False
             EndVideoTimerCheck=0
          End If
       End if
       If ActiveChar=9 then
          If EndVideoTimerCheck >= 900 Then
             Pupevent 700
             CoinLockOut=0
             EndVideoTimer.Enabled = False
             EndVideoTimerCheck=0
          End If
       End If
       If ActiveChar=10 then
          If EndVideoTimerCheck >= 872 Then
             Pupevent 700
             CoinLockOut=0
             EndVideoTimer.Enabled = False
             EndVideoTimerCheck=0
          End If
       End If
       If ActiveChar=11 then
          If EndVideoTimerCheck >= 758 Then
             Pupevent 700
             CoinLockOut=0
             EndVideoTimer.Enabled = False
             EndVideoTimerCheck=0
          End If
       End If
       If ActiveChar=12 then
          If EndVideoTimerCheck >= 721 Then
             Pupevent 700
             CoinLockOut=0
             EndVideoTimer.Enabled = False
             EndVideoTimerCheck=0
          End If
       End If
       If ActiveChar=13 then
          If EndVideoTimerCheck >= 1267 Then
             Pupevent 700
             CoinLockOut=0
             EndVideoTimer.Enabled = False
             EndVideoTimerCheck=0
          End If
       End If
       If ActiveChar=14 then
          If EndVideoTimerCheck >= 1161 Then
             Pupevent 700
             CoinLockOut=0
             EndVideoTimer.Enabled = False
             EndVideoTimerCheck=0
          End If
       End If
       If ActiveChar=15 then
          If EndVideoTimerCheck >= 858 Then
             Pupevent 700
             CoinLockOut=0
             EndVideoTimer.Enabled = False
             EndVideoTimerCheck=0
          End If
       End if
       If ActiveChar=16 then
          If EndVideoTimerCheck >= 831 Then
             Pupevent 700
             CoinLockOut=0
             EndVideoTimer.Enabled = False
             EndVideoTimerCheck=0
          End If
       End If
       If EndVideoTimerCheck >= 1400 Then
             Pupevent 700
             CoinLockOut=0
             EndVideoTimer.Enabled = False
             EndVideoTimerCheck=0
       End if
       EndVideoTimerCheck=EndVideoTimerCheck+1
       PupEvent 410
       eg=1
       state=false
      End If
End Sub

Sub UpdateApron
		if altapron=1 Then
                altapron=1
				Wapron.Image="jp-apron2"
                Wall6.Image="jp-apron2"
		        Pcover.Image="plunger2"
                PrimWalls.Image="primwalls-black"
        End if
		if altapron=0 Then
                altapron=0
			    Wapron.Image="jp-apron1"
                Wall6.Image="jp-apron1"
		        Pcover.Image="plunger1"
                PrimWalls.Image="primwalls-white"
		end if
End Sub
