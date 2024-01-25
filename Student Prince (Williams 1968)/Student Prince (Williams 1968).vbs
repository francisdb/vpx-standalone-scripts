Option Explicit  'Force explicit variable declaration.
Randomize

ShowDT = false
	


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName = "studentprince_1968"

Dim score(4)
dim truesc(4)
dim reel(4)
dim ballrelenabled
dim state
dim playno
dim credit
dim eg
dim currpl
dim plno(4)
dim play(4)
dim rst
dim ballinplay
dim match(10)
dim tilt
dim tiltsens
dim rep(4)
dim plm(4)
dim matchnumb
dim cred
dim scn
dim scn1
dim bell
dim points
dim tempscore
dim replay1
dim replay2
dim replay3
dim hisc
dim up(4)
dim rv
dim laf
dim bgl1,bgl2,bgl3,bgl4
dim ShowDT
dim i

Sub Timer1001_timer
	if B2SOn then
		if bgl1=0 Then
			Controller.B2SSetData 10,1
			bgl1=1
		Else
			Controller.B2SSetData 10,0
			bgl1=0
		end If
	end If
end Sub

Sub Timer1002_timer
	if B2SOn then
		if bgl2=0 Then
			Controller.B2SSetData 11,1
			bgl2=1
		Else
			Controller.B2SSetData 11,0
			bgl2=0
		end If
	end If
end Sub

Sub Timer003_timer
	if B2SOn then
		if bgl3=0 Then
			Controller.B2SSetData 12,1
			bgl3=1
		Else
			Controller.B2SSetData 12,0
			bgl3=0
		end If
	end If
end Sub

Sub Timer004_timer
	if B2SOn then
		if bgl4=0 Then
			Controller.B2SSetData 13,1
			bgl4=1
		Else
			Controller.B2SSetData 13,0
			bgl4=0
		end If
	end If
end Sub


sub table1_init

	dim obj
	LoadEM
	If ShowDT=false Then
		for each obj in DesktopItems
			obj.visible=False
		Next
	end If
    bgl1=0
	bgl2=0
	bgl3=0
	bgl4=0

    play(1)=up1
    play(2)=up2
    play(3)=up3
    play(4)=up4
    match(0)=m0
    match(1)=m1
    match(2)=m2
    match(3)=m3
    match(4)=m4
    match(5)=m5
    match(6)=m6
    match(7)=m7
    match(8)=m8
    match(9)=m9
    up(1)=up1
    up(2)=up2
    up(3)=up3
    up(4)=up4
    reel(1)=reel1
    reel(2)=reel2
    reel(3)=reel3
    reel(4)=reel4
    replay1=4200
    replay2=6500
    replay3=9800
    'loadhs
    if rv="" then rv=1
    creel.setvalue(rv-1)
    if hisc="" then hisc=1000
    hisctxt.text=hisc
    if credit="" then credit=0 
    credittxt.text=credit
    if matchnumb="" then matchnumb=int(rnd(1)*9)
		select case(matchnumb)
			case 0:
				m0.text="0"
			case 1:
				m1.text="1"
			case 2:
				m2.text="2"
			case 3:
				m3.text="3"
			case 4:
				m4.text="4"
			case 5:
				m5.text="5"
			case 6:
				m6.text="6"
			case 7:
				m7.text="7"
			case 8:
				m8.text="8"
			case 9:
				m9.text="9"
		end select
    for i=1 to 4
		currpl=i
		reel(i).setvalue(score(i))
	next
    currpl=0
    cred=0
    credtimer.enabled=true

	If B2SOn Then
		timer001.enabled=True
		timer002.enabled=True
		timer003.enabled=True
		timer004.enabled=True
		Controller.B2SSetScoreRolloverPlayer1 0
		Controller.B2SSetScoreRolloverPlayer2 0
		Controller.B2SSetScoreRolloverPlayer3 0
		Controller.B2SSetScoreRolloverPlayer4 0
		Controller.B2SSetShootAgain 0
		Controller.B2SSetTilt 0
		Controller.B2SSetCredits credit
		Controller.B2SSetGameOver 0
		Controller.B2SSetData 81,0
		Controller.B2SSetData 82,0
		Controller.B2SSetData 83,0
		Controller.B2SSetData 84,0
	End If
	for i=1 to 4
		If B2SOn Then 
			Controller.B2SSetScorePlayer i, 0
		End If
	next
	MusicOn
end sub    

Sub MusicOn
    ' Dim FileSystemObject, folder, r, ct, file, musicPath, myMusicFolder
	' myMusicFolder = "StudentPrince" ' the directory name where your mp3 files are storied, must be a subfolder of Visual Pinball\music
	' Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
    ' musicPath = FileSystemObject.GetAbsolutePathName(".") ' get path to Visual Pinball\table
    ' musicPath = Left(musicPath, Len(musicPath) - 6) + "music\" 'get path to Visual Pinball\music
	' if (FileSystemObject.FolderExists(musicPath + myMusicFolder)) then 
	' 	Set folder = FileSystemObject.GetFolder(musicPath + myMusicFolder)
	' 	Randomize
	' 	r = INT(folder.Files.Count * Rnd + 1)
	' 	ct=1
	' 	For Each file in folder.Files 'get every file in myMusicFolder, for each one countt it and see if the count matches the random number
	' 		if ct = r Then  ' random file found
	' 			if (LCase(Right(file,4))) = ".mp3" Then ' can only play mp3 files
	' 				PlayMusic Mid(file, Len(musicPath) + 1, 1000) ' PlayMusic defaults to Visual Pinball\music\s, need to get jmyMusicFolder\song name
	' 			End If 
	' 		End If
	' 	ct = ct + 1
	' 	Next
	' end if
 End Sub

 Sub Table1_MusicDone()
    MusicOn
 End Sub


Sub Table1_KeyDown(ByVal keycode)
    If keycode = RightMagnaSave Then MusicOn
    If keycode = LeftMagnaSave Then EndMusic
    dim bip    
    If keycode = PlungerKey Then
		Plunger.PullBack
	End If

    if keycode = 6 or keycode = 5 then
		playsound "coin3"
		coindelay.enabled=true
	end if

	if keycode = 2 and credit>0 and state=false and playno=0 then
		credit=credit-1
		credittxt.text=credit
		eg=0
		playno=1
		playno1.state=1
		currpl=1
		play(currpl).state=1
		playsound "click"
		playsound "initialize"
		rst=0
		bip=1 
		resettimer.enabled=true
		If B2SOn Then
			Controller.B2SSetCredits Credit
			Controller.B2SSetCanPlay playno
		end if

		resettimer.enabled=true
	end if
    
	if keycode = 2 and credit>0 and state=true and playno>0 and playno<4 and bip<2 then
		credit=credit-1
        credittxt.text=credit
		playno=playno+1
		if playno=2 then playno2.state=1
		if playno=3 then playno3.state=1
		if playno=4 then playno4.state=1
		If B2SOn Then
			Controller.B2SSetCredits Credit
			Controller.B2SSetCanPlay playno
		end if

		playsound "click"
    end if

    if state=true and tilt=false then
		If keycode = LeftFlipperKey Then
			LeftFlipper.RotateToEnd
			LeftFlipper1.RotateToEnd
			PlaySound "FlipperUp"
			playsound "buzz"
		End If
    
		If keycode = RightFlipperKey Then
			RightFlipper.RotateToEnd
			RightFlipper1.RotateToEnd
			PlaySound "FlipperUp"
			playsound "buzz"
		End If
   
		If keycode = LeftTiltKey Then
			Nudge 90, 1
			checktilt
		End If
    
		If keycode = RightTiltKey Then
			Nudge 270, 1
			checktilt
		End If
    
		If keycode = CenterTiltKey Then
			Nudge 0, 1
			checktilt
		End If
	end if
    
End Sub


Sub Table1_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.Fire
		playsound "plunger"
	End If
    
	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToStart
		LeftFlipper1.RotateToStart
		stopsound "buzz"
		if state=true and tilt=false then PlaySound "FlipperDown"
	End If
    
	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
		RightFlipper1.RotateToStart
		stopsound "buzz"
        if state=true and tilt=false then PlaySound "FlipperDown"
	End If

End Sub

Sub flippertimer_Timer()
	LFlip.RotY = LeftFlipper.CurrentAngle+90
	RFlip.RotY = RightFlipper.CurrentAngle+90
	LFlip1.RotY = LeftFlipper.CurrentAngle+70
	RFlip1.RotY = RightFlipper.CurrentAngle+110
	
End Sub

sub coindelay_timer
	playsound "click" 
	credit=credit+1
	credittxt.text=credit 
	If B2SOn Then Controller.B2SSetCredits credit
    coindelay.enabled=false
end sub

sub resettimer_timer
    rst=rst+1
    reel1.resettozero
    reel2.resettozero
    reel3.resettozero
    reel4.resettozero
	If B2SOn Then
		Controller.B2SSetScorePlayer1 0
		Controller.B2SSetScorePlayer2 0
		Controller.B2SSetScorePlayer3 0
		Controller.B2SSetScorePlayer4 0
		Controller.B2SSetScorePlayer5 0
	end if
    if rst=14 then
    playsound "kickerkick"
    end if
    if rst=18 then
    newgame
    resettimer.enabled=false
    end if
end sub

sub addscore(points)
    if tilt=false then
		bell=0
		if points=10 or points=1 then 
			if points=1 then
				creel.addvalue(1)
				rv=rv+1 
				if rv>5 then rv=1
				if laf=1 then
					if (loutl.state)=1 then
						loutl.state=0
						routl.state=1
					else
						loutl.state=1
						routl.state=0
					end if
				end if
			end if
    end if
    if points = 1 or points = 10 or points = 100 then scn=1
    if points = 100 then 
    reel(currpl).addvalue(100)
    bell=100
    end if
    if points = 10 then 
    reel(currpl).addvalue(10)
    bell=10
    end if
    if points = 1 then 
    reel(currpl).addvalue(1)
    bell=1
    end if
    if points = 300 then 
    reel(currpl).addvalue(300)
    scn=3
    bell=100
    end if
    if points = 200 then
    reel(currpl).addvalue(200) 
    scn=2
    bell=100
    end if
    if points = 500 then
    reel(currpl).addvalue(500)
    scn=5
    bell=100
    end if
    if points = 400 then
    reel(currpl).addvalue(400)
    scn=4
    bell=100
    end if
    if points = 30 then 
    reel(currpl).addvalue(30)
    scn=3
    bell=10
    end if
    if points = 20 then
    reel(currpl).addvalue(20) 
    scn=2
    bell=10
    end if
    if points = 50 then
    reel(currpl).addvalue(50)
    scn=5
    bell=10
    end if
    if points = 40 then
    reel(currpl).addvalue(40)
    scn=4
    bell=10
    end if
    scn1=0
    scntimer.enabled=true
    score(currpl)=score(currpl)+points
    truesc(currpl)=truesc(currpl)+points
    if score(currpl)>9999 then
    score(currpl)=score(currpl)-10000
    rep(currpl)=0
    end if
	If B2SOn Then
		Controller.B2SSetScore currpl,score(currpl)
	end if

    if score(currpl)=>replay1 and rep(currpl)=0 then
    credit=credit+1
    playsound "knocke"
    credittxt.text=credit
    rep(currpl)=1
    playsound "click"
    end if
    if score(currpl)=>replay2 and rep(currpl)=1 then
    credit=credit+1
    playsound "knocke"
    credittxt.text=credit
    rep(currpl)=2
    playsound "click"
    end if
    if score(currpl)=>replay3 and rep(currpl)=2 then
    credit=credit+1
    playsound "knocke"
    credittxt.text=credit
    rep(currpl)=3
    playsound "click"
    end if
    end if
end sub 

sub scntimer_timer
    scn1=scn1 + 1
    if bell=1 then playsound "bell1"
    if bell=10 then playsound "bell10"
    if bell=100 then playsound "bell100"
    if scn1=scn then scntimer.enabled=false
end sub

sub newgame
	credtimer.enabled=false
	state=true
	'bipreel.setvalue(1)
	rep(1)=0
	rep(2)=0
	rep(3)=0
	eg=0
	laf=0
	score(1)=0
	score(2)=0
	score(3)=0
	score(4)=0
	truesc(1)=0
	truesc(2)=0
	truesc(3)=0
	truesc(4)=0
   	bumper1.force=11
	bumper2.force=11
	bumper3.force=11
	bumper4.force=11
	bumperlight1.state=0
	bumperlight2.state=0
	bumperlight3.state=0
	bumperlight4.state=0
	bltl2.state=1
	brtl2.state=1
	bltl1.state=0
	brtl1.state=0
	tl1.state=1
	tl2.state=1
	tl3.state=1
	tl4.state=1
	tl5.state=1
	bbl1.state=0
	bbl2.state=0
	bbl3.state=0
	bbl4.state=0
	bbl5.state=0
	loutl.state=0
	routl.state=0
	gatel.state=0
	gc.isdropped=false  
	bip5.text=" "
	bip1.text="1"
	for i=0 to 9
	match(i).text=" "
	next
	tilttext.text=" " 
	gamov.text=" " 
	tilt=false
	tiltsens=0
	ballinplay=1
	nb.CreateBall
	nb.kick 135,6
	If B2SOn then 
		Controller.B2SSetBallInPlay 1
		Controller.B2SSetPlayerUp 1
		Controller.B2SSetData (81),1
	end if

end sub

Sub Drain_Hit()
	playsound "drainshorter"
	Drain.DestroyBall
	flipopen 
	nextball
End Sub

sub nextball	
	if tilt=true then
		tilt=false
		tilttext.text=" "
		bumper1.force=11
		bumper2.force=11
		bumper3.force=11
		bumper4.force=11
		bumperlight1.state=0
		bumperlight2.state=0
		bumperlight3.state=0
		bumperlight4.state=0
		tiltseq.stopplay
		lsling.isdropped=false
		rsling.isdropped=false
	end if
	if (shootagain.state)=1 then
		playsound "kickerkick"
		newball
		ballreltimer.enabled=true
	else
		currpl=currpl+1
		if B2SOn Then
			Controller.B2SSetData 81,0
			Controller.B2SSetData 82,0
			Controller.B2SSetData 83,0
			Controller.B2SSetData 84,0
			Controller.B2SSetTilt 0
			if currpl > playno then 
				Controller.B2SSetPlayerUp 1
				Controller.B2SSetData (81),1
			else
				Controller.B2SSetPlayerUp currpl
				Controller.B2SSetData (80+currpl),1

			end if
		end if

	end if
	if currpl>playno then
		ballinplay=ballinplay+1
		if ballinplay>5 then
			If B2SOn then Controller.B2SSetBallInPlay 0
			playsound "motorleer"
			eg=1
			ballreltimer.enabled=true
		else
			if state=true and tilt=false then
				play(currpl-1).state=0
				currpl=1
				play(currpl).state=1
				newball
				playsound "kickerkick"
				ballreltimer.enabled=true
			end if
		If B2SOn then Controller.B2SSetBallInPlay ballinplay
		select case (ballinplay)
			case 1:
				bip1.text="1"
			case 2: 
				bip1.text=" "
				bip2.text="2"
			case 3:
				bip2.text=" "
				bip3.text="3"
			case 4:
				bip3.text=" "
				bip4.text="4"
			case 5:
				bip4.text=" "
				bip5.text="5"
			end select
		end if
	end if
	if currpl>1 and currpl<(playno+1) then
		if state=true and tilt=false then
			play(currpl-1).state=0
			play(currpl).state=1
			newball
			playsound "kickerkick"
			ballreltimer.enabled=true  
		end if
	end if
end Sub

sub ballreltimer_timer
	if eg=1 then
		matchnum
		bip3.text=" "
		bip5.text=" "
		'bipreel.setvalue(0)
		state=false
		for i=1 to 4
			if truesc(i)>hisc then 
				hisc=truesc(i)
				hisctxt.text=hisc
			end if
		next
		play(currpl-1).state=0
		playno=0
		gamov.text="GAME OVER"
		If B2SOn Then
			Controller.B2SSetTilt 0
			Controller.B2SSetGameOver 1
			if matchnumb=0 then
				Controller.B2SSetMatch 10
			else
				Controller.B2SSetMatch matchnumb
			end if
		end if
		'savehs
		cred=0
		credtimer.enabled=true
		ballreltimer.enabled=false
	else
		nb.CreateBall
		nb.kick 135,6
		ballreltimer.enabled=false
    end if
end sub

sub newball
    bltl1.state=0
    brtl1.state=0
    bltl2.state=1
    brtl2.state=1
    tl1.state=1
    tl2.state=1
    tl3.state=1
    tl4.state=1
    tl5.state=1
    bbl1.state=0
    bbl2.state=0
    bbl3.state=0
    bbl4.state=0
    bbl5.state=0
    bumperlight1.state=0
    bumperlight2.state=0
    bumperlight3.state=0
    bumperlight4.state=0
    gatel.state=0
    if (gc.isdropped)=true then playsound "gateopen"
    gc.isdropped=false
    stopsound "buzz"
    laf=0
    loutl.state=0
    routl.state=0
end sub

sub matchnum
    select case(matchnumb)
		case 0:
			m0.text="0"
		case 1:
			m1.text="1"
		case 2:
			m2.text="2"
		case 3:
			m3.text="3"
		case 4:
			m4.text="4"
		case 5:
			m5.text="5"
		case 6:
			m6.text="6"
		case 7:
			m7.text="7"
		case 8:
			m8.text="8"
		case 9:
			m9.text="9"
	end select
    for i=1 to playno
		if CStr(matchnumb)=Right(CStr(score(i)),1) then
			credit=credit+1
            If B2SOn Then Controller.B2SSetCredits Credit
			playsound "knocke"
			credittxt.text= credit
			playsound "click"
		end if
    next 
end sub

Sub CheckTilt
	If Tilttimer.Enabled = True Then
	TiltSens = TiltSens + 1
	if TiltSens = 2 Then
	Tilt = True
	tilttext.text="TILT"
		If B2SOn Then
			Controller.B2SSetTilt 1
		end if

	tiltsens = 0
	playsound "tilt"
	turnoff 
	End If
	Else
	TiltSens = 0
	Tilttimer.Enabled = True
	End If
End Sub

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

sub turnoff
    flipopen
    bumper1.force=0
    bumper2.force=0
    bumper3.force=0
    bumper4.force=0
    bumperlight1.state=0
    bumperlight2.state=0
    bumperlight3.state=0
    bumperlight4.state=0
    tiltseq.play seqalloff
    'logo1.setvalue(0)
    'logo2.setvalue(0)
    'bipreel.setvalue(0)
    lsling.isdropped=true
    rsling.isdropped=true
end sub    

sub flipclose
    if tilt=false then
    if (rightflipper.visible)=true then playsound "zclose"
    rflip.visible=false
    rflip1.visible=true 
    lflip.visible=false
    lflip1.visible=true
    leftflipper.enabled=false
    leftflipper1.enabled=true
    rightflipper.enabled=false
    rightflipper1.enabled=true
    fp1.isdropped=false
    fp2.isdropped=false
    end if
end sub    
                
sub flipopen
    if (rightflipper.visible)=false then playsound "zopen"
    if (rflip.visible)=false then playsound "zopen"
    rflip.visible=true
    rflip1.visible=false 
    lflip.visible=true
    lflip1.visible=false
    leftflipper.enabled=true
    leftflipper1.enabled=false
    rightflipper.enabled=true
    rightflipper1.enabled=false
    fp1.isdropped=true
    fp2.isdropped=true
    ctl.state=0
end sub

sub blt_hit
    addscore 10
    if(bltl1.state)=1 then
    flipopen
    bltl1.state=0
    bltl2.state=1
    brtl1.state=0
    brtl2.state=1
    ctl.state=0
    else
    flipclose
    bltl1.state=1
    bltl2.state=0
    brtl1.state=1
    brtl2.state=0
    ctl.state=1
    end if
end sub     

sub brt_hit
    addscore 10
    if(brtl1.state)=1 then
    flipopen
    bltl1.state=0
    bltl2.state=1
    brtl1.state=0
    brtl2.state=1
    ctl.state=0
    else
    flipclose
    bltl1.state=1
    bltl2.state=0
    brtl1.state=1
    brtl2.state=0
    ctl.state=1
    end if
end sub

Sub LSling_Slingshot()
	if tilt=false then PlaySound "Bumper"
    addscore 1 
End Sub

Sub RSling_Slingshot()
	if tilt=False then PlaySound "Bumper"
	addscore 1
End Sub


Sub lsling2_Slingshot()
	if tilt=false then addscore 1 
End Sub

Sub rsling2_Slingshot()
	if tilt=false then addscore 1
End Sub

Sub lsling3_Slingshot()
	if tilt=false then addscore 1 
End Sub

Sub rsling3_Slingshot()
	if tilt=false then addscore 1
End Sub

Sub lsling4_Slingshot()
	if tilt=false then addscore 1 
End Sub

Sub rsling4_Slingshot()
	if tilt=false then addscore 1
End Sub

Sub lsling5_Slingshot()
	if tilt=false then addscore 1 
End Sub

Sub rsling5_Slingshot()
	if tilt=false then addscore 1
End Sub

sub rout_hit
    if (routl.state)=1 then
    addscore 300
    else
    addscore 100
    end if
end sub

sub lout_hit
    if (loutl.state)=1 then
    addscore 300
    else
    addscore 100
    end if
end sub        

sub bb1_hit
    if (bbl1.state)=1 then
    addscore 10
    end if
end sub    
    
sub bb2_hit
    if (bbl2.state)=1 then
    addscore 10
    end if
end sub

sub bb3_hit
    if (bbl3.state)=1 then
    addscore 10
    end if
end sub

sub bb4_hit
    if (bbl4.state)=1 then
    addscore 10
    end if
end sub

sub bb5_hit
    if (bbl5.state)=1 then
    addscore 10
    end if
end sub

sub upl2_hit
    if (tl1.state)=1 then
    tl1.state=0
    bbl1.state=1
    bumperlight1.state=1
    checkaward
    end if
    addscore 100
end sub

sub upl1_hit
    if (tl2.state)=1 then
    tl2.state=0
    bbl2.state=1
    bumperlight2.state=1
    checkaward
    end if
    addscore 100
end sub    

sub upc_hit
    if (tl3.state)=1 then
    tl3.state=0
    bbl3.state=1
    checkaward
    end if
    addscore 100
    if (gc.isdropped)=false then 
    playsound "gateopen"
    playsound "buzz"
    end if
    gc.isdropped=true
    gatel.state=1    
end sub

sub upr1_hit
    if (tl4.state)=1 then
    tl4.state=0
    bbl4.state=1
    bumperlight3.state=1
    checkaward
    end if
    addscore 100
end sub

sub upr2_hit
    if (tl5.state)=1 then
    tl5.state=0
    bbl5.state=1
    bumperlight4.state=1
    checkaward
    end if
    addscore 100
end sub

sub gtrig_hit
    addscore 300
    gatel.state=0
    gc.isdropped=false
    playsound "gateopen"
    stopsound "buzz"
end sub    

sub tt_hit
    if (gc.isdropped)=false then 
    playsound "gateopen"
    playsound "buzz"
    end if
    gc.isdropped=true
    gatel.state=1
    addscore 100
end sub

sub bumper1_hit
    if tilt=false then playsound "jet1"
    if bumperlight1.state=1 then addscore 10 else addscore 1
end sub

sub bumper2_hit
    if tilt=false then playsound "jet1"
    if bumperlight2.state=1 then addscore 10 else addscore 1
end sub

sub bumper3_hit
    if tilt=false then playsound "jet1"
    if bumperlight3.state=1 then addscore 10 else addscore 1
end sub

sub bumper4_hit
    if tilt=false then playsound "jet1"
    if bumperlight4.state=1 then addscore 10 else addscore 1
end sub

sub checkaward
    if (bbl1.state)=1 and (bbl2.state)=1 and (bbl3.state)=1 and (bbl4.state)=1 and (bbl5.state)=1 then
    shootagain.state=1
    spsa.text="SAME PLAYER SHOOTS AGAIN"
    laf=1
    loutl.state=1
    end if
end sub  

sub ct_hit
    if tilt=false then reelscore
end sub

sub topb_hit
    if tilt=false then reelscore
end sub

sub reelscore        
    select case (rv)

    case 1:
    bbl1.state=1
    tl1.state=0
    bumperlight1.state=1
    if (ctl.state)=1 then 
    addscore (rv*100)
    else
    addscore (rv*10)
    end if

    case 2:
    bbl2.state=1
    tl2.state=0
    bumperlight2.state=1    
    if (ctl.state)=1 then 
    addscore (rv*100)
    else
    addscore (rv*10)
    end if

    case 3:
    bbl3.state=1
    tl3.state=0
    if (ctl.state)=1 then 
    addscore (rv*100)
    else
    addscore (rv*10)
    end if

    case 4:
    bbl4.state=1
    tl4.state=0
    bumperlight3.state=1
    if (ctl.state)=1 then 
    addscore (rv*100)
    else
    addscore (rv*10)
    end if

    case 5:
    bbl5.state=1
    tl5.state=0
    bumperlight4.state=1
    if (ctl.state)=1 then 
    addscore (rv*100)
    else
    addscore (rv*10)
    end if

    end select
    checkaward
end sub     


sub savehs
	' Based on Black's Highscore routines
	Dim FileObj
	Dim ScoreFile
	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then 
		Exit Sub
	End if
	Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "Studentprince.txt",True)
		ScoreFile.WriteLine score(1)
		ScoreFile.WriteLine score(2)
		ScoreFile.WriteLine score(3)
		ScoreFile.WriteLine score(4)
		scorefile.writeline credit
		scorefile.writeline matchnumb
        scorefile.writeline hisc
        scorefile.writeline rv
		ScoreFile.Close
	Set ScoreFile=Nothing
	Set FileObj=Nothing
end sub

sub loadhs
    ' Based on Black's Highscore routines
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
	dim TextStr
	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then 
	Exit Sub
	End if
	If Not FileObj.FileExists(UserDirectory & "Studentprince.txt") then
	Exit Sub
	End if
	Set ScoreFile=FileObj.GetFile(UserDirectory & "Studentprince.txt")
	Set TextStr=ScoreFile.OpenAsTextStream(1,0)
		If (TextStr.AtEndOfStream=True) then
		Exit Sub
		End if
		temp1=TextStr.ReadLine
		temp2=textstr.readline
		temp3=textstr.readline
		temp4=Textstr.ReadLine
		temp5=Textstr.ReadLine
		temp6=Textstr.ReadLine
		temp7=Textstr.ReadLine 
		temp8=Textstr.ReadLine
		TextStr.Close
	    score(1) = CDbl(temp1)
	    score(2)= CDbl(temp2)
	    score(3) = CDbl(temp3)
	    score(4) = CDbl(temp4)
	    credit= CDbl(temp5)
	    matchnumb= CDbl(temp6)
	    hisc=cdbl(temp7)
	    rv=cdbl(temp8)
	    Set ScoreFile=Nothing
	    Set FileObj=Nothing
end sub

sub ballhome_hit
    bb.isdropped=true
    ballrelenabled=1
end sub

sub ballout_hit
    bb.isdropped=false
    shootagain.state=0
    spsa.text=" "
end sub
         
sub ballrel_hit
    if ballrelenabled=1 then 
    playsound "launchball"
    ballrelenabled=0
    end if
end sub

