' ******************************************************************
'       VPX - version by Klodo81 2024, Ben Hur version 1.1
'
'                 VISUAL PINBALL X EM Script Based on
'               JPSalas Basic EM script up to 4 players
'		                 JPSalas Physics v3.0
'
' Sound for Score,Coin,Startgame From Cleopatra table From Goldchicco
' 	
'              Ben Hur / IPD No.2855 / 1977 / 4 Players   
'
' ******************************************************************
'
'V1.1 mod table
'- Mod lighting cycle of arrows lights of bonus
'V1.2. VRRoom Added by Ext2k
'
' ******************************************************************

Option Explicit
Randomize

' DOF config
'
' Flippers L/R/BL - 101/102/103
' Bumpers - 105
' Targets - 111
' Targets Round - 203,208,209 
' Triggers - 201,202,203
' Knocker - 110
' Drain - 250 / 104

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

'****************** Options you can change *******************

Const BallsPerGame = 3     ' to play with 3 or 5 balls
Const AttractMode = 1      ' 0 = no attract mode
Const FreePlay = False     ' Free play or coins

'*************************************************************

' Constants
Const TableName = "BenHur_77VPX"   ' file name to save highscores and other variables
Const cGameName = "BenHur1977"   ' B2S name
Const MaxPlayers = 4         ' 1 to 4 can play
Const Special1 = 190000      ' 3 Balls award credit
Const Special2 = 280000      ' 3 Balls award credit
Const Special3 = 370000      ' 3 Balls award credit
Const Special4 = 320000      ' 5 Balls award credit
Const Special5 = 480000      ' 5 Balls award credit
Const Special6 = 620000      ' 5 Balls award credit

' Global variables
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim Bonus
Dim BallsRemaining(4)
Dim BonusMultiplier
Dim ExtraBallsAwards(4)
Dim Special1Awarded(4)
Dim Special2Awarded(4)
Dim Special3Awarded(4)
Dim Special4Awarded(4)
Dim Special5Awarded(4)
Dim Special6Awarded(4)
Dim SpecialHSAwarded
Dim Score(4)
Dim HighScore
Dim Match
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim Add10
Dim Add100
Dim Add1000
Dim Add10000
Dim x

Dim SpecialAwarded
Dim CenterBonusCounter

' Control variables
Dim BallsOnPlayfield

' Boolean variables
Dim bAttractMode
Dim bFreePlay
Dim bGameInPlay
Dim bOnTheFirstBall
Dim bExtraBallWonThisBall
Dim bJustStarted
Dim bBallInPlungerLane
Dim bBallSaverActive
Dim bAttractHSScore
Dim bSpecial

' core.vbs variables

' *********************************************************************
'                Common rutines to all the tables
' *********************************************************************

Sub Table1_Init()
    Dim x

    ' Init some objects, like walls, targets
    VPObjects_Init
    LoadEM

    ' load highscores and credits
    Loadhs
    If B2SOn then
        vpmtimer.addtimer 4000,"ResetScoresInit '"
    End If
    UpdateCredits

    ' init all the global variables
    bFreePlay = Freeplay
    bAttractMode = False
    bOnTheFirstBall = False
    bGameInPlay = False
    bBallInPlungerLane = False
    BallsOnPlayfield = 0
    Tilt = 0
    TiltSensitivity = 6
    Tilted = False
    Match = 0
    bJustStarted = True
    Add10 = 0
    Add100 = 0
    Add1000 = 0
	Add10000 = 0
	bSpecial = False
	CenterBonusCounter = 0

	' select Card Instruction on Apron
	If BallsPerGame = 3 Then
		Flasher3ballsL.visible = 1
		Flasher5ballsL.visible = 0
		Flasher3ballsR.visible = 1
		Flasher5ballsR.visible = 0
	Else 'BallsPerGame = 5
		Flasher3ballsL.visible = 0
		Flasher5ballsL.visible = 1
		Flasher3ballsR.visible = 0
		Flasher5ballsR.visible = 1
	End If

    ' setup table in game over mode
    EndOfGame

    'turn on GI lights
    vpmtimer.addtimer 1000, "GiOn '"

    ' Remove desktop items in FS mode
    If Table1.ShowDT then
        For each x in aReels
            x.Visible = 1
        Next
    Else
        For each x in aReels
            x.Visible = 0
        Next
    End If

    ' LUT - Darkness control
    LoadLUT
End Sub

'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)

    If EnteringInitials then
        CollectInitials(keycode)
        Exit Sub
    End If
	If keycode = LeftFlipperKey Then
		VR_CabFlipperLeft.X = VR_CabFlipperLeft.X +10
	End If
	If keycode = StartGameKey then 		
		VR_Cab_StartButton.y = VR_Cab_StartButton.y -5
	End If

    If keycode = LeftMagnaSave Then bLutActive = True:SetLUTLine "Color LUT image " & table1.ColorGradeImage
    If keycode = RightFlipperKey Then
		VR_CabFlipperRight.X = VR_CabFlipperRight.X - 10
        If bLutActive Then NextLUT:End If
    End If

    ' add coins
    If Keycode = AddCreditKey Then
		VR_Cab_StartButton.y = VR_Cab_StartButton.y -5
        If(Tilted = False) Then
            AddCredits 1
            PlaySound "coin"
        End If
    End If

    ' plunger
    If keycode = PlungerKey Then
        Plunger.Pullback
        PlaySoundAt "fx_plungerpull", plunger
    End If

    ' tilt keys
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound "fx_nudge", 0, 1, -0.1, 0.25:CheckTilt
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound "fx_nudge", 0, 1, 0.1, 0.25:CheckTilt
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound "fx_nudge", 0, 1, 1, 0.25:CheckTilt

    ' keys during game

    If bGameInPlay AND NOT Tilted Then

        If keycode = LeftFlipperKey Then SolLFlipper 1
        If keycode = RightFlipperKey Then SolRFlipper 1

        If keycode = StartGameKey Then
            If((PlayersPlayingGame <MaxPlayers) AND(bOnTheFirstBall = True) ) Then

                If(bFreePlay = True) Then
                    PlayersPlayingGame = PlayersPlayingGame + 1    
                    UpdateBallInPlay
					UpdatePlayers
                    PlaySound "startgame"
                Else
                    If(Credits> 0) then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        Credits = Credits - 1
                        UpdateCredits
                        UpdateBallInPlay
						UpdatePlayers
						PlaySound "startgame"
                    End If
                End If
            End If
        End If
        Else ' If (GameInPlay)

            If keycode = StartGameKey Then
                If(bFreePlay = True) Then
                    If(BallsOnPlayfield = 0) Then
                        ResetScores
                        ResetForNewGame()
						PlaySound "startgame"
                    End If
                Else
                    If(Credits> 0) Then
                        If(BallsOnPlayfield = 0) Then
                            Credits = Credits - 1
                            UpdateCredits
                            ResetScores
                            ResetForNewGame()
						PlaySound "startgame"
                        End If
                    End If
                End If
            End If
    End If ' If (GameInPlay)
End Sub

Sub Table1_KeyUp(ByVal keycode)

    If EnteringInitials then
        Exit Sub
    End If

    If keycode = LeftMagnaSave Then bLutActive = False:HideLUT

	If keycode = LeftFlipperKey Then
		VR_CabFlipperLeft.X = VR_CabFlipperLeft.X -10
	End If
    If keycode = RightFlipperKey Then
		VR_CabFlipperRight.X = VR_CabFlipperRight.X +10
	End If
	If Keycode = StartGameKey Then	
		VR_Cab_StartButton.y = VR_Cab_StartButton.y +5
	End If
	If keycode = AddCreditKey or keycode = 4 then
		VR_Cab_StartButton.y = VR_Cab_StartButton.y +5
	End If

    If bGameInPlay AND NOT Tilted Then

        If keycode = LeftFlipperKey Then SolLFlipper 0
        If keycode = RightFlipperKey Then SolRFlipper 0
    End If

    If keycode = PlungerKey Then
        Plunger.Fire
        If bBallInPlungerLane Then
            PlaySoundAt "fx_plunger", plunger
        Else
            PlaySoundAt "fx_plunger_empty", plunger
        End If
    End If
End Sub

'******************
' Table stop/pause
'******************

Sub table1_Paused
End Sub

Sub table1_unPaused
End Sub

Sub table1_Exit
    Savehs
	'Controller.Stop
End Sub

'************************************
'       LUT - Darkness control
' 10 normal level & 10 warmer levels
'************************************

Dim bLutActive, LUTImage

Sub LoadLUT
    Dim x
    bLutActive = False
    x = LoadValue(cGameName, "LUTImage")
    If(x <> "") Then LUTImage = x Else LUTImage = 0
    UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT:LUTImage = (LUTImage + 1) MOD 22:UpdateLUT:SaveLUT:SetLUTLine "Color LUT image " & table1.ColorGradeImage:End Sub

Sub UpdateLUT
    Select Case LutImage
        Case 0:table1.ColorGradeImage = "LUT0"
        Case 1:table1.ColorGradeImage = "LUT1"
        Case 2:table1.ColorGradeImage = "LUT2"
        Case 3:table1.ColorGradeImage = "LUT3"
        Case 4:table1.ColorGradeImage = "LUT4"
        Case 5:table1.ColorGradeImage = "LUT5"
        Case 6:table1.ColorGradeImage = "LUT6"
        Case 7:table1.ColorGradeImage = "LUT7"
        Case 8:table1.ColorGradeImage = "LUT8"
        Case 9:table1.ColorGradeImage = "LUT9"
        Case 10:table1.ColorGradeImage = "LUT10"
        Case 11:table1.ColorGradeImage = "LUT Warm 0"
        Case 12:table1.ColorGradeImage = "LUT Warm 1"
        Case 13:table1.ColorGradeImage = "LUT Warm 2"
        Case 14:table1.ColorGradeImage = "LUT Warm 3"
        Case 15:table1.ColorGradeImage = "LUT Warm 4"
        Case 16:table1.ColorGradeImage = "LUT Warm 5"
        Case 17:table1.ColorGradeImage = "LUT Warm 6"
        Case 18:table1.ColorGradeImage = "LUT Warm 7"
        Case 19:table1.ColorGradeImage = "LUT Warm 8"
        Case 20:table1.ColorGradeImage = "LUT Warm 9"
        Case 21:table1.ColorGradeImage = "LUT Warm 10"
    End Select
End Sub

' New LUT postit
Sub SetLUTLine(String)
    Dim Index
    Dim xFor
    Index = 1
    LUBack.imagea = "PostItNote"
    String = CL2(String)
    For xFor = 1 to 40
        Eval("LU" &xFor).imageA = GetHSChar(String, Index)
        Index = Index + 1
    Next
End Sub

Sub HideLUT
    SetLUTLine ""
    LUBack.imagea = "PostitBL"
End Sub

Function CL2(NumString) 'center line
    Dim Temp, TempStr
    If Len(NumString)> 40 Then NumString = Left(NumString, 40)
    Temp = (40 - Len(NumString) ) \ 2
    TempStr = Space(Temp) & NumString & Space(Temp)
    CL2 = TempStr
End Function

'**********************************************
'    Flipper adjustments - enable tricks
'             by JLouLouLou
'**********************************************

Dim FlipperPower
Dim FlipperElasticity
Dim SOSTorque, SOSAngle
Dim FullStrokeEOS_Torque, LiveStrokeEOS_Torque
Dim LeftFlipperOn
Dim RightFlipperOn

Dim LLiveCatchTimer
Dim RLiveCatchTimer
Dim LiveCatchSensivity

FlipperPower = 5000
FlipperElasticity = 0.85
FullStrokeEOS_Torque = 0.3 	' EOS Torque when flipper hold up ( EOS Coil is fully charged. Ampere increase due to flipper can't move or when it pushed back when "On". EOS Coil have more power )
LiveStrokeEOS_Torque = 0.2	' EOS Torque when flipper rotate to end ( When flipper move, EOS coil have less Ampere due to flipper can freely move. EOS Coil have less power )

LeftFlipper.EOSTorqueAngle = 10
RightFlipper.EOSTorqueAngle = 10

SOSTorque = 0.1
SOSAngle = 6

LiveCatchSensivity = 10

LLiveCatchTimer = 0
RLiveCatchTimer = 0

LeftFlipper.TimerInterval = 1
LeftFlipper.TimerEnabled = 1

Sub LeftFlipper_Timer 'flipper's tricks timer
'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If LeftFlipper.CurrentAngle >= LeftFlipper.StartAngle - SOSAngle Then LeftFlipper.Strength = FlipperPower * SOSTorque else LeftFlipper.Strength = FlipperPower : End If
 
'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
	If LeftFlipperOn = 1 Then
		If LeftFlipper.CurrentAngle = LeftFlipper.EndAngle then
			LeftFlipper.EOSTorque = FullStrokeEOS_Torque
			LLiveCatchTimer = LLiveCatchTimer + 1
			If LLiveCatchTimer < LiveCatchSensivity Then
				LeftFlipper.Elasticity = 0
			Else
				LeftFlipper.Elasticity = FlipperElasticity
				LLiveCatchTimer = LiveCatchSensivity
			End If
		End If
	Else
		LeftFlipper.Elasticity = FlipperElasticity
		LeftFlipper.EOSTorque = LiveStrokeEOS_Torque
		LLiveCatchTimer = 0
	End If
	

'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If RightFlipper.CurrentAngle <= RightFlipper.StartAngle + SOSAngle Then RightFlipper.Strength = FlipperPower * SOSTorque else RightFlipper.Strength = FlipperPower : End If
 
'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
 	If RightFlipperOn = 1 Then
		If RightFlipper.CurrentAngle = RightFlipper.EndAngle Then
			RightFlipper.EOSTorque = FullStrokeEOS_Torque
			RLiveCatchTimer = RLiveCatchTimer + 1
			If RLiveCatchTimer < LiveCatchSensivity Then
				RightFlipper.Elasticity = 0
			Else
				RightFlipper.Elasticity = FlipperElasticity
				RLiveCatchTimer = LiveCatchSensivity
			End If
		End If
	Else
		RightFlipper.Elasticity = FlipperElasticity
		RightFlipper.EOSTorque = LiveStrokeEOS_Torque
		RLiveCatchTimer = 0
	End If

End Sub

'*******************
'  Flipper Subs
'*******************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFXDOF ("LeftflipperupH-2dB", 101, DOFOn, DOFFlippers), LeftFlipper
		DOF 103, DOFPulse
        LeftFlipper.RotateToEnd
        LeftFlipper2.RotateToEnd
		LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFXDOF ("LeftflipperdownH", 101, DOFOff, DOFFlippers), LeftFlipper
		DOF 103, DOFPulse
        LeftFlipper.RotateToStart
        LeftFlipper2.RotateToStart
		LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFXDOF ("RightflipperupH-2dB", 102, DOFOn, DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFXDOF ("RightflipperdownH", 102, DOFOff, DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipperOn = 0
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub LeftFlipper2_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub


'*******************
' GI lights
'*******************

Sub GiOn 'GI lights on
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff 'GI lights off
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
    Next
End Sub

'**************
'     TILT
'**************

Sub CheckTilt
    Tilt = Tilt + TiltSensitivity
    TiltDecreaseTimer.Enabled = True
    If Tilt> 15 Then
        Tilted = True
        TiltReel.SetValue 1
        If B2SOn then
            Controller.B2SSetTilt 1
        end if
        DisableTable True
        'BallsRemaining(CurrentPlayer) = 0 'player looses the game
        TiltRecoveryTimer.Enabled = True 'wait for all the balls to drain
    End If
End Sub

Sub TiltDecreaseTimer_Timer
    ' DecreaseTilt
    If Tilt> 0 Then
        Tilt = Tilt - 0.1
    Else
        TiltDecreaseTimer.Enabled = False
    End If
End Sub

Sub DisableTable(Enabled)
    If Enabled Then
        GiOff
        'Disable slings, bumpers etc
        LeftFlipper.RotateToStart
        LeftFlipper2.RotateToStart
        RightFlipper.RotateToStart
        Bumper001.Threshold = 100
        DOF 101, DOFOff
        DOF 102, DOFOff
    Else
        GiOn
        Bumper001.Threshold = 1
    End If
End Sub

Sub TiltRecoveryTimer_Timer()
    ' all the balls have drained ..
    If(BallsOnPlayfield = 0) Then
        EndOfBall
        TiltRecoveryTimer.Enabled = False
    End If
' otherwise repeat
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
Const maxvel = 28 'max ball velocity
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
        aBallShadow(b).Y = 2000 'under the apron 'may differ from table to table
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' draw the ball shadow
    For b = lob to UBound(BOT)
        aBallShadow(b).X = BOT(b).X
        aBallShadow(b).Y = BOT(b).Y
        aBallShadow(b).Height = BOT(b).Z - Ballsize / 2 + 1

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

Sub aMetals_Hit(idx):PlaySoundAtBall "fx_MetalHit":End Sub
Sub aMetalWires_Hit(idx):PlaySoundAtBall "fx_MetalWire":End Sub
Sub aRubber_Bands_Hit(idx):PlaySoundAtBall "fx_rubber_band":End Sub
Sub aRubber_LongBands_Hit(idx):PlaySoundAtBall "fx_rubber_longband":End Sub
Sub aRubber_Posts_Hit(idx):PlaySoundAtBall "fx_rubber_post":End Sub
Sub aRubber_Pins_Hit(idx):PlaySoundAtBall "fx_rubber_pin":End Sub
Sub aPlastics_Hit(idx):PlaySoundAtBall "fx_PlasticHit":End Sub
Sub aGates_Hit(idx):PlaySoundAtBall "fx_Gate":End Sub
Sub aWoods_Hit(idx):PlaySoundAtBall "fx_Woodhit":End Sub

'***************************************************************************************************
' Only for VPX 10.2 and higher.
' FlashForMs will blink light or a flasher for TotalPeriod(ms) at rate of BlinkPeriod(ms)
' When TotalPeriod done, light or flasher will be set to FinalState value where
' Final State values are:   0=Off, 1=On, 2=Return to previous State
'***************************************************************************************************

Sub FlashForMs(MyLight, TotalPeriod, BlinkPeriod, FinalState)

    If TypeName(MyLight) = "Light" Then
        If FinalState = 2 Then
            FinalState = MyLight.State
        End If
        MyLight.BlinkInterval = BlinkPeriod
        MyLight.Duration 2, TotalPeriod, FinalState
    ElseIf TypeName(MyLight) = "Flasher" Then
        Dim steps
        ' Store all blink information
        steps = Int(TotalPeriod / BlinkPeriod + .5) 
        If FinalState = 2 Then                      
            FinalState = ABS(MyLight.Visible)
        End If
        MyLight.UserValue = steps * 10 + FinalState        
        MyLight.TimerInterval = BlinkPeriod
        MyLight.TimerEnabled = 0
        MyLight.TimerEnabled = 1
        ExecuteGlobal "Sub " & MyLight.Name & "_Timer:" & "Dim tmp, steps, fstate:tmp=me.UserValue:fstate = tmp MOD 10:steps= tmp\10 -1:Me.Visible = steps MOD 2:me.UserValue = steps *10 + fstate:If Steps = 0 then Me.Visible = fstate:Me.TimerEnabled=0:End if:End Sub"
    End If
End Sub

'****************************************
' Init table for a new game
'****************************************

Sub ResetForNewGame()
    'debug.print "ResetForNewGame"
    Dim i

    bGameInPLay = True
    bBallSaverActive = False

    StopAttractMode
    GiOn

    CurrentPlayer = 1
    PlayersPlayingGame = 1
    bOnTheFirstBall = True
    For i = 1 To MaxPlayers
        Score(i) = 0
        ExtraBallsAwards(i) = 0
        Special1Awarded(i) = False
        Special2Awarded(i) = False
        Special3Awarded(i) = False
        Special4Awarded(i) = False
        Special5Awarded(i) = False
        Special6Awarded(i) = False
        BallsRemaining(i) = BallsPerGame
    Next
    SpecialHSAwarded = False
    BonusMultiplier = 1
    Bonus = 0
    UpdateBallInPlay
	UpdatePlayers
    Clear_Match

    ' init other variables
    Tilt = 0

    ' init game variables
    Game_Init()

    ' start a music?
    ' first ball
    vpmtimer.addtimer 2000, "FirstBall '"
End Sub

Sub FirstBall
    'debug.print "FirstBall"
    ' reset table for a new ball, rise droptargets, etc
    ResetForNewPlayerBall()
    CreateNewBall()
End Sub

' (Re-)init table for a new ball or player

Sub ResetForNewPlayerBall()
    'debug.print "ResetForNewPlayerBall"
    UpdatePlayers
    AddScore 0

    ' reset multiplier to 1x
    BonusMultiplier = 1

    ' turn on lights, and variables
    bExtraBallWonThisBall = False
    ResetNewBallVariables
    ResetNewBallLights
End Sub

' Create new ball

Sub CreateNewBall()
    BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass
    BallsOnPlayfield = BallsOnPlayfield + 1
    UpdateBallInPlay
	vpmtimer.addtimer 1000, "PlaySoundAt SoundFXDOF (""fx_Ballrel"", 104, DOFPulse, DOFContactors), Plunger : BallRelease.Kick 90, 4 '"
End Sub

' player lost the ball

Sub EndOfBall
    'debug.print "EndOfBall"

    ' Lost the first ball, now it cannot accept more players
    bOnTheFirstBall = False
	
	' Bonus Count
'	BonusMultiplier = 2 ' For test
    If NOT Tilted Then
		BonusCountTimer.Interval = 300
        BonusCountTimer.Enabled = 1
    Else 
		vpmtimer.addtimer 250, "EndOfBall2 '"
	End If
End Sub

Sub BonusCountTimer_Timer 'The bonus are count when the ball is lost
    'debug.print "BonusCount_Timer"
    If Bonus> 0 Then
        AddScore 10000 * BonusMultiplier
        Bonus = Bonus -1
        UpdateBonusLights
    Else 
        BonusCountTimer.Enabled = 0
        vpmtimer.addtimer 250, "EndOfBall2 '"
    End If
End Sub

Sub EndOfBall2
    'debug.print "EndOfBall2"

    Tilted = False
    Tilt = 0
    TiltReel.SetValue 0
    If B2SOn then
        Controller.B2SSetTilt 0
    end if
    DisableTable False

   ' win extra ball?
    If(ExtraBallsAwards(CurrentPlayer)> 0) Then
        'debug.print "Extra Ball"

        ' if so then give it
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) - 1

        ' turn off light if no more extra balls
        If(ExtraBallsAwards(CurrentPlayer) = 0) Then
			LightShootAgain.State = 0
			ShootAgainR.SetValue 0
			If B2SOn then
				Controller.B2SSetShootAgain 0
			End If
        End If

        ' extra ball sound or light?

        ' reset as in a new ball
        ResetForNewPlayerBall()
        CreateNewBall()

    Else ' no extra ball

        BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer) - 1

        ' last ball?
        If(BallsRemaining(CurrentPlayer) <= 0) Then
            CheckHighScore()
        End If

        ' this is not the last ball, check for new player
        EndOfBallComplete()
    End If
End Sub

Sub EndOfBallComplete()
    'debug.print "EndOfBallComplete"
    Dim NextPlayer

    ' other players?
    If(PlayersPlayingGame> 1) Then
        NextPlayer = CurrentPlayer + 1
        ' if it is the last player then go to the first one
        If(NextPlayer> PlayersPlayingGame) Then
            NextPlayer = 1
        End If
    Else
        NextPlayer = CurrentPlayer
    End If

    'debug.print "Next Player = " & NextPlayer

    ' end of game?
    If((BallsRemaining(CurrentPlayer) <= 0) AND(BallsRemaining(NextPlayer) <= 0) ) Then

      ' match if playing with coins
        If bFreePlay = False Then
            Verification_Match
        End If

        ' end of game
        EndOfGame()
    Else
        ' next player
        CurrentPlayer = NextPlayer

        ' update score
        AddScore 0

        ' reset table for new player
        ResetForNewPlayerBall()
        CreateNewBall()
    End If
End Sub

' Called at the end of the game

Sub EndOfGame()
    'debug.print "EndOfGame"

    bGameInPLay = False
    bJustStarted = False

    ' turn off flippers
    SolLFlipper 0
    SolRFlipper 0

    ' start the attract mode
    StartAttractMode
End Sub

' Function to calculate the balls left
Function Balls
    Dim tmp
    tmp = BallsPerGame - BallsRemaining(CurrentPlayer) + 1
    If tmp> BallsPerGame Then
        Balls = BallsPerGame
    Else
        Balls = tmp
    End If
End Function

' check the highscore
Sub CheckHighscore
    Dim playertops, si, sj, i, stemp, stempplayers
    For i = 1 to 4
        sortscores(i) = 0
        sortplayers(i) = 0
    Next
    playertops = 0
    For i = 1 to PlayersPlayingGame
        sortscores(i) = Score(i)
        sortplayers(i) = i
    Next
    For si = 1 to PlayersPlayingGame
        For sj = 1 to PlayersPlayingGame-1
            If sortscores(sj)> sortscores(sj + 1) then
                stemp = sortscores(sj + 1)
                stempplayers = sortplayers(sj + 1)
                sortscores(sj + 1) = sortscores(sj)
                sortplayers(sj + 1) = sortplayers(sj)
                sortscores(sj) = stemp
                sortplayers(sj) = stempplayers
            End If
        Next
    Next
    HighScoreTimer.interval = 100
    HighScoreTimer.enabled = True
    ScoreChecker = 4
    CheckAllScores = 1
    NewHighScore sortscores(ScoreChecker), sortplayers(ScoreChecker)
End Sub

'******************
'      Match
'******************

Sub Verification_Match()
    Match = INT(RND(1) * 10) * 10 ' random between 00 and 90
    Display_Match
    If(Score(CurrentPlayer) MOD 100) = Match Then
        AwardSpecial      
    End If
End Sub

Sub Clear_Match()
	ScoreReel6.SetValue 0
	LBallInPlay.State = 1
	MatchR.SetValue 0
    If B2SOn then
		Controller.B2SSetReel 27, 0
		Controller.B2SSetReel 28, 0
        Controller.B2SSetMatch 0
        Controller.B2SSetBallInPlay 1
        Controller.B2SSetCanPlay 1
    end if
End Sub

Sub Display_Match()
	ScoreReel6.SetValue Match
	LBallInPlay.State = 0
	MatchR.SetValue 1
    If B2SOn then
        Controller.B2SSetScore 6, Match
        Controller.B2SSetMatch 1
        Controller.B2SSetBallInPlay 0
        Controller.B2SSetCanPlay 0
    end if
End Sub

' *********************************************************************
'                      Drain / Plunger Functions
' *********************************************************************

Sub Drain_Hit()
    Drain.DestroyBall
    BallsOnPlayfield = BallsOnPlayfield - 1
    PlaySoundAt "fx_drain", Drain
	DOF 250, DOFPulse

    'tilted?
    If Tilted Then
        StopEndOfBallMode
    End If

    ' if still playing and not tilted
    If(bGameInPLay = True) AND (Tilted = False) Then

        ' ballsaver?
        If(bBallSaverActive = True) Then
            CreateNewBall()
        Else
            ' last ball?
            If(BallsOnPlayfield = 0) Then
                StopEndOfBallMode
                vpmtimer.addtimer 500, "EndOfBall '" 
                Exit Sub
            End If
        End If
    End If
End Sub

Sub swPlungerRest_Hit()
    bBallInPlungerLane = True
End Sub

Sub swPlungerRest_UnHit()
    bBallInPlungerLane = False
End Sub

' ***************************************
'               Score functions
' ***************************************

Sub AddScore(Points)
    If Tilted Then Exit Sub
    Select Case Points
        Case 10, 100, 1000 , 10000
            Score(CurrentPlayer) = Score(CurrentPlayer) + points
            UpdateScore points
        ' sounds
			If Points = 100 AND(Score(CurrentPlayer) MOD 1000) \ 100 = 0 Then            ' new reel 1000
                PlaySound SoundFXDOF ("1000pts", 303, DOFPulse, DOFChimes)
            ElseIf Points = 10 AND(Score(CurrentPlayer) MOD 100) \ 10 = 0 Then           ' new reel 100
                PlaySound SoundFXDOF ("100pts", 302, DOFPulse, DOFChimes)
			ElseIf Points = 10000 Then
                PlaySound SoundFXDOF ("1000pts", 303, DOFPulse, DOFChimes)  
			ElseIf Points = 1000 Then
                PlaySound SoundFXDOF ("1000pts", 303, DOFPulse, DOFChimes)
			ElseIf Points = 100 Then
                PlaySound SoundFXDOF ("100pts", 302, DOFPulse, DOFChimes)          
			Else
                PlaySound SoundFXDOF ("10pts", 301, DOFPulse, DOFChimes)
            End If 
        Case 20, 30, 40, 50, 60, 70, 80, 90
            Add10 = Add10 + Points \ 10
            AddScore10Timer.Enabled = TRUE       
        Case 200, 300, 400, 500, 600, 700, 800, 900
            Add100 = Add100 + Points \ 100
            AddScore100Timer.Enabled = TRUE   
        Case 2000, 3000, 4000, 5000, 6000, 7000, 8000, 9000
            Add1000 = Add1000 + Points \ 1000
            AddScore1000Timer.Enabled = TRUE
        Case 20000, 30000, 40000, 50000
            Add10000 = Add10000 + Points \ 10000
            AddScore10000Timer.Enabled = TRUE
    End Select

    ' check for higher score and specials for 3 Balls
	If BallsPerGame = 3 Then
		If Score(CurrentPlayer) >= Special1 AND Special1Awarded(CurrentPlayer) = False Then
			AwardSpecial
			Special1Awarded(CurrentPlayer) = True
		End If
		If Score(CurrentPlayer) >= Special2 AND Special2Awarded(CurrentPlayer) = False Then
			AwardSpecial
			Special2Awarded(CurrentPlayer) = True
		End If
		If Score(CurrentPlayer) >= Special3 AND Special3Awarded(CurrentPlayer) = False Then
			AwardSpecial
			Special3Awarded(CurrentPlayer) = True
		End If
	End If
   ' check for higher score and specials for 5 Balls
	If BallsPerGame = 5 Then
		If Score(CurrentPlayer) >= Special4 AND Special4Awarded(CurrentPlayer) = False Then
			AwardSpecial
			Special4Awarded(CurrentPlayer) = True
		End If
		If Score(CurrentPlayer) >= Special5 AND Special5Awarded(CurrentPlayer) = False Then
			AwardSpecial
			Special5Awarded(CurrentPlayer) = True
		End If
		If Score(CurrentPlayer) >= Special6 AND Special6Awarded(CurrentPlayer) = False Then
			AwardSpecial
			Special6Awarded(CurrentPlayer) = True
		End If
	End If
	' check for higher HSScore
	If Score(CurrentPlayer) >= HSScore(1) AND SpecialHSAwarded = False Then
		AwardSpecial
        vpmtimer.addtimer 300, "AwardSpecial '"
		SpecialHSAwarded = True
	End If

End Sub

'************************************
'       Score sound Timers
'************************************

Sub AddScore10Timer_Timer()
    if Add10> 0 then
        AddScore 10
        Add10 = Add10 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore100Timer_Timer()
    if Add100> 0 then
        AddScore 100
        Add100 = Add100 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore1000Timer_Timer()
    if Add1000> 0 then
        AddScore 1000
        Add1000 = Add1000 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore10000Timer_Timer()
    if Add10000> 0 then
        AddScore 10000
        Add10000 = Add10000 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

'*******************
'     Bonus
'*******************

Sub AddBonus
    If Tilted Then Exit Sub
	If Bonus < 5 Then
		Bonus = Bonus + 1
    End If
End Sub

Sub UpdateBonusLights
    If LightBonus5.State = 1 Then 
		LightBonus5.State = 0
		Exit Sub
	End If
    If LightBonus4.State = 1 Then 
		LightBonus4.State = 0
		Exit Sub
	End If
    If LightBonus3.State = 1 Then 
		LightBonus3.State = 0
		Exit Sub
	End If
    If LightBonus2.State = 1 Then 
		LightBonus2.State = 0
		Exit Sub
	End If
    If LightBonus1.State = 1 Then 
		LightBonus1.State = 0
		Exit Sub
	End If
End Sub

'***********************************************************************************
'        Score EM reels - puntuaciones - y actualiza otras luces del backdrop
'***********************************************************************************

Sub UpdateScore(playerpoints)
    Select Case CurrentPlayer
        Case 1:ScoreReel1.Addvalue playerpoints
		Case 2:ScoreReel2.Addvalue playerpoints
		Case 3:ScoreReel3.Addvalue playerpoints
		Case 4:ScoreReel4.Addvalue playerpoints
    End Select
    If B2SOn then
        Controller.B2SSetScore CurrentPlayer, Score(CurrentPlayer)
    end if
End Sub

Sub ResetScores
	ScoreReel1.SetValue 0 'ResetToZero
	ScoreReel2.SetValue 0
	ScoreReel3.SetValue 0
	ScoreReel4.SetValue 0
    HIScoreR.SetValue 0
    If B2SOn then
        Controller.B2SSetScorePlayer1 0
        Controller.B2SSetScorePlayer2 0
        Controller.B2SSetScorePlayer3 0
        Controller.B2SSetScorePlayer4 0
		Controller.B2SSetData 40,0
    end if
End Sub

Sub ResetScoresInit 'Only Init
    If B2SOn then
		Controller.B2SSetScore 5,Credits
		Controller.B2SSetReel 27, 0
		Controller.B2SSetReel 28, 0
    end if
End Sub

Sub AddCredits(value) 'limit to 30 credits
    If Credits <30 Then
        Credits = Credits + value
        UpdateCredits
    end if
End Sub

Sub UpdateCredits
    If Credits> 0 Then
        CreditLight.State = 1
    Else
        CreditLight.State = 0
    End If
	ScoreReel5.SetValue credits
    If B2SOn Then
		Controller.B2SSetScore 5,Credits
    end if
End Sub

Sub UpdateBallInPlay
	ScoreReel6.SetValue Balls + (PlayersPlayingGame*10)
    If B2SOn Then
		Controller.B2SSetReel 27, PlayersPlayingGame
		Controller.B2SSetReel 28, Balls
    End If
End Sub

Sub UpdatePlayers
    Select case CurrentPlayer
        Case 0:pl1.State = 0:pl2.State = 0:pl3.State = 0:pl4.State = 0
        Case 1:pl1.State = 1:pl2.State = 0:pl3.State = 0:pl4.State = 0
        Case 2:pl1.State = 0:pl2.State = 1:pl3.State = 0:pl4.State = 0
        Case 3:pl1.State = 0:pl2.State = 0:pl3.State = 1:pl4.State = 0
        Case 4:pl1.State = 0:pl2.State = 0:pl3.State = 0:pl4.State = 1
    End Select
    If B2SOn Then	
		Controller.B2SSetPlayerUp CurrentPlayer
    End If 
End Sub


'*************************
'        Specials
'*************************

Sub AwardExtraBall()
    If NOT bExtraBallWonThisBall Then
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
        bExtraBallWonThisBall = True
		LightShootAgain.State = 1
		ShootAgainR.SetValue 1
        If B2SOn Then
            Controller.B2SSetShootAgain 1
        End If		
    End If
End Sub

Sub AwardSpecial()
    PlaySound SoundfXDOF ("fx_knocker", 110, DOFPulse, DOFKnocker)
    AddCredits 1
End Sub

' ********************************
'        Attract Mode
' ********************************
' use the"Blink Pattern" of each light

Sub StartAttractMode()
	Dim x
	If AttractMode = 1 Then
		bAttractMode = True
		For each x in aLights
			x.State = 2
		Next
	End If
    If B2SOn then
        Controller.B2SSetGameOver 1
		Controller.B2SSetShootAgain 0
		Controller.B2SSetPlayerUp 0
		Controller.B2SSetTilt 0
    End If
    GameOverR.SetValue 1
	ShootAgainR.SetValue 0
	pl1.State = 0:pl2.State = 0:pl3.State = 0:pl4.State = 0
	TiltReel.SetValue 0
	bAttractHSScore = 0
	HSScoreAttrackTimer.Enabled = 1
End Sub

Sub HSScoreAttrackTimer_Timer
	If bAttractHSScore = 0 Then
		HSScoreON
		bAttractHSScore = 1
	Else
		ScoreON
		bAttractHSScore = 0
	End If	
	If bAttractMode = False Then
		ResetScores
		HSScoreAttrackTimer.Enabled = 0
	End If
End Sub

Sub StopAttractMode()
    Dim x
    bAttractMode = False
    TurnOffPlayfieldLights
    GameOverR.SetValue 0
    If B2SOn then
        Controller.B2SSetGameOver 0
    end if
End Sub

Sub HSScoreON()
	ScoreReel1.SetValue HSScore(1)
	ScoreReel2.SetValue 0
	ScoreReel3.SetValue 0
	ScoreReel4.SetValue 0
    HIScoreR.SetValue 1 'Lite HSSCore
	If B2SOn then
        Controller.B2SSetScorePlayer1 HSScore(1)
        Controller.B2SSetScorePlayer2 0
        Controller.B2SSetScorePlayer3 0
        Controller.B2SSetScorePlayer4 0
		Controller.B2SSetData 40,1 'Lite HSSCore
	End If
End Sub

Sub ScoreON()
	ScoreReel1.SetValue Score(1)
	ScoreReel2.SetValue Score(2)
	ScoreReel3.SetValue Score(3)
	ScoreReel4.SetValue Score(4)
    HIScoreR.SetValue 0 'UnLite HSSCore
	If B2SOn then
        Controller.B2SSetScorePlayer1 Score(1)
        Controller.B2SSetScorePlayer2 Score(2)
        Controller.B2SSetScorePlayer3 Score(3)
        Controller.B2SSetScorePlayer4 Score(4)
		Controller.B2SSetData 40,0 'Unlite HSSCore
	End If
End Sub

'************************************************
'    Load / Save / Highscore
'************************************************

Sub Loadhs
    ' Based on Black's Highscore routines
    Dim FileObj
    Dim ScoreFile, TextStr
    Dim temp1
    Dim temp2
    Dim temp3
    Dim temp4
    Dim temp5
    Dim temp6
    Dim temp8
    Dim temp9
    Dim temp10
    Dim temp11
    Dim temp12
    Dim temp13
    Dim temp14
    Dim temp15
    Dim temp16
    Dim temp17

    Set FileObj = CreateObject("Scripting.FileSystemObject")
    If Not FileObj.FolderExists(UserDirectory) then
        Credits = 0
        Exit Sub
    End If
    If Not FileObj.FileExists(UserDirectory & TableName& ".txt") then
        Credits = 0
        Exit Sub
    End If
    Set ScoreFile = FileObj.GetFile(UserDirectory & TableName& ".txt")
    Set TextStr = ScoreFile.OpenAsTextStream(1, 0)
    If(TextStr.AtEndOfStream = True) then
        Exit Sub
    End If
    temp1 = TextStr.ReadLine
    temp2 = textstr.readline

    HighScore = cdbl(temp1)
    If HighScore <1 then
        temp8 = textstr.readline
        temp9 = textstr.readline
        temp10 = textstr.readline
        temp11 = textstr.readline
        temp12 = textstr.readline
        temp13 = textstr.readline
        temp14 = textstr.readline
        temp15 = textstr.readline
        temp16 = textstr.readline
        temp17 = textstr.readline
    End If
    TextStr.Close
    Credits = cdbl(temp2)

    If HighScore <1 then
        HSScore(1) = int(temp8)
        HSScore(2) = int(temp9)
        HSScore(3) = int(temp10)
        HSScore(4) = int(temp11)
        HSScore(5) = int(temp12)

        HSName(1) = temp13
        HSName(2) = temp14
        HSName(3) = temp15
        HSName(4) = temp16
        HSName(5) = temp17
    End If
    Set ScoreFile = Nothing
    Set FileObj = Nothing
    SortHighscore 'added to fix a previous error
End Sub

Sub Savehs
    ' Based on Black's Highscore routines
    Dim FileObj
    Dim ScoreFile
    Dim xx
    Set FileObj = CreateObject("Scripting.FileSystemObject")
    If Not FileObj.FolderExists(UserDirectory) then
        Exit Sub
    End If
    Set ScoreFile = FileObj.CreateTextFile(UserDirectory & TableName& ".txt", True)
    ScoreFile.WriteLine 0
    ScoreFile.WriteLine Credits
    For xx = 1 to 5
        scorefile.writeline HSScore(xx)
    Next
    For xx = 1 to 5
        scorefile.writeline HSName(xx)
    Next
    ScoreFile.Close
    Set ScoreFile = Nothing
    Set FileObj = Nothing
End Sub

Sub SortHighscore
    Dim tmp, tmp2, i, j
    For i = 1 to 5
        For j = 1 to 4
            If HSScore(j) <HSScore(j + 1) Then
                tmp = HSScore(j + 1)
                tmp2 = HSName(j + 1)
                HSScore(j + 1) = HSScore(j)
                HSName(j + 1) = HSName(j)
                HSScore(j) = tmp
                HSName(j) = tmp2
            End If
        Next
    Next
End Sub

'****************************************
' Realtime updates
'****************************************

Sub GameTimer_Timer
    RollingUpdate
    LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle
    RightFlipperTop.RotZ = RightFlipper.CurrentAngle
	Pgate1.rotz = (Gate1.currentangle*.75)+25
End Sub

'***********************************************************************
' *********************************************************************
'  *********     G A M E  C O D E  S T A R T S  H E R E      *********
' *********************************************************************
'***********************************************************************

Sub VPObjects_Init 'init objects
    TurnOffPlayfieldLights()
End Sub

' Dim all the variables

Sub Game_Init() 'called at the start of a new game
    'Start music?
    'Init variables?
	SpecialAwarded = False
	CenterBonusCounter = 0
    'Start or init timers
    'Init lights?
    TurnOffPlayfieldLights()
End Sub

Sub StopEndOfBallMode() 'called when the last ball is drained

End Sub

Sub ResetNewBallVariables() 'init variables new ball/player		
	vpmtimer.addtimer 500,"ResetTarget5 '" 
	Bonus = 0
	BonusMultiplier	= 1	
	SpecialAwarded = False
	bSpecial = False
End Sub

Sub ResetNewBallLights()    'init lights for new ball/player
	TurnOffPlayfieldLights()
	TurnOnNewBallLights()
	UpdateCenterLights()	
End Sub

Sub TurnOnNewBallLights()
    Dim a
    For each a in aNewBallLights
        a.State = 1
    Next
End Sub

Sub TurnOffPlayfieldLights()
    Dim a
    For each a in aLights
        a.State = 0
    Next
End Sub

' *********************************************************************
'                        Table Object Hit Events
' *********************************************************************
' Any target hit Sub will do this:
' - play a sound
' - do some physical movement
' - add a score, bonus
' - check some variables/modes this trigger is a member of
' - set the "LastSwicthHit" variable in case it is needed later
' *********************************************************************

'***********************
'       Rubbers
'***********************

Dim Rub1, Rub2, Rub3, Rub4, Rub5


Sub RubberBand001_Hit 'top targets
    If Tilted then Exit Sub
	UpdateCenterLights
	AddScore 10
    Rub1 = 1:RubberBand001_Timer
End Sub

Sub RubberBand001_Timer
    Select Case Rub1
        Case 1:r024.Visible = 0:r029.Visible = 1:RubberBand001.TimerEnabled = 1
        Case 2:r029.Visible = 0:r030.Visible = 1
        Case 3:r030.Visible = 0:r024.Visible = 1:RubberBand001.TimerEnabled = 0
    End Select
    Rub1 = Rub1 + 1
End Sub

Sub RubberBand002_Hit 'center target left
    If Tilted then Exit Sub
	UpdateCenterLights
	AddScore 10
    Rub2 = 1:RubberBand002_Timer
End Sub

Sub RubberBand002_Timer
    Select Case Rub2
        Case 1:r017.Visible = 0:r033.Visible = 1:RubberBand002.TimerEnabled = 1
        Case 2:r033.Visible = 0:r034.Visible = 1
        Case 3:r034.Visible = 0:r017.Visible = 1:RubberBand002.TimerEnabled = 0
    End Select
    Rub2 = Rub2 + 1
End Sub

Sub RubberBand003_Hit 'center target back
    If Tilted then Exit Sub
	UpdateCenterLights
	AddScore 10
    Rub3 = 1:RubberBand003_Timer
End Sub

Sub RubberBand003_Timer
    Select Case Rub3
        Case 1:r017.Visible = 0:r031.Visible = 1:RubberBand003.TimerEnabled = 1
        Case 2:r031.Visible = 0:r032.Visible = 1
        Case 3:r032.Visible = 0:r017.Visible = 1:RubberBand003.TimerEnabled = 0
    End Select
    Rub3 = Rub3 + 1
End Sub

Sub RubberBand004_Hit 'top right
    If Tilted then Exit Sub
	UpdateCenterLights
	AddScore 10
    Rub4 = 1:RubberBand004_Timer
End Sub

Sub RubberBand004_Timer
    Select Case Rub4
        Case 1:r027.Visible = 0:r035.Visible = 1:RubberBand004.TimerEnabled = 1
        Case 2:r035.Visible = 0:r036.Visible = 1
        Case 3:r036.Visible = 0:r027.Visible = 1:RubberBand004.TimerEnabled = 0
    End Select
    Rub4 = Rub4 + 1
End Sub

Sub RubberBand005_Hit 'bottom left
    If Tilted then Exit Sub
	UpdateCenterLights
	AddScore 10
    Rub5 = 1:RubberBand005_Timer
End Sub

Sub RubberBand005_Timer
    Select Case Rub5
        Case 1:r009.Visible = 0:r037.Visible = 1:RubberBand005.TimerEnabled = 1
        Case 2:r037.Visible = 0:r038.Visible = 1
        Case 3:r038.Visible = 0:r009.Visible = 1:RubberBand005.TimerEnabled = 0
    End Select
    Rub5 = Rub5 + 1
End Sub

'*********
' Bumpers
'*********

Sub Bumper001_Hit
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF ("fx_Bumper",105,DOFPulse,DOFContactors), bumper001
    If BumperLight.State = 1 Then AddScore 500 Else AddScore 100
End Sub

'*****************
'     Triggers
'*****************

' bottom lane

Sub Trigger001_Hit 'outlane left
    PlaySoundAt "fx_sensor", Trigger001
	DOF 201, DOFPulse
    If Tilted Then Exit Sub
	Addscore 5000
End Sub

Sub Trigger002_Hit 'black horseleft
    PlaySoundAt "fx_sensor", Trigger002
	DOF 202, DOFPulse
    If Tilted Then Exit Sub	
	If LHorseBlack2.State = 1 Then AddScore 5000 Else AddScore 500
End Sub

Sub Trigger003_Hit 'bonus 1 left
    PlaySoundAt "fx_sensor", Trigger003
	DOF 203, DOFPulse
    If Tilted Then Exit Sub
	If LB12.State = 1 Then
		LB11.State = 0
		LB12.State = 0
		LB13.State = 0
		LightBonus1.State = 1
		Lighttarget20.State = 0
		AddBonus
		CheckExtraBall
	End If
		AddScore 1000
End Sub

Sub Trigger004_Hit 'bonus 1 right
    PlaySoundAt "fx_sensor", Trigger004
	DOF 203, DOFPulse
    If Tilted Then Exit Sub
	If LB13.State = 1 Then
		LB11.State = 0
		LB12.State = 0
		LB13.State = 0
		LightBonus1.State = 1
		Lighttarget20.State = 0
		AddBonus
		CheckExtraBall
	End If
		AddScore 1000
End Sub

Sub Trigger005_Hit 'black horse right
    PlaySoundAt "fx_sensor", Trigger005
	DOF 202, DOFPulse
    If Tilted Then Exit Sub
	If LHorseBlack3.State = 1 Then AddScore 5000 Else AddScore 500
End Sub

Sub Trigger006_Hit 'outlane right
    PlaySoundAt "fx_sensor", Trigger006
	DOF 201, DOFPulse
    If Tilted Then Exit Sub	
	Addscore 5000
End Sub


' center

Sub Trigger007_Hit 'bonus 5 right
    PlaySoundAt "fx_sensor", Trigger007
	DOF 203, DOFPulse
    If Tilted Then Exit Sub
	If LB52.State = 1 Then
		LB51.State = 0
		LB52.State = 0
		LightBonus5.State = 1
		Lighttarget24.State = 0
		AddBonus
		CheckExtraBall
	End If
		AddScore 1000
End Sub

Sub Trigger008_Hit 'bonus 4 right
    PlaySoundAt "fx_sensor", Trigger008
	DOF 203, DOFPulse
    If Tilted Then Exit Sub
	If LB41.State = 1 Then
		LB41.State = 0
		LB42.State = 0
		LightBonus4.State = 1
		Lighttarget23.State = 0
		AddBonus
		CheckExtraBall
	End If
		AddScore 1000
End Sub

Sub Trigger009_Hit 'bonus 5 left
    PlaySoundAt "fx_sensor", Trigger009
	DOF 203, DOFPulse
    If Tilted Then Exit Sub
	If LB51.State = 1 Then
		LB51.State = 0
		LB52.State = 0
		LightBonus5.State = 1
		Lighttarget24.State = 0
		AddBonus
		CheckExtraBall
	End If
		AddScore 1000
End Sub

' top lane

Sub Trigger010_Hit 'bonus 1 left
    PlaySoundAt "fx_sensor", Trigger010
	DOF 203, DOFPulse
    If Tilted Then Exit Sub
	If LB11.State = 1 Then
		LB11.State = 0
		LB12.State = 0
		LB13.State = 0
		LightBonus1.State = 1
		Lighttarget20.State = 0
		AddBonus
		CheckExtraBall
	End If
		AddScore 1000
End Sub

Sub Trigger011_Hit 'black horse
    PlaySoundAt "fx_sensor", Trigger011
	DOF 202, DOFPulse
    If Tilted Then Exit Sub
	If LHorseBlack1.State = 1 Then
		LHorseBlack1.State = 0
		LHorseBlack2.State = 1
		LHorseBlack3.State = 1
	End If
	Addscore 500
	BumperLight.State = 1
	BumperLight1.State = 1
End Sub

Sub Trigger012_Hit 'bonus 2 right
    PlaySoundAt "fx_sensor", Trigger012
	DOF 203, DOFPulse
    If Tilted Then Exit Sub
	If LB21.State = 1 Then
		LB21.State = 0
		LB22.State = 0
		LightBonus2.State = 1
		Lighttarget21.State = 0
		AddBonus
		CheckExtraBall
	End If
		AddScore 1000
End Sub

'************************
'       Targets
'************************

' Round Targets

Sub Target1_hit 'bonus 2 left
    PlaySoundAtBall "fx_target"
	DOF 203, DOFPulse
    If Tilted Then Exit Sub
	If LB22.State = 1 Then
		LB21.State = 0
		LB22.State = 0
		LightBonus2.State = 1
		Lighttarget21.State = 0
		AddBonus
		CheckExtraBall
	End If
	Addscore 1000
End Sub

Sub Target2_hit 'bonus 3 left
    PlaySoundAtBall "fx_target"
	DOF 203, DOFPulse
    If Tilted Then Exit Sub
	If LB31.State = 1 Then
		LB31.State = 0
		LB32.State = 0
		LightBonus3.State = 1
		Lighttarget22.State = 0
		AddBonus
		CheckExtraBall
	End If
	Addscore 1000
End Sub

Sub Target3_hit  'bonus 4 right
    PlaySoundAtBall "fx_target"
	DOF 203, DOFPulse
    If Tilted Then Exit Sub	
	If LB42.State = 1 Then
		LB41.State = 0
		LB42.State = 0
		LightBonus4.State = 1
		Lighttarget23.State = 0
		AddBonus
		CheckExtraBall
	End If
	Addscore 1000
End Sub

Sub Target4_hit  'bonus 3 right
    PlaySoundAtBall "fx_target"
	DOF 203, DOFPulse
    If Tilted Then Exit Sub	
	If LB32.State = 1 Then
		LB31.State = 0
		LB32.State = 0
		LightBonus3.State = 1
		Lighttarget22.State = 0
		AddBonus
		CheckExtraBall
	End If
	Addscore 1000
End Sub

Sub Target5_hit  'Top Special
    PlaySoundAtBall "fx_target"
	DOF 208, DOFPulse
    If Tilted Then Exit Sub
	If LightSpecial.State = 1 Then
		AwardSpecial
		SpecialAwarded = True
		LightSpecial.State = 0
		Addscore 20000
	Else
		Addscore 1000 * BonusMultiplier
	End If
End Sub

' 4 Center Round Targets  

Sub Target10_hit 'Left
	PlaySoundAtBall "fx_target"
	DOF 209, DOFPulse
    If Tilted Then Exit Sub
	If LightTarget10.State = 0 Then AddScore 1000 * BonusMultiplier
	If LightTarget10.State = 1 Then AddScore 10000 * BonusMultiplier
End Sub

Sub Target11_hit
	PlaySoundAtBall "fx_target"
	DOF 209, DOFPulse
    If Tilted Then Exit Sub
	If LightTarget11.State = 0 Then AddScore 1000 * BonusMultiplier
	If LightTarget11.State = 1 Then AddScore 10000 * BonusMultiplier
End Sub

Sub Target12_hit
	PlaySoundAtBall "fx_target"
	DOF 209, DOFPulse
    If Tilted Then Exit Sub
	If LightTarget12.State = 0 Then AddScore 1000 * BonusMultiplier
	If LightTarget12.State = 1 Then AddScore 10000 * BonusMultiplier
End Sub

Sub Target13_hit 'Right
	PlaySoundAtBall "fx_target"
	DOF 209, DOFPulse
    If Tilted Then Exit Sub
	If LightTarget13.State = 0 Then AddScore 1000 * BonusMultiplier
	If LightTarget13.State = 1 Then AddScore 10000 * BonusMultiplier
End Sub

' 5 Top Targets

Sub Target20_hit ' Left
    PlaySoundAt SoundFXDOF ("fx_droptarget", 111, DOFPulse, DOFContactors), Target20
    If Tilted Then Exit Sub
	CheckDoubleBonus
	AddScore 1000
End Sub

Sub Target21_hit
    PlaySoundAt SoundFXDOF ("fx_droptarget", 111, DOFPulse, DOFContactors), Target21
    If Tilted Then Exit Sub
	CheckDoubleBonus
	AddScore 1000
End Sub

Sub Target22_hit
    PlaySoundAt SoundFXDOF ("fx_droptarget", 111, DOFPulse, DOFContactors), Target22
    If Tilted Then Exit Sub
	CheckDoubleBonus
	AddScore 1000
End Sub

Sub Target23_hit
    PlaySoundAt SoundFXDOF ("fx_droptarget", 111, DOFPulse, DOFContactors), Target23
    If Tilted Then Exit Sub
	CheckDoubleBonus
	AddScore 1000
End Sub

Sub Target24_hit'Right
    PlaySoundAt SoundFXDOF ("fx_droptarget", 111, DOFPulse, DOFContactors), Target24
    If Tilted Then Exit Sub 
	CheckDoubleBonus
	AddScore 1000
End Sub

'Reset 5 top Targets

Sub ResetTarget5
    If Tilted Then Exit Sub
	Target20.Isdropped= 0 : Target21.Isdropped= 0 : Target22.Isdropped= 0 : Target23.Isdropped= 0 : Target24.Isdropped= 0	
	PlaySoundAt SoundFXDOF ("fx_resetdrop", 111, DOFPulse, DOFDropTargets), Bumper001
End Sub


'*******************
'     Rollover
'*******************

Sub sw001_Hit
    If Tilted Then Exit Sub  
	LTrigger1.State = 0
	LTriggerTimer.Enabled = True
	AddScore 500 	
	BumperLight.State = 1
	BumperLight1.State = 1
	If bSpecial = False Then UpdateCenterLights 
End Sub

Sub sw002_Hit
    If Tilted Then Exit Sub
	LTrigger2.State = 0
	LTriggerTimer.Enabled = True
	AddScore 500  
	BumperLight.State = 1
	BumperLight1.State = 1
	If bSpecial = False Then UpdateCenterLights 
End Sub

Sub sw003_Hit
    If Tilted Then Exit Sub
	LTrigger3.State = 0
	LTriggerTimer.Enabled = True
	AddScore 500  
	BumperLight.State = 0
	BumperLight1.State = 0
	If bSpecial = False Then UpdateCenterLights
End Sub

Sub sw004_Hit
    If Tilted Then Exit Sub
	LTrigger4.State = 0
	LTriggerTimer.Enabled = True
	AddScore 500  
	BumperLight.State = 0 
	BumperLight1.State = 0
	If bSpecial = False Then UpdateCenterLights
End Sub

Sub LTriggerTimer_Timer
	LTrigger1.State = 1
	LTrigger2.State = 1
	LTrigger3.State = 1
	LTrigger4.State = 1
	LTriggerTimer.Enabled = False
End Sub

'******************************
'      Extra routines
'******************************

Sub CheckDoubleBonus()
	If LightTarget20.State = 1 And LightTarget21.State = 1 And LightTarget22.State = 1 And LightTarget23.State = 1 And LightTarget24.State = 1 Then
		LightDoubleBonus.State = 1
		BonusMultiplier	= 2	
		CheckSpecial		
	End If
End Sub

Sub CheckExtraBall()
	If Bonus = 5 Then
		AwardExtraBall
		CheckSpecial
	End If
End Sub

Sub CheckSpecial()
	If LightDoubleBonus.State = 1 And LightShootAgain.State = 1 Then
		LightSpecial.State = 1
		LightTarget10.State = 1
		LightTarget11.State = 1
		LightTarget12.State = 1
		LightTarget13.State = 1
		bSpecial = True
	End If
End Sub

Sub UpdateCenterLights()
	If CenterBonusCounter < 4 Then CenterBonusCounter = CenterBonusCounter + 1 Else CenterBonusCounter = 1
	Select Case CenterBonusCounter
	Case 1 : LightTarget10.State = 1 : LightTarget11.State = 0 : LightTarget12.State = 0 : LightTarget13.State = 0
	Case 2 : LightTarget10.State = 0 : LightTarget11.State = 1 : LightTarget12.State = 0 : LightTarget13.State = 0
	Case 3 : LightTarget10.State = 0 : LightTarget11.State = 0 : LightTarget12.State = 1 : LightTarget13.State = 0
	Case 4 : LightTarget10.State = 0 : LightTarget11.State = 0 : LightTarget12.State = 0 : LightTarget13.State = 1
	End Select
End Sub


Sub CT_timer() 'Targets Side Walls
	If Target20.IsDropped=false Then
		Target20A.IsDropped=false
		Target20B.IsDropped=false
		else
		Target20A.IsDropped=true
		Target20B.IsDropped=true
	end If
	If Target21.IsDropped=false Then
		Target21A.IsDropped=false
		Target21B.IsDropped=false
		else
		Target21A.IsDropped=true
		Target21B.IsDropped=true
	end If
	If Target22.IsDropped=false Then
		Target22A.IsDropped=false
		Target22B.IsDropped=false
		else
		Target22A.IsDropped=true
		Target22B.IsDropped=true
	end If
	If Target23.IsDropped=false Then
		Target23A.IsDropped=false
		Target23B.IsDropped=false
		else
		Target23A.IsDropped=true
		Target23B.IsDropped=true
	end If
	If Target24.IsDropped=false Then
		Target24A.IsDropped=false
		Target24B.IsDropped=false
		else
		Target24A.IsDropped=true
		Target24B.IsDropped=true
	end If
End Sub

' ============================================================================================
' GNMOD - Multiple High Score Display and Collection
' jpsalas: changed ramps by flashers
' ============================================================================================

Dim EnteringInitials ' Normally zero, set to non-zero to enter initials
EnteringInitials = False
Dim ScoreChecker
ScoreChecker = 0
Dim CheckAllScores
CheckAllScores = 0
Dim sortscores(4)
Dim sortplayers(4)

Dim PlungerPulled
PlungerPulled = 0

Dim SelectedChar   ' character under the "cursor" when entering initials

Dim HSTimerCount   ' Pass counter For HS timer, scores are cycled by the timer
HSTimerCount = 5   ' Timer is initially enabled, it'll wrap from 5 to 1 when it's displayed

Dim InitialString  ' the string holding the player's initials as they're entered

Dim AlphaString    ' A-Z, 0-9, space (_) and backspace (<)
Dim AlphaStringPos ' pointer to AlphaString, move Forward and backward with flipper keys
AlphaString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_<"

Dim HSNewHigh      ' The new score to be recorded

Dim HSScore(5)     ' High Scores read in from config file
Dim HSName(5)      ' High Score Initials read in from config file

' default high scores, remove this when the scores are available from the config file
HSScore(1) = 150000
HSScore(2) = 130000
HSScore(3) = 110000
HSScore(4) = 90000
HSScore(5) = 70000

HSName(1) = "AAA"
HSName(2) = "ZZZ"
HSName(3) = "XXX"
HSName(4) = "ABC"
HSName(5) = "BBB"

Sub HighScoreTimer_Timer
    If EnteringInitials then
        If HSTimerCount = 1 then
            SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
            HSTimerCount = 2
        Else
            SetHSLine 3, InitialString
            HSTimerCount = 1
        End If
    ElseIf bGameInPlay then
        SetHSLine 1, "HIGH SCORE1"
        SetHSLine 2, HSScore(1)
        SetHSLine 3, HSName(1)
        HSTimerCount = 5 ' set so the highest score will show after the game is over
        HighScoreTimer.enabled = false
    ElseIf CheckAllScores then
        NewHighScore sortscores(ScoreChecker), sortplayers(ScoreChecker)
    Else
        ' cycle through high scores
        HighScoreTimer.interval = 2000
        HSTimerCount = HSTimerCount + 1
        If HsTimerCount> 5 then
            HSTimerCount = 1
        End If
        SetHSLine 1, "HIGH SCORE" + FormatNumber(HSTimerCount, 0)
        SetHSLine 2, HSScore(HSTimerCount)
        SetHSLine 3, HSName(HSTimerCount)
    End If
End Sub

Function GetHSChar(String, Index)
    Dim ThisChar
    Dim FileName
    ThisChar = Mid(String, Index, 1)
    FileName = "PostIt"
    If ThisChar = " " or ThisChar = "" then
        FileName = FileName & "BL"
    ElseIf ThisChar = "<" then
        FileName = FileName & "LT"
    ElseIf ThisChar = "_" then
        FileName = FileName & "SP"
    Else
        FileName = FileName & ThisChar
    End If
    GetHSChar = FileName
End Function

Sub SetHsLine(LineNo, String)
    Dim Letter
    Dim ThisDigit
    Dim ThisChar
    Dim StrLen
    Dim LetterLine
    Dim Index
    Dim StartHSArray
    Dim EndHSArray
    Dim LetterName
    Dim xFor
    StartHSArray = array(0, 1, 12, 22)
    EndHSArray = array(0, 11, 21, 31)
    StrLen = len(string)
    Index = 1

    For xFor = StartHSArray(LineNo) to EndHSArray(LineNo)
        Eval("HS" &xFor).imageA = GetHSChar(String, Index)
        Index = Index + 1
    Next
End Sub

Sub NewHighScore(NewScore, PlayNum)
    If NewScore> HSScore(5) then
        HighScoreTimer.interval = 500
        HSTimerCount = 1
        AlphaStringPos = 1      ' start with first character "A"
        EnteringInitials = true ' intercept the control keys while entering initials
        InitialString = ""      ' initials entered so far, initialize to empty
        SetHSLine 1, "PLAYER " + FormatNumber(PlayNum, 0)
        SetHSLine 2, "ENTER NAME"
        SetHSLine 3, MID(AlphaString, AlphaStringPos, 1)
        HSNewHigh = NewScore
        PlaySound "sfx_Enter"
    End If
    ScoreChecker = ScoreChecker-1
    If ScoreChecker = 0 then
        CheckAllScores = 0
    End If
End Sub

Sub CollectInitials(keycode)
    Dim i
    If keycode = LeftFlipperKey Then
        ' back up to previous character
        AlphaStringPos = AlphaStringPos - 1
        If AlphaStringPos <1 then
            AlphaStringPos = len(AlphaString) ' handle wrap from beginning to End
            If InitialString = "" then
                ' Skip the backspace If there are no characters to backspace over
                AlphaStringPos = AlphaStringPos - 1
            End If
        End If
        SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
        PlaySound "sfx_Previous"
    ElseIf keycode = RightFlipperKey Then
        ' advance to Next character
        AlphaStringPos = AlphaStringPos + 1
        If AlphaStringPos> len(AlphaString) or(AlphaStringPos = len(AlphaString) and InitialString = "") then
            ' Skip the backspace If there are no characters to backspace over
            AlphaStringPos = 1
        End If
        SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
        PlaySound "sfx_Next"
    ElseIf keycode = StartGameKey or keycode = PlungerKey Then
        SelectedChar = MID(AlphaString, AlphaStringPos, 1)
        If SelectedChar = "_" then
            InitialString = InitialString & " "
            PlaySound("sfx_Esc")
        ElseIf SelectedChar = "<" then
            InitialString = MID(InitialString, 1, len(InitialString) - 1)
            If len(InitialString) = 0 then
                ' If there are no more characters to back over, don't leave the < displayed
                AlphaStringPos = 1
            End If
            PlaySound("sfx_Esc")
        Else
            InitialString = InitialString & SelectedChar
            PlaySound("sfx_Enter")
        End If
        If len(InitialString) <3 then
            SetHSLine 3, InitialString & SelectedChar
        End If
    End If
    If len(InitialString) = 3 then
        ' save the score
        For i = 5 to 1 step -1
            If i = 1 or(HSNewHigh> HSScore(i) and HSNewHigh <= HSScore(i - 1) ) then
                ' Replace the score at this location
                If i <5 then
                    HSScore(i + 1) = HSScore(i)
                    HSName(i + 1) = HSName(i)
                End If
                EnteringInitials = False
                HSScore(i) = HSNewHigh
                HSName(i) = InitialString
                HSTimerCount = 5
                HighScoreTimer_Timer
                HighScoreTimer.interval = 2000
                PlaySound("fx_Bong")
                Exit Sub
            ElseIf i <5 then
                ' move the score in this slot down by 1, it's been exceeded by the new score
                HSScore(i + 1) = HSScore(i)
                HSName(i + 1) = HSName(i)
            End If
        Next
    End If
End Sub
' End GNMOD


' ***************** VR Clock code below - THANKS RASCAL ******************
Dim CurrentMinute ' for VR clock 
' VR Clock code below....
Sub ClockTimerAnalog_Timer()

    'ClockHands Below
	VR_Clock_minutes.RotAndTra2 = (Minute(Now())+(Second(Now())/100))*6
	VR_Clock_Hours.RotAndTra2 = Hour(Now())*30+(Minute(Now())/2)
      VR_Clock_Seconds.RotAndTra2 = (Second(Now()))*6
	CurrentMinute=Minute(Now())

End Sub

 ' ********************** END CLOCK CODE   *********************************

'----- VR Room Auto-Detect -----
Dim VR_Obj, VRRoom

VRRoom = 2

Sub Table1_OptionEvent(ByVal eventId)
'VR Room
VRRoom = Table1.Option("VR Room", 1, 3, 1, 3, 0, Array("Basti's ROOM", "Mixed Reality", "Pool Bar"))
	SetupVRRoom
'Ambient
    Ambience = Table1.Option("Ambience", 1, 2, 1, 2, 0, Array("ON", "OFF"))
	SetupAmbience

'Music Background
    MusicChoice = Table1.Option("Background Music", 1, 2, 1, 2, 0, Array("ON", "OFF"))
	SetupMusic

'Coin Slots
    CoinChoice = Table1.Option("Coin Slots", 1, 2, 1, 2, 0, Array("USA", "Spain"))
	Setupcoinslots
End Sub

Sub SetupVRRoom()
	If RenderingMode = 2 Then
		If VRRoom = 3 Then
			For Each VR_Obj in VR_MinimalRoom : VR_Obj.Visible = 0 : Next
			For Each VR_Obj in VR_PoolBar : VR_Obj.Visible = 1 : Next
			For Each VR_Obj in VR_Sphere : VR_Obj.Visible = 0 : Next
			For Each x in aReels : x.Visible = 0 : Next
			Flasher1.visible = 0
			Flasher2.visible = 0
			lrail1.visible = 0
			rrail1.visible = 0
			rrail.visible = 0
			lrail.visible = 0		
		ElseIf VRRoom = 2 Then
			For Each VR_Obj in VR_MinimalRoom : VR_Obj.Visible = 0 : Next
			For Each VR_Obj in VR_PoolBar : VR_Obj.Visible = 0 : Next
			For Each VR_Obj in VR_Sphere : VR_Obj.Visible = 1 : Next			
			For Each x in aReels : x.Visible = 0 : Next
			Flasher1.visible = 0
			Flasher2.visible = 0
			lrail1.visible = 0
			rrail1.visible = 0
			rrail.visible = 0
			lrail.visible = 0			
		Else
			For Each VR_Obj in VR_MinimalRoom : VR_Obj.Visible = 1 : Next
			For Each VR_Obj in VR_PoolBar : VR_Obj.Visible = 0 : Next
			For Each VR_Obj in VR_Sphere : VR_Obj.Visible = 0 : Next
			For Each x in aReels : x.Visible = 0 : Next
			Flasher1.visible = 0
			Flasher2.visible = 0
			lrail1.visible = 0
			rrail1.visible = 0
			rrail.visible = 0
			lrail.visible = 0			
		End If
	Else
		VRRoom = 0
		For Each VR_Obj in VR_MinimalRoom : VR_Obj.Visible = 0 : Next
		For Each VR_Obj in VR_PoolBar : VR_Obj.Visible = 0 : Next
		For Each VR_Obj in VR_Sphere : VR_Obj.Visible = 0 : Next
		For Each VR_Obj in VR_Table : VR_Obj.Visible = 0 : Next
	If Table1.ShowDT then	
		Flasher1.visible = 1
		Flasher2.visible = 1
		lrail1.visible = 1
		rrail1.visible = 1
		rrail.visible = 1
		lrail.visible = 1		
		For each x in aReels
            x.Visible = 1
        Next
    Else
		Flasher1.visible = 0
		Flasher2.visible = 0
		lrail1.visible = 0
		rrail1.visible = 0
		rrail.visible = 0
		lrail.visible = 0		
        For each x in aReels
            x.Visible = 0
        Next
    End If
End If
End Sub

'**********
' Digital Clock
'**********


Sub ClockTimer_Timer()

    ' ROB AND WALTER'S Digital Clock below**************************************
	dim n
    n=Hour(now) MOD 12
    if n = 0 then n = 12
	hour1.imagea="digit" & CStr(n \ 10)
    hour2.imagea="digit" & CStr(n mod 10)
	n=Minute(now)
	minute1.imagea="digit" & CStr(n \ 10)
    minute2.imagea="digit" & CStr(n mod 10)
	'n=Second(now)
	'second1.imagea="digit" & CStr(n \ 10)
    'second2.imagea="digit" & CStr(n mod 10)
End Sub


'******
' TV Timer
'******

Const TVCounterMax = 16

TimerTV.enabled = True

Dim TVCounter: TVCounter = 1

Sub TimerTV_Timer()

if TVCounter < 10 then 
        TV.Image = "ezgif-frame-" & "00" & TVCounter
    elseif TVCounter < 16 then
        TV.Image = "ezgif-frame-" & "0" & TVCounter 
    else'if TVCounter < 150 then
        'TV.Image = "ezgif-frame-" & "" & TVCounter 
    'else 
        TV.image = "ezgif-frame-" & TVCounter
     end if

    TVCounter = TVCounter + 1 

    If TVCounter > TVCounterMax Then
        TVCounter = 1
   End If

End Sub

'****************
'Fan Animation'
'****************

FanTimer.enabled = True

sub Fantimer_timer()
Fan.Objrotz = Fan.Objrotz + 2
End Sub

'*********************************************************************************************************
' Song & Music
'*********************************************************************************************************


Dim Ambience : Ambience = 1				'1 - ON, 2 - OFF

Sub Setupambience()
If Ambience = 1 then  PlaySound"Ambience"
If Ambience = 2 then  StopSound"Ambience"

End Sub

Dim Musicchoice : Musicchoice = 1				'1 - ON, 2 - OFF

Sub Setupmusic()
If Musicchoice = 1 then  PlaySound"Song70"', -1, 0.2     '0.3 is the volume, just change it to your preferences
If Musicchoice = 2 then  StopSound"Song70"

End Sub

'*********************************************************************************************************
' Backbox & CoinSlots
'*********************************************************************************************************


Dim Coinchoice : Coinchoice = 1 '1 - Spain, 2 - USA

Sub Setupcoinslots()
If Coinchoice = 1 then  
	Coin_Slots_Recel.image = "Coin_Slot_Recel_USA"
End If
If Coinchoice = 2 then  
	Coin_Slots_Recel.image = "Coin_Slot_Recel"
End If

End Sub


' VR PLUNGER ANIMATION
'
' Code needed to animate the plunger. If you pull the plunger it will move in VR.
' IMPORTANT: there are two numeric values in the code that define the postion of the plunger and the 
' range in which it can move. The first numeric value is the actual y position of the plunger primitive
' and the second is the actual y position + 100 to determine the range in which it can move.
'
' You need to to select the VR_Primary_plunger primitive you copied from the
' template and copy the value of the Y position 
' (e.g. 2130) into the code. The value that determines the range of the plunger is always the y 
' position + 100 (e.g. 2230).
'

Sub TimerPlunger_Timer

  If VR_Primary_plunger.Y < 2141.896 then
      VR_Primary_plunger.Y = VR_Primary_plunger.Y + 5
  End If
End Sub

Sub TimerPlunger2_Timer
	VR_Primary_plunger.Y = 2106.896 + (5* Plunger.Position) -20
End Sub

' ***** Beer Bubble Code - Rawd *****
Sub BeerTimer_Timer()

Randomize(21)
BeerBubble1.z = BeerBubble1.z + Rnd(1)*0.5
if BeerBubble1.z > -771 then BeerBubble1.z = -955
BeerBubble2.z = BeerBubble2.z + Rnd(1)*1
if BeerBubble2.z > -768 then BeerBubble2.z = -955
BeerBubble3.z = BeerBubble3.z + Rnd(1)*1
if BeerBubble3.z > -768 then BeerBubble3.z = -955
BeerBubble4.z = BeerBubble4.z + Rnd(1)*0.75
if BeerBubble4.z > -774 then BeerBubble4.z = -955
BeerBubble5.z = BeerBubble5.z + Rnd(1)*1
if BeerBubble5.z > -771 then BeerBubble5.z = -955
BeerBubble6.z = BeerBubble6.z + Rnd(1)*1
if BeerBubble6.z > -774 then BeerBubble6.z = -955
BeerBubble7.z = BeerBubble7.z + Rnd(1)*0.8
if BeerBubble7.z > -768 then BeerBubble7.z = -955
BeerBubble8.z = BeerBubble8.z + Rnd(1)*1
if BeerBubble8.z > -771 then BeerBubble8.z = -955
End Sub




