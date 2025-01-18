'//////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////
'
'
'      .       .           ._._.    _                     .===.
'     |`      |`        ..'\ /`.. |H|        .--.      .:'   `:.
'    //\-...-/|\         |- o -|  |H|`.     /||||\     ||     ||
' ._.'//////,'|||`._.    '`./|\.'` |\\||:. .'||||||`.   `:.   .:'
' ||||||||||||[ ]||||      /_T_\   |:`:.--'||||||||||`--..`=:='... 
'
'
'    ____        ____          ______                 __               ______                           
'   / __ \____  / / /__  _____/ ____/___  ____ ______/ /____  _____   /_  __/_  ___________  ____  ____ 
'  / /_/ / __ \/ / / _ \/ ___/ /   / __ \/ __ `/ ___/ __/ _ \/ ___/    / / / / / / ___/ __ \/ __ \/ __ \
' / _, _/ /_/ / / /  __/ /  / /___/ /_/ / /_/ (__  ) /_/  __/ /       / / / /_/ / /__/ /_/ / /_/ / / / /
'/_/ |_|\____/_/_/\___/_/   \____/\____/\__,_/____/\__/\___/_/       /_/  \__, /\___/\____/\____/_/ /_/ 
'                                                                        /____/     

' Stern, August 2002

' https://www.ipdb.org/machine.cgi?id=4536

'//////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////

' VPW TEAM
' --------
' Gedankekojote97 - Scratch table build, 3D modeling and rendering, Sound effects, Rollercoaster operator
' Apophis - Scripting, physics, Ride mechanic and janitor
' Rawd - VR Room and scripting, Cotton candy and funnel cake maker 
' Testing - Bietekwiet, Eighties8, Benji, JLou, PinStratsDan, Studlygoorite, DGrimmReaper, AstroNasty, TastyWasps, iDigStuff, HauntFreaks, Wylte, ClarkKent, Primetime5k, Somatik, VPW Team


Option Explicit
Randomize
SetLocale 1033

'******************************
'Constants and Global Variables
'******************************

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height
Dim plungerIM, x 

Const BallSize = 50
Const BallMass = 1
Const tnob = 4
Const lob = 0

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const TestVRonDT = false
Dim UseVPMDMD, VRRoom, DesktopMode, VarHidden, UseVPMColoredDMD
DesktopMode = Table1.ShowDT
If RenderingMode = 2 OR TestVRonDT=True Then UseVPMDMD = True Else UseVPMDMD = DesktopMode
If DesktopMode = true then
    UseVPMColoredDMD = true
    VarHidden = 1
Else
    UseVPMColoredDMD = False
    VarHidden = 0
End If

'*************************************
'* VLM Stuff
'*************************************

' VLM  Arrays - Start
' Arrays per baked part
Dim BP_Bumper1: BP_Bumper1=Array(BM_Bumper1, LM_Flashers_L131_Bumper1, LM_Flashers_L59_Bumper1, LM_Spotlights_gi04_Bumper1, LM_Gi_Bumper1)
Dim BP_Bumper2: BP_Bumper2=Array(BM_Bumper2, LM_Flashers_L127_Bumper2, LM_Flashers_L57_Bumper2, LM_Inserts_L79_Bumper2, LM_Gi_Bumper2)
Dim BP_Bumper3: BP_Bumper3=Array(BM_Bumper3, LM_Flashers_L127_Bumper3, LM_Flashers_L130_Bumper3, LM_Flashers_L131_Bumper3, LM_Flashers_L58_Bumper3, LM_Backlight_L76_Bumper3, LM_Inserts_L79_Bumper3, LM_Spotlights_gi04_Bumper3)
Dim BP_Dummy: BP_Dummy=Array(BM_Dummy, LM_Flashers_L131_Dummy, LM_Flashers_L58_Dummy, LM_Backlight_L67_Dummy, LM_Backlight_L68_Dummy, LM_Spotlights_gi04_Dummy, LM_Gi_Dummy)
Dim BP_Gate002: BP_Gate002=Array(BM_Gate002, LM_Spotlights_gi04_Gate002)
Dim BP_Gate003: BP_Gate003=Array(BM_Gate003, LM_Spotlights_gi04_Gate003, LM_Gi_Gate003)
Dim BP_Gate004: BP_Gate004=Array(BM_Gate004, LM_Gi_Gate004)
Dim BP_Gate005: BP_Gate005=Array(BM_Gate005, LM_Spotlights_gi04_Gate005)
Dim BP_Gate1: BP_Gate1=Array(BM_Gate1, LM_Flashers_L122_Gate1, LM_Gi_Gate1)
Dim BP_Gate3: BP_Gate3=Array(BM_Gate3, LM_Flashers_L131_Gate3, LM_Gi_Gate3)
Dim BP_Gate5: BP_Gate5=Array(BM_Gate5, LM_Flashers_L123_Gate5, LM_Flashers_L129_Gate5, LM_Gi_Gate5)
Dim BP_Gate6: BP_Gate6=Array(BM_Gate6, LM_Spotlights_gi04_Gate6, LM_Gi_Gate6)
Dim BP_Gate7: BP_Gate7=Array(BM_Gate7)
Dim BP_Gate8: BP_Gate8=Array(BM_Gate8, LM_Gi_Gate8)
Dim BP_GhostTarget: BP_GhostTarget=Array(BM_GhostTarget, LM_Flashers_L129_GhostTarget, LM_Flashers_L130_GhostTarget, LM_Flashers_L131_GhostTarget, LM_Inserts_L50_GhostTarget, LM_Inserts_L51_GhostTarget, LM_Inserts_L54_GhostTarget, LM_Inserts_L55_GhostTarget, LM_Inserts_L56_GhostTarget, LM_Spotlights_gi04_GhostTarget, LM_Gi_GhostTarget)
Dim BP_Layer1: BP_Layer1=Array(BM_Layer1, LM_Inserts_L11_Layer1, LM_Flashers_L121_Layer1, LM_Flashers_L122_Layer1, LM_Flashers_L123_Layer1, LM_Flashers_L127_Layer1, LM_Flashers_L129_Layer1, LM_Flashers_L130_Layer1, LM_Flashers_L131_Layer1, LM_Flashers_L132_Layer1, LM_Inserts_L18_Layer1, LM_Inserts_L19_Layer1, LM_Inserts_L35_Layer1, LM_Inserts_L36_Layer1, LM_Flashers_L48_Layer1, LM_Inserts_L51_Layer1, LM_Flashers_L57_Layer1, LM_Flashers_L58_Layer1, LM_Flashers_L59_Layer1, LM_Backlight_L65_Layer1, LM_Backlight_L66_Layer1, LM_Backlight_L69_Layer1, LM_Inserts_L70_Layer1, LM_Inserts_L71_Layer1, LM_Inserts_L72_Layer1, LM_Backlight_L73_Layer1, LM_Backlight_L74_Layer1, LM_Backlight_L75_Layer1, LM_Backlight_L76_Layer1, LM_Backlight_L77_Layer1, LM_Inserts_L78_Layer1, LM_Inserts_L79_Layer1, LM_Inserts_L9_Layer1, LM_Spotlights_gi04_Layer1, LM_Gi1_gi01_Layer1, LM_Gi1_gi02_Layer1, LM_Gi1_gi03_Layer1, LM_Gi_Layer1, LM_Gi1_gi21_Layer1, LM_Gi1_gi22_Layer1, LM_Gi1_gi23_Layer1)
Dim BP_Layer2: BP_Layer2=Array(BM_Layer2, LM_Flashers_L121_Layer2, LM_Flashers_L122_Layer2, LM_Flashers_L127_Layer2, LM_Flashers_L130_Layer2, LM_Flashers_L131_Layer2, LM_Flashers_L58_Layer2, LM_Backlight_L67_Layer2, LM_Backlight_L68_Layer2, LM_Backlight_L69_Layer2, LM_Inserts_L71_Layer2, LM_Inserts_L72_Layer2, LM_Backlight_L73_Layer2, LM_Backlight_L74_Layer2, LM_Backlight_L75_Layer2, LM_Backlight_L76_Layer2, LM_Inserts_L79_Layer2, LM_Inserts_L9_Layer2, LM_Spotlights_gi04_Layer2, LM_Gi1_gi01_Layer2, LM_Gi1_gi02_Layer2, LM_Gi1_gi03_Layer2, LM_Gi_Layer2)
Dim BP_LeftFlipper: BP_LeftFlipper=Array(BM_LeftFlipper, LM_Flashers_L121_LeftFlipper, LM_Inserts_L1_LeftFlipper, LM_Inserts_L23_LeftFlipper, LM_Inserts_L2_LeftFlipper, LM_Gi1_gi01_LeftFlipper, LM_Gi1_gi02_LeftFlipper, LM_Gi1_gi03_LeftFlipper, LM_Gi_LeftFlipper)
Dim BP_LeftFlipper1: BP_LeftFlipper1=Array(BM_LeftFlipper1, LM_Flashers_L121_LeftFlipper1, LM_Flashers_L123_LeftFlipper1, LM_Flashers_L132_LeftFlipper1, LM_Gi_LeftFlipper1)
Dim BP_LeftFlipper1U: BP_LeftFlipper1U=Array(BM_LeftFlipper1U, LM_Flashers_L121_LeftFlipper1U, LM_Flashers_L132_LeftFlipper1U, LM_Inserts_L17_LeftFlipper1U, LM_Inserts_L18_LeftFlipper1U, LM_Gi_LeftFlipper1U)
Dim BP_LeftFlipperU: BP_LeftFlipperU=Array(BM_LeftFlipperU, LM_Flashers_L121_LeftFlipperU, LM_Flashers_L122_LeftFlipperU, LM_Inserts_L1_LeftFlipperU, LM_Inserts_L23_LeftFlipperU, LM_Inserts_L2_LeftFlipperU, LM_Inserts_L3_LeftFlipperU, LM_Gi1_gi01_LeftFlipperU, LM_Gi1_gi02_LeftFlipperU, LM_Gi1_gi03_LeftFlipperU, LM_Gi_LeftFlipperU, LM_Gi1_gi21_LeftFlipperU, LM_Gi1_gi23_LeftFlipperU)
Dim BP_LeftSling1: BP_LeftSling1=Array(BM_LeftSling1, LM_Inserts_L10_LeftSling1, LM_Flashers_L121_LeftSling1, LM_Flashers_L122_LeftSling1, LM_Inserts_L1_LeftSling1, LM_Gi1_gi01_LeftSling1, LM_Gi1_gi02_LeftSling1, LM_Gi1_gi03_LeftSling1, LM_Gi_LeftSling1)
Dim BP_LeftSling2: BP_LeftSling2=Array(BM_LeftSling2, LM_Inserts_L10_LeftSling2, LM_Flashers_L121_LeftSling2, LM_Flashers_L122_LeftSling2, LM_Gi1_gi01_LeftSling2, LM_Gi1_gi02_LeftSling2, LM_Gi1_gi03_LeftSling2, LM_Gi_LeftSling2)
Dim BP_LeftSling3: BP_LeftSling3=Array(BM_LeftSling3, LM_Inserts_L10_LeftSling3, LM_Flashers_L121_LeftSling3, LM_Flashers_L122_LeftSling3, LM_Gi1_gi01_LeftSling3, LM_Gi1_gi02_LeftSling3, LM_Gi1_gi03_LeftSling3, LM_Gi_LeftSling3)
Dim BP_LeftSling4: BP_LeftSling4=Array(BM_LeftSling4, LM_Inserts_L10_LeftSling4, LM_Flashers_L121_LeftSling4, LM_Flashers_L122_LeftSling4, LM_Gi1_gi01_LeftSling4, LM_Gi1_gi02_LeftSling4, LM_Gi1_gi03_LeftSling4, LM_Gi_LeftSling4)
Dim BP_Lemk: BP_Lemk=Array(BM_Lemk, LM_Inserts_L10_Lemk, LM_Flashers_L121_Lemk, LM_Flashers_L122_Lemk, LM_Gi1_gi01_Lemk, LM_Gi1_gi02_Lemk, LM_Gi1_gi03_Lemk)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_DMD_DMX1_Parts, LM_DMD_DMX10_Parts, LM_DMD_DMX100_Parts, LM_DMD_DMX101_Parts, LM_DMD_DMX102_Parts, LM_DMD_DMX103_Parts, LM_DMD_DMX104_Parts, LM_DMD_DMX105_Parts, LM_DMD_DMX11_Parts, LM_DMD_DMX12_Parts, LM_DMD_DMX13_Parts, LM_DMD_DMX14_Parts, LM_DMD_DMX15_Parts, LM_DMD_DMX16_Parts, LM_DMD_DMX17_Parts, LM_DMD_DMX18_Parts, LM_DMD_DMX19_Parts, LM_DMD_DMX2_Parts, LM_DMD_DMX20_Parts, LM_DMD_DMX21_Parts, LM_DMD_DMX22_Parts, LM_DMD_DMX23_Parts, LM_DMD_DMX24_Parts, LM_DMD_DMX25_Parts, LM_DMD_DMX26_Parts, LM_DMD_DMX27_Parts, LM_DMD_DMX28_Parts, LM_DMD_DMX29_Parts, LM_DMD_DMX3_Parts, LM_DMD_DMX30_Parts, LM_DMD_DMX31_Parts, LM_DMD_DMX32_Parts, LM_DMD_DMX33_Parts, LM_DMD_DMX34_Parts, LM_DMD_DMX35_Parts, LM_DMD_DMX36_Parts, LM_DMD_DMX37_Parts, LM_DMD_DMX38_Parts, LM_DMD_DMX39_Parts, LM_DMD_DMX4_Parts, LM_DMD_DMX40_Parts, LM_DMD_DMX41_Parts, LM_DMD_DMX42_Parts, LM_DMD_DMX43_Parts, LM_DMD_DMX44_Parts, LM_DMD_DMX45_Parts, LM_DMD_DMX46_Parts, LM_DMD_DMX47_Parts, LM_DMD_DMX48_Parts, _
	LM_DMD_DMX49_Parts, LM_DMD_DMX5_Parts, LM_DMD_DMX50_Parts, LM_DMD_DMX51_Parts, LM_DMD_DMX52_Parts, LM_DMD_DMX53_Parts, LM_DMD_DMX54_Parts, LM_DMD_DMX55_Parts, LM_DMD_DMX56_Parts, LM_DMD_DMX57_Parts, LM_DMD_DMX58_Parts, LM_DMD_DMX59_Parts, LM_DMD_DMX6_Parts, LM_DMD_DMX60_Parts, LM_DMD_DMX61_Parts, LM_DMD_DMX62_Parts, LM_DMD_DMX63_Parts, LM_DMD_DMX64_Parts, LM_DMD_DMX65_Parts, LM_DMD_DMX66_Parts, LM_DMD_DMX67_Parts, LM_DMD_DMX68_Parts, LM_DMD_DMX69_Parts, LM_DMD_DMX7_Parts, LM_DMD_DMX70_Parts, LM_DMD_DMX71_Parts, LM_DMD_DMX72_Parts, LM_DMD_DMX73_Parts, LM_DMD_DMX74_Parts, LM_DMD_DMX75_Parts, LM_DMD_DMX76_Parts, LM_DMD_DMX77_Parts, LM_DMD_DMX78_Parts, LM_DMD_DMX79_Parts, LM_DMD_DMX8_Parts, LM_DMD_DMX80_Parts, LM_DMD_DMX81_Parts, LM_DMD_DMX82_Parts, LM_DMD_DMX83_Parts, LM_DMD_DMX84_Parts, LM_DMD_DMX85_Parts, LM_DMD_DMX86_Parts, LM_DMD_DMX87_Parts, LM_DMD_DMX88_Parts, LM_DMD_DMX89_Parts, LM_DMD_DMX9_Parts, LM_DMD_DMX90_Parts, LM_DMD_DMX91_Parts, LM_DMD_DMX92_Parts, LM_DMD_DMX93_Parts, LM_DMD_DMX94_Parts, _
	LM_DMD_DMX95_Parts, LM_DMD_DMX96_Parts, LM_DMD_DMX97_Parts, LM_DMD_DMX98_Parts, LM_DMD_DMX99_Parts, LM_Inserts_L10_Parts, LM_Flashers_L121_Parts, LM_Flashers_L122_Parts, LM_Flashers_L123_Parts, LM_Flashers_L127_Parts, LM_Flashers_L129_Parts, LM_Flashers_L130_Parts, LM_Flashers_L131_Parts, LM_Flashers_L132_Parts, LM_Inserts_L14_Parts, LM_Inserts_L17_Parts, LM_Inserts_L18_Parts, LM_Inserts_L1_Parts, LM_Inserts_L24_Parts, LM_Inserts_L26_Parts, LM_Inserts_L27_Parts, LM_Inserts_L28_Parts, LM_Inserts_L29_Parts, LM_Inserts_L30_Parts, LM_Inserts_L33_Parts, LM_Inserts_L35_Parts, LM_Inserts_L36_Parts, LM_Inserts_L37_Parts, LM_Inserts_L38_Parts, LM_Inserts_L39_Parts, LM_Inserts_L3_Parts, LM_Inserts_L41_Parts, LM_Inserts_L42_Parts, LM_Inserts_L43_Parts, LM_Inserts_L44_Parts, LM_Inserts_L45_Parts, LM_Inserts_L46_Parts, LM_Inserts_L47_Parts, LM_Flashers_L48_Parts, LM_Inserts_L49_Parts, LM_Inserts_L4_Parts, LM_Inserts_L50_Parts, LM_Inserts_L51_Parts, LM_Inserts_L52_Parts, LM_Inserts_L53_Parts, LM_Inserts_L54_Parts, _
	LM_Inserts_L55_Parts, LM_Inserts_L56_Parts, LM_Flashers_L57_Parts, LM_Flashers_L58_Parts, LM_Flashers_L59_Parts, LM_Inserts_L60_Parts, LM_Inserts_L62_Parts, LM_Backlight_L65_Parts, LM_Backlight_L66_Parts, LM_Backlight_L67_Parts, LM_Backlight_L68_Parts, LM_Backlight_L69_Parts, LM_Inserts_L70_Parts, LM_Inserts_L71_Parts, LM_Inserts_L72_Parts, LM_Backlight_L73_Parts, LM_Backlight_L74_Parts, LM_Backlight_L75_Parts, LM_Backlight_L76_Parts, LM_Backlight_L77_Parts, LM_Inserts_L78_Parts, LM_Inserts_L79_Parts, LM_Inserts_L8_Parts, LM_Inserts_L9_Parts, LM_Opto_Parts, LM_Spotlights_gi04_Parts, LM_Gi1_gi01_Parts, LM_Gi1_gi02_Parts, LM_Gi1_gi03_Parts, LM_Gi_Parts, LM_Gi1_gi21_Parts, LM_Gi1_gi22_Parts, LM_Gi1_gi23_Parts)
Dim BP_Parts2: BP_Parts2=Array(BM_Parts2, LM_Flashers_L121_Parts2, LM_Gi1_gi01_Parts2, LM_Gi1_gi02_Parts2, LM_Gi1_gi03_Parts2, LM_Gi_Parts2)
Dim BP_PinCab_Rails: BP_PinCab_Rails=Array(BM_PinCab_Rails, LM_Flashers_L121_PinCab_Rails, LM_Flashers_L122_PinCab_Rails)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_Inserts_L10_Playfield, LM_Inserts_L11_Playfield, LM_Flashers_L121_Playfield, LM_Flashers_L122_Playfield, LM_Flashers_L123_Playfield, LM_Flashers_L127_Playfield, LM_Flashers_L129_Playfield, LM_Inserts_L12_Playfield, LM_Flashers_L130_Playfield, LM_Flashers_L131_Playfield, LM_Flashers_L132_Playfield, LM_Inserts_L13_Playfield, LM_Inserts_L14_Playfield, LM_Inserts_L15_Playfield, LM_Inserts_L16_Playfield, LM_Inserts_L17_Playfield, LM_Inserts_L18_Playfield, LM_Inserts_L19_Playfield, LM_Inserts_L1_Playfield, LM_Inserts_L20_Playfield, LM_Inserts_L21_Playfield, LM_Inserts_L22_Playfield, LM_Inserts_L23_Playfield, LM_Inserts_L24_Playfield, LM_Inserts_L25_Playfield, LM_Inserts_L26_Playfield, LM_Inserts_L27_Playfield, LM_Inserts_L28_Playfield, LM_Inserts_L29_Playfield, LM_Inserts_L2_Playfield, LM_Inserts_L30_Playfield, LM_Inserts_L31_Playfield, LM_Inserts_L32_Playfield, LM_Inserts_L33_Playfield, LM_Inserts_L34_Playfield, LM_Inserts_L35_Playfield, _
	LM_Inserts_L36_Playfield, LM_Inserts_L37_Playfield, LM_Inserts_L38_Playfield, LM_Inserts_L39_Playfield, LM_Inserts_L3_Playfield, LM_Inserts_L40_Playfield, LM_Inserts_L41_Playfield, LM_Inserts_L42_Playfield, LM_Inserts_L43_Playfield, LM_Inserts_L44_Playfield, LM_Inserts_L45_Playfield, LM_Inserts_L46_Playfield, LM_Inserts_L47_Playfield, LM_Flashers_L48_Playfield, LM_Inserts_L49_Playfield, LM_Inserts_L4_Playfield, LM_Inserts_L50_Playfield, LM_Inserts_L51_Playfield, LM_Inserts_L52_Playfield, LM_Inserts_L53_Playfield, LM_Inserts_L54_Playfield, LM_Inserts_L55_Playfield, LM_Inserts_L56_Playfield, LM_Flashers_L57_Playfield, LM_Flashers_L58_Playfield, LM_Flashers_L59_Playfield, LM_Inserts_L5_Playfield, LM_Inserts_L60_Playfield, LM_Inserts_L61_Playfield, LM_Inserts_L62_Playfield, LM_Inserts_L63_Playfield, LM_Backlight_L66_Playfield, LM_Backlight_L67_Playfield, LM_Backlight_L68_Playfield, LM_Backlight_L69_Playfield, LM_Inserts_L6_Playfield, LM_Inserts_L70_Playfield, LM_Inserts_L71_Playfield, LM_Inserts_L72_Playfield, _
	LM_Backlight_L73_Playfield, LM_Backlight_L75_Playfield, LM_Backlight_L76_Playfield, LM_Inserts_L78_Playfield, LM_Inserts_L79_Playfield, LM_Inserts_L7_Playfield, LM_Inserts_L8_Playfield, LM_Inserts_L9_Playfield, LM_Opto_Playfield, LM_Spotlights_gi04_Playfield, LM_Gi1_gi01_Playfield, LM_Gi1_gi02_Playfield, LM_Gi1_gi03_Playfield, LM_Gi_Playfield, LM_Gi1_gi21_Playfield, LM_Gi1_gi22_Playfield, LM_Gi1_gi23_Playfield)
Dim BP_Post: BP_Post=Array(BM_Post, LM_Inserts_L70_Post, LM_Gi_Post)
Dim BP_Remk: BP_Remk=Array(BM_Remk, LM_Flashers_L121_Remk, LM_Flashers_L122_Remk, LM_Gi1_gi21_Remk, LM_Gi1_gi22_Remk, LM_Gi1_gi23_Remk)
Dim BP_RightFlipper: BP_RightFlipper=Array(BM_RightFlipper, LM_Flashers_L122_RightFlipper, LM_Inserts_L23_RightFlipper, LM_Inserts_L4_RightFlipper, LM_Inserts_L5_RightFlipper, LM_Gi1_gi21_RightFlipper, LM_Gi1_gi22_RightFlipper, LM_Gi1_gi23_RightFlipper)
Dim BP_RightFlipper1: BP_RightFlipper1=Array(BM_RightFlipper1, LM_Flashers_L129_RightFlipper1, LM_Flashers_L130_RightFlipper1, LM_Flashers_L131_RightFlipper1, LM_Flashers_L48_RightFlipper1, LM_Spotlights_gi04_RF1, LM_Gi_RightFlipper1)
Dim BP_RightFlipper1U: BP_RightFlipper1U=Array(BM_RightFlipper1U, LM_Flashers_L123_RightFlipper1U, LM_Flashers_L129_RightFlipper1U, LM_Flashers_L130_RightFlipper1U, LM_Flashers_L131_RightFlipper1U, LM_Inserts_L41_RightFlipper1U, LM_Inserts_L42_RightFlipper1U, LM_Inserts_L44_RightFlipper1U, LM_Inserts_L45_RightFlipper1U, LM_Inserts_L47_RightFlipper1U, LM_Flashers_L48_RightFlipper1U, LM_Spotlights_gi04_RF1U, LM_Gi_RightFlipper1U)
Dim BP_RightFlipperU: BP_RightFlipperU=Array(BM_RightFlipperU, LM_Flashers_L121_RightFlipperU, LM_Flashers_L122_RightFlipperU, LM_Inserts_L23_RightFlipperU, LM_Inserts_L3_RightFlipperU, LM_Inserts_L4_RightFlipperU, LM_Inserts_L5_RightFlipperU, LM_Inserts_L6_RightFlipperU, LM_Gi1_gi01_RightFlipperU, LM_Gi_RightFlipperU, LM_Gi1_gi21_RightFlipperU, LM_Gi1_gi22_RightFlipperU, LM_Gi1_gi23_RightFlipperU)
Dim BP_RightSling1: BP_RightSling1=Array(BM_RightSling1, LM_Flashers_L121_RightSling1, LM_Flashers_L122_RightSling1, LM_Inserts_L5_RightSling1, LM_Inserts_L6_RightSling1, LM_Gi_RightSling1, LM_Gi1_gi21_RightSling1, LM_Gi1_gi22_RightSling1, LM_Gi1_gi23_RightSling1)
Dim BP_RightSling2: BP_RightSling2=Array(BM_RightSling2, LM_Flashers_L121_RightSling2, LM_Flashers_L122_RightSling2, LM_Inserts_L6_RightSling2, LM_Gi_RightSling2, LM_Gi1_gi21_RightSling2, LM_Gi1_gi22_RightSling2, LM_Gi1_gi23_RightSling2)
Dim BP_RightSling3: BP_RightSling3=Array(BM_RightSling3, LM_Flashers_L121_RightSling3, LM_Flashers_L122_RightSling3, LM_Inserts_L6_RightSling3, LM_Gi_RightSling3, LM_Gi1_gi21_RightSling3, LM_Gi1_gi22_RightSling3, LM_Gi1_gi23_RightSling3)
Dim BP_RightSling4: BP_RightSling4=Array(BM_RightSling4, LM_Flashers_L121_RightSling4, LM_Flashers_L122_RightSling4, LM_Inserts_L6_RightSling4, LM_Gi_RightSling4, LM_Gi1_gi21_RightSling4, LM_Gi1_gi22_RightSling4, LM_Gi1_gi23_RightSling4)
Dim BP_ScrambledEgg: BP_ScrambledEgg=Array(BM_ScrambledEgg, LM_Flashers_L121_ScrambledEgg, LM_Flashers_L132_ScrambledEgg, LM_Flashers_L59_ScrambledEgg, LM_Gi_ScrambledEgg)
Dim BP_sw17: BP_sw17=Array(BM_sw17, LM_Flashers_L129_sw17, LM_Flashers_L130_sw17, LM_Flashers_L131_sw17, LM_Flashers_L58_sw17, LM_Inserts_L61_sw17, LM_Spotlights_gi04_sw17)
Dim BP_sw18: BP_sw18=Array(BM_sw18, LM_Flashers_L123_sw18, LM_Flashers_L130_sw18, LM_Flashers_L131_sw18, LM_Inserts_L61_sw18, LM_Spotlights_gi04_sw18, LM_Gi_sw18)
Dim BP_sw19: BP_sw19=Array(BM_sw19, LM_Flashers_L123_sw19, LM_Flashers_L131_sw19, LM_Spotlights_gi04_sw19)
Dim BP_sw21: BP_sw21=Array(BM_sw21)
Dim BP_sw22: BP_sw22=Array(BM_sw22, LM_Flashers_L121_sw22, LM_Flashers_L132_sw22, LM_Inserts_L17_sw22, LM_Inserts_L18_sw22, LM_Inserts_L20_sw22, LM_Inserts_L33_sw22, LM_Gi_sw22)
Dim BP_sw25: BP_sw25=Array(BM_sw25, LM_Gi_sw25)
Dim BP_sw26: BP_sw26=Array(BM_sw26, LM_Gi_sw26)
Dim BP_sw27: BP_sw27=Array(BM_sw27, LM_Spotlights_gi04_sw27, LM_Gi_sw27)
Dim BP_sw28: BP_sw28=Array(BM_sw28, LM_Backlight_L66_sw28, LM_Spotlights_gi04_sw28, LM_Gi_sw28)
Dim BP_sw29: BP_sw29=Array(BM_sw29, LM_Inserts_L78_sw29, LM_Spotlights_gi04_sw29)
Dim BP_sw30: BP_sw30=Array(BM_sw30, LM_Flashers_L127_sw30, LM_Flashers_L129_sw30, LM_Flashers_L130_sw30, LM_Flashers_L131_sw30, LM_Flashers_L58_sw30, LM_Flashers_L59_sw30, LM_Inserts_L60_sw30, LM_Backlight_L68_sw30, LM_Inserts_L79_sw30, LM_Spotlights_gi04_sw30)
Dim BP_sw31: BP_sw31=Array(BM_sw31, LM_Flashers_L129_sw31, LM_Flashers_L130_sw31, LM_Flashers_L131_sw31, LM_Flashers_L48_sw31, LM_Flashers_L58_sw31, LM_Inserts_L60_sw31, LM_Inserts_L79_sw31, LM_Spotlights_gi04_sw31)
Dim BP_sw32: BP_sw32=Array(BM_sw32, LM_Flashers_L129_sw32, LM_Flashers_L130_sw32, LM_Flashers_L131_sw32, LM_Flashers_L48_sw32, LM_Inserts_L51_sw32, LM_Inserts_L52_sw32, LM_Flashers_L58_sw32, LM_Inserts_L60_sw32, LM_Spotlights_gi04_sw32, LM_Gi_sw32)
Dim BP_sw37: BP_sw37=Array(BM_sw37, LM_Flashers_L48_sw37)
Dim BP_sw39: BP_sw39=Array(BM_sw39, LM_Flashers_L129_sw39, LM_Flashers_L130_sw39, LM_Flashers_L131_sw39, LM_Inserts_L25_sw39, LM_Inserts_L26_sw39, LM_Inserts_L27_sw39, LM_Inserts_L28_sw39, LM_Inserts_L44_sw39, LM_Inserts_L45_sw39, LM_Inserts_L46_sw39, LM_Inserts_L47_sw39, LM_Flashers_L48_sw39, LM_Opto_sw39, LM_Spotlights_gi04_sw39, LM_Gi1_gi02_sw39, LM_Gi1_gi03_sw39, LM_Gi_sw39)
Dim BP_sw40: BP_sw40=Array(BM_sw40, LM_Flashers_L129_sw40, LM_Flashers_L130_sw40, LM_Flashers_L131_sw40, LM_Inserts_L42_sw40, LM_Inserts_L43_sw40, LM_Flashers_L48_sw40, LM_Inserts_L54_sw40, LM_Inserts_L55_sw40, LM_Inserts_L56_sw40, LM_Spotlights_gi04_sw40, LM_Gi_sw40)
Dim BP_sw42: BP_sw42=Array(BM_sw42, LM_Flashers_L121_sw42, LM_Gi_sw42)
Dim BP_sw43: BP_sw43=Array(BM_sw43, LM_Flashers_L121_sw43)
Dim BP_sw44: BP_sw44=Array(BM_sw44, LM_Flashers_L122_sw44, LM_Flashers_L123_sw44, LM_Inserts_L28_sw44, LM_Inserts_L29_sw44, LM_Inserts_L30_sw44, LM_Gi_sw44)
Dim BP_sw45: BP_sw45=Array(BM_sw45, LM_Flashers_L123_sw45, LM_Inserts_L29_sw45, LM_Inserts_L30_sw45, LM_Inserts_L31_sw45, LM_Gi_sw45)
Dim BP_sw46: BP_sw46=Array(BM_sw46, LM_Flashers_L123_sw46, LM_Inserts_L30_sw46, LM_Inserts_L31_sw46, LM_Inserts_L32_sw46, LM_Gi_sw46)
Dim BP_sw57: BP_sw57=Array(BM_sw57, LM_Flashers_L121_sw57, LM_Gi1_gi02_sw57, LM_Gi1_gi03_sw57)
Dim BP_sw58: BP_sw58=Array(BM_sw58, LM_Flashers_L121_sw58, LM_Flashers_L122_sw58, LM_Inserts_L9_sw58, LM_Gi1_gi01_sw58, LM_Gi1_gi02_sw58, LM_Gi1_gi03_sw58, LM_Gi_sw58)
Dim BP_sw60: BP_sw60=Array(BM_sw60, LM_Flashers_L122_sw60, LM_Gi_sw60, LM_Gi1_gi21_sw60, LM_Gi1_gi22_sw60)
Dim BP_sw61: BP_sw61=Array(BM_sw61, LM_Flashers_L121_sw61, LM_Flashers_L122_sw61, LM_Inserts_L8_sw61, LM_Gi_sw61, LM_Gi1_gi21_sw61, LM_Gi1_gi22_sw61, LM_Gi1_gi23_sw61)
Dim BP_swPlunger: BP_swPlunger=Array(BM_swPlunger, LM_Flashers_L122_swPlunger)
' Arrays per lighting scenario
'Dim BL_Backlight_L65: BL_Backlight_L65=Array(LM_Backlight_L65_Layer1, LM_Backlight_L65_Parts)
'Dim BL_Backlight_L66: BL_Backlight_L66=Array(LM_Backlight_L66_Layer1, LM_Backlight_L66_Parts, LM_Backlight_L66_Playfield, LM_Backlight_L66_sw28)
'Dim BL_Backlight_L67: BL_Backlight_L67=Array(LM_Backlight_L67_Dummy, LM_Backlight_L67_Layer2, LM_Backlight_L67_Parts, LM_Backlight_L67_Playfield)
'Dim BL_Backlight_L68: BL_Backlight_L68=Array(LM_Backlight_L68_Dummy, LM_Backlight_L68_Layer2, LM_Backlight_L68_Parts, LM_Backlight_L68_Playfield, LM_Backlight_L68_sw30)
'Dim BL_Backlight_L69: BL_Backlight_L69=Array(LM_Backlight_L69_Layer1, LM_Backlight_L69_Layer2, LM_Backlight_L69_Parts, LM_Backlight_L69_Playfield)
'Dim BL_Backlight_L73: BL_Backlight_L73=Array(LM_Backlight_L73_Layer1, LM_Backlight_L73_Layer2, LM_Backlight_L73_Parts, LM_Backlight_L73_Playfield)
'Dim BL_Backlight_L74: BL_Backlight_L74=Array(LM_Backlight_L74_Layer1, LM_Backlight_L74_Layer2, LM_Backlight_L74_Parts)
'Dim BL_Backlight_L75: BL_Backlight_L75=Array(LM_Backlight_L75_Layer1, LM_Backlight_L75_Layer2, LM_Backlight_L75_Parts, LM_Backlight_L75_Playfield)
'Dim BL_Backlight_L76: BL_Backlight_L76=Array(LM_Backlight_L76_Bumper3, LM_Backlight_L76_Layer1, LM_Backlight_L76_Layer2, LM_Backlight_L76_Parts, LM_Backlight_L76_Playfield)
'Dim BL_Backlight_L77: BL_Backlight_L77=Array(LM_Backlight_L77_Layer1, LM_Backlight_L77_Parts)
'Dim BL_DMD_DMX1: BL_DMD_DMX1=Array(LM_DMD_DMX1_Parts)
'Dim BL_DMD_DMX10: BL_DMD_DMX10=Array(LM_DMD_DMX10_Parts)
'Dim BL_DMD_DMX100: BL_DMD_DMX100=Array(LM_DMD_DMX100_Parts)
'Dim BL_DMD_DMX101: BL_DMD_DMX101=Array(LM_DMD_DMX101_Parts)
'Dim BL_DMD_DMX102: BL_DMD_DMX102=Array(LM_DMD_DMX102_Parts)
'Dim BL_DMD_DMX103: BL_DMD_DMX103=Array(LM_DMD_DMX103_Parts)
'Dim BL_DMD_DMX104: BL_DMD_DMX104=Array(LM_DMD_DMX104_Parts)
'Dim BL_DMD_DMX105: BL_DMD_DMX105=Array(LM_DMD_DMX105_Parts)
'Dim BL_DMD_DMX11: BL_DMD_DMX11=Array(LM_DMD_DMX11_Parts)
'Dim BL_DMD_DMX12: BL_DMD_DMX12=Array(LM_DMD_DMX12_Parts)
'Dim BL_DMD_DMX13: BL_DMD_DMX13=Array(LM_DMD_DMX13_Parts)
'Dim BL_DMD_DMX14: BL_DMD_DMX14=Array(LM_DMD_DMX14_Parts)
'Dim BL_DMD_DMX15: BL_DMD_DMX15=Array(LM_DMD_DMX15_Parts)
'Dim BL_DMD_DMX16: BL_DMD_DMX16=Array(LM_DMD_DMX16_Parts)
'Dim BL_DMD_DMX17: BL_DMD_DMX17=Array(LM_DMD_DMX17_Parts)
'Dim BL_DMD_DMX18: BL_DMD_DMX18=Array(LM_DMD_DMX18_Parts)
'Dim BL_DMD_DMX19: BL_DMD_DMX19=Array(LM_DMD_DMX19_Parts)
'Dim BL_DMD_DMX2: BL_DMD_DMX2=Array(LM_DMD_DMX2_Parts)
'Dim BL_DMD_DMX20: BL_DMD_DMX20=Array(LM_DMD_DMX20_Parts)
'Dim BL_DMD_DMX21: BL_DMD_DMX21=Array(LM_DMD_DMX21_Parts)
'Dim BL_DMD_DMX22: BL_DMD_DMX22=Array(LM_DMD_DMX22_Parts)
'Dim BL_DMD_DMX23: BL_DMD_DMX23=Array(LM_DMD_DMX23_Parts)
'Dim BL_DMD_DMX24: BL_DMD_DMX24=Array(LM_DMD_DMX24_Parts)
'Dim BL_DMD_DMX25: BL_DMD_DMX25=Array(LM_DMD_DMX25_Parts)
'Dim BL_DMD_DMX26: BL_DMD_DMX26=Array(LM_DMD_DMX26_Parts)
'Dim BL_DMD_DMX27: BL_DMD_DMX27=Array(LM_DMD_DMX27_Parts)
'Dim BL_DMD_DMX28: BL_DMD_DMX28=Array(LM_DMD_DMX28_Parts)
'Dim BL_DMD_DMX29: BL_DMD_DMX29=Array(LM_DMD_DMX29_Parts)
'Dim BL_DMD_DMX3: BL_DMD_DMX3=Array(LM_DMD_DMX3_Parts)
'Dim BL_DMD_DMX30: BL_DMD_DMX30=Array(LM_DMD_DMX30_Parts)
'Dim BL_DMD_DMX31: BL_DMD_DMX31=Array(LM_DMD_DMX31_Parts)
'Dim BL_DMD_DMX32: BL_DMD_DMX32=Array(LM_DMD_DMX32_Parts)
'Dim BL_DMD_DMX33: BL_DMD_DMX33=Array(LM_DMD_DMX33_Parts)
'Dim BL_DMD_DMX34: BL_DMD_DMX34=Array(LM_DMD_DMX34_Parts)
'Dim BL_DMD_DMX35: BL_DMD_DMX35=Array(LM_DMD_DMX35_Parts)
'Dim BL_DMD_DMX36: BL_DMD_DMX36=Array(LM_DMD_DMX36_Parts)
'Dim BL_DMD_DMX37: BL_DMD_DMX37=Array(LM_DMD_DMX37_Parts)
'Dim BL_DMD_DMX38: BL_DMD_DMX38=Array(LM_DMD_DMX38_Parts)
'Dim BL_DMD_DMX39: BL_DMD_DMX39=Array(LM_DMD_DMX39_Parts)
'Dim BL_DMD_DMX4: BL_DMD_DMX4=Array(LM_DMD_DMX4_Parts)
'Dim BL_DMD_DMX40: BL_DMD_DMX40=Array(LM_DMD_DMX40_Parts)
'Dim BL_DMD_DMX41: BL_DMD_DMX41=Array(LM_DMD_DMX41_Parts)
'Dim BL_DMD_DMX42: BL_DMD_DMX42=Array(LM_DMD_DMX42_Parts)
'Dim BL_DMD_DMX43: BL_DMD_DMX43=Array(LM_DMD_DMX43_Parts)
'Dim BL_DMD_DMX44: BL_DMD_DMX44=Array(LM_DMD_DMX44_Parts)
'Dim BL_DMD_DMX45: BL_DMD_DMX45=Array(LM_DMD_DMX45_Parts)
'Dim BL_DMD_DMX46: BL_DMD_DMX46=Array(LM_DMD_DMX46_Parts)
'Dim BL_DMD_DMX47: BL_DMD_DMX47=Array(LM_DMD_DMX47_Parts)
'Dim BL_DMD_DMX48: BL_DMD_DMX48=Array(LM_DMD_DMX48_Parts)
'Dim BL_DMD_DMX49: BL_DMD_DMX49=Array(LM_DMD_DMX49_Parts)
'Dim BL_DMD_DMX5: BL_DMD_DMX5=Array(LM_DMD_DMX5_Parts)
'Dim BL_DMD_DMX50: BL_DMD_DMX50=Array(LM_DMD_DMX50_Parts)
'Dim BL_DMD_DMX51: BL_DMD_DMX51=Array(LM_DMD_DMX51_Parts)
'Dim BL_DMD_DMX52: BL_DMD_DMX52=Array(LM_DMD_DMX52_Parts)
'Dim BL_DMD_DMX53: BL_DMD_DMX53=Array(LM_DMD_DMX53_Parts)
'Dim BL_DMD_DMX54: BL_DMD_DMX54=Array(LM_DMD_DMX54_Parts)
'Dim BL_DMD_DMX55: BL_DMD_DMX55=Array(LM_DMD_DMX55_Parts)
'Dim BL_DMD_DMX56: BL_DMD_DMX56=Array(LM_DMD_DMX56_Parts)
'Dim BL_DMD_DMX57: BL_DMD_DMX57=Array(LM_DMD_DMX57_Parts)
'Dim BL_DMD_DMX58: BL_DMD_DMX58=Array(LM_DMD_DMX58_Parts)
'Dim BL_DMD_DMX59: BL_DMD_DMX59=Array(LM_DMD_DMX59_Parts)
'Dim BL_DMD_DMX6: BL_DMD_DMX6=Array(LM_DMD_DMX6_Parts)
'Dim BL_DMD_DMX60: BL_DMD_DMX60=Array(LM_DMD_DMX60_Parts)
'Dim BL_DMD_DMX61: BL_DMD_DMX61=Array(LM_DMD_DMX61_Parts)
'Dim BL_DMD_DMX62: BL_DMD_DMX62=Array(LM_DMD_DMX62_Parts)
'Dim BL_DMD_DMX63: BL_DMD_DMX63=Array(LM_DMD_DMX63_Parts)
'Dim BL_DMD_DMX64: BL_DMD_DMX64=Array(LM_DMD_DMX64_Parts)
'Dim BL_DMD_DMX65: BL_DMD_DMX65=Array(LM_DMD_DMX65_Parts)
'Dim BL_DMD_DMX66: BL_DMD_DMX66=Array(LM_DMD_DMX66_Parts)
'Dim BL_DMD_DMX67: BL_DMD_DMX67=Array(LM_DMD_DMX67_Parts)
'Dim BL_DMD_DMX68: BL_DMD_DMX68=Array(LM_DMD_DMX68_Parts)
'Dim BL_DMD_DMX69: BL_DMD_DMX69=Array(LM_DMD_DMX69_Parts)
'Dim BL_DMD_DMX7: BL_DMD_DMX7=Array(LM_DMD_DMX7_Parts)
'Dim BL_DMD_DMX70: BL_DMD_DMX70=Array(LM_DMD_DMX70_Parts)
'Dim BL_DMD_DMX71: BL_DMD_DMX71=Array(LM_DMD_DMX71_Parts)
'Dim BL_DMD_DMX72: BL_DMD_DMX72=Array(LM_DMD_DMX72_Parts)
'Dim BL_DMD_DMX73: BL_DMD_DMX73=Array(LM_DMD_DMX73_Parts)
'Dim BL_DMD_DMX74: BL_DMD_DMX74=Array(LM_DMD_DMX74_Parts)
'Dim BL_DMD_DMX75: BL_DMD_DMX75=Array(LM_DMD_DMX75_Parts)
'Dim BL_DMD_DMX76: BL_DMD_DMX76=Array(LM_DMD_DMX76_Parts)
'Dim BL_DMD_DMX77: BL_DMD_DMX77=Array(LM_DMD_DMX77_Parts)
'Dim BL_DMD_DMX78: BL_DMD_DMX78=Array(LM_DMD_DMX78_Parts)
'Dim BL_DMD_DMX79: BL_DMD_DMX79=Array(LM_DMD_DMX79_Parts)
'Dim BL_DMD_DMX8: BL_DMD_DMX8=Array(LM_DMD_DMX8_Parts)
'Dim BL_DMD_DMX80: BL_DMD_DMX80=Array(LM_DMD_DMX80_Parts)
'Dim BL_DMD_DMX81: BL_DMD_DMX81=Array(LM_DMD_DMX81_Parts)
'Dim BL_DMD_DMX82: BL_DMD_DMX82=Array(LM_DMD_DMX82_Parts)
'Dim BL_DMD_DMX83: BL_DMD_DMX83=Array(LM_DMD_DMX83_Parts)
'Dim BL_DMD_DMX84: BL_DMD_DMX84=Array(LM_DMD_DMX84_Parts)
'Dim BL_DMD_DMX85: BL_DMD_DMX85=Array(LM_DMD_DMX85_Parts)
'Dim BL_DMD_DMX86: BL_DMD_DMX86=Array(LM_DMD_DMX86_Parts)
'Dim BL_DMD_DMX87: BL_DMD_DMX87=Array(LM_DMD_DMX87_Parts)
'Dim BL_DMD_DMX88: BL_DMD_DMX88=Array(LM_DMD_DMX88_Parts)
'Dim BL_DMD_DMX89: BL_DMD_DMX89=Array(LM_DMD_DMX89_Parts)
'Dim BL_DMD_DMX9: BL_DMD_DMX9=Array(LM_DMD_DMX9_Parts)
'Dim BL_DMD_DMX90: BL_DMD_DMX90=Array(LM_DMD_DMX90_Parts)
'Dim BL_DMD_DMX91: BL_DMD_DMX91=Array(LM_DMD_DMX91_Parts)
'Dim BL_DMD_DMX92: BL_DMD_DMX92=Array(LM_DMD_DMX92_Parts)
'Dim BL_DMD_DMX93: BL_DMD_DMX93=Array(LM_DMD_DMX93_Parts)
'Dim BL_DMD_DMX94: BL_DMD_DMX94=Array(LM_DMD_DMX94_Parts)
'Dim BL_DMD_DMX95: BL_DMD_DMX95=Array(LM_DMD_DMX95_Parts)
'Dim BL_DMD_DMX96: BL_DMD_DMX96=Array(LM_DMD_DMX96_Parts)
'Dim BL_DMD_DMX97: BL_DMD_DMX97=Array(LM_DMD_DMX97_Parts)
'Dim BL_DMD_DMX98: BL_DMD_DMX98=Array(LM_DMD_DMX98_Parts)
'Dim BL_DMD_DMX99: BL_DMD_DMX99=Array(LM_DMD_DMX99_Parts)
Dim BL_Flashers_L121: BL_Flashers_L121=Array(LM_Flashers_L121_Layer1, LM_Flashers_L121_Layer2, LM_Flashers_L121_LeftFlipper, LM_Flashers_L121_LeftFlipper1, LM_Flashers_L121_LeftFlipper1U, LM_Flashers_L121_LeftFlipperU, LM_Flashers_L121_LeftSling1, LM_Flashers_L121_LeftSling2, LM_Flashers_L121_LeftSling3, LM_Flashers_L121_LeftSling4, LM_Flashers_L121_Lemk, LM_Flashers_L121_Parts, LM_Flashers_L121_Parts2, LM_Flashers_L121_PinCab_Rails, LM_Flashers_L121_Playfield, LM_Flashers_L121_Remk, LM_Flashers_L121_RightFlipperU, LM_Flashers_L121_RightSling1, LM_Flashers_L121_RightSling2, LM_Flashers_L121_RightSling3, LM_Flashers_L121_RightSling4, LM_Flashers_L121_ScrambledEgg, LM_Flashers_L121_sw22, LM_Flashers_L121_sw42, LM_Flashers_L121_sw43, LM_Flashers_L121_sw57, LM_Flashers_L121_sw58, LM_Flashers_L121_sw61)
Dim BL_Flashers_L122: BL_Flashers_L122=Array(LM_Flashers_L122_Gate1, LM_Flashers_L122_Layer1, LM_Flashers_L122_Layer2, LM_Flashers_L122_LeftFlipperU, LM_Flashers_L122_LeftSling1, LM_Flashers_L122_LeftSling2, LM_Flashers_L122_LeftSling3, LM_Flashers_L122_LeftSling4, LM_Flashers_L122_Lemk, LM_Flashers_L122_Parts, LM_Flashers_L122_PinCab_Rails, LM_Flashers_L122_Playfield, LM_Flashers_L122_Remk, LM_Flashers_L122_RightFlipper, LM_Flashers_L122_RightFlipperU, LM_Flashers_L122_RightSling1, LM_Flashers_L122_RightSling2, LM_Flashers_L122_RightSling3, LM_Flashers_L122_RightSling4, LM_Flashers_L122_sw44, LM_Flashers_L122_sw58, LM_Flashers_L122_sw60, LM_Flashers_L122_sw61, LM_Flashers_L122_swPlunger)
Dim BL_Flashers_L123: BL_Flashers_L123=Array(LM_Flashers_L123_Gate5, LM_Flashers_L123_Layer1, LM_Flashers_L123_LeftFlipper1, LM_Flashers_L123_Parts, LM_Flashers_L123_Playfield, LM_Flashers_L123_RightFlipper1U, LM_Flashers_L123_sw18, LM_Flashers_L123_sw19, LM_Flashers_L123_sw44, LM_Flashers_L123_sw45, LM_Flashers_L123_sw46)
Dim BL_Flashers_L127: BL_Flashers_L127=Array(LM_Flashers_L127_Bumper2, LM_Flashers_L127_Bumper3, LM_Flashers_L127_Layer1, LM_Flashers_L127_Layer2, LM_Flashers_L127_Parts, LM_Flashers_L127_Playfield, LM_Flashers_L127_sw30)
Dim BL_Flashers_L129: BL_Flashers_L129=Array(LM_Flashers_L129_Gate5, LM_Flashers_L129_GhostTarget, LM_Flashers_L129_Layer1, LM_Flashers_L129_Parts, LM_Flashers_L129_Playfield, LM_Flashers_L129_RightFlipper1, LM_Flashers_L129_RightFlipper1U, LM_Flashers_L129_sw17, LM_Flashers_L129_sw30, LM_Flashers_L129_sw31, LM_Flashers_L129_sw32, LM_Flashers_L129_sw39, LM_Flashers_L129_sw40)
Dim BL_Flashers_L130: BL_Flashers_L130=Array(LM_Flashers_L130_Bumper3, LM_Flashers_L130_GhostTarget, LM_Flashers_L130_Layer1, LM_Flashers_L130_Layer2, LM_Flashers_L130_Parts, LM_Flashers_L130_Playfield, LM_Flashers_L130_RightFlipper1, LM_Flashers_L130_RightFlipper1U, LM_Flashers_L130_sw17, LM_Flashers_L130_sw18, LM_Flashers_L130_sw30, LM_Flashers_L130_sw31, LM_Flashers_L130_sw32, LM_Flashers_L130_sw39, LM_Flashers_L130_sw40)
Dim BL_Flashers_L131: BL_Flashers_L131=Array(LM_Flashers_L131_Bumper1, LM_Flashers_L131_Bumper3, LM_Flashers_L131_Dummy, LM_Flashers_L131_Gate3, LM_Flashers_L131_GhostTarget, LM_Flashers_L131_Layer1, LM_Flashers_L131_Layer2, LM_Flashers_L131_Parts, LM_Flashers_L131_Playfield, LM_Flashers_L131_RightFlipper1, LM_Flashers_L131_RightFlipper1U, LM_Flashers_L131_sw17, LM_Flashers_L131_sw18, LM_Flashers_L131_sw19, LM_Flashers_L131_sw30, LM_Flashers_L131_sw31, LM_Flashers_L131_sw32, LM_Flashers_L131_sw39, LM_Flashers_L131_sw40)
Dim BL_Flashers_L132: BL_Flashers_L132=Array(LM_Flashers_L132_Layer1, LM_Flashers_L132_LeftFlipper1, LM_Flashers_L132_LeftFlipper1U, LM_Flashers_L132_Parts, LM_Flashers_L132_Playfield, LM_Flashers_L132_ScrambledEgg, LM_Flashers_L132_sw22)
Dim BL_Flashers_L48: BL_Flashers_L48=Array(LM_Flashers_L48_Layer1, LM_Flashers_L48_Parts, LM_Flashers_L48_Playfield, LM_Flashers_L48_RightFlipper1, LM_Flashers_L48_RightFlipper1U, LM_Flashers_L48_sw31, LM_Flashers_L48_sw32, LM_Flashers_L48_sw37, LM_Flashers_L48_sw39, LM_Flashers_L48_sw40)
Dim BL_Flashers_L57: BL_Flashers_L57=Array(LM_Flashers_L57_Bumper2, LM_Flashers_L57_Layer1, LM_Flashers_L57_Parts, LM_Flashers_L57_Playfield)
Dim BL_Flashers_L58: BL_Flashers_L58=Array(LM_Flashers_L58_Bumper3, LM_Flashers_L58_Dummy, LM_Flashers_L58_Layer1, LM_Flashers_L58_Layer2, LM_Flashers_L58_Parts, LM_Flashers_L58_Playfield, LM_Flashers_L58_sw17, LM_Flashers_L58_sw30, LM_Flashers_L58_sw31, LM_Flashers_L58_sw32)
Dim BL_Flashers_L59: BL_Flashers_L59=Array(LM_Flashers_L59_Bumper1, LM_Flashers_L59_Layer1, LM_Flashers_L59_Parts, LM_Flashers_L59_Playfield, LM_Flashers_L59_ScrambledEgg, LM_Flashers_L59_sw30)
'Dim BL_Gi: BL_Gi=Array(LM_Gi_Bumper1, LM_Gi_Bumper2, LM_Gi_Dummy, LM_Gi_Gate003, LM_Gi_Gate004, LM_Gi_Gate1, LM_Gi_Gate3, LM_Gi_Gate5, LM_Gi_Gate6, LM_Gi_Gate8, LM_Gi_GhostTarget, LM_Gi_Layer1, LM_Gi_Layer2, LM_Gi_LeftFlipper, LM_Gi_LeftFlipper1, LM_Gi_LeftFlipper1U, LM_Gi_LeftFlipperU, LM_Gi_LeftSling1, LM_Gi_LeftSling2, LM_Gi_LeftSling3, LM_Gi_LeftSling4, LM_Gi_Parts, LM_Gi_Parts2, LM_Gi_Playfield, LM_Gi_Post, LM_Gi_RightFlipper1, LM_Gi_RightFlipper1U, LM_Gi_RightFlipperU, LM_Gi_RightSling1, LM_Gi_RightSling2, LM_Gi_RightSling3, LM_Gi_RightSling4, LM_Gi_ScrambledEgg, LM_Gi_sw18, LM_Gi_sw22, LM_Gi_sw25, LM_Gi_sw26, LM_Gi_sw27, LM_Gi_sw28, LM_Gi_sw32, LM_Gi_sw39, LM_Gi_sw40, LM_Gi_sw42, LM_Gi_sw44, LM_Gi_sw45, LM_Gi_sw46, LM_Gi_sw58, LM_Gi_sw60, LM_Gi_sw61)
'Dim BL_Gi1_gi01: BL_Gi1_gi01=Array(LM_Gi1_gi01_Layer1, LM_Gi1_gi01_Layer2, LM_Gi1_gi01_LeftFlipper, LM_Gi1_gi01_LeftFlipperU, LM_Gi1_gi01_LeftSling1, LM_Gi1_gi01_LeftSling2, LM_Gi1_gi01_LeftSling3, LM_Gi1_gi01_LeftSling4, LM_Gi1_gi01_Lemk, LM_Gi1_gi01_Parts, LM_Gi1_gi01_Parts2, LM_Gi1_gi01_Playfield, LM_Gi1_gi01_RightFlipperU, LM_Gi1_gi01_sw58)
'Dim BL_Gi1_gi02: BL_Gi1_gi02=Array(LM_Gi1_gi02_Layer1, LM_Gi1_gi02_Layer2, LM_Gi1_gi02_LeftFlipper, LM_Gi1_gi02_LeftFlipperU, LM_Gi1_gi02_LeftSling1, LM_Gi1_gi02_LeftSling2, LM_Gi1_gi02_LeftSling3, LM_Gi1_gi02_LeftSling4, LM_Gi1_gi02_Lemk, LM_Gi1_gi02_Parts, LM_Gi1_gi02_Parts2, LM_Gi1_gi02_Playfield, LM_Gi1_gi02_sw39, LM_Gi1_gi02_sw57, LM_Gi1_gi02_sw58)
'Dim BL_Gi1_gi03: BL_Gi1_gi03=Array(LM_Gi1_gi03_Layer1, LM_Gi1_gi03_Layer2, LM_Gi1_gi03_LeftFlipper, LM_Gi1_gi03_LeftFlipperU, LM_Gi1_gi03_LeftSling1, LM_Gi1_gi03_LeftSling2, LM_Gi1_gi03_LeftSling3, LM_Gi1_gi03_LeftSling4, LM_Gi1_gi03_Lemk, LM_Gi1_gi03_Parts, LM_Gi1_gi03_Parts2, LM_Gi1_gi03_Playfield, LM_Gi1_gi03_sw39, LM_Gi1_gi03_sw57, LM_Gi1_gi03_sw58)
'Dim BL_Gi1_gi21: BL_Gi1_gi21=Array(LM_Gi1_gi21_Layer1, LM_Gi1_gi21_LeftFlipperU, LM_Gi1_gi21_Parts, LM_Gi1_gi21_Playfield, LM_Gi1_gi21_Remk, LM_Gi1_gi21_RightFlipper, LM_Gi1_gi21_RightFlipperU, LM_Gi1_gi21_RightSling1, LM_Gi1_gi21_RightSling2, LM_Gi1_gi21_RightSling3, LM_Gi1_gi21_RightSling4, LM_Gi1_gi21_sw60, LM_Gi1_gi21_sw61)
'Dim BL_Gi1_gi22: BL_Gi1_gi22=Array(LM_Gi1_gi22_Layer1, LM_Gi1_gi22_Parts, LM_Gi1_gi22_Playfield, LM_Gi1_gi22_Remk, LM_Gi1_gi22_RightFlipper, LM_Gi1_gi22_RightFlipperU, LM_Gi1_gi22_RightSling1, LM_Gi1_gi22_RightSling2, LM_Gi1_gi22_RightSling3, LM_Gi1_gi22_RightSling4, LM_Gi1_gi22_sw60, LM_Gi1_gi22_sw61)
'Dim BL_Gi1_gi23: BL_Gi1_gi23=Array(LM_Gi1_gi23_Layer1, LM_Gi1_gi23_LeftFlipperU, LM_Gi1_gi23_Parts, LM_Gi1_gi23_Playfield, LM_Gi1_gi23_Remk, LM_Gi1_gi23_RightFlipper, LM_Gi1_gi23_RightFlipperU, LM_Gi1_gi23_RightSling1, LM_Gi1_gi23_RightSling2, LM_Gi1_gi23_RightSling3, LM_Gi1_gi23_RightSling4, LM_Gi1_gi23_sw61)
'Dim BL_Inserts_L1: BL_Inserts_L1=Array(LM_Inserts_L1_LeftFlipper, LM_Inserts_L1_LeftFlipperU, LM_Inserts_L1_LeftSling1, LM_Inserts_L1_Parts, LM_Inserts_L1_Playfield)
'Dim BL_Inserts_L10: BL_Inserts_L10=Array(LM_Inserts_L10_LeftSling1, LM_Inserts_L10_LeftSling2, LM_Inserts_L10_LeftSling3, LM_Inserts_L10_LeftSling4, LM_Inserts_L10_Lemk, LM_Inserts_L10_Parts, LM_Inserts_L10_Playfield)
'Dim BL_Inserts_L11: BL_Inserts_L11=Array(LM_Inserts_L11_Layer1, LM_Inserts_L11_Playfield)
'Dim BL_Inserts_L12: BL_Inserts_L12=Array(LM_Inserts_L12_Playfield)
'Dim BL_Inserts_L13: BL_Inserts_L13=Array(LM_Inserts_L13_Playfield)
'Dim BL_Inserts_L14: BL_Inserts_L14=Array(LM_Inserts_L14_Parts, LM_Inserts_L14_Playfield)
'Dim BL_Inserts_L15: BL_Inserts_L15=Array(LM_Inserts_L15_Playfield)
'Dim BL_Inserts_L16: BL_Inserts_L16=Array(LM_Inserts_L16_Playfield)
'Dim BL_Inserts_L17: BL_Inserts_L17=Array(LM_Inserts_L17_LeftFlipper1U, LM_Inserts_L17_Parts, LM_Inserts_L17_Playfield, LM_Inserts_L17_sw22)
'Dim BL_Inserts_L18: BL_Inserts_L18=Array(LM_Inserts_L18_Layer1, LM_Inserts_L18_LeftFlipper1U, LM_Inserts_L18_Parts, LM_Inserts_L18_Playfield, LM_Inserts_L18_sw22)
'Dim BL_Inserts_L19: BL_Inserts_L19=Array(LM_Inserts_L19_Layer1, LM_Inserts_L19_Playfield)
'Dim BL_Inserts_L2: BL_Inserts_L2=Array(LM_Inserts_L2_LeftFlipper, LM_Inserts_L2_LeftFlipperU, LM_Inserts_L2_Playfield)
'Dim BL_Inserts_L20: BL_Inserts_L20=Array(LM_Inserts_L20_Playfield, LM_Inserts_L20_sw22)
'Dim BL_Inserts_L21: BL_Inserts_L21=Array(LM_Inserts_L21_Playfield)
'Dim BL_Inserts_L22: BL_Inserts_L22=Array(LM_Inserts_L22_Playfield)
'Dim BL_Inserts_L23: BL_Inserts_L23=Array(LM_Inserts_L23_LeftFlipper, LM_Inserts_L23_LeftFlipperU, LM_Inserts_L23_Playfield, LM_Inserts_L23_RightFlipper, LM_Inserts_L23_RightFlipperU)
'Dim BL_Inserts_L24: BL_Inserts_L24=Array(LM_Inserts_L24_Parts, LM_Inserts_L24_Playfield)
'Dim BL_Inserts_L25: BL_Inserts_L25=Array(LM_Inserts_L25_Playfield, LM_Inserts_L25_sw39)
'Dim BL_Inserts_L26: BL_Inserts_L26=Array(LM_Inserts_L26_Parts, LM_Inserts_L26_Playfield, LM_Inserts_L26_sw39)
'Dim BL_Inserts_L27: BL_Inserts_L27=Array(LM_Inserts_L27_Parts, LM_Inserts_L27_Playfield, LM_Inserts_L27_sw39)
'Dim BL_Inserts_L28: BL_Inserts_L28=Array(LM_Inserts_L28_Parts, LM_Inserts_L28_Playfield, LM_Inserts_L28_sw39, LM_Inserts_L28_sw44)
'Dim BL_Inserts_L29: BL_Inserts_L29=Array(LM_Inserts_L29_Parts, LM_Inserts_L29_Playfield, LM_Inserts_L29_sw44, LM_Inserts_L29_sw45)
'Dim BL_Inserts_L3: BL_Inserts_L3=Array(LM_Inserts_L3_LeftFlipperU, LM_Inserts_L3_Parts, LM_Inserts_L3_Playfield, LM_Inserts_L3_RightFlipperU)
'Dim BL_Inserts_L30: BL_Inserts_L30=Array(LM_Inserts_L30_Parts, LM_Inserts_L30_Playfield, LM_Inserts_L30_sw44, LM_Inserts_L30_sw45, LM_Inserts_L30_sw46)
'Dim BL_Inserts_L31: BL_Inserts_L31=Array(LM_Inserts_L31_Playfield, LM_Inserts_L31_sw45, LM_Inserts_L31_sw46)
'Dim BL_Inserts_L32: BL_Inserts_L32=Array(LM_Inserts_L32_Playfield, LM_Inserts_L32_sw46)
'Dim BL_Inserts_L33: BL_Inserts_L33=Array(LM_Inserts_L33_Parts, LM_Inserts_L33_Playfield, LM_Inserts_L33_sw22)
'Dim BL_Inserts_L34: BL_Inserts_L34=Array(LM_Inserts_L34_Playfield)
'Dim BL_Inserts_L35: BL_Inserts_L35=Array(LM_Inserts_L35_Layer1, LM_Inserts_L35_Parts, LM_Inserts_L35_Playfield)
'Dim BL_Inserts_L36: BL_Inserts_L36=Array(LM_Inserts_L36_Layer1, LM_Inserts_L36_Parts, LM_Inserts_L36_Playfield)
'Dim BL_Inserts_L37: BL_Inserts_L37=Array(LM_Inserts_L37_Parts, LM_Inserts_L37_Playfield)
'Dim BL_Inserts_L38: BL_Inserts_L38=Array(LM_Inserts_L38_Parts, LM_Inserts_L38_Playfield)
'Dim BL_Inserts_L39: BL_Inserts_L39=Array(LM_Inserts_L39_Parts, LM_Inserts_L39_Playfield)
'Dim BL_Inserts_L4: BL_Inserts_L4=Array(LM_Inserts_L4_Parts, LM_Inserts_L4_Playfield, LM_Inserts_L4_RightFlipper, LM_Inserts_L4_RightFlipperU)
'Dim BL_Inserts_L40: BL_Inserts_L40=Array(LM_Inserts_L40_Playfield)
'Dim BL_Inserts_L41: BL_Inserts_L41=Array(LM_Inserts_L41_Parts, LM_Inserts_L41_Playfield, LM_Inserts_L41_RightFlipper1U)
'Dim BL_Inserts_L42: BL_Inserts_L42=Array(LM_Inserts_L42_Parts, LM_Inserts_L42_Playfield, LM_Inserts_L42_RightFlipper1U, LM_Inserts_L42_sw40)
'Dim BL_Inserts_L43: BL_Inserts_L43=Array(LM_Inserts_L43_Parts, LM_Inserts_L43_Playfield, LM_Inserts_L43_sw40)
'Dim BL_Inserts_L44: BL_Inserts_L44=Array(LM_Inserts_L44_Parts, LM_Inserts_L44_Playfield, LM_Inserts_L44_RightFlipper1U, LM_Inserts_L44_sw39)
'Dim BL_Inserts_L45: BL_Inserts_L45=Array(LM_Inserts_L45_Parts, LM_Inserts_L45_Playfield, LM_Inserts_L45_RightFlipper1U, LM_Inserts_L45_sw39)
'Dim BL_Inserts_L46: BL_Inserts_L46=Array(LM_Inserts_L46_Parts, LM_Inserts_L46_Playfield, LM_Inserts_L46_sw39)
'Dim BL_Inserts_L47: BL_Inserts_L47=Array(LM_Inserts_L47_Parts, LM_Inserts_L47_Playfield, LM_Inserts_L47_RightFlipper1U, LM_Inserts_L47_sw39)
'Dim BL_Inserts_L49: BL_Inserts_L49=Array(LM_Inserts_L49_Parts, LM_Inserts_L49_Playfield)
'Dim BL_Inserts_L5: BL_Inserts_L5=Array(LM_Inserts_L5_Playfield, LM_Inserts_L5_RightFlipper, LM_Inserts_L5_RightFlipperU, LM_Inserts_L5_RightSling1)
'Dim BL_Inserts_L50: BL_Inserts_L50=Array(LM_Inserts_L50_GhostTarget, LM_Inserts_L50_Parts, LM_Inserts_L50_Playfield)
'Dim BL_Inserts_L51: BL_Inserts_L51=Array(LM_Inserts_L51_GhostTarget, LM_Inserts_L51_Layer1, LM_Inserts_L51_Parts, LM_Inserts_L51_Playfield, LM_Inserts_L51_sw32)
'Dim BL_Inserts_L52: BL_Inserts_L52=Array(LM_Inserts_L52_Parts, LM_Inserts_L52_Playfield, LM_Inserts_L52_sw32)
'Dim BL_Inserts_L53: BL_Inserts_L53=Array(LM_Inserts_L53_Parts, LM_Inserts_L53_Playfield)
'Dim BL_Inserts_L54: BL_Inserts_L54=Array(LM_Inserts_L54_GhostTarget, LM_Inserts_L54_Parts, LM_Inserts_L54_Playfield, LM_Inserts_L54_sw40)
'Dim BL_Inserts_L55: BL_Inserts_L55=Array(LM_Inserts_L55_GhostTarget, LM_Inserts_L55_Parts, LM_Inserts_L55_Playfield, LM_Inserts_L55_sw40)
'Dim BL_Inserts_L56: BL_Inserts_L56=Array(LM_Inserts_L56_GhostTarget, LM_Inserts_L56_Parts, LM_Inserts_L56_Playfield, LM_Inserts_L56_sw40)
'Dim BL_Inserts_L6: BL_Inserts_L6=Array(LM_Inserts_L6_Playfield, LM_Inserts_L6_RightFlipperU, LM_Inserts_L6_RightSling1, LM_Inserts_L6_RightSling2, LM_Inserts_L6_RightSling3, LM_Inserts_L6_RightSling4)
'Dim BL_Inserts_L60: BL_Inserts_L60=Array(LM_Inserts_L60_Parts, LM_Inserts_L60_Playfield, LM_Inserts_L60_sw30, LM_Inserts_L60_sw31, LM_Inserts_L60_sw32)
'Dim BL_Inserts_L61: BL_Inserts_L61=Array(LM_Inserts_L61_Playfield, LM_Inserts_L61_sw17, LM_Inserts_L61_sw18)
'Dim BL_Inserts_L62: BL_Inserts_L62=Array(LM_Inserts_L62_Parts, LM_Inserts_L62_Playfield)
'Dim BL_Inserts_L63: BL_Inserts_L63=Array(LM_Inserts_L63_Playfield)
'Dim BL_Inserts_L7: BL_Inserts_L7=Array(LM_Inserts_L7_Playfield)
'Dim BL_Inserts_L70: BL_Inserts_L70=Array(LM_Inserts_L70_Layer1, LM_Inserts_L70_Parts, LM_Inserts_L70_Playfield, LM_Inserts_L70_Post)
'Dim BL_Inserts_L71: BL_Inserts_L71=Array(LM_Inserts_L71_Layer1, LM_Inserts_L71_Layer2, LM_Inserts_L71_Parts, LM_Inserts_L71_Playfield)
'Dim BL_Inserts_L72: BL_Inserts_L72=Array(LM_Inserts_L72_Layer1, LM_Inserts_L72_Layer2, LM_Inserts_L72_Parts, LM_Inserts_L72_Playfield)
'Dim BL_Inserts_L78: BL_Inserts_L78=Array(LM_Inserts_L78_Layer1, LM_Inserts_L78_Parts, LM_Inserts_L78_Playfield, LM_Inserts_L78_sw29)
'Dim BL_Inserts_L79: BL_Inserts_L79=Array(LM_Inserts_L79_Bumper2, LM_Inserts_L79_Bumper3, LM_Inserts_L79_Layer1, LM_Inserts_L79_Layer2, LM_Inserts_L79_Parts, LM_Inserts_L79_Playfield, LM_Inserts_L79_sw30, LM_Inserts_L79_sw31)
'Dim BL_Inserts_L8: BL_Inserts_L8=Array(LM_Inserts_L8_Parts, LM_Inserts_L8_Playfield, LM_Inserts_L8_sw61)
'Dim BL_Inserts_L9: BL_Inserts_L9=Array(LM_Inserts_L9_Layer1, LM_Inserts_L9_Layer2, LM_Inserts_L9_Parts, LM_Inserts_L9_Playfield, LM_Inserts_L9_sw58)
'Dim BL_Lit_Room: BL_Lit_Room=Array(BM_Bumper1, BM_Bumper2, BM_Bumper3, BM_Dummy, BM_Gate002, BM_Gate003, BM_Gate004, BM_Gate005, BM_Gate1, BM_Gate3, BM_Gate5, BM_Gate6, BM_Gate7, BM_Gate8, BM_GhostTarget, BM_Layer1, BM_Layer2, BM_LeftFlipper, BM_LeftFlipper1, BM_LeftFlipper1U, BM_LeftFlipperU, BM_LeftSling1, BM_LeftSling2, BM_LeftSling3, BM_LeftSling4, BM_Lemk, BM_Parts, BM_Parts2, BM_PinCab_Rails, BM_Playfield, BM_Post, BM_Remk, BM_RightFlipper, BM_RightFlipper1, BM_RightFlipper1U, BM_RightFlipperU, BM_RightSling1, BM_RightSling2, BM_RightSling3, BM_RightSling4, BM_ScrambledEgg, BM_sw17, BM_sw18, BM_sw19, BM_sw21, BM_sw22, BM_sw25, BM_sw26, BM_sw27, BM_sw28, BM_sw29, BM_sw30, BM_sw31, BM_sw32, BM_sw37, BM_sw39, BM_sw40, BM_sw42, BM_sw43, BM_sw44, BM_sw45, BM_sw46, BM_sw57, BM_sw58, BM_sw60, BM_sw61, BM_swPlunger)
'Dim BL_Opto: BL_Opto=Array(LM_Opto_Parts, LM_Opto_Playfield, LM_Opto_sw39)
'Dim BL_Spotlights_gi04: BL_Spotlights_gi04=Array(LM_Spotlights_gi04_Bumper1, LM_Spotlights_gi04_Bumper3, LM_Spotlights_gi04_Dummy, LM_Spotlights_gi04_Gate002, LM_Spotlights_gi04_Gate003, LM_Spotlights_gi04_Gate005, LM_Spotlights_gi04_Gate6, LM_Spotlights_gi04_GhostTarget, LM_Spotlights_gi04_Layer1, LM_Spotlights_gi04_Layer2, LM_Spotlights_gi04_Parts, LM_Spotlights_gi04_Playfield, LM_Spotlights_gi04_RF1, LM_Spotlights_gi04_RF1U, LM_Spotlights_gi04_sw17, LM_Spotlights_gi04_sw18, LM_Spotlights_gi04_sw19, LM_Spotlights_gi04_sw27, LM_Spotlights_gi04_sw28, LM_Spotlights_gi04_sw29, LM_Spotlights_gi04_sw30, LM_Spotlights_gi04_sw31, LM_Spotlights_gi04_sw32, LM_Spotlights_gi04_sw39, LM_Spotlights_gi04_sw40)
' Global arrays
'Dim BG_Bakemap: BG_Bakemap=Array(BM_Bumper1, BM_Bumper2, BM_Bumper3, BM_Dummy, BM_Gate002, BM_Gate003, BM_Gate004, BM_Gate005, BM_Gate1, BM_Gate3, BM_Gate5, BM_Gate6, BM_Gate7, BM_Gate8, BM_GhostTarget, BM_Layer1, BM_Layer2, BM_LeftFlipper, BM_LeftFlipper1, BM_LeftFlipper1U, BM_LeftFlipperU, BM_LeftSling1, BM_LeftSling2, BM_LeftSling3, BM_LeftSling4, BM_Lemk, BM_Parts, BM_Parts2, BM_PinCab_Rails, BM_Playfield, BM_Post, BM_Remk, BM_RightFlipper, BM_RightFlipper1, BM_RightFlipper1U, BM_RightFlipperU, BM_RightSling1, BM_RightSling2, BM_RightSling3, BM_RightSling4, BM_ScrambledEgg, BM_sw17, BM_sw18, BM_sw19, BM_sw21, BM_sw22, BM_sw25, BM_sw26, BM_sw27, BM_sw28, BM_sw29, BM_sw30, BM_sw31, BM_sw32, BM_sw37, BM_sw39, BM_sw40, BM_sw42, BM_sw43, BM_sw44, BM_sw45, BM_sw46, BM_sw57, BM_sw58, BM_sw60, BM_sw61, BM_swPlunger)
'Dim BG_Lightmap: BG_Lightmap=Array(LM_Backlight_L65_Layer1, LM_Backlight_L65_Parts, LM_Backlight_L66_Layer1, LM_Backlight_L66_Parts, LM_Backlight_L66_Playfield, LM_Backlight_L66_sw28, LM_Backlight_L67_Dummy, LM_Backlight_L67_Layer2, LM_Backlight_L67_Parts, LM_Backlight_L67_Playfield, LM_Backlight_L68_Dummy, LM_Backlight_L68_Layer2, LM_Backlight_L68_Parts, LM_Backlight_L68_Playfield, LM_Backlight_L68_sw30, LM_Backlight_L69_Layer1, LM_Backlight_L69_Layer2, LM_Backlight_L69_Parts, LM_Backlight_L69_Playfield, LM_Backlight_L73_Layer1, LM_Backlight_L73_Layer2, LM_Backlight_L73_Parts, LM_Backlight_L73_Playfield, LM_Backlight_L74_Layer1, LM_Backlight_L74_Layer2, LM_Backlight_L74_Parts, LM_Backlight_L75_Layer1, LM_Backlight_L75_Layer2, LM_Backlight_L75_Parts, LM_Backlight_L75_Playfield, LM_Backlight_L76_Bumper3, LM_Backlight_L76_Layer1, LM_Backlight_L76_Layer2, LM_Backlight_L76_Parts, LM_Backlight_L76_Playfield, LM_Backlight_L77_Layer1, LM_Backlight_L77_Parts, LM_DMD_DMX1_Parts, LM_DMD_DMX10_Parts, _
'	LM_DMD_DMX100_Parts, LM_DMD_DMX101_Parts, LM_DMD_DMX102_Parts, LM_DMD_DMX103_Parts, LM_DMD_DMX104_Parts, LM_DMD_DMX105_Parts, LM_DMD_DMX11_Parts, LM_DMD_DMX12_Parts, LM_DMD_DMX13_Parts, LM_DMD_DMX14_Parts, LM_DMD_DMX15_Parts, LM_DMD_DMX16_Parts, LM_DMD_DMX17_Parts, LM_DMD_DMX18_Parts, LM_DMD_DMX19_Parts, LM_DMD_DMX2_Parts, LM_DMD_DMX20_Parts, LM_DMD_DMX21_Parts, LM_DMD_DMX22_Parts, LM_DMD_DMX23_Parts, LM_DMD_DMX24_Parts, LM_DMD_DMX25_Parts, LM_DMD_DMX26_Parts, LM_DMD_DMX27_Parts, LM_DMD_DMX28_Parts, LM_DMD_DMX29_Parts, LM_DMD_DMX3_Parts, LM_DMD_DMX30_Parts, LM_DMD_DMX31_Parts, LM_DMD_DMX32_Parts, LM_DMD_DMX33_Parts, LM_DMD_DMX34_Parts, LM_DMD_DMX35_Parts, LM_DMD_DMX36_Parts, LM_DMD_DMX37_Parts, LM_DMD_DMX38_Parts, LM_DMD_DMX39_Parts, LM_DMD_DMX4_Parts, LM_DMD_DMX40_Parts, LM_DMD_DMX41_Parts, LM_DMD_DMX42_Parts, LM_DMD_DMX43_Parts, LM_DMD_DMX44_Parts, LM_DMD_DMX45_Parts, LM_DMD_DMX46_Parts, LM_DMD_DMX47_Parts, LM_DMD_DMX48_Parts, LM_DMD_DMX49_Parts, LM_DMD_DMX5_Parts, LM_DMD_DMX50_Parts, LM_DMD_DMX51_Parts, _
'	LM_DMD_DMX52_Parts, LM_DMD_DMX53_Parts, LM_DMD_DMX54_Parts, LM_DMD_DMX55_Parts, LM_DMD_DMX56_Parts, LM_DMD_DMX57_Parts, LM_DMD_DMX58_Parts, LM_DMD_DMX59_Parts, LM_DMD_DMX6_Parts, LM_DMD_DMX60_Parts, LM_DMD_DMX61_Parts, LM_DMD_DMX62_Parts, LM_DMD_DMX63_Parts, LM_DMD_DMX64_Parts, LM_DMD_DMX65_Parts, LM_DMD_DMX66_Parts, LM_DMD_DMX67_Parts, LM_DMD_DMX68_Parts, LM_DMD_DMX69_Parts, LM_DMD_DMX7_Parts, LM_DMD_DMX70_Parts, LM_DMD_DMX71_Parts, LM_DMD_DMX72_Parts, LM_DMD_DMX73_Parts, LM_DMD_DMX74_Parts, LM_DMD_DMX75_Parts, LM_DMD_DMX76_Parts, LM_DMD_DMX77_Parts, LM_DMD_DMX78_Parts, LM_DMD_DMX79_Parts, LM_DMD_DMX8_Parts, LM_DMD_DMX80_Parts, LM_DMD_DMX81_Parts, LM_DMD_DMX82_Parts, LM_DMD_DMX83_Parts, LM_DMD_DMX84_Parts, LM_DMD_DMX85_Parts, LM_DMD_DMX86_Parts, LM_DMD_DMX87_Parts, LM_DMD_DMX88_Parts, LM_DMD_DMX89_Parts, LM_DMD_DMX9_Parts, LM_DMD_DMX90_Parts, LM_DMD_DMX91_Parts, LM_DMD_DMX92_Parts, LM_DMD_DMX93_Parts, LM_DMD_DMX94_Parts, LM_DMD_DMX95_Parts, LM_DMD_DMX96_Parts, LM_DMD_DMX97_Parts, LM_DMD_DMX98_Parts, _
'	LM_DMD_DMX99_Parts, LM_Flashers_L121_Layer1, LM_Flashers_L121_Layer2, LM_Flashers_L121_LeftFlipper, LM_Flashers_L121_LeftFlipper1, LM_Flashers_L121_LeftFlipper1U, LM_Flashers_L121_LeftFlipperU, LM_Flashers_L121_LeftSling1, LM_Flashers_L121_LeftSling2, LM_Flashers_L121_LeftSling3, LM_Flashers_L121_LeftSling4, LM_Flashers_L121_Lemk, LM_Flashers_L121_Parts, LM_Flashers_L121_Parts2, LM_Flashers_L121_PinCab_Rails, LM_Flashers_L121_Playfield, LM_Flashers_L121_Remk, LM_Flashers_L121_RightFlipperU, LM_Flashers_L121_RightSling1, LM_Flashers_L121_RightSling2, LM_Flashers_L121_RightSling3, LM_Flashers_L121_RightSling4, LM_Flashers_L121_ScrambledEgg, LM_Flashers_L121_sw22, LM_Flashers_L121_sw42, LM_Flashers_L121_sw43, LM_Flashers_L121_sw57, LM_Flashers_L121_sw58, LM_Flashers_L121_sw61, LM_Flashers_L122_Gate1, LM_Flashers_L122_Layer1, LM_Flashers_L122_Layer2, LM_Flashers_L122_LeftFlipperU, LM_Flashers_L122_LeftSling1, LM_Flashers_L122_LeftSling2, LM_Flashers_L122_LeftSling3, LM_Flashers_L122_LeftSling4, _
'	LM_Flashers_L122_Lemk, LM_Flashers_L122_Parts, LM_Flashers_L122_PinCab_Rails, LM_Flashers_L122_Playfield, LM_Flashers_L122_Remk, LM_Flashers_L122_RightFlipper, LM_Flashers_L122_RightFlipperU, LM_Flashers_L122_RightSling1, LM_Flashers_L122_RightSling2, LM_Flashers_L122_RightSling3, LM_Flashers_L122_RightSling4, LM_Flashers_L122_sw44, LM_Flashers_L122_sw58, LM_Flashers_L122_sw60, LM_Flashers_L122_sw61, LM_Flashers_L122_swPlunger, LM_Flashers_L123_Gate5, LM_Flashers_L123_Layer1, LM_Flashers_L123_LeftFlipper1, LM_Flashers_L123_Parts, LM_Flashers_L123_Playfield, LM_Flashers_L123_RightFlipper1U, LM_Flashers_L123_sw18, LM_Flashers_L123_sw19, LM_Flashers_L123_sw44, LM_Flashers_L123_sw45, LM_Flashers_L123_sw46, LM_Flashers_L127_Bumper2, LM_Flashers_L127_Bumper3, LM_Flashers_L127_Layer1, LM_Flashers_L127_Layer2, LM_Flashers_L127_Parts, LM_Flashers_L127_Playfield, LM_Flashers_L127_sw30, LM_Flashers_L129_Gate5, LM_Flashers_L129_GhostTarget, LM_Flashers_L129_Layer1, LM_Flashers_L129_Parts, LM_Flashers_L129_Playfield, _
'	LM_Flashers_L129_RightFlipper1, LM_Flashers_L129_RightFlipper1U, LM_Flashers_L129_sw17, LM_Flashers_L129_sw30, LM_Flashers_L129_sw31, LM_Flashers_L129_sw32, LM_Flashers_L129_sw39, LM_Flashers_L129_sw40, LM_Flashers_L130_Bumper3, LM_Flashers_L130_GhostTarget, LM_Flashers_L130_Layer1, LM_Flashers_L130_Layer2, LM_Flashers_L130_Parts, LM_Flashers_L130_Playfield, LM_Flashers_L130_RightFlipper1, LM_Flashers_L130_RightFlipper1U, LM_Flashers_L130_sw17, LM_Flashers_L130_sw18, LM_Flashers_L130_sw30, LM_Flashers_L130_sw31, LM_Flashers_L130_sw32, LM_Flashers_L130_sw39, LM_Flashers_L130_sw40, LM_Flashers_L131_Bumper1, LM_Flashers_L131_Bumper3, LM_Flashers_L131_Dummy, LM_Flashers_L131_Gate3, LM_Flashers_L131_GhostTarget, LM_Flashers_L131_Layer1, LM_Flashers_L131_Layer2, LM_Flashers_L131_Parts, LM_Flashers_L131_Playfield, LM_Flashers_L131_RightFlipper1, LM_Flashers_L131_RightFlipper1U, LM_Flashers_L131_sw17, LM_Flashers_L131_sw18, LM_Flashers_L131_sw19, LM_Flashers_L131_sw30, LM_Flashers_L131_sw31, LM_Flashers_L131_sw32, _
'	LM_Flashers_L131_sw39, LM_Flashers_L131_sw40, LM_Flashers_L132_Layer1, LM_Flashers_L132_LeftFlipper1, LM_Flashers_L132_LeftFlipper1U, LM_Flashers_L132_Parts, LM_Flashers_L132_Playfield, LM_Flashers_L132_ScrambledEgg, LM_Flashers_L132_sw22, LM_Flashers_L48_Layer1, LM_Flashers_L48_Parts, LM_Flashers_L48_Playfield, LM_Flashers_L48_RightFlipper1, LM_Flashers_L48_RightFlipper1U, LM_Flashers_L48_sw31, LM_Flashers_L48_sw32, LM_Flashers_L48_sw37, LM_Flashers_L48_sw39, LM_Flashers_L48_sw40, LM_Flashers_L57_Bumper2, LM_Flashers_L57_Layer1, LM_Flashers_L57_Parts, LM_Flashers_L57_Playfield, LM_Flashers_L58_Bumper3, LM_Flashers_L58_Dummy, LM_Flashers_L58_Layer1, LM_Flashers_L58_Layer2, LM_Flashers_L58_Parts, LM_Flashers_L58_Playfield, LM_Flashers_L58_sw17, LM_Flashers_L58_sw30, LM_Flashers_L58_sw31, LM_Flashers_L58_sw32, LM_Flashers_L59_Bumper1, LM_Flashers_L59_Layer1, LM_Flashers_L59_Parts, LM_Flashers_L59_Playfield, LM_Flashers_L59_ScrambledEgg, LM_Flashers_L59_sw30, LM_Gi_Bumper1, LM_Gi_Bumper2, LM_Gi_Dummy, _
'	LM_Gi_Gate003, LM_Gi_Gate004, LM_Gi_Gate1, LM_Gi_Gate3, LM_Gi_Gate5, LM_Gi_Gate6, LM_Gi_Gate8, LM_Gi_GhostTarget, LM_Gi_Layer1, LM_Gi_Layer2, LM_Gi_LeftFlipper, LM_Gi_LeftFlipper1, LM_Gi_LeftFlipper1U, LM_Gi_LeftFlipperU, LM_Gi_LeftSling1, LM_Gi_LeftSling2, LM_Gi_LeftSling3, LM_Gi_LeftSling4, LM_Gi_Parts, LM_Gi_Parts2, LM_Gi_Playfield, LM_Gi_Post, LM_Gi_RightFlipper1, LM_Gi_RightFlipper1U, LM_Gi_RightFlipperU, LM_Gi_RightSling1, LM_Gi_RightSling2, LM_Gi_RightSling3, LM_Gi_RightSling4, LM_Gi_ScrambledEgg, LM_Gi_sw18, LM_Gi_sw22, LM_Gi_sw25, LM_Gi_sw26, LM_Gi_sw27, LM_Gi_sw28, LM_Gi_sw32, LM_Gi_sw39, LM_Gi_sw40, LM_Gi_sw42, LM_Gi_sw44, LM_Gi_sw45, LM_Gi_sw46, LM_Gi_sw58, LM_Gi_sw60, LM_Gi_sw61, LM_Gi1_gi01_Layer1, LM_Gi1_gi01_Layer2, LM_Gi1_gi01_LeftFlipper, LM_Gi1_gi01_LeftFlipperU, LM_Gi1_gi01_LeftSling1, LM_Gi1_gi01_LeftSling2, LM_Gi1_gi01_LeftSling3, LM_Gi1_gi01_LeftSling4, LM_Gi1_gi01_Lemk, LM_Gi1_gi01_Parts, LM_Gi1_gi01_Parts2, LM_Gi1_gi01_Playfield, LM_Gi1_gi01_RightFlipperU, LM_Gi1_gi01_sw58, _
'	LM_Gi1_gi02_Layer1, LM_Gi1_gi02_Layer2, LM_Gi1_gi02_LeftFlipper, LM_Gi1_gi02_LeftFlipperU, LM_Gi1_gi02_LeftSling1, LM_Gi1_gi02_LeftSling2, LM_Gi1_gi02_LeftSling3, LM_Gi1_gi02_LeftSling4, LM_Gi1_gi02_Lemk, LM_Gi1_gi02_Parts, LM_Gi1_gi02_Parts2, LM_Gi1_gi02_Playfield, LM_Gi1_gi02_sw39, LM_Gi1_gi02_sw57, LM_Gi1_gi02_sw58, LM_Gi1_gi03_Layer1, LM_Gi1_gi03_Layer2, LM_Gi1_gi03_LeftFlipper, LM_Gi1_gi03_LeftFlipperU, LM_Gi1_gi03_LeftSling1, LM_Gi1_gi03_LeftSling2, LM_Gi1_gi03_LeftSling3, LM_Gi1_gi03_LeftSling4, LM_Gi1_gi03_Lemk, LM_Gi1_gi03_Parts, LM_Gi1_gi03_Parts2, LM_Gi1_gi03_Playfield, LM_Gi1_gi03_sw39, LM_Gi1_gi03_sw57, LM_Gi1_gi03_sw58, LM_Gi1_gi21_Layer1, LM_Gi1_gi21_LeftFlipperU, LM_Gi1_gi21_Parts, LM_Gi1_gi21_Playfield, LM_Gi1_gi21_Remk, LM_Gi1_gi21_RightFlipper, LM_Gi1_gi21_RightFlipperU, LM_Gi1_gi21_RightSling1, LM_Gi1_gi21_RightSling2, LM_Gi1_gi21_RightSling3, LM_Gi1_gi21_RightSling4, LM_Gi1_gi21_sw60, LM_Gi1_gi21_sw61, LM_Gi1_gi22_Layer1, LM_Gi1_gi22_Parts, LM_Gi1_gi22_Playfield, LM_Gi1_gi22_Remk, _
'	LM_Gi1_gi22_RightFlipper, LM_Gi1_gi22_RightFlipperU, LM_Gi1_gi22_RightSling1, LM_Gi1_gi22_RightSling2, LM_Gi1_gi22_RightSling3, LM_Gi1_gi22_RightSling4, LM_Gi1_gi22_sw60, LM_Gi1_gi22_sw61, LM_Gi1_gi23_Layer1, LM_Gi1_gi23_LeftFlipperU, LM_Gi1_gi23_Parts, LM_Gi1_gi23_Playfield, LM_Gi1_gi23_Remk, LM_Gi1_gi23_RightFlipper, LM_Gi1_gi23_RightFlipperU, LM_Gi1_gi23_RightSling1, LM_Gi1_gi23_RightSling2, LM_Gi1_gi23_RightSling3, LM_Gi1_gi23_RightSling4, LM_Gi1_gi23_sw61, LM_Inserts_L1_LeftFlipper, LM_Inserts_L1_LeftFlipperU, LM_Inserts_L1_LeftSling1, LM_Inserts_L1_Parts, LM_Inserts_L1_Playfield, LM_Inserts_L10_LeftSling1, LM_Inserts_L10_LeftSling2, LM_Inserts_L10_LeftSling3, LM_Inserts_L10_LeftSling4, LM_Inserts_L10_Lemk, LM_Inserts_L10_Parts, LM_Inserts_L10_Playfield, LM_Inserts_L11_Layer1, LM_Inserts_L11_Playfield, LM_Inserts_L12_Playfield, LM_Inserts_L13_Playfield, LM_Inserts_L14_Parts, LM_Inserts_L14_Playfield, LM_Inserts_L15_Playfield, LM_Inserts_L16_Playfield, LM_Inserts_L17_LeftFlipper1U, LM_Inserts_L17_Parts, _
'	LM_Inserts_L17_Playfield, LM_Inserts_L17_sw22, LM_Inserts_L18_Layer1, LM_Inserts_L18_LeftFlipper1U, LM_Inserts_L18_Parts, LM_Inserts_L18_Playfield, LM_Inserts_L18_sw22, LM_Inserts_L19_Layer1, LM_Inserts_L19_Playfield, LM_Inserts_L2_LeftFlipper, LM_Inserts_L2_LeftFlipperU, LM_Inserts_L2_Playfield, LM_Inserts_L20_Playfield, LM_Inserts_L20_sw22, LM_Inserts_L21_Playfield, LM_Inserts_L22_Playfield, LM_Inserts_L23_LeftFlipper, LM_Inserts_L23_LeftFlipperU, LM_Inserts_L23_Playfield, LM_Inserts_L23_RightFlipper, LM_Inserts_L23_RightFlipperU, LM_Inserts_L24_Parts, LM_Inserts_L24_Playfield, LM_Inserts_L25_Playfield, LM_Inserts_L25_sw39, LM_Inserts_L26_Parts, LM_Inserts_L26_Playfield, LM_Inserts_L26_sw39, LM_Inserts_L27_Parts, LM_Inserts_L27_Playfield, LM_Inserts_L27_sw39, LM_Inserts_L28_Parts, LM_Inserts_L28_Playfield, LM_Inserts_L28_sw39, LM_Inserts_L28_sw44, LM_Inserts_L29_Parts, LM_Inserts_L29_Playfield, LM_Inserts_L29_sw44, LM_Inserts_L29_sw45, LM_Inserts_L3_LeftFlipperU, LM_Inserts_L3_Parts, _
'	LM_Inserts_L3_Playfield, LM_Inserts_L3_RightFlipperU, LM_Inserts_L30_Parts, LM_Inserts_L30_Playfield, LM_Inserts_L30_sw44, LM_Inserts_L30_sw45, LM_Inserts_L30_sw46, LM_Inserts_L31_Playfield, LM_Inserts_L31_sw45, LM_Inserts_L31_sw46, LM_Inserts_L32_Playfield, LM_Inserts_L32_sw46, LM_Inserts_L33_Parts, LM_Inserts_L33_Playfield, LM_Inserts_L33_sw22, LM_Inserts_L34_Playfield, LM_Inserts_L35_Layer1, LM_Inserts_L35_Parts, LM_Inserts_L35_Playfield, LM_Inserts_L36_Layer1, LM_Inserts_L36_Parts, LM_Inserts_L36_Playfield, LM_Inserts_L37_Parts, LM_Inserts_L37_Playfield, LM_Inserts_L38_Parts, LM_Inserts_L38_Playfield, LM_Inserts_L39_Parts, LM_Inserts_L39_Playfield, LM_Inserts_L4_Parts, LM_Inserts_L4_Playfield, LM_Inserts_L4_RightFlipper, LM_Inserts_L4_RightFlipperU, LM_Inserts_L40_Playfield, LM_Inserts_L41_Parts, LM_Inserts_L41_Playfield, LM_Inserts_L41_RightFlipper1U, LM_Inserts_L42_Parts, LM_Inserts_L42_Playfield, LM_Inserts_L42_RightFlipper1U, LM_Inserts_L42_sw40, LM_Inserts_L43_Parts, LM_Inserts_L43_Playfield, _
'	LM_Inserts_L43_sw40, LM_Inserts_L44_Parts, LM_Inserts_L44_Playfield, LM_Inserts_L44_RightFlipper1U, LM_Inserts_L44_sw39, LM_Inserts_L45_Parts, LM_Inserts_L45_Playfield, LM_Inserts_L45_RightFlipper1U, LM_Inserts_L45_sw39, LM_Inserts_L46_Parts, LM_Inserts_L46_Playfield, LM_Inserts_L46_sw39, LM_Inserts_L47_Parts, LM_Inserts_L47_Playfield, LM_Inserts_L47_RightFlipper1U, LM_Inserts_L47_sw39, LM_Inserts_L49_Parts, LM_Inserts_L49_Playfield, LM_Inserts_L5_Playfield, LM_Inserts_L5_RightFlipper, LM_Inserts_L5_RightFlipperU, LM_Inserts_L5_RightSling1, LM_Inserts_L50_GhostTarget, LM_Inserts_L50_Parts, LM_Inserts_L50_Playfield, LM_Inserts_L51_GhostTarget, LM_Inserts_L51_Layer1, LM_Inserts_L51_Parts, LM_Inserts_L51_Playfield, LM_Inserts_L51_sw32, LM_Inserts_L52_Parts, LM_Inserts_L52_Playfield, LM_Inserts_L52_sw32, LM_Inserts_L53_Parts, LM_Inserts_L53_Playfield, LM_Inserts_L54_GhostTarget, LM_Inserts_L54_Parts, LM_Inserts_L54_Playfield, LM_Inserts_L54_sw40, LM_Inserts_L55_GhostTarget, LM_Inserts_L55_Parts, _
'	LM_Inserts_L55_Playfield, LM_Inserts_L55_sw40, LM_Inserts_L56_GhostTarget, LM_Inserts_L56_Parts, LM_Inserts_L56_Playfield, LM_Inserts_L56_sw40, LM_Inserts_L6_Playfield, LM_Inserts_L6_RightFlipperU, LM_Inserts_L6_RightSling1, LM_Inserts_L6_RightSling2, LM_Inserts_L6_RightSling3, LM_Inserts_L6_RightSling4, LM_Inserts_L60_Parts, LM_Inserts_L60_Playfield, LM_Inserts_L60_sw30, LM_Inserts_L60_sw31, LM_Inserts_L60_sw32, LM_Inserts_L61_Playfield, LM_Inserts_L61_sw17, LM_Inserts_L61_sw18, LM_Inserts_L62_Parts, LM_Inserts_L62_Playfield, LM_Inserts_L63_Playfield, LM_Inserts_L7_Playfield, LM_Inserts_L70_Layer1, LM_Inserts_L70_Parts, LM_Inserts_L70_Playfield, LM_Inserts_L70_Post, LM_Inserts_L71_Layer1, LM_Inserts_L71_Layer2, LM_Inserts_L71_Parts, LM_Inserts_L71_Playfield, LM_Inserts_L72_Layer1, LM_Inserts_L72_Layer2, LM_Inserts_L72_Parts, LM_Inserts_L72_Playfield, LM_Inserts_L78_Layer1, LM_Inserts_L78_Parts, LM_Inserts_L78_Playfield, LM_Inserts_L78_sw29, LM_Inserts_L79_Bumper2, LM_Inserts_L79_Bumper3, _
'	LM_Inserts_L79_Layer1, LM_Inserts_L79_Layer2, LM_Inserts_L79_Parts, LM_Inserts_L79_Playfield, LM_Inserts_L79_sw30, LM_Inserts_L79_sw31, LM_Inserts_L8_Parts, LM_Inserts_L8_Playfield, LM_Inserts_L8_sw61, LM_Inserts_L9_Layer1, LM_Inserts_L9_Layer2, LM_Inserts_L9_Parts, LM_Inserts_L9_Playfield, LM_Inserts_L9_sw58, LM_Opto_Parts, LM_Opto_Playfield, LM_Opto_sw39, LM_Spotlights_gi04_Bumper1, LM_Spotlights_gi04_Bumper3, LM_Spotlights_gi04_Dummy, LM_Spotlights_gi04_Gate002, LM_Spotlights_gi04_Gate003, LM_Spotlights_gi04_Gate005, LM_Spotlights_gi04_Gate6, LM_Spotlights_gi04_GhostTarget, LM_Spotlights_gi04_Layer1, LM_Spotlights_gi04_Layer2, LM_Spotlights_gi04_Parts, LM_Spotlights_gi04_Playfield, LM_Spotlights_gi04_RF1, LM_Spotlights_gi04_RF1U, LM_Spotlights_gi04_sw17, LM_Spotlights_gi04_sw18, LM_Spotlights_gi04_sw19, LM_Spotlights_gi04_sw27, LM_Spotlights_gi04_sw28, LM_Spotlights_gi04_sw29, LM_Spotlights_gi04_sw30, LM_Spotlights_gi04_sw31, LM_Spotlights_gi04_sw32, LM_Spotlights_gi04_sw39, _
'	LM_Spotlights_gi04_sw40)
'Dim BG_All: BG_All=Array(BM_Bumper1, BM_Bumper2, BM_Bumper3, BM_Dummy, BM_Gate002, BM_Gate003, BM_Gate004, BM_Gate005, BM_Gate1, BM_Gate3, BM_Gate5, BM_Gate6, BM_Gate7, BM_Gate8, BM_GhostTarget, BM_Layer1, BM_Layer2, BM_LeftFlipper, BM_LeftFlipper1, BM_LeftFlipper1U, BM_LeftFlipperU, BM_LeftSling1, BM_LeftSling2, BM_LeftSling3, BM_LeftSling4, BM_Lemk, BM_Parts, BM_Parts2, BM_PinCab_Rails, BM_Playfield, BM_Post, BM_Remk, BM_RightFlipper, BM_RightFlipper1, BM_RightFlipper1U, BM_RightFlipperU, BM_RightSling1, BM_RightSling2, BM_RightSling3, BM_RightSling4, BM_ScrambledEgg, BM_sw17, BM_sw18, BM_sw19, BM_sw21, BM_sw22, BM_sw25, BM_sw26, BM_sw27, BM_sw28, BM_sw29, BM_sw30, BM_sw31, BM_sw32, BM_sw37, BM_sw39, BM_sw40, BM_sw42, BM_sw43, BM_sw44, BM_sw45, BM_sw46, BM_sw57, BM_sw58, BM_sw60, BM_sw61, BM_swPlunger, LM_Backlight_L65_Layer1, LM_Backlight_L65_Parts, LM_Backlight_L66_Layer1, LM_Backlight_L66_Parts, LM_Backlight_L66_Playfield, LM_Backlight_L66_sw28, LM_Backlight_L67_Dummy, LM_Backlight_L67_Layer2, _
'	LM_Backlight_L67_Parts, LM_Backlight_L67_Playfield, LM_Backlight_L68_Dummy, LM_Backlight_L68_Layer2, LM_Backlight_L68_Parts, LM_Backlight_L68_Playfield, LM_Backlight_L68_sw30, LM_Backlight_L69_Layer1, LM_Backlight_L69_Layer2, LM_Backlight_L69_Parts, LM_Backlight_L69_Playfield, LM_Backlight_L73_Layer1, LM_Backlight_L73_Layer2, LM_Backlight_L73_Parts, LM_Backlight_L73_Playfield, LM_Backlight_L74_Layer1, LM_Backlight_L74_Layer2, LM_Backlight_L74_Parts, LM_Backlight_L75_Layer1, LM_Backlight_L75_Layer2, LM_Backlight_L75_Parts, LM_Backlight_L75_Playfield, LM_Backlight_L76_Bumper3, LM_Backlight_L76_Layer1, LM_Backlight_L76_Layer2, LM_Backlight_L76_Parts, LM_Backlight_L76_Playfield, LM_Backlight_L77_Layer1, LM_Backlight_L77_Parts, LM_DMD_DMX1_Parts, LM_DMD_DMX10_Parts, LM_DMD_DMX100_Parts, LM_DMD_DMX101_Parts, LM_DMD_DMX102_Parts, LM_DMD_DMX103_Parts, LM_DMD_DMX104_Parts, LM_DMD_DMX105_Parts, LM_DMD_DMX11_Parts, LM_DMD_DMX12_Parts, LM_DMD_DMX13_Parts, LM_DMD_DMX14_Parts, LM_DMD_DMX15_Parts, LM_DMD_DMX16_Parts, _
'	LM_DMD_DMX17_Parts, LM_DMD_DMX18_Parts, LM_DMD_DMX19_Parts, LM_DMD_DMX2_Parts, LM_DMD_DMX20_Parts, LM_DMD_DMX21_Parts, LM_DMD_DMX22_Parts, LM_DMD_DMX23_Parts, LM_DMD_DMX24_Parts, LM_DMD_DMX25_Parts, LM_DMD_DMX26_Parts, LM_DMD_DMX27_Parts, LM_DMD_DMX28_Parts, LM_DMD_DMX29_Parts, LM_DMD_DMX3_Parts, LM_DMD_DMX30_Parts, LM_DMD_DMX31_Parts, LM_DMD_DMX32_Parts, LM_DMD_DMX33_Parts, LM_DMD_DMX34_Parts, LM_DMD_DMX35_Parts, LM_DMD_DMX36_Parts, LM_DMD_DMX37_Parts, LM_DMD_DMX38_Parts, LM_DMD_DMX39_Parts, LM_DMD_DMX4_Parts, LM_DMD_DMX40_Parts, LM_DMD_DMX41_Parts, LM_DMD_DMX42_Parts, LM_DMD_DMX43_Parts, LM_DMD_DMX44_Parts, LM_DMD_DMX45_Parts, LM_DMD_DMX46_Parts, LM_DMD_DMX47_Parts, LM_DMD_DMX48_Parts, LM_DMD_DMX49_Parts, LM_DMD_DMX5_Parts, LM_DMD_DMX50_Parts, LM_DMD_DMX51_Parts, LM_DMD_DMX52_Parts, LM_DMD_DMX53_Parts, LM_DMD_DMX54_Parts, LM_DMD_DMX55_Parts, LM_DMD_DMX56_Parts, LM_DMD_DMX57_Parts, LM_DMD_DMX58_Parts, LM_DMD_DMX59_Parts, LM_DMD_DMX6_Parts, LM_DMD_DMX60_Parts, LM_DMD_DMX61_Parts, LM_DMD_DMX62_Parts, _
'	LM_DMD_DMX63_Parts, LM_DMD_DMX64_Parts, LM_DMD_DMX65_Parts, LM_DMD_DMX66_Parts, LM_DMD_DMX67_Parts, LM_DMD_DMX68_Parts, LM_DMD_DMX69_Parts, LM_DMD_DMX7_Parts, LM_DMD_DMX70_Parts, LM_DMD_DMX71_Parts, LM_DMD_DMX72_Parts, LM_DMD_DMX73_Parts, LM_DMD_DMX74_Parts, LM_DMD_DMX75_Parts, LM_DMD_DMX76_Parts, LM_DMD_DMX77_Parts, LM_DMD_DMX78_Parts, LM_DMD_DMX79_Parts, LM_DMD_DMX8_Parts, LM_DMD_DMX80_Parts, LM_DMD_DMX81_Parts, LM_DMD_DMX82_Parts, LM_DMD_DMX83_Parts, LM_DMD_DMX84_Parts, LM_DMD_DMX85_Parts, LM_DMD_DMX86_Parts, LM_DMD_DMX87_Parts, LM_DMD_DMX88_Parts, LM_DMD_DMX89_Parts, LM_DMD_DMX9_Parts, LM_DMD_DMX90_Parts, LM_DMD_DMX91_Parts, LM_DMD_DMX92_Parts, LM_DMD_DMX93_Parts, LM_DMD_DMX94_Parts, LM_DMD_DMX95_Parts, LM_DMD_DMX96_Parts, LM_DMD_DMX97_Parts, LM_DMD_DMX98_Parts, LM_DMD_DMX99_Parts, LM_Flashers_L121_Layer1, LM_Flashers_L121_Layer2, LM_Flashers_L121_LeftFlipper, LM_Flashers_L121_LeftFlipper1, LM_Flashers_L121_LeftFlipper1U, LM_Flashers_L121_LeftFlipperU, LM_Flashers_L121_LeftSling1, _
'	LM_Flashers_L121_LeftSling2, LM_Flashers_L121_LeftSling3, LM_Flashers_L121_LeftSling4, LM_Flashers_L121_Lemk, LM_Flashers_L121_Parts, LM_Flashers_L121_Parts2, LM_Flashers_L121_PinCab_Rails, LM_Flashers_L121_Playfield, LM_Flashers_L121_Remk, LM_Flashers_L121_RightFlipperU, LM_Flashers_L121_RightSling1, LM_Flashers_L121_RightSling2, LM_Flashers_L121_RightSling3, LM_Flashers_L121_RightSling4, LM_Flashers_L121_ScrambledEgg, LM_Flashers_L121_sw22, LM_Flashers_L121_sw42, LM_Flashers_L121_sw43, LM_Flashers_L121_sw57, LM_Flashers_L121_sw58, LM_Flashers_L121_sw61, LM_Flashers_L122_Gate1, LM_Flashers_L122_Layer1, LM_Flashers_L122_Layer2, LM_Flashers_L122_LeftFlipperU, LM_Flashers_L122_LeftSling1, LM_Flashers_L122_LeftSling2, LM_Flashers_L122_LeftSling3, LM_Flashers_L122_LeftSling4, LM_Flashers_L122_Lemk, LM_Flashers_L122_Parts, LM_Flashers_L122_PinCab_Rails, LM_Flashers_L122_Playfield, LM_Flashers_L122_Remk, LM_Flashers_L122_RightFlipper, LM_Flashers_L122_RightFlipperU, LM_Flashers_L122_RightSling1, _
'	LM_Flashers_L122_RightSling2, LM_Flashers_L122_RightSling3, LM_Flashers_L122_RightSling4, LM_Flashers_L122_sw44, LM_Flashers_L122_sw58, LM_Flashers_L122_sw60, LM_Flashers_L122_sw61, LM_Flashers_L122_swPlunger, LM_Flashers_L123_Gate5, LM_Flashers_L123_Layer1, LM_Flashers_L123_LeftFlipper1, LM_Flashers_L123_Parts, LM_Flashers_L123_Playfield, LM_Flashers_L123_RightFlipper1U, LM_Flashers_L123_sw18, LM_Flashers_L123_sw19, LM_Flashers_L123_sw44, LM_Flashers_L123_sw45, LM_Flashers_L123_sw46, LM_Flashers_L127_Bumper2, LM_Flashers_L127_Bumper3, LM_Flashers_L127_Layer1, LM_Flashers_L127_Layer2, LM_Flashers_L127_Parts, LM_Flashers_L127_Playfield, LM_Flashers_L127_sw30, LM_Flashers_L129_Gate5, LM_Flashers_L129_GhostTarget, LM_Flashers_L129_Layer1, LM_Flashers_L129_Parts, LM_Flashers_L129_Playfield, LM_Flashers_L129_RightFlipper1, LM_Flashers_L129_RightFlipper1U, LM_Flashers_L129_sw17, LM_Flashers_L129_sw30, LM_Flashers_L129_sw31, LM_Flashers_L129_sw32, LM_Flashers_L129_sw39, LM_Flashers_L129_sw40, _
'	LM_Flashers_L130_Bumper3, LM_Flashers_L130_GhostTarget, LM_Flashers_L130_Layer1, LM_Flashers_L130_Layer2, LM_Flashers_L130_Parts, LM_Flashers_L130_Playfield, LM_Flashers_L130_RightFlipper1, LM_Flashers_L130_RightFlipper1U, LM_Flashers_L130_sw17, LM_Flashers_L130_sw18, LM_Flashers_L130_sw30, LM_Flashers_L130_sw31, LM_Flashers_L130_sw32, LM_Flashers_L130_sw39, LM_Flashers_L130_sw40, LM_Flashers_L131_Bumper1, LM_Flashers_L131_Bumper3, LM_Flashers_L131_Dummy, LM_Flashers_L131_Gate3, LM_Flashers_L131_GhostTarget, LM_Flashers_L131_Layer1, LM_Flashers_L131_Layer2, LM_Flashers_L131_Parts, LM_Flashers_L131_Playfield, LM_Flashers_L131_RightFlipper1, LM_Flashers_L131_RightFlipper1U, LM_Flashers_L131_sw17, LM_Flashers_L131_sw18, LM_Flashers_L131_sw19, LM_Flashers_L131_sw30, LM_Flashers_L131_sw31, LM_Flashers_L131_sw32, LM_Flashers_L131_sw39, LM_Flashers_L131_sw40, LM_Flashers_L132_Layer1, LM_Flashers_L132_LeftFlipper1, LM_Flashers_L132_LeftFlipper1U, LM_Flashers_L132_Parts, LM_Flashers_L132_Playfield, _
'	LM_Flashers_L132_ScrambledEgg, LM_Flashers_L132_sw22, LM_Flashers_L48_Layer1, LM_Flashers_L48_Parts, LM_Flashers_L48_Playfield, LM_Flashers_L48_RightFlipper1, LM_Flashers_L48_RightFlipper1U, LM_Flashers_L48_sw31, LM_Flashers_L48_sw32, LM_Flashers_L48_sw37, LM_Flashers_L48_sw39, LM_Flashers_L48_sw40, LM_Flashers_L57_Bumper2, LM_Flashers_L57_Layer1, LM_Flashers_L57_Parts, LM_Flashers_L57_Playfield, LM_Flashers_L58_Bumper3, LM_Flashers_L58_Dummy, LM_Flashers_L58_Layer1, LM_Flashers_L58_Layer2, LM_Flashers_L58_Parts, LM_Flashers_L58_Playfield, LM_Flashers_L58_sw17, LM_Flashers_L58_sw30, LM_Flashers_L58_sw31, LM_Flashers_L58_sw32, LM_Flashers_L59_Bumper1, LM_Flashers_L59_Layer1, LM_Flashers_L59_Parts, LM_Flashers_L59_Playfield, LM_Flashers_L59_ScrambledEgg, LM_Flashers_L59_sw30, LM_Gi_Bumper1, LM_Gi_Bumper2, LM_Gi_Dummy, LM_Gi_Gate003, LM_Gi_Gate004, LM_Gi_Gate1, LM_Gi_Gate3, LM_Gi_Gate5, LM_Gi_Gate6, LM_Gi_Gate8, LM_Gi_GhostTarget, LM_Gi_Layer1, LM_Gi_Layer2, LM_Gi_LeftFlipper, LM_Gi_LeftFlipper1, _
'	LM_Gi_LeftFlipper1U, LM_Gi_LeftFlipperU, LM_Gi_LeftSling1, LM_Gi_LeftSling2, LM_Gi_LeftSling3, LM_Gi_LeftSling4, LM_Gi_Parts, LM_Gi_Parts2, LM_Gi_Playfield, LM_Gi_Post, LM_Gi_RightFlipper1, LM_Gi_RightFlipper1U, LM_Gi_RightFlipperU, LM_Gi_RightSling1, LM_Gi_RightSling2, LM_Gi_RightSling3, LM_Gi_RightSling4, LM_Gi_ScrambledEgg, LM_Gi_sw18, LM_Gi_sw22, LM_Gi_sw25, LM_Gi_sw26, LM_Gi_sw27, LM_Gi_sw28, LM_Gi_sw32, LM_Gi_sw39, LM_Gi_sw40, LM_Gi_sw42, LM_Gi_sw44, LM_Gi_sw45, LM_Gi_sw46, LM_Gi_sw58, LM_Gi_sw60, LM_Gi_sw61, LM_Gi1_gi01_Layer1, LM_Gi1_gi01_Layer2, LM_Gi1_gi01_LeftFlipper, LM_Gi1_gi01_LeftFlipperU, LM_Gi1_gi01_LeftSling1, LM_Gi1_gi01_LeftSling2, LM_Gi1_gi01_LeftSling3, LM_Gi1_gi01_LeftSling4, LM_Gi1_gi01_Lemk, LM_Gi1_gi01_Parts, LM_Gi1_gi01_Parts2, LM_Gi1_gi01_Playfield, LM_Gi1_gi01_RightFlipperU, LM_Gi1_gi01_sw58, LM_Gi1_gi02_Layer1, LM_Gi1_gi02_Layer2, LM_Gi1_gi02_LeftFlipper, LM_Gi1_gi02_LeftFlipperU, LM_Gi1_gi02_LeftSling1, LM_Gi1_gi02_LeftSling2, LM_Gi1_gi02_LeftSling3, LM_Gi1_gi02_LeftSling4, _
'	LM_Gi1_gi02_Lemk, LM_Gi1_gi02_Parts, LM_Gi1_gi02_Parts2, LM_Gi1_gi02_Playfield, LM_Gi1_gi02_sw39, LM_Gi1_gi02_sw57, LM_Gi1_gi02_sw58, LM_Gi1_gi03_Layer1, LM_Gi1_gi03_Layer2, LM_Gi1_gi03_LeftFlipper, LM_Gi1_gi03_LeftFlipperU, LM_Gi1_gi03_LeftSling1, LM_Gi1_gi03_LeftSling2, LM_Gi1_gi03_LeftSling3, LM_Gi1_gi03_LeftSling4, LM_Gi1_gi03_Lemk, LM_Gi1_gi03_Parts, LM_Gi1_gi03_Parts2, LM_Gi1_gi03_Playfield, LM_Gi1_gi03_sw39, LM_Gi1_gi03_sw57, LM_Gi1_gi03_sw58, LM_Gi1_gi21_Layer1, LM_Gi1_gi21_LeftFlipperU, LM_Gi1_gi21_Parts, LM_Gi1_gi21_Playfield, LM_Gi1_gi21_Remk, LM_Gi1_gi21_RightFlipper, LM_Gi1_gi21_RightFlipperU, LM_Gi1_gi21_RightSling1, LM_Gi1_gi21_RightSling2, LM_Gi1_gi21_RightSling3, LM_Gi1_gi21_RightSling4, LM_Gi1_gi21_sw60, LM_Gi1_gi21_sw61, LM_Gi1_gi22_Layer1, LM_Gi1_gi22_Parts, LM_Gi1_gi22_Playfield, LM_Gi1_gi22_Remk, LM_Gi1_gi22_RightFlipper, LM_Gi1_gi22_RightFlipperU, LM_Gi1_gi22_RightSling1, LM_Gi1_gi22_RightSling2, LM_Gi1_gi22_RightSling3, LM_Gi1_gi22_RightSling4, LM_Gi1_gi22_sw60, LM_Gi1_gi22_sw61, _
'	LM_Gi1_gi23_Layer1, LM_Gi1_gi23_LeftFlipperU, LM_Gi1_gi23_Parts, LM_Gi1_gi23_Playfield, LM_Gi1_gi23_Remk, LM_Gi1_gi23_RightFlipper, LM_Gi1_gi23_RightFlipperU, LM_Gi1_gi23_RightSling1, LM_Gi1_gi23_RightSling2, LM_Gi1_gi23_RightSling3, LM_Gi1_gi23_RightSling4, LM_Gi1_gi23_sw61, LM_Inserts_L1_LeftFlipper, LM_Inserts_L1_LeftFlipperU, LM_Inserts_L1_LeftSling1, LM_Inserts_L1_Parts, LM_Inserts_L1_Playfield, LM_Inserts_L10_LeftSling1, LM_Inserts_L10_LeftSling2, LM_Inserts_L10_LeftSling3, LM_Inserts_L10_LeftSling4, LM_Inserts_L10_Lemk, LM_Inserts_L10_Parts, LM_Inserts_L10_Playfield, LM_Inserts_L11_Layer1, LM_Inserts_L11_Playfield, LM_Inserts_L12_Playfield, LM_Inserts_L13_Playfield, LM_Inserts_L14_Parts, LM_Inserts_L14_Playfield, LM_Inserts_L15_Playfield, LM_Inserts_L16_Playfield, LM_Inserts_L17_LeftFlipper1U, LM_Inserts_L17_Parts, LM_Inserts_L17_Playfield, LM_Inserts_L17_sw22, LM_Inserts_L18_Layer1, LM_Inserts_L18_LeftFlipper1U, LM_Inserts_L18_Parts, LM_Inserts_L18_Playfield, LM_Inserts_L18_sw22, _
'	LM_Inserts_L19_Layer1, LM_Inserts_L19_Playfield, LM_Inserts_L2_LeftFlipper, LM_Inserts_L2_LeftFlipperU, LM_Inserts_L2_Playfield, LM_Inserts_L20_Playfield, LM_Inserts_L20_sw22, LM_Inserts_L21_Playfield, LM_Inserts_L22_Playfield, LM_Inserts_L23_LeftFlipper, LM_Inserts_L23_LeftFlipperU, LM_Inserts_L23_Playfield, LM_Inserts_L23_RightFlipper, LM_Inserts_L23_RightFlipperU, LM_Inserts_L24_Parts, LM_Inserts_L24_Playfield, LM_Inserts_L25_Playfield, LM_Inserts_L25_sw39, LM_Inserts_L26_Parts, LM_Inserts_L26_Playfield, LM_Inserts_L26_sw39, LM_Inserts_L27_Parts, LM_Inserts_L27_Playfield, LM_Inserts_L27_sw39, LM_Inserts_L28_Parts, LM_Inserts_L28_Playfield, LM_Inserts_L28_sw39, LM_Inserts_L28_sw44, LM_Inserts_L29_Parts, LM_Inserts_L29_Playfield, LM_Inserts_L29_sw44, LM_Inserts_L29_sw45, LM_Inserts_L3_LeftFlipperU, LM_Inserts_L3_Parts, LM_Inserts_L3_Playfield, LM_Inserts_L3_RightFlipperU, LM_Inserts_L30_Parts, LM_Inserts_L30_Playfield, LM_Inserts_L30_sw44, LM_Inserts_L30_sw45, LM_Inserts_L30_sw46, LM_Inserts_L31_Playfield, _
'	LM_Inserts_L31_sw45, LM_Inserts_L31_sw46, LM_Inserts_L32_Playfield, LM_Inserts_L32_sw46, LM_Inserts_L33_Parts, LM_Inserts_L33_Playfield, LM_Inserts_L33_sw22, LM_Inserts_L34_Playfield, LM_Inserts_L35_Layer1, LM_Inserts_L35_Parts, LM_Inserts_L35_Playfield, LM_Inserts_L36_Layer1, LM_Inserts_L36_Parts, LM_Inserts_L36_Playfield, LM_Inserts_L37_Parts, LM_Inserts_L37_Playfield, LM_Inserts_L38_Parts, LM_Inserts_L38_Playfield, LM_Inserts_L39_Parts, LM_Inserts_L39_Playfield, LM_Inserts_L4_Parts, LM_Inserts_L4_Playfield, LM_Inserts_L4_RightFlipper, LM_Inserts_L4_RightFlipperU, LM_Inserts_L40_Playfield, LM_Inserts_L41_Parts, LM_Inserts_L41_Playfield, LM_Inserts_L41_RightFlipper1U, LM_Inserts_L42_Parts, LM_Inserts_L42_Playfield, LM_Inserts_L42_RightFlipper1U, LM_Inserts_L42_sw40, LM_Inserts_L43_Parts, LM_Inserts_L43_Playfield, LM_Inserts_L43_sw40, LM_Inserts_L44_Parts, LM_Inserts_L44_Playfield, LM_Inserts_L44_RightFlipper1U, LM_Inserts_L44_sw39, LM_Inserts_L45_Parts, LM_Inserts_L45_Playfield, _
'	LM_Inserts_L45_RightFlipper1U, LM_Inserts_L45_sw39, LM_Inserts_L46_Parts, LM_Inserts_L46_Playfield, LM_Inserts_L46_sw39, LM_Inserts_L47_Parts, LM_Inserts_L47_Playfield, LM_Inserts_L47_RightFlipper1U, LM_Inserts_L47_sw39, LM_Inserts_L49_Parts, LM_Inserts_L49_Playfield, LM_Inserts_L5_Playfield, LM_Inserts_L5_RightFlipper, LM_Inserts_L5_RightFlipperU, LM_Inserts_L5_RightSling1, LM_Inserts_L50_GhostTarget, LM_Inserts_L50_Parts, LM_Inserts_L50_Playfield, LM_Inserts_L51_GhostTarget, LM_Inserts_L51_Layer1, LM_Inserts_L51_Parts, LM_Inserts_L51_Playfield, LM_Inserts_L51_sw32, LM_Inserts_L52_Parts, LM_Inserts_L52_Playfield, LM_Inserts_L52_sw32, LM_Inserts_L53_Parts, LM_Inserts_L53_Playfield, LM_Inserts_L54_GhostTarget, LM_Inserts_L54_Parts, LM_Inserts_L54_Playfield, LM_Inserts_L54_sw40, LM_Inserts_L55_GhostTarget, LM_Inserts_L55_Parts, LM_Inserts_L55_Playfield, LM_Inserts_L55_sw40, LM_Inserts_L56_GhostTarget, LM_Inserts_L56_Parts, LM_Inserts_L56_Playfield, LM_Inserts_L56_sw40, LM_Inserts_L6_Playfield, _
'	LM_Inserts_L6_RightFlipperU, LM_Inserts_L6_RightSling1, LM_Inserts_L6_RightSling2, LM_Inserts_L6_RightSling3, LM_Inserts_L6_RightSling4, LM_Inserts_L60_Parts, LM_Inserts_L60_Playfield, LM_Inserts_L60_sw30, LM_Inserts_L60_sw31, LM_Inserts_L60_sw32, LM_Inserts_L61_Playfield, LM_Inserts_L61_sw17, LM_Inserts_L61_sw18, LM_Inserts_L62_Parts, LM_Inserts_L62_Playfield, LM_Inserts_L63_Playfield, LM_Inserts_L7_Playfield, LM_Inserts_L70_Layer1, LM_Inserts_L70_Parts, LM_Inserts_L70_Playfield, LM_Inserts_L70_Post, LM_Inserts_L71_Layer1, LM_Inserts_L71_Layer2, LM_Inserts_L71_Parts, LM_Inserts_L71_Playfield, LM_Inserts_L72_Layer1, LM_Inserts_L72_Layer2, LM_Inserts_L72_Parts, LM_Inserts_L72_Playfield, LM_Inserts_L78_Layer1, LM_Inserts_L78_Parts, LM_Inserts_L78_Playfield, LM_Inserts_L78_sw29, LM_Inserts_L79_Bumper2, LM_Inserts_L79_Bumper3, LM_Inserts_L79_Layer1, LM_Inserts_L79_Layer2, LM_Inserts_L79_Parts, LM_Inserts_L79_Playfield, LM_Inserts_L79_sw30, LM_Inserts_L79_sw31, LM_Inserts_L8_Parts, LM_Inserts_L8_Playfield, _
'	LM_Inserts_L8_sw61, LM_Inserts_L9_Layer1, LM_Inserts_L9_Layer2, LM_Inserts_L9_Parts, LM_Inserts_L9_Playfield, LM_Inserts_L9_sw58, LM_Opto_Parts, LM_Opto_Playfield, LM_Opto_sw39, LM_Spotlights_gi04_Bumper1, LM_Spotlights_gi04_Bumper3, LM_Spotlights_gi04_Dummy, LM_Spotlights_gi04_Gate002, LM_Spotlights_gi04_Gate003, LM_Spotlights_gi04_Gate005, LM_Spotlights_gi04_Gate6, LM_Spotlights_gi04_GhostTarget, LM_Spotlights_gi04_Layer1, LM_Spotlights_gi04_Layer2, LM_Spotlights_gi04_Parts, LM_Spotlights_gi04_Playfield, LM_Spotlights_gi04_RF1, LM_Spotlights_gi04_RF1U, LM_Spotlights_gi04_sw17, LM_Spotlights_gi04_sw18, LM_Spotlights_gi04_sw19, LM_Spotlights_gi04_sw27, LM_Spotlights_gi04_sw28, LM_Spotlights_gi04_sw29, LM_Spotlights_gi04_sw30, LM_Spotlights_gi04_sw31, LM_Spotlights_gi04_sw32, LM_Spotlights_gi04_sw39, LM_Spotlights_gi04_sw40)
' VLM  Arrays - End

' VLM T Arrays - Start
' Arrays per baked part
Dim BP_TTParts: BP_TTParts=Array(TBM_TParts, TLM_TLett_T10S_TParts, TLM_TLett_T11T_TParts, TLM_TLett_T12E_TParts, TLM_TLett_T13R_TParts, TLM_TLett_T1R_TParts, TLM_TLett_T2O_TParts, TLM_TLett_T3L_TParts, TLM_TLett_T4L_TParts, TLM_TLett_T5E_TParts, TLM_TLett_T6R_TParts, TLM_TLett_T7C_TParts, TLM_TLett_T8O_TParts, TLM_TLett_T9A_TParts, TLM_TFlash_L129_TParts, TLM_TFlash_L130_TParts, TLM_TFlash_L131_TParts, TLM_TGi_TParts)
' Arrays per lighting scenario
Dim BL_TLit_Room: BL_TLit_Room=Array(TBM_TParts)
Dim BL_TTFlash_L129: BL_TTFlash_L129=Array(TLM_TFlash_L129_TParts)
Dim BL_TTFlash_L130: BL_TTFlash_L130=Array(TLM_TFlash_L130_TParts)
Dim BL_TTFlash_L131: BL_TTFlash_L131=Array(TLM_TFlash_L131_TParts)
Dim BL_TTGi: BL_TTGi=Array(TLM_TGi_TParts)
Dim BL_TTLett_T10S: BL_TTLett_T10S=Array(TLM_TLett_T10S_TParts)
Dim BL_TTLett_T11T: BL_TTLett_T11T=Array(TLM_TLett_T11T_TParts)
Dim BL_TTLett_T12E: BL_TTLett_T12E=Array(TLM_TLett_T12E_TParts)
Dim BL_TTLett_T13R: BL_TTLett_T13R=Array(TLM_TLett_T13R_TParts)
Dim BL_TTLett_T1R: BL_TTLett_T1R=Array(TLM_TLett_T1R_TParts)
Dim BL_TTLett_T2O: BL_TTLett_T2O=Array(TLM_TLett_T2O_TParts)
Dim BL_TTLett_T3L: BL_TTLett_T3L=Array(TLM_TLett_T3L_TParts)
Dim BL_TTLett_T4L: BL_TTLett_T4L=Array(TLM_TLett_T4L_TParts)
Dim BL_TTLett_T5E: BL_TTLett_T5E=Array(TLM_TLett_T5E_TParts)
Dim BL_TTLett_T6R: BL_TTLett_T6R=Array(TLM_TLett_T6R_TParts)
Dim BL_TTLett_T7C: BL_TTLett_T7C=Array(TLM_TLett_T7C_TParts)
Dim BL_TTLett_T8O: BL_TTLett_T8O=Array(TLM_TLett_T8O_TParts)
Dim BL_TTLett_T9A: BL_TTLett_T9A=Array(TLM_TLett_T9A_TParts)
'' Global arrays
'Dim BGT_Bakemap: BGT_Bakemap=Array(TBM_TParts)
'Dim BGT_Lightmap: BGT_Lightmap=Array(TLM_TFlash_L129_TParts, TLM_TFlash_L130_TParts, TLM_TFlash_L131_TParts, TLM_TGi_TParts, TLM_TLett_T10S_TParts, TLM_TLett_T11T_TParts, TLM_TLett_T12E_TParts, TLM_TLett_T13R_TParts, TLM_TLett_T1R_TParts, TLM_TLett_T2O_TParts, TLM_TLett_T3L_TParts, TLM_TLett_T4L_TParts, TLM_TLett_T5E_TParts, TLM_TLett_T6R_TParts, TLM_TLett_T7C_TParts, TLM_TLett_T8O_TParts, TLM_TLett_T9A_TParts)
Dim BGT_All: BGT_All=Array(TBM_TParts, TLM_TFlash_L129_TParts, TLM_TFlash_L130_TParts, TLM_TFlash_L131_TParts, TLM_TGi_TParts, TLM_TLett_T10S_TParts, TLM_TLett_T11T_TParts, TLM_TLett_T12E_TParts, TLM_TLett_T13R_TParts, TLM_TLett_T1R_TParts, TLM_TLett_T2O_TParts, TLM_TLett_T3L_TParts, TLM_TLett_T4L_TParts, TLM_TLett_T5E_TParts, TLM_TLett_T6R_TParts, TLM_TLett_T7C_TParts, TLM_TLett_T8O_TParts, TLM_TLett_T9A_TParts)
'' VLM T Arrays - End

Const FlasherOp1 = 150
Const FlasherOp2 = 60

SetFlasherOpacity
Sub SetFlasherOpacity
	Dim BL
	For each BL in BL_Flashers_L121: Bl.opacity = FlasherOp1: next
	For each BL in BL_Flashers_L122: Bl.opacity = FlasherOp1: next
	For each BL in BL_Flashers_L123: Bl.opacity = FlasherOp1: next
	For each BL in BL_Flashers_L127: Bl.opacity = FlasherOp1: next
	For each BL in BL_Flashers_L129: Bl.opacity = FlasherOp1: next
	For each BL in BL_Flashers_L130: Bl.opacity = FlasherOp1: next
	For each BL in BL_Flashers_L131: Bl.opacity = FlasherOp1: next
	For each BL in BL_Flashers_L132: Bl.opacity = FlasherOp1: next
'	For each BL in BL_Flashers_L48: Bl.opacity = FlasherOp1: next
	For each BL in BL_Flashers_L57: Bl.opacity = FlasherOp1: next
	For each BL in BL_Flashers_L58: Bl.opacity = FlasherOp1: next
	For each BL in BL_Flashers_L59: Bl.opacity = FlasherOp1: next
	For each BL in BL_TTLett_T1R: Bl.opacity = FlasherOp2: next
	For each BL in BL_TTLett_T2O: Bl.opacity = FlasherOp2: next
	For each BL in BL_TTLett_T3L: Bl.opacity = FlasherOp2: next
	For each BL in BL_TTLett_T4L: Bl.opacity = FlasherOp2: next
	For each BL in BL_TTLett_T5E: Bl.opacity = FlasherOp2: next
	For each BL in BL_TTLett_T6R: Bl.opacity = FlasherOp2: next
	For each BL in BL_TTLett_T7C: Bl.opacity = FlasherOp2: next
	For each BL in BL_TTLett_T8O: Bl.opacity = FlasherOp2: next
	For each BL in BL_TTLett_T9A: Bl.opacity = FlasherOp2: next
	For each BL in BL_TTLett_T10S: Bl.opacity = FlasherOp2: next
	For each BL in BL_TTLett_T11T: Bl.opacity = FlasherOp2: next
	For each BL in BL_TTLett_T12E: Bl.opacity = FlasherOp2: next
	For each BL in BL_TTLett_T13R: Bl.opacity = FlasherOp2: next
End Sub


	

'************
' Load stuff
'************

Const cGameName = "rctycn"


Const UseVPMModSol = 2 
Const UseSolenoids = 2
Const UseLamps = 1
Const UseGI = 0
Const UseSync = 0 'set it to 1 if the table runs too fast
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoidon"
Const SSolenoidOff = "fx_Solenoidoff"
Const SFlipperOn = ""
Const SFlipperOff = ""
Const SCoin = ""

LoadVPM "03060000", "sega.vbs", 3.10


'************
' Table init.
'************

Dim SpinnerBall, RCTBall1, RCTBall2, RCTBall3, RCTBall4, gBOT

Sub table1_Init
	sw37.enabled = 0
	Op01.state = 1
	Op02.state = 1
	Diverter_left.Open = 1
	Diverter_Right.Open = 1
	Gate1.Collidable = 0

    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "RollerCoaster Tycoon - Stern 2002" & vbNewLine & "VPX table by JPSalas v4.0.1"
        .Games(cGameName).Settings.Value("sound") = 1
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
        '.SetDisplayPosition 0,0, GetPlayerHWnd 'restore dmd window position
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
        .Switch(36) = 1
        .Switch(38) = 1
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 56
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)


	'Trough
	Set RCTBall1 = sw14.CreateSizedballWithMass(Ballsize/2,Ballmass)
	Set RCTBall2 = sw13.CreateSizedballWithMass(Ballsize/2,Ballmass)
	Set RCTBall3 = sw12.CreateSizedballWithMass(Ballsize/2,Ballmass)
	Set RCTBall4 = sw11.CreateSizedballWithMass(Ballsize/2,Ballmass)

	gBOT = Array(RCTBall1,RCTBall2,RCTBall3,RCTBall4)


	Controller.Switch(11) = 1
	Controller.Switch(12) = 1
	Controller.Switch(13) = 1
	Controller.Switch(14) = 1

'	'Spinner
	Set SpinnerBall = SpinnerKick.CreateSizedballWithMass(65/2,Ballmass*0.2)
	SpinnerBall.visible = False
	Spinnerkick.kick 0,0,0
	Spinnerkick.enabled = False

    ' Impulse Plunger - used as the autoplunger
    Const IMPowerSetting = 39 ' Plunger Power
    Const IMTime = 0.7        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swPlunger, IMPowerSetting, IMTime
        .Random 0.3
        .switch 16
        .InitExitSnd SoundFX("", DOFContactors), SoundFX("", DOFContactors)
        .CreateEvents "plungerIM"
    End With

	vpmMapLights InsertLamps  		' Map all lamps to the corresponding ROM output using the value of TimerInterval of each light object


    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    SolDiverterRight 0
    SolLock 0

	'Initialize slings
	RStep = 0:RightSlingShot.Timerenabled=True
	LStep = 0:LeftSlingShot.Timerenabled=True

	Dim BP: For each BP in BP_ScrambledEgg: BP.transz=1: Next


End Sub

Sub table1_Paused:		Controller.Pause = 1: End Sub
Sub table1_unPaused:	Controller.Pause = 0: End Sub
Sub table1_exit:		Controller.Stop: End Sub



'******************
' Timers
'******************

' FIXME for the time being, the cor timer interval must be 10 ms (so below 60FPS framerate)
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub

' The frame timer interval is -1, so executes at the display frame rate
Sub FrameTimer_Timer()
	BSUpdate 'update ball shadows
	UpdateBallBrightness
	RollingUpdate					'update rolling sounds
	DoDTAnim 						'handle drop target animations
	UpdateDropTargets
	DoSTAnim
	UpdateStandupTargets
	UpdateLeds
End Sub

' Hack to return Narnia ball back in play
Sub Narnia_Timer
    Dim b
	For b = 0 to UBound(gBOT)
		if gBOT(b).z < -300 Then
			'msgbox "Ball " &b& " in Narnia X: " & gBOT(b).x &" Y: "&gBOT(b).y & " Z: "&gBOT(b).z
			'debug.print "Move narnia ball ("& gBOT(b).x &" Y: "&gBOT(b).y & " Z: "&gBOT(b).z&") to right scoop"
			gBOT(b).x = 749
			gBOT(b).y = 1161
			gBOT(b).z = -50
		end if
	next
end sub


'******************
' User Options
'******************



'*******************************************
'  ZOPT: User Options
'*******************************************

'Do not adjust the options here (it won't work). Use the in game options menu.
Dim RefractOpt : RefractOpt = 1				'0 - No Refraction (best performance), 1 - Sharp Refractions (improved performance), 2 - Rough Refractions (best visual)
Dim LightLevel : LightLevel = 0.40          'LightLevel - Value between 0 and 1 (0=Dark ... 1=Bright)
Dim ColorLUT : ColorLUT = 1					'LUT saturation option
Dim RecordedVolume : RecordedVolume = 0.6   'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim CoasterVolume : CoasterVolume = 0.6   'Coaster sound effect volume. Recommended values should be no greater than 1.
Dim VolumeDial : VolumeDial = 0.7			'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5 	'Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5 	'Level of ramp rolling volume. Value between 0 and 1
Dim BackglassBuzz : BackglassBuzz = 0 		'Backglass buzz sound (0-off, 1-On)
Dim PlungerVis : PlungerVis = 1				'Plunger position visualization (0-off, 1-On)
Dim VRRoomChoice: VRRoomChoice = 2 			
Dim RailVis: RailVis=1

' Called when options are tweaked by the player. 
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are: 
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Dim dspTriggered : dspTriggered = False
Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If
	Dim v

	RefractOpt = Table1.Option("Refraction Setting", 0, 2, 1, 2, 0, Array("No Refraction (best performance)", "Sharp Refractions (improved performance)", "Rough Refractions (best visuals)"))
	SetRefractionProbes RefractOpt

	'Playfield Reflections
	v = Table1.Option("Playfield Reflections", 0, 3, 1, 1, 0, Array("Off", "Ball Only", "Static", "Dynamic"))
	ReflectionToggle(v)

	'Siderail Visibility
	RailVis = Table1.Option("Siderails", 0, 1, 1, 1, 0, Array("Not Visible", "Visible"))

	'Ball image
	v = Table1.Option("Ball Image", 0, 1, 1, 1, 0, Array("Standard","RCT Room"))
	Select Case v
		Case 0: RCTBall1.image = "ball_Test8B": RCTBall2.image = "ball_Test8B": RCTBall3.image = "ball_Test8B": RCTBall4.image = "ball_Test8B"
		Case 1: RCTBall1.image = "RCT_Ball_VR": RCTBall2.image = "RCT_Ball_VR": RCTBall3.image = "RCT_Ball_VR": RCTBall4.image = "RCT_Ball_VR"
	End Select

    ' Glass	dirt
	v = Table1.Option("Glass dirt", 0, 4, 1, 3, 0, Array("Hide", "25%", "50%", "75%", "100%"))
	F_Glass_spec.Visible = v>0
	F_Glass_dif.Visible = v>0
	Select Case v
		Case 1: F_Glass_spec.Opacity = 50000*0.25: F_Glass_dif.Opacity = 150*0.25
		Case 2: F_Glass_spec.Opacity = 50000*0.50: F_Glass_dif.Opacity = 150*0.50
		Case 3: F_Glass_spec.Opacity = 50000*0.75: F_Glass_dif.Opacity = 150*0.75
		Case 4: F_Glass_spec.Opacity = 50000*1.00: F_Glass_dif.Opacity = 150*1.00
	End Select

    ' Sound volumes
    RecordedVolume = Table1.Option("RCT Mech Volume", 0, 1, 0.01, 0.4, 1)
    CoasterVolume = Table1.Option("RCT Coaster Volume", 0, 1, 0.01, 0.4, 1)
	BackglassBuzz = Table1.Option("Backglass Attract Buzz", 0, 1, 1, 0, 0, Array("Off", "On"))
    VolumeDial = Table1.Option("Other Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
	RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

	If BackglassBuzz = 1 Then
		PlaySoundAtVolLoops ("z_RCT_Attract"), RelaySound , RecordedVolume , -1
	Else
		StopSound "z_RCT_Attract"
	End If

	' Room brightness
'	LightLevel = NightDay/100
	LightLevel = Table1.Option("Table Brightness (Ambient Light Level)", 0, 1, 0.01, .5, 1)
	SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.

    ' Plunger Position Visualization
    PlungerVis = Table1.Option("Plunger Position Visualization", 0, 1, 1, 1, 0, Array("Disabled", "Enabled"))
	PlungerLine.visible = PlungerVis

	'VR Room
	VRRoomChoice = Table1.Option("VR Room", 0, 2, 1, 2, 0, Array("Ultra Minimal", "Minimal", "Coaster"))
	SetupRoom

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub


Sub SetupRoom
	Dim Thing, BP, ball
	TimerPlunger2.enabled = true
	' VR Mode
	If RenderingMode = 2 OR TestVRonDT = True Then
		If VRRoomChoice = 2 Then 					' Coaster
			F_Glass_spec.rotx = - 7.85
			F_Glass_dif.rotx = - 7.85
			F_Glass_spec.height = 284
			F_Glass_dif.height = 284
			VRATimer.enabled = true
			VRTimer.enabled = true
			BM_PinCab_Rails.visible = false
			For each Thing in VRCoasterRoom: Thing.visible = true: next
			For each Thing in VRCab: Thing.visible = true: next
			For each BP in BGT_All: BP.visible = true: next
			MiniRoom.visible = false
			MiniRoom_Wall.visible = false
		For Each BP in BP_PinCab_Rails : BP.visible = False: Next

		ElseIf VRRoomChoice = 1 Then  	' Minimal
			MiniRoom.visible = true
			MiniRoom_Wall.visible = true
			F_Glass_spec.rotx = - 7.85
			F_Glass_dif.rotx = - 7.85
			F_Glass_spec.height = 284
			F_Glass_dif.height = 284
			'Turn off VRRoom = 1 stuff
			VRATimer.enabled = false
			VRTimer.enabled = false
			BM_PinCab_Rails.visible = false
			For each Thing in VRCoasterRoom: Thing.visible = false: next
			For each Thing in VRCab: Thing.visible = true: next
			For each BP in BGT_All: BP.visible = false: next
			StopSound "Train_Arrive_Fix"
			StopSound "Train_Leave_Fix"
			'Turn on VRRoom = 2 stuff
			DMD.visible = true
		For Each BP in BP_PinCab_Rails : BP.visible = False: Next
		Else 					' Ultra Minimal
			MiniRoom.visible = false
			MiniRoom_Wall.visible = false
		For Each BP in BP_PinCab_Rails : BP.visible = RailVis: Next
			For each Thing in VRCoasterRoom: Thing.visible = false: next
			For each Thing in VRCab: Thing.visible = false: next
			For each BP in BGT_All: BP.visible = false: next
		End If
	Else
			MiniRoom.visible = false
			MiniRoom_Wall.visible = false
		For Each BP in BP_PinCab_Rails : BP.visible = RailVis: Next
			For each Thing in VRCoasterRoom: Thing.visible = false: next
			For each Thing in VRCab: Thing.visible = false: next
			For each BP in BGT_All: BP.visible = false: next
	End If
End Sub


Sub SetRefractionProbes(Opt)
	On Error Resume Next
		Select Case Opt
			Case 0:
				BM_Layer2.RefractionProbe = ""
				BM_Layer1.RefractionProbe = ""
			Case 1:
				BM_Layer2.RefractionProbe = "Refractions"
				BM_Layer1.RefractionProbe = "Refractions"
			Case 2:
				BM_Layer2.RefractionProbe = "Refractions Rough"
				BM_Layer1.RefractionProbe = "Refractions Rough"
		End Select
	On Error Goto 0
End Sub

Function ReflectionToggle(state)
	Dim ReflFullObjArray: ReflFullObjArray = Array(BP_Parts, BP_LeftFlipper, BP_LeftFlipperU, BP_LeftFlipper1, BP_LeftFlipper1U, BP_RightFlipper, BP_RightFlipperU, BP_RightFlipper1, BP_RightFlipper1U, BP_GhostTarget, BP_sw17, BP_sw18, BP_sw19, BP_sw22, BP_sw40, BP_sw44, BP_sw45, BP_sw46, BP_sw28, BP_sw29, BP_sw30, BP_sw31, BP_sw32, BP_sw39, BP_ScrambledEgg, BP_Layer1, BP_Layer2) 
	Dim ReflPartObjArray: ReflPartObjArray = Array(BM_Parts, BM_LeftFlipper, BM_LeftFlipperU, BM_LeftFlipper1, BM_LeftFlipper1U, BM_RightFlipper, BM_RightFlipperU, BM_RightFlipper1, BM_RightFlipper1U, BM_GhostTarget, BM_sw17, BM_sw18, BM_sw19, BM_sw22, BM_sw40, BM_sw44, BM_sw45, BM_sw46, BM_sw28, BM_sw29, BM_sw30, BM_sw31, BM_sw32, BM_sw39, BM_ScrambledEgg, BM_Layer1, BM_Layer2) 
	Dim IBP, BP, v1, v2
	Select Case state
		Case 0: 'No reflections
			v1 = false
			v2 = false 
			BM_Playfield.ReflectionProbe = ""
		Case 1: 'Ball Only reflection
			v1 = false
			v2 = false 
			BM_Playfield.ReflectionProbe = "Playfield Reflections"			
		Case 2: 'Only reflect static prims (that matters as defined in ReflPartObjArray)
			v1 = false
			v2 = true 
			BM_Playfield.ReflectionProbe = "Playfield Reflections"			
		Case 3: 'Reflect everything (that matters as defined in ReflFullObjArray)
			v1 = true
			v2 = true
			BM_Playfield.ReflectionProbe = "Playfield Reflections"
	End Select

	For Each IBP in ReflFullObjArray
		For Each BP in IBP
			BP.ReflectionEnabled = v1
		Next
	Next

	For Each BP in ReflPartObjArray
		BP.ReflectionEnabled = v2
	Next

	BM_Playfield.ReflectionEnabled = False
End Function



'****************************
'	Room Brightness
'****************************

' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid","_noXtraShadingVR","_noXtraShadingVRLight", _
          "Plastic Black","Plastic Red","Metal Black","Metal","Metal Wire","Rubber Black","Plastic","Metal Dark") 

Sub SetRoomBrightness(lvl)
	If lvl > 1 Then lvl = 1
	If lvl < 0 Then lvl = 0

	' Lighting level
	Dim v: v=(lvl * 245 + 10)/255

	Dim i: For i = 0 to UBound(RoomBrightnessMtlArray)
		ModulateMaterialBaseColor RoomBrightnessMtlArray(i), i, v
	Next
End Sub

Dim SavedMtlColorArray
SaveMtlColors
Sub SaveMtlColors
	ReDim SavedMtlColorArray(UBound(RoomBrightnessMtlArray))
	Dim i: For i = 0 to UBound(RoomBrightnessMtlArray)
		SaveMaterialBaseColor RoomBrightnessMtlArray(i), i
	Next
End Sub

Sub SaveMaterialBaseColor(name, idx)
	Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
	GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
	SavedMtlColorArray(idx) = round(base,0)
End Sub


Sub ModulateMaterialBaseColor(name, idx, val)
	Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
	Dim red, green, blue, saved_base, new_base
 
	'First get the existing material properties
	GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle

	'Get saved color
	saved_base = SavedMtlColorArray(idx)
    
	'Next extract the r,g,b values from the base color
	red = saved_base And &HFF
	green = (saved_base \ &H100) And &HFF
	blue = (saved_base \ &H10000) And &HFF
	'msgbox red & " " & green & " " & blue

	'Create new color scaled down by 'val', and update the material
	new_base = RGB(red*val, green*val, blue*val)
	UpdateMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, new_base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
End Sub



'******************************************************
' 	BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  1         'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)
Const PLOffset = 0.3
Dim PLGain: PLGain = (1-PLOffset)/(1400-2000)


Sub UpdateBallBrightness
	Dim s, b_base, b_r, b_g, b_b, d_w
	b_base = 170 * BallBrightness + 45*LightLevel + 40*gilvl

	For s = 0 To UBound(gBOT)
		' Handle z direction
		d_w = b_base*(1 - (gBOT(s).z-25)/500)
		If d_w < 30 Then d_w = 30
		' Handle right plunger lane
		If InRect(gBOT(s).x,gBOT(s).y,870,2000,870,1400,930,1400,930,2000) Then 
			d_w = d_w*(PLOffset+PLGain*(gBOT(s).y-2000))
		End If
		' Handle left plunger lane
		If InRect(gBOT(s).x,gBOT(s).y,20,2000,20,1400,85,1400,85,2000) Then 
			d_w = d_w*(PLOffset+PLGain*(gBOT(s).y-2000))
		End If
		' Assign color
		b_r = Int(d_w)
		b_g = Int(d_w)
		b_b = Int(d_w)
		If b_r > 255 Then b_r = 255
		If b_g > 255 Then b_g = 255
		If b_b > 255 Then b_b = 255
		gBOT(s).color = b_r + (b_g * 256) + (b_b * 256 * 256)
		'debug.print "--- ball.color level="&b_r
	Next
End Sub




'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)

	If keycode = RightFlipperKey then VR_FlipperRight.x = VR_FlipperRight.x - 5
	If keycode = LeftFlipperKey then VR_FlipperLeft.x = VR_FlipperLeft.x + 5
	If keycode = StartGameKey then StartButton.y = StartButton.y - 3

    If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
		Select Case Int(rnd*3)
			Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
			Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
			Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
		End Select
	End If
    If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress End If
    If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress End If
	If keycode = LeftTiltKey Then Nudge 90, 1 : SoundNudgeLeft
	If keycode = RightTiltKey Then Nudge 270, 1 : SoundNudgeRight
	If keycode = CenterTiltKey Then Nudge 0, 1 : SoundNudgeCenter
    If keycode = PlungerKey Then
		SoundPlungerPull()
		Plunger.Pullback
		'if VRRoom = 1 Then
			TimerPlunger.Enabled = True
			TimerPlunger2.Enabled = False
		    PinCab_Shooter.TransY = 0
		'end If
	End If
    If vpmKeyDown(keycode)Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
	If keycode = RightFlipperKey then VR_FlipperRight.x = VR_FlipperRight.x + 5
	If keycode = LeftFlipperKey  then VR_FlipperLeft.x = VR_FlipperLeft.x - 5
	If keycode = StartGameKey  then StartButton.y = StartButton.y + 3

	If keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress End If 
	If keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress End If

    If keycode = PlungerKey Then 
		SoundPlungerReleaseBall()
		Plunger.Fire
		'if VRRoom = 1 Then
			TimerPlunger.Enabled = False
			TimerPlunger2.Enabled = True
		'end If
	End If
    If vpmKeyUp(keycode)Then Exit Sub
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
	LS.VelocityCorrect(ActiveBall)
PlaySoundAtVol SoundFX("z_RCT_Sling_Left", DOFContactors), BM_Lemk, RecordedVolume 
    LStep = 0
    vpmTimer.PulseSw 59
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
	Dim BP
	Dim v1, v2, v3, v4, x
	v1 = False: v2 = False: v3 = False: v4 = True: x = -30
    Select Case LStep
        Case 2:v1 = False: v2 = False: v3 = True:  v4 = False: x = -20
        Case 3:v1 = False: v2 = True:  v3 = False: v4 = False: x = -10
        Case 4:v1 = True:  v2 = False: v3 = False: v4 = False: x = 0: LeftSlingShot.TimerEnabled = 0
    End Select

	For Each BP in BP_Leftsling1 : BP.Visible = v1: Next
	For Each BP in BP_Leftsling2 : BP.Visible = v2: Next
	For Each BP in BP_Leftsling3 : BP.Visible = v3: Next
	For Each BP in BP_Leftsling4 : BP.Visible = v4: Next
	For Each BP in BP_Lemk : BP.transx = x: Next

    LStep = LStep + 1
End Sub


Sub RightSlingShot_Slingshot
	RS.VelocityCorrect(ActiveBall)
PlaySoundAtVol SoundFX("z_RCT_Sling_Right", DOFContactors), BM_Remk, RecordedVolume
    RStep = 0
    vpmTimer.PulseSw 62
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
	Dim BP
	Dim v1, v2, v3, v4, x
	v1 = False: v2 = False: v3 = False: v4 = True: x = -30
    Select Case RStep
        Case 2:v1 = False: v2 = False: v3 = True:  v4 = False: x = -20
        Case 3:v1 = False: v2 = True:  v3 = False: v4 = False: x = -10
        Case 4:v1 = True:  v2 = False: v3 = False: v4 = False: x = 0: RightSlingShot.TimerEnabled = 0
    End Select

	For Each BP in BP_Rightsling1 : BP.Visible = v1: Next
	For Each BP in BP_Rightsling2 : BP.Visible = v2: Next
	For Each BP in BP_Rightsling3 : BP.Visible = v3: Next
	For Each BP in BP_Rightsling4 : BP.Visible = v4: Next
	For Each BP in BP_Remk : BP.transx = x: Next

    RStep = RStep + 1
End Sub



' Bumpers
Sub Bumper1_Hit
	vpmTimer.PulseSw 51
	PlaySoundAtVol SoundFX("z_RCT_Bumper_Bottom", DOFContactors), Bumper1, RecordedVolume
End Sub

Sub Bumper2_Hit
	vpmTimer.PulseSw 49
	PlaySoundAtVol SoundFX("z_RCT_Bumper_Left", DOFContactors), Bumper2, RecordedVolume
End Sub

Sub Bumper3_Hit
	vpmTimer.PulseSw 50
	PlaySoundAtVol SoundFX("z_RCT_Bumper_Right", DOFContactors), Bumper3, RecordedVolume
End Sub

' Bumper strength variation hack
Const DebugBumpers=False

'Bumper1
dim L59on, L59count: L59count = 0
Bumper1.TimerInterval = 1500: Bumper1.TimerEnabled = False
Sub L59_animate  
	If L59.state > 0.4 and L59on = False Then 
		L59on = True
		L59count = L59count + 1
		If L59count > 1 Then
			Bumper1.Force = 15
			If DebugBumpers Then debug.print "Bumper Forces = "&Bumper1.Force&"  "&Bumper2.Force&"  "&Bumper3.Force
		End If
	Elseif L59.state < 0.3 and L59on = True Then
		L59on = False
		Bumper1.TimerEnabled = False: Bumper1.TimerEnabled = True
	End If
End Sub
Sub Bumper1_Timer
	Bumper1.TimerEnabled = False
	L59count = 0
	If L59on Then
		Bumper1.Force = 13
	Else
		Bumper1.Force = 11
	End If
	If DebugBumpers Then debug.print "Bumper Forces = "&Bumper1.Force&"  "&Bumper2.Force&"  "&Bumper3.Force
End Sub

'Bumper2
dim L57on, L57count: L57count = 0
Bumper2.TimerInterval = 1500: Bumper2.TimerEnabled = False
Sub L57_animate  'bumper1
	If L57.state > 0.4 and L57on = False Then 
		L57on = True
		L57count = L57count + 1
		If L57count > 1 Then
			Bumper2.Force = 15
			If DebugBumpers Then debug.print "Bumper Forces = "&Bumper1.Force&"  "&Bumper2.Force&"  "&Bumper3.Force
		End If
	Elseif L57.state < 0.3 and L57on = True Then
		L57on = False
		Bumper2.TimerEnabled = False: Bumper2.TimerEnabled = True
	End If
End Sub
Sub Bumper2_Timer
	Bumper2.TimerEnabled = False
	L57count = 0
	If L57on Then
		Bumper2.Force = 13
	Else
		Bumper2.Force = 11
	End If
	If DebugBumpers Then debug.print "Bumper Forces = "&Bumper1.Force&"  "&Bumper2.Force&"  "&Bumper3.Force
End Sub


'Bumper3
dim L58on, L58count: L58count = 0
Bumper3.TimerInterval = 1500: Bumper3.TimerEnabled = False
Sub L58_animate  'bumper1
	If L58.state > 0.4 and L58on = False Then 
		L58on = True
		L58count = L58count + 1
		If L58count > 1 Then
			Bumper3.Force = 15
			If DebugBumpers Then debug.print "Bumper Forces = "&Bumper1.Force&"  "&Bumper2.Force&"  "&Bumper3.Force
		End If
	Elseif L58.state < 0.3 and L58on = True Then
		L58on = False
		Bumper3.TimerEnabled = False: Bumper3.TimerEnabled = True
	End If
End Sub
Sub Bumper3_Timer
	Bumper3.TimerEnabled = False
	L58count = 0
	If L58on Then
		Bumper3.Force = 13
	Else
		Bumper3.Force = 11
	End If
	If DebugBumpers Then debug.print "Bumper Forces = "&Bumper1.Force&"  "&Bumper2.Force&"  "&Bumper3.Force
End Sub



' Rollovers
Sub sw57_Hit:Controller.Switch(57) = 1:End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:End Sub
Sub sw61_UnHit:Controller.Switch(61) = 0:End Sub

Sub sw60_Hit:Controller.Switch(60) = 1:End Sub
Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub

Sub sw21_Hit:Controller.Switch(21) = 1:End Sub
Sub sw21_UnHit:Controller.Switch(21) = 0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub

Sub sw26_Hit:Controller.Switch(26) = 1:End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

Sub sw37_Hit:Controller.Switch(37) = 1:End Sub
Sub sw37_UnHit:Controller.Switch(37) = 0:End Sub

Sub sw37trigger_hit
sw37.enabled = 1
End sub

Sub sw37trigger0_hit
sw37.enabled = 0
End sub

Sub sw24_Hit:Controller.Switch(24) = 1:End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub

Sub sw33_Hit:Controller.Switch(33) = 1:End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub

Sub sw34_Hit:Controller.Switch(34) = 1:End Sub
Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub

Sub sw35_Hit:Controller.Switch(35) = 1:End Sub
Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub

Sub sw23_Hit:Controller.Switch(23) = 1:End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub

Sub sw38_Hit:Controller.Switch(38) = 1:End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

Sub sw42_Hit:Controller.Switch(42) = 1:End Sub
Sub sw42_UnHit:Controller.Switch(42) = 0:End Sub


'//////////////////////////////////////////////////////////////////////

'Targets
Sub sw17_Hit:STHit 17:End Sub
Sub sw18_Hit:STHit 18:End Sub
Sub sw19_Hit:STHit 19:End Sub
Sub sw22_Hit:STHit 22:End Sub
Sub sw28_Hit:STHit 28:End Sub
Sub sw29_Hit:STHit 29:End Sub
Sub sw40_Hit:STHit 40:End Sub
Sub sw44_Hit:STHit 44:End Sub
Sub sw45_Hit:STHit 45:End Sub
Sub sw46_Hit:STHit 46:End Sub

'' Scoop Hole with animation
'Dim aBall, aZpos
'Sub swLock_Hit:bsLock.AddBall Me:End Sub
'Sub swRocket_in_Hit:PlaySoundAt "Eject_Enter_2", swRocket_in:bsRocket.AddBall Me:End Sub


'******************************************************
'* BUMPER ANIMATIONS 
'******************************************************

Sub Bumper1_Animate
	Dim z, BP
	z = Bumper1.CurrentRingOffset
	For Each BP in BP_Bumper1 : BP.transz = z: Next
End Sub

Sub Bumper2_Animate
	Dim z, BP
	z = Bumper2.CurrentRingOffset
	For Each BP in BP_Bumper2 : BP.transz = z: Next
End Sub

Sub Bumper3_Animate
	Dim z, BP
	z = Bumper3.CurrentRingOffset
	For Each BP in BP_Bumper3 : BP.transz = z: Next
End Sub



'******************************************************
'	SWITCH ANIMATIONS
'******************************************************

Sub sw21_Animate
	Dim z : z = sw21.CurrentAnimOffset
	Dim BP : For Each BP in BP_sw21 : BP.transz = z: Next
End Sub

Sub sw25_Animate
	Dim z : z = sw25.CurrentAnimOffset
	Dim BP : For Each BP in BP_sw25 : BP.transz = z: Next
End Sub

Sub sw26_Animate
	Dim z : z = sw26.CurrentAnimOffset
	Dim BP : For Each BP in BP_sw26 : BP.transz = z: Next
End Sub

Sub sw27_Animate
	Dim z : z = sw27.CurrentAnimOffset
	Dim BP : For Each BP in BP_sw27 : BP.transz = z: Next
End Sub

Sub sw37_Animate
	Dim z : z = sw37.CurrentAnimOffset
	Dim BP : For Each BP in BP_sw37 : BP.transz = z: Next
End Sub

Sub sw42_Animate
	Dim z : z = sw42.CurrentAnimOffset
	Dim BP : For Each BP in BP_sw42 : BP.transz = z: Next
End Sub

Sub sw43_Animate
	Dim z : z = sw43.CurrentAnimOffset
	Dim BP : For Each BP in BP_sw43 : BP.transz = z: Next
End Sub

Sub sw57_Animate
	Dim z : z = sw57.CurrentAnimOffset
	Dim BP : For Each BP in BP_sw57 : BP.transz = z: Next
End Sub

Sub sw58_Animate
	Dim z : z = sw58.CurrentAnimOffset
	Dim BP : For Each BP in BP_sw58 : BP.transz = z: Next
End Sub

Sub sw60_Animate
	Dim z : z = sw60.CurrentAnimOffset
	Dim BP : For Each BP in BP_sw60 : BP.transz = z: Next
End Sub

Sub sw61_Animate
	Dim z : z = sw61.CurrentAnimOffset
	Dim BP : For Each BP in BP_sw61 : BP.transz = z: Next
End Sub


'******************************************************
'	GATE ANIMATIONS
'******************************************************

Sub Gate1_Animate
	Dim a : a = Gate1.CurrentAngle
	Dim BP : For Each BP in BP_Gate1 : BP.rotx = a: Next
End Sub

Sub Gate3_Animate
	Dim a : a = Gate3.CurrentAngle
	Dim BP : For Each BP in BP_Gate3 : BP.rotx = a: Next
End Sub

Sub Gate5_Animate
	Dim a : a = Gate5.CurrentAngle
	Dim BP : For Each BP in BP_Gate5 : BP.rotx = a: Next
End Sub

Sub Gate6_Animate
	Dim a : a = Gate6.CurrentAngle
	Dim BP : For Each BP in BP_Gate6 : BP.rotx = a: Next
End Sub

Sub Gate7_Animate
	Dim a : a = Gate7.CurrentAngle
	Dim BP : For Each BP in BP_Gate7 : BP.rotx = a: Next
End Sub

Sub Gate8_Animate
	Dim a : a = Gate8.CurrentAngle
	Dim BP : For Each BP in BP_Gate8 : BP.rotx = a: Next
End Sub

Sub Gate002_Animate
	Dim a : a = Gate002.CurrentAngle
	Dim BP : For Each BP in BP_Gate002 : BP.rotx = -a: Next
End Sub

Sub Gate003_Animate
	Dim a : a = Gate003.CurrentAngle
	Dim BP : For Each BP in BP_Gate003 : BP.rotx = a: Next
End Sub

Sub Gate004_Animate
	Dim a : a = Gate004.CurrentAngle
	Dim BP : For Each BP in BP_Gate004 : BP.rotx = a: Next
End Sub

Sub Gate005_Animate
	Dim a : a = Gate005.CurrentAngle
	Dim BP : For Each BP in BP_Gate005 : BP.rotx = a: Next
End Sub

'Ghost
Sub GhostTarget_Animate
	Dim a : a = GhostTarget.CurrentAngle
	Dim BP : For Each BP in BP_GhostTarget : BP.rotx = a: Next
End Sub

''******************************************************
''	DIVERTER ANIMATIONS
''******************************************************
'
'Sub Diverter_left_Animate
'	Dim a : a = Diverter_left.CurrentAngle
'	Dim BP : For Each BP in BP_Diverter_left : BP.rotx = a: Next
'End Sub
'
'Sub Diverter_right_Animate
'	Dim a : a = Diverter_right.CurrentAngle
'	Dim BP : For Each BP in BP_Diverter_right : BP.rotx = a: Next
'End Sub


'******************************************************
'	DUMMY ANIMATION
'******************************************************

Sub DummyFlipper_Animate
 	Dim BP
	For Each BP in BP_Dummy : BP.TransZ = DummyFlipper.CurrentAngle: Next
End Sub





'*********
'Solenoids
'*********
SolCallBack(1) = "SolRelease"
SolCallBack(2) = "Auto_Plunger"
SolCallback(3) = "BallLock"
SolCallback(4) = "SolDTSweeperUp"
SolCallback(5) = "SolDTSweeperDown"
SolCallback(6) = "SolDTDropDown"
SolCallback(7) = "Rocket"
SolCallback(8) = "KioskScoop"
SolCallback(12) = "SolDTDropUp"
SolCallback(19) = "SolGhost"
SolCallback(20) = "SolPost"
SolCallback(25) = "SolLock"
SolCallback(26) = "SolDiverterRight" 'Right Ramp Diverter
SolCallBack(28) = "DummyAnim"    'Dummy


Sub Auto_Plunger(Enabled)
    If Enabled Then
		PlaySoundAtVol SoundFX("z_RCT_Autolaunch", DOFContactors), Plunger, RecordedVolume
		PlungerIM.AutoFire
    End If
End Sub

Sub SolGhost(Enabled)
    If Enabled Then
		'debug.print "SolGhost true: GhostTarget.open  = 0" 
        GhostTarget.open  = 0
		GhostTarget_Animate
        Controller.Switch(36) = 1
    PlaySoundAtVol SoundFX("z_RCT_Ghost_Release", DOFContactors), GhostTarget, RecordedVolume
    End if
End Sub

Sub SolPost(Enabled)
	Dim BP
	If Enabled Then
    PlaySoundAtVol SoundFX("z_RCT_UpPost", DOFContactors), PostSound, RecordedVolume
		Post.IsDropped = False
		For Each BP in BP_Post : BP.TransZ = 0: Next
	Else
		Post.IsDropped = True
		For Each BP in BP_Post : BP.TransZ = -54: Next
	End If
End Sub

Sub SolLock(Enabled)
    PlaySoundAtVol SoundFX("z_RCT_Diverter_Left", DOFContactors), Diverter_left, RecordedVolume
    If Enabled Then
        Diverter_left.Open = 0
    Else
        Diverter_left.Open = 1
    End If
End Sub

Sub SolDiverterRight(Enabled)
    PlaySoundAtVol SoundFX("z_RCT_Diverter_Right", DOFContactors), Diverter_right, RecordedVolume
    If Enabled Then
        Diverter_right.Open = 0
    Else
        Diverter_right.Open = 1
    End If
End Sub

Sub sw36_Hit
	if GhostTarget.open = 0 and activeball.vely < 0 Then
		'debug.print "sw36_Hit: GhostTarget.open  = 1" 
		GhostTarget.open  = 1
		GhostTarget_Animate
		Controller.Switch(36) = 0
		PlaySoundAtVol SoundFX("z_RCT_Ghost_Hit", DOFContactors), sw36, RecordedVolume - 0.1
	end if
End Sub

Sub DummyAnim(Enabled)
    PlaySoundAtVol SoundFX("z_RCT_Dummy", DOFContactors), BM_Dummy, RecordedVolume
    DummyFlipper.RotateToEnd
    vpmtimer.addtimer 150, "DummyOff '"
End Sub

Sub DummyOff
    DummyFlipper.RotateToStart
End Sub


'Sub GhostTest(flag)
'	if flag then
'		GhostTarget.open  = 1
'		GhostTarget_Animate
'		PlaySoundAtBall SoundFX("Drop_Target_Down_3", DOFDropTargets)
'	else
'        GhostTarget.open  = 0
'		GhostTarget_Animate
'        PlaySound SoundFX("z_Ghost_Release", DOFContactors)
'	End if
'end sub


'*************************************************************
'PWM Flasher Stuff
'*************************************************************
Const DebugFlashers = False


'flashers
SolModCallback(21) = "FlashMod121" 'Lockup 121
SolModCallback(22) = "FlashMod122" 'Shooter 122
SolModCallback(23) = "FlashMod123" 'Kiosk 123
SolModCallback(27) = "FlashMod127" 'Bumpers 127
SolModCallback(29) = "FlashMod129" 'Sign Right 129
SolModCallback(30) = "FlashMod130" 'Sign Middle 130
SolModCallback(31) = "FlashMod131" 'Sign Left 131
SolModCallback(32) = "FlashMod132" 'Middle Left 132


' Modulated (pwm) flasher subs
Sub FlashMod121(pwm)
	If DebugFlashers Then Debug.print "FlashMod121 level=" & pwm
	L121.State = pwm
End Sub

Sub FlashMod122(pwm)
	If DebugFlashers Then Debug.print "FlashMod122 level=" & pwm
	L122.State = pwm
End Sub

Sub FlashMod123(pwm)
	If DebugFlashers Then Debug.print "FlashMod123 level=" & pwm
	L123.State = pwm
End Sub

Sub FlashMod127(pwm)
	If DebugFlashers Then Debug.print "FlashMod127 level=" & pwm
	L127.State = pwm
End Sub

Sub FlashMod129(pwm)
	If DebugFlashers Then Debug.print "FlashMod129 level=" & pwm
	L129.State = pwm
End Sub

Sub FlashMod130(pwm)
	If DebugFlashers Then Debug.print "FlashMod130 level=" & pwm
	L130.State = pwm
End Sub

Sub FlashMod131(pwm)
	If DebugFlashers Then Debug.print "FlashMod131 level=" & pwm
	L131.State = pwm
End Sub

Sub FlashMod132(pwm)
	If DebugFlashers Then Debug.print "FlashMod132 level=" & pwm
	L132.State = pwm
End Sub


'*************************
' GI - needs new vpinmame
'*************************

dim gilvl:gilvl = 0

Set GICallback2 = GetRef("GIUpdate2")


Sub GIUpdate2(no, level)
	'debug.print "GIUpdate2 no="&no&" level="&level
	Dim bulb
	For each bulb in aGiLights: bulb.State = level: Next
	If level >= 0.5 And gilvl < 0.5 Then
		Sound_GI_Relay 1, RelaySound
	ElseIf level <= 0.4 And gilvl > 0.4 Then
		Sound_GI_Relay 0, RelaySound
	End If
	gilvl = level
End Sub

    


'//////////////////////////////////////////////////////////////////////
'  Drop Target Controls
'//////////////////////////////////////////////////////////////////////

Sub SolDTDropUp(Enabled)
	If Enabled Then
		DTRaise 30
		DTRaise 31
		DTRaise 32
PlaySoundAtVol SoundFX("z_RCT_3Bank_Reset", DOFContactors), DropTargetSoundLeft, RecordedVolume
	End If
End Sub

Sub SolDTDropDown(Enabled)
	If Enabled Then
		DTDrop 31
PlaySoundAtVol SoundFX("z_RCT_1Bank_Drop", DOFContactors), DropTargetSoundLeft, RecordedVolume
	End If
End Sub

Sub SolDTSweeperUp(Enabled)
	If Enabled Then
		Dt39Kick2.RotateToEnd
		Dt39Kick2.Enabled = True
		DTRaise 39
PlaySoundAtVol SoundFX("z_RCT_1Bank_Reset", DOFContactors), sw38, RecordedVolume
		Else
		Dt39Kick2.RotateToStart
		Dt39Kick2.Enabled = False
	End If
End Sub

Sub SolDTSweeperDown(Enabled)
	If Enabled Then
		Dt39Kick.RotateToEnd
		Dt39Kick.Enabled = True
		DTDrop 39
PlaySoundAtVol SoundFX("z_RCT_1Bank_Drop", DOFContactors), sw38, RecordedVolume
	Else
		Dt39Kick.RotateToStart
		Dt39Kick.Enabled = False
	End If
End Sub


Sub sw30_hit
	DTHit 30
PlaySoundAtVol SoundFX("z_RCT_1Bank_Drop", DOFContactors), DropTargetSoundLeft, RecordedVolume
End Sub

Sub sw31_hit
	DTHit 31
PlaySoundAtVol SoundFX("z_RCT_1Bank_Drop", DOFContactors), DropTargetSoundLeft, RecordedVolume

End Sub

Sub sw32_hit
	DTHit 32
PlaySoundAtVol SoundFX("z_RCT_1Bank_Drop", DOFContactors), DropTargetSoundLeft, RecordedVolume
End Sub

Sub sw39_hit
	DTHit 39
PlaySoundAtVol SoundFX("z_RCT_1Bank_Drop", DOFContactors), sw38, RecordedVolume
End Sub

'//////////////////////////////////////////////////////////////////////
'// FLIPPERS 
'//////////////////////////////////////////////////////////////////////

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sURFlipper) = "SolURFlipper"
SolCallback(sULFlipper) = "SolULFlipper"

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled)
	If Enabled Then
		LF.Fire 
		PlaySoundAtVol SoundFX("z_RCT_LeftFlipper_Up", DOFContactors), LeftFlipper, RecordedVolume	
		PlaySoundAtVolLoops ("z_RCT_FlipperHold_Left"), LeftFlipper , RecordedVolume, -1
	Else
        LeftFlipper.RotateToStart
		PlaySoundAtVol SoundFX("z_RCT_LeftFlipper_Down", DOFContactors), LeftFlipper, RecordedVolume
		PlaySoundAtVol SoundFX("z_RCT_mini_LeftFlipper_Down", DOFContactors), LeftFlipper1, RecordedVolume
        StopSound "z_RCT_FlipperHold_Left"
		FlipperLeftHitParm = FlipperUpSoundLevel
	End If
End Sub

Sub SolRFlipper(Enabled)
	If Enabled Then
		RF.Fire  'rightflipper.rotatetoend
        PlaySoundAtVol SoundFX("z_RCT_RightFlipper_Up", DOFContactors), RightFlipper, RecordedVolume
        PlaySoundAtVolLoops ("z_RCT_FlipperHold_Right"), RightFlipper , RecordedVolume, -1
	Else
		RightFlipper.RotateToStart
        PlaySoundAtVol SoundFX("z_RCT_RightFlipper_Down", DOFContactors), RightFlipper, RecordedVolume
        StopSound "z_RCT_FlipperHold_Right"
		FlipperRightHitParm = FlipperUpSoundLevel
	End If
End Sub

Sub SolULFlipper(Enabled)
	If Enabled Then
		LeftFlipper1.RotateToEnd
		PlaySoundAtVol SoundFX("z_RCT_mini_LeftFlipper_Up", DOFContactors), LeftFlipper1, RecordedVolume
	Else
        LeftFlipper1.RotateToStart
		PlaySoundAtVol SoundFX("z_RCT_mini_LeftFlipper_Down", DOFContactors), LeftFlipper1, RecordedVolume
		FlipperLeftHitParm = FlipperUpSoundLevel
	End If
End Sub


Sub SolURFlipper(Enabled)
	If Enabled Then
		RightFlipper1.RotateToEnd 'rightflipper.rotatetoend
        PlaySoundAtVol SoundFX("z_RCT_Top_Flipper_Up", DOFContactors), RightFlipper1, RecordedVolume
        PlaySoundAtVolLoops ("z_RCT_FlipperHold_Right"), RightFlipper , RecordedVolume, -1
	Else
        RightFlipper1.RotateToStart
        PlaySoundAtVol SoundFX("z_RCT_Top_Flipper_Down", DOFContactors), RightFlipper1, RecordedVolume
		FlipperRightHitParm = FlipperUpSoundLevel
	End If
End Sub


'Flipper collide subs
Sub LeftFlipper_Collide(parm)
	CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
	LF.ReProcessBalls ActiveBall
	LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
	CheckLiveCatch Activeball, RightFlipper, RFCount, parm
	RF.ReProcessBalls ActiveBall
	RightFlipperCollide parm
End Sub


Sub LeftFlipper_Animate
	Dim a : a = LeftFlipper.CurrentAngle
    FlipperLSh.RotZ = a

	' Darken light from lane bulbs when bats are up
	Dim v, BP
	v = 255.0 * (120.0 -  LeftFlipper.CurrentAngle) / (120.0 -  68.0)

	For each BP in BP_LeftFlipper
		BP.Rotz = a
		BP.visible = v < 128.0
	Next
	For each BP in BP_LeftFlipperU
		BP.Rotz = a
		BP.visible = v >= 128.0
	Next
End Sub

Sub RightFlipper_Animate
	Dim a : a = RightFlipper.CurrentAngle
    FlipperRSh.RotZ = a

	' Darken light from lane bulbs when bats are up
	Dim v, BP
	v = 255.0 * (-120.0 - RightFlipper.CurrentAngle) / (-120.0 + 68.0)

	For each BP in BP_RightFlipper
		BP.Rotz = a
		BP.visible = v < 128.0
	Next
	For each BP in BP_RightFlipperU
		BP.Rotz = a
		BP.visible = v >= 128.0
	Next
End Sub


Sub LeftFlipper1_Animate
	Dim a : a = LeftFlipper1.CurrentAngle
    FlipperLSh1.RotZ = a

	' Darken light from lane bulbs when bats are up
	Dim v, BP
	v = 255.0 * (170.0 -  LeftFlipper1.CurrentAngle) / (170.0 -  117.0)

	For each BP in BP_LeftFlipper1
		BP.Rotz = a
		BP.visible = v < 128.0
	Next
	For each BP in BP_LeftFlipper1U
		BP.Rotz = a
		BP.visible = v >= 128.0
	Next
End Sub


Sub RightFlipper1_Animate
	Dim a : a = RightFlipper1.CurrentAngle
    FlipperRSh1.RotZ = a

	' Darken light from lane bulbs when bats are up
	Dim v, BP
	v = 255.0 * (-157.0 - RightFlipper1.CurrentAngle) / (-157.0 + 105.0)

	For each BP in BP_RightFlipper1
		BP.Rotz = a
		BP.visible = v < 128.0
	Next
	For each BP in BP_RightFlipper1U
		BP.Rotz = a
		BP.visible = v >= 128.0
	Next
End Sub



dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

'*******************************************
' Early 90's and after

Sub InitPolarity()
	Dim x, a
	a = Array(LF, RF)
	For Each x In a
		x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
		x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
		x.enabled = True
		x.TimeDelay = 60
		x.DebugOn=False ' prints some info in debugger

		x.AddPt "Polarity", 0, 0, 0
		x.AddPt "Polarity", 1, 0.05, - 5.5
		x.AddPt "Polarity", 2, 0.16, - 5.5
		x.AddPt "Polarity", 3, 0.20, - 0.75
		x.AddPt "Polarity", 4, 0.25, - 1.25
		x.AddPt "Polarity", 5, 0.3, - 1.75
		x.AddPt "Polarity", 6, 0.4, - 3.5
		x.AddPt "Polarity", 7, 0.5, - 5.25
		x.AddPt "Polarity", 8, 0.7, - 4.0
		x.AddPt "Polarity", 9, 0.75, - 3.5
		x.AddPt "Polarity", 10, 0.8, - 3.0
		x.AddPt "Polarity", 11, 0.85, - 2.5
		x.AddPt "Polarity", 12, 0.9, - 2.0
		x.AddPt "Polarity", 13, 0.95, - 1.5
		x.AddPt "Polarity", 14, 1, - 1.0
		x.AddPt "Polarity", 15, 1.05, -0.5
		x.AddPt "Polarity", 16, 1.1, 0
		x.AddPt "Polarity", 17, 1.3, 0

		x.AddPt "Velocity", 0, 0, 0.85
		x.AddPt "Velocity", 1, 0.23, 0.85
		x.AddPt "Velocity", 2, 0.27, 1
		x.AddPt "Velocity", 3, 0.3, 1
		x.AddPt "Velocity", 4, 0.35, 1
		x.AddPt "Velocity", 5, 0.6, 1 '0.982
		x.AddPt "Velocity", 6, 0.62, 1.0
		x.AddPt "Velocity", 7, 0.702, 0.968
		x.AddPt "Velocity", 8, 0.95,  0.968
		x.AddPt "Velocity", 9, 1.03,  0.945
		x.AddPt "Velocity", 10, 1.5,  0.945

	Next
	
	' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
	LF.SetObjects "LF", LeftFlipper, TriggerLF
	RF.SetObjects "RF", RightFlipper, TriggerRF
End Sub


'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

' modified 2023 by nFozzy
' Removed need for 'endpoint' objects
' Added 'createvents' type thing for TriggerLF / TriggerRF triggers.
' Removed AddPt function which complicated setup imo
' made DebugOn do something (prints some stuff in debugger)
'   Otherwise it should function exactly the same as before\
' modified 2024 by rothbauerw
' Added Reprocessballs for flipper collisions (LF.Reprocessballs Activeball and RF.Reprocessballs Activeball must be added to the flipper collide subs
' Improved handling to remove correction for backhand shots when the flipper is raised

Class FlipperPolarity
	Public DebugOn, Enabled
	Private FlipAt		'Timer variable (IE 'flip at 723,530ms...)
	Public TimeDelay		'delay before trigger turns off and polarity is disabled
	Private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef, FlipStartAngle
	Private Balls(20), balldata(20)
	Private Name
	
	Dim PolarityIn, PolarityOut
	Dim VelocityIn, VelocityOut
	Dim YcoefIn, YcoefOut
	Public Sub Class_Initialize
		ReDim PolarityIn(0)
		ReDim PolarityOut(0)
		ReDim VelocityIn(0)
		ReDim VelocityOut(0)
		ReDim YcoefIn(0)
		ReDim YcoefOut(0)
		Enabled = True
		TimeDelay = 50
		LR = 1
		Dim x
		For x = 0 To UBound(balls)
			balls(x) = Empty
			Set Balldata(x) = new SpoofBall
		Next
	End Sub
	
	Public Sub SetObjects(aName, aFlipper, aTrigger)
		
		If TypeName(aName) <> "String" Then MsgBox "FlipperPolarity: .SetObjects error: first argument must be a String (And name of Object). Found:" & TypeName(aName) End If
		If TypeName(aFlipper) <> "Flipper" Then MsgBox "FlipperPolarity: .SetObjects error: Second argument must be a flipper. Found:" & TypeName(aFlipper) End If
		If TypeName(aTrigger) <> "Trigger" Then MsgBox "FlipperPolarity: .SetObjects error: third argument must be a trigger. Found:" & TypeName(aTrigger) End If
		If aFlipper.EndAngle > aFlipper.StartAngle Then LR = -1 Else LR = 1 End If
		Name = aName
		Set Flipper = aFlipper
		FlipperStart = aFlipper.x
		FlipperEnd = Flipper.Length * Sin((Flipper.StartAngle / 57.295779513082320876798154814105)) + Flipper.X ' big floats for degree to rad conversion
		FlipperEndY = Flipper.Length * Cos(Flipper.StartAngle / 57.295779513082320876798154814105)*-1 + Flipper.Y
		
		Dim str
		str = "Sub " & aTrigger.name & "_Hit() : " & aName & ".AddBall ActiveBall : End Sub'"
		ExecuteGlobal(str)
		str = "Sub " & aTrigger.name & "_UnHit() : " & aName & ".PolarityCorrect ActiveBall : End Sub'"
		ExecuteGlobal(str)
		
	End Sub
	
	' Legacy: just no op
	Public Property Let EndPoint(aInput)
		
	End Property
	
	Public Sub AddPt(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
		Select Case aChooseArray
			Case "Polarity"
				ShuffleArrays PolarityIn, PolarityOut, 1
				PolarityIn(aIDX) = aX
				PolarityOut(aIDX) = aY
				ShuffleArrays PolarityIn, PolarityOut, 0
			Case "Velocity"
				ShuffleArrays VelocityIn, VelocityOut, 1
				VelocityIn(aIDX) = aX
				VelocityOut(aIDX) = aY
				ShuffleArrays VelocityIn, VelocityOut, 0
			Case "Ycoef"
				ShuffleArrays YcoefIn, YcoefOut, 1
				YcoefIn(aIDX) = aX
				YcoefOut(aIDX) = aY
				ShuffleArrays YcoefIn, YcoefOut, 0
		End Select
	End Sub
	
	Public Sub AddBall(aBall)
		Dim x
		For x = 0 To UBound(balls)
			If IsEmpty(balls(x)) Then
				Set balls(x) = aBall
				Exit Sub
			End If
		Next
	End Sub
	
	Private Sub RemoveBall(aBall)
		Dim x
		For x = 0 To UBound(balls)
			If TypeName(balls(x) ) = "IBall" Then
				If aBall.ID = Balls(x).ID Then
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
		Dim x
		For x = 0 To UBound(balls)
			If Not IsEmpty(balls(x)) Then
				pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
			End If
		Next
	End Property
	
	Public Sub ProcessBalls() 'save data of balls in flipper range
		FlipAt = GameTime
		Dim x
		For x = 0 To UBound(balls)
			If Not IsEmpty(balls(x)) Then
				balldata(x).Data = balls(x)
			End If
		Next
		FlipStartAngle = Flipper.currentangle
		PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
		PartialFlipCoef = abs(PartialFlipCoef-1)
	End Sub

	Public Sub ReProcessBalls(aBall) 'save data of balls in flipper range
		If FlipperOn() Then
			Dim x
			For x = 0 To UBound(balls)
				If Not IsEmpty(balls(x)) Then
					if balls(x).ID = aBall.ID Then
						If isempty(balldata(x).ID) Then
							balldata(x).Data = balls(x)
						End If
					End If
				End If
			Next
		End If
	End Sub

	'Timer shutoff for polaritycorrect
	Private Function FlipperOn()
		If GameTime < FlipAt+TimeDelay Then
			FlipperOn = True
		End If
	End Function
	
	Public Sub PolarityCorrect(aBall)
		If FlipperOn() Then
			Dim tmp, BallPos, x, IDX, Ycoef, BalltoFlip, BalltoBase, NoCorrection, checkHit
			Ycoef = 1
			
			'y safety Exit
			If aBall.VelY > -8 Then 'ball going down
				RemoveBall aBall
				Exit Sub
			End If
			
			'Find balldata. BallPos = % on Flipper
			For x = 0 To UBound(Balls)
				If aBall.id = BallData(x).id And Not IsEmpty(BallData(x).id) Then
					idx = x
					BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
					BalltoFlip = DistanceFromFlipperAngle(BallData(x).x, BallData(x).y, Flipper, FlipStartAngle)
					If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)								'find safety coefficient 'ycoef' data
				End If
			Next
			
			If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
				BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
				If ballpos > 0.65 Then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)												'find safety coefficient 'ycoef' data
				NoCorrection = 1
			Else
				checkHit = 50 + (20 * BallPos) 

				If BalltoFlip > checkHit or (PartialFlipCoef < 0.5 and BallPos > 0.22) Then
					NoCorrection = 1
				Else
					NoCorrection = 0
				End If
			End If
			
			'Velocity correction
			If Not IsEmpty(VelocityIn(0) ) Then
				Dim VelCoef
				VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)
				
				'If partialflipcoef < 1 Then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)
				
				If Enabled Then aBall.Velx = aBall.Velx*VelCoef
				If Enabled Then aBall.Vely = aBall.Vely*VelCoef
			End If
			
			'Polarity Correction (optional now)
			If Not IsEmpty(PolarityIn(0) ) Then
				Dim AddX
				AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR
				
				If Enabled and NoCorrection = 0 Then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef*VelCoef)
			End If
			If DebugOn Then debug.print "PolarityCorrect" & " " & Name & " @ " & GameTime & " " & Round(BallPos*100) & "%" & " AddX:" & Round(AddX,2) & " Vel%:" & Round(VelCoef*100)
		End If
		RemoveBall aBall
	End Sub
End Class

'******************************************************
'  FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
	Dim x, aCount
	aCount = 0
	ReDim a(UBound(aArray) )
	For x = 0 To UBound(aArray)		'Shuffle objects in a temp array
		If Not IsEmpty(aArray(x) ) Then
			If IsObject(aArray(x)) Then
				Set a(aCount) = aArray(x)
			Else
				a(aCount) = aArray(x)
			End If
			aCount = aCount + 1
		End If
	Next
	If offset < 0 Then offset = 0
	ReDim aArray(aCount-1+offset)		'Resize original array
	For x = 0 To aCount-1				'set objects back into original array
		If IsObject(a(x)) Then
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
	BallSpeed = Sqr(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)		'Set up line via two points, no clamping. Input X, output Y
	Dim x, y, b, m
	x = input
	m = (Y2 - Y1) / (X2 - X1)
	b = Y2 - m*X2
	Y = M*x+b
	PSlope = Y
End Function

' Used for flipper correction
Class spoofball
	Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
	Public Property Let Data(aBall)
		With aBall
			x = .x
			y = .y
			z = .z
			velx = .velx
			vely = .vely
			velz = .velz
			id = .ID
			mass = .mass
			radius = .radius
		End With
	End Property
	Public Sub Reset()
		x = Empty
		y = Empty
		z = Empty
		velx = Empty
		vely = Empty
		velz = Empty
		id = Empty
		mass = Empty
		radius = Empty
	End Sub
End Class

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
	Dim y 'Y output
	Dim L 'Line
	'find active line
	Dim ii
	For ii = 1 To UBound(xKeyFrame)
		If xInput <= xKeyFrame(ii) Then
			L = ii
			Exit For
		End If
	Next
	If xInput > xKeyFrame(UBound(xKeyFrame) ) Then L = UBound(xKeyFrame)		'catch line overrun
	Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )
	
	If xInput <= xKeyFrame(LBound(xKeyFrame) ) Then Y = yLvl(LBound(xKeyFrame) )		 'Clamp lower
	If xInput >= xKeyFrame(UBound(xKeyFrame) ) Then Y = yLvl(UBound(xKeyFrame) )		'Clamp upper
	
	LinearEnvelope = Y
End Function

'******************************************************
'  FLIPPER TRICKS
'******************************************************
' To add the flipper tricks you must
'	 - Include a call to FlipperCradleCollision from within OnBallBallCollision subroutine
'	 - Include a call the CheckLiveCatch from the LeftFlipper_Collide and RightFlipper_Collide subroutines
'	 - Include FlipperActivate and FlipperDeactivate in the Flipper solenoid subs

RightFlipper.timerinterval = 1
Rightflipper.timerenabled = True

Sub RightFlipper_timer()
	FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
	FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
	FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
	FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
End Sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
	Dim b
	'   Dim BOT
	'   BOT = GetBalls
	
	If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
		EOSNudge1 = 1
		'   debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
		If Flipper2.currentangle = EndAngle2 Then
			For b = 0 To UBound(gBOT)
				If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper1) Then
					'Debug.Print "ball in flip1. exit"
					Exit Sub
				End If
			Next
			For b = 0 To UBound(gBOT)
				If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper2) Then
					gBOT(b).velx = gBOT(b).velx / 1.3
					gBOT(b).vely = gBOT(b).vely - 0.5
				End If
			Next
		End If
	Else
		If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 Then EOSNudge1 = 0
	End If
End Sub


Dim FCCDamping: FCCDamping = 0.4

Sub FlipperCradleCollision(ball1, ball2, velocity)
	if velocity < 0.7 then exit sub		'filter out gentle collisions
    Dim DoDamping, coef
    DoDamping = false
    'Check left flipper
    If LeftFlipper.currentangle = LFEndAngle Then
		If FlipperTrigger(ball1.x, ball1.y, LeftFlipper) OR FlipperTrigger(ball2.x, ball2.y, LeftFlipper) Then DoDamping = true
    End If
    'Check right flipper
    If RightFlipper.currentangle = RFEndAngle Then
		If FlipperTrigger(ball1.x, ball1.y, RightFlipper) OR FlipperTrigger(ball2.x, ball2.y, RightFlipper) Then DoDamping = true
    End If
    If DoDamping Then
		coef = FCCDamping
        ball1.velx = ball1.velx * coef: ball1.vely = ball1.vely * coef: ball1.velz = ball1.velz * coef
        ball2.velx = ball2.velx * coef: ball2.vely = ball2.vely * coef: ball2.velz = ball2.velz * coef
    End If
End Sub
	


'*****************
' Maths
'*****************

Dim PI
PI = 4 * Atn(1)

Function dSin(degrees)
	dsin = Sin(degrees * Pi / 180)
End Function

Function dCos(degrees)
	dcos = Cos(degrees * Pi / 180)
End Function

Function Atn2(dy, dx)
	If dx > 0 Then
		Atn2 = Atn(dy / dx)
	ElseIf dx < 0 Then
		If dy = 0 Then
			Atn2 = pi
		Else
			Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
		End If
	ElseIf dx = 0 Then
		If dy = 0 Then
			Atn2 = 0
		Else
			Atn2 = Sgn(dy) * pi / 2
		End If
	End If
End Function


'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
	Distance = Sqr((ax - bx) ^ 2 + (ay - by) ^ 2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) 'Distance between a point and a line where point Is px,py
	DistancePL = Abs((by - ay) * px - (bx - ax) * py + bx * ay - by * ax) / Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
	Radians = Degrees * PI / 180
End Function

Function AnglePP(ax,ay,bx,by)
	AnglePP = Atn2((by - ay),(bx - ax)) * 180 / PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
	DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle + 90)) + Flipper.x, Sin(Radians(Flipper.currentangle + 90)) + Flipper.y)
End Function

Function DistanceFromFlipperAngle(ballx, bally, Flipper, Angle)
	DistanceFromFlipperAngle = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Angle + 90)) + Flipper.x, Sin(Radians(angle + 90)) + Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
	Dim DiffAngle
	DiffAngle = Abs(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
	If DiffAngle > 180 Then DiffAngle = DiffAngle - 360
	
	If DistanceFromFlipper(ballx,bally,Flipper) < 48 And DiffAngle <= 90 And Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
		FlipperTrigger = True
	Else
		FlipperTrigger = False
	End If
End Function

'*************************************************
'  End - Check ball distance from Flipper for Rem
'*************************************************

Dim LFPress, RFPress, LFCount, RFCount
Dim LFState, RFState
Dim EOST, EOSA,Frampup, FElasticity,FReturn
Dim RFEndAngle, LFEndAngle

Const FlipperCoilRampupMode = 0 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
'Const EOSTnew = 1.5 'EM's to late 80's - new recommendation by rothbauerw (previously 1)
Const EOSTnew = 1.2 '90's and later - new recommendation by rothbauerw (previously 0.8)
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode
	Case 0
		SOSRampup = 2.5
	Case 1
		SOSRampup = 6
	Case 2
		SOSRampup = 8.5
End Select

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
'   Const EOSReturn = 0.055  'EM's
'   Const EOSReturn = 0.045  'late 70's to mid 80's
Const EOSReturn = 0.035  'mid 80's to early 90's
'   Const EOSReturn = 0.025  'mid 90's and later

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
	Flipper.eostorque = EOST * EOSReturn / FReturn
	
	If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
		Dim b', BOT
		'		BOT = GetBalls
		
		For b = 0 To UBound(gBOT)
			If Distance(gBOT(b).x, gBOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
				If gBOT(b).vely >= - 0.4 Then gBOT(b).vely =  - 0.4
			End If
		Next
	End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
	Dim Dir
	Dir = Flipper.startangle / Abs(Flipper.startangle) '-1 for Right Flipper
	
	If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
		If FState <> 1 Then
			Flipper.rampup = SOSRampup
			Flipper.endangle = FEndAngle - 3 * Dir
			Flipper.Elasticity = FElasticity * SOSEM
			FCount = 0
			FState = 1
		End If
	ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) And FlipperPress = 1 Then
		If FCount = 0 Then FCount = GameTime
		
		If FState <> 2 Then
			Flipper.eostorqueangle = EOSAnew
			Flipper.eostorque = EOSTnew
			Flipper.rampup = EOSRampup
			Flipper.endangle = FEndAngle
			FState = 2
		End If
	ElseIf Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 And FlipperPress = 1 Then
		If FState <> 3 Then
			Flipper.eostorque = EOST
			Flipper.eostorqueangle = EOSA
			Flipper.rampup = Frampup
			Flipper.Elasticity = FElasticity
			FState = 3
		End If
	End If
End Sub

Const LiveDistanceMin = 5  'minimum distance In vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114 'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)
Const BaseDampen = 0.55

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
    Dim Dir, LiveDist
    Dir = Flipper.startangle / Abs(Flipper.startangle)    '-1 for Right Flipper
    Dim LiveCatchBounce   'If live catch is not perfect, it won't freeze ball totally
    Dim CatchTime
    CatchTime = GameTime - FCount
    LiveDist = Abs(Flipper.x - ball.x)

    If CatchTime <= LiveCatch And parm > 3 And LiveDist > LiveDistanceMin And LiveDist < LiveDistanceMax Then
        If CatchTime <= LiveCatch * 0.5 Then   'Perfect catch only when catch time happens in the beginning of the window
            LiveCatchBounce = 0
        Else
            LiveCatchBounce = Abs((LiveCatch / 2) - CatchTime)  'Partial catch when catch happens a bit late
        End If
        
        If LiveCatchBounce = 0 And ball.velx * Dir > 0 And LiveDist > 30 Then ball.velx = 0

        If ball.velx * Dir > 0 And LiveDist < 30 Then
            ball.velx = BaseDampen * ball.velx
            ball.vely = BaseDampen * ball.vely
            ball.angmomx = BaseDampen * ball.angmomx
            ball.angmomy = BaseDampen * ball.angmomy
            ball.angmomz = BaseDampen * ball.angmomz
        Elseif LiveDist > 30 Then
            ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
            ball.angmomx = 0
            ball.angmomy = 0
            ball.angmomz = 0
        End If
    Else
        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf ActiveBall, parm
    End If
End Sub

'******************************************************
'****  END FLIPPER CORRECTIONS
'******************************************************



'******************************************************
' 	ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1	  '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.9	 'Level of bounces. Recommmended value of 0.7

Sub TargetBouncer(aBall,defvalue)
	Dim zMultiplier, vel, vratio
	If TargetBouncerEnabled = 1 And aball.z < 30 Then
		'   debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
		vel = BallSpeed(aBall)
		If aBall.velx = 0 Then vratio = 1 Else vratio = aBall.vely / aBall.velx
		Select Case Int(Rnd * 6) + 1
			Case 1
				zMultiplier = 0.2 * defvalue
			Case 2
				zMultiplier = 0.25 * defvalue
			Case 3
				zMultiplier = 0.3 * defvalue
			Case 4
				zMultiplier = 0.4 * defvalue
			Case 5
				zMultiplier = 0.45 * defvalue
			Case 6
				zMultiplier = 0.5 * defvalue
		End Select
		aBall.velz = Abs(vel * zMultiplier * TargetBouncerFactor)
		aBall.velx = Sgn(aBall.velx) * Sqr(Abs((vel ^ 2 - aBall.velz ^ 2) / (1 + vratio ^ 2)))
		aBall.vely = aBall.velx * vratio
		'   debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
		'   debug.print "conservation check: " & BallSpeed(aBall)/vel
	End If
End Sub



'******************************************************
' 	ZDMP:  RUBBER  DAMPENERS
'******************************************************
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

Sub dPosts_Hit(idx)
	RubbersD.dampen ActiveBall
	TargetBouncer ActiveBall, 1
End Sub

Sub dSleeves_Hit(idx)
	SleevesD.Dampen ActiveBall
	TargetBouncer ActiveBall, 0.7
End Sub

Sub zCol_Rubber_RightSling_Hit
	RubbersD.dampen ActiveBall
End Sub

Sub zCol_Rubber_LeftSling_Hit
	RubbersD.dampen ActiveBall
End Sub

Dim RubbersD				'frubber
Set RubbersD = New Dampener
RubbersD.name = "Rubbers"
RubbersD.debugOn = False	'shows info in textbox "TBPout"
RubbersD.Print = False	  'debug, reports In debugger (In vel, out cor); cor bounce curve (linear)

'for best results, try to match in-game velocity as closely as possible to the desired curve
'   RubbersD.addpoint 0, 0, 0.935   'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.1		 'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967	'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64	   'there's clamping so interpolate up to 56 at least

Dim SleevesD	'this is just rubber but cut down to 85%...
Set SleevesD = New Dampener
SleevesD.name = "Sleeves"
SleevesD.debugOn = False	'shows info in textbox "TBPout"
SleevesD.Print = False	  'debug, reports In debugger (In vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'######################### Adjust these values to increase or lessen the elasticity

Dim FlippersD
Set FlippersD = New Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
	Public Print, debugOn   'tbpOut.text
	Public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
	Public ModIn, ModOut
	Private Sub Class_Initialize
		ReDim ModIn(0)
		ReDim Modout(0)
	End Sub
	
	Public Sub AddPoint(aIdx, aX, aY)
		ShuffleArrays ModIn, ModOut, 1
		ModIn(aIDX) = aX
		ModOut(aIDX) = aY
		ShuffleArrays ModIn, ModOut, 0
		If GameTime > 100 Then Report
	End Sub
	
	Public Sub Dampen(aBall)
		If threshold Then
			If BallSpeed(aBall) < threshold Then Exit Sub
		End If
		Dim RealCOR, DesiredCOR, str, coef
		DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
		RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
		coef = desiredcor / realcor
		If debugOn Then str = name & " In vel:" & Round(cor.ballvel(aBall.id),2 ) & vbNewLine & "desired cor: " & Round(desiredcor,4) & vbNewLine & _
		"actual cor: " & Round(realCOR,4) & vbNewLine & "ballspeed coef: " & Round(coef, 3) & vbNewLine
		If Print Then Debug.print Round(cor.ballvel(aBall.id),2) & ", " & Round(desiredcor,3)
		
		aBall.velx = aBall.velx * coef
		aBall.vely = aBall.vely * coef
		aBall.velz = aBall.velz * coef
		If debugOn Then TBPout.text = str
	End Sub
	
	Public Sub Dampenf(aBall, parm) 'Rubberizer is handle here
		Dim RealCOR, DesiredCOR, str, coef
		DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
		RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
		coef = desiredcor / realcor
		If Abs(aball.velx) < 2 And aball.vely < 0 And aball.vely >  - 3.75 Then
			aBall.velx = aBall.velx * coef
			aBall.vely = aBall.vely * coef
			aBall.velz = aBall.velz * coef
		End If
	End Sub
	
	Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
		Dim x
		For x = 0 To UBound(aObj.ModIn)
			addpoint x, aObj.ModIn(x), aObj.ModOut(x) * aCoef
		Next
	End Sub
	
	Public Sub Report() 'debug, reports all coords in tbPL.text
		If Not debugOn Then Exit Sub
		Dim a1, a2
		a1 = ModIn
		a2 = ModOut
		Dim str, x
		For x = 0 To UBound(a1)
			str = str & x & ": " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
		Next
		TBPout.text = str
	End Sub
End Class

'******************************************************
'  TRACK ALL BALL VELOCITIES
'  FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

Dim cor
Set cor = New CoRTracker

Class CoRTracker
	Public ballvel, ballvelx, ballvely
	
	Private Sub Class_Initialize
		ReDim ballvel(0)
		ReDim ballvelx(0)
		ReDim ballvely(0)
	End Sub
	
	Public Sub Update()	'tracks in-ball-velocity
		Dim str, b, AllBalls, highestID
		allBalls = GetBalls
		
		For Each b In allballs
			If b.id >= HighestID Then highestID = b.id
		Next
		
		If UBound(ballvel) < highestID Then ReDim ballvel(highestID)	'set bounds
		If UBound(ballvelx) < highestID Then ReDim ballvelx(highestID)	'set bounds
		If UBound(ballvely) < highestID Then ReDim ballvely(highestID)	'set bounds
		
		For Each b In allballs
			ballvel(b.id) = BallSpeed(b)
			ballvelx(b.id) = b.velx
			ballvely(b.id) = b.vely
		Next
	End Sub
End Class

'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************



'//////////////////////////////////////////////////////////////////////
'// RAMP ROLLING SFX
'//////////////////////////////////////////////////////////////////////

'Ball tracking ramp SFX 1.0
'   Reqirements:
'          * Import A Sound File for each ball on the table for plastic ramps.  Call It RampLoop<Ball_Number> ex: RampLoop1, RampLoop2, ...
'          * Import a Sound File for each ball on the table for wire ramps. Call it WireLoop<Ball_Number> ex: WireLoop1, WireLoop2, ...
'          * Create a Timer called RampRoll, that is enabled, with a interval of 100
'          * Set RampBAlls and RampType variable to Total Number of Balls
'	Usage:
'          * Setup hit events and call WireRampOn True or WireRampOn False (True = Plastic ramp, False = Wire Ramp)
'          * To stop tracking ball
'                 * call WireRampOff
'                 * Otherwise, the ball will auto remove if it's below 30 vp units
'

dim RampMinLoops : RampMinLoops = 4

' RampBalls
'      Setup:        Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  if tnob = 5, then RammBalls(6,2)
'      Description:  
dim RampBalls(6,2)
'x,0 = ball x,1 = ID,	2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(6)	

Sub WireRampOn(input)  : Waddball ActiveBall, input : RampRollUpdate: End Sub
Sub WireRampOff() : WRemoveBall ActiveBall.ID	: End Sub


' WaddBall (Active Ball, Boolean)
'     Description: This subroutine is called from WireRampOn to Add Balls to the RampBalls Array
Sub Waddball(input, RampInput)	'Add ball
	' This will loop through the RampBalls array checking each element of the array x, position 1
	' To see if the the ball was already added to the array.
	' If the ball is found then exit the subroutine
	dim x : for x = 1 to uBound(RampBalls)	'Check, don't add balls twice
		if RampBalls(x, 1) = input.id then 
			if Not IsEmpty(RampBalls(x,1) ) then Exit Sub	'Frustating issue with BallId 0. Empty variable = 0
		End If
	Next

	' This will itterate through the RampBalls Array.
	' The first time it comes to a element in the array where the Ball Id (Slot 1) is empty.  It will add the current ball to the array
	' The RampBalls assigns the ActiveBall to element x,0 and ball id of ActiveBall to 0,1
	' The RampType(BallId) is set to RampInput
	' RampBalls in 0,0 is set to True, this will enable the timer and the timer is also turned on
	For x = 1 to uBound(RampBalls)
		if IsEmpty(RampBalls(x, 1)) then 
			Set RampBalls(x, 0) = input
			RampBalls(x, 1)	= input.ID
			RampType(x) = RampInput
			RampBalls(x, 2)	= 0
			'exit For
			RampBalls(0,0) = True
			RampRoll.Enabled = 1	 'Turn on timer
			'RampRoll.Interval = RampRoll.Interval 'reset timer
			exit Sub
		End If
		if x = uBound(RampBalls) then 	'debug
			Debug.print "WireRampOn error, ball queue is full: " & vbnewline & _
			RampBalls(0, 0) & vbnewline & _
			Typename(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & "type:" & RampType(1) & vbnewline & _
			Typename(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & "type:" & RampType(2) & vbnewline & _
			Typename(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & "type:" & RampType(3) & vbnewline & _
			Typename(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & "type:" & RampType(4) & vbnewline & _
			Typename(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & "type:" & RampType(5) & vbnewline & _
			" "
		End If
	next
End Sub

' WRemoveBall (BallId)
'    Description: This subroutine is called from the RampRollUpdate subroutine 
'                 and is used to remove and stop the ball rolling sounds
Sub WRemoveBall(ID)		'Remove ball
	'Debug.Print "In WRemoveBall() + Remove ball from loop array"
	dim ballcount : ballcount = 0
	dim x : for x = 1 to Ubound(RampBalls)
		if ID = RampBalls(x, 1) then 'remove ball
			Set RampBalls(x, 0) = Nothing
			RampBalls(x, 1) = Empty
			RampType(x) = Empty
			StopSound("RampLoop" & x)
			StopSound("wireloop" & x)
		end If
		'if RampBalls(x,1) = Not IsEmpty(Rampballs(x,1) then ballcount = ballcount + 1
		if not IsEmpty(Rampballs(x,1)) then ballcount = ballcount + 1
	next
	if BallCount = 0 then RampBalls(0,0) = False	'if no balls in queue, disable timer update
End Sub

Sub RampRoll_Timer():RampRollUpdate:End Sub

Sub RampRollUpdate()		'Timer update
	dim x : for x = 1 to uBound(RampBalls)
		if Not IsEmpty(RampBalls(x,1) ) then 
			if BallVel(RampBalls(x,0) ) > 1 then ' if ball is moving, play rolling sound
				If RampType(x) then 
					PlaySound("RampLoop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))				
					StopSound("wireloop" & x)
				Else
					StopSound("RampLoop" & x)
					PlaySound("wireloop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
				End If
				RampBalls(x, 2)	= RampBalls(x, 2) + 1
			Else
				StopSound("RampLoop" & x)
				StopSound("wireloop" & x)
			end if
			if RampBalls(x,0).Z < 30 and RampBalls(x, 2) > RampMinLoops then	'if ball is on the PF, remove  it
				StopSound("RampLoop" & x)
				StopSound("wireloop" & x)
				Wremoveball RampBalls(x,1)
			End If
		Else
			StopSound("RampLoop" & x)
			StopSound("wireloop" & x)
		end if
	next
	if not RampBalls(0,0) then RampRoll.enabled = 0

End Sub

' This can be used to debug the Ramp Roll time.  You need to enable the tbWR timer on the TextBox
Sub tbWR_Timer()	'debug textbox
	me.text =	"on? " & RampBalls(0, 0) & " timer: " & RampRoll.Enabled & vbnewline & _
	"1 " & Typename(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & " type:" & RampType(1) & " Loops:" & RampBalls(1, 2) & vbnewline & _
	"2 " & Typename(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & " type:" & RampType(2) & " Loops:" & RampBalls(2, 2) & vbnewline & _
	"3 " & Typename(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & " type:" & RampType(3) & " Loops:" & RampBalls(3, 2) & vbnewline & _
	"4 " & Typename(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & " type:" & RampType(4) & " Loops:" & RampBalls(4, 2) & vbnewline & _
	"5 " & Typename(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & " type:" & RampType(5) & " Loops:" & RampBalls(5, 2) & vbnewline & _
	"6 " & Typename(RampBalls(6, 0)) & " ID:" & RampBalls(6, 1) & " type:" & RampType(6) & " Loops:" & RampBalls(6, 2) & vbnewline & _
	" "
End Sub


Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
    BallPitch = pSlope(BallVel(ball), 1, -1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
	BallPitchV = pSlope(BallVel(ball), 1, -4000, 60, 7000)
End Function

'//////////////////////////////////////////////////////////////////////
'// RAMP TRIGGERS
'//////////////////////////////////////////////////////////////////////

Sub ramptrigger01_hit()
	WireRampOn True 'Play Plastic Ramp Sound
End Sub

Sub ramptrigger001_hit()
	WireRampOn True 'Play Plastic Ramp Sound
End Sub
Sub rampTrigger001_UnHit()
	if activeball.vely > 0 then WireRampOff
End Sub

Sub ramptrigger0001_hit()
	WireRampOn True 'Play Plastic Ramp Sound
End Sub


Sub ramptrigger004_hit()
	PlaysoundAtVol "Loop2",rampTrigger004, 0.04
End Sub

Sub ramptrigger005_hit()
	PlaysoundAtVol "Loop2",rampTrigger005,  0.04
End Sub

Sub rampTrigger006_Hit
	WireRampOn True 'Play Plastic Ramp Sound
End Sub
Sub rampTrigger006_UnHit
	if activeball.vely > 0 then WireRampOff
End Sub


'Wire_Start

Sub ramptrigger02_hit()
	WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

Sub ramptrigger02_unhit()
	WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
End Sub

Sub ramptrigger002_hit()
	WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

Sub ramptrigger002_unhit()
	WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
End Sub

Sub ramptrigger0002_hit()
	WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

Sub ramptrigger0002_unhit()
	WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
End Sub

Sub rampTrigger007_Hit()
	RandomSoundMetal
	WireRampOn False ' Wire ramp sound for subway
End Sub

Sub ScoopDrop_hit
    PlaySoundAtVol SoundFX("z_RCT_Scoop_Fall_Skillshot", DOFContactors), sw48, RecordedVolume
End Sub


'Wire_End

Sub ramptrigger03_hit()
	WireRampOff ' Exiting Wire Ramp Stop Playing Sound
End Sub

Sub ramptrigger03_unhit()
	PlaySoundAt "WireRamp_Stop", ramptrigger03
End Sub

Sub ramptrigger003_hit()
	WireRampOff ' Exiting Wire Ramp Stop Playing Sound
End Sub

Sub ramptrigger003_unhit()
	PlaySoundAt "WireRamp_Stop", ramptrigger003
End Sub

Sub ramptrigger0003_hit()
	WireRampOff ' Exiting Wire Ramp Stop Playing Sound
End Sub

Sub ramptrigger0003_unhit()
	PlaySoundAt "WireRamp_Stop", ramptrigger0003
End Sub


'//////////////////////////////////////////////////////////////////////
'// Ball Rolling
'//////////////////////////////////////////////////////////////////////

ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

Sub InitRolling
	Dim i
	For i = 0 To tnob
		rolling(i) = False
	Next
End Sub

Sub RollingUpdate()
	Dim b
	'   Dim BOT
	'   BOT = GetBalls
	
	' stop the sound of deleted balls
	For b = UBound(gBOT) + 1 To tnob - 1
		rolling(b) = False
		StopSound("BallRoll_" & b)
	Next
	
	' exit the sub if no balls on the table
	If UBound(gBOT) =  - 1 Then Exit Sub
	
	' play the rolling sound for each ball
	For b = 0 To UBound(gBOT)
		If BallVel(gBOT(b)) > 1 And gBOT(b).z < 30 Then
			rolling(b) = True
			PlaySound ("BallRoll_" & b), - 1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))
		Else
			If rolling(b) = True Then
				StopSound("BallRoll_" & b)
				rolling(b) = False
			End If
		End If
		
		' Ball Drop Sounds
		If gBOT(b).VelZ <  - 1 And gBOT(b).z < 55 And gBOT(b).z > 27 Then 'height adjust for ball drop sounds
			If DropCount(b) >= 5 Then
				DropCount(b) = 0
				If gBOT(b).velz >  - 7 Then
					RandomSoundBallBouncePlayfieldSoft gBOT(b)
				Else
					RandomSoundBallBouncePlayfieldHard gBOT(b)
				End If
			End If
		End If
		
		If DropCount(b) < 5 Then
			DropCount(b) = DropCount(b) + 1
		End If

	Next
End Sub

'//////////////////////////////////////////////////////////////////////
'// Mechanic Sounds
'//////////////////////////////////////////////////////////////////////

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for the TOM table

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.  
' Create the following new collections:
' 	Metals (all metal objects, metal walls, metal posts, metal wire guides)
' 	Apron (the apron walls and plunger wall)
' 	Walls (all wood or plastic walls)
' 	Rollovers (wire rollover triggers, star triggers, or button triggers)
' 	Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
' 	Gates (plate gates)
' 	GatesWire (wire gates)
' 	Rubbers (all rubbers including posts, sleeves, pegs, and bands)
' When creating the collections, make sure "Fire events for this collection" is checked.  
' You'll also need to make sure "Has Hit Event" is checked for each object placed in these collections (not necessary for gates and triggers).  
' Once the collections and objects are added, the save, close, and restart VPX.
'
' Many places in the script need to be modified to include the correct sound effect subroutine calls. The tutorial videos linked below demonstrate 
' how to make these updates. But in summary the following needs to be updated:	
'	- Nudging, plunger, coin-in, start button sounds will be added to the keydown and keyup subs.
'	- Flipper sounds in the flipper solenoid subs. Flipper collision sounds in the flipper collide subs.
'	- Bumpers, slingshots, drain, ball release, knocker, spinner, and saucers in their respective subs
'	- Ball rolling sounds sub
'
' Tutorial vides by Apophis
' Part 1: 	https://youtu.be/PbE2kNiam3g
' Part 2: 	https://youtu.be/B5cm1Y8wQsk
' Part 3: 	https://youtu.be/eLhWyuYOyGg


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
BallReleaseSoundLevel = 0.35												'volume level; range [0, 1]
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
	VolPlayfieldRoll = RollingSoundFactor * 0.0005 * Csng(BallVel(ball) ^3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
	PitchPlayfieldRoll = BallVel(ball) ^2 * 15
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
	PlaySoundAtLevelStatic ("Drain_" & Int(Rnd*4)+1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
	PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd*6)+1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundSlingshotLeft(sling)
	PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd*1)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
	PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd*1)+1,DOFContactors), SlingshotSoundLevel, Sling
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
'Sub OnBallBallCollision(ball1, ball2, velocity)
'	Dim snd
'	Select Case Int(Rnd*7)+1
'		Case 1 : snd = "Ball_Collide_1"
'		Case 2 : snd = "Ball_Collide_2"
'		Case 3 : snd = "Ball_Collide_3"
'		Case 4 : snd = "Ball_Collide_4"
'		Case 5 : snd = "Ball_Collide_5"
'		Case 6 : snd = "Ball_Collide_6"
'		Case 7 : snd = "Ball_Collide_7"
'	End Select
'
'	PlaySound (snd), 0, Csng(velocity) ^2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
'End Sub


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
'****  DROP TARGETS by Rothbauerw
'******************************************************
' This solution improves the physics for drop targets to create more realistic behavior. It allows the ball 
' to move through the target enabling the ability to score more than one target with a well placed shot.
' It also handles full drop target animation, including deflection on hit and a slight lift when the drop 
' targets raise, switch handling, bricking, and popping the ball up if it's over the drop target when it raises.
'
'Add a Timer named DTAnim to editor to handle drop & standup target animations, or run them off an always-on 10ms timer (GameTimer)
'DTAnim.interval = 10
'DTAnim.enabled = True

'Sub DTAnim_Timer
'	DoDTAnim
'	DoSTAnim
'End Sub

' For each drop target, we'll use two wall objects for physics calculations and one primitive for visuals and   
' animation. We will not use target objects.  Place your drop target primitive the same as you would a VP drop target. 
' The primitive should have it's pivot point centered on the x and y axis and at or just below the playfield 
' level on the z axis. Orientation needs to be set using Rotz and bending deflection using Rotx. You'll find a hooded 
' target mesh in this table's example. It uses the same texture map as the VP drop targets.

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
Dim DT30, DT31, DT32, DT39

'Set array with drop target objects
'
'DropTargetvar = Array(primary, secondary, prim, swtich, animate)
'   primary:	primary target wall to determine drop
'   secondary:  wall used to simulate the ball striking a bent or offset target after the initial Hit
'   prim:	   primitive target used for visuals and animation
'				   IMPORTANT!!!
'				   rotz must be used for orientation
'				   rotx to bend the target back
'				   transz to move it up and down
'				   the pivot point should be in the center of the target on the x, y and at or below the playfield (0) on z
'   switch:	 ROM switch number
'   animate:	Array slot for handling the animation instrucitons, set to 0
'				   Values for animate: 1 - bend target (hit to primary), 2 - drop target (hit to secondary), 3 - brick target (high velocity hit to secondary), -1 - raise target
'   isDropped:  Boolean which determines whether a drop target is dropped. Set to false if they are initially raised, true if initially dropped.
'					Use the function DTDropped(switchid) to check a target's drop status.

Set DT30 = (new DropTarget)(sw30, sw30a, BM_sw30, 30, 0, False)
Set DT31 = (new DropTarget)(sw31, sw31a, BM_sw31, 31, 0, False)
Set DT32 = (new DropTarget)(sw32, sw32a, BM_sw32, 32, 0, False)
Set DT39 = (new DropTarget)(sw39, sw39a, BM_sw39, 39, 0, False)

Dim DTArray
DTArray = Array(DT30,DT31,DT32,DT39)


'Configure the behavior of Drop Targets.
Const DTDropSpeed = 80 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 55 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick
Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
Const DTMass = 0.2 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance

'******************************************************
'  DROP TARGETS FUNCTIONS
'******************************************************

Sub DTHit(switch)
	Dim i
	i = DTArrayID(switch)
	
	DTArray(i).animate = DTCheckBrick(ActiveBall,DTArray(i).prim)
	If DTArray(i).animate = 1 Or DTArray(i).animate = 3 Or DTArray(i).animate = 4 Then
		DTBallPhysics ActiveBall, DTArray(i).prim.rotz, DTMass
	End If
	DoDTAnim

End Sub

Sub DTRaise(switch)
	Dim i
	i = DTArrayID(switch)
	
	DTArray(i).animate =  - 1
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
	For i = 0 To UBound(DTArray)
		If DTArray(i).sw = switch Then
			DTArrayID = i
			Exit Function
		End If
	Next
End Function

Sub DTBallPhysics(aBall, angle, mass)
	Dim rangle,bangle,calc1, calc2, calc3
	rangle = (angle - 90) * 3.1416 / 180
	bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
	
	calc1 = cor.BallVel(aball.id) * Cos(bangle - rangle) * (aball.mass - mass) / (aball.mass + mass)
	calc2 = cor.BallVel(aball.id) * Sin(bangle - rangle) * Cos(rangle + 4 * Atn(1) / 2)
	calc3 = cor.BallVel(aball.id) * Sin(bangle - rangle) * Sin(rangle + 4 * Atn(1) / 2)
	
	aBall.velx = calc1 * Cos(rangle) + calc2
	aBall.vely = calc1 * Sin(rangle) + calc3
End Sub

'Check if target is hit on it's face or sides and whether a 'brick' occurred
Function DTCheckBrick(aBall, dtprim)
	Dim bangle, bangleafter, rangle, rangle2, Xintersect, Yintersect, cdist, perpvel, perpvelafter, paravel, paravelafter
	rangle = (dtprim.rotz - 90) * 3.1416 / 180
	rangle2 = dtprim.rotz * 3.1416 / 180
	bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
	bangleafter = Atn2(aBall.vely,aball.velx)
	
	Xintersect = (aBall.y - dtprim.y - Tan(bangle) * aball.x + Tan(rangle2) * dtprim.x) / (Tan(rangle2) - Tan(bangle))
	Yintersect = Tan(rangle2) * Xintersect + (dtprim.y - Tan(rangle2) * dtprim.x)
	
	cdist = Distance(dtprim.x, dtprim.y, Xintersect, Yintersect)
	
	perpvel = cor.BallVel(aball.id) * Cos(bangle - rangle)
	paravel = cor.BallVel(aball.id) * Sin(bangle - rangle)
	
	perpvelafter = BallSpeed(aBall) * Cos(bangleafter - rangle)
	paravelafter = BallSpeed(aBall) * Sin(bangleafter - rangle)
	
	If perpvel > 0 And  perpvelafter <= 0 Then
		If DTEnableBrick = 1 And  perpvel > DTBrickVel And DTBrickVel <> 0 And cdist < 8 Then
			DTCheckBrick = 3
		Else
			DTCheckBrick = 1
		End If
	ElseIf perpvel > 0 And ((paravel > 0 And paravelafter > 0) Or (paravel < 0 And paravelafter < 0)) Then
		DTCheckBrick = 4
	Else
		DTCheckBrick = 0
	End If
End Function

Sub DoDTAnim()
	Dim i
	For i = 0 To UBound(DTArray)
		DTArray(i).animate = DTAnimate(DTArray(i).primary,DTArray(i).secondary,DTArray(i).prim,DTArray(i).sw,DTArray(i).animate)
	Next
If DTDropped(39) then
WallDrop39.collidable = 0
Else
WallDrop39.collidable = 1
End if

If DTDropped(32) then
WallDrop32.collidable = 0
Else
WallDrop32.collidable = 1
End if

If DTDropped(31) then
WallDrop31.collidable = 0
Else
WallDrop31.collidable = 1
End if

If DTDropped(30) then
WallDrop30.collidable = 0
Else
WallDrop30.collidable = 1
End if
End Sub

Function DTAnimate(primary, secondary, prim, switch, animate)
	Dim transz, switchid
	Dim animtime, rangle
	
	switchid = switch
	
	Dim ind
	ind = DTArrayID(switchid)
	
	rangle = prim.rotz * PI / 180
	
	DTAnimate = animate
	
	If animate = 0 Then
		primary.uservalue = 0
		DTAnimate = 0
		Exit Function
	ElseIf primary.uservalue = 0 Then
		primary.uservalue = GameTime
	End If
	
	animtime = GameTime - primary.uservalue
	
	If (animate = 1 Or animate = 4) And animtime < DTDropDelay Then
		primary.collidable = 0
		If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0
		prim.rotx = DTMaxBend * Cos(rangle)
		prim.roty = DTMaxBend * Sin(rangle)
		DTAnimate = animate
		Exit Function
	ElseIf (animate = 1 Or animate = 4) And animtime > DTDropDelay Then
		primary.collidable = 0
		If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0
		prim.rotx = DTMaxBend * Cos(rangle)
		prim.roty = DTMaxBend * Sin(rangle)
		animate = 2
'		SoundDropTargetDrop prim
	End If
	
	If animate = 2 Then
		transz = (animtime - DTDropDelay) / DTDropSpeed * DTDropUnits *  - 1
		If prim.transz >  - DTDropUnits  Then
			prim.transz = transz
		End If
		
		prim.rotx = DTMaxBend * Cos(rangle) / 2
		prim.roty = DTMaxBend * Sin(rangle) / 2
		
		If prim.transz <= - DTDropUnits Then
			prim.transz =  - DTDropUnits
			secondary.collidable = 0
			DTArray(ind).isDropped = True 'Mark target as dropped
			controller.Switch(Switchid) = 1
			primary.uservalue = 0
			DTAnimate = 0
			Exit Function
		Else
			DTAnimate = 2
			Exit Function
		End If
	End If
	
	If animate = 3 And animtime < DTDropDelay Then
		primary.collidable = 0
		secondary.collidable = 1
		prim.rotx = DTMaxBend * Cos(rangle)
		prim.roty = DTMaxBend * Sin(rangle)
	ElseIf animate = 3 And animtime > DTDropDelay Then
		primary.collidable = 1
		secondary.collidable = 0
		prim.rotx = 0
		prim.roty = 0
		primary.uservalue = 0
		DTAnimate = 0
		Exit Function
	End If
	
	If animate =  - 1 Then
		transz = (1 - (animtime) / DTDropUpSpeed) * DTDropUnits *  - 1
		
		If prim.transz =  - DTDropUnits Then
			Dim b
			Dim gBOT
			gBOT = GetBalls
			
			For b = 0 To UBound(gBOT)
				If InRotRect(gBOT(b).x,gBOT(b).y,prim.x, prim.y, prim.rotz, - 25, - 10,25, - 10,25,25, - 25,25) And gBOT(b).z < prim.z + DTDropUnits + 25 Then
					gBOT(b).velz = 20
				End If
			Next
		End If
		
		If prim.transz < 0 Then
			prim.transz = transz
		ElseIf transz > 0 Then
			prim.transz = transz
		End If
		
		If prim.transz > DTDropUpUnits Then
			DTAnimate =  - 2
			prim.transz = DTDropUpUnits
			prim.rotx = 0
			prim.roty = 0
			primary.uservalue = GameTime
		End If
		primary.collidable = 0
		secondary.collidable = 1
		DTArray(ind).isDropped = False 'Mark target as not dropped
		controller.Switch(Switchid) = 0
	End If
	
	If animate =  - 2 And animtime > DTRaiseDelay Then
		prim.transz = (animtime - DTRaiseDelay) / DTDropSpeed * DTDropUnits *  - 1 + DTDropUpUnits
		If prim.transz < 0 Then
			prim.transz = 0
			primary.uservalue = 0
			DTAnimate = 0
			
			primary.collidable = 1
			secondary.collidable = 0
		End If
	End If
End Function

Function DTDropped(switchid)
	Dim ind
	ind = DTArrayID(switchid)
	
	DTDropped = DTArray(ind).isDropped
End Function

Sub UpdateDropTargets
	dim BP, tz, rx, ry

    tz = BM_sw30.transz
	rx = BM_sw30.rotx
	ry = BM_sw30.roty
	For each BP in BP_sw30 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

    tz = BM_sw31.transz
	rx = BM_sw31.rotx
	ry = BM_sw31.roty
	For each BP in BP_sw31 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

    tz = BM_sw32.transz
	rx = BM_sw32.rotx
	ry = BM_sw32.roty
	For each BP in BP_sw32 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

    tz = BM_sw39.transz
	rx = BM_sw39.rotx
	ry = BM_sw39.roty
	For each BP in BP_sw39 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next



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
'	ZRST: STAND-UP TARGETS by Rothbauerw
'******************************************************

Class StandupTarget
  Private m_primary, m_prim, m_sw, m_animate

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(val): Set m_primary = val: End Property

  Public Property Get Prim(): Set Prim = m_prim: End Property
  Public Property Let Prim(val): Set m_prim = val: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(val): m_sw = val: End Property

  Public Property Get Animate(): Animate = m_animate: End Property
  Public Property Let Animate(val): m_animate = val: End Property

  Public default Function init(primary, prim, sw, animate)
    Set m_primary = primary
    Set m_prim = prim
    m_sw = sw
    m_animate = animate

    Set Init = Me
  End Function
End Class


'Define a variable for each stand-up target
Dim ST17, ST18, ST19, ST22, ST28, ST29, ST40, ST44, ST45, ST46

'Set array with stand-up target objects
'
'StandupTargetvar = Array(primary, prim, swtich)
'   primary:	vp target to determine target hit
'   prim:	   primitive target used for visuals and animation
'				   IMPORTANT!!! 
'				   transy must be used to offset the target animation
'   switch:	 ROM switch number
'   animate:	Arrary slot for handling the animation instrucitons, set to 0
' 
'You will also need to add a secondary hit object for each stand up (name sw11o, sw12o, and sw13o on the example Table1)
'these are inclined primitives to simulate hitting a bent target and should provide so z velocity on high speed impacts

Set ST17 = (new StandupTarget)(sw17, BM_sw17, 17, 0)
Set ST18 = (new StandupTarget)(sw18, BM_sw18, 18, 0)
Set ST19 = (new StandupTarget)(sw19, BM_sw19, 19, 0)
Set ST22 = (new StandupTarget)(sw22, BM_sw22, 22, 0)
Set ST28 = (new StandupTarget)(sw28, BM_sw28, 28, 0)
Set ST29 = (new StandupTarget)(sw29, BM_sw29, 29, 0)
Set ST40 = (new StandupTarget)(sw40, BM_sw40, 40, 0)
Set ST44 = (new StandupTarget)(sw44, BM_sw44, 44, 0)
Set ST45 = (new StandupTarget)(sw45, BM_sw45, 45, 0)
Set ST46 = (new StandupTarget)(sw46, BM_sw46, 46, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST17, ST18, ST19, ST22, ST28, ST29, ST40, ST44, ST45, ST46)

'Configure the behavior of Stand-up Targets
Const STAnimStep = 1.5  'vpunits per animation step (control return to Start)
Const STMaxOffset = 9   'max vp units target moves when hit

Const STMass = 0.2	  'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

'******************************************************
'				STAND-UP TARGETS FUNCTIONS
'******************************************************

Sub STHit(switch)
	Dim i
	i = STArrayID(switch)
	
	PlayTargetSound
	STArray(i).animate = STCheckHit(Activeball,STArray(i).primary)
	
	If STArray(i).animate <> 0 Then
		DTBallPhysics Activeball, STArray(i).primary.orientation, STMass
	End If
	DoSTAnim
End Sub

Function STArrayID(switch)
	Dim i
	For i = 0 To UBound(STArray)
		If STArray(i).sw = switch Then 
			STArrayID = i
			Exit Function
		End If
	Next
End Function

Function STCheckHit(aBall, target) 'Check if target is hit on it's face
	Dim bangle, bangleafter, rangle, rangle2, perpvel, perpvelafter, paravel, paravelafter
	rangle = (target.orientation - 90) * 3.1416 / 180
	bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
	bangleafter = Atn2(aBall.vely,aball.velx)
	
	perpvel = cor.BallVel(aball.id) * Cos(bangle - rangle)
	paravel = cor.BallVel(aball.id) * Sin(bangle - rangle)
	
	perpvelafter = BallSpeed(aBall) * Cos(bangleafter - rangle)
	paravelafter = BallSpeed(aBall) * Sin(bangleafter - rangle)
	
	If perpvel > 0 And  perpvelafter <= 0 Then
		STCheckHit = 1
	ElseIf perpvel > 0 And ((paravel > 0 And paravelafter > 0) Or (paravel < 0 And paravelafter < 0)) Then
		STCheckHit = 1
	Else
		STCheckHit = 0
	End If
End Function

Sub DoSTAnim()
	Dim i
	For i = 0 To UBound(STArray)
		STArray(i).animate = STAnimate(STArray(i).primary,STArray(i).prim,STArray(i).sw,STArray(i).animate)
	Next
End Sub

Function STAnimate(primary, prim, switch,  animate)
	Dim animtime
	
	STAnimate = animate
	
	If animate = 0  Then
		primary.uservalue = 0
		STAnimate = 0
		Exit Function
	ElseIf primary.uservalue = 0 Then
		primary.uservalue = gametime
	End If
	
	animtime = gametime - primary.uservalue
	
	If animate = 1 Then
		primary.collidable = 0
		prim.transy =  - STMaxOffset
		vpmTimer.PulseSw switch
		STAnimate = 2
		Exit Function
	ElseIf animate = 2 Then
		prim.transy = prim.transy + STAnimStep
		If prim.transy >= 0 Then
			prim.transy = 0
			primary.collidable = 1
			STAnimate = 0
			Exit Function
		Else
			STAnimate = 2
		End If
	End If
End Function


Sub UpdateStandupTargets
	dim BP, y

    y = BM_sw17.transy
	For Each BP in BP_sw17 : BP.transy = y: Next

    y = BM_sw18.transy
	For Each BP in BP_sw18 : BP.transy = y: Next

    y = BM_sw19.transy
	For Each BP in BP_sw19 : BP.transy = y: Next

    y = BM_sw22.transy
	For Each BP in BP_sw22 : BP.transy = y: Next

    y = BM_sw28.transy
	For Each BP in BP_sw28 : BP.transy = y: Next

    y = BM_sw29.transy
	For Each BP in BP_sw29 : BP.transy = y: Next

    y = BM_sw40.transy
	For Each BP in BP_sw40 : BP.transy = y: Next

    y = BM_sw44.transy
	For Each BP in BP_sw44 : BP.transy = y: Next

    y = BM_sw45.transy
	For Each BP in BP_sw45 : BP.transy = y: Next

    y = BM_sw46.transy
	For Each BP in BP_sw46 : BP.transy = y: Next

End Sub


'******************************************************
'****   END STAND-UP TARGETS
'******************************************************




'******************************************************
'  SLINGSHOT CORRECTION FUNCTIONS
'******************************************************
' To add these slingshot corrections:
' 	- On the table, add the endpoint primitives that define the two ends of the Slingshot
'	- Initialize the SlingshotCorrection objects in InitSlingCorrection
' 	- Call the .VelocityCorrect methods from the respective _Slingshot event sub


dim LS : Set LS = New SlingshotCorrection
dim RS : Set RS = New SlingshotCorrection

InitSlingCorrection

Sub InitSlingCorrection

	LS.Object = LeftSlingshot
	LS.EndPoint1 = EndPoint1LS
	LS.EndPoint2 = EndPoint2LS

	RS.Object = RightSlingshot
	RS.EndPoint1 = EndPoint1RS
	RS.EndPoint2 = EndPoint2RS

	'Slingshot angle corrections (pt, BallPos in %, Angle in deg)
	' These values are best guesses. Retune them if needed based on specific table research.
	AddSlingsPt 0, 0.00,	-4
	AddSlingsPt 1, 0.45,	-7
	AddSlingsPt 2, 0.48,	0
	AddSlingsPt 3, 0.52,	0
	AddSlingsPt 4, 0.55,	7
	AddSlingsPt 5, 1.00,	4

End Sub


Sub AddSlingsPt(idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
	dim a : a = Array(LS, RS)
	dim x : for each x in a
		x.addpoint idx, aX, aY
	Next
End Sub

'' The following sub are needed, however they may exist somewhere else in the script. Uncomment below if needed
'Dim PI: PI = 4*Atn(1)
'Function dSin(degrees)
'	dsin = sin(degrees * Pi/180)
'End Function
'Function dCos(degrees)
'	dcos = cos(degrees * Pi/180)
'End Function
'
'Function RotPoint(x,y,angle)
'    dim rx, ry
'    rx = x*dCos(angle) - y*dSin(angle)
'    ry = x*dSin(angle) + y*dCos(angle)
'    RotPoint = Array(rx,ry)
'End Function

Class SlingshotCorrection
	Public DebugOn, Enabled
	private Slingshot, SlingX1, SlingX2, SlingY1, SlingY2

	Public ModIn, ModOut
	Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): Enabled = True : End Sub 

	Public Property let Object(aInput) : Set Slingshot = aInput : End Property
	Public Property Let EndPoint1(aInput) : SlingX1 = aInput.x: SlingY1 = aInput.y: End Property
	Public Property Let EndPoint2(aInput) : SlingX2 = aInput.x: SlingY2 = aInput.y: End Property

	Public Sub AddPoint(aIdx, aX, aY) 
		ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
		If gametime > 100 then Report
	End Sub

	Public Sub Report()         'debug, reports all coords in tbPL.text
		If not debugOn then exit sub
		dim a1, a2 : a1 = ModIn : a2 = ModOut
		dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
		TBPout.text = str
	End Sub


	Public Sub VelocityCorrect(aBall)
		dim BallPos, XL, XR, YL, YR
		
		'Assign right and left end points
		If SlingX1 < SlingX2 Then 
			XL = SlingX1 : YL = SlingY1 : XR = SlingX2 : YR = SlingY2
		Else
			XL = SlingX2 : YL = SlingY2 : XR = SlingX1 : YR = SlingY1
		End If

		'Find BallPos = % on Slingshot
		If Not IsEmpty(aBall.id) Then 
			If ABS(XR-XL) > ABS(YR-YL) Then 
				BallPos = PSlope(aBall.x, XL, 0, XR, 1)
			Else
				BallPos = PSlope(aBall.y, YL, 0, YR, 1)
			End If
			If BallPos < 0 Then BallPos = 0
			If BallPos > 1 Then BallPos = 1
		End If

		'Velocity angle correction
		If not IsEmpty(ModIn(0) ) then
			Dim Angle, RotVxVy
			Angle = LinearEnvelope(BallPos, ModIn, ModOut)
			'debug.print " BallPos=" & BallPos &" Angle=" & Angle 
			'debug.print " BEFORE: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely 
			RotVxVy = RotPoint(aBall.Velx,aBall.Vely,Angle)
			If Enabled then aBall.Velx = RotVxVy(0)
			If Enabled then aBall.Vely = RotVxVy(1)
			'debug.print " AFTER: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely 
			'debug.print " " 
		End If
	End Sub

End Class





'***************************************************************
'	ZSHA: Ambient ball shadows
'***************************************************************

' For dynamic ball shadows, Check the "Raytraced ball shadows" box for the specific light. 
' Also make sure the light's z position is around 25 (mid ball)

'Ambient (Room light source)
Const AmbientBSFactor = 0.9    '0 To 1, higher is darker
Const AmbientMovement = 1	   '1+ higher means more movement as the ball moves left and right
Const offsetX = 0			   'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY = 0			   'Offset y position under ball (^^for example 5,5 if the light is in the back left corner)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!  (will throw errors if there aren't enough objects)
Dim objBallShadow(7)

'Initialization
BSInit

Sub BSInit()
	Dim iii
	'Prepare the shadow objects before play begins
	For iii = 0 To tnob - 1
		Set objBallShadow(iii) = Eval("BallShadow" & iii)
		objBallShadow(iii).material = "BallShadow" & iii
		UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
		objBallShadow(iii).Z = 3 + iii / 1000
		objBallShadow(iii).visible = 0
	Next
End Sub


Sub BSUpdate
	Dim s: For s = lob To UBound(gBOT)
		' *** Normal "ambient light" ball shadow
		
		'Primitive shadow on playfield, flasher shadow in ramps
		'** If on main and upper pf
		If gBOT(s).Z > 20 And gBOT(s).Z < 30 Then
			objBallShadow(s).visible = 1
			objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
			objBallShadow(s).Y = gBOT(s).Y + offsetY
			'objBallShadow(s).Z = gBOT(s).Z + s/1000 + 1.04 - 25	

		'** No shadow if ball is off the main playfield (this may need to be adjusted per table)
		Else
			objBallShadow(s).visible = 0
		End If
	Next
End Sub



'************
' Led board
'************
Dim LED(46)

LED(0) = Array(dmx85, dmx64, dmx43, dmx22, dmx1)
LED(1) = Array(dmx86, dmx65, dmx44, dmx23, dmx2)
LED(2) = Array(dmx87, dmx66, dmx45, dmx24, dmx3)
LED(3) = Array(dmx88, dmx67, dmx46, dmx25, dmx4)
LED(4) = Array(dmx89, dmx68, dmx47, dmx26, dmx5)
LED(5) = Array(dmx90, dmx69, dmx48, dmx27, dmx6)
LED(6) = Array(dmx91, dmx70, dmx49, dmx28, dmx7)
LED(7) = Array(dmx92, dmx71, dmx50, dmx29, dmx8)
LED(8) = Array(dmx93, dmx72, dmx51, dmx30, dmx9)
LED(9) = Array(dmx94, dmx73, dmx52, dmx31, dmx10)
LED(10) = Array(dmx95, dmx74, dmx53, dmx32, dmx11)
LED(11) = Array(dmx96, dmx75, dmx54, dmx33, dmx12)
LED(12) = Array(dmx97, dmx76, dmx55, dmx34, dmx13)
LED(13) = Array(dmx98, dmx77, dmx56, dmx35, dmx14)
LED(14) = Array(dmx99, dmx78, dmx57, dmx36, dmx15)
LED(15) = Array(dmx100, dmx79, dmx58, dmx37, dmx16)
LED(16) = Array(dmx101, dmx80, dmx59, dmx38, dmx17)
LED(17) = Array(dmx102, dmx81, dmx60, dmx39, dmx18)
LED(18) = Array(dmx103, dmx82, dmx61, dmx40, dmx19)
LED(19) = Array(dmx104, dmx83, dmx62, dmx41, dmx20)
LED(20) = Array(dmx105, dmx84, dmx63, dmx42, dmx21)


Sub UpdateLeds
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLED=Controller.ChangedLEDs(&H00000000, &Hffffffff)
	If Not IsEmpty(ChgLED) Then
		For ii = 0 To UBound(chgLED)
			num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
				For Each obj In LED(num)
					If chg And 1 Then obj.state = stat And 1 
					chg = chg\4:stat = stat\4
				Next
		Next
	End If

End Sub

Sub L64_animate
    Dim p : p = L64.GetInPlayIntensity / L64.Intensity
    startbutton.BlendDisableLighting = p
End Sub



'******************************************************
'	ZDRN: Drain, Trough, and Ball Release
'******************************************************

'********************* TROUGH *************************

Sub sw11_Hit   : Controller.Switch(11) = 1 : UpdateTrough : End Sub
Sub sw11_UnHit : Controller.Switch(11) = 0 : UpdateTrough : End Sub
Sub sw12_Hit   : Controller.Switch(12) = 1 : UpdateTrough : End Sub
Sub sw12_UnHit : Controller.Switch(12) = 0 : UpdateTrough : End Sub
Sub sw13_Hit   : Controller.Switch(13) = 1 : UpdateTrough : End Sub
Sub sw13_UnHit : Controller.Switch(13) = 0 : UpdateTrough : End Sub
Sub sw14_Hit   : Controller.Switch(14) = 1 : UpdateTrough : End Sub
Sub sw14_UnHit : Controller.Switch(14) = 0 : UpdateTrough : End Sub
Sub Drain_Hit  : UpdateTrough : RandomSoundDrain Drain : End Sub

Sub UpdateTrough
	UpdateTroughTimer.Interval = 100
	UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer
	If sw14.BallCntOver = 0 Then sw13.kick 57, 10
	If sw13.BallCntOver = 0 Then sw12.kick 57, 10
	If sw12.BallCntOver = 0 Then sw11.kick 57, 10
	If sw11.BallCntOver = 0 Then Drain.kick 57, 10
	UpdateTroughTimer.Enabled = 0
End Sub


'*****************  DRAIN & RELEASE  ******************

Sub SolRelease(enabled)
	If enabled Then 
		sw14.kick 57, 10	
		Controller.Switch(15) = 1
		sw14.Timerenabled = True
		PlaySoundAtVol SoundFX("z_RCT_Ballrelease", DOFContactors), sw14, RecordedVolume
	End If
End Sub

Sub sw14_Timer
	sw14.Timerenabled = False
	Controller.Switch(15) = 0
End Sub

'************************************************************
'	ZVUK: VUKs and Kickers
'************************************************************

Dim KickerBall47, KickerBall43, KickerBall52

Sub KickBall(kball, kangle, kvel, kvelz, kzlift)
	dim rangle
	rangle = PI * (kangle - 90) / 180
    
	kball.z = kball.z + kzlift
	kball.velz = kvelz
	kball.velx = cos(rangle)*kvel
	kball.vely = sin(rangle)*kvel
End Sub

'************************************************************
'	Scoop
'************************************************************

Sub sw47_Hit
	WireRampOff
    PlaySoundAtVol SoundFX("z_RCT_Scoop_Fall", DOFContactors), sw47, RecordedVolume
    set KickerBall47 = activeball
    Controller.Switch(47) = 1
End Sub

Sub sw47_UnHit
    Controller.Switch(47) = 0
End Sub

Sub KioskScoop(Enable)
    If Enable then
		If Controller.Switch(47) <> 0 Then
            KickBall KickerBall47, 20, 92, 0, 0
    PlaySoundAtVol SoundFX("z_RCT_Scoop", DOFContactors), sw47, RecordedVolume
		End If
	End If
End Sub

'************************************************************
'	Ball Lock
'************************************************************

Sub sw43_Hit
    set KickerBall43 = activeball
    Controller.Switch(43) = 1
End Sub

Sub sw43_UnHit
    Controller.Switch(43) = 0
End Sub

Sub BallLock(Enable)
    If Enable then
		If Controller.Switch(43) <> 0 Then
            KickBall KickerBall43, 0, 50, 0, 0
			PlaySoundAtVol SoundFX("z_RCT_Autoplunger_Left", DOFContactors), sw43, RecordedVolume
		End If
	End If
End Sub

'************************************************************
'	Rocket
'************************************************************

Sub sw52_Hit
    PlaySoundAtVol SoundFX("z_RCT_Scoop_Fall", DOFContactors), sw47, RecordedVolume
    set KickerBall52 = activeball
    Controller.Switch(52) = 1
End Sub

Sub sw52_UnHit
    Controller.Switch(52) = 0
End Sub

Sub Rocket(Enable)
    If Enable then
		If Controller.Switch(52) <> 0 Then
            KickBall KickerBall52, 0, 0, 38, 0
    PlaySoundAtVol SoundFX("z_RCT_Rocket_Launch", DOFContactors), sw52, RecordedVolume
		End If
	End If
End Sub


'************************************************************
'	Gate1
'************************************************************

Sub GateSET_Hit 
	GateSE.Collidable = 0
	If activeball.vely < 0 Then Gate1.Collidable = 1
End Sub

Sub GateSET_unHit 
	GateSE.Collidable = 1
	Gate1.Collidable = 0
End Sub

'Fix for stuck ball at Gate1
Sub GateSETfix_Hit: GateSETfix.Timerenabled = true: End Sub
Sub GateSETfix_UnHit: GateSETfix.Timerenabled = false: End Sub
Sub GateSETfix_Timer: Gate1.Collidable = 0: GateSETfix.Timerenabled = false: End Sub


'***************
' Scrambled Eggs
'***************

Dim discPosition, discSpinSpeed, discLastPos, SpinCounter, maxvel
dim spinAngle, degAngle, startAngle, postSpeedFactor
dim discX, discY
discLastPos = 1
startAngle = 0
discX = 109.0194
discY = 863.4606
PostSpeedFactor = 130'90

Const cDiscSpeedMult = 40 '35 					' Affects speed transfer to object (deg/sec)
Const cDiscFriction = 1.0    		         	' Friction coefficient (deg/sec/sec)
Const cDiscMinSpeed = 0.05 						' Object stops at this speed (deg/sec)
'Const cDiscMinSpeed = 5 						' use this value if you want to enable DOF for lamp below in script
Const cDiscRadius = 37

'Wobble
Const discSpringConst = -70 '-100
Const discSpringAngle = 30
Const discSpringRange = 25
'End Wobble

Dim SEPulseCount: SEPulseCount=0

Sub SpinnerBallTimer_Timer()
	Dim oldDiscSpeed, discFriction
	oldDiscSpeed = discSpinSpeed

	discPosition = discPosition + discSpinSpeed * Me.Interval / 1000

	if ABS(discSpinSpeed) < 200 Then
		discFriction = 6 ' was 6
	else
		discFriction = cDiscFriction
	end if 
	discSpinSpeed = discSpinSpeed * (1 - discFriction * Me.Interval / 1000)

	Do While discPosition < 0 : discPosition = discPosition + 360 : Loop
	Do While discPosition > 360 : discPosition = discPosition - 360 : Loop

	'Wobble

	Dim UpperRange, LowerRange
	UpperRange = discSpringAngle + discSpringRange
	LowerRange = discSpringAngle - discSpringRange

	If abs(discSpinSpeed) < 400 Then
		If discPosition > LowerRange and discPosition < discSpringAngle Then
			discSpinSpeed = newDiscSpinSpeed(discSpinSpeed ,discPosition - LowerRange, Me.Interval / 1000)
		ElseIf discPosition > discSpringAngle and discPosition < UpperRange  Then
			discSpinSpeed = newDiscSpinSpeed(discSpinSpeed ,discPosition - UpperRange, Me.Interval / 1000)
		ElseIf discPosition > LowerRange+180 and discPosition < discSpringAngle+180 Then
			discSpinSpeed = newDiscSpinSpeed(discSpinSpeed ,discPosition - LowerRange - 180, Me.Interval / 1000)
		ElseIf discPosition > discSpringAngle+180 and discPosition < UpperRange+180  Then
			discSpinSpeed = newDiscSpinSpeed(discSpinSpeed ,discPosition - UpperRange - 180, Me.Interval / 1000)
		End If
	End If
	'End Wobble

	If Abs(discSpinSpeed) < cDiscMinSpeed Then
		discSpinSpeed = 0
		'DOF 103,DOFOff
	Else
		'DOF 103,DOFOn
	End If

	If 0 < discPosition And discPosition <= 30 and discLastPos <> 1 Then
		vpmTimer.PulseSw 20 
		discLastPos = 1 ': debug.print "discLastPos = " & discLastPos
	ElseIf 30 < discPosition And discPosition <= 60 and discLastPos <> 2 Then
		vpmTimer.PulseSw 20 
		discLastPos = 2 ': debug.print "discLastPos = " & discLastPos
	ElseIf 60 < discPosition And discPosition <= 90 and discLastPos <> 3 Then
		vpmTimer.PulseSw 20 
		discLastPos = 3 ': debug.print "discLastPos = " & discLastPos
	ElseIf 90 < discPosition And discPosition <= 120 and discLastPos <> 4 Then
		vpmTimer.PulseSw 20 
		discLastPos = 4 ': debug.print "discLastPos = " & discLastPos
	ElseIf 120 < discPosition And discPosition <= 150 and discLastPos <> 5 Then
		vpmTimer.PulseSw 20 
		discLastPos = 5 ': debug.print "discLastPos = " & discLastPos
	ElseIf 150 < discPosition And discPosition <= 180 and discLastPos <> 6 Then
		vpmTimer.PulseSw 20 
		discLastPos = 6 ': debug.print "discLastPos = " & discLastPos
	ElseIf 180 < discPosition And discPosition <= 210 and discLastPos <> 7 Then
		vpmTimer.PulseSw 20
		discLastPos = 7 ': debug.print "discLastPos = " & discLastPos
	ElseIf 210 < discPosition And discPosition <= 240 and discLastPos <> 8 Then
		vpmTimer.PulseSw 20 
		discLastPos = 8 ': debug.print "discLastPos = " & discLastPos
	ElseIf 240 < discPosition And discPosition <= 270 and discLastPos <> 9 Then
		vpmTimer.PulseSw 20 
		discLastPos = 9 ': debug.print "discLastPos = " & discLastPos
	ElseIf 270 < discPosition And discPosition <= 300 and discLastPos <> 10 Then
		vpmTimer.PulseSw 20 
		discLastPos = 10 ': debug.print "discLastPos = " & discLastPos
	ElseIf 300 < discPosition And discPosition <= 330 and discLastPos <> 11 Then
		vpmTimer.PulseSw 20 
		discLastPos = 11 ': debug.print "discLastPos = " & discLastPos
	ElseIf 330 < discPosition And discPosition <= 360 and discLastPos <> 12 Then
		vpmTimer.PulseSw 20 
		discLastPos = 12 ': debug.print "discLastPos = " & discLastPos
	End If


	degAngle = -90  + startAngle + discPosition
	spinAngle = PI * (degAngle) / 180
	

	SpinnerBall.x = discX + (cDiscRadius * Cos(spinAngle))
	SpinnerBall.y = discY + (cDiscRadius * Sin(spinAngle))
	SpinnerBall.z = 33


	If ABS(discSpinSpeed*sin(spinAngle)/postSpeedFactor) < 0.05 Then
		SpinnerBall.velx = 0.05
	Else
		SpinnerBall.velx = - discSpinSpeed*sin(spinAngle)/postSpeedFactor
	End If

	If Abs(discSpinSpeed*cos(spinAngle)/postSpeedFactor) < 0.05 Then
		SpinnerBall.vely = 0.05
	Else
		SpinnerBall.vely = discSpinSpeed*cos(spinAngle)/postSpeedFactor		'0.05
	End If

	SpinnerBall.velz = 0

'***************

	Dim BP : For Each BP in BP_ScrambledEgg : BP.objrotz = discPosition + 25: Next

End Sub


Function newDiscSpinSpeed(spinspeed, springangle, springtime)
	newDiscSpinSpeed = spinspeed + discSpringConst * springangle * springtime
End Function





'***************
' Ball Collision
'***************


Sub OnBallBallCollision(ball1, ball2, velocity)
	dim collAngle,bvelx,bvely,hitball

	FlipperCradleCollision ball1, ball2, velocity

	If ball1.radius > 27 or ball2.radius > 27 then
		
		If ball1.radius > 27 Then
			collAngle = GetCollisionAngle(ball1.x,ball1.y,ball2.x,ball2.y)
			set hitball = ball2
		else 
			collAngle = GetCollisionAngle(ball2.x,ball2.y,ball1.x,ball1.y)
			set hitball = ball1
		End If

		dim discAngle
		discAngle = NormAngle(spinAngle)

'		discSpinSpeed = discspinspeed + ecvel(0,1.5,sin(collAngle - discAngle)*velocity,BallMass * ABS(sin(collAngle - discAngle))) * cDiscSpeedMult

'		PlaySound "fx_lamphit", 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 1, 0, AudioFade(ball1)

		dim sineOfAngle, sineOfAngleSqr
		sineOfAngle = sin(collAngle - discAngle)

		discSpinSpeed = discspinspeed + ecvel(0,1.5,sineOfAngle*velocity,BallMass) * cDiscSpeedMult

	Else
		If ball1.z > 10 and ball2.z > 10 Then
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
		End If
	End If
End Sub

Function GetCollisionAngle(ax, ay, bx, by)
	Dim ang
	Dim collisionV:Set collisionV = new jVector
	collisionV.SetXY ax - bx, ay - by
	GetCollisionAngle = collisionV.ang
End Function

Function NormAngle(angle)
	NormAngle = angle
	Dim pi:pi = 3.14159265358979323846
	Do While NormAngle>2 * pi
		NormAngle = NormAngle - 2 * pi
	Loop
	Do While NormAngle <0
		NormAngle = NormAngle + 2 * pi
	Loop
End Function
 
Class jVector
     Private m_mag, m_ang, pi
 
     Sub Class_Initialize
         m_mag = CDbl(0)
         m_ang = CDbl(0)
         pi = CDbl(3.14159265358979323846)
     End Sub
 
     Public Function add(anothervector)
         Dim tx, ty, theta
         If TypeName(anothervector) = "jVector" then
             Set add = new jVector
             add.SetXY x + anothervector.x, y + anothervector.y
         End If
     End Function
 
     Public Function multiply(scalar)
         Set multiply = new jVector
         multiply.SetXY x * scalar, y * scalar
     End Function
 
     Sub ShiftAxes(theta)
         ang = ang - theta
     end Sub
 
     Sub SetXY(tx, ty)
 
         if tx = 0 And ty = 0 Then
             ang = 0
          elseif tx = 0 And ty <0 then
             ang = - pi / 180 ' -90 degrees
          elseif tx = 0 And ty>0 then
             ang = pi / 180   ' 90 degrees
         else
             ang = atn(ty / tx)
             if tx <0 then ang = ang + pi ' Add 180 deg if in quadrant 2 or 3
         End if
 
         mag = sqr(tx ^2 + ty ^2)
     End Sub
 
     Property Let mag(nmag)
         m_mag = nmag
     End Property
 
     Property Get mag
         mag = m_mag
     End Property
 
     Property Let ang(nang)
         m_ang = nang
         Do While m_ang>2 * pi
             m_ang = m_ang - 2 * pi
         Loop
         Do While m_ang <0
             m_ang = m_ang + 2 * pi
         Loop
     End Property
 
     Property Get ang
         Do While m_ang>2 * pi
             m_ang = m_ang - 2 * pi
         Loop
         Do While m_ang <0
             m_ang = m_ang + 2 * pi
         Loop
         ang = m_ang
     End Property
 
     Property Get x
         x = m_mag * cos(ang)
     End Property
 
     Property Get y
         y = m_mag * sin(ang)
     End Property
 
     Property Get dump
         dump = "vector "
         Select Case CInt(ang + pi / 8)
             case 0, 8:dump = dump & "->"
             case 1:dump = dump & "/'"
             case 2:dump = dump & "/\"
             case 3:dump = dump & "'\"
             case 4:dump = dump & "<-"
             case 5:dump = dump & ":/"
             case 6:dump = dump & "\/"
             case 7:dump = dump & "\:"
         End Select
 
         dump = dump & " mag:" & CLng(mag * 10) / 10 & ", ang:" & CLng(ang * 180 / pi) & ", x:" & CLng(x * 10) / 10 & ", y:" & CLng(y * 10) / 10
     End Property
End Class

Function ECVel(Velocity1, Mass1, Velocity2, Mass2)
	ECVel = (Mass1 - Mass2)/(Mass1 + Mass2) * Velocity1  + 2 * Mass2/(Mass1 + Mass2)*Velocity2 
End Function





' ********* VR Stuff - Rawd ***************


Dim TrainArrived: TrainArrived = false
Dim CarSpeed: CarSpeed = 24
Dim TrainSoundOn: TrainSoundOn = False
Dim TrainSoundLeaveOn: TrainSoundLeaveOn = False
Dim GateMove: GateMove = false
Dim GreenLightOn: GreenLightOn = false
Dim Cnt
Dim LightSpeed: LightSpeed = 1


Sub VRATimer_Timer()  'Controls timing of animations in the VRTimer Sub.
	Cnt = Cnt + 1
End Sub


Sub VRTimer_Timer()

	if TrainArrived = False then 
		if Cnt => 50 then 
			VRTrain.x = VRTrain.x + CarSpeed
			VRTrainSeats1.x = VRTrainSeats1.x + CarSpeed
			VRTrainSeats2.x = VRTrainSeats2.x + CarSpeed
			VRTrainSeats3.x = VRTrainSeats3.x + CarSpeed
			VRTrainSeats4.x = VRTrainSeats4.x + CarSpeed
			VRTrainSeats5.x = VRTrainSeats5.x + CarSpeed

			If VRTRain.x > -9000 then 
				If TrainSoundOn = False then 
					TrainSoundOn = True
					PlaySoundAtVol SoundFX("Train_Arrive_Fix", DOFContactors), DMD, CoasterVolume
				End If
			End if

			If VRTRain.x > -1420 then CarSpeed = 3.1	
			If VRTRain.x > -100 then 
				GateMove = true
				CarSpeed = 0	
			end if
		end if

		if Cnt => 255 then
			if VRTrainSeats1.rotz < 90 then 
				VRTrainSeats1.rotz = VRTrainSeats1.rotz + 0.7
				VRTrainSeats2.rotz = VRTrainSeats2.rotz + 0.7
				VRTrainSeats3.rotz = VRTrainSeats3.rotz + 0.7
				VRTrainSeats4.rotz = VRTrainSeats4.rotz + 0.7
				VRTrainSeats5.rotz = VRTrainSeats5.rotz + 0.7
			end If
		end If

		if Cnt => 293 then
			If VRGateMove5.roty < 90 then
				VRGateMove5.roty = VRGateMove5.roty + 0.5
				VRGateMove4.roty = VRGateMove4.roty + 0.5
				VRGateMove3.roty = VRGateMove3.roty + 0.5
				VRGateMove2.roty = VRGateMove2.roty + 0.5
				VRGateMove1.roty = VRGateMove1.roty + 0.5
			end if
		End if
	End If 'Train Arrived

	' Start Coaster leaving animation sequence
	if Cnt => 500 then
		TrainArrived = True  ' this is so we can set the seats and gates to go the other way.. above code wont run anymore.

		If TrainSoundLeaveOn = False then 
			TrainSoundLeaveOn = True
			PlaySoundAtVol SoundFX("Train_Leave_Fix", DOFContactors), DMD, CoasterVolume
		End If
	end if

	if Cnt => 512 then
		If VRGateMove5.roty => 0 then
			VRGateMove5.roty = VRGateMove5.roty - 0.5
			VRGateMove4.roty = VRGateMove4.roty - 0.5
			VRGateMove3.roty = VRGateMove3.roty - 0.5
			VRGateMove2.roty = VRGateMove2.roty - 0.5
			VRGateMove1.roty = VRGateMove1.roty - 0.5
		end if
	end if

	if Cnt => 529 then
		if VRTrainSeats1.rotz => 0 then 
			VRTrainSeats1.rotz = VRTrainSeats1.rotz - 0.8
			VRTrainSeats2.rotz = VRTrainSeats2.rotz - 0.8
			VRTrainSeats3.rotz = VRTrainSeats3.rotz - 0.8
			VRTrainSeats4.rotz = VRTrainSeats4.rotz - 0.8
			VRTrainSeats5.rotz = VRTrainSeats5.rotz - 0.8
		end If
	end If

	if Cnt = 555 then CarSpeed = 10'only run once

	if Cnt => 555 then
		VRTrain.x = VRTrain.x + CarSpeed
		VRTrainSeats1.x = VRTrainSeats1.x + CarSpeed
		VRTrainSeats2.x = VRTrainSeats2.x + CarSpeed
		VRTrainSeats3.x = VRTrainSeats3.x + CarSpeed
		VRTrainSeats4.x = VRTrainSeats4.x + CarSpeed
		VRTrainSeats5.x = VRTrainSeats5.x + CarSpeed
	end If

	if Cnt = 608 then CarSpeed = 28
	'if Cnt > 610 then CarSpeed = CarSpeed + 0.1 ' speed up train as it exits

	if VRTrain.x => 60000 Then  ' Reset All   (Increase this variable to increase time until next Train)
		cnt = 0
		TrainSoundOn = False
		TrainArrived = False
		TrainSoundLeaveOn = False
		CarSpeed = 24
		VRTrain.x = -33259.07
		VRTrainSeats1.x = -36339
		VRTrainSeats2.x = -38563.95
		VRTrainSeats3.x = -40797.7
		VRTrainSeats4.x = -43033.4
		VRTrainSeats5.x = -45261.55
	End If
End Sub


' VR PLUNGER ANIMATION AND VISUALIZATION
PlungerLine.blenddisablelighting = 3

Sub TimerPlunger_Timer
	If PinCab_Shooter.TransY < 90 then
		PinCab_Shooter.TransY = PinCab_Shooter.TransY + 5
	End If
	PlungerLine.TransY = (6.5* Plunger.Position) - 20
	PlungerLine.TransZ = -0.5*(Plunger.Position-4.1)
End Sub

Sub TimerPlunger2_Timer
	PinCab_Shooter.TransY = (5* Plunger.Position) -20
	PlungerLine.TransY = (6.5* Plunger.Position) - 20
	PlungerLine.TransZ = -0.5*(Plunger.Position-4.1)
End Sub


'************ End VR Stuff ************************************


' CHANGELOG

'01-02 - Gedankekojote97 - 	Table rebuild according to the manual. Added Fleep Sounds and nFozzy Physics. Added Lampz. Removed unneeded stuff.
'                          	Added ball back kick from upper right target when behind.
'03 - apophis - 	       	Updated physics scripts. Minor table level physics tweaks. Fixed scrambled egg. Added new PWM flasher support. 
'					       	Prep for toolkit: removed old dyn shadows (kept ambient shadows). Removed lampz & updated lamp fading
'04 - Gedankekojote97 -    	Fixed skillshot hole. Added new ramp collision models and flaps. Added new recorded sounds for some Sol.
'                          	Tweaked some Physics. Added System for the small gate infront the scrambled egg 
'05 - Gedankekojote97 -    	Reworked all ramps. Reworked scoop. Fixed (hopeful) the ball lock kicker. Implemented new scrambled egg (First try) 
'06 - apophis - 	       	Converted dot matrix objects to lights. Updated/fixed scrambled egg. Fixed OnBallBallCollision sounds. Fix sw43 kick speed. 
'							Increased pf fric. Updated ball rolling code. Updated ball images and env image. Tweaked lamp fading. Updated desktop POV. 
'07 - apophis - 	       	Tweaked scrabled egg ball size, sw43 strength, GateSET trigger and logic.
'08 - apophis - 	       	Changed sw36 from a wall to a gate called "GhostTarget". Fixed sw44.
'09 - Gedankekojote97 - 	Toolkit import: initial 2k batch. Updated playfield_mesh. Added collidable wire ramps.
'10 - apophis - 	       	Hid all light objects and collidable objects. Added and applied VLM materials. Fixed dup playfield_mesh issue. 
'							Added LightLevel option. Added ball brightness code. Added animations: flippers, bumpers, triggers
'11 - apophis - 	       	Fixed DMX lights order. Fix ghost Target gate. Fix scrambled egg switch pulsing
'12 - apophis - 	       	Updated drop target code. Added animations: gates, ghosttarget, scrambled egg, slings, dummy, diverters, drop targets, standup targets, rails
'13 - Gedankekojote97 - 	Toolkit import 4k batch: Fixed flipper orientations, added Dummy hair, removed diverters.
'14 - apophis - 	       	Toolkit script updates. Added Flex options menu. Updated room brightness code. Fixed RT shadow issues. Removed unused images.
'15 - Gedankekojote97 - 	Toolkit import 4k batch: Fixed DMD lights and name length, dummy hair. Added movable "post".
'16 - apophis - 	       	Toolkit script updates. Animated Post. Disabled reflections for all lightmaps. Fixed GhostTarget open/close angle. 
'17 - apophis - 	       	Added staged flipper support (thanks Primetime5k). Added Roth ST stuff. Fixed some ramptriggers. Fixed ball stuck locations.
'							Set Parts and PF bakemaps to static and make Light Level option require a restart.
'18 - apophis - 	       	Fixed scrambled egg switch timing. Tweaked flipper strength. Put UpdateLeds on 10ms timer.
'19 - Gedankekojote97 - 	Toolkit import 4k batch. Optimizations.
'20 - apophis - 	       	Toolkit script updates. Added flipper texture swapping script. Added refractions to plastic ramps. Fix for stuck ball at Gate1. 
'                           Added plinger line visualization option.
'21 - apophis - 	       	Added desktop backdrop (made by Gedank). Put plunger visualization in Flex options menu. Fixed troll's hair. Fixed Gate002 animation. 
'22 - apophis - 	       	Added Narnia catcher. Fixed post passes.
'23 - Sixtoe  -				sw47 made non-visible, altered, combined, expanded and tidied up collidable walls.
'24 - Gedankekojote97 - 	Toolkit import 4k batch. Updated inserts, spotlights, flipper up positions, separated parts above ramp.
'25 - apophis - 	       	Updated ball rolling sounds (louder). Made ball darker in plunger lanes. Ramp refraction thickness 7, roughness 1. 
'							Updated flipper animations. Parts2 active material. Nestmap0 alpha mask 75.
'26 - Gedankekojote97 - 	Toolkit import with some fixes on left ramp and LED board.
'27 - apophis - 	       	Added option to turn off PWM flashers. Fixed flipper shadow DB. Re-fixed: ramp refractions, Parts2 active material, 
'							Parts and playfield meshes set to static. Nestmap0 alpha mask 75.
'28 - Gedankekojote97 - 	Soundvolume tweaks. Flippers, Plunger and Autoplunger tweaks. Collision wall on left metal ramp. Removed right scoop from metal sound collection
'29 - apophis - 	       	Fixed script for standalone.
'30 - Rawd - 				Added Gedank's VR room, animated and coded.
'31 - Gedankekojote97 - 	Toolkit import with some fixes
'32 - apophis - 	       	Updated VLM arrays. Added SetLocale. Added pf image for ball reflections. Added FlipperCradleCollision flipper trick. Put VRRoomChoice in option menu.
'                           Re-fix: playfield_vis static and reflection probe applied. Layer1 refraction applied. Parts set to static. Parts2 active material applied. 
'							Nestmap0 alpha mask 75 (troll hair issue).
'33 - apophis - 	       	Keydown/up VR animations fixed. Turn off VR green light in Minimal Room. Deleted Warmup prims. Added VR Room to light level option. 
'							Rawd updates: Updated VR wall texture fix (webp). VR object DL tweaks. 
'34 - Rawd - 				Unhooked VR Exit sign, and cab lights from the VRlighting options.  Green light bandaid fix on VR lighting changes.
'35 - apophis - 	       	Hooked VR flipper buttons to room lighting. Fixed dummy LM zfighting issue (L131 and Spotlights). Fixed center ramp sfx. 
'36 - Gedankekojote97 -     New Nestmaps with fixed left scoop and rocket scoop materials. New desktop backdrop.
'RC1 - apophis - 	       	Initial flasher flash. Added VPM desktop DMD visibility  option. 
'RC2 - apophis - 	       	Fixed trough switch 15. Updated default ball image and ball brightness equations. 
'RC3 - Rawd - 	       		Added "Metal Dark" to room brightness. Fixes to cabinet bottom in VR. 
'RC4 - apophis - 	       	Adjusted right autoplunger vol. White insert reflection colors updated. Added some walls to Metals collection. Cab pov updated.
'v1.0 Release
'1.0.1 - apophis 			Updates for UseVPMModSol=2. Render probe roughnesses set to 0. Added DisableStaticPreRendering functionality to options menu.
'1.0.2 - apophis 			Fixed timers causing stutters.
'2.0.0 - Gedankekojote97 -  Changed version numbering. Completely new recorded sounds for all mechanics on the table from my real machine
'2.0.1 - Gedankekojote97 -  Fixed sw38 behavior and timer. Changed ROM from 7.2 -> 7.1 (since 7.2 is broken). 
'							Upper Flippers code moved to lower flippers (not controled by their own coil)
'2.0.2 - Gedankekojote97 -  Adjusted mechanical sound Volume
'2.0.3 - Gedankekojote97 -  Adjusted Layout in VPX to match blend file. Removed drop target fleep routine. 
'							Made droptargets collidable from behind 
'2.0.4 - Gedankekojote97 -  Code Cleanup. New rendered backglass for VR Room
'2.0.5 - Gedankekojote97 -  New blender table bake import. New topper VR Room import 
'2.0.6 - apophis - 			PWM not optional anymore. Updated all VLM array names. Converted options menu to Tweak. Fixed plunger visualization option. Updated ambient ball shadow code.
'2.0.7 - Gedankekojote97 -  fixes to the layout where the ball got stuck, made "static" for the movables unchecked, changed dummy to vlm.active material (because of his hair), new Nestmaps
'2.0.8 - apophis - 			Wired up topper to VR room, and some lights controlled by inserts. Fixed dummy DB. Fixed some ball rt shadows. All light faders set to None. Fixed plunger visualization. Added RCT mech sound effects volume option. Fixed flipper length, updated triggers, strength and scripts. Tuned flasher brightness.
'2.0 RC1 - apophis - 		Added Table Info, including rules card. Fixed troll's hair. Fixed too long LM names. Enabled PF reflections. Added ramp refraction options. Hide parts behind for Layer1 and Layer2. Fixed pincab rail visibility in VR.
'2.0.RC2 - Gedankekojote97- Fixed troll hair Nestmap, Fixed troll movement
'2.0 RC3 - apophis - 		Autoplunger power to 39. Adjusted gate physics 
'2.0.RC4 - Gedankekojote97- Reverted upper flippers callbacks back to its own solenoids. Adjusted Drop target triggers and collision. Droptarget drop heigh lowered.
'2.0.RC5 - apophis -  		Added bumper strength variation hack. Updated DisableStaticPreRendering functionality to be compatible with VPX 10.8.1 API.
'Release 2.0
'2.0.1 - apophis - 			Fixed LM flicker on Scrambled Egg. 
'2.0.2 - apophis - 			Backglass buzz sound effect default to OFF. Fixed orientation of sw29. Fixed transparency of the lower left flasher platform. 
'Release 2.1
'2.1.1 - Gedankekojote97-   Fixed ball stuck on skillshot near diverter
'Release 2.2
'2.2.1 - Gedankekojote97-   Removed LUT Settings. Added Reflection Settings. Added toggleable Glass. Set BG to always on. Changed position of GI Relay Sound.
'2.2.2 - Gedankekojote97-   Added new Environment image. Added new Ball image and scratches
'2.2.3 - Gedankekojote97-   Added new rendered Ball image (Made with its VR Room), - Added new Ball Scratches, - Added new Environment (VR Room)
'2.2.4 - apophis -   		Fixed plunger visualization line.
'2.2.5 - Gedankekojote97-   Cleaned up Sound files that are not needed. Added VR Room choice (Full, Mini, Ultra MiniRoom). Added Option to enable/disable Siderails.
'							Added new Full VR Room, bakes/normals. Added Minimal VR Room, bakes. Changed some VR Code. Added Option to change volume of the coaster.
'2.2.6 - Rawd- 				Train seat and Gate animation fix in VRroom.  Reset Coaster position and code.
'2.2.7 - Gedankekojote97-   New LOD Meshes and renders for ULTRA Vr room. Added Black Wall at the back of Minimal Vr room. Added new ball image. Some small fixes
'2.2.8 - Rawd -             Re-aligned new room (that was rotated 3 degrees and scaled) - Reset coaster seats and gate positions
'2.2.8b- Rawd - 			Separated VR rails from posters and other objects and raised the height of the rails and gates in VR
'2.2.9 - apophis - 			Fixed(?) plunger line visualization. Fixed (?) ghost target animation. Added a ball image option.

'Release 2.3

