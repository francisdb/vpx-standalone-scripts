'			VPW are proud to present our 50th Machine!
'   _____                   _   _ _   _____                     
'  / ____|                 | \ | ( ) |  __ \                    
' | |  __ _   _ _ __  ___  |  \| |/  | |__) |___  ___  ___  ___ 
' | | |_ | | | | '_ \/ __| | . ` |   |  _  // _ \/ __|/ _ \/ __|
' | |__| | |_| | | | \__ \ | |\  |   | | \ \ (_) \__ \  __/\__ \
'  \_____|\__,_|_| |_|___/ |_| \_|   |_|  \_\___/|___/\___||___/
'
' Guns N' Roses (Data East 1994)
' https://www.ipdb.org/machine.cgi?id=1100
'
' VPW Roadies
' ===========
' Niwak - Rendering Magic
' Sixtoe - Table Wrangling
' HiRez00 - All Artwork Scans & Redraws
' iaakki - Tweaks and support
' leojreimroc - Animated VR Backglass.
' Tastywasps - VR Room & Guitar.
' Primetime5k - Staged Flipper Support.
' Hauntfreaks - Backglass images.
' Apophis - Tweaks and support.
' Flupper - Rendering Support and Advice
' Testing - VPW Team
'
' All options are available from in game menu accessed by holding both left then right magnasave.
' ** The option menu needs FlexDMD to be installed **
'
' Version Log at end of script.

Option Explicit
Randomize
SetLocale 1033

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'******************************************************
'  Constants and Global Variables
'******************************************************
Const TableVersion = "1.0"	'Table version (also shown in option UI)

Const BallSize = 50			'Ball size must be 50
Const BallMass = 1			'Ball mass must be 1

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height
Const UseVPMModSol = True

'We guess that upper flipper is not supported by "fast flips".
' So this option skips Solenoid 47 and controls upper flipper from left flipper solenoid.
Const ForceFastFlips = true

'*************************************************************
' VR Room Auto-Detect
'*************************************************************
Dim VR_Obj, UseVPMDMD

Dim DesktopMode: DesktopMode = Table1.ShowDT

If RenderingMode = 2 Then
	'Pincab_Bottom.Visible = 1
	For Each VR_Obj in VRCabinet : VR_Obj.Visible = 1 : Next
	For Each VR_Obj in VRRoom : VR_Obj.Visible = 1 : Next
	For each VR_Obj in VRBackglass : VR_Obj.visible = 1 : Next
	UseVPMDMD = True
	Setbackglass
Else
	'Pincab_Bottom.Visible = 0
	For Each VR_Obj in VRCabinet : VR_Obj.Visible = 0 : Next
	For Each VR_Obj in VRRoom : VR_Obj.Visible = 0 : Next
	For each VR_Obj in VRBackglass : VR_Obj.visible = 0 : Next
	UseVPMDMD = DesktopMode
	Pincab_Rails.visible = DesktopMode
End If

LoadVPM "01120100", "de.vbs", 3.02

'******************************************************
' Standard definitions
'******************************************************
Const cGameName = "gnr_300"
Const UseSolenoids = 2
Const UseLamps = 1
Const UseSync = 0
Const HandleMech = 0
Const SSolenoidOn = ""
Const SSolenoidOff = ""
Const SFlipperOn = ""
Const SFlipperOff = ""
Const SCoin = ""

'******************************************************
' VPX lightmapper definitions
'******************************************************

' VLM Arrays - Start
' Arrays per baked part
Dim BP_BP_Empty: BP_BP_Empty=Array(BM_BP_Empty, LM_GI_l141_BP_Empty, LM_GI_l141d_BP_Empty, LM_GI_l141e_BP_Empty, LM_GI_l141f_BP_Empty, LM_GI_l141g_BP_Empty, LM_GI_l141h_BP_Empty, LM_GI_l141i_BP_Empty, LM_GI_l141l_BP_Empty, LM_GI_l141k_BP_Empty, LM_Flashers_f1_BP_Empty, LM_Flashers_f4_BP_Empty, LM_Inserts_l57_BP_Empty, LM_Inserts_l58_BP_Empty, LM_Inserts_l59_BP_Empty, LM_Inserts_l60_BP_Empty)
Dim BP_BP_Jungle: BP_BP_Jungle=Array(BM_BP_Jungle, LM_GI_l141_BP_Jungle, LM_GI_l141d_BP_Jungle, LM_GI_l141e_BP_Jungle, LM_GI_l141f_BP_Jungle, LM_GI_l141g_BP_Jungle, LM_GI_l141h_BP_Jungle, LM_GI_l141i_BP_Jungle, LM_GI_l141l_BP_Jungle, LM_GI_l141k_BP_Jungle, LM_Flashers_f1_BP_Jungle, LM_Flashers_f4_BP_Jungle, LM_Inserts_l29_BP_Jungle, LM_Inserts_l58_BP_Jungle)
Dim BP_BR1: BP_BR1=Array(BM_BR1, LM_GI_l141_BR1, LM_Flashers_f1_BR1, LM_Flashers_f4_BR1, LM_Flashers_f8_BR1)
Dim BP_BR2: BP_BR2=Array(BM_BR2, LM_GI_l141_BR2, LM_GI_l141d_BR2, LM_GI_l141e_BR2, LM_GI_l141f_BR2, LM_Flashers_f1_BR2)
Dim BP_BR3: BP_BR3=Array(BM_BR3, LM_GI_l141_BR3, LM_GI_l141e_BR3, LM_GI_l141f_BR3, LM_Flashers_f1_BR3, LM_Flashers_f4_BR3, LM_Flashers_f8_BR3)
Dim BP_BS1: BP_BS1=Array(BM_BS1, LM_GI_l141_BS1, LM_GI_l141d_BS1, LM_Flashers_f1_BS1, LM_Flashers_f4_BS1, LM_Flashers_f7_BS1, LM_Flashers_f8_BS1)
Dim BP_BS2: BP_BS2=Array(BM_BS2, LM_GI_l141_BS2, LM_GI_l141d_BS2, LM_GI_l141e_BS2, LM_GI_l141f_BS2, LM_GI_l141l_BS2, LM_Flashers_f1_BS2, LM_Flashers_f4_BS2, LM_Flashers_f7_BS2)
Dim BP_BS3: BP_BS3=Array(BM_BS3, LM_GI_l141_BS3, LM_GI_l141d_BS3, LM_GI_l141e_BS3, LM_GI_l141f_BS3, LM_GI_l141g_BS3, LM_GI_l141l_BS3, LM_Flashers_f1_BS3, LM_Flashers_f4_BS3, LM_Flashers_f6_BS3, LM_Flashers_f7_BS3, LM_Flashers_f8_BS3)
Dim BP_B_Caps_Bot: BP_B_Caps_Bot=Array(BM_B_Caps_Bot, LM_GI_l141_B_Caps_Bot, LM_GI_l141d_B_Caps_Bot, LM_GI_l141e_B_Caps_Bot, LM_GI_l141f_B_Caps_Bot, LM_GI_l141g_B_Caps_Bot, LM_Flashers_f1_B_Caps_Bot, LM_Flashers_f4_B_Caps_Bot, LM_Flashers_f8_B_Caps_Bot, LM_Inserts_l31_B_Caps_Bot, LM_Inserts_l33_B_Caps_Bot, LM_Inserts_l48_B_Caps_Bot)
Dim BP_B_Caps_Top: BP_B_Caps_Top=Array(BM_B_Caps_Top, LM_GI_l141_B_Caps_Top, LM_GI_l141d_B_Caps_Top, LM_GI_l141e_B_Caps_Top, LM_GI_l141f_B_Caps_Top, LM_GI_l141g_B_Caps_Top, LM_GI_l141i_B_Caps_Top, LM_GI_l141l_B_Caps_Top, LM_GI_l141k_B_Caps_Top, LM_Flashers_f1_B_Caps_Top, LM_Flashers_f4_B_Caps_Top, LM_Flashers_f8_B_Caps_Top, LM_Inserts_l31_B_Caps_Top, LM_Inserts_l48_B_Caps_Top, LM_Inserts_l58_B_Caps_Top, LM_Inserts_l59_B_Caps_Top, LM_Inserts_l60_B_Caps_Top)
Dim BP_Blades: BP_Blades=Array(BM_Blades, LM_GI0_l140a_Blades, LM_GI0_l140b_Blades, LM_GI0_l140c_Blades, LM_GI0_l140d_Blades, LM_GI0_l140e_Blades, LM_GI0_l140f_Blades, LM_GI_l141a_Blades, LM_GI_l141_Blades, LM_GI_l141b_Blades, LM_GI_l141c_Blades, LM_GI_l141h_Blades, LM_GI_l141i_Blades, LM_GI_l141j_Blades, LM_GI_l141k_Blades, LM_Flashers_f1_Blades, LM_Flashers_f2_Blades, LM_Flashers_f3_Blades, LM_Flashers_f4_Blades, LM_Flashers_f5_Blades, LM_Flashers_f6_Blades, LM_Flashers_f8_Blades, LM_Inserts_l55_Blades, LM_Inserts_l62_Blades)
Dim BP_LEMK: BP_LEMK=Array(BM_LEMK, LM_GI0_l140b_LEMK, LM_GI0_l140c_LEMK, LM_GI0_l140d_LEMK, LM_Flashers_f3_LEMK)
Dim BP_LEMK1: BP_LEMK1=Array(BM_LEMK1, LM_GI_l141_LEMK1, LM_Flashers_f1_LEMK1, LM_Flashers_f4_LEMK1, LM_Flashers_f8_LEMK1)
Dim BP_LFlipper: BP_LFlipper=Array(BM_LFlipper, LM_GI0_l140a_LFlipper, LM_GI0_l140b_LFlipper, LM_GI0_l140c_LFlipper, LM_GI0_l140d_LFlipper, LM_GI0_l140f_LFlipper, LM_GI_l141_LFlipper, LM_Flashers_f2_LFlipper, LM_Flashers_f3_LFlipper, LM_Inserts_l46_LFlipper)
Dim BP_LFlipper1: BP_LFlipper1=Array(BM_LFlipper1, LM_GI_l141_LFlipper1, LM_GI_l141j_LFlipper1, LM_Flashers_f1_LFlipper1, LM_Flashers_f2_LFlipper1, LM_Flashers_f3_LFlipper1, LM_Flashers_f6_LFlipper1, LM_Inserts_l11_LFlipper1, LM_Inserts_l43_LFlipper1, LM_Inserts_l44_LFlipper1, LM_Inserts_l50_LFlipper1, LM_Inserts_l56_LFlipper1)
Dim BP_LFlipper1U: BP_LFlipper1U=Array(BM_LFlipper1U, LM_GI_l141a_LFlipper1U, LM_GI_l141_LFlipper1U, LM_GI_l141i_LFlipper1U, LM_GI_l141j_LFlipper1U, LM_Flashers_f1_LFlipper1U, LM_Flashers_f2_LFlipper1U, LM_Flashers_f3_LFlipper1U, LM_Flashers_f6_LFlipper1U, LM_Flashers_f8_LFlipper1U, LM_Inserts_l11_LFlipper1U, LM_Inserts_l27_LFlipper1U, LM_Inserts_l43_LFlipper1U, LM_Inserts_l44_LFlipper1U, LM_Inserts_l45_LFlipper1U, LM_Inserts_l50_LFlipper1U, LM_Inserts_l51_LFlipper1U, LM_Inserts_l52_LFlipper1U, LM_Inserts_l56_LFlipper1U)
Dim BP_LFlipperU: BP_LFlipperU=Array(BM_LFlipperU, LM_GI0_l140a_LFlipperU, LM_GI0_l140b_LFlipperU, LM_GI0_l140c_LFlipperU, LM_GI0_l140d_LFlipperU, LM_GI0_l140f_LFlipperU, LM_GI_l141_LFlipperU, LM_Flashers_f2_LFlipperU, LM_Flashers_f3_LFlipperU, LM_Inserts_l37_LFlipperU, LM_Inserts_l46_LFlipperU)
Dim BP_LSling1: BP_LSling1=Array(BM_LSling1, LM_GI0_l140c_LSling1, LM_GI0_l140d_LSling1, LM_Flashers_f2_LSling1, LM_Flashers_f3_LSling1)
Dim BP_LSling2: BP_LSling2=Array(BM_LSling2, LM_GI0_l140c_LSling2, LM_GI0_l140d_LSling2, LM_Flashers_f2_LSling2)
Dim BP_Layer_1: BP_Layer_1=Array(BM_Layer_1, LM_GI0_l140a_Layer_1, LM_GI0_l140b_Layer_1, LM_GI0_l140c_Layer_1, LM_GI0_l140d_Layer_1, LM_GI0_l140e_Layer_1, LM_GI0_l140f_Layer_1, LM_GI_l141_Layer_1, LM_GI_l141d_Layer_1, LM_GI_l141e_Layer_1, LM_GI_l141f_Layer_1, LM_GI_l141g_Layer_1, LM_GI_l141i_Layer_1, LM_GI_l141l_Layer_1, LM_GI_l141k_Layer_1, LM_Flashers_f1_Layer_1, LM_Flashers_f2_Layer_1, LM_Flashers_f3_Layer_1, LM_Flashers_f4_Layer_1, LM_Flashers_f7_Layer_1, LM_Flashers_f8_Layer_1, LM_Inserts_l17_Layer_1, LM_Inserts_l18_Layer_1, LM_Inserts_l25_Layer_1, LM_Inserts_l26_Layer_1, LM_Inserts_l29_Layer_1, LM_Inserts_l30_Layer_1, LM_Inserts_l52_Layer_1, LM_Inserts_l53_Layer_1, LM_Inserts_l56_Layer_1, LM_Inserts_l57_Layer_1, LM_Inserts_l61_Layer_1)
Dim BP_Layer_2: BP_Layer_2=Array(BM_Layer_2, LM_GI0_l140a_Layer_2, LM_GI0_l140b_Layer_2, LM_GI0_l140c_Layer_2, LM_GI0_l140d_Layer_2, LM_GI0_l140e_Layer_2, LM_GI0_l140f_Layer_2, LM_GI_l141_Layer_2, LM_GI_l141f_Layer_2, LM_GI_l141k_Layer_2, LM_Flashers_f1_Layer_2, LM_Flashers_f2_Layer_2, LM_Flashers_f3_Layer_2, LM_Flashers_f4_Layer_2)
Dim BP_Layer_3: BP_Layer_3=Array(BM_Layer_3, LM_GI_l141_Layer_3, LM_GI_l141i_Layer_3, LM_GI_l141k_Layer_3, LM_Flashers_f1_Layer_3)
Dim BP_MSling1: BP_MSling1=Array(BM_MSling1, LM_GI_l141_MSling1, LM_GI_l141g_MSling1, LM_GI_l141k_MSling1, LM_Flashers_f1_MSling1, LM_Flashers_f3_MSling1, LM_Flashers_f4_MSling1, LM_Flashers_f8_MSling1, LM_Inserts_l29_MSling1, LM_Inserts_l30_MSling1)
Dim BP_MSling2: BP_MSling2=Array(BM_MSling2, LM_GI_l141_MSling2, LM_GI_l141g_MSling2, LM_Flashers_f1_MSling2, LM_Flashers_f4_MSling2, LM_Inserts_l29_MSling2, LM_Inserts_l30_MSling2)
Dim BP_PF: BP_PF=Array(BM_PF, LM_GI0_l140a_PF, LM_GI0_l140b_PF, LM_GI0_l140c_PF, LM_GI0_l140d_PF, LM_GI0_l140e_PF, LM_GI0_l140f_PF, LM_GI_l141a_PF, LM_GI_l141_PF, LM_GI_l141b_PF, LM_GI_l141c_PF, LM_GI_l141d_PF, LM_GI_l141e_PF, LM_GI_l141f_PF, LM_GI_l141g_PF, LM_GI_l141h_PF, LM_GI_l141i_PF, LM_GI_l141j_PF, LM_GI_l141l_PF, LM_GI_l141k_PF, LM_Flashers_f1_PF, LM_Flashers_f2_PF, LM_Flashers_f3_PF, LM_Flashers_f4_PF, LM_Flashers_f5_PF, LM_Flashers_f6_PF, LM_Flashers_f7_PF, LM_Flashers_f8_PF, LM_Inserts_l10_PF, LM_Inserts_l11_PF, LM_Inserts_l12_PF, LM_Inserts_l13_PF, LM_Inserts_l14_PF, LM_Inserts_l15_PF, LM_Inserts_l16_PF, LM_Inserts_l17_PF, LM_Inserts_l18_PF, LM_Inserts_l19_PF, LM_Inserts_l1_PF, LM_Inserts_l20_PF, LM_Inserts_l21_PF, LM_Inserts_l22_PF, LM_Inserts_l23_PF, LM_Inserts_l24_PF, LM_Inserts_l25_PF, LM_Inserts_l26_PF, LM_Inserts_l27_PF, LM_Inserts_l28_PF, LM_Inserts_l29_PF, LM_Inserts_l2_PF, LM_Inserts_l30_PF, LM_Inserts_l31_PF, LM_Inserts_l33_PF, LM_Inserts_l34_PF, LM_Inserts_l35_PF, LM_Inserts_l36_PF, _
	LM_Inserts_l37_PF, LM_Inserts_l39_PF, LM_Inserts_l3_PF, LM_Inserts_l40_PF, LM_Inserts_l41_PF, LM_Inserts_l42_PF, LM_Inserts_l43_PF, LM_Inserts_l44_PF, LM_Inserts_l45_PF, LM_Inserts_l46_PF, LM_Inserts_l47_PF, LM_Inserts_l48_PF, LM_Inserts_l49_PF, LM_Inserts_l4_PF, LM_Inserts_l50_PF, LM_Inserts_l51_PF, LM_Inserts_l52_PF, LM_Inserts_l53_PF, LM_Inserts_l54_PF, LM_Inserts_l55_PF, LM_Inserts_l56_PF, LM_Inserts_l57_PF, LM_Inserts_l58_PF, LM_Inserts_l59_PF, LM_Inserts_l5_PF, LM_Inserts_l60_PF, LM_Inserts_l61_PF, LM_Inserts_l62_PF, LM_Inserts_l6_PF, LM_Inserts_l7_PF, LM_Inserts_l8_PF, LM_Inserts_l9_PF)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_GI0_l140a_Parts, LM_GI0_l140b_Parts, LM_GI0_l140c_Parts, LM_GI0_l140d_Parts, LM_GI0_l140e_Parts, LM_GI0_l140f_Parts, LM_GI_l141a_Parts, LM_GI_l141_Parts, LM_GI_l141b_Parts, LM_GI_l141c_Parts, LM_GI_l141d_Parts, LM_GI_l141e_Parts, LM_GI_l141f_Parts, LM_GI_l141g_Parts, LM_GI_l141h_Parts, LM_GI_l141i_Parts, LM_GI_l141j_Parts, LM_GI_l141l_Parts, LM_GI_l141k_Parts, LM_Flashers_f1_Parts, LM_Flashers_f2_Parts, LM_Flashers_f3_Parts, LM_Flashers_f4_Parts, LM_Flashers_f5_Parts, LM_Flashers_f6_Parts, LM_Flashers_f7_Parts, LM_Flashers_f8_Parts, LM_Inserts_l10_Parts, LM_Inserts_l11_Parts, LM_Inserts_l12_Parts, LM_Inserts_l13_Parts, LM_Inserts_l14_Parts, LM_Inserts_l15_Parts, LM_Inserts_l16_Parts, LM_Inserts_l17_Parts, LM_Inserts_l18_Parts, LM_Inserts_l19_Parts, LM_Inserts_l1_Parts, LM_Inserts_l20_Parts, LM_Inserts_l21_Parts, LM_Inserts_l22_Parts, LM_Inserts_l23_Parts, LM_Inserts_l24_Parts, LM_Inserts_l25_Parts, LM_Inserts_l26_Parts, LM_Inserts_l27_Parts, LM_Inserts_l28_Parts, _
	LM_Inserts_l29_Parts, LM_Inserts_l2_Parts, LM_Inserts_l30_Parts, LM_Inserts_l31_Parts, LM_Inserts_l32_Parts, LM_Inserts_l33_Parts, LM_Inserts_l34_Parts, LM_Inserts_l35_Parts, LM_Inserts_l36_Parts, LM_Inserts_l37_Parts, LM_Inserts_l38_Parts, LM_Inserts_l39_Parts, LM_Inserts_l3_Parts, LM_Inserts_l40_Parts, LM_Inserts_l41_Parts, LM_Inserts_l42_Parts, LM_Inserts_l43_Parts, LM_Inserts_l44_Parts, LM_Inserts_l45_Parts, LM_Inserts_l46_Parts, LM_Inserts_l47_Parts, LM_Inserts_l48_Parts, LM_Inserts_l49_Parts, LM_Inserts_l4_Parts, LM_Inserts_l50_Parts, LM_Inserts_l51_Parts, LM_Inserts_l52_Parts, LM_Inserts_l53_Parts, LM_Inserts_l54_Parts, LM_Inserts_l55_Parts, LM_Inserts_l56_Parts, LM_Inserts_l57_Parts, LM_Inserts_l58_Parts, LM_Inserts_l59_Parts, LM_Inserts_l5_Parts, LM_Inserts_l60_Parts, LM_Inserts_l61_Parts, LM_Inserts_l62_Parts, LM_Inserts_l6_Parts, LM_Inserts_l7_Parts, LM_Inserts_l8_Parts, LM_Inserts_l9_Parts)
Dim BP_Post_L: BP_Post_L=Array(BM_Post_L, LM_GI0_l140b_Post_L, LM_GI0_l140c_Post_L, LM_GI0_l140d_Post_L, LM_Flashers_f2_Post_L, LM_Flashers_f3_Post_L, LM_Inserts_l17_Post_L, LM_Inserts_l18_Post_L)
Dim BP_Post_R: BP_Post_R=Array(BM_Post_R, LM_GI0_l140e_Post_R, LM_GI_l141_Post_R, LM_GI_l141b_Post_R, LM_Flashers_f2_Post_R, LM_Inserts_l26_Post_R)
Dim BP_REMK: BP_REMK=Array(BM_REMK, LM_GI0_l140a_REMK, LM_GI0_l140e_REMK, LM_GI0_l140f_REMK, LM_Flashers_f2_REMK)
Dim BP_RFlipper: BP_RFlipper=Array(BM_RFlipper, LM_GI0_l140a_RFlipper, LM_GI0_l140b_RFlipper, LM_GI0_l140c_RFlipper, LM_GI0_l140e_RFlipper, LM_GI0_l140f_RFlipper, LM_Flashers_f2_RFlipper, LM_Flashers_f3_RFlipper, LM_Inserts_l46_RFlipper)
Dim BP_RFlipperU: BP_RFlipperU=Array(BM_RFlipperU, LM_GI0_l140a_RFlipperU, LM_GI0_l140b_RFlipperU, LM_GI0_l140c_RFlipperU, LM_GI0_l140e_RFlipperU, LM_GI0_l140f_RFlipperU, LM_GI_l141b_RFlipperU, LM_Flashers_f2_RFlipperU, LM_Flashers_f3_RFlipperU, LM_Inserts_l37_RFlipperU, LM_Inserts_l46_RFlipperU)
Dim BP_RSling1: BP_RSling1=Array(BM_RSling1, LM_GI0_l140e_RSling1, LM_GI0_l140f_RSling1, LM_Flashers_f3_RSling1)
Dim BP_RSling2: BP_RSling2=Array(BM_RSling2, LM_GI0_l140e_RSling2, LM_GI0_l140f_RSling2, LM_Flashers_f3_RSling2)
Dim BP_R_Green: BP_R_Green=Array(BM_R_Green, LM_GI_l141a_R_Green, LM_GI_l141_R_Green, LM_GI_l141g_R_Green, LM_GI_l141i_R_Green, LM_GI_l141j_R_Green, LM_GI_l141k_R_Green, LM_Flashers_f1_R_Green, LM_Flashers_f2_R_Green, LM_Flashers_f3_R_Green, LM_Flashers_f4_R_Green, LM_Flashers_f6_R_Green, LM_Flashers_f8_R_Green, LM_Inserts_l11_R_Green, LM_Inserts_l31_R_Green, LM_Inserts_l33_R_Green, LM_Inserts_l50_R_Green, LM_Inserts_l51_R_Green, LM_Inserts_l52_R_Green, LM_Inserts_l53_R_Green, LM_Inserts_l56_R_Green)
Dim BP_R_Red: BP_R_Red=Array(BM_R_Red, LM_GI_l141_R_Red, LM_GI_l141c_R_Red, LM_GI_l141e_R_Red, LM_GI_l141h_R_Red, LM_Flashers_f1_R_Red, LM_Flashers_f2_R_Red, LM_Flashers_f5_R_Red, LM_Flashers_f7_R_Red, LM_Flashers_f8_R_Red, LM_Inserts_l48_R_Red, LM_Inserts_l60_R_Red)
Dim BP_R_Yellow: BP_R_Yellow=Array(BM_R_Yellow, LM_GI_l141_R_Yellow, LM_GI_l141e_R_Yellow, LM_GI_l141f_R_Yellow, LM_GI_l141g_R_Yellow, LM_GI_l141i_R_Yellow, LM_GI_l141j_R_Yellow, LM_GI_l141l_R_Yellow, LM_GI_l141k_R_Yellow, LM_Flashers_f1_R_Yellow, LM_Flashers_f3_R_Yellow, LM_Flashers_f4_R_Yellow, LM_Flashers_f5_R_Yellow, LM_Flashers_f6_R_Yellow, LM_Flashers_f7_R_Yellow, LM_Flashers_f8_R_Yellow, LM_Inserts_l27_R_Yellow, LM_Inserts_l29_R_Yellow, LM_Inserts_l30_R_Yellow, LM_Inserts_l31_R_Yellow, LM_Inserts_l32_R_Yellow, LM_Inserts_l33_R_Yellow, LM_Inserts_l38_R_Yellow, LM_Inserts_l39_R_Yellow, LM_Inserts_l3_R_Yellow, LM_Inserts_l40_R_Yellow, LM_Inserts_l41_R_Yellow, LM_Inserts_l43_R_Yellow, LM_Inserts_l44_R_Yellow, LM_Inserts_l45_R_Yellow, LM_Inserts_l47_R_Yellow, LM_Inserts_l49_R_Yellow, LM_Inserts_l50_R_Yellow, LM_Inserts_l51_R_Yellow, LM_Inserts_l52_R_Yellow, LM_Inserts_l53_R_Yellow, LM_Inserts_l55_R_Yellow, LM_Inserts_l56_R_Yellow)
Dim BP_Trap: BP_Trap=Array(BM_Trap)
Dim BP_gate1: BP_gate1=Array(BM_gate1, LM_GI_l141_gate1, LM_GI_l141d_gate1, LM_GI_l141e_gate1, LM_GI_l141f_gate1, LM_GI_l141l_gate1, LM_Flashers_f1_gate1, LM_Inserts_l60_gate1)
Dim BP_gate2: BP_gate2=Array(BM_gate2, LM_GI_l141_gate2, LM_GI_l141d_gate2, LM_GI_l141e_gate2, LM_GI_l141h_gate2, LM_GI_l141l_gate2, LM_Flashers_f1_gate2, LM_Inserts_l59_gate2, LM_Inserts_l60_gate2)
Dim BP_sw16: BP_sw16=Array(BM_sw16)
Dim BP_sw17: BP_sw17=Array(BM_sw17, LM_GI_l141_sw17, LM_GI_l141g_sw17, LM_GI_l141k_sw17, LM_Flashers_f1_sw17, LM_Flashers_f4_sw17, LM_Flashers_f8_sw17, LM_Inserts_l29_sw17, LM_Inserts_l30_sw17, LM_Inserts_l31_sw17)
Dim BP_sw18: BP_sw18=Array(BM_sw18, LM_GI_l141_sw18, LM_GI_l141f_sw18, LM_GI_l141g_sw18, LM_GI_l141k_sw18, LM_Flashers_f1_sw18, LM_Flashers_f4_sw18, LM_Flashers_f8_sw18, LM_Inserts_l29_sw18, LM_Inserts_l30_sw18, LM_Inserts_l31_sw18, LM_Inserts_l33_sw18)
Dim BP_sw19: BP_sw19=Array(BM_sw19, LM_GI_l141_sw19, LM_GI_l141f_sw19, LM_GI_l141g_sw19, LM_GI_l141k_sw19, LM_Flashers_f1_sw19, LM_Flashers_f4_sw19, LM_Flashers_f8_sw19, LM_Inserts_l29_sw19, LM_Inserts_l30_sw19, LM_Inserts_l31_sw19, LM_Inserts_l33_sw19)
Dim BP_sw20: BP_sw20=Array(BM_sw20, LM_GI_l141_sw20, LM_GI_l141g_sw20, LM_GI_l141k_sw20, LM_Flashers_f1_sw20, LM_Flashers_f4_sw20, LM_Flashers_f8_sw20, LM_Inserts_l29_sw20, LM_Inserts_l30_sw20, LM_Inserts_l31_sw20, LM_Inserts_l32_sw20, LM_Inserts_l33_sw20)
Dim BP_sw21: BP_sw21=Array(BM_sw21, LM_GI_l141_sw21, LM_GI_l141e_sw21, LM_GI_l141f_sw21, LM_GI_l141g_sw21, LM_Flashers_f1_sw21)
Dim BP_sw22: BP_sw22=Array(BM_sw22, LM_GI_l141d_sw22, LM_GI_l141e_sw22, LM_GI_l141f_sw22, LM_GI_l141g_sw22, LM_Flashers_f1_sw22)
Dim BP_sw23: BP_sw23=Array(BM_sw23, LM_GI_l141_sw23, LM_GI_l141d_sw23, LM_GI_l141e_sw23, LM_GI_l141f_sw23, LM_Flashers_f1_sw23)
Dim BP_sw24: BP_sw24=Array(BM_sw24, LM_Flashers_f3_sw24)
Dim BP_sw33: BP_sw33=Array(BM_sw33, LM_GI0_l140c_sw33, LM_GI0_l140d_sw33, LM_GI_l141a_sw33, LM_GI_l141_sw33, LM_Flashers_f2_sw33, LM_Flashers_f3_sw33, LM_Flashers_f6_sw33, LM_Inserts_l11_sw33, LM_Inserts_l2_sw33, LM_Inserts_l3_sw33)
Dim BP_sw34: BP_sw34=Array(BM_sw34, LM_GI0_l140b_sw34, LM_GI0_l140c_sw34, LM_GI0_l140d_sw34, LM_GI_l141a_sw34, LM_GI_l141_sw34, LM_Flashers_f2_sw34, LM_Flashers_f3_sw34, LM_Flashers_f5_sw34, LM_Inserts_l11_sw34, LM_Inserts_l2_sw34, LM_Inserts_l3_sw34)
Dim BP_sw35: BP_sw35=Array(BM_sw35, LM_GI0_l140a_sw35, LM_GI0_l140e_sw35, LM_GI0_l140f_sw35, LM_GI_l141_sw35, LM_GI_l141b_sw35, LM_Flashers_f2_sw35, LM_Flashers_f3_sw35, LM_Flashers_f5_sw35, LM_Inserts_l33_sw35, LM_Inserts_l4_sw35, LM_Inserts_l5_sw35, LM_Inserts_l61_sw35)
Dim BP_sw36: BP_sw36=Array(BM_sw36, LM_GI0_l140d_sw36, LM_GI0_l140e_sw36, LM_GI0_l140f_sw36, LM_GI_l141_sw36, LM_GI_l141b_sw36, LM_Flashers_f2_sw36, LM_Flashers_f3_sw36, LM_Flashers_f5_sw36, LM_Inserts_l4_sw36, LM_Inserts_l5_sw36, LM_Inserts_l61_sw36)
Dim BP_sw48: BP_sw48=Array(BM_sw48, LM_GI_l141_sw48, LM_GI_l141d_sw48, LM_Flashers_f1_sw48)
Dim BP_sw49: BP_sw49=Array(BM_sw49, LM_GI_l141_sw49, LM_Flashers_f5_sw49, LM_Flashers_f7_sw49)
Dim BP_sw51: BP_sw51=Array(BM_sw51, LM_Flashers_f1_sw51, LM_Flashers_f3_sw51, LM_Flashers_f6_sw51, LM_Flashers_f8_sw51, LM_Inserts_l39_sw51)
Dim BP_sw52: BP_sw52=Array(BM_sw52, LM_Flashers_f8_sw52)
Dim BP_sw53: BP_sw53=Array(BM_sw53, LM_GI0_l140b_sw53, LM_GI0_l140c_sw53, LM_GI0_l140d_sw53, LM_Flashers_f3_sw53)
Dim BP_sw54: BP_sw54=Array(BM_sw54, LM_GI0_l140b_sw54, LM_GI0_l140c_sw54, LM_GI0_l140d_sw54, LM_Flashers_f2_sw54, LM_Flashers_f3_sw54)
Dim BP_sw55: BP_sw55=Array(BM_sw55, LM_GI0_l140e_sw55, LM_GI0_l140f_sw55, LM_GI_l141b_sw55, LM_Flashers_f3_sw55)
Dim BP_sw56: BP_sw56=Array(BM_sw56, LM_GI0_l140e_sw56, LM_GI0_l140f_sw56, LM_GI_l141b_sw56, LM_Flashers_f3_sw56)
Dim BP_sw57: BP_sw57=Array(BM_sw57, LM_GI0_l140a_sw57, LM_GI0_l140e_sw57, LM_GI0_l140f_sw57, LM_GI_l141_sw57, LM_GI_l141b_sw57, LM_Flashers_f2_sw57, LM_Flashers_f3_sw57, LM_Flashers_f5_sw57, LM_Inserts_l4_sw57, LM_Inserts_l5_sw57, LM_Inserts_l61_sw57, LM_Inserts_l7_sw57)
Dim BP_sw58: BP_sw58=Array(BM_sw58, LM_GI_l141h_sw58, LM_Flashers_f1_sw58)
Dim BP_sw59: BP_sw59=Array(BM_sw59, LM_GI0_l140d_sw59, LM_GI_l141a_sw59, LM_GI_l141_sw59, LM_GI_l141j_sw59, LM_Flashers_f1_sw59, LM_Flashers_f2_sw59, LM_Flashers_f3_sw59, LM_Flashers_f5_sw59, LM_Flashers_f6_sw59, LM_Inserts_l11_sw59, LM_Inserts_l2_sw59, LM_Inserts_l3_sw59)
Dim BP_sw60: BP_sw60=Array(BM_sw60, LM_GI_l141_sw60, LM_GI_l141i_sw60, LM_GI_l141k_sw60, LM_Flashers_f1_sw60, LM_Flashers_f4_sw60)
' Arrays per lighting scenario
Dim BL_Flashers_f1: BL_Flashers_f1=Array(LM_Flashers_f1_B_Caps_Bot, LM_Flashers_f1_B_Caps_Top, LM_Flashers_f1_BP_Empty, LM_Flashers_f1_BP_Jungle, LM_Flashers_f1_BR1, LM_Flashers_f1_BR2, LM_Flashers_f1_BR3, LM_Flashers_f1_BS1, LM_Flashers_f1_BS2, LM_Flashers_f1_BS3, LM_Flashers_f1_Blades, LM_Flashers_f1_LEMK1, LM_Flashers_f1_LFlipper1, LM_Flashers_f1_LFlipper1U, LM_Flashers_f1_Layer_1, LM_Flashers_f1_Layer_2, LM_Flashers_f1_Layer_3, LM_Flashers_f1_MSling1, LM_Flashers_f1_MSling2, LM_Flashers_f1_PF, LM_Flashers_f1_Parts, LM_Flashers_f1_R_Green, LM_Flashers_f1_R_Red, LM_Flashers_f1_R_Yellow, LM_Flashers_f1_gate1, LM_Flashers_f1_gate2, LM_Flashers_f1_sw17, LM_Flashers_f1_sw18, LM_Flashers_f1_sw19, LM_Flashers_f1_sw20, LM_Flashers_f1_sw21, LM_Flashers_f1_sw22, LM_Flashers_f1_sw23, LM_Flashers_f1_sw48, LM_Flashers_f1_sw51, LM_Flashers_f1_sw58, LM_Flashers_f1_sw59, LM_Flashers_f1_sw60)
Dim BL_Flashers_f2: BL_Flashers_f2=Array(LM_Flashers_f2_Blades, LM_Flashers_f2_LFlipper, LM_Flashers_f2_LFlipper1, LM_Flashers_f2_LFlipper1U, LM_Flashers_f2_LFlipperU, LM_Flashers_f2_LSling1, LM_Flashers_f2_LSling2, LM_Flashers_f2_Layer_1, LM_Flashers_f2_Layer_2, LM_Flashers_f2_PF, LM_Flashers_f2_Parts, LM_Flashers_f2_Post_L, LM_Flashers_f2_Post_R, LM_Flashers_f2_R_Green, LM_Flashers_f2_R_Red, LM_Flashers_f2_REMK, LM_Flashers_f2_RFlipper, LM_Flashers_f2_RFlipperU, LM_Flashers_f2_sw33, LM_Flashers_f2_sw34, LM_Flashers_f2_sw35, LM_Flashers_f2_sw36, LM_Flashers_f2_sw54, LM_Flashers_f2_sw57, LM_Flashers_f2_sw59)
Dim BL_Flashers_f3: BL_Flashers_f3=Array(LM_Flashers_f3_Blades, LM_Flashers_f3_LEMK, LM_Flashers_f3_LFlipper, LM_Flashers_f3_LFlipper1, LM_Flashers_f3_LFlipper1U, LM_Flashers_f3_LFlipperU, LM_Flashers_f3_LSling1, LM_Flashers_f3_Layer_1, LM_Flashers_f3_Layer_2, LM_Flashers_f3_MSling1, LM_Flashers_f3_PF, LM_Flashers_f3_Parts, LM_Flashers_f3_Post_L, LM_Flashers_f3_R_Green, LM_Flashers_f3_R_Yellow, LM_Flashers_f3_RFlipper, LM_Flashers_f3_RFlipperU, LM_Flashers_f3_RSling1, LM_Flashers_f3_RSling2, LM_Flashers_f3_sw24, LM_Flashers_f3_sw33, LM_Flashers_f3_sw34, LM_Flashers_f3_sw35, LM_Flashers_f3_sw36, LM_Flashers_f3_sw51, LM_Flashers_f3_sw53, LM_Flashers_f3_sw54, LM_Flashers_f3_sw55, LM_Flashers_f3_sw56, LM_Flashers_f3_sw57, LM_Flashers_f3_sw59)
Dim BL_Flashers_f4: BL_Flashers_f4=Array(LM_Flashers_f4_B_Caps_Bot, LM_Flashers_f4_B_Caps_Top, LM_Flashers_f4_BP_Empty, LM_Flashers_f4_BP_Jungle, LM_Flashers_f4_BR1, LM_Flashers_f4_BR3, LM_Flashers_f4_BS1, LM_Flashers_f4_BS2, LM_Flashers_f4_BS3, LM_Flashers_f4_Blades, LM_Flashers_f4_LEMK1, LM_Flashers_f4_Layer_1, LM_Flashers_f4_Layer_2, LM_Flashers_f4_MSling1, LM_Flashers_f4_MSling2, LM_Flashers_f4_PF, LM_Flashers_f4_Parts, LM_Flashers_f4_R_Green, LM_Flashers_f4_R_Yellow, LM_Flashers_f4_sw17, LM_Flashers_f4_sw18, LM_Flashers_f4_sw19, LM_Flashers_f4_sw20, LM_Flashers_f4_sw60)
Dim BL_Flashers_f5: BL_Flashers_f5=Array(LM_Flashers_f5_Blades, LM_Flashers_f5_PF, LM_Flashers_f5_Parts, LM_Flashers_f5_R_Red, LM_Flashers_f5_R_Yellow, LM_Flashers_f5_sw34, LM_Flashers_f5_sw35, LM_Flashers_f5_sw36, LM_Flashers_f5_sw49, LM_Flashers_f5_sw57, LM_Flashers_f5_sw59)
Dim BL_Flashers_f6: BL_Flashers_f6=Array(LM_Flashers_f6_BS3, LM_Flashers_f6_Blades, LM_Flashers_f6_LFlipper1, LM_Flashers_f6_LFlipper1U, LM_Flashers_f6_PF, LM_Flashers_f6_Parts, LM_Flashers_f6_R_Green, LM_Flashers_f6_R_Yellow, LM_Flashers_f6_sw33, LM_Flashers_f6_sw51, LM_Flashers_f6_sw59)
Dim BL_Flashers_f7: BL_Flashers_f7=Array(LM_Flashers_f7_BS1, LM_Flashers_f7_BS2, LM_Flashers_f7_BS3, LM_Flashers_f7_Layer_1, LM_Flashers_f7_PF, LM_Flashers_f7_Parts, LM_Flashers_f7_R_Red, LM_Flashers_f7_R_Yellow, LM_Flashers_f7_sw49)
Dim BL_Flashers_f8: BL_Flashers_f8=Array(LM_Flashers_f8_B_Caps_Bot, LM_Flashers_f8_B_Caps_Top, LM_Flashers_f8_BR1, LM_Flashers_f8_BR3, LM_Flashers_f8_BS1, LM_Flashers_f8_BS3, LM_Flashers_f8_Blades, LM_Flashers_f8_LEMK1, LM_Flashers_f8_LFlipper1U, LM_Flashers_f8_Layer_1, LM_Flashers_f8_MSling1, LM_Flashers_f8_PF, LM_Flashers_f8_Parts, LM_Flashers_f8_R_Green, LM_Flashers_f8_R_Red, LM_Flashers_f8_R_Yellow, LM_Flashers_f8_sw17, LM_Flashers_f8_sw18, LM_Flashers_f8_sw19, LM_Flashers_f8_sw20, LM_Flashers_f8_sw51, LM_Flashers_f8_sw52)
Dim BL_GI_l141: BL_GI_l141=Array(LM_GI_l141_B_Caps_Bot, LM_GI_l141_B_Caps_Top, LM_GI_l141_BP_Empty, LM_GI_l141_BP_Jungle, LM_GI_l141_BR1, LM_GI_l141_BR2, LM_GI_l141_BR3, LM_GI_l141_BS1, LM_GI_l141_BS2, LM_GI_l141_BS3, LM_GI_l141_Blades, LM_GI_l141_LEMK1, LM_GI_l141_LFlipper, LM_GI_l141_LFlipper1, LM_GI_l141_LFlipper1U, LM_GI_l141_LFlipperU, LM_GI_l141_Layer_1, LM_GI_l141_Layer_2, LM_GI_l141_Layer_3, LM_GI_l141_MSling1, LM_GI_l141_MSling2, LM_GI_l141_PF, LM_GI_l141_Parts, LM_GI_l141_Post_R, LM_GI_l141_R_Green, LM_GI_l141_R_Red, LM_GI_l141_R_Yellow, LM_GI_l141_gate1, LM_GI_l141_gate2, LM_GI_l141_sw17, LM_GI_l141_sw18, LM_GI_l141_sw19, LM_GI_l141_sw20, LM_GI_l141_sw21, LM_GI_l141_sw23, LM_GI_l141_sw33, LM_GI_l141_sw34, LM_GI_l141_sw35, LM_GI_l141_sw36, LM_GI_l141_sw48, LM_GI_l141_sw49, LM_GI_l141_sw57, LM_GI_l141_sw59, LM_GI_l141_sw60)
Dim BL_GI_l141a: BL_GI_l141a=Array(LM_GI_l141a_Blades, LM_GI_l141a_LFlipper1U, LM_GI_l141a_PF, LM_GI_l141a_Parts, LM_GI_l141a_R_Green, LM_GI_l141a_sw33, LM_GI_l141a_sw34, LM_GI_l141a_sw59)
Dim BL_GI_l141b: BL_GI_l141b=Array(LM_GI_l141b_Blades, LM_GI_l141b_PF, LM_GI_l141b_Parts, LM_GI_l141b_Post_R, LM_GI_l141b_RFlipperU, LM_GI_l141b_sw35, LM_GI_l141b_sw36, LM_GI_l141b_sw55, LM_GI_l141b_sw56, LM_GI_l141b_sw57)
Dim BL_GI_l141c: BL_GI_l141c=Array(LM_GI_l141c_Blades, LM_GI_l141c_PF, LM_GI_l141c_Parts, LM_GI_l141c_R_Red)
Dim BL_GI_l141d: BL_GI_l141d=Array(LM_GI_l141d_B_Caps_Bot, LM_GI_l141d_B_Caps_Top, LM_GI_l141d_BP_Empty, LM_GI_l141d_BP_Jungle, LM_GI_l141d_BR2, LM_GI_l141d_BS1, LM_GI_l141d_BS2, LM_GI_l141d_BS3, LM_GI_l141d_Layer_1, LM_GI_l141d_PF, LM_GI_l141d_Parts, LM_GI_l141d_gate1, LM_GI_l141d_gate2, LM_GI_l141d_sw22, LM_GI_l141d_sw23, LM_GI_l141d_sw48)
Dim BL_GI_l141e: BL_GI_l141e=Array(LM_GI_l141e_B_Caps_Bot, LM_GI_l141e_B_Caps_Top, LM_GI_l141e_BP_Empty, LM_GI_l141e_BP_Jungle, LM_GI_l141e_BR2, LM_GI_l141e_BR3, LM_GI_l141e_BS2, LM_GI_l141e_BS3, LM_GI_l141e_Layer_1, LM_GI_l141e_PF, LM_GI_l141e_Parts, LM_GI_l141e_R_Red, LM_GI_l141e_R_Yellow, LM_GI_l141e_gate1, LM_GI_l141e_gate2, LM_GI_l141e_sw21, LM_GI_l141e_sw22, LM_GI_l141e_sw23)
Dim BL_GI_l141f: BL_GI_l141f=Array(LM_GI_l141f_B_Caps_Bot, LM_GI_l141f_B_Caps_Top, LM_GI_l141f_BP_Empty, LM_GI_l141f_BP_Jungle, LM_GI_l141f_BR2, LM_GI_l141f_BR3, LM_GI_l141f_BS2, LM_GI_l141f_BS3, LM_GI_l141f_Layer_1, LM_GI_l141f_Layer_2, LM_GI_l141f_PF, LM_GI_l141f_Parts, LM_GI_l141f_R_Yellow, LM_GI_l141f_gate1, LM_GI_l141f_sw18, LM_GI_l141f_sw19, LM_GI_l141f_sw21, LM_GI_l141f_sw22, LM_GI_l141f_sw23)
Dim BL_GI_l141g: BL_GI_l141g=Array(LM_GI_l141g_B_Caps_Bot, LM_GI_l141g_B_Caps_Top, LM_GI_l141g_BP_Empty, LM_GI_l141g_BP_Jungle, LM_GI_l141g_BS3, LM_GI_l141g_Layer_1, LM_GI_l141g_MSling1, LM_GI_l141g_MSling2, LM_GI_l141g_PF, LM_GI_l141g_Parts, LM_GI_l141g_R_Green, LM_GI_l141g_R_Yellow, LM_GI_l141g_sw17, LM_GI_l141g_sw18, LM_GI_l141g_sw19, LM_GI_l141g_sw20, LM_GI_l141g_sw21, LM_GI_l141g_sw22)
Dim BL_GI_l141h: BL_GI_l141h=Array(LM_GI_l141h_BP_Empty, LM_GI_l141h_BP_Jungle, LM_GI_l141h_Blades, LM_GI_l141h_PF, LM_GI_l141h_Parts, LM_GI_l141h_R_Red, LM_GI_l141h_gate2, LM_GI_l141h_sw58)
Dim BL_GI_l141i: BL_GI_l141i=Array(LM_GI_l141i_B_Caps_Top, LM_GI_l141i_BP_Empty, LM_GI_l141i_BP_Jungle, LM_GI_l141i_Blades, LM_GI_l141i_LFlipper1U, LM_GI_l141i_Layer_1, LM_GI_l141i_Layer_3, LM_GI_l141i_PF, LM_GI_l141i_Parts, LM_GI_l141i_R_Green, LM_GI_l141i_R_Yellow, LM_GI_l141i_sw60)
Dim BL_GI_l141j: BL_GI_l141j=Array(LM_GI_l141j_Blades, LM_GI_l141j_LFlipper1, LM_GI_l141j_LFlipper1U, LM_GI_l141j_PF, LM_GI_l141j_Parts, LM_GI_l141j_R_Green, LM_GI_l141j_R_Yellow, LM_GI_l141j_sw59)
Dim BL_GI_l141k: BL_GI_l141k=Array(LM_GI_l141k_B_Caps_Top, LM_GI_l141k_BP_Empty, LM_GI_l141k_BP_Jungle, LM_GI_l141k_Blades, LM_GI_l141k_Layer_1, LM_GI_l141k_Layer_2, LM_GI_l141k_Layer_3, LM_GI_l141k_MSling1, LM_GI_l141k_PF, LM_GI_l141k_Parts, LM_GI_l141k_R_Green, LM_GI_l141k_R_Yellow, LM_GI_l141k_sw17, LM_GI_l141k_sw18, LM_GI_l141k_sw19, LM_GI_l141k_sw20, LM_GI_l141k_sw60)
Dim BL_GI_l141l: BL_GI_l141l=Array(LM_GI_l141l_B_Caps_Top, LM_GI_l141l_BP_Empty, LM_GI_l141l_BP_Jungle, LM_GI_l141l_BS2, LM_GI_l141l_BS3, LM_GI_l141l_Layer_1, LM_GI_l141l_PF, LM_GI_l141l_Parts, LM_GI_l141l_R_Yellow, LM_GI_l141l_gate1, LM_GI_l141l_gate2)
Dim BL_GI0_l140a: BL_GI0_l140a=Array(LM_GI0_l140a_Blades, LM_GI0_l140a_LFlipper, LM_GI0_l140a_LFlipperU, LM_GI0_l140a_Layer_1, LM_GI0_l140a_Layer_2, LM_GI0_l140a_PF, LM_GI0_l140a_Parts, LM_GI0_l140a_REMK, LM_GI0_l140a_RFlipper, LM_GI0_l140a_RFlipperU, LM_GI0_l140a_sw35, LM_GI0_l140a_sw57)
Dim BL_GI0_l140b: BL_GI0_l140b=Array(LM_GI0_l140b_Blades, LM_GI0_l140b_LEMK, LM_GI0_l140b_LFlipper, LM_GI0_l140b_LFlipperU, LM_GI0_l140b_Layer_1, LM_GI0_l140b_Layer_2, LM_GI0_l140b_PF, LM_GI0_l140b_Parts, LM_GI0_l140b_Post_L, LM_GI0_l140b_RFlipper, LM_GI0_l140b_RFlipperU, LM_GI0_l140b_sw34, LM_GI0_l140b_sw53, LM_GI0_l140b_sw54)
Dim BL_GI0_l140c: BL_GI0_l140c=Array(LM_GI0_l140c_Blades, LM_GI0_l140c_LEMK, LM_GI0_l140c_LFlipper, LM_GI0_l140c_LFlipperU, LM_GI0_l140c_LSling1, LM_GI0_l140c_LSling2, LM_GI0_l140c_Layer_1, LM_GI0_l140c_Layer_2, LM_GI0_l140c_PF, LM_GI0_l140c_Parts, LM_GI0_l140c_Post_L, LM_GI0_l140c_RFlipper, LM_GI0_l140c_RFlipperU, LM_GI0_l140c_sw33, LM_GI0_l140c_sw34, LM_GI0_l140c_sw53, LM_GI0_l140c_sw54)
Dim BL_GI0_l140d: BL_GI0_l140d=Array(LM_GI0_l140d_Blades, LM_GI0_l140d_LEMK, LM_GI0_l140d_LFlipper, LM_GI0_l140d_LFlipperU, LM_GI0_l140d_LSling1, LM_GI0_l140d_LSling2, LM_GI0_l140d_Layer_1, LM_GI0_l140d_Layer_2, LM_GI0_l140d_PF, LM_GI0_l140d_Parts, LM_GI0_l140d_Post_L, LM_GI0_l140d_sw33, LM_GI0_l140d_sw34, LM_GI0_l140d_sw36, LM_GI0_l140d_sw53, LM_GI0_l140d_sw54, LM_GI0_l140d_sw59)
Dim BL_GI0_l140e: BL_GI0_l140e=Array(LM_GI0_l140e_Blades, LM_GI0_l140e_Layer_1, LM_GI0_l140e_Layer_2, LM_GI0_l140e_PF, LM_GI0_l140e_Parts, LM_GI0_l140e_Post_R, LM_GI0_l140e_REMK, LM_GI0_l140e_RFlipper, LM_GI0_l140e_RFlipperU, LM_GI0_l140e_RSling1, LM_GI0_l140e_RSling2, LM_GI0_l140e_sw35, LM_GI0_l140e_sw36, LM_GI0_l140e_sw55, LM_GI0_l140e_sw56, LM_GI0_l140e_sw57)
Dim BL_GI0_l140f: BL_GI0_l140f=Array(LM_GI0_l140f_Blades, LM_GI0_l140f_LFlipper, LM_GI0_l140f_LFlipperU, LM_GI0_l140f_Layer_1, LM_GI0_l140f_Layer_2, LM_GI0_l140f_PF, LM_GI0_l140f_Parts, LM_GI0_l140f_REMK, LM_GI0_l140f_RFlipper, LM_GI0_l140f_RFlipperU, LM_GI0_l140f_RSling1, LM_GI0_l140f_RSling2, LM_GI0_l140f_sw35, LM_GI0_l140f_sw36, LM_GI0_l140f_sw55, LM_GI0_l140f_sw56, LM_GI0_l140f_sw57)
Dim BL_Inserts_l1: BL_Inserts_l1=Array(LM_Inserts_l1_PF, LM_Inserts_l1_Parts)
Dim BL_Inserts_l10: BL_Inserts_l10=Array(LM_Inserts_l10_PF, LM_Inserts_l10_Parts)
Dim BL_Inserts_l11: BL_Inserts_l11=Array(LM_Inserts_l11_LFlipper1, LM_Inserts_l11_LFlipper1U, LM_Inserts_l11_PF, LM_Inserts_l11_Parts, LM_Inserts_l11_R_Green, LM_Inserts_l11_sw33, LM_Inserts_l11_sw34, LM_Inserts_l11_sw59)
Dim BL_Inserts_l12: BL_Inserts_l12=Array(LM_Inserts_l12_PF, LM_Inserts_l12_Parts)
Dim BL_Inserts_l13: BL_Inserts_l13=Array(LM_Inserts_l13_PF, LM_Inserts_l13_Parts)
Dim BL_Inserts_l14: BL_Inserts_l14=Array(LM_Inserts_l14_PF, LM_Inserts_l14_Parts)
Dim BL_Inserts_l15: BL_Inserts_l15=Array(LM_Inserts_l15_PF, LM_Inserts_l15_Parts)
Dim BL_Inserts_l16: BL_Inserts_l16=Array(LM_Inserts_l16_PF, LM_Inserts_l16_Parts)
Dim BL_Inserts_l17: BL_Inserts_l17=Array(LM_Inserts_l17_Layer_1, LM_Inserts_l17_PF, LM_Inserts_l17_Parts, LM_Inserts_l17_Post_L)
Dim BL_Inserts_l18: BL_Inserts_l18=Array(LM_Inserts_l18_Layer_1, LM_Inserts_l18_PF, LM_Inserts_l18_Parts, LM_Inserts_l18_Post_L)
Dim BL_Inserts_l19: BL_Inserts_l19=Array(LM_Inserts_l19_PF, LM_Inserts_l19_Parts)
Dim BL_Inserts_l2: BL_Inserts_l2=Array(LM_Inserts_l2_PF, LM_Inserts_l2_Parts, LM_Inserts_l2_sw33, LM_Inserts_l2_sw34, LM_Inserts_l2_sw59)
Dim BL_Inserts_l20: BL_Inserts_l20=Array(LM_Inserts_l20_PF, LM_Inserts_l20_Parts)
Dim BL_Inserts_l21: BL_Inserts_l21=Array(LM_Inserts_l21_PF, LM_Inserts_l21_Parts)
Dim BL_Inserts_l22: BL_Inserts_l22=Array(LM_Inserts_l22_PF, LM_Inserts_l22_Parts)
Dim BL_Inserts_l23: BL_Inserts_l23=Array(LM_Inserts_l23_PF, LM_Inserts_l23_Parts)
Dim BL_Inserts_l24: BL_Inserts_l24=Array(LM_Inserts_l24_PF, LM_Inserts_l24_Parts)
Dim BL_Inserts_l25: BL_Inserts_l25=Array(LM_Inserts_l25_Layer_1, LM_Inserts_l25_PF, LM_Inserts_l25_Parts)
Dim BL_Inserts_l26: BL_Inserts_l26=Array(LM_Inserts_l26_Layer_1, LM_Inserts_l26_PF, LM_Inserts_l26_Parts, LM_Inserts_l26_Post_R)
Dim BL_Inserts_l27: BL_Inserts_l27=Array(LM_Inserts_l27_LFlipper1U, LM_Inserts_l27_PF, LM_Inserts_l27_Parts, LM_Inserts_l27_R_Yellow)
Dim BL_Inserts_l28: BL_Inserts_l28=Array(LM_Inserts_l28_PF, LM_Inserts_l28_Parts)
Dim BL_Inserts_l29: BL_Inserts_l29=Array(LM_Inserts_l29_BP_Jungle, LM_Inserts_l29_Layer_1, LM_Inserts_l29_MSling1, LM_Inserts_l29_MSling2, LM_Inserts_l29_PF, LM_Inserts_l29_Parts, LM_Inserts_l29_R_Yellow, LM_Inserts_l29_sw17, LM_Inserts_l29_sw18, LM_Inserts_l29_sw19, LM_Inserts_l29_sw20)
Dim BL_Inserts_l3: BL_Inserts_l3=Array(LM_Inserts_l3_PF, LM_Inserts_l3_Parts, LM_Inserts_l3_R_Yellow, LM_Inserts_l3_sw33, LM_Inserts_l3_sw34, LM_Inserts_l3_sw59)
Dim BL_Inserts_l30: BL_Inserts_l30=Array(LM_Inserts_l30_Layer_1, LM_Inserts_l30_MSling1, LM_Inserts_l30_MSling2, LM_Inserts_l30_PF, LM_Inserts_l30_Parts, LM_Inserts_l30_R_Yellow, LM_Inserts_l30_sw17, LM_Inserts_l30_sw18, LM_Inserts_l30_sw19, LM_Inserts_l30_sw20)
Dim BL_Inserts_l31: BL_Inserts_l31=Array(LM_Inserts_l31_B_Caps_Bot, LM_Inserts_l31_B_Caps_Top, LM_Inserts_l31_PF, LM_Inserts_l31_Parts, LM_Inserts_l31_R_Green, LM_Inserts_l31_R_Yellow, LM_Inserts_l31_sw17, LM_Inserts_l31_sw18, LM_Inserts_l31_sw19, LM_Inserts_l31_sw20)
Dim BL_Inserts_l32: BL_Inserts_l32=Array(LM_Inserts_l32_Parts, LM_Inserts_l32_R_Yellow, LM_Inserts_l32_sw20)
Dim BL_Inserts_l33: BL_Inserts_l33=Array(LM_Inserts_l33_B_Caps_Bot, LM_Inserts_l33_PF, LM_Inserts_l33_Parts, LM_Inserts_l33_R_Green, LM_Inserts_l33_R_Yellow, LM_Inserts_l33_sw18, LM_Inserts_l33_sw19, LM_Inserts_l33_sw20, LM_Inserts_l33_sw35)
Dim BL_Inserts_l34: BL_Inserts_l34=Array(LM_Inserts_l34_PF, LM_Inserts_l34_Parts)
Dim BL_Inserts_l35: BL_Inserts_l35=Array(LM_Inserts_l35_PF, LM_Inserts_l35_Parts)
Dim BL_Inserts_l36: BL_Inserts_l36=Array(LM_Inserts_l36_PF, LM_Inserts_l36_Parts)
Dim BL_Inserts_l37: BL_Inserts_l37=Array(LM_Inserts_l37_LFlipperU, LM_Inserts_l37_PF, LM_Inserts_l37_Parts, LM_Inserts_l37_RFlipperU)
Dim BL_Inserts_l38: BL_Inserts_l38=Array(LM_Inserts_l38_Parts, LM_Inserts_l38_R_Yellow)
Dim BL_Inserts_l39: BL_Inserts_l39=Array(LM_Inserts_l39_PF, LM_Inserts_l39_Parts, LM_Inserts_l39_R_Yellow, LM_Inserts_l39_sw51)
Dim BL_Inserts_l4: BL_Inserts_l4=Array(LM_Inserts_l4_PF, LM_Inserts_l4_Parts, LM_Inserts_l4_sw35, LM_Inserts_l4_sw36, LM_Inserts_l4_sw57)
Dim BL_Inserts_l40: BL_Inserts_l40=Array(LM_Inserts_l40_PF, LM_Inserts_l40_Parts, LM_Inserts_l40_R_Yellow)
Dim BL_Inserts_l41: BL_Inserts_l41=Array(LM_Inserts_l41_PF, LM_Inserts_l41_Parts, LM_Inserts_l41_R_Yellow)
Dim BL_Inserts_l42: BL_Inserts_l42=Array(LM_Inserts_l42_PF, LM_Inserts_l42_Parts)
Dim BL_Inserts_l43: BL_Inserts_l43=Array(LM_Inserts_l43_LFlipper1, LM_Inserts_l43_LFlipper1U, LM_Inserts_l43_PF, LM_Inserts_l43_Parts, LM_Inserts_l43_R_Yellow)
Dim BL_Inserts_l44: BL_Inserts_l44=Array(LM_Inserts_l44_LFlipper1, LM_Inserts_l44_LFlipper1U, LM_Inserts_l44_PF, LM_Inserts_l44_Parts, LM_Inserts_l44_R_Yellow)
Dim BL_Inserts_l45: BL_Inserts_l45=Array(LM_Inserts_l45_LFlipper1U, LM_Inserts_l45_PF, LM_Inserts_l45_Parts, LM_Inserts_l45_R_Yellow)
Dim BL_Inserts_l46: BL_Inserts_l46=Array(LM_Inserts_l46_LFlipper, LM_Inserts_l46_LFlipperU, LM_Inserts_l46_PF, LM_Inserts_l46_Parts, LM_Inserts_l46_RFlipper, LM_Inserts_l46_RFlipperU)
Dim BL_Inserts_l47: BL_Inserts_l47=Array(LM_Inserts_l47_PF, LM_Inserts_l47_Parts, LM_Inserts_l47_R_Yellow)
Dim BL_Inserts_l48: BL_Inserts_l48=Array(LM_Inserts_l48_B_Caps_Bot, LM_Inserts_l48_B_Caps_Top, LM_Inserts_l48_PF, LM_Inserts_l48_Parts, LM_Inserts_l48_R_Red)
Dim BL_Inserts_l49: BL_Inserts_l49=Array(LM_Inserts_l49_PF, LM_Inserts_l49_Parts, LM_Inserts_l49_R_Yellow)
Dim BL_Inserts_l5: BL_Inserts_l5=Array(LM_Inserts_l5_PF, LM_Inserts_l5_Parts, LM_Inserts_l5_sw35, LM_Inserts_l5_sw36, LM_Inserts_l5_sw57)
Dim BL_Inserts_l50: BL_Inserts_l50=Array(LM_Inserts_l50_LFlipper1, LM_Inserts_l50_LFlipper1U, LM_Inserts_l50_PF, LM_Inserts_l50_Parts, LM_Inserts_l50_R_Green, LM_Inserts_l50_R_Yellow)
Dim BL_Inserts_l51: BL_Inserts_l51=Array(LM_Inserts_l51_LFlipper1U, LM_Inserts_l51_PF, LM_Inserts_l51_Parts, LM_Inserts_l51_R_Green, LM_Inserts_l51_R_Yellow)
Dim BL_Inserts_l52: BL_Inserts_l52=Array(LM_Inserts_l52_LFlipper1U, LM_Inserts_l52_Layer_1, LM_Inserts_l52_PF, LM_Inserts_l52_Parts, LM_Inserts_l52_R_Green, LM_Inserts_l52_R_Yellow)
Dim BL_Inserts_l53: BL_Inserts_l53=Array(LM_Inserts_l53_Layer_1, LM_Inserts_l53_PF, LM_Inserts_l53_Parts, LM_Inserts_l53_R_Green, LM_Inserts_l53_R_Yellow)
Dim BL_Inserts_l54: BL_Inserts_l54=Array(LM_Inserts_l54_PF, LM_Inserts_l54_Parts)
Dim BL_Inserts_l55: BL_Inserts_l55=Array(LM_Inserts_l55_Blades, LM_Inserts_l55_PF, LM_Inserts_l55_Parts, LM_Inserts_l55_R_Yellow)
Dim BL_Inserts_l56: BL_Inserts_l56=Array(LM_Inserts_l56_LFlipper1, LM_Inserts_l56_LFlipper1U, LM_Inserts_l56_Layer_1, LM_Inserts_l56_PF, LM_Inserts_l56_Parts, LM_Inserts_l56_R_Green, LM_Inserts_l56_R_Yellow)
Dim BL_Inserts_l57: BL_Inserts_l57=Array(LM_Inserts_l57_BP_Empty, LM_Inserts_l57_Layer_1, LM_Inserts_l57_PF, LM_Inserts_l57_Parts)
Dim BL_Inserts_l58: BL_Inserts_l58=Array(LM_Inserts_l58_B_Caps_Top, LM_Inserts_l58_BP_Empty, LM_Inserts_l58_BP_Jungle, LM_Inserts_l58_PF, LM_Inserts_l58_Parts)
Dim BL_Inserts_l59: BL_Inserts_l59=Array(LM_Inserts_l59_B_Caps_Top, LM_Inserts_l59_BP_Empty, LM_Inserts_l59_PF, LM_Inserts_l59_Parts, LM_Inserts_l59_gate2)
Dim BL_Inserts_l6: BL_Inserts_l6=Array(LM_Inserts_l6_PF, LM_Inserts_l6_Parts)
Dim BL_Inserts_l60: BL_Inserts_l60=Array(LM_Inserts_l60_B_Caps_Top, LM_Inserts_l60_BP_Empty, LM_Inserts_l60_PF, LM_Inserts_l60_Parts, LM_Inserts_l60_R_Red, LM_Inserts_l60_gate1, LM_Inserts_l60_gate2)
Dim BL_Inserts_l61: BL_Inserts_l61=Array(LM_Inserts_l61_Layer_1, LM_Inserts_l61_PF, LM_Inserts_l61_Parts, LM_Inserts_l61_sw35, LM_Inserts_l61_sw36, LM_Inserts_l61_sw57)
Dim BL_Inserts_l62: BL_Inserts_l62=Array(LM_Inserts_l62_Blades, LM_Inserts_l62_PF, LM_Inserts_l62_Parts)
Dim BL_Inserts_l7: BL_Inserts_l7=Array(LM_Inserts_l7_PF, LM_Inserts_l7_Parts, LM_Inserts_l7_sw57)
Dim BL_Inserts_l8: BL_Inserts_l8=Array(LM_Inserts_l8_PF, LM_Inserts_l8_Parts)
Dim BL_Inserts_l9: BL_Inserts_l9=Array(LM_Inserts_l9_PF, LM_Inserts_l9_Parts)
Dim BL_Room: BL_Room=Array(BM_B_Caps_Bot, BM_B_Caps_Top, BM_BP_Empty, BM_BP_Jungle, BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_Blades, BM_LEMK, BM_LEMK1, BM_LFlipper, BM_LFlipper1, BM_LFlipper1U, BM_LFlipperU, BM_LSling1, BM_LSling2, BM_Layer_1, BM_Layer_2, BM_Layer_3, BM_MSling1, BM_MSling2, BM_PF, BM_Parts, BM_Post_L, BM_Post_R, BM_R_Green, BM_R_Red, BM_R_Yellow, BM_REMK, BM_RFlipper, BM_RFlipperU, BM_RSling1, BM_RSling2, BM_Trap, BM_gate1, BM_gate2, BM_sw16, BM_sw17, BM_sw18, BM_sw19, BM_sw20, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw33, BM_sw34, BM_sw35, BM_sw36, BM_sw48, BM_sw49, BM_sw51, BM_sw52, BM_sw53, BM_sw54, BM_sw55, BM_sw56, BM_sw57, BM_sw58, BM_sw59, BM_sw60)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_B_Caps_Bot, BM_B_Caps_Top, BM_BP_Empty, BM_BP_Jungle, BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_Blades, BM_LEMK, BM_LEMK1, BM_LFlipper, BM_LFlipper1, BM_LFlipper1U, BM_LFlipperU, BM_LSling1, BM_LSling2, BM_Layer_1, BM_Layer_2, BM_Layer_3, BM_MSling1, BM_MSling2, BM_PF, BM_Parts, BM_Post_L, BM_Post_R, BM_R_Green, BM_R_Red, BM_R_Yellow, BM_REMK, BM_RFlipper, BM_RFlipperU, BM_RSling1, BM_RSling2, BM_Trap, BM_gate1, BM_gate2, BM_sw16, BM_sw17, BM_sw18, BM_sw19, BM_sw20, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw33, BM_sw34, BM_sw35, BM_sw36, BM_sw48, BM_sw49, BM_sw51, BM_sw52, BM_sw53, BM_sw54, BM_sw55, BM_sw56, BM_sw57, BM_sw58, BM_sw59, BM_sw60)
Dim BG_Lightmap: BG_Lightmap=Array(LM_Flashers_f1_B_Caps_Bot, LM_Flashers_f1_B_Caps_Top, LM_Flashers_f1_BP_Empty, LM_Flashers_f1_BP_Jungle, LM_Flashers_f1_BR1, LM_Flashers_f1_BR2, LM_Flashers_f1_BR3, LM_Flashers_f1_BS1, LM_Flashers_f1_BS2, LM_Flashers_f1_BS3, LM_Flashers_f1_Blades, LM_Flashers_f1_LEMK1, LM_Flashers_f1_LFlipper1, LM_Flashers_f1_LFlipper1U, LM_Flashers_f1_Layer_1, LM_Flashers_f1_Layer_2, LM_Flashers_f1_Layer_3, LM_Flashers_f1_MSling1, LM_Flashers_f1_MSling2, LM_Flashers_f1_PF, LM_Flashers_f1_Parts, LM_Flashers_f1_R_Green, LM_Flashers_f1_R_Red, LM_Flashers_f1_R_Yellow, LM_Flashers_f1_gate1, LM_Flashers_f1_gate2, LM_Flashers_f1_sw17, LM_Flashers_f1_sw18, LM_Flashers_f1_sw19, LM_Flashers_f1_sw20, LM_Flashers_f1_sw21, LM_Flashers_f1_sw22, LM_Flashers_f1_sw23, LM_Flashers_f1_sw48, LM_Flashers_f1_sw51, LM_Flashers_f1_sw58, LM_Flashers_f1_sw59, LM_Flashers_f1_sw60, LM_Flashers_f2_Blades, LM_Flashers_f2_LFlipper, LM_Flashers_f2_LFlipper1, LM_Flashers_f2_LFlipper1U, _
	LM_Flashers_f2_LFlipperU, LM_Flashers_f2_LSling1, LM_Flashers_f2_LSling2, LM_Flashers_f2_Layer_1, LM_Flashers_f2_Layer_2, LM_Flashers_f2_PF, LM_Flashers_f2_Parts, LM_Flashers_f2_Post_L, LM_Flashers_f2_Post_R, LM_Flashers_f2_R_Green, LM_Flashers_f2_R_Red, LM_Flashers_f2_REMK, LM_Flashers_f2_RFlipper, LM_Flashers_f2_RFlipperU, LM_Flashers_f2_sw33, LM_Flashers_f2_sw34, LM_Flashers_f2_sw35, LM_Flashers_f2_sw36, LM_Flashers_f2_sw54, LM_Flashers_f2_sw57, LM_Flashers_f2_sw59, LM_Flashers_f3_Blades, LM_Flashers_f3_LEMK, LM_Flashers_f3_LFlipper, LM_Flashers_f3_LFlipper1, LM_Flashers_f3_LFlipper1U, LM_Flashers_f3_LFlipperU, LM_Flashers_f3_LSling1, LM_Flashers_f3_Layer_1, LM_Flashers_f3_Layer_2, LM_Flashers_f3_MSling1, LM_Flashers_f3_PF, LM_Flashers_f3_Parts, LM_Flashers_f3_Post_L, LM_Flashers_f3_R_Green, LM_Flashers_f3_R_Yellow, LM_Flashers_f3_RFlipper, LM_Flashers_f3_RFlipperU, LM_Flashers_f3_RSling1, LM_Flashers_f3_RSling2, LM_Flashers_f3_sw24, LM_Flashers_f3_sw33, _
	LM_Flashers_f3_sw34, LM_Flashers_f3_sw35, LM_Flashers_f3_sw36, LM_Flashers_f3_sw51, LM_Flashers_f3_sw53, LM_Flashers_f3_sw54, LM_Flashers_f3_sw55, LM_Flashers_f3_sw56, LM_Flashers_f3_sw57, LM_Flashers_f3_sw59, LM_Flashers_f4_B_Caps_Bot, LM_Flashers_f4_B_Caps_Top, LM_Flashers_f4_BP_Empty, LM_Flashers_f4_BP_Jungle, LM_Flashers_f4_BR1, LM_Flashers_f4_BR3, LM_Flashers_f4_BS1, LM_Flashers_f4_BS2, LM_Flashers_f4_BS3, LM_Flashers_f4_Blades, LM_Flashers_f4_LEMK1, LM_Flashers_f4_Layer_1, LM_Flashers_f4_Layer_2, LM_Flashers_f4_MSling1, LM_Flashers_f4_MSling2, LM_Flashers_f4_PF, LM_Flashers_f4_Parts, LM_Flashers_f4_R_Green, LM_Flashers_f4_R_Yellow, LM_Flashers_f4_sw17, LM_Flashers_f4_sw18, LM_Flashers_f4_sw19, LM_Flashers_f4_sw20, LM_Flashers_f4_sw60, LM_Flashers_f5_Blades, LM_Flashers_f5_PF, LM_Flashers_f5_Parts, LM_Flashers_f5_R_Red, LM_Flashers_f5_R_Yellow, LM_Flashers_f5_sw34, LM_Flashers_f5_sw35, LM_Flashers_f5_sw36, LM_Flashers_f5_sw49, LM_Flashers_f5_sw57, LM_Flashers_f5_sw59, LM_Flashers_f6_BS3, _
	LM_Flashers_f6_Blades, LM_Flashers_f6_LFlipper1, LM_Flashers_f6_LFlipper1U, LM_Flashers_f6_PF, LM_Flashers_f6_Parts, LM_Flashers_f6_R_Green, LM_Flashers_f6_R_Yellow, LM_Flashers_f6_sw33, LM_Flashers_f6_sw51, LM_Flashers_f6_sw59, LM_Flashers_f7_BS1, LM_Flashers_f7_BS2, LM_Flashers_f7_BS3, LM_Flashers_f7_Layer_1, LM_Flashers_f7_PF, LM_Flashers_f7_Parts, LM_Flashers_f7_R_Red, LM_Flashers_f7_R_Yellow, LM_Flashers_f7_sw49, LM_Flashers_f8_B_Caps_Bot, LM_Flashers_f8_B_Caps_Top, LM_Flashers_f8_BR1, LM_Flashers_f8_BR3, LM_Flashers_f8_BS1, LM_Flashers_f8_BS3, LM_Flashers_f8_Blades, LM_Flashers_f8_LEMK1, LM_Flashers_f8_LFlipper1U, LM_Flashers_f8_Layer_1, LM_Flashers_f8_MSling1, LM_Flashers_f8_PF, LM_Flashers_f8_Parts, LM_Flashers_f8_R_Green, LM_Flashers_f8_R_Red, LM_Flashers_f8_R_Yellow, LM_Flashers_f8_sw17, LM_Flashers_f8_sw18, LM_Flashers_f8_sw19, LM_Flashers_f8_sw20, LM_Flashers_f8_sw51, LM_Flashers_f8_sw52, LM_GI_l141_B_Caps_Bot, LM_GI_l141_B_Caps_Top, LM_GI_l141_BP_Empty, LM_GI_l141_BP_Jungle, LM_GI_l141_BR1, _
	LM_GI_l141_BR2, LM_GI_l141_BR3, LM_GI_l141_BS1, LM_GI_l141_BS2, LM_GI_l141_BS3, LM_GI_l141_Blades, LM_GI_l141_LEMK1, LM_GI_l141_LFlipper, LM_GI_l141_LFlipper1, LM_GI_l141_LFlipper1U, LM_GI_l141_LFlipperU, LM_GI_l141_Layer_1, LM_GI_l141_Layer_2, LM_GI_l141_Layer_3, LM_GI_l141_MSling1, LM_GI_l141_MSling2, LM_GI_l141_PF, LM_GI_l141_Parts, LM_GI_l141_Post_R, LM_GI_l141_R_Green, LM_GI_l141_R_Red, LM_GI_l141_R_Yellow, LM_GI_l141_gate1, LM_GI_l141_gate2, LM_GI_l141_sw17, LM_GI_l141_sw18, LM_GI_l141_sw19, LM_GI_l141_sw20, LM_GI_l141_sw21, LM_GI_l141_sw23, LM_GI_l141_sw33, LM_GI_l141_sw34, LM_GI_l141_sw35, LM_GI_l141_sw36, LM_GI_l141_sw48, LM_GI_l141_sw49, LM_GI_l141_sw57, LM_GI_l141_sw59, LM_GI_l141_sw60, LM_GI_l141a_Blades, LM_GI_l141a_LFlipper1U, LM_GI_l141a_PF, LM_GI_l141a_Parts, LM_GI_l141a_R_Green, LM_GI_l141a_sw33, LM_GI_l141a_sw34, LM_GI_l141a_sw59, LM_GI_l141b_Blades, LM_GI_l141b_PF, LM_GI_l141b_Parts, LM_GI_l141b_Post_R, LM_GI_l141b_RFlipperU, LM_GI_l141b_sw35, LM_GI_l141b_sw36, LM_GI_l141b_sw55, _
	LM_GI_l141b_sw56, LM_GI_l141b_sw57, LM_GI_l141c_Blades, LM_GI_l141c_PF, LM_GI_l141c_Parts, LM_GI_l141c_R_Red, LM_GI_l141d_B_Caps_Bot, LM_GI_l141d_B_Caps_Top, LM_GI_l141d_BP_Empty, LM_GI_l141d_BP_Jungle, LM_GI_l141d_BR2, LM_GI_l141d_BS1, LM_GI_l141d_BS2, LM_GI_l141d_BS3, LM_GI_l141d_Layer_1, LM_GI_l141d_PF, LM_GI_l141d_Parts, LM_GI_l141d_gate1, LM_GI_l141d_gate2, LM_GI_l141d_sw22, LM_GI_l141d_sw23, LM_GI_l141d_sw48, LM_GI_l141e_B_Caps_Bot, LM_GI_l141e_B_Caps_Top, LM_GI_l141e_BP_Empty, LM_GI_l141e_BP_Jungle, LM_GI_l141e_BR2, LM_GI_l141e_BR3, LM_GI_l141e_BS2, LM_GI_l141e_BS3, LM_GI_l141e_Layer_1, LM_GI_l141e_PF, LM_GI_l141e_Parts, LM_GI_l141e_R_Red, LM_GI_l141e_R_Yellow, LM_GI_l141e_gate1, LM_GI_l141e_gate2, LM_GI_l141e_sw21, LM_GI_l141e_sw22, LM_GI_l141e_sw23, LM_GI_l141f_B_Caps_Bot, LM_GI_l141f_B_Caps_Top, LM_GI_l141f_BP_Empty, LM_GI_l141f_BP_Jungle, LM_GI_l141f_BR2, LM_GI_l141f_BR3, LM_GI_l141f_BS2, LM_GI_l141f_BS3, LM_GI_l141f_Layer_1, LM_GI_l141f_Layer_2, LM_GI_l141f_PF, LM_GI_l141f_Parts, _
	LM_GI_l141f_R_Yellow, LM_GI_l141f_gate1, LM_GI_l141f_sw18, LM_GI_l141f_sw19, LM_GI_l141f_sw21, LM_GI_l141f_sw22, LM_GI_l141f_sw23, LM_GI_l141g_B_Caps_Bot, LM_GI_l141g_B_Caps_Top, LM_GI_l141g_BP_Empty, LM_GI_l141g_BP_Jungle, LM_GI_l141g_BS3, LM_GI_l141g_Layer_1, LM_GI_l141g_MSling1, LM_GI_l141g_MSling2, LM_GI_l141g_PF, LM_GI_l141g_Parts, LM_GI_l141g_R_Green, LM_GI_l141g_R_Yellow, LM_GI_l141g_sw17, LM_GI_l141g_sw18, LM_GI_l141g_sw19, LM_GI_l141g_sw20, LM_GI_l141g_sw21, LM_GI_l141g_sw22, LM_GI_l141h_BP_Empty, LM_GI_l141h_BP_Jungle, LM_GI_l141h_Blades, LM_GI_l141h_PF, LM_GI_l141h_Parts, LM_GI_l141h_R_Red, LM_GI_l141h_gate2, LM_GI_l141h_sw58, LM_GI_l141i_B_Caps_Top, LM_GI_l141i_BP_Empty, LM_GI_l141i_BP_Jungle, LM_GI_l141i_Blades, LM_GI_l141i_LFlipper1U, LM_GI_l141i_Layer_1, LM_GI_l141i_Layer_3, LM_GI_l141i_PF, LM_GI_l141i_Parts, LM_GI_l141i_R_Green, LM_GI_l141i_R_Yellow, LM_GI_l141i_sw60, LM_GI_l141j_Blades, LM_GI_l141j_LFlipper1, LM_GI_l141j_LFlipper1U, LM_GI_l141j_PF, LM_GI_l141j_Parts, LM_GI_l141j_R_Green, _
	LM_GI_l141j_R_Yellow, LM_GI_l141j_sw59, LM_GI_l141k_B_Caps_Top, LM_GI_l141k_BP_Empty, LM_GI_l141k_BP_Jungle, LM_GI_l141k_Blades, LM_GI_l141k_Layer_1, LM_GI_l141k_Layer_2, LM_GI_l141k_Layer_3, LM_GI_l141k_MSling1, LM_GI_l141k_PF, LM_GI_l141k_Parts, LM_GI_l141k_R_Green, LM_GI_l141k_R_Yellow, LM_GI_l141k_sw17, LM_GI_l141k_sw18, LM_GI_l141k_sw19, LM_GI_l141k_sw20, LM_GI_l141k_sw60, LM_GI_l141l_B_Caps_Top, LM_GI_l141l_BP_Empty, LM_GI_l141l_BP_Jungle, LM_GI_l141l_BS2, LM_GI_l141l_BS3, LM_GI_l141l_Layer_1, LM_GI_l141l_PF, LM_GI_l141l_Parts, LM_GI_l141l_R_Yellow, LM_GI_l141l_gate1, LM_GI_l141l_gate2, LM_GI0_l140a_Blades, LM_GI0_l140a_LFlipper, LM_GI0_l140a_LFlipperU, LM_GI0_l140a_Layer_1, LM_GI0_l140a_Layer_2, LM_GI0_l140a_PF, LM_GI0_l140a_Parts, LM_GI0_l140a_REMK, LM_GI0_l140a_RFlipper, LM_GI0_l140a_RFlipperU, LM_GI0_l140a_sw35, LM_GI0_l140a_sw57, LM_GI0_l140b_Blades, LM_GI0_l140b_LEMK, LM_GI0_l140b_LFlipper, LM_GI0_l140b_LFlipperU, LM_GI0_l140b_Layer_1, LM_GI0_l140b_Layer_2, LM_GI0_l140b_PF, LM_GI0_l140b_Parts, _
	LM_GI0_l140b_Post_L, LM_GI0_l140b_RFlipper, LM_GI0_l140b_RFlipperU, LM_GI0_l140b_sw34, LM_GI0_l140b_sw53, LM_GI0_l140b_sw54, LM_GI0_l140c_Blades, LM_GI0_l140c_LEMK, LM_GI0_l140c_LFlipper, LM_GI0_l140c_LFlipperU, LM_GI0_l140c_LSling1, LM_GI0_l140c_LSling2, LM_GI0_l140c_Layer_1, LM_GI0_l140c_Layer_2, LM_GI0_l140c_PF, LM_GI0_l140c_Parts, LM_GI0_l140c_Post_L, LM_GI0_l140c_RFlipper, LM_GI0_l140c_RFlipperU, LM_GI0_l140c_sw33, LM_GI0_l140c_sw34, LM_GI0_l140c_sw53, LM_GI0_l140c_sw54, LM_GI0_l140d_Blades, LM_GI0_l140d_LEMK, LM_GI0_l140d_LFlipper, LM_GI0_l140d_LFlipperU, LM_GI0_l140d_LSling1, LM_GI0_l140d_LSling2, LM_GI0_l140d_Layer_1, LM_GI0_l140d_Layer_2, LM_GI0_l140d_PF, LM_GI0_l140d_Parts, LM_GI0_l140d_Post_L, LM_GI0_l140d_sw33, LM_GI0_l140d_sw34, LM_GI0_l140d_sw36, LM_GI0_l140d_sw53, LM_GI0_l140d_sw54, LM_GI0_l140d_sw59, LM_GI0_l140e_Blades, LM_GI0_l140e_Layer_1, LM_GI0_l140e_Layer_2, LM_GI0_l140e_PF, LM_GI0_l140e_Parts, LM_GI0_l140e_Post_R, LM_GI0_l140e_REMK, LM_GI0_l140e_RFlipper, LM_GI0_l140e_RFlipperU, _
	LM_GI0_l140e_RSling1, LM_GI0_l140e_RSling2, LM_GI0_l140e_sw35, LM_GI0_l140e_sw36, LM_GI0_l140e_sw55, LM_GI0_l140e_sw56, LM_GI0_l140e_sw57, LM_GI0_l140f_Blades, LM_GI0_l140f_LFlipper, LM_GI0_l140f_LFlipperU, LM_GI0_l140f_Layer_1, LM_GI0_l140f_Layer_2, LM_GI0_l140f_PF, LM_GI0_l140f_Parts, LM_GI0_l140f_REMK, LM_GI0_l140f_RFlipper, LM_GI0_l140f_RFlipperU, LM_GI0_l140f_RSling1, LM_GI0_l140f_RSling2, LM_GI0_l140f_sw35, LM_GI0_l140f_sw36, LM_GI0_l140f_sw55, LM_GI0_l140f_sw56, LM_GI0_l140f_sw57, LM_Inserts_l1_PF, LM_Inserts_l1_Parts, LM_Inserts_l10_PF, LM_Inserts_l10_Parts, LM_Inserts_l11_LFlipper1, LM_Inserts_l11_LFlipper1U, LM_Inserts_l11_PF, LM_Inserts_l11_Parts, LM_Inserts_l11_R_Green, LM_Inserts_l11_sw33, LM_Inserts_l11_sw34, LM_Inserts_l11_sw59, LM_Inserts_l12_PF, LM_Inserts_l12_Parts, LM_Inserts_l13_PF, LM_Inserts_l13_Parts, LM_Inserts_l14_PF, LM_Inserts_l14_Parts, LM_Inserts_l15_PF, LM_Inserts_l15_Parts, LM_Inserts_l16_PF, LM_Inserts_l16_Parts, LM_Inserts_l17_Layer_1, LM_Inserts_l17_PF, LM_Inserts_l17_Parts, _
	LM_Inserts_l17_Post_L, LM_Inserts_l18_Layer_1, LM_Inserts_l18_PF, LM_Inserts_l18_Parts, LM_Inserts_l18_Post_L, LM_Inserts_l19_PF, LM_Inserts_l19_Parts, LM_Inserts_l2_PF, LM_Inserts_l2_Parts, LM_Inserts_l2_sw33, LM_Inserts_l2_sw34, LM_Inserts_l2_sw59, LM_Inserts_l20_PF, LM_Inserts_l20_Parts, LM_Inserts_l21_PF, LM_Inserts_l21_Parts, LM_Inserts_l22_PF, LM_Inserts_l22_Parts, LM_Inserts_l23_PF, LM_Inserts_l23_Parts, LM_Inserts_l24_PF, LM_Inserts_l24_Parts, LM_Inserts_l25_Layer_1, LM_Inserts_l25_PF, LM_Inserts_l25_Parts, LM_Inserts_l26_Layer_1, LM_Inserts_l26_PF, LM_Inserts_l26_Parts, LM_Inserts_l26_Post_R, LM_Inserts_l27_LFlipper1U, LM_Inserts_l27_PF, LM_Inserts_l27_Parts, LM_Inserts_l27_R_Yellow, LM_Inserts_l28_PF, LM_Inserts_l28_Parts, LM_Inserts_l29_BP_Jungle, LM_Inserts_l29_Layer_1, LM_Inserts_l29_MSling1, LM_Inserts_l29_MSling2, LM_Inserts_l29_PF, LM_Inserts_l29_Parts, LM_Inserts_l29_R_Yellow, LM_Inserts_l29_sw17, LM_Inserts_l29_sw18, LM_Inserts_l29_sw19, LM_Inserts_l29_sw20, LM_Inserts_l3_PF, _
	LM_Inserts_l3_Parts, LM_Inserts_l3_R_Yellow, LM_Inserts_l3_sw33, LM_Inserts_l3_sw34, LM_Inserts_l3_sw59, LM_Inserts_l30_Layer_1, LM_Inserts_l30_MSling1, LM_Inserts_l30_MSling2, LM_Inserts_l30_PF, LM_Inserts_l30_Parts, LM_Inserts_l30_R_Yellow, LM_Inserts_l30_sw17, LM_Inserts_l30_sw18, LM_Inserts_l30_sw19, LM_Inserts_l30_sw20, LM_Inserts_l31_B_Caps_Bot, LM_Inserts_l31_B_Caps_Top, LM_Inserts_l31_PF, LM_Inserts_l31_Parts, LM_Inserts_l31_R_Green, LM_Inserts_l31_R_Yellow, LM_Inserts_l31_sw17, LM_Inserts_l31_sw18, LM_Inserts_l31_sw19, LM_Inserts_l31_sw20, LM_Inserts_l32_Parts, LM_Inserts_l32_R_Yellow, LM_Inserts_l32_sw20, LM_Inserts_l33_B_Caps_Bot, LM_Inserts_l33_PF, LM_Inserts_l33_Parts, LM_Inserts_l33_R_Green, LM_Inserts_l33_R_Yellow, LM_Inserts_l33_sw18, LM_Inserts_l33_sw19, LM_Inserts_l33_sw20, LM_Inserts_l33_sw35, LM_Inserts_l34_PF, LM_Inserts_l34_Parts, LM_Inserts_l35_PF, LM_Inserts_l35_Parts, LM_Inserts_l36_PF, LM_Inserts_l36_Parts, LM_Inserts_l37_LFlipperU, LM_Inserts_l37_PF, LM_Inserts_l37_Parts, _
	LM_Inserts_l37_RFlipperU, LM_Inserts_l38_Parts, LM_Inserts_l38_R_Yellow, LM_Inserts_l39_PF, LM_Inserts_l39_Parts, LM_Inserts_l39_R_Yellow, LM_Inserts_l39_sw51, LM_Inserts_l4_PF, LM_Inserts_l4_Parts, LM_Inserts_l4_sw35, LM_Inserts_l4_sw36, LM_Inserts_l4_sw57, LM_Inserts_l40_PF, LM_Inserts_l40_Parts, LM_Inserts_l40_R_Yellow, LM_Inserts_l41_PF, LM_Inserts_l41_Parts, LM_Inserts_l41_R_Yellow, LM_Inserts_l42_PF, LM_Inserts_l42_Parts, LM_Inserts_l43_LFlipper1, LM_Inserts_l43_LFlipper1U, LM_Inserts_l43_PF, LM_Inserts_l43_Parts, LM_Inserts_l43_R_Yellow, LM_Inserts_l44_LFlipper1, LM_Inserts_l44_LFlipper1U, LM_Inserts_l44_PF, LM_Inserts_l44_Parts, LM_Inserts_l44_R_Yellow, LM_Inserts_l45_LFlipper1U, LM_Inserts_l45_PF, LM_Inserts_l45_Parts, LM_Inserts_l45_R_Yellow, LM_Inserts_l46_LFlipper, LM_Inserts_l46_LFlipperU, LM_Inserts_l46_PF, LM_Inserts_l46_Parts, LM_Inserts_l46_RFlipper, LM_Inserts_l46_RFlipperU, LM_Inserts_l47_PF, LM_Inserts_l47_Parts, LM_Inserts_l47_R_Yellow, LM_Inserts_l48_B_Caps_Bot, _
	LM_Inserts_l48_B_Caps_Top, LM_Inserts_l48_PF, LM_Inserts_l48_Parts, LM_Inserts_l48_R_Red, LM_Inserts_l49_PF, LM_Inserts_l49_Parts, LM_Inserts_l49_R_Yellow, LM_Inserts_l5_PF, LM_Inserts_l5_Parts, LM_Inserts_l5_sw35, LM_Inserts_l5_sw36, LM_Inserts_l5_sw57, LM_Inserts_l50_LFlipper1, LM_Inserts_l50_LFlipper1U, LM_Inserts_l50_PF, LM_Inserts_l50_Parts, LM_Inserts_l50_R_Green, LM_Inserts_l50_R_Yellow, LM_Inserts_l51_LFlipper1U, LM_Inserts_l51_PF, LM_Inserts_l51_Parts, LM_Inserts_l51_R_Green, LM_Inserts_l51_R_Yellow, LM_Inserts_l52_LFlipper1U, LM_Inserts_l52_Layer_1, LM_Inserts_l52_PF, LM_Inserts_l52_Parts, LM_Inserts_l52_R_Green, LM_Inserts_l52_R_Yellow, LM_Inserts_l53_Layer_1, LM_Inserts_l53_PF, LM_Inserts_l53_Parts, LM_Inserts_l53_R_Green, LM_Inserts_l53_R_Yellow, LM_Inserts_l54_PF, LM_Inserts_l54_Parts, LM_Inserts_l55_Blades, LM_Inserts_l55_PF, LM_Inserts_l55_Parts, LM_Inserts_l55_R_Yellow, LM_Inserts_l56_LFlipper1, LM_Inserts_l56_LFlipper1U, LM_Inserts_l56_Layer_1, LM_Inserts_l56_PF, LM_Inserts_l56_Parts, _
	LM_Inserts_l56_R_Green, LM_Inserts_l56_R_Yellow, LM_Inserts_l57_BP_Empty, LM_Inserts_l57_Layer_1, LM_Inserts_l57_PF, LM_Inserts_l57_Parts, LM_Inserts_l58_B_Caps_Top, LM_Inserts_l58_BP_Empty, LM_Inserts_l58_BP_Jungle, LM_Inserts_l58_PF, LM_Inserts_l58_Parts, LM_Inserts_l59_B_Caps_Top, LM_Inserts_l59_BP_Empty, LM_Inserts_l59_PF, LM_Inserts_l59_Parts, LM_Inserts_l59_gate2, LM_Inserts_l6_PF, LM_Inserts_l6_Parts, LM_Inserts_l60_B_Caps_Top, LM_Inserts_l60_BP_Empty, LM_Inserts_l60_PF, LM_Inserts_l60_Parts, LM_Inserts_l60_R_Red, LM_Inserts_l60_gate1, LM_Inserts_l60_gate2, LM_Inserts_l61_Layer_1, LM_Inserts_l61_PF, LM_Inserts_l61_Parts, LM_Inserts_l61_sw35, LM_Inserts_l61_sw36, LM_Inserts_l61_sw57, LM_Inserts_l62_Blades, LM_Inserts_l62_PF, LM_Inserts_l62_Parts, LM_Inserts_l7_PF, LM_Inserts_l7_Parts, LM_Inserts_l7_sw57, LM_Inserts_l8_PF, LM_Inserts_l8_Parts, LM_Inserts_l9_PF, LM_Inserts_l9_Parts)
Dim BG_All: BG_All=Array(BM_B_Caps_Bot, BM_B_Caps_Top, BM_BP_Empty, BM_BP_Jungle, BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_Blades, BM_LEMK, BM_LEMK1, BM_LFlipper, BM_LFlipper1, BM_LFlipper1U, BM_LFlipperU, BM_LSling1, BM_LSling2, BM_Layer_1, BM_Layer_2, BM_Layer_3, BM_MSling1, BM_MSling2, BM_PF, BM_Parts, BM_Post_L, BM_Post_R, BM_R_Green, BM_R_Red, BM_R_Yellow, BM_REMK, BM_RFlipper, BM_RFlipperU, BM_RSling1, BM_RSling2, BM_Trap, BM_gate1, BM_gate2, BM_sw16, BM_sw17, BM_sw18, BM_sw19, BM_sw20, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw33, BM_sw34, BM_sw35, BM_sw36, BM_sw48, BM_sw49, BM_sw51, BM_sw52, BM_sw53, BM_sw54, BM_sw55, BM_sw56, BM_sw57, BM_sw58, BM_sw59, BM_sw60, LM_Flashers_f1_B_Caps_Bot, LM_Flashers_f1_B_Caps_Top, LM_Flashers_f1_BP_Empty, LM_Flashers_f1_BP_Jungle, LM_Flashers_f1_BR1, LM_Flashers_f1_BR2, LM_Flashers_f1_BR3, LM_Flashers_f1_BS1, LM_Flashers_f1_BS2, LM_Flashers_f1_BS3, LM_Flashers_f1_Blades, LM_Flashers_f1_LEMK1, LM_Flashers_f1_LFlipper1, _
	LM_Flashers_f1_LFlipper1U, LM_Flashers_f1_Layer_1, LM_Flashers_f1_Layer_2, LM_Flashers_f1_Layer_3, LM_Flashers_f1_MSling1, LM_Flashers_f1_MSling2, LM_Flashers_f1_PF, LM_Flashers_f1_Parts, LM_Flashers_f1_R_Green, LM_Flashers_f1_R_Red, LM_Flashers_f1_R_Yellow, LM_Flashers_f1_gate1, LM_Flashers_f1_gate2, LM_Flashers_f1_sw17, LM_Flashers_f1_sw18, LM_Flashers_f1_sw19, LM_Flashers_f1_sw20, LM_Flashers_f1_sw21, LM_Flashers_f1_sw22, LM_Flashers_f1_sw23, LM_Flashers_f1_sw48, LM_Flashers_f1_sw51, LM_Flashers_f1_sw58, LM_Flashers_f1_sw59, LM_Flashers_f1_sw60, LM_Flashers_f2_Blades, LM_Flashers_f2_LFlipper, LM_Flashers_f2_LFlipper1, LM_Flashers_f2_LFlipper1U, LM_Flashers_f2_LFlipperU, LM_Flashers_f2_LSling1, LM_Flashers_f2_LSling2, LM_Flashers_f2_Layer_1, LM_Flashers_f2_Layer_2, LM_Flashers_f2_PF, LM_Flashers_f2_Parts, LM_Flashers_f2_Post_L, LM_Flashers_f2_Post_R, LM_Flashers_f2_R_Green, LM_Flashers_f2_R_Red, LM_Flashers_f2_REMK, LM_Flashers_f2_RFlipper, LM_Flashers_f2_RFlipperU, _
	LM_Flashers_f2_sw33, LM_Flashers_f2_sw34, LM_Flashers_f2_sw35, LM_Flashers_f2_sw36, LM_Flashers_f2_sw54, LM_Flashers_f2_sw57, LM_Flashers_f2_sw59, LM_Flashers_f3_Blades, LM_Flashers_f3_LEMK, LM_Flashers_f3_LFlipper, LM_Flashers_f3_LFlipper1, LM_Flashers_f3_LFlipper1U, LM_Flashers_f3_LFlipperU, LM_Flashers_f3_LSling1, LM_Flashers_f3_Layer_1, LM_Flashers_f3_Layer_2, LM_Flashers_f3_MSling1, LM_Flashers_f3_PF, LM_Flashers_f3_Parts, LM_Flashers_f3_Post_L, LM_Flashers_f3_R_Green, LM_Flashers_f3_R_Yellow, LM_Flashers_f3_RFlipper, LM_Flashers_f3_RFlipperU, LM_Flashers_f3_RSling1, LM_Flashers_f3_RSling2, LM_Flashers_f3_sw24, LM_Flashers_f3_sw33, LM_Flashers_f3_sw34, LM_Flashers_f3_sw35, LM_Flashers_f3_sw36, LM_Flashers_f3_sw51, LM_Flashers_f3_sw53, LM_Flashers_f3_sw54, LM_Flashers_f3_sw55, LM_Flashers_f3_sw56, LM_Flashers_f3_sw57, LM_Flashers_f3_sw59, LM_Flashers_f4_B_Caps_Bot, LM_Flashers_f4_B_Caps_Top, LM_Flashers_f4_BP_Empty, LM_Flashers_f4_BP_Jungle, LM_Flashers_f4_BR1, _
	LM_Flashers_f4_BR3, LM_Flashers_f4_BS1, LM_Flashers_f4_BS2, LM_Flashers_f4_BS3, LM_Flashers_f4_Blades, LM_Flashers_f4_LEMK1, LM_Flashers_f4_Layer_1, LM_Flashers_f4_Layer_2, LM_Flashers_f4_MSling1, LM_Flashers_f4_MSling2, LM_Flashers_f4_PF, LM_Flashers_f4_Parts, LM_Flashers_f4_R_Green, LM_Flashers_f4_R_Yellow, LM_Flashers_f4_sw17, LM_Flashers_f4_sw18, LM_Flashers_f4_sw19, LM_Flashers_f4_sw20, LM_Flashers_f4_sw60, LM_Flashers_f5_Blades, LM_Flashers_f5_PF, LM_Flashers_f5_Parts, LM_Flashers_f5_R_Red, LM_Flashers_f5_R_Yellow, LM_Flashers_f5_sw34, LM_Flashers_f5_sw35, LM_Flashers_f5_sw36, LM_Flashers_f5_sw49, LM_Flashers_f5_sw57, LM_Flashers_f5_sw59, LM_Flashers_f6_BS3, LM_Flashers_f6_Blades, LM_Flashers_f6_LFlipper1, LM_Flashers_f6_LFlipper1U, LM_Flashers_f6_PF, LM_Flashers_f6_Parts, LM_Flashers_f6_R_Green, LM_Flashers_f6_R_Yellow, LM_Flashers_f6_sw33, LM_Flashers_f6_sw51, LM_Flashers_f6_sw59, LM_Flashers_f7_BS1, LM_Flashers_f7_BS2, LM_Flashers_f7_BS3, LM_Flashers_f7_Layer_1, LM_Flashers_f7_PF, _
	LM_Flashers_f7_Parts, LM_Flashers_f7_R_Red, LM_Flashers_f7_R_Yellow, LM_Flashers_f7_sw49, LM_Flashers_f8_B_Caps_Bot, LM_Flashers_f8_B_Caps_Top, LM_Flashers_f8_BR1, LM_Flashers_f8_BR3, LM_Flashers_f8_BS1, LM_Flashers_f8_BS3, LM_Flashers_f8_Blades, LM_Flashers_f8_LEMK1, LM_Flashers_f8_LFlipper1U, LM_Flashers_f8_Layer_1, LM_Flashers_f8_MSling1, LM_Flashers_f8_PF, LM_Flashers_f8_Parts, LM_Flashers_f8_R_Green, LM_Flashers_f8_R_Red, LM_Flashers_f8_R_Yellow, LM_Flashers_f8_sw17, LM_Flashers_f8_sw18, LM_Flashers_f8_sw19, LM_Flashers_f8_sw20, LM_Flashers_f8_sw51, LM_Flashers_f8_sw52, LM_GI_l141_B_Caps_Bot, LM_GI_l141_B_Caps_Top, LM_GI_l141_BP_Empty, LM_GI_l141_BP_Jungle, LM_GI_l141_BR1, LM_GI_l141_BR2, LM_GI_l141_BR3, LM_GI_l141_BS1, LM_GI_l141_BS2, LM_GI_l141_BS3, LM_GI_l141_Blades, LM_GI_l141_LEMK1, LM_GI_l141_LFlipper, LM_GI_l141_LFlipper1, LM_GI_l141_LFlipper1U, LM_GI_l141_LFlipperU, LM_GI_l141_Layer_1, LM_GI_l141_Layer_2, LM_GI_l141_Layer_3, LM_GI_l141_MSling1, LM_GI_l141_MSling2, LM_GI_l141_PF, LM_GI_l141_Parts, _
	LM_GI_l141_Post_R, LM_GI_l141_R_Green, LM_GI_l141_R_Red, LM_GI_l141_R_Yellow, LM_GI_l141_gate1, LM_GI_l141_gate2, LM_GI_l141_sw17, LM_GI_l141_sw18, LM_GI_l141_sw19, LM_GI_l141_sw20, LM_GI_l141_sw21, LM_GI_l141_sw23, LM_GI_l141_sw33, LM_GI_l141_sw34, LM_GI_l141_sw35, LM_GI_l141_sw36, LM_GI_l141_sw48, LM_GI_l141_sw49, LM_GI_l141_sw57, LM_GI_l141_sw59, LM_GI_l141_sw60, LM_GI_l141a_Blades, LM_GI_l141a_LFlipper1U, LM_GI_l141a_PF, LM_GI_l141a_Parts, LM_GI_l141a_R_Green, LM_GI_l141a_sw33, LM_GI_l141a_sw34, LM_GI_l141a_sw59, LM_GI_l141b_Blades, LM_GI_l141b_PF, LM_GI_l141b_Parts, LM_GI_l141b_Post_R, LM_GI_l141b_RFlipperU, LM_GI_l141b_sw35, LM_GI_l141b_sw36, LM_GI_l141b_sw55, LM_GI_l141b_sw56, LM_GI_l141b_sw57, LM_GI_l141c_Blades, LM_GI_l141c_PF, LM_GI_l141c_Parts, LM_GI_l141c_R_Red, LM_GI_l141d_B_Caps_Bot, LM_GI_l141d_B_Caps_Top, LM_GI_l141d_BP_Empty, LM_GI_l141d_BP_Jungle, LM_GI_l141d_BR2, LM_GI_l141d_BS1, LM_GI_l141d_BS2, LM_GI_l141d_BS3, LM_GI_l141d_Layer_1, LM_GI_l141d_PF, LM_GI_l141d_Parts, LM_GI_l141d_gate1, _
	LM_GI_l141d_gate2, LM_GI_l141d_sw22, LM_GI_l141d_sw23, LM_GI_l141d_sw48, LM_GI_l141e_B_Caps_Bot, LM_GI_l141e_B_Caps_Top, LM_GI_l141e_BP_Empty, LM_GI_l141e_BP_Jungle, LM_GI_l141e_BR2, LM_GI_l141e_BR3, LM_GI_l141e_BS2, LM_GI_l141e_BS3, LM_GI_l141e_Layer_1, LM_GI_l141e_PF, LM_GI_l141e_Parts, LM_GI_l141e_R_Red, LM_GI_l141e_R_Yellow, LM_GI_l141e_gate1, LM_GI_l141e_gate2, LM_GI_l141e_sw21, LM_GI_l141e_sw22, LM_GI_l141e_sw23, LM_GI_l141f_B_Caps_Bot, LM_GI_l141f_B_Caps_Top, LM_GI_l141f_BP_Empty, LM_GI_l141f_BP_Jungle, LM_GI_l141f_BR2, LM_GI_l141f_BR3, LM_GI_l141f_BS2, LM_GI_l141f_BS3, LM_GI_l141f_Layer_1, LM_GI_l141f_Layer_2, LM_GI_l141f_PF, LM_GI_l141f_Parts, LM_GI_l141f_R_Yellow, LM_GI_l141f_gate1, LM_GI_l141f_sw18, LM_GI_l141f_sw19, LM_GI_l141f_sw21, LM_GI_l141f_sw22, LM_GI_l141f_sw23, LM_GI_l141g_B_Caps_Bot, LM_GI_l141g_B_Caps_Top, LM_GI_l141g_BP_Empty, LM_GI_l141g_BP_Jungle, LM_GI_l141g_BS3, LM_GI_l141g_Layer_1, LM_GI_l141g_MSling1, LM_GI_l141g_MSling2, LM_GI_l141g_PF, LM_GI_l141g_Parts, LM_GI_l141g_R_Green, _
	LM_GI_l141g_R_Yellow, LM_GI_l141g_sw17, LM_GI_l141g_sw18, LM_GI_l141g_sw19, LM_GI_l141g_sw20, LM_GI_l141g_sw21, LM_GI_l141g_sw22, LM_GI_l141h_BP_Empty, LM_GI_l141h_BP_Jungle, LM_GI_l141h_Blades, LM_GI_l141h_PF, LM_GI_l141h_Parts, LM_GI_l141h_R_Red, LM_GI_l141h_gate2, LM_GI_l141h_sw58, LM_GI_l141i_B_Caps_Top, LM_GI_l141i_BP_Empty, LM_GI_l141i_BP_Jungle, LM_GI_l141i_Blades, LM_GI_l141i_LFlipper1U, LM_GI_l141i_Layer_1, LM_GI_l141i_Layer_3, LM_GI_l141i_PF, LM_GI_l141i_Parts, LM_GI_l141i_R_Green, LM_GI_l141i_R_Yellow, LM_GI_l141i_sw60, LM_GI_l141j_Blades, LM_GI_l141j_LFlipper1, LM_GI_l141j_LFlipper1U, LM_GI_l141j_PF, LM_GI_l141j_Parts, LM_GI_l141j_R_Green, LM_GI_l141j_R_Yellow, LM_GI_l141j_sw59, LM_GI_l141k_B_Caps_Top, LM_GI_l141k_BP_Empty, LM_GI_l141k_BP_Jungle, LM_GI_l141k_Blades, LM_GI_l141k_Layer_1, LM_GI_l141k_Layer_2, LM_GI_l141k_Layer_3, LM_GI_l141k_MSling1, LM_GI_l141k_PF, LM_GI_l141k_Parts, LM_GI_l141k_R_Green, LM_GI_l141k_R_Yellow, LM_GI_l141k_sw17, LM_GI_l141k_sw18, LM_GI_l141k_sw19, LM_GI_l141k_sw20, _
	LM_GI_l141k_sw60, LM_GI_l141l_B_Caps_Top, LM_GI_l141l_BP_Empty, LM_GI_l141l_BP_Jungle, LM_GI_l141l_BS2, LM_GI_l141l_BS3, LM_GI_l141l_Layer_1, LM_GI_l141l_PF, LM_GI_l141l_Parts, LM_GI_l141l_R_Yellow, LM_GI_l141l_gate1, LM_GI_l141l_gate2, LM_GI0_l140a_Blades, LM_GI0_l140a_LFlipper, LM_GI0_l140a_LFlipperU, LM_GI0_l140a_Layer_1, LM_GI0_l140a_Layer_2, LM_GI0_l140a_PF, LM_GI0_l140a_Parts, LM_GI0_l140a_REMK, LM_GI0_l140a_RFlipper, LM_GI0_l140a_RFlipperU, LM_GI0_l140a_sw35, LM_GI0_l140a_sw57, LM_GI0_l140b_Blades, LM_GI0_l140b_LEMK, LM_GI0_l140b_LFlipper, LM_GI0_l140b_LFlipperU, LM_GI0_l140b_Layer_1, LM_GI0_l140b_Layer_2, LM_GI0_l140b_PF, LM_GI0_l140b_Parts, LM_GI0_l140b_Post_L, LM_GI0_l140b_RFlipper, LM_GI0_l140b_RFlipperU, LM_GI0_l140b_sw34, LM_GI0_l140b_sw53, LM_GI0_l140b_sw54, LM_GI0_l140c_Blades, LM_GI0_l140c_LEMK, LM_GI0_l140c_LFlipper, LM_GI0_l140c_LFlipperU, LM_GI0_l140c_LSling1, LM_GI0_l140c_LSling2, LM_GI0_l140c_Layer_1, LM_GI0_l140c_Layer_2, LM_GI0_l140c_PF, LM_GI0_l140c_Parts, LM_GI0_l140c_Post_L, _
	LM_GI0_l140c_RFlipper, LM_GI0_l140c_RFlipperU, LM_GI0_l140c_sw33, LM_GI0_l140c_sw34, LM_GI0_l140c_sw53, LM_GI0_l140c_sw54, LM_GI0_l140d_Blades, LM_GI0_l140d_LEMK, LM_GI0_l140d_LFlipper, LM_GI0_l140d_LFlipperU, LM_GI0_l140d_LSling1, LM_GI0_l140d_LSling2, LM_GI0_l140d_Layer_1, LM_GI0_l140d_Layer_2, LM_GI0_l140d_PF, LM_GI0_l140d_Parts, LM_GI0_l140d_Post_L, LM_GI0_l140d_sw33, LM_GI0_l140d_sw34, LM_GI0_l140d_sw36, LM_GI0_l140d_sw53, LM_GI0_l140d_sw54, LM_GI0_l140d_sw59, LM_GI0_l140e_Blades, LM_GI0_l140e_Layer_1, LM_GI0_l140e_Layer_2, LM_GI0_l140e_PF, LM_GI0_l140e_Parts, LM_GI0_l140e_Post_R, LM_GI0_l140e_REMK, LM_GI0_l140e_RFlipper, LM_GI0_l140e_RFlipperU, LM_GI0_l140e_RSling1, LM_GI0_l140e_RSling2, LM_GI0_l140e_sw35, LM_GI0_l140e_sw36, LM_GI0_l140e_sw55, LM_GI0_l140e_sw56, LM_GI0_l140e_sw57, LM_GI0_l140f_Blades, LM_GI0_l140f_LFlipper, LM_GI0_l140f_LFlipperU, LM_GI0_l140f_Layer_1, LM_GI0_l140f_Layer_2, LM_GI0_l140f_PF, LM_GI0_l140f_Parts, LM_GI0_l140f_REMK, LM_GI0_l140f_RFlipper, LM_GI0_l140f_RFlipperU, _
	LM_GI0_l140f_RSling1, LM_GI0_l140f_RSling2, LM_GI0_l140f_sw35, LM_GI0_l140f_sw36, LM_GI0_l140f_sw55, LM_GI0_l140f_sw56, LM_GI0_l140f_sw57, LM_Inserts_l1_PF, LM_Inserts_l1_Parts, LM_Inserts_l10_PF, LM_Inserts_l10_Parts, LM_Inserts_l11_LFlipper1, LM_Inserts_l11_LFlipper1U, LM_Inserts_l11_PF, LM_Inserts_l11_Parts, LM_Inserts_l11_R_Green, LM_Inserts_l11_sw33, LM_Inserts_l11_sw34, LM_Inserts_l11_sw59, LM_Inserts_l12_PF, LM_Inserts_l12_Parts, LM_Inserts_l13_PF, LM_Inserts_l13_Parts, LM_Inserts_l14_PF, LM_Inserts_l14_Parts, LM_Inserts_l15_PF, LM_Inserts_l15_Parts, LM_Inserts_l16_PF, LM_Inserts_l16_Parts, LM_Inserts_l17_Layer_1, LM_Inserts_l17_PF, LM_Inserts_l17_Parts, LM_Inserts_l17_Post_L, LM_Inserts_l18_Layer_1, LM_Inserts_l18_PF, LM_Inserts_l18_Parts, LM_Inserts_l18_Post_L, LM_Inserts_l19_PF, LM_Inserts_l19_Parts, LM_Inserts_l2_PF, LM_Inserts_l2_Parts, LM_Inserts_l2_sw33, LM_Inserts_l2_sw34, LM_Inserts_l2_sw59, LM_Inserts_l20_PF, LM_Inserts_l20_Parts, LM_Inserts_l21_PF, LM_Inserts_l21_Parts, LM_Inserts_l22_PF, _
	LM_Inserts_l22_Parts, LM_Inserts_l23_PF, LM_Inserts_l23_Parts, LM_Inserts_l24_PF, LM_Inserts_l24_Parts, LM_Inserts_l25_Layer_1, LM_Inserts_l25_PF, LM_Inserts_l25_Parts, LM_Inserts_l26_Layer_1, LM_Inserts_l26_PF, LM_Inserts_l26_Parts, LM_Inserts_l26_Post_R, LM_Inserts_l27_LFlipper1U, LM_Inserts_l27_PF, LM_Inserts_l27_Parts, LM_Inserts_l27_R_Yellow, LM_Inserts_l28_PF, LM_Inserts_l28_Parts, LM_Inserts_l29_BP_Jungle, LM_Inserts_l29_Layer_1, LM_Inserts_l29_MSling1, LM_Inserts_l29_MSling2, LM_Inserts_l29_PF, LM_Inserts_l29_Parts, LM_Inserts_l29_R_Yellow, LM_Inserts_l29_sw17, LM_Inserts_l29_sw18, LM_Inserts_l29_sw19, LM_Inserts_l29_sw20, LM_Inserts_l3_PF, LM_Inserts_l3_Parts, LM_Inserts_l3_R_Yellow, LM_Inserts_l3_sw33, LM_Inserts_l3_sw34, LM_Inserts_l3_sw59, LM_Inserts_l30_Layer_1, LM_Inserts_l30_MSling1, LM_Inserts_l30_MSling2, LM_Inserts_l30_PF, LM_Inserts_l30_Parts, LM_Inserts_l30_R_Yellow, LM_Inserts_l30_sw17, LM_Inserts_l30_sw18, LM_Inserts_l30_sw19, LM_Inserts_l30_sw20, LM_Inserts_l31_B_Caps_Bot, _
	LM_Inserts_l31_B_Caps_Top, LM_Inserts_l31_PF, LM_Inserts_l31_Parts, LM_Inserts_l31_R_Green, LM_Inserts_l31_R_Yellow, LM_Inserts_l31_sw17, LM_Inserts_l31_sw18, LM_Inserts_l31_sw19, LM_Inserts_l31_sw20, LM_Inserts_l32_Parts, LM_Inserts_l32_R_Yellow, LM_Inserts_l32_sw20, LM_Inserts_l33_B_Caps_Bot, LM_Inserts_l33_PF, LM_Inserts_l33_Parts, LM_Inserts_l33_R_Green, LM_Inserts_l33_R_Yellow, LM_Inserts_l33_sw18, LM_Inserts_l33_sw19, LM_Inserts_l33_sw20, LM_Inserts_l33_sw35, LM_Inserts_l34_PF, LM_Inserts_l34_Parts, LM_Inserts_l35_PF, LM_Inserts_l35_Parts, LM_Inserts_l36_PF, LM_Inserts_l36_Parts, LM_Inserts_l37_LFlipperU, LM_Inserts_l37_PF, LM_Inserts_l37_Parts, LM_Inserts_l37_RFlipperU, LM_Inserts_l38_Parts, LM_Inserts_l38_R_Yellow, LM_Inserts_l39_PF, LM_Inserts_l39_Parts, LM_Inserts_l39_R_Yellow, LM_Inserts_l39_sw51, LM_Inserts_l4_PF, LM_Inserts_l4_Parts, LM_Inserts_l4_sw35, LM_Inserts_l4_sw36, LM_Inserts_l4_sw57, LM_Inserts_l40_PF, LM_Inserts_l40_Parts, LM_Inserts_l40_R_Yellow, LM_Inserts_l41_PF, _
	LM_Inserts_l41_Parts, LM_Inserts_l41_R_Yellow, LM_Inserts_l42_PF, LM_Inserts_l42_Parts, LM_Inserts_l43_LFlipper1, LM_Inserts_l43_LFlipper1U, LM_Inserts_l43_PF, LM_Inserts_l43_Parts, LM_Inserts_l43_R_Yellow, LM_Inserts_l44_LFlipper1, LM_Inserts_l44_LFlipper1U, LM_Inserts_l44_PF, LM_Inserts_l44_Parts, LM_Inserts_l44_R_Yellow, LM_Inserts_l45_LFlipper1U, LM_Inserts_l45_PF, LM_Inserts_l45_Parts, LM_Inserts_l45_R_Yellow, LM_Inserts_l46_LFlipper, LM_Inserts_l46_LFlipperU, LM_Inserts_l46_PF, LM_Inserts_l46_Parts, LM_Inserts_l46_RFlipper, LM_Inserts_l46_RFlipperU, LM_Inserts_l47_PF, LM_Inserts_l47_Parts, LM_Inserts_l47_R_Yellow, LM_Inserts_l48_B_Caps_Bot, LM_Inserts_l48_B_Caps_Top, LM_Inserts_l48_PF, LM_Inserts_l48_Parts, LM_Inserts_l48_R_Red, LM_Inserts_l49_PF, LM_Inserts_l49_Parts, LM_Inserts_l49_R_Yellow, LM_Inserts_l5_PF, LM_Inserts_l5_Parts, LM_Inserts_l5_sw35, LM_Inserts_l5_sw36, LM_Inserts_l5_sw57, LM_Inserts_l50_LFlipper1, LM_Inserts_l50_LFlipper1U, LM_Inserts_l50_PF, LM_Inserts_l50_Parts, _
	LM_Inserts_l50_R_Green, LM_Inserts_l50_R_Yellow, LM_Inserts_l51_LFlipper1U, LM_Inserts_l51_PF, LM_Inserts_l51_Parts, LM_Inserts_l51_R_Green, LM_Inserts_l51_R_Yellow, LM_Inserts_l52_LFlipper1U, LM_Inserts_l52_Layer_1, LM_Inserts_l52_PF, LM_Inserts_l52_Parts, LM_Inserts_l52_R_Green, LM_Inserts_l52_R_Yellow, LM_Inserts_l53_Layer_1, LM_Inserts_l53_PF, LM_Inserts_l53_Parts, LM_Inserts_l53_R_Green, LM_Inserts_l53_R_Yellow, LM_Inserts_l54_PF, LM_Inserts_l54_Parts, LM_Inserts_l55_Blades, LM_Inserts_l55_PF, LM_Inserts_l55_Parts, LM_Inserts_l55_R_Yellow, LM_Inserts_l56_LFlipper1, LM_Inserts_l56_LFlipper1U, LM_Inserts_l56_Layer_1, LM_Inserts_l56_PF, LM_Inserts_l56_Parts, LM_Inserts_l56_R_Green, LM_Inserts_l56_R_Yellow, LM_Inserts_l57_BP_Empty, LM_Inserts_l57_Layer_1, LM_Inserts_l57_PF, LM_Inserts_l57_Parts, LM_Inserts_l58_B_Caps_Top, LM_Inserts_l58_BP_Empty, LM_Inserts_l58_BP_Jungle, LM_Inserts_l58_PF, LM_Inserts_l58_Parts, LM_Inserts_l59_B_Caps_Top, LM_Inserts_l59_BP_Empty, LM_Inserts_l59_PF, LM_Inserts_l59_Parts, _
	LM_Inserts_l59_gate2, LM_Inserts_l6_PF, LM_Inserts_l6_Parts, LM_Inserts_l60_B_Caps_Top, LM_Inserts_l60_BP_Empty, LM_Inserts_l60_PF, LM_Inserts_l60_Parts, LM_Inserts_l60_R_Red, LM_Inserts_l60_gate1, LM_Inserts_l60_gate2, LM_Inserts_l61_Layer_1, LM_Inserts_l61_PF, LM_Inserts_l61_Parts, LM_Inserts_l61_sw35, LM_Inserts_l61_sw36, LM_Inserts_l61_sw57, LM_Inserts_l62_Blades, LM_Inserts_l62_PF, LM_Inserts_l62_Parts, LM_Inserts_l7_PF, LM_Inserts_l7_Parts, LM_Inserts_l7_sw57, LM_Inserts_l8_PF, LM_Inserts_l8_Parts, LM_Inserts_l9_PF, LM_Inserts_l9_Parts)
' VLM Arrays - End

'******************************************************
'  Timers
'******************************************************
Sub GameTimer_Timer()
	Cor.Update 						'update ball tracking
	RollingUpdate					'update rolling sounds
	DoSTAnim
	DoDTAnim
	Pincab_Rose_Rod.Y = -65 + (5* PlungerRose.Position) -20
	Pincab_Rose_Rose.Y = -65 + (5* PlungerRose.Position) -20
End Sub

sub hundreds_timer
	CaptiveBallLost
end sub

sub CaptiveBallLost
	if isempty(GNRCaptiveBall) then exit sub
	if not InRect(GNRCaptiveBall.x,GNRCaptiveBall.y,100,215,480,186,377,703,262,733) then
'		debug.print "back in the gage!"
		GNRCaptiveBall.x = 262
		GNRCaptiveBall.y = 386
		GNRCaptiveBall.z = 25
	end if
end sub

Dim Bumpers : Bumpers = Array(Bumper1, Bumper2, Bumper3)

Sub FrameTimer_Timer()
	' If DynamicBallShadowsOn=1 Then DynamicBSUpdate 'update ball shadows

	Options_UpdateDMD

	' Compute light color from baked area brightness/color (completely experimental and mathematically wrong, but not that bad...)
	Dim room, gi0, gi
	Dim s, x, y, z
	Dim r, g, b
	Dim sr, sg, sb
	Dim roomLvl, giLvl, gi0Lvl
	roomLvl = (LightLevel / 100.0)^2.2
	gi0Lvl = l140a.GetInPlayIntensity / l140a.Intensity
	giLvl = l141.GetInPlayIntensity / l141.Intensity
	For s = 0 to UBound(gBOT)
		x = CInt(16.0 * gBOT(s).x / 1081.0)
		y = CInt(32.0 * gBOT(s).y / 2162.0)
		' TODO add blending along z: z = Int(gBot(s).z / 2162.0)
		' TODO blend between 4 samples
		if x < 0 then x = 0
		if x > 15 Then x = 15
		if y < 0 then y = 0
		if y > 31 Then y = 31
		sr = 15
		sg = 15
		sb = 15
		room = TableIrr(x + y * 16) ' Value for room light when ball is on playfield
		r = (room And &hFF0000&) / 65536
		g = (room And &h00FF00&) / 256
		b = (room And &h0000FF&) 
		sr = sr + CInt(roomLvl * r)
		sg = sg + CInt(roomLvl * g)
		sb = sb + CInt(roomLvl * b)
 		gi = TableIrr(16*32*2 + x + y * 16) ' Value for gi light when ball is on playfield
		r = (gi And &hFF0000&) / 65536
		g = (gi And &h00FF00&) / 256
		b = (gi And &h0000FF&) 
		sr = sr + CInt(giLvl * r)
		sg = sg + CInt(giLvl * g)
		sb = sb + CInt(giLvl * b)
		gi0 = TableIrr(16*32*4 + x + y * 16) ' Value for gi0 light when ball is on playfield
		r = (gi0 And &hFF0000&) / 65536
		g = (gi0 And &h00FF00&) / 256
		b = (gi0 And &h0000FF&) 
		sr = sr + CInt(gi0Lvl * r)
		sg = sg + CInt(gi0Lvl * g)
		sb = sb + CInt(gi0Lvl * b)
		if sr < 0 then sr = 0
		if sr > 255 Then sr = 255
		if sg < 0 then sg = 0
		if sg > 255 Then sg = 255
		if sb < 0 then sb = 0
		if sb > 255 Then sb = 255
		gBOT(s).color = RGB(sr, sg, sb)
'		Debug.Print "Ball #" & s & ": " & sr & ", " & sg & ", " & sb & " // " & roomLvl & ", " & giLvl & ", " & gi0Lvl & " // " & hex(room) & " - " & hex(gi) & " - " & hex(gi0)
	Next
	
	For Each s in GIRefl
		s.State = giLvl
	Next
	
	' Animate Bumper switch
	' 83 = Bumper skirt radius + 25 => start pressing
	' 63 = Bumper radius + 25 => fully pressed
	Dim i, nearest
	For i = 0 To 2
		nearest = 10000.
		For s = 0 to UBound(gBOT)
			If gBOT(s).z < 30 Then ' Ball on playfield
				x = Bumpers(i).x - gBOT(s).x
				y = Bumpers(i).y - gBOT(s).y
				b = x * x + y * y
				If b < nearest Then nearest = b
			End If
		Next
		z = 0
		If nearest < 83 * 83 Then
			z = (83.0 - sqr(nearest)) / (83.0 - 63.0)
			z = - z * z * 5.0
			if (z < -5) Then z = -5
		End If
		z = z -5
		If i = 0 And BM_BS3.z <> z Then For Each x in BP_BS3: x.Z = z: Next
		If i = 1 And BM_BS1.z <> z Then For Each x in BP_BS1: x.Z = z: Next
		If i = 2 And BM_BS2.z <> z Then For Each x in BP_BS2: x.Z = z: Next
	Next
End Sub

'******************************************************
' Table init.
'******************************************************
Dim Mag1, Mag2, Mag3, PlungerIM
Dim GNRBall1, GNRBall2, GNRBall3, GNRBall4, GNRBall5, GNRBall6, GNRCaptiveBall, gBOT, allBOT

' Copy/Paste from PinMame source for PWM output settings
Const VPM_MODOUT_DEFAULT              =   0 ' Uses default driver modulated solenoid implementation
Const VPM_MODOUT_PWM_RATIO            =   1 ' pulse ratio over the last integration period, allow (approximated) device emulation by the calling app
Const VPM_MODOUT_BULB_44_6_3V_AC      = 100 ' Incandescent #44/555 Bulb connected to 6.3V, commonly used for GI
Const VPM_MODOUT_BULB_47_6_3V_AC      = 101 ' Incandescent #47 Bulb connected to 6.3V, commonly used for (darker) GI with less heat
Const VPM_MODOUT_BULB_86_6_3V_AC      = 102 ' Incandescent #86 Bulb connected to 6.3V, seldom use: TZ, CFTBL,...
Const VPM_MODOUT_BULB_44_18V_DC_WPC   = 201 ' Incandescent #44/555 Bulb connected to 18V, commonly used for lamp matrix with short strobing
Const VPM_MODOUT_BULB_44_18V_DC_GTS3  = 202 ' Incandescent #44/555 Bulb connected to 18V, commonly used for lamp matrix with short strobing
Const VPM_MODOUT_BULB_44_18V_DC_S11   = 203 ' Incandescent #44/555 Bulb connected to 18V, commonly used for lamp matrix with short strobing
Const VPM_MODOUT_BULB_89_20V_DC_WPC   = 301 ' Incandescent #89/906 Bulb connected to 12V, commonly used for flashers
Const VPM_MODOUT_BULB_89_20V_DC_GTS3  = 302 ' Incandescent #89/906 Bulb connected to 12V, commonly used for flashers
Const VPM_MODOUT_BULB_89_32V_DC_S11   = 303 ' Incandescent #89/906 Bulb connected to 32V, used for flashers on S11 with output strobing
Const VPM_MODOUT_LED                  = 400 ' LED PWM (in fact mostly human eye reaction, since LED are nearly instantaneous)
Const VPM_MODOUT_MOTOR_LINEAR         = 500 ' Linear integration for motor (like Twilight Zone clock). Note: evaluate PWM implementation for mech for this type of devices.

Sub Table1_Init
	Dim x: For Each x in BG_Lightmap: x.Opacity = 0: Next
	For Each x in BL_GI0_l140a: x.Opacity = 100: Next
	For Each x in BL_GI0_l140b: x.Opacity = 100: Next
	For Each x in BL_GI0_l140c: x.Opacity = 100: Next
	For Each x in BL_GI0_l140d: x.Opacity = 100: Next
	For Each x in BL_GI0_l140e: x.Opacity = 100: Next
	For Each x in BL_GI0_l140f: x.Opacity = 100: Next

	vpminit me
	vpmMapLights AllLamps ' Map all lamps to the corresponding ROM output suing the value of TimerInterval of each light object
	With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
		.SplashInfoLine = "Guns'n Roses (Data East 1994)" & vbNewLine & "VPW"
		.HandleKeyboard = 0
		.ShowTitle = 0
		.ShowDMDOnly = 1
		.ShowFrame = 0
		.HandleMechanics = 0
	End With
	On Error Resume Next
	Controller.Run
	If Err Then MsgBox Err.Description
	On Error Goto 0

	'Balls
	Set GNRBall1 = sw14.CreateSizedballWithMass(Ballsize/2,Ballmass)
	Set GNRBall2 = sw13.CreateSizedballWithMass(Ballsize/2,Ballmass)
	Set GNRBall3 = sw12.CreateSizedballWithMass(Ballsize/2,Ballmass)
	Set GNRBall4 = sw11.CreateSizedballWithMass(Ballsize/2,Ballmass)
	Set GNRBall5 = sw10.CreateSizedballWithMass(Ballsize/2,Ballmass)
	Set GNRBall6 = sw9.CreateSizedballWithMass(Ballsize/2,Ballmass)
	Set GNRCaptiveBall = captiveBall.CreateSizedballWithMass(Ballsize/2,Ballmass)
	gBOT = Array(GNRCaptiveBall,GNRBall1,GNRBall2,GNRBall3,GNRBall4,GNRBall5,GNRBall6)
	InitRolling
	Controller.Switch(9) = 1
	Controller.Switch(10) = 1
	Controller.Switch(11) = 1
	Controller.Switch(12) = 1
	Controller.Switch(13) = 1
	Controller.Switch(14) = 1

	'Nudging
	vpmNudge.TiltSwitch=-7
	vpmNudge.Sensitivity=4
	vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot, TopSlingShot)

	'Main Timer init
	PinMAMETimer.Interval = PinMAMEInterval
	PinMAMETimer.Enabled = 1

	vpmTimer.AddTimer 300, "captiveball.kick 180,1 '"
	vpmTimer.AddTimer 310, "captiveball.enabled= 0 '"

	'Magnets
	Set mag1= New cvpmMagnet
	With mag1
		.InitMagnet MagLeft, 16  
		.GrabCenter = False 
		.solenoid=51
		.CreateEvents "mag1"
	End With

	Set mag2= New cvpmMagnet
	With mag2
		.InitMagnet MagRight, 16  
		.GrabCenter = False 
		.solenoid=53
		.CreateEvents "mag2"
	End With

	Set mag3= New cvpmMagnet
	With mag3
		.InitMagnet MagMid, 16  
		.GrabCenter = False 
		.solenoid=52
		.CreateEvents "mag3"
	End With

	'Plunger.Pullback
	KickBack.Pullback

	SolTrapDoor 0

	' Flasher PWM and bulb fading (must be made after controller is started by calling Run)
	On Error Resume Next ' Do not fail if using an older version of VPinMame not supporting PWM integration
	With Controller
		.SolMask(1025) = VPM_MODOUT_BULB_89_32V_DC_S11
		.SolMask(1026) = VPM_MODOUT_BULB_89_32V_DC_S11
		.SolMask(1027) = VPM_MODOUT_BULB_89_32V_DC_S11
		.SolMask(1028) = VPM_MODOUT_BULB_89_32V_DC_S11
		.SolMask(1029) = VPM_MODOUT_BULB_89_32V_DC_S11
		.SolMask(1030) = VPM_MODOUT_BULB_89_32V_DC_S11
		.SolMask(1031) = VPM_MODOUT_BULB_89_32V_DC_S11
		.SolMask(1032) = VPM_MODOUT_BULB_89_32V_DC_S11
	End With
	On Error Goto 0
	
'	Controller.SolMask(0) = 0
'	vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
'	Controller.Run

	Options_Load
	UpdateMods
	SetRoomBrightness LightLevel
End Sub

Sub Table_Paused:Controller.Pause = 1:End Sub
Sub Table_unPaused:Controller.Pause = 0:End Sub
Sub Table_exit()
	Controller.Pause = False
	Controller.Stop
End Sub


'******************************************************
'	Keys
'******************************************************

Sub Table1_KeyDown(ByVal keycode)
	If keycode = LeftFlipperKey Then PinCab_LeftFlipperButton.X = PinCab_LeftFlipperButton.X + 10
	If Keycode = RightFlipperKey Then PinCab_RightFlipperButton.X = PinCab_RightFlipperButton.X - 10
	If bInOptions Then
		Options_KeyDown keycode
		Exit Sub
	End If
    If keycode = LeftMagnaSave Then
		If bOptionsMagna Then Options_Open() Else bOptionsMagna = True
    ElseIf keycode = RightMagnaSave Then
		If bOptionsMagna Then Options_Open() Else bOptionsMagna = True
	End If
	If keycode = LeftFlipperKey Then 
		FlipperActivate LeftFlipper, LFPress
		If StagedFlipperMod <> 1 Then 
			FlipperActivate LeftFlipper1, LFPress1
			SolULFlipper true
		End If
	End If
	If StagedFlipperMod = 1 Then
		If keycode = KeyUpperLeft Then
			FlipperActivate LeftFlipper1, LFPress1
			SolULFlipper true
		End If
	End If
	If Keycode = RightFlipperKey Then
		FlipperActivate RightFlipper, RFPress
	End If
	If keycode = PlungerKey Then 
		if Not BIPL then	'sw16
			PlungerRose.Pullback
			SoundPlungerPull 'Pulls gun trigger and pulls rose plunger.
		end if
		vpmTimer.PulseSw 62
	end if
	If keycode = LeftTiltKey Then Nudge 90, 1 : SoundNudgeLeft
	If keycode = RightTiltKey Then Nudge 270, 1 : SoundNudgeRight
	If keycode = CenterTiltKey Then Nudge 0, 1 : SoundNudgeCenter
	If keycode = StartGameKey Then SoundStartButton
	If keycode = AddCreditKey or keycode = AddCreditKey2 Then SoundCoinIn
	If vpmKeyDown(keycode) Then Exit Sub   
End Sub
 
Sub Table1_KeyUp(ByVal keycode)
	If keycode = LeftFlipperKey Then PinCab_LeftFlipperButton.X = PinCab_LeftFlipperButton.X - 10
	If Keycode = RightFlipperKey Then PinCab_RightFlipperButton.X = PinCab_RightFlipperButton.X + 10
    If keycode = LeftMagnaSave And Not bInOptions Then bOptionsMagna = False
    If keycode = RightMagnaSave And Not bInOptions Then bOptionsMagna = False
	If keycode = LeftFlipperKey Then 
		FlipperDeActivate LeftFlipper, LFPress
		If StagedFlipperMod <> 1 Then 
			FlipperDeActivate LeftFlipper1, LFPress1
			SolULFlipper false
		End If
	End If
	If StagedFlipperMod = 1 Then
		If keycode = KeyUpperLeft Then
			FlipperDeActivate LeftFlipper1, LFPress1
			SolULFlipper false
		End If
	End If
	If Keycode = RightFlipperKey Then
		FlipperDeActivate RightFlipper, RFPress
	End If
	If Keycode = StartGameKey Then Controller.Switch(16) = 0
	If keycode = PlungerKey Then
		if Not BIPL then 'sw16
			PlungerRose.Fire
			if BIRP then	'sw24
				SoundPlungerReleaseBallRose
			Else
				SoundPlungerReleaseNoBallRose
			end if
		else
'			SoundPlungerReleaseBall	'not needed as sound is done in kicker
		end if
	end if
	If vpmKeyUp(keycode) Then Exit Sub
End Sub


'******************************************************
'	Solenoids & Flashers
'******************************************************

SolCallback(1) = "SolTrough"
SolCallback(2) = "SolRelease"
SolCallback(3) = "SolAutoFire"
SolCallback(4) = "KickerUpperLeft"
SolCallback(5) = "KickerUpperRight"
SolCallback(6) = "ScoopKicker"
SolCallback(7) = "SolTrapDoor"
SolCallback(8) = "vpmSolSound SoundFX(""knocker"",DOFKnocker),"
SolCallback(9) = "ResetDropsR"
SolCallback(10) = "K1Relay"
SolCallback(11) = "GIRelay"
SolCallback(12) = "ResetDropsL"
SolCallback(14) = "SolKickBack"

SolCallback(23) = "SolEnableFlips"	'Used to enable/disable upper flip


' Data East hardware has 8 flashers which are switched from normal solenoids output using a left/right relay. Pinmame outputs them on 25-32 solenoids
SolModCallback(25) = "FlashPWM 1, f01, BL_Flashers_f1, "
SolModCallback(26) = "FlashPWM 2, f02, BL_Flashers_f2, "
SolModCallback(27) = "FlashPWM 3, f03, BL_Flashers_f3, "
SolModCallback(28) = "FlashPWM 4, f04, BL_Flashers_f4, "
SolModCallback(29) = "FlashPWM 5, f05, BL_Flashers_f5, "
SolModCallback(30) = "FlashPWM 6, f06, BL_Flashers_f6, "
SolModCallback(31) = "FlashPWM 7, f07, BL_Flashers_f7, "
SolModCallback(32) = "FlashPWM 8, f08, BL_Flashers_f8, "

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"


if not ForceFastFlips then
	SolCallback(47) = "SolULFlipper"                        				'Upper Left Flipper
end if

Sub FlashPWM(idx, light, lightmaps, pwm)
	Dim p : p = pwm / 255.0
	light.State = p
	dim r, g, b
	if p < 0 then p = 0
	if p > 1 then p = 1
	r = 1.0
	g = p^0.15
	b = (p - 0.04) / 0.96
	if b < 0 then b = 0
	Dim i : i = .2126 * r^2.2 + .7152 * g^2.2 + .0722 * b^2.2
	i = 100.0 * p / i
	Dim v : v = RGB(Int(255 * r), Int(255 * g), Int(255 * b))
	' Debug.print light.name & " = " & pwm & " => " & i & "/" & r & ", " & g & ", " & b
	dim x : For Each x in lightmaps : x.color = v : x.opacity = i : Next

	' For ball reflections
	If idx = 4 Then f24a.State = p : f24b.State = p
End Sub

Sub K1Relay(Enabled)	'This is solenoid mux relay. Added only for SSF.
	Dim v
	If Enabled Then v = 0 Else v = 1
	Sound_K1_Relay v, K1RelayPosition
End Sub

Sub GIRelay(Enabled)
	Dim v, vrlight
	If Enabled Then v = 0 Else v = 1
	Sound_GI_Relay v, GIRelayPosition
	l141.State = v : l141b.State = v : l141c.State = v : l141d.State = v : l141e.State = v : l141f.State = v
	l141g.State = v : l141h.State = v : l141i.State = v : l141j.State = v : l141k.State = v : l141l.State = v
	lgi.State = v ' For ball reflections
End Sub

'******************************************************
'	TROUGH
'******************************************************

Sub sw14_Hit   : Controller.Switch(14) = 1 : UpdateTrough : End Sub
Sub sw14_UnHit : Controller.Switch(14) = 0 : UpdateTrough : End Sub
Sub sw13_Hit   : Controller.Switch(13) = 1 : UpdateTrough : End Sub
Sub sw13_UnHit : Controller.Switch(13) = 0 : UpdateTrough : End Sub
Sub sw12_Hit   : Controller.Switch(12) = 1 : UpdateTrough : End Sub
Sub sw12_UnHit : Controller.Switch(12) = 0 : UpdateTrough : End Sub
Sub sw11_Hit   : Controller.Switch(11) = 1 : UpdateTrough : End Sub
Sub sw11_UnHit : Controller.Switch(11) = 0 : UpdateTrough : End Sub
Sub sw10_Hit   : Controller.Switch(10) = 1 : UpdateTrough : End Sub
Sub sw10_UnHit : Controller.Switch(10) = 0 : UpdateTrough : End Sub
Sub sw9_Hit    : Controller.Switch(9)  = 1 : UpdateTrough : RandomSoundDrain sw9 : End Sub
Sub sw9_UnHit  : Controller.Switch(9)  = 0 : UpdateTrough : End Sub

Sub UpdateTrough
	UpdateTroughTimer.Interval = 300
	UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer
	If sw14.BallCntOver = 0 Then sw13.kick 57, 10
	If sw13.BallCntOver = 0 Then sw12.kick 57, 10
	If sw12.BallCntOver = 0 Then sw11.kick 57, 10
	If sw11.BallCntOver = 0 Then sw10.kick 57, 10
	If sw10.BallCntOver = 0 Then sw9.kick 57, 10
	Me.Enabled = 0
End Sub


'******************************************************
'	DRAIN & RELEASE
'******************************************************

Sub SolTrough(enabled)
	If enabled Then 
		sw14.kick 57, 20
		Controller.Switch(15) = 1
	End If
End Sub

Sub SolRelease(enabled)
	If enabled Then 
		sw15.kick 57, 10		
		Controller.Switch(15) = 0
		RandomSoundBallRelease sw15
	End If
End Sub


'******************************************************
'	FLIPPERS
'******************************************************
Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
	If Enabled Then
		LF.Fire
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
		RF.Fire
		If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
			RandomSoundReflipUpRight RightFlipper
		Else 
			SoundFlipperUpAttackRight RightFlipper
			RandomSoundFlipperUpRight RightFlipper
		End If
	Else
		RightFlipper.RotateToStart
		If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
			RandomSoundFlipperDownRight RightFlipper
		End If	
		FlipperRightHitParm = FlipperUpSoundLevel
	End If
End Sub

dim FlipsEnabled : FlipsEnabled = false
sub SolEnableFlips(Enabled)
'	msgbox Enabled
	If Enabled Then
		FlipsEnabled = True
	Else
		FlipsEnabled = False
		SolULFlipper False
	end if
end sub

Sub SolULFlipper(Enabled)
	If Enabled And FlipsEnabled Then
		LeftFlipper1.RotateToEnd
		If leftflipper1.currentangle < leftflipper1.endangle + ReflipAngle Then 
			RandomSoundReflipUpLeft LeftFlipper1
		Else 
			SoundFlipperUpAttackLeft LeftFlipper1
			RandomSoundFlipperUpLeft LeftFlipper1
		End If		
	Else
		LeftFlipper1.RotateToStart
		If LeftFlipper1.currentangle < LeftFlipper1.startAngle - 5 Then
			RandomSoundFlipperDownLeft LeftFlipper1
		End If
		FlipperLeft1HitParm = FlipperUpSoundLevel
	End If
End Sub

'Flipper collide subs
Sub LeftFlipper_Collide(parm)
	CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
	LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
	CheckLiveCatch Activeball, RightFlipper, RFCount, parm
	RightFlipperCollide parm
End Sub

Sub LeftFlipper1_Collide(parm)
	LeftFlipper1Collide parm
End Sub

Function Gray2RGB(gray)
	if gray < 0 then gray = 0
	if gray > 255 Then gray = 255
	Gray2RGB = RGB(gray, gray, gray)
End Function

Sub LeftFlipper_Animate
	Dim lfa : lfa = LeftFlipper.CurrentAngle - 90
    FlipperLSh.RotZ = lfa + 90
	' Darken light from lane bulbs when bats are up
	Dim v, w, u,  LM
	v = 255.0 * (121.0 -  LeftFlipper.CurrentAngle) / (121.0 -  70.0)
	w = Gray2RGB(255.0 - v)
	u = Gray2RGB(v)
	For each LM in BP_LFlipper
		LM.Rotz = lfa
		LM.visible = v < 128.0
		LM.color = w
	Next
	For each LM in BP_LFlipperU
		LM.Rotz = lfa
		LM.visible = v >= 128.0
		LM.color = u
	Next
	BM_LFlipper.color = RGB(255, 255, 255)
	BM_LFlipperU.color = RGB(255, 255, 255)
End Sub

Sub LeftFlipper1_Animate
	Dim lfa1 : lfa1 = LeftFlipper1.CurrentAngle - 90
    FlipperL1Sh.RotZ = lfa1 + 90
	' Darken light from lane bulbs when bats are up
	Dim v, w, u, LM
	v = 255.0 * (155.0 - LeftFlipper1.CurrentAngle) / (155.0 - 105.0)
	w = Gray2RGB(255.0 - v)
	u = Gray2RGB(v)
	For each LM in BP_LFlipper1
		LM.Rotz = lfa1
		LM.visible = v < 128.0
		LM.color = w
	Next
	For each LM in BP_LFlipper1U
		LM.Rotz = lfa1
		LM.visible = v >= 128.0
		LM.color = u
	Next
	BM_LFlipper1.color = RGB(255, 255, 255)
	BM_LFlipper1U.color = RGB(255, 255, 255)
End Sub

Sub RightFlipper_Animate
	Dim rfa : rfa = RightFlipper.CurrentAngle - 90
    FlipperRSh.RotZ = rfa + 90
	' Darken light from lane bulbs when bats are up
	Dim v, w, u, LM
	v = 255.0 * (-121 - RightFlipper.CurrentAngle) / (-121 + 70)
	w = Gray2RGB(255.0 - v)
	u = Gray2RGB(v)
	For each LM in BP_RFlipper
		LM.Rotz = rfa
		LM.visible = v < 128.0
		LM.color = w
	Next
	For each LM in BP_RFlipperU
		LM.Rotz = rfa
		LM.visible = v >= 128.0
		LM.color = u
	Next
	BM_RFlipper.color = RGB(255, 255, 255)
	BM_RFlipperU.color = RGB(255, 255, 255)
End Sub

'Impulse Plunger
Sub SolAutofire(Enabled)
	If Enabled Then
		PlungerIM.AutoFire
		SoundPlungerReleaseBall
	End If
End Sub

Const IMPowerSetting = 65
Const IMTime = 0.6
Set plungerIM = New cvpmImpulseP
With plungerIM
    .InitImpulseP swplunger, IMPowerSetting, IMTime
    .Random 0.5
    .CreateEvents "plungerIM"
End With

'************************************************************
'	Sling Shot Animations
'************************************************************
Dim LStep : LStep = 4 : LeftSlingShot_Timer
Dim RStep : RStep = 4 : RightSlingShot_Timer
Dim TStep : TStep = 4 : TopSlingShot_Timer

Sub RightSlingShot_Slingshot
	RS.VelocityCorrect(Activeball)
	vpmTimer.PulseSw(28)
	RandomSoundSlingshotRight zCol_Rubber_Post044
	RStep = -1 : RightSlingShot_Timer ' Initialize Step to 0
	RightSlingShot.TimerEnabled = 1
	RightSlingShot.TimerInterval = 10
End Sub

Sub RightSlingShot_Timer
	Dim x, x1, x2, y: x1 = True:x2 = False:y = 24
    Select Case RStep
        Case 3:x1 = False:x2 = True:y = 10
        Case 4:x1 = False:x2 = False:y = 0:RightSlingShot.TimerEnabled = 0
    End Select
	For Each x in BP_RSling1: x.Visible = x1: Next
	For Each x in BP_RSling2: x.Visible = x2: Next
	For Each x in BP_REMK: x.transx = y: Next
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	LS.VelocityCorrect(Activeball)
	vpmTimer.PulseSw(29)
	RandomSoundSlingshotLeft zCol_Rubber_Post047
	LStep = -1 : LeftSlingShot_Timer ' Initialize Step to 0
	LeftSlingShot.TimerEnabled = 1
	LeftSlingShot.TimerInterval = 10
End Sub

Sub LeftSlingShot_Timer
	Dim x, x1, x2, y: x1 = True:x2 = False:y = 24
    Select Case LStep
        Case 3:x1 = False:x2 = True:y = 10
        Case 4:x1 = False:x2 = False:y = 0:LeftSlingShot.TimerEnabled = 0
    End Select
	For Each x in BP_LSling1: x.Visible = x1: Next
	For Each x in BP_LSling2: x.Visible = x2: Next
	For Each x in BP_LEMK: x.transx = y: Next
    LStep = LStep + 1
End Sub

Sub TopSlingShot_Slingshot
	TS.VelocityCorrect(Activeball)
	vpmTimer.PulseSw(30)
	RandomSoundSlingshotLeft EndPoint1TS
	TStep = -1 : TopSlingShot_Timer ' Initialize Step to 0
	TopSlingShot.TimerEnabled = 1
	TopSlingShot.TimerInterval = 10
End Sub

Sub TopSlingShot_Timer
	Dim x, x1, x2, y: x1 = True:x2 = False:y = 24
    Select Case TStep
        Case 3:x1 = False:x2 = True:y = 10
        Case 4:x1 = False:x2 = False:y = 0:TopSlingShot.TimerEnabled = 0
    End Select
	For Each x in BP_MSling1: x.Visible = x1: Next
	For Each x in BP_MSling2: x.Visible = x2: Next
	For Each x in BP_LEMK1: x.transy = y: Next
    TStep = TStep + 1
End Sub

'******************************************************
'  SLINGSHOT CORRECTION FUNCTIONS
'******************************************************

dim LS : Set LS = New SlingshotCorrection
dim RS : Set RS = New SlingshotCorrection
dim TS : Set TS = New SlingshotCorrection

InitSlingCorrection

Sub InitSlingCorrection

	LS.Object = LeftSlingshot
	LS.EndPoint1 = EndPoint1LS
	LS.EndPoint2 = EndPoint2LS

	RS.Object = RightSlingshot
	RS.EndPoint1 = EndPoint1RS
	RS.EndPoint2 = EndPoint2RS

	TS.Object = TopSlingshot
	TS.EndPoint1 = EndPoint1TS
	TS.EndPoint2 = EndPoint2TS

	'Slingshot angle corrections (pt, BallPos in %, Angle in deg)
	AddSlingsPt 0, 0.00,	-6
	AddSlingsPt 1, 0.45,	-7
	AddSlingsPt 2, 0.48,	0
	AddSlingsPt 3, 0.52,	0
	AddSlingsPt 4, 0.55,	7
	AddSlingsPt 5, 1.00,	6

End Sub

Sub AddSlingsPt(idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
	dim a : a = Array(LS, RS)
	dim x : for each x in a
		x.addpoint idx, aX, aY
	Next
End Sub

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

'	Public Sub Report()         'debug, reports all coords in tbPL.text
'		If not debugOn then exit sub
'		dim a1, a2 : a1 = ModIn : a2 = ModOut
'		dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
'		TBPout.text = str
'	End Sub


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
'			debug.print " BallPos=" & BallPos &" Angle=" & Angle 
'			debug.print " BEFORE: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely 
			Angle = LinearEnvelope(BallPos, ModIn, ModOut)
			RotVxVy = RotPoint(aBall.Velx,aBall.Vely,Angle)
			If Enabled then aBall.Velx = RotVxVy(0)
			If Enabled then aBall.Vely = RotVxVy(1)
'			debug.print " AFTER: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely 
'			debug.print " " 
		End If
	End Sub

End Class

'************************************************************
'************************* SWITCHES *************************
'************************************************************


'************************* Bumpers **************************
Sub Bumper1_Hit():vpmTimer.PulseSw 25:RandomSoundBumperTop Bumper1:End Sub
Sub Bumper2_Hit():vpmTimer.PulseSw 26:RandomSoundBumperMiddle Bumper2:End Sub
Sub Bumper3_Hit():vpmTimer.PulseSw 27:RandomSoundBumperBottom Bumper3:End Sub

Sub Bumper1_Animate: Dim a, x: a = Bumper1.CurrentRingOffset: For Each x in BP_BR3: x.Z = a: Next: End Sub
Sub Bumper2_Animate: Dim a, x: a = Bumper2.CurrentRingOffset: For Each x in BP_BR1: x.Z = a: Next: End Sub
Sub Bumper3_Animate: Dim a, x: a = Bumper3.CurrentRingOffset: For Each x in BP_BR2: x.Z = a: Next: End Sub


'************************ Rollovers *************************
dim BIPL : BIPL = false
Sub sw16_Hit():Controller.Switch(16) = 1:BIPL = true:End Sub
Sub sw16_UnHit
	'RampTriggerPE will make BIPL false, but we add a failsafe using a timer. Otherwise it may break Rose Plunger.
	sw16.timerinterval = 3000
	sw16.timerenabled = true
	Controller.Switch(16) = 0
End Sub	

Sub sw16_timer
'	debug.print "timer bipl false"
	BIPL = false
	me.timerenabled = false
end sub


Sub sw21_Hit():Controller.Switch(21) = 1:MetalWallCollSounds = true:End Sub
Sub sw21_UnHit:Controller.Switch(21) = 0:End Sub
Sub sw22_Hit():Controller.Switch(22) = 1:MetalWallCollSounds = true:End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub
Sub sw23_Hit():Controller.Switch(23) = 1:MetalWallCollSounds = true:End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub
dim BIRP: BIRP = false
Sub sw24_Hit():Controller.Switch(24) = 1:BIRP = true:End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:BIRP = false:End Sub

Sub sw48_Hit():Controller.Switch(48) = 1:End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

Sub sw53_Hit():leftInlaneSpeedLimit:Controller.Switch(53) = 1:End Sub	'Inlane Rollover Left Return
Sub sw53_UnHit:Controller.Switch(53) = 0:End Sub

Sub sw54_Hit():Controller.Switch(54) = 1:End Sub
Sub sw54_UnHit:Controller.Switch(54) = 0:End Sub
Sub sw55_Hit():Controller.Switch(55) = 1:End Sub
Sub sw55_UnHit:Controller.Switch(55) = 0:End Sub

Sub sw56_Hit():rightInlaneSpeedLimit:Controller.Switch(56) = 1:End Sub	'Inlane Rollover Right Return
Sub sw56_UnHit:Controller.Switch(56) = 0:End Sub

Sub sw58_Hit():Controller.Switch(58) = 1:End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub
Sub sw60_Hit():Controller.Switch(60) = 1:End Sub
Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub

Sub leftInlaneSpeedLimit
	'Wylte's implementation
'    debug.print "Spin in: "& activeball.AngMomZ
'    debug.print "Speed in: "& activeball.vely
	if activeball.vely < 0 then exit sub 							'don't affect upwards movement
    activeball.AngMomZ = -abs(activeball.AngMomZ) * RndNum(3,6)
    If abs(activeball.AngMomZ) > 60 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If abs(activeball.AngMomZ) > 80 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If activeball.AngMomZ > 100 Then activeball.AngMomZ = RndNum(80,100)
    If activeball.AngMomZ < -100 Then activeball.AngMomZ = RndNum(-80,-100)

    if abs(activeball.vely) > 5 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 10 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 15 then activeball.vely = 0.8 * activeball.vely
    if activeball.vely > 16 then activeball.vely = RndNum(14,16)
    if activeball.vely < -16 then activeball.vely = RndNum(-14,-16)
'    debug.print "Spin out: "& activeball.AngMomZ
'    debug.print "Speed out: "& activeball.vely
End Sub


Sub rightInlaneSpeedLimit
	'Wylte's implementation
'    debug.print "Spin in: "& activeball.AngMomZ
'    debug.print "Speed in: "& activeball.vely
	if activeball.vely < 0 then exit sub 							'don't affect upwards movement
    activeball.AngMomZ = abs(activeball.AngMomZ) * RndNum(2,4)
    If abs(activeball.AngMomZ) > 60 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If abs(activeball.AngMomZ) > 80 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If activeball.AngMomZ > 100 Then activeball.AngMomZ = RndNum(80,100)
    If activeball.AngMomZ < -100 Then activeball.AngMomZ = RndNum(-80,-100)

	if abs(activeball.vely) > 5 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 10 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 15 then activeball.vely = 0.8 * activeball.vely
    if activeball.vely > 16 then activeball.vely = RndNum(14,16)
    if activeball.vely < -16 then activeball.vely = RndNum(-14,-16)
'    debug.print "Spin out: "& activeball.AngMomZ
'    debug.print "Speed out: "& activeball.vely
End Sub

Sub sw16_Animate: Dim a, x: a = sw16.CurrentAnimOffset: For Each x in BP_sw16: x.transz = a: Next: End Sub
Sub sw21_Animate: Dim a, x: a = sw21.CurrentAnimOffset: For Each x in BP_sw21: x.transz = a: Next: End Sub
Sub sw22_Animate: Dim a, x: a = sw22.CurrentAnimOffset: For Each x in BP_sw22: x.transz = a: Next: End Sub
Sub sw23_Animate: Dim a, x: a = sw23.CurrentAnimOffset: For Each x in BP_sw23: x.transz = a: Next: End Sub
Sub sw24_Animate: Dim a, x: a = sw24.CurrentAnimOffset: For Each x in BP_sw24: x.transz = a: Next: End Sub
Sub sw48_Animate: Dim a, x: a = sw48.CurrentAnimOffset: For Each x in BP_sw48: x.transz = a: Next: End Sub
Sub sw53_Animate: Dim a, x: a = sw53.CurrentAnimOffset: For Each x in BP_sw53: x.transz = a: Next: End Sub
Sub sw54_Animate: Dim a, x: a = sw54.CurrentAnimOffset: For Each x in BP_sw54: x.transz = a: Next: End Sub
Sub sw55_Animate: Dim a, x: a = sw55.CurrentAnimOffset: For Each x in BP_sw55: x.transz = a: Next: End Sub
Sub sw56_Animate: Dim a, x: a = sw56.CurrentAnimOffset: For Each x in BP_sw56: x.transz = a: Next: End Sub
Sub sw58_Animate: Dim a, x: a = sw58.CurrentAnimOffset: For Each x in BP_sw58: x.transz = a: Next: End Sub
Sub sw60_Animate: Dim a, x: a = sw60.CurrentAnimOffset: For Each x in BP_sw60: x.transz = a: Next: End Sub

Sub Gate5_Animate   : Dim a, x : a = Gate5.CurrentAngle   : For Each x in BP_sw52 : x.RotX = a: Next: End Sub
Sub GateSW49_Animate: Dim a, x : a = GateSW49.CurrentAngle: For Each x in BP_sw49 : x.RotX = a: Next: End Sub
Sub GateSW51_Animate: Dim a, x : a = GateSW51.CurrentAngle: For Each x in BP_sw51 : x.RotX = a: Next: End Sub
Sub Gate1_Animate   : Dim a, x : a = Gate1.CurrentAngle   : For Each x in BP_gate1: x.RotX = a: Next: End Sub
Sub Gate2_Animate   : Dim a, x : a = Gate2.CurrentAngle   : For Each x in BP_gate2: x.RotX = a: Next: End Sub

'********************** Ramp Switches ***********************
Sub sw40_Hit():Controller.Switch(40) = 1:End Sub		'Snake Funnel
Sub sw40_UnHit:Controller.Switch(40) = 0:End Sub

Sub sw49_Hit():Controller.Switch(49) = 1:End Sub		'R Ramp Entrance
Sub sw49_UnHit:Controller.Switch(49) = 0:End Sub

Sub sw50_Hit():Controller.Switch(50) = 1:End Sub		'R Ramp Made
Sub sw50_UnHit:Controller.Switch(50) = 0:End Sub

Sub sw51_Hit():Controller.Switch(51) = 1:End Sub		'G Ramp Entrance
Sub sw51_UnHit:Controller.Switch(51) = 0:End Sub

Sub sw52_Hit():Controller.Switch(52) = 1:End Sub		'G Ramp Made
Sub sw52_UnHit:Controller.Switch(52) = 0:End Sub


'********************* Standup Targets **********************
Sub sw17_hit:STHit 17:End Sub
Sub sw18_hit:STHit 18:End Sub
Sub sw19_hit:STHit 19:End Sub
Sub sw20_hit:STHit 20:End Sub


'********************** Drop Targets ************************
Sub Sw33_Hit:DTHit 33:TargetBouncer Activeball, 1.5:End Sub
Sub Sw34_Hit:DTHit 34:TargetBouncer Activeball, 1.5:End Sub
Sub Sw59_Hit:DTHit 59:TargetBouncer Activeball, 1.5:End Sub

Sub ResetDropsL(enabled)
	if enabled then
		PlaySoundAt SoundFX(DTResetSound,DOFContactors),BM_sw34
		DTRaise 33
		DTRaise 34
		DTRaise 59
	end if
End Sub

Sub Sw35_Hit:DTHit 35:TargetBouncer Activeball, 1.5:End Sub
Sub Sw36_Hit:DTHit 36:TargetBouncer Activeball, 1.5:End Sub
Sub Sw57_Hit:DTHit 57:TargetBouncer Activeball, 1.5:End Sub

Sub ResetDropsR(enabled)
	if enabled then
		PlaySoundAt SoundFX(DTResetSound,DOFContactors),BM_sw35
		DTRaise 35
		DTRaise 36
		DTRaise 57
	end if
End Sub


'************************* VUKs *****************************
Dim KickerBall37, KickerBall38, KickerBall39

Sub KickBall(kball, kangle, kvel, kvelz, kzlift)
	dim rangle
	rangle = PI * (kangle - 90) / 180
    
	kball.z = kball.z + kzlift
	kball.velz = kvelz
	kball.velx = cos(rangle)*kvel
	kball.vely = sin(rangle)*kvel
End Sub

'Upper Left VUK
Sub sw37_Hit
'	msgbox activeball.z
    set KickerBall37 = activeball
    Controller.Switch(37) = 1
    SoundSaucerLock
	sw37.timerinterval=10
	sw37.timerenabled=true
End Sub

Sub sw37_unHit
	Controller.Switch(37) = 0
	sw37.timerenabled=false
end sub

sub sw37_timer
'	debug.print "vuk left: " & KickerBall37.z
	'prevents ball wiggle in saucer
	KickerBall37.x=sw37.x
	KickerBall37.y=sw37.y
	KickerBall37.z=6.5 'a bit higher so it wont touch the collidable bottom
	KickerBall37.angmomx=0
	KickerBall37.angmomy=0
	KickerBall37.angmomz=0
	KickerBall37.velx=0
	KickerBall37.vely=0
	KickerBall37.velz=0
	If Controller.Switch(37) = 0 Then me.timerenabled = false
end sub

Sub KickerUpperLeft(Enable)
	sw37.timerenabled=false
    If Enable then
		If Controller.Switch(37) <> 0 Then
			KickBall KickerBall37, 220, 15, 5, 10
			SoundSaucerKick 1, sw37
'			Controller.Switch(37) = 0 'this may not be safe here
		End If
	End If
End Sub

'Middle Scoop
Sub sw38_Hit
    set KickerBall38 = activeball
    Controller.Switch(38) = 1
    SoundSaucerLock
'	debug.print "scoop hit"
End Sub

Sub ScoopKicker(Enable)
    If Enable then
		If Controller.Switch(38) <> 0 Then
			KickBall KickerBall38, 15, 70, 0, 0
			SoundSaucerKick 1, sw38
			Controller.Switch(38) = 0
		End If
	End If
End Sub

'Upper Right VUK
Sub sw39_Hit
'	msgbox activeball.z
    set KickerBall39 = activeball
    Controller.Switch(39) = 1
    SoundSaucerLock
	sw39.timerinterval=10
	sw39.timerenabled=true
End Sub

Sub sw39_unHit
	Controller.Switch(39) = 0
	sw39.timerenabled=false
end sub

sub sw39_timer
'	debug.print "vuk left: " & KickerBall37.z
	'prevents ball wiggle in saucer
	KickerBall39.x=sw39.x
	KickerBall39.y=sw39.y
	KickerBall39.z=6.5 'a bit higher so it wont touch the collidable bottom
	KickerBall39.angmomx=0
	KickerBall39.angmomy=0
	KickerBall39.angmomz=0
	KickerBall39.velx=0
	KickerBall39.vely=0
	KickerBall39.velz=0
	If Controller.Switch(39) = 0 Then me.timerenabled = false
end sub

Sub KickerUpperRight(Enable)
	sw39.timerenabled=false
    If Enable then
		If Controller.Switch(39) <> 0 Then
			KickBall KickerBall39, 23, 0, 50, 10
			SoundSaucerKick 1, sw39
'			Controller.Switch(39) = 0 'not safe
		End If
	End If
End Sub


'******************* Kicker / Trapdoor **********************
Sub SolKickBack(enabled)
    If enabled Then
		Kickback.Fire
		RandomSoundLockingKickerSolenoid
	Else
		Kickback.PullBack
	End If
End Sub

Sub SolTrapDoor(Enabled)
	If Enabled Then
		SnakeTrapDoor.IsDropped = 1
		TrapDoorOpen.IsDropped = 0
		TrapFlipper.RotateToEnd
	Else
		SnakeTrapDoor.IsDropped = 0
		TrapDoorOpen.IsDropped = 1
		TrapFlipper.RotateToStart
	End If
End Sub

Sub TrapFlipper_Animate
	Dim a, x: a = TrapFlipper.CurrentAngle:	For Each x in BP_Trap: x.RotX = a: Next
End Sub


'******************************************************
'****  BALL ROLLING AND DROP SOUNDS
'******************************************************
'
' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

Const AmbientBallShadowOn   = 1
Const AmbientBSFactor 		= 1   	'0 to 1, higher is darker
Const fovY					= 0		'Offset y position under ball to account for layback or inclination (more pronounced need further back)
Const AmbientMovement		= 2		'1 to 4, higher means more movement as the ball moves left and right
dim objBallShadow: objBallShadow = Array (BallShadow0, BallShadow1, BallShadow2, BallShadow3, BallShadow4, BallShadow5, BallShadow6)
Dim BallShadowA: BallShadowA = Array (BallShadowA0, BallShadowA1, BallShadowA2, BallShadowA3, BallShadowA4, BallShadowA5, BallShadowA6)
Dim DropCount, rolling

Sub InitRolling
	Dim i
	ReDim rolling(UBound(gBOT))
	ReDim DropCount(UBound(gBOT))
	For i = 0 to UBound(gBOT)
		rolling(i) = False
		objBallShadow(i).material = "BallShadow" & i
		UpdateMaterial objBallShadow(i).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
		objBallShadow(i).Z = 1 + i/1000 + 0.04
		objBallShadow(i).visible = 0
		BallShadowA(i).Opacity = 100*AmbientBSFactor
		BallShadowA(i).visible = 0
	Next
End Sub

Sub RollingUpdate()
	Dim b
	For b = 0 to UBound(gBOT)
		If BallVel(gBOT(b)) > 1 AND gBOT(b).z < 30 Then
			rolling(b) = True
			PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))

		Else
			If rolling(b) = True Then
				StopSound("BallRoll_" & b)
				rolling(b) = False
			End If
		End If

		' Ball Drop Sounds
		If gBOT(b).VelZ < -1 and gBOT(b).z < 55 and gBOT(b).z > 27 Then 'height adjust for ball drop sounds
			If DropCount(b) >= 5 Then
				DropCount(b) = 0
				If gBOT(b).velz > -7 Then
					RandomSoundBallBouncePlayfieldSoft gBOT(b)
				Else
					RandomSoundBallBouncePlayfieldHard gBOT(b)
				End If				
			End If
		End If
		If DropCount(b) < 5 Then
			DropCount(b) = DropCount(b) + 1
		End If

		' Ambient Ball Shadows
		If AmbientBallShadowOn = 1 Then			'Primitive shadow on playfield, flasher shadow in ramps
			If gBOT(b).Z > 30 Then							'The flasher follows the ball up ramps while the primitive is on the pf
				If gBOT(b).X < tablewidth/2 Then
					objBallShadow(b).X = ((gBOT(b).X) - (Ballsize/10) + ((gBOT(b).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
				Else
					objBallShadow(b).X = ((gBOT(b).X) + (Ballsize/10) + ((gBOT(b).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
				End If
				objBallShadow(b).Y = gBOT(b).Y + BallSize/10 + fovY
				objBallShadow(b).visible = 1
				objBallShadow(b).visible = 0 ' Disabled since it messes up with the lightmap and would need some more work

				BallShadowA(b).X = gBOT(b).X
				BallShadowA(b).Y = gBOT(b).Y + BallSize/5 + fovY
				BallShadowA(b).height=gBOT(b).z - BallSize/4		'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
				BallShadowA(b).visible = 1
				BallShadowA(b).visible = 0 ' Disabled since it messes up with the lightmap and would need some more work
			Elseif gBOT(b).Z <= 30 And gBOT(b).Z > 20 Then	'On pf, primitive only
				objBallShadow(b).visible = 1
				If gBOT(b).X < tablewidth/2 Then
					objBallShadow(b).X = ((gBOT(b).X) - (Ballsize/10) + ((gBOT(b).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
				Else
					objBallShadow(b).X = ((gBOT(b).X) + (Ballsize/10) + ((gBOT(b).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
				End If
				objBallShadow(b).Y = gBOT(b).Y + fovY
				BallShadowA(b).visible = 0
			Else											'Under pf, no shadows
				objBallShadow(b).visible = 0
				BallShadowA(b).visible = 0
			end if

		Elseif AmbientBallShadowOn = 2 Then		'Flasher shadow everywhere
			If gBOT(b).Z > 30 Then							'In a ramp
				BallShadowA(b).X = gBOT(b).X
				BallShadowA(b).Y = gBOT(b).Y + BallSize/5 + fovY
				BallShadowA(b).height=gBOT(b).z - BallSize/4		'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
				BallShadowA(b).visible = 1
			Elseif gBOT(b).Z <= 30 And gBOT(b).Z > 20 Then	'On pf
				BallShadowA(b).visible = 1
				If gBOT(b).X < tablewidth/2 Then
					BallShadowA(b).X = ((gBOT(b).X) - (Ballsize/10) + ((gBOT(b).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
				Else
					BallShadowA(b).X = ((gBOT(b).X) + (Ballsize/10) + ((gBOT(b).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
				End If
				BallShadowA(b).Y = gBOT(b).Y + Ballsize/10 + fovY
				BallShadowA(b).height=gBOT(b).z - BallSize/2 + 5
			Else											'Under pf
				BallShadowA(b).visible = 0
			End If
		End If
	Next
End Sub


'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************



'******************************************************
'**** RAMP ROLLING SFX
'******************************************************

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
dim RampBalls(5,2)
'x,0 = ball x,1 = ID,	2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(5)	

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

'******************************************************
'Ramp triggers
'******************************************************

'Ramp Snake
Sub RampTriggerS_Hit
	If activeball.vely < 0 Then
		WireRampOn True
		RandomSoundRampFlapUp()
	Else
		WireRampOff
		RandomSoundRampFlapDown()
	End If
End Sub

Sub RampTriggerSE_Hit
    if abs(activeball.AngMomZ) > 70 then activeball.AngMomZ = 50
    activeball.AngMomZ = -abs(activeball.AngMomZ) * 3
    WireRampOff
End Sub

'Ramp G
Sub RampTriggerG_Hit
	If activeball.vely < 0 Then
		WireRampOn True
	Else
		WireRampOff
	End If
End Sub

Sub RampTriggerGM_Hit
	WireRampOff
	WireRampOn False
End Sub

Sub RampTriggerGE_Hit
    if abs(activeball.AngMomZ) > 70 then activeball.AngMomZ = 40
    activeball.AngMomZ = -abs(activeball.AngMomZ) * 3
    WireRampOff
End Sub

'Ramp R
Sub RampTriggerR_Hit
	If activeball.vely < 0 Then
		WireRampOn True
	Else
		WireRampOff
	End If
End Sub

Sub RampTriggerRM_Hit
	WireRampOff
	WireRampOn False
End Sub

Sub RampTriggerRE_Hit
    if abs(activeball.AngMomZ) > 70 then activeball.AngMomZ = 40
    activeball.AngMomZ = -abs(activeball.AngMomZ) * 3
    WireRampOff
End Sub

'Ramp Plunger
Sub RampTriggerP_Hit
	If activeball.vely < 0 Then
		WireRampOn False
	Else
		WireRampOff
	End If
End Sub

'Ramp Plunger
Sub RampTriggerP1_Hit
	WireRampOn False
	PlaySoundAtLevelStatic ("TOM_R_Ramp_3"), RightRampSoundLevel*3, RampTriggerP1	'metal to metal wire
End Sub


Sub RampTriggerPE_Hit
    if abs(activeball.AngMomZ) > 70 then activeball.AngMomZ = 50
    activeball.AngMomZ = -abs(activeball.AngMomZ) * 3
    WireRampOff
	BIPL = False
End Sub



'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************



'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************

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
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperLeft1HitParm, FlipperRightHitParm
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
Dim SaucerLockSoundLevel, SaucerKickSoundLevel, PostSoundLevel

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
PostSoundLevel = 0.8

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
	PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
	PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
	PlaySound playsoundparams, - 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
	PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
	PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
	PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
	PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
	PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
	PlaySound playsoundparams, - 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
	PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
	PlaySound soundname, 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
	PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
	PlaySound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
	PlaySound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
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
	PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, PlungerRose
End Sub

Sub SoundPlungerReleaseBall()
	PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, sw16	
End Sub

Sub SoundPlungerReleaseBallRose()
	PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, PlungerRose	
End Sub

Sub SoundPlungerReleaseNoBall()
	PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, sw16
End Sub

Sub SoundPlungerReleaseNoBallRose()
	PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, sw24
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

Sub LeftFlipper1Collide(parm)
	FlipperLeft1HitParm = parm/10
	If FlipperLeft1HitParm > 1 Then
		FlipperLeft1HitParm = 1
	End If
	FlipperLeft1HitParm = FlipperUpSoundLevel * FlipperLeft1HitParm
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
'these are done later in the script
Sub SoundPlayfieldGate()			
	PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd*2)+1), GateSoundLevel, Activeball
End Sub
'
'Sub SoundHeavyGate()
'	PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, Activeball
'End Sub
'
'Sub Gates_hit(idx)
'	SoundHeavyGate
'End Sub
'
Sub GatesWire_hit(idx)	
	SoundPlayfieldGate
End Sub	

'/////////////////////////////  ORB ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
	PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
	PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub

dim MetalWallCollSounds : MetalWallCollSounds = true
sub LeftOrbEnter_hit
	if activeball.vely < -15 then
'		debug.print "left vely: " & activeball.vely
		RandomSoundLeftArch
		MetalWallCollSounds = false
	else
		MetalWallCollSounds = true
	end if
end sub

sub MiddleOrbEnter_hit
	if activeball.velx > 15 then	'going towards right fast enough
'		debug.print "middle velx: " & activeball.velx
		'hand picked shorter sound		
		PlaySoundAtLevelActiveBall ("Arch_R2"), Vol(ActiveBall) * ArchSoundFactor
		MetalWallCollSounds = false
	else
		MetalWallCollSounds = true
	end if
end sub

sub RightOrbEnter_hit
	if activeball.vely < -15 then
'		debug.print "right vely: " & activeball.vely
		RandomSoundRightArch
		MetalWallCollSounds = false
	else
		MetalWallCollSounds = true
	end if
end sub


sub TopOrb_hit
'	debug.print "top velx: " & activeball.velx
	if activeball.velx < -10 then	'going fast from right to left
		RandomSoundLeftArch
		StopSound "Arch_R1"
		StopSound "Arch_R2"
		StopSound "Arch_R3"
		StopSound "Arch_R4"
		MetalWallCollSounds = false
	else							'going from left to right which is stopped by the gate
		StopSound "Arch_L1"
		StopSound "Arch_L2"
		StopSound "Arch_L3"
		StopSound "Arch_L4"
		MetalWallCollSounds = true
	end if

end sub

sub Wall018_hit		'top wall, not in metals collection anymore
	if MetalWallCollSounds then
		randomsoundmetal
	end if
end sub

sub Wall264_hit
	if MetalWallCollSounds then
'		debug.print "metal sound yes"
		randomsoundmetal
	end if
end sub


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

	FlipperCradleCollision ball1, ball2, velocity
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
Const RelayGISoundLevel = 0.1									'volume level; range [0, 1];
Const RelayK1SoundLevel = 0.45									'volume level; range [0, 1];


Sub Sound_K1_Relay(toggle, obj)
'	debug.print "K1: " & toggle
	Select Case toggle
		Case 1
			PlaySoundAtLevelStatic ("Relay_Flash_On"), RelayK1SoundLevel, obj
		Case 0
			PlaySoundAtLevelStatic ("Relay_Flash_Off"), RelayK1SoundLevel, obj
	End Select
End Sub

Sub Sound_GI_Relay(toggle, obj)
'	debug.print "GI: " & toggle
	Select Case toggle
		Case 1
			PlaySoundAtLevelStatic ("Relay_GI_On"), RelayGISoundLevel, obj
		Case 0
			PlaySoundAtLevelStatic ("Relay_GI_Off"), RelayGISoundLevel, obj
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

'/////////////////////////////  COIN IN  ////////////////////////////

Sub SoundCoinIn
	Select Case Int(rnd*3)
		Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
		Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
		Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
	End Select
End Sub




'/////////////////////////////  RAMP COLLISIONS  ////////////////////////////

dim RRHit1_volume, RRHit2_volume, RRHit3_volume, RRHit4_volume
dim LRHit1_volume, LRHit2_volume, LRHit3_volume, LRHit4_volume

Dim LeftRampSoundLevel
Dim RightRampSoundLevel
Dim RampFallbackSoundLevel

Dim FlapSoundLevel

'///////////////////////-----Ramps-----///////////////////////
'///////////////////////-----Plastic Ramps-----///////////////////////
LeftRampSoundLevel = 0.1												'volume level; range [0, 1]
RightRampSoundLevel = 0.1												'volume level; range [0, 1]
RampFallbackSoundLevel = 0.2											'volume level; range [0, 1]

'///////////////////////-----Ramp Flaps-----///////////////////////
FlapSoundLevel = 0.8													'volume level; range [0, 1]



'/////////////////////////////  PLASTIC RIGHT RAMP SOUNDS  ////////////////////////////
sub RRHit1_Hit()
	RRHit1_volume = RightRampSoundLevel
	PlaySoundAtLevelStatic ("TOM_C_Ramp_1_Improved2"), RRHit1_volume, RRHit1
end Sub

Sub RRHit2_Hit()
	RRHit2_volume = RightRampSoundLevel
	RRHit1.TimerInterval = 5
	RRHit1.TimerEnabled = 1
	PlaySoundAtLevelStatic ("TOM_C_Ramp_2_Improved2"), RRHit2_volume, RRHit2
End Sub

Sub RRHit3_Hit()
	RRHit3_volume = RightRampSoundLevel
	RRHit2.TimerInterval = 5
	RRHit2.TimerEnabled = 1
	PlaySoundAtLevelStatic ("TOM_C_Ramp_3_Improved2"), RRHit3_volume, RRHit3
End Sub

Sub RRHit4_Hit()
	RRHit4_volume = RightRampSoundLevel
	RRHit3.TimerInterval = 5
	RRHit3.TimerEnabled = 1
	PlaySoundAtLevelStatic ("TOM_C_Ramp_4_Improved2"), RRHit4_volume, RRHit4
End Sub


'/////////////////////////////  PLASTIC RIGHT RAMP SOUND TIMERS  ////////////////////////////
Sub RRHit1_Timer()
'	debug.print "1: " & RRHit1_volume
	If RRHit1_volume > 0 Then
		RRHit1_volume = RRHit1_volume - 0.05
		PlaySoundAtLevelExistingStatic ("TOM_C_Ramp_1_Improved2"), RRHit1_volume, RRHit1
	Else
		Me.TimerEnabled = 0
	End If
End Sub

Sub RRHit2_Timer()
'	debug.print "2: " & RRHit2_volume
	If RRHit2_volume > 0 Then
		RRHit2_volume = RRHit2_volume - 0.05
		PlaySoundAtLevelExistingStatic ("TOM_C_Ramp_2_Improved2"), RRHit2_volume, RRHit2
	Else
		Me.TimerEnabled = 0
	End If
End Sub

Sub RRHit3_Timer()
'	debug.print "3: " & RRHit3_volume
	If RRHit3_volume > 0 Then
		RRHit3_volume = RRHit3_volume - 0.05
		PlaySoundAtLevelExistingStatic ("TOM_C_Ramp_3_Improved2"), RRHit3_volume, RRHit3
	Else
		Me.TimerEnabled = 0
	End If
End Sub


'/////////////////////////////  PLASTIC LEFT RAMP SOUNDS  ////////////////////////////

sub LRHit1_Hit()
	LRHit1_volume = LeftRampSoundLevel
	PlaySoundAtLevelStatic ("TOM_C_Ramp_1_Improved2"), LRHit1_volume, LRHit1
end Sub

Sub LRHit2_Hit()
	LRHit2_volume = LeftRampSoundLevel
	LRHit1.TimerInterval = 5
	LRHit1.TimerEnabled = 1
	PlaySoundAtLevelStatic ("TOM_C_Ramp_2_Improved2"), LRHit2_volume, LRHit2
End Sub

Sub LRHit3_Hit()
	LRHit3_volume = LeftRampSoundLevel
	LRHit2.TimerInterval = 5
	LRHit2.TimerEnabled = 1
	PlaySoundAtLevelStatic ("TOM_C_Ramp_3_Improved2"), LRHit3_volume, LRHit3
End Sub

Sub LRHit4_Hit()
	LRHit4_volume = LeftRampSoundLevel
	LRHit3.TimerInterval = 5
	LRHit3.TimerEnabled = 1
	PlaySoundAtLevelStatic ("TOM_R_Ramp_2"), LRHit4_volume, LRHit4	'to metal wire
End Sub

'Plate by the snake ramp
Sub LRHitPlate_Hit()
	PlaySoundAtLevelStatic ("TOM_R_Ramp_1"), LeftRampSoundLevel, LRHitPlate	'metal plate
End Sub

'before snake ramp vortex
Sub SRHit1_Hit()
	PlaySoundAtLevelStatic ("TOM_C_Ramp_2_Improved2"), LeftRampSoundLevel, SRHit1
end sub


'/////////////////////////////  PLASTIC LEFT RAMP SOUND TIMERS ////////////////////////////
Sub LRHit1_Timer()
'	debug.print "1: " & LRHit1_volume
	If LRHit1_volume > 0 Then
		LRHit1_volume = LRHit1_volume - 0.05
		PlaySoundAtLevelExistingStatic ("TOM_C_Ramp_1_Improved2"), LRHit1_volume, LRHit1
	Else
		Me.TimerEnabled = 0
	End If
End Sub

Sub LRHit2_Timer()
'	debug.print "2: " & LRHit2_volume
	If LRHit2_volume > 0 Then
		LRHit2_volume = LRHit2_volume - 0.05
		PlaySoundAtLevelExistingStatic ("TOM_C_Ramp_2_Improved2"), LRHit2_volume, LRHit2
	Else
		Me.TimerEnabled = 0
	End If
End Sub

Sub LRHit3_Timer()
'	debug.print "3: " & LRHit3_volume
	If LRHit3_volume > 0 Then
		LRHit3_volume = LRHit3_volume - 0.05
		PlaySoundAtLevelExistingStatic ("TOM_C_Ramp_3_Improved2"), LRHit3_volume, LRHit3
	Else
		Me.TimerEnabled = 0
	End If
End Sub


'/////////////////////////////  PLASTIC RAMPS FLAPS  ////////////////////////////
'/////////////////////////////  PLASTIC RAMPS FLAPS - EVENTS  ////////////////////////////
Sub RampHelper1_Hit()
'	debug.print "RampTriggerGFlapHit: " & ActiveBall.VelY
	If (ActiveBall.VelY > 0) Then 
		'ball is traveling down the playfield
		RandomSoundRampFlapDown()
		
	ElseIf (ActiveBall.VelY < 0) Then
		'ball is traveling up the playfield
		RandomSoundRampFlapUp()
		
	End If
End Sub

Sub RampHelper2_Hit()
'	debug.print "RampTriggerRFlapHit: " & ActiveBall.VelY
	If (ActiveBall.VelY > 0) Then 
		'ball is traveling down the playfield
		RandomSoundRampFlapDown()
	ElseIf (ActiveBall.VelY < 0) Then
		RandomSoundRampFlapUp()

	End If
End Sub



'/////////////////////////////  PLASTIC RAMPS FLAPS - SOUNDS  ////////////////////////////
Sub RandomSoundRampFlapUp()
'	debug.print "flap up"
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Up_1"), FlapSoundLevel
		Case 2 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Up_2"), FlapSoundLevel
		Case 3 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Up_3"), FlapSoundLevel
	End Select
End Sub

Sub RandomSoundRampFlapDown()
'	debug.print "flap down"	
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Down_1"), FlapSoundLevel
		Case 2 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Down_2"), FlapSoundLevel
		Case 3 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Down_3"), FlapSoundLevel
	End Select
End Sub



'/////////////////////////////  GATE SOUNDS  ////////////////////////////
Sub Gate5_Hit:SoundBallGate1WayPass Gate5:End Sub	'using this as the ball speed is usually quite low

dim Gate1JustHit:Gate1JustHit = 0
Sub Gate1Helper_Hit 'Gate1_Hit
	if Gate1JustHit = 1 then exit sub	'exit if gate was just hit
	Gate1JustHit = 1
	Gate1Helper.TimerInterval = 100
	Gate1Helper.TimerEnabled = 1
'	debug.print "Gate1: " & activeball.velx
	if activeball.velx > 0 then 'blocked
		SoundBallGate1WayBlockedSmall Gate1
	Else
		SoundBallGate1WayPass Gate1
	end if
End Sub

sub Gate1Helper_timer
	Gate1JustHit = 0
	Me.TimerEnabled = 0
end sub

dim Gate2JustHit:Gate2JustHit = 0
Sub Gate2Helper_Hit 'Gate2_Hit
	if Gate2JustHit = 1 then exit sub	'exit if gate was just hit
	Gate2JustHit = 1
	Gate2Helper.TimerInterval = 100
	Gate2Helper.TimerEnabled = 1
'	debug.print "Gate2: " & activeball.velx
	if activeball.velx > 0 then 'blocked
		SoundBallGate1WayBlockedLarge Gate2
	Else
		SoundBallGate1WayPass Gate2
	end if
End Sub

sub Gate2Helper_timer
	Gate2JustHit = 0
	Me.TimerEnabled = 0
end sub

Sub SoundBallGate1WayPass(obj)
	PlaySoundAtLevelStatic ("TOM_Gate2_5"), GateSoundLevel, obj
End Sub

Sub SoundBallGate2Way(obj)
	PlaySoundAtLevelStatic ("TOM_Gate3_3"), GateSoundLevel, obj
End Sub

Sub SoundBallGate1WayBlockedLarge(obj)
	PlaySoundAtLevelStatic ("TOM_Gate_FastTrigger_1"), GateSoundLevel * 0.5, obj
End Sub

Sub SoundBallGate1WayBlockedSmall(obj)
	PlaySoundAtLevelStatic ("TOM_Gate_FastTrigger_4"), GateSoundLevel * 0.2, obj
End Sub


'///////////////////////////  LOCKING KICKER SOLENOID  //////////////////////////
Dim Solenoid_LeftLockingKickback_SoundLevel
Solenoid_LeftLockingKickback_SoundLevel = 1
Const Cartridge_Kickers					= "WS_WHD_REV01"

Sub RandomSoundLockingKickerSolenoid()
	PlaySoundAtLevelStatic (Cartridge_Kickers & "_Locking_Kickback_" & Int(Rnd*4)+1), Solenoid_LeftLockingKickback_SoundLevel, KickBack
End Sub



'/////////////////////////////////////////////////////////////////
'					End Mechanical Sounds
'/////////////////////////////////////////////////////////////////


'******************************************************
'****  END FLEEP MECHANICAL SOUNDS
'******************************************************

'******************************************************
'****  FLIPPER CORRECTIONS by nFozzy
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
	dim x, a : a = Array(LF, RF)
	for each x in a
		x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
		x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
		x.enabled = True
		x.TimeDelay = 60
	Next

	AddPt "Polarity", 0, 0, 0
	AddPt "Polarity", 1, 0.05, -5
	AddPt "Polarity", 2, 0.4, -5
	AddPt "Polarity", 3, 0.6, -4.5
	AddPt "Polarity", 4, 0.65, -4.0
	AddPt "Polarity", 5, 0.7, -3.5
	AddPt "Polarity", 6, 0.75, -3.0
	AddPt "Polarity", 7, 0.8, -2.5
	AddPt "Polarity", 8, 0.85, -2.0
	AddPt "Polarity", 9, 0.9,-1.5
	AddPt "Polarity", 10, 0.95, -1.0
	AddPt "Polarity", 11, 1, -0.5
	AddPt "Polarity", 12, 1.1, 0
	AddPt "Polarity", 13, 1.3, 0

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
	Dim b

	If Flipper1.currentangle = Endangle1 and EOSNudge1 <> 1 Then
		EOSNudge1 = 1
		'debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
		If Flipper2.currentangle = EndAngle2 Then 
			For b = 0 to Ubound(gBOT)
				If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper1) Then
					'Debug.Print "ball in flip1. exit"
					exit Sub
				end If
			Next
			For b = 0 to Ubound(gBOT)
				If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper2) Then
					gBOT(b).velx = gBOT(b).velx / 1.3
					gBOT(b).vely = gBOT(b).vely - 0.5
				end If
			Next
		End If
	Else 
		If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 then EOSNudge1 = 0
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

Function max(a,b)
	If a > b Then
		max = a
	Else
		max = b
	End If
End Function

Function min(a,b)
	If a > b Then
		min = b
	Else
		min = a
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

dim LFPress, LFPress1, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
'Const EOSTnew = 1 'EM's to late 80's
Const EOSTnew = 0.8 '90's and later
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Const FlipperCoilRampupMode = 0   	'0 = fast, 1 = medium, 2 = slow (tap passes should work)
Select Case FlipperCoilRampupMode 
	Case 0:
		SOSRampup = 2.5
		LeftFlipper1.rampup = 1.5
	Case 1:
		SOSRampup = 6
		LeftFlipper1.rampup = 2
	Case 2:
		SOSRampup = 8.5
		LeftFlipper1.rampup = 2.5
End Select

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
'Const EOSReturn = 0.055  'EM's
'Const EOSReturn = 0.045  'late 70's to mid 80's
Const EOSReturn = 0.035  'mid 80's to early 90's
'Const EOSReturn = 0.025  'mid 90's and later
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
		Dim b

		For b = 0 to UBound(gBOT)
			If Distance(gBOT(b).x, gBOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
				If gBOT(b).vely >= -0.4 Then gBOT(b).vely = -0.4
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

' sling corners targetbouncer disabled
sub zCol_Rubber_Post046_hit
	RubbersD.dampen Activeball
end sub

sub zCol_Rubber_Post041_hit
	RubbersD.dampen Activeball
end sub

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

	Public Sub Update	'tracks in-ball-velocity
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

  Public Property Get Prim(): Prim = m_prim: End Property
  Public Property Let Prim(input): m_prim = input: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(input): m_sw = input: End Property

  Public Property Get Animate(): Animate = m_animate: End Property
  Public Property Let Animate(input): m_animate = input: End Property

  Public Property Get IsDropped(): IsDropped = m_isDropped: End Property
  Public Property Let IsDropped(input): m_isDropped = input: End Property

  Public default Function init(primary, secondary, prim, sw, animate, isDropped)
    Set m_primary = primary
    Set m_secondary = secondary
    m_prim = prim
    m_sw = sw
    m_animate = animate
    m_isDropped = isDropped

    Set Init = Me
  End Function
End Class

'Define a variable for each drop target
Dim DT33, DT34, DT35, DT36, DT57, DT59

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
'	animate:			Array slot for handling the animation instrucitons, set to 0
'						Values for animate: 1 - bend target (hit to primary), 2 - drop target (hit to secondary), 3 - brick target (high velocity hit to secondary), -1 - raise target 
'   isDropped:			Boolean which determines whether a drop target is dropped. Set to false if they are initially raised, true if initially dropped.

Set DT33 = (new DropTarget)(sw33, sw33a, BP_sw33, 33, 0, false)
Set DT34 = (new DropTarget)(sw34, sw34a, BP_sw34, 34, 0, false)
Set DT35 = (new DropTarget)(sw35, sw35a, BP_sw35, 35, 0, false)
Set DT36 = (new DropTarget)(sw36, sw36a, BP_sw36, 36, 0, false)
Set DT57 = (new DropTarget)(sw57, sw57a, BP_sw57, 57, 0, false)
Set DT59 = (new DropTarget)(sw59, sw59a, BP_sw59, 59, 0, false)

Dim DTArray
DTArray = Array(DT33, DT34, DT35, DT36, DT57, DT59)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 110 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
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
	Dim prim : prim = DTArray(i).prim
	DTArray(i).animate =  DTCheckBrick(Activeball,prim(0))
	If DTArray(i).animate = 1 or DTArray(i).animate = 3 or DTArray(i).animate = 4 Then
		DTBallPhysics Activeball, prim(0).rotz, DTMass
	End If
	DoDTAnim
End Sub

Sub DTRaise(switch)
	Dim i
	i = DTArrayID(switch)

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
		If DTArray(i).sw = switch Then
			DTArrayID = i
			Exit Function 
		End If
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
	rangle2 = (dtprim.rotz) * 3.1416 / 180
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

Function DTAnimate(primary, secondary, lightmaps, switch, animate)
	dim transz, switchid
	Dim animtime, rangle
	Dim prim: Set prim = lightmaps(0)

	switchid = switch

	Dim ind
	ind = DTArrayID(switchid)

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
	ElseIf (animate = 1 or animate = 4) and animtime > DTDropDelay Then
		primary.collidable = 0
		If animate = 1 then secondary.collidable = 1 else secondary.collidable= 0
		prim.rotx = DTMaxBend * cos(rangle)
		prim.roty = DTMaxBend * sin(rangle)
		animate = 2
		SoundDropTargetDrop prim
	End If

	if animate = 2 Then
		transz = (animtime - DTDropDelay)/DTDropSpeed *  DTDropUnits * -1
		if prim.transz > -DTDropUnits  Then prim.transz = transz
		prim.rotx = DTMaxBend * cos(rangle)/2
		prim.roty = DTMaxBend * sin(rangle)/2
		if prim.transz <= -DTDropUnits Then 
			prim.transz = -DTDropUnits
			secondary.collidable = 0
			DTArray(ind).isDropped = true 'Mark target as dropped
			controller.Switch(Switchid) = 1
			primary.uservalue = 0
			DTAnimate = 0
		Else
			DTAnimate = 2
		end If 
	ElseIf animate = 3 and animtime < DTDropDelay Then
		primary.collidable = 0
		secondary.collidable = 1
		prim.rotx = DTMaxBend * cos(rangle)
		prim.roty = DTMaxBend * sin(rangle)
	ElseIf animate = 3 and animtime > DTDropDelay Then
		primary.collidable = 1
		secondary.collidable = 0
		prim.rotx = 0
		prim.roty = 0
		primary.uservalue = 0
		DTAnimate = 0
	ElseIf animate = -1 Then
		transz = (1 - (animtime)/DTDropUpSpeed) *  DTDropUnits * -1
		If prim.transz = -DTDropUnits Then
			Dim b', BOT
'			BOT = GetBalls
			For b = 0 to UBound(gBOT)
				If InRotRect(gBOT(b).x,gBOT(b).y,prim.x, prim.y, prim.rotz, -25,-10,25,-10,25,25,-25,25) and gBOT(b).z < prim.z+DTDropUnits+25 Then
					gBOT(b).velz = 20
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
		DTArray(ind).isDropped = false 'Mark target as not dropped
		controller.Switch(Switchid) = 0
	ElseIf animate = -2 and animtime > DTRaiseDelay Then
		prim.transz = (animtime - DTRaiseDelay)/DTDropSpeed *  DTDropUnits * -1 + DTDropUpUnits 
		if prim.transz < 0 then
			prim.transz = 0
			primary.uservalue = 0
			DTAnimate = 0
			primary.collidable = 1
			secondary.collidable = 0
		end If 
	End If
	
	Dim x: For each x in lightmaps: x.transz=prim.transz: x.rotx=prim.rotx: x.roty=prim.roty: Next
End Function

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
'		STAND-UP TARGET INITIALIZATION
'******************************************************

Class StandupTarget
  Private m_primary, m_prim, m_sw, m_animate

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(input): Set m_primary = input: End Property

  Public Property Get Prim(): Prim = m_prim: End Property
  Public Property Let Prim(input): m_prim = input: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(input): m_sw = input: End Property

  Public Property Get Animate(): Animate = m_animate: End Property
  Public Property Let Animate(input): m_animate = input: End Property

  Public default Function init(primary, prim, sw, animate)
    Set m_primary = primary
    m_prim = prim
    m_sw = sw
    m_animate = animate

    Set Init = Me
  End Function
End Class

'Define a variable for each stand-up target
Dim ST17, ST18, ST19, ST20

'Set array with stand-up target objects
'
'StandupTargetvar = Array(primary, prim, swtich)
' 	primary: 			vp target to determine target hit
'	prims:				array of primitive target used for visuals and animation
'							IMPORTANT!!! 
'							transy must be used to offset the target animation
'	switch:				ROM switch number
'	animate:			Arrary slot for handling the animation instrucitons, set to 0
' 
'You will also need to add a secondary hit object for each stand up (name sw11o, sw12o, and sw13o on the example Table1)
'these are inclined primitives to simulate hitting a bent target and should provide so z velocity on high speed impacts

Set ST17 = (new StandupTarget)(sw17, BP_sw17, 17, 0)
Set ST18 = (new StandupTarget)(sw18, BP_sw18, 18, 0)
Set ST19 = (new StandupTarget)(sw19, BP_sw19, 19, 0)
Set ST20 = (new StandupTarget)(sw20, BP_sw20, 20, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST17, ST18, ST19, ST20)

'Configure the behavior of Stand-up Targets
Const STAnimStep =  1.5 		'vpunits per animation step (control return to Start)
Const STMaxOffset = 9 			'max vp units target moves when hit

Const STMass = 0.2				'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

'******************************************************
'				STAND-UP TARGETS FUNCTIONS
'******************************************************

Sub STHit(switch)
	Dim i
	i = STArrayID(switch)

	PlayTargetSound
	STArray(i).animate =  STCheckHit(Activeball,STArray(i).primary)

	If STArray(i).animate <> 0 Then
		DTBallPhysics Activeball, STArray(i).primary.orientation, STMass
	End If
	DoSTAnim
End Sub

Function STArrayID(switch)
	Dim i
	For i = 0 to uBound(STArray) 
		If STArray(i).sw = switch Then STArrayID = i:Exit Function 
	Next
End Function

'Check if target is hit on it's face
Function STCheckHit(aBall, target) 
	dim bangle, bangleafter, rangle, rangle2, perpvel, perpvelafter, paravel, paravelafter
	rangle = (target.orientation - 90) * 3.1416 / 180	
	bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
	bangleafter = Atn2(aBall.vely,aball.velx)

	perpvel = cor.BallVel(aball.id) * cos(bangle-rangle)
	paravel = cor.BallVel(aball.id) * sin(bangle-rangle)

	perpvelafter = BallSpeed(aBall) * cos(bangleafter - rangle) 
	paravelafter = BallSpeed(aBall) * sin(bangleafter - rangle)

	If perpvel > 0 and  perpvelafter <= 0 Then
		STCheckHit = 1
	ElseIf perpvel > 0 and ((paravel > 0 and paravelafter > 0) or (paravel < 0 and paravelafter < 0)) Then
		STCheckHit = 1
	Else 
		STCheckHit = 0
	End If
End Function

Sub DoSTAnim()
	Dim i
	For i=0 to Ubound(STArray)
		STArray(i).animate = STAnimate(STArray(i).primary,STArray(i).prim,STArray(i).sw,STArray(i).animate)
	Next
End Sub

Function STAnimate(primary, lightmaps, switch,  animate)
	Dim animtime
	Dim prim: Set prim = lightmaps(0)

	STAnimate = animate

	if animate = 0  Then
		primary.uservalue = 0
		STAnimate = 0
		Exit Function
	Elseif primary.uservalue = 0 then 
		primary.uservalue = gametime
	end if

	animtime = gametime - primary.uservalue

	If animate = 1 Then
		primary.collidable = 0
		prim.transy = -STMaxOffset
		vpmTimer.PulseSw switch
		STAnimate = 2
		Exit Function
	elseif animate = 2 Then
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
	Dim x: For each x in lightmaps: x.transy=prim.transy: Next
End Function

'******************************************************
'		END STAND-UP TARGETS
'******************************************************


'******************************************************
' VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1 		'0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7 	'Level of bounces. Recommmended value of 0.7

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

'****************************************************************
'****  END VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'****************************************************************


'****************************
'*    Room Brightness
'****************************

' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid","VLM.Bake.BumperCaps","VLM.Bake.RampGreen","VLM.Bake.RampRed","VLM.Bake.RampYellow", _
	"VRMetal","VRGlass","VRBeer","Playfield","Plastic with an image","Plastic White","Plastic Black","Opacity80", "MetalVR","Metal_Black_Powdercoat","Metal with an image", _
	"Colour_Red","colormaxnoreflectionhalf","CoinDoor", "Cabinet1","BeerBubble","_noXtraShading","_noXtraShadingVR")
Dim SavedMtlColorArray:     SavedMtlColorArray     = Array(0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)


Sub SetRoomBrightness(Level)
    If Level > 100 Then Level = 100
    If Level < 0 Then Level = 0

    ' Lighting level
    Dim v:  v = (5 + 250 * (Level^2.2)/(100.0^2.2)) / 255

    Dim i: For i = 0 to UBound(RoomBrightnessMtlArray)
        ModulateMaterialBaseColor RoomBrightnessMtlArray(i), i, v
    Next
End Sub

SaveMtlColors
Sub SaveMtlColors
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



'***********************************************************************
'* TABLE OPTIONS *******************************************************
'***********************************************************************

Dim BackMod
Dim BladeMod
Dim OutPostMod
Dim StagedFlipperMod
Dim Cabinetmode								'0 - Siderails On, 1 - Siderails Off, 2 - Extended Blades (only applies when not using VR headset)
Dim VRRoomChoice : VRRoomChoice = 1         '1 - Minimal Room, 2 - Ultra Room (only applies when using VR headset)
Dim LightLevel : LightLevel = 50
Dim ColorLUT : ColorLUT = 1
Dim VolumeDial : VolumeDial = 0.8			'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5 	'Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5 	'Level of ramp rolling volume. Value between 0 and 1

' Options
Const Opt_Light = 0
Const Opt_LUT = 1
Const Opt_Outpost = 2
Const Opt_Volume = 3
Const Opt_Volume_Ramp = 4
Const Opt_Volume_Ball = 5
Const Opt_Cabinet = 6 ' Only when not in VR
Const Opt_VR_Room = 6 ' Only when in VR
Const Opt_BackMod = 7
Const Opt_BladeMod = 8
Const Opt_Staged_Flipper = 9
' Informations
Const Opt_Info_1 = 10
Const Opt_Info_2 = 11

Const NOptions = 12

Const FlexDMD_RenderMode_DMD_GRAY_2 = 0
Const FlexDMD_RenderMode_DMD_GRAY_4 = 1
Const FlexDMD_RenderMode_DMD_RGB = 2
Const FlexDMD_RenderMode_SEG_2x16Alpha = 3
Const FlexDMD_RenderMode_SEG_2x20Alpha = 4
Const FlexDMD_RenderMode_SEG_2x7Alpha_2x7Num = 5
Const FlexDMD_RenderMode_SEG_2x7Alpha_2x7Num_4x1Num = 6
Const FlexDMD_RenderMode_SEG_2x7Num_2x7Num_4x1Num = 7
Const FlexDMD_RenderMode_SEG_2x7Num_2x7Num_10x1Num = 8
Const FlexDMD_RenderMode_SEG_2x7Num_2x7Num_4x1Num_gen7 = 9
Const FlexDMD_RenderMode_SEG_2x7Num10_2x7Num10_4x1Num = 10
Const FlexDMD_RenderMode_SEG_2x6Num_2x6Num_4x1Num = 11
Const FlexDMD_RenderMode_SEG_2x6Num10_2x6Num10_4x1Num = 12
Const FlexDMD_RenderMode_SEG_4x7Num10 = 13
Const FlexDMD_RenderMode_SEG_6x4Num_4x1Num = 14
Const FlexDMD_RenderMode_SEG_2x7Num_4x1Num_1x16Alpha = 15
Const FlexDMD_RenderMode_SEG_1x16Alpha_1x16Num_1x7Num = 16

Const FlexDMD_Align_TopLeft = 0
Const FlexDMD_Align_Top = 1
Const FlexDMD_Align_TopRight = 2
Const FlexDMD_Align_Left = 3
Const FlexDMD_Align_Center = 4
Const FlexDMD_Align_Right = 5
Const FlexDMD_Align_BottomLeft = 6
Const FlexDMD_Align_Bottom = 7
Const FlexDMD_Align_BottomRight = 8

Const FlexDMD_Scaling_Fit = 0
Const FlexDMD_Scaling_Fill = 1
Const FlexDMD_Scaling_FillX = 2
Const FlexDMD_Scaling_FillY = 3
Const FlexDMD_Scaling_Stretch = 4
Const FlexDMD_Scaling_StretchX = 5
Const FlexDMD_Scaling_StretchY = 6
Const FlexDMD_Scaling_None = 7

Const FlexDMD_Interpolation_Linear = 0
Const FlexDMD_Interpolation_ElasticIn = 1
Const FlexDMD_Interpolation_ElasticOut = 2
Const FlexDMD_Interpolation_ElasticInOut = 3
Const FlexDMD_Interpolation_QuadIn = 4
Const FlexDMD_Interpolation_QuadOut = 5
Const FlexDMD_Interpolation_QuadInOut = 6
Const FlexDMD_Interpolation_CubeIn = 7
Const FlexDMD_Interpolation_CubeOut = 8
Const FlexDMD_Interpolation_CubeInOut = 9
Const FlexDMD_Interpolation_QuartIn = 10
Const FlexDMD_Interpolation_QuartOut = 11
Const FlexDMD_Interpolation_QuartInOut = 12
Const FlexDMD_Interpolation_QuintIn = 13
Const FlexDMD_Interpolation_QuintOut = 14
Const FlexDMD_Interpolation_QuintInOut = 15
Const FlexDMD_Interpolation_SineIn = 16
Const FlexDMD_Interpolation_SineOut = 17
Const FlexDMD_Interpolation_SineInOut = 18
Const FlexDMD_Interpolation_BounceIn = 19
Const FlexDMD_Interpolation_BounceOut = 20
Const FlexDMD_Interpolation_BounceInOut = 21
Const FlexDMD_Interpolation_CircIn = 22
Const FlexDMD_Interpolation_CircOut = 23
Const FlexDMD_Interpolation_CircInOut = 24
Const FlexDMD_Interpolation_ExpoIn = 25
Const FlexDMD_Interpolation_ExpoOut = 26
Const FlexDMD_Interpolation_ExpoInOut = 27
Const FlexDMD_Interpolation_BackIn = 28
Const FlexDMD_Interpolation_BackOut = 29
Const FlexDMD_Interpolation_BackInOut = 30

Dim OptionDMD: Set OptionDMD = Nothing
Dim bOptionsMagna, bInOptions : bOptionsMagna = False
Dim OptPos, OptSelected, OptN, OptTop, OptBot, OptSel
Dim OptFontHi, OptFontLo

Sub Options_Open
	bOptionsMagna = False
	On Error Resume Next
	Set OptionDMD = CreateObject("FlexDMD.FlexDMD")
	On Error Goto 0
	If OptionDMD is Nothing Then
		Debug.Print "FlexDMD is not installed"
		Debug.Print "Option UI can not be opened"
		MsgBox "You need to install FlexDMD to access table options"
		Exit Sub
	End If
	If Table1.ShowDT Then OptionDMDFlasher.RotX = -(Table1.Inclination + Table1.Layback)
	bInOptions = True
	OptPos = 0
	OptSelected = False
	OptionDMD.Show = False
	OptionDMD.RenderMode = FlexDMD_RenderMode_DMD_GRAY_4
	OptionDMD.Width = 128
	OptionDMD.Height = 32
	OptionDMD.Clear = True
	OptionDMD.Run = True
	Dim a, scene, font
	Set scene = OptionDMD.NewGroup("Scene")
	Set OptFontHi = OptionDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0)
	Set OptFontLo = OptionDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(100, 100, 100), RGB(100, 100, 100), 0)
	Set OptSel = OptionDMD.NewGroup("Sel")
	Set a = OptionDMD.NewLabel(">", OptFontLo, ">>>")
	a.SetAlignedPosition 1, 16, FlexDMD_Align_Left
	OptSel.AddActor a
	Set a = OptionDMD.NewLabel(">", OptFontLo, "<<<")
	a.SetAlignedPosition 127, 16, FlexDMD_Align_Right
	OptSel.AddActor a
	scene.AddActor OptSel
	OptSel.SetBounds 0, 0, 128, 32
	OptSel.Visible = False
	
	Set a = OptionDMD.NewLabel("Info1", OptFontLo, "MAGNA EXIT/ENTER")
	a.SetAlignedPosition 1, 32, FlexDMD_Align_BottomLeft
	scene.AddActor a
	Set a = OptionDMD.NewLabel("Info2", OptFontLo, "FLIPPER SELECT")
	a.SetAlignedPosition 127, 32, FlexDMD_Align_BottomRight
	scene.AddActor a
	Set OptN = OptionDMD.NewLabel("Pos", OptFontLo, "LINE 1")
	Set OptTop = OptionDMD.NewLabel("Top", OptFontLo, "LINE 1")
	Set OptBot = OptionDMD.NewLabel("Bottom", OptFontLo, "LINE 2")
	scene.AddActor OptN
	scene.AddActor OptTop
	scene.AddActor OptBot
	Options_OnOptChg
	OptionDMD.LockRenderThread
	OptionDMD.Stage.AddActor scene
	OptionDMD.UnlockRenderThread
	OptionDMDFlasher.Visible = True
End Sub

Sub Options_UpdateDMD
	If OptionDMD is Nothing Then Exit Sub
	Dim DMDp: DMDp = OptionDMD.DmdPixels
	If Not IsEmpty(DMDp) Then
		OptionDMDFlasher.DMDWidth = OptionDMD.Width
		OptionDMDFlasher.DMDHeight = OptionDMD.Height
		OptionDMDFlasher.DMDPixels = DMDp
	End If
End Sub

Sub Options_Close
	bInOptions = False
	OptionDMDFlasher.Visible = False
	If OptionDMD is Nothing Then Exit Sub
	OptionDMD.Run = False
	Set OptionDMD = Nothing
End Sub

Function Options_OnOffText(opt)
	If opt Then
		Options_OnOffText = "ON"
	Else
		Options_OnOffText = "OFF"
	End If
End Function

Sub Options_OnOptChg
	If OptionDMD is Nothing Then Exit Sub
	OptionDMD.LockRenderThread
	OptN.Text = (OptPos+1) & "/" & NOptions
	If OptSelected Then
		OptTop.Font = OptFontLo
		OptBot.Font = OptFontHi
		OptSel.Visible = True
	Else
		OptTop.Font = OptFontHi
		OptBot.Font = OptFontLo
		OptSel.Visible = False
	End If
	If OptPos = Opt_Light Then
		If RenderingMode = 2 Then 
			OptTop.Text = "LIGHT LEVEL"
		Else
			OptTop.Text = "LIGHT LEVEL (Req Restart)"
		End If
		OptBot.Text = "LEVEL " & LightLevel
		SaveValue cGameName, "LIGHT", LightLevel
	ElseIf OptPos = Opt_LUT Then
		OptTop.Text = "COLOR SATURATION"
		if ColorLUT = 1 Then OptBot.text = "DISABLED"
		if ColorLUT = 2 Then OptBot.text = "DESATURATED -10%"
		if ColorLUT = 3 Then OptBot.text = "DESATURATED -20%"
		if ColorLUT = 4 Then OptBot.text = "DESATURATED -30%"
		if ColorLUT = 5 Then OptBot.text = "DESATURATED -40%"
		if ColorLUT = 6 Then OptBot.text = "DESATURATED -50%"
		if ColorLUT = 7 Then OptBot.text = "DESATURATED -60%"
		if ColorLUT = 8 Then OptBot.text = "DESATURATED -70%"
		if ColorLUT = 9 Then OptBot.text = "DESATURATED -80%"
		if ColorLUT = 10 Then OptBot.text = "DESATURATED -90%"
		if ColorLUT = 11 Then OptBot.text = "BLACK'N WHITE"
		SaveValue cGameName, "LUT", ColorLUT
	ElseIf OptPos = Opt_Outpost Then
		OptTop.Text = "OUT POST DIFFICULTY"
		If OutPostMod = 0 Then
			OptBot.Text = "EASY"
		ElseIf OutPostMod = 1 Then
			OptBot.Text = "MEDIUM"
		ElseIf OutPostMod = 2 Then
			OptBot.Text = "HARD"
		End If
		SaveValue cGameName, "OUTPOST", OutPostMod
	ElseIf OptPos = Opt_Volume Then
		OptTop.Text = "MECH VOLUME"
		OptBot.Text = "LEVEL " & CInt(VolumeDial * 100)
		SaveValue cGameName, "VOLUME", VolumeDial
	ElseIf OptPos = Opt_Volume_Ramp Then
		OptTop.Text = "RAMP VOLUME"
		OptBot.Text = "LEVEL " & CInt(RampRollVolume * 100)
		SaveValue cGameName, "RAMPVOLUME", RampRollVolume
	ElseIf OptPos = Opt_Volume_Ball Then
		OptTop.Text = "BALL VOLUME"
		OptBot.Text = "LEVEL " & CInt(BallRollVolume * 100)
		SaveValue cGameName, "BALLVOLUME", BallRollVolume
	ElseIf OptPos = Opt_BackMod Then
		OptTop.Text = "BACK MOD"
		OptBot.Text = Options_OnOffText(BackMod)
		SaveValue cGameName, "BACK_MOD", BackMod
	ElseIf OptPos = Opt_BladeMod Then
		OptTop.Text = "BLADE MOD"
		OptBot.Text = Options_OnOffText(BladeMod)
		SaveValue cGameName, "BLADE_MOD", BladeMod
	ElseIf OptPos = Opt_Cabinet Then
		OptTop.Text = "CABINET MODE"
		If CabinetMode = 0 Then OptBot.Text = "OFF"
		If CabinetMode = 1 Then OptBot.Text = "ON"
'		If CabinetMode = 2 Then OptBot.Text = "EXT. BLADES"
		SaveValue cGameName, "CABINET", CabinetMode
	ElseIf OptPos = Opt_VR_Room Then
		OptTop.Text = "VR ROOM"
		If VRRoomChoice = 0 Then OptBot.Text = "MINIMAL"
		If VRRoomChoice = 1 Then OptBot.Text = "ULTRA"
		SaveValue cGameName, "VRROOM", VRRoomChoice
	ElseIf OptPos = Opt_Staged_Flipper Then
		OptTop.Text = "STAGED FLIPPER"
		OptBot.Text = Options_OnOffText(StagedFlipperMod)
		SaveValue cGameName, "STAGED", StagedFlipperMod
	ElseIf OptPos = Opt_Info_1 Then
		OptTop.Text = "VPX " & VersionMajor & "." & VersionMinor & "." & VersionRevision
		OptBot.Text = "GUNS'N ROSES " & TableVersion
	ElseIf OptPos = Opt_Info_2 Then
		OptTop.Text = "RENDER MODE"
		If RenderingMode = 0 Then OptBot.Text = "DEFAULT"
		If RenderingMode = 1 Then OptBot.Text = "STEREO 3D"
		If RenderingMode = 2 Then OptBot.Text = "VR"
	End If
	OptTop.Pack
	OptTop.SetAlignedPosition 127, 1, FlexDMD_Align_TopRight
	OptBot.SetAlignedPosition 64, 16, FlexDMD_Align_Center
	OptionDMD.UnlockRenderThread
	UpdateMods
End Sub

Sub Options_Toggle(amount)
	If OptionDMD is Nothing Then Exit Sub
	If OptPos = Opt_Light Then
		LightLevel = LightLevel + amount * 10
		If LightLevel < 0 Then LightLevel = 100
		If LightLevel > 100 Then LightLevel = 0
	ElseIf OptPos = Opt_LUT Then
		ColorLUT = ColorLUT + amount * 1
		If ColorLUT < 1 Then ColorLUT = 11
		If ColorLUT > 11 Then ColorLUT = 1
	ElseIf OptPos = Opt_Outpost Then
		OutPostMod = OutPostMod + amount
		If OutPostMod < 0 Then OutPostMod = 2
		If OutPostMod > 2 Then OutPostMod = 0
	ElseIf OptPos = Opt_Volume Then
		VolumeDial = VolumeDial + amount * 0.1
		If VolumeDial < 0 Then VolumeDial = 1
		If VolumeDial > 1 Then VolumeDial = 0
	ElseIf OptPos = Opt_Volume_Ramp Then
		RampRollVolume = RampRollVolume + amount * 0.1
		If RampRollVolume < 0 Then RampRollVolume = 1
		If RampRollVolume > 1 Then RampRollVolume = 0
	ElseIf OptPos = Opt_Volume_Ball Then
		BallRollVolume = BallRollVolume + amount * 0.1
		If BallRollVolume < 0 Then BallRollVolume = 1
		If BallRollVolume > 1 Then BallRollVolume = 0
	ElseIf OptPos = Opt_BackMod Then
		BackMod = 1 - BackMod
	ElseIf OptPos = Opt_BladeMod Then
		BladeMod = 1 - BladeMod
	ElseIf OptPos = Opt_Staged_Flipper Then
		StagedFlipperMod = 1 - StagedFlipperMod
	ElseIf OptPos = Opt_Cabinet Then
		CabinetMode = CabinetMode + amount
		If CabinetMode < 0 Then CabinetMode = 1
		If CabinetMode > 1 Then CabinetMode = 0
	ElseIf OptPos = Opt_VR_Room Then
		VRRoomChoice = VRRoomChoice + amount
		If VRRoomChoice < 0 Then VRRoomChoice = 2
		If VRRoomChoice > 2 Then VRRoomChoice = 0
	End If
End Sub

Sub Options_KeyDown(ByVal keycode)
	If OptSelected Then
		If keycode = LeftMagnaSave Then ' Exit / Cancel
			OptSelected = False
		ElseIf keycode = RightMagnaSave Then ' Enter / Select
			OptSelected = False
		ElseIf keycode = LeftFlipperKey Then ' Next / +
			Options_Toggle	-1
		ElseIf keycode = RightFlipperKey Then ' Prev / -
			Options_Toggle	1
		End If
	Else
		If keycode = LeftMagnaSave Then ' Exit / Cancel
			Options_Close
		ElseIf keycode = RightMagnaSave Then ' Enter / Select
			If OptPos < Opt_Info_1 Then OptSelected = True
		ElseIf keycode = LeftFlipperKey Then ' Next / +
			OptPos = OptPos - 1
			If OptPos = Opt_VR_Room And RenderingMode <> 2 Then OptPos = OptPos - 1 ' Skip VR option in non VR mode
			If OptPos = Opt_Cabinet And RenderingMode = 2 Then OptPos = OptPos - 1 ' Skip non VR option in VR mode
			If OptPos < 0 Then OptPos = NOptions - 1
		ElseIf keycode = RightFlipperKey Then ' Prev / -
			OptPos = OptPos + 1
			If OptPos = Opt_VR_Room And RenderingMode <> 2 Then OptPos = OptPos + 1 ' Skip VR option in non VR mode
			If OptPos = Opt_Cabinet And RenderingMode = 2 Then OptPos = OptPos + 1 ' Skip non VR option in VR mode
			If OptPos >= NOPtions Then OptPos = 0
		End If
	End If
	Options_OnOptChg
End Sub

Sub Options_Load
	Dim x
    x = LoadValue(cGameName, "LIGHT") : If x <> "" Then LightLevel = CInt(x) Else LightLevel = 50
    x = LoadValue(cGameName, "LUT") : If x <> "" Then ColorLUT = CInt(x) Else ColorLUT = 1
    x = LoadValue(cGameName, "OUTPOST") : If x <> "" Then OutPostMod = CInt(x) Else OutPostMod = 1
    x = LoadValue(cGameName, "VOLUME") : If x <> "" Then VolumeDial = CNCDbl(x) Else VolumeDial = 0.8
    x = LoadValue(cGameName, "RAMPVOLUME") : If x <> "" Then RampRollVolume = CNCDbl(x) Else RampRollVolume = 0.5
    x = LoadValue(cGameName, "BALLVOLUME") : If x <> "" Then BallRollVolume = CNCDbl(x) Else BallRollVolume = 0.5
    x = LoadValue(cGameName, "BACK_MOD") : If x <> "" Then BackMod = CInt(x) Else BackMod = 0
    x = LoadValue(cGameName, "BLADE_MOD") : If x <> "" Then BladeMod = CInt(x) Else BladeMod = 0
    x = LoadValue(cGameName, "CABINET") : If x <> "" Then CabinetMode = CInt(x) Else CabinetMode = 0
    x = LoadValue(cGameName, "VRROOM") : If x <> "" Then VRRoomChoice = CInt(x) Else VRRoomChoice = 0
    x = LoadValue(cGameName, "STAGED") : If x <> "" Then StagedFlipperMod = CInt(x) Else StagedFlipperMod = 0
End Sub

Sub UpdateMods
	Dim v, x
	
'	' Lighting level
'	v = Int((5 + 250 * (LightLevel^2.2)/(100.0^2.2)) / 2)
'	UpdateMaterial "VLM.Bake.Active", 0, 0, 0, 0, 0, 0, 1, RGB(v, v, v), RGB(0, 0, 0), RGB(0, 0, 0), False, True, 0, 0, 0, 0
'	UpdateMaterial "VLM.Bake.Solid", 0, 0, 0, 0, 0, 0, 1, RGB(v, v, v), RGB(0, 0, 0), RGB(0, 0, 0), False, False, 0, 0, 0, 0
'	UpdateMaterial "VLM.Bake.BumperCaps", 0, 0, 0, 0, 0, 0, 1, RGB(v, v, v), RGB(0, 0, 0), RGB(0, 0, 0), False, True, 0, 0, 0, 0
'	UpdateMaterial "VLM.Bake.RampGreen", 0, 0, 0, 0, 0, 0, 1, RGB(v, v, v), RGB(0, 0, 0), RGB(0, 0, 0), False, True, 0, 0, 0, 0
'	UpdateMaterial "VLM.Bake.RampRed", 0, 0, 0, 0, 0, 0, 1, RGB(v, v, v), RGB(0, 0, 0), RGB(0, 0, 0), False, True, 0, 0, 0, 0
'	UpdateMaterial "VLM.Bake.RampYellow", 0, 0, 0, 0, 0, 0, 1, RGB(v, v, v), RGB(0, 0, 0), RGB(0, 0, 0), False, True, 0, 0, 0, 0

   ' Room Lighting level
	If RenderingMode = 2 Then SetRoomBrightness LightLevel

	' Desaturation
	if ColorLUT = 1 Then Table1.ColorGradeImage = ""
	if ColorLUT = 2 Then Table1.ColorGradeImage = "colorgradelut256x16-10"
	if ColorLUT = 3 Then Table1.ColorGradeImage = "colorgradelut256x16-20"
	if ColorLUT = 4 Then Table1.ColorGradeImage = "colorgradelut256x16-30"
	if ColorLUT = 5 Then Table1.ColorGradeImage = "colorgradelut256x16-40"
	if ColorLUT = 6 Then Table1.ColorGradeImage = "colorgradelut256x16-50"
	if ColorLUT = 7 Then Table1.ColorGradeImage = "colorgradelut256x16-60"
	if ColorLUT = 8 Then Table1.ColorGradeImage = "colorgradelut256x16-70"
	if ColorLUT = 9 Then Table1.ColorGradeImage = "colorgradelut256x16-80"
	if ColorLUT = 10 Then Table1.ColorGradeImage = "colorgradelut256x16-90"
	if ColorLUT = 11 Then Table1.ColorGradeImage = "colorgradelut256x16-100"

	v = 1 - BackMod
	For Each x in BP_BP_Empty: x.Visible = v: Next
	
	v = BackMod
	For Each x in BP_BP_Jungle: x.Visible = v: Next

	v = BladeMod
	For Each x in BP_Blades: x.Visible = v: Next

	' VR Mode
	Dim VRRoom
	If RenderingMode = 2 Then VRRoom = VRRoomChoice Else VRRoom = 0
	DIM VRThings
	If VRRoom > 0 Then

	Else

	End if

	' Side rails (always for VR, and if not in cabinet mode)
	v = VRRoom > 0 Or CabinetMode = 0
	Pincab_Rails.Visible = v

	' Extended side rails for dekstop
'	v = VRRoom = 0 And CabinetMode = 0 And Table1.ShowDT
'	Pincab_Rails.Visible = v

	' Outlane  difficulty
	v = OutPostMod = 0 Or OutPostMod = 1
	If v Then zCol_Rubber_Post_L1.Collidable = 1 Else zCol_Rubber_Post_L1.Collidable = 0
	For Each x in BP_Post_L: x.Visible = v: Next
	If OutPostMod = 0 Then zCol_Rubber_Post_R1.Collidable = 1 Else zCol_Rubber_Post_R1.Collidable = 0
	If OutPostMod = 1 Then zCol_Rubber_Post_R2.Collidable = 1 Else zCol_Rubber_Post_R2.Collidable = 0
	If OutPostMod = 2 Then zCol_Rubber_Post_R3.Collidable = 1 Else zCol_Rubber_Post_R3.Collidable = 0
	For Each x in BP_Post_R: x.X = 940.0 + 13.0 * OutPostMod: Next

End Sub

'Set Up VR Backglass
Sub SetBackglass()
	Dim obj, vrbglight

	For Each obj In VRBackglass
		obj.x = obj.x
		obj.height = - obj.y + 347
		obj.y = 10
		obj.rotX = -86.5
	Next

End Sub


' Culture neutral string to double conversion (handles situation where you don't know how the string was written)
Function CNCDbl(str)
    Dim strt, Sep, i
    If IsNumeric(str) Then
        CNCDbl = CDbl(str)
    Else
        Sep = Mid(CStr(0.5), 2, 1)
        Select Case Sep
        Case "."
            i = InStr(1, str, ",")
        Case ","
            i = InStr(1, str, ".")
        End Select
        If i = 0 Then     
            CNCDbl = Empty
        Else
            strt = Mid(str, 1, i - 1) & Sep & Mid(str, i + 1)
            If IsNumeric(strt) Then
                CNCDbl = CDbl(strt)
            Else
                CNCDbl = Empty
            End If
        End If
    End If
End Function

' Lava Lamp code below.  Thank you STEELY!
Dim Lspeed(7), Lbob(7), blobAng(7), blobRad(7), blobSiz(7), Blob, Bcnt   ' for VR Lava Lamp

Bcnt = 0
For Each Blob in Lspeed
	Lbob(Bcnt) = .2
	Lspeed(Bcnt) = Int((5 * Rnd) + 2) * .05
	Bcnt = Bcnt + 1
Next

Sub LavaTimer_Timer()
	Bcnt = 0
	For Each Blob in Lava
		If Blob.TransZ <= VRLavaBase.Size_Z * 1.5 Then 	'Change blob direction to up
			Lspeed(Bcnt) = Int((5 * Rnd) + 2) * .05		'travel speed
			blobAng(Bcnt) = Int((359 * Rnd) + 1)			'blob location/angle from center
			blobRad(Bcnt) = Int((40 * Rnd) + 10)			'blob distance from center, radius
			blobSiz(Bcnt) = Int((150 * Rnd) + 100)		'blob size
			Blob.Size_x = blobSiz(Bcnt):Blob.Size_y = blobSiz(Bcnt):Blob.Size_z = blobSiz(Bcnt)
			Blob.X = Round(Cos(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + VRLavaBase.X	'place blob
			Blob.Y = Round(Sin(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + VRLavaBase.Y
		End If

		If Blob.TransZ => VRLavaBase.Size_Z*5 Then		'Change blob direction to down
			blobAng(Bcnt) = Int((359 * Rnd) + 1)			'blob location/angle from center
			blobRad(Bcnt) = Int((40 * Rnd) + 10)			'blob distance from center,radius
			Blob.X = Round(Cos(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + VRLavaBase.X	'place blob
			Blob.Y = Round(Sin(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + VRLavaBase.Y
			Lspeed(Bcnt) = Int((8 * Rnd) + 5) * .05:Lspeed(Bcnt) = Lspeed(Bcnt) * -1				'travel speed
		End If

		' Make blob wobble
		If Blob.Size_x > blobSiz(Bcnt) + blobSiz(Bcnt)*.15 or Blob.Size_x < blobSiz(Bcnt) - blobSiz(Bcnt)*.15  Then Lbob(Bcnt) = Lbob(Bcnt) * -1
		Blob.Size_x = Blob.Size_x + Lbob(Bcnt)
		Blob.Size_y = Blob.Size_y + Lbob(Bcnt)
		Blob.Size_z = Blob.Size_Z - Lbob(Bcnt) * .66
		Blob.TransZ = Blob.TransZ + Lspeed(Bcnt)		'Move blob
		Bcnt = Bcnt + 1
	Next
End Sub

' Make bubbles random speed
Sub BeerTimer_Timer()

	Randomize(21)
	VRBeerBubble1.z = VRBeerBubble1.z + Rnd(1)*0.5
	if VRBeerBubble1.z > -771 then VRBeerBubble1.z = -955

	VRBeerBubble2.z = VRBeerBubble2.z + Rnd(1)*1
	if VRBeerBubble2.z > -768 then VRBeerBubble2.z = -955

	VRBeerBubble3.z = VRBeerBubble3.z + Rnd(1)*1
	if VRBeerBubble3.z > -768 then VRBeerBubble3.z = -955

	VRBeerBubble4.z = VRBeerBubble4.z + Rnd(1)*0.75
	if VRBeerBubble4.z > -774 then VRBeerBubble4.z = -955

	VRBeerBubble5.z = VRBeerBubble5.z + Rnd(1)*1
	if VRBeerBubble5.z > -771 then VRBeerBubble5.z = -955

	VRBeerBubble6.z = VRBeerBubble6.z + Rnd(1)*1
	if VRBeerBubble6.z > -774 then VRBeerBubble6.z = -955

	VRBeerBubble7.z = VRBeerBubble7.z + Rnd(1)*0.8
	if VRBeerBubble7.z > -768 then VRBeerBubble7.z = -955

	VRBeerBubble8.z = VRBeerBubble8.z + Rnd(1)*1
	if VRBeerBubble8.z > -771 then VRBeerBubble8.z = -955

End Sub

Dim TableIrr: TableIrr = Array(_
	&h364235, &h3a3e31, &h4b4c3f, &h5a554b, &h6a645a, &h777668, &h7e7971, &h8a7e78, &h918881, &h948a82, &h8d817b, &h857a73, &h746964, &h676053, &h4e4943, &h3e3d3f, _
	&h35482f, &h535a3c, &h7c8566, &h94987d, &h97937d, &h9c9784, &h989188, &hb8a69d, &hd7bfb4, &hd3bfb4, &hb7aaa0, &ha2948a, &h90827a, &h8b7875, &h57514f, &h40413e, _
	&h405436, &h627048, &h838e6b, &ha2a084, &h938b74, &h948a7a, &h8e847f, &hbba59d, &he7c9bc, &hdec1b7, &hb0a29d, &ha29085, &h847167, &h89716b, &h74615c, &h4d483f, _
	&h42623f, &h607c49, &h767d5d, &ha59d83, &ha5977b, &ha89680, &hab9d92, &hcaa59b, &hebc4b4, &hd8b8ab, &h978c85, &h9e8a7f, &ha88e80, &h745b56, &h745956, &h5a4a44, _
	&h547544, &h5e7642, &h696b52, &haa9c82, &hb09d7b, &hbba282, &hd0bfa7, &he0c5b4, &hf1d4c1, &hd2b8a4, &h948374, &h8e7365, &hac8b7c, &h70524d, &h674644, &h624947, _
	&h677949, &h677145, &h6c6a54, &hb5a287, &hb89f74, &hab9574, &he1cbab, &hebd0ba, &hf4ddc6, &hc9b19a, &h8d7763, &h927261, &had8377, &h77504c, &h623d3c, &h604542, _
	&h6e7e53, &h717650, &h6e6853, &ha79376, &hbfa276, &hcdb185, &he8cda6, &he2bf92, &hedd3b0, &hd1b196, &hb18f76, &hae846f, &h9a6f63, &h7a4e4a, &h704a48, &h5e4441, _
	&h6a7b5d, &h7e8462, &h817b63, &h7a6c53, &ha28a65, &hd1b486, &hcbae7f, &hb59b77, &he5cdae, &hd6b6a0, &hc89d89, &hb08172, &h744e46, &h814f4c, &h804b49, &h553432, _
	&h55574e, &h959478, &ha39a7f, &h796d54, &h8e7a59, &h9d8864, &h958163, &he3cbab, &hffefd7, &he8c9ba, &hc09487, &h865952, &h6d443f, &h8f5c58, &h82514c, &h4b312d, _
	&h464743, &h959179, &ha4987d, &h8b7c61, &h8e7c5f, &ha28e73, &h9f907e, &hf1dccb, &hfff5e7, &hccada3, &h906861, &h804340, &h7e4e4a, &ha1706a, &h9b6c66, &h4b3732, _
	&h3f423a, &h7d745e, &h9e8d70, &hb5a182, &ha89882, &hcab6a5, &he0cabd, &hf4e1d4, &he5d0c6, &h89716a, &h8c5855, &h8b4645, &h835551, &h96716c, &h9d736f, &h4d3834, _
	&h353930, &h5c5240, &ha28c6d, &hd2bc9b, &hddcab6, &heed8c8, &hf4ddcb, &hd1bfaf, &h897b73, &h76645e, &ha27c78, &h8c5b59, &h86615d, &h755e59, &h80635e, &h423331, _
	&h353a34, &h716141, &hb89869, &he2cba6, &heed9c6, &hfae1cf, &hf3ddc8, &hb8a795, &h807269, &h867772, &hac948f, &h90746f, &h9c837d, &h8b7b76, &h61514c, &h3b2f2e, _
	&h3b4039, &h7d7054, &hb79d7a, &he3d2b7, &hefddcd, &hf3dfce, &hebd8c6, &hc8b6a6, &hc5b4a7, &hdfcfc4, &he8d6cc, &hbaa7a0, &hc5b4ab, &hb4a79e, &h4f4640, &h393433, _
	&h415038, &h80805e, &hcbbba1, &he8d8c9, &hf4e0d2, &hf9e5d6, &hf5e2d5, &hf2e0d3, &hffefe2, &hfff8eb, &hffecdf, &hd2c0b6, &hcabbb1, &hc0b0a6, &h534b45, &h413d3c, _
	&h415b39, &h707c58, &hbbae9b, &he8d5c9, &hf5e1d2, &hfce8db, &hfff0e4, &hfff7eb, &hfffcf0, &hfffaee, &hffeee1, &hd8c6bb, &hc8bab0, &hbbaca2, &h504542, &h3a3334, _
	&h406339, &h696f54, &h9b8b81, &he5d1c7, &hf4e2d5, &hfcebde, &hfff7e9, &hfff9ec, &hfffaee, &hfffaed, &hffefe2, &hdeccc1, &hcabcb1, &hbaaca1, &h473f3b, &h332e30, _
	&h386239, &h667b59, &haea698, &he8d5cb, &hf2e1d4, &hfdede0, &hfff7ea, &hfff9ec, &hfffaed, &hfff9ec, &hffefe2, &he2d0c5, &hc8bab0, &hbfb1a7, &h4f4642, &h3b3637, _
	&h355f34, &h788367, &hbfb0a6, &head7ce, &hf2e0d5, &hffefe2, &hfff7ea, &hfff9ec, &hfff9ec, &hfff9eb, &hffefe1, &he6d4ca, &hc7b9b0, &hbaaea3, &h5b524b, &h4e4946, _
	&h345336, &h6f8060, &ha89f92, &he6d3cb, &hf0ded3, &hfcebde, &hfff7e9, &hfff9ec, &hfff9ec, &hfff8ea, &hfeede0, &he9d7cc, &hc9bbb1, &hb7aaa0, &h5e5650, &h69625e, _
	&h324f38, &h5e6d52, &h8f897b, &hdfcdc3, &heddcd0, &hfae9dc, &hfff5e8, &hfff9ec, &hfff9ec, &hfff7ea, &hfdecdf, &he9d7cc, &hc7b8ae, &hb8aaa1, &h69605a, &h746c68, _
	&h355239, &h5e6952, &haba395, &hd9c9bf, &hd7c8bd, &hfae8dc, &hfff4e6, &hfff9eb, &hfff9ec, &hfff8ea, &hfeebdf, &he5d3c9, &hb7aaa1, &hb9aba3, &h716862, &h796f68, _
	&h405942, &h5b634f, &hc3b6ab, &hdccac1, &hac9f96, &heeddd1, &hfff5e8, &hfff8eb, &hfff9ed, &hfff7eb, &hfeebdf, &hd2c1b6, &ha0958c, &hb2a49d, &h6f6660, &h7b7167, _
	&h54754d, &h626c54, &hcabdb1, &hd8c8c0, &h988d86, &hdfcdc2, &hfff5e8, &hfff6ea, &hfff9ed, &hfff7eb, &hfae8dc, &hb3a59c, &h91867f, &ha99b94, &h756d66, &h7e766d, _
	&h859e7e, &h6a715c, &hc2bbad, &hc8bab1, &h9b8f89, &hcdbbb3, &hfff2e5, &hfff8ec, &hfff9ee, &hfff8ec, &hf9e7db, &hb0a29c, &h988d87, &h9f938c, &h776d67, &h80786d, _
	&h9b9e92, &h67675b, &hada79b, &hc7b9ad, &hcfc0b6, &he3d2c8, &hfff3e6, &hfffbef, &hfffbf0, &hfff8eb, &hf9e6db, &hd3c1ba, &hc0b2aa, &ha3978e, &h6c645d, &h736d65, _
	&h86847e, &h615e56, &h817c74, &hb1a39a, &hebd9cd, &hefdacf, &hffece0, &hfff9ec, &hfffaee, &hfff5e9, &hebd8cd, &hd7c5bc, &hbfb1a9, &h837a74, &h4d4440, &h514d47, _
	&h6e6f69, &h55504b, &h494641, &h59524e, &hbbaca6, &he7d4ca, &hfff7ea, &hfff8eb, &hfffbed, &hfff3e8, &hecdace, &hafa197, &h6f6560, &h3c3732, &h4b4640, &h4a4842, _
	&h636561, &h4e4a47, &h46423d, &h443f3c, &h564f4b, &h92867f, &hfff2e7, &hfffcee, &hffede0, &he4d3c8, &h9f958d, &h4b4540, &h38312e, &h3d3733, &h4a413d, &h48433f, _
	&h4b5049, &h474540, &h47433f, &h48423f, &h4a443f, &h574d48, &h776d67, &hc0afa5, &ha4978e, &h807770, &h413d38, &h342f2a, &h413c37, &h4b463f, &h4e4741, &h484341, _
	&h434d45, &h4c4944, &h4b4744, &h514a46, &h59514b, &h5d524c, &h605752, &h7d6f69, &h746861, &h544f47, &h4c473f, &h574d48, &h5a524c, &h58514a, &h594e4b, &h504b4c, _
	&h505851, &h4f514d, &h54514e, &h645c56, &h69635c, &h6a635e, &h706964, &h776f68, &h83756c, &h777468, &h696558, &h6d6458, &h696254, &h616050, &h595e51, &h4d5651, _
	&hd5cbbf, &hded7d2, &hf5f0dd, &hfff7e8, &hfff7ea, &hfff5e7, &hfff3e6, &hfff3e7, &hfff4e8, &hfff4e7, &hfff4e8, &hfff5ec, &hfff1e2, &hffecd5, &hd4c9b9, &hb9b7b6, _
	&hcec6b3, &he4d5c2, &hf1e6d8, &hfff2e4, &hfff3e5, &hfff6e9, &hfff6e8, &hfff6e8, &hfff7e9, &hfff6e9, &hfff6e9, &hfff7e9, &hfff7e9, &hfee7df, &hc9bbb0, &hb2a9a1, _
	&hdeccbe, &he8d7c5, &hf3e4d7, &hffefe0, &hfff3e5, &hfff6e8, &hfff8e9, &hfff9e9, &hfff9ea, &hfff8ea, &hfff8ea, &hfff8eb, &hfff4e8, &he3d2c6, &hb6a79e, &haca699, _
	&hd0c5b7, &hdacdb9, &heddfce, &hfcecdb, &hfff1e1, &hfff6e8, &hfff7ea, &hfff7ea, &hfff8eb, &hfff8ec, &hfff8ec, &hfff9ee, &hfff8eb, &hddcbc0, &hbbaaa1, &hb5a89d, _
	&hc5af9f, &hcabba6, &he8d8c7, &hfaead9, &hffeddc, &hfff5e6, &hfff8eb, &hfff7ea, &hfff8eb, &hfff8ec, &hfff8eb, &hfffbee, &hfff4e7, &hcbbcb1, &hb0a399, &hb3a39e, _
	&hc2a770, &hd7c3a4, &hefdece, &hffeddd, &hfff1e2, &hfff6e8, &hfff7eb, &hfff7ea, &hfff8eb, &hfff9ec, &hfff8eb, &hfffaed, &hfff5e9, &hdcccc2, &hbcaca3, &hb8a6a3, _
	&hcbb394, &hd2beaa, &hf2e1d3, &hfff3e4, &hfff7e8, &hfff8e9, &hfff8ea, &hfff8ea, &hfff9eb, &hfff9ec, &hfff8eb, &hfffaed, &hfffcef, &hfdebde, &hcebeb3, &hbcaaa7, _
	&hd3bda9, &hd5bda3, &hfae6d7, &hfff5e9, &hfff8ea, &hfff7ea, &hfff8ea, &hfff8ea, &hfff9ec, &hfffaec, &hfff9eb, &hfffaed, &hfffaed, &hfff2e5, &hd4c3b8, &hbba9a5, _
	&hcabcaa, &hc8b59e, &hfae9d9, &hfff6ea, &hfff8eb, &hfff7ea, &hfff7ea, &hfff8eb, &hfff9ec, &hfffaed, &hfff9ec, &hfff9ec, &hfffbee, &hfff2e7, &hd4c4b9, &hbcaba7, _
	&hc6b8aa, &haa9c92, &head9cd, &hfff6e8, &hfff8eb, &hfff8eb, &hfff7ea, &hfff8eb, &hfff9ec, &hfff9ec, &hfff9ec, &hfff9ec, &hfffbed, &hfff6ea, &hd9c7bc, &hbdada8, _
	&hb4ab9b, &hc5b9b0, &hffefe3, &hfff8eb, &hfff9eb, &hfff8eb, &hfff7ea, &hfff8eb, &hfff9ec, &hfff9ec, &hfffaed, &hfff9ed, &hfffbee, &hfff8ec, &he1cdc1, &hbcaba6, _
	&hd3c8b7, &he9dace, &hfff5e9, &hfff6ea, &hfff8ea, &hfff8eb, &hfff8eb, &hfff8ec, &hfff9ec, &hfff9ec, &hfffaed, &hfff9ed, &hfffcef, &hfff9ec, &he6d4c8, &hbdada8, _
	&hdccdc2, &hf4e3d6, &hfff4e7, &hfff7ea, &hfff8ea, &hfff9eb, &hfff9ec, &hfff9ec, &hfffaed, &hfffaed, &hfffaed, &hfff9ed, &hfffaed, &hfffaed, &hfae7da, &hc2b2ad, _
	&hdbcbc2, &hfeeddd, &hfff8e9, &hfff9ea, &hfff8ea, &hfff9eb, &hfff9ec, &hfffaed, &hfffaed, &hfffaed, &hfff9ed, &hfff9ed, &hfffaee, &hfffbee, &hfff0e3, &hc1b0ad, _
	&he0cfc5, &hfff2e3, &hfff8ea, &hfff8eb, &hfff9eb, &hfffaec, &hfffaed, &hfffaed, &hfffaed, &hfffaed, &hfff9ec, &hfff9ec, &hfffaed, &hfffcef, &hfff5e6, &hc5b5b1, _
	&he4d3c7, &hfff5e4, &hfff8ea, &hfff8eb, &hfff9ec, &hfffaed, &hfffbee, &hfffbed, &hfffbee, &hfffaee, &hfff9ed, &hfff9ed, &hfff9ec, &hfffcef, &hfff5e7, &hcbb9b6, _
	&he8d6ca, &hfff7e7, &hfff9ec, &hfff8ec, &hfffaed, &hfffbee, &hfffbee, &hfffbee, &hfffbee, &hfffbed, &hfffaed, &hfffaed, &hfff9ec, &hfffbef, &hfff6e8, &hd0bfbb, _
	&hebd8cd, &hfff9e9, &hfffcef, &hfffaed, &hfffaed, &hfffbee, &hfffbee, &hfffbee, &hfffbed, &hfffbed, &hfffaed, &hfffaed, &hfffaed, &hfffbef, &hfff6e8, &hd9c7c3, _
	&hf3e0d4, &hfff9e9, &hfffbee, &hfffaed, &hfffbed, &hfffcee, &hfffbee, &hfffbee, &hfffbed, &hfffaed, &hfffaed, &hfffaed, &hfffaec, &hfffaed, &hfff7e8, &he7d5cf, _
	&hfeecde, &hfffaeb, &hfffaed, &hfffaed, &hfffbee, &hfffbee, &hfffbee, &hfffbee, &hfffbed, &hfffaed, &hfffaed, &hfffbee, &hfffaed, &hfffaed, &hfffaeb, &hfeece4, _
	&hfff4e6, &hfffbed, &hfff9ec, &hfffaed, &hfffaed, &hfffbee, &hfffbee, &hfffbee, &hfffbed, &hfffaed, &hfffbed, &hfffbee, &hfffbee, &hfffaec, &hfffaeb, &hfff8ef, _
	&hfff7e9, &hfffcee, &hfff9ec, &hfffaed, &hfffaed, &hfffbee, &hfffbee, &hfffbee, &hfffbed, &hfffbed, &hfffbed, &hfffbef, &hfffbee, &hfff9ec, &hfff8ea, &hfff8f0, _
	&hfff8e9, &hfffced, &hfff9ec, &hfffaed, &hfffaed, &hfffbed, &hfffbee, &hfffbee, &hfffbee, &hfffbed, &hfffaed, &hfffbef, &hfffbef, &hfffaed, &hfffaed, &hfff9f1, _
	&hfff6e8, &hfffbed, &hfffaec, &hfffaed, &hfffaed, &hfffaed, &hfffbed, &hfffaed, &hfffbed, &hfffaed, &hfffaed, &hfffbef, &hfffbee, &hfffaed, &hfffaed, &hfff8f0, _
	&hfff7eb, &hfffbed, &hfffaed, &hfffaed, &hfff9ec, &hfffaed, &hfffaed, &hfffaed, &hfffbed, &hfffaed, &hfffaed, &hfffbef, &hfffbee, &hfffaed, &hfff9ec, &hfff7ef, _
	&hfffaec, &hfffbed, &hfff9ec, &hfffaed, &hfff9ec, &hfffaed, &hfffaed, &hfffaed, &hfffaed, &hfffaed, &hfffaed, &hfffaee, &hfffbee, &hfffbee, &hfff9eb, &hfff9ef, _
	&hfff8ee, &hfffaee, &hfff8eb, &hfffaed, &hfffaed, &hfffaed, &hfffaed, &hfffaed, &hfffaed, &hfffaed, &hfffaed, &hfffaee, &hfffbee, &hfffbee, &hfff9eb, &hfffbf0, _
	&hfff9ee, &hfffbf0, &hfffaed, &hfffbee, &hfffbee, &hfffbee, &hfffbee, &hfffbee, &hfffaed, &hfffaed, &hfffbed, &hfffbee, &hfffced, &hfffeef, &hfffbee, &hfffcee, _
	&hfffcf0, &hfffbf1, &hfffced, &hfffeef, &hfffcee, &hfffcee, &hfffcee, &hfffbee, &hfffbed, &hfffbed, &hfffbed, &hfffbee, &hfffdef, &hfffff0, &hfff6ec, &hfff8ea, _
	&hfff7ec, &hfffff3, &hffffee, &hfffdee, &hfffcee, &hfffbee, &hfffcee, &hfffcee, &hfffbee, &hfffcee, &hfffbed, &hfffaee, &hfffcee, &hfffbeb, &hfffff0, &hfffff0, _
	&hfffced, &hfffff1, &hfffff2, &hfffff2, &hfffff1, &hfffcef, &hfffbee, &hfffcee, &hfffbed, &hfffaee, &hfff9ee, &hfff9ee, &hfffef1, &hfff7eb, &hfffff5, &hfffff9, _
	&hf1f3e9, &hfaf8f3, &hfff7ee, &hfff9ed, &hfff9f0, &hfffaf2, &hfffcf1, &hfffdf1, &hfffcf1, &hfffcf1, &hfffcf0, &hfffbef, &hfffaf0, &hfffbea, &hfffff5, &hedeeea, _
	&hfcbc6c, &hffffa0, &hffff9a, &hf5a368, &hb88059, &h9e7959, &ha17351, &h9f7758, &ha77e5f, &hae8166, &hc79678, &he4ac88, &heca376, &hffff93, &hffffaf, &heecd9b, _
	&hffff8e, &hffffd9, &hfff68a, &he7b87c, &hd3a573, &ha87f58, &h8f6a4a, &h916b4d, &h957152, &h9e7b5f, &hbc9475, &hcba283, &hc5956d, &he19c65, &hffffb0, &hffff94, _
	&hffffae, &hffff88, &hceeb70, &hc89b6c, &hfff987, &hffd478, &h9a6a41, &h8e5d3d, &h926243, &h976f53, &hc19976, &hb88f6d, &h7a583a, &h7e512f, &h955b31, &he0a056, _
	&hffd268, &hd9954c, &hd7b45a, &hb5914e, &heac060, &hffc067, &h916037, &h885233, &h955436, &ha46748, &hcb9570, &ha37153, &h643f27, &h57351d, &h6e3e23, &h765733, _
	&h926e2b, &ha86b35, &hffbd69, &he59b53, &h966d34, &hc98e4a, &h906437, &h7a502f, &h834f2e, &h8b593a, &hc9966e, &h88593d, &h613822, &h4d2f1d, &h573622, &h573e26, _
	&h53562d, &h57451f, &h835b2e, &ha56f37, &h7c5423, &hae7739, &ha56d3a, &h83532d, &h7c4a26, &h8d5130, &hc57e52, &h8a522f, &h67371e, &h482713, &h452714, &h48301d, _
	&h517221, &h504b1b, &h583f1c, &h785026, &h6e481a, &h634215, &h764a21, &h814f27, &h824e26, &h884524, &h924d2c, &h8d5029, &h9d5e34, &h552e18, &h3e210f, &h3c2815, _
	&h3a3b1a, &h48421b, &h3e3110, &h452d0c, &h5c3910, &h5d380e, &h58310d, &h603715, &h6f401d, &h6b3416, &h7f3f1c, &hffb25c, &hffd773, &h854a26, &h562d16, &h5c3921, _
	&h443216, &h5a491f, &h36320c, &h3c2b09, &h51330a, &h563108, &h502906, &h4c2807, &h4d2a0b, &h562a0c, &haa5f2b, &hffffa9, &hfff37e, &ha3612e, &h815125, &hb17841, _
	&h5f4420, &h4e3815, &h32290a, &h352307, &h412705, &h432704, &h402003, &h3e2004, &h3f2107, &h4c280c, &ha95c2b, &hffb459, &haa6733, &h844d25, &h935d2a, &hffac5a, _
	&hb1783b, &h936133, &h33240a, &h2a1804, &h2c1603, &h301803, &h371903, &h391b03, &h391d06, &h3e2209, &h623013, &h8c4021, &h5b3116, &h583619, &h8b5d2b, &hffba5f, _
	&h976530, &h9f6836, &h2d2008, &h221404, &h251002, &h261002, &h2e1302, &h311703, &h301905, &h351b07, &h411c08, &h471d0c, &h3b1d0b, &h3f2910, &h9b6a34, &heca854, _
	&h4b341b, &h4c3214, &h211804, &h1d1103, &h201002, &h211002, &h241102, &h261203, &h251304, &h2b1505, &h361908, &h351a0b, &h321a09, &h3b280d, &he1944c, &hfdb555, _
	&h32240f, &h3f270a, &h221604, &h1b0f03, &h1c0e02, &h1c0e02, &h1d0e02, &h1d0d02, &h1b0c02, &h1c0c02, &h210d03, &h241005, &h291405, &h453012, &hfff57b, &hffff90, _
	&h28240e, &h37250b, &h271906, &h1b0e02, &h1a0c01, &h190c01, &h190b01, &h180a01, &h160901, &h150701, &h160701, &h190b03, &h2b1707, &h583c1a, &hffff87, &hffff8f, _
	&h34300f, &h4f3614, &h37220c, &h1c0e02, &h170b01, &h170a01, &h170901, &h160901, &h160901, &h150801, &h150802, &h160c04, &h31200d, &h7b5a2c, &hffffa7, &hfffe7e, _
	&h4a4b17, &h85652a, &h402b0f, &h1d0f02, &h160b02, &h150901, &h140801, &h140801, &h150801, &h150801, &h150802, &h150a03, &h271806, &h4e3e1c, &hfff871, &hfffd7a, _
	&h727d27, &he1b14c, &h634321, &h211103, &h160b02, &h130801, &h120701, &h130701, &h140802, &h150902, &h150a02, &h150a03, &h211104, &h463917, &hffdb6d, &hffff8f, _
	&h5f641e, &h826d2a, &h4f3315, &h211003, &h150902, &h110701, &h100601, &h100601, &h130802, &h160a02, &h160b02, &h150b04, &h1f0d03, &h513c1d, &hffd474, &hffff91, _
	&h443713, &h745a24, &h8e5e32, &h2a1907, &h150802, &h0f0501, &h0e0401, &h0e0501, &h100601, &h130802, &h140a02, &h140a03, &h231106, &ha26e37, &hffff97, &hffb969, _
	&h464015, &hbd8c3a, &he4954f, &h35230e, &h160a02, &h0d0401, &h0b0401, &h0b0401, &h0b0401, &h0d0501, &h0f0601, &h120802, &h241406, &h855b37, &hffb965, &h875634, _
	&h2c2911, &h5f4a1c, &h593e1d, &h211505, &h100702, &h090301, &h080300, &h080300, &h080200, &h090300, &h0b0301, &h0f0501, &h180b02, &h422810, &h44240d, &h2f1d0e, _
	&h14160e, &h221705, &h1f1103, &h0f0701, &h090401, &h070200, &h070200, &h070200, &h070200, &h070200, &h090301, &h0e0602, &h110802, &h220d04, &h200e03, &h1a1109, _
	&h030902, &h0e0901, &h100701, &h0b0401, &h080301, &h060200, &h060200, &h060200, &h060200, &h060200, &h090301, &h0f0702, &h110902, &h110602, &h0b0301, &h0b0804, _
	&h030803, &h0b0a02, &h0d0802, &h0c0501, &h080301, &h060200, &h060200, &h060200, &h060200, &h060200, &h0a0401, &h120802, &h140a02, &h100602, &h050201, &h080704, _
	&h010502, &h070701, &h0c0602, &h0c0501, &h090401, &h070300, &h060200, &h060200, &h060200, &h060301, &h0b0501, &h130902, &h150b02, &h0f0601, &h050301, &h070504, _
	&h010102, &h040301, &h090401, &h0c0501, &h0c0401, &h090301, &h080301, &h080301, &h070301, &h070301, &h0b0501, &h110802, &h110802, &h0b0501, &h040200, &h050304, _
	&h010102, &h030201, &h060201, &h080301, &h0a0401, &h0a0401, &h0a0401, &h0b0401, &h0b0401, &h0b0501, &h0c0501, &h0d0501, &h090401, &h040201, &h030100, &h030103, _
	&h010102, &h030201, &h030100, &h030100, &h050201, &h070301, &h0b0401, &h0f0501, &h110602, &h100602, &h0d0401, &h070301, &h040200, &h030100, &h020000, &h020001, _
	&h000001, &h010101, &h020101, &h020000, &h020000, &h020000, &h040100, &h090301, &h0c0401, &h0a0401, &h060200, &h030100, &h020000, &h010000, &h010000, &h000000, _
	&h000001, &h000100, &h010101, &h010000, &h000000, &h000000, &h000000, &h010100, &h020100, &h030200, &h020100, &h010000, &h000000, &h000000, &h000000, &h000001, _
	&h000000, &h000000, &h010101, &h010101, &h010100, &h000000, &h010100, &h010100, &h010100, &h010100, &h010100, &h010100, &h010000, &h000000, &h000000, &h000001, _
	&h1c170e, &h1b1810, &h211c0f, &h2c2416, &h3c2c1e, &h39271a, &h392215, &h2e1d14, &h2d1c10, &h332014, &h302015, &h332215, &h392516, &h3c250d, &h3d1309, &h271c13, _
	&h191a0b, &h1e190e, &h221d11, &h302d1c, &h3e3224, &h3f3225, &h3a2619, &h3b2315, &h412514, &h3e2719, &h3a2a1e, &h3a2b1d, &h3a2819, &h341e0f, &h290e05, &h362512, _
	&h1c1b0d, &h1e1b0e, &h282114, &h3e3423, &h514230, &h503d2d, &h543725, &h974421, &h9e5125, &h743f24, &h4e3421, &h412e1d, &h442e1b, &h2a170a, &h230b04, &h190c04, _
	&h171b0b, &h1d1e0d, &h2c2312, &h42331f, &h5c452c, &h573d24, &h5c341e, &h8b4422, &haa4c29, &ha75531, &h845632, &h48311c, &h2b1809, &h220f05, &h1b0802, &h110602, _
	&h171208, &h1b1307, &h281807, &h3e2509, &h6f491c, &h8f5a1c, &h814919, &h603212, &h6a3515, &h75421e, &h985c28, &h59371a, &h271205, &h1c0903, &h130501, &h080301, _
	&h160f03, &h1b1003, &h221102, &h381c02, &h542f06, &hc7731b, &hab621c, &h562e0b, &h552f0c, &h814b18, &he28c2d, &h774715, &h2d1103, &h190701, &h100401, &h070201, _
	&h151001, &h231402, &h251303, &h522805, &h2f1703, &h371c04, &h5a330b, &h8c5517, &h975e1b, &h683c0f, &h8d5515, &h4c2807, &h250c02, &h170601, &h100401, &h080301, _
	&h180a01, &h341b04, &h1f0f02, &h220e02, &h261102, &h200e01, &h4d2705, &hab651e, &hbd7321, &h492807, &h311904, &h220f02, &h1f0801, &h160601, &h100400, &h0a0301, _
	&h300e01, &h3b1b04, &h1c0c02, &h170901, &h1b0a01, &h1c0c01, &h281302, &h412405, &h452706, &h321503, &h270d02, &h200b01, &h1b0801, &h150601, &h110400, &h0c0301, _
	&h0f0601, &h160c02, &h150a02, &h140801, &h150801, &h170901, &h1a0b01, &h180b01, &h1a0c01, &h200d01, &h2c0b01, &h240a01, &h190801, &h130601, &h100400, &h0c0401, _
	&h050201, &h060200, &h0d0501, &h100601, &h120601, &h130701, &h150801, &h160901, &h170901, &h1b0a01, &h220b01, &h1e0a01, &h150801, &h130601, &h110501, &h0c0401, _
	&h040300, &h060300, &h0a0400, &h0f0501, &h100601, &h110601, &h120601, &h130701, &h140701, &h150801, &h190901, &h1a0901, &h170801, &h190701, &h180601, &h100601, _
	&h030300, &h050300, &h080400, &h0d0400, &h0f0500, &h0f0500, &h0f0500, &h100500, &h100501, &h110601, &h130701, &h160701, &h190801, &h200a02, &h1c0a02, &h130802, _
	&h030201, &h050300, &h070300, &h0a0300, &h0c0400, &h0c0400, &h0c0400, &h0d0400, &h0e0400, &h0e0500, &h100501, &h130601, &h170902, &h1b0a02, &h190b02, &h1a0f03, _
	&h030301, &h060301, &h070301, &h080300, &h0a0300, &h0a0300, &h0a0300, &h0b0300, &h0b0400, &h0c0400, &h0e0500, &h110601, &h150902, &h170902, &h1a0c02, &h241605, _
	&h050401, &h0d0801, &h0f0702, &h0a0401, &h090300, &h090300, &h090300, &h090300, &h090300, &h090300, &h0b0300, &h0d0501, &h110701, &h120701, &h150801, &h190d03, _
	&h060701, &h110b02, &h170c02, &h0c0501, &h080300, &h090300, &h080300, &h080300, &h080200, &h070200, &h070200, &h080300, &h0b0401, &h0d0501, &h110601, &h120901, _
	&h070702, &h100b02, &h170d03, &h0c0601, &h070300, &h070300, &h070300, &h070200, &h070200, &h070200, &h070200, &h070200, &h090301, &h0b0300, &h160902, &h180c03, _
	&h060602, &h0f0a02, &h130b02, &h0a0501, &h060300, &h060200, &h070200, &h070200, &h070200, &h070200, &h060200, &h060200, &h080300, &h0b0401, &h1c0d03, &h2e1b08, _
	&h050501, &h0a0701, &h0d0701, &h070401, &h050300, &h050200, &h060200, &h060200, &h060200, &h060200, &h060200, &h060200, &h070200, &h0c0401, &h221004, &h402710, _
	&h040401, &h050400, &h060300, &h040300, &h040200, &h040200, &h050200, &h050200, &h050200, &h050200, &h050200, &h050200, &h060200, &h0e0501, &h220f05, &h311f0d, _
	&h030200, &h030200, &h030200, &h030100, &h030100, &h030100, &h040100, &h040100, &h040100, &h030100, &h030100, &h040100, &h050200, &h0b0401, &h140702, &h221506, _
	&h020200, &h030100, &h020100, &h030100, &h030100, &h030100, &h030100, &h030100, &h030100, &h030100, &h030100, &h030100, &h040100, &h080200, &h0b0301, &h0e0602, _
	&h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h030100, &h030100, &h050200, &h070201, &h070301, _
	&h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h030100, &h040100, &h060201, &h050201, _
	&h010100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h030100, &h040100, &h050200, &h040201, _
	&h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h030100, &h030100, &h040200, &h040201, _
	&h010100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h030100, &h030100, &h040100, &h040201, _
	&h010100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h030100, &h030100, &h030101, _
	&h010000, &h010100, &h010100, &h020100, &h010100, &h010100, &h010100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h030100, &h030100, &h030100, _
	&h000000, &h000000, &h010100, &h010100, &h010100, &h010100, &h010100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020101, _
	&h000000, &h000000, &h010000, &h010100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h010100, &h010101, _
	&h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, _
	&h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, _
	&h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, _
	&h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h010000, &h010000, &h010000, _
	&h010000, &h010000, &h010000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h010000, &h010000, &h010000, _
	&h010000, &h010000, &h010000, &h010000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h010000, &h010000, &h010000, &h010000, _
	&h010100, &h010100, &h010100, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h020000, &h020000, &h010000, _
	&h010100, &h010100, &h020100, &h020100, &h010100, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h020000, &h020000, &h010000, &h000000, _
	&h010100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h010000, &h010000, &h010000, &h020000, &h030100, &h030100, &h010000, &h000000, _
	&h000100, &h020100, &h030100, &h030100, &h030100, &h030100, &h030100, &h020100, &h020100, &h020100, &h020000, &h030100, &h040100, &h040100, &h020000, &h000000, _
	&h010100, &h020100, &h030100, &h030100, &h030100, &h030100, &h030100, &h030100, &h030100, &h030100, &h030100, &h030100, &h050100, &h050100, &h030100, &h010000, _
	&h010100, &h020100, &h030100, &h030100, &h030100, &h030100, &h030100, &h030100, &h030100, &h030100, &h030100, &h040100, &h050100, &h060200, &h040100, &h010000, _
	&h010100, &h020100, &h030100, &h030100, &h030100, &h030100, &h030100, &h030100, &h030100, &h030100, &h040100, &h040100, &h050200, &h060200, &h030100, &h010000, _
	&h010100, &h020100, &h030200, &h030100, &h030100, &h030100, &h030100, &h030100, &h030100, &h030100, &h030100, &h040100, &h040200, &h050201, &h030100, &h010100, _
	&h020200, &h030200, &h040200, &h040200, &h030200, &h030100, &h030100, &h030100, &h030100, &h030100, &h030100, &h030100, &h040200, &h050201, &h040100, &h020100, _
	&h020301, &h040300, &h050301, &h040200, &h040200, &h040100, &h030100, &h030100, &h030100, &h030100, &h030100, &h030100, &h040200, &h060201, &h050200, &h030100, _
	&h020301, &h060401, &h080401, &h060301, &h050200, &h040200, &h040200, &h030100, &h030100, &h030100, &h030100, &h040100, &h050200, &h070301, &h070201, &h050201, _
	&h030501, &h0c0a01, &h180d03, &h0b0501, &h060301, &h050201, &h050200, &h040200, &h040200, &h040200, &h040200, &h050200, &h060201, &h080301, &h090301, &h060301, _
	&h030601, &h161603, &h442e0e, &h180c03, &h080401, &h060301, &h060301, &h060301, &h060301, &h060201, &h060301, &h060301, &h080301, &h0b0501, &h0e0602, &h0c0602, _
	&h040801, &h161703, &h402d0e, &h211205, &h0c0501, &h090401, &h0b0501, &h0d0601, &h0c0501, &h0b0501, &h0a0401, &h0a0401, &h0d0602, &h150a03, &h1a0c03, &h110802, _
	&h030601, &h0f0e02, &h2b1c08, &h2d1808, &h180b03, &h130902, &h150a02, &h140902, &h110702, &h0f0602, &h120802, &h180c03, &h1e1005, &h241506, &h271406, &h120902, _
	&h030601, &h141103, &h523415, &h613819, &h533216, &h3c240e, &h221205, &h170c03, &h120802, &h140a03, &h281808, &h412911, &h38240e, &h271808, &h201005, &h090401, _
	&h050901, &h141002, &h69441d, &ha26731, &hffb05a, &h9c6433, &h321c0a, &h1d1004, &h180c03, &h201205, &h503317, &hb77a3c, &hbe803e, &h5c3e1c, &h281607, &h070301, _
	&h060902, &h150f03, &h64411c, &hb57538, &hffdf72, &hd78c48, &h452912, &h201105, &h1f1205, &h311f0b, &h86592c, &hffce68, &hf8a150, &h5b3a1a, &h271507, &h0a0401, _
	&h070702, &h140c03, &h704722, &hb9753a, &hffb65e, &he7984d, &h593819, &h261607, &h291808, &h412a11, &h986531, &hffff89, &hffaa58, &h613f1d, &h271607, &h0b0502, _
	&h060402, &h150b03, &h91592f, &hc07b3e, &hb97a3e, &h8f5f2e, &h4f3315, &h2f1c0a, &h2f1d0a, &h412810, &h6c431e, &hea994c, &hffcf6d, &ha96f36, &h301c0a, &h0b0502, _
	&h050301, &h140803, &h804d27, &hbd763b, &h8d5c2d, &h66421e, &h442a11, &h321e0b, &h301d0a, &h3c240e, &h4d2c12, &h7c4f24, &hb57b3c, &h945c30, &h0e0502, &h070301, _
	&h040201, &h0a0301, &h201005, &h3b210d, &h6c421f, &h603c1b, &h462c12, &h35200c, &h2f1d0a, &h39230e, &h492d13, &h6d4721, &h59381a, &h130703, &h070201, &h050201, _
	&h030201, &h050201, &h060301, &h0d0402, &h1c0c04, &h361f0d, &h3e2710, &h37220c, &h35200b, &h33200b, &h351f0b, &h221305, &h0e0502, &h050101, &h040100, &h040101, _
	&h010100, &h020100, &h020100, &h020100, &h040100, &h060201, &h0d0502, &h241206, &h351e0a, &h2b1907, &h1b0e03, &h070301, &h030100, &h020000, &h020000, &h010100, _
	&h000100, &h010100, &h010000, &h010100, &h020100, &h020100, &h020100, &h040100, &h170a03, &h1e1105, &h110702, &h020100, &h010000, &h010000, &h010000, &h010000, _
	&h000000, &h010000, &h010000, &h020100, &h050301, &h030201, &h020100, &h020100, &h040100, &h050101, &h050100, &h020000, &h020000, &h010000, &h010000, &h000000, _
	&h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, _
	&h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, _
	&h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, _
	&h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, _
	&h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, _
	&h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, _
	&h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, &h000000, _
	&h000000, &h000000, &h000000, &h000000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h000000, &h010000, &h000000, &h010000, &h010000, &h000000, _
	&h000000, &h000000, &h000000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h000000, _
	&h000000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h000000, _
	&h000000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h000000, _
	&h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, _
	&h010100, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, _
	&h010100, &h010100, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, _
	&h020100, &h020100, &h010100, &h010100, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, &h010100, &h020100, &h020100, _
	&h020100, &h020100, &h020100, &h010100, &h010100, &h010100, &h010100, &h010000, &h010000, &h010000, &h010000, &h010000, &h010100, &h020100, &h020100, &h020100, _
	&h030100, &h030100, &h020100, &h020100, &h020100, &h020100, &h010100, &h010100, &h010100, &h010100, &h010100, &h010100, &h020100, &h020100, &h030100, &h030100, _
	&h030200, &h030200, &h030100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h030100, &h030200, _
	&h040201, &h040200, &h030200, &h030100, &h030100, &h030100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h020100, &h030100, &h040100, &h040200, _
	&h050201, &h060201, &h050201, &h050201, &h040200, &h040200, &h040100, &h040200, &h040200, &h040200, &h040200, &h040200, &h040200, &h040200, &h050200, &h040201, _
	&h050301, &h070301, &h070301, &h090401, &h090401, &h070301, &h070301, &h070301, &h070301, &h070301, &h070301, &h080301, &h070301, &h070301, &h080301, &h060201, _
	&h070301, &h0a0401, &h0d0601, &h100802, &h120902, &h100702, &h0c0501, &h0a0401, &h0b0401, &h0d0501, &h100802, &h150c04, &h120903, &h0f0702, &h0f0602, &h080301, _
	&h0b0601, &h100702, &h150a02, &h1a0e03, &h1d1004, &h1d1004, &h160b03, &h0c0501, &h0c0501, &h100602, &h1b1005, &h291d0f, &h23190c, &h120b04, &h0c0502, &h060301, _
	&h0e0702, &h120902, &h190e03, &h221405, &h271606, &h291707, &h211204, &h100601, &h0d0501, &h100702, &h1d1005, &h281a0b, &h261809, &h170d04, &h0d0502, &h060201, _
	&h0c0602, &h110802, &h190e03, &h251505, &h2c1907, &h2c1808, &h241405, &h110702, &h0d0501, &h100702, &h1c1005, &h2a1909, &h2a1808, &h1d0e04, &h100702, &h060301, _
	&h090501, &h0f0701, &h1b0e03, &h261505, &h2c1807, &h2b1808, &h241305, &h100702, &h0c0401, &h0f0601, &h1a0e04, &h261607, &h231306, &h1e0e04, &h160903, &h080301, _
	&h070301, &h0b0501, &h1b0f03, &h281606, &h2c1807, &h261506, &h1b0d03, &h0d0501, &h0a0401, &h0d0501, &h130902, &h1a0e04, &h1d1004, &h1d0e04, &h160902, &h070301, _
	&h030200, &h050200, &h0d0601, &h1c0e03, &h1d0e04, &h150a02, &h0d0501, &h090301, &h070201, &h090301, &h0e0501, &h140a03, &h190d03, &h130802, &h090301, &h040200, _
	&h020100, &h020100, &h030100, &h060201, &h080301, &h070201, &h060200, &h050200, &h050200, &h060200, &h090301, &h090301, &h060201, &h050200, &h040100, &h030100, _
	&h010100, &h010100, &h010100, &h010100, &h020100, &h020100, &h030100, &h040100, &h040100, &h040100, &h030100, &h020100, &h020100, &h020100, &h020100, &h020100, _
	&h010100, &h010000, &h010000, &h010000, &h010000, &h010000, &h020100, &h030100, &h030100, &h020100, &h010000, &h010000, &h010000, &h010000, &h010000, &h010000, _
	&h000000, &h000000, &h000000, &h010000, &h010000, &h010000, &h010000, &h020100, &h020100, &h020100, &h010000, &h010000, &h010000, &h010000, &h000000, &h000000 _
)


'VPW Version Log
'===============
'0.001 - Sixtoe - Initial Table Build,
'0.002 - Sixtoe - Added Roth drop & standup targets (with temporary target prims), Plungers Sorted, Kickers organised and prepped for playfield mesh, tweaked ramps.
'0.003 - Sixtoe - Added primitive collidable ramps and playfield, implemented kickers and scoop, fixed triggers, fixed lots of little things, now fully playable, flipper power increased a lot as G ramp is now very steep, scoop and right VUK could probably do with some love.
'0.004 - Sixtoe - Changed ramp entrances to gates, sorted out triggers, sorted out the blocker walls, realigned physical objects to new blueprint, remade kicker, scoop and vuk, much happier now with them, redid primitive ramp ends
'0.005 - Niwak - Imported all graphics from Blender, linked the graphics with the physics parts
'0.006 - Sixtoe - Added fleep audio and ramprolling sounds, hooked up movable targets (but not lightmaps yet)
'0.007 - Niwak - Graphic fixes, drop targets, option menu, room lighting level, dynamic ball brightness
'0.008 - Niwak - Add dynamic ball color from bakes, fix drop target lightmaps, add PinMame 3.7 PWM flasher with tinting
'0.009 - Niwak - Better lights (incandescent from VPX & PinMame), dual bake for flipper bats, better inserts translucency, ramp exit, slightly runed scoop, increased kickback strength, animate trap door
'0.010 - Sixtoe - Fixed physical shooter lane ramp to match new visible one, upped shooter power.
'0.011 - Niwak - Remove refraction (too buggy for the time being), use new 10.8 ball mapping, add low res playfield for ball reflection, use new camera modes
'0.012 - apophis - Fixed flipper triggers. Fixed flipper animations. Fixed bumper ring animation.
'0.013 - Niwak - Readd refractions, fix lighting making transparent part dark, change ball image, adjust ramp transparency, add reflections
'0.014 - Sixtoe - Mystery scoop area rebuilt (including playfield mesh), flipper inlanes tweaked to allow post passes, inlane material changed to plastic, left kicker upped in power.
'0.015 - Niwak - Remade bumper caps, new batch using latest toolkit (fix unwrapped baked parts), overall script improvments, fix mods (HiRez blades and back panel)
'0.016 - Niwak - Fix flashers fading, add insert reflection on balls, fix scoop trigger visible, add outlane difficulty, add ambient ball shadow, fix gate orientation
'0.017 - Niwak - Fix missing base GI, fix wrong ball shadows, made parts and GI0 static for quality and performance
'0.018 - Niwak - Basic animation for bumper skirts, add blurry ball reflection on playfield
'0.019 - iaakki - Ramp flap and bump sounds added, Gate sounds improvements, Plunger sounds reworked, kickback sounds added (sw16, sw24, RampTriggerPE), gi relay sound added
'0.020 - iaakki - Failsafe added to plunger usage, K1 Relay sounds added to Sol10, GI sound tuned, Kickback sound bug fixed
'0.021 - Sixtoe - Adjusted G and R physical ramps so they act correctly, turned down rose plunger and increased snake ramp friction slightly, adjusted left side orbit wall and end rubber post, adjusted band member entrance walls and rubbers, adjusted mission approach walls, turned down band member VUK as much as it will allow.
'0.022 - iaakki - captive ball returns back in place if it escapes, left and right vuk ball wiggling removed, slight change to sling threshold
'0.023 - iaakki - added plungerlane grooves to plungerlanes
'0.024 - iaakki - Fixed fast flips and ball roll sounds when ball in vuks
'0.025 - Primetime5k - added staged flipper support, staged flipper option in options menu
'0.026 - iaakki - Added FlipsEnabled solenoid to #23
'0.027 - iaakki - fixed duplicated sound triggers, improved orb wall sounds
'0.028 - TastyWasps - Draft VR Cab addded
'0.029 - TastyWasps - VR Room addded
'0.030 - Sixtoe - VR cabinet reconfigured, front and side artwork redone a bit, new temp backglass, vr blocker walls installed, gun and rose shooters added.
'0.031 - Niwak - New render batch with improved flashers and GI, fixed red insert colors, better bumper socket anim, new bumper cap, fixed sling anim, adjusted ball light
'0.032 - Upload mistake
'0.033 - TastyWasps - Guitar added to VR Room, VR flipper animations
'0.034 - Niwak - Fix snake plunger material, improve ball image, use new insert reflections, add GI bulb reflection, adjust GI light positions
'0.035 - Niwak - Fix blueish ball, ball shadow, add some reflections, adjusted refraction thickness
'0.036 - leojreimroc - VR Backglass
'RC1 - Sixtoe - Removed old toolkit rails, lockdown and associated code, added new pincab_rails, added new l55 snake ramp lamp primitive, added top and bottom glass heights, animated rose plunger, added flipper cradle code, added volume min code, turned off debugs, added inlane slowdown, fixed gate multiple sounds, numerous other fixes.
'RC2 - apophis - Updated LightLevel code, including affect on VR room. Fixed drifting VR button bug. Fixed FlipperCradleCollision call. Fixed TS sling correction call. Added SetLocale 1033. Set "hide parts behind" for BM_PF. 