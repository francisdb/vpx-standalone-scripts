Option Explicit

'|============================================================================|
'|     _                _                _  _  _       _ _     _              |
'|    | |              (_)              | || || |     | | |   | |             |
'|     \ \   ____ _   _ _ ____   ____   | || || | ____| | | _ | | ____ ____   |
'|      \ \ / _  | | | | |  _ \ / _  |  | ||_|| |/ _  | | |/ || |/ _  )  _ \  |
'|  _____) | ( | |\ V /| | | | ( ( | |  | |___| ( ( | | | ( (_| ( (/ /| | | | |
'| (______/ \_||_| \_/ |_|_| |_|\_|| |   \______|\_||_|_|_|\____|\____)_| |_| |
'|                             (_____|                                        |
'|                                                                            |
'|============================================================================|
Const TABLE_VERSION = "0.3.0-alpha"

' Project lead / Game Code / Game Rules / Music: Arelyel Krele
' Initial table layout: DonkeyKlonk
' Voice actors: Graham Growat, Harmony, Smaug
' 3D objects prep work: Remdwaas
' Light State Controller: Flux
'
' Bug Hunters: PinStratsDan, bietekwiet, apophis, iaakki, bennaboo, oqqsan,
' Wylte
'
' Additional credit for libraries used listed in the respective sections of the 
' script.

' CHANGELOG is at the bottom of the script.
'
' RELEASE NOTE: Prior to every public release, set usePUP to false, 
' UseFlexDMD to 0, ScorbitActive to 0, skip initialization to false, all debug 
' section consts to false, debug mode to 1000, skip dialog to false, allow 
' zen to true, allow cooperative play to true, and unless this is a
' tournament release, tournament mode to 0 and global seed to Null.
'
'******************************************************************************
' LICENSE STATEMENT
' * Do not mod this table without the permission of Arelyel Krele.
' * Do not distribute this table outside of the Digital Pinball Network 
'   (VP Universe, VP Forums, Pinball Nirvana, Rogue Pinball, VPDB).
' * Do not distribute this table in a table pack or archive of multiple tables.
' * Do not charge money to play this table. This includes arcade or tournament
'   cover charges.
' * Do not modify the code so the table accepts or simulates "credits" / money.
' * If you paid money in any way to get or play this table, GET A REFUND!
' * If you were given this table from a purchase you made, GET A REFUND!
' * Please report violations to Arelyel Krele and the respective sites.
'******************************************************************************
'
'******************************************************************************
' STREAMING STATEMENT
' This table should be safe in its entirety to stream or post videos of 
' gameplay. The music is original*. The sound effects are either original* or
' Creative Commons. And the artwork is either original, Creative Commons, or AI
' generated (to be replaced by original or CC art later). If you received a
' copyright or content ID notice, please notify Arelyel Krele immediately.
' *made in MAGIX Music Maker with purchased loops.
'******************************************************************************
'
'******************************************************************************
' === TABLE OF CONTENTS ===
' Search the 4-letter code to easily jump to that section!
' Prefix W is specific for Wallden; Z is a library.
'
' WOPT: Basic configurable options
'
' --- Spoilers below this point!!! ---
'
' WCON: Constants
'	- WDOF: Switch Assignments
' WTRK: Game play tracking variables
' WQUE: Queues / priority assignments
' ZERR: Error logs
' WDMD: Display (Pinup and JPSalas)
' ZFLP: Flipper correction and tricks
' ZTAR: Target bouncer
' ZPHY: Physics dampeners
' ZRDT: Drop Targets
' ZSTD: Stand-up targets
' ZBAL: Ball rolling and drop sounds
' ZRMP: Ramp rolling sounds
' ZMCH: Fleep mechanical sounds
' ZDRN: Drain, Trough, and Ball Release
' ZSLG: Slingshot animations
' WLHT: Supporting functions for lights
' ZRBW: Rainbow lights
' ZQUE: VPW Advanced Queuing System
' ZDAT: VPW Data Manager
' ZCON: VPW Constants
' ZCLK: VPW Clocks
' ZMUS: VPW Music
' ZTRK: VPW Tracks / Fading
' ZLAM: nFozzy Lamps
' ZLSC: Light State Controller
'	- WLSC: Custom Light Sequences
' ZSHA: Dynamic Ball Shadows
' ZDOM: Flupper Domes
' ZSCO: Scorbit
' WSEG: Segment Displays
' WHLP: Wallden Helper Functions
' WSOL: Solenoids
' WSWI: Switch Handler
' WROU: Wallden Routines
'	- WNUG: Nudging
'	- WCLK: Clock Routines
'	- WMDE: Table Modes
'	- WGEM: Gems and Powerups
'	- WHPD: +HP Drop Targets
'	- WBON: Bumper Inlanes / Bonus X
'	- WBSM: Blacksmith
' WVER: Change Log
'******************************************************************************

Randomize

'*****************************************
'   WOPT: BASIC CONFIGURABLE OPTIONS
'   Modify as desired
'*****************************************

'-----------------------------------------
' *** Relative Sound volumes ***
'-----------------------------------------
' Table
Const VolumeDial = 0.5			'Mechanical sounds from Fleep. Float between 0 and 1.
Const BallRollVolume = 0.2		'Level of ball rolling volume. Float between 0 and 1.
Const RampRollVolume = 0.2		'Level of ramp rolling volume. Float between 0 and 1.

' Backglass
Const VOLUME_BG_MUSIC = 0.5		'Level of volume for the gameplay (background) music. Float between 0 and 1.
Const VOLUME_SFX = 0.8 			'Level of volume for sound effects. Float between 0 and 1.
Const VOLUME_CALLOUTS = 1.0		'Level of volume for voice callouts and narration. Float between 0 and 1.

'-----------------------------------------
' *** PinUP Player ***
'-----------------------------------------
Const usePUP = False		'Set to true to enable PinUP (TODO: not currently supported for Wallden; maybe later? ;) ). Make sure you have the savingWallden PUP-Pack in your PUPVideos folder.

Dim PuPDMDDriverType
PuPDMDDriverType = 2   	'0=LCD DMD, 1=RealDMD, 2=FULLDMD (large/High LCD). TODO: implement
Dim useRealDMDScale
useRealDMDScale = 1    	'0 or 1 for RealDMD scaling.  Choose which one you prefer. TODO: Implement
Dim useDMDVideos
useDMDVideos = True   	'true or false to use DMD splash videos. TODO: Implement

'-----------------------------------------
' *** FlexDMD / JPSalas DMD ***
'-----------------------------------------
'A JPSalas DMD will always be visible on
'the playfield above the flippers. You can
'enable FlexDMD which will show what the
'JP DMD shows.

Const UseFlexDMD = 0				'TODO: Seems to be broken at the moment; 0 = Do not use FlexDMD; 1 = Always use FlexDMD; 2 = Use FlexDMD when not running in desktop mode
Const FlexDMDHighQuality = False 	'Set to true if your DMD can handle 256x64 resolution opposed to the default 128x32 resolution

'-----------------------------------------
' *** Scorbit (TODO: not working ATM) ***
'-----------------------------------------
'Scorbit is an app which allows you to
'share high scores with others and
'participate in leader-boards and
'tournaments. You can also earn table
'achievements. See scorbit.io. It's free!
'Saving Wallden supports Scorbit and is
'listed in their system as a machine.

Dim ScorbitActive
ScorbitActive					= 0		' Set to 1 to enable Scorbit. In Saving Wallden, Scorbit is not initialized until the Attract Sequence begins. Ignored if usePUP = False or PUP fails to load.
Const     ScorbitShowClaimQR	= 1 	' If Scorbit is active and this is 1, then each player will see a QR code they can scan in the app to claim their slot the first time they are up.
Const     ScorbitQRSmall		= 0 	' Set to 1 when using a large HD display if you want QR codes to show in the score area (takes up less room) than the main video area.
Const     ScorbitUploadLog		= 1 	' Store local log and upload after the game is over
Const     ScorbitAlternateUUID  = 0 	' Force Alternate UUID from Windows Machine and saves it in VPX Users directory (C:\Visual Pinball\User\ScorbitUUID.dat)

'-----------------------------------------
' *** Dynamic Shadows ***
' You should probably disable these when
' running VPX 10.8 or later as 10.8 has
' native support for dynamic ball shadows.
'-----------------------------------------

Const DynamicBallShadowsOn = 0		'0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow

Const AmbientBallShadowOn = 1       '0 = Static shadow under ball ("flasher" image, like JP's)
'1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
'2 = flasher image shadow, but it moves like ninuzzu's

'-----------------------------------------
' *** Physics ***
'-----------------------------------------

const bSlingSpin = True 			'Sling corner spin feature. This is a random thing that happens with pinball.

'-----------------------------------------
' *** Debugging ***
'-----------------------------------------
Const DEBUG_SKIP_INITIALIZATION = False 	'Skip the self test sequence so the game enters attract sequence faster

' Logging (via txt file and output)
Const DEBUG_LOG_SCORES = False            	'Log scoring events (Warning! Can get long quickly)
Const DEBUG_LOG_BALL_SEARCH = False			'Log ball search operations
Const DEBUG_LOG_FLIPPER_DAMPENER = False	'Log flipper dampener to output (no txt file logging)
Const DEBUG_LOG_WIRE_RAMPS = False			'Log wire ramp errors
Const DEBUG_LOG_TROUGH = False				'Log activity with the ball trough
Const DEBUG_LOG_TRACK_FADING = False		'Log track fading (Warning! Can get long quickly)
Const DEBUG_LOG_SCORBIT = False				'Log Scorbit activity
Const DEBUG_LOG_RANDOM_NUMBERS = False		'Log info about generated random numbers from the custom Wallden tournament-friendly generator
Const DEBUG_LOG_MODES = False				'Log mode changes / statuses
Const DEBUG_LOG_SWITCHES = False			'Log switches when they are enabled / disabled (Warning! Can get long quickly)
Const DEBUG_LOG_QUEUE = False				'Log main queue statuses (Warning! Can get long quickly)
Const DEBUG_LOG_SOLENOIDS = False			'Log when solenoids fire / disengage
Const DEBUG_LOG_JP_DMD = False				'Log JP DMD activity (Warning! Can get long quickly and cause stutter)

' Data resets
Const RESET_HIGH_SCORES = False            	'Set to true, load/exit the game, and then set back to false to reset all the high scores (except tournament scores) to defaults
Const RESET_TOURNAMENT_HIGH_SCORES = False 	'Set to true, load/exit the game, and then set back to false to reset all the tournament high scores to 0
Const RESET_GAME_TRACKING = False        	'Set to true, load/exit the game, and then set back to false to reset the game tracking data in vpReg (especially after downloading the latest table version)

' Game Play
Const DEBUG_INFINITE_SHIELD = False			'Set to true for infinite ball shield (when there is 1 ball in play); disables Scorbit, high-score entries, and final wizards if enabled! TODO: REMOVE IN PRODUCTION!

' Light State Controller
Const DEBUG_COMPILE_LIGHT_SHOWS = False		'Set to true to load / compile light show files on table init.

'-------------------------------------------
' === DEBUG MODE ===
' Test a specific mode by having it
' immediately start at the start of a
' game (after the skill shot).
'
' Set this to the numerical value of
' the appropriate mode; MODE_ constants
' are further down.
'
' Set to 1000 (MODE_NORMAL) if not
' debugging a mode.
'
' Set to 0 to debug features specific to
' this development release.
'
' TODO: REMOVE IN PRODUCTION!
'------------------------------------------
Const DEBUG_MODE = 1000

'-----------------------------------------
' *** Game Settings ***
'-----------------------------------------
Const SKIP_DIALOG = False	'Set to true if you want the game to skip lore dialog to speed up game play

'-----------------------------------------
' Credits:
' Saving Wallden does not support credits;
' it is forcefully a free to play table.
' This is to discourage use of the table
' in vcabs in arcades or charging money
' to play the game, which is against
' the Terms of Service of VPX. Please do
' not modify this behavior.
'-----------------------------------------

' Allow zen difficulty, which is a
' special difficulty where players can 
' play as long as they want or until they 
' complete a final wizard mode. HP is 
' hidden and players take a score penalty
' once they end the game depending on how
' much HP damage they took.
Const ALLOW_ZEN = True

' Allow players the choice to play as a
' team towards the table objectives instead
' of against each other (competitive play).
' They will share the same scores and
' objective progress but will have separate
' HP and AC stats. If true and
' TOURNAMENT_MODE is non-zero, then
' cooperative play is forced; a choice is
' not given (team-based tournaments).
'
' Coop mode is activated after 
' initializing player 2 and then pressing
' the action button.
Const ALLOW_COOPERATIVE_PLAY = True
' TODO: Add skillshot DMD screen for adding players when this is false or not allowed via tournament mode

'-----------------------------------------
' *** Tournament Options ***
'-----------------------------------------

' DISCLAIMER: As Wallden is only a digital
' pinball table, it is near impossible to
' ensure a fair tournament competition for
' this or any other virtual table. This is
' due to the many variables surrounding
' vcab controls from vcab to vcab and also
' differences between vcab and desktop
' play. Unfortunately, a standard cannot
' be enforced unless all tournament play
' occurs within a controlled facility.
' Despite Wallden having options for
' tournament play, it is not entirely
' vetted for "fair competition"
' tournaments, and as such, should be 
' limited only to "for fun" tournaments
' unless used within a controlled
' facility. We take no responsibility for
' biases resulting from not acknowledging
' this disclaimer.

'-----------------------------------------
' === TOURNAMENT MODE ===
' Normally, the first player of a new game
' can choose which difficulty they wish to
' play. Furthermore, leaderboards on the
' attract sequence are separated by
' difficulty.
'
' To activate tournament mode, choose a
' difficulty which will be forced for
' every game. This will also result in a
' separate tournament mode leaderboard to
' display on the attract sequence (the
' other leaderboards will be hidden).
'
' You should also set a GLOBAL_SEED
' when using tournament mode. And you
' should set ALLOW_COOPERATIVE_PLAY
' according to whether the tournament is
' team-based (True) or individual based
' (False).
'
' Zen cannot be chosen for
' tournaments because it is not designed
' for tournament play.
'
' 0 = No tournament mode
'	  (Player chooses difficulty)
'
' 1 = Tournament mode - Easy
' 2 = Tournament mode - Normal
' 3 = Tournament mode - Hard
' 4 = Tournament mode - Impossible
'-----------------------------------------
Const TOURNAMENT_MODE = 0
'TODO: caution! Use dataGame.data.Item("tournamentMode") for deciding high score entries, not this Const (except if tournamentMode was a tournament and TOURNAMENT_MODE is 0, consider tournament expired and treat as normal high score entry)

'-----------------------------------------
' === RANDOM NUMBER GENERATOR ===
' While Saving Wallden uses randomness to
' offer variations in gameplay, it is also
' tournament-friendly thanks to how it
' generates random numbers. Each
' component of the game utilizing random
' numbers has its own seed derived from
' the GLOBAL_SEED. When a component uses
' a seed to generate a random number, the
' seed changes in a procedural, non-random
' way to ensure the same randomness every
' game that uses the same GLOBAL_SEED.
' This also allows for the game to present
' the same randomness to every player in a
' competitive multi-player game to ensure
' fairness.
'
' If playing this game for a tournament,
' you should use the same GLOBAL_SEED in
' every game to ensure the same order of
' modes, mode shots, skill shots,
' jackpots, hurryups, and so on.
' Otherwise, you can set GLOBAL_SEED to
' Null, which means a new seed will be
' randomly generated on each game.
'
' This should be a positive float with no 
' more than 15 decimal places, or Null.
'-----------------------------------------
Const GLOBAL_SEED = Null

'*********************************************************************************
' !!! DO NOT EDIT ANYTHING BELOW THIS LINE unless you know what you are doing !!!
'*********************************************************************************

'+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
'
'  #####  ######  ####### ### #       ####### ######     #     #    #    ######  #     # ### #     #  #####
' #     # #     # #     #  #  #       #       #     #    #  #  #   # #   #     # ##    #  #  ##    # #     #
' #       #     # #     #  #  #       #       #     #    #  #  #  #   #  #     # # #   #  #  # #   # #
'  #####  ######  #     #  #  #       #####   ######     #  #  # #     # ######  #  #  #  #  #  #  # #  ####
'       # #       #     #  #  #       #       #   #      #  #  # ####### #   #   #   # #  #  #   # # #     #
' #     # #       #     #  #  #       #       #    #     #  #  # #     # #    #  #    ##  #  #    ## #     #
'  #####  #       ####### ### ####### ####### #     #     ## ##  #     # #     # #     # ### #     #  #####
'
' Proceeding further down will heavily spoil some of the table's gameplay and
' features. Proceed at your OWN RISK! Whitespace was added below intentionally.
'+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-



































































'*****************************************
'    WCON: CONSTANTS
'*****************************************

Const UsingROM = False	'Saving Wallden is an original, ROM-less table

'-----------------------------------------
'    General / Table
'-----------------------------------------
Const cGameName = "savingWallden" ' Game name (should be the same as cPuPPack and should be filename friendly)
Const cPuPPack = "savingWallden" ' PinUP Pack name

Const BallSize = 50  'Ball diameter
Const BallMass = 1 'Ball mass

Const ReflipAngle = 20 'Flippers

Const tnob = 6 ' Total number of balls the machine contains including captive balls (Wallden has 5 trough balls and 1 captive ball)
Const lob = 1 ' Locked balls (in Wallden, we have one captive ball which we consider a "locked ball")

Dim tablewidth
tablewidth = Table1.width
Dim tableheight
tableheight = Table1.height

Dim calloutActive	'Name of the callout sound last played
calloutActive = ""

'-----------------------------------------
'   WDOF: Switch assignments
'	MUST follow numberical conventions
'	specified as logic is specifically
'	coded for it in doSwitch.
'-----------------------------------------

'Buttons (<100; does not fail skillshot
'when triggered, unlike >=100 switches.)
Const SWITCH_LEFT_FLIPPER_BUTTON = 1
Const SWITCH_RIGHT_FLIPPER_BUTTON = 2
Const SWITCH_COIN_1 = 3
Const SWITCH_COIN_2 = 4
Const SWITCH_START_BUTTON = 5
Const SWITCH_ACTION_BUTTON = 6


' Post drops (1xx)
Const SWITCH_POST_DROP_0 = 100
Const SWITCH_POST_DROP_1 = 101
Const SWITCH_POST_DROP_2 = 102
Const SWITCH_POST_DROP_3 = 103
Const SWITCH_POST_DROP_4 = 104
Const SWITCH_POST_DROP_5 = 105
Const SWITCH_POST_DROP_6 = 106
Const SWITCH_POST_DROP_7 = 107
Const SWITCH_POST_DROP_8 = 108
Const SWITCH_POST_DROP_9 = 109
Const SWITCH_POST_DROP_10 = 110
Const SWITCH_POST_DROP_11 = 111
Const SWITCH_POST_DROP_12 = 112
Const SWITCH_POST_DROP_13 = 113
Const SWITCH_POST_DROP_14 = 114
Const SWITCH_POST_DROP_15 = 115
Const SWITCH_POST_DROP_16 = 116
Const SWITCH_POST_DROP_17 = 117
Const SWITCH_POST_DROP_18 = 118
Const SWITCH_POST_DROP_19 = 119
Const SWITCH_POST_DROP_20 = 120
Const SWITCH_POST_DROP_21 = 121
Const SWITCH_POST_DROP_22 = 122
Const SWITCH_POST_DROP_23 = 123
Const SWITCH_POST_DROP_24 = 124
Const SWITCH_POST_DROP_25 = 125
Const SWITCH_POST_DROP_26 = 126
Const SWITCH_POST_DROP_27 = 127
Const SWITCH_POST_DROP_28 = 128
Const SWITCH_POST_DROP_29 = 129

'Crossbow Standups (20x)
Const SWITCH_CROSSBOW_LEFT_STANDUP = 200
Const SWITCH_CROSSBOW_CENTER_STANDUP = 201
Const SWITCH_CROSSBOW_RIGHT_STANDUP = 202

'TRAIN standups (21x)
Const SWITCH_STANDUP_T = 210
Const SWITCH_STANDUP_R = 211
Const SWITCH_STANDUP_A = 212
Const SWITCH_STANDUP_I = 213
Const SWITCH_STANDUP_N = 214

'HP Drops (22x)
Const SWITCH_DROPS_PLUS = 220
Const SWITCH_DROPS_H = 221
Const SWITCH_DROPS_P = 222
Const SWITCH_DROPS_BANK_COMPLETE = 223

'Switches with quick-hit protection (4xx)
'(when these switches are enabled via
'doSwitch, they cannot be enabled again
'until they are first disabled)
Const SWITCH_LEFT_ORBIT_COMPLETE = 401
Const SWITCH_LEFT_HORSESHOE_COMPLETE = 402
Const SWITCH_RIGHT_HORSESHOE_COMPLETE = 403
Const SWITCH_RIGHT_ORBIT_COMPLETE = 405
Const SWITCH_LEFT_RAMP_ENTER = 406
Const SWITCH_LEFT_ORBIT_ENTER = 407
Const SWITCH_LEFT_HORSESHOE_ENTER = 408
Const SWITCH_RIGHT_RAMP_ENTER = 409
Const SWITCH_RIGHT_HORSESHOE_ENTER = 410
Const SWITCH_RIGHT_ORBIT_ENTER = 411

'Other switches (5xx)
Const SWITCH_HOLE = 503
Const SWITCH_SCOOP = 508
Const SWITCH_UPPER_LEFT_INLANE = 509
Const SWITCH_UPPER_CENTER_INLANE = 510
Const SWITCH_UPPER_RIGHT_INLANE = 511
Const SWITCH_RED_BUMPER = 512
Const SWITCH_GREEN_BUMPER = 513
Const SWITCH_BLUE_BUMPER = 514
Const SWITCH_LEFT_SLINGSHOT = 515
Const SWITCH_RIGHT_SLINGSHOT = 516
Const SWITCH_LEFT_OUTLANE = 517
Const SWITCH_LEFT_INLANE = 518
Const SWITCH_RIGHT_INLANE = 519
Const SWITCH_RIGHT_OUTLANE = 520
Const SWITCH_DRAIN = 521
Const SWITCH_BALL_LAUNCH = 522
Const SWITCH_SPINNER = 523
Const SWITCH_LEFT_RAMP_COMPLETE = 524
Const SWITCH_RIGHT_RAMP_COMPLETE = 525
Const SWITCH_VUK = 526
Const SWITCH_TILT_BOBBER = 527
Const SWITCH_HP_DING_WALL = 528
Const SWITCH_SUBWAY_GATE = 529

'-----------------------------------------
'    Scoring
'    CAUTION!: Always use multiples of 10;
'    the ones place is reserved for
'    difficulty level played.
'-----------------------------------------

Dim Scoring
Set Scoring = New WalldenScoring
Class WalldenScoring
	Private cBasic
	Private cBonus
	Private cPenalties
	Private cBlacksmith
	
	' Settings are contained in here (be sure to edit PUP where necessary as well)
	Private Sub Class_Initialize
		
		' Basic Scoring
		Set cBasic = New vpwConstants
		With cBasic
			.Add "bumperInlane", 5000
			.Add "bumperInlanes.complete", 20000 ' TODO
			.Add "bumper", 1770
			.Add "ramp.enter", 5000
			.Add "ramp.complete", 20000
			.Add "orbit.enter", 5000
			.Add "orbit.complete", 10000
			.Add "crossbow", 50000
			.Add "trainStandup", 2500
			.Add "hpTarget", 10000
			.Add "hpDingWall", 7770
			.Add "hpTargets.complete", 50000
			.Add "horseshoe.enter", 5000
			.Add "horseshoe.complete", 5000
			.Add "spinner", 770
			.Add "scoop", 25000
			.Add "outlane", 25000
			.Add "inlane", 5000
			.Add "slingshot", 330
			
			.Add "skillshot", 3000000
			.Add "spell_BONUSX", 250000
			.Add "bonusX", 1000000
			.Add "combo", 250000
		End With
		
		' End of game bonus
		Set cBonus = New vpwConstants
		With cBonus
			.Add "CASTLE_phase1", 100000 			'Points per training session in phase 1
			.Add "CASTLE_boss1", 250000 			'Points for completing the first mini-boss
			.Add "CASTLE_phase2", 200000 			'Points per training session in phase 2
			.Add "CASTLE_boss2", 500000 			'Points for completing the second mini-boss
			.Add "CASTLE_phase3", 300000 			'Points per training session in phase 3
			.Add "CASTLE_boss3", 750000 			'Points for completing the third mini-boss
			.Add "CASTLE_phase4", 400000			'Points for each additional training session after phase 3
			
			.Add "CHASER", 50000					'Points for every jackpot collected in CHASER multiballs (multipliers included)
			
			.Add "SNIPER", 200000					'Points for every collected SNIPER hurry-up
			
			.Add "DRAGON", 2000000					'Points for every completed dragon battle
			
			.Add "VIKING", 250000					'Points for every killed viking
			
			.Add "ESCAPE", 2000000					'Points for every successful ESCAPE
			
			.Add "BLACKSMITH_AC", 150000			'Points for each AC class
			
			.Add "FINAL_BOSS", 5000000				'Points for starting the final boss
			
			.Add "FINAL_JUDGMENT", 10000000			'Points for completing the final boss / getting a final judgment
		End With

		'Score penalties
		Set cPenalties = New vpwConstants
		With cPenalties
			.Add "zen", 50000				'Number of points player loses per HP damage taken in Zen (also points gained per HP if player had positive HP)
		End With

		'Blacksmith objective
		Set cBlacksmith = New vpwConstants
		With cBlacksmith
			.Add "wizard_kill_spell", 10000000	'Points awarded each time a BLACKSMITH letter is spelled in blacksmith boss battle (kill)
			.Add "wizard_kill_completed", 50000000	'Bonus points awarded for killing the Blacksmith
			.Add "wizard_spare", Array(4000000, 10000000, 18000000, 28000000, 40000000, 54000000, 70000000, 88000000, 108000000, 150000000) 'Jackpots depending on number of BLACKSMITH letters spelled
		End With
	End Sub
	
	Public Property Get basic(Key)
		basic = cBasic.Item(Key)
	End Property
	
	Public Property Get bonus(Key)
		bonus = cBonus.Item(Key)
	End Property

	Public Property Get penalties(Key)
		penalties = cPenalties.Item(Key)
	End Property

	Public Property Get blacksmith(Key)
		blacksmith = cBlacksmith.Item(Key)
	End Property
End Class

'-----------------------------------------
'    Health / Drain Damage
'-----------------------------------------

Dim Health
Set Health = New WalldenHealth
Class WalldenHealth
	Private cHP
	Private cDrainDamage
	Private cArmorClass
	
	' Settings are contained in here as arrays according to game difficulty: Array(zen, easy, normal, hard, impossible)
	Private Sub Class_Initialize
		
		' HP / Health
		Set cHP = New vpwConstants
		With cHP
			.Add "start", Array(0, 50, 50, 50, 100)	                'Health at the start of a game
			.Add "heal.dropTargets.completed", Array(1, 1, 1, 1, 3) ' Health healed for every completion of the +HP drop target bank
			.Add "heal.dropTargets.bonus", Array(1, 1, 1, 1, 3)	' Bonus HP for every rubber hit before the drop targets reset
		End With
		
		' Drain Damage
		Set cDrainDamage = New vpwConstants
		With cDrainDamage
			.Add "start", Array(10, 10, 15, 20, 20) 											' Starting drain damage at the beginning of a game
			.Add "nudge", Array(2, 2, 3, 4, 4) 													' Drain damage added when nudging too hard
			.Add "training.noPoints", Array(2, 2, 3, 4, 4) 										' Drain damage added when ending a training session having earned 0 points from it
			.Add "training.greatJob", Array(3, 3, 4, 5, 5) 										' Drain damage subtracted when ending a training session having earned training.greatJob.threshold points
			.Add "training.greatJob.threshold", Array(7500000, 15000000, 22500000, 30000000) 	' Minimum points required for a training session to remove drain damage (Array according to phase #: Array(1, 2, 3, 4 [bonus])
			.Add "poison", Array(2, 2, 3, 4, 4) 												' Drain damage added when hitting an outlane with poison lit
		End With
		
		' Armor Class
		Set cArmorClass = New vpwConstants
		With cArmorClass
			.Add "start", Array(12, 11, 10, 9, 8) ' Starting AC at the beginning of a game
		End With
	End Sub
	
	Public Property Get HP(Key)
		HP = cHP.Item(Key)
	End Property
	Public Property Get drainDamage(Key)
		drainDamage = cDrainDamage.Item(Key)
	End Property
	Public Property Get AC(Key)
		AC = cArmorClass.Item(Key)
	End Property
End Class

'-----------------------------------------
'    Modes (see startMode)
'    Flippers disabled for < 1000.
'    No scoring for < 900.
'    Bumpers / slings disabled for < 1000
'	 (except blue bumper to prevent stuck
'	 Balls).
'-----------------------------------------

' Idle modes where a game is not in progress (0xx)
Const MODE_FAULT = 0 				'Table experienced a fault
Const MODE_INITIALIZE = 1
Const MODE_ATTRACT = 2 				'No game in progress / attract sequence
Const MODE_RESUME_GAME_PROMPT = 3	'Prompt if an in-progress game should be resumed
Const MODE_HIGH_SCORE_FINAL = 4

' Pre-game modes (1xx)
Const MODE_CHOOSE_DIFFICULTY = 101 	'Player is selecting which difficulty to use
Const MODE_START_GAME = 102 		'New game sequence

' Post-game modes (2xx)
Const MODE_GAME_OVER = 200			'Game over sequence
Const MODE_ZEN_END = 201			'Player requested to end a zen game, and we are waiting for the balls to drain
Const MODE_HIGH_SCORE_ENTRY = 202

' Pre-gameplay modes (intro sequences) (300 - 499)
Const MODE_DEATH_SAVE_INTRO = 300

Const MODE_BLACKSMITH_WIZARD_INTRO = 310
Const MODE_BLACKSMITH_WIZARD_SELECTION = 311
Const MODE_BLACKSMITH_WIZARD_INTRO_KILL = 312
Const MODE_BLACKSMITH_WIZARD_INTRO_SPARE = 313

Const MODE_GLITCH_WIZARD_INTRO = 320

' Post-gameplay modes (outtro sequences) (500 - 699)
Const MODE_END_OF_TURN = 500		'When a player drains their ball, resulting in the end of their turn
Const MODE_GLITCH_WIZARD_OUTTRO = 501
Const MODE_BLACKSMITH_WIZARD_KILL_FAIL = 502
Const MODE_BLACKSMITH_WIZARD_KILL_SUCCESS = 503

' Non-gameplay modes where a game is in progress (7xx)
Const MODE_MISSING_BALL = 700		'Ball search failed / operator requested

' Gameplay modes where we are waiting for all the balls to drain (8xx; ball search will monitor these modes)
Const MODE_GLITCH_WIZARD_END = 800
Const MODE_BLACKSMITH_WIZARD_END = 801

' Gameplay modes where table elements are disabled but scoring is active (e.g. "video modes") (9xx)
Const MODE_DEATH_SAVE = 900
Const MODE_END_OF_GAME_BONUS = 901

' Basic gameplay (1xxx)
Const MODE_NORMAL = 1000			'Basic mode where no features are active and the player is spelling letters towards objectives
Const MODE_SKILLSHOT = 1001

' CASTLE objective (2xxx)

' CHASER objective (3xxx)

' DRAGON objective (4xxx)

' VIKING objective (5xxx)

' ESCAPE objective (6xxx)

' SNIPER objective (7xxx)

' BLAKSMITH objective (8xxx)
Const MODE_BLACKSMITH_WIZARD_SPARE = 8010
Const MODE_BLACKSMITH_WIZARD_KILL = 8011

' FINAL BOSS objective (9xxx)

' FINAL JUDGMENT objective (10xxx)

' other / secret objectives (11xxx)
Const MODE_GLITCH_WIZARD = 11000

'*****************************************
'   WTRK: GAME PLAY TRACKING VARIABLES
'*****************************************

'-----------------------------------------
'    Global / Machine
'-----------------------------------------

Dim GAME_OFFICIALLY_STARTED
GAME_OFFICIALLY_STARTED = False ' Set to true after the first ball of the game is plunged and the first switch is hit; new players can no longer be added to the game.

'-----------------------------------------
'    Game play
'-----------------------------------------

' Ball search
Dim BALL_SEARCH_TIME
BALL_SEARCH_TIME = GameTime
Dim BALL_SEARCH_STAGE
BALL_SEARCH_STAGE = 0
Dim BALL_SEARCH_MODE_MEMORY			' Remember which mode we were in so we can re-initialize it after ball search
Dim BALL_SEARCH_MUSIC_MEMORY		' Remember what music was playing so we can go back to it after ball search
Dim LeftFlipperPressed
LeftFlipperPressed = False
Dim RightFlipperPressed
RightFlipperPressed = False

Dim LAST_CLOCK
LAST_CLOCK = GameTime			  	' The GameTime when the last clock execution took place (for counting shield times, mode times, etc.

' Bonus X / Bumper Inlane Tracking
Dim bumperInlanes(3)

' Mode tracking
Dim CURRENT_MODE
CURRENT_MODE = MODE_INITIALIZE    	' Which MODE_* constant is currently active

Dim MODE_SHOTS_ORDER_A 				' Used to track the order in which shots should be activated for modes and multiballs (MUST BE ARRAY)
Dim MODE_SHOTS_ORDER_B 				' Used to track the order in which shots should be activated for modes and multiballs (MUST BE ARRAY)
MODE_SHOTS_ORDER_A = Array()
MODE_SHOTS_ORDER_B = Array()

Dim MODE_SHOTS_A 					' Used to track which shots are lit or completed (MUST BE ARRAY)
Dim MODE_SHOTS_B 					' Used to track which shots are lit or completed (MUST BE ARRAY)
MODE_SHOTS_A = Array()
MODE_SHOTS_B = Array()

Dim MODE_VALUES						' Used to track values for mode shots (should NOT be reset to 0 if used to track mode total except at the start of a new mode)
Dim MODE_COUNTERS_A					' Used to track other internal mode information. When checking, should ALWAYS be guarded by a PARENT if/case clause for the CURRENT_MODE to prevent type errors.
Dim MODE_COUNTERS_B					' Used to track other internal mode information. When checking, should ALWAYS be guarded by a PARENT if/case clause for the CURRENT_MODE to prevent type errors.
Dim MODE_COUNTERS_C					' Used to track other internal mode information. When checking, should ALWAYS be guarded by a PARENT if/case clause for the CURRENT_MODE to prevent type errors.
Dim MODE_COUNTERS_D					' Used to track other internal mode information. When checking, should ALWAYS be guarded by a PARENT if/case clause for the CURRENT_MODE to prevent type errors.

'Switch tracking
Dim switchTime	'Dictionary of switch hit times; key is switch number and value is GameTime last hit
Set switchTime = CreateObject("Scripting.Dictionary")

'Flippers
Dim flippersReversed	'Whether the flippers should be reversed (left button = right flipper, right button = left Flipper). Use the reverseFlippers routine instead of changing this directly.
flippersReversed = False

'*****************************************
'   VBS FILES
'*****************************************

LoadCoreFiles ' Must be called after initializing our constants as BallSize and a couple others must already be defined before loading core.vbs

Sub LoadCoreFiles
	'First, try to load the Controller.vbs (DOF), which helps controlling additional hardware like lights, gears, knockers, bells and chimes (to increase realism)
	'This table uses DOF via the 'SoundFX' calls that are inserted in some of the PlaySound commands, which will then fire an additional event, instead of just playing a sample/sound effect
	On Error Resume Next
	ExecuteGlobal GetTextFile("controller.vbs")
	If Err Then 
		MsgBox "You need the Controller.vbs file In order To run this table (installed With the VPX package In the scripts folder)"
		QuitPlayer 1	'Quit the game as controller.vbs is required
		Err.Clear
	End If

	ExecuteGlobal GetTextFile("core.vbs")
	If Err Then 
		MsgBox "You need the Core.vbs file In order To run this table (installed With the VPX package In the scripts folder)"
		QuitPlayer 1	'Quit the game as controller.vbs is required
		Err.Clear
	End If
	On Error GoTo 0
End Sub

'*****************************************
'	VPM CLASSES
'*****************************************

' Impulse / Auto Plunger
Dim impulseP
Set impulseP = New cvpmImpulseP
With impulseP
	.InitImpulseP Trigger1, 35, 0
	.Random 4
	.InitExitSnd "Autofire", "Autofire"
End With

'*****************************************
'    WQUE: QUEUES
'*****************************************

'-----------------------------------------
'	Main queue priorities (1 - 9):
'	1 = Basically queue_empty
'	2 = Attract or lp intros / scenes
'   3 = Scenes
'	9 = Features are Ready

'	Low importance (10 - 99)
'   10 = expireMode (mode totals)
'   11 = Poisoned outlane hit
'   12 = General Letter Spelling
'   20 = BONUS X / BLACKSMITH / +HP / Combo
'
'	High importance (100 - 999)
'	100 = Regular Jackpot
'	200 = Super Jackpot
'
'	300 = Powerup Deployments; sequence / follow-up screen
'	301 = Powerup Deployments; info screen
'
'	UI responses (1000)
'
'	Extremely high importance (1001+)
'-----------------------------------------
Dim queue ' Main queue
Set queue = New vpwQueueManager
queue.QueueEmpty = "queue_empty"
If DEBUG_LOG_QUEUE = True Then queue.debugOn = "Main Queue"

Dim queueB ' Secondary queue
Set queueB = New vpwQueueManager
If DEBUG_LOG_QUEUE = True Then queueB.debugOn = "Secondary Queue"

Dim troughQueue ' Trough and auto-fire queue
Set troughQueue = New vpwQueueManager
'	troughQueue.debugOn = "troughQueue"

Dim flasherQueue ' Queue for the flashers
Set flasherQueue = New vpwQueueManager
flasherQueue.QueueEmpty = "flasherQueue_empty"
If DEBUG_LOG_QUEUE = True Then flasherQueue.debugOn = "Flasher Queue"

Dim postDropQueue ' Queue for firing the drop post solenoids
Set postDropQueue = New vpwQueueManager
'	postDropQueue.debugOn = "postDropQueue"

'*****************************************
'    CLOCKS
'*****************************************

Dim Clocks
Set Clocks = New vpwClocks

With Clocks
	.Add "shield"
	.Add "mode"
	.Add "leftBumperDiverter"
	.Add "rightBumperDiverter"
	.Add "hpTargets"
	.Add "bumperInlanes"
	.Add "combo"
	.Add "impossibleHP"
End With

'*****************************************
'    MUSIC AND SOUND MANAGERS
'*****************************************

Dim Music
Set Music = New vpwMusic
With Music
	.Add "music_choose_difficulty", -1, VOLUME_BG_MUSIC, 500, 500
	.Add "music_normal_firstball", -1, VOLUME_BG_MUSIC, 500, 1000
	.Add "music_normal_0", -1, VOLUME_BG_MUSIC, 500, 1000
	.Add "music_normal_1", -1, VOLUME_BG_MUSIC, 500, 1000
	.Add "music_normal_2", -1, VOLUME_BG_MUSIC, 500, 1000
	.Add "music_zen_0", -1, VOLUME_BG_MUSIC, 500, 1000
	.Add "music_zen_1", -1, VOLUME_BG_MUSIC, 500, 1000
	.Add "music_zen_2", -1, VOLUME_BG_MUSIC, 500, 1000
	.Add "music_eerie_chant", -1, VOLUME_BG_MUSIC, 3000, 1000
	.Add "music_dead_bonus", 1, VOLUME_BG_MUSIC, 0, 0
	.Add "music_game_over_failed", 1, VOLUME_BG_MUSIC, 0, 500
	.Add "music_resume_game", 1, VOLUME_BG_MUSIC, 0, 0
	'.Add "music_zen_suspense", 1, VOLUME_BG_MUSIC, 0, 0
	'.Add "music_death_save", -1, VOLUME_BG_MUSIC, 0, 2000
	.Add "music_glitch", -1, VOLUME_BG_MUSIC, 0, 3000
	.Add "music_blacksmith_lore", -1, VOLUME_BG_MUSIC, 0, 0
	'.Add "music_blacksmith_wizard_prompt", -1, VOLUME_BG_MUSIC, 0, 0
	.Add "music_tension", -1, VOLUME_BG_MUSIC, 0, 3000
	.Add "music_blacksmith_kill", -1, VOLUME_BG_MUSIC, 500, 1000
	.Add "music_startup", 1, VOLUME_BG_MUSIC, 0, 1000
	.Add "music_blacksmith_kill_fail", 1, VOLUME_BG_MUSIC, 0, 0
	.Add "music_blacksmith_kill_success", 1, VOLUME_BG_MUSIC, 0, 0
	.Add "music_blacksmith_death", 1, VOLUME_BG_MUSIC, 1500, 1500
	.Add "music_blacksmith_spare", -1, VOLUME_BG_MUSIC, 500, 1000
	.Add "music_blacksmith_spare_lore", -1, VOLUME_BG_MUSIC, 0, 0
	.Add "music_ball_search_intro", 1, VOLUME_BG_MUSIC, 0, 0
	.Add "music_ball_search_main", -1, VOLUME_BG_MUSIC, 0, 0
	.Add "music_ball_search_ultimate", 1, VOLUME_BG_MUSIC, 0, 0
	.Add "sting_victory", 1, VOLUME_BG_MUSIC, 0, 0
End With

Dim Tracks
Set Tracks = New vpwTracks

'*****************************************
'   LAMP CONTROLLERS
'*****************************************

Dim lampC
Set lampC = New LStateController

Dim NullFader
Set NullFader = New NullFadingObject
Dim Lampz
Set Lampz = New LampFader

InitLampsNF()

'*****************************************
'    OTHER DIMs
'*****************************************

Dim BQueue				'If > 0, we want to launch balls until BIP reaches this number. Generally do not modify directly; use the targetBIP method.
BQueue = 0 				

Dim BAutoPlunge			'Tracks when we want queued balls to be auto-plunged. Generally do not modify directly; use the targetBIP method.
BAutoPlunge = False						

Dim BIPL				'Balls in the shooter lane
BIPL = 0

Dim BIS					'Balls in the subway
BIS = 0
Dim ballCameFromSubway	'Whether a ball which entered the trough probably came from the subway and not the drain
ballCameFromSubway = False

Dim ballsInTrough		'Number of balls currently in the trough
ballsInTrough = (tnob - lob)

Dim BallInScoop			'Is there a ball in the scoop?
BallInScoop = False		

Dim BallInKicker		'Is there a ball in the VUK?
BallInKicker = False	

' Flipper polarity correction
Dim LF
Set LF = New FlipperPolarity
Dim RF
Set RF = New FlipperPolarity

' Trough and captive ball
Dim CBall1, ETBall1, ETBall2, ETBall3, ETBall4, ETBall5, gBOT

'*******************************************************
'   DATA MANAGERS
'   Manages saved / loaded data and game play tracking.
'*******************************************************

' High Scores
Dim dataHighScores(6)

' Zen leaderboard
Set dataHighScores(0) = New vpwData
With dataHighScores(0)
	.gameName = cGameName
	.dataName = "HSmyst"
	
	' Scores
	.Load "HS1", 2500000000
	.Load "HS2", 2000000000
	.Load "HS3", 1500000000
	.Load "HS4", 1000000000
	.Load "HS5", 500000000
	
	' Dates
	.Load "HSD1", "DEMO SCORE"
	.Load "HSD2", "DEMO SCORE"
	.Load "HSD3", "DEMO SCORE"
	.Load "HSD4", "DEMO SCORE"
	.Load "HSD5", "DEMO SCORE"
	
	' Names
	.Load "HSN1", "CRAZY BOB"
	.Load "HSN2", "DAVE"
	.Load "HSN3", "COW"
	.Load "HSN4", "CACTUS"
	.Load "HSN5", "TEAM REDTEDD"
	
	.reset = RESET_HIGH_SCORES
End With

' Easy leaderboard
Set dataHighScores(1) = New vpwData
With dataHighScores(1)
	.gameName = cGameName
	.dataName = "HSeasy"
	
	' Scores
	.Load "HS1", 1250000000
	.Load "HS2", 1000000000
	.Load "HS3", 750000000
	.Load "HS4", 500000000
	.Load "HS5", 250000000
	
	' Dates
	.Load "HSD1", "DEMO SCORE"
	.Load "HSD2", "DEMO SCORE"
	.Load "HSD3", "DEMO SCORE"
	.Load "HSD4", "DEMO SCORE"
	.Load "HSD5", "DEMO SCORE"
	
	' Names
	.Load "HSN1", "BLACKSMITH"
	.Load "HSN2", "SNIPER"
	.Load "HSN3", "GOBLINS"
	.Load "HSN4", "VIKINGS"
	.Load "HSN5", "MINI BOSSES"
	
	.reset = RESET_HIGH_SCORES
End With

' Normal leaderboard
Set dataHighScores(2) = New vpwData
With dataHighScores(2)
	.gameName = cGameName
	.dataName = "HSnormal"
	
	' Scores
	.Load "HS1", 1000000000
	.Load "HS2", 800000000
	.Load "HS3", 600000000
	.Load "HS4", 400000000
	.Load "HS5", 200000000
	
	' Dates
	.Load "HSD1", "DEMO SCORE"
	.Load "HSD2", "DEMO SCORE"
	.Load "HSD3", "DEMO SCORE"
	.Load "HSD4", "DEMO SCORE"
	.Load "HSD5", "DEMO SCORE"
	
	' Names
	.Load "HSN1", "TRAINER DAVE"
	.Load "HSN2", "CHASER"
	.Load "HSN3", "ARELYEL"
	.Load "HSN4", "WALLDEN"
	.Load "HSN5", "VPW"
	
	.reset = RESET_HIGH_SCORES
End With

' Hard leaderboard
Set dataHighScores(3) = New vpwData
With dataHighScores(3)
	.gameName = cGameName
	.dataName = "HShard"
	
	' Scores
	.Load "HS1", 750000000
	.Load "HS2", 600000000
	.Load "HS3", 450000000
	.Load "HS4", 300000000
	.Load "HS5", 150000000
	
	' Dates
	.Load "HSD1", "DEMO SCORE"
	.Load "HSD2", "DEMO SCORE"
	.Load "HSD3", "DEMO SCORE"
	.Load "HSD4", "DEMO SCORE"
	.Load "HSD5", "DEMO SCORE"
	
	' Names
	.Load "HSN1", "FINAL BOSS"
	.Load "HSN2", "YOUR MAJESTY"
	.Load "HSN3", "SPIRITS"
	.Load "HSN4", "HARMONY"
	.Load "HSN5", "LOVINITY"
	
	.reset = RESET_HIGH_SCORES
End With

' Impossible leaderboard
Set dataHighScores(4) = New vpwData
With dataHighScores(4)
	.gameName = cGameName
	.dataName = "HSimpos"
	
	' Scores
	.Load "HS1", 500000000
	.Load "HS2", 400000000
	.Load "HS3", 300000000
	.Load "HS4", 200000000
	.Load "HS5", 100000000
	
	' Dates
	.Load "HSD1", "DEMO SCORE"
	.Load "HSD2", "DEMO SCORE"
	.Load "HSD3", "DEMO SCORE"
	.Load "HSD4", "DEMO SCORE"
	.Load "HSD5", "DEMO SCORE"
	
	' Names
	.Load "HSN1", "SINS"
	.Load "HSN2", "REPENT"
	.Load "HSN3", "NO MERCY"
	.Load "HSN4", "HELL"
	.Load "HSN5", "FIRE"
	
	.reset = RESET_HIGH_SCORES
End With

' Tournament leaderboard
Set dataHighScores(5) = New vpwData
With dataHighScores(5)
	.gameName = cGameName
	.dataName = "HStourn"
	
	' Scores
	.Load "HS1", 0
	.Load "HS2", 0
	.Load "HS3", 0
	.Load "HS4", 0
	.Load "HS5", 0
	
	' Dates
	.Load "HSD1", "NO SCORE YET"
	.Load "HSD2", "NO SCORE YET"
	.Load "HSD3", "NO SCORE YET"
	.Load "HSD4", "NO SCORE YET"
	.Load "HSD5", "NO SCORE YET"
	
	' Names
	.Load "HSN1", ""
	.Load "HSN2", ""
	.Load "HSN3", ""
	.Load "HSN4", ""
	.Load "HSN5", ""
	
	.reset = RESET_TOURNAMENT_HIGH_SCORES
End With

' Game tracking
Dim dataGame
Set dataGame = New vpwData
With dataGame
	.gameName = cGameName
	.dataName = "game"
	
	' Basic tracking
	.Load "seed", 0				'Global seed number for this game
	.Load "resumeStatus", 4 	'When resuming a game: 0 = current player takes drain penalty / go to next player. 1 = Just go to next player without penalty. 2 = Resume current player (skillshot); 3 = resume current player (normal); 4 = resume current player (normal via game crash)
	.Load "tournamentMode", 0	'Which tournament mode this game was played (see TOURNAMENT_MODE) (0 = no tournament)
	.Load "numPlayers", 0
	.Load "playerUp", 0 		'0 = no game in progress, else player number who is up
	.Load "gameMode", 0 		'0 = competitive, 1 = collaborative / coop
	.Load "gameDifficulty", 2 	'0 = zen, 1 = easy, 2 = normal, 3 = hard, 4 = impossible
	.Load "ball", -1 			'The game itself uses HP, but this is used for Scorbit, tracking when players can end zen games, and when an in-progress game was never finished
	.Load "jester", 0			'Is this a "Jester" game?
	
	' Final Boss Battle
	.Load "fBossStage", 0
	.Load "fBossHP", 0
	.Load "fBossSpinTime", 0
	
	.reset = RESET_GAME_TRACKING
End With

' Player tracking (use currentPlayer and currentPlayerSet to read and adjust these values)
Dim dataPlayer(4) ' Game play progress
Dim dataSeeds(4) ' Random seeds
Dim bkP
For bkP = 0 To 3
	' Player tracking
	Set dataPlayer(bkP) = New vpwData
	With dataPlayer(bkP)
		.gameName = cGameName
		.dataName = "player" & bkP
		
		'Basic Tracking
		.Load "colorSlot", 0 ' 0 = Not in Play, 1 = Red, 2 = Yellow, 3 = Green, 4 = Blue
		.Load "name", ""
		.Load "score", 0
		.Load "HP", 0
		.Load "deathSaves", 0
		.Load "dead", 0		'0 = Alive, 1 = dead
		.Load "drainDamage", 0
		.Load "AC", 0
		.Load "bonusX", 1
		.Load "gems", 0 'Powerups
		.Load "poisonL", 0
		.Load "poisonR", 0
		
		'spell_* for tracking which letters were spelled
		.Load "spell_BONUSX", 0
		.Load "spell_TRAIN_T", 0 'binary
		.Load "spell_TRAIN_R", 0 'binary
		.Load "spell_TRAIN_A", 0 'binary
		.Load "spell_TRAIN_I", 0 'binary
		.Load "spell_TRAIN_N", 0 'binary
		.Load "spell_CHASER", 0
		.Load "spell_SNIPER", 0
		.Load "spell_DRAGON", 0
		.Load "spell_VIKING", 0
		.Load "spell_ESCAPE", 0
		.Load "spell_BLACK", 0
		.Load "spell_SMITH", 0
		
		'obj_* objective progress
		.Load "obj_CASTLE", 0 ' 1-3, 5-7, 9-11 = training sessions; 4, 8, 12 = mini bosses completed; 12 = objective complete; >12 = bonus training sessions
		.Load "obj_CHASER", 0 ' 1 per jackpot; 50 = objective completed; >50 = additional jackpots collected
		.Load "obj_SNIPER", 0 ' 1 per sniper phase completed; 3 = objective completed. >3 = additional hurry-ups completed
		.Load "obj_DRAGON", 0 ' 1 = objective completed; >1 = additional dragon battles won
		.Load "obj_VIKING", 0 ' 1 per viking killed; 8 = objective completed; >8 = additional vikings killed
		.Load "obj_ESCAPE", 0 ' 1 per escape completed; 1 = objective completed; >1 = additional escapes completed
		.Load "obj_BLACKSMITH", 0 ' 1 = completed; >1 = additional blacksmiths spelled
		.Load "obj_FINAL_BOSS", 0 ' 1 = Final Boss Started, 2 = Final Boss Defeated (and final judgment achieved)

		'spare_* Whether we spared or killed something
		.Load "spare_CASTLE", 0 ' 0 = killed, 1 = spared
		.Load "spare_CHASER", 0 ' 0 = killed, 1 = spared
		.Load "spare_SNIPER", 0	' 0 = killed, 1 = spared
		.Load "spare_DRAGON", 0 ' 0 = killed, 1 = spared
		.Load "spare_VIKING", 0 ' 0 = killed, 1 = spared
		.Load "spare_ESCAPE", 0 ' 0 = abandoned, 1 = saved
		.Load "spare_BLACKSMITH", 0	'0 = killed, 1 = spared, 2 = failed to kill
		.Load "spare_FINAL_BOSS", 0 ' 0 = killed, 1 = spared
		
		'Viking HP tracking
		.Load "viking_HP_0", 5
		.Load "viking_HP_1", 5
		.Load "viking_HP_2", 5
		.Load "viking_HP_3", 5
		.Load "viking_HP_4", 5
		.Load "viking_HP_5", 5
		.Load "viking_HP_6", 5
		.Load "viking_HP_7", 5

		'Dragon HP tracking
		.Load "dragon_HP_red", 250
		.Load "dragon_HP_green", 250
		.Load "dragon_HP_blue", 250
		
		.reset = RESET_GAME_TRACKING
	End With
	
	' Seeds for our random number generator
	Set dataSeeds(bkP) = New vpwData
	With dataSeeds(bkP)
		.gameName = cGameName
		.dataName = "seeds" & bkP
		
		.Load "skillshot", 0 		'Determines which standup is lit for the skill shot at the start of each turn
		.Load "gems", 0				'Determines the order of gem shots to collect a powerup.
		.Load "train_sessions", 0 	'Determines the order In which training sessions will be played
		.Load "train_shots", 0 		'Determine shots within modes; valid shots should be shuffled at the start of each mode with this seed (incremented after each mode)
		.Load "chaser_type", 0 		'Determines the order in which chaser multiball variations will be played
		.Load "chaser_shots", 0 	'Determines shots in each chaser multiball; should have an array of shots shuffled at the start of the multiball with this seed (incremented after each multiball)
		.Load "dragon", 0 			'Determines shots for dragon battles
		.Load "sniper", 0 			'Determines shots for sniper hurry-ups (for stages with multiple shots, should be generated in a shuffled array at the start of the phase.
		
		.reset = RESET_GAME_TRACKING
	End With
Next

'*****************************************************************************************************************************************
'  ZERR: ERROR LOGS by baldgeek
'*****************************************************************************************************************************************

' Log File Usage:
'   WriteToLog "Label 1", "Message 1 "
'   WriteToLog "Label 2", "Message 2 "


Class DebugLogFile
	
	Private Filename
	Private TxtFileStream
	
	Private Function LZ(ByVal Number, ByVal Places)
		Dim Zeros
		Zeros = String(CInt(Places), "0")
		LZ = Right(Zeros & CStr(Number), Places)
	End Function
	
	Private Function GetTimeStamp
		Dim CurrTime, Elapsed, MilliSecs
		CurrTime = Now()
		Elapsed = Timer()
		MilliSecs = Int((Elapsed - Int(Elapsed)) * 1000)
		GetTimeStamp = _
		LZ(Year(CurrTime),   4) & "-" _
		& LZ(Month(CurrTime),  2) & "-" _
		& LZ(Day(CurrTime),    2) & " " _
		& LZ(Hour(CurrTime),   2) & ":" _
		& LZ(Minute(CurrTime), 2) & ":" _
		& LZ(Second(CurrTime), 2) & ":" _
		& LZ(MilliSecs, 4)
	End Function
	
	' *** Debug.Print the time with milliseconds, and a message of your choice
	Public Sub WriteToLog(label, message, code)
		Dim FormattedMsg, Timestamp
		'Filename = UserDirectory + "\" + cGameName + "_debug_log.txt"
		Filename = cGameName + "_debug_log.txt"
		
		Set TxtFileStream = CreateObject("Scripting.FileSystemObject").OpenTextFile(Filename, code, True)
		Timestamp = GetTimeStamp
		FormattedMsg = GetTimeStamp + " : " + label + " : " + message
		TxtFileStream.WriteLine FormattedMsg
		TxtFileStream.Close
		Debug.print GetTimeStamp & " " & label & " : " & message
	End Sub
	
End Class

Sub WriteToLog(label, message)
	Dim LogFileObj
	Set LogFileObj = New DebugLogFile
	LogFileObj.WriteToLog label, message, 8
End Sub

Sub NewLog()
	Dim LogFileObj
	Set LogFileObj = New DebugLogFile
	LogFileObj.WriteToLog "NEW Log", " ", 2
End Sub

'****************************************************
'   WDMD: PINUP / DISPLAY
'	Uses both / either PinUP or JPSalas' reduced
'	DMD script.
'****************************************************
Dim DMD
Set DMD = New dmdWallden

'---------------------------
' Video Constants
' Value = D* (<900) PUP trigger
' CURRENT NUMBER: 118
'---------------------------

'Initialize
Const VIDEO_CLEAR = 0
Const VIDEO_SELF_TEST = 1
Const VIDEO_VPW_LOGO = 2
Const VIDEO_TESTS_PASSED = 3
Const VIDEO_BLACK_BG = 4

'Attract sequence
Const VIDEO_FREE_PLAY = 5
Const VIDEO_VPW_PRESENTS = 6
Const VIDEO_RGB_TEST = 7
Const VIDEO_SAVING_WALLDEN = 8
Const VIDEO_PREVIOUS_SCORES = 75
Const VIDEO_TUTORIAL_INTRO = 14
Const VIDEO_TUTORIAL_CASTLE = 15
Const VIDEO_TUTORIAL_CHASER = 16
Const VIDEO_TUTORIAL_DRAGON = 17
Const VIDEO_TUTORIAL_VIKING = 18
Const VIDEO_TUTORIAL_ESCAPE = 19
Const VIDEO_TUTORIAL_SNIPER = 20
Const VIDEO_TUTORIAL_BLACKSMITH = 21
Const VIDEO_TUTORIAL_FINAL_BOSS = 22
Const VIDEO_TUTORIAL_FINAL_JUDGMENT = 23
Const VIDEO_CONTRIBUTORS_INTRO = 24
Const VIDEO_CONTRIBUTORS_ARELYEL = 25
Const VIDEO_CONTRIBUTORS_DONKEYKLONK = 26
Const VIDEO_CONTRIBUTORS_FLUX = 27
Const VIDEO_CONTRIBUTORS_GRAHAM_GROWAT = 28
Const VIDEO_CONTRIBUTORS_HARMONY = 29
Const VIDEO_CONTRIBUTORS_SMAUG = 76
Const VIDEO_CONTRIBUTORS_OTHER = 30

'High Scores (should always ascend in number except tournament!)
Const VIDEO_HIGH_SCORES_ZEN = 9
Const VIDEO_HIGH_SCORES_EASY = 10
Const VIDEO_HIGH_SCORES_NORMAL = 11
Const VIDEO_HIGH_SCORES_HARD = 12
Const VIDEO_HIGH_SCORES_IMPOSSIBLE = 13
Const VIDEO_HIGH_SCORES_TOURNAMENT = 55
Const VIDEO_HIGH_SCORE_FINAL = 117

'Difficulty Selection
Const VIDEO_DIFFICULTY_ZEN = 31
Const VIDEO_DIFFICULTY_EASY = 32
Const VIDEO_DIFFICULTY_NORMAL = 33
Const VIDEO_DIFFICULTY_HARD = 34
Const VIDEO_DIFFICULTY_IMPOSSIBLE = 35

'Skillshot
Const VIDEO_SKILLSHOT_ADD_PLAYERS = 37
Const VIDEO_SKILLSHOT = 38
Const VIDEO_SKILLSHOT_COMPLETE = 39
Const VIDEO_SKILLSHOT_FAILED = 40
Const VIDEO_SKILLSHOT_ZEN_END = 71

'Combo
Const VIDEO_COMBO = 98

Const VIDEO_START_GAME = 42
Const VIDEO_PLAYER_STATS = 80

'Normal / General gameplay
Const VIDEO_GAMEPLAY_NORMAL = 41
Const VIDEO_GAMEPLAY_INSTRUCTIONS = 43
Const VIDEO_BALL_SHIELD = 44
Const VIDEO_DAMAGE_RECEIVED = 45
Const VIDEO_PLAYER_1 = 46
Const VIDEO_PLAYER_2 = 47
Const VIDEO_PLAYER_3 = 48
Const VIDEO_PLAYER_4 = 49
Const VIDEO_BALL_SEARCH = 50
Const VIDEO_BALL_MISSING = 51
Const VIDEO_POWERUP_COLLECTED = 56
Const VIDEO_DRAINING_BALLS = 102
Const VIDEO_LAST_HURRAH = 112
Const VIDEO_POISONED_OUTLANE = 115
Const VIDEO_FINAL_SCORE = 116
Const VIDEO_TILT = 118

'Death Save
Const VIDEO_DEATH_SAVE_INTRO = 72
Const VIDEO_DEATH_SAVE_PROGRESS = 73
Const VIDEO_NO_MORE_DEATH_SAVES = 74

'End of game
Const VIDEO_DEAD = 53
Const VIDEO_GAME_OVER_FAILED = 57
Const VIDEO_ZEN_END_WAIT = 69
Const VIDEO_ZEN_DAMAGE_PENALTY = 70

'Bonus X / Glitch
Const VIDEO_BONUSX_1 = 58
Const VIDEO_BONUSX_2 = 59
Const VIDEO_BONUSX_3 = 60
Const VIDEO_BONUSX_4 = 61
Const VIDEO_BONUSX_5 = 62
Const VIDEO_BONUSX_6 = 63
Const VIDEO_BONUSX_GLITCH_3 = 79
Const VIDEO_GLITCH_WIZARD_READY = 81
Const VIDEO_GLITCH_WIZARD_INTRO = 82
Const VIDEO_GLITCH_WIZARD_SCREEN = 83
Const VIDEO_GLITCH_WIZARD_JACKPOT = 84
Const VIDEO_GLITCH_WIZARD_TOTAL = 85
Const VIDEO_GLITCH_OUTTRO = 86
Const VIDEO_GLITCH_JACKPOT_DOUBLED = 97

'Resume Game
Const VIDEO_RESUME_GAME = 64
Const VIDEO_RESUME_GAME_SCORBIT = 65
Const VIDEO_RESUME_GAME_DRAIN_PENALTY = 66
Const VIDEO_RESUME_GAME_NEXT_PLAYER = 67
Const VIDEO_RESUME_GAME_CURRENT_PLAYER = 68

'Blacksmith
Const VIDEO_BLACKSMITH_SPELL = 77
Const VIDEO_BLACKSMITH_WIZARD_READY = 78
Const VIDEO_BLACKSMITH_LORE = 87
Const VIDEO_BLACKSMITH_LORE_STILL = 90
Const VIDEO_BLACKSMITH_LORE_KILL = 91
Const VIDEO_BLACKSMITH_WIZARD_KILL_INTRO = 92
Const VIDEO_BLACKSMITH_WIZARD_KILL = 93
Const VIDEO_BLACKSMITH_WIZARD_KILL_DAMAGE = 95
Const VIDEO_BLACKSMITH_WIZARD_KILL_LOSE_AC = 96
Const VIDEO_BLACKSMITH_WIZARD_KILL_FAIL = 99
Const VIDEO_BLACKSMITH_WIZARD_KILL_FAIL_RUN = 100
Const VIDEO_BLACKSMITH_WIZARD_KILL_FAIL_TOTAL = 101
Const VIDEO_BLACKSMITH_WIZARD_KILL_SUCCESS = 103
Const VIDEO_BLACKSMITH_WIZARD_KILL_SUCCESS_2 = 104
Const VIDEO_BLACKSMITH_WIZARD_KILL_SUCCESS_3 = 105
Const VIDEO_BLACKSMITH_WIZARD_KILL_SUCCESS_TOTAL = 106
Const VIDEO_BLACKSMITH_LORE_SPARE = 107
Const VIDEO_BLACKSMITH_WIZARD_SPARE_INTRO = 108
Const VIDEO_BLACKSMITH_WIZARD_SPARE = 109
Const VIDEO_BLACKSMITH_WIZARD_SPARE_JACKPOT = 110
Const VIDEO_BLACKSMITH_WIZARD_SPARE_SUPER_JACKPOT = 111
Const VIDEO_BLACKSMITH_WIZARD_SPARE_TOTAL = 113

'Powerups
Const VIDEO_POWERUP_GLITCH = 88
Const VIDEO_POWERUP_NORMAL = 89
Const VIDEO_POWERUP_BLACKSMITH_KILL = 94
Const VIDEO_POWERUP_BLACKSMITH_SPARE = 114

'---------------------------
' Overlay Constants
' Value = D* (900-999) PUP trigger
'---------------------------

'Basic / General
Const OVERLAY_REMOVE = 900
Const OVERLAY_SCORE_COMPETITIVE = 901
Const OVERLAY_SCORE_COOPERATIVE = 902
Const OVERLAY_KILL_OR_SPARE = 903
Const OVERLAY_KILL = 904

'Scorbit
Const OVERLAY_SCORBIT_PAIR_SMALL = 905
Const OVERLAY_SCORBIT_CLAIM_SMALL = 906
Const OVERLAY_SCORBIT_PAIR = 907
Const OVERLAY_SCORBIT_CLAIM = 908


'--------------------------
' Pinup constants
'--------------------------
' Labels
Const pPageBlank = 0
Const pPageDefault = 1

' Screens
Const pDMD = 1
Const pBackglass = 2

' Fonts
Const pFontArial = "Arial"
Const pFontRobotoBk = "Roboto Bk"

'--------------------------
' Reduced DMD constants
'--------------------------
Const eNone = 0        'Instantly displayed
Const eScrollLeft = 1  'scroll on from the right
Const eScrollRight = 2 'scroll on from the left
Const eBlink = 3       'Blink (blinks for 'TimeOn')
Const eBlinkFast = 4   'Blink (blinks for 'TimeOn') at user specified intervals (fast speed)

Const dqSize = 64

'--------------------------
' Controller
'--------------------------
Class dmdWallden
	'---------------------------
	'	PinUP
	'---------------------------
	Private PUP                 'Pinup Player Object
	Private PUPStatus           'Whether Pinup has been loaded
	
	Private pPageCur            'The current label page
	Private pDisplayPriority    'The display priority (TODO: Is this used?)
	Private curVideo            'The most recent video played
	Private curOverlay          'The current overlay active

	'---------------------------
	'	FlexDMD / Reduced DMD
	'---------------------------
	Private FlexDMD				'FlexDMD object
	Private FlexDMDStatus       'Whether FlexDMD has been loaded

	'Modified for Wallden: dqHead = earliest index in the queue we want to run; dqTail = furthest index in the queue we want to run; dqPos = Current index being run in the queue.
	Private dqHead
	Private dqTail
	Private dqPos

	Private deSpeed
	Private deBlinkSlowRate
	Private deBlinkFastRate

	Private dLine(2)
	Private deCount(2)
	Private deCountEnd(2)
	Private deBlinkCycle(2)

	Private dqText(2, 64)
	Private dqEffect(2, 64)
	Private dqTimeOn(64)
	Private dqbFlush(64)
	Private dqSound(64)

	Private DMDScene

	Public jpLocked	'Used by triggerVideo; set to true (AFTER jpFlush, if applicable) before/when calling jpDMD and we don't want triggerVideo to flush it

	Private Digits, Chars(255), Images(255)

	'---------------------------
    '   Backglass
    '---------------------------

    Public B2S

	'---------------------------
	'	Other
	'---------------------------
	Private activeSeqAttract    'Which light sequence is currently active (for attract mode tutorial sequences)
	
	Private Sub Class_Initialize
		Dim i

		PUPStatus = False
		FlexDMDStatus = False
		pPageCur = 0
		pDisplayPriority =  - 1
		activeSeqAttract = Null
		curVideo = 0
		curOverlay = 0

		'Prep reduced DMD digits / chars / images
		Digits = Array(digit001, digit002, digit003, digit004, digit005, digit006, digit007, digit008, digit009, digit010, _
			digit011, digit012, digit013, digit014, digit015, digit016, digit017, digit018, digit019, digit020,            _
			digit021, digit022, digit023, digit024, digit025, digit026, digit027, digit028, digit029, digit030,            _
			digit031, digit032, digit033, digit034, digit035, digit036, digit037, digit038, digit039, digit040,            _
			digit041)
		For i = 0 to 255:Chars(i) = "d_empty":Next

		For i = 32 to 126:Chars(i) = "d_char_" & i:Next 'ASCII

		'Special Icons
		Chars(128) = "d_char_heart"
		Chars(130) = "d_char_shield"
		Chars(131) = "d_char_skull"
		Chars(132) = "d_char_jackpot"
		Chars(133) = "d_char_time"
		Chars(134) = "d_char_monster_heart"
	End Sub
	
	Public Sub Init
		Dim i, j, DMDImage

		'Init PUP
		If usePUP = True And PUPStatus = False Then
			Set PUP = CreateObject("PinUpPlayer.PinDisplay")
			If PUP Is Nothing Then
				usePUP = False
				PUPStatus = False
			Else
				On Error Resume Next
				PUP.B2SInit "",cPuPPack 'start the Pup-Pack
				If Err.Number <> 0 Then 
					PUPStatus = False
					usePUP = False
					Err.Clear
				Else
					PUPStatus = True
				End If
				On Error GoTo 0

				PUP_initPages
			End If

			If PUPStatus = False Then
				msgbox "Warning: usePUP = True in the table script, but either PinUp Player could not be loaded or the " & cPuPPack & " PUP Pack was not found. If you do not want to use PinUp, please set usePUP = False in the table script."
			End If
		End If

		'JP's reduced DMD Init via FlexDMD
		If (UseFlexDMD = 1 Or (UseFlexDMD = 2 And Table1.ShowDT = False)) And FlexDMDStatus = False Then
			Set FlexDMD = CreateObject("FlexDMD.FlexDMD")
			If Not FlexDMD is Nothing Then
				If FlexDMDHighQuality Then
					FlexDMD.TableFile = Table1.Filename & ".vpx"
					FlexDMD.RenderMode = 2
					FlexDMD.Width = 256
					FlexDMD.Height = 64
					FlexDMD.Clear = True
					FlexDMD.GameName = cGameName
					FlexDMD.Run = True
					FlexDMDStatus = True
					Set DMDScene = FlexDMD.NewGroup("Scene")
					DMDScene.AddActor FlexDMD.NewImage("Back", "VPX.d_border")
					DMDScene.GetImage("Back").SetSize FlexDMD.Width, FlexDMD.Height
					For i = 0 to 40
						DMDScene.AddActor FlexDMD.NewImage("Dig" & i, "VPX.d_empty&dmd=2")
					Next
					For i = 0 to 19 ' Top
						DMDScene.GetImage("Dig" & i).SetBounds 8 + i * 12, 6, 12, 22
					Next
					For i = 20 to 39 ' Bottom
						DMDScene.GetImage("Dig" & i).SetBounds 8 + (i - 20) * 12, 34, 12, 22
					Next
					FlexDMD.LockRenderThread
					FlexDMD.Stage.AddActor DMDScene
					FlexDMD.UnlockRenderThread
				Else
					FlexDMD.TableFile = Table1.Filename & ".vpx"
					FlexDMD.RenderMode = 2
					FlexDMD.Width = 128
					FlexDMD.Height = 32
					FlexDMD.Clear = True
					FlexDMD.GameName = cGameName
					FlexDMD.Run = True
					FlexDMDStatus = True
					Set DMDScene = FlexDMD.NewGroup("Scene")
					DMDScene.AddActor FlexDMD.NewImage("Back", "VPX.d_border")
					DMDScene.GetImage("Back").SetSize FlexDMD.Width, FlexDMD.Height
					For i = 0 to 40
						DMDScene.AddActor FlexDMD.NewImage("Dig" & i, "VPX.d_empty&dmd=2")
					Next
					For i = 0 to 19 ' Top
						DMDScene.GetImage("Dig" & i).SetBounds 4 + i * 6, 3, 6, 11
					Next
					For i = 20 to 39 ' Bottom
						DMDScene.GetImage("Dig" & i).SetBounds 4 + (i - 20) * 6, 17, 6, 11
					Next
					FlexDMD.LockRenderThread
					FlexDMD.Stage.AddActor DMDScene
					FlexDMD.UnlockRenderThread
				End If
			Else
				MsgBox "Warning! UseFlexDMD is non-zero, but FlexDMD could not be loaded. Is it installed correctly? Please set UseFlexDMD to 0 in the table script if you do not want to use FlexDMD."
			End If
		End If

		'JP's reduced DMD Init via Flashers
		jpFlush()
		deSpeed = 20
		deBlinkSlowRate = 10
		deBlinkFastRate = 5
		For i = 0 to 2
			dLine(i) = Space(20)
			deCount(i) = 0
			deCountEnd(i) = 0
			deBlinkCycle(i) = 0
			dqTimeOn(i) = 0
			dqbFlush(i) = True
			dqSound(i) = ""
		Next
		dLine(2) = " "
		For i = 0 to 2
			For j = 0 to 64
				dqText(i, j) = ""
				dqEffect(i, j) = eNone
			Next
		Next
		jpDMD dLine(0), dLine(1), dLine(2), eNone, eNone, eNone, 25, True, ""
		If DEBUG_LOG_JP_DMD = True Then WriteToLog "JP_DMD", "jpDMD initialized"

		'Backglass
		Set B2S = CreateObject("B2S.Server")
		B2S.B2SName = "Saving Wallden"
		B2S.Run
	End Sub
	
	' Construct PUP pages for labels
	Private Sub PUP_initPages
		If (usePUP = False Or PUPStatus = False) Then Exit Sub
		
		PUP.LabelInit pDMD
		
		'labelNew <screen#>, <Labelname>, <fontName>,<size%>,<colour>,<rotation>,<xalign>,<yalign>,<xpos>,<ypos>,<PageNum>,<visible>
		'***********************************************************************'
		'<screen#>, in standard we’d set this to pDMD ( or 1)
		'<Labelname>, your name of the label. keep it short no spaces (like 8 chars) although you can call it anything really. When setting the label you will use this labelname to access the label.
		'<fontName> Windows font name, this must be exact match of OS front name. if you are using custom TTF fonts then double check the name of font names.
		'<size%>, Height as a percent of display height. 20=20% of screen height.
		'<colour>, integer value of windows color.
		'<rotation>, degrees in tenths   (900=90 degrees)
		'<xAlign>, 0= horizontal left align, 1 = center horizontal, 2= right horizontal
		'<yAlign>, 0 = top, 1 = center, 2=bottom vertical alignment
		'<xpos>, this should be 0, but if you want to ‘force’ a position you can set this. it is a % of horizontal width. 20=20% of screen width.
		'<ypos> same as xpos.
		'<PageNum> IMPORTANT… this will assign this label to this ‘page’ or group.
		'<visible> initial state of label. visible=1 show, 0 = off.
		
		'Score overlay
		PUP.LabelNew pDMD,"gameplay",pFontRobotoBk,6.5,RGB(255,255,255),0,1,0,0,75.46,1,0

		PUP.LabelNew pDMD,"p1label",pFontRobotoBk,5,RGB(255,255,64),0,1,0,13,81.29,1,0
		PUP.LabelNew pDMD,"p2label",pFontRobotoBk,5,RGB(255,255,64),0,1,0,38,81.29,1,0
		PUP.LabelNew pDMD,"p3label",pFontRobotoBk,5,RGB(255,255,64),0,1,0,63,81.29,1,0
		PUP.LabelNew pDMD,"p4label",pFontRobotoBk,5,RGB(255,255,64),0,1,0,88,81.29,1,0
		PUP.LabelNew pDMD,"p1hp",pFontRobotoBk,5,RGB(255,192,192),0,0,0,1,86.85,1,0
		PUP.LabelNew pDMD,"p2hp",pFontRobotoBk,5,RGB(255,192,192),0,0,0,26,86.85,1,0
		PUP.LabelNew pDMD,"p3hp",pFontRobotoBk,5,RGB(255,192,192),0,0,0,51,86.85,1,0
		PUP.LabelNew pDMD,"p4hp",pFontRobotoBk,5,RGB(255,192,192),0,0,0,76,86.85,1,0
		PUP.LabelNew pDMD,"p1ac",pFontRobotoBk,5,RGB(255,192,192),0,0,0,10.33,86.85,1,0
		PUP.LabelNew pDMD,"p2ac",pFontRobotoBk,5,RGB(255,192,192),0,0,0,35.33,86.85,1,0
		PUP.LabelNew pDMD,"p3ac",pFontRobotoBk,5,RGB(255,192,192),0,0,0,60.33,86.85,1,0
		PUP.LabelNew pDMD,"p4ac",pFontRobotoBk,5,RGB(255,192,192),0,0,0,85.33,86.85,1,0
		PUP.LabelNew pDMD,"p1dd",pFontRobotoBk,5,RGB(255,192,192),0,0,0,17.67,86.85,1,0
		PUP.LabelNew pDMD,"p2dd",pFontRobotoBk,5,RGB(255,192,192),0,0,0,42.67,86.85,1,0
		PUP.LabelNew pDMD,"p3dd",pFontRobotoBk,5,RGB(255,192,192),0,0,0,67.67,86.85,1,0
		PUP.LabelNew pDMD,"p4dd",pFontRobotoBk,5,RGB(255,192,192),0,0,0,92.67,86.85,1,0
		PUP.LabelNew pDMD,"p1score",pFontRobotoBk,6.5,RGB(255,255,255),0,1,0,13,92.75,1,0
		PUP.LabelNew pDMD,"p2score",pFontRobotoBk,6.5,RGB(255,255,255),0,1,0,38,92.75,1,0
		PUP.LabelNew pDMD,"p3score",pFontRobotoBk,6.5,RGB(255,255,255),0,1,0,63,92.75,1,0
		PUP.LabelNew pDMD,"p4score",pFontRobotoBk,6.5,RGB(255,255,255),0,1,0,88,92.75,1,0
		PUP.LabelNew pDMD,"coopscore",pFontRobotoBk,9.25,RGB(255,255,255),0,1,0,50,90.27,1,0

		' Scorbit
		PUP.LabelNew pDMD,"ScorbitQR",pFontRobotoBk,1,RGB(247, 170, 51),0,2,0 ,0,0,1,1
		PUP.LabelNew pDMD,"p1scorbit",pFontRobotoBk,1,RGB(255,255,255),0,2,0,0,0,1,0
		PUP.LabelNew pDMD,"p2scorbit",pFontRobotoBk,1,RGB(255,255,255),0,2,0,0,0,1,0
		PUP.LabelNew pDMD,"p3scorbit",pFontRobotoBk,1,RGB(255,255,255),0,2,0,0,0,1,0
		PUP.LabelNew pDMD,"p4scorbit",pFontRobotoBk,1,RGB(255,255,255),0,2,0,0,0,1,0
		
		' High Scores - Dates
		PUP.LabelNew pDMD,"HSD1",pFontRobotoBk,5.5,RGB(255,255,255),0,0,0,18.2,27.1,1,0
		PUP.LabelNew pDMD,"HSD2",pFontRobotoBk,5.5,RGB(255,255,255),0,0,0,18.2,36.5,1,0
		PUP.LabelNew pDMD,"HSD3",pFontRobotoBk,5.5,RGB(255,255,255),0,0,0,18.2,46.0,1,0
		PUP.LabelNew pDMD,"HSD4",pFontRobotoBk,5.5,RGB(255,255,255),0,0,0,18.2,55.5,1,0
		PUP.LabelNew pDMD,"HSD5",pFontRobotoBk,5.5,RGB(255,255,255),0,0,0,18.2,65.0,1,0
		
		' High Scores - Scores
		PUP.LabelNew pDMD,"HS1",pFontRobotoBk,5.5,RGB(255,255,255),0,2,0,65.1,27.1,1,0
		PUP.LabelNew pDMD,"HS2",pFontRobotoBk,5.5,RGB(255,255,255),0,2,0,65.1,36.5,1,0
		PUP.LabelNew pDMD,"HS3",pFontRobotoBk,5.5,RGB(255,255,255),0,2,0,65.1,46.0,1,0
		PUP.LabelNew pDMD,"HS4",pFontRobotoBk,5.5,RGB(255,255,255),0,2,0,65.1,55.5,1,0
		PUP.LabelNew pDMD,"HS5",pFontRobotoBk,5.5,RGB(255,255,255),0,2,0,65.1,65.0,1,0
		
		' High Scores - Names
		PUP.LabelNew pDMD,"HSN1",pFontRobotoBk,5.5,RGB(255,255,255),0,0,0,68.7,27.1,1,0
		PUP.LabelNew pDMD,"HSN2",pFontRobotoBk,5.5,RGB(255,255,255),0,0,0,68.7,36.5,1,0
		PUP.LabelNew pDMD,"HSN3",pFontRobotoBk,5.5,RGB(255,255,255),0,0,0,68.7,46.0,1,0
		PUP.LabelNew pDMD,"HSN4",pFontRobotoBk,5.5,RGB(255,255,255),0,0,0,68.7,55.5,1,0
		PUP.LabelNew pDMD,"HSN5",pFontRobotoBk,5.5,RGB(255,255,255),0,0,0,68.7,65.0,1,0

		' Previous Scores - Scores
		PUP.LabelNew pDMD,"PS1",pFontRobotoBk,5.5,RGB(255,255,255),0,2,0,94.01,27.1,1,0
		PUP.LabelNew pDMD,"PS2",pFontRobotoBk,5.5,RGB(255,255,255),0,2,0,94.01,36.5,1,0
		PUP.LabelNew pDMD,"PS3",pFontRobotoBk,5.5,RGB(255,255,255),0,2,0,94.01,46.0,1,0
		PUP.LabelNew pDMD,"PS4",pFontRobotoBk,5.5,RGB(255,255,255),0,2,0,94.01,55.5,1,0
		
		' Previous Scores - Names
		PUP.LabelNew pDMD,"PSN1",pFontRobotoBk,5.5,RGB(255,255,255),0,0,0,5.3,27.1,1,0
		PUP.LabelNew pDMD,"PSN2",pFontRobotoBk,5.5,RGB(255,255,255),0,0,0,5.3,36.5,1,0
		PUP.LabelNew pDMD,"PSN3",pFontRobotoBk,5.5,RGB(255,255,255),0,0,0,5.3,46.0,1,0
		PUP.LabelNew pDMD,"PSN4",pFontRobotoBk,5.5,RGB(255,255,255),0,0,0,5.3,55.5,1,0
		
		' Center labels
		PUP.LabelNew pDMD,"center_top",pFontRobotoBk,8.8,RGB(255,255,255),0,1,0,50,26.1,1,0
		PUP.LabelNew pDMD,"center_bottom",pFontRobotoBk,17.7,RGB(255,255,255),0,1,0,50,45.5,1,0

		' Death Save
		PUP.LabelNew pDMD,"death_res",pFontRobotoBk,20.74,RGB(192,255,192),0,1,0,25,46.29,1,0
		PUP.LabelNew pDMD,"death_strikes",pFontRobotoBk,20.74,RGB(255,192,192),0,1,0,75,46.29,1,0
		
		' Init with initial page
		setPage pDMD, pPageDefault
	End Sub

	Public Sub close() 'Close DMD programs
		If FlexDMDStatus = True then FlexDMD.Run = False 'Close FlexDMD
	End Sub

	'*************************
	' REDUCED DMD
	' Notice! Heavily modified
	' for Saving Wallden.
	'*************************

	Public Sub jpAdvance() 'Advance to the next item in the queue
		Dim pos
		pos = dqPos
		dqPos = dqPos + 1
		If DEBUG_LOG_JP_DMD = True Then WriteToLog "JP_DMD", "jpAdvance; dqHead = " & dqHead & "; dqPos = " & dqPos & "; dqTail = " & dqTail
		If (dqPos >= dqTail) Then
			jpLocked = False
			if (dqbFlush(pos) = True) Then
				If dqPos <= dqTail and dqPos < 64 Then
					dqHead = dqPos
				Else
					jpFlush()
				End If
			End If
			dqPos = dqHead
			if (dqPos >= dqTail) Then
				jpScoreNow()
			Else
				jpHead()
			End If
		Else
			if(dqbFlush(pos) = True) Then
				If dqPos <= dqTail and dqPos < 64 Then
					dqHead = dqPos
					jpHead()
				Else
					jpScoreNow()
				End If
			Else
				jpHead()
			End If
		End If
	End Sub

	Public Sub jpFlush()
		If DEBUG_LOG_JP_DMD = True Then WriteToLog "JP_DMD", "jpFlush"
		Dim i
		DMDTimer.Enabled = False
		DMDEffectTimer.Enabled = False
		dqPos = 0
		dqHead = 0
		dqTail = 0
		For i = 0 to 2
			deCount(i) = 0
			deCountEnd(i) = 0
			deBlinkCycle(i) = 0
		Next
		jpLocked = False
	End Sub

	Private Sub jpScore()
		If Not dqTail = 0 Then Exit Sub

		If DEBUG_LOG_JP_DMD = True Then WriteToLog "JP_DMD", "jpScore"
		If CURRENT_MODE < 1000 Then
			If DEBUG_LOG_JP_DMD = True Then WriteToLog "JP_DMD", "jpScore; no modes active"
			DMD.jpDMD "_", "_", "_", eNone, eNone, eNone, 100, True, ""
			Exit Sub
		End If

		Dim textBg
		textBg = ""

		'Determine background to use
		Select Case CURRENT_MODE
			Case MODE_NORMAL:
				If currentPlayer("obj_CASTLE") = 0 Then 'First adventure
					textBg = "d_E143_0"
				Else
					textBg = "d_E141_0"
				End If
		End Select

		'Actualizer
		Select Case CURRENT_MODE
			Case MODE_SKILLSHOT:
				If GAME_OFFICIALLY_STARTED = False Then
					DMD.jpDMD "", "", "d_E137_0", eNone, eNone, eNone, 500, True, ""
				ElseIf dataGame.data.Item("gameDifficulty") = 0 And dataGame.data.Item("ball") > 1 Then 'Zen; player can end the game with action button
					DMD.jpDMD DMD.jpCL(Left(currentPlayer("name"), 15)), "", "d_E171_0", eNone, eNone, eNone, 500, True, ""
				Else
					DMD.jpDMD DMD.jpCL(Left(currentPlayer("name"), 15)), "", "d_E138_0", eNone, eNone, eNone, 500, True, ""
				End If

			Case MODE_GLITCH_WIZARD:
				If (GameTime / 3000) Mod 2 = 0 Then
					DMD.jpDMD DMD.jpCL("Glitch Mini-wizard"), DMD.jpFL(Chr(132) & FormatScore(currentPlayerBonus("all", False) * MODE_COUNTERS_B), Chr(133) & int(Clocks.data.Item("mode").timeLeft / 1000)), "d_E183_0", eNone, eNone, eNone, 250, True, ""
				Else
					DMD.jpDMD DMD.jpCL("Light the lanterns"), DMD.jpFL(Chr(132) & FormatScore(currentPlayerBonus("all", False) * MODE_COUNTERS_B), Chr(133) & int(Clocks.data.Item("mode").timeLeft / 1000)), "d_E183_0", eNone, eNone, eNone, 250, True, ""
				End If

			Case MODE_BLACKSMITH_WIZARD_KILL
				If (GameTime / 3000) Mod 2 = 0 Then
					DMD.jpDMD DMD.jpCL("Blacksmith boss"), DMD.jpFL(FormatScore(currentPlayer("score")), Chr(130) & currentPlayer("AC")), "d_E193_0", eNone, eNone, eNone, 250, True, ""
				ElseIf MODE_COUNTERS_A = 0 Then
					DMD.jpDMD DMD.jpCL("Shoot Right Ramp"), DMD.jpFL(FormatScore(currentPlayer("score")), Chr(130) & currentPlayer("AC")), "d_E193_0", eNone, eNone, eNone, 250, True, ""
				Else
					DMD.jpDMD DMD.jpCL("Shoot Horseshoe"), DMD.jpFL(FormatScore(currentPlayer("score")), Chr(130) & currentPlayer("AC")), "d_E193_0", eNone, eNone, eNone, 250, True, ""
				End If

			Case MODE_BLACKSMITH_WIZARD_SPARE
				If (GameTime / 3000) Mod 3 = 0 Then
					DMD.jpDMD DMD.jpCL("Raid the Shop"), FormatScore(currentPlayer("score")), "d_E1109_0", eNone, eNone, eNone, 250, True, ""
				ElseIf (GameTime / 3000) Mod 3 = 1 Then
					DMD.jpDMD DMD.jpCL("Horeshoe raises " & Chr(132)), Chr(132) & FormatScore(MODE_COUNTERS_A), "d_E1109_0", eNone, eNone, eNone, 250, True, ""
				Else
					DMD.jpDMD DMD.jpCL("Hole collects " & Chr(132)), Chr(132) & FormatScore(MODE_COUNTERS_A), "d_E1109_0", eNone, eNone, eNone, 250, True, ""
				End If

			Case Else
				If DEBUG_LOG_JP_DMD = True Then WriteToLog "JP_DMD", "jpScore; showing current player score"

				If (GameTime / 3000) Mod 2 = 0 Then 'Show score
					DMD.jpDMD Left(currentPlayer("name"), 15), DMD.jpRL(FormatScore(currentPlayer("score"))), textBg, eNone, eNone, eNone, 100, True, ""
				Else 'Show stats
					Dim playerHP
					playerHP = currentPlayer("HP")
					If playerHP <= 0 Then
						If currentPlayer("dead") = 0 Then
							playerHP = "SOS"	'Show SOS when a player is about to die
						Else
							playerHP = "DEAD"
						End If
					End If
					If dataGame.data.Item("gameDifficulty") = 0 And currentPlayer("dead") = 0 Then playerHP = "?" 'Don't show player health If in zen mode

					DMD.jpDMD Left(currentPlayer("name"), 15), DMD.jpCL(Chr(128) & PlayerHP & " " & Chr(130) & currentPlayer("AC") & " " & Chr(131) & currentPlayer("drainDamage")), textBg, eNone, eNone, eNone, 100, True, ""
				End If
		End Select
	End Sub

	Public Sub jpScoreNow
		If DEBUG_LOG_JP_DMD = True Then WriteToLog "JP_DMD", "jpScoreNow"
		jpFlush
		jpScore
	End Sub

	' Modified for Wallden: bFlush = True means flush all lines queued before, and including, this one once finished (original behavior: flush all lines if queue is finished).
	Public Sub jpDMD(Text0, Text1, Text2, Effect0, Effect1, Effect2, TimeOn, bFlush, Sound)
		if(dqTail < dqSize) Then
			If DEBUG_LOG_JP_DMD = True Then WriteToLog "JP_DMD", "jpDMD | " & Text0 & " | " & Text1 & " | " & Text2 & " | dqTail = " & dqTail
			if(Text0 = "_") Then
				dqEffect(0, dqTail) = eNone
				dqText(0, dqTail) = "_"
			Else
				dqEffect(0, dqTail) = Effect0
				dqText(0, dqTail) = jpExpandLine(Text0)
			End If

			if(Text1 = "_") Then
				dqEffect(1, dqTail) = eNone
				dqText(1, dqTail) = "_"
			Else
				dqEffect(1, dqTail) = Effect1
				dqText(1, dqTail) = jpExpandLine(Text1)
			End If

			if(Text2 = "_") Then
				dqEffect(2, dqTail) = eNone
				dqText(2, dqTail) = "_"
			Else
				dqEffect(2, dqTail) = Effect2
				dqText(2, dqTail) = Text2
			End If

			dqTimeOn(dqTail) = TimeOn
			dqbFlush(dqTail) = bFlush
			dqSound(dqTail) = Sound
			dqTail = dqTail + 1
			if(dqTail = 1) Then
				jpHead()
			End If
		End If
	End Sub

	Public Sub jpHead()
		If DEBUG_LOG_JP_DMD = True Then WriteToLog "JP_DMD", "jpHead; dqHead = " & dqHead & "; dqPos = " & dqPos & "; deSpeed = " & deSpeed
		Dim i
		deCount(0) = 0
		deCount(1) = 0
		deCount(2) = 0
		DMDEffectTimer.Interval = deSpeed

		For i = 0 to 2
			Select Case dqEffect(i, dqPos)
				Case eNone:deCountEnd(i) = 1
				Case eScrollLeft:deCountEnd(i) = Len(dqText(i, dqPos) )
				Case eScrollRight:deCountEnd(i) = Len(dqText(i, dqPos) )
				Case eBlink:deCountEnd(i) = int(dqTimeOn(dqPos) / deSpeed)
					deBlinkCycle(i) = 0
				Case eBlinkFast:deCountEnd(i) = int(dqTimeOn(dqPos) / deSpeed)
					deBlinkCycle(i) = 0
			End Select
		Next
		if(dqSound(dqPos) <> "") Then
			' Disabled for Wallden; we use the queue system etc
			'	PlaySound(dqSound(dqPos) )
		End If
		DMDEffectTimer.Enabled = True
	End Sub

	Public Sub jpProcessEffectOn()
		Dim i
		Dim BlinkEffect
		Dim Temp

		BlinkEffect = False

		For i = 0 to 2
			if(deCount(i) <> deCountEnd(i) ) Then
				deCount(i) = deCount(i) + 1

				select case(dqEffect(i, dqPos) )
					case eNone:
						Temp = dqText(i, dqPos)
					case eScrollLeft:
						Temp = Right(dLine(i), 19)
						Temp = Temp & Mid(dqText(i, dqPos), deCount(i), 1)
					case eScrollRight:
						Temp = Mid(dqText(i, dqPos), 21 - deCount(i), 1)
						Temp = Temp & Left(dLine(i), 19)
					case eBlink:
						BlinkEffect = True
						if((deCount(i) MOD deBlinkSlowRate) = 0) Then
							deBlinkCycle(i) = deBlinkCycle(i) xor 1
						End If

						if(deBlinkCycle(i) = 0) Then
							Temp = dqText(i, dqPos)
						Else
							Temp = Space(20)
						End If
					case eBlinkFast:
						BlinkEffect = True
						if((deCount(i) MOD deBlinkFastRate) = 0) Then
							deBlinkCycle(i) = deBlinkCycle(i) xor 1
						End If

						if(deBlinkCycle(i) = 0) Then
							Temp = dqText(i, dqPos)
						Else
							Temp = Space(20)
						End If
				End Select

				if(dqText(i, dqPos) <> "_") Then
					dLine(i) = Temp
					jpUpdate i
				End If
			End If
		Next

		if(deCount(0) = deCountEnd(0) ) and(deCount(1) = deCountEnd(1) ) and(deCount(2) = deCountEnd(2) ) Then
			If DEBUG_LOG_JP_DMD = True Then WriteToLog "JP_DMD", "jpProcessEffectOn; effect finished for dqPos = " & dqPos

			if(dqTimeOn(dqPos) = 0) Then
				jpScoreNow()
			Else
				if(BlinkEffect = True) Then
					DMDTimer.Interval = 10
				Else
					DMDTimer.Interval = dqTimeOn(dqPos)
				End If

				DMDTimer.Enabled = True
			End If
		Else
			DMDEffectTimer.Enabled = True
		End If
	End Sub

	Public Function jpExpandLine(TempStr) 'id is the number of the dmd line
		If TempStr = "" Then
			TempStr = Space(20)
		Else
			if Len(TempStr) > Space(20) Then
				TempStr = Left(TempStr, Space(20) )
			Else
				if(Len(TempStr) < 20) Then
					TempStr = TempStr & Space(20 - Len(TempStr) )
				End If
			End If
		End If
		jpExpandLine = TempStr
	End Function

	Public Function jpFormatScore(ByVal Num) 'it returns a string with commas (as in Black's original font)
		dim i
		dim NumString

		NumString = CStr(abs(Num) )

		For i = Len(NumString) -3 to 1 step -3
			if IsNumeric(mid(NumString, i, 1) ) then
				NumString = left(NumString, i-1) & chr(asc(mid(NumString, i, 1) ) + 112) & right(NumString, Len(NumString) - i)
			end if
		Next
		jpFormatScore = NumString
	End function

	Public Function jpFL(NumString1, NumString2) 'Fill line
		Dim Temp, TempStr
		Temp = 20 - Len(NumString1) - Len(NumString2)
		TempStr = NumString1 & Space(Temp) & NumString2
		jpFL = TempStr
	End Function

	Public Function jpCL(NumString) 'center line
		Dim Temp, TempStr
		Temp = (20 - Len(NumString) ) \ 2
		TempStr = Space(Temp) & NumString & Space(Temp)
		jpCL = TempStr
	End Function

	Public Function jpRL(NumString) 'right line
		Dim Temp, TempStr
		Temp = 20 - Len(NumString)
		TempStr = Space(Temp) & NumString
		jpRL = TempStr
	End Function

	Private Sub jpUpdate(id)
		If DEBUG_LOG_JP_DMD = True Then WriteToLog "JP_DMD", "jpUpdate " & id
		Dim digit, value
		If FlexDMDStatus = True Then FlexDMD.LockRenderThread
		Select Case id
			Case 0 'top text line
				For digit = 0 to 19
					jpDisplayChar mid(dLine(0), digit + 1, 1), digit
				Next
			Case 1 'bottom text line
				For digit = 20 to 39
					jpDisplayChar mid(dLine(1), digit -19, 1), digit
				Next
			Case 2 ' back image - back animations
				If dLine(2) = "" OR dLine(2) = " " Then dLine(2) = "d_border"
				Digits(40).ImageA = dLine(2)
				If FlexDMDStatus = True Then DMDScene.GetImage("Back").Bitmap = FlexDMD.NewImage("", "VPX." & dLine(2) & "&dmd=2").Bitmap
		End Select
		If FlexDMDStatus = True Then FlexDMD.UnlockRenderThread
	End Sub

	Private Sub jpDisplayChar(achar, adigit)
		If achar = "" Then achar = " "
		achar = ASC(achar)
		Digits(adigit).ImageA = Chars(achar)
		If FlexDMDStatus = True Then DMDScene.GetImage("Dig" & adigit).Bitmap = FlexDMD.NewImage("", "VPX." & Chars(achar) & "&dmd=2&add").Bitmap
	End Sub

	'*************************
	' PUP FUNCTIONS
	'*************************

	Public Function getPUPRoot()
		If (usePUP = False Or PUPStatus = False) Then Exit Function

		getPUPRoot = PUP.getroot
	End Function
	
	'*************************
	' EVENTS
	'*************************
	
	Public Sub triggerVideo(EventNum) 'Trigger a video event (D* event in Pinup; background images for JP DMD). Should be called BEFORE setting any labels that are auto-reset in this routine.
		Dim i

		'PinUP
		If usePUP = True And PUPStatus = True Then
			'*** Specialized resets ***
			'Automatically clear labels when their respective video is no longer being shown.
			
			'High score label resets
			If ((curVideo >= VIDEO_HIGH_SCORES_ZEN And curVideo <= VIDEO_HIGH_SCORES_IMPOSSIBLE) Or curVideo = VIDEO_HIGH_SCORES_TOURNAMENT) And _
			(EventNum < VIDEO_HIGH_SCORES_ZEN Or (EventNum > VIDEO_HIGH_SCORES_IMPOSSIBLE And Not EventNum = VIDEO_HIGH_SCORES_TOURNAMENT)) Then
				PUP.LabelSet pDMD,"HSD1","",0,""
				PUP.LabelSet pDMD,"HSD2","",0,""
				PUP.LabelSet pDMD,"HSD3","",0,""
				PUP.LabelSet pDMD,"HSD4","",0,""
				PUP.LabelSet pDMD,"HSD5","",0,""
				PUP.LabelSet pDMD,"HS1","",0,""
				PUP.LabelSet pDMD,"HS2","",0,""
				PUP.LabelSet pDMD,"HS3","",0,""
				PUP.LabelSet pDMD,"HS4","",0,""
				PUP.LabelSet pDMD,"HS5","",0,""
				PUP.LabelSet pDMD,"HSN1","",0,""
				PUP.LabelSet pDMD,"HSN2","",0,""
				PUP.LabelSet pDMD,"HSN3","",0,""
				PUP.LabelSet pDMD,"HSN4","",0,""
				PUP.LabelSet pDMD,"HSN5","",0,""
			End If

			'Previous Score label resets
			If curVideo = VIDEO_PREVIOUS_SCORES Then
				PUP.LabelSet pDMD,"PS1","",0,""
				PUP.LabelSet pDMD,"PS2","",0,""
				PUP.LabelSet pDMD,"PS3","",0,""
				PUP.LabelSet pDMD,"PS4","",0,""
				PUP.LabelSet pDMD,"PSN1","",0,""
				PUP.LabelSet pDMD,"PSN2","",0,""
				PUP.LabelSet pDMD,"PSN3","",0,""
				PUP.LabelSet pDMD,"PSN4","",0,""
			End If
			
			' Center Label resets
			If curVideo = VIDEO_DEAD Or _
			curVideo = VIDEO_BLACKSMITH_SPELL Or _
			curVideo = VIDEO_GLITCH_WIZARD_SCREEN Or _
			curVideo = VIDEO_GLITCH_WIZARD_JACKPOT Or _
			curVideo = VIDEO_BLACKSMITH_WIZARD_KILL_DAMAGE Or _
			curVideo = VIDEO_BLACKSMITH_WIZARD_KILL_LOSE_AC Or _
			curVideo = VIDEO_GLITCH_WIZARD_TOTAL Then
				PUP.LabelSet pDMD,"center_top","",0,""
				PUP.LabelSet pDMD,"center_bottom","",0,""
			End If

			' Death Save labels Reset
			If curVideo = VIDEO_DEATH_SAVE_PROGRESS Then
				PUP.LabelSet pDMD,"death_res","",0,""
				PUP.LabelSet pDMD,"death_strikes","",0,""
			End If

			'*** B2S Event ***

			'We do not use PUP priorities for Wallden since we defer priorities to the queue system. When triggerVideo is called, it is always assumed to take priority over the current video, so call the stopPlayer event.
			If Not EventNum = 0 Then PUP.B2SData "D0",1
			
			PUP.B2SData "D" & EventNum,1  'send event to Pup-Pack

			'*** Specialized labels ***
			'When we ALWAYS want labels to show when certain videos are triggered, they are put here. They should be cleared in the "specialized resets" section when a different video is triggered.
			Select Case EventNum
				Case VIDEO_PREVIOUS_SCORES:
					For i = 1 to 4
						If i <= dataGame.data.Item("numPlayers") Then
							If dataGame.data.Item("gameMode") = 1 Then 'coop mode
								PUP.LabelSet pDMD,"PS" & i,FormatScore(dataPlayer(0).data.Item("score")),1,""
							Else
								PUP.LabelSet pDMD,"PS" & i,FormatScore(dataPlayer(i - 1).data.Item("score")),1,""
							End If
							PUP.LabelSet pDMD,"PSN" & i,dataPlayer(i - 1).data.Item("name"),1,""
						End If
					Next

				Case VIDEO_HIGH_SCORES_ZEN:
					PUP.LabelSet pDMD,"HSD1",dataHighScores(0).data.Item("HSD1"),1,""
					PUP.LabelSet pDMD,"HSD2",dataHighScores(0).data.Item("HSD2"),1,""
					PUP.LabelSet pDMD,"HSD3",dataHighScores(0).data.Item("HSD3"),1,""
					PUP.LabelSet pDMD,"HSD4",dataHighScores(0).data.Item("HSD4"),1,""
					PUP.LabelSet pDMD,"HSD5",dataHighScores(0).data.Item("HSD5"),1,""
					PUP.LabelSet pDMD,"HS1",FormatScore(dataHighScores(0).data.Item("HS1")),1,""
					PUP.LabelSet pDMD,"HS2",FormatScore(dataHighScores(0).data.Item("HS2")),1,""
					PUP.LabelSet pDMD,"HS3",FormatScore(dataHighScores(0).data.Item("HS3")),1,""
					PUP.LabelSet pDMD,"HS4",FormatScore(dataHighScores(0).data.Item("HS4")),1,""
					PUP.LabelSet pDMD,"HS5",FormatScore(dataHighScores(0).data.Item("HS5")),1,""
					PUP.LabelSet pDMD,"HSN1",dataHighScores(0).data.Item("HSN1"),1,""
					PUP.LabelSet pDMD,"HSN2",dataHighScores(0).data.Item("HSN2"),1,""
					PUP.LabelSet pDMD,"HSN3",dataHighScores(0).data.Item("HSN3"),1,""
					PUP.LabelSet pDMD,"HSN4",dataHighScores(0).data.Item("HSN4"),1,""
					PUP.LabelSet pDMD,"HSN5",dataHighScores(0).data.Item("HSN5"),1,""

				Case VIDEO_HIGH_SCORES_EASY:
					PUP.LabelSet pDMD,"HSD1",dataHighScores(1).data.Item("HSD1"),1,""
					PUP.LabelSet pDMD,"HSD2",dataHighScores(1).data.Item("HSD2"),1,""
					PUP.LabelSet pDMD,"HSD3",dataHighScores(1).data.Item("HSD3"),1,""
					PUP.LabelSet pDMD,"HSD4",dataHighScores(1).data.Item("HSD4"),1,""
					PUP.LabelSet pDMD,"HSD5",dataHighScores(1).data.Item("HSD5"),1,""
					PUP.LabelSet pDMD,"HS1",FormatScore(dataHighScores(1).data.Item("HS1")),1,""
					PUP.LabelSet pDMD,"HS2",FormatScore(dataHighScores(1).data.Item("HS2")),1,""
					PUP.LabelSet pDMD,"HS3",FormatScore(dataHighScores(1).data.Item("HS3")),1,""
					PUP.LabelSet pDMD,"HS4",FormatScore(dataHighScores(1).data.Item("HS4")),1,""
					PUP.LabelSet pDMD,"HS5",FormatScore(dataHighScores(1).data.Item("HS5")),1,""
					PUP.LabelSet pDMD,"HSN1",dataHighScores(1).data.Item("HSN1"),1,""
					PUP.LabelSet pDMD,"HSN2",dataHighScores(1).data.Item("HSN2"),1,""
					PUP.LabelSet pDMD,"HSN3",dataHighScores(1).data.Item("HSN3"),1,""
					PUP.LabelSet pDMD,"HSN4",dataHighScores(1).data.Item("HSN4"),1,""
					PUP.LabelSet pDMD,"HSN5",dataHighScores(1).data.Item("HSN5"),1,""

				Case VIDEO_HIGH_SCORES_NORMAL:
					PUP.LabelSet pDMD,"HSD1",dataHighScores(2).data.Item("HSD1"),1,""
					PUP.LabelSet pDMD,"HSD2",dataHighScores(2).data.Item("HSD2"),1,""
					PUP.LabelSet pDMD,"HSD3",dataHighScores(2).data.Item("HSD3"),1,""
					PUP.LabelSet pDMD,"HSD4",dataHighScores(2).data.Item("HSD4"),1,""
					PUP.LabelSet pDMD,"HSD5",dataHighScores(2).data.Item("HSD5"),1,""
					PUP.LabelSet pDMD,"HS1",FormatScore(dataHighScores(2).data.Item("HS1")),1,""
					PUP.LabelSet pDMD,"HS2",FormatScore(dataHighScores(2).data.Item("HS2")),1,""
					PUP.LabelSet pDMD,"HS3",FormatScore(dataHighScores(2).data.Item("HS3")),1,""
					PUP.LabelSet pDMD,"HS4",FormatScore(dataHighScores(2).data.Item("HS4")),1,""
					PUP.LabelSet pDMD,"HS5",FormatScore(dataHighScores(2).data.Item("HS5")),1,""
					PUP.LabelSet pDMD,"HSN1",dataHighScores(2).data.Item("HSN1"),1,""
					PUP.LabelSet pDMD,"HSN2",dataHighScores(2).data.Item("HSN2"),1,""
					PUP.LabelSet pDMD,"HSN3",dataHighScores(2).data.Item("HSN3"),1,""
					PUP.LabelSet pDMD,"HSN4",dataHighScores(2).data.Item("HSN4"),1,""
					PUP.LabelSet pDMD,"HSN5",dataHighScores(2).data.Item("HSN5"),1,""

				Case VIDEO_HIGH_SCORES_HARD:
					PUP.LabelSet pDMD,"HSD1",dataHighScores(3).data.Item("HSD1"),1,""
					PUP.LabelSet pDMD,"HSD2",dataHighScores(3).data.Item("HSD2"),1,""
					PUP.LabelSet pDMD,"HSD3",dataHighScores(3).data.Item("HSD3"),1,""
					PUP.LabelSet pDMD,"HSD4",dataHighScores(3).data.Item("HSD4"),1,""
					PUP.LabelSet pDMD,"HSD5",dataHighScores(3).data.Item("HSD5"),1,""
					PUP.LabelSet pDMD,"HS1",FormatScore(dataHighScores(3).data.Item("HS1")),1,""
					PUP.LabelSet pDMD,"HS2",FormatScore(dataHighScores(3).data.Item("HS2")),1,""
					PUP.LabelSet pDMD,"HS3",FormatScore(dataHighScores(3).data.Item("HS3")),1,""
					PUP.LabelSet pDMD,"HS4",FormatScore(dataHighScores(3).data.Item("HS4")),1,""
					PUP.LabelSet pDMD,"HS5",FormatScore(dataHighScores(3).data.Item("HS5")),1,""
					PUP.LabelSet pDMD,"HSN1",dataHighScores(3).data.Item("HSN1"),1,""
					PUP.LabelSet pDMD,"HSN2",dataHighScores(3).data.Item("HSN2"),1,""
					PUP.LabelSet pDMD,"HSN3",dataHighScores(3).data.Item("HSN3"),1,""
					PUP.LabelSet pDMD,"HSN4",dataHighScores(3).data.Item("HSN4"),1,""
					PUP.LabelSet pDMD,"HSN5",dataHighScores(3).data.Item("HSN5"),1,""

				Case VIDEO_HIGH_SCORES_IMPOSSIBLE:
					PUP.LabelSet pDMD,"HSD1",dataHighScores(4).data.Item("HSD1"),1,""
					PUP.LabelSet pDMD,"HSD2",dataHighScores(4).data.Item("HSD2"),1,""
					PUP.LabelSet pDMD,"HSD3",dataHighScores(4).data.Item("HSD3"),1,""
					PUP.LabelSet pDMD,"HSD4",dataHighScores(4).data.Item("HSD4"),1,""
					PUP.LabelSet pDMD,"HSD5",dataHighScores(4).data.Item("HSD5"),1,""
					PUP.LabelSet pDMD,"HS1",FormatScore(dataHighScores(4).data.Item("HS1")),1,""
					PUP.LabelSet pDMD,"HS2",FormatScore(dataHighScores(4).data.Item("HS2")),1,""
					PUP.LabelSet pDMD,"HS3",FormatScore(dataHighScores(4).data.Item("HS3")),1,""
					PUP.LabelSet pDMD,"HS4",FormatScore(dataHighScores(4).data.Item("HS4")),1,""
					PUP.LabelSet pDMD,"HS5",FormatScore(dataHighScores(4).data.Item("HS5")),1,""
					PUP.LabelSet pDMD,"HSN1",dataHighScores(4).data.Item("HSN1"),1,""
					PUP.LabelSet pDMD,"HSN2",dataHighScores(4).data.Item("HSN2"),1,""
					PUP.LabelSet pDMD,"HSN3",dataHighScores(4).data.Item("HSN3"),1,""
					PUP.LabelSet pDMD,"HSN4",dataHighScores(4).data.Item("HSN4"),1,""
					PUP.LabelSet pDMD,"HSN5",dataHighScores(4).data.Item("HSN5"),1,""

				Case VIDEO_HIGH_SCORES_TOURNAMENT:
					PUP.LabelSet pDMD,"HSD1",dataHighScores(5).data.Item("HSD1"),1,""
					PUP.LabelSet pDMD,"HSD2",dataHighScores(5).data.Item("HSD2"),1,""
					PUP.LabelSet pDMD,"HSD3",dataHighScores(5).data.Item("HSD3"),1,""
					PUP.LabelSet pDMD,"HSD4",dataHighScores(5).data.Item("HSD4"),1,""
					PUP.LabelSet pDMD,"HSD5",dataHighScores(5).data.Item("HSD5"),1,""
					PUP.LabelSet pDMD,"HS1",FormatScore(dataHighScores(5).data.Item("HS1")),1,""
					PUP.LabelSet pDMD,"HS2",FormatScore(dataHighScores(5).data.Item("HS2")),1,""
					PUP.LabelSet pDMD,"HS3",FormatScore(dataHighScores(5).data.Item("HS3")),1,""
					PUP.LabelSet pDMD,"HS4",FormatScore(dataHighScores(5).data.Item("HS4")),1,""
					PUP.LabelSet pDMD,"HS5",FormatScore(dataHighScores(5).data.Item("HS5")),1,""
					PUP.LabelSet pDMD,"HSN1",dataHighScores(5).data.Item("HSN1"),1,""
					PUP.LabelSet pDMD,"HSN2",dataHighScores(5).data.Item("HSN2"),1,""
					PUP.LabelSet pDMD,"HSN3",dataHighScores(5).data.Item("HSN3"),1,""
					PUP.LabelSet pDMD,"HSN4",dataHighScores(5).data.Item("HSN4"),1,""
					PUP.LabelSet pDMD,"HSN5",dataHighScores(5).data.Item("HSN5"),1,""

				Case VIDEO_BLACKSMITH_SPELL:
					PUP.LabelSet pDMD,"center_top","BLACKSMITH spelled",1,""
					PUP.LabelSet pDMD,"center_bottom","Armor Class " & currentPlayer("AC") & " / 20",1,""

				Case VIDEO_GLITCH_WIZARD_SCREEN:
					PUP.LabelSet pDMD,"center_bottom",FormatScore(currentPlayerBonus("all", False) * MODE_COUNTERS_B),1,""

				Case VIDEO_GLITCH_WIZARD_TOTAL:
					PUP.LabelSet pDMD,"center_top","Total",1,""
					PUP.LabelSet pDMD,"center_bottom",FormatScore(MODE_VALUES),1,""
			End Select
		End If

		'JP DMD
		If DMD.jpLocked = False Then DMD.jpFlush 'It is always assumed triggerVideo takes priority over the current DMD sequence unless jpLocked is true.
		Select Case (EventNum)
			Case VIDEO_CLEAR:
				DMD.jpDMD "", "", "", eNone, eNone, eNone, 25, True, ""

			Case VIDEO_SELF_TEST:
				DMD.jpDMD "", "", "d_E11_0", eNone, eNone, eNone, 666*4, False, ""
				DMD.jpDMD "", "", "d_E11_1", eNone, eNone, eNone, 666*4, False, ""
				DMD.jpDMD "", "", "d_E11_2", eNone, eNone, eNone, 666*4, False, ""
				DMD.jpDMD "", "", "d_E11_3", eNone, eNone, eNone, 666*2, False, ""
				DMD.jpDMD "", "", "d_E11_4", eNone, eNone, eNone, 666, False, ""
				DMD.jpDMD "", "", "d_E11_5", eNone, eNone, eNone, 666, False, ""
				DMD.jpDMD "", "", "d_E11_6", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD "", "", "d_E11_7", eNone, eNone, eNone, 8000, True, ""

			Case VIDEO_VPW_LOGO:
				DMD.jpDMD "", "", "d_E12_0", eNone, eNone, eNone, 3000, True, ""

			Case VIDEO_TESTS_PASSED:
				DMD.jpDMD DMD.jpCL("Self test"), DMD.jpCL("Passed"), "", eNone, eNone, eNone, 3000, True, ""

			Case VIDEO_BLACK_BG:

			Case VIDEO_FREE_PLAY:
				DMD.jpDMD DMD.jpCL("Free Play"), DMD.jpCL("Press Start"), "", eNone, eBlink, eNone, 5000, True, ""

			Case VIDEO_VPW_PRESENTS:
				DMD.jpDMD "", "", "d_E16_0", eNone, eNone, eNone, 3000, True, ""

			Case VIDEO_RGB_TEST:
				DMD.jpDMD "", "", "d_E17_0", eNone, eNone, eNone, 500, False, ""
				DMD.jpDMD "", "", "d_E17_1", eNone, eNone, eNone, 500, False, ""
				DMD.jpDMD "", "", "d_E17_2", eNone, eNone, eNone, 500, True, ""

			Case VIDEO_SAVING_WALLDEN:
				DMD.jpDMD "", "", "d_E18_0", eNone, eNone, eNone, 3000, True, ""

			Case VIDEO_PREVIOUS_SCORES:
				DMD.jpDMD DMD.jpCL("Previous"), DMD.jpCL("Game Scores"), "d_E175_0", eNone, eNone, eNone, 2500, False, ""
				For i = 1 to 4
					If i <= dataGame.data.Item("numPlayers") Then
						If dataGame.data.Item("gameMode") = 1 Then 'coop mode
							DMD.jpDMD "" & Left(dataPlayer(i - 1).data.Item("name"), 15), DMD.jpRL(FormatScore(dataPlayer(0).data.Item("score"))), "d_E175_0", eScrollLeft, eScrollLeft, eNone, 2500, False, ""
						Else
							DMD.jpDMD "" & Left(dataPlayer(i - 1).data.Item("name"), 15), DMD.jpRL(FormatScore(dataPlayer(i - 1).data.Item("score"))), "d_E175_0", eScrollLeft, eScrollLeft, eNone, 2500, False, ""
						End If
					End If
				Next

			Case VIDEO_HIGH_SCORES_ZEN:
				DMD.jpDMD DMD.jpCL("High Scores"), DMD.jpCL("Zen"), "d_E19_0", eNone, eBlink, eNone, 2500, False, ""
				DMD.jpDMD "1) " & Left(dataHighScores(0).data.Item("HSN1"), 15), DMD.jpRL(FormatScore(dataHighScores(0).data.Item("HS1"))), "d_E19_0", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
				DMD.jpDMD "2) " & Left(dataHighScores(0).data.Item("HSN2"), 15), DMD.jpRL(FormatScore(dataHighScores(0).data.Item("HS2"))), "d_E19_0", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
				DMD.jpDMD "3) " & Left(dataHighScores(0).data.Item("HSN3"), 15), DMD.jpRL(FormatScore(dataHighScores(0).data.Item("HS3"))), "d_E19_0", eScrollLeft, eScrollLeft, eNone, 2000, True, ""

			Case VIDEO_HIGH_SCORES_EASY:
				DMD.jpDMD DMD.jpCL("High Scores"), DMD.jpCL("Easy"), "d_E110_0", eNone, eBlink, eNone, 2500, False, ""
				DMD.jpDMD "1) " & Left(dataHighScores(1).data.Item("HSN1"), 15), DMD.jpRL(FormatScore(dataHighScores(1).data.Item("HS1"))), "d_E110_0", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
				DMD.jpDMD "2) " & Left(dataHighScores(1).data.Item("HSN2"), 15), DMD.jpRL(FormatScore(dataHighScores(1).data.Item("HS2"))), "d_E110_0", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
				DMD.jpDMD "3) " & Left(dataHighScores(1).data.Item("HSN3"), 15), DMD.jpRL(FormatScore(dataHighScores(1).data.Item("HS3"))), "d_E110_0", eScrollLeft, eScrollLeft, eNone, 2000, True, ""

			Case VIDEO_HIGH_SCORES_NORMAL:
				DMD.jpDMD DMD.jpCL("High Scores"), DMD.jpCL("Normal"), "d_E111_0", eNone, eBlink, eNone, 2500, False, ""
				DMD.jpDMD "1) " & Left(dataHighScores(2).data.Item("HSN1"), 15), DMD.jpRL(FormatScore(dataHighScores(2).data.Item("HS1"))), "d_E111_0", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
				DMD.jpDMD "2) " & Left(dataHighScores(2).data.Item("HSN2"), 15), DMD.jpRL(FormatScore(dataHighScores(2).data.Item("HS2"))), "d_E111_0", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
				DMD.jpDMD "3) " & Left(dataHighScores(2).data.Item("HSN3"), 15), DMD.jpRL(FormatScore(dataHighScores(2).data.Item("HS3"))), "d_E111_0", eScrollLeft, eScrollLeft, eNone, 2000, True, ""

			Case VIDEO_HIGH_SCORES_HARD:
				DMD.jpDMD DMD.jpCL("High Scores"), DMD.jpCL("Hard"), "d_E112_0", eNone, eBlink, eNone, 2500, False, ""
				DMD.jpDMD "1) " & Left(dataHighScores(3).data.Item("HSN1"), 15), DMD.jpRL(FormatScore(dataHighScores(3).data.Item("HS1"))), "d_E112_0", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
				DMD.jpDMD "2) " & Left(dataHighScores(3).data.Item("HSN2"), 15), DMD.jpRL(FormatScore(dataHighScores(3).data.Item("HS2"))), "d_E112_0", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
				DMD.jpDMD "3) " & Left(dataHighScores(3).data.Item("HSN3"), 15), DMD.jpRL(FormatScore(dataHighScores(3).data.Item("HS3"))), "d_E112_0", eScrollLeft, eScrollLeft, eNone, 2000, True, ""

			Case VIDEO_HIGH_SCORES_IMPOSSIBLE:
				DMD.jpDMD DMD.jpCL("High Scores"), DMD.jpCL("Impossible"), "d_E113_0", eNone, eBlink, eNone, 2500, False, ""
				DMD.jpDMD "1) " & Left(dataHighScores(4).data.Item("HSN1"), 15), DMD.jpRL(FormatScore(dataHighScores(4).data.Item("HS1"))), "d_E113_0", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
				DMD.jpDMD "2) " & Left(dataHighScores(4).data.Item("HSN2"), 15), DMD.jpRL(FormatScore(dataHighScores(4).data.Item("HS2"))), "d_E113_0", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
				DMD.jpDMD "3) " & Left(dataHighScores(4).data.Item("HSN3"), 15), DMD.jpRL(FormatScore(dataHighScores(4).data.Item("HS3"))), "d_E113_0", eScrollLeft, eScrollLeft, eNone, 2000, True, ""

			Case VIDEO_HIGH_SCORES_TOURNAMENT:
				DMD.jpDMD DMD.jpCL("Tournament"), DMD.jpCL("High Scores"), "d_E155_0", eBlink, eNone, eNone, 2500, False, ""
				DMD.jpDMD "1) " & Left(dataHighScores(5).data.Item("HSN1"), 15), DMD.jpRL(FormatScore(dataHighScores(5).data.Item("HS1"))), "d_E155_0", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
				DMD.jpDMD "2) " & Left(dataHighScores(5).data.Item("HSN2"), 15), DMD.jpRL(FormatScore(dataHighScores(5).data.Item("HS2"))), "d_E155_0", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
				DMD.jpDMD "3) " & Left(dataHighScores(5).data.Item("HSN3"), 15), DMD.jpRL(FormatScore(dataHighScores(5).data.Item("HS3"))), "d_E155_0", eScrollLeft, eScrollLeft, eNone, 2000, True, ""

			Case VIDEO_TUTORIAL_INTRO:
				DMD.jpDMD "", "", "d_E114_0", eNone, eNone, eNone, 5000, True, ""

			Case VIDEO_TUTORIAL_CASTLE:
				DMD.jpDMD DMD.jpCL("Objective:"), DMD.jpCL("CASTLE"), "d_E115_0", eNone, eBlink, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Spell TRAIN to"), DMD.jpCL("light scoop modes"), "d_E115_0", eNone, eNone, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Do 3 modes to"), DMD.jpCL("battle a mini-boss"), "d_E115_0", eNone, eNone, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Battle 3 bosses"), DMD.jpCL("to complete CASTLE"), "d_E115_0", eNone, eNone, eNone, 2500, True, ""

			Case VIDEO_TUTORIAL_CHASER:
				DMD.jpDMD DMD.jpCL("Objective:"), DMD.jpCL("CHASER"), "d_E116_0", eNone, eBlink, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Spell CHASER at"), DMD.jpCL("drop targets"), "d_E116_0", eNone, eNone, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("to start"), DMD.jpCL("hurry-up phase"), "d_E116_0", eNone, eNone, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Complete 3 phases"), DMD.jpCL("to complete CHASER"), "d_E116_0", eNone, eNone, eNone, 2500, True, ""

			Case VIDEO_TUTORIAL_DRAGON:
				DMD.jpDMD DMD.jpCL("Objective:"), DMD.jpCL("DRAGON"), "d_E117_0", eNone, eBlink, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Scoop spells DRAGON"), DMD.jpCL("to start battle"), "d_E117_0", eNone, eNone, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Shoot crossbow to"), DMD.jpCL("deal dragon damage"), "d_E117_0", eNone, eNone, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Defeat the 3 dragons"), DMD.jpCL("to complete DRAGON"), "d_E117_0", eNone, eNone, eNone, 2500, True, ""

			Case VIDEO_TUTORIAL_VIKING:
				DMD.jpDMD DMD.jpCL("Objective:"), DMD.jpCL("VIKING"), "d_E118_0", eNone, eBlink, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Orbits spell VIKING"), DMD.jpCL("for timed multiball"), "d_E118_0", eNone, eNone, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Shoot lit shots"), DMD.jpCL("to defeat vikings"), "d_E118_0", eNone, eNone, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Defeat the 8 vikings"), DMD.jpCL("to complete VIKING"), "d_E118_0", eNone, eNone, eNone, 2500, True, ""

			Case VIDEO_TUTORIAL_ESCAPE:
				DMD.jpDMD DMD.jpCL("Objective:"), DMD.jpCL("ESCAPE"), "d_E119_0", eNone, eBlink, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Ramps spell ESCAPE"), DMD.jpCL("for multiball"), "d_E119_0", eNone, eNone, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Get all 3 balls"), DMD.jpCL("in center hole"), "d_E119_0", eNone, eNone, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("without draining"), DMD.jpCL("to complete ESCAPE"), "d_E119_0", eNone, eNone, eNone, 2500, True, ""

			Case VIDEO_TUTORIAL_SNIPER:
				DMD.jpDMD DMD.jpCL("Objective:"), DMD.jpCL("SNIPER"), "d_E120_0", eNone, eBlink, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Spell SNIPER"), DMD.jpCL("at crossbow"), "d_E120_0", eNone, eNone, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Lock balls to"), DMD.jpCL("start a multiball"), "d_E120_0", eNone, eNone, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Collect 50 jackpots"), DMD.jpCL("to complete SNIPER"), "d_E120_0", eNone, eNone, eNone, 2500, True, ""

			Case VIDEO_TUTORIAL_BLACKSMITH:
				DMD.jpDMD DMD.jpCL("Objective:"), DMD.jpCL("BLACKSMITH"), "d_E121_0", eNone, eBlink, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Horseshoe spells"), DMD.jpCL("BLACKSMITH"), "d_E121_0", eNone, eNone, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("to increase your"), DMD.jpCL("AC / ball shield"), "d_E121_0", eNone, eNone, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Reach 20 AC to"), DMD.jpCL("complete BLACKSMITH"), "d_E121_0", eNone, eNone, eNone, 2500, True, ""

			Case VIDEO_TUTORIAL_FINAL_BOSS:
				DMD.jpDMD DMD.jpCL("Objective:"), DMD.jpCL("FINAL BOSS"), "d_E122_0", eNone, eBlink, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Complete the other"), DMD.jpCL("7 objectives"), "d_E122_0", eNone, eNone, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("to battle the"), DMD.jpCL("FINAL BOSS"), "d_E122_0", eNone, eNone, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("and save"), DMD.jpCL("Wallden"), "d_E122_0", eNone, eNone, eNone, 2500, True, ""

			Case VIDEO_TUTORIAL_FINAL_JUDGMENT:
				DMD.jpDMD DMD.jpCL("Objective:"), DMD.jpCL("?"), "d_E123_0", eNone, eBlink, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("After battling"), DMD.jpCL("the FINAL BOSS,"), "d_E123_0", eNone, eNone, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("your heart"), DMD.jpCL("will be judged"), "d_E123_0", eNone, eNone, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("by the archangel"), DMD.jpCL("of wallden"), "d_E123_0", eNone, eNone, eNone, 2500, True, ""

			Case VIDEO_CONTRIBUTORS_INTRO:
				DMD.jpDMD DMD.jpCL("Special thanks to"), DMD.jpCL("these contributors"), "", eNone, eNone, eNone, 5000, True, ""

			Case VIDEO_CONTRIBUTORS_ARELYEL:
				DMD.jpDMD DMD.jpCL("Arelyel Krele"), DMD.jpCL("Project Lead"), "d_E125_0", eScrollLeft, eScrollRight, eNone, 4000, False, ""
				DMD.jpDMD DMD.jpCL("Arelyel Krele"), DMD.jpCL("Scripting"), "d_E125_0", eNone, eScrollRight, eNone, 4000, True, ""

			Case VIDEO_CONTRIBUTORS_DONKEYKLONK:
				DMD.jpDMD DMD.jpCL("DonkeyKlonk"), DMD.jpCL("Table Layout"), "", eScrollLeft, eScrollRight, eNone, 8000, True, ""

			Case VIDEO_CONTRIBUTORS_FLUX:
				DMD.jpDMD DMD.jpCL("Flux"), DMD.jpCL("Light Controller"), "", eScrollLeft, eScrollRight, eNone, 8000, True, ""

			Case VIDEO_CONTRIBUTORS_GRAHAM_GROWAT:
				DMD.jpDMD DMD.jpCL("Graham Growat"), DMD.jpCL("Voice - Trainer"), "", eScrollLeft, eScrollRight, eNone, 8000, True, ""

			Case VIDEO_CONTRIBUTORS_HARMONY:
				DMD.jpDMD DMD.jpCL("Harmony"), DMD.jpCL("Voice - Instruction"), "", eScrollLeft, eScrollRight, eNone, 8000, True, ""

			Case VIDEO_CONTRIBUTORS_SMAUG:
				DMD.jpDMD DMD.jpCL("Smaug"), DMD.jpCL("Voice - Blacksmith"), "d_E176_0", eScrollLeft, eScrollRight, eNone, 8000, True, ""

			Case VIDEO_CONTRIBUTORS_OTHER:
				DMD.jpDMD DMD.jpCL("For other credits"), DMD.jpCL("and contributors,"), "", eNone, eNone, eNone, 7500, False, ""
				DMD.jpDMD DMD.jpCL("check out the"), DMD.jpCL("table script"), "", eNone, eNone, eNone, 7500, True, ""

			Case VIDEO_DIFFICULTY_ZEN:
				DMD.jpDMD "", "", "d_E131_0", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("Play as long"), DMD.jpCL("as you like"), "", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("Take a score penalty"), DMD.jpCL("for HP damage"), "", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("Starting AC:"), DMD.jpCL("12 / 20"), "", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("Drain damage:"), DMD.jpCL("@ - -"), "", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("Battle difficulty:"), DMD.jpCL("@ - -"), "", eNone, eNone, eNone, 3000, False, ""

			Case VIDEO_DIFFICULTY_EASY:
				DMD.jpDMD "", "", "d_E132_0", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("Max death saves:"), DMD.jpCL("5"), "", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("Starting AC:"), DMD.jpCL("11 / 20"), "", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("Drain damage:"), DMD.jpCL("@ - -"), "", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("Battle difficulty:"), DMD.jpCL("@ - -"), "", eNone, eNone, eNone, 3000, False, ""

			Case VIDEO_DIFFICULTY_NORMAL:
				DMD.jpDMD "", "", "d_E133_0", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("Max death saves:"), DMD.jpCL("3"), "", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("Starting AC:"), DMD.jpCL("10 / 20"), "", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("Drain damage:"), DMD.jpCL("@ @ -"), "", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("Battle difficulty:"), DMD.jpCL("@ @ -"), "", eNone, eNone, eNone, 3000, False, ""

			Case VIDEO_DIFFICULTY_HARD:
				DMD.jpDMD "", "", "d_E134_0", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("Max death saves:"), DMD.jpCL("1"), "", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("Starting AC:"), DMD.jpCL("9 / 20"), "", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("Drain damage:"), DMD.jpCL("@ @ @"), "", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("Battle difficulty:"), DMD.jpCL("@ @ @"), "", eNone, eNone, eNone, 3000, False, ""

			Case VIDEO_DIFFICULTY_IMPOSSIBLE:
				DMD.jpDMD "", "", "d_E135_0", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("Lose 1 HP"), DMD.jpCL("every 5 seconds"), "", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("but more starting"), DMD.jpCL("and healing HP."), "", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("No"), DMD.jpCL("death saves!"), "", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("Starting AC:"), DMD.jpCL("8 / 20"), "", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("Drain damage:"), DMD.jpCL("@ @ @"), "", eNone, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("Battle difficulty:"), DMD.jpCL("@ @ @"), "", eNone, eNone, eNone, 3000, False, ""

			Case VIDEO_SKILLSHOT_ADD_PLAYERS:
				DMD.jpDMD "", "", "d_E137_0", eNone, eNone, eNone, 500, True, ""

			Case VIDEO_SKILLSHOT:
				DMD.jpDMD DMD.jpCL(Left(currentPlayer("name"), 15)), "", "d_E138_0", eNone, eNone, eNone, 500, True, ""

			Case VIDEO_SKILLSHOT_COMPLETE:
				DMD.jpDMD DMD.jpCL("SKILLSHOT!"), DMD.jpCL("3,000,000 + Letters"), "", eNone, eBlink, eNone, 3000, True, ""

			Case VIDEO_SKILLSHOT_FAILED:
				DMD.jpDMD DMD.jpCL("Skillshot"), DMD.jpCL("failed!"), "", eNone, eBlink, eNone, 3000, True, ""

			'Case VIDEO_GAMEPLAY_NORMAL: (uses Else because we want this to show player's score)

			'Case VIDEO_GAMEPLAY_INSTRUCTIONS: (uses promptInstructions)

			Case VIDEO_START_GAME:
				DMD.jpDMD "", "", "d_E142_0", eNone, eNone, eNone, 700, False, ""
				DMD.jpDMD "", "", "d_E142_1", eNone, eNone, eNone, 1000, False, ""
				DMD.jpDMD "", "", "d_E142_2", eNone, eNone, eNone, 920, False, ""
				DMD.jpDMD "", "", "d_E142_3", eNone, eNone, eNone, 1000, False, ""
				DMD.jpDMD DMD.jpCL("Wall'd"), DMD.jpCL("en!"), "d_E142_3", eBlink, eBlink, eNone, 6630, True, ""

			Case VIDEO_BALL_SHIELD:
				DMD.jpDMD DMD.jpCL("Ball"), DMD.jpCL("Shield"), "d_E144_0", eBlinkFast, eBlinkFast, eNone, 2000, True, ""

			'Case VIDEO_DAMAGE_RECEIVED: Handled in DMD.damageReceived because we want to show how much damage the player took

			Case VIDEO_PLAYER_1:
				DMD.jpDMD DMD.jpCL("PLAYER 1"), DMD.jpCL("IS UP"), "", eBlink, eBlink, eNone, 3000, True, ""

			Case VIDEO_PLAYER_2:
				DMD.jpDMD DMD.jpCL("PLAYER 2"), DMD.jpCL("IS UP"), "", eBlink, eBlink, eNone, 3000, True, ""

			Case VIDEO_PLAYER_3:
				DMD.jpDMD DMD.jpCL("PLAYER 3"), DMD.jpCL("IS UP"), "", eBlink, eBlink, eNone, 3000, True, ""

			Case VIDEO_PLAYER_4:
				DMD.jpDMD DMD.jpCL("PLAYER 4"), DMD.jpCL("IS UP"), "", eBlink, eBlink, eNone, 3000, True, ""

			Case VIDEO_BALL_SEARCH:
				DMD.jpDMD DMD.jpCL("Ball Search"), DMD.jpCL("Leave table alone"), "d_E150_0", eNone, eScrollLeft, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Ball Search"), DMD.jpCL("Don't touch flippers"), "d_E150_0", eNone, eScrollLeft, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Ball Search"), DMD.jpCL("Do not nudge"), "d_E150_0", eNone, eScrollLeft, eNone, 2500, False, "" 'Sequence should loop; do not flush

			Case VIDEO_BALL_MISSING:
				DMD.jpDMD DMD.jpCL("Ball Missing!"), DMD.jpCL("Game paused"), "d_E151_0", eNone, eScrollLeft, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Ball Missing!"), DMD.jpCL("Consult operator"), "d_E151_0", eNone, eScrollLeft, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Ball Missing!"), DMD.jpCL("Remove glass"), "d_E151_0", eNone, eScrollLeft, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Ball Missing!"), DMD.jpCL("Dislodge ball"), "d_E151_0", eNone, eScrollLeft, eNone, 2500, False, "" 'Sequence should loop; do not flush

			Case VIDEO_DEAD:
				DMD.jpDMD DMD.jpCL(Left(currentPlayer("name"), 15) & ","), DMD.jpCL("You're dead!"), "d_E153_0", eNone, eBlink, eNone, 3000, True, ""

			Case VIDEO_POWERUP_COLLECTED:
				DMD.jpDMD "", "", "d_E156_0", eNone, eNone, eNone, 5000, True, ""

			Case VIDEO_GAME_OVER_FAILED:
				DMD.jpDMD "", "", "d_E157_0", eNone, eNone, eNone, 2000, False, ""
				DMD.jpDMD DMD.jpCL("GAME OVER"), "", "d_E157_0", eScrollLeft, eNone, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("GAME OVER"), DMD.jpCL("MISSION FAILED!"), "d_E157_0", eNone, eScrollRight, eNone, 500, False, ""
				DMD.jpDMD DMD.jpCL("GAME OVER"), DMD.jpCL("MISSION FAILED!"), "d_E157_0", eNone, eBlink, eNone, 3500, False, ""
				DMD.jpDMD DMD.jpCL("GAME OVER"), DMD.jpCL("MISSION FAILED!"), "d_E157_0", eNone, eNone, eNone, 4000, True, ""

			Case VIDEO_BONUSX_1:
				DMD.jpDMD DMD.jpCL("BONUS X"), DMD.jpCL("^      "), "d_E158_0", eNone, eBlinkFast, eNone, 3000, True, ""

			Case VIDEO_BONUSX_2:
				DMD.jpDMD DMD.jpCL("BONUS X"), DMD.jpCL("^^     "), "d_E158_0", eNone, eBlinkFast, eNone, 3000, True, ""

			Case VIDEO_BONUSX_3:
				DMD.jpDMD DMD.jpCL("BONUS X"), DMD.jpCL("^^^    "), "d_E158_0", eNone, eBlinkFast, eNone, 3000, True, ""

			Case VIDEO_BONUSX_4:
				DMD.jpDMD DMD.jpCL("BONUS X"), DMD.jpCL("^^^^   "), "d_E158_0", eNone, eBlinkFast, eNone, 3000, True, ""

			Case VIDEO_BONUSX_5:
				DMD.jpDMD DMD.jpCL("BONUS X"), DMD.jpCL("^^^^^  "), "d_E158_0", eNone, eBlinkFast, eNone, 3000, True, ""

			Case VIDEO_BONUSX_6:
				DMD.jpDMD DMD.jpCL("BONUS X"), DMD.jpCL("" & currentPlayer("bonusX") & " X"), "d_E158_0", eNone, eBlink, eNone, 5000, True, ""

			'Case VIDEO_RESUME_GAME: Handled separately to display a countdown

			Case VIDEO_RESUME_GAME_SCORBIT:
				DMD.jpDMD DMD.jpCL("Scorbit disabled"), DMD.jpCL("for resumed games"), "", eBlink, eNone, eNone, 5000, True, ""

			Case VIDEO_RESUME_GAME_DRAIN_PENALTY:
				DMD.jpDMD DMD.jpCL("Current player"), DMD.jpCL("gets drain penalty"), "", eNone, eNone, eNone, 10000, True, ""

			Case VIDEO_RESUME_GAME_NEXT_PLAYER:
				DMD.jpDMD DMD.jpCL("Resuming on"), DMD.jpCL("next player"), "", eNone, eNone, eNone, 10000, True, ""

			Case VIDEO_RESUME_GAME_CURRENT_PLAYER:
				DMD.jpDMD DMD.jpCL("Resuming on"), DMD.jpCL("current player"), "", eNone, eNone, eNone, 10000, True, ""

			Case VIDEO_ZEN_END_WAIT:
				DMD.jpDMD "", "", "d_E169_0", eNone, eNone, eNone, 5000, False, ""

			Case VIDEO_ZEN_DAMAGE_PENALTY:
				DMD.jpDMD DMD.jpCL("Damage Penalty"), DMD.jpCL("???,???,???"), "d_E170_0", eNone, eNone, eNone, 5000, False, ""

			Case VIDEO_SKILLSHOT_ZEN_END:
				DMD.jpDMD DMD.jpCL(Left(currentPlayer("name"), 15)), "", "d_E171_0", eNone, eNone, eNone, 500, True, ""

			Case VIDEO_DEATH_SAVE_INTRO:
				DMD.jpDMD DMD.jpCL("Your quest"), DMD.jpCL("isn't over!"), "d_E172_0", eBlink, eBlink, eNone, 2030, False, ""
				DMD.jpDMD DMD.jpCL("You"), DMD.jpCL("must live!"), "d_E172_0", eBlink, eBlink, eNone, 2000, False, ""
				DMD.jpDMD DMD.jpCL("Press action button"), DMD.jpCL("to heart beat"), "d_E172_0", eScrollLeft, eScrollRight, eNone, 2700, False, ""
				DMD.jpDMD DMD.jpCL("to revive"), DMD.jpCL("your character!"), "d_E172_0", eScrollLeft, eScrollRight, eNone, 1600, True, ""

			'Case VIDEO_DEATH_SAVE_PROGRESS: Handled elsewhere so we can display progress

			Case VIDEO_NO_MORE_DEATH_SAVES:
				DMD.jpDMD "", "", "d_E174_0", eNone, eNone, eNone, 10000, True, ""

			Case VIDEO_BLACKSMITH_SPELL:
				DMD.jpDMD DMD.jpCL("BLACKSMITH"), DMD.jpCL("Armor Class " & currentPlayer("AC") & " / 20"), "d_E177_0", eNone, eBlink, eNone, 8000, True, ""

			Case VIDEO_BLACKSMITH_WIZARD_READY:
				DMD.jpDMD DMD.jpCL("BLACKSMITH wizard"), DMD.jpCL("ready!"), "d_E178_0", eNone, eBlinkFast, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Shoot the"), DMD.jpCL("hole"), "d_E178_0", eNone, eNone, eNone, 2500, False, "" 'Sequence should loop; do not flush

			Case VIDEO_BONUSX_GLITCH_3:
				DMD.jpDMD "11111111111111111111", "11111111111111111111", "d_E179_0", eNone, eNone, eNone, 500, False, ""
				DMD.jpDMD "22222222222222222222", "22222222222222222222", "d_E179_1", eNone, eNone, eNone, 500, False, ""
				DMD.jpDMD "33333333333333333333", "33333333333333333333", "d_E179_2", eNone, eNone, eNone, 500, False, ""
				DMD.jpDMD "44444444444444444444", "44444444444444444444", "d_E179_3", eNone, eNone, eNone, 500, False, ""
				DMD.jpDMD "AAAAAAAAAAAAAAAAAAAA", "AAAAAAAAAAAAAAAAAAAA", "d_E179_0", eNone, eNone, eNone, 500, False, ""
				DMD.jpDMD "BBBBBBBBBBBBBBBBBBBB", "BBBBBBBBBBBBBBBBBBBB", "d_E179_1", eNone, eNone, eNone, 500, False, ""
				DMD.jpDMD "CCCCCCCCCCCCCCCCCCCC", "CCCCCCCCCCCCCCCCCCCC", "d_E179_2", eNone, eNone, eNone, 500, False, ""
				DMD.jpDMD "DDDDDDDDDDDDDDDDDDDD", "DDDDDDDDDDDDDDDDDDDD", "d_E179_3", eNone, eNone, eNone, 500, True, ""

			Case VIDEO_PLAYER_STATS:
				DMD.jpDMD DMD.jpFL(Chr(128) & "Health (HP)", Chr(130) & "AC"), DMD.jpCL(Chr(131) & "Drain Damage (-HP)"), "", eNone, eNone, eNone, 5000, True, ""

			Case VIDEO_GLITCH_WIZARD_READY:
				DMD.jpDMD DMD.jpCL("Glitch wizard"), DMD.jpCL("ready!"), "d_E181_0", eNone, eBlinkFast, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Shoot the"), DMD.jpCL("hole"), "d_E181_0", eNone, eNone, eNone, 2500, False, "" 'Sequence should loop; do not flush

			Case VIDEO_GLITCH_WIZARD_INTRO:
				For i = 0 to 11
					DMD.jpDMD "", "", "d_E182_0", eNone, eNone, eNone, 200, False, ""
					DMD.jpDMD "", "", "d_E182_1", eNone, eNone, eNone, 200, False, ""
					DMD.jpDMD "", "", "d_E182_2", eNone, eNone, eNone, 200, False, ""
					DMD.jpDMD "", "", "d_E182_3", eNone, eNone, eNone, 200, False, ""
					DMD.jpDMD "", "", "d_E182_4", eNone, eNone, eNone, 200, False, ""
				Next
				DMD.jpDMD DMD.jpCL("Glitch"), DMD.jpCL("Mini-wizard!"), "d_E182_5", eBlink, eBlink, eNone, 4000, True, ""
				DMD.jpDMD DMD.jpCL("Lanterns"), DMD.jpCL("award jackpot"), "d_E182_5", eBlinkFast, eBlinkFast, eNone, 7000, True, ""

			'Case VIDEO_GLITCH_WIZARD_SCREEN: (handled by jpScoreNow)
			'Case VIDEO_GLITCH_WIZARD_JACKPOT: (handled by the routine awarding the jackpot)

			Case VIDEO_GLITCH_JACKPOT_DOUBLED:
				DMD.jpDMD DMD.jpCL("Glitch jackpot"), DMD.jpCL("Doubled!"), "", eBlink, eBlink, eNone, 4000, True, ""

			Case VIDEO_GLITCH_WIZARD_TOTAL:
				DMD.jpDMD DMD.jpCL("Glitch Total"), DMD.jpCL(FormatScore(MODE_VALUES)), "", eNone, eBlink, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("Stand By;"), DMD.jpCL("draining balls"), "", eNone, eNone, eNone, 3000, False, "" 'Keep looping; once balls are drained, something will intercept this.

			Case VIDEO_GLITCH_OUTTRO:
				DMD.jpDMD ":~$", "", "", eNone, eNone, eNone, 1000, False, ""
				DMD.jpDMD ":~$ sudo", "", "", eNone, eNone, eNone, 1000, False, ""
				DMD.jpDMD ":~$ sudo reboot", "", "", eNone, eNone, eNone, 1500, True, ""

			'Case VIDEO_BLACKSMITH_LORE: Handled by DMD.blacksmithLore

			Case VIDEO_POWERUP_GLITCH:
				DMD.jpDMD "", "", "d_E188_0", eNone, eNone, eNone, 2000, True, ""

			Case VIDEO_POWERUP_NORMAL
				DMD.jpDMD "", "", "d_E189_0", eNone, eNone, eNone, 2000, True, ""

			Case VIDEO_BLACKSMITH_WIZARD_KILL_INTRO
				DMD.jpDMD "", "", "d_E192_0", eNone, eNone, eNone, 2500, False, ""
				DMD.jpDMD DMD.jpCL("Blacksmith"), DMD.jpCL("boss battle"), "d_E192_0", eBlink, eBlink, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL("Right Ramp"), DMD.jpCL("readies attack"), "d_E192_0", eScrollLeft, eScrollRight, eNone, 3200, False, ""
				DMD.jpDMD DMD.jpCL("Horseshoe deals"), DMD.jpCL("blacksmith damage"), "d_E192_0", eScrollLeft, eScrollRight, eNone, 3400, False, ""
				DMD.jpDMD DMD.jpCL("Drain a ball ="), DMD.jpCL("lose your AC"), "d_E192_0", eScrollLeft, eScrollRight, eNone, 4000, True, ""

			Case VIDEO_POWERUP_BLACKSMITH_KILL
				DMD.jpDMD "", "", "d_E194_0", eNone, eNone, eNone, 2000, True, ""

			'Case VIDEO_BLACKSMITH_WIZARD_KILL_DAMAGE handled by separate method

			'Case VIDEO_BLACKSMITH_WIZARD_KILL_LOSE_AC handled by separate method

			Case VIDEO_COMBO:
				DMD.jpDMD DMD.jpCL("Combo!"), DMD.jpCL(FormatScore((MODE_COUNTERS_A - 1) * Scoring.basic("combo"))), "d_E198_0", eNone, eBlink, eNone, 3000, True, ""

			'Case VIDEO_BLACKSMITH_WIZARD_KILL_FAIL handled by blacksmithLore

			Case VIDEO_BLACKSMITH_WIZARD_KILL_FAIL_RUN
				DMD.jpDMD DMD.jpCL("The Blacksmith"), DMD.jpCL("Fled!"), "", eBlinkFast, eBlinkFast, eNone, 5000, True, ""

			Case VIDEO_BLACKSMITH_WIZARD_KILL_FAIL_TOTAL
			Case VIDEO_BLACKSMITH_WIZARD_KILL_SUCCESS_TOTAL
			Case VIDEO_BLACKSMITH_WIZARD_SPARE_TOTAL
				DMD.jpDMD DMD.jpCL("Blacksmith Total:"), DMD.jpCL(FormatScore(MODE_VALUES)), "", eNone, eBlinkFast, eNone, 5000, True, ""

			Case VIDEO_DRAINING_BALLS:
				DMD.jpDMD DMD.jpCL("Stand By;"), DMD.jpCL("draining balls"), "", eNone, eNone, eNone, 3000, False, "" 'Keep looping; once balls are drained, something will intercept this.

			Case VIDEO_BLACKSMITH_WIZARD_KILL_SUCCESS:
			Case VIDEO_BLACKSMITH_WIZARD_KILL_SUCCESS_3:
				DMD.jpDMD DMD.jpCL("You killed"), DMD.jpCL("the Blacksmith!"), "", eBlink, eBlink, eNone, 3000, False, ""

			'Case VIDEO_BLACKSMITH_WIZARD_KILL_SUCCESS_2 handled by blacksmithLore

			Case VIDEO_BLACKSMITH_WIZARD_SPARE_INTRO:
				DMD.jpDMD DMD.jpCL("Raid the Blacksmith"), DMD.jpCL("Mini-Wizard"), "d_E1109_0", eBlink, eBlink, eNone, 2800, False, ""
				DMD.jpDMD DMD.jpCL("Shoot the"), DMD.jpCL("horseshoe"), "d_E1109_0", eScrollLeft, eScrollRight, eNone, 1100, False, ""
				DMD.jpDMD DMD.jpCL("to raise"), DMD.jpCL("the jackpot"), "d_E1109_0", eScrollLeft, eScrollRight, eNone, 1758, False, ""
				DMD.jpDMD DMD.jpCL("Center hole"), DMD.jpCL("collects jackpot"), "d_E1109_0", eScrollLeft, eScrollRight, eNone, 2990, True, ""

			'Case VIDEO_BLACKSMITH_WIZARD_SPARE_JACKPOT handled by blacksmithWizardSpareJackpot

			'Case VIDEO_BLACKSMITH_WIZARD_SPARE_SUPER_JACKPOT handled by blacksmithWizardSpareSuperJackpot

			Case VIDEO_LAST_HURRAH:
				DMD.jpDMD DMD.jpCL("Time is"), DMD.jpCL("running out!"), "", eBlinkFast, eBlinkFast, eNone, 2500, True, ""

			Case VIDEO_POWERUP_BLACKSMITH_SPARE
				DMD.jpDMD "", "", "d_E1114_0", eNone, eNone, eNone, 2000, True, ""

			Case VIDEO_POISONED_OUTLANE:
				DMD.jpDMD DMD.jpCL("Poisoned outlane!"), DMD.jpCL(Chr(131) & currentPlayer("drainDamage")), "d_E1115_0", eNone, eNone, eNone, 3000, True, ""

			Case VIDEO_FINAL_SCORE:
				DMD.jpDMD DMD.jpCL("Final Score"), DMD.jpCL(FormatScore(currentPlayer("score"))), "", eNone, eBlinkFast, eNone, 5000, True, ""

			'Case VIDEO_HIGH_SCORE_FINAL: (handled in a separate routine)

			Case VIDEO_TILT:
				DMD.jpDMD DMD.jpCL("Too much nudge!"), DMD.jpCL(Chr(131) & currentPlayer("drainDamage")), "d_E1115_0", eNone, eNone, eNone, 3000, True, ""

			Case Else:
				If DMD.jpLocked = False Then DMD.jpScoreNow
		End Select
		
		curVideo = EventNum
	End Sub
	
	Public Sub triggerOverlay(EventNum) 	'Trigger a PUP overlay event (D* event 900-999). Should be called BEFORE setting any labels that are auto-reset in this routine.
		If (usePUP = False Or PUPStatus = False) Then
			' TODO
		Else
			' specialized resets
			If Not curOverlay = EventNum Then
				Select Case curOverlay
					Case OVERLAY_SCORE_COMPETITIVE, OVERLAY_SCORE_COOPERATIVE, OVERLAY_SCORBIT_CLAIM, OVERLAY_SCORBIT_PAIR
						If Not EventNum = OVERLAY_SCORE_COMPETITIVE And Not EventNum = OVERLAY_SCORE_COOPERATIVE And Not EventNum = OVERLAY_SCORBIT_CLAIM And Not EventNum = OVERLAY_SCORBIT_PAIR Then
							PUP.LabelSet pDMD,"p1label","",0,""
							PUP.LabelSet pDMD,"p2label","",0,""
							PUP.LabelSet pDMD,"p3label","",0,""
							PUP.LabelSet pDMD,"p4label","",0,""
							PUP.LabelSet pDMD,"p1hp","",0,""
							PUP.LabelSet pDMD,"p2hp","",0,""
							PUP.LabelSet pDMD,"p3hp","",0,""
							PUP.LabelSet pDMD,"p4hp","",0,""
							PUP.LabelSet pDMD,"p1ac","",0,""
							PUP.LabelSet pDMD,"p2ac","",0,""
							PUP.LabelSet pDMD,"p3ac","",0,""
							PUP.LabelSet pDMD,"p4ac","",0,""
							PUP.LabelSet pDMD,"p1dd","",0,""
							PUP.LabelSet pDMD,"p2dd","",0,""
							PUP.LabelSet pDMD,"p3dd","",0,""
							PUP.LabelSet pDMD,"p4dd","",0,""
							PUP.LabelSet pDMD,"p1score","",0,""
							PUP.LabelSet pDMD,"p2score","",0,""
							PUP.LabelSet pDMD,"p3score","",0,""
							PUP.LabelSet pDMD,"p4score","",0,""
							PUP.LabelSet pDMD,"coopscore","",0,""
							PUP.LabelSet pDMD,"p1scorbit","",0,""
							PUP.LabelSet pDMD,"p2scorbit","",0,""
							PUP.LabelSet pDMD,"p3scorbit","",0,""
							PUP.LabelSet pDMD,"p4scorbit","",0,""
							PUP.LabelSet pDMD,"gameplay","",0,""
						End If
				End Select
			End If

			PUP.B2SData "D" & EventNum,1  'send event to Pup-Pack
		End If
		
		curOverlay = EventNum
	End Sub
	
	' Trigger a DOF event in PUP (E* event)
	Public Sub triggerEvent(EventNum)
		If (usePUP = False Or PUPStatus = False) Then Exit Sub
		PUP.B2SData "E" & EventNum,1  'send event to Pup-Pack
	End Sub

	' Trigger a switch event in PUP (W* event)
	Public Sub triggerSwitch(switchNum, enabled)
		If (usePUP = False Or PUPStatus = False) Then Exit Sub
		If enabled = true Then
			PUP.B2SData "W" & switchNum,1
		Else
			PUP.B2SData "W" & switchNum,0
		End If
	End Sub

	' Trigger a solenoid event in PUP (W* event)
	Public Sub triggerSolenoid(solenoidNum, enabled)
		If (usePUP = False Or PUPStatus = False) Then Exit Sub
		If enabled = true Then
			PUP.B2SData "S" & solenoidNum,1
		Else
			PUP.B2SData "S" & solenoidNum,0
		End If
	End Sub
	
	' Set the page number visible
	Public Sub setPage(pScreen, pagenum)
		If (usePUP = False Or PUPStatus = False) Then Exit Sub
		PUP.LabelShowPage pScreen,pagenum,0,""
		pPageCur = pagenum
	End Sub
	
	' Hide a label
	Public Sub labelHide(pScreen, labName)
		If (usePUP = False Or PUPStatus = False) Then Exit Sub
		PUP.LabelSet pScreen,labName,"",0,""
	End Sub

	Public Sub centerText(textTop, textBottom) 'Show text on the center_top and center_bottom PUP labels
		If usePUP = True And PUPStatus = True Then
			PUP.LabelSet pDMD,"center_top",textTop,1,""
			PUP.LabelSet pDMD,"center_bottom",textBottom,1,""
		End If
	End Sub

	Public Sub centerTextRemove() 'Clear the center_top and center_bottom PUP labels
		If usePUP = True And PUPStatus = True Then
			PUP.LabelSet pDMD,"center_top","",0,""
			PUP.LabelSet pDMD,"center_bottom","",0,""
		End If
	End Sub

	'*************************
	' SCORBIT
	'*************************

	Public Sub ScorbitCheckPairing()
		If (usePUP = False Or PUPStatus = False Or Not ScorbitActive = 1) Then Exit Sub

		If (Scorbit.bNeedsPairing And CURRENT_MODE = MODE_ATTRACT) Then 
			If ScorbitQRSmall = 1 Then
				triggerOverlay OVERLAY_SCORBIT_PAIR_SMALL
				PUP.LabelSet pDMD, "ScorbitQR","PuPOverlays\\QRcode.png",1,"{'mt':2,'width':13.28125, 'height':23.611111111,'xalign':0,'yalign':0,'ypos':75.555555555,'xpos':86.40625}"
			Else
				triggerOverlay OVERLAY_SCORBIT_PAIR
				PUP.LabelSet pDMD, "ScorbitQR","PuPOverlays\\QRcode.png",1,"{'mt':2,'width':40.364583333, 'height':71.759259259,'xalign':0,'yalign':0,'ypos':1.666666666,'xpos':0.9375}"
			End If
		Else
			PUP.LabelSet pDMD, "ScorbitQR","",0,""
			If CURRENT_MODE = MODE_ATTRACT Then 
				DMD.sShowScores
			End If
		End If
	End Sub

	Public Sub ScorbitPlayerClaimed(PlayerNum, PlayerName)
		If (usePUP = False Or PUPStatus = False Or Not ScorbitActive = 1) Then Exit Sub

		'Reset score labels to display name and scorbit icons
		sUpdateScores True
	End Sub

	Public Sub ScorbitClaimQR(show)
		If (usePUP = False Or PUPStatus = False Or Not ScorbitActive = 1) Then Exit Sub

		If show = True And dataGame.data.Item("gameMode") = 0 Then 'No player claiming for coop mode
			If ScorbitQRSmall = 1 Then
				triggerOverlay OVERLAY_SCORBIT_CLAIM_SMALL
				PUP.LabelSet pDMD, "ScorbitQR","PuPOverlays\\QRclaim.png",1,"{'mt':2,'width':13.28125, 'height':23.611111111,'xalign':0,'yalign':0,'ypos':75.555555555,'xpos':86.40625}"
			Else
				triggerOverlay OVERLAY_SCORBIT_CLAIM
				PUP.LabelSet pDMD, "ScorbitQR","PuPOverlays\\QRclaim.png",1,"{'mt':2,'width':40.364583333, 'height':71.759259259,'xalign':0,'yalign':0,'ypos':1.666666666,'xpos':0.9375}"
			End If
		Else
			DMD.sShowScores
			PUP.LabelSet pDMD, "ScorbitQR","",0,""
		End If
	End Sub
	
	'*************************
	' SELF TEST SEQUENCE
	'*************************
	
	' Queue up the table initialization / boot sequence
	Public Sub sInitialization()
		' TODO: Elegant solution for PUP load delay?
		queue.Add "Initialization_0","DMD.triggerVideo VIDEO_RGB_TEST",10,0,2000,2000,0,False
		
		' Self-test sequence
		If DEBUG_SKIP_INITIALIZATION = False Then
			' Flippers / lights / flashers / shaker motor impacts
			queue.Add "Initialization_1","DMD.sInitializationStep 1",10,0,0,2466,0,False
			queue.Add "Initialization_2","DMD.sInitializationStep 2",10,0,0,1333,0,False
			queue.Add "Initialization_3","DMD.sInitializationStep 3",10,0,0,1333,0,False
			queue.Add "Initialization_4","DMD.sInitializationStep 4",10,0,0,1333,0,False
			queue.Add "Initialization_5","DMD.sInitializationStep 5",10,0,0,1333,0,False
			queue.Add "Initialization_6","DMD.sInitializationStep 6",10,0,0,1333,0,False
			queue.Add "Initialization_7","DMD.sInitializationStep 7",10,0,0,1333,0,False
			queue.Add "Initialization_8","DMD.sInitializationStep 8",10,0,0,666,0,False
			queue.Add "Initialization_9","DMD.sInitializationStep 9",10,0,0,666,0,False
			queue.Add "Initialization_10","DMD.sInitializationStep 10",10,0,0,333,0,False
			queue.Add "Initialization_11","DMD.sInitializationStep 11",10,0,0,333,0,False
			queue.Add "Initialization_12","DMD.sInitializationStep 12",10,0,0,333,0,False
			queue.Add "Initialization_13","DMD.sInitializationStep 13",10,0,0,333,0,False
			
			' Bumpers
			queue.Add "Initialization_14","DMD.sInitializationStep 14",10,0,0,330,0,False
			queue.Add "Initialization_15","DMD.sInitializationStep 15",10,0,0,330,0,False
			queue.Add "Initialization_16","DMD.sInitializationStep 16",10,0,0,330,0,False
			queue.Add "Initialization_17","DMD.sInitializationStep 17",10,0,0,330,0,False
			queue.Add "Initialization_18","DMD.sInitializationStep 18",10,0,0,1330,0,False
			
			' Slingshots
			queue.Add "Initialization_19","DMD.sInitializationStep 19",10,0,0,330,0,False
			queue.Add "Initialization_20","DMD.sInitializationStep 20",10,0,0,330,0,False
			queue.Add "Initialization_21","DMD.sInitializationStep 21",10,0,0,330,0,False
			queue.Add "Initialization_22","DMD.sInitializationStep 22",10,0,0,330,0,False
			queue.Add "Initialization_23","DMD.sInitializationStep 23",10,0,0,1330,0,False
			
			' Kickers / Drops
			queue.Add "Initialization_24","DMD.sInitializationStep 24",10,0,0,330,0,False
			queue.Add "Initialization_25","DMD.sInitializationStep 25",10,0,0,330,0,False
			queue.Add "Initialization_26","DMD.sInitializationStep 26",10,0,0,330,0,False
			queue.Add "Initialization_27","DMD.sInitializationStep 27",10,0,0,330,0,False
			queue.Add "Initialization_28","DMD.sInitializationStep 28",10,0,0,1330,0,False
			
			' Diverter
			queue.Add "Initialization_29","DMD.sInitializationStep 29",10,0,0,330,0,False
			queue.Add "Initialization_30","DMD.sInitializationStep 30",10,0,0,330,0,False
			queue.Add "Initialization_31","DMD.sInitializationStep 31",10,0,0,330,0,False
			queue.Add "Initialization_32","DMD.sInitializationStep 32",10,0,0,330,0,False
			queue.Add "Initialization_33","DMD.sInitializationStep 33",10,0,0,1330,0,False
			
			' Empty (lights)
			queue.Add "Initialization_34","DMD.sInitializationStep 34",10,0,0,330,0,False
			queue.Add "Initialization_35","DMD.sInitializationStep 35",10,0,0,330,0,False
			queue.Add "Initialization_36","DMD.sInitializationStep 36",10,0,0,330,0,False
			queue.Add "Initialization_37","DMD.sInitializationStep 37",10,0,0,330,0,False
			queue.Add "Initialization_38","DMD.sInitializationStep 38",10,0,0,1330,0,False
			
			' Empty (lights)
			queue.Add "Initialization_39","DMD.sInitializationStep 39",10,0,0,330,0,False
			queue.Add "Initialization_40","DMD.sInitializationStep 40",10,0,0,330,0,False
			queue.Add "Initialization_41","DMD.sInitializationStep 41",10,0,0,330,0,False
			queue.Add "Initialization_42","DMD.sInitializationStep 42",10,0,0,330,0,False
			queue.Add "Initialization_43","DMD.sInitializationStep 43",10,0,0,1330,0,False
			
			' Done
			queue.Add "Initialization_44","DMD.sTestsPassed",10,0,0,5333,0,False
		End If

		queue.Add "Initialization_scorbit","DMD.sInitializationStep 0",10,0,0,100,0,False

		If Not dataGame.data.Item("ball") = -1 And dataGame.data.Item("tournamentMode") = 0 Then 'Non-tournament game was in progress; ask if player wants to resume (tournament games cannot be resumed)
			queue.Add "Initialization_45","startMode MODE_RESUME_GAME_PROMPT",10,0,0,0,0,False
		Else
			' Start the attract mode
			queue.Add "Initialization_45","startMode MODE_ATTRACT",10,0,0,0,0,False
		End If
	End Sub
	
	' Execute a step in the self test / initialization sequence
	Public Sub sInitializationStep(sStep)
		Select Case (sStep)
			Case 0:
				If usePUP = True And PUPStatus = True Then
					If ScorbitActive = 1 Then 'Putting this on the same if condition as DoInit will still call DoInit even if this is not true. Got to love VBScript ;)
						'Init Scorbit (machine ID is 4113 for production, 4114 for staging, for Saving Wallden)
						If Scorbit.DoInit(4113, "PuPOverlays", TABLE_VERSION, "wallden-vpin") then
							tmrScorbit.Interval=2000
							tmrScorbit.UserValue = 0
							tmrScorbit.Enabled=True
						End if
					End If
				End If
			Case 1:
				DMD.sInitialization_romTest
			Case 2:
				DMD.triggerVideo VIDEO_SELF_TEST
				Music.PlayTrack "music_startup", Null, Null, Null, Null, 0, 0, 0, 0
				DMD.sInitialization_impact Array(dred, red)
				DMD.sInitialization_pegDrops 0, 4, True
				Flash1 True
			Case 3, 5, 7, 9, 11, 13:
				DMD.sInitialization_impactOff
			Case 4:
				DMD.sInitialization_impact Array(dorange, orange)
				DMD.sInitialization_pegDrops 5, 9, True
				Flash3 True
			Case 6:
				DMD.sInitialization_impact Array(dyellow, yellow)
				DMD.sInitialization_pegDrops 10, 14, True
				Flash4 True
			Case 8:
				DMD.sInitialization_impact Array(dgreen, green)
				DMD.sInitialization_pegDrops 15, 19, True
				Flash5 True
			Case 10:
				DMD.sInitialization_impact Array(dblue, blue)
				DMD.sInitialization_pegDrops 20, 24, True
				Flash2 True
			Case 12:
				DMD.sInitialization_impact Array(dpurple, purple)
				DMD.sInitialization_pegDrops 25, 29, True
				Flash6 True
			Case 14:
				LightsOn GI, Array(dbase, base), 100
				DMD.sInitialization_rainbow
				SolPegDrop 0, False
				DMD.sInitialization_lights1a
			Case 15:
				SolBumper1 True
				SolPegDrop 1, False
				DMD.sInitialization_lights1b
			Case 16:
				SolBumper2 True
				SolPegDrop 2, False
				DMD.sInitialization_lights1a
			Case 17:
				SolBumper3 True
				SolPegDrop 3, False
				DMD.sInitialization_lights1b
			Case 18:
				SolPegDrop 4, False
				StartLightSequence LightSeqAttract, Null, 25, SeqRandom, 25, 1
			Case 19:
				SolLSling True
				SolPegDrop 5, False
				StopLightSequence LightSeqAttract
				DMD.sInitialization_lights2a
				Flash7 True
			Case 20:
				SolPegDrop 6, False
				DMD.sInitialization_lights2b
			Case 21:
				SolRSling True
				SolPegDrop 7, False
				DMD.sInitialization_lights2a
				Flash7 False
				Flash8 True
			Case 22:
				SolPegDrop 8, False
				DMD.sInitialization_lights2b
			Case 23:
				SolLSling False
				SolRSling False
				SolPegDrop 9, False
				StartLightSequence LightSeqAttract, Null, 25, SeqRandom, 25, 1
				Flash8 False
			Case 24:
				SolScoop True
				SolPegDrop 10, False
				StopLightSequence LightSeqAttract
				DMD.sInitialization_lights3a
			Case 25:
				SolVUpKicker True
				SolPegDrop 11, False
				DMD.sInitialization_lights3b
			Case 26:
				SolHPBank True
				SolPegDrop 12, False
				DMD.sInitialization_lights3a
			Case 27:
				SolPegDrop 13, False
				DMD.sInitialization_lights3b
			Case 28:
				SolPegDrop 14, False
				StartLightSequence LightSeqAttract, Null, 25, SeqRandom, 25, 1
			Case 29:
				SolVUKDiverter True
				SolPegDrop 15, False
				StopLightSequence LightSeqAttract
				StartLightSequence LightSeqAttract, Null, 14, SeqScrewLeftOn, 15, 1
			Case 30:
				SolPegDrop 16, False
			Case 31:
				SolVUKDiverter False
				SolPegDrop 17, False
			Case 32:
				SolPegDrop 18, False
			Case 33:
				SolPegDrop 19, False
			Case 34:
				SolPegDrop 20, False
				StopLightSequence LightSeqAttract
				DMD.sInitialization_lights1a
			Case 35:
				SolPegDrop 21, False
				DMD.sInitialization_lights1b
			Case 36:
				SolPegDrop 22, False
				DMD.sInitialization_lights1a
			Case 37:
				SolPegDrop 23, False
				DMD.sInitialization_lights1b
			Case 38:
				SolPegDrop 24, False
				StartLightSequence LightSeqAttract, Null, 25, SeqRandom, 25, 1
			Case 39:
				SolPegDrop 25, False
				StopLightSequence LightSeqAttract
				DMD.sInitialization_lights2a
			Case 40:
				SolPegDrop 26, False
				DMD.sInitialization_lights2b
			Case 41:
				SolPegDrop 27, False
				DMD.sInitialization_lights2a
			Case 42:
				SolPegDrop 28, False
				DMD.sInitialization_lights2b
			Case 43:
				SolPegDrop 29, False
		End Select
	End Sub
	
	' Controller for the peg drop test
	Public Sub sInitialization_pegDrops(istart, iend, enabled)
		Dim index
		For index = istart To iend
			SolPegDrop index, enabled
		Next
	End Sub
	
	' Controller for flashing all flashers
	Public Sub sInitialization_Flashers(Enabled)
		Flash1 Enabled
		Flash2 Enabled
		Flash3 Enabled
		Flash4 Enabled
		Flash5 Enabled
		Flash6 Enabled
		Flash7 Enabled
		Flash8 Enabled
	End Sub
	
	' Randomized flashing / test tone / logo
	Public Sub sInitialization_romTest()
		triggerVideo VIDEO_VPW_LOGO
		
		LightsOn GI, Array(dwhite, white), 100
		LightsOn aLights, Array(dwhite, white), 100
		StartLightSequence LightSeqAttract, Array(dwhite,white), 25, SeqRandom, 25, 1

		PlaySound "sfx_testtone", 1, VOLUME_SFX
	End Sub
	
	' Turn all table lights on the specified color, activate flipper solenoids, and turn on the shaker motor
	Public Sub sInitialization_impact(color)
		StopLightSequence LightSeqAttract
		LightsOn coloredLights, color, 100
		LightsOn bumperLights, Null, 100
		LightsOn GI, color, 100
		Dim aLight
		For Each aLight In segmentLights
			aLight.state = 1
		Next
		SolLFlipper True
		SolRFlipper True
		SolUFlipper True
		
		ShakerMotorTimer.Enabled = 1
	End Sub
	
	' Turn lights off, release the flippers, and stop the shaker motor
	Public Sub sInitialization_impactOff()
		StopLightSequence LightSeqAttract
		LightsOff GI
		LightsOff aLights
		Dim aLight
		For Each aLight In segmentLights
			aLight.state = 0
		Next
		SolLFlipper False
		SolRFlipper False
		SolUFlipper False
		
		ShakerMotorTimer.Enabled = 0
		sInitialization_Flashers False
	End Sub
	
	' Begin rainbow
	Public Sub sInitialization_rainbow()
		StartRainbow coloredLights
	End Sub
	
	' Rhythmic light sequence - Top half of table
	Public Sub sInitialization_lights1a()
		Dim aLight
		For Each aLight In aLights
			If aLight.X <= (Table1.Width / 2) Then
				lampC.LightOn(aLight)
			Else
				lampC.LightOff(aLight)
			End If
		Next
	End Sub
	
	' Rhythmic light sequence - Bottom half of table
	Public Sub sInitialization_lights1b()
		Dim aLight
		For Each aLight In aLights
			If aLight.X >= (Table1.Width / 2) Then
				lampC.LightOn(aLight)
			Else
				lampC.LightOff(aLight)
			End If
		Next
	End Sub
	
	' Rhythmic light sequence - Left half of table
	Public Sub sInitialization_lights2a()
		Dim aLight
		For Each aLight In aLights
			If aLight.Y <= (Table1.Height / 2) Then
				lampC.LightOn(aLight)
			Else
				lampC.LightOff(aLight)
			End If
		Next
	End Sub
	
	' Rhythmic light sequence - Right half of table
	Public Sub sInitialization_lights2b()
		Dim aLight
		For Each aLight In aLights
			If aLight.Y >= (Table1.Height / 2) Then
				lampC.LightOn(aLight)
			Else
				lampC.LightOff(aLight)
			End If
		Next
	End Sub
	
	' Rhythmic light sequence - Top-left half of table
	Public Sub sInitialization_lights3a()
		Dim aLight
		For Each aLight In aLights
			If (aLight.X + aLight.Y) <= ((Table1.Height + Table1.Width) / 2) Then
				lampC.LightOn(aLight)
			Else
				lampC.LightOff(aLight)
			End If
		Next
	End Sub
	
	' Rhythmic light sequence - Bottom-right half of table
	Public Sub sInitialization_lights3b()
		Dim aLight
		For Each aLight In aLights
			If (aLight.X + aLight.Y) >= ((Table1.Height + Table1.Width) / 2) Then
				lampC.LightOn(aLight)
			Else
				lampC.LightOff(aLight)
			End If
		Next
	End Sub
	
	' Tests passed; turn everything green
	Public Sub sTestsPassed()
		StopRainbow coloredLights
		StopLightSequence LightSeqAttract
		LightsOn coloredLights, Array(dgreen, green), 100
		LightsOn GI, Array(dgreen, green), 100
		triggerVideo VIDEO_TESTS_PASSED
	End Sub
	
	' Start / initialize attract sequence
	Public Sub sBeginAttract()
		triggerVideo VIDEO_BLACK_BG
		sShowScores
		sAttractSequence
		If (usePUP = True And PUPStatus = True) Then gameStatus "GAME OVER"
	End Sub
	
	' Queue up the attract sequence
	Public Sub sAttractSequence()
		StartAttractLightSequence
		StartRainbow coloredLights
		
		' Intro / High Scores / Stats
		queue.Add "attracta_1","DMD.sAttractCredits",2,0,0,5000,0,False
		queue.Add "attracta_2","DMD.sAttractVPinPresents",2,0,0,5000,0,False
		queue.Add "attracta_3","DMD.sAttractSavingWallden",2,0,0,5000,0,False
		queue.Add "attracta_4","DMD.sAttractPreviousScores",2,0,0,2500 + (2500 * dataGame.data.Item("numPlayers")),0,False
		If TOURNAMENT_MODE = 0 Then
			queue.Add "attracta_5","StopRainbow coloredLights : DMD.sAttractHighScoresZen",2,0,0,10000,0,False
			queue.Add "attracta_6","DMD.sAttractHighScoresEasy",2,0,0,10000,0,False
			queue.Add "attracta_7","DMD.sAttractHighScoresNormal",2,0,0,10000,0,False
			queue.Add "attracta_8","DMD.sAttractHighScoresHard",2,0,0,10000,0,False
			queue.Add "attracta_9","DMD.sAttractHighScoresImpossible",2,0,0,10000,0,False
		Else
			queue.Add "attracta_5","StopRainbow coloredLights : DMD.sAttractHighScoresTournament",2,0,0,10000,0,False
		End If
		
		' Tutorial
		queue.Add "attractb_1","DMD.sAttractCredits : DMD.sAttractCallout",2,0,0,5000,0,False
		queue.Add "attractb_2","DMD.sAttractTutorialIntro",2,0,0,5000,0,False
		queue.Add "attractb_3","DMD.sAttractTutorialCastle",2,0,0,10000,0,False
		queue.Add "attractb_4","DMD.sAttractTutorialChaser",2,0,0,10000,0,False
		queue.Add "attractb_5","DMD.sAttractTutorialDragon",2,0,0,10000,0,False
		queue.Add "attractb_6","DMD.sAttractTutorialViking",2,0,0,10000,0,False
		queue.Add "attractb_7","DMD.sAttractTutorialEscape",2,0,0,10000,0,False
		queue.Add "attractb_8","DMD.sAttractTutorialSniper",2,0,0,10000,0,False
		queue.Add "attractb_9","DMD.sAttractTutorialBlacksmith",2,0,0,10000,0,False
		queue.Add "attractb_10","DMD.sAttractTutorialFinalBoss",2,0,0,10000,0,False
		queue.Add "attractb_11","DMD.sAttractTutorialFinalWizard",2,0,0,10000,0,False
		
		' Contributors / Special Thanks (Don't forget to also update sAttractSequence_gameOverFailed and gameOverSuccess)
		queue.Add "attractc_1","StartAttractLightSequence : StartRainbow coloredLights : DMD.sAttractCredits : DMD.sAttractCallout",2,0,0,5000,0,False
		queue.Add "attractc_scores","DMD.sAttractPreviousScores",2,0,0,2500 + (2500 * dataGame.data.Item("numPlayers")),0,False
		queue.Add "attractc_2","DMD.triggerVideo VIDEO_CONTRIBUTORS_INTRO",2,0,0,5000,0,False
		queue.Add "attractc_3","DMD.triggerVideo VIDEO_CONTRIBUTORS_ARELYEL",2,0,0,8000,0,False
		queue.Add "attractc_4","DMD.triggerVideo VIDEO_CONTRIBUTORS_DONKEYKLONK",2,0,0,8000,0,False
		queue.Add "attractc_5","DMD.triggerVideo VIDEO_CONTRIBUTORS_FLUX",2,0,0,8000,0,False
		queue.Add "attractc_6","DMD.triggerVideo VIDEO_CONTRIBUTORS_GRAHAM_GROWAT",2,0,0,8000,0,False
		queue.Add "attractc_7","DMD.triggerVideo VIDEO_CONTRIBUTORS_HARMONY",2,0,0,8000,0,False
		queue.Add "attractc_8","DMD.triggerVideo VIDEO_CONTRIBUTORS_SMAUG",2,0,0,8000,0,False
		queue.Add "attractc_end","DMD.triggerVideo VIDEO_CONTRIBUTORS_OTHER",2,0,0,15000,0,False
		
		' Do it all again!
		queue.Add "attractd_1","DMD.sAttractSequence : DMD.sAttractCallout",2,0,0,0,0,False
	End Sub

	' Queue up the attract sequence after the game is over and the mission was failed
	Public Sub sAttractSequence_gameOverFailed()
		'Loop through the credits at the end of a game first
		queue.Add "attract_failed_1","DMD.sAttractSequence_gameOverFailed_video",2,0,0,15000,0,False
		queue.Add "attract_failed_2","StartAttractLightSequence : StartRainbow coloredLights : startMode MODE_ATTRACT : DMD.sAttractPreviousScores",2,0,0,2500 + (2500 * dataGame.data.Item("numPlayers")),0,False
		queue.Add "attract_failed_3","DMD.triggerVideo VIDEO_CONTRIBUTORS_INTRO",2,0,0,5000,0,False
		queue.Add "attract_failed_4","DMD.triggerVideo VIDEO_CONTRIBUTORS_ARELYEL",2,0,0,8000,0,False
		queue.Add "attract_failed_5","DMD.triggerVideo VIDEO_CONTRIBUTORS_DONKEYKLONK",2,0,0,8000,0,False
		queue.Add "attract_failed_6","DMD.triggerVideo VIDEO_CONTRIBUTORS_FLUX",2,0,0,8000,0,False
		queue.Add "attract_failed_7","DMD.triggerVideo VIDEO_CONTRIBUTORS_GRAHAM_GROWAT",2,0,0,8000,0,False
		queue.Add "attract_failed_8","DMD.triggerVideo VIDEO_CONTRIBUTORS_HARMONY",2,0,0,8000,0,False
		queue.Add "attract_failed_9","DMD.triggerVideo VIDEO_CONTRIBUTORS_SMAUG",2,0,0,8000,0,False
		queue.Add "attract_failed_end","DMD.triggerVideo VIDEO_CONTRIBUTORS_OTHER",2,0,0,15000,0,False

		'Start the main attract sequence
		queue.Add "attract_failed_restart","DMD.sAttractSequence : DMD.sAttractCallout",2,0,0,0,0,False
	End Sub

	Sub sAttractSequence_gameOverFailed_video()
		'DMD.triggerVideo VIDEO_CLEAR
		DMD.triggerVideo VIDEO_GAME_OVER_FAILED
		Music.PlayTrack "music_game_over_failed", 1, Null, Null, Null, 0, 0, 0, 0
		PlayCallout "callout_nar_game_over_mission_failed", 1, VOLUME_CALLOUTS
	End Sub
	
	' Play a random callout
	Public Sub sAttractCallout
		PlayCallout "callout_attract_" & Int(Rnd*11), 1, VOLUME_CALLOUTS
	End Sub
	
	' Update credits display (FREE PLAY)
	Public Sub sAttractCredits()
		triggerVideo VIDEO_FREE_PLAY
		If (usePUP = False Or PUPStatus = False) Then
			' TODO
		Else
			If curOverlay = OVERLAY_SCORE_COMPETITIVE Or curOverlay = OVERLAY_SCORE_COOPERATIVE Then sCreditsGameplay
		End If
	End Sub
	
	' Show credits (FREE PLAY) on the gameplay line
	Public Sub sCreditsGameplay()
		If (usePUP = False Or PUPStatus = False) Then
			' TODO
		Else
			If CURRENT_MODE < 1000 Then gameStatus "FREE PLAY / PRESS START"
			If CURRENT_MODE = MODE_SKILLSHOT And GAME_OFFICIALLY_STARTED = False Then gameStatus "PRESS START Now To ADD PLAYERS"
		End If
	End Sub
	
	' VPin Presents slide
	Public Sub sAttractVPinPresents()
		triggerVideo VIDEO_VPW_PRESENTS
	End Sub
	
	' Saving Wallden slide
	Public Sub sAttractSavingWallden()
		triggerVideo VIDEO_SAVING_WALLDEN
	End Sub

	' Previous Scores slide
	Public Sub sAttractPreviousScores()
		triggerVideo VIDEO_PREVIOUS_SCORES
	End Sub
	
	' Show high scores for zen difficulty
	Public Sub sAttractHighScoresZen()
		LightsOn coloredLights, Array(dpurple, purple), 100
		triggerVideo VIDEO_HIGH_SCORES_ZEN
	End Sub
	
	' Show high scores for easy difficulty
	Public Sub sAttractHighScoresEasy()
		LightsOn coloredLights, Array(dgreen, green), 100
		triggerVideo VIDEO_HIGH_SCORES_EASY
	End Sub
	
	' Show high scores for normal difficulty
	Public Sub sAttractHighScoresNormal()
		LightsOn coloredLights, Array(dyellow, yellow), 100
		triggerVideo VIDEO_HIGH_SCORES_NORMAL
	End Sub
	
	' Show high scores for hard difficulty
	Public Sub sAttractHighScoresHard()
		LightsOn coloredLights, Array(dorange, orange), 100
		triggerVideo VIDEO_HIGH_SCORES_HARD
	End Sub
	
	' Show high scores for impossible difficulty
	Public Sub sAttractHighScoresImpossible()
		LightsOn coloredLights, Array(dred, red), 100
		triggerVideo VIDEO_HIGH_SCORES_IMPOSSIBLE
	End Sub
	
	' Show high scores for tournaments
	Public Sub sAttractHighScoresTournament()
		LightsOn coloredLights, Array(ddarkblue, darkblue), 100
		triggerVideo VIDEO_HIGH_SCORES_TOURNAMENT
	End Sub
	
	' Intro to tutorial slides; turn all lights on base color
	Public Sub sAttractTutorialIntro()
		triggerVideo VIDEO_TUTORIAL_INTRO
		StopLightSequence LightSeqAttract
		StopLightSequence LightSeqAttractB
		StopRainbow coloredLights
		LightsOn GI, Array(dbase, base), 100
		LightsOn coloredLights, Array(dbase, base), 100
		LightsOn bumperLights, Null, 100
	End Sub
	
	' Turn all tutorial lights off. This should also be called whenever we are interrupting the attract mode!
	Public Sub resetTutorialSeq()
		If Not IsNull(activeSeqAttract) Then
			lampC.RemoveAllLightSeq activeSeqAttract
			activeSeqAttract = Null
		End If
		
		LightsOff aLights
		
		StopLightSequence LightSeqAttract
		StopLightSequence LightSeqAttractB
		StopRainbow coloredLights
	End Sub
	
	' TRAIN / CASTLE tutorial
	Public Sub sAttractTutorialCastle()
		resetTutorialSeq
		triggerVideo VIDEO_TUTORIAL_CASTLE
		
		LightsOn GI, Array(dyellow, yellow), 100
		LightsBlink tutorialLights1B, Array(dyellow, yellow), 500, 100
		
		Dim lampASeq
		Set lampASeq = New LCSeq
		lampASeq.Name = "sAttractTutorialCastle_train"
		lampASeq.Sequence = AppendArray(Array(), Array(Array("StandupsTL|100","StandupsRL|20","StandupsAL|20","StandupsIL|20","StandupsNL|20")))
		lampASeq.Sequence = AppendArray(lampASeq.Sequence, Array(Array("StandupsTL|20","StandupsRL|100","StandupsAL|20","StandupsIL|20","StandupsNL|20")))
		lampASeq.Sequence = AppendArray(lampASeq.Sequence, Array(Array("StandupsTL|20","StandupsRL|20","StandupsAL|100","StandupsIL|20","StandupsNL|20")))
		lampASeq.Sequence = AppendArray(lampASeq.Sequence, Array(Array("StandupsTL|20","StandupsRL|20","StandupsAL|20","StandupsIL|100","StandupsNL|20")))
		lampASeq.Sequence = AppendArray(lampASeq.Sequence, Array(Array("StandupsTL|20","StandupsRL|20","StandupsAL|20","StandupsIL|20","StandupsNL|100")))
		lampASeq.UpdateInterval = 200
		lampASeq.Color = Array(dyellow, yellow)
		lampASeq.Repeat = True
		
		lampC.AddLightSeq "attract", lampASeq
		activeSeqAttract = "attract"
	End Sub
	
	' CHASER tutorial
	Public Sub sAttractTutorialChaser()
		resetTutorialSeq
		triggerVideo VIDEO_TUTORIAL_CHASER
		
		LightsOn GI, Array(dgreen, green), 100
		LightsBlink tutorialLights2, Array(dgreen, green), 500, 100
	End Sub
	
	' DRAGON tutorial
	Public Sub sAttractTutorialDragon()
		resetTutorialSeq
		triggerVideo VIDEO_TUTORIAL_DRAGON
		
		LightsOn GI, Array(ddarkblue, darkblue), 100
		LightsBlink tutorialLights3, Array(ddarkblue, darkblue), 500, 100
		LightsBlink tutorialLights3B, Array(ddarkblue, darkblue), 500, 100
		LightsBlink tutorialLights3C, Null, 500, 100
	End Sub
	
	' VIKING tutorial
	Public Sub sAttractTutorialViking()
		resetTutorialSeq
		triggerVideo VIDEO_TUTORIAL_VIKING
		
		LightsOn GI, Array(dblue, blue), 100
		LightsBlink tutorialLights4, Array(dblue, blue), 500, 100
	End Sub
	
	' ESCAPE tutorial
	Public Sub sAttractTutorialEscape()
		resetTutorialSeq
		triggerVideo VIDEO_TUTORIAL_ESCAPE
		
		LightsOn GI, Array(dorange, orange), 100
		LightsBlink tutorialLights5, Array(dorange, orange), 500, 100
	End Sub
	
	' SNIPER tutorial
	Public Sub sAttractTutorialSniper()
		resetTutorialSeq
		triggerVideo VIDEO_TUTORIAL_SNIPER
		
		LightsOn GI, Array(damber, amber), 100
		LightsBlink tutorialLights6, Array(damber, amber), 500, 100
	End Sub
	
	' BLACKSMITH tutorial
	Public Sub sAttractTutorialBlacksmith()
		resetTutorialSeq
		triggerVideo VIDEO_TUTORIAL_BLACKSMITH
		
		LightsOn GI, Array(dpurple, purple), 100
		LightsBlink tutorialLights7, Array(dpurple, purple), 500, 100
	End Sub
	
	' FINAL BOSS tutorial
	Public Sub sAttractTutorialFinalBoss()
		resetTutorialSeq
		triggerVideo VIDEO_TUTORIAL_FINAL_BOSS
		
		LightsOn GI, Array(dred, red), 100
		LightsBlink tutorialLights8, Array(dred, red), 500, 100
		LightsOn tutorialLights8B, Array(dgreen, green), 33
	End Sub
	
	' ? tutorial
	Public Sub sAttractTutorialFinalWizard()
		resetTutorialSeq
		triggerVideo VIDEO_TUTORIAL_FINAL_JUDGMENT
		
		SetLightColorCollection tutorialLights9, white, 2
		SetLightColorCollection GI, white, 1
		SetLightColorCollection tutorialLights9B, green, 1
		
		LightsBlink tutorialLights9, Array(dwhite, white), 500, 100
		LightsOn tutorialLights9B, Array(dgreen, green), 33
	End Sub
	
	'*************************
	' SCORING / CREDITS
	'*************************
	
	' Display score overlay
	Public Sub sShowScores()
		If dataGame.data.Item("gameMode") = 1 Then
			triggerOverlay OVERLAY_SCORE_COOPERATIVE
		Else
			triggerOverlay OVERLAY_SCORE_COMPETITIVE
		End If

		If (usePUP = False Or PUPStatus = False) Then
			' TODO
		Else
			PUP.LabelShowPage pDMD,1,0,""
		End If
		
		sUpdateScores True
	End Sub
	
	' Update PUP with the current scores and player info; clearFirst (Boolean) set to true when we want to reset all the labels (such as a change in player turn or number of players)
	Public Sub sUpdateScores(clearFirst)
		'If Not curOverlay = OVERLAY_SCORBIT_CLAIM And Not curOverlay = OVERLAY_SCORBIT_PAIR And Not curOverlay = OVERLAY_SCORE_COMPETITIVE And Not curOverlay = OVERLAY_SCORE_COOPERATIVE Then Exit Sub

		' Clear current player info
		If (usePUP = False Or PUPStatus = False) Then
			' TODO
		Else
			' Clear Everything when applicable
			If clearFirst = True Then
				PUP.LabelSet pDMD,"p1label","",1,""
				PUP.LabelSet pDMD,"p2label","",1,""
				PUP.LabelSet pDMD,"p3label","",1,""
				PUP.LabelSet pDMD,"p4label","",1,""
				PUP.LabelSet pDMD,"p1hp","",1,""
				PUP.LabelSet pDMD,"p2hp","",1,""
				PUP.LabelSet pDMD,"p3hp","",1,""
				PUP.LabelSet pDMD,"p4hp","",1,""
				PUP.LabelSet pDMD,"p1ac","",1,""
				PUP.LabelSet pDMD,"p2ac","",1,""
				PUP.LabelSet pDMD,"p3ac","",1,""
				PUP.LabelSet pDMD,"p4ac","",1,""
				PUP.LabelSet pDMD,"p1dd","",1,""
				PUP.LabelSet pDMD,"p2dd","",1,""
				PUP.LabelSet pDMD,"p3dd","",1,""
				PUP.LabelSet pDMD,"p4dd","",1,""

				If dataGame.data.Item("gameMode") = 1 Then
					PUP.LabelSet pDMD,"p1score","",0,""
					PUP.LabelSet pDMD,"p2score","",0,""
					PUP.LabelSet pDMD,"p3score","",0,""
					PUP.LabelSet pDMD,"p4score","",0,""
					PUP.LabelSet pDMD,"coopscore","",1,""
				Else
					PUP.LabelSet pDMD,"p1score","",1,""
					PUP.LabelSet pDMD,"p2score","",1,""
					PUP.LabelSet pDMD,"p3score","",1,""
					PUP.LabelSet pDMD,"p4score","",1,""
					PUP.LabelSet pDMD,"coopscore","",0,""
				End If

				PUP.LabelSet pDMD,"p1scorbit","",0,""
				PUP.LabelSet pDMD,"p2scorbit","",0,""
				PUP.LabelSet pDMD,"p3scorbit","",0,""
				PUP.LabelSet pDMD,"p4scorbit","",0,""
			End If
		End If
		
		' Set player information
		Dim numPlayer
		For numPlayer = 1 To 4
			If clearFirst = True Or dataGame.data.Item("playerUp") = numPlayer Then 'Only update current player unless we are clearing everything (helps prevent frequent stutters)
				Dim playerColor
				playerColor = dataPlayer(numPlayer - 1).data.Item("colorSlot")
				
				'HP
				Dim playerHP
				If dataGame.data.Item("gameDifficulty") = 0 And dataPlayer(numPlayer - 1).data.Item("dead") = 0 Then 
					playerHP = "?" 'Don't show player health If in zen mode
					B2S.B2SSetScore (numPlayer + 4), 999
				Else
					playerHP = dataPlayer(numPlayer - 1).data.Item("HP")
					If playerHP <= 0 Then
						If dataGame.data.Item("numPlayers") >= numPlayer And dataPlayer(numPlayer - 1).data.Item("dead") = 0 Then
							playerHP = "SOS"	'Show SOS when a player is about to die
							B2S.B2SSetScore (numPlayer + 4), 1
						Else
							playerHP = "DEAD"	'Show DEAD when a player is actually dead or not playing the game
							B2S.B2SSetScore (numPlayer + 4), 0
						End If
					Else
						B2S.B2SSetScore (numPlayer + 4), playerHP
					End If
				End If
				
				' Default muted colors for inactive player
				Dim playerLabelColor
				playerLabelColor = RGB(128,128,0)
				Dim hpColor
				hpColor = RGB(96,128,96)
				Dim acColor
				acColor = RGB(128,96,128)
				Dim ddColor
				ddColor = RGB(128,96,96)

				' If player is active / up, use brighter colors
				If dataGame.data.Item("playerUp") = numPlayer Then
					playerLabelColor = RGB(255,255,0)
					hpColor = RGB(192,255,192)
					acColor = RGB(255,192,255)
					ddColor = RGB(255,192,192)
					B2S.B2SSetPlayerUp numPlayer
				ElseIf dataGame.data.Item("playerUp") = 0 And numPlayer = 1 Then
					B2S.B2SSetPlayerUp 0
				End If

				B2S.B2SSetScore (numPlayer + 8), dataPlayer(numPlayer - 1).data.Item("AC")
				B2S.B2SSetScore (numPlayer + 12), dataPlayer(numPlayer - 1).data.Item("drainDamage")
				
				If (usePUP = False Or PUPStatus = False) Then
					' TODO
				Else
					PUP.LabelSet pDMD,"p" & playerColor & "label",Left(dataPlayer(numPlayer - 1).data.Item("name"), 15),1,"{'mt':2,'color':" & playerLabelColor & "}"
					PUP.LabelSet pDMD,"p" & playerColor & "hp","❤" & playerHP,1,"{'mt':2,'color':" & hpColor & "}"
					PUP.LabelSet pDMD,"p" & playerColor & "ac","🛡️" & dataPlayer(numPlayer - 1).data.Item("AC"),1,"{'mt':2,'color':" & acColor & "}"
					PUP.LabelSet pDMD,"p" & playerColor & "dd","💀" & dataPlayer(numPlayer - 1).data.Item("drainDamage"),1,"{'mt':2,'color':" & ddColor & "}"
				End If
			End If

			'Scores must be kept separate in case of coop mode
			Dim playerScore
			Dim scoreColor
			If (clearFirst = True Or dataGame.data.Item("playerUp") = numPlayer) And dataGame.data.Item("gameMode") = 0 Then
				playerScore = FormatScore(dataPlayer(numPlayer - 1).data.Item("score"))

				scoreColor = RGB(192,192,255)
				If dataGame.data.Item("playerUp") = numPlayer Then scoreColor = RGB(255,255,255)

				If (usePUP = False Or PUPStatus = False) Then
					' TODO
				Else
					PUP.LabelSet pDMD,"p" & playerColor & "score","" & playerScore,1,"{'mt':2,'color':" & scoreColor & "}"
				End If

				'Update Backglass
				B2S.B2SSetScorePlayer numPlayer, dataPlayer(numPlayer - 1).data.Item("score")
			ElseIf dataGame.data.Item("gameMode") = 1 Then
				playerScore = FormatScore(dataPlayer(0).data.Item("score"))
				scoreColor = RGB(255,255,255)

				If (usePUP = False Or PUPStatus = False) Then
					' TODO
				Else
					PUP.LabelSet pDMD,"coopscore","" & playerScore,1,"{'mt':2,'color':" & scoreColor & "}"
				End If

				'Update Backglass
				B2S.B2SSetScorePlayer numPlayer, dataPlayer(0).data.Item("score")
			End If
		Next

		' Scorbit icons for players who linked their Scorbit
		If clearFirst = True And usePUP = True And PUPStatus = True And dataGame.data.Item("gameMode") = 0 Then
			If Not dataPlayer(0).data.Item("name") = "" And Not dataPlayer(0).data.Item("name") = "PLAYER 1" Then 
				PUP.LabelSet pDMD, "p1scorbit","PuPOverlays\\scorbit_logo.png",1,"{'mt':2,'width':2.5,'height': 4.074074074,'xalign':0,'yalign':0,'ypos':92.222222222,'xpos': 0.3125}"
			End If
			If Not dataPlayer(1).data.Item("name") = "" And Not dataPlayer(1).data.Item("name") = "PLAYER 2" Then 
				PUP.LabelSet pDMD, "p2scorbit","PuPOverlays\\scorbit_logo.png",1,"{'mt':2,'width':2.5,'height': 4.074074074,'xalign':0,'yalign':0,'ypos':92.222222222,'xpos': 25.3125}"
			End If
			If Not dataPlayer(2).data.Item("name") = "" And Not dataPlayer(2).data.Item("name") = "PLAYER 3" Then 
				PUP.LabelSet pDMD, "p3scorbit","PuPOverlays\\scorbit_logo.png",1,"{'mt':2,'width':2.5,'height': 4.074074074,'xalign':0,'yalign':0,'ypos':92.222222222,'xpos': 50.3125}"
			End If
			If Not dataPlayer(3).data.Item("name") = "" And Not dataPlayer(3).data.Item("name") = "PLAYER 4" Then 
				PUP.LabelSet pDMD, "p4scorbit","PuPOverlays\\scorbit_logo.png",1,"{'mt':2,'width':2.5,'height': 4.074074074,'xalign':0,'yalign':0,'ypos':92.222222222,'xpos': 75.3125}"
			End If
		End If
	End Sub
	
	'*************************
	' BASIC GAMEPLAY
	'*************************

	Public Sub promptInstructions()
		PlayCallout "callout_nar_firstball_instr", 1, VOLUME_CALLOUTS
		triggerVideo VIDEO_GAMEPLAY_INSTRUCTIONS

		DMD.jpFlush
		DMD.jpLocked = True
		DMD.jpDMD DMD.jpCL("Shoot the"), DMD.jpCL("flashing letters"), "", eNone, eNone, eNone, 1620, False, ""
		DMD.jpDMD DMD.jpCL("to spell"), DMD.jpCL("words"), "", eNone, eNone, eNone, 1580, False, ""
		DMD.jpDMD DMD.jpCL("Spell words to"), DMD.jpCL("start objectives"), "", eBlink, eBlink, eNone, 2620, True, ""
	End Sub
	
	' Set the line in the display indicating the current game play mode / status.
	Public Sub gameStatus(statusLabel)
		If (usePUP = False Or PUPStatus = False) Then
			' TODO
		Else
			PUP.LabelSet pDMD,"gameplay",statusLabel,1,""
		End If
	End Sub
	
	' Queue / run a ball search routine
	Public Sub searchForBall()
		Dim i
		Dim i2
		
		If (GameTime - BALL_SEARCH_TIME) < 10000 Or (CURRENT_MODE < 800 Or (CURRENT_MODE >= 900 And CURRENT_MODE < 1000)) Or BIPL > 0 Then
			If Not BALL_SEARCH_STAGE = 0 Then updateClocks
			BALL_SEARCH_STAGE = 0
			Exit Sub
		End If

		If DEBUG_LOG_BALL_SEARCH = True Then WriteToLog "Ball Search", "Stage " & BALL_SEARCH_STAGE & " of 8 started."

		'Reset timers to give the player a grace; give 10 seconds ball shield (+2 seconds grace) and reset mode timer to 10 seconds if each is below 10 seconds.
		If Clocks.data.Item("shield").timeLeft < 12000 Then setBallShield 10000
		If Clocks.data.Item("mode").timeLeft < 10000 And Clocks.data.Item("mode").timeLeft > 0 Then Clocks.data.Item("mode").timeLeft = 10000
		
		If BALL_SEARCH_STAGE >= 8 Then
			If DEBUG_LOG_BALL_SEARCH = True Then WriteToLog "Ball Search", "Could not find ball. Ball search terminated and game paused."
			BALL_SEARCH_TIME = GameTime
			BALL_SEARCH_STAGE = 0
			startMode MODE_MISSING_BALL
			Exit Sub
		ElseIf BALL_SEARCH_STAGE = 0 Then 'Determine if we have a discrepancy between number of actual balls in play and number of balls in play the table thinks; generally should never happen as BIP is updated with the trough
			BALL_SEARCH_STAGE = BALL_SEARCH_STAGE + 1
			updateClocks
			PlaySound "fx_solenoid", 1, VolumeDial
			
			Dim actualBIP
			actualBIP = tnob - lob - actualBallsInTrough - BIS

			If DEBUG_LOG_BALL_SEARCH = True Then WriteToLog "Ball Search", "We want " & BQueue & " balls in play (if 0, we are not pending to launch any); there are " & BIPL & " balls in the plunger lane."
			If DEBUG_LOG_BALL_SEARCH = True Then WriteToLog "Ball Search", "Balls we think are in play: " & BIP & "; Balls actually in play: " & actualBIP	
			If BQueue > 0 And BIPL = 0 Then	'Balls waiting to be released from the trough; probably a stuck release timer, so enable the timer
				BALL_SEARCH_STAGE = 0
				updateClocks

				UpdateTrough
				swTrough1.TimerEnabled = True
			ElseIf actualBIP < BIP Then 'Balls drained and the table missed it
				BALL_SEARCH_STAGE = 0
				updateClocks
				
				'Run the drain hit event
				BIPWhenDrained = actualBIP
				checkDrainedBalls True
			Else 'Balls in play count is correct; wait a little longer and then activate ball search
				queue.Add "ball_search_again","DMD.searchForBall",1,0,5000,100,0,False
			End If
		Else
			If BALL_SEARCH_STAGE = 1 Then PlaySound "ui_ball_search", 1, VOLUME_SFX
			
			triggerVideo VIDEO_BALL_SEARCH
			BALL_SEARCH_STAGE = BALL_SEARCH_STAGE + 1
			updateClocks
			
			PlayCallout "callout_lov_ball_search_" & BALL_SEARCH_STAGE, 1, VOLUME_CALLOUTS
			
			If BALL_SEARCH_STAGE = 4 Then
				BALL_SEARCH_MUSIC_MEMORY = Music.nowPlaying
				Music.PlayTrack "music_ball_search_intro", Null, Null, Null, Null, 0, 0, 0, 0
				queue.Add "ball_search_delay","DMD.ballSearchStep 0",1,0,8000,0,0,False
				queue.Add "ball_search_1","DMD.ballSearchStep 1",1,0,0,1000,0,False
			ElseIf BALL_SEARCH_STAGE = 8 Then
				StopLightSequence StandardSeq
				Music.PlayTrack "music_ball_search_intro", Null, Null, Null, Null, 0, 0, 0, 0
				queue.Add "ball_search_delay","DMD.ballSearchStep 2",1,0,8000,0,0,False
				queue.Add "ball_search_1","DMD.ballSearchStep 3",1,0,0,200,0,False
				
				For i2 = 0 To 3
					queue.Add "ball_search_3_" & i2,"DMD.ballSearchStep 4",1,0,0,500,0,False
					queue.Add "ball_search_4_" & i2,"DMD.ballSearchStep 5",1,0,0,500,0,False
					queue.Add "ball_search_5_" & i2,"DMD.ballSearchStep 6",1,0,0,500,0,False
					queue.Add "ball_search_6_" & i2,"DMD.balLSearchStep 7",1,0,0,500,0,False
					queue.Add "ball_search_7_" & i2,"DMD.ballSearchStep 8",1,0,0,500,0,False
					queue.Add "ball_search_9_" & i2,"DMD.ballSearchStep 9",1,0,0,500,0,False
					queue.Add "ball_search_10_" & i2,"DMD.ballSearchStep 10",1,0,0,500,0,False
					queue.Add "ball_search_11_" & i2,"DMD.ballSearchStep 11",1,0,0,500,0,False
				Next
				
				queue.Add "ball_search_13","DMD.ballSearchStep 12",1,0,0,500,0,False
				queue.Add "ball_search_again","DMD.balLSearchStep 13",1,0,7500,0,0,False
				
				Exit Sub
			Else
				queue.Add "ball_search_1","DMD.ballSearchStep 1",1,0,3000,1000,0,False
			End If

			'If you change anything below, also modify glitchProgression
			
			queue.Add "ball_search_2","DMD.ballSearchStep 14",1,0,0,200,0,False
			
			queue.Add "ball_search_3","DMD.ballSearchStep 9",1,0,0,200,0,False
			queue.Add "ball_search_4","DMD.ballSearchStep 10",1,0,0,200,0,False
			queue.Add "ball_search_5","DMD.ballSearchStep 11",1,0,0,200,0,False
			
			queue.Add "ball_search_6","DMD.ballSearchStep 15",1,0,0,200,0,False
			queue.Add "ball_search_7","DMD.ballSearchStep 16",1,0,0,200,0,False
			queue.Add "ball_search_8","DMD.ballSearchStep 17",1,0,0,200,0,False
			
			queue.Add "ball_search_9","DMD.ballSearchStep 18",1,0,0,200,0,False
			queue.Add "ball_search_10","DMD.ballSearchStep 19",1,0,0,200,0,False
			
			If BALL_SEARCH_STAGE > 3 Then
				queue.Add "ball_search_11","DMD.ballSearchStep 20",1,0,0,200,0,False
				queue.Add "ball_search_12","DMD.ballSearchStep 21",1,0,0,200,0,False
				queue.Add "ball_search_13","DMD.ballSearchStep 22",1,0,0,200,0,False
				queue.Add "ball_search_14","DMD.ballSearchStep 23",1,0,0,200,0,False
				queue.Add "ball_search_15","DMD.ballSearchStep 24",1,0,0,6000,0,False
				
				queue.Add "ball_search_again","DMD.searchForBall",1,0,7000,100,0,False
			Else
				queue.Add "ball_search_again","DMD.searchForBall",1,0,3000,100,0,False
			End If
		End If
	End Sub
	
	' Perform a step in the ball searching
	Public Sub ballSearchStep(sStep)
		Dim i
		Dim i2
		
		Select Case sStep
			Case 0:
				Music.PlayTrack "music_ball_search_main", Null, Null, Null, Null, 0, 0, 0, 0
				StartLightSequence StandardSeq, Array(dred, red), 250, SeqRandom, 25, 1
			Case 1:
				ShakerMotorTimer.Enabled = 1
				SolVUKDiverter True
			Case 2:
				Music.PlayTrack "music_ball_search_ultimate", Null, Null, Null, Null, 0, 0, 0, 0
				StartLightSequence StandardSeq, Array(dred, red), 100, SeqRandom, 25, 1
			Case 3:
				ShakerMotorTimer.Enabled = 1
				
				' Post drops run on their own queue
				For i2 = 0 To 1
					For i = 0 To 29
						postDropQueue.Add "ball_search_2_" & i & "_" & i2,"SolPegDrop " & i & ", True",1,0,0,100,0,False
					Next
					postDropQueue.Add "ball_search_2_delay_1_" & i2,"",1,0,0,1000,0,False
					For i = 0 To 29
						postDropQueue.Add "ball_search_2b_" & i & "_" & i2,"SolPegDrop " & i & ", False",1,0,0,100,0,False
					Next
					postDropQueue.Add "ball_search_2_delay_2_" & i2,"",1,0,0,1000,0,False
				Next
			Case 4:
				SolVUKDiverter True
			Case 5:
				SolLSling True
				SolRFlipper True
			Case 6:
				SolRSling True
				SolRFlipper False
				SolLFlipper True
			Case 7:
				SolLSling False
				SolRSling False
				SolLFlipper False
				SolUFlipper True
			Case 8:
				SolVUKDiverter False
				SolUFlipper False
				SolVUpKicker True
				SolScoop True
			Case 9:
				SolBumper1 True
			Case 10:
				SolBumper2 True
			Case 11:
				SolBumper3 True
			Case 12:
				ShakerMotorTimer.Enabled = 0
			Case 13:
				StopLightSequence StandardSeq
				DMD.searchForBall
			Case 14:
				ShakerMotorTimer.Enabled = 0
				SolVUKDiverter False
			Case 15:
				SolLSling True
			Case 16:
				SolRSling True
			Case 17:
				SolLSling False
				SolRSling False
			Case 18:
				SolScoop True
			Case 19:
				SolVUpKicker True
			Case 20:
				SolLFlipper True
			Case 21:
				SolLFlipper False
				SolRFlipper True
			Case 22:
				SolRFlipper False
				SolUFlipper True
			Case 23:
				SolUFlipper False
			Case 24:
				For i = 0 To 29
					postDropQueue.Add "ball_search_24_" & i,"SolPegDrop " & i & ", True",1,0,0,100,0,False
				Next
				For i = 0 To 29
					postDropQueue.Add "ball_search_24b_" & i,"SolPegDrop " & i & ", False",1,0,0,100,0,False
				Next
		End Select
	End Sub
	
	' Called by startMode MODE_MISSING_BALL to "summon" the operator to find the ball
	Public Sub missingBallSummon()
		PlayCallout "callout_lov_ball_search_summon", 1, VOLUME_CALLOUTS
		
		' Rainbow all the lights
		LightsOn coloredLights, Array(dbase, base), 100
		LightsOn GI, Array(dbase, base), 100
		StartRainbow coloredLights
		
		' Show the screen for the operator
		triggerVideo VIDEO_BALL_MISSING
		
		' Play some epic chanting
		Music.PlayTrack "music_eerie_chant", Null, Null, Null, Null, 0, 0, 0, 0
	End Sub
	
	' Called by the flipper keys to toggle which difficulty to select
	Public Sub selectDifficulty(newDifficulty)
		'Do not allow choosing Zen if explicitly disabled in script
		If newDifficulty = 0 And ALLOW_ZEN = False Then Exit Sub
		
		dataGame.data.Item("gameDifficulty") = newDifficulty
		
		Select Case newDifficulty
			Case 0
				'zen
				PlayCallout "callout_nar_zen", 1, VOLUME_CALLOUTS
				triggerVideo VIDEO_DIFFICULTY_ZEN
				LightsOn coloredLights, Array(dpurple, purple), 100
				LightsOn GI, Array(dpurple, purple), 100
			Case 1
				'easy
				PlayCallout "callout_nar_easy", 1, VOLUME_CALLOUTS
				triggerVideo VIDEO_DIFFICULTY_EASY
				LightsOn coloredLights, Array(dgreen, green), 100
				LightsOn GI, Array(dgreen, green), 100
			Case 2
				'normal
				PlayCallout "callout_nar_normal", 1, VOLUME_CALLOUTS
				triggerVideo VIDEO_DIFFICULTY_NORMAL
				LightsOn coloredLights, Array(dyellow, yellow), 100
				LightsOn GI, Array(dyellow, yellow), 100
			Case 3
				'hard
				PlayCallout "callout_nar_hard", 1, VOLUME_CALLOUTS
				triggerVideo VIDEO_DIFFICULTY_HARD
				LightsOn coloredLights, Array(dorange, orange), 100
				LightsOn GI, Array(dorange, orange), 100
			Case 4
				'impossible
				PlayCallout "callout_nar_impossible", 1, VOLUME_CALLOUTS
				triggerVideo VIDEO_DIFFICULTY_IMPOSSIBLE
				LightsOn coloredLights, Array(dred, red), 100
				LightsOn GI, Array(dred, red), 100
		End Select
	End Sub
	
	' Sequence on a skillshot
	Public Sub skillshot(ssComplete, trainLight)
		If ssComplete = True Then
			triggerVideo VIDEO_SKILLSHOT_COMPLETE
			PlaySound "sfx_skillshot_complete", 1, VOLUME_SFX
			PlayCallout "callout_nar_skillshot", 1, VOLUME_CALLOUTS
			
			StandardSeq.CenterX = trainLight.X
			StandardSeq.CenterY = trainLight.Y
			StartLightSequence StandardSeq, Array(dpurple, purple), 10, SeqCircleOutOn, 0, 1
			flasherQueue.Add "skillshot_1","Flash6 True",1,0,0,100,0,False
			flasherQueue.Add "skillshot_2","Flash6 False",1,0,0,200,0,False
			flasherQueue.Add "skillshot_3","Flash6 True",1,0,0,100,0,False
			flasherQueue.Add "skillshot_4","Flash6 False",1,0,0,200,0,False
			flasherQueue.Add "skillshot_5","Flash6 True",1,0,0,100,0,False
			flasherQueue.Add "skillshot_6","Flash6 False",1,0,0,200,0,False
		Else
			triggerVideo VIDEO_SKILLSHOT_FAILED
			PlaySound "sfx_skillshot_failed", 1, VOLUME_SFX
			PlayCallout "callout_nar_skillshot_failed_" & Int(Rnd*3), 1, VOLUME_CALLOUTS
		End If
	End Sub
	
	' Sequence played via queue whenever the player takes damage
	Public Sub damageReceived(damageHP)
		'TODO: Show how much damage was taken on PUP label
		If BIP <= 1 Then 
			triggerVideo VIDEO_DAMAGE_RECEIVED
			DMD.jpLocked = True
			DMD.jpDMD DMD.jpCL("You Took Damage!"), DMD.jpCL("-" & damageHP & " HP"), "d_E145_0", eBlinkFast, eNone, eNone, 2500, True, ""
		End If
		StartLightSequence StandardSeq, Array(dred, red), 250, SeqBlinking, 0, 1
	End Sub
	
	' End the current player's turn via a drain (startNextPlayer needs to be called separately)
	Public Sub currentPlayerEnd()
		'Queue up end of turn sequence
		Music.StopTrack Null
		'	queue.Add "drain_player_L","DMD.triggerVideo VIDEO_CLEAR : DMD.triggerVideo 52",2,0,0,1500,0,False
		queue.Add "drain_player_endB","PlaySound ""sting_endofturn"", 1, VOLUME_SFX : DMD.currentPlayerEndCallout",2,0,0,4000,0,False
		
		'Queue up end of game bonus if the player died, else queue up next player
		If currentPlayer("HP") < 0 And Not dataGame.data.Item("gameDifficulty") = 0 Then
			queue.Add "drain_player_endC","startMode MODE_END_OF_GAME_BONUS",1,0,0,100,0,False
		Else
			queue.Add "drain_player_endC","DMD.startNextPlayer False",1,0,0,100,0,False
		End If
		
		' Lights Off
		LightsOff GI
		LightsOff aLights
		
		' TODO: Drain.Y throws an error for some stupid reason...
		StandardSeq.CenterX = 430
		StandardSeq.CenterY = 1863
		StartLightSequence StandardSeq, Array(dred, red), 10, SeqCircleInOff, 0, 1

		' Stop +HP timer; if it's active, drops will be reset on the next player / skillshot
		Clocks.data.Item("hpTargets").timeLeft = 0
	End Sub
	
	' Play a random callout for ending a player
	Public Sub currentPlayerEndCallout()
		PlayCallout "callout_nar_drain_" & Int(Rnd*3), 1, VOLUME_CALLOUTS
	End Sub
	
	' Initialize the next player in sequential order (but since we use HP, we have to work some magic)
	Public Sub startNextPlayer(cameFromBonus)
		If dataGame.data.Item("numPlayers") < 1 Then Exit Sub
		
		Dim i
		For i = 1 To 4 ' Maximum number of players supported
			dataGame.data.Item("playerUp") = dataGame.data.Item("playerUp") + 1
			If dataGame.data.Item("playerUp") > dataGame.data.Item("numPlayers") Then 
				dataGame.data.Item("playerUp") = 1
				dataGame.data.Item("ball") = dataGame.data.Item("ball") + 1
			End If
			
			' If the selected player still is alive, then it is their turn. Set up the skillshot etc. and exit the sub.
			If currentPlayer("dead") = 0 Then
				'queue.Add "start_player_0","DMD.triggerVideo VIDEO_CLEAR",1,0,0,100,0,False
				If dataGame.data.Item("numPlayers") > 1 Then
					Select Case dataGame.data.Item("playerUp")
						Case 1:
							queue.Add "start_player_A","DMD.triggerVideo VIDEO_PLAYER_1 : StartLightSequence StandardSeq, Array(dred, red), 1500, SeqBlinking, 0, 1 : PlayCallout ""callout_nar_player_1"", 1, VOLUME_CALLOUTS",1,0,0,3000,0,False
						Case 2:
							queue.Add "start_player_A","DMD.triggerVideo VIDEO_PLAYER_2 : StartLightSequence StandardSeq, Array(dyellow, yellow), 1500, SeqBlinking, 0, 1 : PlayCallout ""callout_nar_player_2"", 1, VOLUME_CALLOUTS",1,0,0,3000,0,False
						Case 3:
							queue.Add "start_player_A","DMD.triggerVideo VIDEO_PLAYER_3 : StartLightSequence StandardSeq, Array(dgreen, green), 1500, SeqBlinking, 0, 1 : PlayCallout ""callout_nar_player_3"", 1, VOLUME_CALLOUTS",1,0,0,3000,0,False
						Case 4:
							queue.Add "start_player_A","DMD.triggerVideo VIDEO_PLAYER_4 : StartLightSequence StandardSeq, Array(dblue, blue), 1500, SeqBlinking, 0, 1 : PlayCallout ""callout_nar_player_4"", 1, VOLUME_CALLOUTS",1,0,0,3000,0,False
					End Select
				End If
				queue.Add "start_player_B","startMode MODE_SKILLSHOT",1,0,0,100,0,False

				DMD.sUpdateScores True
				Exit Sub
			End If
		Next
		
		'If we did not exit the sub by this point, then everyone died and the game is over. In coop mode, we have to force the end of game bonus first if we did not already come from it.
		If dataGame.data.Item("gameMode") = 0 Or cameFromBonus = True Then
			startMode MODE_GAME_OVER
		Else
			CURRENT_MODE = MODE_END_OF_GAME_BONUS
			DMD.endOfGameBonus
		End If
	End Sub
	
	' Prepare to eject the ball from the scoop by giving a warning first
	Public Sub SolScoopWarning()
		If Scoop.TimerEnabled = False Then
			flasherQueue.Add "ball_eject_1","PlaySound ""sfx_ball_ejection"", 1, VOLUME_SFX, 0.5 : Scoop.TimerEnabled = True : Flash4 True",1,0,0,50,0,False
			flasherQueue.Add "ball_eject_2","Flash4 False",1,0,0,100,0,False
			flasherQueue.Add "ball_eject_3","Flash4 True",1,0,0,50,0,False
			flasherQueue.Add "ball_eject_4","Flash4 False",1,0,0,100,0,False
			flasherQueue.Add "ball_eject_5","Flash4 True",1,0,0,50,0,False
			flasherQueue.Add "ball_eject_6","Flash4 False",1,0,0,100,0,False
			flasherQueue.Add "ball_eject_7","Flash4 True",1,0,0,50,0,False
			flasherQueue.Add "ball_eject_8","Flash4 False",1,0,0,100,0,False
		End If
	End Sub

	Public Sub TiltWarning()
		triggerVideo VIDEO_TILT

		PlaySound "sfx_nudge_too_hard", 1, VOLUME_SFX
		flasherQueue.Add "nudge_warn_1","Flash1 True",1,0,0,200,0,False
		flasherQueue.Add "nudge_warn_2","Flash1 False",1,0,0,200,0,False
		flasherQueue.Add "nudge_warn_3","Flash1 True",1,0,0,200,0,False
		flasherQueue.Add "nudge_warn_4","Flash1 False",1,0,0,200,0,False
		flasherQueue.Add "nudge_warn_5","Flash1 True",1,0,0,200,0,False
		flasherQueue.Add "nudge_warn_6","Flash1 False",1,0,0,200,0,False
		flasherQueue.Add "nudge_warn_7","Flash1 True",1,0,0,200,0,False
		flasherQueue.Add "nudge_warn_8","Flash1 False",1,0,0,200,0,False
	End Sub
	
	'*************************
	' END OF GAME / DEATH SAVE
	'*************************
	
	'Mark the player as dead and proceed on bonuses / penalties
	Public Sub playerIsDead()
		'Mark dead
		currentPlayerSet "dead", 1
		DMD.sUpdateScores False

		'In Zen, we have to determine HP damage penalty
		If dataGame.data.Item("gameDifficulty") = 0 Then
			DMD.endOfZenPenalty
		End If

		'In coop mode, end of game bonus should not be awarded until all players are dead.
		If dataGame.data.Item("gameMode") = 0 Then 
			queue.Add "end_of_game_bonus", "DMD.endOfGameBonus", 1, 0, 0, 3800, 0, False
		Else
			queue.Add "coop_next_player", "DMD.startNextPlayer False", 1, 0, 0, 3800, 0, False
		End If
	End Sub

	'Tally the HP damage penalty the player should take at the end of a Zen game
	Public Sub endOfZenPenalty()
		'Sequence
		Music.PlayTrack "music_zen_suspense", Null, Null, Null, Null, 0, 0, 0, 0
		'DMD.triggerVideo VIDEO_CLEAR
		DMD.triggerVideo VIDEO_ZEN_DAMAGE_PENALTY

		'Show suspense labels
		If (usePUP = False Or PUPStatus = False) Then
			' TODO
		Else
			PUP.LabelSet pDMD,"center_top","??? HP Damage",1,""
			PUP.LabelSet pDMD,"center_bottom","???,???,???",1,""
		End If

		'Queue next part
		queue.Add "end_of_zen_penalty_2", "DMD.endOfZenPenaltyActualization", 1, 0, 14500, 5000, 0, False
	End Sub

	'Actually apply and show the HP penalty
	Public Sub endOfZenPenaltyActualization()
		Dim i
		'Calculate penalty
		Dim penaltyHP, penaltyScore
		penaltyHP = -(currentPlayer("HP"))
		penaltyScore = penaltyHP * Scoring.penalties("zen")

		'Apply penalty
		addScore 0 - penaltyScore, "Zen HP Damage Penalty"

		'Do lights for suspense
		LightsOn coloredLights, Array(dred, red), 100
		For i = 0 to 9
			flasherQueue.Add "zen_penalty_on_" & i,"Flash1 True",1,0,0,100,0,False
			flasherQueue.Add "zen_penalty_off_" & i,"Flash1 False",1,0,0,100,0,False
		Next

		'Show labels
		If usePUP = True And PUPStatus = True Then
			PUP.LabelSet pDMD,"center_top","" & penaltyHP & " HP Damage",1,""
			PUP.LabelSet pDMD,"center_bottom",FormatScore(penaltyScore),1,""
		End If

		DMD.jpFlush
		DMD.jpDMD DMD.jpCL(penaltyHP & " HP Damage"), DMD.jpCL(FormatScore(penaltyScore)), "d_E170_0", eNone, eBlink, eNone, 5000, False, ""
	End Sub
	
	' Tally end of game bonus / show that the player is dead
	Public Sub endOfGameBonus()
		Dim i

		'GI
		LightsOff aLights
		LightsOn GI, Array(dred, red), 25

		'Sequence
		Music.PlayTrack "music_dead_bonus", Null, Null, Null, Null, 0, 0, 0, 0
		'DMD.triggerVideo VIDEO_CLEAR
		DMD.triggerVideo VIDEO_DEAD

		'Hide labels at first (in case we are coming from Zen penalty)
		If (usePUP = False Or PUPStatus = False) Then
			' TODO
		Else
			PUP.LabelSet pDMD,"center_top","",0,""
			PUP.LabelSet pDMD,"center_bottom","",0,""
		End If
		
		'Queue bonus tallies
		queue.Add "CASTLE_bonus","DMD.endOfGameBonus_display ObjectiveCastleL, Null, ""CASTLE Bonus"", """ & FormatScore(currentPlayerBonus("CASTLE", False)) & """",1,0,0,2237,0,False
		queue.Add "CHASER_bonus","DMD.endOfGameBonus_display ObjectiveChaserL, ObjectiveCastleL, ""CHASER Bonus"", """ & FormatScore(currentPlayerBonus("CHASER", False)) & """",1,0,0,2237,0,False
		queue.Add "DRAGON_bonus","DMD.endOfGameBonus_display ObjectiveDragonL, ObjectiveChaserL, ""DRAGON Bonus"", """ & FormatScore(currentPlayerBonus("DRAGON", False)) & """",1,0,0,2237,0,False
		queue.Add "VIKING_bonus","DMD.endOfGameBonus_display ObjectiveVikingL, ObjectiveDragonL, ""VIKING Bonus"", """ & FormatScore(currentPlayerBonus("VIKING", False)) & """",1,0,0,2237,0,False
		queue.Add "ESCAPE_bonus","DMD.endOfGameBonus_display ObjectiveEscapeL, ObjectiveVikingL, ""ESCAPE Bonus"", """ & FormatScore(currentPlayerBonus("ESCAPE", False)) & """",1,0,0,2237,0,False
		queue.Add "SNIPER_bonus","DMD.endOfGameBonus_display ObjectiveSniperL, ObjectiveEscapeL, ""SNIPER Bonus"", """ & FormatScore(currentPlayerBonus("SNIPER", False)) & """",1,0,0,2237,0,False
		queue.Add "BLACKSMITH_bonus","DMD.endOfGameBonus_display ObjectiveBlacksmithL, ObjectiveSniperL, ""BLACKSMITH Bonus"", """ & FormatScore(currentPlayerBonus("BLACKSMITH", False)) & """",1,0,0,2237,0,False
		queue.Add "FINAL_BOSS_bonus","DMD.endOfGameBonus_display ObjectiveFinalBossL, ObjectiveBlacksmithL, ""FINAL BOSS Bonus"", """ & FormatScore(currentPlayerBonus("FINAL BOSS", False)) & """",1,0,0,2237,0,False
		queue.Add "ULTIMATE_WIZARD_bonus","DMD.endOfGameBonus_display ObjectiveFinalJudgmentL, ObjectiveFinalBossL, ""JUDGMENT Bonus"", """ & FormatScore(currentPlayerBonus("FINAL JUDGMENT", False)) & """",1,0,0,2237,0,False
		queue.Add "bonusX","DMD.endOfGameBonus_display Null, ObjectiveFinalJudgmentL, ""Bonus Multiplier"", """ & currentPlayer("bonusX") & " X""",1,0,0,2237,0,False
		
		' Total Bonus
		queue.Add "total_bonus","DMD.endOfGameBonus_display Null, Null, ""Total Bonus"", """ & FormatScore(currentPlayerBonus("all", True)) & """ : addScore " & currentPlayerBonus("all", True) & ", ""End of Game Bonus""",1,0,0,2237,0,False
		queue.Add "check_for_highscore","DMD.checkForHighScore",1,0,0,100,0,False
	End Sub
	
	Public Sub endOfGameBonus_display(objBlink, objOn, textTop, textBottom)
		If usePUP = True And PUPStatus = True Then
			PUP.LabelSet pDMD,"center_top",textTop,1,""
			PUP.LabelSet pDMD,"center_bottom",textBottom,1,""
		End If

		DMD.jpFlush
		DMD.jpDMD DMD.jpCL(textTop), DMD.jpCL(textBottom), "d_E153_0", eScrollRight, eScrollLeft, eNone, 3000, True, ""
		
		If Not IsNull(objBlink) Then
			LampC.LightColor objBlink, Array(dYellow, yellow)
			LampC.Blink objBlink
		End If
		
		If Not IsNull(objOn) Then
			LampC.LightOnWithColor objOn, Array(dgreen, green)
		End If
	End Sub

	Public Sub checkForHighScore()
		If Not dataGame.data.Item("tournamentMode") = 0 And currentPlayer("score") > dataHighScores(5).data.Item("HS5") Then
			highScoreSequence
		ElseIf currentPlayer("score") > dataHighScores(dataGame.data.Item("gameDifficulty")).data.Item("HS5") Then
			highScoreSequence
		Else
			queue.Add "show_final_score","DMD.triggerVideo VIDEO_FINAL_SCORE",1,0,0,5000,0,False
			queue.Add "next_player","DMD.startNextPlayer True",1,0,0,100,0,False
		End If
	End Sub

	Public Sub highScoreSequence()
		'TODO: triggerVideo
		PlaySound "sfx_highscore_buildup", 1, VOLUME_SFX

		DMD.jpFlush
		DMD.jpLocked = True

		'Add our fancy table freak-out routine
		queue.Add "highscore_sequence_end","",1,0,0,18500,0,False
		queueB.Add "highscore_sequence_1a","DMD.highScoreSequence_on",1,300,0,100,0,False
		queueB.Add "highscore_sequence_1b","DMD.highScoreSequence_off",1,1686,0,100,0,False
		queueB.Add "highscore_sequence_2a","DMD.highScoreSequence_on",1,3738,0,100,0,False
		queueB.Add "highscore_sequence_2b","DMD.highScoreSequence_off",1,5152,0,100,0,False
		queueB.Add "highscore_sequence_3a","DMD.highScoreSequence_on",1,6915,0,100,0,False
		queueB.Add "highscore_sequence_3b","DMD.highScoreSequence_off",1,8318,0,100,0,False
		queueB.Add "highscore_sequence_4a","DMD.highScoreSequence_on",1,9516,0,100,0,False
		queueB.Add "highscore_sequence_4b","DMD.highScoreSequence_off",1,10908,0,100,0,False
		queueB.Add "highscore_sequence_5a","DMD.highScoreSequence_on",1,11596,0,100,0,False
		queueB.Add "highscore_sequence_5b","DMD.highScoreSequence_off",1,12865,0,100,0,False
		queueB.Add "highscore_sequence_6a","DMD.highScoreSequence_on",1,13104,0,100,0,False
		queueB.Add "highscore_sequence_6b","DMD.highScoreSequence_off",1,13836,0,100,0,False
		queueB.Add "highscore_sequence_7a","DMD.highScoreSequence_on",1,13975,0,100,0,False
		queueB.Add "highscore_sequence_7b","DMD.highScoreSequence_off",1,14557,0,100,0,False
		queueB.Add "highscore_sequence_8a","DMD.highScoreSequence_on",1,14607,0,100,0,False
		queueB.Add "highscore_sequence_8b","DMD.highScoreSequence_off",1,15161,0,100,0,False
		queueB.Add "highscore_sequence_9a","DMD.highScoreSequence_on",1,15211,0,100,0,False
		queueB.Add "highscore_sequence_9b","DMD.highScoreSequence_off",1,15605,0,100,0,False
		queueB.Add "highscore_sequence_10a","DMD.highScoreSequence_on",1,15655,0,100,0,False
		queueB.Add "highscore_sequence_10b","DMD.highScoreSequence_off",1,16048,0,100,0,False
		queueB.Add "highscore_sequence_11a","DMD.highScoreSequence_on",1,16098,0,100,0,False
		queueB.Add "highscore_sequence_11b","DMD.highScoreSequence_off",1,16298,0,100,0,False
		queueB.Add "highscore_sequence_12a","DMD.highScoreSequence_on",1,16348,0,100,0,False
		queueB.Add "highscore_sequence_12b","DMD.highScoreSequence_off",1,16531,0,100,0,False
		queueB.Add "highscore_sequence_13a","DMD.highScoreSequence_on",1,16581,0,100,0,False
		queueB.Add "highscore_sequence_13b","DMD.highScoreSequence_off",1,16780,0,100,0,False
		queueB.Add "highscore_sequence_14a","DMD.highScoreSequence_on",1,16830,0,100,0,False
		queueB.Add "highscore_sequence_14b","DMD.highScoreSequence_off",1,17019,0,100,0,False
		queueB.Add "highscore_sequence_15a","DMD.highScoreSequence_on",1,17069,0,100,0,False
		queueB.Add "highscore_sequence_15b","DMD.highScoreSequence_off",1,17142,0,100,0,False
		Dim i
		For i = 1 to 7
			queueB.Add "highscore_sequence_" & (16 + i) & "a","DMD.highScoreSequence_on",1,17069 + (146 * i),0,100,0,False
			queueB.Add "highscore_sequence_" & (16 + i) & "b","DMD.highScoreSequence_off",1,17142 + (146 * i),0,100,0,False
		Next
		queueB.Add "highscore_sequence_end_a","DMD.highScoreSequence_on : Music.PlayTrack ""sting_victory"", Null, Null, Null, Null, 0, 0, 0, 0",1,18280,0,100,0,False
		queueB.Add "highscore_sequence_end_b","DMD.highScoreSequence_off : startMode MODE_HIGH_SCORE_ENTRY",1,20280,0,100,0,False
	End Sub

	Public Sub highScoreSequence_on()
		'Lights
		StartLightSequence StandardSeq, Array(dwhite, white), 25, seqRandom, 25, 1
		lightsOn GI, Array(dwhite, white), 100

		'Flippers
		SolLFlipper True
		SolRFlipper True
		SolUFlipper True

		'Shaker
		ShakerMotorTimer.Enabled = True

		'DMD
		DMD.jpFlush
		DMD.jpLocked = True
		DMD.jpDMD DMD.jpCL(FormatScore(currentPlayer("score"))), DMD.jpCL(FormatScore(currentPlayer("score"))), "", eScrollRight, eScrollLeft, eNone, 10000, False, ""
	End Sub

	Public Sub highScoreSequence_off()
		'Lights
		StopLightSequence StandardSeq
		lightsOff GI

		'Flippers
		SolLFlipper False
		SolRFlipper False
		SolUFlipper False

		'Shaker
		ShakerMotorTimer.Enabled = False

		'DMD
		DMD.jpFlush
		DMD.jpLocked = True
		DMD.jpDMD "", "", "", eScrollRight, eScrollLeft, eNone, 10000, False, ""
	End Sub

	Public Sub highScoreSequence_InsertName(doFlushing)
		Dim leaderboard
		Dim placeB
		'Determine placement
		If Not dataGame.data.Item("tournamentMode") = 0 Then
			leaderboard = 5
		Else
			leaderboard = dataGame.data.Item("gameDifficulty")
		End If

		If currentPlayer("score") > dataHighScores(leaderboard).data.Item("HS5") Then placeB = 5
		If currentPlayer("score") > dataHighScores(leaderboard).data.Item("HS4") Then placeB = 4
		If currentPlayer("score") > dataHighScores(leaderboard).data.Item("HS3") Then placeB = 3
		If currentPlayer("score") > dataHighScores(leaderboard).data.Item("HS2") Then placeB = 2
		If currentPlayer("score") > dataHighScores(leaderboard).data.Item("HS1") Then placeB = 1

		If doFlushing Then
			DMD.jpFlush
			DMD.jpLocked = True
		End If

		DMD.jpDMD placeB & ") " & currentPlayer("name") & MODE_COUNTERS_A, DMD.jpCL(FormatScore(currentPlayer("score"))), "", eNone, eNone, eNone, 250, False, ""
		DMD.jpDMD placeB & ") " & currentPlayer("name") & "_", DMD.jpCL(FormatScore(currentPlayer("score"))), "", eNone, eNone, eNone, 250, False, ""
	End Sub

	Public Sub highScoreSequence_toggleChar(direction)
		'Set up
		Dim hsChars
		hsChars = " <ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
		Dim charPos
		charPos = InStr(hsChars, MODE_COUNTERS_A)
		If charPos = 0 Then charPos = 1

		'Process toggle
		charPos = charPos + direction
		If charPos < 1 Then charPos = 38
		If charPos > 38 Then charPos = 1
		MODE_COUNTERS_A = Mid(hsChars, charPos, 1)

		'Update display
		highScoreSequence_InsertName True
	End Sub

	Public Sub highScoreSequence_selectChar()
		'Backspace
		If MODE_COUNTERS_A = "<" And Len(currentPlayer("name")) > 0 Then
			currentPlayerSet "name", Left(currentPlayer("name"), Len(currentPlayer("name")) - 1)
			highScoreSequence_InsertName True
			Exit Sub
		End If

		'Double-Space = done entering our name
		If MODE_COUNTERS_A = " " And Right(currentPlayer("name"), 1) = " " Then
			currentPlayerSet "name", Left(currentPlayer("name"), Len(currentPlayer("name")) - 1)
			startMode MODE_HIGH_SCORE_FINAL
			Exit Sub
		End If

		currentPlayerSet "name", currentPlayer("name") & MODE_COUNTERS_A

		'15 characters is limit; also done entering name
		If Len(currentPlayer("name")) >= 15 Then
			startMode MODE_HIGH_SCORE_FINAL
			Exit Sub
		End If

		highScoreSequence_InsertName True
	End Sub

	Public Sub highScoreSequence_final()
		DMD.triggerVideo VIDEO_HIGH_SCORE_FINAL

		Dim leaderboard
		Dim place
		Dim placeB
		
		'Determine placement
		If Not dataGame.data.Item("tournamentMode") = 0 Then
			leaderboard = 5
		Else
			leaderboard = dataGame.data.Item("gameDifficulty")
		End If

		If currentPlayer("score") > dataHighScores(leaderboard).data.Item("HS5") Then
			place = "5th place"
			placeB = 5
		End If
		If currentPlayer("score") > dataHighScores(leaderboard).data.Item("HS4") Then
			place = "4th place"
			placeB = 4
		End If
		If currentPlayer("score") > dataHighScores(leaderboard).data.Item("HS3") Then
			place = "3rd place"
			placeB = 3
		End If
		If currentPlayer("score") > dataHighScores(leaderboard).data.Item("HS2") Then
			place = "2nd place"
			placeB = 2
		End If
		If currentPlayer("score") > dataHighScores(leaderboard).data.Item("HS1") Then
			place = "1st place"
			placeB = 1
		End If

		'DMD
		DMD.jpFlush
		DMD.jpLocked = True
		DMD.jpDMD DMD.jpCL(currentPlayer("name")), DMD.jpCL("High score!"), "", eNone, eBlink, eNone, 3000, False, ""
		DMD.jpDMD DMD.jpCL(currentPlayer("name")), DMD.jpCL(FormatScore(currentPlayer("score"))), "", eNone, eBlinkFast, eNone, 5000, False, ""
		DMD.jpDMD DMD.jpCL(currentPlayer("name")), DMD.jpCL(place), "", eNone, eScrollLeft, eNone, 5000, True, ""

		'Actually save the high score
		dataHighScores(leaderboard).data.Item("HS" & placeB) = currentPlayer("score")
		dataHighScores(leaderboard).data.Item("HSD" & placeB) = FormatDateTime(Now(), vbShortDate)
		dataHighScores(leaderboard).data.Item("HSN" & placeB) = currentPlayer("name")
		dataHighScores(leaderboard).SaveAll

		'Call going to the next player
		queue.Add "next_player","DMD.startNextPlayer True",1,0,15000,100,0,False
	End Sub

	Public Sub deathSaveIntro()
		'Show Sequence on DMD 
		DMD.gameStatus "DEATH SAVE"
		'DMD.triggerVideo VIDEO_CLEAR
		DMD.triggerVideo VIDEO_DEATH_SAVE_INTRO
		PlayCallout "callout_nar_death_save_instr", 1, VOLUME_CALLOUTS
	End Sub

	Public Sub deathSave_strike() 'Score a strike towards death save
		Dim i

		'Note the failed heart beat
		MODE_VALUES = 0
		MODE_COUNTERS_B(1) = MODE_COUNTERS_B(1) + 1
		DMD.deathSave_progress MODE_COUNTERS_B(0), MODE_COUNTERS_B(1)

		'Annoy the player with a buzzer, shaker, and flashers
		PlaySound "sfx_loud_buzzer", 1, VOLUME_SFX
		ShakerMotorTimer.Enabled = True
		For i = 0 to 2
			flasherQueue.Add "death_save_strike_1_" & i, "Flash7 True : Flash8 True", 100, 0, 0, 100, 0, False
			flasherQueue.Add "death_save_strike_2_" & i, "Flash7 False : Flash8 False", 100, 0, 0, 100, 0, False
		Next
		queueB.Add "death_save_strike_2_" & i, "ShakerMotorTimer.Enabled = False", 100, 0, 500, 0, 0, False

		'You failed! Prepare the bonus sequence
		If MODE_COUNTERS_B(1) >= 3 Then
			deathSaveTimer.Enabled = False
			LightsOn coloredLights, Array(dred, red), 100
			Music.StopTrack 0
			PlaySound "sfx_third_strike", 1, VOLUME_SFX
			queue.Add "death_save_dead", "startMode MODE_END_OF_GAME_BONUS", 2, 0, 5000, 100, 0, False

			For i = 0 to 8
				flasherQueue.Add "death_save_failed_1_" & i, "Flash1 True", 100, 0, 0, 200, 0, False
				flasherQueue.Add "death_save_failed_2_" & i, "Flash1 False", 100, 0, 0, 200, 0, False
			Next
		End If
	End Sub

	Public Sub deathSave_CPR() 'Score a successful CPR
		Dim i

		MODE_VALUES = 1
		MODE_COUNTERS_B(0) = MODE_COUNTERS_B(0) - 1

		PlaySound "sfx_ecg", 1, VOLUME_SFX

		'Show Progress
		DMD.deathSave_progress MODE_COUNTERS_B(0), MODE_COUNTERS_B(1)

		'Did we succeed the death save?
		If MODE_COUNTERS_B(0) <= 0 Then
			deathSaveTimer.Enabled = False
			LightsOn coloredLights, Array(dgreen, green), 100
			Music.StopTrack 0

			currentPlayerSet "HP", 1	'Make sure the player now has exactly 1 HP
			DMD.sUpdateScores False

			For i = 0 to 8
				flasherQueue.Add "death_save_success_1_" & i, "Flash5 True", 100, 0, 0, 200, 0, False
				flasherQueue.Add "death_save_success_2_" & i, "Flash5 False", 100, 0, 0, 200, 0, False
			Next

			PlayCallout "callout_nar_death_save_success", 1, VOLUME_CALLOUTS 'TODO: re-record with harmony as the voice

			If currentPlayer("deathSaves") + ((dataGame.data.Item("gameDifficulty") - 2) * 2) >= 5 Then 'Was that the last death save we are allowed?
				queue.Add "death_saves_none_left", "DMD.triggerVideo VIDEO_DEATH_SAVE_PROGRESS : PlayCallout ""callout_nar_death_saves_none_left"", 1, VOLUME_CALLOUTS", 2, 0, 5000, 2000, 0, False 'TODO: re-record with harmony as the voice
			End If

			queue.Add "death_save_success", "startMode MODE_END_OF_TURN", 2, 0, 5000, 100, 0, False
		End If
	End Sub

	Public Sub deathSave_countdown(cNum)
		If cNum = 0 Then cNum = "GO!"

		If IsNull(cNum) Then 
			If usePUP = True And PUPStatus = True Then PUP.LabelSet pDMD,"center_bottom","",0,""
		Else
			If usePUP = True And PUPStatus = True Then PUP.LabelSet pDMD,"center_bottom",cNum,1,""
			DMD.jpFlush
			DMD.jpDMD DMD.jpCL(cNum), DMD.jpCL(cNum), "d_E172_0", eNone, eNone, eNone, 2000, True, ""
		End If

	End Sub

	Public Sub deathSave_progress(resLeft, strikes)
		If usePUP = True And PUPStatus = True Then
			PUP.LabelSet pDMD,"death_res",resLeft,1,""
			PUP.LabelSet pDMD,"death_strikes",strikes,1,""
		End If

		DMD.jpFlush
		DMD.jpLocked = True
		DMD.jpDMD DMD.jpFL("CPRs left: ", resLeft), DMD.jpFL("Strikes: ", strikes), "d_E173_0", eNone, eNone, eNone, 2000, True, ""
	End Sub

	'*************************
	' BONUS X / GLITCH
	'*************************

	Public Sub glitchProgression(stage)
		Dim i
		PlaySound "sfx_glitch_" & stage, 1, VOLUME_SFX
		PlayCallout "callout_game_glitch_" & stage, 1, VOLUME_CALLOUTS

		Select Case stage
			Case 1: 'Shaker glitch
				ShakerMotorTimer.Enabled = True
				queueB.Add "glitch_1_end", "ShakerMotorTimer.Enabled = False", 100, 0, 3000, 0, 0, False

			Case 2: 'Flasher glitch
				For i = 0 to 14
					flasherQueue.Add "glitch_2_on_" & i, "Flash" & int(rnd*8)+1 & " True", 1000, 0, 0, 100, 0, False
					flasherQueue.Add "glitch_2_off_" & i, "Flash" & int(rnd*8)+1 & " False", 1000, 0, 0, 100, 0, False
				Next
				flasherQueue.Add "glitch_2_end", "Flash1 False : Flash2 False : Flash3 False : Flash4 False : Flash5 False : Flash6 False : Flash7 False : Flash8 False", 1000, 0, 0, 200, 0, False

			Case 3: 'Glitch DMD text and videos
				DMD.triggerVideo VIDEO_BONUSX_GLITCH_3

				If usePUP = True And PUPStatus = True Then
					queue.Add "glitch_3_1", "DMD.centerText ""11111111111111111111"", ""11111111111111111111""", 1000, 0, 500, 0, 0, False
					queue.Add "glitch_3_2", "DMD.centerText ""22222222222222222222"", ""22222222222222222222""", 1000, 0, 500, 0, 0, False
					queue.Add "glitch_3_3", "DMD.centerText ""33333333333333333333"", ""33333333333333333333""", 1000, 0, 500, 0, 0, False
					queue.Add "glitch_3_4", "DMD.centerText ""44444444444444444444"", ""44444444444444444444""", 1000, 0, 500, 0, 0, False
					queue.Add "glitch_3_A", "DMD.centerText ""AAAAAAAAAAAAAAAAAAAA"", ""AAAAAAAAAAAAAAAAAAAA""", 1000, 0, 500, 0, 0, False
					queue.Add "glitch_3_B", "DMD.centerText ""BBBBBBBBBBBBBBBBBBBB"", ""BBBBBBBBBBBBBBBBBBBB""", 1000, 0, 500, 0, 0, False
					queue.Add "glitch_3_C", "DMD.centerText ""CCCCCCCCCCCCCCCCCCCC"", ""CCCCCCCCCCCCCCCCCCCC""", 1000, 0, 500, 0, 0, False
					queue.Add "glitch_3_D", "DMD.centerText ""DDDDDDDDDDDDDDDDDDDD"", ""DDDDDDDDDDDDDDDDDDDD""", 1000, 0, 500, 0, 0, False
					queue.Add "glitch_3_end", "DMD.centerTextRemove", 1000, 0, 500, 0, 0, False
				Else
					queue.Add "glitch_3_end", "", 1000, 0, 4000, 0, 0, False 'Pause queue for the triggered video
				End If

			Case 4: 'Glitch lights and segment displays
				segmentTimer.Enabled = False
				StartLightSequence StandardSeq, Array(ddarkgreen,darkgreen), 25, SeqRandom, 25, 1
				For i = 0 to 12
					queueB.Add "glitch_4_seg_" & i, "prepareSegments Array(YellowSegmentA, YellowSegmentB, YellowSegmentC), int(rnd*1000), False, Null : prepareSegments Array(RedSegmentA, RedSegmentB, RedSegmentC), int(rnd*1000), False, Null : prepareSegments Array(BonusXSegments), int(rnd*10), False, Null", 1000, 0, 0, 250, 0, False
				Next
				queueB.Add "glitch_4_end", "segmentTimer.Enabled = True : StopLightSequence StandardSeq", 999, 0, 0, 0, 0, False
				

			Case 5: 'Blackout glitch
				DMD.triggerVideo VIDEO_CLEAR
				DMD.triggerOverlay OVERLAY_REMOVE

				DMD.jpFlush
				DMD.jpLocked = True
				DMD.jpDMD "", "", "", eNone, eNone, eNone, 3000, True, ""

				StartLightSequence LightSeqAttract, Array(ddarkgreen,darkgreen), 25, SeqAllOff, 25, 1
				queueB.Add "glitch_5_end", "StopLightSequence LightSeqAttract : DMD.sShowScores", 1000, 0, 3000, 0, 0, False

			Case 6: 'Ball search glitch
				StartLightSequence StandardSeq, Array(ddarkgreen,darkgreen), 25, SeqAllOn, 25, 1
				For i = 0 to 14
					flasherQueue.Add "glitch_6_on_" & i, "Flash5 True", 1000, 0, 0, 100, 0, False
					flasherQueue.Add "glitch_6_off_" & i, "Flash5 False", 1000, 0, 0, 100, 0, False
				Next

				queue.Add "ball_search_glitch_2","DMD.ballSearchStep 14",1,0,0,200,0,False
				
				queue.Add "ball_search_glitch_3","DMD.ballSearchStep 9",1,0,0,200,0,False
				queue.Add "ball_search_glitch_4","DMD.ballSearchStep 10",1,0,0,200,0,False
				queue.Add "ball_search_glitch_5","DMD.ballSearchStep 11",1,0,0,200,0,False
				
				queue.Add "ball_search_glitch_6","DMD.ballSearchStep 15",1,0,0,200,0,False
				queue.Add "ball_search_glitch_7","DMD.ballSearchStep 16",1,0,0,200,0,False
				queue.Add "ball_search_glitch_8","DMD.ballSearchStep 17",1,0,0,200,0,False
				
				queue.Add "ball_search_glitch_9","DMD.ballSearchStep 18",1,0,0,200,0,False
				queue.Add "ball_search_glitch_10","DMD.ballSearchStep 19",1,0,0,200,0,False
				
				queue.Add "ball_search_glitch_11","DMD.ballSearchStep 20",1,0,0,200,0,False
				queue.Add "ball_search_glitch_12","DMD.ballSearchStep 21",1,0,0,200,0,False
				queue.Add "ball_search_glitch_13","DMD.ballSearchStep 22",1,0,0,200,0,False
				queue.Add "ball_search_glitch_14","DMD.ballSearchStep 23",1,0,0,200,0,False
				queue.Add "ball_search_glitch_15","StopLightSequence StandardSeq",1,0,0,200,0,False

		End Select
	End Sub

	Public Sub wizardReady_glitch()
		'TODO
	End Sub

	Public Sub glitchWizardJackpot(jackpotAwarded)
		DMD.triggerVideo VIDEO_GLITCH_WIZARD_JACKPOT
		If usePUP = True And PUPStatus = True Then
			PUP.LabelSet pDMD,"center_bottom",FormatScore(jackpotAwarded),1,""
		End If

		DMD.jpFlush
		DMD.jpLocked = True
		DMD.jpDMD DMD.jpCL("Glitch Jackpot!"), DMD.jpCL(FormatScore(jackpotAwarded)), "", eNone, eBlinkFast, eNone, 5000, True, ""
	End Sub

	'*************************
	' CASTLE
	'*************************

	'*************************
	' BLACKSMITH
	'*************************

	Public Sub wizardReady_blacksmith()
		PlayCallout "callout_blacksmith_wizard_ready", 1, VOLUME_CALLOUTS
		'TODO
	End Sub

	Public Sub wizardIntro_blacksmithKill()
		'Show score overlay again
		sShowScores

		'Intro video / sequence
		triggerVideo VIDEO_BLACKSMITH_WIZARD_KILL_INTRO

		'SFX
		Music.StopTrack 500
		PlaySound "sfx_blacksmith_wizard_kill_start", 1, VOLUME_SFX

		'GI / lighting
		LightsOff coloredLights
		lampC.AddLightSeq "attract", lSeqblacksmithWizardKillIntroA

		'Callout
		queue.Add "wizardIntro_blacksmithKill_callout", "PlayCallout ""callout_nar_instr_blacksmith_battle"", 1, VOLUME_CALLOUTS", 1, 0, 2000, 14000, 0, False

		'Start mode
		queue.Add "wizardIntro_blacksmithKill_start", "startMode MODE_BLACKSMITH_WIZARD_KILL", 1, 0, 0, 100, 0, False

		'Shaker motor + thunder clap + music
		queueB.Add "wizardIntro_blacksmithKill_shaker_on", "ShakerMotorTimer.Enabled = True : DMD.wizardIntro_blacksmithKill_ThunderClap : Music.PlayTrack ""music_blacksmith_kill"", Null, Null, Null, Null, 0, 0, 0, 0", 1, 0, 3000, 300, 0, False
		queueB.Add "wizardIntro_blacksmithKill_shaker_off", "ShakerMotorTimer.Enabled = False", 1, 2000, 0, 100, 0, False

		'Instructions
		queueB.Add "wizardIntro_blacksmithKill_instr_1", "LampC.LightColor RightRampShotL, Array(dpurple, purple) : LampC.Blink RightRampShotL : LampC.UpdateBlinkInterval RightRampShotL, 250", 1, 5500, 0, 100, 0, False
		queueB.Add "wizardIntro_blacksmithKill_instr_2", "LampC.LightOff RightRampShotL : LightsBlink blacksmithLights, Array(dpurple, purple), 250, 100", 1, 8500, 0, 100, 0, False
		queueB.Add "wizardIntro_blacksmithKill_instr_3", "LightsOff blacksmithLights", 1, 11750, 0, 100, 0, False
	End Sub

	Public Sub wizardIntro_blacksmithKill_ThunderClap() 'Thunder clap effect; we're trying to save on line counts in Light State Controller
		Dim i
		Dim j
		Dim k
		Dim l

		LightsOn coloredLights, Array(dwhite, white), 100
		Flash7 True
		Flash8 True

		queueB.Add "wizardIntro_blacksmithKill_flashers_off", "Flash7 False : Flash8 False", 1, 0, 0, 200, 0, False

		For i = 1 to 100 step 3
			l = 100 - i
			queueB.Add "wizardIntro_blacksmithKill_thunder_" & i, "LightsOn coloredLights, Array(dwhite, white), " & l, 1, 0, 0, 30, 0, False
		Next
	End Sub

	Public Sub blacksmithLore(stage) 'Show/handle blacksmith lore dialog
		Select Case stage
			'*** Intro ***
			Case 1
				'DMD
				DMD.jpLocked = False
				DMD.jpFlush
				DMD.jpLocked = True
				DMD.jpDMD DMD.jpCL("I see you in my"), DMD.jpCL("shop all the time."), "d_E187_0", eScrollLeft, eScrollLeft, eNone, 3240, False, ""
				DMD.jpDMD DMD.jpCL("It's time to tell"), DMD.jpCL("you a little story."), "d_E187_0", eBlink, eBlink, eNone, 3230, True, ""

				'Callout
				PlayCallout "callout_blacksmith_lore_1", 1, VOLUME_CALLOUTS

			Case 2
				'Start music
				Music.PlayTrack "music_blacksmith_lore", 1, Null, Null, Null, 0, 0, 0, 0

				'Light Sequence
				lampC.AddLightSeq "accomplishment", lSeqblacksmithMiniWizard

			Case 3
				PlayCallout "callout_blacksmith_lore_2", 1, VOLUME_CALLOUTS

				DMD.jpLocked = False
				DMD.jpFlush
				DMD.jpLocked = True
				DMD.jpDMD DMD.jpCL("A long time ago,"), DMD.jpCL("I used to be"), "d_E187_0", eNone, eNone, eNone, 1763, False, ""
				DMD.jpDMD DMD.jpCL("just"), DMD.jpCL("like you."), "d_E187_0", eBlinkFast, eBlinkFast, eNone, 1857, True, ""

			Case 4
				DMD.jpLocked = False
				DMD.jpFlush
				DMD.jpLocked = True
				DMD.jpDMD DMD.jpCL("I was an adventurer;"), DMD.jpCL("a fighter."), "d_E187_0", eNone, eBlinkFast, eNone, 2398, False, ""
				DMD.jpDMD DMD.jpCL("I wanted to defeat"), DMD.jpCL("all that was evil"), "d_E187_0", eNone, eNone, eNone, 2257, False, ""
				DMD.jpDMD DMD.jpCL("and be the hero"), DMD.jpCL("in my town."), "d_E187_0", eNone, eNone, eNone, 3000, True, ""

			Case 5
				DMD.jpLocked = False
				DMD.jpFlush
				DMD.jpLocked = True
				DMD.jpDMD DMD.jpCL("but little did I"), DMD.jpCL("know that my"), "d_E187_0", eNone, eNone, eNone, 1645, False, ""
				DMD.jpDMD DMD.jpCL("aspirations would"), DMD.jpCL("be my downfall."), "d_E187_0", eNone, eBlink, eNone, 5383, True, ""

			Case 6
				DMD.jpLocked = False
				DMD.jpFlush
				DMD.jpLocked = True
				DMD.jpDMD DMD.jpCL("It's not always"), DMD.jpCL("black and white"), "d_E187_0", eNone, eNone, eNone, 3549, False, ""
				DMD.jpDMD DMD.jpCL("who is truly your"), DMD.jpCL("ally"), "d_E187_0", eNone, eBlinkFast, eNone, 2022, False, ""
				DMD.jpDMD DMD.jpCL("or who is actually"), DMD.jpCL("your foe."), "d_E187_0", eNone, eBlink, eNone, 4444, True, ""

			Case 7
				Music.StopTrack 8360

				DMD.jpLocked = False
				DMD.jpFlush
				DMD.jpLocked = True
				DMD.jpDMD DMD.jpCL("Pay close attention"), DMD.jpCL("to your intuition"), "d_E187_0", eNone, eNone, eNone, 3596, False, ""
				DMD.jpDMD DMD.jpCL("and don't make"), DMD.jpCL("allies with the"), "d_E187_0", eNone, eNone, eNone, 1739, False, ""
				DMD.jpDMD DMD.jpCL("wrong"), DMD.jpCL("people."), "d_E187_0", eBlinkFast, eBlinkFast, eNone, 3025, True, ""

			'*** Kill Begin ***
			Case 8:
				'Callout
				PlayCallout "callout_blacksmith_kill", 1, VOLUME_CALLOUTS

				'Music
				Music.PlayTrack "music_tension", Null, Null, Null, Null, 0, 0, 0, 0

				DMD.jpLocked = False
				DMD.jpFlush
				DMD.jpLocked = True
				DMD.jpDMD DMD.jpCL("Well, to say I am"), DMD.jpCL("disappointed"), "d_E187_0", eNone, eBlinkFast, eNone, 1928, False, ""
				DMD.jpDMD DMD.jpCL("is an"), DMD.jpCL("understatement."), "d_E187_0", eNone, eBlink, eNone, 2300, True, ""

			Case 9:
				DMD.jpLocked = False
				DMD.jpFlush
				DMD.jpLocked = True
				DMD.jpDMD DMD.jpCL("I've supported you"), DMD.jpCL("through all your"), "d_E187_0", eNone, eNone, eNone, 1763, False, ""
				DMD.jpDMD DMD.jpCL("adventures"), DMD.jpCL("adventures"), "d_E187_0", eNone, eNone, eNone, 750, False, ""
				DMD.jpDMD DMD.jpCL("and *THIS* is how"), DMD.jpCL("you repay me?"), "d_E187_0", eNone, eNone, eNone, 3424, True, ""

			Case 10:
				DMD.jpLocked = False
				DMD.jpFlush
				DMD.jpLocked = True
				DMD.jpDMD DMD.jpCL("Very"), DMD.jpCL("well,"), "d_E187_0", eNone, eNone, eNone, 1248, False, ""
				DMD.jpDMD DMD.jpCL("but you have no"), DMD.jpCL("idea whom you are"), "d_E187_0", eNone, eNone, eNone, 1856, False, ""
				DMD.jpDMD DMD.jpCL("fighting"), DMD.jpCL("fighting"), "d_E187_0", eBlinkFast, eBlinkFast, eNone, 1300, True, ""

			Case 11:
				DMD.jpLocked = False
				DMD.jpFlush
				DMD.jpLocked = True
				DMD.jpDMD DMD.jpCL("I will take back"), DMD.jpCL("from you what"), "d_E187_0", eNone, eNone, eNone, 1443, False, ""
				DMD.jpDMD DMD.jpCL("rightfully belongs"), DMD.jpCL("to me for this!"), "d_E187_0", eNone, eNone, eNone, 2400, True, ""

			'*** Kill fail ***
			case 12:
				PlayCallout "callout_blacksmith_kill_fail", 1, VOLUME_CALLOUTS

				DMD.jpLocked = False
				DMD.jpFlush
				DMD.jpLocked = True
				DMD.jpDMD DMD.jpCL("Now that I have"), DMD.jpCL("taken everything"), "d_E187_0", eNone, eNone, eNone, 3446, False, ""
				DMD.jpDMD DMD.jpCL("back what no"), DMD.jpCL("longer belongs"), "d_E187_0", eNone, eNone, eNone, 1697, False, ""
				DMD.jpDMD DMD.jpCL("to"), DMD.jpCL("you,"), "d_E187_0", eNone, eNone, eNone, 1032, False, ""
				DMD.jpDMD DMD.jpCL("I will no longer"), DMD.jpCL("entertain myself"), "d_E187_0", eNone, eNone, eNone, 2151, False, ""
				DMD.jpDMD DMD.jpCL("with this"), DMD.jpCL("battle."), "d_E187_0", eNone, eBlink, eNone, 1609, False, ""
				DMD.jpDMD DMD.jpCL("So long,"), DMD.jpCL("adventurers!"), "d_E187_0", eBlinkFast, eBlinkFast, eNone, 1679, False, ""
				DMD.jpDMD DMD.jpCL("And"), DMD.jpCL("good luck..."), "d_E187_0", eNone, eNone, eNone, 1399, False, ""
				DMD.jpDMD DMD.jpCL("because you're"), DMD.jpCL("gonna need it!"), "d_E187_0", eNone, eBlink, eNone, 1714, True, ""

			'*** Kill success ***
			case 13:
				DMD.triggerVideo VIDEO_BLACKSMITH_WIZARD_KILL_SUCCESS_2
				PlayCallout "callout_blacksmith_death", 1, VOLUME_CALLOUTS

				DMD.jpLocked = False
				DMD.jpFlush
				DMD.jpLocked = True
				DMD.jpDMD DMD.jpCL("So this is"), DMD.jpCL("how it ends..."), "d_E1104_0", eScrollLeft, eScrolLRight, eNone, 3163, False, ""
				DMD.jpDMD DMD.jpCL("The elder, wiser,"), DMD.jpCL("stronger one"), "d_E1104_0", eNone, eNone, eNone, 2562, False, ""
				DMD.jpDMD DMD.jpCL("still fails to"), DMD.jpCL("the inexperienced."), "d_E1104_0", eNone, eBlink, eNone, 5110, False, ""
				DMD.jpDMD DMD.jpCL("What a"), DMD.jpCL("cruel world!"), "d_E1104_0", eBlinkFast, eBlinkFast, eNone, 4423, False, ""
				DMD.jpDMD DMD.jpCL("You have a"), DMD.jpCL("lot to learn,"), "d_E1104_0", eNone, eNone, eNone, 2662, False, ""
				DMD.jpDMD DMD.jpCL("and I shall not"), DMD.jpCL("be there for you"), "d_E1104_0", eNone, eNone, eNone, 2319, False, ""
				DMD.jpDMD DMD.jpCL("any"), DMD.jpCL("longer..."), "d_E1104_0", eBlink, eBlink, eNone, 2805, False, ""
				DMD.jpDMD "", "", "d_E1104_0", eNone, eNone, eNone, 2000, True, ""

			'*** Spare ***
			Case 14:
				PlayCallout "callout_blacksmith_spare", 1, VOLUME_CALLOUTS

				DMD.jpLocked = False
				DMD.jpFlush
				DMD.jpLocked = True
				DMD.jpDMD DMD.jpCL("I am honored"), DMD.jpCL("that you"), "d_E121_0", eScrollLeft, eScrolLRight, eNone, 1312, False, ""
				DMD.jpDMD DMD.jpCL("followed your"), DMD.jpCL("intuition."), "d_E121_0", eNone, eBlink, eNone, 2463, False, ""
				DMD.jpDMD DMD.jpCL("I've been by your"), DMD.jpCL("side since the"), "d_E121_0", eNone, eNone, eNone, 1828, False, ""
				DMD.jpDMD DMD.jpCL("very"), DMD.jpCL("beginning,"), "d_E121_0", eBlinkFast, eBlinkFast, eNone, 1194, False, ""
				DMD.jpDMD DMD.jpCL("and I will"), DMD.jpCL("continue to"), "d_E121_0", eNone, eNone, eNone, 1119, False, ""
				DMD.jpDMD DMD.jpCL("be"), DMD.jpCL("so."), "d_E121_0", eNone, eNone, eNone, 1462, False, ""
				DMD.jpDMD DMD.jpCL("Here is"), DMD.jpCL("my gift:"), "d_E121_0", eNone, eNone, eNone, 1406, True, ""
		End Select
	End Sub

	Public Sub blacksmithWizardKillDamage(jackpotAwarded)
		PlayCallout "callout_blacksmith_hit_" & int(rnd * 3), 1, VOLUME_CALLOUTS

		DMD.triggerVideo VIDEO_BLACKSMITH_WIZARD_KILL_DAMAGE
		If usePUP = True And PUPStatus = True Then
			PUP.LabelSet pDMD,"center_top","Blacksmith hurt!",1,""
			PUP.LabelSet pDMD,"center_bottom",FormatScore(jackpotAwarded),1,""
		End If

		DMD.jpFlush
		DMD.jpLocked = True
		DMD.jpDMD DMD.jpCL("Blacksmith hurt!"), DMD.jpCL(FormatScore(jackpotAwarded)), "d_E195_0", eNone, eBlinkFast, eNone, 5000, True, ""
	End Sub

	Public Sub blacksmithWizardKillLostAC()
		PlayCallout "callout_blacksmith_drain_" & int(rnd * 4), 1, VOLUME_CALLOUTS

		DMD.triggerVideo VIDEO_BLACKSMITH_WIZARD_KILL_LOSE_AC
		If usePUP = True And PUPStatus = True Then
			PUP.LabelSet pDMD,"center_top","Armor Stolen!",1,""
			PUP.LabelSet pDMD,"center_bottom","AC: " & currentPlayer("AC"),1,""
		End If

		DMD.jpFlush
		DMD.jpLocked = True
		DMD.jpDMD DMD.jpCL("Armor Stolen!"), DMD.jpCL(Chr(130) & currentPlayer("AC")), "d_E196_0", eNone, eBlinkFast, eNone, 5000, True, ""
	End Sub

	Public Sub wizardIntro_blacksmithSpare()
		'Show score overlay again
		sShowScores

		'Intro video / sequence
		triggerVideo VIDEO_BLACKSMITH_WIZARD_SPARE_INTRO

		'Music
		Music.PlayTrack "music_blacksmith_spare", Null, Null, Null, Null, 0, 0, 0, 0
		

		'GI / lighting
		LightsOff coloredLights
		lampC.AddLightSeq "attract", lSeqblacksmith_spare

		'Callout
		queue.Add "wizardIntro_blacksmithSpare_callout", "PlayCallout ""callout_nar_blacksmith_raid_instr"", 1, VOLUME_CALLOUTS", 1, 0, 0, 9000, 0, False

		'Start mode after callout
		queue.Add "wizardIntro_blacksmithSpare_start", "startMode MODE_BLACKSMITH_WIZARD_SPARE", 1, 0, 0, 100, 0, False

		'Instructions
		queueB.Add "wizardIntro_blacksmithSpare_instr_1", "LightsBlink blacksmithLights, Array(dpurple, purple), 250, 100", 1, 2800, 0, 100, 0, False
		queueB.Add "wizardIntro_blacksmithSpare_instr_2", "LightsOff blacksmithLights : LampC.LightColor HoleShotL, Array(dpurple, purple) : LampC.Blink HoleShotL : LampC.UpdateBlinkInterval HoleShotL, 250", 1, 5680, 0, 100, 0, False
		queueB.Add "wizardIntro_blacksmithSpare_instr_3", "LampC.LightOff HoleShotL", 1, 9000, 0, 100, 0, False
	End Sub

	Public Sub blacksmithWizardSpareJackpot(jackpotAwarded)
		DMD.triggerVideo VIDEO_BLACKSMITH_WIZARD_SPARE_JACKPOT
		If usePUP = True And PUPStatus = True Then
			PUP.LabelSet pDMD,"center_bottom",FormatScore(jackpotAwarded),1,""
		End If

		PlayCallout "callout_blacksmith_jackpot", 1, VOLUME_CALLOUTS

		DMD.jpFlush
		DMD.jpLocked = True
		DMD.jpDMD DMD.jpCL("Raid Jackpot!"), DMD.jpCL(FormatScore(jackpotAwarded)), "", eNone, eBlinkFast, eNone, 3000, True, ""
	End Sub

	Public Sub blacksmithWizardSpareSuperJackpot(jackpotAwarded)
		DMD.triggerVideo VIDEO_BLACKSMITH_WIZARD_SPARE_SUPER_JACKPOT
		If usePUP = True And PUPStatus = True Then
			PUP.LabelSet pDMD,"center_bottom",FormatScore(jackpotAwarded),1,""
		End If

		PlayCallout "callout_blacksmith_super_jackpot", 1, VOLUME_CALLOUTS

		DMD.jpFlush
		DMD.jpLocked = True
		DMD.jpDMD DMD.jpCL("Super Jackpot!"), DMD.jpCL(FormatScore(jackpotAwarded)), "", eNone, eBlinkFast, eNone, 5000, True, ""
	End Sub
End Class

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

'******************************************************
'  ZFLP: FLIPPER CORRECTION FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
	Dim a
	a = Array(LF, RF)
	Dim x
	For Each x In a
		x.addpoint aStr, idx, aX, aY
	Next
End Sub

Class FlipperPolarity
	Public DebugOn, Enabled
	Private FlipAt        'Timer variable (IE 'flip at 723,530ms...)
	Public TimeDelay        'delay before trigger turns off and polarity is disabled TODO set time!
	Private Flipper, FlipperStart,FlipperEnd, FlipperEndY, LR, PartialFlipCoef
	Private Balls(20), balldata(20)
	
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
			Set Balldata(x) = New SpoofBall
		Next
	End Sub
	
	Public Property Let Object(aInput)
		Set Flipper = aInput
		StartPoint = Flipper.x
	End Property
	Public Property Let StartPoint(aInput)
		If IsObject(aInput) Then
			FlipperStart = aInput.x
		Else
			FlipperStart = aInput
		End If
	End Property
	Public Property Get StartPoint
		StartPoint = FlipperStart
	End Property
	Public Property Let EndPoint(aInput)
		FlipperEnd = aInput.x
		FlipperEndY = aInput.y
	End Property
	Public Property Get EndPoint
		EndPoint = FlipperEnd
	End Property
	Public Property Get EndPointY
		EndPointY = FlipperEndY
	End Property
	
	Public Sub AddPoint(aChooseArray, aIDX, aX, aY) 'Index #, X position, (In) y Position (out)
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
		If GameTime > 100 Then Report aChooseArray
	End Sub
	
	Public Sub Report(aChooseArray)         'debug, reports all coords in tbPL.text
		If Not DebugOn Then Exit Sub
		Dim a1, a2
		Select Case aChooseArray
			Case "Polarity"
				a1 = PolarityIn
				a2 = PolarityOut
			Case "Velocity"
				a1 = VelocityIn
				a2 = VelocityOut
			Case "Ycoef"
				a1 = YcoefIn
				a2 = YcoefOut
			Case Else
				tbpl.text = "wrong String"
				Exit Sub
		End Select
		Dim str, x
		For x = 0 To UBound(a1)
			str = str & aChooseArray & " x: " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
		Next
		tbpl.text = str
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
			If Not IsEmpty(balls(x) ) Then
				pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
			End If
		Next
	End Property
	
	Public Sub ProcessBalls() 'save data of balls in flipper range
		FlipAt = GameTime
		Dim x
		For x = 0 To UBound(balls)
			If Not IsEmpty(balls(x) ) Then
				balldata(x).Data = balls(x)
			End If
		Next
		PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
		PartialFlipCoef = Abs(PartialFlipCoef - 1)
	End Sub
	
	Private Function FlipperOn() 'Timer shutoff for polaritycorrect
		If GameTime < FlipAt + TimeDelay Then FlipperOn = True
	End Function
	
	Public Sub PolarityCorrect(aBall)
		If FlipperOn() Then
			Dim tmp, BallPos, x, IDX, Ycoef
			Ycoef = 1
			
			'y safety Exit
			If aBall.VelY >  - 8 Then 'ball going down
				RemoveBall aBall
				Exit Sub
			End If
			
			'Find balldata. BallPos = % on Flipper
			For x = 0 To UBound(Balls)
				If aBall.id = BallData(x).id And Not IsEmpty(BallData(x).id) Then
					idx = x
					BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
					If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                                'find safety coefficient 'ycoef' data
				End If
			Next
			
			If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
				BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
				If ballpos > 0.65 Then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                                                'find safety coefficient 'ycoef' data
			End If
			
			'Velocity correction
			If Not IsEmpty(VelocityIn(0) ) Then
				Dim VelCoef
				VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)
				
				If partialflipcoef < 1 Then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)
				
				If Enabled Then aBall.Velx = aBall.Velx * VelCoef
				If Enabled Then aBall.Vely = aBall.Vely * VelCoef
			End If
			
			'Polarity Correction (optional now)
			If Not IsEmpty(PolarityIn(0) ) Then
				If StartPoint > EndPoint Then LR =  - 1        'Reverse polarity if left flipper
				Dim AddX
				AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR
				
				If Enabled Then aBall.VelX = aBall.VelX + 1 * (AddX * ycoef * PartialFlipcoef)
			End If
		End If
		RemoveBall aBall
	End Sub
End Class

'******************************************************
'  SLINGSHOT CORNER SPIN
'******************************************************

if bSlingSpin Then
	LSlingSpin.enabled = true
	RSlingSpin.enabled = true
end if

const MaxSlingSpin = 350
const MinCollVelX = 7
const MinCorVelocity = 10

dim lspinball
sub LSlingSpin_hit
'	debug.print "lhit: " & activeball.angmomz & " velx: " & activeball.velx & " vely: " & activeball.vely & " corvel: " & cor.ballvel(activeball.id)
	if activeball.velx < -MinCollVelX And cor.ballvel(activeball.id) > MinCorVelocity then
'		debug.print "lspinball set, vel: " & activeball.angmomz
		lspinball = activeball.id
	end if

end sub

sub LSlingSpin_unhit
	if lspinball <> -1 then
		activeball.angmomz = (cor.ballvel(activeball.id)/2) * activeball.angmomz
		if activeball.angmomz  > MaxSlingSpin then activeball.angmomz  = MaxSlingSpin
		if activeball.angmomz  < -MaxSlingSpin then activeball.angmomz  = -MaxSlingSpin 
'		debug.print "lunhit: " & activeball.angmomz & " velx: " & activeball.velx & " vely: " & activeball.vely & ", vel: " & cor.ballvel(activeball.id)
		lspinball = -1
	end if
end sub


dim rspinball
sub RSlingSpin_hit
'	debug.print "rhit: " & activeball.angmomz & " velx: " & activeball.velx & " vely: " & activeball.vely & " corvel: " & cor.ballvel(activeball.id)
	if activeball.velx > MinCollVelX And cor.ballvel(activeball.id) > MinCorVelocity then
'		debug.print "rspinball set, vel: " & activeball.angmomz
		rspinball = activeball.id
	end if

end sub

sub RSlingSpin_unhit
	if rspinball <> -1 then
		activeball.angmomz = (cor.ballvel(activeball.id)/2) * activeball.angmomz
		if activeball.angmomz  > MaxSlingSpin then activeball.angmomz  = MaxSlingSpin
		if activeball.angmomz  < -MaxSlingSpin then activeball.angmomz  = -MaxSlingSpin  
'		debug.print "runhit: " & activeball.angmomz & " velx: " & activeball.velx & " vely: " & activeball.vely & ", vel: " & cor.ballvel(activeball.id)
		rspinball = -1
	end if
end sub


'******************************************************
'  FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, ByVal offset) 'shuffle 1d array
	Dim x, aCount
	aCount = 0
	ReDim a(UBound(aArray) )
	For x = 0 To UBound(aArray)        'Shuffle objects in a temp array
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
	ReDim aArray(aCount - 1 + offset)        'Resize original array
	For x = 0 To aCount - 1                'set objects back into original array
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
	BallSpeed = Sqr(ball.VelX ^ 2 + ball.VelY ^ 2 + ball.VelZ ^ 2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)        'Set up line via two points, no clamping. Input X, output Y
	Dim x, y, b, m
	x = input
	m = (Y2 - Y1) / (X2 - X1)
	b = Y2 - m * X2
	Y = M * x + b
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
	Dim ii
	For ii = 1 To UBound(xKeyFrame)        'find active line
		If xInput <= xKeyFrame(ii) Then
			L = ii
			Exit For
		End If
	Next
	If xInput > xKeyFrame(UBound(xKeyFrame) ) Then L = UBound(xKeyFrame)        'catch line overrun
	Y = pSlope(xInput, xKeyFrame(L - 1), yLvl(L - 1), xKeyFrame(L), yLvl(L) )
	
	If xInput <= xKeyFrame(LBound(xKeyFrame) ) Then Y = yLvl(LBound(xKeyFrame) )         'Clamp lower
	If xInput >= xKeyFrame(UBound(xKeyFrame) ) Then Y = yLvl(UBound(xKeyFrame) )        'Clamp upper
	
	LinearEnvelope = Y
End Function

Sub InitPolarity()
	dim x, a : a = Array(LF, RF)
	for each x in a
			x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
			x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
			x.enabled = True
			x.TimeDelay = 60
	Next

	AddPt "Polarity", 0, 0, 0
	AddPt "Polarity", 1, 0.05, -5.5
	AddPt "Polarity", 2, 0.4, -5.5
	AddPt "Polarity", 3, 0.6, -5.0
	AddPt "Polarity", 4, 0.65, -4.5
	AddPt "Polarity", 5, 0.7, -4.0
	AddPt "Polarity", 6, 0.75, -3.5
	AddPt "Polarity", 7, 0.8, -3.0
	AddPt "Polarity", 8, 0.85, -2.5
	AddPt "Polarity", 9, 0.9,-2.0
	AddPt "Polarity", 10, 0.95, -1.5
	AddPt "Polarity", 11, 1, -1.0
	AddPt "Polarity", 12, 1.05, -0.5
	AddPt "Polarity", 13, 1.1, 0
	AddPt "Polarity", 14, 1.3, 0

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

'******************************************************
'  FLIPPER TRICKS
'******************************************************

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
	Dim b, BOT
	BOT = gBOT
	
	If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
		EOSNudge1 = 1
		'debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
		If Flipper2.currentangle = EndAngle2 Then
			For b = 0 To UBound(BOT)
				If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
					'Debug.Print "ball in flip1. exit"
					Exit Sub
				End If
			Next
			For b = 0 To UBound(BOT)
				If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
					BOT(b).velx = BOT(b).velx / 1.3
					BOT(b).vely = BOT(b).vely - 0.5
				End If
			Next
		End If
	Else
		If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 Then EOSNudge1 = 0
	End If
End Sub

'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

Function max(a,b)
	If a > b Then
		max = a
	Else
		max = b
	End If
End Function

Function Distance(ax,ay,bx,by)
	Distance = Sqr((ax - bx) ^ 2 + (ay - by) ^ 2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) ' Distance between a point and a line where point Is px,py
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

Const FlipperCoilRampupMode = 0       '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 0.8
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
Const EOSReturn = 0.025  'mid 90's and later

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
		Dim b, BOT
		BOT = gBOT
		
		For b = 0 To UBound(BOT)
			If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
				If BOT(b).vely >= - 0.4 Then BOT(b).vely =  - 0.4
			End If
		Next
	End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
	Dim Dir
	Dir = Flipper.startangle / Abs(Flipper.startangle)        '-1 for Right Flipper
	
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

Const LiveDistanceMin = 30  'minimum distance In vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114  'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
	Dim Dir
	Dir = Flipper.startangle / Abs(Flipper.startangle)    '-1 for Right Flipper
	Dim LiveCatchBounce                                                                                                                        'If live catch is not perfect, it won't freeze ball totally
	Dim CatchTime
	CatchTime = GameTime - FCount
	
	If CatchTime <= LiveCatch And parm > 6 And Abs(Flipper.x - ball.x) > LiveDistanceMin And Abs(Flipper.x - ball.x) < LiveDistanceMax Then
		If CatchTime <= LiveCatch * 0.5 Then                                                'Perfect catch only when catch time happens in the beginning of the window
			LiveCatchBounce = 0
		Else
			LiveCatchBounce = Abs((LiveCatch / 2) - CatchTime)        'Partial catch when catch happens a bit late
		End If
		
		If LiveCatchBounce = 0 And ball.velx * Dir > 0 Then ball.velx = 0
		ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
		ball.angmomx = 0
		ball.angmomy = 0
		ball.angmomz = 0
	Else
		If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf ActiveBall, parm
	End If
End Sub

'******************************************************
' ZTAR: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1         '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7     'Level of bounces. Recommmended value of 0.7

Sub TargetBouncer(aBall,defvalue)
	Dim zMultiplier, vel, vratio
	If TargetBouncerEnabled = 1 And aball.z < 30 Then
		'debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
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
		'debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
		'debug.print "conservation check: " & BallSpeed(aBall)/vel
	End If
End Sub

' Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit
	TargetBouncer ActiveBall, 1
End Sub

'******************************************************
'****  ZPHY: PHYSICS DAMPENERS
'******************************************************
'
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


Dim RubbersD
Set RubbersD = New Dampener        'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = True        'shows info in textbox "TBPout"
RubbersD.Print = True       'debug, reports In debugger (In vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967        'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64        'there's clamping so interpolate up to 56 at least

Dim SleevesD
Set SleevesD = New Dampener        'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = True        'shows info in textbox "TBPout"
SleevesD.Print = True        'debug, reports In debugger (In vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'#########################    Adjust these values to increase or lessen the elasticity

Dim FlippersD
Set FlippersD = New Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = DEBUG_LOG_FLIPPER_DAMPENER
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
	Public Print, debugOn 'tbpOut.text
	Public name, Threshold         'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
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
		If debugOn Then TBPout.text = str
	End Sub
	
	Public Sub Dampenf(aBall, parm) 'Rubberizer is handle here
		Dim RealCOR, DesiredCOR, str, coef
		DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
		RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
		coef = desiredcor / realcor
		If debugOn Then str = name & " In vel:" & Round(cor.ballvel(aBall.id),2 ) & vbNewLine & "desired cor: " & Round(desiredcor,4) & vbNewLine & _
		"actual cor: " & Round(realCOR,4) & vbNewLine & "ballspeed coef: " & Round(coef, 3) & vbNewLine
		If Print Then Debug.print Round(cor.ballvel(aBall.id),2) & ", " & Round(desiredcor,3)
		If Abs(aball.velx) < 2 And aball.vely < 0 And aball.vely >  - 3.75 Then
			aBall.velx = aBall.velx * coef
			aBall.vely = aBall.vely * coef
		End If
	End Sub
	
	Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
		Dim x
		For x = 0 To UBound(aObj.ModIn)
			addpoint x, aObj.ModIn(x), aObj.ModOut(x) * aCoef
		Next
	End Sub
	
	
	Public Sub Report()         'debug, reports all coords in tbPL.text
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
	
	Public Sub Update()    'tracks in-ball-velocity
		Dim str, b, AllBalls, highestID
		allBalls = GetBalls
		
		For Each b In allballs
			If b.id >= HighestID Then highestID = b.id
		Next
		
		If UBound(ballvel) < highestID Then ReDim ballvel(highestID)    'set bounds
		If UBound(ballvelx) < highestID Then ReDim ballvelx(highestID)    'set bounds
		If UBound(ballvely) < highestID Then ReDim ballvely(highestID)    'set bounds
		
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

'******************************************************
' 	ZRDT:  DROP TARGETS by Rothbauerw
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

'Define a variable for each drop target
Dim DT1, DT2, DT3

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

DT1 = Array(HP1a, HP1b, HP1, SWITCH_DROPS_PLUS, 0, False)
DT2 = Array(HP2a, HP2b, HP2, SWITCH_DROPS_H, 0, False)
DT3 = Array(HP3a, HP3b, HP3, SWITCH_DROPS_P, 0, False)

Dim DTArray
DTArray = Array(DT1, DT2, DT3)


Dim DTArray0, DTArray1, DTArray2, DTArray3, DTArray4, DTArray5
ReDim DTArray0(UBound(DTArray)), DTArray1(UBound(DTArray)), DTArray2(UBound(DTArray)), DTArray3(UBound(DTArray)), DTArray4(UBound(DTArray)), DTArray5(UBound(DTArray))

Dim DTIdx
For DTIdx = 0 To UBound(DTArray)
	Set DTArray0(DTIdx) = DTArray(DTIdx)(0)
	Set DTArray1(DTIdx) = DTArray(DTIdx)(1)
	Set DTArray2(DTIdx) = DTArray(DTIdx)(2)
	DTArray3(DTIdx) = DTArray(DTIdx)(3)
	DTArray4(DTIdx) = DTArray(DTIdx)(4)
	DTArray5(DTIdx) = DTArray(DTIdx)(5)
Next


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
Const DTMass = 0.2 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance

'******************************************************
'  DROP TARGETS FUNCTIONS
'******************************************************

Sub DTHit(switch)
	Dim i
	i = DTArrayID(switch)
	
	PlayTargetSound
	DTArray4(i) = DTCheckBrick(ActiveBall,DTArray2(i))
	If DTArray4(i) = 1 Or DTArray4(i) = 3 Or DTArray4(i) = 4 Then
		DTBallPhysics ActiveBall, DTArray2(i).rotz, DTMass
	End If
	DoDTAnim
End Sub

Sub DTRaise(switch)
	Dim i
	i = DTArrayID(switch)
	
	DTArray4(i) =  - 1
	DoDTAnim
End Sub

Sub DTDrop(switch)
	Dim i
	i = DTArrayID(switch)
	
	DTArray4(i) = 1
	DoDTAnim
End Sub

Function DTArrayID(switch)
	Dim i
	For i = 0 To UBound(DTArray)
		If DTArray3(i) = switch Then
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
		DTArray4(i) = DTAnimate(DTArray0(i),DTArray1(i),DTArray2(i),DTArray3(i),DTArray4(i))
	Next
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
		SoundDropTargetDrop prim
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
			DTArray5(ind) = True 'Mark target as dropped
			If UsingROM Then
				controller.Switch(Switchid) = 1
			Else
				DTAction switchid
			End If
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
		DTArray5(ind) = False 'Mark target as not dropped
		If UsingROM Then controller.Switch(Switchid) = 0
		RandomSoundDropTargetReset prim 'Custom added
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
	
	DTDropped = DTArray5(ind)
End Function

'******************************************************
'  DROP TARGET
'  SUPPORTING FUNCTIONS
'******************************************************

' Used for drop targets
Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy) 'Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order
	Dim AB, BC, CD, DA
	AB = (bx * py) - (by * px) - (ax * py) + (ay * px) + (ax * by) - (ay * bx)
	BC = (cx * py) - (cy * px) - (bx * py) + (by * px) + (bx * cy) - (by * cx)
	CD = (dx * py) - (dy * px) - (cx * py) + (cy * px) + (cx * dy) - (cy * dx)
	DA = (ax * py) - (ay * px) - (dx * py) + (dy * px) + (dx * ay) - (dy * ax)
	
	If (AB <= 0 And BC <= 0 And CD <= 0 And DA <= 0) Or (AB >= 0 And BC >= 0 And CD >= 0 And DA >= 0) Then
		InRect = True
	Else
		InRect = False
	End If
End Function

Function InRotRect(ballx,bally,px,py,angle,ax,ay,bx,by,cx,cy,dx,dy)
	Dim rax,ray,rbx,rby,rcx,rcy,rdx,rdy
	Dim rotxy
	rotxy = RotPoint(ax,ay,angle)
	rax = rotxy(0) + px
	ray = rotxy(1) + py
	rotxy = RotPoint(bx,by,angle)
	rbx = rotxy(0) + px
	rby = rotxy(1) + py
	rotxy = RotPoint(cx,cy,angle)
	rcx = rotxy(0) + px
	rcy = rotxy(1) + py
	rotxy = RotPoint(dx,dy,angle)
	rdx = rotxy(0) + px
	rdy = rotxy(1) + py
	
	InRotRect = InRect(ballx,bally,rax,ray,rbx,rby,rcx,rcy,rdx,rdy)
End Function

Function RotPoint(x,y,angle)
	Dim rx, ry
	rx = x * dCos(angle) - y * dSin(angle)
	ry = x * dSin(angle) + y * dCos(angle)
	RotPoint = Array(rx,ry)
End Function

'******************************************************
'****  END DROP TARGETS
'******************************************************


'******************************************************
'	ZSTD: STAND-UP TARGET INITIALIZATION
'******************************************************

'Define a variable for each stand-up target
Dim STa1, STa2, STa3, STa4, STa5
Dim STb1, Stb2, STb3

'Set array with stand-up target objects
'
'StandupTargetvar = Array(primary, prim, swtich)
'     primary:             vp target to determine target hit
'    prim:                primitive target used for visuals and animation
'                            IMPORTANT!!!
'                            transy must be used to offset the target animation
'    switch:                ROM switch number
'    animate:            Arrary slot for handling the animation instrucitons, set to 0
'
'You will also need to add a secondary hit object for each stand up (name sw11o, sw12o, and sw13o on the example Table1)
'these are inclined primitives to simulate hitting a bent target and should provide so z velocity on high speed impacts

' TRAIN standups
STa1 = Array(StandupT, StandupTP, SWITCH_STANDUP_T, 0)
STa2 = Array(StandupR, StandupRP, SWITCH_STANDUP_R, 0)
STa3 = Array(StandupA, StandupAP, SWITCH_STANDUP_A, 0)
STa4 = Array(StandupI, StandupIP, SWITCH_STANDUP_I, 0)
STa5 = Array(StandupN, StandupNP, SWITCH_STANDUP_N, 0)

' Crossbow Standups
STb1 = Array(StandupCL, StandupCLP, SWITCH_CROSSBOW_LEFT_STANDUP, 0)
STb2 = Array(StandupCM, StandupCMP, SWITCH_CROSSBOW_CENTER_STANDUP, 0)
STb3 = Array(StandupCR, StandupCRP, SWITCH_CROSSBOW_RIGHT_STANDUP, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(STa1, STa2, STa3, STa4, STa5, STb1, STb2, STb3)

Dim STArray0, STArray1, STArray2, STArray3
ReDim STArray0(UBound(STArray)), STArray1(UBound(STArray)), STArray2(UBound(STArray)), STArray3(UBound(STArray))

Dim STIdx
For STIdx = 0 To UBound(STArray)
	Set STArray0(STIdx) = STArray(STIdx)(0)
	Set STArray1(STIdx) = STArray(STIdx)(1)
	STArray2(STIdx) = STArray(STIdx)(2)
	STArray3(STIdx) = STArray(STIdx)(3)
Next

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
	STArray3(i) = STCheckHit(ActiveBall,STArray0(i))
	
	If STArray3(i) <> 0 Then
		DTBallPhysics ActiveBall, STArray0(i).orientation, STMass
	End If
	DoSTAnim
End Sub

Function STArrayID(switch)
	Dim i
	For i = 0 To UBound(STArray)
		If STArray2(i) = switch Then
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
		STArray3(i) = STAnimate(STArray0(i),STArray1(i),STArray2(i),STArray3(i))
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
		primary.uservalue = GameTime
	End If
	
	animtime = GameTime - primary.uservalue
	
	If animate = 1 Then
		primary.collidable = 0
		prim.transy =  - STMaxOffset
		If UsingROM Then
			vpmTimer.PulseSw switch
		Else
			doSwitch switch, True
		End If
		STAnimate = 2
		Exit Function
	ElseIf animate = 2 Then
		prim.transy = prim.transy + STAnimStep
		If prim.transy >= 0 Then
			prim.transy = 0
			primary.collidable = 1
			doSwitch switch, False
			STAnimate = 0
			Exit Function
		Else
			STAnimate = 2
		End If
	End If
End Function

'******************************************************
'****   END STAND-UP TARGETS
'******************************************************

'******************************************************
'****  ZBAL: BALL ROLLING AND DROP SOUNDS
'******************************************************
'
' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

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
	Dim b', BOT
	'    BOT = GetBalls
	
	' stop the sound of deleted balls
	For b = UBound(gBOT) + 1 To tnob - 1
		' Comment the next line if you are not implementing Dyanmic Ball Shadows
		If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
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
		
		' "Static" Ball Shadows
		' Comment the next If block, if you are not implementing the Dyanmic Ball Shadows
		If AmbientBallShadowOn = 0 Then
			If gBOT(b).Z > 30 Then
				BallShadowA(b).height = gBOT(b).z - BallSize / 4        'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
			Else
				BallShadowA(b).height = 0.1
			End If
			BallShadowA(b).Y = gBOT(b).Y + offsetY
			BallShadowA(b).X = gBOT(b).X + offsetX
			BallShadowA(b).visible = 1
		End If
	Next
End Sub


'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************

'******************************************************
'**** ZRMP: RAMP ROLLING SFX
'******************************************************

'Ball tracking ramp SFX 1.0
'   Reqirements:
'          * Import A Sound File for each ball on the table for plastic ramps.  Call It RampLoop<Ball_Number> ex: RampLoop1, RampLoop2, ...
'          * Import a Sound File for each ball on the table for wire ramps. Call it WireLoop<Ball_Number> ex: WireLoop1, WireLoop2, ...
'          * Create a Timer called RampRoll, that is enabled, with a interval of 100
'          * Set RampBAlls and RampType variable to Total Number of Balls
'    Usage:
'          * Setup hit events and call WireRampOn True or WireRampOn False (True = Plastic ramp, False = Wire Ramp)
'          * To stop tracking ball
'                 * call WireRampOff
'                 * Otherwise, the ball will auto remove if it's below 30 vp units
'

Dim RampMinLoops
RampMinLoops = 4

' RampBalls
'      Setup:        Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  if tnob = 5, then RammBalls(6,2)
'      Description:
Dim RampBalls(7,2)
'x,0 = ball x,1 = ID,    2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
Dim RampType(7)

Sub WireRampOn(input)
	Waddball ActiveBall, input
	RampRollUpdate
End Sub
Sub WireRampOff()
	WRemoveBall ActiveBall.ID
End Sub


' WaddBall (Active Ball, Boolean)
'     Description: This subroutine is called from WireRampOn to Add Balls to the RampBalls Array
Sub Waddball(input, RampInput)    'Add ball
	' This will loop through the RampBalls array checking each element of the array x, position 1
	' To see if the the ball was already added to the array.
	' If the ball is found then exit the subroutine
	Dim x
	For x = 1 To UBound(RampBalls)    'Check, don't add balls twice
		If RampBalls(x, 1) = input.id Then
			If Not IsEmpty(RampBalls(x,1) ) Then Exit Sub    'Frustating issue with BallId 0. Empty variable = 0
		End If
	Next
	
	' This will itterate through the RampBalls Array.
	' The first time it comes to a element in the array where the Ball Id (Slot 1) is empty.  It will add the current ball to the array
	' The RampBalls assigns the ActiveBall to element x,0 and ball id of ActiveBall to 0,1
	' The RampType(BallId) is set to RampInput
	' RampBalls in 0,0 is set to True, this will enable the timer and the timer is also turned on
	For x = 1 To UBound(RampBalls)
		If IsEmpty(RampBalls(x, 1)) Then
			Set RampBalls(x, 0) = input
			RampBalls(x, 1) = input.ID
			RampType(x) = RampInput
			RampBalls(x, 2) = 0
			'exit For
			RampBalls(0,0) = True
			RampRoll.Enabled = 1     'Turn on timer
			'RampRoll.Interval = RampRoll.Interval 'reset timer
			Exit Sub
		End If
		If x = UBound(RampBalls) Then     'debug
			If DEBUG_LOG_WIRE_RAMPS = True Then WriteToLog "Wire Ramps", "WireRampOn error, ball queue Is full: " & vbNewLine & _
			RampBalls(0, 0) & vbNewLine & _
			TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & "type:" & RampType(1) & vbNewLine & _
			TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & "type:" & RampType(2) & vbNewLine & _
			TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & "type:" & RampType(3) & vbNewLine & _
			TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & "type:" & RampType(4) & vbNewLine & _
			TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & "type:" & RampType(5) & vbNewLine & _
			" "
		End If
	Next
End Sub

' WRemoveBall (BallId)
'    Description: This subroutine is called from the RampRollUpdate subroutine
'                 and is used to remove and stop the ball rolling sounds
Sub WRemoveBall(ID)        'Remove ball
	'Debug.Print "In WRemoveBall() + Remove ball from loop array"
	Dim ballcount
	ballcount = 0
	Dim x
	For x = 1 To UBound(RampBalls)
		If ID = RampBalls(x, 1) Then 'remove ball
			Set RampBalls(x, 0) = Nothing
			RampBalls(x, 1) = Empty
			RampType(x) = Empty
			StopSound("RampLoop" & x)
			StopSound("wireloop" & x)
		End If
		'if RampBalls(x,1) = Not IsEmpty(Rampballs(x,1) then ballcount = ballcount + 1
		If Not IsEmpty(Rampballs(x,1)) Then ballcount = ballcount + 1
	Next
	If BallCount = 0 Then RampBalls(0,0) = False    'if no balls in queue, disable timer update
End Sub

Sub RampRoll_Timer()
	RampRollUpdate
End Sub

Sub RampRollUpdate()        'Timer update
	Dim x
	For x = 1 To UBound(RampBalls)
		If Not IsEmpty(RampBalls(x,1) ) Then
			If BallVel(RampBalls(x,0) ) > 1 Then ' if ball is moving, play rolling sound
				If RampType(x) Then
					PlaySound("RampLoop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
					StopSound("wireloop" & x)
				Else
					StopSound("RampLoop" & x)
					PlaySound("wireloop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
				End If
				RampBalls(x, 2) = RampBalls(x, 2) + 1
			Else
				StopSound("RampLoop" & x)
				StopSound("wireloop" & x)
			End If
			If RampBalls(x,0).Z < 30 And RampBalls(x, 2) > RampMinLoops Then    'if ball is on the PF, remove  it
				StopSound("RampLoop" & x)
				StopSound("wireloop" & x)
				Wremoveball RampBalls(x,1)
			End If
		Else
			StopSound("RampLoop" & x)
			StopSound("wireloop" & x)
		End If
	Next
	If Not RampBalls(0,0) Then RampRoll.enabled = 0
	
End Sub

' This can be used to debug the Ramp Roll time.  You need to enable the tbWR timer on the TextBox
Sub tbWR_Timer()    'debug textbox
	Me.text = "on? " & RampBalls(0, 0) & " timer: " & RampRoll.Enabled & vbNewLine & _
	"1 " & TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & " type:" & RampType(1) & " Loops:" & RampBalls(1, 2) & vbNewLine & _
	"2 " & TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & " type:" & RampType(2) & " Loops:" & RampBalls(2, 2) & vbNewLine & _
	"3 " & TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & " type:" & RampType(3) & " Loops:" & RampBalls(3, 2) & vbNewLine & _
	"4 " & TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & " type:" & RampType(4) & " Loops:" & RampBalls(4, 2) & vbNewLine & _
	"5 " & TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & " type:" & RampType(5) & " Loops:" & RampBalls(5, 2) & vbNewLine & _
	"6 " & TypeName(RampBalls(6, 0)) & " ID:" & RampBalls(6, 1) & " type:" & RampType(6) & " Loops:" & RampBalls(6, 2) & vbNewLine & _
	" "
End Sub


Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
	BallPitch = pSlope(BallVel(ball), 1, - 1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
	BallPitchV = pSlope(BallVel(ball), 1, - 4000, 60, 7000)
End Function



'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************

'******************************************************
'****  ZMCH: FLEEP MECHANICAL SOUNDS
'******************************************************

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for the TOM table

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.
' Create the following new collections:
'     Metals (all metal objects, metal walls, metal posts, metal wire guides)
'     Apron (the apron walls and plunger wall)
'     Walls (all wood or plastic walls)
'     Rollovers (wire rollover triggers, star triggers, or button triggers)
'     Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
'     Gates (plate gates)
'     GatesWire (wire gates)
'     Rubbers (all rubbers including posts, sleeves, pegs, and bands)
' When creating the collections, make sure "Fire events for this collection" is checked.
' You'll also need to make sure "Has Hit Event" is checked for each object placed in these collections (not necessary for gates and triggers).
' Once the collections and objects are added, the save, close, and restart VPX.
'
' Many places in the script need to be modified to include the correct sound effect subroutine calls. The tutorial videos linked below demonstrate
' how to make these updates. But in summary the following needs to be updated:
'    - Nudging, plunger, coin-in, start button sounds will be added to the keydown and keyup subs.
'    - Flipper sounds in the flipper solenoid subs. Flipper collision sounds in the flipper collide subs.
'    - Bumpers, slingshots, drain, ball release, knocker, spinner, and saucers in their respective subs
'    - Ball rolling sounds sub
'
' Tutorial vides by Apophis
' Part 1:     https://youtu.be/PbE2kNiam3g
' Part 2:     https://youtu.be/B5cm1Y8wQsk
' Part 3:     https://youtu.be/eLhWyuYOyGg


'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1                                                        'volume level; range [0, 1]
NudgeLeftSoundLevel = 1                                                    'volume level; range [0, 1]
NudgeRightSoundLevel = 1                                                'volume level; range [0, 1]
NudgeCenterSoundLevel = 1                                                'volume level; range [0, 1]
StartButtonSoundLevel = 0.1                                                'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr                                            'volume level; range [0, 1]
PlungerPullSoundLevel = 1                                                'volume level; range [0, 1]
RollingSoundFactor = 1.1 / 5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010                                   'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635                                'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0                                                'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45                                              'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel                                'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel                                'sound helper; not configurable
SlingshotSoundLevel = 0.95                                                'volume level; range [0, 1]
BumperSoundFactor = 4.25                                                'volume multiplier; must not be zero
KnockerSoundLevel = 1                                                     'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2                                    'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055 / 5                                            'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075 / 5                                            'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075 / 5                                        'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025                                    'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025                                    'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8                                    'volume level; range [0, 1]
WallImpactSoundFactor = 0.075                                            'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075 / 3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5 / 5                                                    'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10                                            'volume multiplier; must not be zero
DTSoundLevel = 0.25                                                        'volume multiplier; must not be zero
RolloverSoundLevel = 0.25                                                  'volume level; range [0, 1]
SpinnerSoundLevel = 0.5                                                  'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8                                                        'volume level; range [0, 1]
BallReleaseSoundLevel = 1                                                'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2                                    'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015                                        'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025 / 5                                                    'volume multiplier; must not be zero


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
	PlaySound playsoundparams, - 1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
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
	PlaySound playsoundparams, - 1, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
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
	PlaySound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
	PlaySound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
	PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


'******************************************************
'  Fleep  Supporting Ball & Sound Functions
'******************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
	Dim tmp
	tmp = tableobj.y * 2 / tableheight - 1
	
	If tmp > 7000 Then
		tmp = 7000
	ElseIf tmp <  - 7000 Then
		tmp =  - 7000
	End If
	
	If tmp > 0 Then
		AudioFade = CSng(tmp ^ 10)
	Else
		AudioFade = CSng( - (( - tmp) ^ 10) )
	End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
	Dim tmp
	tmp = tableobj.x * 2 / tablewidth - 1
	
	If tmp > 7000 Then
		tmp = 7000
	ElseIf tmp <  - 7000 Then
		tmp =  - 7000
	End If
	
	If tmp > 0 Then
		AudioPan = CSng(tmp ^ 10)
	Else
		AudioPan = CSng( - (( - tmp) ^ 10) )
	End If
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
	Vol = CSng(BallVel(ball) ^ 2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
	Volz = CSng((ball.velz) ^ 2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
	Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
	BallVel = Int(Sqr((ball.VelX ^ 2) + (ball.VelY ^ 2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
	VolPlayfieldRoll = RollingSoundFactor * 0.0005 * CSng(BallVel(ball) ^ 3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
	PitchPlayfieldRoll = BallVel(ball) ^ 2 * 15
End Function

Function RndInt(min, max)
	RndInt = Int(Rnd() * (max - min + 1) + min)' Sets a random number integer between min and max
End Function

Function RndNum(min, max)
	RndNum = Rnd() * (max - min) + min' Sets a random number between min and max
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////
Sub SoundStartButton()
	PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
	PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeLeftSoundLevel * VolumeDial, - 0.1, 0.25
End Sub

Sub SoundNudgeRight()
	PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
	PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
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
	PlaySoundAtLevelStatic ("Drain_" & Int(Rnd * 11) + 1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
	PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd * 7) + 1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundSlingshotLeft(sling)
	PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd * 10) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
	PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd * 8) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////
' MODIFICATION: Normally, "Vol(ActiveBall) *" is before the BumperSoundFactor. But because we run bumpers manually (self test, ball search, etc), ActiveBall may be undefined. So we don't use it in Saving Wallden.

Sub RandomSoundBumperTop(Bump)
	PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd * 5) + 1,DOFContactors), BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
	PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd * 5) + 1,DOFContactors), BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
	PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd * 5) + 1,DOFContactors), BumperSoundFactor, Bump
End Sub

'/////////////////////////////  SPINNER SOUNDS  ////////////////////////////
Sub SoundSpinner(spinnerswitch)
	PlaySoundAtLevelStatic ("Spinner"), SpinnerSoundLevel, spinnerswitch
End Sub


'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////
Sub SoundFlipperUpAttackLeft(flipper)
	FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
	PlaySoundAtLevelStatic SoundFX("Flipper_Attack-L01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
	FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
	PlaySoundAtLevelStatic SoundFX("Flipper_Attack-R01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////
Sub RandomSoundFlipperUpLeft(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd * 7) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd * 8) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
	FlipperLeftHitParm = parm / 10
	If FlipperLeftHitParm > 1 Then
		FlipperLeftHitParm = 1
	End If
	FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
	RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
	FlipperRightHitParm = parm / 10
	If FlipperRightHitParm > 1 Then
		FlipperRightHitParm = 1
	End If
	FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
	RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
	PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd * 7) + 1), parm * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////
Sub RandomSoundRollover()
	PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd * 4) + 1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
	RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////
Sub Rubbers_Hit(idx)
	Dim finalspeed
	finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
	If finalspeed > 5 Then
		RandomSoundRubberStrong 1
	End If
	If finalspeed <= 5 Then
		RandomSoundRubberWeak()
	End If
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////
Sub RandomSoundRubberStrong(voladj)
	Select Case Int(Rnd * 10) + 1
		Case 1
			PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 2
			PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 3
			PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 4
			PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 5
			PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 6
			PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 7
			PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 8
			PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 9
			PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 10
			PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6 * voladj
	End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////
Sub RandomSoundRubberWeak()
	PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd * 9) + 1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////
Sub Walls_Hit(idx)
	RandomSoundWall()
End Sub

Sub RandomSoundWall()
	Dim finalspeed
	finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
	If finalspeed > 16 Then
		Select Case Int(Rnd * 5) + 1
			Case 1
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 2
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 3
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 4
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 5
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
		End Select
	End If
	If finalspeed >= 6 And finalspeed <= 16 Then
		Select Case Int(Rnd * 4) + 1
			Case 1
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 2
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 3
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 4
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
		End Select
	End If
	If finalspeed < 6 Then
		Select Case Int(Rnd * 3) + 1
			Case 1
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 2
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 3
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
		End Select
	End If
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////
Sub RandomSoundMetal()
	PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd * 13) + 1), Vol(ActiveBall) * MetalImpactSoundFactor
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
	Dim finalspeed
	finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
	If finalspeed > 16 Then
		PlaySoundAtLevelActiveBall ("Apron_Bounce_" & Int(Rnd * 2) + 1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
	End If
	If finalspeed >= 6 And finalspeed <= 16 Then
		Select Case Int(Rnd * 2) + 1
			Case 1
				PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
			Case 2
				PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
		End Select
	End If
	If finalspeed < 6 Then
		Select Case Int(Rnd * 2) + 1
			Case 1
				PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
			Case 2
				PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
		End Select
	End If
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////
Sub RandomSoundBottomArchBallGuideHardHit()
	PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd * 3) + 1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
	If Abs(cor.ballvelx(ActiveBall.id) < 4) And cor.ballvely(ActiveBall.id) > 7 Then
		RandomSoundBottomArchBallGuideHardHit()
	Else
		RandomSoundBottomArchBallGuide
	End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////
Sub RandomSoundFlipperBallGuide()
	Dim finalspeed
	finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
	If finalspeed > 16 Then
		Select Case Int(Rnd * 2) + 1
			Case 1
				PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
			Case 2
				PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
		End Select
	End If
	If finalspeed >= 6 And finalspeed <= 16 Then
		PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd * 3) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
	End If
	If finalspeed < 6 Then
		PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd * 7) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
	End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////
Sub RandomSoundTargetHitStrong()
	PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
	PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
	Dim finalspeed
	finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
	If finalspeed > 10 Then
		RandomSoundTargetHitStrong()
		RandomSoundBallBouncePlayfieldSoft ActiveBall
	Else
		RandomSoundTargetHitWeak()
	End If
End Sub

Sub Targets_Hit (idx)
	PlayTargetSound
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////
Sub RandomSoundBallBouncePlayfieldSoft(aBall)
	Select Case Int(Rnd * 9) + 1
		Case 1
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
		Case 2
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
		Case 3
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
		Case 4
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
		Case 5
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
		Case 6
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
		Case 7
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
		Case 8
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
		Case 9
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
	End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
	PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd * 7) + 1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////
Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
	Select Case Int(Rnd * 5) + 1
		Case 1
			PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 2
			PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 3
			PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 4
			PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 5
			PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
	End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()
	PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd * 2) + 1), GateSoundLevel, ActiveBall
End Sub

Sub SoundHeavyGate()
	PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, ActiveBall
End Sub

Sub Gates_hit(idx)
	SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
	SoundPlayfieldGate
End Sub

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
	PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
	PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub


Sub Arch1_hit()
	If ActiveBall.velx > 1 Then SoundPlayfieldGate
	StopSound "Arch_L1"
	StopSound "Arch_L2"
	StopSound "Arch_L3"
	StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
	If ActiveBall.velx <  - 8 Then
		RandomSoundRightArch
	End If
End Sub

Sub Arch2_hit()
	If ActiveBall.velx < 1 Then SoundPlayfieldGate
	StopSound "Arch_R1"
	StopSound "Arch_R2"
	StopSound "Arch_R3"
	StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
	If ActiveBall.velx > 10 Then
		RandomSoundLeftArch
	End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
	PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd * 2) + 1), SaucerLockSoundLevel, ActiveBall
End Sub

Sub SoundSaucerKick(scenario, saucer)
	Select Case scenario
		Case 0
			PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
		Case 1
			PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
	End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////
Sub OnBallBallCollision(ball1, ball2, velocity)
	Dim snd
	Select Case Int(Rnd * 7) + 1
		Case 1
			snd = "Ball_Collide_1"
		Case 2
			snd = "Ball_Collide_2"
		Case 3
			snd = "Ball_Collide_3"
		Case 4
			snd = "Ball_Collide_4"
		Case 5
			snd = "Ball_Collide_5"
		Case 6
			snd = "Ball_Collide_6"
		Case 7
			snd = "Ball_Collide_7"
	End Select
	
	PlaySound (snd), 0, CSng(velocity) ^ 2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
	PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd * 6) + 1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
	PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd * 6) + 1), 200, obj
End Sub

'/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

Const RelayFlashSoundLevel = 0.315                                    'volume level; range [0, 1];
Const RelayGISoundLevel = 1.05                                    'volume level; range [0, 1];

Sub Sound_GI_Relay(toggle, obj)
	Select Case toggle
		Case 1
			PlaySoundAtLevelStatic ("Relay_GI_On"), 0.025 * RelayGISoundLevel, obj
		Case 0
			PlaySoundAtLevelStatic ("Relay_GI_Off"), 0.025 * RelayGISoundLevel, obj
	End Select
End Sub

Sub Sound_Flash_Relay(toggle, obj)
	Select Case toggle
		Case 1
			PlaySoundAtLevelStatic ("Relay_Flash_On"), 0.025 * RelayFlashSoundLevel, obj
		Case 0
			PlaySoundAtLevelStatic ("Relay_Flash_Off"), 0.025 * RelayFlashSoundLevel, obj
	End Select
End Sub

'/////////////////////////////////////////////////////////////////
'                    End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************


'*******************************************
'  ZDRN: Drain, Trough, and Ball Release
'*******************************************
' It is best practice to never destroy balls. This leads to more stable and accurate pinball game simulations.
' The following code supports a "physical trough" where balls are not destroyed.
' To use this,
'    - The trough geometry needs to be modeled with walls, and a set of kickers needs to be added to
'     the trough. The number of kickers depends on the number of physical balls on the table.
'     - A timer called "UpdateTroughTimer" needs to be added to the table. It should have an interval of 300 and be initially disabled.
'   - The balls need to be created within the Table1_Init sub. A global ball array (gBOT) can be created and used throughout the script


dim updateTroughState
updateTroughState = 0

dim BIPWhenDrained
BIPWhenDrained = 0 'Number of balls that were in play when a ball last drained (0 = no drain occurred at this time)

Sub swTrough1_Hit
	UpdateTrough
End Sub
Sub swTrough1_UnHit
	UpdateTrough
End Sub
Sub swTrough2_Hit
	UpdateTrough
End Sub
Sub swTrough2_UnHit
	UpdateTrough
End Sub
Sub swTrough3_Hit
	UpdateTrough
End Sub
Sub swTrough3_UnHit
	UpdateTrough
End Sub
Sub swTrough4_Hit
	UpdateTrough
End Sub
Sub swTrough4_UnHit
	UpdateTrough
End Sub
Sub swTrough5_Hit
	ballsInTrough = ballsInTrough + 1
	If ballCameFromSubway = True Then
		ballCameFromSubway = False
		BIS = BIS - 1
		If BIS < 0 Then BIS = 0
		If DEBUG_LOG_TROUGH = True Then WriteToLog "Trough", "Ball left the subway. " & BIS & " balls in subway; " & BIP & " balls in play."
	End If
	UpdateTrough
End Sub
Sub swTrough5_UnHit
	UpdateTrough
End Sub
Sub swTrough5b_Hit
	RandomSoundDrain swTrough5b
	UpdateTrough
End Sub
Sub swTrough5b_UnHit
	UpdateTrough
End Sub

Sub UpdateTrough
	updateTroughState = 1
	UpdateTroughTimer.Interval = 300
	UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer 'Modified for Wallden
	'Advance pinballs in the trough
	If swTrough1.BallCntOver = 0 Then swTrough2.kick 57, 10
	If swTrough2.BallCntOver = 0 Then swTrough3.kick 57, 10
	If swTrough3.BallCntOver = 0 Then swTrough4.kick 57, 10
	If swTrough4.BallCntOver = 0 Then swTrough5.kick 57, 10
	If swTrough5.BallCntOver = 0 Then swTrough5b.kickZ 0, 20, 0, 60

	If DEBUG_LOG_TROUGH = True Then WriteToLog "Trough", "Trough updating. " & ballsInTrough & " balls in the trough (presumed); " & BIP & " balls in play; " & BIS & " balls in the subway."

	'If this timer was not called by a change in trough ball count, and there are no balls in the drain, consider the trough "settled" and run routines like drain checks.
	If updateTroughState = 0 And Drain.BallCntOver = 0 Then
		If DEBUG_LOG_TROUGH = True Then WriteToLog "Trough", "Trough balls settled. Balls actually in trough now: " & actualBallsInTrough
		ballsInTrough = actualBallsInTrough
		Me.Enabled = 0
		If BIPWhenDrained > 0 Then
			checkDrainedBalls True
		End If
		BIPWhenDrained = 0
		updateClocks
	End If

	updateTroughState = 0
End Sub

Function BIP
	BIP = tnob - lob - ballsInTrough - BIS
End Function

Function actualBallsInTrough
	actualBallsInTrough = swTrough1.BallCntOver + swTrough2.BallCntOver + swTrough3.BallCntOver + swTrough4.BallCntOver + swTrough5.BallCntOver
	ballsInTrough = actualBallsInTrough
End Function

'**********Sling Shot Animations
' ZSLG: Rstep and Lstep  are the variables that increment the animation
'****************
Dim LStep, Rstep

Sub LeftSlingShot_Timer()
	Select Case Lstep
		Case 3
			LSLing1.Visible = 0
			LSLing2.Visible = 1
			sling2.rotx = 10
		Case 4
			LSLing2.Visible = 0
			LSLing.Visible = 1
			sling2.rotx = 0
			LeftSlingShot.TimerEnabled = 0
	End Select
	Lstep = Lstep + 1
End Sub

Sub RightSlingShot_Timer()
	Select Case RStep
		Case 3
			RSLing1.Visible = 0
			RSLing2.Visible = 1
			sling1.rotx = 10
		Case 4
			RSLing2.Visible = 0
			RSLing.Visible = 1
			sling1.rotx = 0
			RightSlingShot.TimerEnabled = 0
	End Select
	RStep = RStep + 1
End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  LIGHTING / RAINBOW LIGHTS
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'


'********************************************************************************************
' Only for VPX 10.2 and higher.
' FlashForMs will blink light or a flasher for TotalPeriod(ms) at rate of BlinkPeriod(ms)
' When TotalPeriod done, light or flasher will be set to FinalState value where
' Final State values are:   0=Off, 1=On, 2=Return to previous State
'********************************************************************************************

Sub FlashForMs(MyLight, TotalPeriod, BlinkPeriod, FinalState) 'thanks gtxjoe for the first version
	
	If TypeName(MyLight) = "Light" Then
		
		If FinalState = 2 Then
			FinalState = MyLight.State 'Keep the current light state
		End If
		MyLight.BlinkInterval = BlinkPeriod
		MyLight.Duration 2, TotalPeriod, FinalState
	ElseIf TypeName(MyLight) = "Flasher" Then
		
		Dim steps
		
		' Store all blink information
		steps = Int(TotalPeriod / BlinkPeriod + .5) 'Number of ON/OFF steps to perform
		If FinalState = 2 Then                      'Keep the current flasher state
			FinalState = Abs(MyLight.Visible)
		End If
		MyLight.UserValue = steps * 10 + FinalState 'Store # of blinks, and final state
		
		' Start blink timer and create timer subroutine
		MyLight.TimerInterval = BlinkPeriod
		MyLight.TimerEnabled = 0
		MyLight.TimerEnabled = 1
		ExecuteGlobal "Sub " & MyLight.Name & "_Timer:" & "Dim tmp, steps, fstate:tmp=me.UserValue:fstate = tmp Mod 10:steps= tmp\10 -1:Me.Visible = steps Mod 2:me.UserValue = steps *10 + fstate:If Steps = 0 Then Me.Visible = fstate:Me.TimerEnabled=0:End If:End Sub"
	End If
End Sub

'******************************************
' WLHT: Change light color - simulate color leds
' changes the light color and state
'******************************************

'full colors
Dim red, orange, amber, yellow, green, darkgreen, blue, darkblue, purple, white, base
red = RGB(255, 0, 0)
orange = RGB(255, 64, 0)
amber = RGB(255, 153, 0)
yellow = RGB(255, 255, 0)
green = RGB(0, 255, 0)
darkgreen = RGB(0, 64, 0)
blue = RGB(0, 128, 255)
darkblue = RGB(0, 0, 255)
purple = RGB(128, 0, 255)
white = RGB(255, 252, 224)
base = RGB(255, 197, 143)

'dark colors
Dim dred, dorange, damber, dyellow, dgreen, ddarkgreen, dblue, ddarkblue, dpurple, dwhite, dbase
dred = RGB(26, 0, 0)
dorange = RGB(26, 6, 0)
damber = RGB(26, 15, 0)
dyellow = RGB(26, 26, 0)
dgreen = RGB(0, 26, 0)
ddarkgreen = RGB(0, 6, 0)
dblue = RGB(0, 13, 26)
ddarkblue = RGB(0, 0, 26)
dpurple = RGB(13, 0, 26)
dwhite = RGB(26, 25, 22)
dbase = RGB(26, 20, 14)

Sub LightsResetColor() 	'Reset light colors to their defaults
	LightsColor shotLights, Array(dwhite, white)
	LightsColor powerupLights, Array(dblue, blue)
	LightsColor escapeLights, Array(dorange, orange)
	LightsColor vikingLights, Array(dblue, blue)
	LightsColor sniperLights, Array(damber, amber)
	LightsColor lockLights, Array(dgreen, green)
	LightsColor toBumpersLights, Array(dbase, base)
	LightsColor trainLights, Array(dyellow, yellow)
	LightsColor hpLights, Array(dred, red)
	LightsColor chaserLights, Array(dgreen, green)
	LightsColor blacksmithLights, Array(dpurple, purple)
	LightsColor dragonLights, Array(ddarkblue, darkblue)
	LightsColor poisonLights, Array(dred, red)
	LightsColor lanternLights, Array(damber, amber)
End Sub

Sub LightIntervalByTime(light, milliseconds, upperValue) 'Update a light's blink interval based on the number of milliseconds provided (such as time left) and the upperTime (number of milliseconds at which lights blink slowest)
	If milliseconds > 0 Then
		If milliseconds < upperValue And upperValue > 0 Then
			LampC.UpdateBlinkInterval light, 333 - (300 * ((upperValue - milliseconds) / upperValue))
		Else
			LampC.UpdateBlinkInterval light, 333
		End If
	End If
End Sub

Sub LightsOn(collection, color, lvl)
	Dim coloredLight
	For Each coloredLight In collection
		If Not IsNull(lvl) Then lampC.LightLevel coloredLight, lvl
		If IsNull(color) Then
			lampC.LightOn coloredLight
		Else
			lampC.LightOnWithColor coloredLight, color
		End If
	Next
End Sub

Sub LightsColor(collection, color)
	Dim coloredLight
	For Each coloredLight In collection
		lampC.LightColor coloredLight, color
	Next
End Sub

Sub LightsBlink(collection, color, interval, lvl)
	Dim coloredLight
	For Each coloredLight In collection
		lampC.LightOff coloredLight
		LampC.UpdateBlinkInterval coloredLight, interval
		If Not IsNull(lvl) Then lampC.LightLevel coloredLight, lvl
		If Not IsNull(color) Then lampC.LightColor coloredLight, color
		lampC.Blink coloredLight
	Next
End Sub

Sub LightsOff(collection)
	Dim coloredLight
	For Each coloredLight In collection
		lampC.LightOff coloredLight
	Next
End Sub

Sub StartLightSequence(lightSeq, seqColor, seqInterval, seqSequence, seqTail, seqRepeat)
	lightSeq.UpdateInterval = seqInterval
	lightSeq.Play seqSequence, seqTail, seqRepeat
	
	lampC.SyncWithVpxLights lightSeq
	If Not IsNull(seqColor) Then
		lampC.SetVpxSyncLightColor seqColor
	End If
End Sub

Sub StopLightSequence(lightSeq)
	lightSeq.StopPlay
	lampC.StopSyncWithVpxLights()
End Sub

Sub StartAttractLightSequence()
	StopLightSequence LightSeqAttract
	StopLightSequence LightSeqAttractB
	
	Dim a
	lampC.SyncWithVpxLights LightSeqAttract
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqUpOn, 50, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqDownOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqCircleOutOn, 15, 2
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqUpOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqDownOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqUpOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqDownOn, 25, 1
	LightSeqAttractB.UpdateInterval = 10
	LightSeqAttractB.Play SeqCircleOutOn, 15, 3
	LightSeqAttractB.UpdateInterval = 5
	LightSeqAttractB.Play SeqRightOn, 50, 1
	LightSeqAttractB.UpdateInterval = 5
	LightSeqAttractB.Play SeqLeftOn, 50, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqRightOn, 50, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqLeftOn, 50, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqRightOn, 40, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqLeftOn, 40, 1
	LightSeqAttractB.UpdateInterval = 10
	LightSeqAttractB.Play SeqRightOn, 30, 1
	LightSeqAttractB.UpdateInterval = 10
	LightSeqAttractB.Play SeqLeftOn, 30, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqRightOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqLeftOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqRightOn, 15, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqLeftOn, 15, 1
	LightSeqAttractB.UpdateInterval = 10
	LightSeqAttractB.Play SeqCircleOutOn, 15, 3
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqLeftOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqRightOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqLeftOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqUpOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqDownOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqUpOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqDownOn, 25, 1
	LightSeqAttractB.UpdateInterval = 5
	LightSeqAttractB.Play SeqStripe1VertOn, 50, 2
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqCircleOutOn, 15, 2
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqStripe1VertOn, 50, 3
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqLeftOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqRightOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqLeftOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqUpOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqDownOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqCircleOutOn, 15, 2
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqStripe2VertOn, 50, 3
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqLeftOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqRightOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqLeftOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqUpOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqDownOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqUpOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqDownOn, 25, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqStripe1VertOn, 25, 3
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqStripe2VertOn, 25, 3
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqUpOn, 15, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqDownOn, 15, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqUpOn, 15, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqDownOn, 15, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqUpOn, 15, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqDownOn, 15, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqRightOn, 15, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqLeftOn, 15, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqRightOn, 15, 1
	LightSeqAttractB.UpdateInterval = 8
	LightSeqAttractB.Play SeqLeftOn, 15, 1
End Sub

Sub LightSeqAttract_PlayDone()
	lampC.StopSyncWithVpxLights()
End Sub

Sub LightSeqAttractB_PlayDone()
	StartAttractLightSequence
End Sub

Sub StandardSeq_PlayDone()
	lampC.StopSyncWithVpxLights()
End Sub

Sub SetLightColor(n, col, stat)
	' DEPRECATED
End Sub

Sub SetLightColorCollection(n, col, stat)
	' DEPRECATED
End Sub

Sub ResetAllLightsColor
	' DEPRECATED
End Sub

Sub UpdateBonusColors
	' DEPRECATED
End Sub

'*************************
' ZRBW: Rainbow Changing Lights
'*************************

Dim RGBStep, RGBFactor, rRed, rGreen, rBlue, RainbowLights

Sub StartRainbow(n)
	Set RainbowLights = n
	RGBStep = 0
	RGBFactor = 5
	rRed = 255
	rGreen = 0
	rBlue = 0
	RainbowTimer.Enabled = 1
End Sub

Sub StopRainbow(n)
	Dim obj
	RainbowTimer.Enabled = 0
	RainbowTimer.Enabled = 0
	If IsObject(RainbowLights) Then
		For Each obj In RainbowLights
			'SetLightColor obj, "white", 0 ' TODO: Change this
		Next
	End If
End Sub

Sub RainbowTimer_Timer 'rainbow led light color changing
	Dim obj
	Select Case RGBStep
		Case 0 'Green
			rGreen = rGreen + RGBFactor
			If rGreen > 255 Then
				rGreen = 255
				RGBStep = 1
			End If
		Case 1 'Red
			rRed = rRed - RGBFactor
			If rRed < 0 Then
				rRed = 0
				RGBStep = 2
			End If
		Case 2 'Blue
			rBlue = rBlue + RGBFactor
			If rBlue > 255 Then
				rBlue = 255
				RGBStep = 3
			End If
		Case 3 'Green
			rGreen = rGreen - RGBFactor
			If rGreen < 0 Then
				rGreen = 0
				RGBStep = 4
			End If
		Case 4 'Red
			rRed = rRed + RGBFactor
			If rRed > 255 Then
				rRed = 255
				RGBStep = 5
			End If
		Case 5 'Blue
			rBlue = rBlue - RGBFactor
			If rBlue < 0 Then
				rBlue = 0
				RGBStep = 0
			End If
	End Select
	For Each obj In RainbowLights
		lampC.LightColor obj, Array(RGB(rRed \ 10, rGreen \ 10, rBlue \ 10), RGB(rRed, rGreen, rBlue))
	Next
	
	' TODO: Cannot be left here as rainbow lights should control specific light sets
	LightsOn GI, Array(RGB(rRed \ 10, rGreen \ 10, rBlue \ 10), RGB(rRed, rGreen, rBlue)), 100
End Sub

'***************************************************************
' ZQUE: VPIN WORKSHOP ADVANCED QUEUING SYSTEM - 1.2.0
'***************************************************************
' WHAT IS IT?
' The VPin Workshop Advanced Queuing System allows table authors
' to put sub routine calls in a queue without creating a bunch
' of timers. There are many use cases for this: queuing sequences
' for light shows and DMD scenes, delaying solenoids until the
' DMD is finished playing all its sequences (such as holding a
' ball in a scoop), managing what actions take priority over
' others (e.g. an extra ball sequence is probably more important
' than a small jackpot), and many more.
'
' This system uses Scripting.Dictionary, a single timer, and the
' GameTime global to keep track of everything in the queue.
' This allows for better stability and a virtually unlimited
' number of items in the queue. It also allows for greater
' versatility, like pre-delays, queue delays, priorities, and
' even modifying items in the queue.
'
' The VPin Workshop Queuing System can replace vpmTimer as a
' proper queue system (each item depends on the previous)
' whereas vpmTimer is a collection of virtual timers that run
' in parallel. It also adds on other advanced functionality.
' However, this queue system does not have ROM support out of
' the box like vpmTimer does.
'
' I recommend reading all the comments before you implement the
' queuing system into your table.
'
' WHAT YOU NEED to use the queuing system:
' 1) Put this VBS file in your scripts folder, or copy / paste
'    the code into your table script (and skip step 2).
' 2) Include this file via Scripting.FileSystemObject, and
'    ExecuteGlobal it.
' 3) Make one or more queues by constructing the vpwQueueManager:
'    Dim queue : Set queue = New vpwQueueManager
' 4) Create (or use) a timer that is always enabled and
'    preferably has an interval of 1 millisecond. Use a
'    higher number for less time precision but less resource
'    use. You only need one timer even if you
'    have multiple queues.
' 5) For each queue you created, call its Tick routine in
'    the timer's *_timer() routine:
'    queue.Tick
' 6) You're done! Refer to the routines in vpwQueueManager to
'    learn how to use the queuing system.
'
' TUTORIAL: https://youtu.be/kpPYgOiUlxQ
'***************************************************************

'===========================================
' vpwQueueManager
' This class manages a queue of
' vpwQueueItems and executes them.
'===========================================
Class vpwQueueManager
	Public qItems ' A dictionary of vpwQueueItems in the queue (do NOT use native Scripting.Dictionary.Add/Remove; use the vpwQueueManager's Add/Remove methods instead!)
	Public preQItems ' A dictionary of vpwQueueItems pending to be added to qItems
	Public debugOn 'Null = no debug. String = activate debug by using this unique label for the queue. REQUIRES baldgeek's error logs.
	
	'----------------------------------------------------------
	' vpwQueueManager.qCurrentItem
	' This contains a string of the key currently active / at
	' the top of the queue. An empty string means no items are
	' active right now.
	' This is an important property; it should be monitored
	' in another timer or routine whenever you Add a queue item
	' with a -1 (indefinite) preDelay or postDelay. Then, for
	' preDelay, ExecuteCurrentItem should be called to run the
	' queue item. And for postDelay, DoNextItem should be
	' called to move to the next item in the queue.
	'
	' For example, let's say you add a queue item with the
	' key "kickTheBall" and an indefinite preDelay. You want
	' to wait until another timer fires before this queue item
	' executes and kicks the ball out of a scoop. In the other
	' timer, you will monitor qCurrentItem. Once it equals
	' "kickTheBall", call ExecuteCurrentItem, which will run
	' the queue item and presumably kick out the ball.
	'
	' WARNING!: If you do not properly execute one of these
	' callback routines on an indefinite delayed item, then
	' the queue will effectively freeze / stop until you do.
	'---------------------------------------------------------
	Public qCurrentItem
	
	Public preDelayTime ' The GameTime the preDelay for the qCurrentItem was started
	Public postDelayTime ' The GameTime the postDelay for the qCurrentItem was started
	
	Private onQueueEmpty ' A string or object to be called every time the queue empties (use the QueueEmpty property to get/set this)
	Private queueWasEmpty ' Boolean to determine if the queue was already empty when firing DoNextItem
	Private preDelayTransfer ' Number of milliseconds of preDelay to transfer over to the next queue item when doNextItem is called
	
	Private Sub Class_Initialize
		Set qItems = CreateObject("Scripting.Dictionary")
		Set preQItems = CreateObject("Scripting.Dictionary")
		qCurrentItem = ""
		onQueueEmpty = ""
		queueWasEmpty = True
		debugOn = Null
		preDelayTransfer = 0
	End Sub
	
	'----------------------------------------------------------
	' vpwQueueManager.Tick
	' This is where all the magic happens! Call this method in
	' your timer's _timer routine to check the queue and
	' execute the necessary methods. We do not iterate over
	' every item in the queue here, which allows for superior
	' performance even if you have hundreds of items in the
	' queue.
	'----------------------------------------------------------
	Public Sub Tick()
		Dim item
		If qItems.Count > 0 Then ' Don't waste precious resources if we have nothing in the queue
			
			' If no items are active, or the currently active item no longer exists, move to the next item in the queue.
			' (This is also a failsafe to ensure the queue continues to work even if an item gets manually deleted from the dictionary).
			If qCurrentItem = "" Or Not qItems.Exists(qCurrentItem) Then
				DoNextItem
			Else ' We are good; do stuff as normal
				Set item = qItems.item(qCurrentItem)
				
				If item.Executed Then
					' If the current item was executed and the post delay passed, go to the next item in the queue
					If item.postDelay >= 0 And GameTime >= (postDelayTime + item.postDelay) Then
						DebugLog qCurrentItem & " - postDelay of " & item.postDelay & " passed."
						DoNextItem
					End If
				Else
					' If the current item expires before it can be executed, go to the next item in the queue
					If item.timeToLive > 0 And GameTime >= (item.queuedOn + item.timeToLive) Then
						DebugLog qCurrentItem & " - expired (Time To live). Moving To the Next queue item."
						DoNextItem
					End If
					
					' If the current item was not executed yet and the pre delay passed, then execute it
					If item.preDelay >= 0 And GameTime >= (preDelayTime + item.preDelay) Then
						DebugLog qCurrentItem & " - preDelay of " & item.preDelay & " passed. Executing callback."
						item.Execute
						preDelayTime = 0
						postDelayTime = GameTime
					End If
				End If
			End If
		End If
		
		' Loop through each item in the pre-queue to find any that is ready to be added
		If preQItems.Count > 0 Then
			Dim k, key
			k = preQItems.Keys
			For Each key In k
				Set item = preQItems.Item(key)
				
				' If a queue item was pre-queued and is ready to be considered as actually in the queue, add it
				If GameTime >= (item.queuedOn + item.preQueueDelay) Then
					DebugLog key & " (preQueue) - preQueueDelay of " & item.preQueueDelay & " passed. Item added To the main queue."
					preQItems.Remove key
					Me.Add key, item.Callback, item.priority, 0, item.preDelay, item.postDelay, item.timeToLive, item.executeNow
				End If
			Next
		End If
	End Sub
	
	'----------------------------------------------------------
	' vpwQueueManager.DoNextItem
	' Goes to the next item in the queue and deletes the
	' currently active one.
	'----------------------------------------------------------
	Public Sub DoNextItem()
		If Not qCurrentItem = "" Then
			If qItems.Exists(qCurrentItem) Then qItems.Remove qCurrentItem ' Remove the current item from the queue if it still exists
			qCurrentItem = ""
		End If
		
		If qItems.Count > 0 Then
			Dim k, key
			Dim nextItem
			Dim nextItemPriority
			Dim item
			nextItemPriority = 0
			nextItem = ""
			
			' Find which item needs to run next based on priority first, queue order second (ignore items with an active preQueueDelay)
			k = qItems.Keys
			For Each key In k
				Set item = qItems.Item(key)
				
				If item.preQueueDelay <= 0 And item.priority > nextItemPriority Then
					nextItem = key
					nextItemPriority = item.priority
				End If
			Next
			
			If qItems.Exists(nextItem) Then
				Set item = qItems.Item(nextItem)
				DebugLog "DoNextItem - checking " & nextItem & " (priority " & item.priority & ")"
				
				' Make sure the item is not expired and not already executed. If it is, remove it and re-call doNextItem
				If (item.timeToLive > 0 And GameTime >= (item.queuedOn + item.timeToLive + preDelayTransfer)) Or item.executed = True Then
					DebugLog "DoNextItem - " & nextItem & " expired (Time To live) Or already executed. Removing And going To the Next item."
					qItems.Remove nextItem
					DoNextItem
					Exit Sub
				End If
				
				'Transfer preDelay time when applicable
				If preDelayTransfer > 0 And item.preDelay > -1 Then
					DebugLog "DoNextItem " & nextItem & " - Transferred remaining postDelay of " & preDelayTransfer & " milliseconds from previously overridden queue item To its preDelay And timeToLive"
					qItems.Item(nextItem).preDelay = item.preDelay + preDelayTransfer
					If item.timeToLive > 0 Then qItems.Item(nextItem).timeToLive = item.timeToLive + preDelayTransfer
					preDelayTransfer = 0
				End If
				
				' Set item as current / active, and execute if it has no pre-delay (otherwise Tick will take care of pre-delay)
				qCurrentItem = nextItem
				If item.preDelay = 0 Then
					DebugLog "DoNextItem - " & nextItem & " Now active. It has no preDelay, so executing callback immediately."
					item.Execute
					preDelayTime = 0
					postDelayTime = GameTime
				Else
					DebugLog "DoNextItem - " & nextItem & " Now active. Waiting For a preDelay of " & item.preDelay & " before executing."
					preDelayTime = GameTime
					postDelayTime = 0
				End If
			End If
		ElseIf queueWasEmpty = False Then
			DebugLog "DoNextItem - Queue Is Now Empty; executing queueEmpty callback."
			CallQueueEmpty() ' Call QueueEmpty if this was the last item in the queue
		End If
	End Sub
	
	'----------------------------------------------------------
	' vpwQueueManager.ExecuteCurrentItem
	' Helper routine that can be used when the current item is
	' on an indefinite preDelay. Call this when you are ready
	' for that item to execute.
	'----------------------------------------------------------
	Public Sub ExecuteCurrentItem()
		If Not qCurrentItem = "" And qItems.Exists(qCurrentItem) Then
			DebugLog "ExecuteCurrentItem - Executing the callback For " & qCurrentItem & "."
			Dim item
			Set item = qItems.Item(qCurrentItem)
			item.Execute
			preDelayTime = 0
			postDelayTime = GameTime
		End If
	End Sub
	
	'----------------------------------------------------------
	' vpwQueueManager.Add
	' REQUIRES Class vpwQueueItem
	'
	' Add an item to the queue.
	'
	' PARAMETERS:
	'
	' key (string) - Unique name for this queue item
	' WARNING: Specifying a key that already exists will
	' overwrite the item in the queue. This is by design. Also
	' note the following behaviors:
	' * Tickers / clocks for tracking delay times will NOT be
	' restarted for this item (but the total duration will be
	' updated. For example, if the old preDelay was 3 seconds
	' and 2 seconds elapsed, but Add was called to update
	' preDelay to 5 seconds, then the queue item will now
	' execute in 3 more seconds (new preDelay - time elapsed)).
	' However, timeToLive WILL be restarted.
	' * Items will maintain their same place in the queue.
	' * If key = qCurrentItem (overwriting the currently active
	' item in the queue) and qCurrentItem already executed
	' the callback (but is waiting for a postDelay), then the
	' current queue item's remaining postDelay will be added to
	' the preDelay of the next item, and this item will be
	' added to the bottom of the queue for re-execution.
	' If you do not want it to re-execute, then add an If
	' guard on your call to the Add method checking
	' "If Not vpwQueueManager.qCurrentItem = key".
	'
	' qCallback (object|string) - An object to be called,
	' or string to be executed globally, when this queue item
	' runs. I highly recommend making sub routines for groups
	' of things that should be executed by the queue so that
	' your qCallback string does not get long, and you can
	' easily organize your callbacks. Also, use double
	' double-quotes when the call itself has quotes in it
	' (VBScript escaping).
	' Example: "playsound ""Plunger"""
	'
	' priority (number) - Items in the queue will be executed
	' in order from highest priority to lowest. Items with the
	' same priority will be executed in order according to
	' when they were added to the queue. Use any number
	' greater than 0. My recommendation is to make a plan for
	' your table on how you will prioritize various types of
	' queue items and what priority number each type should
	' have. Also, you should reserve priority 1 (lowest) to
	' items which should wait until everything else in the
	' queue is done (such as ejecting a ball from a scoop).
	'
	' preQueueDelay (number) - The number of
	' milliseconds before the queue actually considers this
	' item as "in the queue" (pretend you started a timer to
	' add this item into the queue after this delay; this
	' logically works in a similar way; the only difference is
	' timeToLive is still considered even when an item is
	' pre-queued.) Set to 0 to add to the queue immediately.
	' NOTE: this should be less than timeToLive.
	'
	' preDelay (number) - The number of milliseconds before
	' the qCallback executes once this item is active (top)
	' in the queue. Set this to 0 to immediately execute the
	' qCallback when this item becomes active.
	' Set this to -1 to have an indefinite delay until
	' vpwQueueManager.ExecuteCurrentItem is called (see the
	' comment for qCurrentItem for more information).
	' NOTE: this should be less than timeToLive. And, if
	' timeToLive runs out before preDelay runs out, the item
	' will be removed and will not execute.
	'
	' postDelay (number) - After the qCallback executes, the
	' number of milliseconds before moving on to the next item
	' in the queue. Set this to -1 to have an indefinite delay
	' until vpwQueueManager.DoNextItem is called (see the
	' comment for qCurrentItem for more information).
	'
	' timeToLive (number) - After this item is added to the
	' queue, the number of milliseconds before this queue item
	' expires / is removed if the qCallback is not executed by
	' then. Set to 0 to never expire. NOTE: If not 0, this
	' should be greater than preDelay + preQueueDelay or the
	' item will expire before the qCallback is executed.
	' Example use case: Maybe a player scored a jackpot, but
	' it would be awkward / irrelevant to play that jackpot
	' sequence if it hasn't played after a few seconds (e.g.
	' other items in the queue took priority).
	'
	' executeNow (boolean) - Specify true if this item
	' should interrupt the queue and run immediately. This
	' will only happen, however, if the currently active item
	' has a priority less than or equal to the item you are
	' adding. Note this does not bypass preQueueDelay nor
	' preDelay if set.
	' Example: If a player scores an extra ball, you might
	' want that to interrupt everything else going on as it
	' is an important milestone.
	'----------------------------------------------------------
	Public Sub Add(key, qCallback, priority, preQueueDelay, preDelay, postDelay, timeToLive, executeNow)
		DebugLog "Adding queue item " & key
		
		'Construct the item class
		Dim newClass
		Set newClass = New vpwQueueItem
		With newClass
			.Callback = qCallback
			.priority = priority
			.preQueueDelay = preQueueDelay
			.preDelay = preDelay
			.postDelay = postDelay
			.timeToLive = timeToLive
			.executeNow = executeNow
		End With
		
		'If we are attempting to overwrite the current queue item which already executed, take the remaining postDelay and add it to the preDelay of the next item. And set us up to immediately go to the next item while re-adding this item to the queue.
		If preQueueDelay <= 0 And qItems.Exists(key) And qCurrentItem = key Then
			If qItems.Item(key).executed = True Then
				DebugLog key & " (Add) - Attempting To overwrite the current queue item which already executed. Immediately re-queuing this item To the bottom of the queue, transferring the remaining postDelay To the Next item, And going To the Next item."
				If qItems.Item(key).postDelay >= 0 Then
					preDelayTransfer = ((postDelayTime + qItems.Item(key).postDelay) - GameTime)
				End If
				
				'Remove current queue item so we can go to the next item, this can be re-queued to the bottom, and the remaining postDelay transferred to the preDelay of the next item
				qItems.Remove qCurrentItem
				qCurrentItem = ""
			End If
		End If
		
		' Determine execution stuff if this item does not have a pre-queue delay
		If preQueueDelay <= 0 Then
			If executeNow = True Then
				' Make sure this item does not immediately execute if the current item has a higher priority
				If Not qCurrentItem = "" And qItems.Exists(qCurrentItem) Then
					Dim item
					Set item = qItems.Item(qCurrentItem)
					If item.priority <= priority Then
						DebugLog key & " (Add) - Execute Now was Set To True And this item's priority (" & priority & ") Is >= the active item's priority (" & item.priority & " from " & qCurrentItem & "). Making it the current active queue item."
						qCurrentItem = key
						If preDelay = 0 And preDelayTransfer = 0 Then
							DebugLog key & " (Add) - No pre-delay. Executing the callback immediately."
							newClass.Execute
							preDelayTime = 0
							postDelayTime = GameTime
						Else
							DebugLog key & " (Add) - Waiting For a pre-delay of " & (preDelay + preDelayTransfer) & " before executing the callback."
							preDelayTime = GameTime
							postDelayTime = 0
						End If
					Else
						DebugLog key & " (Add) - Execute Now was Set To True, but this item's priority (" & priority & ") Is Not >= the active item's priority (" & item.priority & " from " & qCurrentItem & "). This item will Not be executed Now And will be added To the queue normally."
					End If
				Else
					DebugLog key & " (Add) - Execute Now was Set To True And no item was active In the queue. Making it the current active queue item."
					qCurrentItem = key
					If preDelay = 0 Then
						DebugLog key & " (Add) - No pre-delay. Executing the callback immediately."
						preDelayTransfer = 0 'No preDelay transfer if we are immediately re-executing the same queue item
						newClass.Execute
						preDelayTime = 0
						postDelayTime = GameTime
					Else
						DebugLog key & " (Add) - Waiting For a pre-delay of " & preDelay & " before executing the callback."
						preDelayTime = GameTime
						postDelayTime = 0
					End If
				End If
			End If
			If qItems.Exists(key) Then 'Overwrite existing item in the queue if it exists
				DebugLog key & " (Add) - Already exists In the queue. Updating the item With the new parameters passed In Add."
				Set qItems.Item(key) = newClass
			Else
				DebugLog key & " (Add) - Added To the queue."
				qItems.Add key, newClass
			End If
			queueWasEmpty = False
		Else
			If preQItems.Exists(key) Then 'Overwrite existing item in the preQueue if it exists
				DebugLog key & " (Add) - Already exists In the preQueue. Updating the item With the new parameters passed In Add."
				Set preQItems.Item(key) = newClass
			Else
				DebugLog key & " (Add) - Added To the preQueue."
				preQItems.Add key, newClass
			End If
		End If
	End Sub
	
	'----------------------------------------------------------
	' vpwQueueManager.Remove
	'
	' Removes an item from the queue. It is better to use this
	' than to remove the item from qItems directly as this sub
	' will also call DoNextItem to advance the queue if
	' the item removed was the active item.
	' NOTE: This only removes items from qItems; to remove
	' an item from preQItems, use the standard
	' Scripting.Dictionary Remove method.
	'
	' PARAMETERS:
	'
	' key (string) - Unique name of the queue item to remove.
	'----------------------------------------------------------
	Public Sub Remove(key)
		If qItems.Exists(key) Then
			DebugLog key & " (Remove)"
			qItems.Remove key
			If qCurrentItem = key Or qCurrentItem = "" Then DoNextItem ' Ensure the queue does not get stuck
		End If
	End Sub
	
	'----------------------------------------------------------
	' vpwQueueManager.RemoveAll
	'
	' Removes all items from the queue / clears the queue.
	' It is better to call this sub than to remove all items
	' from qItems directly because this sub cleans up the queue
	' to ensure it continues to work properly.
	'
	' PARAMETERS:
	'
	' preQueue (boolean) - Also clear the pre-queue.
	'----------------------------------------------------------
	Public Sub RemoveAll(preQueue)
		DebugLog "Queue was emptied via RemoveAll."
		
		' Loop through each item in the queue and remove it
		Dim k, key
		k = qItems.Keys
		For Each key In k
			qItems.Remove key
		Next
		qCurrentItem = ""
		
		If queueWasEmpty = False Then CallQueueEmpty() ' Queue is now empty, so call our callback if applicable
		
		If preQueue Then
			k = preQItems.Keys
			For Each key In k
				preQItems.Remove key
			Next
		End If
	End Sub
	
	'----------------------------------------------------------
	' Get vpwQueueManager.QueueEmpty
	' Get the current callback for when the queue is empty.
	'----------------------------------------------------------
	Public Property Get QueueEmpty()
		If IsObject(onQueueEmpty) Then
			Set QueueEmpty = onQueueEmpty
		Else
			QueueEmpty = onQueueEmpty
		End If
	End Property
	
	'----------------------------------------------------------
	' Let vpwQueueManager.QueueEmpty
	' Set the callback to call every time the queue empties.
	' This could be useful for setting a sub routine to be
	' called each time the queue empties for doing things such
	' as ejecting balls from scoops. Unlike using the Add
	' method, this callback is immune from getting removed by
	' higher priority items in the queue and will be called
	' every time the queue is emptied, not just once.
	'
	' PARAMETERS:
	'
	' callback (object|string) - The callback to call every
	' time the queue empties.
	'----------------------------------------------------------
	Public Property Let QueueEmpty(callback)
		If IsObject(callback) Then
			Set onQueueEmpty = callback
		ElseIf VarType(callback) = vbString Then
			onQueueEmpty = callback
		End If
	End Property
	
	'----------------------------------------------------------
	' Get vpwQueueManager.CallQueueEmpty
	' Private method that actually calls the QueueEmpty
	' callback.
	'----------------------------------------------------------
	Private Sub CallQueueEmpty()
		If queueWasEmpty = True Then Exit Sub
		queueWasEmpty = True
		
		If IsObject(onQueueEmpty) Then
			Call onQueueEmpty(0)
		ElseIf VarType(onQueueEmpty) = vbString Then
			If onQueueEmpty > "" Then ExecuteGlobal onQueueEmpty
		End If
	End Sub
	
	'----------------------------------------------------------
	' DebugLog
	' Log something if debugOn is not null.
	' REQUIRES / uses the WriteToLog sub from Baldgeek's
	' error log library.
	'----------------------------------------------------------
	Private Sub DebugLog(message)
		If Not IsNull(debugOn) Then
			WriteToLog "VPW Queue " & debugOn, message
		End If
	End Sub
End Class

'===========================================
' vpwQueueItem
' Represents a single item for the queue
' system. Do NOT use this class directly.
' Instead, use the vpwQueueManager.Add
' routine.

' You can, however, access an individual
' item in the queue via
' vpwQueueManager.qItems and then modify
' its properties while it is still in the
' queue.
'===========================================
Class vpwQueueItem  ' Do not construct this class directly; use vpwQueueManager.Add instead, and vpwQueueManager.qItems.Item(key) to modify an item's properties.
	Public priority ' The item's set priority
	Public timeToLive ' The item's set timeToLive milliseconds requested
	Public preQueueDelay ' The item's pre-queue milliseconds requested
	Public preDelay ' The item's pre delay milliseconds requested
	Public postDelay ' The item's post delay milliseconds requested
	Public executeNow ' Whether the item was set to Execute immediately
	Private qCallback ' The item's callback object or string (use the Callback property on the class to get/set it)
	
	Public executed ' Whether or not this item's qCallback was executed yet
	Public queuedOn ' The game time this item was added to the queue
	Public executedOn ' The game time this item was executed
	
	Private Sub Class_Initialize
		' Defaults
		priority = 0
		timeToLive = 0
		preQueueDelay = 0
		preDelay = 0
		postDelay = 0
		qCallback = ""
		executeNow = False
		
		queuedOn = GameTime
		executedOn = 0
	End Sub
	
	'----------------------------------------------------------
	' vpwQueueItem.Execute
	' Executes the qCallback on this item if it was not yet
	' already executed.
	'----------------------------------------------------------
	Public Sub Execute()
		If executed Then Exit Sub ' Do not allow an item's qCallback to ever Execute more than one time
		
		'Mark as execute before actually executing callback; that way, if callback recursively adds the item back into the queue, then we can properly handle it.
		executed = True
		executedOn = GameTime
		
		' Execute qCallback
		If IsObject(qCallback) Then
			Call qCallback(0)
		ElseIf VarType(qCallback) = vbString Then
			If qCallback > "" Then ExecuteGlobal qCallback
		End If
	End Sub
	
	Public Property Get Callback()
		If IsObject(qCallback) Then
			Set Callback = qCallback
		Else
			Callback = qCallback
		End If
	End Property
	
	Public Property Let Callback(cb)
		If IsObject(cb) Then
			Set qCallback = cb
		ElseIf VarType(cb) = vbString Then
			qCallback = cb
		End If
	End Property
End Class

'***************************************************************
' END VPIN WORKSHOP ADVANCED QUEUING SYSTEM
'***************************************************************

'***************************************************************
' ZDAT: VPIN WORKSHOP DATA MANAGER - 1.1.0 BETA
'***************************************************************
' WHAT IS IT?
' The VPin Workshop Data Manager is a tool to manage key/value
' data and sync it with the VPReg. This allows you to set a
' standard data structure, modify it during games and whatnot,
' save it to the VPReg, and automatically load the saved value
' when you next load the game and initialize the class.
'
' This class may be useful for managing data such as high
' scores, settings from operator menus, game play progress
' (especially if you want to support saving / resuming games),
' and table / game statistics.
'
' NOTE: This library might not work on 32-bit set-ups! This is
' because data loaded into the Dictionary from the VPReg is
' converted to type Double if it is numeric.
'
' HOW TO USE
' 1) Put this VBS file in your VPX Scripts folder, or
'    copy / paste the code into your table script
'    (and skip step 2).
' 2) Include this VBS file in your table via ExecuteGlobal and
'    Scripting.FileSystemObject.
' 3) Construct one or more classes (example:)
'    Dim highScores
'    Set highScores = New vpwData
'    With highScores
'         .gameName = "tableName"
'         .dataName = "nameOfDataset"
'         .Load "keyName", "defaultValue"
'         .Load "anotherKey", 0
'         .Reset = true
'    End With
' 4) For each vpwData, call SaveAll in the table_exit
'    event:
'    highScores.SaveAll
' 5) Refer to the methods below for more info.
'---------------------------------------------------------------
' TIPS FOR CREATING / USING DATA CLASSES
' 1) Use vpwData.data.Item(keyName) to get the value of a
'    key. Use vpwData.data.Item(keyName) = newValue to set
'    a new value for a key, and vpwData.Save keyName if
'    you want to save the new value via SaveValue immediately.
'    ONLY use strings or numbers as values!
' 2) When tracking player-specific data, such as game progress,
'    you may want to use a For loop to create a dataset for
'    each player (ex. if your table can support up to 4 players,
'    create 4 datasets):
'    Dim dataPlayer(4)
'    Dim i
'    For i = 0 to 3
'        Set dataPlayer(i) = New vpwData
'        With dataPlayer(i)
'             .gameName = "thisTable"
'             .dataName = "player" & i
'            .Load "score", 0
'            .Load "modesCompleted", 0
'        End With
'    Next
' 3) WARNING: the total character length of each key may not
'    exceed (31 - character length of pDataName in Init call)
'    characters. This is a VPReg key length limitation.
' 4) WARNING: Due to VPReg limitations, only strings and
'    numbers are supported for values. Do not use booleans
'    (use 0 or 1 instead), arrays, objects, etc.
' 5) Avoid repeatedly saving data keys that frequently change;
'    you should only save for significant changes (credits
'    inserted, extra balls, etc) and when exiting the table.
'    You can also SaveAll during downtime in the game, such as
'    when the ball is held in a saucer while something is
'    playing on the DMD, or a change in player turns / balls.
'
' TUTORIAL: https://youtu.be/STAc5ykyWl4
'***************************************************************

Class vpwData
	Public data ' Scripting.Dictionary of the current key/value data
	Public gameName ' Alphanumeric name of this table
	Public dataName ' Name of this dataset, alphanumeric; MUST be set when initializing the class and MUST be unique for each dataset.
	
	Private defaults ' Dictionary of data.Keys to their default values
	Private dataReg	 ' Dictionary of data.Keys to the value loaded from the VP Registry (for tracking which ones to save in SaveAll)
	
	'----------------------------------------------------------
	' Class_Initialize
	' Class initializer
	'----------------------------------------------------------
	Public Sub Class_Initialize
		Set data = CreateObject("Scripting.Dictionary")
		Set defaults = CreateObject("Scripting.Dictionary")
		Set dataReg = CreateObject("Scripting.Dictionary")
	End Sub
	
	'----------------------------------------------------------
	' vpwData.Load
	' Load a data key into the data manager.
	'
	' You MUST set gameName and dataName first!
	'
	' PARAMETERS
	' key (String) - The name of the key to load
	' defaultValue (String|Number) - The default value to use
	' for this key when reset or when the value does not yet
	' exist in VPReg.
	'
	' NOTE: The data is loaded immediately after call into
	' vpwData.data with either the value that was saved
	' in VPReg or the specified default value. Once Save or
	' SaveAll is called, whatever value is set for this key in
	' vpwData.data will be saved to VPReg (even if it is
	' the default value! This means you will need to set the
	' value of the key to an empty string if you want to reset
	' it to the defined default value [such as if you change
	' the default] or set .reset = true on the dataset).
	'----------------------------------------------------------
	Public Sub Load(key, defaultValue)
		' Note the default value requested
		defaults.Add key, defaultValue
		
		' Try loading the saved value from VPReg
		Dim savedValue
		savedValue = LoadValue(gameName, dataName & "." & key)
		If savedValue = "" Then ' Blank / no value; use default
			data.Add key, defaultValue
			dataReg.Add key, ""
		Else
			If IsNumeric(savedValue) Then ' Convert numerical values to double
				data.Add key, CDbl(savedValue)
				dataReg.Add key, CDbl(savedValue)
			Else ' Convert non-numerical values to string
				data.Add key, CStr(savedValue)
				dataReg.Add key, CStr(savedValue)
			End If
		End If
	End Sub
	
	'----------------------------------------------------------
	' vpwData.Default
	' Get the default value of a key.
	'
	' PARAMETERS
	' Key (string) - The key of which to get the default value
	'----------------------------------------------------------
	Public Property Get Default(Key)
		Default = defaults.Item(Key)
	End Property
	
	'----------------------------------------------------------
	' vpwData.Save
	' Save a key's current value to the VPX registry via
	' SaveValue, even if we believe it has not changed.
	'
	' PARAMETERS
	' pKey (string) - The name of the key to save
	'----------------------------------------------------------
	Public Sub Save(pKey)
		Dim i
		If data.Exists(pKey) Then
			i = data.Item(pKey)
			SaveValue gameName, dataName & "." & pKey, i
			dataReg.Item(pKey) = i
		End If
	End Sub
	
	'----------------------------------------------------------
	' vpwData.SaveAll
	' Save the values of all keys which we believe have changed
	' to the VP registry (via SaveValue).
	'----------------------------------------------------------
	Public Sub SaveAll()
		Dim i
		Dim i2
		For Each i In data.Keys
			i2 = data.Item(i)
			
			'Optimize by only saving keys which we think have a different value
			If Not dataReg.Item(i) = i2 Then
				SaveValue gameName, dataName & "." & i, i2
				dataReg.Item(i) = i2
			End If
		Next
	End Sub
	
	'----------------------------------------------------------
	' vpwData.Reset = True|False
	' Reset all of the key's values back to their defaults and
	' save (if True). Should not be called until all keys have
	' been loaded via the Load routine.
	'----------------------------------------------------------
	Public Property Let Reset(doIt)
		If doIt = True Then
			Dim key
			For Each key In defaults.Keys
				data.Item(key) = defaults.Item(key)
			Next
		End If
	End Property
End Class

'**************************************************************
'   END VPW DATA MANAGER
'**************************************************************

'***************************************************************
' ZCON: VPIN WORKSHOP CONSTANTS - 1.0.0 BETA
'***************************************************************
' WHAT IS IT?
' The VPin Workshop Constants class is a quick / easy class that
' uses Scripting.Dictionary to define "constants" with types you
' cannot normally use on a Visual Basic constant (Const)
' (such as arrays or objects).
'
' HOW TO USE
' 1) Put this VBS file in your VPX Scripts folder, or
'    copy / paste the code into your table script
'    (and skip step 2).
' 2) Include this VBS file in your table via ExecuteGlobal and
'    Scripting.FileSystemObject.
' 3) Initialize the class on a variable:
'    Dim Scoring : set Scoring = new vpwConstants
' 4) Set values as you normally would with Scripting.Dictionary:
'    Scoring.Item("Bumpers") = 5000
'    -OR-
'    Scoring.Add "Bumpers", 5000
' 5) Once a key has a non-empty value, it cannot be set again.
'
' TUTORIAL: https://youtu.be/w1ZhxoNogWE
'***************************************************************

Class vpwConstants
	' Create the private dictionary
	Private mDict
	Private Sub Class_Initialize
		Set mDict = CreateObject("Scripting.Dictionary")
	End Sub
	
	' Get the count of the number of constants / keys
	Public Property Get Count
		Count = mDict.Count
	End Property
	
	' Get the value of a constant key
	Public Property Get Item(aKey)
		Item = Empty
		If mDict.Exists(aKey) Then
			If IsObject(mDict(aKey)) Then
				Set Item = mDict(aKey)
			Else
				Item = mDict(aKey)
			End If
		End If
	End Property
	
	' Set a constant value
	Public Property Let Item(aKey, aData)
		setConstant aKey, aData
	End Property
	
	Public Property Set Key(aKey)
		' This function is (and always has been) a no-op. Previous definition
		' just looked up aKey in the keys list, and if found, set the key to itself.
	End Property
	
	' Add a constant key (same as setting its value)
	Public Sub Add(aKey, aData)
		setConstant aKey, aData
	End Sub
	
	' The actualizer for setting a constant
	Private Sub setConstant(aKey, aData)
		' If the constant key does not already exist, create / set it. Otherwise, do nothing (disallow overwriting / changing constants!)
		If Not mDict.Exists(aKey) Then
			If IsObject(aData) Then
				Set mDict(aKey) = aData
			Else
				mDict(aKey) = aData
			End If
		End If
	End Sub
	
	' No remove functions because we want to disallow removing constants!
	
	Public Function Exists(aKey)
		Exists = mDict.Exists(aKey)
	End Function
	Public Function Items
		Items = mDict.Items
	End Function
	Public Function Keys
		Keys = mDict.Keys
	End Function
End Class

'**************************************************************
'   END VPWCONSTANTS
'**************************************************************

'***************************************************************
' ZCLK: VPIN WORKSHOP CLOCKS - 0.2.1 ALPHA
'***************************************************************
' WHAT IS IT?
' The VPin Workshop Clocks class is a simple utility class for
' tracking time-based operations in the table with GameTime
' precision. It can also be used to fire a callback when a
' clock expires. And clocks can be paused / unpaused.
'
' This is intended for recurring clocks, such as mode times,
' ball saves, and so on. It is not intended for temporary
' or one-off timers which would be better suited using either
' VPX timers, vpmTimer, or vpwQueueManager (VPW queue system).
'
' HOW TO USE
' 1) Initialize the class on a variable:
'    Dim Clocks : set Clocks = new vpwClocks
' 2) Use a VPX timer to tick the clocks with vpwClocks.Tick.
'    You can use any interval, but the lower the interval, the
'    more precise the clocks will expire.
' 3) Add your clocks via vpwClocks.Add.
' 4) Refer to the vpwClock for how to control your clock.
'    Control your clock by changing its properties via
'    vpwClocks.data.Item(clockName).
'
' TUTORIAL: https://youtu.be/QvYl0P09Uw4
'***************************************************************
Class vpwClocks
	Public data ' Dictionary of clocks we are using
	Public lastTick ' GameTime which Tick was last called
	
	Private Sub Class_Initialize
		lastTick = GameTime
		Set data = CreateObject("Scripting.Dictionary")
	End Sub
	
	'----------------------------------------------------------
	' vpwClocks.Add
	' Register a new clock.
	'
	' PARAMETERS:
	' cName (String) - The unique name for this clock.
	' Caution! If you use the name of a clock that already
	' exists, it will be overwritten!
	'
	' After adding a clock, use vpwClocks.data.Item(cName)
	' to set properties on the clock (see vpwClock class).
	'----------------------------------------------------------
	Public Sub Add(cName)
		' Remove duplicates
		If data.Exists(cName) Then data.Remove cName
		
		' Add our new clock
		Dim i
		Set i = New vpwClock
		data.Add cName, i
	End Sub
	
	Public Sub Tick()
		If data.Count = 0 Then Exit Sub ' Nothing to Do
		
		' Tick every clock
		Dim k, key, item
		k = data.Keys
		For Each key In k
			Set item = data.Item(key)
			item.Tick Int(GameTime - lastTick)
		Next
		
		' Set when we last ticked
		lastTick = GameTime
	End Sub
End Class

Class vpwClock ' An individual clock.
	Public timeLeft ' Integer - The number of milliseconds until the clock expires (0 = currently expired)
	Public isPaused ' Boolean - Whether or not the clock Is paused (timeLeft will not decrease when True)
	Public canExpire ' Boolean - Whether or not the clock can actually expire right now (if false, timeLeft will never go below 1 until this is true)
	
	Private c_expiryCallback
	Private c_tickCallback
	
	Private Sub Class_Initialize
		timeLeft = 0
		isPaused = False
		canExpire = True
		c_expiryCallback = Null
	End Sub
	
	Public Property Get expiryCallback() ' String|Object|Null - The expiry callback for this clock
		If IsObject(c_expiryCallback) Then
			Set expiryCallback = c_expiryCallback
		Else
			expiryCallback = c_expiryCallback
		End If
	End Property
	
	'----------------------------------------------------------
	' Let|Set vpwClock.expiryCallback
	' Assign a callback to be fired when the clock expires.
	'
	' PARAMETERS:
	' callback (String|Object|Null) - Either a string of the
	' routine to fire, an Object of the routine, or Null if
	' no callback should be fired when this clock expires.
	'----------------------------------------------------------
	Public Property Let expiryCallback(callback)
		If IsObject(callback) Then
			Set c_expiryCallback = callback
		Else
			c_expiryCallback = callback
		End If
	End Property
	
	Public Property Set expiryCallback(callback)
		Set c_expiryCallback = callback
	End Property
	
	Public Property Get tickCallback() ' String|Null - The tick callback for this clock
		tickCallback = c_tickCallback
	End Property
	
	'----------------------------------------------------------
	' Let vpwClock.tickCallback
	' Assign a callback to be fired when the clock ticks.
	' NOTE: Will stop firing when clock expires, is paused, or
	' out of time but cannot yet expire. Tick callback fires
	' before expiry callback.
	'
	' PARAMETERS:
	' callback (String|Null) - Either a string of the
	' routine to fire or Null if no callback should be fired
	' when this clock ticks.
	'
	' WARNING! Should be specified as the name of a sub
	' without any parameters in the string; the following
	' parameters are passed automatically into the sub:
	' timeLeft (number) - Time left in milliseconds
	' timeElapsed (number) - Milliseconds that elapsed since
	' the previous tick
	'----------------------------------------------------------
	Public Property Let tickCallback(callback)
		c_tickCallback = callback
	End Property
	
	Public Sub Tick(milli) ' Tick this clock and handle expiration
		' Do not tick if expired or paused
		If timeLeft <= 0 Or isPaused = True Then Exit Sub
		
		' Tick down the specified number of milliseconds
		timeLeft = timeLeft - Int(milli)
		If timeLeft < 0 Then
			milli = Int(milli) + timeLeft
			timeLeft = 0
		End If
		
		' Did the clock expire?
		If timeLeft <= 0 Then
			' The clock can expire, so expire it
			If canExpire = True Then
				timeLeft = 0
				
				'Fire tick callback if applicable
				If VarType(c_tickCallback) = vbString Then ExecuteGlobal c_tickCallback & " " & timeLeft & ", " & milli
				
				' Fire callback if applicable
				If IsObject(c_expiryCallback) Then
					Call c_expiryCallback(0)
				ElseIf VarType(c_expiryCallback) = vbString Then
					ExecuteGlobal c_expiryCallback
				End If
			Else
				timeLeft = 1 ' clock cannot expire yet
			End If
		'Else
			'Fire tick callback if applicable
			'If VarType(c_tickCallback) = vbString Then ExecuteGlobal c_tickCallback & " " & timeLeft & ", " & milli
		End If
	End Sub
End Class

'**************************************************************
'   END VPWCLOCKS
'**************************************************************

'***************************************************************
' ZMUS: VPIN WORKSHOP MUSIC - 0.2.3 ALPHA
'***************************************************************
' WHAT IS IT?
' The VPin Workshop Music class is a small wrapper around the
' VPin Workshop Tracks class. It allows managing a group of
' sound tracks as music tracks with automatic fading and
' stopping of previous music tracks when starting a new one.
'
' REQUIRES vpwTracks and vpwTrack
'
' HOW TO USE
' 1) Add vpwTracks and vpwTrack scripts to your table. You do
'    not need to follow set-up instructions beyond including 
'    the script unless you will be using fading on individual
'    non-music sounds as well.
' 2) Put this VBS file in your VPX Scripts folder, or
'    copy / paste the code into your table script
'    (and skip step 3).
' 3) Include this VBS file in your table via ExecuteGlobal and
'    Scripting.FileSystemObject.
' 4) Initialize the class on a variable:
'    	Dim Music : set Music = new vpwMusic
' 5) Use a VPX timer to tick the fade with vpwMusic.Tick.
'    You can use any interval, but the lower the interval, the
'    more smooth the fading. If you already have a timer for
'    vpwTracks, it is highly recommend to just put
'    vpwMusic.Tick in that timer's timer event instead of
'    creating another timer.
' 6) Add your music tracks to the group via vpwMusic.Add.
' 7) Refer to the rest of this class for instructions.
'
' Note: You can initialize multiple classes for multiple
' groups of music, but normally you would not do this as you
' would probably only ever want one music track playing at a
' time.
'
' TUTORIAL: https://youtu.be/C0J1LT5-wqA
'***************************************************************

Class vpwMusic
	Public trackManager		'The vpwTracks class to use for managing the music
	Public nowPlaying 		'The name of the music track currently playing or last played, or null if nothing is playing
	Private data 			'Dictionary of music: key is the name of the sound in VPX, value is An array of default volume, fadeIn, fadeOut.
	
	Private Sub Class_Initialize
		Set trackManager = New vpwTracks
		Set data = CreateObject("Scripting.Dictionary")
		nowPlaying = Null
	End Sub
	
	'----------------------------------------------------------
	' vpwMusic.Add
	' Add a sound to this music manager. This means it will be
	' stopped when any other music track in the same class is
	' started, or when StopTrack is called.
	'
	' PARAMETERS:
	' mName (String) - The name of the sound in VPX's sound
	' manager.
	' loopCount (Integer) - The number of times this music
	' track should play (at maximum). Use -1 for unlimited.
	' volume (float) - The default volume to use (0 - 1).
	' fadeIn (Integer) - The number of milliseconds to fade In
	' this track when starting to play it.
	' fadeOut (Integer) - The number of milliseconds to fade
	' Out this track when stopping it.
	'----------------------------------------------------------
	Public Sub Add(mName, loopCount, volume, fadeIn, fadeOut)
		data.Item(mName) = Array(loopCount, volume, fadeIn, fadeOut)
	End Sub
	
	'----------------------------------------------------------
	' vpwMusic.PlayTrack
	' Begin playing a music track while stopping the currently
	' playing track (if applicable).
	'
	' PARAMETERS:
	' mName (String) - The name of the sound to play.
	' loopCount (Integer) - The number of times this music
	' track should play. Use -1 for unlimited. Null = default.
	' volume (float) - The volume to use (0 - 1). Null = use
	' default.
	' fade (Integer) - The fade duration in milliseconds
	' for this track. Null = use default.
	' fadeOut (Integer) - The fade out duration in milliseconds
	' for the currently playing track. Null = use default.
	' Pan (float) - ranges from -1.0 (left) over 0.0 (both) to
	' 1.0 (right).
	' randomPitch (float) - ranges from 0.0 (no randomization)
	' to 1.0 (vary between half speed to double speed).
	' pitch (integer) - can be positive or negative and directly
	' adds onto the standard sample frequency.
	' frontRearFade (float) - similar to pan but fades between
	' the front and rear speakers.
	'----------------------------------------------------------
	Public Sub PlayTrack(mName, loopCount, volume, fade, fadeOut, pan, randomPitch, pitch, frontRearFade)
		If Not data.Exists(mName) Then Exit Sub
		Dim i: i = data.Item(mName)
		' Determine our volumes and fading
		Dim loopCountToUse, volumeToUse, fadeToUse, fadeOutToUse
		If IsNull(loopCount) Then
			loopCountToUse = i(0)
		Else
			loopCountToUse = loopCount
		End If
		If IsNull(volume) Then
			volumeToUse = i(1)
		Else
			volumeToUse = volume
		End If
		If IsNull(fade) Then
			fadeToUse = i(2)
		Else
			fadeToUse = fade
		End If
		
		If Not IsNull(nowPlaying) Then
			'Stop / fade out the current track if it is not the same as the track we are requesting
			If Not mName = nowPlaying Then
				If IsNull(fadeOut) Then
					Dim np: np = data.Item(nowPlaying)
					fadeOutToUse = np(3)
				Else
					fadeOutToUse = fadeOut
				End If
				
				If fadeOutToUse = 0 Then ' No fade-out, stop immediately
					trackManager.StopTrack nowPlaying
				Else
					trackManager.AddFade nowPlaying, "volume", 0, fadeOutToUse
				End If
			End If
		End If
		
		nowPlaying = mName
		
		' Start the sound
		If fadeToUse = 0 Then ' No fade-in, start immediately at set volume
			trackManager.PlayTrack mName, loopCountToUse, volumeToUse, pan, randomPitch, pitch, True, True, frontRearFade
		Else
			trackManager.PlayTrack mName, loopCountToUse, 0, pan, randomPitch, pitch, True, True, frontRearFade
			trackManager.AddFade mName, "volume", volumeToUse, fadeToUse
		End If
	End Sub
	
	'----------------------------------------------------------
	' vpwMusic.StopTrack
	' Stop the currently-playing track. Use PlayTrack if you
	' are immediately starting a different track.
	'
	' PARAMETERS:
	' fadeOut (Integer) - The fade out duration in milliseconds
	' for the currently playing track. Null = use default.
	'----------------------------------------------------------
	Public Sub StopTrack(fadeOut)
		If IsNull(nowPlaying) Then Exit Sub ' Nothing to stop
		
		' Get our fade-out duration
		Dim fadeOutToUse
		If IsNull(fadeOut) Then
			Dim np: np = data.Item(nowPlaying)
			fadeOutToUse = np(2)
		Else
			fadeOutToUse = fadeOut
		End If
		
		' Stop the sound
		If fadeOutToUse = 0 Then ' No fade-out, stop immediately
			trackManager.StopTrack nowPlaying
		Else
			trackManager.AddFade nowPlaying, "volume", 0, fadeOutToUse
		End If
		
		nowPlaying = Null
	End Sub
	
	Public Sub Tick()
		trackManager.Tick
	End Sub
End Class

'**************************************************************
'   END VPWMUSIC
'**************************************************************

'***************************************************************
' ZTRK: VPIN WORKSHOP DYNAMIC SOUND FADER - 0.1.2 ALPHA
'***************************************************************
' WHAT IS IT?
' The VPin Workshop Dynamic Sound Fader is a utility class
' for managing sound played through VPX's sound manager
' library. This class allows for easy management of track
' fading, pitching, and/or panning over a period of time.
'
' HOW TO USE
' 1) Put this VBS file in your VPX Scripts folder, or
'    copy / paste the code into your table script
'    (and skip step 2).
' 2) Include this VBS file in your table via ExecuteGlobal and
'    Scripting.FileSystemObject.
' 3) Initialize the class on a variable:
'    Dim Tracks : set Tracks = new vpwTracks
' 4) Use a VPX timer to tick the fade with vpwTracks.Tick.
'    You can use any interval, but the lower the interval, the
'    more smooth the fading.
' 5) Instead of using VPX's PlaySound or PlayMusic, use
'    With vpwTracks
'        .PlayTrack "name_of_sound", loopCount, initialVolume, initialPan, randomPitch, initialPitch, useExisting, restart, initialFrontRearFade (essentially, same parameters as PlaySound, except all are required)
'        .AddFade "name_of_sound", "volume", targetVolume, fadeTimeMilliseconds (for fading the volume. You can also use pan, pitch, or frontRearFade. Note that if you used randomPitch then pitch fading will NOT work correctly.)
'        (You can add any number of additional AddFades here,
'        or even add more AddFades later so long as the track
'        is still playing. If it's not playing, you should call
'        PlayTrack again first. Note that you cannot overlap
'        multiple fades of the same type [param 2]; they will
'        overwrite each other if the previous did not yet finish)
'    End With
' 6) Refer to the rest of this class for instructions.
'
' Consider also using vpwMusic to manage your music.
'
' Note: For sounds for which you will never use dynamic
' fading, use PlaySound / StopSound instead of vpwTracks.
' Using tracks in vpwTracks which will never be faded will
' increase the performance / resource use of vpwTracks
' for no good reason.
'***************************************************************

Class vpwTracks
	Public data             'A dictionary of VPX sound tracks currently being processed (key is the sound track name in VPX, value is the vpwTrack class)
	Private lastTick        'GameTime of when Tick was last called
	
	Private Sub Class_Initialize
		Set data = CreateObject("Scripting.Dictionary")
		lastTick = GameTime
	End Sub
	
	Private Function GetOrCreateTrack(trackName) 'Find or create a VPX sound track
		If IsNull(trackName) Then Exit Function 'Might be used with vpwMusic nowPlaying, which could be null

		If data.Exists(trackName) Then
			Set GetOrCreateTrack = data.Item(trackName)
			Exit Function
		End If
		
		Dim track
		Set track = new vpwTrack
		track.name = trackName
		Set data.Item(trackName) = track
		Set GetOrCreateTrack = data.Item(trackName)
	End Function
	
	'----------------------------------------------------------
	' vpwTracks.PlayTrack
	' Replacement for PlaySound; play a VPX sound.
	'
	' PARAMETERS:
	' trackName (String) - The name of the VPX sound to play.
	' loopCount (Integer) - The number of times to play the
	' sound. Use -1 to play infinitely until stopped.
	' volume (float) - The initial volume to use (0 - 1).
	' Use null to maintain current volume if played before.
	' Pan (float) - ranges from -1.0 (left) over 0.0 (both) to
	' 1.0 (right). Use null to maintain current pan if
	' played before.
	' randomPitch (float) - ranges from 0.0 (no randomization)
	' to 1.0 (vary between half speed to double speed).
	' pitch (integer) - can be positive or negative and directly
	' adds onto the standard sample frequency. Use null for the
	' current pitch if played before.
	' useExisting (Boolean) - Instead of playing the sound on a
	' new channel, use the existing channel if it is currently
	' playing.
	' restart (Boolean) - Start the sound over from the
	' beginning.
	' frontRearFade (float) - similar to pan but fades between
	' the front and rear speakers. Use null to maintain the
	' current fade if played before.
	'
	' WARNING: Use of AddFade may break randomPitch if used.
	'
	' Note that all parameters are required even though that
	' is not the case for PlaySound.
	'----------------------------------------------------------
	Public Sub PlayTrack(trackName, loopCount, volume, pan, randomPitch, pitch, useExisting, restart, frontRearFade)
		If IsNull(trackName) Then Exit Sub 'Might be used with vpwMusic nowPlaying, which could be null

		Dim track
		Set track = GetOrCreateTrack(trackName)
		track.PlayTrack loopCount, volume, pan, randomPitch, pitch, useExisting, restart, frontRearFade
	End Sub
	
	'----------------------------------------------------------
	' vpwTracks.AddFade
	' Add a fade operation to a track that is playing.
	' Note that fade operations run in parallel, and specifying
	' the same fadeName before a previous one finishes will
	' overwrite it.
	' If a track is not playing, you should call PlayTrack
	' first or the track might not sound right when it starts.
	' Fading might break randomPitch if it was used!
	'
	' PARAMETERS:
	' trackName (String) - Name of sound in VPX to add the
	' fade operation.
	' fadeName (String) - volume, pan, pitch, or frontRearFade.
	' targetValue (float) - The [fadeName] of the track should
	' gradually change from its current value to the
	' targetValue over the specified duration period.
	' duration (Integer) - Number of milliseconds the fade
	' should take.
	'----------------------------------------------------------
	Public Sub AddFade(trackName, fadeName, targetValue, duration)
		If IsNull(trackName) Then Exit Sub 'Might be used with vpwMusic nowPlaying, which could be null
		
		Dim track
		Set track = GetOrCreateTrack(trackName)
		track.AddFade fadeName, targetValue, duration
	End Sub
	
	'----------------------------------------------------------
	' vpwTracks.StopTrack
	' Immediately stop playing a track.
	'
	' PARAMETERS:
	' trackName (String) - Name of sound in VPX to stop.
	'----------------------------------------------------------
	Public Sub StopTrack(trackName)
		If Not data.Exists(trackName) Then Exit Sub
		data.Item(trackName).StopTrack
	End Sub
	
	'----------------------------------------------------------
	' vpwTracks.Tick
	' Tick each track (which re-processes fading levels).
	'----------------------------------------------------------
	Public Sub Tick()
		' Calculate time since last call
		Dim timeElapsed
		timeElapsed = GameTime - lastTick
		lastTick = GameTime
		
		' Do nothing if there are no processing tracks
		If data.Count <= 0 Then Exit Sub
		
		' Run the Tick operation on each track
		Dim k, key
		k = data.Keys
		For Each key In k
			data.Item(key).Tick timeElapsed
		Next
	End Sub
End Class

Class vpwTrack ' Do not use directly; use vpwTracks instead!
	Public name                     'The name of the sound in the VPX sound manager
	Public tLoopCount               'The number of times to loop
	Public tVolume                  'The current volume of the track (float, 0 - 1)
	Public tPan                     'pan ranges from -1.0 (left) over 0.0 (both) to 1.0 (right)
	Public tPitch                   'can be positive or negative and directly adds onto the standard sample frequency
	Public tFrontRearFade           'similar to pan but fades between the front and rear speakers
	
	Public data 'Dictionary of active fading operations. Key is operation; value is array (initialValue, targetValue, totalTime, timeLeft, stopSoundWhenDone)
	
	Private Sub Class_Initialize ' You MUST set name (String) after initialization!
		name = Null
		tVolume = 0.0
		tLoopCount = 1
		tPan = 0.0
		tPitch = 0.0
		tFrontRearFade = 0.0
		Set data = CreateObject("Scripting.Dictionary")
	End Sub
	
	Public Sub PlayTrack(loopCount, volume, pan, randomPitch, pitch, useExisting, restart, frontRearFade) ' Play the track, making note of its parameters
		' Determine what volume, loopCount, pan, pitch, and frontRearFade to use; and note their values in the class.
		Dim pVolume
		If IsNull(volume) Then
			pVolume = tVolume
		Else
			pVolume = volume
			tVolume = volume
		End If
		Dim pLoopCount
		If IsNull(loopCount) Then
			pLoopCount = tLoopCount
		Else
			pLoopCount = loopCount
			tLoopCount = loopCount
		End If
		Dim pPan
		If IsNull(pan) Then
			pPan = tPan
		Else
			pPan = pan
			tPan = pan
		End If
		Dim pPitch
		If IsNull(pitch) Then
			pPitch = tPitch
		Else
			If pitch = 0 Then pitch = 1 'We treat a pitch of 0 as "use original pitch", however VPX treats 0 as "maintain current pitch"
			pPitch = pitch
			tPitch = pitch
		End If

		Dim pFrontRearFade
		If IsNull(frontRearFade) Then
			pFrontRearFade = tFrontRearFade
		Else
			pFrontRearFade = frontRearFade
			tFrontRearFade = frontRearFade
		End If
		
		' Play the sound
		PlaySound name, pLoopCount, pVolume, pPan, randomPitch, pPitch, useExisting, restart, pfrontRearFade
	End Sub
	
	Public Sub AddFade(fadeName, targetValue, duration) ' Add a fade operation to the track
		' Note: When adding a fade operation, we are overwriting any existing operations of the same fade type.
		' We use the current volume/pan/pitch/fade as the initial value so when overwrites happen, the fade will start at the value left off
		' to prevent "jumping".
		Select Case fadeName
			Case "volume":
				data.Item(fadeName) = Array(tVolume, targetValue, 0, duration)
				
			Case "pan":
				data.Item(fadeName) = Array(tPan, targetValue, 0, duration)
				
			Case "pitch":
				data.Item(fadeName) = Array(tPitch, targetValue, 0, duration)
				
			Case "frontRearFade":
				data.Item(fadeName) = Array(tFrontRearFade, targetValue, 0, duration)
		End Select
	End Sub
	
	Public Sub Tick(timeElapsed) ' Process fading on the track
		If data.Count = 0 Then Exit Sub
		
		' loop through each fade operation
		Dim k, key, fadeValue, totalTimeElapsed, timeRemaining, valueDifference, percentProgress, newValue, reTriggerPlaySound
		k = data.Keys
		reTriggerPlaySound = False
		For Each key In k
			Dim i: i = data.Item(key)
			' Determine the total time that has elapsed, and the time remaining, for the fade operation
			totalTimeElapsed = i(2) + timeElapsed
			timeRemaining = i(3) - totalTimeElapsed
			If timeRemaining < 0 Then timeRemaining = 0
			If totalTimeElapsed > i(3) Then totalTimeElapsed = i(3)
			
			' Determine the difference between initial and target value, the percent complete from time elapsed to total fade time, and thus what value we should be using for the fade right now.
			valueDifference = Abs(i(0) - i(1))
			If i(3) <= 0 Then ' Division by zero / negative duration protection
				percentProgress = 1
			Else
				percentProgress = totalTimeElapsed / i(3)
			End If
			If i(0) > i(1) Then ' Decreasing the value over time
				newValue = i(0) - (valueDifference * percentProgress)
			Else ' Increasing the value over time
				newValue = i(0) + (valueDifference * percentProgress)
			End If
			
			' Save our new parameters
			data.Item(key) = Array(i(0), i(1), totalTimeElapsed, i(3))
			
			' Update the track's audio values in the class, and mark that we must call PlaySound
			Select Case key
				Case "volume":
					tVolume = newValue
					reTriggerPlaySound = True
					
				Case "pan":
					tPan = newValue
					reTriggerPlaySound = True
					
				Case "pitch":
					tPitch = newValue
					If tPitch = 0 Then tPitch = 1 'We treat a pitch of 0 as "use original pitch", however VPX treats 0 as "maintain current pitch"
					reTriggerPlaySound = True
					
				Case "frontRearFade":
					tFrontRearFade = newValue
					reTriggerPlaySound = True
			End Select
			
			'Debug.print name & ": " & key & ": " & timeRemaining & " / " & valueDifference & " / " & percentProgress & " / " & newValue
			
			' Remove the fade operation if it is done
			If timeRemaining <= 0 Then data.Remove key
		Next
		
		' Call PlaySound with the new values, re-using the current channel without restarting it
		If reTriggerPlaySound = True Then PlaySound name, tLoopCount, tVolume, tPan, 0, tPitch, True, False, tFrontRearFade
		
		' We want To actually stop playing the track if the volume is 0 and there are no more fade operations
		If tVolume <= 0 And data.Count <= 0 Then 
			StopSound name
		End If
	End Sub
	
	Public Sub StopTrack() ' Immediately stop playing the track, clear out all fade operations, and set volume To 0
		data.RemoveAll
		StopSound name
		tVolume = 0.0
	End Sub
End Class

'**************************************************************
'   END VPWTRACKS
'**************************************************************

'******************************************************
' ZLAM:  LAMPZ by nFozzy
'
' 2021.07.01 Added modulated flashers
'******************************************************
'
' Lampz is a utility designed to manage and fade the lights and light-related objects on a table that is being driven by a ROM.
' To set up Lampz, one must populate the Lampz.MassAssign array with VPX Light objects, where the index of the MassAssign array
' corrisponds to the ROM index of the associated light. More that one Light object can be associated with a single MassAssign index (not shown in this example)
' Optionally, callbacks can be assigned for each index using the Lampz.Callback array. This is very useful for allowing 3D Insert primitives
' to be controlled by the ROM. Note, the aLvl parameter (i.e. the fading level that ranges between 0 and 1) is appended to the callback call.

'Helper functions

Function ColtoArray(aDict)	'converts a collection To an indexed array. Indexes will come out random probably.
	ReDim a(999)
	Dim count
	count = 0
	Dim x
	For Each x In aDict
		Set a(Count) = x
		count = count + 1
	Next
	ReDim preserve a(count-1)
	ColtoArray = a
End Function

'**********************************************************************
'Class jungle nf
'**********************************************************************

'No-op object instead of adding more conditionals to the main loop
'It also prevents errors if empty lamp numbers are called, and it's only one object
'should be g2g?

Class NullFadingObject
	Public Property Let IntensityScale(input)
		
	End Property
End Class

'version 0.11 - Mass Assign, Changed modulate style
'version 0.12 - Update2 (single -1 timer update) update method for core.vbs
'Version 0.12a - Filter can now be accessed via 'FilterOut'
'Version 0.12b - Changed MassAssign from a sub to an indexed property (new syntax: lampfader.MassAssign(15) = Light1 )
'Version 0.13 - No longer requires setlocale. Callback() can be assigned multiple times per index
' Note: if using multiple 'LampFader' objects, set the 'name' variable to avoid conflicts with callbacks
'Version 0.14 - Updated to support modulated signals - Niwak
'Version 0.15 - Added IsLight property - apophis

Class LampFader
	Public IsLight(150)
	Public FadeSpeedDown(150), FadeSpeedUp(150)
	Private Lock(150), Loaded(150), OnOff(150)
	Public UseFunc
	Private cFilter
	Public UseCallback(150), cCallback(150)
	Public Lvl(150), Obj(150)
	Private Mult(150)
	Public FrameTime
	Private InitFrame
	Public Name
	
	Sub Class_Initialize()
		InitFrame = 0
		Dim x
		For x = 0 To UBound(OnOff)	 'Set up fade speeds
			FadeSpeedDown(x) = 1 / 100	'fade speed down
			FadeSpeedUp(x) = 1 / 80		'Fade speed up
			UseFunc = False
			lvl(x) = 0
			OnOff(x) = 0
			Lock(x) = True
			Loaded(x) = False
			Mult(x) = 1
			IsLight(x) = False
		Next
		Name = "LampFaderNF" 'NEEDS TO BE CHANGED IF THERE'S MULTIPLE OF THESE OBJECTS, OTHERWISE CALLBACKS WILL INTERFERE WITH EACH OTHER!!
		For x = 0 To UBound(OnOff)		 'clear out empty obj
			If IsEmpty(obj(x) ) Then Set Obj(x) = NullFader' : Loaded(x) = True
		Next
	End Sub
	
	Public Property Get Locked(idx)
		Locked = Lock(idx)
		'   debug.print Lampz.Locked(100)	'debug
	End Property
	
	Public Property Get state(idx)
		state = OnOff(idx)
	End Property
	
	Public Property Let Filter(String)
		Set cFilter = GetRef(String)
		UseFunc = True
	End Property
	
	Public Function FilterOut(aInput)
		If UseFunc Then
			FilterOut = cFilter(aInput)
		Else
			FilterOut = aInput
		End If
	End Function
	
	'   Public Property Let Callback(idx, String)
	'	   cCallback(idx) = String
	'	   UseCallBack(idx) = True
	'   End Property
	
	Public Property Let Callback(idx, String)
		UseCallBack(idx) = True
		'   cCallback(idx) = String 'old execute method
		
		'New method: build wrapper subs using ExecuteGlobal, then call them
		cCallback(idx) = cCallback(idx) & "___" & String	'multiple strings dilineated by 3x _
		
		Dim tmp
		tmp = Split(cCallback(idx), "___")
		
		Dim str, x
		For x = 0 To UBound(tmp)	'build proc contents
			'If Not tmp(x)="" then str = str & "	" & tmp(x) & " aLVL" & "	'" & x & vbnewline	'more verbose
			If Not tmp(x) = "" Then str = str & tmp(x) & " aLVL:"
		Next
		'   msgbox "Sub " & name & idx & "(aLvl):" & str & "End Sub"
		Dim out
		out = "Sub " & name & idx & "(aLvl):" & str & "End Sub"
		ExecuteGlobal Out
	End Property
	
	Public Property Let state(ByVal idx, input) 'Major update path
		If TypeName(input) <> "Double" And TypeName(input) <> "Integer"  And TypeName(input) <> "Long" Then
			If input Then
				input = 1
			Else
				input = 0
			End If
		End If
		If Input <> OnOff(idx) Then  'discard redundant updates
			OnOff(idx) = input
			Lock(idx) = False
			Loaded(idx) = False
		End If
	End Property
	
	'Sub MassAssign(aIdx, aInput)
	Public Property Let MassAssign(aIdx, aInput) 'Mass assign, Builds arrays where necessary
		If TypeName(obj(aIdx)) = "NullFadingObject" Then 'if empty, use Set
			If IsArray(aInput) Then
				obj(aIdx) = aInput
			Else
				Set obj(aIdx) = aInput
				If TypeName(aInput) = "Light" Then
					IsLight(aIdx) = True
				End If
			End If
		Else
			Obj(aIdx) = AppendArray(obj(aIdx), aInput)
		End If
	End Property
	
	Sub SetLamp(aIdx, aOn)  'If obj contains any light objects, set their states to 1 (Fading is our job!)
		state(aIdx) = aOn
	End Sub
	
	Public Sub TurnOnStates() 'turn state to 1
		Dim debugstr
		Dim idx
		For idx = 0 To UBound(obj)
			If IsArray(obj(idx)) Then
				'debugstr = debugstr & "array found at " & idx & "..."
				Dim x, tmp
				tmp = obj(idx) 'set tmp to array in order to access it
				For x = 0 To UBound(tmp)
					If TypeName(tmp(x)) = "Light" Then DisableState tmp(x)' : debugstr = debugstr & tmp(x).name & " state'd" & vbnewline
					tmp(x).intensityscale = 0.001 ' this can prevent init stuttering
				Next
			Else
				If TypeName(obj(idx)) = "Light" Then DisableState obj(idx)' : debugstr = debugstr & obj(idx).name & " state'd (not array)" & vbnewline
				obj(idx).intensityscale = 0.001 ' this can prevent init stuttering
			End If
		Next
		'   debug.print debugstr
	End Sub
	
	Private Sub DisableState(ByRef aObj)
		aObj.FadeSpeedUp = 1000
		aObj.State = 1
	End Sub
	
	Public Sub Init() 'Just runs TurnOnStates right now
		TurnOnStates
	End Sub
	
	Public Property Let Modulate(aIdx, aCoef)
		Mult(aIdx) = aCoef
		Lock(aIdx) = False
		Loaded(aIdx) = False
	End Property
	Public Property Get Modulate(aIdx)
		Modulate = Mult(aIdx)
	End Property
	
	Public Sub Update1() 'Handle all boolean numeric fading. If done fading, Lock(x) = True. Update on a '1' interval Timer!
		Dim x
		For x = 0 To UBound(OnOff)
			If Not Lock(x) Then 'and not Loaded(x) then
				If OnOff(x) > 0 Then 'Fade Up
					Lvl(x) = Lvl(x) + FadeSpeedUp(x)
					If Lvl(x) >= OnOff(x) Then
						Lvl(x) = OnOff(x)
						Lock(x) = True
					End If
				Else 'fade down
					Lvl(x) = Lvl(x) - FadeSpeedDown(x)
					If Lvl(x) <= 0 Then
						Lvl(x) = 0
						Lock(x) = True
					End If
				End If
			End If
		Next
	End Sub
	
	Public Sub Update2() 'Both updates on -1 timer (Lowest latency, but less accurate fading at 60fps vsync)
		FrameTime = GameTime - InitFrame
		InitFrame = GameTime	'Calculate frametime
		Dim x
		For x = 0 To UBound(OnOff)
			If Not Lock(x) Then 'and not Loaded(x) then
				If OnOff(x) > 0 Then 'Fade Up
					Lvl(x) = Lvl(x) + FadeSpeedUp(x) * FrameTime
					If Lvl(x) >= OnOff(x) Then
						Lvl(x) = OnOff(x)
						Lock(x) = True
					End If
				Else 'fade down
					Lvl(x) = Lvl(x) - FadeSpeedDown(x) * FrameTime
					If Lvl(x) <= 0 Then
						Lvl(x) = 0
						Lock(x) = True
					End If
				End If
			End If
		Next
		Update
	End Sub
	
	Public Sub Update() 'Handle object updates. Update on a -1 Timer! If done fading, loaded(x) = True
		Dim x,xx, aLvl
		For x = 0 To UBound(OnOff)
			If Not Loaded(x) Then
				aLvl = Lvl(x) * Mult(x)
				If IsArray(obj(x) ) Then	'if array
					If UseFunc Then
						For Each xx In obj(x)
							xx.IntensityScale = cFilter(aLvl)
						Next
					Else
						For Each xx In obj(x)
							xx.IntensityScale = aLvl
						Next
					End If
				Else						'if single lamp or flasher
					If UseFunc Then
						obj(x).Intensityscale = cFilter(aLvl)
					Else
						obj(x).Intensityscale = aLvl
					End If
				End If
				'   if TypeName(lvl(x)) <> "Double" and typename(lvl(x)) <> "Integer" and typename(lvl(x)) <> "Long" then msgbox "uhh " & 2 & " = " & lvl(x)
				'   If UseCallBack(x) then execute cCallback(x) & " " & (Lvl(x))	'Callback
				If UseCallBack(x) Then Proc name & x,aLvl	'Proc
				If Lock(x) Then
					If Lvl(x) = OnOff(x) Or Lvl(x) = 0 Then Loaded(x) = True	'finished fading
				End If
			End If
		Next
	End Sub
End Class

'Lamp Filter
Function LampFilter(aLvl)
	LampFilter = aLvl^1.6	'exponential curve?
End Function

'Helper functions
Sub Proc(String, Callback)	'proc using a String and one argument
	Dim p
	Set P = GetRef(String)
	P Callback
	If err.number = 13 Then  MsgBox "Proc error! No such procedure: " & vbNewLine & String
	If err.number = 424 Then MsgBox "Proc error! No such Object"
End Sub

Function AppendArray(ByVal aArray, aInput)	'append one value, object, or Array onto the End of a 1 dimensional array
	If IsArray(aInput) Then 'Input is an array...
		Dim tmp
		tmp = aArray
		If Not IsArray(aArray) Then	'if Not array, create an array
			tmp = aInput
		Else						'Append existing array with aInput array
			ReDim Preserve tmp(UBound(aArray) + UBound(aInput)+1)	'If existing array, increase bounds by UBound of incoming array
			Dim x
			For x = 0 To UBound(aInput)
				If IsObject(aInput(x)) Then
					Set tmp(x+UBound(aArray)+1 ) = aInput(x)
				Else
					tmp(x+UBound(aArray)+1 ) = aInput(x)
				End If
			Next
			AppendArray = tmp	 'return new array
		End If
	Else 'Input is NOT an array...
		If Not IsArray(aArray) Then	'if Not array, create an array
			aArray = Array(aArray, aInput)
		Else
			ReDim Preserve aArray(UBound(aArray)+1)	'If array, increase bounds by 1
			If IsObject(aInput) Then
				Set aArray(UBound(aArray)) = aInput
			Else
				aArray(UBound(aArray)) = aInput
			End If
		End If
		AppendArray = aArray 'return new array
	End If
End Function

Function InArray(arr, target) 'Determine if a value exists within an array
    Dim found
	Dim item
    found = False

    For Each item In arr
        If item = target Then
            found = True
            Exit For
        End If
    Next

    InArray = found
End Function
'******************************************************
'****  END LAMPZ
'******************************************************

'ZLSC
'Do NOT use v 9.0.0 as it is all kinds of broken!

'***********************************************************************************************************************
' Lights State Controller - 0.8.4
'  
' A light state controller for original vpx tables.
'
' Documentation: https://github.com/mpcarr/vpx-light-controller
'
'***********************************************************************************************************************

Class LStateController

    Private m_currentFrameState, m_on, m_off, m_seqRunners, m_lights, m_seqs, m_vpxLightSyncRunning, m_vpxLightSyncClear, m_vpxLightSyncCollection, m_tableSeqColor, m_tableSeqFadeUp, m_tableSeqFadeDown, m_frametime, m_initFrameTime, m_pulse, m_pulseInterval, useVpxLights, m_lightmaps, m_seqOverrideRunners

    Private Sub Class_Initialize()
        Set m_lights = CreateObject("Scripting.Dictionary")
        Set m_on = CreateObject("Scripting.Dictionary")
        Set m_off = CreateObject("Scripting.Dictionary")
        Set m_seqRunners = CreateObject("Scripting.Dictionary")
        Set m_seqOverrideRunners = CreateObject("Scripting.Dictionary")
        Set m_currentFrameState = CreateObject("Scripting.Dictionary")
        Set m_seqs = CreateObject("Scripting.Dictionary")
        Set m_pulse = CreateObject("Scripting.Dictionary")
        Set m_on = CreateObject("Scripting.Dictionary")
        m_vpxLightSyncRunning = False
        m_vpxLightSyncCollection = Null
		m_initFrameTime = 0
        m_frameTime = 0
        m_pulseInterval = 26
        m_vpxLightSyncClear = False
        m_tableSeqColor = Null
        m_tableSeqFadeUp = Null
        m_tableSeqFadeDown = Null
        useVpxLights = False
        Set m_lightmaps = CreateObject("Scripting.Dictionary")
    End Sub

    Private Sub AssignStateForFrame(key, state)
        If m_currentFrameState.Exists(key) Then
            m_currentFrameState.Remove key
        End If
        m_currentFrameState.Add key, state
    End Sub

    Public Sub LoadLightShows()
        Dim oFile
        Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
        Dim objFileToWrite : Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(cGameName & "_LightShows/lights-out.txt",2,true)
        For Each oFile In oFSO.GetFolder(cGameName & "_LightShows").Files
            If LCase(oFSO.GetExtensionName(oFile.Name)) = "yaml" And Not Left(oFile.Name,6) = "lights" Then
                Dim textStream : Set textStream = oFSO.OpenTextFile(oFile.Path, 1)
                Dim show : show = textStream.ReadAll
                Dim fileName : fileName = "lSeq" & Replace(oFSO.GetFileName(oFile.Name), "."&oFSO.GetExtensionName(oFile.Name), "")
                Dim lcSeq : lcSeq = "Dim " & fileName & " : Set " & fileName & " = New LCSeq"&vbCrLf
                lcSeq = lcSeq + fileName & ".Name = """&fileName&""""&vbCrLf
                Dim seq : seq = ""
                Dim re : Set re = New RegExp
                With re
                    .Pattern    = "- time:.*?\n"
                    .IgnoreCase = False
                    .Global     = True
                End With
                Dim matches : Set matches = re.execute(show)
                Dim steps : steps = matches.Count
                Dim match, nextMatchIndex, uniqueLights
                Set uniqueLights = CreateObject("Scripting.Dictionary")
                nextMatchIndex = 1
                For Each match in matches
                    Dim lightStep
                    If Not nextMatchIndex < steps Then
                        lightStep = Mid(show, match.FirstIndex, Len(show))
                    Else
                        lightStep = Mid(show, match.FirstIndex, matches(nextMatchIndex).FirstIndex - match.FirstIndex)
                        nextMatchIndex = nextMatchIndex + 1
                    End If

                    Dim re1 : Set re1 = New RegExp
                    With re1
                        .Pattern        = ".*:?: '([A-Fa-f0-9]{6})'"
                        .IgnoreCase     = True
                        .Global         = True
                    End With

                    Dim lightMatches : Set lightMatches = re1.execute(lightStep)
                    If lightMatches.Count > 0 Then
                        Dim lightMatch, lightStr, lightSplit
                        lightStr = "Array("
                        lightSplit = 0
                        For Each lightMatch in lightMatches
                            Dim sParts : sParts = Split(lightMatch.Value, ":")
                            Dim lightName : lightName = Trim(sParts(0))
                            Dim color : color = Trim(Replace(sParts(1),"'", ""))
                            If color = "000000" Then
                                lightStr = lightStr + """"&lightName&"|0|000000"","
                            Else
                                lightStr = lightStr + """"&lightName&"|100|"&color&""","
                            End If

                            If Len(lightStr)+20 > 2000 And lightSplit = 0 Then                           
                                lightSplit = Len(lightStr)
                            End If

                            uniqueLights(lightname) = 0
                        Next
                        lightStr = Left(lightStr, Len(lightStr) - 1)
                        lightStr = lightStr & ")"
                        
                        If lightSplit > 0 Then
                            lightStr = Left(lightStr, lightSplit) & " _ " & vbCrLF & Right(lightStr, Len(lightStr)-lightSplit)
                        End If

                        seq = seq + lightStr & ", _"&vbCrLf
                    Else
                        seq = seq + "Array(), _"&vbCrLf
                    End If

                    
                    Set re1 = Nothing
                Next
                
                lcSeq = lcSeq + filename & ".Sequence = Array( " & Left(seq, Len(seq) - 5) & ")"&vbCrLf
                'lcSeq = lcSeq + seq & vbCrLf
                lcSeq = lcSeq + fileName & ".UpdateInterval = 20"&vbCrLf
                lcSeq = lcSeq + fileName & ".Color = Null"&vbCrLf
                lcSeq = lcSeq + fileName & ".Repeat = False"&vbCrLf

                'MsgBox(lcSeq)
                objFileToWrite.WriteLine(lcSeq)
                ExecuteGlobal lcSeq
                Set re = Nothing

                textStream.Close
            End if
        Next
        'Clean up
        objFileToWrite.Close
        Set objFileToWrite = Nothing
        Set oFile = Nothing
        Set oFSO = Nothing
    End Sub

    Public Sub CompileLights(collection, name)
        Dim light
        Dim lights : lights = "light:" & vbCrLf
        For Each light in collection
            lights = lights + light.name & ":"&vbCrLf
            lights = lights + "   x: "& light.x/tablewidth & vbCrLf
            lights = lights + "   y: "& light.y/tableheight & vbCrLf
        Next
        Dim objFileToWrite : Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(cGameName & "_LightShows/lights-"&name&".yaml",2,true)
	    objFileToWrite.WriteLine(lights)
	    objFileToWrite.Close
	    Set objFileToWrite = Nothing
        Debug.print("Lights YAML File saved to: " & cGameName & "_LightShows/lights-"&name&".yaml")
    End Sub

    Public Sub RegisterLights(mode)

        Dim idx,tmp,vpxLight,lcItem
        If mode = "Lampz" Then
            
            For idx = 0 to UBound(Lampz.obj)
                If Lampz.IsLight(idx) Then
                    Set lcItem = new LCItem
                    If IsArray(Lampz.obj(idx)) Then
                        tmp = Lampz.obj(idx)
                        Set vpxLight = tmp(0)
                    Else
                        Set vpxLight = Lampz.obj(idx)
                        
                    End If
                    Lampz.Modulate(idx) = 1/100
                    Lampz.FadeSpeedUp(idx) = 100/30 : Lampz.FadeSpeedDown(idx) = 100/120
                    lcItem.Init idx, vpxLight.BlinkInterval, Array(vpxLight.color, vpxLight.colorFull), vpxLight.name, vpxLight.x, vpxLight.y
                    
                    m_lights.Add vpxLight.Name, lcItem
                    m_seqRunners.Add "lSeqRunner" & CStr(vpxLight.name), new LCSeqRunner
                End If
            Next        
        ElseIf mode = "VPX" Then
            useVpxLights = True


            For idx = 0 to UBound(Lights)
                vpxLight = Null
                Set lcItem = new LCItem
                If IsArray(Lights(idx)) Then
                    tmp = Lights(idx)
                    Set vpxLight = tmp(0)
                ElseIf IsObject(Lights(idx)) Then
                    Set vpxLight = Lights(idx)
                End If
                If Not IsNull(vpxLight) Then
                    Dim e, lmStr: lmStr = "lmArr = Array("    
                    For Each e in GetElements()
                        If Right(e.Name, Len(vpxLight.Name)+1) = "_" & vpxLight.Name Then
                            Debug.Print(e.Name)
                            lmStr = lmStr & e.Name & ","
                        End If
                    Next
                    lmStr = lmStr & "Null)"
                    lmStr = Replace(lmStr, ",Null)", ")")
			        ExecuteGlobal "Dim lmArr : "&lmStr
                    m_lightmaps.Add vpxLight.Name, lmArr
                    Debug.print("Registering Light: "& vpxLight.name) 
                    lcItem.Init idx, vpxLight.BlinkInterval, Array(vpxLight.color, vpxLight.colorFull), vpxLight.name, vpxLight.x, vpxLight.y
                    m_lights.Add vpxLight.Name, lcItem
                    m_seqRunners.Add "lSeqRunner" & CStr(vpxLight.name), new LCSeqRunner
                End If
            Next  
        End If
    End Sub

    Private Function ColtoArray(aDict)	'converts a collection to an indexed array. Indexes will come out random probably.
        redim a(999)
        dim count : count = 0
        dim x  : for each x in aDict : set a(Count) = x : count = count + 1 : Next
        redim preserve a(count-1) : ColtoArray = a
    End Function

	Public Sub AddLight(light, idx)
        If m_lights.Exists(light.name) Then
            Exit Sub
        End If
        Dim lcItem : Set lcItem = new LCItem
        lcItem.Init idx, light.BlinkInterval, Array(light.color, light.colorFull), light.name, light.x, light.y
        m_lights.Add light.Name, lcItem
        m_seqRunners.Add "lSeqRunner" & CStr(light.name), new LCSeqRunner
    End Sub

    Public Sub LightState(light, state)
        m_lightOff(light.name)
        If state = 1 Then
            m_lightOn(light.name)
        ElseIF state = 2 Then
            Blink(light)
        End If
    End Sub

    Public Sub LightOn(light)
        m_LightOn(light.name)
    End Sub

    Public Sub LightOnWithColor(light, color)
        m_LightOnWithColor light.name, color
    End Sub

    Public Sub FlickerOn(light)
        Dim name : name = light.name
        If m_lights.Exists(name) Then
            m_lightOn(name)

            If m_pulse.Exists(name) Then 
                Exit Sub
            End If
            m_pulse.Add name, (new PulseState)(m_lights(name), Array(37,100,24,0,70), 0, m_pulseInterval, 1)
        End If
    End Sub  
    
    Public Sub LightColor(light, color)
        If m_lights.Exists(light.name) Then
            m_lights(light.name).Color = color
            'Update internal blink seq for light
            If m_seqs.Exists(light.name & "Blink") Then
                m_seqs(light.name & "Blink").Color = color
            End If

        End If
    End Sub

    Private Sub m_LightOn(name)
        If m_lights.Exists(name) Then
            If m_off.Exists(name) Then 
                m_off.Remove(name)
            End If
            If m_seqs.Exists(name & "Blink") Then
                m_seqRunners("lSeqRunner"&CStr(name)).RemoveItem m_seqs(name & "Blink")
            End If
            If m_on.Exists(name) Then 
                Exit Sub
            End If
            m_on.Add name, m_lights(name)
        End If
    End Sub

    Private Sub m_LightOnWithColor(name, color)
        If m_lights.Exists(name) Then
            m_lights(name).Color = color
            If m_off.Exists(name) Then 
                m_off.Remove(name)
            End If

            If m_seqs.Exists(name & "Blink") Then
                m_seqRunners("lSeqRunner"&CStr(name)).RemoveItem m_seqs(name & "Blink")
            End If

            If m_on.Exists(name) Then 
                Exit Sub
            End If
            m_on.Add name, m_lights(name)
        End If
    End Sub

    Public Sub LightOff(light)
        m_lightOff(light.name)
    End Sub

    Private Sub m_lightOff(name)
        If m_lights.Exists(name) Then
            If m_on.Exists(name) Then 
                m_on.Remove(name)
            End If

            If m_seqs.Exists(name & "Blink") Then
                m_seqRunners("lSeqRunner"&CStr(name)).RemoveItem m_seqs(name & "Blink")
            End If

            If m_off.Exists(name) Then 
                Exit Sub
            End If
            m_off.Add name, m_lights(name)
        End If
    End Sub

    Public Sub UpdateBlinkInterval(light, interval)
        If m_lights.Exists(light.name) Then
            light.BlinkInterval = interval
            If m_seqs.Exists(light.name & "Blink") Then
                m_seqs.Item(light.name & "Blink").UpdateInterval = interval
            End If
        End If
    End Sub


    Public Sub Pulse(light, repeatCount)
        Dim name : name = light.name
        If m_lights.Exists(name) Then
            If m_off.Exists(name) Then 
                m_off.Remove(name)
            End If
            If m_pulse.Exists(name) Then 
                Exit Sub
            End If
            'Array(100,94,32,13,6,3,0)
            m_pulse.Add name, (new PulseState)(m_lights(name), Array(37,100,24,0,70,100,12,0), 0, m_pulseInterval, repeatCount)
        End If
    End Sub

    Public Sub PulseWithProfile(light, profile, repeatCount)
        Dim name : name = light.name
        If m_lights.Exists(name) Then
            If m_off.Exists(name) Then 
                m_off.Remove(name)
            End If
            If m_pulse.Exists(name) Then 
                Exit Sub
            End If
            m_pulse.Add name, (new PulseState)(m_lights(name), profile, 0, m_pulseInterval, repeatCount)
        End If
    End Sub       

    Public Sub PulseWithState(pulse)
        
        If m_lights.Exists(pulse.Light) Then
            If m_off.Exists(pulse.Light) Then 
                m_off.Remove(pulse.Light)
            End If
            If m_pulse.Exists(pulse.Light) Then 
                Exit Sub
            End If
            m_pulse.Add name, pulse
        End If
    End Sub

    Public Sub LightLevel(light, lvl)
        If m_lights.Exists(light.name) Then
            m_lights(light.name).Level = lvl

            If m_seqs.Exists(light.name & "Blink") Then
                m_seqs(light.name & "Blink").Sequence = m_buildBlinkSeq(light)
            End If
        End If
    End Sub


    Public Sub AddShot(name, light, color)
        If m_lights.Exists(light.name) Then
            If m_seqs.Exists(name & light.name) Then
                m_seqs(name & light.name).Color = color
                m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem m_seqs(name & light.name)
            Else
                Dim stateOn : stateOn = light.name&"|100"
                Dim stateOff : stateOff = light.name&"|0"
                Dim seq : Set seq = new LCSeq
                seq.Name = name
                seq.Sequence = Array(stateOn, stateOff,stateOn, stateOff)
                seq.Color = color
                seq.UpdateInterval = light.BlinkInterval
                seq.Repeat = True

                m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem seq
                m_seqs.Add name & light.name, seq
            End If
            If m_on.Exists(light.name) Then
                m_on.Remove light.name
            End If
        End If
    End Sub

    Public Sub RemoveShot(name, light)
        If m_lights.Exists(light.name) And m_seqs.Exists(name & light.name) Then
            m_seqRunners("lSeqRunner"&CStr(light.name)).RemoveItem m_seqs(name & light.name)
            If IsNUll(m_seqRunners("lSeqRunner"&CStr(light.name)).CurrentItem) Then
               LightOff(light)
            End If
        End If
    End Sub

    Public Sub RemoveAllShots()
        Dim light
        For Each light in m_lights.Keys()
            m_seqRunners("lSeqRunner"&CStr(light)).RemoveAll
            AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
        Next
    End Sub

    Public Sub RemoveShotsFromLight(light)
        If m_lights.Exists(light.name) Then
            m_seqRunners("lSeqRunner"&CStr(light.name)).RemoveAll   
            m_lightOff(light.name)  
        End If
    End Sub

    Public Sub Blink(light)
        If m_lights.Exists(light.name) Then

            If m_seqs.Exists(light.name & "Blink") Then
                m_seqs(light.name & "Blink").ResetInterval
                m_seqs(light.name & "Blink").CurrentIdx = 0 
                m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem m_seqs(light.name & "Blink")
            Else
                Dim seq : Set seq = new LCSeq
                seq.Name = light.name & "Blink"
                seq.Sequence = m_buildBlinkSeq(light)
                seq.Color = Null
                seq.UpdateInterval = light.BlinkInterval
                seq.Repeat = True

                m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem seq
                m_seqs.Add light.name & "Blink", seq
            End If
            If m_on.Exists(light.name) Then
                m_on.Remove light.name
            End If
        End If
    End Sub

    Public Function GetLightState(light)
        GetLightState = 0
        If(m_lights.Exists(light.name)) Then
            If m_on.Exists(light.name) Then
                GetLightState = 1
            Else
                If m_seqs.Exists(light.name & "Blink") Then
                    GetLightState = 2
                End If
            End If
        End If
    End Function

    Public Function IsShotLit(name, light)
        IsShotLit = False
        If(m_lights.Exists(light.name)) Then
            If m_seqRunners("lSeqRunner"&CStr(light.name)).HasSeq(name) Then
                IsShotLit = True
            End If
        End If
    End Function

    Public Sub CreateSeqRunner(name)
        If m_seqRunners.Exists(name) Then
            Exit Sub
        End If
        Dim seqRunner : Set seqRunner = new LCSeqRunner
        seqRunner.Name = name
        m_seqRunners.Add name, seqRunner
    End Sub

    Private Sub CreateOverrideSeqRunner(name)
        If m_seqOverrideRunners.Exists(name) Then
            Exit Sub
        End If
        Dim seqRunner : Set seqRunner = new LCSeqRunner
        seqRunner.Name = name
        m_seqOverrideRunners.Add name, seqRunner
    End Sub

    Public Sub AddLightSeq(lcSeqRunner, lcSeq)
        If Not m_seqRunners.Exists(lcSeqRunner) Then
            Exit Sub
        End If

        m_seqRunners(lcSeqRunner).AddItem lcSeq
    End Sub

    Public Sub RemoveLightSeq(lcSeqRunner, lcSeq)
        If Not m_seqRunners.Exists(lcSeqRunner) Then
            Exit Sub
        End If

        Dim light
        For Each light in lcSeq.LightsInSeq
            If(m_lights.Exists(light)) Then
                    AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
            End If
        Next

        m_seqRunners(lcSeqRunner).RemoveItem lcSeq
    End Sub

    Public Sub RemoveAllLightSeq(lcSeqRunner)
        If Not m_seqRunners.Exists(lcSeqRunner) Then
            Exit Sub
        End If
        Dim lcSeqKey, light, seqs, lcSeq
        Set seqs = m_seqRunners(lcSeqRunner).Items()
        For Each lcSeqKey in seqs.Keys()
			Set lcSeq = seqs(lcSeqKey)
            For Each light in lcSeq.LightsInSeq
                If(m_lights.Exists(light)) Then
                    AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
                End If
            Next
        Next

        m_seqRunners(lcSeqRunner).RemoveAll
    End Sub

    Public Sub AddTableLightSeq(name, lcSeq)
        CreateOverrideSeqRunner(name)

        Dim seqOverride, hasOverride
        hasOverride = False
        For Each seqOverride In m_seqOverrideRunners.Keys()
            If Not IsNull(m_seqOverrideRunners(seqOverride).CurrentItem) Then
                hasOverride = True
            End If
        Next
        If hasOverride = False Then
            Dim light
            For Each light in m_lights.Keys()
                AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
            Next
        End If
        m_seqOverrideRunners(name).AddItem lcSeq
    End Sub

    Public Sub RemoveTableLightSeq(name, lcSeq)
        If Not m_seqOverrideRunners.Exists(name) Then
            Exit Sub
        End If
        m_seqOverrideRunners(name).RemoveItem lcSeq
        Dim seqOverride, hasOverride
        hasOverride = False
        For Each seqOverride In m_seqOverrideRunners.Keys()
            If Not IsNull(m_seqOverrideRunners(seqOverride).CurrentItem) Then
                hasOverride = True
            End If
        Next
        If hasOverride = False Then
            Dim light
            For Each light in m_lights.Keys()
                AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
            Next
        End If
    End Sub

    Public Sub RemoveAllTableLightSeqs()
        Dim light, runner
        For Each runner in m_seqOverrideRunners.Keys()
            m_seqOverrideRunners(runner).RemoveAll()
        Next
		For Each light in m_lights.Keys()
            AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
        Next
    End Sub

   Public Sub SyncLightMapColors()
        dim light,lm
        For Each light in m_lights.Keys()
            If m_lightmaps.Exists(light) Then
                For Each lm in m_lightmaps(light)
                    dim color : color = m_lights(light).Color
                    If not IsNull(lm) Then
						lm.Color = color(0)
					End If
                Next
            End If
        Next
    End Sub

    Public Sub SyncWithVpxLights(lightSeq)
        m_vpxLightSyncCollection = ColToArray(eval(lightSeq.collection))
        m_vpxLightSyncRunning = True
    End Sub

    Public Sub StopSyncWithVpxLights()
        m_vpxLightSyncRunning = False
        m_vpxLightSyncClear = True
		m_tableSeqColor = Null
        m_tableSeqFadeUp = Null
        m_tableSeqFadeDown = Null
    End Sub

	Public Sub SetVpxSyncLightColor(color)
		m_tableSeqColor = color
	End Sub

    Public Sub SetTableSequenceFade(fadeUp, fadeDown)
		m_tableSeqFadeUp = fadeUp
        m_tableSeqFadeDown = fadeDown
	End Sub

    Public Sub UseToolkitColoredLightMaps()
        If useVpxLights = True Then
            Exit Sub
        End If

        Dim sUpdateLightMap
        sUpdateLightMap = "Sub UpdateLightMap(idx, lightmap, intensity, ByVal aLvl)" + vbCrLf    
        sUpdateLightMap = sUpdateLightMap + "   if Lampz.UseFunc then aLvl = Lampz.FilterOut(aLvl)	'Callbacks don't get this filter automatically" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "   lightmap.Opacity = aLvl * intensity" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "   If IsArray(Lampz.obj(idx) ) Then" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "       lightmap.Color = Lampz.obj(idx)(0).color" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "   Else" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "       lightmap.color = Lampz.obj(idx).color" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "   End If" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "End Sub" + vbCrLf

        ExecuteGlobal sUpdateLightMap

        Dim x
        For x=0 to Ubound(Lampz.cCallback)
            Lampz.cCallback(x) = Replace(Lampz.cCallback(x), "UpdateLightMap ", "UpdateLightMap " & x & ",")
            Lampz.Callback(x) = "" 'Force Callback Sub to be build
        Next
    End Sub

    Private Function m_buildBlinkSeq(light)
        Dim i, buff : buff = Array()
        ReDim buff(Len(light.BlinkPattern)-1)
        For i = 0 To Len(light.BlinkPattern)-1
            
            If Mid(light.BlinkPattern, i+1, 1) = 1 Then
                buff(i) = light.name & "|100"
            Else
                buff(i) = light.name & "|0"
            End If
        Next
        m_buildBlinkSeq=buff
    End Function

    Private Function GetTmpLight(idx)
        If useVpxLights = True Then
          If IsArray(Lights(idx) ) Then	'if array
                Set GetTmpLight = Lights(idx)(0)
            Else
                Set GetTmpLight = Lights(idx)
            End If
        Else
            If IsArray(Lampz.obj(idx) ) Then	'if array
                Set GetTmpLight = Lampz.obj(idx)(0)
            Else
                Set GetTmpLight = Lampz.obj(idx)
            End If
        End If
        
    End Function

    Public Sub ResetLights()
        Dim light
        For Each light in m_lights.Keys()
            m_seqRunners("lSeqRunner"&CStr(light)).RemoveAll
            m_lightOff(light) 
            AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
        Next
        RemoveAllTableLightSeqs()
        Dim k
        For Each k in m_seqRunners.Keys()
            Dim lsRunner: Set lsRunner = m_seqRunners(k)
            lsRunner.RemoveAll
        Next

    End Sub

    Public Sub Update()

		m_frameTime = gametime - m_initFrameTime : m_initFrameTime = gametime
		Dim x
        Dim lk
        dim color
        Dim lightKey
        Dim lcItem
        Dim tmpLight
        Dim seqOverride, hasOverride
        hasOverride = False
        For Each seqOverride In m_seqOverrideRunners.Keys()
            If Not IsNull(m_seqOverrideRunners(seqOverride).CurrentItem) Then
                RunLightSeq m_seqOverrideRunners(seqOverride)
                hasOverride = True
            End If
        Next
        If hasOverride = False Then
        
            If HasKeys(m_on) Then   
                For Each lightKey in m_on.Keys()
                    Set lcItem = m_on(lightKey)
                    AssignStateForFrame lightKey, (new FrameState)(lcItem.level, m_on(lightKey).Color, m_on(lightKey).Idx)
                Next
            End If

            If HasKeys(m_pulse) Then   
                For Each lightKey in m_pulse.Keys()
                    AssignStateForFrame lightKey, (new FrameState)(m_pulse(lightKey).PulseAt(m_pulse(lightKey).idx), m_pulse(lightKey).light.Color, m_pulse(lightKey).light.Idx)
                    Dim pulseUpdateInt : pulseUpdateInt = m_pulse(lightKey).interval - m_frameTime
                    Dim pulseIdx : pulseIdx = m_pulse(lightKey).idx
                    If pulseUpdateInt <= 0 Then
                        pulseUpdateInt = m_pulseInterval
                        pulseIdx = pulseIdx + 1
                    End If
                    
                    Dim pulses : pulses = m_pulse(lightKey).pulses
					Dim pulseCount : pulseCount = m_pulse(lightKey).Cnt
                    If pulseIdx > UBound(m_pulse(lightKey).pulses) Then
						m_pulse.Remove lightKey    
						If pulseCount > 0 Then
                            pulseCount = pulseCount - 1
                            pulseIdx = 0
                            m_pulse.Add lightKey, (new PulseState)(m_lights(lightKey),pulses, pulseIdx, pulseUpdateInt, pulseCount)
                        End If
                    Else
						m_pulse.Remove lightKey
                        m_pulse.Add lightKey, (new PulseState)(m_lights(lightKey),pulses, pulseIdx, pulseUpdateInt, pulseCount)
                    End If
                Next
            End If

            If HasKeys(m_off) Then
                For Each lightKey in m_off.Keys()
                    Set lcItem = m_off(lightKey)
                    AssignStateForFrame lightKey, (new FrameState)(0, Null, lcItem.Idx)
                Next
            End If

            If HasKeys(m_seqRunners) Then
                Dim k
                For Each k in m_seqRunners.Keys()
                    Dim lsRunner: Set lsRunner = m_seqRunners(k)
                    If Not IsNull(lsRunner.CurrentItem) Then
                            RunLightSeq lsRunner
                    End If
                Next
            End If

            If m_vpxLightSyncRunning = True Then
                Dim lx
                If Not IsNull(m_vpxLightSyncCollection) Then
                    For Each lx in m_vpxLightSyncCollection
                        'sync each light being ran by the vpx LS
                        dim syncLight : syncLight = Null
                        If m_lights.Exists(lx.name) Then
                            'found a light
                            Set syncLight = m_lights(lx.name)
                        End If
                        If Not IsNull(syncLight) Then
                            'Found a light to sync.
                            Dim lightState

                            If IsNull(m_tableSeqColor) Then
                                color = syncLight.Color
                            Else
                                If Not IsArray(m_tableSeqColor) Then
                                    color = Array(m_TableSeqColor, Null)
                                Else
                                    color = m_tableSeqColor
                                End If
                            End If

                            'TODO - Fix VPX Fade
                            If Not useVpxLights = True Then
                                If Not IsNull(m_tableSeqFadeUp) Then
                                    Lampz.FadeSpeedUp(syncLight.Idx) = m_tableSeqFadeUp
                                End If
                                If Not IsNull(m_tableSeqFadeDown) Then
                                    Lampz.FadeSpeedDown(syncLight.Idx) = m_tableSeqFadeDown
                                End If
                            End If
                    
                            AssignStateForFrame syncLight.name, (new FrameState)(lx.GetInPlayState*100,color, syncLight.Idx)                     
                        End If
                    Next
		        End If
            End If

            If m_vpxLightSyncClear = True Then  
                If Not IsNull(m_vpxLightSyncCollection) Then
                    For Each lk in m_vpxLightSyncCollection
                        'sync each light being ran by the vpx LS
                        dim syncClearLight : syncClearLight = Null
                        If m_lights.Exists(lk.name) Then
                            'found a light
                            Set syncClearLight = m_lights(lk.name)
                        End If
                        If Not IsNull(syncClearLight) Then
                            AssignStateForFrame syncClearLight.name, (new FrameState)(0, Null, syncClearLight.idx) 
                            'TODO - Only do fade speed for lampz
                            If Not useVpxLights = True Then
                                Lampz.FadeSpeedUp(syncClearLight.Idx) = 100/30
                                Lampz.FadeSpeedDown(syncClearLight.Idx) = 100/120
                            End If
                        End If
                    Next
                End If
               
                m_vpxLightSyncClear = False
            End If
        End If
        

        If HasKeys(m_currentFrameState) Then
			
            Dim frameStateKey
            For Each frameStateKey in m_currentFrameState.Keys()
                Dim idx : idx = m_currentFrameState(frameStateKey).idx
                
                Dim newColor : newColor = m_currentFrameState(frameStateKey).colors
                Dim bUpdate

                If Not IsNull(newColor) Then
                    'Check current color is the new color coming in, if not, set the new color.
                    
                    Set tmpLight = GetTmpLight(idx)

					Dim c, cf
					c = newColor(0)
					cf= newColor(1)

					If Not IsNull(c) Then
						If Not CStr(tmpLight.Color) = CStr(c) Then
							bUpdate = True
						End If
					End If

					If Not IsNull(cf) Then
						If Not CStr(tmpLight.ColorFull) = CStr(cf) Then
							bUpdate = True
						End If
					End If
            	End If

                If useVpxLights = False Then
                    If bUpdate Then
                        'Update lamp color
                        If IsArray(Lampz.obj(idx)) Then
                            for each x in Lampz.obj(idx)
                                If Not IsNull(c) Then
                                    x.color = c
                                End If
                                If Not IsNull(cf) Then
                                    x.colorFull = cf
                                End If
                            Next
                        Else
                            If Not IsNull(c) Then
                                Lampz.obj(idx).color = c
                            End If
                            If Not IsNull(cf) Then
                                Lampz.obj(idx).colorFull = cf
                            End If
                        End If
                        If Lampz.UseCallBack(idx) then Proc Lampz.name & idx,Lampz.Lvl(idx)*Lampz.Modulate(idx)	'Force Callbacks Proc
                    End If
                    Lampz.state(idx) = CInt(m_currentFrameState(frameStateKey).level) 'Lampz will handle redundant updates
                Else
                    Dim lm
                    If IsArray(Lights(idx)) Then
                        For Each x in Lights(idx)
                            If bUpdate Then 
                                If Not IsNull(c) Then
                                    x.color = c
                                End If
                                If Not IsNull(cf) Then
                                    x.colorFull = cf
                                End If
                                If m_lightmaps.Exists(x.Name) Then
                                    For Each lm in m_lightmaps(x.Name)
                                        lm.Color = c
                                    Next
                                End If
                            End If
                            x.State = m_currentFrameState(frameStateKey).level/100
                        Next
                    Else
                        If bUpdate Then    
                            If Not IsNull(c) Then
                                Lights(idx).color = c
                            End If
                            If Not IsNull(cf) Then
                                Lights(idx).colorFull = cf
                            End If
                            If m_lightmaps.Exists(Lights(idx).Name) Then
                                For Each lm in m_lightmaps(Lights(idx).Name)
                                    If Not IsNull(lm) Then
                                        lm.Color = c
                                    End If
                                Next
                            End If
                        End If
                        Lights(idx).State = m_currentFrameState(frameStateKey).level/100
                    End If
                End If


           
                
				 
            Next
        End If
        m_currentFrameState.RemoveAll
        m_off.RemoveAll

    End Sub

    Private Function HexToInt(hex)
        HexToInt = CInt("&H" & hex)
    End Function

    Private Function HasKeys(o)
        Dim Success
        Success = False

        On Error Resume Next
            o.Keys()
            Success = (Err.Number = 0)
        On Error Goto 0
        HasKeys = Success
    End Function

    Private Sub RunLightSeq(seqRunner)

        Dim lcSeq: Set lcSeq = seqRunner.CurrentItem
        dim lsName, isSeqEnd
        If UBound(lcSeq.Sequence)<lcSeq.CurrentIdx Then
            isSeqEnd = True
        Else
            isSeqEnd = False
        End If

        dim lightInSeq
        For each lightInSeq in lcSeq.LightsInSeq
        
            If isSeqEnd Then

                

            'Needs a guard here for something, but i've forgotten. 
            'I remember: Only reset the light if there isn't frame data for the light. 
            'e.g. a previous seq has affected the light, we don't want to clear that here on this frame
                If m_lights.Exists(lightInSeq) = True AND NOT m_currentFrameState.Exists(lightInSeq) Then


                    If lcSeq.Name = "lSeqRgbRandomRed" AND lightInSeq = "l46" Then
                     Debug.print("Reseting l46")
                    End If
                   AssignStateForFrame lightInSeq, (new FrameState)(0, Null, m_lights(lightInSeq).Idx)
                End If
            Else
                


                If m_currentFrameState.Exists(lightInSeq) Then

                    
                    'already frame data for this light.
                    'replace with the last known state from this seq
                    If Not IsNull(lcSeq.LastLightState(lightInSeq)) Then
                        If lcSeq.Name = "lSeqRgbRandomRed" AND lightInSeq = "l46" Then
                            Debug.print("Assigning Previous State for l46")
                        End If
						AssignStateForFrame lightInSeq, lcSeq.LastLightState(lightInSeq)
                    End If
                End If

            End If
        Next

        If isSeqEnd Then
            lcSeq.CurrentIdx = 0
            seqRunner.NextItem()
        End If

        If Not IsNull(seqRunner.CurrentItem) Then
            Dim framesRemaining, seq, color
            Set lcSeq = seqRunner.CurrentItem
            seq = lcSeq.Sequence
            

            Dim name
            Dim ls, x
            If IsArray(seq(lcSeq.CurrentIdx)) Then
                For x = 0 To UBound(seq(lcSeq.CurrentIdx))
                    lsName = Split(seq(lcSeq.CurrentIdx)(x),"|")
                    name = lsName(0)
                    If m_lights.Exists(name) Then
                        Set ls = m_lights(name)
                        
						color = lcSeq.Color

                        If IsNull(color) Then
							color = ls.Color
                        End If
						
                        If Ubound(lsName) = 2 Then
							If lsName(2) = "FFFFFF" Then
                                AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
                            Else
                                AssignStateForFrame name, (new FrameState)(lsName(1), Array( RGB( HexToInt(Left(lsName(2), 2)), HexToInt(Mid(lsName(2), 3, 2)), HexToInt(Right(lsName(2), 2)) ), RGB(0,0,0)), ls.Idx)
                            End If
                        Else
                            AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
                        End If
                        lcSeq.SetLastLightState name, m_currentFrameState(name) 
                    End If
                Next       
            Else
                lsName = Split(seq(lcSeq.CurrentIdx),"|")
                name = lsName(0)
                If m_lights.Exists(name) Then
                    Set ls = m_lights(name)
                    
					color = lcSeq.Color
                    If IsNull(color) Then
                        color = ls.Color
                    End If
                    If Ubound(lsName) = 2 Then
                        If lsName(2) = "FFFFFF" Then
                            AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
                        Else
                            AssignStateForFrame name, (new FrameState)(lsName(1), Array( RGB( HexToInt(Left(lsName(2), 2)), HexToInt(Mid(lsName(2), 3, 2)), HexToInt(Right(lsName(2), 2)) ), RGB(0,0,0)), ls.Idx)
                        End If
                    Else
                        AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
                    End If
                    lcSeq.SetLastLightState name, m_currentFrameState(name) 
                End If
            End If

            framesRemaining = lcSeq.Update(m_frameTime)
            If framesRemaining < 0 Then
                lcSeq.ResetInterval()
                lcSeq.NextFrame()
            End If
            
        End If
    End Sub

End Class

Class FrameState
    Private m_level, m_colors, m_idx

    Public Property Get Level(): Level = m_level: End Property
    Public Property Let Level(input): m_level = input: End Property

    Public Property Get Colors(): Colors = m_colors: End Property
    Public Property Let Colors(input): m_colors = input: End Property

    Public Property Get Idx(): Idx = m_idx: End Property
    Public Property Let Idx(input): m_idx = input: End Property

    Public default function init(level, colors, idx)
		m_level = level
		m_colors = colors
		m_idx = idx 

		Set Init = Me
    End Function

    Public Function ColorAt(idx)
        ColorAt = m_colors(idx) 
    End Function
End Class
 
Class PulseState
    Private m_light, m_pulses, m_idx, m_interval, m_cnt

    Public Property Get Light(): Set Light = m_light: End Property
    Public Property Let Light(input): Set m_light = input: End Property

    Public Property Get Pulses(): Pulses = m_pulses: End Property
    Public Property Let Pulses(input): m_pulses = input: End Property

    Public Property Get Idx(): Idx = m_idx: End Property
    Public Property Let Idx(input): m_idx = input: End Property

    Public Property Get Interval(): Interval = m_interval: End Property
    Public Property Let Interval(input): m_interval = input: End Property

    Public Property Get Cnt(): Cnt = m_cnt: End Property
    Public Property Let Cnt(input): m_cnt = input: End Property

    Public default function init(light, pulses, idx, interval, cnt)
		Set m_light = light
		m_pulses = pulses
		'debug.Print(Join(Pulses))
		m_idx = idx 
		m_interval = interval
		m_cnt = cnt

		Set Init = Me
    End Function

    Public Function PulseAt(idx)
        PulseAt = m_pulses(idx) 
    End Function
End Class

Class LCItem
	
	Private m_Idx, m_State, m_blinkSeq, m_color, m_name, m_level, m_x, m_y

        Public Property Get Idx()
            Idx=m_Idx
        End Property

        Public Property Get Color()
            Color=m_color
        End Property

        Public Property Let Color(input)
            If IsNull(input) Then
				m_Color = Null
			Else
				If Not IsArray(input) Then
					input = Array(input, null)
				End If
				m_Color = input
			End If
	    End Property

        Public Property Let Level(input)
            m_level = input
	    End Property

        Public Property Get Level()
            Level=m_level
        End Property

        Public Property Get Name()
            Name=m_name
        End Property

        Public Property Get X()
            X=m_x
        End Property

        Public Property Get Y()
            Y=m_y
        End Property

        Public Sub Init(idx, intervalMs, color, name, x, y)
            m_Idx = idx
            If Not IsArray(color) Then
                m_color = Array(color, null)
            Else
                m_color = color
            End If
            m_name = name
            m_level = 100
            m_x = x
            m_y = y
	    End Sub

End Class

Class LCSeq
	
	Private m_currentIdx, m_sequence, m_name, m_image, m_color, m_updateInterval, m_Frames, m_repeat, m_lightsInSeq, m_lastLightStates

    Public Property Get CurrentIdx()
        CurrentIdx=m_currentIdx
    End Property

    Public Property Let CurrentIdx(input)
		m_lastLightStates.RemoveAll()
        m_currentIdx = input
    End Property

    Public Property Get LightsInSeq()
        LightsInSeq=m_lightsInSeq.Keys()
    End Property

    Public Property Get Sequence()
        Sequence=m_sequence
    End Property
    
	Public Property Let Sequence(input)
		m_sequence = input
        dim item, light, lightItem
        for each item in input
            If IsArray(item) Then
                for each light in item
                    lightItem = Split(light,"|")
                    If Not m_lightsInSeq.Exists(lightItem(0)) Then
                        m_lightsInSeq.Add lightItem(0), True
                    End If    
                next
            Else
                lightItem = Split(item,"|")
                If Not m_lightsInSeq.Exists(lightItem(0)) Then
                    m_lightsInSeq.Add lightItem(0), True
                End If
            End If
        next
	End Property

    Public Property Get LastLightState(light)
		If m_lastLightStates.Exists(light) Then
			dim c : Set c = m_lastLightStates(light)
			Set LastLightState = c
		Else
			LastLightState = Null
		End If
    End Property

    Public Property Let LastLightState(light, input)
        If m_lastLightStates.Exists(light) Then
            m_lastLightStates.Remove light
        End If
		If input.level > 0 Then
			m_lastLightStates.Add light, input
		End If
    End Property

    Public Sub SetLastLightState(light, input)	
        If m_lastLightStates.Exists(light) Then	
            m_lastLightStates.Remove light	
        End If	
        If input.level > 0 Then	
                m_lastLightStates.Add light, input	
        End If	
    End Sub

    Public Property Get Color()
        Color=m_color
    End Property
    
	Public Property Let Color(input)
		If IsNull(input) Then
			m_Color = Null
		Else
			If Not IsArray(input) Then
				input = Array(input, null)
			End If
			m_Color = input
		End If
	End Property

    Public Property Get Name()
        Name=m_name
    End Property
    
	Public Property Let Name(input)
		m_name = input
	End Property        

    Public Property Get UpdateInterval()
        UpdateInterval=m_updateInterval
    End Property

    Public Property Let UpdateInterval(input)
        m_updateInterval = input
        'm_Frames = input
    End Property

    Public Property Get Repeat()
        Repeat=m_repeat
    End Property

    Public Property Let Repeat(input)
        m_repeat = input
    End Property

    Private Sub Class_Initialize()
        m_currentIdx = 0
        m_color = Array(Null, Null)
        m_updateInterval = 180
        m_repeat = False
        m_Frames = 180
        Set m_lightsInSeq = CreateObject("Scripting.Dictionary")
        Set m_lastLightStates = CreateObject("Scripting.Dictionary")
    End Sub

    Public Property Get Update(framesPassed)
        m_Frames = m_Frames - framesPassed
        Update = m_Frames
    End Property

    Public Sub NextFrame()
        m_currentIdx = m_currentIdx + 1
    End Sub

    Public Sub ResetInterval()

        m_Frames = m_updateInterval
        Exit Sub

        If Not IsNull(m_sequence) And UBound(m_sequence) > 1 Then

        'For i = 0 To totalSteps - 1
        '    currentStep = i
        '    duration = 20 ' Base duration of 20ms
            'Debug.print("TotalSteps: " & UBound(m_sequence)-1)
            Dim easeAmount : easeAmount = Round(m_currentIdx / UBound(m_sequence), 2) ' Normalize current step
            if easeAmount < 0 then
                easeAmount = 0
            elseif easeAmount > 1 then
                easeAmount = 1
            end if
            'Debug.print("Step: " & m_currentIdx)
            'Debug.print("Ease Amount: "& easeAmount)
            Dim newDuration : newDuration = 100 - Lerp(20, 80, EaseIn(easeAmount) )' Apply EaseInOut to duration
            'Debug.print("Duration: "& Round(newDuration))
            'Dim newDuration : newDuration = 100- Lerp(20, 80, Spike(easeAmount) )' Apply EaseInOut to duration
            
            m_frames = newDuration
        Else
            m_Frames = m_updateInterval
        End If
    End Sub

End Class

Class LCSeqRunner
	
	Private m_name, m_items,m_currentItemIdx

    Public Property Get Name()
        Name=m_name
    End Property
    
	Public Property Let Name(input)
		m_name = input
	End Property

    Public Property Get Items()
		Set Items = m_items
	End Property

    Public Property Get CurrentItem()
        Dim items: items = m_items.Items()
        If m_currentItemIdx > UBound(items) Then
            m_currentItemIdx = 0
        End If
        If UBound(items) = -1 Then       
            CurrentItem  = Null
        Else
            Set CurrentItem = items(m_currentItemIdx)                
        End If
    End Property

    Private Sub Class_Initialize()    
        Set m_items = CreateObject("Scripting.Dictionary")
        m_currentItemIdx = 0
    End Sub

    Public Sub AddItem(item)
        If Not IsNull(item) Then
            If Not m_items.Exists(item.Name) Then
                m_items.Add item.Name, item
            End If
        End If
    End Sub

    Public Sub RemoveAll()
        Dim item
        For Each item in m_items.Keys()
            m_items(item).ResetInterval
            m_items(item).CurrentIdx = 0
            m_items.Remove item
        Next
    End Sub

    Public Sub RemoveItem(item)
        If Not IsNull(item) Then
            If m_items.Exists(item.Name) Then
                    item.ResetInterval
                    item.CurrentIdx = 0
                    m_items.Remove item.Name
            End If
        End If
    End Sub

    Public Sub NextItem()
        Dim items: items = m_items.Items
        If items(m_currentItemIdx).Repeat = False Then
            RemoveItem(items(m_currentItemIdx))
        Else
            m_currentItemIdx = m_currentItemIdx + 1
        End If
        
        If m_currentItemIdx > UBound(m_items.Items) Then   
            m_currentItemIdx = 0
        End If
    End Sub

    Public Function HasSeq(name)
        If m_items.Exists(name) Then
            HasSeq = True
        Else
            HasSeq = False
        End If
    End Function

End Class


Function Lerp(startValue, endValue, amount)
    Lerp = startValue + (endValue - startValue) * amount
End Function

Function Flip(x)
    Flip = 1 - x
End Function

Function EaseIn(amount)
    EaseIn = amount * amount
End Function

Function EaseOut(amount)
    EaseOut = Flip(Sqr(Flip(amount)))
End Function

Function EaseInOut(amount)
    EaseInOut = Lerp(EaseIn(amount), EaseOut(amount), amount)
End Function

Function Spike(t)
    If t <= 0.5 Then
        Spike = EaseIn(t / 0.5)
    Else
        Spike = EaseIn(Flip(t)/0.5)
    End If
End Function

'****************************************************************
'	WLSC: Custom Light Seqs for Saving Wallden
'****************************************************************

'*** BEGIN lights-out ***

Dim lSeqblacksmithMiniWizard : Set lSeqblacksmithMiniWizard = New LCSeq
lSeqblacksmithMiniWizard.Name = "lSeqblacksmithMiniWizard"
lSeqblacksmithMiniWizard.Sequence = Array( Array("StandupsTL|100|847585","DropTargetsCHL|100|FDDFFE","RightRampSL|100|0A090A","HoleShotL|100|473F47"), _
Array("StandupsTL|0|000000","DropTargetsPlusL|100|0E0D0E","DropTargetsASL|100|FDDFFE","DropTargetsCHL|100|0C0B0C","RightRampSL|100|FDDFFE","HoleShotL|100|4E444E"), _
Array("StandupsRL|100|F8DAF9","DropTargetsHL|100|3F383F","DropTargetsPlusL|100|FDDFFE","DropTargetsERL|100|040404","DropTargetsCHL|0|000000","RightRampAL|100|FADCFB","RightRampPowerupL|100|FDDFFE","HoleShotL|100|5B505B"), _
Array("CrossbowRightStandupL|100|010101","StandupsAL|100|675B67","StandupsRL|0|000000","StandupsTL|100|FCDEFD","DropTargetsHL|100|FDDFFE","DropTargetsPlusL|100|030303","DropTargetsASL|0|000000","DropTargetsERL|100|FDDFFE","RightRampAL|100|FDDFFE","RightRampSL|100|2F292F","HoleShotL|100|5A4F5A"), _
Array("CrossbowRightStandupL|100|FDDFFE","StandupsAL|100|443C45","StandupsTL|100|FDDFFE","DropTargetsPL|100|FDDFFE","DropTargetsHL|100|DCC1DC","DropTargetsPlusL|0|000000","RightRampEL|100|FDDFFE","RightRampSL|0|000000","HoleShotL|100|524953","RightRampShotL|100|FDDFFE","Light003|100|FDDFFE","Light002|100|FDDFFE"), _
Array("CrossbowRightStandupL|0|000000","CrossbowCenterStandupL|100|151316","StandupsIL|100|FDDFFE","StandupsAL|0|000000","StandupsRL|100|E7CCE8","DropTargetsHL|0|000000","DropTargetsERL|0|000000","DropTargetsCHL|100|A894A9","RightRampPowerupL|0|000000","HoleShotL|100|584E59","Light003|100|060606"), _
Array("CrossbowCenterStandupL|100|FDDFFE","StandupsIL|0|000000","StandupsRL|100|FDDFFE","DropTargetsPL|100|010101","DropTargetsCHL|100|FDDFFE","LeftHorseshoeKL|100|FDDFFE","RightRampAL|0|000000","HoleShotL|100|665A66","Light003|0|000000","Light002|100|DEC4DF"), _
Array("CrossbowCenterStandupL|100|1F1C1F","CrossbowLeftStandupL|100|594E59","StandupsNL|100|FDDFFE","StandupsAL|100|030304","DropTargetsPL|0|000000","LeftHorseshoeCL|100|3F373F","HoleShotL|100|6F6270","RightRampShotL|100|D6BDD7","Light002|0|000000"), _
Array("CrossbowCenterStandupL|0|000000","CrossbowLeftStandupL|100|FDDFFE","StandupsNL|0|000000","StandupsAL|100|FDDFFE","DropTargetsPlusL|100|665A66","DropTargetsASL|100|050505","LeftHorseshoeCL|100|FDDFFE","RightRampEL|100|DAC0DB","HoleShotL|100|776977","RightRampShotL|0|000000","GICrossbowLeft|100|F5D8F6"), _
Array("CrossbowRightStandupL|100|D7BDD8","CrossbowLeftStandupL|100|FCDEFD","StandupsTL|100|EBD0EC","DropTargetsPlusL|100|FDDFFE","DropTargetsASL|100|EFD3F0","LeftHorseshoeKL|100|F0D3F1","LeftHorseshoeAL|100|FDDFFE","RightRampEL|0|000000","HoleShotL|100|7C6E7D","GICrossbowLeft|100|FDDFFE","Light003|100|AF9AB0"), _
Array("CrossbowRightStandupL|100|FDDFFE","CrossbowLeftStandupL|0|000000","StandupsIL|100|F2D5F3","StandupsTL|100|0E0C0E","DropTargetsASL|100|FDDFFE","LeftHorseshoeKL|0|000000","LeftHorseshoeLL|100|020102","LeftOrbitGL|100|FDDFFE","RightRampSL|100|030303","HoleShotL|100|817281","GICrossbowLeft|100|D4BBD5","Light003|100|FDDFFE"), _
Array("StandupsIL|100|FDDFFE","StandupsTL|0|000000","DropTargetsHL|100|9A879A","LeftHorseshoeCL|100|FCDEFD","LeftHorseshoeLL|100|FDDFFE","LeftOrbitI2L|100|FCDEFD","RightRampSL|100|594E59","LeftOrbitPowerupL|100|FDDFFE","HoleShotL|100|847585","CrossbowShotL|100|FDDFFE","GICrossbowLeft|0|000000"), _
Array("DropTargetsHL|100|FDDFFE","RightHorseshoeHL|100|FDDFFE","LeftHorseshoeCL|0|000000","LeftHorseshoeBL|100|E8CCE9","LeftOrbitBumpersL|100|B39EB4","LeftOrbitLockL|100|FDDFFE","LeftOrbitGL|100|5F5460","LeftOrbitI2L|100|FDDFFE","LeftOrbitI1L|100|403941","RightRampSL|100|F4D7F5","HoleShotL|100|877788","LeftOrbitShotL|100|FDDFFE"), _
Array("CrossbowCenterStandupL|100|E0C5E1","StandupsNL|100|FDDFFE","DropTargetsERL|100|F8DAF9","RightHorseshoeTL|100|FDDFFE","LeftHorseshoeBL|100|FDDFFE","LeftOrbitBumpersL|100|FDDFFE","LeftOrbitGL|0|000000","LeftOrbitI1L|100|FDDFFE","RightRampSL|100|FDDFFE","RightRampPowerupL|100|020102","HoleShotL|100|8D7C8E","CrossbowShotL|0|000000"), _
Array("CrossbowCenterStandupL|100|FDDFFE","StandupsRL|100|837484","DropTargetsPL|100|060506","DropTargetsERL|100|FDDFFE","RightHorseshoeIL|100|C1AAC2","LeftHorseshoeAL|0|000000","LeftOrbitLockL|100|AB96AB","LeftOrbitI2L|0|000000","RightRampPowerupL|100|897989","CrossbowPowerupL|100|FDDFFE","LeftOrbitPowerupL|100|030203","HoleShotL|100|928193","Light002|100|0B0A0B"), _
Array("StandupsRL|100|010101","DropTargetsPL|100|D2B9D3","RightHorseshoeIL|100|FDDFFE","RightHorseshoeML|100|090809","RightOrbitBumpersL|100|817181","LeftOrbitBumpersL|0|000000","LeftOrbitLockL|0|000000","RightOrbitVL|100|FDDFFE","LeftOrbitI1L|100|E1C6E1","RightRampPowerupL|100|FDDFFE","LeftOrbitPowerupL|0|000000","HoleShotL|100|968497","LeftHorseshoeShotL|100|FDDFFE","LeftOrbitShotL|100|615561","Light002|100|FDDFFE"), _
Array("StandupsRL|0|000000","DropTargetsPL|100|FDDFFE","RightHorseshoeML|100|FDDFFE","LeftHorseshoeLL|100|453D45","RightOrbitBumpersL|100|FDDFFE","CrossbowERL|100|FDDFFE","CrossbowIPL|100|9E8B9F","RightOrbitKL|100|F9DBF9","LeftOrbitI1L|0|000000","CrossbowPowerupL|100|4E454F","HoleShotL|100|9A889B","LeftOrbitShotL|0|000000","GIRightRamp|100|FDDFFE","GILeftOrbit2|100|FDDFFE","GILeftOrbit1|100|FDDFFE"), _
Array("CrossbowLeftStandupL|100|EACEEB","DropTargetsCHL|100|E8CCE8","RightHorseshoeSL|100|FDDFFE","LeftHorseshoeLL|0|000000","CrossbowSNL|100|FDDFFE","CrossbowIPL|100|FDDFFE","RightOrbitKL|100|FDDFFE","RightRampAL|100|746674","CrossbowPowerupL|0|000000","HoleShotL|100|9E8B9E","GIBumperInlanes4|100|887889"), _
Array("CrossbowLeftStandupL|100|FDDFFE","StandupsAL|100|FADCFB","DropTargetsCHL|100|99879A","LeftHorseshoeBL|100|DDC3DE","CrossbowERL|0|000000","RightOrbitNL|100|FDDFFE","RightRampAL|100|EACEEA","RightOrbitPowerupL|100|FDDFFE","HoleShotL|100|A18EA2","RightRampShotL|100|020202","GICrossbowLeft|100|EDD1EE","GIBumperInlanes4|100|FDDFFE","GIBumperInlanes3|100|FDDFFE","GILeftOrbit3|100|151315"), _
Array("StandupsAL|100|5B505C","DropTargetsCHL|100|322C33","RightHorseshoeHL|100|E8CDE9","LeftHorseshoeBL|0|000000","CrossbowSNL|100|BFA8C0","CrossbowIPL|0|000000","RightRampAL|100|FDDFFE","HoleShotL|100|A490A4","RightHorseshoeShotL|100|FCDEFD","RightRampShotL|100|8B7B8C","GICrossbowLeft|100|FDDFFE","GIBumperInlanes2|100|EDD1EE","GILeftOrbit3|100|FDDFFE"), _
Array("StandupsAL|100|020202","DropTargetsCHL|0|000000","RightHorseshoeHL|0|000000","ScoopDRL|100|D7BDD8","CrossbowSNL|0|000000","LeftOrbitGL|100|040404","HoleShotL|100|A692A6","RightOrbitShotL|100|F9DBFA","RightHorseshoeShotL|100|FDDFFE","RightRampShotL|100|F7DAF8","CrossbowShotL|100|3D353D","GIBumperInlanes2|100|FDDFFE","GILeftOrbit1|100|1E1B1E"), _
Array("StandupsAL|0|000000","RightHorseshoeTL|100|403840","LeftHorseshoeKL|100|010101","ScoopDRL|100|FDDFFE","LeftOrbitGL|100|F1D4F2","ScoopPowerupL|100|665A67","HoleShotL|100|A894A9","ScoopShotL|100|786A78","RightOrbitShotL|100|FDDFFE","RightRampShotL|100|FDDFFE","CrossbowShotL|100|FDDFFE","GIBumperInlanes1|100|AB97AC","GILeftOrbit2|0|000000","GILeftOrbit1|0|000000"), _
Array("RightHorseshoeTL|0|000000","LeftHorseshoeKL|100|8C7C8D","ScoopAGL|100|FDDFFE","LeftOrbitGL|100|FDDFFE","RightRampEL|100|090809","ScoopPowerupL|100|FDDFFE","HoleShotL|100|AA96AB","ScoopShotL|100|FDDFFE","LeftHorseshoeShotL|100|342E34","LeftRampShotL|100|FCDEFD","Light003|100|FBDDFC","GIBumperInlanes1|100|FDDFFE"), _
Array("SpinnerShotL|100|EACEEB","StandupsIL|100|F3D7F4","RightHorseshoeIL|100|201C20","LeftHorseshoeKL|100|FBDEFC","ScoopLockL|100|FDDFFE","LeftOrbitLockL|100|2B262C","ScoopONL|100|3D363D","LeftOrbitI2L|100|8D7C8D","RightRampEL|100|6A5D6A","HoleShotL|100|AC98AD","LeftHorseshoeShotL|0|000000","LeftRampShotL|100|FDDFFE","GIScoop1|100|C6AFC7","Light003|100|B29DB2"), _
Array("SpinnerShotL|100|FDDFFE","CrossbowRightStandupL|100|907F91","StandupsIL|100|695D6A","RightHorseshoeIL|0|000000","RightHorseshoeML|100|FBDDFC","LeftHorseshoeKL|100|FDDFFE","LeftOrbitBumpersL|100|E5C9E5","LeftOrbitLockL|100|FDDFFE","ScoopONL|100|FDDFFE","RightOrbitVL|100|F5D8F6","LeftOrbitI2L|100|FDDFFE","RightRampEL|100|D7BED8","LeftRampPL|100|141214","CrossbowPowerupL|100|010101","LeftOrbitPowerupL|100|BFA8BF","LeftRampPowerupL|100|FDDFFE","HoleShotL|100|AE99AF","GIScoop1|100|FDDFFE","GIRightRamp|100|453D46","Light003|100|171517","GILeftOrbit3|100|FBDDFC","GIRightRubbersUpper|100|1D1A1D"), _
Array("CrossbowRightStandupL|100|1D1A1D","StandupsIL|0|000000","DropTargetsPlusL|100|D7BED8","DropTargetsASL|100|FBDEFC","RightHorseshoeML|100|090809","LeftOrbitBumpersL|100|FDDFFE","RightOrbitVL|100|252025","RightRampEL|100|FCDEFD","LeftRampPL|100|FDDFFE","LeftRampCL|100|080708","CrossbowPowerupL|100|F5D8F6","LeftOrbitPowerupL|100|FDDFFE","HoleShotL|100|AF9AB0","LeftRampShotL|100|F8DBF9","GIScoop2|100|FDDFFE","GIRightRamp|0|000000","Light003|0|000000","GIBumperInlane2|100|6F626F","GIVUK|100|1E1B1E","GIBumperInlanes4|100|211D21","GIBumperInlanes3|100|5D525E","GILeftOrbit3|0|000000","GIRightRubbersUpper|100|010101"), _
Array("CrossbowRightStandupL|0|000000","DropTargetsPlusL|100|A38FA3","DropTargetsASL|100|D5BCD6","RightHorseshoeML|0|000000","RightHorseshoeSL|100|DDC2DD","LeftHorseshoeCL|100|3D363E","RightOrbitBumpersL|100|BBA4BB","RightOrbitVL|0|000000","LeftOrbitI1L|100|EBD0EC","RightRampEL|100|FDDFFE","LeftRampCL|100|FDDFFE","LeftRampEL|100|6A5E6B","CrossbowPowerupL|100|FDDFFE","HoleShotL|100|B19CB2","LeftOrbitShotL|100|010101","LeftRampShotL|0|000000","GIBumperInlane2|100|FDDFFE","GIVUK|100|FDDFFE","GIBumperInlanes4|0|000000","GIBumperInlanes3|0|000000","GILeftRubbersUpper|100|F0D3F1","GIRightRubbersUpper|0|000000"), _
Array("ObjectiveFinalJudgmentL|100|AE9AAF","DropTargetsPlusL|100|4C434C","DropTargetsASL|100|A894A9","RightHorseshoeSL|0|000000","LeftHorseshoeCL|100|DAC0DB","RightOrbitBumpersL|100|010001","CrossbowERL|100|F5D8F6","RightOrbitKL|100|817181","LeftOrbitI1L|100|FDDFFE","LeftRampEL|100|FDDFFE","LeftRampPowerupL|100|FBDDFC","HoleShotL|100|B29DB3","LeftOrbitShotL|100|C3ACC4","GIBumperInlanes2|100|3A333A","GILeftRubbersUpper|100|FDDFFE"), _
Array("ObjectiveFinalJudgmentL|100|FDDFFE","StandupsNL|100|F7D9F8","DropTargetsPlusL|100|010101","DropTargetsASL|100|594E59","LeftHorseshoeCL|100|FCDEFD","RightOrbitBumpersL|0|000000","CrossbowERL|100|FDDFFE","CrossbowIPL|100|5E535E","RightOrbitKL|0|000000","LeftRampPL|100|413941","LeftRampPowerupL|0|000000","HoleShotL|100|B39EB4","LeftOrbitShotL|100|FDDFFE","GIRightOrbit1|100|6A5D6A","GIBumperInlanes2|0|000000"), _
Array("ObjectiveSniperL|100|FDDFFE","ObjectiveBlacksmithL|100|FDDFFE","StandupsNL|100|837383","DropTargetsPlusL|0|000000","DropTargetsASL|100|0F0E0F","LeftHorseshoeCL|100|FDDFFE","CrossbowSNL|100|5E535E","CrossbowIPL|100|FDDFFE","RightRampSL|100|F6D9F7","LeftRampPL|0|000000","LeftRampCL|100|B39EB4","RightOrbitPowerupL|100|89798A","HoleShotL|100|B59FB5","RightHorseshoeShotL|100|F3D6F4","GIUpperFlipper|100|3F383F","GIRightOrbit3|100|AC98AD","GIRightOrbit2|100|695D6A","GIRightOrbit1|100|FDDFFE","GILeftOrbit4|100|FDDFFE","Light001|100|8C7C8D","GIRightRubbersLower|100|907F90"), _
Array("ObjectiveFinalJudgmentL|100|827383","StandupsNL|100|040404","StandupsTL|100|020202","DropTargetsASL|0|000000","LeftHorseshoeAL|100|010101","CrossbowSNL|100|FDDFFE","RightOrbitNL|100|181518","RightRampSL|100|E7CCE8","LeftRampCL|0|000000","LeftRampEL|100|B09BB1","RightOrbitPowerupL|0|000000","HoleShotL|100|B6A0B6","RightHorseshoeShotL|100|040404","GIUpperFlipper|100|FDDFFE","GIRightOrbit3|100|FDDFFE","GIRightOrbit2|100|FDDFFE","GIBumperInlanes1|100|0A090A","Light001|100|FDDFFE","GIRightRubbersLower|100|C0A9C1"), _
Array("ObjectiveFinalJudgmentL|0|000000","ObjectiveVikingL|100|EACEEB","ObjectiveFinalBossL|100|FDDFFE","ObjectiveEscapeL|100|938193","StandupsNL|0|000000","StandupsTL|100|040404","DropTargetsHL|100|F3D6F4","LeftHorseshoeAL|100|5D525D","RightOrbitNL|0|000000","RightRampSL|100|D7BED8","LeftRampEL|0|000000","HoleShotL|100|B7A1B8","RightHorseshoeShotL|0|000000","GIBumperInlanes1|0|000000","GILeftRubbersUpper|100|060506","GIRightRubbersLower|0|000000"), _
Array("ObjectiveSniperL|100|030304","ObjectiveBlacksmithL|100|F2D5F3","ObjectiveVikingL|100|FDDFFE","ObjectiveEscapeL|100|FDDFFE","ObjectiveChaserL|100|6F626F","ObjectiveDragonL|100|615662","StandupsTL|100|080708","DropTargetsHL|100|DCC2DD","LeftHorseshoeAL|100|EACEEB","RightRampSL|100|C6AFC7","HoleShotL|100|B7A2B8","RightOrbitShotL|100|A995A9","GILeftRubbersUpper|0|000000"), _
Array("ObjectiveSniperL|0|000000","ObjectiveBlacksmithL|0|000000","ObjectiveVikingL|100|F2D5F3","ObjectiveChaserL|100|FDDFFE","ObjectiveDragonL|100|FDDFFE","CrossbowCenterStandupL|100|BFA8BF","StandupsTL|100|2A252A","DropTargetsHL|100|B7A2B8","LeftHorseshoeAL|100|FCDFFE","RightRampSL|100|B49FB5","HoleShotL|100|B9A3B9","RightOrbitShotL|100|020102","GILeftOrbit1|100|0F0D0F"), _
Array("ObjectiveVikingL|0|000000","ObjectiveFinalBossL|100|0A090A","ObjectiveCastleL|100|FDDFFE","CrossbowCenterStandupL|100|5C515C","StandupsTL|100|4B424C","DropTargetsHL|100|7D6E7E","DropTargetsERL|100|F6D9F7","LeftHorseshoeAL|100|FDDFFE","LeftHorseshoeLL|100|010001","RightRampSL|100|998799","RightRampPowerupL|100|FCDEFD","HoleShotL|100|B9A3BA","RightOrbitShotL|0|000000","GILeftOrbit1|100|E7CCE8","Light001|100|887888","GIRightRubbersUpper|100|100E10","GIRightSlingshotUpper|100|FDDFFE","GILeftSlingshot|100|110F11"), _
Array("ObjectiveFinalBossL|0|000000","ObjectiveEscapeL|100|AB97AC","ObjectiveChaserL|100|E1C6E2","CrossbowCenterStandupL|100|161316","StandupsTL|100|6C5F6D","DropTargetsHL|100|484048","DropTargetsERL|100|EBCFEC","LeftHorseshoeLL|100|221E22","RightRampSL|100|796A79","RightRampPowerupL|100|FBDEFC","HoleShotL|100|BAA4BB","GIBumperInlane2|100|F4D8F5","GILeftOrbit1|100|FDDFFE","Light001|0|000000","GIRightRubbersUpper|100|FBDDFC","GILeftSlingshot|100|FDDFFE"), _
Array("SpinnerShotL|100|B9A3BA","ObjectiveEscapeL|0|000000","ObjectiveChaserL|0|000000","ObjectiveDragonL|100|847585","CrossbowCenterStandupL|0|000000","StandupsTL|100|8B7B8C","DropTargetsHL|100|2B262B","DropTargetsERL|100|C8B0C9","LeftHorseshoeLL|100|AF9BB0","ScoopDRL|100|B8A2B9","RightRampSL|100|5C515C","RightRampPowerupL|100|E8CCE9","HoleShotL|100|BBA5BC","ScoopShotL|100|ECD0ED","GIBumperInlane2|100|020102","GILeftOrbit2|100|8A7A8B","GIRightRubbersUpper|100|FDDFFE","GIRightSlingshotUpper|100|121012"), _
Array("SpinnerShotL|100|090809","ObjectiveDragonL|0|000000","ObjectiveCastleL|100|E4C9E5","StandupsTL|100|A793A7","DropTargetsHL|100|0F0D0F","DropTargetsERL|100|8F7E8F","LeftHorseshoeLL|100|F3D6F4","ScoopDRL|100|060506","RightRampSL|100|423A42","ScoopPowerupL|100|F1D5F2","RightRampPowerupL|100|C4ACC4","ScoopShotL|100|413941","GIScoop1|100|E1C7E2","GIBumperInlane2|0|000000","GIVUK|100|332D33","GILeftOrbit2|100|F7DAF8","GIRightSlingshotUpper|0|000000"), _
Array("SpinnerShotL|0|000000","ObjectiveCastleL|0|000000","StandupsTL|100|BAA4BB","DropTargetsHL|0|000000","DropTargetsERL|100|504650","LeftHorseshoeLL|100|FDDFFE","ScoopDRL|0|000000","ScoopAGL|100|FCDEFD","RightRampSL|100|2A252B","ScoopPowerupL|100|2F2A2F","RightRampPowerupL|100|A28EA2","HoleShotL|100|BCA6BD","ScoopShotL|0|000000","LeftRampShotL|100|060506","GIScoop1|100|262226","GIVUK|0|000000","GILeftOrbit2|100|FDDFFE","GIRightSlingshotLow|100|A591A5"), _
Array("StandupsTL|100|CBB3CC","DropTargetsPL|100|FBDDFC","DropTargetsERL|100|262126","LeftHorseshoeBL|100|060606","ScoopAGL|100|8B7B8C","RightRampSL|100|151215","ScoopPowerupL|0|000000","RightRampPowerupL|100|817281","HoleShotL|100|BDA6BD","LeftRampShotL|100|D3BAD4","GIScoop1|0|000000","GILeftSlingshotLower|100|F4D7F4","GILeftSlingshot|100|BEA8BF","GIRightSlingshotLow|100|FDDFFE"), _
Array("StandupsTL|100|DAC0DB","DropTargetsPL|100|F8DBF9","DropTargetsERL|100|131013","LeftHorseshoeBL|100|322D33","ScoopLockL|100|FADCFB","ScoopAGL|0|000000","RightRampAL|100|FCDEFD","RightRampSL|100|010101","RightRampPowerupL|100|615662","HoleShotL|100|BEA7BE","LeftRampShotL|100|FDDFFE","Light002|100|E7CBE7","GILeftOrbit4|100|4E444E","GILeftSlingshotLower|100|FDDFFE","GILeftSlingshot|0|000000"), _
Array("RightOutlanePoisonL|100|C5ADC6","StandupsTL|100|E7CCE8","DropTargetsPL|100|F6D8F7","DropTargetsERL|100|060606","LeftHorseshoeBL|100|CAB2CB","ScoopLockL|100|A491A5","ScoopONL|100|E3C8E4","RightRampAL|100|F9DBFA","RightRampSL|0|000000","LeftRampPL|100|292429","RightRampPowerupL|100|443C44","LeftRampPowerupL|100|443C44","HoleShotL|100|BEA8BF","Light002|100|CAB3CB","GILeftOrbit4|0|000000","GIRightSlingshotLow|100|615661"), _
Array("LeftOutlanePoisonL|100|BCA6BD","RightOutlanePoisonL|100|FDDFFE","ObjectiveFinalJudgmentL|100|040304","CrossbowLeftStandupL|100|E3C8E4","StandupsRL|100|040404","StandupsTL|100|F3D6F4","DropTargetsPL|100|C9B1CA","DropTargetsERL|0|000000","LeftHorseshoeBL|100|F4D7F5","ScoopLockL|100|141214","ScoopONL|100|443C44","RightRampAL|100|F5D8F6","LeftRampPL|100|F3D6F3","RightRampPowerupL|100|282328","LeftRampPowerupL|100|F7DAF8","HoleShotL|100|BFA8BF","GIScoop2|100|5C515C","Light002|100|A995AA","GIRightOrbit1|100|685C69","GIRightRubbersLower|100|948294","GIRightInlaneUpper|100|C4ADC5","GIRightSlingshotLow|0|000000"), _
Array("LeftOutlanePoisonL|100|FDDFFE","RightOutlanePoisonL|100|B49FB5","ObjectiveFinalJudgmentL|100|968496","CrossbowLeftStandupL|100|8F7E90","StandupsRL|100|0F0D0F","StandupsTL|100|FDDFFE","DropTargetsPL|100|9F8C9F","RightHorseshoeHL|100|030303","LeftHorseshoeBL|100|FDDFFE","ScoopLockL|0|000000","ScoopONL|100|010101","RightRampAL|100|F0D3F1","LeftRampPL|100|FDDFFE","LeftRampCL|100|514751","RightRampPowerupL|100|0F0E0F","LeftRampPowerupL|100|FDDFFE","GIScoop2|0|000000","GICrossbowLeft|100|F3D6F4","Light002|100|6C5F6C","GIRightOrbit3|100|F9DCFA","GIRightOrbit1|0|000000","GIRightRubbersLower|100|FDDFFE","GIRightInlaneUpper|100|FDDFFE","GIRightInlaneLower|100|FDDFFE","GILeftInlaneLower|100|060506","GILeftInlaneUpper|100|C0A9C1"), _
Array("RightOutlanePoisonL|0|000000","ObjectiveFinalJudgmentL|100|FDDFFE","ObjectiveSniperL|100|030303","CrossbowLeftStandupL|100|4C434C","StandupsRL|100|2A252A","DropTargetsPL|100|766876","RightHorseshoeHL|100|262226","ScoopONL|0|000000","RightRampAL|100|E0C5E1","LeftRampCL|100|FDDFFE","LeftRampEL|100|040304","RightRampPowerupL|100|0C0A0C","HoleShotL|100|BFA9C0","CrossbowShotL|100|F9DBFA","GICrossbowLeft|100|A693A7","Light002|100|332D33","GIRightOrbit3|100|6A5E6B","GIRightOrbit2|100|FADDFB","GILeftOrbit3|100|2B262B","GILeftInlaneLower|100|F6D9F7","GILeftInlaneUpper|100|FDDFFE","GILeftSlingshotLower|100|655965"), _
Array("ObjectiveSniperL|100|B09BB1","CrossbowLeftStandupL|100|151215","StandupsRL|100|473F47","DropTargetsPL|100|4E454F","RightHorseshoeHL|100|695D69","RightRampAL|100|CEB5CE","LeftRampEL|100|A591A6","RightRampPowerupL|100|080708","HoleShotL|100|C0A9C1","CrossbowShotL|100|D6BCD6","GICrossbowLeft|100|5F545F","Light002|0|000000","GIRightOrbit3|0|000000","GIRightOrbit2|100|655965","GILeftOrbit3|100|BDA7BE","GIRightInlaneUpper|100|231F23","GILeftInlaneLower|100|FDDFFE","GILeftSlingshotLower|0|000000"), _
Array("ObjectiveSniperL|100|FDDFFE","ObjectiveBlacksmithL|100|141214","CrossbowLeftStandupL|0|000000","StandupsRL|100|665A66","DropTargetsPL|100|292429","RightHorseshoeHL|100|BAA4BB","RightRampAL|100|B5A0B6","LeftRampEL|100|FDDFFE","RightRampPowerupL|100|050505","HoleShotL|100|C0AAC1","LeftHorseshoeShotL|100|040404","CrossbowShotL|100|7D6E7D","GICrossbowLeft|100|241F24","GIRightOrbit2|0|000000","GILeftOrbit3|100|FADCFB","GIRightInlaneUpper|0|000000","GIRightInlaneLower|100|A18EA2"), _
Array("ObjectiveBlacksmithL|100|D3BAD4","ObjectiveVikingL|100|847484","ObjectiveFinalBossL|100|020102","StandupsRL|100|867686","DropTargetsPL|100|040304","RightHorseshoeHL|100|F1D5F2","RightHorseshoeTL|100|040304","RightRampAL|100|9C899C","RightRampPowerupL|100|020202","HoleShotL|100|C1AAC2","RightRampShotL|100|FADDFB","LeftHorseshoeShotL|100|1B181C","CrossbowShotL|100|2B262B","GICrossbowLeft|0|000000","GILeftOrbit3|100|FDDFFE","GILeftRubbersUpper|100|403941","GIRightInlaneLower|0|000000"), _
Array("LeftOutlanePoisonL|100|B19CB2","ObjectiveBlacksmithL|100|FDDFFE","ObjectiveVikingL|100|FDDFFE","ObjectiveFinalBossL|100|7D6E7E","StandupsRL|100|A491A5","DropTargetsPL|0|000000","RightHorseshoeHL|100|F9DCFA","RightHorseshoeTL|100|625763","RightRampAL|100|807181","RightRampPowerupL|0|000000","HoleShotL|100|C1ABC2","RightRampShotL|100|F3D6F4","LeftHorseshoeShotL|100|958396","CrossbowShotL|100|030203","GILeftRubbersUpper|100|E3C8E4","GILeftInlaneUpper|100|FADDFB","ActionLeftL|100|594E59","ActionRightL|100|C9B2CA"), _
Array("LeftOutlanePoisonL|0|000000","ObjectiveFinalBossL|100|FDDFFE","ObjectiveChaserL|100|050405","StandupsRL|100|C2ABC2","DropTargetsCHL|100|010101","RightHorseshoeHL|100|FDDFFE","RightHorseshoeTL|100|D9BFD9","LeftOrbitGL|100|E8CDE9","RightRampAL|100|635864","RightRampShotL|100|EACEEB","LeftHorseshoeShotL|100|E5CAE6","CrossbowShotL|0|000000","GILeftRubbersUpper|100|FDDFFE","GILeftInlaneLower|100|C2ABC3","GILeftInlaneUpper|100|473E47","GIRightSlingshotUpper|100|030303","ActionLeftL|100|FDDFFE","ActionRightL|100|FDDFFE"), _
Array("ObjectiveChaserL|100|6F6270","StandupsRL|100|CCB4CD","DropTargetsCHL|100|050405","RightHorseshoeTL|100|F2D5F3","RightHorseshoeIL|100|030303","LeftOrbitGL|100|AB97AC","RightRampAL|100|736674","HoleShotL|100|A894A8","RightRampShotL|100|CBB3CC","LeftHorseshoeShotL|100|FDDFFE","GIBumperInlanes4|100|080708","GIBumperInlanes3|100|3E363E","GILeftInlaneLower|100|050405","GILeftInlaneUpper|0|000000","GIRightSlingshotUpper|100|5C515C"), _
Array("ObjectiveEscapeL|100|5E535E","ObjectiveChaserL|100|F6D9F7","ObjectiveDragonL|100|0A080A","StandupsRL|100|DBC1DC","DropTargetsCHL|100|060506","RightHorseshoeTL|100|FDDFFE","RightHorseshoeIL|100|2C272D","LeftOrbitGL|100|6D606D","RightRampAL|100|564C57","HoleShotL|100|A995A9","RightRampShotL|100|B39EB4","GIUpperFlipper|100|F4D7F5","GIRightRamp|100|030203","GIBumperInlanes4|100|322C33","GIBumperInlanes3|100|A28FA3","Light001|100|0C0B0D","GILeftInlaneLower|0|000000","GIRightSlingshotUpper|100|FDDFFE"), _
Array("ObjectiveEscapeL|100|F2D5F3","ObjectiveChaserL|100|FDDFFE","ObjectiveDragonL|100|BAA4BA","StandupsAL|100|010101","StandupsRL|100|E8CCE8","DropTargetsCHL|100|070607","RightHorseshoeIL|100|6F626F","LeftOrbitGL|100|322C32","RightRampAL|100|393239","CrossbowPowerupL|100|FBDEFC","HoleShotL|100|A995AA","RightRampShotL|100|8B7B8C","GIUpperFlipper|100|958395","GIRightRamp|100|312C32","GIBumperInlanes4|100|938194","GIBumperInlanes3|100|FDDFFE","GIBumperInlanes2|100|201C20","Light001|100|948395","ActionRightL|100|E0C6E1"), _
Array("ObjectiveEscapeL|100|FDDFFE","ObjectiveDragonL|100|FCDEFD","ObjectiveCastleL|100|5F5460","StandupsAL|100|040404","StandupsRL|100|F1D5F2","DropTargetsCHL|100|080708","RightHorseshoeIL|100|BDA7BE","LeftOrbitBumpersL|100|F7DAF8","LeftOrbitGL|0|000000","RightRampEL|100|FCDEFD","RightRampAL|100|221E22","CrossbowPowerupL|100|F3D6F4","HoleShotL|100|AA96AB","RightRampShotL|100|695C69","GIUpperFlipper|100|1D191D","GIRightRamp|100|9B899C","GIBumperInlanes4|100|E0C5E0","GIBumperInlanes2|100|776977","Light001|100|FDDFFE","ActionRightL|0|000000"), _
Array("ObjectiveDragonL|100|FDDFFE","ObjectiveCastleL|100|F1D5F2","StandupsAL|100|090809","StandupsRL|100|F7DAF8","RightHorseshoeIL|100|F1D5F2","RightHorseshoeML|100|010101","LeftOrbitBumpersL|100|E8CCE9","LeftOrbitLockL|100|DEC3DF","LeftOrbitI2L|100|F8DBF9","RightRampAL|100|0E0C0E","CrossbowPowerupL|100|B39EB4","HoleShotL|100|AB97AC","RightRampShotL|100|49414A","GIUpperFlipper|0|000000","GIRightRamp|100|DDC3DE","GIBumperInlanes4|100|FDDFFE","GIBumperInlanes2|100|DDC2DE","ActionLeftL|100|9B889B"), _
Array("ObjectiveCastleL|100|FDDFFE","StandupsAL|100|0E0D0E","StandupsRL|100|F9DCFA","DropTargetsCHL|100|100E10","RightHorseshoeIL|100|FADCFB","RightHorseshoeML|100|4E454E","LeftOrbitBumpersL|100|BBA4BB","LeftOrbitLockL|100|B19CB1","LeftOrbitI2L|100|C0AAC1","RightRampEL|100|FADCFB","RightRampAL|100|0A090A","CrossbowPowerupL|100|5B505B","HoleShotL|100|AC98AD","RightRampShotL|100|2D282E","GIRightRamp|100|F8DAF9","GIBumperInlanes2|100|F9DBFA","GIBumperInlanes1|100|010101","ActionLeftL|100|040304"), _
Array("StandupsAL|100|151215","StandupsRL|100|FADDFB","DropTargetsCHL|100|171417","RightHorseshoeIL|100|FDDFFE","RightHorseshoeML|100|A591A5","LeftHorseshoeKL|100|F9DCFA","LeftOrbitBumpersL|100|685C69","LeftOrbitLockL|100|716371","RightOrbitVL|100|0F0E10","LeftOrbitI2L|100|877788","RightRampEL|100|F7DAF8","RightRampAL|100|080708","CrossbowPowerupL|100|171417","HoleShotL|100|AD98AD","RightRampShotL|100|151215","GIRightRamp|100|FDDFFE","GIBumperInlanes2|100|FDDFFE","GIBumperInlanes1|100|0C0B0C","ActionLeftL|0|000000"), _
Array("StandupsAL|100|211D21","StandupsRL|100|FCDEFD","DropTargetsCHL|100|1E1B1F","RightHorseshoeML|100|D8BFD9","RightHorseshoeSL|100|080709","LeftHorseshoeKL|100|EACEEB","LeftOrbitBumpersL|100|191619","LeftOrbitLockL|100|312B31","CrossbowERL|100|D6BDD7","RightOrbitVL|100|211D21","LeftOrbitI2L|100|504751","RightRampEL|100|F3D7F4","RightRampAL|100|060506","CrossbowPowerupL|0|000000","HoleShotL|100|AD99AE","RightRampShotL|100|0F0D0F","GIBumperInlanes1|100|6C5F6C","GILeftSlingshot|100|262126","GIRightSlingshotLow|100|373138"), _
Array("StandupsAL|100|3A333B","DropTargetsCHL|100|252125","RightHorseshoeML|100|F2D5F3","RightHorseshoeSL|100|211D22","LeftHorseshoeKL|100|D9BFDA","LeftOrbitBumpersL|100|050405","LeftOrbitLockL|0|000000","CrossbowERL|100|867787","RightOrbitVL|100|8A798A","LeftOrbitI2L|100|1D191D","RightRampEL|100|F0D3F0","RightRampAL|100|040404","LeftOrbitPowerupL|100|FBDDFC","HoleShotL|100|AE9AAF","RightRampShotL|100|0B0A0B","GIBumperInlanes1|100|EACEEB","GILeftSlingshot|100|CDB4CE","GIRightSlingshotLow|100|E8CCE9"), _
Array("StandupsAL|100|544A55","StandupsRL|100|FDDFFE","DropTargetsCHL|100|2C272C","RightHorseshoeML|100|FDDFFE","RightHorseshoeSL|100|5C515C","LeftHorseshoeKL|100|C7B0C8","RightOrbitBumpersL|100|030303","LeftOrbitBumpersL|100|010101","CrossbowERL|100|3A333B","RightOrbitVL|100|EFD3F0","LeftOrbitI2L|0|000000","RightRampEL|100|EBCFEC","RightRampAL|100|030203","LeftOrbitPowerupL|100|BDA7BE","HoleShotL|100|AF9AAF","RightRampShotL|100|080708","GIBumperInlanes1|100|F8DBF9","GILeftSlingshot|100|FDDFFE","GIRightSlingshotLow|100|FDDFFE"), _
Array("RightOutlanePoisonL|100|655965","StandupsAL|100|6F616F","DropTargetsCHL|100|332D33","RightHorseshoeSL|100|AA96AB","LeftHorseshoeKL|100|B19CB2","RightOrbitBumpersL|100|0A090A","LeftOrbitBumpersL|0|000000","CrossbowERL|100|090809","CrossbowIPL|100|E7CCE8","RightOrbitVL|100|F9DBFA","RightOrbitKL|100|030303","RightRampEL|100|E2C8E3","RightRampAL|100|020102","LeftOrbitPowerupL|100|817282","HoleShotL|100|B09BB0","RightRampShotL|100|050505","GIBumperInlanes1|100|FDDFFE"), _
Array("RightOutlanePoisonL|100|FDDFFE","StandupsAL|100|897989","DropTargetsCHL|100|393239","RightHorseshoeSL|100|EBCFEC","LeftHorseshoeKL|100|8A7A8B","RightOrbitBumpersL|100|2A252A","CrossbowERL|100|010101","CrossbowIPL|100|AC98AD","RightOrbitVL|100|FCDEFD","RightOrbitKL|100|0A090A","LeftOrbitI1L|100|D6BDD7","RightRampEL|100|D1B9D2","RightRampAL|100|010101","LeftOrbitPowerupL|100|524852","HoleShotL|100|B09BB1","RightRampShotL|100|030303","ActionButton|100|CDB5CE"), _
Array("StandupsAL|100|A28FA3","DropTargetsCHL|100|3F383F","RightHorseshoeSL|100|FADCFB","LeftHorseshoeKL|100|645865","RightOrbitBumpersL|100|7E6F7E","CrossbowERL|0|000000","CrossbowSNL|100|F5D8F6","CrossbowIPL|100|544A55","RightOrbitVL|100|FDDFFE","RightOrbitKL|100|2D272D","LeftOrbitI1L|100|9F8CA0","RightRampEL|100|B9A3B9","RightRampAL|0|000000","LeftOrbitPowerupL|100|282328","HoleShotL|100|B19CB1","RightRampShotL|100|010101","GIRightInlaneUpper|100|231F23","ActionButton|100|FDDFFE"), _
Array("StandupsAL|100|B8A2B9","DropTargetsCHL|100|453D45","RightHorseshoeSL|100|FCDEFD","LeftHorseshoeKL|100|403840","RightOrbitBumpersL|100|CFB6D0","CrossbowSNL|100|CFB7D0","CrossbowIPL|100|151215","RightOrbitKL|100|827383","LeftOrbitI1L|100|6A5E6B","RightRampEL|100|A18EA1","RightOrbitPowerupL|100|010001","LeftOrbitPowerupL|0|000000","HoleShotL|100|B19CB2","RightHorseshoeShotL|100|0D0B0D","RightRampShotL|0|000000","GIRightRubbersUpper|100|CCB4CD","GIRightInlaneUpper|100|C1AAC2","GILeftSlingshotLower|100|151215"), _
Array("StandupsAL|100|CAB2CB","DropTargetsCHL|100|4B424B","RightHorseshoeSL|100|FDDFFE","LeftHorseshoeKL|100|1C191C","RightOrbitBumpersL|100|E2C7E3","CrossbowSNL|100|8F7E90","CrossbowIPL|100|050405","RightOrbitKL|100|D0B8D1","LeftOrbitI1L|100|3E363E","RightRampEL|100|8A7A8A","RightOrbitPowerupL|100|010101","HoleShotL|100|B29DB3","RightHorseshoeShotL|100|2A252A","LeftOrbitShotL|100|FADCFB","GIRightRubbersUpper|100|675B68","GIRightInlaneUpper|100|FDDFFE","GIRightInlaneLower|100|2F292F","GILeftSlingshotLower|100|B19CB1"), _
Array("StandupsIL|100|010101","StandupsAL|100|DAC0DA","DropTargetsCHL|100|504751","LeftHorseshoeKL|100|080708","RightOrbitBumpersL|100|F3D6F4","CrossbowSNL|100|3F383F","CrossbowIPL|0|000000","RightOrbitKL|100|E1C7E2","LeftOrbitI1L|100|141214","RightRampEL|100|756775","RightOrbitPowerupL|100|151215","RightHorseshoeShotL|100|615561","LeftOrbitShotL|100|F7DAF8","GIRightRubbersUpper|100|19161A","GIRightInlaneLower|100|C8B1C9","GILeftSlingshotLower|100|FDDFFE"), _
Array("StandupsIL|100|040304","StandupsAL|100|E7CBE8","DropTargetsCHL|100|564C57","LeftHorseshoeKL|100|050405","LeftHorseshoeCL|100|FADDFB","RightOrbitBumpersL|100|FCDEFD","CrossbowSNL|0|000000","RightOrbitKL|100|F3D6F4","RightOrbitNL|100|050405","LeftOrbitI1L|0|000000","RightRampEL|100|615561","RightOrbitPowerupL|100|292429","HoleShotL|100|B39EB4","RightHorseshoeShotL|100|A894A9","LeftOrbitShotL|100|EBCFEC","GIRightRubbersUpper|0|000000","GIRightInlaneLower|100|FDDFFE","ActionButton|100|FBDDFC"), _
Array("StandupsIL|100|070607","StandupsAL|100|F2D5F3","DropTargetsCHL|100|5C515D","LeftHorseshoeKL|100|020202","LeftHorseshoeCL|100|F0D4F1","RightOrbitKL|100|FADDFB","RightOrbitNL|100|0A090A","RightRampEL|100|4E454E","RightOrbitPowerupL|100|665A66","RightHorseshoeShotL|100|E2C8E3","LeftOrbitShotL|100|BDA7BE","ActionButton|100|A591A5"), _
Array("StandupsIL|100|0F0D0F","StandupsAL|100|FBDEFC","DropTargetsCHL|100|665A66","LeftHorseshoeKL|100|010101","LeftHorseshoeCL|100|E5CAE6","RightOrbitBumpersL|100|FDDFFE","RightOrbitKL|100|FCDEFD","RightOrbitNL|100|383138","RightRampEL|100|3C343C","RightOrbitPowerupL|100|B7A1B7","HoleShotL|100|B49EB4","RightHorseshoeShotL|100|FDDFFE","LeftOrbitShotL|100|918091","ActionButton|100|010101"), _
Array("StandupsIL|100|1E1A1E","StandupsAL|100|FDDFFE","DropTargetsCHL|100|6F6270","LeftHorseshoeKL|0|000000","LeftHorseshoeCL|100|D9BFDA","RightOrbitKL|100|FDDFFE","RightOrbitNL|100|7F7080","RightRampEL|100|2B252B","RightOrbitPowerupL|100|F6D9F7","HoleShotL|100|B49FB5","LeftOrbitShotL|100|675B68","Light003|100|09080A","GIBumperInlane2|100|030303","ActionButton|0|000000"), _
Array("LeftOutlanePoisonL|100|0B0A0B","StandupsIL|100|2F292F","DropTargetsCHL|100|786A79","LeftHorseshoeCL|100|CCB4CD","RightOrbitNL|100|BEA7BF","RightRampEL|100|1C191C","RightOrbitPowerupL|100|F9DBFA","HoleShotL|100|B59FB5","RightOrbitShotL|100|040404","LeftOrbitShotL|100|403941","Light003|100|171417","GIBumperInlane2|100|433B43","GILeftInlaneLower|100|070607","GILeftInlaneUpper|100|0A090A"), _
Array("LeftOutlanePoisonL|100|766877","CrossbowRightStandupL|100|040404","StandupsIL|100|423A42","DropTargetsCHL|100|817282","LeftHorseshoeCL|100|B29DB3","RightOrbitNL|100|D4BBD5","RightRampEL|100|121012","RightOrbitPowerupL|100|FCDEFD","HoleShotL|100|B59FB6","RightOrbitShotL|100|090809","LeftOrbitShotL|100|1A171A","Light003|100|242024","GIBumperInlane2|100|887888","GILeftInlaneLower|100|564C56","GILeftInlaneUpper|100|766876"), _
Array("LeftOutlanePoisonL|100|EFD3F0","CrossbowRightStandupL|100|0B0A0B","StandupsIL|100|564C57","DropTargetsCHL|100|89798A","LeftHorseshoeCL|100|918091","RightOrbitNL|100|EBCFEC","RightRampEL|100|0A090A","RightOrbitPowerupL|100|FDDFFE","HoleShotL|100|B6A0B7","RightOrbitShotL|100|312B31","LeftOrbitShotL|0|000000","Light003|100|322C33","GIBumperInlane2|100|D5BBD6","GILeftInlaneLower|100|DDC3DE","GILeftInlaneUpper|100|E8CCE9"), _
Array("LeftOutlanePoisonL|100|FDDFFE","CrossbowRightStandupL|100|121012","StandupsIL|100|6D606E","DropTargetsCHL|100|928092","LeftHorseshoeCL|100|6F6270","RightOrbitNL|100|F8DBF9","RightRampEL|100|040304","RightRampPowerupL|100|010101","HoleShotL|100|B6A1B7","RightOrbitShotL|100|837383","Light003|100|413941","GIBumperInlane2|100|FBDEFC","GIVUK|100|020202","GILeftOrbit4|100|010101","GILeftInlaneLower|100|FDDFFE","GILeftInlaneUpper|100|FDDFFE","ActionRightL|100|584E58"), _
Array("CrossbowRightStandupL|100|1A171A","StandupsIL|100|847485","DropTargetsASL|100|010101","DropTargetsCHL|100|9A889B","LeftHorseshoeCL|100|4F464F","RightOrbitNL|100|FADCFB","RightRampEL|0|000000","HoleShotL|100|B7A1B8","RightOrbitShotL|100|D4BBD5","Light003|100|504751","GIBumperInlane2|100|FDDFFE","GIVUK|100|070607","GILeftOrbit4|100|050505","ActionRightL|100|ECD0ED"), _
Array("CrossbowRightStandupL|100|211D21","StandupsIL|100|9B889B","DropTargetsASL|100|040304","DropTargetsCHL|100|A28FA2","LeftHorseshoeCL|100|302A30","RightOrbitNL|100|FCDEFD","RightOrbitShotL|100|EDD1EE","Light003|100|5F5460","GIVUK|100|221E22","GILeftOrbit4|100|322C32","ActionRightL|100|FDDFFE"), _
Array("CrossbowRightStandupL|100|282328","StandupsNL|100|010101","StandupsIL|100|B09BB1","DropTargetsASL|100|060606","DropTargetsCHL|100|A995AA","LeftHorseshoeCL|100|141114","RightOrbitNL|100|FDDFFE","RightRampPowerupL|100|020202","RightOrbitShotL|100|F6D9F7","Light003|100|6F6270","GIVUK|100|514751","GILeftOrbit4|100|968496","ActionLeftL|100|090809"), _
Array("CrossbowRightStandupL|100|2F2A2F","StandupsIL|100|C3ACC4","DropTargetsPlusL|100|010101","DropTargetsASL|100|090809","DropTargetsCHL|100|B19CB2","LeftHorseshoeCL|100|0F0D0F","LeftHorseshoeAL|100|FADDFB","HoleShotL|100|B8A2B8","RightOrbitShotL|100|FDDFFE","Light003|100|807080","GIVUK|100|907F90","GILeftOrbit4|100|DEC4DF","GIRightRubbersLower|100|F5D8F6","ActionLeftL|100|504650"), _
Array("CrossbowRightStandupL|100|373037","StandupsNL|100|030303","StandupsIL|100|D2B9D3","DropTargetsPlusL|100|020202","DropTargetsASL|100|0C0A0C","DropTargetsCHL|100|B9A3BA","LeftHorseshoeCL|100|0A090A","LeftHorseshoeAL|100|F3D6F4","RightRampPowerupL|100|030203","HoleShotL|100|B8A2B9","Light003|100|8F7E90","GIVUK|100|D7BDD8","GILeftOrbit4|100|F4D7F5","GIRightRubbersLower|100|BBA5BC","ActionLeftL|100|EDD1EE"), _
Array("SpinnerShotL|100|181518","CrossbowRightStandupL|100|3E373E","StandupsNL|100|060606","StandupsIL|100|DFC5E0","DropTargetsPlusL|100|040304","DropTargetsASL|100|0E0D0E","DropTargetsCHL|100|C0A9C1","LeftHorseshoeCL|100|070607","LeftHorseshoeAL|100|EBCFEC","RightRampPowerupL|100|030303","HoleShotL|100|B8A3B9","Light003|100|9B899C","GIVUK|100|FDDFFE","GILeftOrbit4|100|FDDFFE","GIRightRubbersLower|100|776977","ActionLeftL|100|FDDFFE"), _
Array("SpinnerShotL|100|524852","CrossbowRightStandupL|100|4C434C","StandupsNL|100|0A090A","StandupsIL|100|EACFEB","DropTargetsPlusL|100|050505","DropTargetsASL|100|110F11","DropTargetsCHL|100|C7AFC7","LeftHorseshoeCL|100|040404","LeftHorseshoeAL|100|E2C8E3","HoleShotL|100|B9A3BA","Light003|100|A793A8","GILeftOrbit1|100|E3C8E4","GIRightRubbersLower|100|322C32"), _
Array("SpinnerShotL|100|968497","CrossbowRightStandupL|100|594F5A","StandupsNL|100|0E0C0E","StandupsIL|100|F4D7F5","DropTargetsPlusL|100|070607","DropTargetsASL|100|141214","DropTargetsCHL|100|CEB5CF","LeftHorseshoeCL|100|020202","LeftHorseshoeAL|100|D9BFDA","RightRampPowerupL|100|040304","GIScoop1|100|010101","Light003|100|B29DB3","GILeftOrbit1|100|B59FB5","GIRightRubbersLower|100|030203"), _
Array("SpinnerShotL|100|C7AFC8","CrossbowRightStandupL|100|665A66","StandupsNL|100|131013","StandupsIL|100|F6D9F7","DropTargetsPlusL|100|090809","DropTargetsASL|100|171417","DropTargetsCHL|100|D4BBD5","LeftHorseshoeCL|0|000000","LeftHorseshoeAL|100|CEB6CF","ScoopDRL|100|010101","RightRampPowerupL|100|040404","HoleShotL|100|BAA4BA","LeftRampShotL|100|F7DAF8","GIScoop1|100|060506","Light003|100|BCA6BD","GILeftOrbit1|100|877788","GIRightRubbersLower|0|000000"), _
Array("SpinnerShotL|100|DFC4E0","CrossbowRightStandupL|100|726573","StandupsNL|100|181518","StandupsIL|100|F7DAF8","DropTargetsPlusL|100|0B090B","DropTargetsASL|100|1A171A","DropTargetsCHL|100|DBC1DC","LeftHorseshoeAL|100|B29DB3","ScoopDRL|100|030203","HoleShotL|100|BAA4BB","ScoopShotL|100|020202","LeftRampShotL|100|E2C8E3","GIScoop1|100|161416","Light003|100|C6AFC7","GILeftOrbit1|100|605460"), _
Array("SpinnerShotL|100|F1D5F2","CrossbowRightStandupL|100|7E6F7E","StandupsNL|100|231F23","StandupsIL|100|F9DBFA","DropTargetsPlusL|100|0D0B0D","DropTargetsASL|100|1D191D","DropTargetsCHL|100|E1C7E2","LeftHorseshoeAL|100|958395","ScoopDRL|100|060506","ScoopPowerupL|100|010101","RightRampPowerupL|100|050405","ScoopShotL|100|191619","LeftRampShotL|100|C2ABC3","GIScoop1|100|262227","Light003|100|CFB6D0","GILeftOrbit1|100|413A42"), _
Array("SpinnerShotL|100|FCDEFD","ObjectiveFinalJudgmentL|100|F6D9F7","CrossbowRightStandupL|100|8A7A8A","StandupsNL|100|383138","StandupsIL|100|FADCFB","DropTargetsPlusL|100|0E0D0F","DropTargetsASL|100|1F1C20","DropTargetsCHL|100|E7CCE8","LeftHorseshoeAL|100|786978","ScoopDRL|100|0A090A","ScoopPowerupL|100|030303","RightRampPowerupL|100|050505","HoleShotL|100|BBA4BB","ScoopShotL|100|352E35","LeftRampShotL|100|938193","GIScoop1|100|4A414A","Light003|100|D8BED8","GIRightOrbit1|100|010001","GILeftOrbit1|100|242024"), _
Array("SpinnerShotL|100|FDDFFE","ObjectiveFinalJudgmentL|100|C0A9C1","CrossbowRightStandupL|100|958396","StandupsNL|100|514852","StandupsIL|100|FBDDFC","DropTargetsPlusL|100|110F11","DropTargetsASL|100|221E22","DropTargetsCHL|100|EDD1EE","LeftHorseshoeAL|100|5C515C","ScoopDRL|100|161316","ScoopPowerupL|100|070607","RightRampPowerupL|100|090809","LeftRampPowerupL|100|F7DAF8","HoleShotL|100|BBA5BC","ScoopShotL|100|544A54","LeftRampShotL|100|5B505B","GIScoop1|100|8A7A8B","Light003|100|E0C5E1","GIRightOrbit1|100|060606","GILeftOrbit1|100|060606"), _
Array("ObjectiveFinalJudgmentL|100|8A7A8A","CrossbowRightStandupL|100|A08DA1","StandupsNL|100|6A5D6A","StandupsIL|100|FCDEFD","DropTargetsHL|100|010101","DropTargetsPlusL|100|131013","DropTargetsASL|100|252125","DropTargetsCHL|100|F3D6F3","LeftHorseshoeAL|100|403840","LeftHorseshoeLL|100|FBDDFC","ScoopDRL|100|3C353C","LeftRampPL|100|CDB5CE","ScoopPowerupL|100|110F11","RightRampPowerupL|100|0D0B0D","LeftRampPowerupL|100|EDD1EE","ScoopShotL|100|7C6D7C","LeftRampShotL|100|2E292E","GIScoop1|100|CAB2CB","Light003|100|E7CCE8","GIRightOrbit1|100|121012","GILeftOrbit2|100|EBCFEC","GILeftOrbit1|0|000000"), _
Array("ObjectiveFinalJudgmentL|100|4F454F","ObjectiveSniperL|100|EED2EF","CrossbowRightStandupL|100|AB97AC","StandupsNL|100|807181","DropTargetsPlusL|100|151215","DropTargetsASL|100|282328","DropTargetsCHL|100|F3D6F4","LeftHorseshoeAL|100|262126","LeftHorseshoeLL|100|F6D8F7","ScoopDRL|100|655965","LeftRampPL|100|918091","ScoopPowerupL|100|2E292E","RightRampPowerupL|100|100E10","LeftRampPowerupL|100|D7BDD7","ScoopShotL|100|A491A5","LeftRampShotL|100|121012","GIScoop1|100|F6D9F7","Light003|100|EED2EF","GIRightOrbit1|100|342D34","GILeftOrbit2|100|D0B7D1"), _
Array("ObjectiveFinalJudgmentL|100|151315","ObjectiveSniperL|100|A28FA3","CrossbowRightStandupL|100|B6A0B7","StandupsNL|100|948395","StandupsIL|100|FDDFFE","DropTargetsHL|100|020202","DropTargetsPlusL|100|171417","DropTargetsASL|100|2B262B","LeftHorseshoeAL|100|171417","LeftHorseshoeLL|100|EFD3F0","ScoopDRL|100|8D7C8D","ScoopAGL|100|020202","LeftRampPL|100|564C56","LeftRampCL|100|FCDEFD","ScoopPowerupL|100|514751","RightRampPowerupL|100|131113","LeftRampPowerupL|100|9C899C","HoleShotL|100|BCA6BD","ScoopShotL|100|CAB2CB","LeftRampShotL|100|070607","GIScoop1|100|F8DBF9","Light003|100|F5D8F6","GIRightOrbit1|100|695C69","GILeftOrbit2|100|B6A0B7"), _
Array("ObjectiveFinalJudgmentL|100|050505","ObjectiveSniperL|100|5A505B","CrossbowRightStandupL|100|C0AAC1","StandupsNL|100|A894A8","DropTargetsPlusL|100|191619","DropTargetsASL|100|2E282E","DropTargetsCHL|100|F4D7F5","LeftHorseshoeAL|100|121012","LeftHorseshoeLL|100|E9CDEA","ScoopDRL|100|B49FB5","ScoopAGL|100|050405","LeftRampPL|100|302A30","LeftRampCL|100|F6D9F7","ScoopPowerupL|100|786A78","RightRampPowerupL|100|161417","LeftRampPowerupL|100|615661","ScoopShotL|100|ECD0ED","LeftRampShotL|100|010101","GIScoop1|100|FBDDFC","Light003|100|FCDEFD","GIRightOrbit3|100|010101","GIRightOrbit1|100|A592A6","GILeftOrbit2|100|9C8A9D"), _
Array("ObjectiveFinalJudgmentL|100|020202","ObjectiveSniperL|100|312B31","ObjectiveVikingL|100|CEB6CF","CrossbowRightStandupL|100|C5AEC6","StandupsNL|100|B9A3BA","DropTargetsHL|100|030203","DropTargetsPlusL|100|1B181B","DropTargetsASL|100|322C32","LeftHorseshoeAL|100|0D0C0D","LeftHorseshoeLL|100|E2C7E3","ScoopDRL|100|D2B9D3","ScoopAGL|100|090809","LeftRampPL|100|121012","LeftRampCL|100|E0C6E1","ScoopPowerupL|100|A28FA2","RightRampPowerupL|100|1A171A","LeftRampPowerupL|100|2D282D","ScoopShotL|100|F3D6F4","LeftRampShotL|0|000000","GIScoop1|100|FDDFFE","Light003|100|FDDFFE","GIRightOrbit3|100|040304","GIRightOrbit1|100|CFB7D0","GILeftOrbit2|100|766877"), _
Array("ObjectiveFinalJudgmentL|0|000000","ObjectiveSniperL|100|080708","ObjectiveBlacksmithL|100|F9DBFA","ObjectiveVikingL|100|918092","CrossbowRightStandupL|100|C9B1CA","StandupsNL|100|C5AEC6","DropTargetsHL|100|030303","DropTargetsPlusL|100|1D1A1D","DropTargetsASL|100|3B343B","LeftHorseshoeAL|100|0A080A","LeftHorseshoeLL|100|DAC0DB","ScoopDRL|100|E7CCE8","ScoopAGL|100|0F0E10","LeftRampPL|100|030203","LeftRampCL|100|A390A3","ScoopPowerupL|100|C0A9C0","RightRampPowerupL|100|1D191D","LeftRampPowerupL|100|121012","HoleShotL|100|BDA6BD","ScoopShotL|100|F8DBF9","GIRightOrbit3|100|090809","GIRightOrbit1|100|EACEEA","GILeftOrbit2|100|514751"), _
Array("ObjectiveSniperL|0|000000","ObjectiveBlacksmithL|100|EFD3F0","ObjectiveVikingL|100|5D525D","CrossbowRightStandupL|100|CDB5CE","StandupsNL|100|D0B7D1","DropTargetsHL|100|040304","DropTargetsPlusL|100|201C20","DropTargetsASL|100|443C45","DropTargetsCHL|100|F5D8F6","LeftHorseshoeAL|100|060606","LeftHorseshoeLL|100|CEB6CF","ScoopDRL|100|F6D9F7","ScoopAGL|100|2F2A2F","LeftRampPL|0|000000","LeftRampCL|100|665A67","LeftRampEL|100|F8DBF9","ScoopPowerupL|100|D8BED9","RightRampPowerupL|100|201C20","LeftRampPowerupL|100|030303","HoleShotL|100|BDA7BE","ScoopShotL|100|FCDEFD","GIScoop2|100|010101","GIRightOrbit3|100|171417","GIRightOrbit1|100|FADCFB","GILeftOrbit2|100|2C272C","ActionButton|100|171417"), _
Array("ObjectiveBlacksmithL|100|DAC0DB","ObjectiveVikingL|100|2F292F","CrossbowRightStandupL|100|D2B9D2","StandupsNL|100|D9BFDA","DropTargetsHL|100|040404","DropTargetsPlusL|100|231E23","DropTargetsASL|100|4D444D","LeftHorseshoeAL|100|040304","LeftHorseshoeLL|100|B39EB4","ScoopDRL|100|FBDDFC","ScoopAGL|100|564C56","LeftRampCL|100|312B31","LeftRampEL|100|F2D5F3","ScoopPowerupL|100|ECD0ED","RightRampPowerupL|100|231F23","LeftRampPowerupL|0|000000","ScoopShotL|100|FDDFFE","GIRightOrbit3|100|413A42","GIRightOrbit1|100|FDDFFE","GILeftOrbit2|100|090809","ActionButton|100|5E535F"), _
Array("ObjectiveBlacksmithL|100|A591A6","ObjectiveVikingL|100|0E0C0E","ObjectiveFinalBossL|100|EBCFEC","CrossbowRightStandupL|100|D6BCD7","StandupsNL|100|E1C7E2","DropTargetsHL|100|050405","DropTargetsPlusL|100|2C262C","DropTargetsASL|100|564C56","DropTargetsCHL|100|F6D8F6","LeftHorseshoeAL|100|010101","LeftHorseshoeLL|100|988698","ScoopONL|100|020202","ScoopDRL|100|FDDFFE","ScoopAGL|100|7B6C7B","LeftRampCL|100|171417","LeftRampEL|100|DAC0DB","ScoopPowerupL|100|FBDEFC","RightRampPowerupL|100|262126","GIScoop2|100|080708","GIRightOrbit3|100|796A79","GIRightOrbit2|100|110F11","GILeftOrbit2|100|070607","GIRightSlingshotUpper|100|E8CDE9","ActionButton|100|D1B8D1"), _
Array("ObjectiveBlacksmithL|100|6F626F","ObjectiveVikingL|100|070707","ObjectiveFinalBossL|100|B8A2B9","CrossbowRightStandupL|100|DAC0DB","StandupsNL|100|E8CDE9","DropTargetsHL|100|050505","DropTargetsPlusL|100|352F35","DropTargetsASL|100|5E535E","DropTargetsCHL|100|F6D9F7","LeftHorseshoeAL|0|000000","LeftHorseshoeLL|100|7E6F7F","ScoopLockL|100|050505","ScoopONL|100|070607","ScoopAGL|100|A08DA1","LeftRampCL|100|070607","LeftRampEL|100|9F8C9F","ScoopPowerupL|100|FDDFFE","RightRampPowerupL|100|292429","GIScoop2|100|373037","GIRightOrbit3|100|AF9AAF","GIRightOrbit2|100|2B262B","GILeftOrbit2|100|050405","GIRightSlingshotUpper|100|978598","ActionButton|100|F9DBFA"), _
Array("ObjectiveBlacksmithL|100|423A42","ObjectiveVikingL|0|000000","ObjectiveFinalBossL|100|857586","CrossbowRightStandupL|100|DEC4DF","CrossbowCenterStandupL|100|040304","StandupsNL|100|EFD2EF","DropTargetsHL|100|060506","DropTargetsPlusL|100|3E373E","DropTargetsASL|100|665A67","DropTargetsERL|100|010101","LeftHorseshoeLL|100|655965","ScoopLockL|100|151316","ScoopONL|100|171417","ScoopAGL|100|C4ADC5","LeftRampCL|0|000000","LeftRampEL|100|665A66","RightRampPowerupL|100|2C272C","HoleShotL|100|BEA7BF","GIScoop2|100|665A66","GIRightOrbit3|100|DBC1DB","GIRightOrbit2|100|514751","GILeftOrbit2|100|030203","GIRightSlingshotUpper|100|473F47","ActionButton|100|FDDFFE"), _
Array("ObjectiveBlacksmithL|100|1C191C","ObjectiveFinalBossL|100|524852","ObjectiveChaserL|100|D2BAD3","CrossbowRightStandupL|100|E2C7E3","CrossbowCenterStandupL|100|090809","StandupsNL|100|F3D7F4","DropTargetsPlusL|100|473F48","DropTargetsASL|100|6E616E","DropTargetsERL|100|020202","DropTargetsCHL|100|F7D9F8","LeftHorseshoeLL|100|4C434C","LeftHorseshoeBL|100|FBDDFC","ScoopLockL|100|2F292F","ScoopONL|100|2C272C","ScoopAGL|100|DFC5E0","LeftRampEL|100|332D33","RightRampPowerupL|100|2F292F","HoleShotL|100|BEA8BF","GIScoop2|100|938294","GIRightOrbit3|100|F7DAF8","GIRightOrbit2|100|807080","GILeftOrbit2|100|010101","GIRightSlingshotUpper|100|282429"), _
Array("ObjectiveBlacksmithL|0|000000","ObjectiveFinalBossL|100|1F1B1F","ObjectiveChaserL|100|918092","CrossbowRightStandupL|100|E7CBE8","CrossbowCenterStandupL|100|0E0C0E","StandupsNL|100|F7DAF8","DropTargetsHL|100|070607","DropTargetsPlusL|100|504751","DropTargetsASL|100|766877","DropTargetsERL|100|040304","DropTargetsCHL|100|F7DAF8","LeftHorseshoeLL|100|342E34","LeftHorseshoeBL|100|F7D9F8","ScoopLockL|100|4A424B","ScoopONL|100|473F47","ScoopAGL|100|F1D4F2","RightRampSL|100|010101","LeftRampEL|100|171517","RightRampPowerupL|100|312B31","GIScoop2|100|B7A2B8","GIRightOrbit3|100|FBDDFC","GIRightOrbit2|100|AC98AD","GILeftOrbit2|0|000000","GIRightSlingshotUpper|100|0B090B"), _
Array("ObjectiveFinalBossL|100|050505","ObjectiveChaserL|100|584D58","CrossbowRightStandupL|100|EBCFEC","CrossbowCenterStandupL|100|131013","StandupsNL|100|FADDFB","DropTargetsPlusL|100|5A4F5A","DropTargetsASL|100|7E6F7E","DropTargetsERL|100|050505","LeftHorseshoeLL|100|1F1B1F","LeftHorseshoeBL|100|F2D5F3","ScoopLockL|100|6A5D6A","ScoopONL|100|675A67","ScoopAGL|100|F9DBFA","RightRampSL|100|030303","LeftRampEL|100|060607","RightRampPowerupL|100|342E34","HoleShotL|100|BFA8BF","GIScoop2|100|CCB4CD","GIRightOrbit3|100|FDDFFE","GIRightOrbit2|100|D3BAD4","GILeftRubbersUpper|100|FCDEFD","GIRightSlingshotUpper|0|000000"), _
Array("ObjectiveFinalBossL|100|040304","ObjectiveChaserL|100|3E373F","CrossbowRightStandupL|100|EFD3F0","CrossbowCenterStandupL|100|181518","DropTargetsHL|100|100F11","DropTargetsPlusL|100|7E6F7E","DropTargetsASL|100|A08DA0","DropTargetsERL|100|0A090A","DropTargetsCHL|100|F9DBFA","LeftHorseshoeLL|100|151215","LeftHorseshoeBL|100|E8CDE9","ScoopLockL|100|766877","ScoopONL|100|766877","ScoopAGL|100|FADDFB","RightRampSL|0|000000","LeftRampEL|100|010101","RightRampPowerupL|100|322C32","HoleShotL|100|B29DB3","GIScoop2|100|DDC3DE","GIRightOrbit2|100|ECD0ED","GILeftRubbersUpper|100|FBDDFC"), _
Array("ObjectiveFinalBossL|100|010101","ObjectiveEscapeL|100|FADDFB","ObjectiveChaserL|100|171417","ObjectiveDragonL|100|FBDDFC","CrossbowRightStandupL|100|F3D6F4","CrossbowCenterStandupL|100|1D191D","StandupsNL|100|FCDEFD","DropTargetsHL|100|141215","DropTargetsPlusL|100|867686","DropTargetsASL|100|A693A7","DropTargetsERL|100|0C0A0C","DropTargetsCHL|100|F9DCFA","LeftHorseshoeLL|100|100E10","LeftHorseshoeBL|100|E2C7E3","ScoopLockL|100|958395","ScoopONL|100|99879A","ScoopAGL|100|FCDEFD","LeftRampEL|0|000000","RightRampPowerupL|100|342E34","GIScoop2|100|F1D4F2","GIRightOrbit2|100|FDDFFE","GILeftRubbersUpper|100|F7DAF8"), _
Array("ObjectiveFinalBossL|0|000000","ObjectiveEscapeL|100|F0D4F1","ObjectiveChaserL|0|000000","ObjectiveDragonL|100|EACEEB","CrossbowRightStandupL|100|F7DAF8","CrossbowCenterStandupL|100|211D21","DropTargetsHL|100|181618","DropTargetsPlusL|100|8E7D8E","DropTargetsASL|100|AE99AE","DropTargetsERL|100|0D0B0D","DropTargetsCHL|100|FADCFB","LeftHorseshoeLL|100|0C0A0C","LeftHorseshoeBL|100|DCC2DD","ScoopLockL|100|B49EB4","ScoopONL|100|BCA5BC","ScoopAGL|100|FDDFFE","RightRampPowerupL|100|373037","HoleShotL|100|B39DB3","GIScoop2|100|FDDFFE","GILeftRubbersUpper|100|F2D5F3"), _
Array("ObjectiveEscapeL|100|E4C9E5","ObjectiveDragonL|100|D8BED9","CrossbowRightStandupL|100|FBDDFC","CrossbowCenterStandupL|100|252126","StandupsNL|100|FDDFFE","DropTargetsHL|100|1C191D","DropTargetsPlusL|100|958396","DropTargetsASL|100|B49FB5","DropTargetsERL|100|0E0D0F","LeftHorseshoeLL|100|080708","LeftHorseshoeBL|100|CBB3CC","ScoopLockL|100|D2B9D3","ScoopONL|100|D4BBD5","RightRampPowerupL|100|393239","HoleShotL|100|B39EB4","GILeftRubbersUpper|100|E5CAE6"), _
Array("ObjectiveEscapeL|100|BFA8BF","ObjectiveDragonL|100|B19CB1","ObjectiveCastleL|100|E3C8E4","CrossbowRightStandupL|100|FDDFFE","CrossbowCenterStandupL|100|2A252A","DropTargetsHL|100|201C20","DropTargetsPlusL|100|9D8A9E","DropTargetsASL|100|BBA4BB","DropTargetsERL|100|100E10","LeftHorseshoeLL|100|050405","LeftHorseshoeBL|100|B49FB5","ScoopLockL|100|E7CCE8","ScoopONL|100|E7CCE8","RightRampSL|100|010101","RightRampPowerupL|100|3E373E","GILeftRubbersUpper|100|BEA7BE"), _
Array("ObjectiveEscapeL|100|827383","ObjectiveDragonL|100|7B6C7B","ObjectiveCastleL|100|B6A0B6","CrossbowCenterStandupL|100|2E282E","DropTargetsPL|100|010101","DropTargetsHL|100|242024","DropTargetsPlusL|100|A491A5","DropTargetsASL|100|C1AAC2","DropTargetsERL|100|110F11","DropTargetsCHL|100|FADDFB","RightHorseshoeHL|100|FCDEFD","LeftHorseshoeLL|100|020202","LeftHorseshoeBL|100|9D8A9E","ScoopLockL|100|F3D6F4","ScoopONL|100|F7D9F8","RightRampSL|100|020202","RightRampPowerupL|100|453D45","HoleShotL|100|B49EB4","GILeftRubbersUpper|100|918091"), _
Array("ObjectiveEscapeL|100|423A42","ObjectiveDragonL|100|433B44","ObjectiveCastleL|100|887888","CrossbowCenterStandupL|100|322C32","DropTargetsHL|100|282328","DropTargetsPlusL|100|AB97AC","DropTargetsASL|100|C7AFC8","DropTargetsERL|100|131113","DropTargetsCHL|100|FBDDFC","RightHorseshoeHL|100|FBDDFC","LeftHorseshoeLL|0|000000","LeftHorseshoeBL|100|857585","ScoopLockL|100|FBDEFC","ScoopONL|100|F9DBFA","RightRampSL|100|030303","RightRampPowerupL|100|4C434D","HoleShotL|100|B49FB5","GILeftRubbersUpper|100|615561"), _
Array("ObjectiveEscapeL|100|1E1A1E","ObjectiveDragonL|100|262227","ObjectiveCastleL|100|594E59","CrossbowCenterStandupL|100|352F36","DropTargetsPL|100|020202","DropTargetsHL|100|2B262C","DropTargetsPlusL|100|B19CB2","DropTargetsASL|100|CDB5CE","DropTargetsERL|100|151215","RightHorseshoeHL|100|F0D3F1","LeftHorseshoeBL|100|6C5F6D","ScoopLockL|100|FDDFFE","ScoopONL|100|FADDFB","RightRampSL|100|040404","RightRampPowerupL|100|534953","GILeftRubbersUpper|100|2F292F"), _
Array("ObjectiveEscapeL|100|100E10","ObjectiveDragonL|100|121012","ObjectiveCastleL|100|2A252A","CrossbowCenterStandupL|100|3F373F","DropTargetsPL|100|030203","DropTargetsHL|100|2F292F","DropTargetsPlusL|100|B8A2B9","DropTargetsASL|100|D3BAD4","DropTargetsERL|100|161316","RightHorseshoeHL|100|E4C9E5","LeftHorseshoeBL|100|554B56","ScoopONL|100|FBDEFC","RightRampSL|100|050505","RightRampPowerupL|100|5A4F5A","GILeftOrbit3|100|FADDFB","GILeftRubbersUpper|100|040404"), _
Array("ObjectiveEscapeL|100|050405","ObjectiveDragonL|100|010001","ObjectiveCastleL|100|070607","CrossbowCenterStandupL|100|484048","DropTargetsPL|100|030303","DropTargetsHL|100|332D33","DropTargetsPlusL|100|BFA8C0","DropTargetsASL|100|D9BFDA","DropTargetsERL|100|181518","DropTargetsCHL|100|FBDEFC","RightHorseshoeHL|100|D7BED8","LeftHorseshoeBL|100|3F3840","ScoopONL|100|FCDEFD","RightRampSL|100|060506","RightRampPowerupL|100|615561","HoleShotL|100|B59FB6","GILeftOrbit3|100|F0D3F0","GILeftRubbersUpper|100|020202","GIRightSlingshotLow|100|CEB5CF"), _
Array("ObjectiveEscapeL|100|010101","ObjectiveDragonL|0|000000","ObjectiveCastleL|100|040304","CrossbowCenterStandupL|100|524852","DropTargetsPL|100|040304","DropTargetsHL|100|363036","DropTargetsPlusL|100|C5AEC6","DropTargetsASL|100|DFC4DF","DropTargetsERL|100|191619","DropTargetsCHL|100|FCDEFD","RightHorseshoeHL|100|CAB2CB","LeftHorseshoeBL|100|2A252A","ScoopONL|100|FDDFFE","RightRampSL|100|070607","RightRampPowerupL|100|675B67","HoleShotL|100|B5A0B6","GILeftOrbit3|100|E2C7E3","GILeftRubbersUpper|100|010101","Light001|100|F4D7F5","GIRightSlingshotLow|100|8E7D8E"), _
Array("RightOutlanePoisonL|100|DBC1DC","ObjectiveCastleL|100|010101","CrossbowCenterStandupL|100|5B505B","DropTargetsPL|100|040404","DropTargetsHL|100|3A333A","DropTargetsPlusL|100|CBB3CC","DropTargetsASL|100|E4C9E5","DropTargetsERL|100|1B181B","RightHorseshoeHL|100|BCA6BD","LeftHorseshoeBL|100|201C20","RightRampSL|100|080708","RightRampPowerupL|100|6E616E","HoleShotL|100|B6A0B6","GILeftOrbit3|100|D1B8D2","Light001|100|DCC1DC","GIRightSlingshotLow|100|534954"), _
Array("RightOutlanePoisonL|100|A693A7","ObjectiveEscapeL|0|000000","ObjectiveCastleL|0|000000","CrossbowCenterStandupL|100|645864","DropTargetsPL|100|050405","DropTargetsHL|100|3D363D","DropTargetsPlusL|100|D1B9D2","DropTargetsASL|100|EACEEB","DropTargetsERL|100|1C191C","RightHorseshoeHL|100|AE9AAF","RightHorseshoeTL|100|FADCFB","LeftHorseshoeBL|100|1B181B","RightRampSL|100|080709","RightRampPowerupL|100|746674","RightRampShotL|100|010101","GILeftOrbit3|100|BCA6BD","GILeftRubbersUpper|0|000000","Light001|100|B9A3BA","GIRightSlingshotLow|100|342E34"), _
Array("RightOutlanePoisonL|100|706370","CrossbowCenterStandupL|100|6D606D","DropTargetsPL|100|050505","DropTargetsHL|100|403941","DropTargetsPlusL|100|D7BED8","DropTargetsASL|100|ECD0ED","DropTargetsERL|100|1E1A1E","RightHorseshoeHL|100|A18EA2","RightHorseshoeTL|100|F7D9F7","LeftHorseshoeBL|100|161316","RightRampSL|100|090809","RightRampPowerupL|100|7A6B7A","HoleShotL|100|B6A0B7","LeftHorseshoeShotL|100|FBDEFC","GILeftOrbit3|100|A693A7","Light001|100|8D7C8E","GIRightSlingshotLow|100|151215"), _
Array("RightOutlanePoisonL|100|494049","CrossbowCenterStandupL|100|766876","DropTargetsPL|100|060506","DropTargetsHL|100|443C44","DropTargetsPlusL|100|DDC3DE","DropTargetsERL|100|211D21","DropTargetsCHL|100|FDDFFE","RightHorseshoeHL|100|948295","RightHorseshoeTL|100|F3D6F4","LeftHorseshoeBL|100|110F11","RightRampSL|100|0A090A","RightRampPowerupL|100|807180","HoleShotL|100|B6A1B7","LeftHorseshoeShotL|100|F8DBF9","GILeftOrbit3|100|897889","Light001|100|5E535F","GIRightSlingshotLow|100|020202"), _
Array("RightOutlanePoisonL|100|241F24","CrossbowCenterStandupL|100|7E6F7E","DropTargetsHL|100|473F47","DropTargetsPlusL|100|E3C8E4","DropTargetsASL|100|EDD1EE","DropTargetsERL|100|292429","RightHorseshoeHL|100|877787","RightHorseshoeTL|100|EFD3F0","LeftHorseshoeBL|100|0D0C0D","RightRampPowerupL|100|867687","RightRampShotL|100|020102","LeftHorseshoeShotL|100|F5D8F6","GILeftOrbit3|100|6C5F6C","Light001|100|352F35","GIRightSlingshotLow|100|010101"), _
Array("RightOutlanePoisonL|100|070607","CrossbowCenterStandupL|100|867687","DropTargetsPL|100|070607","DropTargetsHL|100|4B424B","DropTargetsPlusL|100|E8CDE9","DropTargetsASL|100|EED1EE","DropTargetsERL|100|302A30","RightHorseshoeHL|100|7A6B7A","RightHorseshoeTL|100|EACFEB","LeftHorseshoeBL|100|0A090A","RightRampSL|100|0B090B","RightRampPowerupL|100|8B7B8C","HoleShotL|100|B7A1B8","RightRampShotL|100|020202","LeftHorseshoeShotL|100|F1D4F2","GILeftOrbit3|100|524852","Light001|100|1D191D","GILeftSlingshot|100|F9DBFA","GIRightSlingshotLow|0|000000"), _
Array("RightOutlanePoisonL|100|040304","CrossbowCenterStandupL|100|8E7D8F","StandupsTL|100|FCDEFD","DropTargetsHL|100|4E444E","DropTargetsPlusL|100|ECD0ED","DropTargetsASL|100|EED2EF","DropTargetsERL|100|373137","RightHorseshoeHL|100|6E616E","RightHorseshoeTL|100|E6CBE7","LeftHorseshoeBL|100|070607","RightRampSL|100|0B0A0B","RightRampPowerupL|100|928092","LeftHorseshoeShotL|100|EDD1EE","GIUpperFlipper|100|020102","GILeftOrbit3|100|3A333A","Light001|100|0E0C0E","GILeftSlingshot|100|E8CDE9"), _
Array("RightOutlanePoisonL|100|010101","CrossbowCenterStandupL|100|968497","DropTargetsPL|100|070708","DropTargetsHL|100|514751","DropTargetsPlusL|100|EDD1EE","DropTargetsASL|100|EFD2F0","DropTargetsERL|100|3F373F","RightHorseshoeHL|100|625763","RightHorseshoeTL|100|E0C5E1","LeftHorseshoeBL|100|040304","RightRampPowerupL|100|978598","RightRampShotL|100|030203","LeftHorseshoeShotL|100|E8CDE9","CrossbowShotL|100|010101","GIUpperFlipper|100|040404","GILeftOrbit3|100|252025","Light001|100|070607","GIRightInlaneUpper|100|CCB4CD","GILeftSlingshot|100|D3BAD4"), _
Array("RightOutlanePoisonL|0|000000","CrossbowCenterStandupL|100|9E8B9F","StandupsTL|100|FBDDFC","DropTargetsPL|100|080708","DropTargetsHL|100|544A55","DropTargetsPlusL|100|EED1EE","DropTargetsASL|100|EFD3F0","DropTargetsERL|100|463E46","RightHorseshoeHL|100|584D58","RightHorseshoeTL|100|D9BFDA","LeftHorseshoeBL|100|010101","RightRampPowerupL|100|9C8A9D","HoleShotL|100|B8A2B8","RightRampShotL|100|030303","LeftHorseshoeShotL|100|E3C8E4","CrossbowShotL|100|020202","GIUpperFlipper|100|0D0B0D","Light002|100|010101","GILeftOrbit3|100|121012","Light001|100|040304","GIRightInlaneUpper|100|8B7B8C","GILeftSlingshot|100|BAA3BA"), _
Array("CrossbowCenterStandupL|100|A692A7","DropTargetsHL|100|574D58","DropTargetsPlusL|100|EED2EF","DropTargetsASL|100|F0D3F1","DropTargetsERL|100|4D434D","RightHorseshoeHL|100|4D444D","RightHorseshoeTL|100|D1B8D2","LeftHorseshoeBL|0|000000","RightRampSL|100|0C0A0C","RightRampPowerupL|100|A28FA2","HoleShotL|100|B8A2B9","LeftHorseshoeShotL|100|DEC4DF","CrossbowShotL|100|040304","GIUpperFlipper|100|1B181B","Light002|100|020202","GILeftOrbit3|100|0B0A0B","Light001|100|020102","GIRightInlaneUpper|100|4A424B","GILeftSlingshot|100|8B7B8C"), _
Array("CrossbowCenterStandupL|100|AD99AE","StandupsTL|100|FADCFB","DropTargetsPL|100|090809","DropTargetsHL|100|5C515C","DropTargetsPlusL|100|EFD2F0","DropTargetsASL|100|F0D4F1","DropTargetsERL|100|544A54","RightHorseshoeHL|100|423B43","RightHorseshoeTL|100|C6AFC7","RightRampPowerupL|100|A793A8","RightRampShotL|100|040304","LeftHorseshoeShotL|100|D2B9D2","CrossbowShotL|100|050405","GIUpperFlipper|100|2D282D","GILeftOrbit3|100|080708","Light001|0|000000","GIRightInlaneUpper|100|2F2A30","GILeftSlingshot|100|5D525D"), _
Array("CrossbowCenterStandupL|100|B49FB5","CrossbowLeftStandupL|100|030203","DropTargetsHL|100|615661","DropTargetsPlusL|100|EFD3F0","DropTargetsASL|100|F1D4F2","DropTargetsERL|100|5B505B","RightHorseshoeHL|100|393239","RightHorseshoeTL|100|B49EB4","RightRampPowerupL|100|AD98AD","LeftHorseshoeShotL|100|BFA8C0","CrossbowShotL|100|070607","GIUpperFlipper|100|423A42","Light002|100|030303","GILeftOrbit3|100|050405","GIRightInlaneUpper|100|171518","GIRightInlaneLower|100|F0D4F1","GILeftSlingshot|100|332D34"), _
Array("CrossbowCenterStandupL|100|B8A2B9","CrossbowLeftStandupL|100|060506","StandupsTL|100|F9DCFA","DropTargetsPL|100|0A080A","DropTargetsHL|100|675B67","DropTargetsPlusL|100|F0D4F1","DropTargetsASL|100|F1D5F2","DropTargetsERL|100|625662","RightHorseshoeHL|100|2F292F","RightHorseshoeTL|100|A18EA2","RightHorseshoeIL|100|FCDEFD","RightRampPowerupL|100|B29DB3","HoleShotL|100|B9A3BA","RightRampShotL|100|050405","LeftHorseshoeShotL|100|AC97AC","CrossbowShotL|100|090809","GIUpperFlipper|100|5E535F","Light002|100|040404","GIBumperInlanes3|100|FCDEFD","GILeftOrbit3|100|030303","GIRightInlaneUpper|0|000000","GIRightInlaneLower|100|BBA5BC","GILeftSlingshot|100|181518"), _
Array("CrossbowCenterStandupL|100|BCA6BD","CrossbowLeftStandupL|100|090809","StandupsTL|100|F9DBFA","DropTargetsPL|100|0D0C0D","DropTargetsHL|100|6C5F6D","DropTargetsASL|100|F2D5F3","DropTargetsERL|100|685C69","RightHorseshoeHL|100|262126","RightHorseshoeTL|100|907F90","RightHorseshoeIL|100|FADCFB","RightRampSL|100|0C0B0C","RightRampPowerupL|100|B7A2B8","RightRampShotL|100|060506","LeftHorseshoeShotL|100|988698","CrossbowShotL|100|0B0A0B","GIUpperFlipper|100|7B6C7B","GIRightRamp|100|FBDEFC","Light002|100|050405","GIBumperInlanes4|100|FCDEFD","GIBumperInlanes3|100|F6D9F7","GILeftOrbit3|100|010101","GIRightInlaneLower|100|877787","GILeftSlingshot|100|0F0D0F"), _
Array("CrossbowCenterStandupL|100|C0A9C1","CrossbowLeftStandupL|100|0C0B0C","StandupsTL|100|F8DBF9","DropTargetsPL|100|110F11","DropTargetsHL|100|726472","DropTargetsPlusL|100|F1D4F2","DropTargetsERL|100|6F6270","RightHorseshoeHL|100|1D191D","RightHorseshoeTL|100|7F7080","RightHorseshoeIL|100|F8DAF9","RightRampSL|100|0C0A0C","RightRampPowerupL|100|BCA6BD","RightRampShotL|100|070607","LeftHorseshoeShotL|100|837484","CrossbowShotL|100|0E0C0E","GIUpperFlipper|100|978597","GIRightRamp|100|F9DCFA","GICrossbowLeft|100|020202","Light002|100|060506","GIBumperInlanes4|100|FBDDFC","GIBumperInlanes3|100|F0D3F1","GILeftOrbit3|0|000000","GIRightInlaneLower|100|5B505B","GILeftSlingshot|100|070707"), _
Array("CrossbowCenterStandupL|100|C4ACC4","CrossbowLeftStandupL|100|0F0E0F","DropTargetsPL|100|141215","DropTargetsHL|100|776977","DropTargetsPlusL|100|F2D5F3","DropTargetsASL|100|F3D6F4","DropTargetsERL|100|766876","RightHorseshoeHL|100|141214","RightHorseshoeTL|100|6F6270","RightHorseshoeIL|100|F5D8F6","RightRampPowerupL|100|C1AAC2","RightRampShotL|100|080708","LeftHorseshoeShotL|100|6E616F","CrossbowShotL|100|121012","GIUpperFlipper|100|B39EB4","GIRightRamp|100|F7DAF8","GICrossbowLeft|100|070607","Light002|100|070607","GIBumperInlanes4|100|FADDFB","GIBumperInlanes3|100|EACEEB","GIRightInlaneLower|100|3C353C","GILeftSlingshot|100|020202"), _
Array("CrossbowCenterStandupL|100|C7B0C8","CrossbowLeftStandupL|100|121012","StandupsTL|100|F8DAF9","DropTargetsPL|100|181518","DropTargetsHL|100|7C6E7D","DropTargetsERL|100|7C6E7D","RightHorseshoeHL|100|0D0B0D","RightHorseshoeTL|100|615561","RightHorseshoeIL|100|F3D6F4","RightRampPowerupL|100|C6AFC7","RightRampShotL|100|090809","LeftHorseshoeShotL|100|594E59","CrossbowShotL|100|171517","GIUpperFlipper|100|CEB6CF","GIRightRamp|100|F5D8F5","GICrossbowLeft|100|0B0A0B","Light002|100|080708","GIBumperInlanes4|100|F9DCFA","GIBumperInlanes3|100|E4C9E5","GIRightInlaneLower|100|1D191D","GILeftSlingshotLower|100|FBDDFC","GILeftSlingshot|0|000000"), _
Array("CrossbowCenterStandupL|100|CBB3CC","CrossbowLeftStandupL|100|151315","StandupsTL|100|F7DAF8","DropTargetsPL|100|1C181C","DropTargetsHL|100|817282","DropTargetsPlusL|100|F3D6F4","DropTargetsASL|100|F4D7F5","DropTargetsERL|100|837383","RightHorseshoeHL|100|0B0A0B","RightHorseshoeTL|100|524953","RightHorseshoeIL|100|E9CEEA","RightRampPowerupL|100|CBB3CC","HoleShotL|100|BAA4BB","RightRampShotL|100|0A090B","LeftHorseshoeShotL|100|453D45","CrossbowShotL|100|1D1A1D","GIUpperFlipper|100|E9CEEA","GIRightRamp|100|F1D5F2","GICrossbowLeft|100|100E10","Light002|100|090809","GIBumperInlanes4|100|F9DBFA","GIBumperInlanes3|100|DEC4DF","GIRightInlaneLower|100|070607","GILeftSlingshotLower|100|F5D8F6"), _
Array("CrossbowCenterStandupL|100|CFB7D0","CrossbowLeftStandupL|100|181518","DropTargetsPL|100|1F1B1F","DropTargetsHL|100|867687","DropTargetsERL|100|887889","RightHorseshoeHL|100|0A090A","RightHorseshoeTL|100|453D45","RightHorseshoeIL|100|D9BFD9","RightRampAL|100|020202","RightRampSL|100|0B0A0B","RightRampPowerupL|100|D0B7D1","RightRampShotL|100|0C0A0C","LeftHorseshoeShotL|100|322C33","CrossbowShotL|100|231F23","GIUpperFlipper|100|F5D8F6","GIRightRamp|100|EED2EF","GICrossbowLeft|100|141214","Light002|100|0A090A","GIBumperInlanes4|100|F8DAF9","GIBumperInlanes3|100|D8BFD9","GIRightInlaneLower|100|040404","GILeftSlingshotLower|100|EBCFEC"), _
Array("CrossbowCenterStandupL|100|D3BAD4","CrossbowLeftStandupL|100|1A171B","StandupsTL|100|F7D9F8","DropTargetsPL|100|221E22","DropTargetsHL|100|8B7A8B","DropTargetsPlusL|100|F4D7F4","DropTargetsERL|100|8E7E8F","RightHorseshoeHL|100|090809","RightHorseshoeTL|100|373138","RightHorseshoeIL|100|C8B0C8","RightOrbitVL|100|FCDFFD","RightRampAL|100|040404","RightRampPowerupL|100|D5BBD5","RightRampShotL|100|0D0C0D","LeftHorseshoeShotL|100|252025","CrossbowShotL|100|2A252A","GIUpperFlipper|100|F7DAF8","GIRightRamp|100|EACFEB","GICrossbowLeft|100|191619","Light002|100|0B0A0B","GIBumperInlanes4|100|F7DAF8","GIBumperInlanes3|100|D2B9D3","GIRightInlaneLower|100|020202","GILeftSlingshotLower|100|C6AEC6"), _
Array("CrossbowCenterStandupL|100|D7BDD8","CrossbowLeftStandupL|100|1D1A1D","StandupsTL|100|F6D9F7","DropTargetsPL|100|262126","DropTargetsHL|100|907F90","DropTargetsPlusL|100|F4D7F5","DropTargetsASL|100|F5D8F6","DropTargetsERL|100|948295","RightHorseshoeHL|100|080708","RightHorseshoeTL|100|2B262B","RightHorseshoeIL|100|B8A2B8","RightOrbitVL|100|FCDEFD","RightRampAL|100|060606","RightRampPowerupL|100|D9BFDA","RightRampShotL|100|0E0D0F","LeftHorseshoeShotL|100|201C20","CrossbowShotL|100|312B31","GIUpperFlipper|100|F9DBFA","GIRightRamp|100|E6CBE7","GICrossbowLeft|100|1D1A1D","Light002|100|0C0B0C","GIBumperInlanes4|100|EFD3F0","GIBumperInlanes3|100|C7AFC8","GIRightInlaneLower|0|000000","GILeftSlingshotLower|100|958396"), _
Array("CrossbowCenterStandupL|100|DAC0DB","CrossbowLeftStandupL|100|201C20","DropTargetsPL|100|292429","DropTargetsHL|100|958395","DropTargetsERL|100|9A889B","RightHorseshoeHL|100|070607","RightHorseshoeTL|100|1F1C20","RightHorseshoeIL|100|A693A7","RightRampAL|100|080708","RightRampSL|100|0B090B","RightRampPowerupL|100|DDC3DE","HoleShotL|100|BBA4BB","RightRampShotL|100|100E10","LeftHorseshoeShotL|100|1B181B","CrossbowShotL|100|383138","GIUpperFlipper|100|FADCFB","GIRightRamp|100|E2C7E3","GICrossbowLeft|100|211D21","Light002|100|0D0C0D","GIBumperInlanes4|100|E1C6E1","GIBumperInlanes3|100|B7A1B7","GIBumperInlanes2|100|F8DAF9","GILeftSlingshotLower|100|665A67"), _
Array("CrossbowCenterStandupL|100|DEC4DF","CrossbowLeftStandupL|100|221E22","DropTargetsPL|100|2C272C","DropTargetsHL|100|99879A","DropTargetsPlusL|100|F5D8F6","DropTargetsASL|100|F6D9F7","DropTargetsERL|100|A08DA0","RightHorseshoeHL|100|060506","RightHorseshoeTL|100|161416","RightHorseshoeIL|100|978597","RightHorseshoeML|100|FCDEFD","RightRampAL|100|0A090A","RightRampSL|100|0A090A","RightRampPowerupL|100|E2C7E3","RightRampShotL|100|110F11","LeftHorseshoeShotL|100|171417","CrossbowShotL|100|403840","GIUpperFlipper|100|FBDDFC","GIRightRamp|100|DAC0DB","GICrossbowLeft|100|252125","Light002|100|0E0D0E","GIBumperInlanes4|100|D3BAD3","GIBumperInlanes3|100|A793A7","GIBumperInlanes2|100|EED2EF","GILeftSlingshotLower|100|393239"), _
Array("CrossbowCenterStandupL|100|E2C7E3","CrossbowLeftStandupL|100|242024","StandupsTL|100|F5D8F6","DropTargetsPL|100|302A30","DropTargetsHL|100|9E8B9E","DropTargetsERL|100|A592A6","RightHorseshoeHL|100|050405","RightHorseshoeTL|100|141214","RightHorseshoeIL|100|867687","RightHorseshoeML|100|F2D5F3","RightOrbitBumpersL|100|FCDFFD","RightRampAL|100|0C0B0C","RightRampPowerupL|100|E7CBE8","HoleShotL|100|BBA5BC","RightRampShotL|100|131113","LeftHorseshoeShotL|100|131113","CrossbowShotL|100|483F48","GIUpperFlipper|100|FCDEFD","GIRightRamp|100|D0B7D0","GICrossbowLeft|100|292429","Light002|100|100E10","GIBumperInlanes4|100|C4ADC5","GIBumperInlanes3|100|978597","GIBumperInlanes2|100|E4C9E5","GILeftSlingshotLower|100|272227"), _
Array("CrossbowCenterStandupL|100|E6CAE6","CrossbowLeftStandupL|100|262226","DropTargetsPL|100|322D33","DropTargetsHL|100|A28FA3","DropTargetsPlusL|100|F6D8F7","DropTargetsERL|100|AB96AB","RightHorseshoeHL|100|040404","RightHorseshoeTL|100|121012","RightHorseshoeIL|100|766877","RightHorseshoeML|100|E7CCE8","RightOrbitBumpersL|100|FCDEFD","RightOrbitVL|100|FBDEFC","RightRampAL|100|0F0D0F","RightRampPowerupL|100|E7CCE8","RightRampShotL|100|171518","LeftHorseshoeShotL|100|0F0D0F","CrossbowShotL|100|524853","GIUpperFlipper|100|FDDFFE","GIRightRamp|100|C0AAC1","GICrossbowLeft|100|2C272C","Light002|100|110F11","GIBumperInlanes4|100|B6A1B7","GIBumperInlanes3|100|877788","GIBumperInlanes2|100|DBC1DB","GILeftSlingshotLower|100|181518"), _
Array("CrossbowCenterStandupL|100|E9CEEA","CrossbowLeftStandupL|100|282328","DropTargetsPL|100|362F36","DropTargetsHL|100|A793A7","DropTargetsPlusL|100|F6D9F7","DropTargetsASL|100|F7D9F8","DropTargetsERL|100|B09BB1","RightHorseshoeHL|100|030303","RightHorseshoeTL|100|100E10","RightHorseshoeIL|100|665A67","RightHorseshoeML|100|DDC2DD","RightOrbitVL|100|FBDDFC","RightRampAL|100|100E10","RightRampPowerupL|100|E8CCE9","RightRampShotL|100|1E1A1E","LeftHorseshoeShotL|100|0C0A0C","CrossbowShotL|100|5E535F","GIRightRamp|100|B09BB0","GICrossbowLeft|100|302A30","Light002|100|121012","GIBumperInlanes4|100|A995AA","GIBumperInlanes3|100|786A79","GIBumperInlanes2|100|D0B8D1","GILeftSlingshotLower|100|0B0A0B"), _
Array("CrossbowCenterStandupL|100|EDD1EE","CrossbowLeftStandupL|100|302A30","DropTargetsPL|100|393239","DropTargetsHL|100|AB97AC","DropTargetsASL|100|F7DAF8","DropTargetsERL|100|B5A0B6","RightHorseshoeHL|100|030203","RightHorseshoeTL|100|0E0C0E","RightHorseshoeIL|100|574D57","RightHorseshoeML|100|D0B7D1","RightRampAL|100|121012","RightRampShotL|100|241F24","LeftHorseshoeShotL|100|090809","CrossbowShotL|100|6A5E6B","GIRightRamp|100|A08DA0","GICrossbowLeft|100|332D34","Light002|100|131113","GIBumperInlanes4|100|9B899C","GIBumperInlanes3|100|695D6A","GIBumperInlanes2|100|C6AEC7","GILeftSlingshotLower|100|020202"), _
Array("CrossbowCenterStandupL|100|F1D4F2","CrossbowLeftStandupL|100|373138","StandupsTL|100|F4D7F5","DropTargetsPL|100|3C353C","DropTargetsHL|100|AF9AB0","DropTargetsPlusL|100|F7D9F8","DropTargetsERL|100|BBA5BC","RightHorseshoeHL|100|020202","RightHorseshoeTL|100|0C0A0C","RightHorseshoeIL|100|473E47","RightHorseshoeML|100|C3ACC4","RightOrbitVL|100|F9DCFA","RightRampAL|100|151215","RightRampPowerupL|100|E8CDE9","CrossbowPowerupL|100|010101","HoleShotL|100|BCA5BC","RightRampShotL|100|29252A","LeftHorseshoeShotL|100|060506","CrossbowShotL|100|766877","GIRightRamp|100|918092","GICrossbowLeft|100|373037","Light002|100|151215","GIBumperInlanes4|100|8E7D8F","GIBumperInlanes3|100|5B505B","GIBumperInlanes2|100|BCA5BC","GILeftSlingshotLower|100|010101","ActionRightL|100|F9DCFA"), _
Array("CrossbowCenterStandupL|100|F4D7F5","CrossbowLeftStandupL|100|3F373F","DropTargetsPL|100|3F373F","DropTargetsHL|100|B39EB4","DropTargetsPlusL|100|F7DAF8","DropTargetsASL|100|F8DAF9","DropTargetsERL|100|BFA9C0","RightHorseshoeHL|100|010101","RightHorseshoeTL|100|0A090A","RightHorseshoeIL|100|373137","RightHorseshoeML|100|B5A0B6","RightHorseshoeSL|100|FADDFB","RightOrbitVL|100|E9CEEA","RightRampAL|100|161416","RightRampSL|100|0B090B","CrossbowPowerupL|100|020202","RightRampShotL|100|2F292F","LeftHorseshoeShotL|100|030303","CrossbowShotL|100|827383","GIRightRamp|100|827383","GICrossbowLeft|100|3A333A","Light002|100|161316","GIBumperInlanes4|100|817282","GIBumperInlanes3|100|4C434C","GIBumperInlanes2|100|B19CB2","GILeftInlaneLower|100|FCDEFD","GILeftSlingshotLower|100|010001","ActionRightL|100|EFD2F0"), _
Array("CrossbowCenterStandupL|100|F8DBF9","CrossbowLeftStandupL|100|463E46","DropTargetsPL|100|423A42","DropTargetsHL|100|B7A2B8","DropTargetsASL|100|F8DBF9","DropTargetsERL|100|C5ADC5","RightHorseshoeHL|0|000000","RightHorseshoeTL|100|080708","RightHorseshoeIL|100|282328","RightHorseshoeML|100|A793A8","RightHorseshoeSL|100|F6D9F7","RightOrbitBumpersL|100|FBDEFC","RightOrbitVL|100|D9BFDA","RightRampAL|100|181518","RightRampPowerupL|100|E9CDEA","CrossbowPowerupL|100|030203","HoleShotL|100|BCA6BD","RightRampShotL|100|352E35","LeftHorseshoeShotL|100|010101","CrossbowShotL|100|8E7D8E","GIRightRamp|100|746675","GICrossbowLeft|100|3F373F","Light002|100|171417","GIBumperInlanes4|100|746674","GIBumperInlanes3|100|433B43","GIBumperInlanes2|100|A693A7","GIBumperInlanes1|100|F9DBFA","GILeftInlaneLower|100|FADCFB","GILeftSlingshotLower|0|000000","ActionRightL|100|E5CAE6"), _
Array("CrossbowCenterStandupL|100|FCDEFD","CrossbowLeftStandupL|100|4D444D","DropTargetsPL|100|443C45","DropTargetsHL|100|BCA5BC","DropTargetsPlusL|100|F8DAF9","DropTargetsERL|100|C9B1CA","RightHorseshoeTL|100|070607","RightHorseshoeIL|100|221E22","RightHorseshoeML|100|998699","RightHorseshoeSL|100|F2D5F3","RightOrbitBumpersL|100|FBDDFC","RightOrbitVL|100|C9B1CA","RightOrbitKL|100|FCDEFD","LeftOrbitGL|100|010101","RightRampEL|100|010101","RightRampAL|100|1A171A","CrossbowPowerupL|100|040404","RightRampShotL|100|3A343B","LeftHorseshoeShotL|0|000000","CrossbowShotL|100|988699","GIRightRamp|100|665A67","GICrossbowLeft|100|463E47","Light002|100|181518","GIBumperInlanes4|100|675B67","GIBumperInlanes3|100|3C343C","GIBumperInlanes2|100|958396","GIBumperInlanes1|100|F4D7F5","GILeftInlaneLower|100|F8DBF9","GILeftInlaneUpper|100|FCDEFD","ActionRightL|100|DBC1DC"), _
Array("CrossbowCenterStandupL|100|FDDFFE","CrossbowLeftStandupL|100|544A55","DropTargetsPL|100|473F48","DropTargetsHL|100|C0A9C0","DropTargetsPlusL|100|F8DBF9","DropTargetsASL|100|F9DBFA","DropTargetsERL|100|CFB6CF","RightHorseshoeTL|100|060506","RightHorseshoeIL|100|1E1B1E","RightHorseshoeML|100|89798A","RightHorseshoeSL|100|EDD1EE","RightOrbitBumpersL|100|F8DBF9","RightOrbitVL|100|BAA4BA","LeftOrbitGL|100|050405","RightRampEL|100|020202","RightRampAL|100|1B181B","RightRampSL|100|0B0A0B","RightRampPowerupL|100|E9CEEA","CrossbowPowerupL|100|050505","HoleShotL|100|BDA6BD","RightRampShotL|100|403840","CrossbowShotL|100|A28FA3","GIRightRamp|100|594F59","GICrossbowLeft|100|4E454E","Light002|100|1A171A","GIBumperInlanes4|100|594F5A","GIBumperInlanes3|100|342E34","GIBumperInlanes2|100|837383","GIBumperInlanes1|100|EED1EF","GILeftInlaneLower|100|EFD2EF","GILeftInlaneUpper|100|FADCFB","ActionRightL|100|AC98AD"), _
Array("CrossbowLeftStandupL|100|5B515C","StandupsRL|100|FCDEFD","StandupsTL|100|F3D7F4","DropTargetsPL|100|4A414A","DropTargetsHL|100|C4ADC5","DropTargetsERL|100|D3BAD4","RightHorseshoeTL|100|050405","RightHorseshoeIL|100|1B181B","RightHorseshoeML|100|7A6B7A","RightHorseshoeSL|100|E8CCE9","RightOrbitBumpersL|100|F1D4F2","RightOrbitVL|100|AA96AB","LeftOrbitGL|100|080708","RightRampEL|100|040304","RightRampAL|100|1D191D","CrossbowPowerupL|100|070607","RightRampShotL|100|463D46","CrossbowShotL|100|AC97AC","GIRightRamp|100|4D444D","GICrossbowLeft|100|554B55","Light002|100|1B181B","GIBumperInlanes4|100|4C434C","GIBumperInlanes3|100|2C272C","GIBumperInlanes2|100|716371","GIBumperInlanes1|100|E7CCE8","GILeftInlaneLower|100|DDC3DE","GILeftInlaneUpper|100|F7DAF8","ActionRightL|100|7E6F7E"), _
Array("LeftOutlanePoisonL|100|FADCFB","CrossbowLeftStandupL|100|625763","StandupsTL|100|F3D6F4","DropTargetsPL|100|4D444D","DropTargetsHL|100|C7B0C8","DropTargetsPlusL|100|F9DBFA","DropTargetsASL|100|F9DCFA","DropTargetsERL|100|D8BED9","RightHorseshoeTL|100|030303","RightHorseshoeIL|100|181518","RightHorseshoeML|100|6B5E6B","RightHorseshoeSL|100|E1C7E2","RightOrbitBumpersL|100|E9CEEA","RightOrbitVL|100|9B899C","LeftOrbitGL|100|0C0B0C","RightRampEL|100|060506","RightRampAL|100|1E1A1E","RightRampPowerupL|100|EACEEB","CrossbowPowerupL|100|090809","HoleShotL|100|BDA6BE","RightRampShotL|100|4B424B","CrossbowShotL|100|B59FB5","GIRightRamp|100|403840","GICrossbowLeft|100|5C515D","Light002|100|1C191C","GIBumperInlanes4|100|3F383F","GIBumperInlanes3|100|252025","GIBumperInlanes2|100|605460","GIBumperInlanes1|100|E0C6E1","GILeftInlaneLower|100|CAB2CA","GILeftInlaneUpper|100|F4D7F5","ActionRightL|100|4F464F"), _
Array("LeftOutlanePoisonL|100|F1D4F2","CrossbowLeftStandupL|100|695D6A","DropTargetsPL|100|4F4650","DropTargetsHL|100|CCB3CC","DropTargetsASL|100|FADCFB","DropTargetsERL|100|DCC2DD","RightHorseshoeTL|100|020202","RightHorseshoeIL|100|161316","RightHorseshoeML|100|5D525D","RightHorseshoeSL|100|DAC0DB","RightOrbitBumpersL|100|E2C7E3","RightOrbitVL|100|8C7B8C","LeftOrbitGL|100|100E10","RightRampEL|100|070607","RightRampAL|100|1F1B1F","CrossbowPowerupL|100|0B090B","RightRampShotL|100|504750","CrossbowShotL|100|BDA6BD","GIRightRamp|100|342E35","GICrossbowLeft|100|635864","Light002|100|201C20","GIBumperInlanes4|100|332D33","GIBumperInlanes3|100|1D1A1D","GIBumperInlanes2|100|4F464F","GIBumperInlanes1|100|D9BFDA","GIRightRubbersUpper|100|010101","GILeftInlaneLower|100|B49FB5","GILeftInlaneUpper|100|DBC1DC","ActionRightL|100|332D33"), _
Array("LeftOutlanePoisonL|100|E3C8E3","CrossbowLeftStandupL|100|706270","DropTargetsPL|100|524852","DropTargetsHL|100|CFB7D0","DropTargetsPlusL|100|F9DCFA","DropTargetsERL|100|E1C7E2","RightHorseshoeTL|100|010101","RightHorseshoeIL|100|131113","RightHorseshoeML|100|4F454F","RightHorseshoeSL|100|D1B8D2","RightOrbitBumpersL|100|DBC1DC","RightOrbitVL|100|7D6E7E","RightOrbitKL|100|FBDEFC","LeftOrbitGL|100|141114","RightRampEL|100|090809","RightRampAL|100|201C20","CrossbowPowerupL|100|0D0B0D","RightRampShotL|100|554B56","CrossbowShotL|100|C4ADC5","GIRightRamp|100|282428","GICrossbowLeft|100|6B5E6B","Light002|100|272227","GIBumperInlanes4|100|262126","GIBumperInlanes3|100|161316","GIBumperInlanes2|100|3E373F","GIBumperInlanes1|100|D1B8D2","GIRightRubbersUpper|100|020202","GILeftInlaneLower|100|907F90","GILeftInlaneUpper|100|B49EB4","ActionRightL|100|221E22"), _
Array("LeftOutlanePoisonL|100|CFB7D0","CrossbowLeftStandupL|100|776977","DropTargetsPL|100|554B55","DropTargetsHL|100|D3BAD4","DropTargetsERL|100|E6CBE7","RightHorseshoeTL|0|000000","RightHorseshoeIL|100|100E10","RightHorseshoeML|100|473F47","RightHorseshoeSL|100|C8B0C8","RightOrbitBumpersL|100|D4BBD5","RightOrbitVL|100|6F626F","LeftOrbitGL|100|171417","RightRampEL|100|0B0A0B","RightRampAL|100|211D21","CrossbowPowerupL|100|0F0D0F","HoleShotL|100|BDA7BE","RightRampShotL|100|5B505B","CrossbowShotL|100|CCB4CD","GIRightRamp|100|1D1A1D","GICrossbowLeft|100|726472","Light002|100|2C272D","GIBumperInlanes4|100|19161A","GIBumperInlanes3|100|0F0D0F","GIBumperInlanes2|100|2F292F","GIBumperInlanes1|100|C9B1CA","GIRightRubbersUpper|100|030303","GILeftInlaneLower|100|655966","GILeftInlaneUpper|100|897989","ActionRightL|100|100E10"), _
Array("LeftOutlanePoisonL|100|B29DB2","CrossbowLeftStandupL|100|7D6E7E","StandupsRL|100|FBDEFC","DropTargetsPL|100|574D58","DropTargetsHL|100|D7BDD8","DropTargetsPlusL|100|FADCFB","DropTargetsASL|100|FADDFB","DropTargetsERL|100|E8CCE9","RightHorseshoeIL|100|0E0C0E","RightHorseshoeML|100|403840","RightHorseshoeSL|100|BEA7BF","RightOrbitBumpersL|100|CDB4CD","LeftOrbitLockL|100|010101","CrossbowERL|100|030303","RightOrbitVL|100|615561","RightOrbitKL|100|FBDDFC","LeftOrbitGL|100|1B181B","RightRampEL|100|0D0B0D","RightRampAL|100|221E22","RightRampPowerupL|100|EACFEB","CrossbowPowerupL|100|131013","RightRampShotL|100|605560","CrossbowShotL|100|D3BAD4","GIRightRamp|100|151315","GICrossbowLeft|100|786A79","Light002|100|332D33","GIBumperInlanes4|100|0D0B0D","GIBumperInlanes3|100|070708","GIBumperInlanes2|100|1F1C1F","GIBumperInlanes1|100|BFA8C0","GIRightRubbersUpper|100|050405","GILeftInlaneLower|100|3E373E","GILeftInlaneUpper|100|5B505C","ActionRightL|0|000000"), _
Array("LeftOutlanePoisonL|100|8C7C8D","CrossbowLeftStandupL|100|837484","StandupsRL|100|FBDDFC","DropTargetsPL|100|5A505B","DropTargetsHL|100|DAC0DB","DropTargetsASL|100|FBDDFC","DropTargetsERL|100|E8CDE9","RightHorseshoeIL|100|0C0A0C","RightHorseshoeML|100|393239","RightHorseshoeSL|100|B39EB4","RightOrbitBumpersL|100|C5AEC6","LeftOrbitLockL|100|020202","CrossbowERL|100|060506","RightOrbitVL|100|534953","LeftOrbitGL|100|1F1B1F","RightRampEL|100|0F0D0F","RightRampAL|100|231F23","RightRampPowerupL|100|EBCFEC","CrossbowPowerupL|100|1B171B","HoleShotL|100|BEA7BE","RightRampShotL|100|655966","CrossbowShotL|100|D9BFDA","GIRightRamp|100|100E10","GICrossbowLeft|100|7F7080","Light002|100|393239","GIBumperInlanes4|100|010101","GIBumperInlanes3|100|010101","GIBumperInlanes2|100|110F11","GIBumperInlanes1|100|B39EB4","GIRightRubbersUpper|100|100E10","GILeftInlaneLower|100|1B171B","GILeftInlaneUpper|100|373137","ActionLeftL|100|FCDEFD"), _
Array("LeftOutlanePoisonL|100|675B68","CrossbowLeftStandupL|100|8A7A8A","StandupsTL|100|F2D6F3","DropTargetsPL|100|5C515D","DropTargetsHL|100|DEC3DF","DropTargetsPlusL|100|FADDFB","DropTargetsERL|100|E9CDEA","RightHorseshoeIL|100|090809","RightHorseshoeML|100|332D33","RightHorseshoeSL|100|A894A8","RightOrbitBumpersL|100|BEA8BF","LeftOrbitLockL|100|040304","CrossbowERL|100|0A080A","RightOrbitVL|100|453D45","RightOrbitKL|100|EDD1EE","LeftOrbitGL|100|231E23","RightRampEL|100|110F11","RightRampAL|100|242024","RightOrbitPowerupL|100|FCDFFD","CrossbowPowerupL|100|231F23","RightRampShotL|100|6B5E6B","CrossbowShotL|100|DFC4E0","GIRightRamp|100|0B090B","GICrossbowLeft|100|867686","Light002|100|3F373F","GIBumperInlanes4|0|000000","GIBumperInlanes3|0|000000","GIBumperInlanes2|100|020202","GIBumperInlanes1|100|9E8B9F","GIRightRubbersUpper|100|201C20","GILeftInlaneLower|100|0B0A0B","GILeftInlaneUpper|100|231F23","ActionLeftL|100|FBDDFC"))
lSeqblacksmithMiniWizard.UpdateInterval = 20
lSeqblacksmithMiniWizard.Color = Null
lSeqblacksmithMiniWizard.Repeat = False

Dim lSeqblacksmithWizardKillIntroA : Set lSeqblacksmithWizardKillIntroA = New LCSeq
lSeqblacksmithWizardKillIntroA.Name = "lSeqblacksmithWizardKillIntroA"
lSeqblacksmithWizardKillIntroA.Sequence = Array( Array("CrossbowShotL|100|4F004F"), _
Array("CrossbowShotL|0|000000"), _
Array(), _
Array("RightRampAL|100|B800B8"), _
Array("RightRampEL|100|290029","RightRampAL|0|000000"), _
Array("RightRampEL|100|FF00FF"), _
Array("RightRampEL|0|000000"), _
Array("StandupsNL|100|DA00DA"), _
Array("StandupsNL|100|FF00FF"), _
Array("StandupsNL|100|EE00EE"), _
Array("CrossbowCenterStandupL|100|070007","StandupsNL|0|000000","StandupsIL|100|090009"), _
Array("CrossbowCenterStandupL|100|FF00FF","StandupsIL|100|FF00FF","CrossbowERL|100|FF00FF","Light003|100|050005"), _
Array("CrossbowRightStandupL|100|0C000C","CrossbowCenterStandupL|0|000000","CrossbowERL|100|E100E1","CrossbowIPL|100|FF00FF","Light003|100|FF00FF"), _
Array("CrossbowRightStandupL|100|FF00FF","StandupsIL|100|090009","StandupsAL|100|990099","CrossbowERL|0|000000","CrossbowSNL|100|FF00FF","CrossbowIPL|100|0A000A","Light003|100|CD00CD"), _
Array("CrossbowRightStandupL|0|000000","StandupsIL|0|000000","StandupsAL|100|FF00FF","CrossbowSNL|100|FC00FC","CrossbowIPL|0|000000","Light003|0|000000"), _
Array("StandupsAL|100|F900F9","StandupsRL|100|CE00CE","CrossbowSNL|0|000000","CrossbowPowerupL|100|FF00FF"), _
Array("StandupsAL|0|000000","StandupsRL|100|FF00FF"), _
Array("StandupsRL|100|9E009E","StandupsTL|100|FF00FF","CrossbowPowerupL|0|000000"), _
Array("StandupsRL|0|000000","StandupsTL|100|FE00FE","ScoopPowerupL|100|FF00FF"), _
Array("StandupsTL|0|000000","DropTargetsPlusL|100|FF00FF","ScoopDRL|100|FF00FF","ScoopPowerupL|100|1F001F"), _
Array("DropTargetsHL|100|FF00FF","DropTargetsPlusL|100|290029","ScoopONL|100|830083","ScoopAGL|100|FF00FF","ScoopPowerupL|0|000000","CrossbowShotL|100|F000F0"), _
Array("DropTargetsPL|100|FF00FF","DropTargetsHL|0|000000","DropTargetsPlusL|0|000000","DropTargetsCHL|100|FF00FF","ScoopONL|100|FF00FF","ScoopDRL|0|000000","ScoopAGL|100|6D006D","HoleShotL|100|FF00FF","CrossbowShotL|100|FF00FF"), _
Array("ObjectiveFinalJudgmentL|100|F400F4","DropTargetsPL|0|000000","DropTargetsASL|100|FF00FF","DropTargetsERL|100|FF00FF","DropTargetsCHL|0|000000","ScoopLockL|100|FF00FF","ScoopONL|0|000000","ScoopAGL|0|000000","HoleShotL|100|2C002C"), _
Array("ObjectiveFinalJudgmentL|100|FF00FF","DropTargetsASL|0|000000","DropTargetsERL|100|F800F8","LeftHorseshoeKL|100|FF00FF","LeftOrbitBumpersL|100|FF00FF","ScoopLockL|100|F800F8","RightRampSL|100|FF00FF","HoleShotL|0|000000","CrossbowShotL|0|000000"), _
Array("ObjectiveFinalJudgmentL|0|000000","DropTargetsERL|0|000000","LeftHorseshoeKL|0|000000","LeftOrbitBumpersL|100|410041","ScoopLockL|0|000000","LeftOrbitLockL|100|7C007C","RightRampAL|100|3D003D"), _
Array("ObjectiveBlacksmithL|100|FF00FF","LeftOrbitBumpersL|0|000000","LeftOrbitLockL|100|FF00FF","RightRampAL|100|FF00FF","RightRampSL|0|000000","RightRampPowerupL|100|FF00FF"), _
Array("ObjectiveBlacksmithL|0|000000","LeftOrbitLockL|0|000000","LeftOrbitI1L|100|FF00FF","RightRampShotL|100|FF00FF","GILeftOrbit1|100|FF00FF"), _
Array("StandupsNL|100|640064","LeftOrbitGL|100|080008","LeftOrbitI2L|100|FF00FF","RightRampEL|100|FF00FF","RightRampAL|0|000000","RightRampPowerupL|0|000000","RightRampShotL|0|000000","GILeftOrbit1|0|000000"), _
Array("StandupsNL|100|FF00FF","LeftOrbitGL|100|FF00FF","LeftOrbitI2L|0|000000","LeftOrbitI1L|0|000000","LeftOrbitPowerupL|100|DF00DF"), _
Array("RightHorseshoeHL|100|FF00FF","RightHorseshoeTL|100|FF00FF","LeftOrbitGL|100|0E000E","RightRampEL|0|000000","LeftOrbitPowerupL|100|FF00FF","GICrossbowLeft|100|FF00FF","GILeftOrbit2|100|FF00FF","GIRightRubbersUpper|100|FF00FF"), _
Array("RightHorseshoeHL|100|FC00FC","RightHorseshoeTL|100|3A003A","LeftOrbitGL|0|000000","LeftRampPL|100|FB00FB","LeftOrbitPowerupL|0|000000","LeftOrbitShotL|100|FF00FF","GILeftOrbit2|0|000000","GIRightRubbersUpper|0|000000"), _
Array("CrossbowLeftStandupL|100|FF00FF","StandupsIL|100|F300F3","RightHorseshoeHL|0|000000","RightHorseshoeTL|0|000000","LeftRampPL|100|FF00FF","LeftRampCL|100|FF00FF","LeftRampEL|100|AC00AC","GICrossbowLeft|0|000000"), _
Array("CrossbowCenterStandupL|100|FF00FF","StandupsNL|0|000000","StandupsIL|100|FF00FF","CrossbowERL|100|C500C5","LeftRampPL|0|000000","LeftRampCL|0|000000","LeftRampEL|100|FF00FF","LeftOrbitShotL|0|000000"), _
Array("CrossbowLeftStandupL|0|000000","CrossbowERL|100|FF00FF","CrossbowIPL|100|8C008C","RightOrbitVL|100|FF00FF","LeftRampEL|0|000000","RightOrbitPowerupL|100|FF00FF","LeftRampPowerupL|100|FF00FF","Light003|100|FF00FF","GIRightRubbersLower|100|FF00FF"), _
Array("CrossbowRightStandupL|100|FF00FF","CrossbowCenterStandupL|0|000000","StandupsAL|100|FB00FB","RightOrbitBumpersL|100|FF00FF","CrossbowIPL|100|FF00FF","RightOrbitVL|0|000000","RightOrbitKL|100|FF00FF","RightOrbitNL|100|FF00FF","RightOrbitPowerupL|0|000000","LeftRampPowerupL|0|000000","GIRightRubbersLower|100|9E009E"), _
Array("StandupsIL|100|330033","StandupsAL|100|FF00FF","CrossbowERL|100|830083","CrossbowSNL|100|FF00FF","RightOrbitKL|0|000000","RightOrbitNL|100|FA00FA","LeftRampShotL|100|FF00FF","Light001|100|FF00FF","GIRightRubbersLower|0|000000"), _
Array("CrossbowRightStandupL|0|000000","StandupsIL|0|000000","StandupsRL|100|F300F3","RightOrbitBumpersL|0|000000","CrossbowERL|0|000000","CrossbowIPL|0|000000","RightOrbitNL|0|000000","CrossbowPowerupL|100|4C004C","Light003|0|000000","Light001|0|000000"), _
Array("StandupsRL|100|FF00FF","CrossbowSNL|100|A600A6","CrossbowPowerupL|100|FF00FF","LeftRampShotL|0|000000","GILeftRubbersUpper|100|F000F0"), _
Array("StandupsAL|0|000000","StandupsTL|100|FF00FF","CrossbowSNL|0|000000","GILeftRubbersUpper|100|FF00FF","GIRightSlingshotUpper|100|CF00CF"), _
Array("StandupsRL|100|C100C1","ScoopShotL|100|FF00FF","GIScoop2|100|FF00FF","GILeftRubbersUpper|0|000000","GIRightSlingshotUpper|100|FD00FD"), _
Array("StandupsRL|0|000000","ScoopPowerupL|100|190019","CrossbowPowerupL|0|000000","ScoopShotL|100|050005","GIScoop2|0|000000","Light002|100|FF00FF","GIRightSlingshotUpper|0|000000"), _
Array("ObjectiveVikingL|100|FF00FF","StandupsTL|0|000000","DropTargetsPlusL|100|FF00FF","ScoopPowerupL|100|FF00FF","ScoopShotL|0|000000"), _
Array("ObjectiveVikingL|0|000000","DropTargetsPL|100|9C009C","DropTargetsHL|100|FF00FF","DropTargetsCHL|100|F700F7","ScoopDRL|100|FF00FF","ScoopAGL|100|FF00FF","ScoopPowerupL|100|5D005D","HoleShotL|100|FF00FF","CrossbowShotL|100|FF00FF","Light002|0|000000"), _
Array("ObjectiveSniperL|100|FF00FF","ObjectiveChaserL|100|FF00FF","DropTargetsPL|100|FF00FF","DropTargetsHL|100|B400B4","DropTargetsPlusL|0|000000","DropTargetsASL|100|FF00FF","DropTargetsCHL|100|FF00FF","ScoopONL|100|FF00FF","RightRampSL|100|680068","ScoopPowerupL|0|000000"), _
Array("ObjectiveSniperL|0|000000","ObjectiveChaserL|0|000000","DropTargetsPL|0|000000","DropTargetsHL|0|000000","DropTargetsERL|100|FF00FF","DropTargetsCHL|100|900090","LeftHorseshoeKL|100|FD00FD","LeftOrbitBumpersL|100|C500C5","ScoopLockL|100|FF00FF","ScoopONL|100|FD00FD","ScoopDRL|0|000000","ScoopAGL|100|C500C5","RightRampSL|100|FF00FF","HoleShotL|100|900090"), _
Array("ObjectiveFinalJudgmentL|100|FF00FF","ObjectiveFinalBossL|100|FF00FF","ObjectiveCastleL|100|FF00FF","DropTargetsASL|0|000000","DropTargetsCHL|0|000000","LeftHorseshoeKL|100|FF00FF","LeftHorseshoeCL|100|FF00FF","LeftHorseshoeAL|100|FF00FF","LeftHorseshoeLL|100|FF00FF","LeftHorseshoeBL|100|CC00CC","LeftOrbitBumpersL|100|FF00FF","ScoopONL|0|000000","ScoopAGL|0|000000","RightRampAL|100|EF00EF","HoleShotL|0|000000","GIUpperFlipper|100|FF00FF"), _
Array("ObjectiveFinalBossL|100|F300F3","ObjectiveCastleL|0|000000","DropTargetsERL|0|000000","LeftHorseshoeKL|0|000000","LeftHorseshoeCL|0|000000","LeftHorseshoeAL|100|460046","LeftHorseshoeBL|100|FF00FF","ScoopLockL|100|FA00FA","LeftOrbitLockL|100|EF00EF","RightRampAL|100|FF00FF","RightRampPowerupL|100|F100F1","LeftHorseshoeShotL|100|FF00FF","CrossbowShotL|0|000000","GIUpperFlipper|100|F000F0"), _
Array("ObjectiveFinalJudgmentL|0|000000","ObjectiveBlacksmithL|100|B600B6","ObjectiveFinalBossL|0|000000","ObjectiveDragonL|100|FF00FF","LeftHorseshoeAL|0|000000","LeftHorseshoeLL|0|000000","LeftHorseshoeBL|0|000000","LeftOrbitBumpersL|0|000000","ScoopLockL|0|000000","LeftOrbitLockL|100|FF00FF","RightRampEL|100|8E008E","RightRampSL|0|000000","RightRampPowerupL|100|FF00FF","RightRampShotL|100|FF00FF","LeftHorseshoeShotL|0|000000","GIUpperFlipper|0|000000"), _
Array("ObjectiveBlacksmithL|100|FF00FF","ObjectiveDragonL|100|F200F2","StandupsNL|100|5F005F","LeftOrbitI2L|100|770077","LeftOrbitI1L|100|FF00FF","RightRampEL|100|FF00FF","GIRightRamp|100|FA00FA","GILeftOrbit1|100|FF00FF"), _
Array("ObjectiveBlacksmithL|0|000000","ObjectiveEscapeL|100|FF00FF","ObjectiveDragonL|0|000000","StandupsNL|100|FF00FF","LeftOrbitLockL|0|000000","LeftOrbitGL|100|9E009E","LeftOrbitI2L|100|FF00FF","RightRampAL|0|000000","RightRampPowerupL|0|000000","RightRampShotL|0|000000","GIRightRamp|100|FF00FF"), _
Array("LeftOrbitGL|100|FF00FF","LeftOrbitI1L|0|000000","LeftOrbitPowerupL|100|D700D7","GIRightRamp|0|000000","GILeftOrbit1|0|000000","GILeftSlingshotLower|100|FF00FF"), _
Array("ObjectiveEscapeL|0|000000","RightHorseshoeHL|100|FF00FF","RightHorseshoeTL|100|FF00FF","RightHorseshoeIL|100|FF00FF","RightHorseshoeML|100|FF00FF","RightHorseshoeSL|100|FF00FF","LeftOrbitI2L|0|000000","RightRampEL|0|000000","LeftOrbitPowerupL|100|FF00FF","RightHorseshoeShotL|100|FF00FF","GICrossbowLeft|100|FF00FF","GILeftOrbit2|100|FF00FF","GIRightRubbersUpper|100|FF00FF","GILeftSlingshotLower|0|000000","GILeftSlingshot|100|330033"), _
Array("CrossbowLeftStandupL|100|280028","StandupsIL|100|AB00AB","RightHorseshoeIL|100|E100E1","RightHorseshoeML|0|000000","RightHorseshoeSL|0|000000","LeftOrbitGL|0|000000","RightHorseshoeShotL|0|000000","LeftOrbitShotL|100|F400F4","GILeftSlingshot|100|FF00FF"), _
Array("SpinnerShotL|100|FF00FF","CrossbowLeftStandupL|100|FF00FF","StandupsIL|100|FF00FF","RightHorseshoeHL|0|000000","RightHorseshoeTL|0|000000","RightHorseshoeIL|0|000000","LeftRampPL|100|FF00FF","LeftOrbitPowerupL|0|000000","LeftOrbitShotL|100|FF00FF","GILeftOrbit2|0|000000","GIRightRubbersUpper|0|000000","GILeftSlingshot|0|000000"), _
Array("SpinnerShotL|0|000000","CrossbowCenterStandupL|100|FF00FF","RightOrbitVL|100|FD00FD","LeftRampCL|100|FF00FF","LeftRampEL|100|FF00FF","RightOrbitShotL|100|FF00FF","LeftOrbitShotL|0|000000","GICrossbowLeft|0|000000","Light003|100|EE00EE","GILeftOrbit3|100|FC00FC"), _
Array("CrossbowRightStandupL|100|630063","CrossbowLeftStandupL|100|340034","StandupsNL|100|040004","StandupsAL|100|980098","CrossbowERL|100|100010","RightOrbitVL|100|FF00FF","RightOrbitKL|100|FF00FF","LeftRampPL|100|410041","LeftRampCL|100|F800F8","RightOrbitPowerupL|100|FF00FF","LeftRampPowerupL|100|250025","RightOrbitShotL|100|A200A2","Light003|100|FF00FF","GILeftOrbit3|100|FF00FF"), _
Array("CrossbowRightStandupL|100|FF00FF","CrossbowLeftStandupL|0|000000","StandupsNL|0|000000","StandupsAL|100|FF00FF","RightOrbitBumpersL|100|FF00FF","CrossbowERL|100|FF00FF","RightOrbitVL|0|000000","RightOrbitNL|100|FF00FF","LeftRampPL|0|000000","LeftRampCL|0|000000","LeftRampEL|0|000000","RightOrbitPowerupL|0|000000","LeftRampPowerupL|100|FF00FF","RightOrbitShotL|0|000000","GIScoop1|100|FF00FF","GILeftOrbit3|0|000000","GIRightRubbersLower|100|FF00FF"), _
Array("CrossbowCenterStandupL|0|000000","CrossbowIPL|100|FF00FF","RightOrbitKL|0|000000","GIScoop1|100|310031","GIRightRubbersLower|100|070007"), _
Array("StandupsIL|100|6E006E","StandupsRL|100|FF00FF","RightOrbitBumpersL|0|000000","CrossbowSNL|100|FF00FF","RightOrbitNL|0|000000","LeftRampPowerupL|0|000000","LeftRampShotL|100|FF00FF","GIScoop1|0|000000","Light003|100|D600D6","Light001|100|FF00FF","GIRightRubbersLower|0|000000"), _
Array("RightOutlanePoisonL|100|920092","CrossbowRightStandupL|0|000000","StandupsIL|0|000000","StandupsTL|100|540054","Light003|0|000000","GIBumperInlanes1|100|180018","Light001|0|000000"), _
Array("RightOutlanePoisonL|100|FF00FF","StandupsTL|100|FF00FF","CrossbowERL|100|490049","CrossbowIPL|100|F400F4","CrossbowPowerupL|100|FF00FF","ScoopShotL|100|FB00FB","LeftRampShotL|100|BD00BD","GIScoop2|100|FA00FA","GIBumperInlanes1|100|FF00FF","GILeftRubbersUpper|100|FF00FF"), _
Array("RightOutlanePoisonL|0|000000","StandupsAL|0|000000","CrossbowERL|0|000000","CrossbowSNL|100|F900F9","CrossbowIPL|0|000000","ScoopShotL|100|FF00FF","LeftRampShotL|0|000000","GIScoop2|100|FF00FF","GIBumperInlanes2|100|F600F6","GIBumperInlanes1|0|000000","GIRightSlingshotUpper|100|FF00FF"), _
Array("StandupsRL|100|8A008A","DropTargetsPlusL|100|100010","ScoopDRL|100|060006","CrossbowSNL|0|000000","ScoopPowerupL|100|FF00FF","ScoopShotL|100|490049","GIScoop2|0|000000","Light002|100|FE00FE","GIBumperInlanes2|100|FF00FF","GILeftRubbersUpper|0|000000","GIRightInlaneUpper|100|FF00FF","GIRightSlingshotUpper|100|8F008F","GIRightSlingshotLow|100|FF00FF"), _
Array("ObjectiveVikingL|100|FF00FF","StandupsRL|0|000000","DropTargetsHL|100|130013","DropTargetsPlusL|100|FF00FF","ScoopDRL|100|FF00FF","ScoopAGL|100|6F006F","HoleShotL|100|560056","ScoopShotL|0|000000","Light002|100|FF00FF","GIBumperInlane2|100|FF00FF","GIBumperInlanes3|100|FF00FF","GIBumperInlanes2|0|000000","GIRightInlaneUpper|0|000000","GIRightInlaneLower|100|FF00FF","GIRightSlingshotUpper|0|000000"), _
Array("StandupsTL|0|000000","DropTargetsPL|100|7B007B","DropTargetsHL|100|FF00FF","DropTargetsCHL|100|F700F7","ScoopONL|100|FF00FF","ScoopAGL|100|FF00FF","ScoopPowerupL|100|350035","CrossbowPowerupL|100|F600F6","HoleShotL|100|FF00FF","Light002|100|D500D5","GIBumperInlane2|0|000000","GIRightInlaneLower|0|000000","GIRightSlingshotLow|0|000000"), _
Array("ObjectiveSniperL|100|FF00FF","ObjectiveVikingL|0|000000","ObjectiveChaserL|100|FF00FF","DropTargetsPL|100|FF00FF","DropTargetsPlusL|100|2D002D","DropTargetsASL|100|FB00FB","DropTargetsCHL|100|FF00FF","ScoopLockL|100|EE00EE","RightRampSL|100|F100F1","ScoopPowerupL|0|000000","CrossbowPowerupL|0|000000","CrossbowShotL|100|9F009F","Light002|0|000000","GIBumperInlanes4|100|F100F1","GIBumperInlanes3|0|000000","ActionRightL|100|FF00FF"), _
Array("DropTargetsHL|0|000000","DropTargetsPlusL|0|000000","DropTargetsASL|100|FF00FF","DropTargetsERL|100|FF00FF","LeftHorseshoeKL|100|F100F1","ScoopLockL|100|FF00FF","ScoopDRL|0|000000","ScoopAGL|100|430043","RightRampSL|100|FF00FF","CrossbowShotL|100|FF00FF","GIUpperFlipper|100|F000F0","GIBumperInlanes4|100|FF00FF","ActionRightL|0|000000"), _
Array("ObjectiveFinalJudgmentL|100|A600A6","ObjectiveSniperL|0|000000","ObjectiveFinalBossL|100|FF00FF","ObjectiveChaserL|0|000000","ObjectiveCastleL|100|FA00FA","DropTargetsPL|0|000000","DropTargetsCHL|0|000000","LeftHorseshoeKL|100|FF00FF","LeftHorseshoeCL|100|FF00FF","LeftHorseshoeAL|100|FF00FF","LeftHorseshoeLL|100|F200F2","ScoopONL|0|000000","ScoopAGL|0|000000","RightRampAL|100|FA00FA","HoleShotL|100|170017","GIUpperFlipper|100|FF00FF","GIVUK|100|6B006B","GIBumperInlanes4|0|000000"), _
Array("ObjectiveFinalJudgmentL|100|FF00FF","ObjectiveCastleL|100|FF00FF","DropTargetsASL|0|000000","DropTargetsERL|100|B400B4","LeftHorseshoeLL|100|FF00FF","LeftHorseshoeBL|100|FF00FF","LeftOrbitBumpersL|100|FF00FF","ScoopLockL|100|B400B4","RightRampAL|100|FF00FF","RightRampPowerupL|100|F700F7","HoleShotL|0|000000","LeftHorseshoeShotL|100|FF00FF","GIUpperFlipper|100|FD00FD","GIVUK|100|FF00FF"), _
Array("ObjectiveFinalBossL|100|A200A2","ObjectiveDragonL|100|A700A7","ObjectiveCastleL|0|000000","DropTargetsERL|0|000000","LeftHorseshoeKL|0|000000","LeftHorseshoeCL|0|000000","LeftHorseshoeAL|0|000000","LeftHorseshoeLL|0|000000","LeftHorseshoeBL|100|650065","ScoopLockL|0|000000","LeftOrbitLockL|100|280028","RightRampEL|100|F300F3","RightRampPowerupL|100|FF00FF","RightRampShotL|100|FF00FF","GIUpperFlipper|0|000000","GIVUK|0|000000"), _
Array("ObjectiveFinalJudgmentL|0|000000","ObjectiveBlacksmithL|100|FF00FF","ObjectiveFinalBossL|0|000000","ObjectiveDragonL|100|FF00FF","LeftHorseshoeBL|0|000000","LeftOrbitLockL|100|FF00FF","RightRampEL|100|FF00FF","RightRampSL|0|000000","LeftHorseshoeShotL|0|000000","CrossbowShotL|100|FC00FC","GIRightRamp|100|610061","ActionLeftL|100|FF00FF"), _
Array("ObjectiveDragonL|100|640064","LeftOrbitBumpersL|100|EE00EE","LeftOrbitI1L|100|FF00FF","CrossbowShotL|100|450045","GIRightRamp|100|FF00FF","GIRightOrbit1|100|FF00FF","GILeftOrbit1|100|FF00FF","ActionLeftL|0|000000"), _
Array("ObjectiveBlacksmithL|0|000000","ObjectiveEscapeL|100|FF00FF","ObjectiveDragonL|0|000000","LeftOrbitBumpersL|0|000000","LeftOrbitI2L|100|FF00FF","RightRampAL|0|000000","RightRampPowerupL|0|000000","RightRampShotL|0|000000","CrossbowShotL|0|000000","GIRightRamp|0|000000","GIRightOrbit1|0|000000","GILeftInlaneLower|100|FF00FF","GILeftSlingshotLower|100|FF00FF"), _
Array("StandupsNL|100|FF00FF","RightHorseshoeHL|100|FF00FF","RightHorseshoeTL|100|FF00FF","RightHorseshoeIL|100|FF00FF","RightHorseshoeML|100|FF00FF","RightHorseshoeSL|100|FF00FF","LeftOrbitLockL|100|4E004E","LeftOrbitGL|100|FF00FF","LeftOrbitPowerupL|100|350035","RightHorseshoeShotL|100|FF00FF","GIRightOrbit3|100|FF00FF","GILeftOrbit1|100|560056","GIRightRubbersUpper|100|FF00FF","GILeftInlaneLower|0|000000","GILeftInlaneUpper|100|FF00FF"), _
Array("ObjectiveEscapeL|0|000000","LeftOrbitLockL|0|000000","LeftOrbitI1L|100|490049","RightRampEL|0|000000","LeftOrbitPowerupL|100|FF00FF","GICrossbowLeft|100|F900F9","GIRightOrbit3|0|000000","GIRightOrbit2|100|700070","GILeftOrbit2|100|FF00FF","GILeftOrbit1|0|000000","GILeftInlaneUpper|0|000000","GILeftSlingshotLower|0|000000","GILeftSlingshot|100|FF00FF"), _
Array("LeftOutlanePoisonL|100|FF00FF","SpinnerShotL|100|FA00FA","RightHorseshoeHL|100|950095","RightHorseshoeTL|0|000000","RightHorseshoeIL|0|000000","RightHorseshoeML|0|000000","RightHorseshoeSL|0|000000","LeftOrbitI2L|0|000000","LeftOrbitI1L|0|000000","LeftRampPL|100|910091","RightHorseshoeShotL|0|000000","LeftOrbitShotL|100|CA00CA","GICrossbowLeft|100|FF00FF","GIRightOrbit2|100|FF00FF","GIRightRubbersUpper|0|000000","GILeftSlingshot|100|FE00FE"), _
Array("LeftOutlanePoisonL|100|0F000F","SpinnerShotL|100|FF00FF","CrossbowLeftStandupL|100|FF00FF","StandupsIL|100|500050","RightHorseshoeHL|0|000000","RightOrbitVL|100|F300F3","LeftOrbitGL|100|850085","LeftRampPL|100|FF00FF","LeftRampCL|100|FF00FF","LeftOrbitPowerupL|100|F200F2","RightOrbitShotL|100|FF00FF","LeftOrbitShotL|100|FF00FF","GIRightOrbit2|0|000000","GILeftOrbit2|100|3E003E","GILeftSlingshot|0|000000"), _
Array("LeftOutlanePoisonL|0|000000","SpinnerShotL|0|000000","CrossbowCenterStandupL|100|FD00FD","StandupsIL|100|FF00FF","CrossbowERL|100|030003","RightOrbitVL|100|FF00FF","RightOrbitKL|100|F900F9","LeftOrbitGL|0|000000","LeftRampEL|100|FF00FF","RightOrbitPowerupL|100|FF00FF","LeftOrbitPowerupL|0|000000","GILeftOrbit3|100|FF00FF","GILeftOrbit2|0|000000","GIRightRubbersLower|100|FF00FF"), _
Array("CrossbowCenterStandupL|100|FF00FF","RightOrbitBumpersL|100|FF00FF","CrossbowERL|100|FF00FF","RightOrbitKL|100|FF00FF","RightOrbitNL|100|FF00FF","LeftRampPL|100|570057","LeftRampCL|100|F200F2","LeftRampPowerupL|100|FF00FF","RightOrbitShotL|0|000000","LeftOrbitShotL|0|000000","GIScoop1|100|FF00FF","GICrossbowLeft|100|180018","Light003|100|FF00FF"), _
Array("CrossbowRightStandupL|100|FF00FF","CrossbowLeftStandupL|100|F000F0","StandupsAL|100|740074","CrossbowIPL|100|FF00FF","RightOrbitVL|0|000000","LeftRampPL|0|000000","LeftRampCL|0|000000","LeftRampEL|0|000000","RightOrbitPowerupL|0|000000","GICrossbowLeft|0|000000","GILeftOrbit3|0|000000","GIRightRubbersLower|0|000000"), _
Array("CrossbowLeftStandupL|0|000000","StandupsAL|100|FF00FF","CrossbowSNL|100|F400F4","RightOrbitKL|0|000000","RightOrbitNL|100|650065","LeftRampPowerupL|100|FA00FA","LeftRampShotL|100|FC00FC","GIScoop1|0|000000","Light001|100|FF00FF"), _
Array("RightOutlanePoisonL|100|FF00FF","CrossbowCenterStandupL|100|D100D1","StandupsNL|100|F400F4","StandupsRL|100|2B002B","RightOrbitBumpersL|0|000000","CrossbowSNL|100|FF00FF","RightOrbitNL|0|000000","LeftRampPowerupL|0|000000","LeftRampShotL|100|FF00FF"), _
Array("CrossbowCenterStandupL|0|000000","StandupsNL|100|800080","StandupsRL|100|FF00FF","CrossbowPowerupL|100|FF00FF","ScoopShotL|100|F300F3","GIScoop2|100|DA00DA","GIBumperInlanes1|100|E800E8","GILeftOrbit4|100|FF00FF","GILeftRubbersUpper|100|0A000A","Light001|0|000000","GIRightSlingshotUpper|100|F300F3"), _
Array("RightOutlanePoisonL|0|000000","CrossbowRightStandupL|100|F900F9","StandupsNL|0|000000","StandupsTL|100|D000D0","CrossbowERL|100|F000F0","CrossbowIPL|100|FD00FD","ScoopShotL|100|FF00FF","LeftRampShotL|100|C300C3","GIScoop2|100|FF00FF","Light003|100|F500F5","GIBumperInlanes1|100|FF00FF","GILeftOrbit4|0|000000","GILeftRubbersUpper|100|FF00FF","GIRightInlaneUpper|100|F200F2","GIRightSlingshotUpper|100|FF00FF"), _
Array("ObjectiveVikingL|100|C200C2","CrossbowRightStandupL|0|000000","StandupsIL|100|F200F2","StandupsTL|100|FF00FF","ScoopDRL|100|5C005C","CrossbowERL|0|000000","CrossbowIPL|100|320032","ScoopPowerupL|100|FF00FF","LeftRampShotL|0|000000","Light003|0|000000","GIBumperInlanes2|100|FF00FF","GIBumperInlanes1|0|000000","GIRightInlaneUpper|100|FF00FF","GIRightSlingshotUpper|100|630063","GIRightSlingshotLow|100|FF00FF"), _
Array("ObjectiveVikingL|100|FF00FF","StandupsIL|100|440044","ScoopDRL|100|FF00FF","ScoopAGL|100|960096","CrossbowSNL|100|290029","CrossbowIPL|0|000000","ScoopShotL|0|000000","GIScoop2|0|000000","Light002|100|CC00CC","GILeftRubbersUpper|0|000000","GIRightInlaneUpper|0|000000","GIRightInlaneLower|100|FF00FF","GIRightSlingshotUpper|0|000000"), _
Array("ObjectiveSniperL|100|F300F3","StandupsIL|0|000000","StandupsAL|100|EF00EF","DropTargetsPlusL|100|FF00FF","ScoopONL|100|FF00FF","ScoopAGL|100|FF00FF","CrossbowSNL|0|000000","Light002|100|FF00FF","GIBumperInlane2|100|FF00FF","GIBumperInlanes3|100|FF00FF","GIBumperInlanes2|0|000000","GIRightInlaneLower|0|000000","GIRightSlingshotLow|0|000000"), _
Array("ObjectiveSniperL|100|FF00FF","ObjectiveVikingL|0|000000","ObjectiveChaserL|100|FF00FF","StandupsAL|0|000000","StandupsRL|100|F500F5","DropTargetsHL|100|FF00FF","DropTargetsCHL|100|200020","ScoopLockL|100|F000F0","ScoopPowerupL|0|000000","HoleShotL|100|FF00FF","CrossbowShotL|100|870087","GIBumperInlane2|0|000000","ActionRightL|100|FF00FF"), _
Array("ObjectiveFinalJudgmentL|100|660066","StandupsRL|100|100010","StandupsTL|100|F200F2","DropTargetsPL|100|FF00FF","DropTargetsASL|100|5C005C","DropTargetsCHL|100|FF00FF","ScoopLockL|100|FF00FF","ScoopDRL|100|830083","CrossbowPowerupL|100|A300A3","CrossbowShotL|100|FF00FF","GIUpperFlipper|100|AC00AC","Light002|100|BA00BA","GIBumperInlanes4|100|FF00FF","GIBumperInlanes3|0|000000","ActionRightL|0|000000"), _
Array("ObjectiveFinalJudgmentL|100|FF00FF","ObjectiveSniperL|0|000000","ObjectiveFinalBossL|100|FF00FF","ObjectiveChaserL|0|000000","ObjectiveCastleL|100|FF00FF","StandupsRL|0|000000","StandupsTL|0|000000","DropTargetsPlusL|100|D100D1","DropTargetsASL|100|FF00FF","DropTargetsERL|100|FF00FF","LeftHorseshoeKL|100|450045","ScoopONL|100|450045","ScoopDRL|0|000000","ScoopAGL|0|000000","RightRampSL|100|FF00FF","CrossbowPowerupL|0|000000","GIUpperFlipper|100|FF00FF","Light002|0|000000"))
lSeqblacksmithWizardKillIntroA.UpdateInterval = 20
lSeqblacksmithWizardKillIntroA.Color = Null
lSeqblacksmithWizardKillIntroA.Repeat = False

Dim lSeqpowerupCollected : Set lSeqpowerupCollected = New LCSeq
lSeqpowerupCollected.Name = "lSeqpowerupCollected"
lSeqpowerupCollected.Sequence = Array( Array("RightRampPowerupL|100|1B593C"), _
Array("RightOrbitPowerupL|100|3C2400"), _
Array("RightOrbitPowerupL|100|C87900"), _
Array(), _
Array(), _
Array(), _
Array(), _
Array(), _
Array("RightRampPowerupL|100|C87900"), _
Array("RightOrbitPowerupL|100|C8401F"), _
Array("LeftOrbitPowerupL|100|1B593C"), _
Array(), _
Array(), _
Array(), _
Array(), _
Array(), _
Array("RightOrbitPowerupL|100|C70101"), _
Array(), _
Array("RightRampPowerupL|100|C8401F","ScoopPowerupL|100|240B05"), _
Array("ScoopPowerupL|100|C8401F"), _
Array(), _
Array("ScoopPowerupL|100|C70101"), _
Array("CrossbowPowerupL|100|C87900"), _
Array("RightOrbitPowerupL|100|BF06A0"), _
Array("RightRampPowerupL|100|C83B1C"), _
Array("RightRampPowerupL|100|C70101"), _
Array(), _
Array(), _
Array("RightRampPowerupL|100|C0059A"), _
Array("LeftOrbitPowerupL|100|997110","CrossbowPowerupL|100|C8401F","RightRampPowerupL|100|BF06A0","ScoopPowerupL|100|BF06A0"), _
Array("LeftOrbitPowerupL|100|C87900","RightOrbitPowerupL|100|7A2AB9"), _
Array("RightRampPowerupL|100|7A2AB9"), _
Array(), _
Array(), _
Array("RightRampPowerupL|100|0000C8"), _
Array(), _
Array(), _
Array("RightOrbitPowerupL|100|0000C8","ScoopPowerupL|100|7A2AB9"), _
Array("LeftRampPowerupL|100|C8401F"), _
Array(), _
Array("RightRampPowerupL|100|0051B6"), _
Array(), _
Array(), _
Array(), _
Array("LeftRampPowerupL|100|C83A1C","ScoopPowerupL|100|0000C8"), _
Array("LeftRampPowerupL|100|C70101","RightOrbitPowerupL|100|001AC2"), _
Array("RightOrbitPowerupL|100|0051B6"), _
Array(), _
Array("RightRampPowerupL|100|165752","RightOrbitPowerupL|0|000000"), _
Array("RightRampPowerupL|100|1B593C"), _
Array(), _
Array("ScoopPowerupL|100|0051B6"), _
Array(), _
Array(), _
Array(), _
Array(), _
Array(), _
Array("LeftOrbitPowerupL|100|C8401F"), _
Array(), _
Array("RightRampPowerupL|100|C87900","ScoopPowerupL|100|1B593C"), _
Array(), _
Array("LeftOrbitPowerupL|100|63200F"), _
Array("LeftOrbitPowerupL|0|000000","RightRampPowerupL|0|000000"), _
Array("ScoopPowerupL|100|164A32"), _
Array("ScoopPowerupL|0|000000"), _
Array(), _
Array(), _
Array(), _
Array(), _
Array(), _
Array(), _
Array("CrossbowPowerupL|100|C70101"), _
Array(), _
Array(), _
Array(), _
Array("LeftRampPowerupL|100|BF06A0"), _
Array(), _
Array(), _
Array(), _
Array(), _
Array(), _
Array(), _
Array(), _
Array("CrossbowPowerupL|0|000000"), _
Array(), _
Array(), _
Array(), _
Array(), _
Array(), _
Array(), _
Array())
lSeqpowerupCollected.UpdateInterval = 20
lSeqpowerupCollected.Color = Null
lSeqpowerupCollected.Repeat = False

Dim lSeqpowerupDeployed : Set lSeqpowerupDeployed = New LCSeq
lSeqpowerupDeployed.Name = "lSeqpowerupDeployed"
lSeqpowerupDeployed.Sequence = Array( Array(), _
Array(), _
Array(), _
Array(), _
Array(), _
Array(), _
Array("ActionButton|100|00D6FF"), _
Array(), _
Array("ActionButton|0|000000"), _
Array(), _
Array("ActionRightL|100|006072"), _
Array("ObjectiveCastleL|100|00D6FF","ActionLeftL|100|00D6FF","ActionRightL|100|00D6FF"), _
Array("ObjectiveFinalBossL|100|00CAF1"), _
Array("ObjectiveFinalBossL|100|00D6FF","ObjectiveCastleL|100|000B0D","ActionButton|100|00D6FF","ActionLeftL|0|000000","ActionRightL|0|000000"), _
Array("ObjectiveFinalJudgmentL|100|00D6FF","ObjectiveDragonL|100|000404","ObjectiveCastleL|0|000000"), _
Array("ObjectiveFinalBossL|0|000000","ObjectiveChaserL|100|00D6FF","ObjectiveDragonL|100|00D6FF","ActionButton|0|000000"), _
Array("ObjectiveFinalJudgmentL|100|000608","ObjectiveSniperL|100|00D6FF","ObjectiveBlacksmithL|100|00D6FF","GIRightInlaneLower|100|00404D"), _
Array("ObjectiveFinalJudgmentL|0|000000","ObjectiveChaserL|0|000000","ObjectiveDragonL|0|000000","GIRightInlaneLower|100|00D6FF","GILeftInlaneLower|100|00D6FF","ActionRightL|100|007D95"), _
Array("ObjectiveSniperL|100|00819A","ObjectiveBlacksmithL|100|00CBF2","ObjectiveVikingL|100|006D81","ObjectiveCastleL|100|00D6FF","GILeftSlingshotLower|100|00080A","GIRightSlingshotLow|100|004C5A","ActionLeftL|100|00D6FF","ActionRightL|100|00D6FF"), _
Array("ObjectiveSniperL|0|000000","ObjectiveBlacksmithL|0|000000","ObjectiveVikingL|100|00D6FF","ObjectiveFinalBossL|100|00D5FE","ObjectiveEscapeL|100|00D6FF","GIRightInlaneLower|0|000000","GILeftInlaneLower|100|000F12","GILeftSlingshotLower|100|00D6FF","GIRightSlingshotLow|100|00D6FF","ActionButton|100|00BCDF","ActionRightL|100|00D2FB"), _
Array("ObjectiveFinalBossL|100|00D6FF","ObjectiveCastleL|100|000E10","GIRightInlaneUpper|100|00D6FF","GILeftInlaneLower|0|000000","GILeftInlaneUpper|100|00D6FF","ActionButton|100|00D6FF","ActionLeftL|0|000000","ActionRightL|0|000000"), _
Array("ObjectiveFinalJudgmentL|100|00D6FF","ObjectiveVikingL|0|000000","ObjectiveEscapeL|0|000000","ObjectiveChaserL|100|005262","ObjectiveCastleL|0|000000","GILeftSlingshotLower|0|000000","GIRightSlingshotLow|0|000000"), _
Array("ObjectiveFinalBossL|0|000000","ObjectiveChaserL|100|00D6FF","ObjectiveDragonL|100|00D6FF","GIRightInlaneUpper|0|000000","GILeftInlaneUpper|100|00687C","GIRightSlingshotUpper|100|00D6FF","GILeftSlingshot|100|00C3E9","ActionButton|0|000000"), _
Array("ObjectiveFinalJudgmentL|0|000000","ObjectiveSniperL|100|00D6FF","ObjectiveBlacksmithL|100|00D6FF","GIRightInlaneLower|100|009DBB","GILeftInlaneUpper|0|000000","GILeftSlingshot|100|00D6FF"), _
Array("ObjectiveChaserL|0|000000","ObjectiveDragonL|0|000000","CrossbowERL|100|00A0BF","LeftRampPL|100|00D6FF","GIRightInlaneLower|100|00D6FF","GILeftInlaneLower|100|00D6FF","GIRightSlingshotUpper|100|004E5D","ActionRightL|100|00C6EC"), _
Array("LeftOutlanePoisonL|100|00D6FF","RightOutlanePoisonL|100|00D6FF","ObjectiveSniperL|100|00343E","ObjectiveBlacksmithL|100|0091AD","ObjectiveVikingL|100|0091AC","ObjectiveEscapeL|100|00171C","ObjectiveCastleL|100|00D6FF","CrossbowERL|100|00D6FF","CrossbowIPL|100|008DA8","LeftRampCL|100|00D6FF","GILeftSlingshotLower|100|000A0C","GIRightSlingshotUpper|0|000000","GILeftSlingshot|0|000000","GIRightSlingshotLow|100|006173","ActionLeftL|100|00D6FF","ActionRightL|100|00D6FF"), _
Array("ObjectiveSniperL|0|000000","ObjectiveBlacksmithL|0|000000","ObjectiveVikingL|100|00D6FF","ObjectiveFinalBossL|100|00D4FD","ObjectiveEscapeL|100|00D6FF","StandupsNL|100|003C48","CrossbowIPL|100|00D6FF","LeftRampPL|100|00B1D3","LeftRampEL|100|00D6FF","CrossbowPowerupL|100|000C0F","GIRightInlaneLower|0|000000","GILeftInlaneLower|0|000000","GILeftSlingshotLower|100|00D6FF","GIRightSlingshotLow|100|00D6FF","ActionRightL|100|00D2FB"), _
Array("LeftOutlanePoisonL|100|00D3FB","RightOutlanePoisonL|100|00191D","ObjectiveFinalBossL|100|00D6FF","ObjectiveCastleL|100|00171B","StandupsNL|100|00D6FF","StandupsIL|100|00D6FF","StandupsAL|100|008AA5","CrossbowERL|0|000000","CrossbowSNL|100|00D6FF","LeftRampPL|0|000000","LeftRampCL|100|003D48","CrossbowPowerupL|100|00D6FF","LeftRampPowerupL|100|00D6FF","GIRightInlaneUpper|100|00D6FF","GILeftInlaneUpper|100|00D6FF","ActionLeftL|0|000000","ActionRightL|0|000000"), _
Array("LeftOutlanePoisonL|0|000000","RightOutlanePoisonL|0|000000","ObjectiveFinalJudgmentL|100|00D6FF","ObjectiveVikingL|0|000000","ObjectiveEscapeL|0|000000","ObjectiveChaserL|100|00D4FD","ObjectiveCastleL|0|000000","StandupsAL|100|00D6FF","StandupsRL|100|00D6FF","StandupsTL|100|00D5FE","CrossbowIPL|0|000000","LeftRampCL|0|000000","LeftRampEL|100|00829B","HoleShotL|100|00D4FD","CrossbowShotL|100|00CEF6","GILeftSlingshotLower|0|000000","GIRightSlingshotLow|0|000000"), _
Array("ObjectiveFinalBossL|0|000000","ObjectiveChaserL|100|00D6FF","ObjectiveDragonL|100|00D6FF","StandupsNL|0|000000","StandupsIL|100|006E83","StandupsTL|100|00D6FF","DropTargetsCHL|100|00667A","CrossbowSNL|100|00161B","RightRampSL|100|00D6FF","LeftRampEL|0|000000","CrossbowPowerupL|0|000000","LeftRampPowerupL|100|00CCF3","HoleShotL|100|00D6FF","CrossbowShotL|100|00D6FF","LeftRampShotL|100|00252C","Light003|100|00C5EB","GIRightInlaneUpper|0|000000","GILeftInlaneUpper|100|006F84","GIRightSlingshotUpper|100|00D6FF","GILeftSlingshot|100|008BA5"), _
Array("ObjectiveFinalJudgmentL|100|000505","ObjectiveSniperL|100|00D6FF","ObjectiveBlacksmithL|100|00D1F9","ObjectiveChaserL|100|00D5FD","StandupsIL|0|000000","StandupsAL|0|000000","StandupsRL|100|000809","StandupsTL|100|00CEF6","DropTargetsHL|100|00D6FF","DropTargetsPlusL|100|006B7F","DropTargetsASL|100|00080A","DropTargetsCHL|100|00D6FF","CrossbowSNL|0|000000","RightRampAL|100|00D6FF","LeftRampPowerupL|0|000000","HoleShotL|100|00D1F9","LeftRampShotL|100|00D6FF","Light003|100|00D6FF","GIRightRubbersLower|100|00B8DB","GILeftInlaneUpper|0|000000","GILeftSlingshot|100|00D6FF"), _
Array("ObjectiveFinalJudgmentL|0|000000","ObjectiveBlacksmithL|100|00D6FF","ObjectiveChaserL|0|000000","ObjectiveDragonL|0|000000","ObjectiveCastleL|100|000C0E","CrossbowRightStandupL|100|00D6FF","StandupsRL|0|000000","StandupsTL|0|000000","DropTargetsPL|100|00D6FF","DropTargetsPlusL|100|00D6FF","DropTargetsASL|100|00D6FF","DropTargetsERL|100|000506","CrossbowERL|100|007389","RightRampEL|100|00D6FF","RightRampSL|100|00090B","LeftRampPL|100|00D6FF","HoleShotL|0|000000","CrossbowShotL|0|000000","GIRightRubbersLower|100|00D6FF","GIRightSlingshotUpper|0|000000"), _
Array("ObjectiveSniperL|0|000000","ObjectiveBlacksmithL|100|00B3D5","ObjectiveVikingL|100|00D2FB","ObjectiveEscapeL|100|000101","ObjectiveCastleL|100|00D6FF","CrossbowCenterStandupL|100|001D22","DropTargetsHL|100|00A2C2","DropTargetsERL|100|00D6FF","DropTargetsCHL|0|000000","LeftOrbitBumpersL|100|00BCE0","CrossbowERL|100|00D6FF","CrossbowIPL|100|008EA9","RightRampAL|0|000000","RightRampSL|0|000000","LeftRampCL|100|00D6FF","RightRampPowerupL|100|00D6FF","LeftRampShotL|0|000000","Light003|0|000000","Light002|100|00748A","Light001|100|00B0D2","GILeftSlingshot|0|000000"), _
Array("ObjectiveBlacksmithL|0|000000","ObjectiveVikingL|100|00D6FF","ObjectiveFinalBossL|100|00D6FF","ObjectiveEscapeL|100|00D6FF","CrossbowRightStandupL|0|000000","CrossbowCenterStandupL|100|00D6FF","StandupsNL|100|000D10","DropTargetsPL|0|000000","DropTargetsHL|0|000000","DropTargetsPlusL|0|000000","DropTargetsASL|0|000000","LeftHorseshoeKL|100|003D49","LeftOrbitBumpersL|100|00D6FF","CrossbowIPL|100|00D6FF","RightRampEL|100|007D95","LeftRampPL|100|0097B4","LeftRampEL|100|00D6FF","CrossbowPowerupL|100|000101","Light002|100|00D6FF","Light001|100|00D6FF","GIRightRubbersLower|0|000000"), _
Array("ObjectiveFinalJudgmentL|100|004A58","ObjectiveVikingL|100|00CEF6","ObjectiveCastleL|0|000000","CrossbowLeftStandupL|100|009BB9","StandupsNL|100|00D6FF","StandupsIL|100|00D6FF","StandupsAL|100|00B3D6","DropTargetsERL|0|000000","LeftHorseshoeKL|100|00D6FF","LeftHorseshoeCL|100|001A1F","LeftOrbitLockL|100|00D6FF","CrossbowERL|0|000000","CrossbowSNL|100|00D6FF","LeftOrbitGL|100|001519","RightRampEL|0|000000","LeftRampPL|0|000000","LeftRampCL|0|000000","RightRampPowerupL|0|000000","CrossbowPowerupL|100|00D6FF","LeftRampPowerupL|100|00D6FF","RightRampShotL|100|00D6FF","GICrossbowLeft|100|00D6FF","GIRightRubbersUpper|100|002830"), _
Array("ObjectiveFinalJudgmentL|100|00D6FF","ObjectiveVikingL|0|000000","ObjectiveFinalBossL|0|000000","ObjectiveEscapeL|0|000000","CrossbowCenterStandupL|0|000000","CrossbowLeftStandupL|100|00D6FF","StandupsAL|100|00D6FF","StandupsRL|100|00D6FF","StandupsTL|100|00D4FD","LeftHorseshoeCL|100|00D6FF","LeftHorseshoeAL|100|003640","LeftOrbitBumpersL|0|000000","CrossbowIPL|0|000000","LeftOrbitGL|100|00D6FF","LeftOrbitI2L|100|009CBA","RightRampSL|100|00090B","LeftRampEL|100|005363","HoleShotL|100|00D6FF","CrossbowShotL|100|00D6FF","Light002|0|000000","GILeftRubbersUpper|100|00D4FD","Light001|0|000000","GIRightRubbersUpper|100|00D6FF"), _
Array("StandupsNL|0|000000","StandupsIL|100|004653","StandupsTL|100|00D6FF","DropTargetsCHL|100|00D3FC","LeftHorseshoeKL|0|000000","LeftHorseshoeAL|100|00D6FF","LeftHorseshoeLL|100|00242B","ScoopLockL|100|00D6FF","LeftOrbitLockL|100|000709","CrossbowSNL|0|000000","LeftOrbitI2L|100|00D6FF","LeftOrbitI1L|100|001014","RightRampAL|100|000708","RightRampSL|100|00D6FF","LeftRampEL|0|000000","CrossbowPowerupL|0|000000","LeftRampPowerupL|100|00AFD1","RightRampShotL|100|006173","LeftRampShotL|100|00191E","GICrossbowLeft|100|004C5B","Light003|100|00D1FA","GIBumperInlanes4|100|00D6FF","GILeftRubbersUpper|100|00D6FF"), _
Array("ObjectiveFinalJudgmentL|0|000000","CrossbowLeftStandupL|0|000000","StandupsIL|0|000000","StandupsAL|0|000000","StandupsRL|0|000000","StandupsTL|100|00CEF5","DropTargetsPL|100|006073","DropTargetsHL|100|00D5FE","DropTargetsPlusL|100|00687B","DropTargetsASL|100|00819A","DropTargetsCHL|100|00D6FF","LeftHorseshoeCL|0|000000","LeftHorseshoeLL|100|00D6FF","LeftOrbitLockL|0|000000","ScoopDRL|100|00ABCC","LeftOrbitGL|0|000000","LeftOrbitI1L|100|00D6FF","RightRampAL|100|00D6FF","LeftOrbitPowerupL|100|00C4E9","LeftRampPowerupL|0|000000","HoleShotL|100|00191E","RightRampShotL|0|000000","CrossbowShotL|100|00D0F8","LeftRampShotL|100|00D6FF","GICrossbowLeft|0|000000","Light003|100|00D6FF","GILeftRubbersUpper|100|00CFF7","GIRightRubbersUpper|0|000000","GIRightRubbersLower|100|00D6FF"), _
Array("CrossbowRightStandupL|100|00D6FF","StandupsTL|0|000000","DropTargetsPL|100|00D6FF","DropTargetsHL|100|00D6FF","DropTargetsPlusL|100|00D6FF","DropTargetsASL|100|00D6FF","DropTargetsERL|100|007288","DropTargetsCHL|100|00A9C9","LeftHorseshoeAL|0|000000","LeftHorseshoeBL|100|00D6FF","ScoopLockL|100|00A3C3","ScoopDRL|100|00D6FF","ScoopAGL|100|00D6FF","CrossbowERL|100|005E70","LeftOrbitI2L|0|000000","RightRampEL|100|00D6FF","RightRampSL|0|000000","LeftRampPL|100|00D6FF","RightRampPowerupL|100|00272E","LeftOrbitPowerupL|100|00D6FF","HoleShotL|0|000000","CrossbowShotL|0|000000","Light003|100|00CCF3","GIBumperInlanes4|100|001C21","GILeftRubbersUpper|0|000000","GIRightRubbersLower|0|000000"), _
Array("CrossbowCenterStandupL|100|000B0D","DropTargetsHL|100|00B7DB","DropTargetsERL|100|00D6FF","DropTargetsCHL|0|000000","LeftHorseshoeLL|0|000000","LeftOrbitBumpersL|100|00BADD","ScoopLockL|0|000000","ScoopONL|100|00C5EB","CrossbowERL|100|00D6FF","CrossbowIPL|100|002128","LeftOrbitI1L|0|000000","RightRampAL|0|000000","LeftRampPL|0|000000","RightRampPowerupL|100|00D6FF","LeftHorseshoeShotL|100|00CEF6","LeftRampShotL|0|000000","Light003|0|000000","Light002|100|00282F","GIBumperInlanes4|0|000000"), _
Array("CrossbowRightStandupL|0|000000","CrossbowCenterStandupL|100|00D6FF","StandupsNL|100|00191E","DropTargetsPL|0|000000","DropTargetsHL|0|000000","DropTargetsPlusL|0|000000","DropTargetsASL|0|000000","LeftHorseshoeKL|100|00D5FE","LeftHorseshoeBL|0|000000","LeftOrbitBumpersL|100|00D6FF","ScoopONL|100|00D6FF","ScoopDRL|0|000000","ScoopAGL|100|00A2C1","CrossbowIPL|100|00D6FF","RightRampEL|0|000000","ScoopPowerupL|100|00B8DC","LeftOrbitPowerupL|0|000000","LeftHorseshoeShotL|100|00D6FF","LeftOrbitShotL|100|00D6FF","Light002|100|00D6FF","GIBumperInlanes3|100|00D6FF"), _
Array("CrossbowLeftStandupL|100|00B4D6","StandupsNL|100|00D6FF","StandupsIL|100|00D6FF","StandupsAL|100|00D1F9","DropTargetsERL|0|000000","RightHorseshoeHL|100|00D6FF","LeftHorseshoeKL|100|00D6FF","LeftHorseshoeCL|100|00B8DC","RightOrbitBumpersL|100|00323C","LeftOrbitLockL|100|00D6FF","ScoopONL|100|00D0F7","ScoopAGL|0|000000","CrossbowERL|0|000000","CrossbowSNL|100|00D6FF","LeftOrbitGL|100|00242A","ScoopPowerupL|100|00D6FF","RightRampPowerupL|0|000000","CrossbowPowerupL|100|00D6FF","RightRampShotL|100|00D6FF","LeftHorseshoeShotL|100|00D4FD","GIUpperFlipper|100|001C21","GICrossbowLeft|100|00D6FF","GILeftOrbit1|100|00D6FF"), _
Array("CrossbowCenterStandupL|0|000000","CrossbowLeftStandupL|100|00D6FF","StandupsAL|100|00D6FF","StandupsRL|100|00D6FF","StandupsTL|100|00D6FE","RightHorseshoeTL|100|00D6FF","LeftHorseshoeKL|100|00CEF6","LeftHorseshoeCL|100|00D6FF","LeftHorseshoeAL|100|00748A","RightOrbitBumpersL|100|00D6FF","LeftOrbitBumpersL|0|000000","ScoopONL|0|000000","CrossbowIPL|0|000000","RightOrbitVL|100|00A8C8","LeftOrbitGL|100|00D6FF","LeftOrbitI2L|100|00BDE1","RightRampSL|100|00080A","HoleShotL|100|00D6FF","LeftHorseshoeShotL|0|000000","CrossbowShotL|100|00D4FD","LeftOrbitShotL|0|000000","GIUpperFlipper|100|00D6FF","Light002|0|000000","GIBumperInlanes3|100|00748A"), _
Array("StandupsNL|0|000000","StandupsIL|100|000505","StandupsAL|100|00D5FE","StandupsTL|100|00D6FF","DropTargetsCHL|100|00D6FF","RightHorseshoeHL|100|001114","RightHorseshoeIL|100|00D6FF","LeftHorseshoeKL|0|000000","LeftHorseshoeCL|100|00D5FE","LeftHorseshoeAL|100|00D6FF","LeftHorseshoeLL|100|006C80","ScoopLockL|100|00D6FF","LeftOrbitLockL|0|000000","CrossbowSNL|0|000000","RightOrbitVL|100|00D6FF","RightOrbitKL|100|00D6FF","LeftOrbitI2L|100|00D6FF","LeftOrbitI1L|100|004C5A","RightRampAL|100|005160","RightRampSL|100|00D6FF","ScoopPowerupL|0|000000","CrossbowPowerupL|0|000000","ScoopShotL|100|00D6FF","RightRampShotL|100|00080A","CrossbowShotL|100|00D6FF","GIRightRamp|100|00D6FF","GICrossbowLeft|100|000202","Light003|100|00D6FF","GIVUK|100|007085","GIBumperInlanes4|100|00D6FF","GIBumperInlanes3|0|000000","GILeftOrbit1|0|000000"), _
Array("CrossbowLeftStandupL|0|000000","StandupsIL|0|000000","StandupsAL|0|000000","StandupsRL|0|000000","StandupsTL|100|00BDE1","DropTargetsPL|100|00D3FB","DropTargetsHL|100|00D6FF","DropTargetsPlusL|100|006477","DropTargetsASL|100|00D6FF","RightHorseshoeHL|0|000000","RightHorseshoeTL|0|000000","RightHorseshoeML|100|00D6FF","LeftHorseshoeCL|0|000000","LeftHorseshoeAL|100|00CEF6","LeftHorseshoeLL|100|00D6FF","LeftHorseshoeBL|100|0087A1","RightOrbitBumpersL|0|000000","ScoopDRL|100|00D6FF","RightOrbitVL|100|00CBF2","RightOrbitNL|100|00D2FA","LeftOrbitGL|0|000000","LeftOrbitI1L|100|00D6FF","RightRampEL|100|000202","RightRampAL|100|00D6FF","LeftOrbitPowerupL|100|00C8EF","HoleShotL|0|000000","RightRampShotL|0|000000","CrossbowShotL|100|00C6EC","GIUpperFlipper|0|000000","GICrossbowLeft|0|000000","GIVUK|100|00D6FF","GIBumperInlanes2|100|00C9F0","GILeftOrbit2|100|00D6FF"), _
Array("CrossbowRightStandupL|100|00D6FF","StandupsTL|0|000000","DropTargetsPL|100|00D6FF","DropTargetsPlusL|100|00D6FF","DropTargetsERL|100|00CBF2","DropTargetsCHL|100|00272F","RightHorseshoeIL|100|003E4A","RightHorseshoeSL|100|00D6FF","LeftHorseshoeAL|0|000000","LeftHorseshoeLL|100|00D3FB","LeftHorseshoeBL|100|00D6FF","ScoopLockL|0|000000","ScoopAGL|100|00D6FF","RightOrbitVL|0|000000","RightOrbitKL|100|005868","RightOrbitNL|100|00D6FF","LeftOrbitI2L|0|000000","RightRampEL|100|00D6FF","RightRampSL|0|000000","RightOrbitPowerupL|100|00D4FC","RightRampPowerupL|100|0095B1","LeftOrbitPowerupL|100|00D6FF","ScoopShotL|0|000000","CrossbowShotL|0|000000","GIRightRamp|0|000000","GIBumperInlane2|100|00282F","GIBumperInlanes4|0|000000","GIBumperInlanes2|100|00D6FF"), _
Array("CrossbowCenterStandupL|100|001114","DropTargetsPL|100|00CEF5","DropTargetsHL|100|00A5C5","DropTargetsASL|100|00B2D4","DropTargetsERL|100|00D6FF","DropTargetsCHL|0|000000","RightHorseshoeIL|0|000000","RightHorseshoeML|100|003F4B","LeftHorseshoeLL|0|000000","ScoopDRL|0|000000","ScoopAGL|0|000000","RightOrbitKL|0|000000","RightOrbitNL|100|00D2FA","LeftOrbitI1L|0|000000","RightRampAL|0|000000","RightOrbitPowerupL|100|00D6FF","RightRampPowerupL|100|00D6FF","LeftOrbitPowerupL|100|00D1F9","RightHorseshoeShotL|100|000506","LeftHorseshoeShotL|100|00D6FF","Light003|0|000000","Light002|100|003D49","GIBumperInlane2|100|00D6FF","GIVUK|0|000000","GILeftOrbit2|100|005666"), _
Array("CrossbowRightStandupL|0|000000","CrossbowCenterStandupL|100|00D6FF","DropTargetsPL|0|000000","DropTargetsHL|0|000000","DropTargetsPlusL|0|000000","DropTargetsASL|0|000000","DropTargetsERL|100|00B4D7","RightHorseshoeML|0|000000","RightHorseshoeSL|100|004451","LeftHorseshoeKL|100|00D6FF","LeftHorseshoeBL|0|000000","RightOrbitNL|0|000000","RightRampEL|0|000000","RightOrbitPowerupL|100|00C6EC","RightRampPowerupL|100|00D1F9","LeftOrbitPowerupL|0|000000","RightHorseshoeShotL|100|00D6FF","RightRampShotL|100|000101","LeftOrbitShotL|100|00D6FF","GIScoop2|100|00D6FF","Light002|100|00D6FF","GIBumperInlanes3|100|00D6FF","GIBumperInlanes2|0|000000","GILeftOrbit2|0|000000"), _
Array("CrossbowLeftStandupL|100|00B0D1","StandupsRL|100|005666","DropTargetsERL|0|000000","RightHorseshoeHL|100|00D6FF","RightHorseshoeTL|100|00080A","RightHorseshoeSL|0|000000","LeftHorseshoeCL|100|00D6FF","RightOrbitBumpersL|100|00D4FD","RightOrbitPowerupL|0|000000","RightRampPowerupL|0|000000","HoleShotL|100|00B6D9","RightOrbitShotL|100|00D6FF","RightRampShotL|100|00D6FF","LeftHorseshoeShotL|100|00343E","GIBumperInlane2|0|000000","GIBumperInlanes1|100|00AFD0","GILeftOrbit3|100|00D6FF"), _
Array("CrossbowCenterStandupL|0|000000","CrossbowLeftStandupL|0|000000","StandupsRL|0|000000","RightHorseshoeTL|100|00D6FF","LeftHorseshoeKL|100|00424F","LeftHorseshoeAL|100|00D6FF","RightOrbitBumpersL|100|00D6FF","RightOrbitVL|100|00D6FF","RightRampSL|100|00D6FF","HoleShotL|100|00D6FF","RightHorseshoeShotL|0|000000","LeftHorseshoeShotL|0|000000","LeftOrbitShotL|0|000000","GIScoop2|0|000000","GIRightRamp|100|000F12","Light002|0|000000","GIBumperInlanes3|100|002930","GIBumperInlanes1|100|00D6FF"), _
Array("SpinnerShotL|100|000506","DropTargetsCHL|100|00D6FF","RightHorseshoeHL|0|000000","RightHorseshoeIL|100|00D6FF","LeftHorseshoeKL|0|000000","LeftHorseshoeCL|100|005261","LeftHorseshoeLL|100|00D6FF","RightOrbitBumpersL|0|000000","RightOrbitKL|100|00D6FF","RightRampSL|0|000000","HoleShotL|0|000000","RightOrbitShotL|0|000000","RightRampShotL|0|000000","GIScoop1|100|00D6FF","GIRightRamp|100|00D6FF","GIVUK|100|00D6FF","GIBumperInlanes4|100|00D6FF","GIBumperInlanes3|0|000000","GILeftOrbit3|100|00C4E9"), _
Array("SpinnerShotL|100|00D6FF","DropTargetsPL|100|00D6FF","DropTargetsHL|100|00D6FF","DropTargetsPlusL|100|00D6FF","DropTargetsASL|100|00D6FF","RightHorseshoeTL|0|000000","RightHorseshoeML|100|00D6FF","LeftHorseshoeCL|0|000000","LeftHorseshoeAL|100|00778E","LeftHorseshoeBL|100|00D1F9","RightOrbitVL|0|000000","RightOrbitKL|0|000000","GIRightOrbit1|100|00D6FF","GIBumperInlanes2|100|00D6FF","GIBumperInlanes1|0|000000","GILeftOrbit3|0|000000"), _
Array("DropTargetsPlusL|0|000000","DropTargetsASL|0|000000","DropTargetsCHL|0|000000","RightHorseshoeIL|0|000000","RightHorseshoeSL|100|00D6FF","LeftHorseshoeAL|0|000000","LeftHorseshoeLL|100|00C3E9","LeftHorseshoeBL|100|00D6FF","GIScoop1|0|000000","GIRightRamp|0|000000","GIBumperInlane2|100|00BDE1","GIVUK|100|00B0D2","GIBumperInlanes4|0|000000"), _
Array("SpinnerShotL|0|000000","DropTargetsPL|0|000000","DropTargetsHL|0|000000","RightHorseshoeML|0|000000","LeftHorseshoeKL|100|00262D","LeftHorseshoeLL|0|000000","LeftHorseshoeBL|100|00A2C0","RightHorseshoeShotL|100|0091AD","LeftHorseshoeShotL|100|00D6FF","GIBumperInlane2|100|00D6FF","GIRightOrbit1|100|002A32","GIVUK|0|000000","GIBumperInlanes2|100|00D5FD"), _
Array("RightHorseshoeSL|0|000000","LeftHorseshoeKL|0|000000","LeftHorseshoeBL|0|000000","RightHorseshoeShotL|100|00D6FF","GIRightOrbit3|100|00D6FF","GIRightOrbit1|0|000000","GIBumperInlanes3|100|00D6FF","GIBumperInlanes2|0|000000"), _
Array("LeftHorseshoeShotL|0|000000","GIBumperInlane2|0|000000","GIRightOrbit2|100|001D23","GIBumperInlanes1|100|006376","GILeftOrbit4|100|0088A2"), _
Array("RightHorseshoeShotL|0|000000","GIRightOrbit3|100|006F84","GIRightOrbit2|100|00D6FF","GIBumperInlanes4|100|00849D","GIBumperInlanes3|100|001317","GIBumperInlanes1|100|00D6FF","GILeftOrbit4|100|00D6FF"), _
Array("GIRightOrbit3|0|000000","GIVUK|100|00D6FF","GIBumperInlanes4|100|00D6FF","GIBumperInlanes3|0|000000"), _
Array("GIRightOrbit2|0|000000","GIRightOrbit1|100|00D6FF","GIBumperInlanes2|100|00D6FF","GIBumperInlanes1|0|000000","GILeftOrbit4|0|000000"), _
Array("GIBumperInlane2|100|00B5D7","GIVUK|100|00191E","GIBumperInlanes4|0|000000","GIBumperInlanes2|0|000000"), _
Array("GIBumperInlane2|100|00D6FF","GIRightOrbit1|100|000101","GIVUK|0|000000"), _
Array("GIRightOrbit1|0|000000"))
lSeqpowerupDeployed.UpdateInterval = 20
lSeqpowerupDeployed.Color = Null
lSeqpowerupDeployed.Repeat = False


' *** END lights-out ***

'Glitch mini-wizard orbit shots
Dim lSeqGlitchWizard
Set lSeqGlitchWizard = New LCSeq
lSeqGlitchWizard.Name = "lSeqGlitchWizard"
lSeqGlitchWizard.Sequence = Array(Array("LeftOrbitBumpersL|100", "RightOrbitBumpersL|100"), _
Array("LeftOrbitBumpersL|0", "RightOrbitBumpersL|0", "LeftOrbitLockL|100"), _
Array("LeftOrbitLockL|0", "LeftOrbitI1L|100", "LeftOrbitI2L|100", "LeftOrbitGL|100", "RightOrbitVL|100", "RightOrbitKL|100", "RightOrbitNL|100"), _
Array("LeftOrbitI1L|0", "LeftOrbitI2L|0", "LeftOrbitGL|0", "RightOrbitVL|0", "RightOrbitKL|0", "RightOrbitNL|0", "LeftOrbitShotL|100", "RightOrbitShotL|100"))
lSeqGlitchWizard.UpdateInterval = 100
lSeqGlitchWizard.Color = Array(ddarkgreen, darkgreen)
lSeqGlitchWizard.Repeat = True

'Inlane GI reverse flipper indication
Dim lSeqReversedFlippers
Set lSeqReversedFlippers = New LCSeq
lSeqReversedFlippers.Name = "lSeqReversedFlippers"
lSeqReversedFlippers.Sequence = Array(Array("GIRightInlaneUpper|100", "GILeftInlaneUpper|100", "GIRightInlaneLower|0", "GILeftInlaneLower|0"), _
Array("GIRightInlaneUpper|0", "GILeftInlaneUpper|0", "GIRightInlaneLower|100", "GILeftInlaneLower|100"))
lSeqReversedFlippers.UpdateInterval = 200
lSeqReversedFlippers.Color = Array(damber, amber)
lSeqReversedFlippers.Repeat = True

'Blacksmith spare start
Dim lSeqblacksmith_spare : Set lSeqblacksmith_spare = New LCSeq
lSeqblacksmith_spare.Name = "lSeqblacksmith_spare"
lSeqblacksmith_spare.Sequence = Array( Array("StandupsTL|100|3D273D"), _
Array("StandupsTL|0|000000","DropTargetsPlusL|100|B977B9","DropTargetsCHL|100|6B456B","HoleShotL|100|C37DC3"), _
Array("StandupsRL|100|714871","StandupsTL|100|C17BC1","DropTargetsPlusL|100|C17BC1","DropTargetsCHL|100|C17BC1","HoleShotL|100|C17BC1"), _
Array("CrossbowRightStandupL|100|BE7ABE","StandupsRL|0|000000","StandupsTL|100|BF7ABF","DropTargetsHL|100|BF7ABF","DropTargetsPlusL|0|000000","DropTargetsASL|100|9F659F","DropTargetsCHL|100|160E16","HoleShotL|100|3C273C","Light002|100|040204"), _
Array("CrossbowRightStandupL|0|000000","CrossbowCenterStandupL|100|010001","StandupsAL|100|8A588A","StandupsTL|100|BD78BD","DropTargetsHL|100|BD78BD","DropTargetsASL|100|BD78BD","DropTargetsCHL|0|000000","RightRampSL|100|573857","HoleShotL|0|000000","Light003|100|B875B8","Light002|100|BD78BD"), _
Array("CrossbowCenterStandupL|100|BA77BA","StandupsAL|0|000000","StandupsRL|100|BA77BA","StandupsTL|100|B674B6","DropTargetsPL|100|BA77BA","DropTargetsHL|100|010101","DropTargetsASL|100|BA77BA","DropTargetsERL|100|BA77BA","RightRampSL|100|BA77BA","Light003|0|000000","Light002|100|B977B9"), _
Array("CrossbowRightStandupL|100|AE6FAE","CrossbowCenterStandupL|0|000000","StandupsIL|100|9A639A","StandupsRL|100|B876B8","StandupsTL|0|000000","DropTargetsPL|100|B876B8","DropTargetsHL|0|000000","DropTargetsASL|0|000000","DropTargetsERL|100|B876B8","RightRampAL|100|080508","RightRampSL|100|B876B8","RightRampPowerupL|100|5A3A5A","Light002|0|000000"), _
Array("CrossbowRightStandupL|100|B674B6","CrossbowLeftStandupL|100|B674B6","StandupsIL|0|000000","StandupsAL|100|694369","StandupsRL|100|B674B6","DropTargetsPL|100|9D649D","DropTargetsPlusL|100|3E273E","DropTargetsERL|100|B674B6","RightRampAL|100|B674B6","RightRampSL|100|B674B6","RightRampPowerupL|100|B674B6","Light003|100|0B070B"), _
Array("CrossbowRightStandupL|100|B473B4","CrossbowLeftStandupL|0|000000","StandupsNL|100|A469A4","StandupsAL|100|B473B4","StandupsRL|100|B473B4","DropTargetsPL|0|000000","DropTargetsPlusL|100|B473B4","DropTargetsERL|0|000000","DropTargetsCHL|100|030203","LeftHorseshoeKL|100|B473B4","RightRampEL|100|090609","RightRampAL|100|B473B4","RightRampSL|100|925D92","RightRampPowerupL|100|B473B4","RightRampShotL|100|B473B4","GICrossbowLeft|100|B473B4","Light003|100|B473B4"), _
Array("CrossbowRightStandupL|100|B271B2","CrossbowCenterStandupL|100|9A629A","StandupsNL|0|000000","StandupsAL|100|B271B2","StandupsRL|100|B271B2","DropTargetsPlusL|100|B271B2","DropTargetsCHL|100|B271B2","LeftHorseshoeKL|100|B271B2","RightRampEL|100|B271B2","RightRampAL|100|B271B2","RightRampSL|0|000000","RightRampPowerupL|100|B271B2","HoleShotL|100|3F283F","RightRampShotL|100|B271B2","GICrossbowLeft|100|3E283E","Light003|100|B271B2"), _
Array("CrossbowRightStandupL|100|AF70AF","CrossbowCenterStandupL|100|AF70AF","StandupsIL|100|AF70AF","StandupsAL|100|AF70AF","StandupsRL|100|050305","DropTargetsPlusL|100|AF70AF","DropTargetsCHL|100|AF70AF","LeftHorseshoeKL|100|AF70AF","LeftHorseshoeCL|100|AF70AF","LeftOrbitGL|100|AF70AF","RightRampEL|100|AF70AF","RightRampAL|100|AF70AF","RightRampPowerupL|100|AD6FAD","HoleShotL|100|AE6FAE","RightRampShotL|100|AF70AF","GICrossbowLeft|0|000000","Light003|100|AF70AF"), _
Array("CrossbowRightStandupL|100|AD6FAD","CrossbowCenterStandupL|100|AD6FAD","StandupsIL|100|AD6FAD","StandupsAL|100|AD6FAD","StandupsRL|0|000000","DropTargetsHL|100|432B43","DropTargetsPlusL|100|AD6FAD","DropTargetsCHL|100|AD6FAD","LeftHorseshoeKL|100|AC6FAC","LeftHorseshoeCL|100|AD6FAD","LeftHorseshoeAL|100|613E61","LeftOrbitGL|100|120C12","LeftOrbitI2L|100|412A41","RightRampEL|100|AD6FAD","RightRampAL|100|744B74","RightRampPowerupL|0|000000","LeftOrbitPowerupL|100|AD6FAD","HoleShotL|100|AD6FAD","RightRampShotL|100|AD6FAD","LeftOrbitShotL|100|734A73","Light003|100|AD6FAD","Light002|100|010001"), _
Array("CrossbowRightStandupL|100|AB6DAB","CrossbowCenterStandupL|100|AB6DAB","CrossbowLeftStandupL|100|A368A3","StandupsNL|100|885688","StandupsIL|100|AB6DAB","StandupsAL|100|AB6DAB","DropTargetsHL|100|AB6DAB","DropTargetsPlusL|100|AB6DAB","DropTargetsCHL|100|AB6DAB","LeftHorseshoeKL|0|000000","LeftHorseshoeCL|100|AB6DAB","LeftHorseshoeAL|100|AB6DAB","LeftOrbitLockL|100|150D15","LeftOrbitGL|0|000000","LeftOrbitI2L|100|AB6DAB","RightRampEL|100|AB6DAB","RightRampAL|0|000000","LeftOrbitPowerupL|100|AA6CAA","HoleShotL|100|AB6DAB","RightRampShotL|100|442C44","LeftOrbitShotL|100|AB6DAB","Light003|100|AB6DAB","Light002|100|AA6CAA"), _
Array("CrossbowRightStandupL|100|080508","CrossbowCenterStandupL|100|A96CA9","CrossbowLeftStandupL|100|A96CA9","StandupsNL|100|A96CA9","StandupsIL|100|A96CA9","StandupsAL|100|A76BA7","StandupsTL|100|040204","DropTargetsHL|100|A96CA9","DropTargetsPlusL|100|A96CA9","DropTargetsASL|100|895789","DropTargetsCHL|100|A96CA9","LeftHorseshoeCL|100|A96CA9","LeftHorseshoeAL|100|A96CA9","LeftHorseshoeLL|100|A96CA9","LeftOrbitBumpersL|100|010001","LeftOrbitLockL|100|A96CA9","LeftOrbitI2L|0|000000","LeftOrbitI1L|100|A96CA9","RightRampEL|100|A96CA9","LeftOrbitPowerupL|0|000000","HoleShotL|100|A96CA9","RightRampShotL|0|000000","CrossbowShotL|100|3A253A","LeftOrbitShotL|100|4E324E","GICrossbowLeft|100|080508","Light003|100|A96CA9","Light002|100|A96CA9"), _
Array("CrossbowRightStandupL|0|000000","CrossbowCenterStandupL|100|A76AA7","CrossbowLeftStandupL|100|A76AA7","StandupsNL|100|A76AA7","StandupsIL|100|A76AA7","StandupsAL|0|000000","StandupsTL|100|825382","DropTargetsHL|100|A76AA7","DropTargetsPlusL|100|A76AA7","DropTargetsASL|100|A76AA7","DropTargetsCHL|100|A76AA7","LeftHorseshoeCL|0|000000","LeftHorseshoeAL|100|A76AA7","LeftHorseshoeLL|100|A76AA7","LeftHorseshoeBL|100|010101","LeftOrbitBumpersL|100|A569A5","LeftOrbitLockL|0|000000","LeftOrbitI1L|100|965F96","RightRampEL|100|030203","HoleShotL|100|A76AA7","CrossbowShotL|0|000000","LeftOrbitShotL|0|000000","GICrossbowLeft|100|A76AA7","Light003|100|895789","Light002|100|A76AA7"), _
Array("CrossbowCenterStandupL|100|A469A4","CrossbowLeftStandupL|100|A469A4","StandupsNL|100|A469A4","StandupsIL|100|A469A4","StandupsTL|100|A469A4","DropTargetsPL|100|8F5C8F","DropTargetsHL|100|A469A4","DropTargetsPlusL|100|A469A4","DropTargetsASL|100|A469A4","DropTargetsCHL|100|A469A4","LeftHorseshoeAL|100|A469A4","LeftHorseshoeLL|100|A469A4","LeftHorseshoeBL|100|A469A4","LeftOrbitBumpersL|0|000000","LeftOrbitI1L|0|000000","RightRampEL|0|000000","HoleShotL|100|A469A4","GICrossbowLeft|100|A469A4","Light003|0|000000","Light002|100|A469A4"), _
Array("CrossbowCenterStandupL|100|A268A2","CrossbowLeftStandupL|100|A268A2","StandupsNL|100|A268A2","StandupsIL|100|A268A2","StandupsTL|100|A268A2","DropTargetsPL|100|A268A2","DropTargetsHL|100|A268A2","DropTargetsPlusL|100|A268A2","DropTargetsASL|100|A268A2","DropTargetsERL|100|6E476E","DropTargetsCHL|100|A268A2","RightHorseshoeHL|100|1D131D","LeftHorseshoeAL|0|000000","LeftHorseshoeLL|100|A268A2","LeftHorseshoeBL|100|A268A2","LeftOrbitGL|100|A268A2","HoleShotL|100|A268A2","GICrossbowLeft|100|A268A2","Light002|100|A268A2","GILeftOrbit2|100|7F517F"), _
Array("CrossbowCenterStandupL|100|9C639C","CrossbowLeftStandupL|100|A066A0","StandupsNL|100|A066A0","StandupsIL|100|311F31","StandupsTL|100|A066A0","DropTargetsPL|100|A066A0","DropTargetsHL|100|A066A0","DropTargetsPlusL|100|A066A0","DropTargetsASL|100|A066A0","DropTargetsERL|100|A066A0","DropTargetsCHL|100|A066A0","RightHorseshoeHL|100|A066A0","RightHorseshoeTL|100|2B1C2B","LeftHorseshoeLL|100|A066A0","LeftHorseshoeBL|100|A066A0","LeftOrbitGL|100|A066A0","HoleShotL|100|A066A0","LeftHorseshoeShotL|100|A066A0","GICrossbowLeft|100|A066A0","Light002|100|A066A0","GILeftOrbit2|100|A066A0","GILeftOrbit1|100|A066A0"), _
Array("CrossbowCenterStandupL|0|000000","CrossbowLeftStandupL|100|9E659E","StandupsNL|100|9E659E","StandupsIL|0|000000","StandupsTL|100|9E659E","DropTargetsPL|100|9E659E","DropTargetsHL|100|9E659E","DropTargetsPlusL|100|885788","DropTargetsASL|100|9E659E","DropTargetsERL|100|9E659E","DropTargetsCHL|100|9E659E","RightHorseshoeHL|100|9E659E","RightHorseshoeTL|100|9E659E","LeftHorseshoeLL|0|000000","LeftHorseshoeBL|100|9E659E","LeftOrbitGL|100|9E659E","LeftOrbitI2L|100|9E659E","RightRampSL|100|3D273D","LeftOrbitPowerupL|100|9E659E","HoleShotL|100|9E659E","LeftHorseshoeShotL|100|9E659E","CrossbowShotL|100|9E659E","GICrossbowLeft|100|9E659E","Light002|100|9E659E","GILeftOrbit2|100|996299","GILeftOrbit1|100|4C314C"), _
Array("CrossbowLeftStandupL|100|9C639C","StandupsNL|100|9C639C","StandupsTL|100|9C639C","DropTargetsPL|100|9C639C","DropTargetsHL|100|9C639C","DropTargetsPlusL|100|0D080D","DropTargetsASL|100|9C639C","DropTargetsERL|100|9C639C","DropTargetsCHL|100|9C639C","RightHorseshoeHL|100|9C639C","RightHorseshoeTL|100|9C639C","RightHorseshoeIL|100|9C639C","LeftHorseshoeBL|100|9C639C","LeftOrbitBumpersL|100|020102","LeftOrbitLockL|100|9C639C","RightOrbitVL|100|784C78","LeftOrbitGL|100|9C639C","LeftOrbitI2L|100|9C639C","RightRampSL|100|9C639C","LeftOrbitPowerupL|100|9C639C","HoleShotL|100|9C639C","LeftHorseshoeShotL|100|9C639C","CrossbowShotL|100|9C639C","LeftOrbitShotL|100|9C639C","GIRightRamp|100|9C639C","GICrossbowLeft|100|9C639C","Light002|100|9C639C","GIBumperInlanes4|100|9C639C","GIBumperInlanes3|100|9C639C","GILeftOrbit3|100|9C639C","GILeftOrbit2|0|000000","GILeftOrbit1|0|000000"), _
Array("CrossbowLeftStandupL|100|996299","StandupsNL|100|714871","StandupsTL|100|996299","DropTargetsPL|100|996299","DropTargetsHL|100|996299","DropTargetsPlusL|0|000000","DropTargetsASL|100|996299","DropTargetsERL|100|996299","DropTargetsCHL|100|976197","RightHorseshoeHL|100|996299","RightHorseshoeTL|100|996299","RightHorseshoeIL|100|996299","RightHorseshoeML|100|996299","LeftHorseshoeBL|0|000000","RightOrbitBumpersL|100|996299","LeftOrbitBumpersL|100|996299","LeftOrbitLockL|100|996299","RightOrbitVL|100|996299","LeftOrbitGL|100|996299","LeftOrbitI2L|100|996299","LeftOrbitI1L|100|996299","RightRampSL|100|996299","LeftOrbitPowerupL|100|996299","HoleShotL|100|996299","LeftHorseshoeShotL|100|996299","CrossbowShotL|100|996299","LeftOrbitShotL|100|996299","GIRightRamp|100|996299","GICrossbowLeft|100|996299","Light002|100|996299","GIBumperInlanes4|100|996299","GIBumperInlanes3|100|996299","GIBumperInlanes2|100|3B263B","GILeftOrbit3|100|996299"), _
Array("CrossbowLeftStandupL|100|976197","StandupsNL|0|000000","StandupsRL|100|6B456B","StandupsTL|100|976197","DropTargetsPL|100|976197","DropTargetsHL|100|976197","DropTargetsASL|100|976197","DropTargetsERL|100|976197","DropTargetsCHL|100|794D79","RightHorseshoeHL|100|976197","RightHorseshoeTL|100|976197","RightHorseshoeIL|100|976197","RightHorseshoeML|100|976197","RightHorseshoeSL|100|976197","LeftHorseshoeKL|100|462D46","RightOrbitBumpersL|100|976197","LeftOrbitBumpersL|100|976197","LeftOrbitLockL|100|976197","RightOrbitVL|100|976197","RightOrbitKL|100|976197","LeftOrbitGL|100|976197","LeftOrbitI2L|100|976197","LeftOrbitI1L|100|976197","RightRampSL|100|976197","RightRampPowerupL|100|160E16","LeftOrbitPowerupL|100|976197","HoleShotL|100|945F94","LeftHorseshoeShotL|100|976197","CrossbowShotL|100|976197","LeftOrbitShotL|100|976197","GIRightRamp|100|976197","GICrossbowLeft|100|976197","Light002|100|976197","GIBumperInlanes4|100|976197","GIBumperInlanes3|100|976197","GIBumperInlanes2|100|976197","GILeftOrbit3|100|976197"), _
Array("CrossbowLeftStandupL|0|000000","StandupsRL|100|955F95","StandupsTL|100|955F95","DropTargetsPL|100|955F95","DropTargetsHL|100|955F95","DropTargetsASL|100|955F95","DropTargetsERL|100|955F95","DropTargetsCHL|100|170F17","RightHorseshoeHL|100|8E5B8E","RightHorseshoeTL|100|955F95","RightHorseshoeIL|100|955F95","RightHorseshoeML|100|955F95","RightHorseshoeSL|100|955F95","LeftHorseshoeKL|100|955F95","RightOrbitBumpersL|100|955F95","LeftOrbitBumpersL|100|955F95","LeftOrbitLockL|100|955F95","RightOrbitVL|100|955F95","RightOrbitKL|100|955F95","LeftOrbitGL|100|955F95","LeftOrbitI2L|100|955F95","LeftOrbitI1L|100|955F95","RightRampAL|100|010101","RightRampSL|100|955F95","RightOrbitPowerupL|100|261826","RightRampPowerupL|100|955F95","CrossbowPowerupL|100|925D92","LeftOrbitPowerupL|100|955F95","HoleShotL|100|7A4E7A","RightHorseshoeShotL|100|020102","LeftHorseshoeShotL|100|955F95","CrossbowShotL|100|955F95","LeftOrbitShotL|100|955F95","GIRightRamp|100|955F95","GICrossbowLeft|100|955F95","Light002|100|955F95","GIBumperInlanes4|100|955F95","GIBumperInlanes3|100|955F95","GIBumperInlanes2|100|955F95","GILeftOrbit3|0|000000"), _
Array("StandupsRL|100|935E93","StandupsTL|100|935E93","DropTargetsPL|100|935E93","DropTargetsHL|100|935E93","DropTargetsASL|100|935E93","DropTargetsERL|100|935E93","DropTargetsCHL|0|000000","RightHorseshoeHL|0|000000","RightHorseshoeTL|100|935E93","RightHorseshoeIL|100|935E93","RightHorseshoeML|100|935E93","RightHorseshoeSL|100|935E93","LeftHorseshoeKL|100|935E93","LeftHorseshoeCL|100|010101","RightOrbitBumpersL|100|935E93","LeftOrbitBumpersL|100|935E93","LeftOrbitLockL|100|935E93","RightOrbitVL|100|935E93","RightOrbitKL|100|935E93","RightOrbitNL|100|935E93","LeftOrbitGL|100|935E93","LeftOrbitI2L|100|935E93","LeftOrbitI1L|100|935E93","RightRampAL|100|734A73","RightRampSL|100|935E93","RightOrbitPowerupL|100|935E93","RightRampPowerupL|100|935E93","CrossbowPowerupL|100|935E93","LeftOrbitPowerupL|100|935E93","HoleShotL|100|3B263B","RightHorseshoeShotL|100|935E93","RightRampShotL|100|442B44","LeftHorseshoeShotL|0|000000","CrossbowShotL|100|935E93","LeftOrbitShotL|100|935E93","GIRightRamp|100|935E93","GICrossbowLeft|100|372337","Light002|100|935E93","GIBumperInlanes4|100|935E93","GIBumperInlanes3|100|8F5C8F","GIBumperInlanes2|100|935E93","GIBumperInlanes1|100|935E93"), _
Array("StandupsRL|100|915C91","StandupsTL|100|915C91","DropTargetsPL|100|915C91","DropTargetsHL|100|8A578A","DropTargetsASL|100|915C91","DropTargetsERL|100|915C91","RightHorseshoeTL|0|000000","RightHorseshoeIL|100|915C91","RightHorseshoeML|100|915C91","RightHorseshoeSL|100|915C91","LeftHorseshoeKL|100|915C91","LeftHorseshoeCL|100|885688","RightOrbitBumpersL|100|915C91","LeftOrbitBumpersL|100|915C91","LeftOrbitLockL|100|915C91","ScoopDRL|100|915C91","CrossbowERL|100|2C1C2C","RightOrbitVL|100|915C91","RightOrbitKL|100|915C91","RightOrbitNL|100|915C91","LeftOrbitGL|100|915C91","LeftOrbitI2L|100|915C91","LeftOrbitI1L|100|915C91","RightRampAL|100|915C91","RightRampSL|100|915C91","RightOrbitPowerupL|100|915C91","RightRampPowerupL|100|915C91","CrossbowPowerupL|100|915C91","LeftOrbitPowerupL|100|915C91","HoleShotL|100|070407","RightOrbitShotL|100|482D48","RightHorseshoeShotL|100|915C91","RightRampShotL|100|915C91","CrossbowShotL|100|915C91","LeftOrbitShotL|100|915C91","GIRightRamp|100|915C91","GICrossbowLeft|0|000000","Light002|100|865586","GIBumperInlanes4|100|1B111B","GIBumperInlanes3|0|000000","GIBumperInlanes2|100|8F5B8F","GIBumperInlanes1|100|915C91","GILeftOrbit1|100|241724"), _
Array("CrossbowRightStandupL|100|160E16","StandupsRL|100|8E5B8E","StandupsTL|100|8E5B8E","DropTargetsPL|100|8E5B8E","DropTargetsHL|100|2C1C2C","DropTargetsASL|100|8E5B8E","DropTargetsERL|100|8E5B8E","RightHorseshoeIL|100|040304","RightHorseshoeML|100|8E5B8E","RightHorseshoeSL|100|8E5B8E","LeftHorseshoeKL|100|8E5B8E","LeftHorseshoeCL|100|8E5B8E","RightOrbitBumpersL|100|8E5B8E","LeftOrbitBumpersL|100|8E5B8E","LeftOrbitLockL|100|8E5B8E","ScoopDRL|100|8E5B8E","CrossbowERL|100|8E5B8E","CrossbowIPL|100|7D507D","RightOrbitVL|100|8E5B8E","RightOrbitKL|100|8E5B8E","RightOrbitNL|100|8E5B8E","LeftOrbitGL|100|8E5B8E","LeftOrbitI2L|100|8E5B8E","LeftOrbitI1L|100|8E5B8E","RightRampAL|100|8E5B8E","RightRampSL|100|8E5B8E","RightOrbitPowerupL|100|8E5B8E","RightRampPowerupL|100|8E5B8E","CrossbowPowerupL|100|8E5B8E","LeftOrbitPowerupL|100|8E5B8E","HoleShotL|0|000000","RightOrbitShotL|100|8E5B8E","RightHorseshoeShotL|100|8E5B8E","RightRampShotL|100|8E5B8E","CrossbowShotL|100|8E5B8E","LeftOrbitShotL|100|8E5B8E","GIRightRamp|100|170F17","Light002|100|140D14","GIBumperInlanes4|0|000000","GIBumperInlanes2|0|000000","GIBumperInlanes1|100|8E5B8E","GILeftOrbit2|100|8B598B","GILeftOrbit1|100|8E5B8E"), _
Array("CrossbowRightStandupL|100|825382","StandupsAL|100|0A060A","StandupsRL|100|8C5A8C","StandupsTL|100|8C5A8C","DropTargetsPL|100|8C5A8C","DropTargetsHL|0|000000","DropTargetsASL|100|8C5A8C","DropTargetsERL|100|8C5A8C","RightHorseshoeIL|0|000000","RightHorseshoeML|100|3E283E","RightHorseshoeSL|100|8C5A8C","LeftHorseshoeKL|100|8C5A8C","LeftHorseshoeCL|100|8C5A8C","LeftHorseshoeAL|100|261826","RightOrbitBumpersL|100|8C5A8C","LeftOrbitBumpersL|100|8C5A8C","ScoopLockL|100|140D14","LeftOrbitLockL|100|8C5A8C","ScoopDRL|100|8C5A8C","ScoopAGL|100|8C5A8C","CrossbowERL|100|8C5A8C","CrossbowSNL|100|8C5A8C","CrossbowIPL|100|8C5A8C","RightOrbitVL|100|2B1B2B","RightOrbitKL|100|8C5A8C","RightOrbitNL|100|8C5A8C","LeftOrbitGL|100|160E16","LeftOrbitI2L|100|8C5A8C","LeftOrbitI1L|100|8C5A8C","RightRampEL|100|090609","RightRampAL|100|8C5A8C","RightRampSL|100|8C5A8C","ScoopPowerupL|100|8C5A8C","RightOrbitPowerupL|100|8C5A8C","RightRampPowerupL|100|8C5A8C","CrossbowPowerupL|100|8C5A8C","LeftOrbitPowerupL|100|8C5A8C","ScoopShotL|100|8C5A8C","RightOrbitShotL|100|8C5A8C","RightHorseshoeShotL|100|8C5A8C","RightRampShotL|100|8C5A8C","CrossbowShotL|100|8C5A8C","LeftOrbitShotL|100|8C5A8C","GIRightRamp|0|000000","Light002|0|000000","GIBumperInlanes1|100|865686","GILeftOrbit2|100|8C5A8C","GILeftOrbit1|100|8C5A8C"), _
Array("SpinnerShotL|100|8A588A","CrossbowRightStandupL|100|8A588A","StandupsAL|100|704870","StandupsRL|100|8A588A","StandupsTL|100|895889","DropTargetsPL|100|8A588A","DropTargetsASL|100|8A588A","DropTargetsERL|100|8A588A","RightHorseshoeML|0|000000","RightHorseshoeSL|100|815281","LeftHorseshoeKL|100|8A588A","LeftHorseshoeCL|100|8A588A","LeftHorseshoeAL|100|8A588A","RightOrbitBumpersL|100|4F334F","LeftOrbitBumpersL|100|8A588A","ScoopLockL|100|8A588A","LeftOrbitLockL|100|8A588A","ScoopDRL|100|8A588A","ScoopAGL|100|8A588A","CrossbowERL|100|8A588A","CrossbowSNL|100|8A588A","CrossbowIPL|100|8A588A","RightOrbitVL|0|000000","RightOrbitKL|100|8A588A","RightOrbitNL|100|8A588A","LeftOrbitGL|0|000000","LeftOrbitI2L|100|8A588A","LeftOrbitI1L|100|8A588A","RightRampEL|100|7B4F7B","RightRampAL|100|8A588A","RightRampSL|100|8A588A","ScoopPowerupL|100|8A588A","RightOrbitPowerupL|100|8A588A","RightRampPowerupL|100|8A588A","CrossbowPowerupL|100|8A588A","LeftOrbitPowerupL|100|8A588A","ScoopShotL|100|8A588A","RightOrbitShotL|100|8A588A","RightHorseshoeShotL|100|8A588A","RightRampShotL|100|8A588A","CrossbowShotL|100|8A588A","LeftOrbitShotL|100|8A588A","GIScoop1|100|3D273D","Light003|100|020202","GIBumperInlane2|100|6A436A","GIBumperInlanes1|0|000000","GILeftOrbit2|100|8A588A","GILeftOrbit1|100|8A588A"), _
Array("SpinnerShotL|100|885788","CrossbowRightStandupL|100|885788","StandupsAL|100|885788","StandupsRL|100|885788","StandupsTL|100|684368","DropTargetsPL|100|885788","DropTargetsASL|100|5E3C5E","DropTargetsERL|100|885788","RightHorseshoeSL|0|000000","LeftHorseshoeKL|100|885788","LeftHorseshoeCL|100|885788","LeftHorseshoeAL|100|885788","RightOrbitBumpersL|0|000000","LeftOrbitBumpersL|100|885788","ScoopLockL|100|885788","LeftOrbitLockL|100|885788","ScoopONL|100|885788","ScoopDRL|100|885788","ScoopAGL|100|885788","CrossbowERL|100|885788","CrossbowSNL|100|885788","CrossbowIPL|100|885788","RightOrbitKL|100|4B304B","RightOrbitNL|100|885788","LeftOrbitI2L|100|704770","LeftOrbitI1L|100|885788","RightRampEL|100|885788","RightRampAL|100|885788","RightRampSL|100|885788","ScoopPowerupL|100|885788","RightOrbitPowerupL|100|885788","RightRampPowerupL|100|885788","CrossbowPowerupL|100|885788","LeftOrbitPowerupL|100|885788","ScoopShotL|100|885788","RightOrbitShotL|100|885788","RightHorseshoeShotL|100|885788","RightRampShotL|100|885788","CrossbowShotL|100|0A070A","LeftOrbitShotL|100|885788","GIScoop1|100|885788","Light003|100|573757","GIBumperInlane2|100|885788","GIVUK|100|855585","GILeftOrbit2|100|885788","GILeftOrbit1|100|885788"), _
Array("SpinnerShotL|100|865586","CrossbowRightStandupL|100|865586","StandupsAL|100|865586","StandupsRL|100|865586","StandupsTL|100|492E49","DropTargetsPL|100|805180","DropTargetsASL|100|0B070B","DropTargetsERL|100|865586","LeftHorseshoeKL|100|865586","LeftHorseshoeCL|100|865586","LeftHorseshoeAL|100|865586","LeftHorseshoeLL|100|7B4E7B","LeftOrbitBumpersL|100|865586","ScoopLockL|100|865586","LeftOrbitLockL|100|593859","ScoopONL|100|865586","ScoopDRL|100|865586","ScoopAGL|100|865586","CrossbowERL|100|865586","CrossbowSNL|100|865586","CrossbowIPL|100|865586","RightOrbitKL|0|000000","RightOrbitNL|100|865586","LeftOrbitI2L|0|000000","LeftOrbitI1L|100|865586","RightRampEL|100|865586","RightRampAL|100|865586","RightRampSL|100|865586","ScoopPowerupL|100|865586","RightOrbitPowerupL|100|865586","RightRampPowerupL|100|865586","CrossbowPowerupL|100|865586","LeftOrbitPowerupL|100|090509","ScoopShotL|100|865586","RightOrbitShotL|100|865586","RightHorseshoeShotL|100|845484","RightRampShotL|100|865586","CrossbowShotL|0|000000","LeftOrbitShotL|100|865586","GIScoop2|100|865586","GIScoop1|100|865586","Light003|100|855485","GIBumperInlane2|100|865586","GIVUK|100|865586","GILeftOrbit3|100|3B253B","GILeftOrbit2|100|865586","GILeftOrbit1|100|865586"), _
Array("SpinnerShotL|100|835483","CrossbowRightStandupL|100|835483","StandupsAL|100|835483","StandupsRL|100|835483","StandupsTL|100|191019","DropTargetsPL|100|3D273D","DropTargetsASL|0|000000","DropTargetsERL|100|835483","LeftHorseshoeKL|100|835483","LeftHorseshoeCL|100|835483","LeftHorseshoeAL|100|835483","LeftHorseshoeLL|100|835483","LeftOrbitBumpersL|100|684268","ScoopLockL|100|835483","LeftOrbitLockL|0|000000","ScoopONL|100|835483","ScoopDRL|100|835483","ScoopAGL|100|835483","CrossbowERL|100|835483","CrossbowSNL|100|835483","CrossbowIPL|100|835483","RightOrbitNL|100|5D3B5D","LeftOrbitI1L|100|835483","RightRampEL|100|835483","RightRampAL|100|835483","RightRampSL|100|835483","ScoopPowerupL|100|835483","RightOrbitPowerupL|0|000000","RightRampPowerupL|100|835483","CrossbowPowerupL|100|835483","LeftOrbitPowerupL|0|000000","ScoopShotL|100|835483","RightOrbitShotL|100|835483","RightHorseshoeShotL|0|000000","RightRampShotL|100|835483","LeftOrbitShotL|100|835483","GIScoop2|100|835483","GIScoop1|100|835483","Light003|100|835483","GIBumperInlane2|100|835483","GIVUK|100|835483","GILeftOrbit3|100|835483","GILeftOrbit2|100|835483","GILeftOrbit1|100|835483"), _
Array("SpinnerShotL|100|815281","CrossbowRightStandupL|100|815281","CrossbowCenterStandupL|100|010001","StandupsIL|100|020202","StandupsAL|100|815281","StandupsRL|100|815281","StandupsTL|100|030203","DropTargetsPL|0|000000","DropTargetsERL|100|7F517F","LeftHorseshoeKL|100|815281","LeftHorseshoeCL|100|815281","LeftHorseshoeAL|100|815281","LeftHorseshoeLL|100|815281","LeftHorseshoeBL|100|764B76","LeftOrbitBumpersL|0|000000","ScoopLockL|100|815281","ScoopONL|100|815281","ScoopDRL|100|815281","ScoopAGL|100|815281","CrossbowERL|100|815281","CrossbowSNL|100|815281","CrossbowIPL|100|815281","RightOrbitNL|0|000000","LeftOrbitI1L|100|1B111B","RightRampEL|100|815281","RightRampAL|100|815281","RightRampSL|100|815281","ScoopPowerupL|100|815281","RightRampPowerupL|100|815281","CrossbowPowerupL|100|815281","ScoopShotL|100|815281","RightOrbitShotL|100|815281","RightRampShotL|100|815281","LeftOrbitShotL|100|090609","GIScoop2|100|815281","GIScoop1|100|815281","Light003|100|815281","GIBumperInlane2|100|815281","GIRightOrbit1|100|120C12","GIVUK|100|815281","GILeftOrbit4|100|794D79","GILeftOrbit3|100|815281","GILeftOrbit2|100|815281","GILeftOrbit1|100|815281"), _
Array("SpinnerShotL|100|7F517F","CrossbowRightStandupL|100|7F517F","CrossbowCenterStandupL|100|2D1D2D","StandupsIL|100|422A42","StandupsAL|100|7F517F","StandupsRL|100|7F517F","StandupsTL|100|020102","DropTargetsERL|100|5A395A","LeftHorseshoeKL|100|7F517F","LeftHorseshoeCL|100|7F517F","LeftHorseshoeAL|100|7F517F","LeftHorseshoeLL|100|7F517F","LeftHorseshoeBL|100|7F517F","ScoopLockL|100|7F517F","ScoopONL|100|7F517F","ScoopDRL|100|7A4E7A","ScoopAGL|100|7F517F","CrossbowERL|100|7F517F","CrossbowSNL|100|7F517F","CrossbowIPL|100|7F517F","LeftOrbitI1L|0|000000","RightRampEL|100|7F517F","RightRampAL|100|7F517F","RightRampSL|100|7F517F","ScoopPowerupL|100|7F517F","RightRampPowerupL|100|7F517F","CrossbowPowerupL|100|7D507D","ScoopShotL|100|7F517F","RightOrbitShotL|100|050305","RightRampShotL|100|7F517F","LeftOrbitShotL|0|000000","LeftRampShotL|100|7E507E","GIScoop2|100|7F517F","GIUpperFlipper|100|030203","GIScoop1|100|7F517F","Light003|100|7F517F","GIBumperInlane2|0|000000","GIRightOrbit3|100|050305","GIRightOrbit1|100|7F517F","GIVUK|100|7F517F","GILeftOrbit4|100|7F517F","GILeftOrbit3|100|7F517F","GILeftOrbit2|100|7F517F","GILeftOrbit1|100|7F517F"), _
Array("SpinnerShotL|100|7D507D","CrossbowRightStandupL|100|7D507D","CrossbowCenterStandupL|100|714871","StandupsIL|100|7C507C","StandupsAL|100|7D507D","StandupsRL|100|7D507D","StandupsTL|0|000000","DropTargetsERL|100|0A060A","LeftHorseshoeKL|100|7D507D","LeftHorseshoeCL|100|7D507D","LeftHorseshoeAL|100|7D507D","LeftHorseshoeLL|100|7D507D","LeftHorseshoeBL|100|7D507D","ScoopLockL|100|7D507D","ScoopONL|100|7D507D","ScoopDRL|0|000000","ScoopAGL|100|7D507D","CrossbowERL|100|7D507D","CrossbowSNL|100|7D507D","CrossbowIPL|100|7D507D","RightRampEL|100|7D507D","RightRampAL|100|7D507D","RightRampSL|100|7D507D","ScoopPowerupL|100|7D507D","RightRampPowerupL|100|7D507D","CrossbowPowerupL|100|020202","ScoopShotL|100|7D507D","RightOrbitShotL|0|000000","RightRampShotL|100|7D507D","LeftRampShotL|100|7D507D","GIScoop2|100|7D507D","GIUpperFlipper|100|7D507D","GIScoop1|100|7D507D","Light003|100|7D507D","GIRightOrbit3|100|7D507D","GIRightOrbit2|100|794D79","GIRightOrbit1|100|7D507D","GIVUK|100|784D78","GIBumperInlanes3|100|0C080C","GILeftOrbit4|100|7D507D","GILeftOrbit3|100|7D507D","GILeftOrbit2|100|7D507D","GILeftOrbit1|100|7D507D"), _
Array("SpinnerShotL|100|7A4D7A","CrossbowRightStandupL|100|7B4E7B","CrossbowCenterStandupL|100|7B4E7B","StandupsIL|100|7B4E7B","StandupsAL|100|7B4E7B","StandupsRL|100|7B4E7B","DropTargetsERL|0|000000","LeftHorseshoeKL|100|7B4E7B","LeftHorseshoeCL|100|7B4E7B","LeftHorseshoeAL|100|7B4E7B","LeftHorseshoeLL|100|7B4E7B","LeftHorseshoeBL|100|7B4E7B","ScoopLockL|100|7B4E7B","ScoopONL|100|7B4E7B","ScoopAGL|100|7A4D7A","CrossbowERL|100|7B4E7B","CrossbowSNL|100|7B4E7B","CrossbowIPL|100|7B4E7B","RightRampEL|100|7B4E7B","RightRampAL|100|7B4E7B","RightRampSL|100|7B4E7B","ScoopPowerupL|100|3E273E","RightRampPowerupL|100|7B4E7B","CrossbowPowerupL|0|000000","LeftRampPowerupL|100|3D273D","ScoopShotL|100|191019","RightRampShotL|100|7B4E7B","LeftHorseshoeShotL|100|3E273E","LeftRampShotL|100|7B4E7B","GIScoop2|100|7B4E7B","GIUpperFlipper|100|7B4E7B","GIScoop1|100|7B4E7B","Light003|100|7B4E7B","GIRightOrbit3|100|7B4E7B","GIRightOrbit2|100|7B4E7B","GIRightOrbit1|100|7B4E7B","GIVUK|0|000000","GIBumperInlanes4|100|010001","GIBumperInlanes3|100|7B4E7B","GIBumperInlanes2|100|603D60","GILeftOrbit4|100|754A75","GILeftOrbit3|100|7B4E7B","GILeftOrbit2|100|7B4E7B","GILeftOrbit1|100|7B4E7B"), _
Array("SpinnerShotL|0|000000","CrossbowRightStandupL|100|784D78","CrossbowCenterStandupL|100|784D78","StandupsIL|100|784D78","StandupsAL|100|784D78","StandupsRL|100|784D78","DropTargetsPlusL|100|010101","LeftHorseshoeKL|100|784D78","LeftHorseshoeCL|100|784D78","LeftHorseshoeAL|100|784D78","LeftHorseshoeLL|100|784D78","LeftHorseshoeBL|100|784D78","ScoopLockL|100|6A446A","ScoopONL|100|784D78","ScoopAGL|100|050305","CrossbowERL|100|4D314D","CrossbowSNL|100|784D78","CrossbowIPL|100|784D78","RightRampEL|100|784D78","RightRampAL|100|784D78","RightRampSL|100|784D78","LeftRampPL|100|372337","ScoopPowerupL|0|000000","RightRampPowerupL|100|784D78","LeftRampPowerupL|100|784D78","ScoopShotL|0|000000","RightRampShotL|100|784D78","LeftHorseshoeShotL|100|784D78","LeftRampShotL|100|784D78","GIScoop2|100|784D78","GIUpperFlipper|100|784D78","GIScoop1|100|6E466E","Light003|100|784D78","GIRightOrbit3|100|784D78","GIRightOrbit2|100|784D78","GIRightOrbit1|100|784D78","GIBumperInlanes4|100|754B75","GIBumperInlanes3|100|784D78","GIBumperInlanes2|100|784D78","GIBumperInlanes1|100|060406","GILeftOrbit4|0|000000","GILeftOrbit3|100|784D78","GILeftOrbit2|100|784D78","GILeftOrbit1|100|784D78"), _
Array("CrossbowRightStandupL|100|764B76","CrossbowCenterStandupL|100|764B76","StandupsNL|100|0C080C","StandupsIL|100|764B76","StandupsAL|100|764B76","StandupsRL|100|764B76","DropTargetsPlusL|100|130C13","RightHorseshoeHL|100|523452","LeftHorseshoeKL|100|764B76","LeftHorseshoeCL|100|764B76","LeftHorseshoeAL|100|764B76","LeftHorseshoeLL|100|764B76","LeftHorseshoeBL|100|764B76","ScoopLockL|0|000000","ScoopONL|100|764B76","ScoopAGL|0|000000","CrossbowERL|0|000000","CrossbowSNL|100|764B76","CrossbowIPL|100|3F283F","RightRampEL|100|764B76","RightRampAL|100|764B76","RightRampSL|100|704770","LeftRampPL|100|764B76","LeftRampCL|100|060406","RightRampPowerupL|100|764B76","LeftRampPowerupL|100|764B76","RightRampShotL|100|764B76","LeftHorseshoeShotL|100|764B76","LeftRampShotL|100|764B76","GIScoop2|100|764B76","GIUpperFlipper|100|764B76","GIScoop1|0|000000","Light003|100|764B76","GIRightOrbit3|100|764B76","GIRightOrbit2|100|764B76","GIRightOrbit1|100|764B76","GIBumperInlanes4|100|764B76","GIBumperInlanes3|100|764B76","GIBumperInlanes2|100|764B76","GIBumperInlanes1|100|764B76","GILeftOrbit3|100|764B76","GILeftOrbit2|100|764B76","GILeftOrbit1|100|764B76","GILeftRubbersUpper|100|060406","GIRightRubbersUpper|100|714871"), _
Array("CrossbowRightStandupL|100|744A74","CrossbowCenterStandupL|100|744A74","CrossbowLeftStandupL|100|010101","StandupsNL|100|613E61","StandupsIL|100|744A74","StandupsAL|100|744A74","StandupsRL|100|744A74","DropTargetsPlusL|100|322032","RightHorseshoeHL|100|744A74","RightHorseshoeTL|100|412941","LeftHorseshoeKL|100|744A74","LeftHorseshoeCL|100|744A74","LeftHorseshoeAL|100|744A74","LeftHorseshoeLL|100|744A74","LeftHorseshoeBL|100|744A74","ScoopONL|100|0A070A","CrossbowSNL|100|2C1C2C","CrossbowIPL|0|000000","RightRampEL|100|744A74","RightRampAL|100|744A74","RightRampSL|100|5E3C5E","LeftRampPL|100|744A74","LeftRampCL|100|744A74","LeftRampEL|100|080508","RightRampPowerupL|100|744A74","LeftRampPowerupL|100|744A74","RightRampShotL|100|744A74","LeftHorseshoeShotL|100|744A74","LeftRampShotL|100|744A74","GIScoop2|100|714871","GIUpperFlipper|100|744A74","Light003|100|744A74","GIRightOrbit3|100|744A74","GIRightOrbit2|100|744A74","GIRightOrbit1|100|744A74","GIBumperInlanes4|100|744A74","GIBumperInlanes3|100|744A74","GIBumperInlanes2|100|744A74","GIBumperInlanes1|100|744A74","GILeftOrbit3|100|744A74","GILeftOrbit2|100|744A74","GILeftOrbit1|100|1A111A","GILeftRubbersUpper|100|744A74","GIRightRubbersUpper|100|744A74"), _
Array("CrossbowRightStandupL|100|724972","CrossbowCenterStandupL|100|724972","CrossbowLeftStandupL|100|281A28","StandupsNL|100|724972","StandupsIL|100|724972","StandupsAL|100|724972","StandupsRL|100|724972","DropTargetsPlusL|100|5D3C5D","RightHorseshoeHL|100|724972","RightHorseshoeTL|100|724972","RightHorseshoeIL|100|231723","LeftHorseshoeKL|100|442C44","LeftHorseshoeCL|100|724972","LeftHorseshoeAL|100|724972","LeftHorseshoeLL|100|724972","LeftHorseshoeBL|100|724972","ScoopONL|0|000000","CrossbowSNL|0|000000","RightRampEL|100|724972","RightRampAL|100|724972","RightRampSL|100|382438","LeftRampPL|100|724972","LeftRampCL|100|724972","LeftRampEL|100|724972","RightRampPowerupL|100|724972","LeftRampPowerupL|100|724972","RightRampShotL|100|724972","LeftHorseshoeShotL|100|724972","LeftRampShotL|100|724972","GIScoop2|100|010101","GIUpperFlipper|100|724972","GIRightRamp|100|6F476F","Light003|100|724972","GIRightOrbit3|100|724972","GIRightOrbit2|100|724972","GIRightOrbit1|100|010101","GIBumperInlanes4|100|724972","GIBumperInlanes3|100|724972","GIBumperInlanes2|100|724972","GIBumperInlanes1|100|724972","GILeftOrbit3|100|724972","GILeftOrbit2|100|2B1B2B","GILeftOrbit1|0|000000","GILeftRubbersUpper|100|724972","GIRightRubbersUpper|100|724972"), _
Array("ObjectiveFinalJudgmentL|100|6F466F","CrossbowRightStandupL|100|704770","CrossbowCenterStandupL|100|704770","CrossbowLeftStandupL|100|694269","StandupsNL|100|704770","StandupsIL|100|704770","StandupsAL|100|704770","StandupsRL|100|6D456D","DropTargetsPlusL|100|6A436A","RightHorseshoeHL|100|704770","RightHorseshoeTL|100|704770","RightHorseshoeIL|100|704770","RightHorseshoeML|100|010101","LeftHorseshoeKL|100|060406","LeftHorseshoeCL|100|704770","LeftHorseshoeAL|100|704770","LeftHorseshoeLL|100|704770","LeftHorseshoeBL|100|704770","RightRampEL|100|704770","RightRampAL|100|704770","RightRampSL|0|000000","LeftRampPL|100|704770","LeftRampCL|100|704770","LeftRampEL|100|704770","RightRampPowerupL|100|704770","LeftRampPowerupL|100|704770","RightRampShotL|100|704770","LeftHorseshoeShotL|100|704770","LeftRampShotL|100|704770","GIScoop2|0|000000","GIUpperFlipper|100|704770","GIRightRamp|100|704770","Light003|100|704770","GIRightOrbit3|100|593859","GIRightOrbit2|100|704770","GIRightOrbit1|0|000000","GIBumperInlanes4|100|704770","GIBumperInlanes3|100|704770","GIBumperInlanes2|100|704770","GIBumperInlanes1|100|704770","GILeftOrbit3|100|704770","GILeftOrbit2|0|000000","GILeftRubbersUpper|100|704770","GIRightRubbersUpper|100|704770"), _
Array("ObjectiveFinalJudgmentL|100|6D466D","CrossbowRightStandupL|100|6D466D","CrossbowCenterStandupL|100|6D466D","CrossbowLeftStandupL|100|6D466D","StandupsNL|100|6D466D","StandupsIL|100|6D466D","StandupsAL|100|6D466D","StandupsRL|100|613E61","DropTargetsPlusL|100|6D466D","RightHorseshoeHL|100|6D466D","RightHorseshoeTL|100|6D466D","RightHorseshoeIL|100|6D466D","RightHorseshoeML|100|6D466D","LeftHorseshoeKL|0|000000","LeftHorseshoeCL|100|6C456C","LeftHorseshoeAL|100|6D466D","LeftHorseshoeLL|100|6D466D","LeftHorseshoeBL|100|6D466D","RightOrbitVL|100|020102","RightRampEL|100|6D466D","RightRampAL|100|6D466D","LeftRampPL|100|6D466D","LeftRampCL|100|6D466D","LeftRampEL|100|6D466D","RightRampPowerupL|100|6D466D","LeftRampPowerupL|100|6D466D","RightRampShotL|100|6D466D","LeftHorseshoeShotL|100|6D466D","LeftRampShotL|100|6D466D","GIUpperFlipper|100|6D466D","GIRightRamp|100|6D466D","GICrossbowLeft|100|100A10","Light003|100|6D466D","GIRightOrbit3|0|000000","GIRightOrbit2|100|311F31","GIBumperInlanes4|100|6D466D","GIBumperInlanes3|100|6D466D","GIBumperInlanes2|100|6D466D","GIBumperInlanes1|100|6D466D","GILeftOrbit3|100|6D466D","GILeftRubbersUpper|100|6D466D","Light001|100|120B12","GIRightRubbersUpper|100|6D466D"), _
Array("ObjectiveFinalJudgmentL|100|6B446B","ObjectiveSniperL|100|593959","ObjectiveBlacksmithL|100|6B446B","CrossbowRightStandupL|100|6B446B","CrossbowCenterStandupL|100|6B446B","CrossbowLeftStandupL|100|6B446B","StandupsNL|100|6B446B","StandupsIL|100|6B446B","StandupsAL|100|6B446B","StandupsRL|100|3A253A","DropTargetsPlusL|100|6B446B","DropTargetsCHL|100|0A060A","RightHorseshoeHL|100|6B446B","RightHorseshoeTL|100|6B446B","RightHorseshoeIL|100|6B446B","RightHorseshoeML|100|6B446B","RightHorseshoeSL|100|603D60","LeftHorseshoeCL|100|372337","LeftHorseshoeAL|100|6B446B","LeftHorseshoeLL|100|6B446B","LeftHorseshoeBL|100|6B446B","RightOrbitVL|100|654065","RightRampEL|100|6B446B","RightRampAL|100|6B446B","LeftRampPL|100|6B446B","LeftRampCL|100|6B446B","LeftRampEL|100|6B446B","RightRampPowerupL|100|4B2F4B","LeftRampPowerupL|100|6B446B","RightRampShotL|100|6A436A","LeftHorseshoeShotL|100|6B446B","LeftRampShotL|100|6B446B","GIUpperFlipper|100|6B446B","GIRightRamp|100|6B446B","GICrossbowLeft|100|623E62","Light003|100|6B446B","GIRightOrbit2|0|000000","GIBumperInlanes4|100|6B446B","GIBumperInlanes3|100|6B446B","GIBumperInlanes2|100|6B446B","GIBumperInlanes1|100|6B446B","GILeftOrbit3|100|6B446B","GILeftRubbersUpper|100|6B446B","Light001|100|6B446B","GIRightRubbersUpper|100|6B446B","GIRightRubbersLower|100|402940"), _
Array("ObjectiveFinalJudgmentL|100|694369","ObjectiveSniperL|100|694369","ObjectiveBlacksmithL|100|694369","ObjectiveFinalBossL|100|070507","CrossbowRightStandupL|100|694369","CrossbowCenterStandupL|100|694369","CrossbowLeftStandupL|100|694369","StandupsNL|100|694369","StandupsIL|100|694369","StandupsAL|100|694369","StandupsRL|100|211521","DropTargetsPlusL|100|694369","DropTargetsCHL|100|150D15","RightHorseshoeHL|100|694369","RightHorseshoeTL|100|694369","RightHorseshoeIL|100|694369","RightHorseshoeML|100|694369","RightHorseshoeSL|100|694369","LeftHorseshoeCL|0|000000","LeftHorseshoeAL|100|694369","LeftHorseshoeLL|100|694369","LeftHorseshoeBL|100|694369","RightOrbitBumpersL|100|0F090F","RightOrbitVL|100|694369","RightRampEL|100|694369","RightRampAL|100|684368","LeftRampPL|100|694369","LeftRampCL|100|694369","LeftRampEL|100|694369","RightRampPowerupL|100|0F0A0F","LeftRampPowerupL|100|694369","RightRampShotL|100|492E49","LeftHorseshoeShotL|100|694369","LeftRampShotL|100|694369","GIUpperFlipper|100|281A28","GIRightRamp|100|694369","GICrossbowLeft|100|684268","Light003|100|694369","Light002|100|100A10","GIBumperInlane2|100|513451","GIBumperInlanes4|100|694369","GIBumperInlanes3|100|694369","GIBumperInlanes2|100|694369","GIBumperInlanes1|100|694369","GILeftOrbit3|100|694369","GILeftRubbersUpper|100|694369","Light001|100|694369","GIRightRubbersUpper|100|694369","GIRightRubbersLower|100|694369"), _
Array("ObjectiveFinalJudgmentL|100|674267","ObjectiveSniperL|100|674267","ObjectiveBlacksmithL|100|674267","ObjectiveVikingL|100|301E30","ObjectiveFinalBossL|100|664266","ObjectiveEscapeL|100|231623","CrossbowRightStandupL|100|674267","CrossbowCenterStandupL|100|674267","CrossbowLeftStandupL|100|674267","StandupsNL|100|674267","StandupsIL|100|674267","StandupsAL|100|674267","StandupsRL|100|0D080D","DropTargetsHL|100|050305","DropTargetsPlusL|100|674267","DropTargetsCHL|100|2F1E2F","RightHorseshoeHL|100|674267","RightHorseshoeTL|100|674267","RightHorseshoeIL|100|674267","RightHorseshoeML|100|674267","RightHorseshoeSL|100|674267","LeftHorseshoeAL|100|674267","LeftHorseshoeLL|100|674267","LeftHorseshoeBL|100|674267","RightOrbitBumpersL|100|674267","RightOrbitVL|100|674267","RightOrbitKL|100|2C1C2C","RightRampEL|100|674267","RightRampAL|100|603E60","LeftRampPL|100|674267","LeftRampCL|100|674267","LeftRampEL|100|674267","RightRampPowerupL|0|000000","LeftRampPowerupL|100|674267","RightHorseshoeShotL|100|563756","RightRampShotL|100|080508","LeftHorseshoeShotL|100|674267","LeftRampShotL|100|674267","GIUpperFlipper|0|000000","GIRightRamp|100|674267","GICrossbowLeft|100|674267","Light003|100|674267","Light002|100|3F283F","GIBumperInlane2|100|674267","GIBumperInlanes4|100|674267","GIBumperInlanes3|100|674267","GIBumperInlanes2|100|674267","GIBumperInlanes1|100|674267","GILeftOrbit3|100|674267","GILeftRubbersUpper|100|674267","Light001|100|674267","GIRightRubbersUpper|100|674267","GIRightRubbersLower|100|674267"), _
Array("ObjectiveFinalJudgmentL|100|654065","ObjectiveSniperL|100|654065","ObjectiveBlacksmithL|100|654065","ObjectiveVikingL|100|654065","ObjectiveFinalBossL|100|654065","ObjectiveEscapeL|100|654065","ObjectiveChaserL|100|0C080C","ObjectiveDragonL|100|030203","CrossbowRightStandupL|100|654065","CrossbowCenterStandupL|100|654065","CrossbowLeftStandupL|100|654065","StandupsNL|100|654065","StandupsIL|100|654065","StandupsAL|100|654065","StandupsRL|100|020102","DropTargetsHL|100|261826","DropTargetsPlusL|100|654065","DropTargetsCHL|100|482E48","RightHorseshoeHL|100|654065","RightHorseshoeTL|100|654065","RightHorseshoeIL|100|654065","RightHorseshoeML|100|654065","RightHorseshoeSL|100|654065","LeftHorseshoeAL|100|261826","LeftHorseshoeLL|100|654065","LeftHorseshoeBL|100|654065","RightOrbitBumpersL|100|654065","RightOrbitVL|100|654065","RightOrbitKL|100|654065","LeftOrbitGL|100|110B11","RightRampEL|100|654065","RightRampAL|100|3B253B","LeftRampPL|100|654065","LeftRampCL|100|654065","LeftRampEL|100|654065","RightOrbitPowerupL|100|1F141F","LeftRampPowerupL|100|654065","RightHorseshoeShotL|100|654065","RightRampShotL|0|000000","LeftHorseshoeShotL|100|654065","LeftRampShotL|100|3F283F","GIRightRamp|100|654065","GICrossbowLeft|100|654065","Light003|100|654065","Light002|100|5B3A5B","GIBumperInlane2|100|654065","GIBumperInlanes4|100|654065","GIBumperInlanes3|100|654065","GIBumperInlanes2|100|654065","GIBumperInlanes1|100|654065","GILeftOrbit4|100|583858","GILeftOrbit3|100|100A10","GILeftRubbersUpper|100|654065","Light001|100|654065","GIRightRubbersUpper|100|654065","GIRightRubbersLower|100|654065"), _
Array("ObjectiveFinalJudgmentL|100|623F62","ObjectiveSniperL|100|623F62","ObjectiveBlacksmithL|100|623F62","ObjectiveVikingL|100|623F62","ObjectiveFinalBossL|100|623F62","ObjectiveEscapeL|100|623F62","ObjectiveChaserL|100|623F62","ObjectiveDragonL|100|623F62","CrossbowRightStandupL|100|5E3D5E","CrossbowCenterStandupL|100|623F62","CrossbowLeftStandupL|100|623F62","StandupsNL|100|623F62","StandupsIL|100|623F62","StandupsAL|100|623F62","StandupsRL|0|000000","DropTargetsHL|100|4A2F4A","DropTargetsPlusL|100|623F62","DropTargetsCHL|100|573857","RightHorseshoeHL|100|623F62","RightHorseshoeTL|100|623F62","RightHorseshoeIL|100|623F62","RightHorseshoeML|100|623F62","RightHorseshoeSL|100|623F62","LeftHorseshoeAL|0|000000","LeftHorseshoeLL|100|623F62","LeftHorseshoeBL|100|623F62","RightOrbitBumpersL|100|623F62","RightOrbitVL|100|623F62","RightOrbitKL|100|623F62","RightOrbitNL|100|060406","LeftOrbitGL|100|4F334F","RightRampEL|100|623F62","RightRampAL|100|0D080D","LeftRampPL|100|623F62","LeftRampCL|100|623F62","LeftRampEL|100|623F62","RightOrbitPowerupL|100|623F62","LeftRampPowerupL|100|623F62","RightHorseshoeShotL|100|623F62","LeftHorseshoeShotL|100|623F62","LeftRampShotL|0|000000","GIRightRamp|100|623F62","GICrossbowLeft|100|623F62","Light003|100|623F62","Light002|100|623F62","GIBumperInlane2|100|623F62","GIVUK|100|4B304B","GIBumperInlanes4|100|623F62","GIBumperInlanes3|100|623F62","GIBumperInlanes2|100|623F62","GIBumperInlanes1|100|623F62","GILeftOrbit4|100|623F62","GILeftOrbit3|0|000000","GILeftRubbersUpper|100|623F62","Light001|100|623F62","GIRightRubbersUpper|100|623F62","GIRightRubbersLower|100|623F62"), _
Array("ObjectiveFinalJudgmentL|100|603D60","ObjectiveSniperL|100|603D60","ObjectiveBlacksmithL|100|603D60","ObjectiveVikingL|100|603D60","ObjectiveFinalBossL|100|603D60","ObjectiveEscapeL|100|603D60","ObjectiveChaserL|100|603D60","ObjectiveDragonL|100|603D60","ObjectiveCastleL|100|372337","CrossbowRightStandupL|100|4D314D","CrossbowCenterStandupL|100|603D60","CrossbowLeftStandupL|100|603D60","StandupsNL|100|603D60","StandupsIL|100|603D60","StandupsAL|100|603D60","DropTargetsHL|100|5C3A5C","DropTargetsPlusL|100|603D60","DropTargetsCHL|100|603D60","RightHorseshoeHL|100|603D60","RightHorseshoeTL|100|603D60","RightHorseshoeIL|100|603D60","RightHorseshoeML|100|603D60","RightHorseshoeSL|100|603D60","LeftHorseshoeLL|100|573757","LeftHorseshoeBL|100|603D60","RightOrbitBumpersL|100|603D60","RightOrbitVL|100|603D60","RightOrbitKL|100|603D60","RightOrbitNL|100|5D3B5D","LeftOrbitGL|100|603D60","RightRampEL|100|603D60","RightRampAL|100|020102","LeftRampPL|100|603D60","LeftRampCL|100|603D60","LeftRampEL|100|603D60","RightOrbitPowerupL|100|603D60","LeftRampPowerupL|100|603D60","HoleShotL|100|010101","RightOrbitShotL|100|352235","RightHorseshoeShotL|100|603D60","LeftHorseshoeShotL|100|603D60","GIRightRamp|100|603D60","GICrossbowLeft|100|603D60","Light003|100|603D60","Light002|100|603D60","GIBumperInlane2|100|603D60","GIVUK|100|603D60","GIBumperInlanes4|100|603D60","GIBumperInlanes3|100|603D60","GIBumperInlanes2|100|603D60","GIBumperInlanes1|100|603D60","GILeftOrbit4|100|603D60","GILeftRubbersUpper|100|603D60","Light001|100|603D60","GIRightRubbersUpper|100|362236","GIRightRubbersLower|100|603D60","GIRightSlingshotUpper|100|5F3C5F"), _
Array("ObjectiveFinalJudgmentL|100|5E3C5E","ObjectiveSniperL|100|5E3C5E","ObjectiveBlacksmithL|100|5E3C5E","ObjectiveVikingL|100|5E3C5E","ObjectiveFinalBossL|100|5E3C5E","ObjectiveEscapeL|100|5E3C5E","ObjectiveChaserL|100|5E3C5E","ObjectiveDragonL|100|5E3C5E","ObjectiveCastleL|100|5E3C5E","CrossbowRightStandupL|100|342134","CrossbowCenterStandupL|100|5E3C5E","CrossbowLeftStandupL|100|5E3C5E","StandupsNL|100|5E3C5E","StandupsIL|100|5E3C5E","StandupsAL|100|5E3C5E","DropTargetsHL|100|5E3C5E","DropTargetsPlusL|100|5E3C5E","DropTargetsCHL|100|5E3C5E","RightHorseshoeHL|100|5E3C5E","RightHorseshoeTL|100|5E3C5E","RightHorseshoeIL|100|5E3C5E","RightHorseshoeML|100|5E3C5E","RightHorseshoeSL|100|5E3C5E","LeftHorseshoeLL|100|080508","LeftHorseshoeBL|100|5E3C5E","RightOrbitBumpersL|100|5E3C5E","RightOrbitVL|100|5E3C5E","RightOrbitKL|100|5E3C5E","RightOrbitNL|100|5E3C5E","LeftOrbitGL|100|5E3C5E","LeftOrbitI2L|100|050305","RightRampEL|100|5E3C5E","RightRampAL|0|000000","LeftRampPL|100|593959","LeftRampCL|100|5E3C5E","LeftRampEL|100|5E3C5E","RightOrbitPowerupL|100|5E3C5E","LeftRampPowerupL|100|080508","HoleShotL|100|050305","RightOrbitShotL|100|5E3C5E","RightHorseshoeShotL|100|5E3C5E","LeftHorseshoeShotL|100|5E3C5E","CrossbowShotL|100|221622","GIRightRamp|100|5E3C5E","GICrossbowLeft|100|5E3C5E","Light003|100|5E3C5E","Light002|100|5E3C5E","GIBumperInlane2|100|5E3C5E","GIVUK|100|5E3C5E","GIBumperInlanes4|100|5E3C5E","GIBumperInlanes3|100|5E3C5E","GIBumperInlanes2|100|5E3C5E","GIBumperInlanes1|100|5E3C5E","GILeftOrbit4|100|5E3C5E","GILeftRubbersUpper|100|5E3C5E","Light001|100|5E3C5E","GIRightRubbersUpper|0|000000","GIRightRubbersLower|100|5E3C5E","GIRightSlingshotUpper|100|5E3C5E","GILeftSlingshot|100|583858"), _
Array("SpinnerShotL|100|130C13","ObjectiveFinalJudgmentL|100|5C3B5C","ObjectiveSniperL|100|5C3B5C","ObjectiveBlacksmithL|100|5C3B5C","ObjectiveVikingL|100|5C3B5C","ObjectiveFinalBossL|100|5C3B5C","ObjectiveEscapeL|100|5C3B5C","ObjectiveChaserL|100|5C3B5C","ObjectiveDragonL|100|5C3B5C","ObjectiveCastleL|100|5C3B5C","CrossbowRightStandupL|100|1D131D","CrossbowCenterStandupL|100|5C3B5C","CrossbowLeftStandupL|100|5C3B5C","StandupsNL|100|5C3B5C","StandupsIL|100|5C3B5C","StandupsAL|100|5A395A","DropTargetsPL|100|010101","DropTargetsHL|100|5C3B5C","DropTargetsPlusL|100|5C3B5C","DropTargetsCHL|100|5C3B5C","RightHorseshoeHL|100|5C3B5C","RightHorseshoeTL|100|5C3B5C","RightHorseshoeIL|100|5C3B5C","RightHorseshoeML|100|5C3B5C","RightHorseshoeSL|100|5C3B5C","LeftHorseshoeLL|0|000000","LeftHorseshoeBL|100|593959","RightOrbitBumpersL|100|5C3B5C","LeftOrbitLockL|100|020102","RightOrbitVL|100|5C3B5C","RightOrbitKL|100|5C3B5C","RightOrbitNL|100|5C3B5C","LeftOrbitGL|100|5C3B5C","LeftOrbitI2L|100|362236","RightRampEL|100|452C45","LeftRampPL|100|020102","LeftRampCL|100|5C3B5C","LeftRampEL|100|5C3B5C","RightOrbitPowerupL|100|5C3B5C","LeftOrbitPowerupL|100|251825","LeftRampPowerupL|0|000000","HoleShotL|100|0E090E","RightOrbitShotL|100|5C3B5C","RightHorseshoeShotL|100|5C3B5C","LeftHorseshoeShotL|100|5C3B5C","CrossbowShotL|100|593959","GIRightRamp|100|5C3B5C","GICrossbowLeft|100|5C3B5C","Light003|100|5C3B5C","Light002|100|5C3B5C","GIBumperInlane2|100|5C3B5C","GIVUK|100|5C3B5C","GIBumperInlanes4|100|5C3B5C","GIBumperInlanes3|100|5C3B5C","GIBumperInlanes2|100|5C3B5C","GIBumperInlanes1|100|5C3B5C","GILeftOrbit4|100|5C3B5C","GILeftRubbersUpper|100|5C3B5C","Light001|100|5C3B5C","GIRightRubbersLower|100|5C3B5C","GIRightSlingshotUpper|100|5C3B5C","GILeftSlingshot|100|5C3B5C"), _
Array("SpinnerShotL|100|5A395A","ObjectiveFinalJudgmentL|100|5A395A","ObjectiveSniperL|100|5A395A","ObjectiveBlacksmithL|100|5A395A","ObjectiveVikingL|100|5A395A","ObjectiveFinalBossL|100|5A395A","ObjectiveEscapeL|100|5A395A","ObjectiveChaserL|100|5A395A","ObjectiveDragonL|100|5A395A","ObjectiveCastleL|100|5A395A","CrossbowRightStandupL|100|0A070A","CrossbowCenterStandupL|100|5A395A","CrossbowLeftStandupL|100|5A395A","StandupsNL|100|5A395A","StandupsIL|100|5A395A","StandupsAL|100|4F324F","DropTargetsPL|100|090609","DropTargetsHL|100|5A395A","DropTargetsPlusL|100|5A395A","DropTargetsASL|100|140C14","DropTargetsCHL|100|5A395A","RightHorseshoeHL|100|5A395A","RightHorseshoeTL|100|5A395A","RightHorseshoeIL|100|5A395A","RightHorseshoeML|100|5A395A","RightHorseshoeSL|100|5A395A","LeftHorseshoeBL|100|150E15","RightOrbitBumpersL|100|5A395A","LeftOrbitLockL|100|2A1B2A","ScoopDRL|100|070407","RightOrbitVL|100|5A395A","RightOrbitKL|100|5A395A","RightOrbitNL|100|5A395A","LeftOrbitGL|100|5A395A","LeftOrbitI2L|100|593859","RightRampEL|100|160E16","LeftRampPL|0|000000","LeftRampCL|100|211521","LeftRampEL|100|5A395A","RightOrbitPowerupL|100|5A395A","LeftOrbitPowerupL|100|523452","HoleShotL|100|1C121C","RightOrbitShotL|100|5A395A","RightHorseshoeShotL|100|5A395A","LeftHorseshoeShotL|100|5A395A","CrossbowShotL|100|5A395A","GIRightRamp|100|5A395A","GICrossbowLeft|100|5A395A","Light003|100|5A395A","Light002|100|5A395A","GIBumperInlane2|100|5A395A","GIVUK|100|5A395A","GIBumperInlanes4|100|5A395A","GIBumperInlanes3|100|5A395A","GIBumperInlanes2|100|5A395A","GIBumperInlanes1|100|5A395A","GILeftOrbit4|100|5A395A","GILeftRubbersUpper|100|583858","Light001|100|5A395A","GIRightRubbersLower|100|5A395A","GIRightSlingshotUpper|100|5A395A","GILeftSlingshot|100|5A395A"), _
Array("SpinnerShotL|100|573857","ObjectiveFinalJudgmentL|100|4B304B","ObjectiveSniperL|100|573857","ObjectiveBlacksmithL|100|573857","ObjectiveVikingL|100|573857","ObjectiveFinalBossL|100|573857","ObjectiveEscapeL|100|573857","ObjectiveChaserL|100|573857","ObjectiveDragonL|100|573857","ObjectiveCastleL|100|573857","CrossbowRightStandupL|0|000000","CrossbowCenterStandupL|100|573857","CrossbowLeftStandupL|100|573857","StandupsNL|100|573857","StandupsIL|100|573857","StandupsAL|100|3B263B","DropTargetsPL|100|281A28","DropTargetsHL|100|573857","DropTargetsPlusL|100|573857","DropTargetsASL|100|332133","DropTargetsCHL|100|573857","RightHorseshoeHL|100|573857","RightHorseshoeTL|100|573857","RightHorseshoeIL|100|573857","RightHorseshoeML|100|573857","RightHorseshoeSL|100|573857","LeftHorseshoeBL|0|000000","RightOrbitBumpersL|100|573857","LeftOrbitBumpersL|100|140D14","LeftOrbitLockL|100|533553","ScoopDRL|100|503350","RightOrbitVL|100|573857","RightOrbitKL|100|573857","RightOrbitNL|100|573857","LeftOrbitGL|100|573857","LeftOrbitI2L|100|573857","RightRampEL|100|020102","LeftRampCL|0|000000","LeftRampEL|100|422A42","ScoopPowerupL|100|010101","RightOrbitPowerupL|100|573857","LeftOrbitPowerupL|100|573857","HoleShotL|100|2D1D2D","ScoopShotL|100|2A1B2A","RightOrbitShotL|100|573857","RightHorseshoeShotL|100|573857","LeftHorseshoeShotL|100|573857","CrossbowShotL|100|573857","LeftOrbitShotL|100|070507","GIScoop1|100|342134","GIRightRamp|100|573857","GICrossbowLeft|100|573857","Light003|100|553755","Light002|100|573857","GIBumperInlane2|100|573857","GIRightOrbit1|100|3C273C","GIVUK|100|573857","GIBumperInlanes4|100|573857","GIBumperInlanes3|100|3C273C","GIBumperInlanes2|100|503450","GIBumperInlanes1|100|573857","GILeftOrbit4|100|573857","GILeftRubbersUpper|100|0D090D","Light001|100|573857","GIRightRubbersLower|100|573857","GIRightSlingshotUpper|100|573857","GILeftSlingshot|100|573857","GIRightSlingshotLow|100|1C121C"), _
Array("SpinnerShotL|100|553655","ObjectiveFinalJudgmentL|0|000000","ObjectiveSniperL|100|553655","ObjectiveBlacksmithL|100|553655","ObjectiveVikingL|100|553655","ObjectiveFinalBossL|100|553655","ObjectiveEscapeL|100|553655","ObjectiveChaserL|100|553655","ObjectiveDragonL|100|553655","ObjectiveCastleL|100|553655","CrossbowCenterStandupL|100|553655","CrossbowLeftStandupL|100|553655","StandupsNL|100|553655","StandupsIL|100|553655","StandupsAL|100|211521","DropTargetsPL|100|4D314D","DropTargetsHL|100|553655","DropTargetsPlusL|100|553655","DropTargetsASL|100|492F49","DropTargetsCHL|100|553655","RightHorseshoeHL|100|553655","RightHorseshoeTL|100|553655","RightHorseshoeIL|100|553655","RightHorseshoeML|100|553655","RightHorseshoeSL|100|553655","RightOrbitBumpersL|100|553655","LeftOrbitBumpersL|100|4B304B","LeftOrbitLockL|100|553655","ScoopDRL|100|553655","ScoopAGL|100|010101","RightOrbitVL|100|553655","RightOrbitKL|100|553655","RightOrbitNL|100|553655","LeftOrbitGL|100|553655","LeftOrbitI2L|100|553655","LeftOrbitI1L|100|130C13","RightRampEL|0|000000","LeftRampEL|0|000000","ScoopPowerupL|100|452C45","RightOrbitPowerupL|100|553655","LeftOrbitPowerupL|100|553655","HoleShotL|100|3C263C","ScoopShotL|100|553655","RightOrbitShotL|100|553655","RightHorseshoeShotL|100|553655","LeftHorseshoeShotL|100|553655","CrossbowShotL|100|553655","LeftOrbitShotL|100|362236","GIScoop1|100|553655","GIRightRamp|100|553655","GICrossbowLeft|100|553655","Light003|100|442B44","Light002|100|553655","GIBumperInlane2|100|553655","GIRightOrbit1|100|553655","GIVUK|100|553655","GIBumperInlanes4|100|553655","GIBumperInlanes3|100|010001","GIBumperInlanes2|100|060406","GIBumperInlanes1|100|553655","GILeftOrbit4|100|553655","GILeftRubbersUpper|0|000000","Light001|100|553655","GIRightRubbersLower|100|553655","GILeftSlingshotLower|100|120B12","GIRightSlingshotUpper|100|553655","GILeftSlingshot|100|553655","GIRightSlingshotLow|100|553655"), _
Array("SpinnerShotL|100|533553","ObjectiveSniperL|100|533553","ObjectiveBlacksmithL|100|533553","ObjectiveVikingL|100|533553","ObjectiveFinalBossL|100|533553","ObjectiveEscapeL|100|533553","ObjectiveChaserL|100|533553","ObjectiveDragonL|100|533553","ObjectiveCastleL|100|533553","CrossbowCenterStandupL|100|533553","CrossbowLeftStandupL|100|533553","StandupsNL|100|533553","StandupsIL|100|533553","StandupsAL|100|0F090F","DropTargetsPL|100|533553","DropTargetsHL|100|533553","DropTargetsPlusL|100|533553","DropTargetsASL|100|523552","DropTargetsERL|100|040304","DropTargetsCHL|100|533553","RightHorseshoeHL|100|533553","RightHorseshoeTL|100|533553","RightHorseshoeIL|100|533553","RightHorseshoeML|100|533553","RightHorseshoeSL|100|533553","RightOrbitBumpersL|100|533553","LeftOrbitBumpersL|100|533553","ScoopLockL|100|010101","LeftOrbitLockL|100|533553","ScoopDRL|100|533553","ScoopAGL|100|392539","RightOrbitVL|100|533553","RightOrbitKL|100|533553","RightOrbitNL|100|533553","LeftOrbitGL|100|533553","LeftOrbitI2L|100|533553","LeftOrbitI1L|100|4E324E","ScoopPowerupL|100|533553","RightOrbitPowerupL|100|533553","LeftOrbitPowerupL|100|533553","HoleShotL|100|472D47","ScoopShotL|100|533553","RightOrbitShotL|100|533553","RightHorseshoeShotL|100|533553","LeftHorseshoeShotL|100|332133","CrossbowShotL|100|533553","LeftOrbitShotL|100|4F334F","GIScoop1|100|533553","GIRightRamp|100|533553","GICrossbowLeft|100|533553","Light003|100|291A29","Light002|100|533553","GIBumperInlane2|100|533553","GIRightOrbit3|100|322032","GIRightOrbit1|100|533553","GIVUK|100|533553","GIBumperInlanes4|100|2E1E2E","GIBumperInlanes3|0|000000","GIBumperInlanes2|0|000000","GIBumperInlanes1|100|221622","GILeftOrbit4|100|533553","Light001|100|533553","GIRightRubbersLower|100|2F1E2F","GILeftSlingshotLower|100|533553","GIRightSlingshotUpper|100|533553","GILeftSlingshot|100|533553","GIRightSlingshotLow|100|533553"), _
Array("RightOutlanePoisonL|100|070407","SpinnerShotL|100|513451","ObjectiveSniperL|100|030203","ObjectiveBlacksmithL|100|130C13","ObjectiveVikingL|100|513451","ObjectiveFinalBossL|100|513451","ObjectiveEscapeL|100|513451","ObjectiveChaserL|100|513451","ObjectiveDragonL|100|513451","ObjectiveCastleL|100|513451","CrossbowCenterStandupL|100|513451","CrossbowLeftStandupL|100|513451","StandupsNL|100|513451","StandupsIL|100|513451","StandupsAL|0|000000","DropTargetsPL|100|513451","DropTargetsHL|100|513451","DropTargetsPlusL|100|4E324E","DropTargetsASL|100|513451","DropTargetsERL|100|191019","DropTargetsCHL|100|513451","RightHorseshoeHL|100|513451","RightHorseshoeTL|100|513451","RightHorseshoeIL|100|513451","RightHorseshoeML|100|513451","RightHorseshoeSL|100|513451","RightOrbitBumpersL|100|513451","LeftOrbitBumpersL|100|513451","ScoopLockL|100|392439","LeftOrbitLockL|100|513451","ScoopONL|100|020102","ScoopDRL|100|513451","ScoopAGL|100|513451","RightOrbitVL|100|513451","RightOrbitKL|100|513451","RightOrbitNL|100|513451","LeftOrbitGL|100|513451","LeftOrbitI2L|100|513451","LeftOrbitI1L|100|513451","ScoopPowerupL|100|513451","RightOrbitPowerupL|100|513451","CrossbowPowerupL|100|2B1B2B","LeftOrbitPowerupL|100|513451","HoleShotL|100|4F334F","ScoopShotL|100|513451","RightOrbitShotL|100|513451","RightHorseshoeShotL|100|513451","LeftHorseshoeShotL|100|040304","CrossbowShotL|100|513451","LeftOrbitShotL|100|513451","GIScoop2|100|140D14","GIScoop1|100|513451","GIRightRamp|100|513451","GICrossbowLeft|100|513451","Light003|100|140D14","Light002|100|513451","GIBumperInlane2|100|513451","GIRightOrbit3|100|513451","GIRightOrbit2|100|1C121C","GIRightOrbit1|100|513451","GIVUK|100|513451","GIBumperInlanes4|0|000000","GIBumperInlanes1|0|000000","GILeftOrbit4|100|513451","Light001|100|513451","GIRightRubbersLower|0|000000","GILeftSlingshotLower|100|513451","GIRightSlingshotUpper|100|513451","GILeftSlingshot|100|513451","GIRightSlingshotLow|100|513451"), _
Array("RightOutlanePoisonL|100|4F324F","SpinnerShotL|100|4F324F","ObjectiveSniperL|0|000000","ObjectiveBlacksmithL|0|000000","ObjectiveVikingL|100|4F324F","ObjectiveFinalBossL|100|412941","ObjectiveEscapeL|100|4F324F","ObjectiveChaserL|100|4F324F","ObjectiveDragonL|100|4F324F","ObjectiveCastleL|100|4F324F","CrossbowCenterStandupL|100|4F324F","CrossbowLeftStandupL|100|4F324F","StandupsNL|100|4F324F","StandupsIL|100|4F324F","DropTargetsPL|100|4F324F","DropTargetsHL|100|4F324F","DropTargetsPlusL|100|3F283F","DropTargetsASL|100|4F324F","DropTargetsERL|100|3C263C","DropTargetsCHL|100|4F324F","RightHorseshoeHL|100|4F324F","RightHorseshoeTL|100|4F324F","RightHorseshoeIL|100|4F324F","RightHorseshoeML|100|4F324F","RightHorseshoeSL|100|4F324F","RightOrbitBumpersL|100|4F324F","LeftOrbitBumpersL|100|4F324F","ScoopLockL|100|4F324F","LeftOrbitLockL|100|4F324F","ScoopONL|100|402940","ScoopDRL|100|4F324F","ScoopAGL|100|4F324F","RightOrbitVL|100|4F324F","RightOrbitKL|100|4F324F","RightOrbitNL|100|4F324F","LeftOrbitGL|100|4F324F","LeftOrbitI2L|100|4F324F","LeftOrbitI1L|100|4F324F","ScoopPowerupL|100|4F324F","RightOrbitPowerupL|100|4F324F","CrossbowPowerupL|100|4D304D","LeftOrbitPowerupL|100|4F324F","HoleShotL|100|4F324F","ScoopShotL|100|4F324F","RightOrbitShotL|100|4F324F","RightHorseshoeShotL|100|4F324F","LeftHorseshoeShotL|0|000000","CrossbowShotL|100|4F324F","LeftOrbitShotL|100|4F324F","GIScoop2|100|4F324F","GIScoop1|100|4F324F","GIRightRamp|100|4F324F","GICrossbowLeft|100|4F324F","Light003|100|090609","Light002|100|4F324F","GIBumperInlane2|100|4F324F","GIRightOrbit3|100|4F324F","GIRightOrbit2|100|4F324F","GIRightOrbit1|100|4F324F","GIVUK|100|4F324F","GILeftOrbit4|100|4F324F","Light001|100|090609","GIRightInlaneUpper|100|0A060A","GILeftSlingshotLower|100|4F324F","GIRightSlingshotUpper|100|4F324F","GILeftSlingshot|100|4F324F","GIRightSlingshotLow|100|4F324F"), _
Array("LeftOutlanePoisonL|100|4C314C","RightOutlanePoisonL|100|4C314C","SpinnerShotL|100|4C314C","ObjectiveVikingL|100|090609","ObjectiveFinalBossL|0|000000","ObjectiveEscapeL|100|462D46","ObjectiveChaserL|100|4C314C","ObjectiveDragonL|100|4C314C","ObjectiveCastleL|100|4C314C","CrossbowCenterStandupL|100|4A304A","CrossbowLeftStandupL|100|4C314C","StandupsNL|100|4C314C","StandupsIL|100|4B304B","DropTargetsPL|100|4C314C","DropTargetsHL|100|4C314C","DropTargetsPlusL|100|301F30","DropTargetsASL|100|4C314C","DropTargetsERL|100|4C314C","DropTargetsCHL|100|4C314C","RightHorseshoeHL|100|4C314C","RightHorseshoeTL|100|4C314C","RightHorseshoeIL|100|4C314C","RightHorseshoeML|100|4C314C","RightHorseshoeSL|100|4C314C","RightOrbitBumpersL|100|4C314C","LeftOrbitBumpersL|100|4C314C","ScoopLockL|100|4C314C","LeftOrbitLockL|100|4C314C","ScoopONL|100|4C314C","ScoopDRL|100|4C314C","ScoopAGL|100|4C314C","CrossbowERL|100|070407","RightOrbitVL|100|4C314C","RightOrbitKL|100|4C314C","RightOrbitNL|100|4C314C","LeftOrbitGL|100|4C314C","LeftOrbitI2L|100|4C314C","LeftOrbitI1L|100|4C314C","ScoopPowerupL|100|4C314C","RightOrbitPowerupL|100|4C314C","CrossbowPowerupL|100|4C314C","LeftOrbitPowerupL|100|4C314C","HoleShotL|100|4C314C","ScoopShotL|100|4C314C","RightOrbitShotL|100|4C314C","RightHorseshoeShotL|100|4C314C","CrossbowShotL|100|4C314C","LeftOrbitShotL|100|4C314C","GIScoop2|100|4C314C","GIScoop1|100|4C314C","GIRightRamp|100|4C314C","GICrossbowLeft|100|4C314C","Light003|0|000000","Light002|100|4C314C","GIBumperInlane2|100|4C314C","GIRightOrbit3|100|4C314C","GIRightOrbit2|100|4C314C","GIRightOrbit1|100|4C314C","GIVUK|100|4C314C","GILeftOrbit4|100|4C314C","Light001|0|000000","GIRightInlaneUpper|100|4C314C","GIRightInlaneLower|100|110B11","GILeftInlaneUpper|100|0E090E","GILeftSlingshotLower|100|4C314C","GIRightSlingshotUpper|100|4C314C","GILeftSlingshot|100|4C314C","GIRightSlingshotLow|100|4C314C"), _
Array("LeftOutlanePoisonL|100|4A2F4A","RightOutlanePoisonL|100|4A2F4A","SpinnerShotL|100|4A2F4A","ObjectiveVikingL|0|000000","ObjectiveEscapeL|100|030203","ObjectiveChaserL|100|3B263B","ObjectiveDragonL|100|4A2F4A","ObjectiveCastleL|100|4A2F4A","CrossbowCenterStandupL|100|422A42","CrossbowLeftStandupL|100|4A2F4A","StandupsNL|100|4A2F4A","StandupsIL|100|462C46","DropTargetsPL|100|4A2F4A","DropTargetsHL|100|4A2F4A","DropTargetsPlusL|100|201520","DropTargetsASL|100|4A2F4A","DropTargetsERL|100|4A2F4A","DropTargetsCHL|100|4A2F4A","RightHorseshoeHL|100|2C1C2C","RightHorseshoeTL|100|4A2F4A","RightHorseshoeIL|100|4A2F4A","RightHorseshoeML|100|4A2F4A","RightHorseshoeSL|100|4A2F4A","RightOrbitBumpersL|100|4A2F4A","LeftOrbitBumpersL|100|4A2F4A","ScoopLockL|100|4A2F4A","LeftOrbitLockL|100|4A2F4A","ScoopONL|100|4A2F4A","ScoopDRL|100|4A2F4A","ScoopAGL|100|4A2F4A","CrossbowERL|100|3F283F","CrossbowIPL|100|020102","RightOrbitVL|100|4A2F4A","RightOrbitKL|100|4A2F4A","RightOrbitNL|100|4A2F4A","LeftOrbitGL|100|4A2F4A","LeftOrbitI2L|100|4A2F4A","LeftOrbitI1L|100|4A2F4A","ScoopPowerupL|100|4A2F4A","RightOrbitPowerupL|100|4A2F4A","CrossbowPowerupL|100|4A2F4A","LeftOrbitPowerupL|100|4A2F4A","HoleShotL|100|4A2F4A","ScoopShotL|100|4A2F4A","RightOrbitShotL|100|4A2F4A","RightHorseshoeShotL|100|4A2F4A","CrossbowShotL|100|4A2F4A","LeftOrbitShotL|100|4A2F4A","GIScoop2|100|4A2F4A","GIScoop1|100|4A2F4A","GIRightRamp|100|3A253A","GICrossbowLeft|100|4A2F4A","Light002|100|4A2F4A","GIBumperInlane2|100|4A2F4A","GIRightOrbit3|100|4A2F4A","GIRightOrbit2|100|4A2F4A","GIRightOrbit1|100|4A2F4A","GIVUK|100|4A2F4A","GILeftOrbit4|100|4A2F4A","GIRightInlaneUpper|100|4A2F4A","GIRightInlaneLower|100|4A2F4A","GILeftInlaneLower|100|402940","GILeftInlaneUpper|100|4A2F4A","GILeftSlingshotLower|100|4A2F4A","GIRightSlingshotUpper|100|4A2F4A","GILeftSlingshot|100|4A2F4A","GIRightSlingshotLow|100|4A2F4A"), _
Array("LeftOutlanePoisonL|100|482E48","RightOutlanePoisonL|100|482E48","SpinnerShotL|100|482E48","ObjectiveEscapeL|0|000000","ObjectiveChaserL|0|000000","ObjectiveDragonL|100|040304","ObjectiveCastleL|100|482E48","CrossbowCenterStandupL|100|301F30","CrossbowLeftStandupL|100|482E48","StandupsNL|100|482E48","StandupsIL|100|322032","DropTargetsPL|100|482E48","DropTargetsHL|100|482E48","DropTargetsPlusL|100|100A10","DropTargetsASL|100|482E48","DropTargetsERL|100|482E48","DropTargetsCHL|100|482E48","RightHorseshoeHL|100|020102","RightHorseshoeTL|100|261926","RightHorseshoeIL|100|472E47","RightHorseshoeML|100|482E48","RightHorseshoeSL|100|482E48","RightOrbitBumpersL|100|482E48","LeftOrbitBumpersL|100|482E48","ScoopLockL|100|482E48","LeftOrbitLockL|100|482E48","ScoopONL|100|482E48","ScoopDRL|100|482E48","ScoopAGL|100|482E48","CrossbowERL|100|482E48","CrossbowIPL|100|2A1B2A","RightOrbitVL|100|482E48","RightOrbitKL|100|482E48","RightOrbitNL|100|482E48","LeftOrbitGL|100|482E48","LeftOrbitI2L|100|482E48","LeftOrbitI1L|100|482E48","ScoopPowerupL|100|482E48","RightOrbitPowerupL|100|482E48","CrossbowPowerupL|100|482E48","LeftOrbitPowerupL|100|482E48","HoleShotL|100|482E48","ScoopShotL|100|482E48","RightOrbitShotL|100|482E48","RightHorseshoeShotL|100|482E48","CrossbowShotL|100|482E48","LeftOrbitShotL|100|482E48","GIScoop2|100|482E48","GIScoop1|100|482E48","GIRightRamp|100|0B070B","GICrossbowLeft|100|482E48","Light002|100|482E48","GIBumperInlane2|100|482E48","GIRightOrbit3|100|482E48","GIRightOrbit2|100|482E48","GIRightOrbit1|100|482E48","GIVUK|100|482E48","GILeftOrbit4|100|482E48","GIRightInlaneUpper|100|482E48","GIRightInlaneLower|100|482E48","GILeftInlaneLower|100|482E48","GILeftInlaneUpper|100|482E48","GILeftSlingshotLower|100|482E48","GIRightSlingshotUpper|100|422A42","GILeftSlingshot|100|482E48","GIRightSlingshotLow|100|482E48"), _
Array("LeftOutlanePoisonL|100|462D46","RightOutlanePoisonL|100|462D46","SpinnerShotL|100|462D46","ObjectiveDragonL|0|000000","ObjectiveCastleL|100|2D1D2D","CrossbowCenterStandupL|100|1F141F","CrossbowLeftStandupL|100|462D46","StandupsNL|100|462D46","StandupsIL|100|191019","DropTargetsPL|100|462D46","DropTargetsHL|100|462D46","DropTargetsPlusL|0|000000","DropTargetsASL|100|462D46","DropTargetsERL|100|462D46","DropTargetsCHL|100|462D46","RightHorseshoeHL|0|000000","RightHorseshoeTL|100|010101","RightHorseshoeIL|100|281928","RightHorseshoeML|100|462D46","RightHorseshoeSL|100|462D46","RightOrbitBumpersL|100|462D46","LeftOrbitBumpersL|100|462D46","ScoopLockL|100|462D46","LeftOrbitLockL|100|462D46","ScoopONL|100|462D46","ScoopDRL|100|462D46","ScoopAGL|100|462D46","CrossbowERL|100|462D46","CrossbowSNL|100|1B111B","CrossbowIPL|100|462D46","RightOrbitVL|100|462D46","RightOrbitKL|100|462D46","RightOrbitNL|100|462D46","LeftOrbitGL|100|462D46","LeftOrbitI2L|100|462D46","LeftOrbitI1L|100|462D46","ScoopPowerupL|100|462D46","RightOrbitPowerupL|100|462D46","CrossbowPowerupL|100|462D46","LeftOrbitPowerupL|100|462D46","HoleShotL|100|462D46","ScoopShotL|100|462D46","RightOrbitShotL|100|462D46","RightHorseshoeShotL|100|462D46","CrossbowShotL|100|462D46","LeftOrbitShotL|100|462D46","GIScoop2|100|462D46","GIScoop1|100|462D46","GIRightRamp|0|000000","GICrossbowLeft|100|462D46","Light002|100|462D46","GIBumperInlane2|100|462D46","GIRightOrbit3|100|462D46","GIRightOrbit2|100|462D46","GIRightOrbit1|100|462D46","GIVUK|100|462D46","GILeftOrbit4|100|462D46","GILeftOrbit1|100|010101","GIRightInlaneUpper|100|462D46","GIRightInlaneLower|100|462D46","GILeftInlaneLower|100|462D46","GILeftInlaneUpper|100|462D46","GILeftSlingshotLower|100|462D46","GIRightSlingshotUpper|0|000000","GILeftSlingshot|100|462D46","GIRightSlingshotLow|100|462D46"), _
Array("LeftOutlanePoisonL|100|442B44","RightOutlanePoisonL|100|442B44","SpinnerShotL|100|442B44","ObjectiveCastleL|0|000000","CrossbowCenterStandupL|100|100A10","CrossbowLeftStandupL|100|442B44","StandupsNL|100|442B44","StandupsIL|100|050305","DropTargetsPL|100|442B44","DropTargetsHL|100|442B44","DropTargetsASL|100|442B44","DropTargetsERL|100|442B44","DropTargetsCHL|100|422A42","RightHorseshoeTL|0|000000","RightHorseshoeIL|100|010101","RightHorseshoeML|100|321F32","RightHorseshoeSL|100|442B44","LeftHorseshoeKL|100|1C121C","RightOrbitBumpersL|100|442B44","LeftOrbitBumpersL|100|442B44","ScoopLockL|100|442B44","LeftOrbitLockL|100|442B44","ScoopONL|100|442B44","ScoopDRL|100|442B44","ScoopAGL|100|442B44","CrossbowERL|100|442B44","CrossbowSNL|100|432A43","CrossbowIPL|100|442B44","RightOrbitVL|100|442B44","RightOrbitKL|100|442B44","RightOrbitNL|100|442B44","LeftOrbitGL|100|442B44","LeftOrbitI2L|100|442B44","LeftOrbitI1L|100|442B44","ScoopPowerupL|100|442B44","RightOrbitPowerupL|100|442B44","CrossbowPowerupL|100|442B44","LeftOrbitPowerupL|100|442B44","HoleShotL|100|442B44","ScoopShotL|100|442B44","RightOrbitShotL|100|442B44","RightHorseshoeShotL|100|442B44","CrossbowShotL|100|442B44","LeftOrbitShotL|100|442B44","GIScoop2|100|442B44","GIUpperFlipper|100|2E1D2E","GIScoop1|100|442B44","GICrossbowLeft|100|442B44","Light002|100|442B44","GIBumperInlane2|100|402840","GIRightOrbit3|100|442B44","GIRightOrbit2|100|442B44","GIRightOrbit1|100|442B44","GIVUK|100|442B44","GILeftOrbit4|100|442B44","GILeftOrbit1|100|1B111B","GIRightInlaneUpper|100|442B44","GIRightInlaneLower|100|442B44","GILeftInlaneLower|100|442B44","GILeftInlaneUpper|100|442B44","GILeftSlingshotLower|100|442B44","GILeftSlingshot|100|3B263B","GIRightSlingshotLow|100|442B44"), _
Array("LeftOutlanePoisonL|100|412A41","RightOutlanePoisonL|100|412A41","SpinnerShotL|100|412A41","CrossbowCenterStandupL|100|040304","CrossbowLeftStandupL|100|412A41","StandupsNL|100|412A41","StandupsIL|0|000000","StandupsTL|100|040304","DropTargetsPL|100|412A41","DropTargetsHL|100|412A41","DropTargetsASL|100|412A41","DropTargetsERL|100|412A41","DropTargetsCHL|100|3B263B","RightHorseshoeIL|0|000000","RightHorseshoeML|100|050305","RightHorseshoeSL|100|352335","LeftHorseshoeKL|100|352235","RightOrbitBumpersL|100|412A41","LeftOrbitBumpersL|100|412A41","ScoopLockL|100|412A41","LeftOrbitLockL|100|412A41","ScoopONL|100|412A41","ScoopDRL|100|412A41","ScoopAGL|100|412A41","CrossbowERL|100|412A41","CrossbowSNL|100|412A41","CrossbowIPL|100|412A41","RightOrbitVL|100|412A41","RightOrbitKL|100|412A41","RightOrbitNL|100|412A41","LeftOrbitGL|100|412A41","LeftOrbitI2L|100|412A41","LeftOrbitI1L|100|412A41","ScoopPowerupL|100|412A41","RightOrbitPowerupL|100|412A41","CrossbowPowerupL|100|412A41","LeftOrbitPowerupL|100|412A41","HoleShotL|100|412A41","ScoopShotL|100|412A41","RightOrbitShotL|100|412A41","RightHorseshoeShotL|100|412A41","CrossbowShotL|100|412A41","LeftOrbitShotL|100|412A41","GIScoop2|100|412A41","GIUpperFlipper|100|412A41","GIScoop1|100|412A41","GICrossbowLeft|100|412A41","Light002|100|402A40","GIBumperInlane2|100|030203","GIRightOrbit3|100|412A41","GIRightOrbit2|100|412A41","GIRightOrbit1|100|412A41","GIVUK|100|412A41","GILeftOrbit4|100|412A41","GILeftOrbit2|100|030203","GILeftOrbit1|100|412A41","GIRightInlaneUpper|100|412A41","GIRightInlaneLower|100|412A41","GILeftInlaneLower|100|412A41","GILeftInlaneUpper|100|412A41","GILeftSlingshotLower|100|412A41","GILeftSlingshot|100|010101","GIRightSlingshotLow|100|412A41"), _
Array("LeftOutlanePoisonL|100|3F283F","RightOutlanePoisonL|100|3F283F","SpinnerShotL|100|3F283F","CrossbowCenterStandupL|0|000000","CrossbowLeftStandupL|100|3F283F","StandupsNL|100|3F283F","StandupsTL|100|070407","DropTargetsPL|100|3F283F","DropTargetsHL|100|3C263C","DropTargetsASL|100|3F283F","DropTargetsERL|100|3F283F","DropTargetsCHL|100|342134","RightHorseshoeML|0|000000","RightHorseshoeSL|100|090609","LeftHorseshoeKL|100|3F283F","LeftHorseshoeCL|100|030203","RightOrbitBumpersL|100|3F283F","LeftOrbitBumpersL|100|3F283F","ScoopLockL|100|3F283F","LeftOrbitLockL|100|3F283F","ScoopONL|100|3F283F","ScoopDRL|100|3F283F","ScoopAGL|100|3F283F","CrossbowERL|100|3F283F","CrossbowSNL|100|3F283F","CrossbowIPL|100|3F283F","RightOrbitVL|100|3D263D","RightOrbitKL|100|3F283F","RightOrbitNL|100|3F283F","LeftOrbitGL|100|3F283F","LeftOrbitI2L|100|3F283F","LeftOrbitI1L|100|3F283F","ScoopPowerupL|100|3F283F","RightOrbitPowerupL|100|3F283F","CrossbowPowerupL|100|3F283F","LeftOrbitPowerupL|100|3F283F","HoleShotL|100|3F283F","ScoopShotL|100|3F283F","RightOrbitShotL|100|3F283F","RightHorseshoeShotL|100|3F283F","CrossbowShotL|100|3F283F","LeftOrbitShotL|100|3F283F","GIScoop2|100|3F283F","GIUpperFlipper|100|3F283F","GIScoop1|100|3F283F","GICrossbowLeft|100|3F283F","Light002|100|3D273D","GIBumperInlane2|0|000000","GIRightOrbit3|100|3F283F","GIRightOrbit2|100|3F283F","GIRightOrbit1|100|3F283F","GIVUK|100|3F283F","GILeftOrbit4|100|261826","GILeftOrbit2|100|2F1E2F","GILeftOrbit1|100|3F283F","GIRightInlaneUpper|100|3F283F","GIRightInlaneLower|100|3F283F","GILeftInlaneLower|100|3F283F","GILeftInlaneUpper|100|3F283F","GILeftSlingshotLower|100|3F283F","GILeftSlingshot|0|000000","GIRightSlingshotLow|100|3F283F"), _
Array("LeftOutlanePoisonL|100|3D273D","RightOutlanePoisonL|100|3D273D","SpinnerShotL|100|3D273D","CrossbowLeftStandupL|100|3D273D","StandupsNL|100|3B263B","StandupsTL|100|100A10","DropTargetsPL|100|3D273D","DropTargetsHL|100|372337","DropTargetsASL|100|3D273D","DropTargetsERL|100|3D273D","DropTargetsCHL|100|2C1C2C","RightHorseshoeSL|0|000000","LeftHorseshoeKL|100|3D273D","LeftHorseshoeCL|100|110B11","RightOrbitBumpersL|100|3D273D","LeftOrbitBumpersL|100|3D273D","ScoopLockL|100|3D273D","LeftOrbitLockL|100|3D273D","ScoopONL|100|3D273D","ScoopDRL|100|3D273D","ScoopAGL|100|3D273D","CrossbowERL|100|3D273D","CrossbowSNL|100|3D273D","CrossbowIPL|100|3D273D","RightOrbitVL|100|0D080D","RightOrbitKL|100|3D273D","RightOrbitNL|100|3D273D","LeftOrbitGL|100|3D273D","LeftOrbitI2L|100|3D273D","LeftOrbitI1L|100|3D273D","RightRampSL|100|030203","ScoopPowerupL|100|3D273D","RightOrbitPowerupL|100|3D273D","CrossbowPowerupL|100|3D273D","LeftOrbitPowerupL|100|3D273D","HoleShotL|100|3D273D","ScoopShotL|100|3D273D","RightOrbitShotL|100|3D273D","RightHorseshoeShotL|100|392539","CrossbowShotL|100|3D273D","LeftOrbitShotL|100|3D273D","GIScoop2|100|3D273D","GIUpperFlipper|100|3D273D","GIScoop1|100|3D273D","GICrossbowLeft|100|3D273D","Light002|100|352235","GIRightOrbit3|100|3D273D","GIRightOrbit2|100|3D273D","GIRightOrbit1|100|3D273D","GIVUK|100|3D273D","GILeftOrbit4|0|000000","GILeftOrbit2|100|3D273D","GILeftOrbit1|100|3D273D","GIRightInlaneUpper|100|3D273D","GIRightInlaneLower|100|3D273D","GILeftInlaneLower|100|3D273D","GILeftInlaneUpper|100|3D273D","GILeftSlingshotLower|100|3D273D","GIRightSlingshotLow|100|3A253A"), _
Array("LeftOutlanePoisonL|100|3B253B","RightOutlanePoisonL|100|3B253B","SpinnerShotL|100|3B253B","CrossbowLeftStandupL|100|3B253B","StandupsNL|100|1E131E","StandupsTL|100|1A101A","DropTargetsPL|100|3B253B","DropTargetsHL|100|2D1C2D","DropTargetsASL|100|3B253B","DropTargetsERL|100|3B253B","DropTargetsCHL|100|1C121C","LeftHorseshoeKL|100|3B253B","LeftHorseshoeCL|100|2C1C2C","RightOrbitBumpersL|100|392339","LeftOrbitBumpersL|100|3B253B","ScoopLockL|100|3B253B","LeftOrbitLockL|100|3B253B","ScoopONL|100|3B253B","ScoopDRL|100|3B253B","ScoopAGL|100|3B253B","CrossbowERL|100|3B253B","CrossbowSNL|100|3B253B","CrossbowIPL|100|3B253B","RightOrbitVL|0|000000","RightOrbitKL|100|3A253A","RightOrbitNL|100|3B253B","LeftOrbitGL|100|3B253B","LeftOrbitI2L|100|3B253B","LeftOrbitI1L|100|3B253B","RightRampSL|100|070507","ScoopPowerupL|100|3B253B","RightOrbitPowerupL|100|3B253B","RightRampPowerupL|100|010101","CrossbowPowerupL|100|3B253B","LeftOrbitPowerupL|100|3B253B","HoleShotL|100|3B253B","ScoopShotL|100|3B253B","RightOrbitShotL|100|3B253B","RightHorseshoeShotL|100|060406","RightRampShotL|100|010001","CrossbowShotL|100|3B253B","LeftOrbitShotL|100|3B253B","GIScoop2|100|3B253B","GIUpperFlipper|100|3B253B","GIScoop1|100|3B253B","GICrossbowLeft|100|3B253B","Light002|100|2A1A2A","GIRightOrbit3|100|3B253B","GIRightOrbit2|100|3B253B","GIRightOrbit1|100|3B253B","GIVUK|100|382338","GILeftOrbit2|100|3B253B","GILeftOrbit1|100|3B253B","GIRightInlaneUpper|100|3B253B","GIRightInlaneLower|100|3B253B","GILeftInlaneLower|100|3B253B","GILeftInlaneUpper|100|3B253B","GILeftSlingshotLower|100|3B253B","GIRightSlingshotLow|100|020102"), _
Array("LeftOutlanePoisonL|100|392439","RightOutlanePoisonL|100|392439","SpinnerShotL|100|392439","CrossbowLeftStandupL|100|342134","StandupsNL|100|070407","StandupsTL|100|211521","DropTargetsPL|100|392439","DropTargetsHL|100|1F141F","DropTargetsASL|100|392439","DropTargetsERL|100|392439","DropTargetsCHL|100|0E090E","LeftHorseshoeKL|100|392439","LeftHorseshoeCL|100|372337","RightOrbitBumpersL|100|0F090F","LeftOrbitBumpersL|100|392439","ScoopLockL|100|392439","LeftOrbitLockL|100|392439","ScoopONL|100|392439","ScoopDRL|100|392439","ScoopAGL|100|392439","CrossbowERL|100|392439","CrossbowSNL|100|392439","CrossbowIPL|100|392439","RightOrbitKL|100|201420","RightOrbitNL|100|392439","LeftOrbitGL|100|392439","LeftOrbitI2L|100|392439","LeftOrbitI1L|100|392439","RightRampSL|100|120B12","ScoopPowerupL|100|392439","RightOrbitPowerupL|100|392439","RightRampPowerupL|100|080508","CrossbowPowerupL|100|392439","LeftOrbitPowerupL|100|392439","HoleShotL|100|392439","ScoopShotL|100|392439","RightOrbitShotL|100|392439","RightHorseshoeShotL|0|000000","RightRampShotL|100|0C070C","CrossbowShotL|100|392439","LeftOrbitShotL|100|392439","GIScoop2|100|392439","GIUpperFlipper|100|392439","GIScoop1|100|392439","GICrossbowLeft|100|392439","Light002|100|1D121D","GIRightOrbit3|100|392439","GIRightOrbit2|100|392439","GIRightOrbit1|100|392439","GIVUK|100|070507","GILeftOrbit2|100|392439","GILeftOrbit1|100|392439","GIRightRubbersUpper|100|2C1C2C","GIRightInlaneUpper|100|392439","GIRightInlaneLower|100|392439","GILeftInlaneLower|100|392439","GILeftInlaneUpper|100|392439","GILeftSlingshotLower|100|1D131D","GIRightSlingshotLow|0|000000"), _
Array("LeftOutlanePoisonL|100|362336","RightOutlanePoisonL|100|1E131E","SpinnerShotL|100|362336","CrossbowLeftStandupL|100|2B1C2B","StandupsNL|0|000000","StandupsTL|100|261926","DropTargetsPL|100|362336","DropTargetsHL|100|0D080D","DropTargetsASL|100|362336","DropTargetsERL|100|362336","DropTargetsCHL|100|030203","LeftHorseshoeKL|100|362336","LeftHorseshoeCL|100|362336","LeftHorseshoeAL|100|080508","RightOrbitBumpersL|0|000000","LeftOrbitBumpersL|100|362336","ScoopLockL|100|362336","LeftOrbitLockL|100|362336","ScoopONL|100|362336","ScoopDRL|100|362336","ScoopAGL|100|362336","CrossbowERL|100|362336","CrossbowSNL|100|362336","CrossbowIPL|100|362336","RightOrbitKL|100|010101","RightOrbitNL|100|362336","LeftOrbitGL|100|362336","LeftOrbitI2L|100|362336","LeftOrbitI1L|100|362336","RightRampSL|100|1F141F","ScoopPowerupL|100|362336","RightOrbitPowerupL|100|191019","RightRampPowerupL|100|181018","CrossbowPowerupL|100|362336","LeftOrbitPowerupL|100|362336","HoleShotL|100|362336","ScoopShotL|100|362336","RightOrbitShotL|100|362336","RightRampShotL|100|201420","CrossbowShotL|100|362336","LeftOrbitShotL|100|362336","GIScoop2|100|362336","GIUpperFlipper|100|362336","GIScoop1|100|362336","GICrossbowLeft|100|362336","Light002|100|100B10","GIRightOrbit3|100|362336","GIRightOrbit2|100|362336","GIRightOrbit1|100|362336","GIVUK|0|000000","GILeftOrbit2|100|362336","GILeftOrbit1|100|362336","GIRightRubbersUpper|100|362336","GIRightInlaneUpper|100|362336","GIRightInlaneLower|100|362336","GILeftInlaneLower|100|362336","GILeftInlaneUpper|100|362336","GILeftSlingshotLower|0|000000"), _
Array("LeftOutlanePoisonL|100|342134","RightOutlanePoisonL|0|000000","SpinnerShotL|100|342134","CrossbowLeftStandupL|100|1C121C","StandupsTL|100|2B1B2B","DropTargetsPL|100|342134","DropTargetsHL|100|060406","DropTargetsASL|100|342134","DropTargetsERL|100|342134","DropTargetsCHL|0|000000","LeftHorseshoeKL|100|342134","LeftHorseshoeCL|100|342134","LeftHorseshoeAL|100|1F131F","LeftOrbitBumpersL|100|342134","ScoopLockL|100|342134","LeftOrbitLockL|100|342134","ScoopONL|100|342134","ScoopDRL|100|342134","ScoopAGL|100|342134","CrossbowERL|100|342134","CrossbowSNL|100|342134","CrossbowIPL|100|342134","RightOrbitKL|0|000000","RightOrbitNL|100|2C1C2C","LeftOrbitGL|100|342134","LeftOrbitI2L|100|342134","LeftOrbitI1L|100|342134","RightRampSL|100|2A1B2A","ScoopPowerupL|100|342134","RightOrbitPowerupL|0|000000","RightRampPowerupL|100|2C1C2C","CrossbowPowerupL|100|342134","LeftOrbitPowerupL|100|342134","HoleShotL|100|342134","ScoopShotL|100|342134","RightOrbitShotL|100|301F30","RightRampShotL|100|301E30","CrossbowShotL|100|342134","LeftOrbitShotL|100|342134","LeftRampShotL|100|150E15","GIScoop2|100|342134","GIUpperFlipper|100|342134","GIScoop1|100|342134","GICrossbowLeft|100|342134","Light002|100|080508","GIRightOrbit3|100|342134","GIRightOrbit2|100|342134","GIRightOrbit1|100|342134","GILeftOrbit2|100|342134","GILeftOrbit1|100|342134","GIRightRubbersUpper|100|342134","GIRightInlaneUpper|100|2A1A2A","GIRightInlaneLower|100|342134","GILeftInlaneLower|100|342134","GILeftInlaneUpper|100|342134"), _
Array("LeftOutlanePoisonL|100|322032","SpinnerShotL|100|322032","CrossbowLeftStandupL|100|110B11","StandupsTL|100|2D1D2D","DropTargetsPL|100|322032","DropTargetsHL|100|030203","DropTargetsASL|100|2F1E2F","DropTargetsERL|100|322032","LeftHorseshoeKL|100|322032","LeftHorseshoeCL|100|322032","LeftHorseshoeAL|100|2E1D2E","LeftOrbitBumpersL|100|322032","ScoopLockL|100|322032","LeftOrbitLockL|100|322032","ScoopONL|100|322032","ScoopDRL|100|322032","ScoopAGL|100|322032","CrossbowERL|100|322032","CrossbowSNL|100|322032","CrossbowIPL|100|322032","RightOrbitNL|100|080508","LeftOrbitGL|100|322032","LeftOrbitI2L|100|322032","LeftOrbitI1L|100|322032","RightRampSL|100|312031","ScoopPowerupL|100|322032","RightRampPowerupL|100|322032","CrossbowPowerupL|100|322032","LeftOrbitPowerupL|100|322032","HoleShotL|100|322032","ScoopShotL|100|322032","RightOrbitShotL|100|070407","RightRampShotL|100|322032","CrossbowShotL|100|322032","LeftOrbitShotL|100|322032","LeftRampShotL|100|322032","GIScoop2|100|322032","GIUpperFlipper|100|322032","GIScoop1|100|322032","GICrossbowLeft|100|301E30","Light002|100|020102","GIRightOrbit3|100|322032","GIRightOrbit2|100|322032","GIRightOrbit1|100|322032","GILeftOrbit3|100|010101","GILeftOrbit2|100|322032","GILeftOrbit1|100|322032","GIRightRubbersUpper|100|322032","GIRightInlaneUpper|0|000000","GIRightInlaneLower|100|2C1C2C","GILeftInlaneLower|100|322032","GILeftInlaneUpper|100|322032"), _
Array("LeftOutlanePoisonL|100|100A10","SpinnerShotL|100|271927","CrossbowLeftStandupL|0|000000","StandupsTL|100|2E1D2E","DropTargetsPL|100|301E30","DropTargetsHL|100|010101","DropTargetsASL|100|281928","DropTargetsERL|100|301E30","LeftHorseshoeKL|100|301E30","LeftHorseshoeCL|100|301E30","LeftHorseshoeAL|100|301E30","LeftHorseshoeLL|100|040304","LeftOrbitBumpersL|100|301E30","ScoopLockL|100|301E30","LeftOrbitLockL|100|301E30","ScoopONL|100|301E30","ScoopDRL|100|301E30","ScoopAGL|100|301E30","CrossbowERL|100|301E30","CrossbowSNL|100|301E30","CrossbowIPL|100|301E30","RightOrbitNL|0|000000","LeftOrbitGL|100|301E30","LeftOrbitI2L|100|301E30","LeftOrbitI1L|100|301E30","RightRampAL|100|050305","RightRampSL|100|301E30","LeftRampPL|100|020102","ScoopPowerupL|100|301E30","RightRampPowerupL|100|301E30","CrossbowPowerupL|100|301E30","LeftOrbitPowerupL|100|301E30","LeftRampPowerupL|100|0C080C","HoleShotL|100|2E1D2E","ScoopShotL|100|301E30","RightOrbitShotL|0|000000","RightRampShotL|100|301E30","CrossbowShotL|100|301E30","LeftOrbitShotL|100|301E30","LeftRampShotL|100|301E30","GIScoop2|100|301E30","GIUpperFlipper|100|301E30","GIScoop1|100|301E30","GICrossbowLeft|100|281928","Light002|0|000000","GIRightOrbit3|100|301E30","GIRightOrbit2|100|301E30","GIRightOrbit1|100|301E30","GILeftOrbit3|100|1A101A","GILeftOrbit2|100|301E30","GILeftOrbit1|100|301E30","GIRightRubbersUpper|100|301E30","GIRightInlaneLower|0|000000","GILeftInlaneLower|100|301E30","GILeftInlaneUpper|100|2D1C2D"), _
Array("LeftOutlanePoisonL|0|000000","SpinnerShotL|100|030203","DropTargetsPL|100|2D1C2D","DropTargetsHL|0|000000","DropTargetsASL|100|1E131E","DropTargetsERL|100|2E1D2E","LeftHorseshoeKL|100|2E1D2E","LeftHorseshoeCL|100|2E1D2E","LeftHorseshoeAL|100|2E1D2E","LeftHorseshoeLL|100|140D14","LeftOrbitBumpersL|100|2E1D2E","ScoopLockL|100|2E1D2E","LeftOrbitLockL|100|2E1D2E","ScoopONL|100|2E1D2E","ScoopDRL|100|2E1D2E","ScoopAGL|100|2E1D2E","CrossbowERL|100|2E1D2E","CrossbowSNL|100|2E1D2E","CrossbowIPL|100|2E1D2E","LeftOrbitGL|100|2E1D2E","LeftOrbitI2L|100|2E1D2E","LeftOrbitI1L|100|2E1D2E","RightRampAL|100|110B11","RightRampSL|100|2E1D2E","LeftRampPL|100|281928","ScoopPowerupL|100|2E1D2E","RightRampPowerupL|100|2E1D2E","CrossbowPowerupL|100|2E1D2E","LeftOrbitPowerupL|100|2E1D2E","LeftRampPowerupL|100|2D1C2D","HoleShotL|100|2B1B2B","ScoopShotL|100|2E1D2E","RightRampShotL|100|2E1D2E","CrossbowShotL|100|2E1D2E","LeftOrbitShotL|100|2E1D2E","LeftRampShotL|100|2E1D2E","GIScoop2|100|2E1D2E","GIUpperFlipper|100|2E1D2E","GIScoop1|100|2E1D2E","GICrossbowLeft|100|140D14","GIRightOrbit3|100|2E1D2E","GIRightOrbit2|100|2E1D2E","GIRightOrbit1|100|291A29","GILeftOrbit3|100|2C1C2C","GILeftOrbit2|100|2E1D2E","GILeftOrbit1|100|2E1D2E","GIRightRubbersUpper|100|2E1D2E","GILeftInlaneLower|100|160E16","GILeftInlaneUpper|100|020102"), _
Array("SpinnerShotL|0|000000","ObjectiveFinalJudgmentL|100|0E090E","StandupsTL|100|2B1C2B","DropTargetsPL|100|281A28","DropTargetsASL|100|120C12","DropTargetsERL|100|2B1C2B","LeftHorseshoeKL|100|2B1C2B","LeftHorseshoeCL|100|2B1C2B","LeftHorseshoeAL|100|2B1C2B","LeftHorseshoeLL|100|251825","LeftOrbitBumpersL|100|2B1C2B","ScoopLockL|100|2B1C2B","LeftOrbitLockL|100|2B1C2B","ScoopONL|100|2B1C2B","ScoopDRL|100|2B1C2B","ScoopAGL|100|2B1C2B","CrossbowERL|100|2B1C2B","CrossbowSNL|100|2B1C2B","CrossbowIPL|100|2B1C2B","LeftOrbitGL|100|2B1C2B","LeftOrbitI2L|100|2B1C2B","LeftOrbitI1L|100|2B1C2B","RightRampAL|100|1D131D","RightRampSL|100|2B1C2B","LeftRampPL|100|2B1C2B","LeftRampCL|100|0E090E","ScoopPowerupL|100|2B1C2B","RightRampPowerupL|100|2B1C2B","CrossbowPowerupL|100|2B1C2B","LeftOrbitPowerupL|100|2B1C2B","LeftRampPowerupL|100|2B1C2B","HoleShotL|100|251825","ScoopShotL|100|2B1C2B","RightRampShotL|100|2B1C2B","CrossbowShotL|100|2B1C2B","LeftOrbitShotL|100|2B1C2B","LeftRampShotL|100|2B1C2B","GIScoop2|100|2B1C2B","GIUpperFlipper|100|2B1C2B","GIScoop1|100|2B1C2B","GICrossbowLeft|100|040304","GIRightOrbit3|100|2B1C2B","GIRightOrbit2|100|2B1C2B","GIRightOrbit1|100|040304","GILeftOrbit3|100|2B1C2B","GILeftOrbit2|100|2B1C2B","GILeftOrbit1|100|2B1C2B","GIRightRubbersUpper|100|2B1C2B","GIRightRubbersLower|100|1D131D","GILeftInlaneLower|0|000000","GILeftInlaneUpper|0|000000"), _
Array("ObjectiveFinalJudgmentL|100|291A29","StandupsRL|100|010101","StandupsTL|100|291A29","DropTargetsPL|100|201520","DropTargetsASL|100|0B070B","DropTargetsERL|100|291A29","LeftHorseshoeKL|100|291A29","LeftHorseshoeCL|100|291A29","LeftHorseshoeAL|100|291A29","LeftHorseshoeLL|100|291A29","LeftHorseshoeBL|100|020102","LeftOrbitBumpersL|100|291A29","ScoopLockL|100|291A29","LeftOrbitLockL|100|291A29","ScoopONL|100|291A29","ScoopDRL|100|291A29","ScoopAGL|100|291A29","CrossbowERL|100|291A29","CrossbowSNL|100|291A29","CrossbowIPL|100|291A29","LeftOrbitGL|100|291A29","LeftOrbitI2L|100|291A29","LeftOrbitI1L|100|291A29","RightRampEL|100|010101","RightRampAL|100|261826","RightRampSL|100|291A29","LeftRampPL|100|291A29","LeftRampCL|100|291A29","LeftRampEL|100|040304","ScoopPowerupL|100|291A29","RightRampPowerupL|100|291A29","CrossbowPowerupL|100|291A29","LeftOrbitPowerupL|100|291A29","LeftRampPowerupL|100|291A29","HoleShotL|100|1E131E","ScoopShotL|100|281928","RightRampShotL|100|291A29","CrossbowShotL|100|291A29","LeftOrbitShotL|100|291A29","LeftRampShotL|100|291A29","GIScoop2|100|291A29","GIUpperFlipper|100|291A29","GIScoop1|100|0F090F","GICrossbowLeft|100|010101","GIRightOrbit3|100|281A28","GIRightOrbit2|100|291A29","GIRightOrbit1|0|000000","GILeftOrbit3|100|291A29","GILeftOrbit2|100|291A29","GILeftOrbit1|100|291A29","GIRightRubbersUpper|100|291A29","GIRightRubbersLower|100|291A29"), _
Array("ObjectiveFinalJudgmentL|100|271927","ObjectiveSniperL|100|1B111B","StandupsRL|100|060406","StandupsTL|100|271927","DropTargetsPL|100|170F17","DropTargetsPlusL|100|010001","DropTargetsASL|100|010101","DropTargetsERL|100|251825","LeftHorseshoeKL|100|271927","LeftHorseshoeCL|100|271927","LeftHorseshoeAL|100|271927","LeftHorseshoeLL|100|271927","LeftHorseshoeBL|100|160E16","LeftOrbitBumpersL|100|271927","ScoopLockL|100|271927","LeftOrbitLockL|100|271927","ScoopONL|100|271927","ScoopDRL|100|130C13","ScoopAGL|100|271927","CrossbowERL|100|271927","CrossbowSNL|100|271927","CrossbowIPL|100|271927","LeftOrbitGL|100|251825","LeftOrbitI2L|100|271927","LeftOrbitI1L|100|271927","RightRampEL|100|080508","RightRampAL|100|271927","RightRampSL|100|271927","LeftRampPL|100|271927","LeftRampCL|100|271927","LeftRampEL|100|211521","ScoopPowerupL|100|271927","RightRampPowerupL|100|271927","CrossbowPowerupL|100|271927","LeftOrbitPowerupL|100|271927","LeftRampPowerupL|100|271927","HoleShotL|100|160E16","ScoopShotL|100|0D080D","RightRampShotL|100|271927","CrossbowShotL|100|271927","LeftOrbitShotL|100|271927","LeftRampShotL|100|271927","GIScoop2|100|271927","GIUpperFlipper|100|271927","GIScoop1|0|000000","GICrossbowLeft|0|000000","GIRightOrbit3|100|110B11","GIRightOrbit2|100|271927","GILeftOrbit3|100|271927","GILeftOrbit2|100|271927","GILeftOrbit1|100|271927","GILeftRubbersUpper|100|0C080C","GIRightRubbersUpper|100|271927","GIRightRubbersLower|100|271927"), _
Array("ObjectiveFinalJudgmentL|100|251725","ObjectiveSniperL|100|251725","ObjectiveBlacksmithL|100|1B111B","ObjectiveVikingL|100|030203","StandupsRL|100|0D080D","StandupsTL|100|251725","DropTargetsPL|100|0E080E","DropTargetsERL|100|211521","LeftHorseshoeKL|100|251725","LeftHorseshoeCL|100|251725","LeftHorseshoeAL|100|251725","LeftHorseshoeLL|100|251725","LeftHorseshoeBL|100|211521","LeftOrbitBumpersL|100|251725","ScoopLockL|100|251725","LeftOrbitLockL|100|251725","ScoopONL|100|251725","ScoopDRL|0|000000","ScoopAGL|100|251725","CrossbowERL|100|251725","CrossbowSNL|100|251725","CrossbowIPL|100|251725","LeftOrbitGL|100|191019","LeftOrbitI2L|100|251725","LeftOrbitI1L|100|251725","RightRampEL|100|140D14","RightRampAL|100|251725","RightRampSL|100|251725","LeftRampPL|100|251725","LeftRampCL|100|251725","LeftRampEL|100|251725","ScoopPowerupL|100|0E090E","RightRampPowerupL|100|251725","CrossbowPowerupL|100|251725","LeftOrbitPowerupL|100|251725","LeftRampPowerupL|100|251725","HoleShotL|100|0E090E","ScoopShotL|0|000000","RightRampShotL|100|251725","CrossbowShotL|100|251725","LeftOrbitShotL|100|251725","LeftRampShotL|100|251725","GIScoop2|100|251725","GIUpperFlipper|100|251725","GIRightOrbit3|0|000000","GIRightOrbit2|100|211421","GILeftOrbit3|100|251725","GILeftOrbit2|100|251725","GILeftOrbit1|100|251725","GILeftRubbersUpper|100|241724","GIRightRubbersUpper|100|251725","GIRightRubbersLower|100|251725"), _
Array("ObjectiveFinalJudgmentL|100|231623","ObjectiveSniperL|100|231623","ObjectiveBlacksmithL|100|231623","ObjectiveVikingL|100|211521","ObjectiveFinalBossL|100|191019","StandupsRL|100|130C13","StandupsTL|100|231623","DropTargetsPL|100|070407","DropTargetsPlusL|0|000000","DropTargetsASL|0|000000","DropTargetsERL|100|1D121D","LeftHorseshoeKL|100|231623","LeftHorseshoeCL|100|231623","LeftHorseshoeAL|100|231623","LeftHorseshoeLL|100|231623","LeftHorseshoeBL|100|231623","LeftOrbitBumpersL|100|231623","ScoopLockL|100|231623","LeftOrbitLockL|100|231623","ScoopONL|100|231623","ScoopAGL|100|180F18","CrossbowERL|100|231623","CrossbowSNL|100|231623","CrossbowIPL|100|231623","LeftOrbitGL|100|0A060A","LeftOrbitI2L|100|231623","LeftOrbitI1L|100|231623","RightRampEL|100|1D121D","RightRampAL|100|231623","RightRampSL|100|231623","LeftRampPL|100|231623","LeftRampCL|100|231623","LeftRampEL|100|231623","ScoopPowerupL|0|000000","RightRampPowerupL|100|231623","CrossbowPowerupL|100|231623","LeftOrbitPowerupL|100|231623","LeftRampPowerupL|100|231623","HoleShotL|100|080508","RightRampShotL|100|231623","CrossbowShotL|100|170E17","LeftOrbitShotL|100|231623","LeftRampShotL|100|231623","GIScoop2|100|231623","GIUpperFlipper|100|231623","GIRightOrbit2|100|040204","GILeftOrbit3|100|231623","GILeftOrbit2|100|231623","GILeftOrbit1|100|231623","GILeftRubbersUpper|100|231623","GIRightRubbersUpper|100|231623","GIRightRubbersLower|100|231623"), _
Array("ObjectiveFinalJudgmentL|100|201520","ObjectiveSniperL|100|201520","ObjectiveBlacksmithL|100|201520","ObjectiveVikingL|100|201520","ObjectiveFinalBossL|100|201520","ObjectiveEscapeL|100|020102","ObjectiveChaserL|100|090609","StandupsRL|100|181018","StandupsTL|100|201520","DropTargetsPL|100|030203","DropTargetsERL|100|120C12","LeftHorseshoeKL|100|201520","LeftHorseshoeCL|100|201520","LeftHorseshoeAL|100|201520","LeftHorseshoeLL|100|201520","LeftHorseshoeBL|100|201520","LeftOrbitBumpersL|100|201520","ScoopLockL|100|1D131D","LeftOrbitLockL|100|201520","ScoopONL|100|201520","ScoopAGL|0|000000","CrossbowERL|100|201520","CrossbowSNL|100|201520","CrossbowIPL|100|201520","LeftOrbitGL|100|010101","LeftOrbitI2L|100|201520","LeftOrbitI1L|100|201520","RightRampEL|100|201520","RightRampAL|100|201520","RightRampSL|100|201520","LeftRampPL|100|201520","LeftRampCL|100|201520","LeftRampEL|100|201520","RightRampPowerupL|100|201520","CrossbowPowerupL|100|201520","LeftOrbitPowerupL|100|201520","LeftRampPowerupL|100|201520","HoleShotL|100|030203","RightRampShotL|100|201520","LeftHorseshoeShotL|100|010001","CrossbowShotL|100|050305","LeftOrbitShotL|100|201520","LeftRampShotL|100|201520","GIScoop2|100|0B070B","GIUpperFlipper|100|201520","GIRightOrbit2|0|000000","GIBumperInlanes3|100|0A070A","GIBumperInlanes2|100|020102","GILeftOrbit3|100|201520","GILeftOrbit2|100|201520","GILeftOrbit1|100|201520","GILeftRubbersUpper|100|201520","Light001|100|040304","GIRightRubbersUpper|100|201520","GIRightRubbersLower|100|201520","GIRightSlingshotUpper|100|010101"), _
Array("ObjectiveFinalJudgmentL|100|1E131E","ObjectiveSniperL|100|1E131E","ObjectiveBlacksmithL|100|1E131E","ObjectiveVikingL|100|1E131E","ObjectiveFinalBossL|100|1E131E","ObjectiveEscapeL|100|191019","ObjectiveChaserL|100|1E131E","ObjectiveDragonL|100|080508","StandupsRL|100|1A101A","StandupsTL|100|1E131E","DropTargetsPL|0|000000","DropTargetsERL|100|090609","LeftHorseshoeKL|100|1E131E","LeftHorseshoeCL|100|1E131E","LeftHorseshoeAL|100|1E131E","LeftHorseshoeLL|100|1E131E","LeftHorseshoeBL|100|1E131E","LeftOrbitBumpersL|100|1E131E","ScoopLockL|100|060406","LeftOrbitLockL|100|1D131D","ScoopONL|100|100A10","CrossbowERL|100|1E131E","CrossbowSNL|100|1E131E","CrossbowIPL|100|1E131E","LeftOrbitGL|0|000000","LeftOrbitI2L|100|1B111B","LeftOrbitI1L|100|1E131E","RightRampEL|100|1E131E","RightRampAL|100|1E131E","RightRampSL|100|1E131E","LeftRampPL|100|1E131E","LeftRampCL|100|1E131E","LeftRampEL|100|1E131E","RightRampPowerupL|100|1E131E","CrossbowPowerupL|100|1E131E","LeftOrbitPowerupL|100|1E131E","LeftRampPowerupL|100|1E131E","HoleShotL|100|010001","RightRampShotL|100|1E131E","LeftHorseshoeShotL|100|030203","CrossbowShotL|0|000000","LeftOrbitShotL|100|1E131E","LeftRampShotL|100|1E131E","GIScoop2|0|000000","GIUpperFlipper|100|1E131E","GIBumperInlanes4|100|010101","GIBumperInlanes3|100|191019","GIBumperInlanes2|100|100A10","GILeftOrbit3|100|1E131E","GILeftOrbit2|100|1E131E","GILeftOrbit1|100|1E131E","GILeftRubbersUpper|100|1E131E","Light001|100|1B111B","GIRightRubbersUpper|100|1E131E","GIRightRubbersLower|100|1E131E","GIRightSlingshotUpper|100|1D121D"), _
Array("ObjectiveFinalJudgmentL|100|1C121C","ObjectiveSniperL|100|1C121C","ObjectiveBlacksmithL|100|1C121C","ObjectiveVikingL|100|1C121C","ObjectiveFinalBossL|100|1C121C","ObjectiveEscapeL|100|1C121C","ObjectiveChaserL|100|1C121C","ObjectiveDragonL|100|1C121C","ObjectiveCastleL|100|080508","StandupsRL|100|1B111B","StandupsTL|100|1C121C","DropTargetsERL|100|040304","LeftHorseshoeKL|100|1C121C","LeftHorseshoeCL|100|1C121C","LeftHorseshoeAL|100|1C121C","LeftHorseshoeLL|100|1C121C","LeftHorseshoeBL|100|1C121C","LeftOrbitBumpersL|100|1C121C","ScoopLockL|0|000000","LeftOrbitLockL|100|170F17","ScoopONL|100|020102","CrossbowERL|100|1C121C","CrossbowSNL|100|1C121C","CrossbowIPL|100|1C121C","LeftOrbitI2L|100|0E090E","LeftOrbitI1L|100|1C121C","RightRampEL|100|1C121C","RightRampAL|100|1C121C","RightRampSL|100|1C121C","LeftRampPL|100|1C121C","LeftRampCL|100|1C121C","LeftRampEL|100|1C121C","RightRampPowerupL|100|1C121C","CrossbowPowerupL|100|1C121C","LeftOrbitPowerupL|100|191019","LeftRampPowerupL|100|1C121C","HoleShotL|0|000000","RightRampShotL|100|1C121C","LeftHorseshoeShotL|100|140D14","LeftOrbitShotL|100|1C121C","LeftRampShotL|100|1C121C","GIUpperFlipper|100|1C121C","GIBumperInlanes4|100|0A070A","GIBumperInlanes3|100|1C121C","GIBumperInlanes2|100|1C121C","GIBumperInlanes1|100|010101","GILeftOrbit3|100|1C121C","GILeftOrbit2|100|1C121C","GILeftOrbit1|100|1C121C","GILeftRubbersUpper|100|1C121C","Light001|100|1C121C","GIRightRubbersUpper|100|1C121C","GIRightRubbersLower|100|1C121C","GIRightSlingshotUpper|100|1C121C"), _
Array("ObjectiveFinalJudgmentL|100|1A101A","ObjectiveSniperL|100|1A101A","ObjectiveBlacksmithL|100|1A101A","ObjectiveVikingL|100|1A101A","ObjectiveFinalBossL|100|1A101A","ObjectiveEscapeL|100|1A101A","ObjectiveChaserL|100|1A101A","ObjectiveDragonL|100|1A101A","ObjectiveCastleL|100|1A101A","StandupsAL|100|020102","StandupsRL|100|1A101A","StandupsTL|100|1A101A","DropTargetsERL|100|020102","LeftHorseshoeKL|100|1A101A","LeftHorseshoeCL|100|1A101A","LeftHorseshoeAL|100|1A101A","LeftHorseshoeLL|100|1A101A","LeftHorseshoeBL|100|1A101A","LeftOrbitBumpersL|100|120B12","LeftOrbitLockL|100|080508","ScoopONL|0|000000","CrossbowERL|100|1A101A","CrossbowSNL|100|1A101A","CrossbowIPL|100|1A101A","LeftOrbitI2L|100|010101","LeftOrbitI1L|100|1A101A","RightRampEL|100|1A101A","RightRampAL|100|1A101A","RightRampSL|100|1A101A","LeftRampPL|100|1A101A","LeftRampCL|100|1A101A","LeftRampEL|100|1A101A","RightRampPowerupL|100|1A101A","CrossbowPowerupL|100|1A101A","LeftOrbitPowerupL|100|0D080D","LeftRampPowerupL|100|1A101A","RightRampShotL|100|1A101A","LeftHorseshoeShotL|100|1A101A","LeftOrbitShotL|100|1A101A","LeftRampShotL|100|1A101A","GIUpperFlipper|100|1A101A","GIBumperInlanes4|100|150D15","GIBumperInlanes3|100|1A101A","GIBumperInlanes2|100|1A101A","GIBumperInlanes1|100|0C070C","GILeftOrbit3|100|1A101A","GILeftOrbit2|100|1A101A","GILeftOrbit1|100|1A101A","GILeftRubbersUpper|100|1A101A","Light001|100|1A101A","GIRightRubbersUpper|100|1A101A","GIRightRubbersLower|100|1A101A","GIRightSlingshotUpper|100|1A101A"), _
Array("ObjectiveFinalJudgmentL|100|180F18","ObjectiveSniperL|100|180F18","ObjectiveBlacksmithL|100|180F18","ObjectiveVikingL|100|180F18","ObjectiveFinalBossL|100|180F18","ObjectiveEscapeL|100|180F18","ObjectiveChaserL|100|180F18","ObjectiveDragonL|100|180F18","ObjectiveCastleL|100|180F18","StandupsAL|100|040204","StandupsRL|100|180F18","StandupsTL|100|180F18","DropTargetsERL|100|010101","RightHorseshoeHL|100|010101","LeftHorseshoeKL|100|180F18","LeftHorseshoeCL|100|180F18","LeftHorseshoeAL|100|180F18","LeftHorseshoeLL|100|180F18","LeftHorseshoeBL|100|180F18","LeftOrbitBumpersL|100|050305","LeftOrbitLockL|0|000000","CrossbowERL|100|180F18","CrossbowSNL|100|180F18","CrossbowIPL|100|180F18","LeftOrbitI2L|0|000000","LeftOrbitI1L|100|180F18","RightRampEL|100|180F18","RightRampAL|100|180F18","RightRampSL|100|180F18","LeftRampPL|100|180F18","LeftRampCL|100|180F18","LeftRampEL|100|180F18","RightRampPowerupL|100|180F18","CrossbowPowerupL|100|100A10","LeftOrbitPowerupL|100|020102","LeftRampPowerupL|100|180F18","RightRampShotL|100|180F18","LeftHorseshoeShotL|100|180F18","LeftOrbitShotL|100|180F18","LeftRampShotL|100|180F18","GIUpperFlipper|100|180F18","GIBumperInlanes4|100|180F18","GIBumperInlanes3|100|180F18","GIBumperInlanes2|100|180F18","GIBumperInlanes1|100|180F18","GILeftOrbit3|100|180F18","GILeftOrbit2|100|180F18","GILeftOrbit1|100|180F18","GILeftRubbersUpper|100|180F18","Light001|100|180F18","GIRightRubbersUpper|100|180F18","GIRightRubbersLower|100|180F18","GIRightSlingshotUpper|100|180F18","GILeftSlingshot|100|010101"), _
Array("ObjectiveFinalJudgmentL|100|150E15","ObjectiveSniperL|100|150E15","ObjectiveBlacksmithL|100|150E15","ObjectiveVikingL|100|150E15","ObjectiveFinalBossL|100|150E15","ObjectiveEscapeL|100|150E15","ObjectiveChaserL|100|150E15","ObjectiveDragonL|100|150E15","ObjectiveCastleL|100|150E15","CrossbowRightStandupL|100|010101","StandupsAL|100|090609","StandupsRL|100|150E15","StandupsTL|100|150E15","DropTargetsERL|0|000000","DropTargetsCHL|100|010101","RightHorseshoeHL|100|080508","RightHorseshoeTL|100|010001","LeftHorseshoeKL|100|150E15","LeftHorseshoeCL|100|150E15","LeftHorseshoeAL|100|150E15","LeftHorseshoeLL|100|150E15","LeftHorseshoeBL|100|150E15","LeftOrbitBumpersL|0|000000","CrossbowERL|100|150E15","CrossbowSNL|100|150E15","CrossbowIPL|100|150E15","LeftOrbitI1L|100|110B11","RightRampEL|100|150E15","RightRampAL|100|150E15","RightRampSL|100|150E15","LeftRampPL|100|150E15","LeftRampCL|100|150E15","LeftRampEL|100|150E15","RightRampPowerupL|100|150E15","CrossbowPowerupL|100|040304","LeftOrbitPowerupL|0|000000","LeftRampPowerupL|100|150E15","RightRampShotL|100|150E15","LeftHorseshoeShotL|100|150E15","LeftOrbitShotL|100|150E15","LeftRampShotL|100|150E15","GIUpperFlipper|100|150E15","GIBumperInlanes4|100|150E15","GIBumperInlanes3|100|150E15","GIBumperInlanes2|100|150E15","GIBumperInlanes1|100|150E15","GILeftOrbit3|100|150E15","GILeftOrbit2|100|150E15","GILeftOrbit1|100|150E15","GILeftRubbersUpper|100|150E15","Light001|100|150E15","GIRightRubbersUpper|100|150E15","GIRightRubbersLower|100|150E15","GIRightSlingshotUpper|100|150E15","GILeftSlingshot|100|150E15","GIRightSlingshotLow|100|070407"), _
Array("ObjectiveFinalJudgmentL|100|130C13","ObjectiveSniperL|100|130C13","ObjectiveBlacksmithL|100|130C13","ObjectiveVikingL|100|130C13","ObjectiveFinalBossL|100|130C13","ObjectiveEscapeL|100|130C13","ObjectiveChaserL|100|130C13","ObjectiveDragonL|100|130C13","ObjectiveCastleL|100|130C13","CrossbowRightStandupL|100|040204","StandupsAL|100|0D080D","StandupsRL|100|130C13","StandupsTL|100|130C13","RightHorseshoeHL|100|100A10","RightHorseshoeTL|100|050305","LeftHorseshoeKL|100|130C13","LeftHorseshoeCL|100|130C13","LeftHorseshoeAL|100|130C13","LeftHorseshoeLL|100|130C13","LeftHorseshoeBL|100|130C13","CrossbowERL|100|0E090E","CrossbowSNL|100|130C13","CrossbowIPL|100|130C13","LeftOrbitI1L|100|060406","RightRampEL|100|130C13","RightRampAL|100|130C13","RightRampSL|100|130C13","LeftRampPL|100|130C13","LeftRampCL|100|130C13","LeftRampEL|100|130C13","RightRampPowerupL|100|130C13","CrossbowPowerupL|0|000000","LeftRampPowerupL|100|130C13","RightRampShotL|100|130C13","LeftHorseshoeShotL|100|130C13","LeftOrbitShotL|100|0D080D","LeftRampShotL|100|130C13","GIUpperFlipper|100|0B070B","GIRightRamp|100|030203","Light003|100|020102","GIBumperInlanes4|100|130C13","GIBumperInlanes3|100|130C13","GIBumperInlanes2|100|130C13","GIBumperInlanes1|100|130C13","GILeftOrbit3|100|130C13","GILeftOrbit2|100|130C13","GILeftOrbit1|100|130C13","GILeftRubbersUpper|100|130C13","Light001|100|130C13","GIRightRubbersUpper|100|130C13","GIRightRubbersLower|100|130C13","GIRightSlingshotUpper|100|130C13","GILeftSlingshot|100|130C13","GIRightSlingshotLow|100|130C13"), _
Array("RightOutlanePoisonL|100|0B070B","ObjectiveFinalJudgmentL|100|110B11","ObjectiveSniperL|100|110B11","ObjectiveBlacksmithL|100|110B11","ObjectiveVikingL|100|110B11","ObjectiveFinalBossL|100|110B11","ObjectiveEscapeL|100|110B11","ObjectiveChaserL|100|110B11","ObjectiveDragonL|100|110B11","ObjectiveCastleL|100|110B11","CrossbowRightStandupL|100|060406","StandupsAL|100|100B10","StandupsRL|100|110B11","StandupsTL|100|110B11","DropTargetsCHL|100|020102","RightHorseshoeHL|100|110B11","RightHorseshoeTL|100|0D080D","RightHorseshoeIL|100|020102","LeftHorseshoeKL|100|110B11","LeftHorseshoeCL|100|110B11","LeftHorseshoeAL|100|110B11","LeftHorseshoeLL|100|110B11","LeftHorseshoeBL|100|110B11","CrossbowERL|100|020202","CrossbowSNL|100|110B11","CrossbowIPL|100|100A10","LeftOrbitI1L|0|000000","RightRampEL|100|110B11","RightRampAL|100|110B11","RightRampSL|100|110B11","LeftRampPL|100|110B11","LeftRampCL|100|110B11","LeftRampEL|100|110B11","RightRampPowerupL|100|110B11","LeftRampPowerupL|100|110B11","RightRampShotL|100|110B11","LeftHorseshoeShotL|100|110B11","LeftOrbitShotL|100|050305","LeftRampShotL|100|110B11","GIUpperFlipper|0|000000","GIRightRamp|100|0B070B","Light003|100|030203","GIBumperInlanes4|100|110B11","GIBumperInlanes3|100|110B11","GIBumperInlanes2|100|110B11","GIBumperInlanes1|100|110B11","GILeftOrbit3|100|110B11","GILeftOrbit2|100|110B11","GILeftOrbit1|100|110B11","GILeftRubbersUpper|100|110B11","Light001|100|110B11","GIRightRubbersUpper|100|110B11","GIRightRubbersLower|100|110B11","GIRightSlingshotUpper|100|110B11","GILeftSlingshot|100|110B11","GIRightSlingshotLow|100|110B11"), _
Array("RightOutlanePoisonL|100|0F090F","ObjectiveFinalJudgmentL|100|0F090F","ObjectiveSniperL|100|0F090F","ObjectiveBlacksmithL|100|0F090F","ObjectiveVikingL|100|0F090F","ObjectiveFinalBossL|100|0F090F","ObjectiveEscapeL|100|0F090F","ObjectiveChaserL|100|0F090F","ObjectiveDragonL|100|0F090F","ObjectiveCastleL|100|0F090F","CrossbowRightStandupL|100|080508","StandupsAL|100|0F090F","StandupsRL|100|0F090F","StandupsTL|100|0F090F","DropTargetsPlusL|100|010001","RightHorseshoeHL|100|0F090F","RightHorseshoeTL|100|0E090E","RightHorseshoeIL|100|070407","LeftHorseshoeKL|100|0F090F","LeftHorseshoeCL|100|0F090F","LeftHorseshoeAL|100|0F090F","LeftHorseshoeLL|100|0F090F","LeftHorseshoeBL|100|0F090F","CrossbowERL|0|000000","CrossbowSNL|100|0E090E","CrossbowIPL|100|050305","RightRampEL|100|0F090F","RightRampAL|100|0F090F","RightRampSL|100|0E090E","LeftRampPL|100|0F090F","LeftRampCL|100|0F090F","LeftRampEL|100|0F090F","RightRampPowerupL|100|0F090F","LeftRampPowerupL|100|0F090F","RightRampShotL|100|0F090F","LeftHorseshoeShotL|100|0F090F","LeftOrbitShotL|100|010001","LeftRampShotL|100|0F090F","GIRightRamp|100|0F090F","Light003|100|060406","GIBumperInlanes4|100|0F090F","GIBumperInlanes3|100|0F090F","GIBumperInlanes2|100|0F090F","GIBumperInlanes1|100|0F090F","GILeftOrbit3|100|0F090F","GILeftOrbit2|100|0F090F","GILeftOrbit1|100|0F090F","GILeftRubbersUpper|100|0F090F","Light001|100|0F090F","GIRightRubbersUpper|100|0F090F","GIRightRubbersLower|100|0F090F","GIRightInlaneUpper|100|010001","GILeftSlingshotLower|100|030203","GIRightSlingshotUpper|100|0F090F","GILeftSlingshot|100|0F090F","GIRightSlingshotLow|100|0F090F"), _
Array("RightOutlanePoisonL|100|0D080D","ObjectiveFinalJudgmentL|100|0D080D","ObjectiveSniperL|100|0D080D","ObjectiveBlacksmithL|100|0D080D","ObjectiveVikingL|100|0D080D","ObjectiveFinalBossL|100|0D080D","ObjectiveEscapeL|100|0D080D","ObjectiveChaserL|100|0D080D","ObjectiveDragonL|100|0D080D","ObjectiveCastleL|100|0D080D","CrossbowRightStandupL|100|090609","StandupsIL|100|020102","StandupsAL|100|0D080D","StandupsRL|100|0D080D","StandupsTL|100|0D080D","DropTargetsPlusL|100|010101","RightHorseshoeHL|100|0D080D","RightHorseshoeTL|100|0D080D","RightHorseshoeIL|100|0B070B","RightHorseshoeML|100|020102","LeftHorseshoeKL|100|0D080D","LeftHorseshoeCL|100|0D080D","LeftHorseshoeAL|100|0D080D","LeftHorseshoeLL|100|0D080D","LeftHorseshoeBL|100|0D080D","CrossbowSNL|100|080508","CrossbowIPL|0|000000","RightRampEL|100|0D080D","RightRampAL|100|0D080D","RightRampSL|100|0A060A","LeftRampPL|100|0D080D","LeftRampCL|100|0D080D","LeftRampEL|100|0D080D","RightRampPowerupL|100|0D080D","LeftRampPowerupL|100|0D080D","RightRampShotL|100|0D080D","LeftHorseshoeShotL|100|0D080D","LeftOrbitShotL|0|000000","LeftRampShotL|100|0D080D","GIRightRamp|100|0D080D","Light003|100|080508","GIBumperInlanes4|100|0D080D","GIBumperInlanes3|100|0D080D","GIBumperInlanes2|100|0D080D","GIBumperInlanes1|100|0D080D","GILeftOrbit3|100|0D080D","GILeftOrbit2|100|0D080D","GILeftOrbit1|100|0D080D","GILeftRubbersUpper|100|0D080D","Light001|100|0D080D","GIRightRubbersUpper|100|0D080D","GIRightRubbersLower|100|0D080D","GIRightInlaneUpper|100|0D080D","GILeftSlingshotLower|100|0D080D","GIRightSlingshotUpper|100|0D080D","GILeftSlingshot|100|0D080D","GIRightSlingshotLow|100|0D080D"), _
Array("RightOutlanePoisonL|100|0A070A","ObjectiveFinalJudgmentL|100|0A070A","ObjectiveSniperL|100|0A070A","ObjectiveBlacksmithL|100|0A070A","ObjectiveVikingL|100|0A070A","ObjectiveFinalBossL|100|0A070A","ObjectiveEscapeL|100|0A070A","ObjectiveChaserL|100|0A070A","ObjectiveDragonL|100|0A070A","ObjectiveCastleL|100|0A070A","CrossbowRightStandupL|100|080608","StandupsIL|100|040304","StandupsAL|100|0A070A","StandupsRL|100|0A070A","StandupsTL|100|0A070A","DropTargetsPlusL|0|000000","DropTargetsCHL|100|010101","RightHorseshoeHL|100|0A070A","RightHorseshoeTL|100|0A070A","RightHorseshoeIL|100|0A070A","RightHorseshoeML|100|070507","LeftHorseshoeKL|100|090609","LeftHorseshoeCL|100|0A070A","LeftHorseshoeAL|100|0A070A","LeftHorseshoeLL|100|0A070A","LeftHorseshoeBL|100|0A070A","CrossbowSNL|0|000000","RightRampEL|100|0A070A","RightRampAL|100|0A070A","RightRampSL|100|060506","LeftRampPL|100|0A070A","LeftRampCL|100|0A070A","LeftRampEL|100|0A070A","RightRampPowerupL|100|090609","LeftRampPowerupL|100|0A070A","RightRampShotL|100|0A070A","LeftHorseshoeShotL|100|0A070A","LeftRampShotL|100|0A070A","GIRightRamp|100|0A070A","Light003|100|080608","GIBumperInlanes4|100|0A070A","GIBumperInlanes3|100|0A070A","GIBumperInlanes2|100|0A070A","GIBumperInlanes1|100|0A070A","GILeftOrbit3|100|0A070A","GILeftOrbit2|100|0A070A","GILeftOrbit1|100|0A070A","GILeftRubbersUpper|100|0A070A","Light001|100|0A070A","GIRightRubbersUpper|100|0A070A","GIRightRubbersLower|100|0A070A","GIRightInlaneUpper|100|0A070A","GIRightInlaneLower|100|090609","GILeftSlingshotLower|100|0A070A","GIRightSlingshotUpper|100|0A070A","GILeftSlingshot|100|0A070A","GIRightSlingshotLow|100|0A070A"), _
Array("RightOutlanePoisonL|100|080508","ObjectiveFinalJudgmentL|100|080508","ObjectiveSniperL|100|080508","ObjectiveBlacksmithL|100|080508","ObjectiveVikingL|100|080508","ObjectiveFinalBossL|100|080508","ObjectiveEscapeL|100|080508","ObjectiveChaserL|100|080508","ObjectiveDragonL|100|080508","ObjectiveCastleL|100|080508","CrossbowRightStandupL|100|070507","StandupsIL|100|050305","StandupsAL|100|080508","StandupsRL|100|080508","StandupsTL|100|070407","DropTargetsCHL|100|010001","RightHorseshoeHL|100|080508","RightHorseshoeTL|100|080508","RightHorseshoeIL|100|080508","RightHorseshoeML|100|080508","RightHorseshoeSL|100|020102","LeftHorseshoeKL|100|060406","LeftHorseshoeCL|100|080508","LeftHorseshoeAL|100|080508","LeftHorseshoeLL|100|080508","LeftHorseshoeBL|100|080508","RightOrbitVL|100|020102","RightRampEL|100|080508","RightRampAL|100|080508","RightRampSL|100|040204","LeftRampPL|100|080508","LeftRampCL|100|080508","LeftRampEL|100|080508","RightRampPowerupL|100|070407","LeftRampPowerupL|100|080508","HoleShotL|100|010001","RightRampShotL|100|080508","LeftHorseshoeShotL|100|080508","LeftRampShotL|100|080508","GIRightRamp|100|080508","Light003|100|080508","GIBumperInlanes4|100|080508","GIBumperInlanes3|100|080508","GIBumperInlanes2|100|080508","GIBumperInlanes1|100|080508","GILeftOrbit3|100|080508","GILeftOrbit2|100|080508","GILeftOrbit1|100|080508","GILeftRubbersUpper|100|080508","Light001|100|080508","GIRightRubbersUpper|100|080508","GIRightRubbersLower|100|080508","GIRightInlaneUpper|100|080508","GIRightInlaneLower|100|080508","GILeftSlingshotLower|100|080508","GIRightSlingshotUpper|100|080508","GILeftSlingshot|100|080508","GIRightSlingshotLow|100|080508"), _
Array("LeftOutlanePoisonL|100|010101","RightOutlanePoisonL|100|060406","ObjectiveFinalJudgmentL|100|060406","ObjectiveSniperL|100|060406","ObjectiveBlacksmithL|100|060406","ObjectiveVikingL|100|060406","ObjectiveFinalBossL|100|060406","ObjectiveEscapeL|100|060406","ObjectiveChaserL|100|060406","ObjectiveDragonL|100|060406","ObjectiveCastleL|100|060406","CrossbowRightStandupL|100|060406","StandupsIL|100|060406","StandupsAL|100|060406","StandupsRL|100|060406","StandupsTL|100|050305","RightHorseshoeHL|100|060406","RightHorseshoeTL|100|060406","RightHorseshoeIL|100|060406","RightHorseshoeML|100|060406","RightHorseshoeSL|100|060406","LeftHorseshoeKL|100|040204","LeftHorseshoeCL|100|060406","LeftHorseshoeAL|100|060406","LeftHorseshoeLL|100|060406","LeftHorseshoeBL|100|060406","RightOrbitVL|100|040204","RightRampEL|100|060406","RightRampAL|100|060406","RightRampSL|100|020102","LeftRampPL|100|060406","LeftRampCL|100|060406","LeftRampEL|100|060406","RightRampPowerupL|100|040304","LeftRampPowerupL|100|060406","HoleShotL|100|010101","RightRampShotL|100|060406","LeftHorseshoeShotL|100|060406","LeftRampShotL|100|060406","GIRightRamp|100|060406","Light003|100|060406","GIBumperInlane2|100|020202","GIBumperInlanes4|100|060406","GIBumperInlanes3|100|060406","GIBumperInlanes2|100|060406","GIBumperInlanes1|100|060406","GILeftOrbit3|100|060406","GILeftOrbit2|100|060406","GILeftOrbit1|100|060406","GILeftRubbersUpper|100|060406","Light001|100|060406","GIRightRubbersUpper|100|010101","GIRightRubbersLower|100|060406","GIRightInlaneUpper|100|060406","GIRightInlaneLower|100|060406","GILeftInlaneUpper|100|010101","GILeftSlingshotLower|100|060406","GIRightSlingshotUpper|100|060406","GILeftSlingshot|100|060406","GIRightSlingshotLow|100|060406"), _
Array("LeftOutlanePoisonL|100|040204","RightOutlanePoisonL|100|040204","ObjectiveFinalJudgmentL|100|040204","ObjectiveSniperL|100|040204","ObjectiveBlacksmithL|100|040204","ObjectiveVikingL|100|040204","ObjectiveFinalBossL|100|040204","ObjectiveEscapeL|100|040204","ObjectiveChaserL|100|040204","ObjectiveDragonL|100|040204","ObjectiveCastleL|100|040204","CrossbowRightStandupL|100|040204","StandupsIL|100|040204","StandupsAL|100|040204","StandupsRL|100|040204","StandupsTL|100|030103","RightHorseshoeHL|100|040204","RightHorseshoeTL|100|040204","RightHorseshoeIL|100|040204","RightHorseshoeML|100|040204","RightHorseshoeSL|100|040204","LeftHorseshoeKL|100|020102","LeftHorseshoeCL|100|040204","LeftHorseshoeAL|100|040204","LeftHorseshoeLL|100|040204","LeftHorseshoeBL|100|040204","RightRampEL|100|040204","RightRampAL|100|040204","RightRampSL|0|000000","LeftRampPL|100|040204","LeftRampCL|100|040204","LeftRampEL|100|040204","RightRampPowerupL|100|020102","LeftRampPowerupL|100|040204","HoleShotL|100|010001","RightRampShotL|100|040204","LeftHorseshoeShotL|100|040204","LeftRampShotL|100|040204","GIRightRamp|100|040204","Light003|100|040204","GIBumperInlane2|100|030203","GIBumperInlanes4|100|040204","GIBumperInlanes3|100|040204","GIBumperInlanes2|100|040204","GIBumperInlanes1|100|040204","GILeftOrbit4|100|010001","GILeftOrbit3|100|040204","GILeftOrbit2|100|040204","GILeftOrbit1|100|020102","GILeftRubbersUpper|100|040204","Light001|100|040204","GIRightRubbersUpper|0|000000","GIRightRubbersLower|100|040204","GIRightInlaneUpper|100|040204","GIRightInlaneLower|100|040204","GILeftInlaneLower|100|040204","GILeftInlaneUpper|100|040204","GILeftSlingshotLower|100|040204","GIRightSlingshotUpper|100|040204","GILeftSlingshot|100|040204","GIRightSlingshotLow|100|040204"), _
Array("LeftOutlanePoisonL|100|020102","RightOutlanePoisonL|100|020102","ObjectiveFinalJudgmentL|100|020102","ObjectiveSniperL|100|020102","ObjectiveBlacksmithL|100|020102","ObjectiveVikingL|100|020102","ObjectiveFinalBossL|100|020102","ObjectiveEscapeL|100|020102","ObjectiveChaserL|100|020102","ObjectiveDragonL|100|020102","ObjectiveCastleL|100|020102","CrossbowRightStandupL|100|020102","StandupsIL|100|020102","StandupsAL|100|020102","StandupsRL|100|020102","StandupsTL|100|010001","DropTargetsCHL|0|000000","RightHorseshoeHL|100|020102","RightHorseshoeTL|100|020102","RightHorseshoeIL|100|020102","RightHorseshoeML|100|020102","RightHorseshoeSL|100|020102","LeftHorseshoeKL|0|000000","LeftHorseshoeCL|100|020102","LeftHorseshoeAL|100|020102","LeftHorseshoeLL|100|020102","LeftHorseshoeBL|100|020102","RightOrbitBumpersL|100|010001","RightOrbitVL|100|020102","RightRampEL|100|020102","RightRampAL|100|020102","LeftRampPL|100|020102","LeftRampCL|100|020102","LeftRampEL|100|020102","RightRampPowerupL|100|010001","LeftRampPowerupL|100|020102","RightHorseshoeShotL|100|010001","RightRampShotL|100|020102","LeftHorseshoeShotL|100|020102","LeftRampShotL|100|020102","GIRightRamp|100|020102","Light003|100|020102","GIBumperInlane2|100|020102","GIBumperInlanes4|100|020102","GIBumperInlanes3|100|020102","GIBumperInlanes2|100|020102","GIBumperInlanes1|100|020102","GILeftOrbit4|100|020102","GILeftOrbit3|100|020102","GILeftOrbit2|100|020102","GILeftOrbit1|0|000000","GILeftRubbersUpper|100|020102","Light001|100|020102","GIRightRubbersLower|100|020102","GIRightInlaneUpper|100|020102","GIRightInlaneLower|100|020102","GILeftInlaneLower|100|020102","GILeftInlaneUpper|100|020102","GILeftSlingshotLower|100|020102","GIRightSlingshotUpper|100|020102","GILeftSlingshot|100|020102","GIRightSlingshotLow|100|020102"), _
Array("LeftOutlanePoisonL|0|000000","RightOutlanePoisonL|0|000000","ObjectiveFinalJudgmentL|0|000000","ObjectiveSniperL|0|000000","ObjectiveBlacksmithL|0|000000","ObjectiveVikingL|0|000000","ObjectiveFinalBossL|0|000000","ObjectiveEscapeL|0|000000","ObjectiveChaserL|0|000000","ObjectiveDragonL|0|000000","ObjectiveCastleL|0|000000","CrossbowRightStandupL|0|000000","StandupsIL|0|000000","StandupsAL|0|000000","StandupsRL|0|000000","StandupsTL|0|000000","RightHorseshoeHL|0|000000","RightHorseshoeTL|0|000000","RightHorseshoeIL|0|000000","RightHorseshoeML|0|000000","RightHorseshoeSL|0|000000","LeftHorseshoeCL|0|000000","LeftHorseshoeAL|0|000000","LeftHorseshoeLL|0|000000","LeftHorseshoeBL|0|000000","RightOrbitBumpersL|0|000000","RightOrbitVL|0|000000","RightRampEL|0|000000","RightRampAL|0|000000","LeftRampPL|0|000000","LeftRampCL|0|000000","LeftRampEL|0|000000","RightRampPowerupL|0|000000","LeftRampPowerupL|0|000000","HoleShotL|0|000000","RightHorseshoeShotL|0|000000","RightRampShotL|0|000000","LeftHorseshoeShotL|0|000000","LeftRampShotL|0|000000","GIRightRamp|0|000000","Light003|0|000000","GIBumperInlane2|0|000000","GIBumperInlanes4|0|000000","GIBumperInlanes3|0|000000","GIBumperInlanes2|0|000000","GIBumperInlanes1|0|000000","GILeftOrbit4|0|000000","GILeftOrbit3|0|000000","GILeftOrbit2|0|000000","GILeftRubbersUpper|0|000000","Light001|0|000000","GIRightRubbersLower|0|000000","GIRightInlaneUpper|0|000000","GIRightInlaneLower|0|000000","GILeftInlaneLower|0|000000","GILeftInlaneUpper|0|000000","GILeftSlingshotLower|0|000000","GIRightSlingshotUpper|0|000000","GILeftSlingshot|0|000000","GIRightSlingshotLow|0|000000"))
lSeqblacksmith_spare.UpdateInterval = 20
lSeqblacksmith_spare.Color = Null
lSeqblacksmith_spare.Repeat = False

'****************************************************************
'	END Custom Light Seqs for Saving Wallden
'****************************************************************

'****************************************************************
'****  ZSHA: VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'****************************************************************

'Ambient (Room light source)
Const AmbientBSFactor = 0.9    '0 To 1, higher is darker
Const AmbientMovement = 2        '1 to 4, higher means more movement as the ball moves left and right
Const offsetX = 0        'Offset x position under ball    (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY = 0        'Offset y position under ball     (for example 5,5 if the light is in the back left corner)
'Dynamic (Table light sources)
Const DynamicBSFactor = 0.95    '0 To 1, higher is darker
Const Wideness = 20    'Sets how wide the dynamic ball shadows can get (20 +5 thinness is technically most accurate for lights at z ~25 hitting a 50 unit ball)
Const Thinness = 5        'Sets minimum as ball moves away from source

' ***                                                        ***

' *** Trim or extend these to *match* the number of balls/primitives/flashers on the table!
Dim objrtx1(6), objrtx2(6)
Dim objBallShadow(6)
Dim OnPF(6)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3,BallShadowA4,BallShadowA5)
Dim DSSources(150), numberofsources', DSGISide(30) 'Adapted for TZ with GI left / GI right

Dim ClearSurface
ClearSurface = True        'Variable for hiding flasher shadow on wire and clear plastic ramps
'Intention is to set this either globally or in a similar manner to RampRolling sounds

'Initialization
DynamicBSInit

Sub DynamicBSInit()
	Dim iii, source
	
	For iii = 0 To tnob - 1                                'Prepares the shadow objects before play begins
		Set objrtx1(iii) = Eval("RtxBallShadow" & iii)
		objrtx1(iii).material = "RtxBallShadow" & iii
		objrtx1(iii).z = 1 + iii / 1000 + 0.01            'Separate z for layering without clipping
		objrtx1(iii).visible = 0
		
		Set objrtx2(iii) = Eval("RtxBall2Shadow" & iii)
		objrtx2(iii).material = "RtxBallShadow2_" & iii
		objrtx2(iii).z = 1 + iii / 1000 + 0.02
		objrtx2(iii).visible = 0
		
		Set objBallShadow(iii) = Eval("BallShadow" & iii)
		objBallShadow(iii).material = "BallShadow" & iii
		UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
		objBallShadow(iii).Z = 1 + iii / 1000 + 0.04
		objBallShadow(iii).visible = 0
		
		BallShadowA(iii).Opacity = 100 * AmbientBSFactor
		BallShadowA(iii).visible = 0
	Next
	
	iii = 0
	
	For Each Source In DynamicSources
		DSSources(iii) = Array(Source.x, Source.y)
		'        If Instr(Source.name , "Left") > 0 Then DSGISide(iii) = 0 Else DSGISide(iii) = 1    'Adapted for TZ with GI left / GI right
		iii = iii + 1
	Next
	numberofsources = iii
End Sub


Sub BallOnPlayfieldNow(yeh, num)        'Only update certain things once, save some cycles
	If yeh Then
		OnPF(num) = True
		'        debug.print "Back on PF"
		UpdateMaterial objBallShadow(num).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
		objBallShadow(num).size_x = 5
		objBallShadow(num).size_y = 4.5
		objBallShadow(num).visible = 1
		BallShadowA(num).visible = 0
	Else
		OnPF(num) = False
		'        debug.print "Leaving PF"
		If Not ClearSurface Then
			BallShadowA(num).visible = 1
			objBallShadow(num).visible = 0
		Else
			objBallShadow(num).visible = 1
		End If
	End If
End Sub

Sub DynamicBSUpdate
	Dim falloff
	falloff = 150 'Max distance to light sources, can be changed dynamically if you have a reason
	Dim ShadowOpacity1, ShadowOpacity2
	Dim s, LSd, iii
	Dim dist1, dist2, src1, src2
	'    Dim gBOT: gBOT=getballs    'Uncomment if you're deleting balls - Don't do it! #SaveTheBalls
	
	'Hide shadow of deleted balls
	For s = UBound(gBOT) + 1 To tnob - 1
		objrtx1(s).visible = 0
		objrtx2(s).visible = 0
		objBallShadow(s).visible = 0
		BallShadowA(s).visible = 0
	Next
	
	If UBound(gBOT) < lob Then Exit Sub        'No balls in play, exit
	
	'The Magic happens now
	For s = lob To UBound(gBOT)
		
		' *** Normal "ambient light" ball shadow
		'Layered from top to bottom. If you had an upper pf at for example 80 units and ramps even above that, your segments would be z>110; z<=110 And z>100; z<=100 And z>30; z<=30 And z>20; Else invisible
		
		If AmbientBallShadowOn = 1 Then            'Primitive shadow on playfield, flasher shadow in ramps
			If gBOT(s).Z > 30 Then                            'The flasher follows the ball up ramps while the primitive is on the pf
				If OnPF(s) Then BallOnPlayfieldNow False, s        'One-time update
				
				If Not ClearSurface Then                            'Don't show this shadow on plastic or wire ramps (table-wide variable, for now)
					BallShadowA(s).X = gBOT(s).X + offsetX
					BallShadowA(s).Y = gBOT(s).Y + BallSize / 5
					BallShadowA(s).height = gBOT(s).z - BallSize / 4        'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
				Else
					If gBOT(s).X < tablewidth / 2 Then
						objBallShadow(s).X = ((gBOT(s).X) - (Ballsize / 10) + ((gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement))) + offsetX + 5
					Else
						objBallShadow(s).X = ((gBOT(s).X) + (Ballsize / 10) + ((gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement))) + offsetX - 5
					End If
					objBallShadow(s).Y = gBOT(s).Y + BallSize / 10 + offsetY
					objBallShadow(s).size_x = 5 * ((gBOT(s).Z + BallSize) / 80)            'Shadow gets larger and more diffuse as it moves up
					objBallShadow(s).size_y = 4.5 * ((gBOT(s).Z + BallSize) / 80)
					UpdateMaterial objBallShadow(s).material,1,0,0,0,0,0,AmbientBSFactor * (30 / (gBOT(s).Z)),RGB(0,0,0),0,0,False,True,0,0,0,0
				End If
				
			ElseIf gBOT(s).Z <= 30 And gBOT(s).Z > 20 Then    'On pf, primitive only
				If Not OnPF(s) Then BallOnPlayfieldNow True, s
				
				If gBOT(s).X < tablewidth / 2 Then
					objBallShadow(s).X = ((gBOT(s).X) - (Ballsize / 10) + ((gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement))) + offsetX + 5
				Else
					objBallShadow(s).X = ((gBOT(s).X) + (Ballsize / 10) + ((gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement))) + offsetX - 5
				End If
				objBallShadow(s).Y = gBOT(s).Y + offsetY
				'                objBallShadow(s).Z = gBOT(s).Z + s/1000 + 0.04        'Uncomment (and adjust If/Elseif height logic) if you want the primitive shadow on an upper/split pf
				
			Else                                                'Under pf, no shadows
				objBallShadow(s).visible = 0
				BallShadowA(s).visible = 0
			End If
			
		ElseIf AmbientBallShadowOn = 2 Then    'Flasher shadow everywhere
			If gBOT(s).Z > 30 Then                            'In a ramp
				If Not ClearSurface Then                            'Don't show this shadow on plastic or wire ramps (table-wide variable, for now)
					BallShadowA(s).X = gBOT(s).X + offsetX
					BallShadowA(s).Y = gBOT(s).Y + BallSize / 5
					BallShadowA(s).height = gBOT(s).z - BallSize / 4        'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
				Else
					BallShadowA(s).X = gBOT(s).X + offsetX
					BallShadowA(s).Y = gBOT(s).Y + offsetY
				End If
			ElseIf gBOT(s).Z <= 30 And gBOT(s).Z > 20 Then    'On pf
				BallShadowA(s).visible = 1
				If gBOT(s).X < tablewidth / 2 Then
					BallShadowA(s).X = ((gBOT(s).X) - (Ballsize / 10) + ((gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement))) + offsetX + 5
				Else
					BallShadowA(s).X = ((gBOT(s).X) + (Ballsize / 10) + ((gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement))) + offsetX - 5
				End If
				BallShadowA(s).Y = gBOT(s).Y + offsetY
				BallShadowA(s).height = 0.1
			Else                                            'Under pf
				BallShadowA(s).visible = 0
			End If
		End If
		
		' *** Dynamic shadows
		If DynamicBallShadowsOn Then
			If gBOT(s).Z < 30 And gBOT(s).X < 850 Then    'Parameters for where the shadows can show, here they are not visible above the table (no upper pf) or in the plunger lane
				dist1 = falloff
				
				dist2 = falloff
				For iii = 0 To numberofsources - 1 ' Search the 2 nearest influencing lights
					LSd = Distance(gBOT(s).x, gBOT(s).y, DSSources(iii)(0), DSSources(iii)(1)) 'Calculating the Linear distance to the Source
					If LSd < falloff And Lampz.State(0) > 0 Then ' Modified for Lampz
						'                    If LSd < dist2 And ((DSGISide(iii) = 0 And Lampz.State(100)>0) Or (DSGISide(iii) = 1 And Lampz.State(104)>0)) Then    'Adapted for TZ with GI left / GI right
						dist2 = dist1
						dist1 = LSd
						src2 = src1
						src1 = iii
					End If
				Next
				ShadowOpacity1 = 0
				If dist1 < falloff Then
					objrtx1(s).visible = 1
					objrtx1(s).X = gBOT(s).X
					objrtx1(s).Y = gBOT(s).Y
					'objrtx1(s).Z = gBOT(s).Z - 25 + s/1000 + 0.01 'Uncomment if you want to add shadows to an upper/lower pf
					objrtx1(s).rotz = AnglePP(DSSources(src1)(0), DSSources(src1)(1), gBOT(s).X, gBOT(s).Y) + 90
					ShadowOpacity1 = 1 - dist1 / falloff
					objrtx1(s).size_y = Wideness * ShadowOpacity1 + Thinness
					UpdateMaterial objrtx1(s).material,1,0,0,0,0,0,ShadowOpacity1 * DynamicBSFactor ^ 3,RGB(0,0,0),0,0,False,True,0,0,0,0
				Else
					objrtx1(s).visible = 0
				End If
				ShadowOpacity2 = 0
				If dist2 < falloff Then
					objrtx2(s).visible = 1
					objrtx2(s).X = gBOT(s).X
					objrtx2(s).Y = gBOT(s).Y + offsetY
					'objrtx2(s).Z = gBOT(s).Z - 25 + s/1000 + 0.02 'Uncomment if you want to add shadows to an upper/lower pf
					objrtx2(s).rotz = AnglePP(DSSources(src2)(0), DSSources(src2)(1), gBOT(s).X, gBOT(s).Y) + 90
					ShadowOpacity2 = 1 - dist2 / falloff
					objrtx2(s).size_y = Wideness * ShadowOpacity2 + Thinness
					UpdateMaterial objrtx2(s).material,1,0,0,0,0,0,ShadowOpacity2 * DynamicBSFactor ^ 3,RGB(0,0,0),0,0,False,True,0,0,0,0
				Else
					objrtx2(s).visible = 0
				End If
				If AmbientBallShadowOn = 1 Then
					'Fades the ambient shadow (primitive only) when it's close to a light
					UpdateMaterial objBallShadow(s).material,1,0,0,0,0,0,AmbientBSFactor * (1 - max(ShadowOpacity1, ShadowOpacity2)),RGB(0,0,0),0,0,False,True,0,0,0,0
				Else
					BallShadowA(s).Opacity = 100 * AmbientBSFactor * (1 - max(ShadowOpacity1, ShadowOpacity2))
				End If
			Else 'Hide dynamic shadows everywhere else, just in case
				objrtx2(s).visible = 0
				objrtx1(s).visible = 0
			End If
		End If
	Next
End Sub

'****************************************************************
'****  END VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'****************************************************************

Sub FlipperVisualUpdate
	FlipperLSh.RotZ = LeftFlipper.CurrentAngle
	FlipperRSh.RotZ = RightFlipper.CurrentAngle
	FlipperUSh.RotZ = UpperFlipper.CurrentAngle
End Sub

'******************************************************
'*****   ZDOM: FLUPPER DOMES
'******************************************************
' Based on FlupperDoms2.2

' What you need in your table to use these flashers:
' Open this table and your table both in VPX
' Export all the materials domebasemat, Flashermaterial0 - 20 and import them in your table
' Export all textures (images) starting with the name "dome" and "ronddome" and import them into your table with the same names
' Export all textures (images) starting with the name "flasherbloom" and import them into your table with the same names
' Copy a set of 4 objects flasherbase, flasherlit, flasherlight and flasherflash from layer 7 to your table
' If you duplicate the four objects for a new flasher dome, be sure that they all end with the same number (in the 0-20 range)
' Copy the flasherbloom flashers from layer 10 to your table. you will need to make one per flasher dome that you plan to make
' Select the correct flasherbloom texture for each flasherbloom flasher, per flasher dome
' Copy the script below

' Place your flasher base primitive where you want the flasher located on your Table
' Then run InitFlasher in the script with the number of your flasher objects and the color of the flasher.  This will align the flasher object, light object, and
' flasher lit primitive.  It will also assign the appropriate flasher bloom images to the flasher bloom object.
'
' Example: InitFlasher 1, "green"
'
' Color Options: "blue", "green", "red", "purple", "yellow", "white", and "orange"

' You can use the RotateFlasher call to align the Rotz/ObjRotz of the flasher primitives with "handles".  Don't set those values in the editor,
' call the RotateFlasher sub instead (this call will likely crash VP if it's call for the flasher primitives without "handles")
'
' Example: RotateFlasher 1, 180         'where 1 is the flasher number and 180 is the angle of Z rotation

' For flashing the flasher use in the script: "ObjLevel(1) = 1 : FlasherFlash1_Timer"
' This should also work for flashers with variable flash levels from the rom, just use ObjLevel(1) = xx from the rom (in the range 0-1)
'
' Notes (please read!!):
' - Setting TestFlashers = 1 (below in the ScriptsDirectory) will allow you to see how the flasher objects are aligned (need the targetflasher image imported to your table)
' - The rotation of the primitives with "handles" is done with a script command, not on the primitive itself (see RotateFlasher below)
' - Color of the objects are set in the script, not on the primitive itself
' - Screws are optional to copy and position manually
' - If your table is not named "Table1" then change the name below in the script
' - Every flasher uses its own material (Flashermaterialxx), do not use it for anything else
' - Lighting > Bloom Strength affects how the flashers look, do not set it too high
' - Change RotY and RotX of flasherbase only when having a flasher something other then parallel to the playfield
' - Leave RotX of the flasherflash object to -45; this makes sure that the flash effect is visible in FS and DT
' - If you want to resize a flasher, be sure to resize flasherbase, flasherlit and flasherflash with the same percentage
' - If you think that the flasher effects are too bright, change flasherlightintensity and/or flasherflareintensity below

' Some more notes for users of the v1 flashers and/or JP's fading lights routines:
' - Delete all textures/primitives/script/materials in your table from the v1 flashers and scripts before you start; they don't mix well with v2
' - Remove flupperflash(m) routines if you have them; they do not work with this new script
' - Do not try to mix this v2 script with the JP fading light routine (that is making it too complicated), just use the example script below

' example script for rom based tables (non modulated):

' SolCallback(25)="FlashRed"
'
' Sub FlashRed(flstate)
'    If Flstate Then
'        ObjTargetLevel(1) = 1
'    Else
'        ObjTargetLevel(1) = 0
'    End If
'   FlasherFlash1_Timer
' End Sub

' example script for rom based tables (modulated):

' SolModCallback(25)="FlashRed"
'
' Sub FlashRed(level)
'    ObjTargetLevel(1) = level/255 : FlasherFlash1_Timer
' End Sub

Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherBloomIntensity, FlasherOffBrightness

' *********************************************************************
TestFlashers = 0                ' *** set this to 1 to check position of flasher object             ***
Set TableRef = Table1           ' *** change this, if your table has another name                   ***
FlasherLightIntensity = 0.2        ' *** lower this, if the VPX lights are too bright (i.e. 0.1)        ***
FlasherFlareIntensity = 0.3        ' *** lower this, if the flares are too bright (i.e. 0.1)            ***
FlasherBloomIntensity = 0.3        ' *** lower this, if the blooms are too bright (i.e. 0.1)            ***
FlasherOffBrightness = 0.3        ' *** brightness of the flasher dome when switched off (range 0-2)    ***
' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20), ObjTargetLevel(20)
'Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height
'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"


Sub InitFlasher(nr, col)
	' store all objects in an array for use in FlashFlasher subroutine
	Set objbase(nr) = Eval("Flasherbase" & nr)
	Set objlit(nr) = Eval("Flasherlit" & nr)
	Set objflasher(nr) = Eval("Flasherflash" & nr)
	Set objlight(nr) = Eval("Flasherlight" & nr)
	Set objbloom(nr) = Eval("Flasherbloom" & nr)
	' If the flasher is parallel to the playfield, rotate the VPX flasher object for POV and place it at the correct height
	If objbase(nr).RotY = 0 Then
		objbase(nr).ObjRotZ = Atn( (tablewidth / 2 - objbase(nr).x) / (objbase(nr).y - tableheight * 1.1)) * 180 / 3.14159
		objflasher(nr).RotZ = objbase(nr).ObjRotZ
		objflasher(nr).height = objbase(nr).z + 40
	End If
	' set all effects to invisible and move the lit primitive at the same position and rotation as the base primitive
	objlight(nr).IntensityScale = 0
	objlit(nr).visible = 0
	objlit(nr).material = "Flashermaterial" & nr
	objlit(nr).RotX = objbase(nr).RotX
	objlit(nr).RotY = objbase(nr).RotY
	objlit(nr).RotZ = objbase(nr).RotZ
	objlit(nr).ObjRotX = objbase(nr).ObjRotX
	objlit(nr).ObjRotY = objbase(nr).ObjRotY
	objlit(nr).ObjRotZ = objbase(nr).ObjRotZ
	objlit(nr).x = objbase(nr).x
	objlit(nr).y = objbase(nr).y
	objlit(nr).z = objbase(nr).z
	objbase(nr).BlendDisableLighting = FlasherOffBrightness
	
	'rothbauerw
	'Adjust the position of the flasher object to align with the flasher base.
	'Comment out these lines if you want to manually adjust the flasher object
	If objbase(nr).roty > 135 Then
		objflasher(nr).y = objbase(nr).y + 50
		objflasher(nr).height = objbase(nr).z + 20
	Else
		objflasher(nr).y = objbase(nr).y + 20
		objflasher(nr).height = objbase(nr).z + 50
	End If
	objflasher(nr).x = objbase(nr).x
	
	'rothbauerw
	'Adjust the position of the light object to align with the flasher base.
	'Comment out these lines if you want to manually adjust the flasher object
	objlight(nr).x = objbase(nr).x
	objlight(nr).y = objbase(nr).y
	objlight(nr).bulbhaloheight = objbase(nr).z - 10
	
	'rothbauerw
	'Assign the appropriate bloom image basked on the location of the flasher base
	'Comment out these lines if you want to manually assign the bloom images
	Dim xthird, ythird
	xthird = tablewidth / 3
	ythird = tableheight / 3
	
	If objbase(nr).x >= xthird And objbase(nr).x <= xthird * 2 Then
		objbloom(nr).imageA = "flasherbloomCenter"
		objbloom(nr).imageB = "flasherbloomCenter"
	ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird Then
		objbloom(nr).imageA = "flasherbloomUpperLeft"
		objbloom(nr).imageB = "flasherbloomUpperLeft"
	ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird Then
		objbloom(nr).imageA = "flasherbloomUpperRight"
		objbloom(nr).imageB = "flasherbloomUpperRight"
	ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird * 2 Then
		objbloom(nr).imageA = "flasherbloomCenterLeft"
		objbloom(nr).imageB = "flasherbloomCenterLeft"
	ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird * 2 Then
		objbloom(nr).imageA = "flasherbloomCenterRight"
		objbloom(nr).imageB = "flasherbloomCenterRight"
	ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird * 3 Then
		objbloom(nr).imageA = "flasherbloomLowerLeft"
		objbloom(nr).imageB = "flasherbloomLowerLeft"
	ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird * 3 Then
		objbloom(nr).imageA = "flasherbloomLowerRight"
		objbloom(nr).imageB = "flasherbloomLowerRight"
	End If
	
	' set the texture and color of all objects
	Select Case objbase(nr).image
		Case "dome2basewhite"
			objbase(nr).image = "dome2base" & col
			objlit(nr).image = "dome2lit" & col
			
		Case "ronddomebasewhite"
			objbase(nr).image = "ronddomebase" & col
			objlit(nr).image = "ronddomelit" & col
		Case "domeearbasewhite"
			objbase(nr).image = "domeearbase" & col
			objlit(nr).image = "domeearlit" & col
	End Select
	If TestFlashers = 0 Then
		objflasher(nr).imageA = "domeflashwhite"
		objflasher(nr).visible = 0
	End If
	Select Case col
		Case "blue"
			objlight(nr).color = RGB(4,120,255)
			objflasher(nr).color = RGB(200,255,255)
			objbloom(nr).color = RGB(4,120,255)
			objlight(nr).intensity = 5000
		Case "green"
			objlight(nr).color = RGB(12,255,4)
			objflasher(nr).color = RGB(12,255,4)
			objbloom(nr).color = RGB(12,255,4)
		Case "red"
			objlight(nr).color = RGB(255,32,4)
			objflasher(nr).color = RGB(255,32,4)
			objbloom(nr).color = RGB(255,32,4)
		Case "purple"
			objlight(nr).color = RGB(230,49,255)
			objflasher(nr).color = RGB(255,64,255)
			objbloom(nr).color = RGB(230,49,255)
		Case "yellow"
			objlight(nr).color = RGB(200,173,25)
			objflasher(nr).color = RGB(255,200,50)
			objbloom(nr).color = RGB(200,173,25)
		Case "white"
			objlight(nr).color = RGB(255,240,150)
			objflasher(nr).color = RGB(100,86,59)
			objbloom(nr).color = RGB(255,240,150)
		Case "orange"
			objlight(nr).color = RGB(255,70,0)
			objflasher(nr).color = RGB(255,70,0)
			objbloom(nr).color = RGB(255,70,0)
	End Select
	objlight(nr).colorfull = objlight(nr).color
	If TableRef.ShowDT And ObjFlasher(nr).RotX =  - 45 Then
		objflasher(nr).height = objflasher(nr).height - 20 * ObjFlasher(nr).y / tableheight
		ObjFlasher(nr).y = ObjFlasher(nr).y + 10
	End If
End Sub

Sub RotateFlasher(nr, angle)
	angle = ((angle + 360 - objbase(nr).ObjRotZ) Mod 180) / 30
	objbase(nr).showframe(angle)
	objlit(nr).showframe(angle)
End Sub

Sub FlashFlasher(nr)
	If Not objflasher(nr).TimerEnabled Then
		objflasher(nr).TimerEnabled = True
		objflasher(nr).visible = 1
		objbloom(nr).visible = 1
		objlit(nr).visible = 1
	End If
	objflasher(nr).opacity = 1000 * FlasherFlareIntensity * ObjLevel(nr) ^ 2.5
	objbloom(nr).opacity = 100 * FlasherBloomIntensity * ObjLevel(nr) ^ 2.5
	objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr) ^ 3
	objbase(nr).BlendDisableLighting = FlasherOffBrightness + 10 * ObjLevel(nr) ^ 3
	objlit(nr).BlendDisableLighting = 10 * ObjLevel(nr) ^ 2
	UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
	If Round(ObjTargetLevel(nr),1) > Round(ObjLevel(nr),1) Then
		ObjLevel(nr) = ObjLevel(nr) + 0.3
		If ObjLevel(nr) > 1 Then ObjLevel(nr) = 1
	ElseIf Round(ObjTargetLevel(nr),1) < Round(ObjLevel(nr),1) Then
		ObjLevel(nr) = ObjLevel(nr) * 0.85 - 0.01
		If ObjLevel(nr) < 0 Then ObjLevel(nr) = 0
	Else
		ObjLevel(nr) = Round(ObjTargetLevel(nr),1)
		objflasher(nr).TimerEnabled = False
	End If
	'ObjLevel(nr) = ObjLevel(nr) * 0.9 - 0.01
	If ObjLevel(nr) < 0 Then
		objflasher(nr).TimerEnabled = False
		objflasher(nr).visible = 0
		objbloom(nr).visible = 0
		objlit(nr).visible = 0
	End If
End Sub



'******************************************************
'******  END FLUPPER DOMES
'******************************************************

'******************************************************
'******  FLUPPER BUMPERS
'******************************************************
' Based on FlupperBumpers 0.145 final

' Explanation of how these bumpers work:
' There are 10 elements involved per bumper:
' - the shadow of the bumper ( a vpx flasher object)
' - the bumper skirt (primitive)
' - the bumperbase (primitive)
' - a vpx light which colors everything you can see through the bumpertop
' - the bulb (primitive)
' - another vpx light which lights up everything around the bumper
' - the bumpertop (primitive)
' - the VPX bumper object
' - the bumper screws (primitive)
' - the bulb highlight VPX flasher object
' All elements have a special name with the number of the bumper at the end, this is necessary for the fading routine and the initialisation.
' For the bulb and the bumpertop there is a unique material as well per bumpertop.
' To use these bumpers you have to first copy all 10 elements to your table.
' Also export the textures (images) with names that start with "Flbumper" and "Flhighlight" and materials with names that start with "bumper".
' Make sure that all the ten objects are aligned on center, if possible with the exact same x,y coordinates
' After that copy the script (below); also copy the BumperTimer vpx object to your table
' Every bumper needs to be initialised with the FlInitBumper command, see example below;
' Colors available are red, white, blue, orange, yellow, green, purple and blacklight.
' In a GI subroutine you can then call set the bumperlight intensity with the "FlBumperFadeTarget(nr) = value" command
' where nr is the number of the bumper, value is between 0 (off) and 1 (full on) (so you can also use 0.3 0.4 etc).

' Notes:
' - There is only one color for the disk; you can photoshop it to a different color
' - The bumpertops are angle independent up to a degree; my estimate is -45 to + 45 degrees horizontally, 0 (topview) to 70-80 degrees (frontview)
' - I built in correction for the day-night slider; this might not work perfectly, depending on your table lighting
' - These elements, textures and materials do NOT integrate with any of the lighting routines I have seen in use in many VPX tables
'   (just find the GI handling routine and insert the FlBumperFadeTarget statement)
' - If you want to use VPX native bumperdisks just copy my bumperdisk but make it invisible



' prepare some global vars to dim/brighten objects when using day-night slider
Dim DayNightAdjust , DNA30, DNA45, DNA90
If NightDay < 10 Then
	DNA30 = 0
	DNA45 = (NightDay - 10) / 20
	DNA90 = 0
	DayNightAdjust = 0.4
Else
	DNA30 = (NightDay - 10) / 30
	DNA45 = (NightDay - 10) / 45
	DNA90 = (NightDay - 10) / 90
	DayNightAdjust = NightDay / 25
End If

Dim FlBumperFadeActual(6), FlBumperFadeTarget(6), FlBumperColor(6), FlBumperTop(6), FlBumperSmallLight(6), Flbumperbiglight(6)
Dim FlBumperDisk(6), FlBumperBase(6), FlBumperBulb(6), FlBumperscrews(6), FlBumperActive(6), FlBumperHighlight(6)
Dim FlBumperControlLight(6) ' CUSTOM LINE (Lampz support)
Dim cnt
For cnt = 1 To 6
	FlBumperActive(cnt) = False
Next

Sub FlInitBumper(nr, col)
	FlBumperActive(nr) = True
	' store all objects in an array for use in FlFadeBumper subroutine
	FlBumperFadeActual(nr) = 1
	FlBumperFadeTarget(nr) = 1.1
	FlBumperColor(nr) = col
	Set FlBumperTop(nr) = Eval("bumpertop" & nr)
	FlBumperTop(nr).material = "bumpertopmat" & nr
	Set FlBumperSmallLight(nr) = Eval("bumpersmalllight" & nr)
	Set Flbumperbiglight(nr) = Eval("bumperbiglight" & nr)
	Set FlBumperDisk(nr) = Eval("bumperdisk" & nr)
	Set FlBumperBase(nr) = Eval("bumperbase" & nr)
	Set FlBumperBulb(nr) = Eval("bumperbulb" & nr)
	FlBumperBulb(nr).material = "bumperbulbmat" & nr
	Set FlBumperscrews(nr) = Eval("bumperscrews" & nr)
	FlBumperscrews(nr).material = "bumperscrew" & col
	Set FlBumperHighlight(nr) = Eval("bumperhighlight" & nr)
	Set FlBumperControlLight(nr) = Eval("bumperlightcontrol" & nr) ' CUSTOM LINE (Lampz support)
	' set the color for the two VPX lights
	Select Case col
		Case "red"
			FlBumperSmallLight(nr).color = RGB(255,4,0)
			FlBumperSmallLight(nr).colorfull = RGB(255,24,0)
			FlBumperBigLight(nr).color = RGB(255,32,0)
			FlBumperBigLight(nr).colorfull = RGB(255,32,0)
			FlBumperHighlight(nr).color = RGB(64,255,0)
			FlBumperSmallLight(nr).BulbModulateVsAdd = 0.98
			FlBumperSmallLight(nr).TransmissionScale = 0
		Case "blue"
			FlBumperBigLight(nr).color = RGB(32,80,255)
			FlBumperBigLight(nr).colorfull = RGB(32,80,255)
			FlBumperSmallLight(nr).color = RGB(0,80,255)
			FlBumperSmallLight(nr).colorfull = RGB(0,80,255)
			FlBumperSmallLight(nr).TransmissionScale = 0
			MaterialColor "bumpertopmat" & nr, RGB(8,120,255)
			FlBumperHighlight(nr).color = RGB(255,16,8)
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1
		Case "green"
			FlBumperSmallLight(nr).color = RGB(8,255,8)
			FlBumperSmallLight(nr).colorfull = RGB(8,255,8)
			FlBumperBigLight(nr).color = RGB(32,255,32)
			FlBumperBigLight(nr).colorfull = RGB(32,255,32)
			FlBumperHighlight(nr).color = RGB(255,32,255)
			MaterialColor "bumpertopmat" & nr, RGB(16,255,16)
			FlBumperSmallLight(nr).TransmissionScale = 0.005
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1
		Case "orange"
			FlBumperHighlight(nr).color = RGB(255,130,255)
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1
			FlBumperSmallLight(nr).TransmissionScale = 0
			FlBumperSmallLight(nr).color = RGB(255,130,0)
			FlBumperSmallLight(nr).colorfull = RGB (255,90,0)
			FlBumperBigLight(nr).color = RGB(255,190,8)
			FlBumperBigLight(nr).colorfull = RGB(255,190,8)
		Case "white"
			FlBumperBigLight(nr).color = RGB(255,230,190)
			FlBumperBigLight(nr).colorfull = RGB(255,230,190)
			FlBumperHighlight(nr).color = RGB(255,180,100)
			
			FlBumperSmallLight(nr).TransmissionScale = 0
			FlBumperSmallLight(nr).BulbModulateVsAdd = 0.99
		Case "blacklight"
			FlBumperBigLight(nr).color = RGB(32,32,255)
			FlBumperBigLight(nr).colorfull = RGB(32,32,255)
			FlBumperHighlight(nr).color = RGB(48,8,255)
			
			FlBumperSmallLight(nr).TransmissionScale = 0
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1
		Case "yellow"
			FlBumperSmallLight(nr).color = RGB(255,230,4)
			FlBumperSmallLight(nr).colorfull = RGB(255,230,4)
			FlBumperBigLight(nr).color = RGB(255,240,50)
			FlBumperBigLight(nr).colorfull = RGB(255,240,50)
			FlBumperHighlight(nr).color = RGB(255,255,220)
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1
			FlBumperSmallLight(nr).TransmissionScale = 0
		Case "purple"
			FlBumperBigLight(nr).color = RGB(80,32,255)
			FlBumperBigLight(nr).colorfull = RGB(80,32,255)
			FlBumperSmallLight(nr).color = RGB(80,32,255)
			FlBumperSmallLight(nr).colorfull = RGB(80,32,255)
			FlBumperSmallLight(nr).TransmissionScale = 0
			
			FlBumperHighlight(nr).color = RGB(32,64,255)
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1
	End Select
End Sub

Sub FlFadeBumper(nr, Z)
	FlBumperBase(nr).BlendDisableLighting = 0.5 * DayNightAdjust
	'    UpdateMaterial(string, float wrapLighting, float roughness, float glossyImageLerp, float thickness, float edge, float edgeAlpha, float opacity,
	'               OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive,
	'               float elasticity, float elasticityFalloff, float friction, float scatterAngle) - updates all parameters of a material
	FlBumperDisk(nr).BlendDisableLighting = (0.5 - Z * 0.3 ) * DayNightAdjust
	
	Select Case FlBumperColor(nr)
			
		Case "blue"
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(38 - 24 * Z,130 - 98 * Z,255), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 20 + 500 * Z / (0.5 + DNA30)
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
			FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 5000 * (0.03 * Z + 0.97 * Z ^ 3)
			Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 10000 * (Z ^ 3) / (0.5 + DNA90)
			
		Case "green"
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(16 + 16 * Sin(Z * 3.14),255,16 + 16 * Sin(Z * 3.14)), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 10 + 150 * Z / (1 + DNA30)
			FlBumperTop(nr).BlendDisableLighting = 2 * DayNightAdjust + 20 * Z
			FlBumperBulb(nr).BlendDisableLighting = 7 * DayNightAdjust + 6000 * (0.03 * Z + 0.97 * Z ^ 10)
			Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 6000 * (Z ^ 3) / (1 + DNA90)
			
		Case "red"
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255, 16 - 11 * Z + 16 * Sin(Z * 3.14),0), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 17 + 100 * Z / (1 + DNA30 ^ 2)
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 18 * Z / (1 + DNA90)
			FlBumperBulb(nr).BlendDisableLighting = 20 * DayNightAdjust + 9000 * (0.03 * Z + 0.97 * Z ^ 10)
			Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 2000 * (Z ^ 3) / (1 + DNA90)
			MaterialColor "bumpertopmat" & nr, RGB(255,20 + Z * 4,8 - Z * 8)
			
		Case "orange"
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255, 100 - 22 * z + 16 * Sin(Z * 3.14),Z * 32), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 17 + 250 * Z / (1 + DNA30 ^ 2)
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z / (1 + DNA90)
			FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 2500 * (0.03 * Z + 0.97 * Z ^ 10)
			Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 4000 * (Z ^ 3) / (1 + DNA90)
			MaterialColor "bumpertopmat" & nr, RGB(255,100 + Z * 50, 0)
			
		Case "white"
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255,230 - 100 * Z, 200 - 150 * Z), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 20 + 180 * Z / (1 + DNA30)
			FlBumperTop(nr).BlendDisableLighting = 5 * DayNightAdjust + 30 * Z
			FlBumperBulb(nr).BlendDisableLighting = 18 * DayNightAdjust + 3000 * (0.03 * Z + 0.97 * Z ^ 10)
			Flbumperbiglight(nr).intensity = 8 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 1000 * (Z ^ 3) / (1 + DNA90)
			FlBumperSmallLight(nr).color = RGB(255,255 - 20 * Z,255 - 65 * Z)
			FlBumperSmallLight(nr).colorfull = RGB(255,255 - 20 * Z,255 - 65 * Z)
			MaterialColor "bumpertopmat" & nr, RGB(255,235 - z * 36,220 - Z * 90)
			
		Case "blacklight"
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 1, RGB(30 - 27 * Z ^ 0.03,30 - 28 * Z ^ 0.01, 255), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 20 + 900 * Z / (1 + DNA30)
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 60 * Z
			FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 30000 * Z ^ 3
			Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 2000 * (Z ^ 3) / (1 + DNA90)
			FlBumperSmallLight(nr).color = RGB(255 - 240 * (Z ^ 0.1),255 - 240 * (Z ^ 0.1),255)
			FlBumperSmallLight(nr).colorfull = RGB(255 - 200 * z,255 - 200 * Z,255)
			MaterialColor "bumpertopmat" & nr, RGB(255 - 190 * Z,235 - z * 180,220 + 35 * Z)
			
		Case "yellow"
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255, 180 + 40 * z, 48 * Z), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 17 + 200 * Z / (1 + DNA30 ^ 2)
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 40 * Z / (1 + DNA90)
			FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 2000 * (0.03 * Z + 0.97 * Z ^ 10)
			Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 1000 * (Z ^ 3) / (1 + DNA90)
			MaterialColor "bumpertopmat" & nr, RGB(255,200, 24 - 24 * z)
			
		Case "purple"
			
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(128 - 118 * Z - 32 * Sin(Z * 3.14), 32 - 26 * Z ,255), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 15 + 200 * Z / (0.5 + DNA30)
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
			FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 10000 * (0.03 * Z + 0.97 * Z ^ 3)
			Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 4000 * (Z ^ 3) / (0.5 + DNA90)
			MaterialColor "bumpertopmat" & nr, RGB(128 - 60 * Z,32,255)
			
			
	End Select
End Sub

Sub BumperTimer_Timer
	Dim nr
	For nr = 1 To 6
		If FlBumperActive(nr) = True Then
			FlBumperFadeTarget(nr) = FlBumperControlLight(nr).IntensityScale ' CUSTOM LINE (Lampz support)
			If FlBumperFadeActual(nr) < FlBumperFadeTarget(nr) And FlBumperActive(nr) Then
				FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.8
				If FlBumperFadeActual(nr) > 0.99 Then FlBumperFadeActual(nr) = 1
				FlFadeBumper nr, FlBumperFadeActual(nr)
			End If
			If FlBumperFadeActual(nr) > FlBumperFadeTarget(nr) And FlBumperActive(nr)  Then
				FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.4 / (FlBumperFadeActual(nr) + 0.1)
				If FlBumperFadeActual(nr) < 0.01 Then FlBumperFadeActual(nr) = 0
				FlFadeBumper nr, FlBumperFadeActual(nr)
			End If
		End If
	Next
End Sub


'******************************************************
'******  END FLUPPER BUMPERS
'******************************************************

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
' ZSCO: SCORBIT Interface
' To Use:
' 1) Define a timer tmrScorbit
' 2) Call DoInit at the end of PupInit or in Table Init if you are nto using pup with the appropriate parameters
'     Replace 2108 with your TableID from Scorbit 
'     Replace GRWvz-MP37P from your table on OPDB - eg: https://opdb.org/machines/2103
'		if Scorbit.DoInit(389, "PupOverlays", "1.0.0", "GRWvz-MP37P") then 
'			tmrScorbit.Interval=2000
'			tmrScorbit.UserValue = 0
'			tmrScorbit.Enabled=True 
'		End if 
' 3) Customize helper functions below for different events if you want or make your own 
' 4) Call 
'		DoInit - After Pup/Screen is setup (PuPInit)
'		StartSession - When a game starts (ResetForNewGame)
'		StopSession - When the game is over (Table1_Exit, EndOfGame)
'		SendUpdate - called when Score Changes (AddScore)
'			SendUpdate(P1Score, P2Score, P3Score, P4Score, CurrentBall, CurrentPlayer, NumberPlayers)
'			Example:  Scorbit.SendUpdate Score(0), Score(1), Score(2), Score(3), Balls, CurrentPlayer+1, PlayersPlayingGame
'		SetGameMode - When different game events happen like starting a mode, MB etc.  (ScorbitBuildGameModes helper function shows you how)
' 5) Drop the binaries sQRCode.exe and sToken.exe in your Pup Root so we can create session tokens and QRCodes.
'	- Drop QRCode Images (QRCodeS.png, QRcodeB.png) in yur pup PuPOverlays if you want to use those 
' 6) Callbacks 
'		Scorbit_Paired   	- Called when machine is successfully paired.  Hide QRCode and play a sound 
'		Scorbit_PlayerClaimed	- Called when player is claimed.  Hide QRCode, play a sound and display name 
'		ScorbitClaimQR		- Call before/after plunge (swPlungerRest_Hit, swPlungerRest_UnHit)
' 7) Other 
'		Set Pair QR Code	- During Attract
'			if (Scorbit.bNeedsPairing) then 
'				PuPlayer.LabelSet pDMDFull, "ScorbitQR_a", "PuPOverlays\\QRcode.png",1,"{'mt':2,'width':32, 'height':64,'xalign':0,'yalign':0,'ypos':5,'xpos':5}"
'				PuPlayer.LabelSet pDMDFull, "ScorbitQRIcon_a", "PuPOverlays\\QRcodeS.png",1,"{'mt':2,'width':36, 'height':85,'xalign':0,'yalign':0,'ypos':3,'xpos':3,'zback':1}"
'			End if 
'		Set Player Names 	- Wherever it makes sense but I do it here: (pPupdateScores)
'		   if ScorbitActive then 
'			if Scorbit.bSessionActive then
'				PlayerName=Scorbit.GetName(CurrentPlayer+1)
'				if PlayerName="" then PlayerName= "Player " & CurrentPlayer+1 
'			End if 
'		   End if 
'
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' TABLE CUSTOMIZATION START HERE 

Sub Scorbit_Paired()								' Scorbit callback when new machine is paired
	If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "PAIRED"
	PlaySound "scorbit_connected", 1, VOLUME_SFX
	DMD.ScorbitCheckPairing
End Sub

Sub Scorbit_PlayerClaimed(PlayerNum, PlayerName)	' Scorbit callback when QR Is Claimed
	If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "LOGGED IN"
	PlaySound "scorbit_claimed", 1, VOLUME_SFX
	ScorbitClaimQR(False)
	If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "PLAYER CLAIMED; #" & PlayerNum & " " & PlayerName

	dataPlayer(PlayerNum - 1).data.Item("name") = Left(UCase(PlayerName), 15)
	DMD.ScorbitPlayerClaimed PlayerNum, PlayerName
End Sub


Sub ScorbitClaimQR(bShow)						'  Show QRCode on first ball for users to claim this position
	If Scorbit.bSessionActive=False Then Exit Sub
	If ScorbitShowClaimQR=False Then Exit Sub
	If Scorbit.bNeedsPairing Then Exit Sub
	If bShow And CurrentPlayer("score") < 10 And Scorbit.GetName(dataGame.data.Item("playerUp"))="" Then
		If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "Player claim QR code visible"
		DMD.ScorbitClaimQR True
	Else
		If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "Player claim QR code hidden"
		DMD.ScorbitClaimQR False
	End If
End Sub

Sub ScorbitBuildGameModes()		' Custom function to build the game modes for better stats 
	dim GameModeStr
	if Scorbit.bSessionActive=False then Exit Sub 

	GameModeStr="NA:"

	If CURRENT_MODE >= 100 And dataGame.data.Item("playerUp") > 0 Then
		If dataGame.data.Item("gameDifficulty") = 0 Then
			GameModeStr="NA{green}:? HP"
		ElseIf currentPlayer("HP") > 0 Then
			GameModeStr="NA{green}:" & currentPlayer("HP") & " HP"
		Else
			GameModeStr="NA{green}:0 HP"
		End If
		GameModeStr=GameModeStr&";NA{purple}:" & currentPlayer("AC") & " AC"
		GameModeStr=GameModeStr&";NA{red}:" & currentPlayer("drainDamage") & " Drain Damage"
		GameModeStr=GameModeStr&";BX" & currentPlayer("bonusX") & ":Bonus " & currentPlayer("bonusX") & "X"
	End If

	Select Case CURRENT_MODE
		Case MODE_DEATH_SAVE, MODE_DEATH_SAVE_INTRO:
			GameModeStr=GameModeStr&";NA{red}:Death Save"
        Case MODE_MISSING_BALL:
			GameModeStr=GameModeStr&";NA{red}:Missing Ball"
        Case MODE_NORMAL:
			GameModeStr=GameModeStr&";NA{blue}:Normal Gameplay"
		Case MODE_SKILLSHOT:
			GameModeStr=GameModeStr&";NA{purple}:Skillshot"
		Case MODE_GLITCH_WIZARD_INTRO, MODE_GLITCH_WIZARD:
			GameModeStr=GameModeStr&";WM:Glitch"
		Case MODE_BLACKSMITH_WIZARD_SPARE:
			GameModeStr=GameModeStr&";WM:Blacksmith (Spare)"
		Case MODE_BLACKSMITH_WIZARD_KILL:
			GameModeStr=GameModeStr&";WM:Blacksmith (Kill)"
	End Select

	Scorbit.SetGameMode(GameModeStr)
End Sub 

Sub Scorbit_LOGUpload(state)	' Callback during the log creation process.  0=Creating Log, 1=Uploading Log, 2=Done
	Select Case state
		Case 0:
			If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "CREATING Log"
		Case 1:
			If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "Uploading Log"
		Case 2:
			If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "Log Complete"
	End Select
End Sub
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' TABLE CUSTOMIZATION END HERE - NO NEED TO EDIT BELOW THIS LINE

' WARNING for Saving wallden: Some edits were made to the class to support Wallden: puplayer.getroot -> DMD.getPUPRoot, and other changes as commented


dim Scorbit : Set Scorbit = New ScorbitIF
' Workaround - Call get a reference to Member Function
Sub tmrScorbit_Timer()								' Timer to send heartbeat 
	Scorbit.DoTimer(tmrScorbit.UserValue)
	tmrScorbit.UserValue=tmrScorbit.UserValue+1
	if tmrScorbit.UserValue>5 then tmrScorbit.UserValue=0
End Sub 
Function ScorbitIF_Callback()
	Scorbit.Callback()
End Function 
Class ScorbitIF

	Public bSessionActive
	Public bNeedsPairing
	Private bUploadLog
	Private bActive
	Private LOGFILE(10000000)
	Private LogIdx

	Private bProduction

	Private TypeLib
	Private MyMac
	Private Serial
	Private MyUUID
	Private TableVersion

	Private SessionUUID
	Private SessionSeq
	Private SessionTimeStart
	Private bRunAsynch
	Private bWaitResp
	Private GameMode
	Private GameModeOrig		' Non escaped version for log
	Private VenueMachineID
	Private CachedPlayerNames(4)
	Private SaveCurrentPlayer

	Public bEnabled
	Private sToken
	Private machineID
	Private dirQRCode
	Private opdbID
	Private wsh

	Private objXmlHttpMain
	Private objXmlHttpMainAsync
	Private fso
	Private Domain

	Public Sub Class_Initialize()
		bActive="false"
		bSessionActive=False
		bEnabled=False 
	End Sub 

	Property Let UploadLog(bValue)
		bUploadLog = bValue
	End Property

	Sub DoTimer(bInterval)	' 2 second interval
		dim holdScores(4)
		dim i
		if bInterval=0 then 
			SendHeartbeat()
		elseif bRunAsynch And bSessionActive = True then ' Game in play (Updated to resolve stutter in CoopMode)
			'Customized for Saving Wallden
			Scorbit.SendUpdate dataPlayer(0).data.Item("score"), dataPlayer(1).data.Item("score"), dataPlayer(2).data.Item("score"), dataPlayer(3).data.Item("score"), dataGame.data.Item("ball"), dataGame.data.Item("playerUp"), dataGame.data.Item("numPlayers")
		End if 
	End Sub 

	Function GetName(PlayerNum)	' Return Parsed Players name  
		if PlayerNum<1 or PlayerNum>4 then 
			GetName=""
		else 
			GetName=CachedPlayerNames(PlayerNum-1)
		End if 
	End Function 

	Function DoInit(MyMachineID, Directory_PupQRCode, Version, opdb)
		dim Nad
		Dim EndPoint
		Dim resultStr 
		Dim UUIDParts 
		Dim UUIDFile

		bProduction=1
'		bProduction=0
		SaveCurrentPlayer=0
		VenueMachineID=""
		bWaitResp=False 
		bRunAsynch=False 
		DoInit=False 
		opdbID=opdb
		dirQrCode=Directory_PupQRCode
		MachineID=MyMachineID
		TableVersion=version
		bNeedsPairing=False 
		if bProduction then 
			domain = "api.scorbit.io"
		else 
			domain = "staging.scorbit.io"
			domain = "scorbit-api-staging.herokuapp.com"
		End if 
		Set fso = CreateObject("Scripting.FileSystemObject")
		dim objLocator:Set objLocator = CreateObject("WbemScripting.SWbemLocator")
		Dim objService:Set objService = objLocator.ConnectServer(".", "root\cimv2")
		Set objXmlHttpMain = CreateObject("Msxml2.ServerXMLHTTP")
		Set objXmlHttpMainAsync = CreateObject("Microsoft.XMLHTTP")
		objXmlHttpMain.onreadystatechange = GetRef("ScorbitIF_Callback")
		Set wsh = CreateObject("WScript.Shell")

		' Get Mac for Serial Number 
		dim Nads: set Nads = objService.ExecQuery("Select * from Win32_NetworkAdapter where physicaladapter=true")
		for each Nad in Nads
			if not isnull(Nad.MACAddress) then
				if left(Nad.MACAddress, 6)<>"00090F" then ' Skip over forticlient MAC
					If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "Using MAC Addresses:" & Nad.MACAddress & " From Adapter:" & Nad.description   
					MyMac=replace(Nad.MACAddress, ":", "")
					Exit For 
				End if 
			End if 
		Next
		Serial=eval("&H" & mid(MyMac, 5))
		if Serial<0 then Serial=eval("&H" & mid(MyMac, 6))		' Mac Address Overflow Special Case 
		if MyMachineID<>2108 then 			' GOTG did it wrong but MachineID should be added to serial number also
			Serial=Serial+MyMachineID
		End if 
'		Serial=123456
		If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "Serial:" & Serial

		' Get System UUID
		set Nads = objService.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct")
		for each Nad in Nads
			If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "Using UUID:" & Nad.UUID   
			MyUUID=Nad.UUID
			Exit For 
		Next

		if MyUUID="" then 
			MsgBox "SCORBIT - Can get UUID, Disabling."
			Exit Function
		elseif MyUUID="03000200-0400-0500-0006-000700080009" or ScorbitAlternateUUID then
			If fso.FolderExists(UserDirectory) then 
				If fso.FileExists(UserDirectory & "ScorbitUUID.dat") then
					Set UUIDFile = fso.OpenTextFile(UserDirectory & "ScorbitUUID.dat",1)
					MyUUID = UUIDFile.ReadLine()
					UUIDFile.Close
					Set UUIDFile = Nothing
				Else 
					MyUUID=GUID()
					Set UUIDFile=fso.CreateTextFile(UserDirectory & "ScorbitUUID.dat",True)
					UUIDFile.WriteLine MyUUID
					UUIDFile.Close
					Set UUIDFile=Nothing
				End if
			End if 
		End if

		' Clean UUID
		UUIDParts=split(MyUUID, "-")
		MyUUID=LCASE(Hex(eval("&h" & UUIDParts(0))+MyMachineID) & UUIDParts(1) &  UUIDParts(2) &  UUIDParts(3) & UUIDParts(4))		 ' Add MachineID to UUID
		MyUUID=LPad(MyUUID, 32, "0")
'		MyUUID=Replace(MyUUID, "-",  "")
'		Debug.print "MyUUID:" & MyUUID 


' Debug
'		myUUID="adc12b19a3504453a7414e722f58737f"
'		Serial="123456778"

		' Authenticate and get our token 
		if getStoken() then 
			bEnabled=True 
'			SendHeartbeat
			DoInit=True
		End if 
	End Function 

	Sub Callback()
		Dim ResponseStr
		Dim i 
		Dim Parts
		Dim Parts2
		Dim Parts3
		if bEnabled=False then Exit Sub 

		if bWaitResp and objXmlHttpMain.readystate=4 then 
'			Debug.print "CALLBACK: " & objXmlHttpMain.Status & " " & objXmlHttpMain.readystate
			if objXmlHttpMain.Status=200 and objXmlHttpMain.readystate = 4 then 
				ResponseStr=objXmlHttpMain.responseText
				If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "CALLBACK RESPONSE: " & ResponseStr

				' Parse Name 
				If bSessionActive = True Then
					if CachedPlayerNames(SaveCurrentPlayer-1)="" then  ' Player doesnt have a name
						if instr(1, ResponseStr, "cached_display_name") <> 0 Then	' There are names in the result
							Parts=Split(ResponseStr,",{")							' split it 
							if ubound(Parts)>=SaveCurrentPlayer-1 then 				' Make sure they are enough avail
								if instr(1, Parts(SaveCurrentPlayer-1), "cached_display_name")<>0 then 	' See if mine has a name 
									CachedPlayerNames(SaveCurrentPlayer-1)=GetJSONValue(Parts(SaveCurrentPlayer-1), "cached_display_name")		' Get my name
									CachedPlayerNames(SaveCurrentPlayer-1)=Replace(CachedPlayerNames(SaveCurrentPlayer-1), """", "")
									Scorbit_PlayerClaimed SaveCurrentPlayer, CachedPlayerNames(SaveCurrentPlayer-1)
	'								Debug.print "Player Claim:" & SaveCurrentPlayer & " " & CachedPlayerNames(SaveCurrentPlayer-1)
								End if 
							End if
						End if 
					else												    ' Check for unclaim 
						if instr(1, ResponseStr, """player"":null")<>0 Then	' Someone doesnt have a name
							Parts=Split(ResponseStr,"[")						' split it 
	'Debug.print "Parts:" & Parts(1)
							Parts2=Split(Parts(1),"}")							' split it 
							for i = 0 to Ubound(Parts2)
	'Debug.print "Parts2:" & Parts2(i)
								if instr(1, Parts2(i), """player"":null")<>0 Then
									CachedPlayerNames(i)=""
								End if 
							Next 
						End if 
					End if
				End If

				'Check heartbeat
				HandleHeartbeatResp ResponseStr
			End if 
			bWaitResp=False
		End if 
	End Sub

	Public Sub StartSession()
		if bEnabled=False then Exit Sub 
		If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "Start Session" 
		CachedPlayerNames(0)=""
		CachedPlayerNames(1)=""
		CachedPlayerNames(2)=""
		CachedPlayerNames(3)=""
		bRunAsynch=True 
		bActive="true"
		bSessionActive=True
		SessionSeq=0
		SessionUUID=GUID()
		SessionTimeStart=GameTime
		LogIdx=0
		SendUpdate 0, 0, 0, 0, 1, 1, 1
	End Sub

	' Custom method to work around stuttering
	Public Sub ForceAsynch(enabled)
		if bEnabled=False then Exit Sub
		if bSessionActive=True then Exit Sub 'Sessions should always control asynch when active
		bRunAsynch=enabled
	End Sub

	Public Sub StopSession(P1Score, P2Score, P3Score, P4Score, NumberPlayers)
		StopSession2 P1Score, P2Score, P3Score, P4Score, NumberPlayers, False
	End Sub 

	Public Sub StopSession2(P1Score, P2Score, P3Score, P4Score, NumberPlayers, bCancel)
		Dim i
		dim objFile
		if bEnabled=False then Exit Sub 
		bRunAsynch=False 'Asynch might have been forced to prevent stutter
		if bSessionActive=False then Exit Sub 
		If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "Stop Session" 

		bActive="false" 
		SendUpdate P1Score, P2Score, P3Score, P4Score, -1, -1, NumberPlayers
		bSessionActive=False
'		SendHeartbeat

		if bUploadLog and LogIdx<>0 and bCancel=False then 
			If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "Creating Scorbit Log: Size " & LogIdx
			Scorbit_LOGUpload(0)
			Set objFile = fso.CreateTextFile(DMD.getPUPRoot&"\" & cGameName & "\sGameLog.csv")
			For i = 0 to LogIdx-1 
				objFile.Writeline LOGFILE(i)
			Next 
			objFile.Close
			LogIdx=0
			Scorbit_LOGUpload(1)
			pvPostFile "https://" & domain & "/api/session_log/", DMD.getPUPRoot&"\" & cGameName & "\sGameLog.csv", False
			Scorbit_LOGUpload(2)
			on error resume next
			fso.DeleteFile(DMD.getPUPRoot&"\" & cGameName & "\sGameLog.csv")
			Err.Clear
			on error goto 0
		End if 

	End Sub 

	Public Sub SetGameMode(GameModeStr)
		GameModeOrig=GameModeStr
		GameMode=GameModeStr
		GameMode=Replace(GameMode, ":", "%3a")
		GameMode=Replace(GameMode, ";", "%3b")
		GameMode=Replace(GameMode, " ", "%20")
		GameMode=Replace(GameMode, "{", "%7B")
		GameMode=Replace(GameMode, "}", "%7D")
	End sub 

	Public Sub SendUpdate(P1Score, P2Score, P3Score, P4Score, CurrentBall, CurrentPlayer, NumberPlayers)
		SendUpdateAsynch P1Score, P2Score, P3Score, P4Score, CurrentBall, CurrentPlayer, NumberPlayers, bRunAsynch
	End Sub 

	Public Sub SendUpdateAsynch(P1Score, P2Score, P3Score, P4Score, CurrentBall, CurrentPlayer, NumberPlayers, bAsynch)
		dim i
		Dim PostData
		Dim resultStr
		dim LogScores(4)

		if bUploadLog then 
			if NumberPlayers>=1 then LogScores(0)=P1Score
			if NumberPlayers>=2 then LogScores(1)=P2Score
			if NumberPlayers>=3 then LogScores(2)=P3Score
			if NumberPlayers>=4 then LogScores(3)=P4Score
			LOGFILE(LogIdx)=DateDiff("S", "1/1/1970", Now()) & "," & LogScores(0) & "," & LogScores(1) & "," & LogScores(2) & "," & LogScores(3) & ",,," &  CurrentPlayer & "," & CurrentBall & ",""" & GameModeOrig & """"
			LogIdx=LogIdx+1
		End if

		if bSessionActive=False then Exit Sub 
		if bEnabled=False then Exit Sub 
		if bWaitResp then exit sub ' Drop message until we get our next response 

'		msgbox "currentplayer: " & CurrentPlayer
		SaveCurrentPlayer=CurrentPlayer
'		PostData = "session_uuid=" & SessionUUID & "&session_time=" & DateDiff("S", "1/1/1970", Now()) & _
'					"&session_sequence=" & SessionSeq & "&active=" & bActive
		PostData = "session_uuid=" & SessionUUID & "&session_time=" & GameTime-SessionTimeStart+1 & _
					"&session_sequence=" & SessionSeq & "&active=" & bActive

		SessionSeq=SessionSeq+1
		if NumberPlayers > 0 then 
			for i = 0 to NumberPlayers-1
				PostData = PostData & "&current_p" & i+1 & "_score="
				if i <= NumberPlayers-1 then 
					if i = 0 then PostData = PostData & P1Score
					if i = 1 then PostData = PostData & P2Score
					if i = 2 then PostData = PostData & P3Score
					if i = 3 then PostData = PostData & P4Score
				else 
					PostData = PostData & "-1"
				End if 
			Next 

			PostData = PostData & "&current_ball=" & CurrentBall & "&current_player=" & CurrentPlayer
			if GameMode<>"" then PostData=PostData & "&game_modes=" & GameMode
		End if 
		resultStr = PostMsg("https://" & domain, "/api/entry/", PostData, bAsynch)
		if resultStr<>"" and DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "SendUpdate Resp: " & resultStr
	End Sub 

' PRIVATE BELOW 
	Private Function LPad(StringToPad, Length, CharacterToPad)
	  Dim x : x = 0
	  If Length > Len(StringToPad) Then x = Length - len(StringToPad)
	  LPad = String(x, CharacterToPad) & StringToPad
	End Function

	Private Function GUID()		
		Dim TypeLib
		Set TypeLib = CreateObject("Scriptlet.TypeLib")
		GUID = Mid(TypeLib.Guid, 2, 36)

'		Set wsh = CreateObject("WScript.Shell")
'		Set fso = CreateObject("Scripting.FileSystemObject")
'
'		dim rc
'		dim result
'		dim objFileToRead
'		Dim sessionID:sessionID=DMD.getPUPRoot&"\" & cGameName & "\sessionID.txt"
'
'		on error resume next
'		fso.DeleteFile(sessionID)
'		On error goto 0 
'
'		rc = wsh.Run("powershell -Command ""(New-Guid).Guid"" | out-file -encoding ascii " & sessionID, 0, True)
'		if FileExists(sessionID) and rc=0 then
'			Set objFileToRead = fso.OpenTextFile(sessionID,1)
'			result = objFileToRead.ReadLine()
'			objFileToRead.Close
'			GUID=result
'		else 
'			MsgBox "Cant Create SessionUUID through powershell. Disabling Scorbit"
'			bEnabled=False 
'		End if

	End Function

	Private Function GetJSONValue(JSONStr, key)
		dim i 
		Dim tmpStrs,tmpStrs2
		if Instr(1, JSONStr, key)<>0 then 
			tmpStrs=split(JSONStr,",")
			for i = 0 to ubound(tmpStrs)
				if instr(1, tmpStrs(i), key)<>0 then 
					tmpStrs2=split(tmpStrs(i),":")
					GetJSONValue=tmpStrs2(1)
					exit for
				End if 
			Next 
		End if 
	End Function

	Private Sub SendHeartbeat()
		Dim resultStr
		if bEnabled=False then Exit Sub 
		resultStr = GetMsgHdr("https://" & domain, "/api/heartbeat/", "Authorization", "SToken " & sToken)
		
		'Customized
		If bRunAsynch = False Then 
			If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "Heartbeat Resp: " & resultStr
			HandleHeartbeatResp ResultStr
		End If
	End Sub 

	'custom method
	Private Sub HandleHeartbeatResp(resultStr)
		dim TmpStr
		Dim Command
		Dim rc
		Dim QRFile:QRFile=DMD.getPUPRoot&"\" & cGameName & "\" & dirQrCode

		If VenueMachineID="" then
			If resultStr<>"" And Not InStr(resultStr, """machine_id"":" & machineID)=0 Then 'We Paired
				bNeedsPairing=False
				Scorbit_Paired()
			ElseIf resultStr<>"" And Not InStr(resultStr, """unpaired"":true")=0 Then 'We Did not Pair
				bNeedsPairing=True
				DMD.ScorbitCheckPairing	'Wallden Customization
			Else
				' Error (or not a heartbeat); do nothing
			End If

			TmpStr=GetJSONValue(resultStr, "venuemachine_id")
			if TmpStr<>"" then 
				VenueMachineID=TmpStr
'Debug.print "VenueMachineID=" & VenueMachineID			
				Command = """" & DMD.getPUPRoot&"\" & cGameName & "\sQRCode.exe"" " & VenueMachineID & " " & opdbID & " """ & QRFile & """"
				rc = wsh.Run(Command, 0, False)
			End if 
		End if
	End Sub

	Private Function getStoken()
		Dim result
		Dim results
'		dim wsh
		Dim tmpUUID:tmpUUID="adc12b19a3504453a7414e722f58736b"
		Dim tmpVendor:tmpVendor="vscorbitron"
		Dim tmpSerial:tmpSerial="999990104"
		Dim QRFile:QRFile=DMD.getPUPRoot&"\" & cGameName & "\" & dirQrCode
		Dim sTokenFile:sTokenFile=DMD.getPUPRoot&"\" & cGameName & "\sToken.dat"

		' Set everything up
		tmpUUID=MyUUID
		tmpVendor="vpin"
		tmpSerial=Serial
		
		on error resume next
		fso.DeleteFile(sTokenFile)
		Err.Clear
		On error goto 0 

		' get sToken and generate QRCode
'		Set wsh = CreateObject("WScript.Shell")
		Dim waitOnReturn: waitOnReturn = True
		Dim windowStyle: windowStyle = 0
		Dim Command 
		Dim rc
		Dim objFileToRead

		Command = """" & DMD.getPUPRoot&"\" & cGameName & "\sToken.exe"" " & tmpUUID & " " & tmpVendor & " " &  tmpSerial & " " & MachineID & " """ & QRFile & """ """ & sTokenFile & """ " & domain
		If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "RUNNING Command: " & Command
		rc = wsh.Run(Command, windowStyle, waitOnReturn)
		If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "Command Returned: " & rc
		if FileExists(DMD.getPUPRoot&"\" & cGameName & "\sToken.dat") and rc=0 then
			Set objFileToRead = fso.OpenTextFile(DMD.getPUPRoot&"\" & cGameName & "\sToken.dat",1)
			result = objFileToRead.ReadLine()
			objFileToRead.Close
			Set objFileToRead = Nothing
'Debug.print result

			if Instr(1, result, "Invalid timestamp")<> 0 then 
				MsgBox "Scorbit Timestamp Error: Please make sure the time on your system is exact"
				getStoken=False
			elseif Instr(1, result, ":")<>0 then 
				results=split(result, ":")
				sToken=results(1)
				sToken=mid(sToken, 3, len(sToken)-4)
				If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "Got TOKEN: " & sToken
				getStoken=True
			Else 
				If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "ERROR: " & result
				getStoken=False
			End if 
		else 
			If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "ERROR No File: " & rc
		End if 

	End Function 

	private Function FileExists(FilePath)
		If fso.FileExists(FilePath) Then
			FileExists=CBool(1)
		Else
			FileExists=CBool(0)
		End If
	End Function

	Private Function GetMsg(URLBase, endpoint)
		GetMsg = GetMsgHdr(URLBase, endpoint, "", "")
	End Function

	Private Function GetMsgHdr(URLBase, endpoint, Hdr1, Hdr1Val)
		Dim Url
		Url = URLBase + endpoint & "?session_active=" & bActive
		If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "Url:" & Url  & "  Async=" & bRunAsynch
		objXmlHttpMain.open "GET", Url, bRunAsynch
'		objXmlHttpMain.setRequestHeader "Content-Type", "text/xml"
		objXmlHttpMain.setRequestHeader "Cache-Control", "no-cache"
		if Hdr1<> "" then objXmlHttpMain.setRequestHeader Hdr1, Hdr1Val

'		on error resume next
			err.clear
			objXmlHttpMain.send ""
			if err.number=-2147012867 then 
				MsgBox "Multiplayer Server is down (" & err.number & ") " & Err.Description
				bEnabled=False
			elseif err.number <> 0 then 
				If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "Server error: (" & err.number & ") " & Err.Description
			End if 
			if bRunAsynch=False then 
				If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "Status: " & objXmlHttpMain.status
				If objXmlHttpMain.status = 200 Then
					GetMsgHdr = objXmlHttpMain.responseText
				Else 
					GetMsgHdr=""
				End if 
			Else 
				bWaitResp=True
				GetMsgHdr=""
			End if 
'		On error goto 0

	End Function

	Private Function PostMsg(URLBase, endpoint, PostData, bAsynch)
		Dim Url

		Url = URLBase + endpoint
		If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "PostMSg: " & Url & " " & PostData

		objXmlHttpMain.open "POST",Url, bAsynch
		objXmlHttpMain.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXmlHttpMain.setRequestHeader "Content-Length", Len(PostData)
		objXmlHttpMain.setRequestHeader "Cache-Control", "no-cache"
		objXmlHttpMain.setRequestHeader "Authorization", "SToken " & sToken
		if bAsynch then bWaitResp=True 

		on error resume next
			objXmlHttpMain.send PostData
			if err.number=-2147012867 then 
				MsgBox "Multiplayer Server is down (" & err.number & ") " & Err.Description
				bEnabled=False
			elseif err.number <> 0 then 
				If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "Multiplayer Server error (" & err.number & ") " & Err.Description
			End if 
			If objXmlHttpMain.status = 200 Then
				PostMsg = objXmlHttpMain.responseText
			else 
				PostMsg="ERROR: " & objXmlHttpMain.status & " >" & objXmlHttpMain.responseText & "<"
			End if
			Err.Clear
		On error goto 0
	End Function

	Private Function pvPostFile(sUrl, sFileName, bAsync)
		If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "Posting File " & sUrl & " " & sFileName & " " & bAsync & " File:" & Mid(sFileName, InStrRev(sFileName, "\") + 1)
		Dim STR_BOUNDARY:STR_BOUNDARY  = GUID()
		Dim nFile  
		Dim baBuffer()
		Dim sPostData
		Dim Response

		'--- read file
		Set nFile = fso.GetFile(sFileName)
		With nFile.OpenAsTextStream()
			sPostData = .Read(nFile.Size)
			.Close
		End With
'		fso.Open sFileName For Binary Access Read As nFile
'		If LOF(nFile) > 0 Then
'			ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
'			Get nFile, , baBuffer
'			sPostData = StrConv(baBuffer, vbUnicode)
'		End If
'		Close nFile

		'--- prepare body
		sPostData = "--" & STR_BOUNDARY & vbCrLf & _
			"Content-Disposition: form-data; name=""uuid""" & vbCrLf & vbCrLf & _
			SessionUUID & vbcrlf & _
			"--" & STR_BOUNDARY & vbCrLf & _
			"Content-Disposition: form-data; name=""log_file""; filename=""" & SessionUUID & ".csv""" & vbCrLf & _
			"Content-Type: application/octet-stream" & vbCrLf & vbCrLf & _
			sPostData & vbCrLf & _
			"--" & STR_BOUNDARY & "--"

'Debug.print "POSTDATA: " & sPostData & vbcrlf

		'--- post
		With objXmlHttpMain
			.Open "POST", sUrl, bAsync
			.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & STR_BOUNDARY
			.SetRequestHeader "Authorization", "SToken " & sToken
			.Send sPostData ' pvToByteArray(sPostData)
			If Not bAsync Then
				Response= .ResponseText
				pvPostFile = Response
				If DEBUG_LOG_SCORBIT = True Then WriteToLog "Scorbit", "Upload Response: " & Response
			End If
		End With

	End Function

	Private Function pvToByteArray(sText)
		pvToByteArray = StrConv(sText, 128)		' vbFromUnicode
	End Function

End Class 
'  END SCORBIT
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

'*****************************************
'   WSEG: SEGMENT DISPLAYS
'*****************************************

' Digit displays; determines which lights to illuminate for each number. segDigit is an Array, key being the number (10 is a dash), value being an Array of each segment light and whether it should be on or off for that number.
Dim segDigit(11)
segDigit(0) = Array(1, 1, 1, 1, 1, 1, 0, 0, 0, 0)
segDigit(1) = Array(0, 0, 0, 0, 0, 0, 1, 1, 0, 0)
segDigit(2) = Array(1, 0, 1, 1, 0, 1, 0, 0, 1, 1)
segDigit(3) = Array(1, 1, 1, 0, 0, 1, 0, 0, 1, 1)
segDigit(4) = Array(1, 1, 0, 0, 1, 0, 0, 0, 1, 1)
segDigit(5) = Array(0, 1, 1, 0, 1, 1, 0, 0, 1, 1)
segDigit(6) = Array(0, 1, 1, 1, 1, 1, 0, 0, 1, 1)
segDigit(7) = Array(1, 1, 0, 0, 0, 1, 0, 0, 0, 0)
segDigit(8) = Array(1, 1, 1, 1, 1, 1, 0, 0, 1, 1)
segDigit(9) = Array(1, 1, 1, 0, 1, 1, 0, 0, 1, 1)
segDigit(10) = Array(0, 0, 0, 0, 0, 0, 0, 0, 1, 1) 'dash

Sub updateSegments() ' Update the segment displays
	If CURRENT_MODE < 100 Then Exit Sub ' Sequencers etc control the segments when the game is idle, so don't touch the segments
	
	Dim segment
	Dim playerUp

	If CURRENT_MODE = MODE_DEATH_SAVE Or CURRENT_MODE = MODE_DEATH_SAVE_INTRO Then
		'Segments should display death save progress
		prepareSegments Array(YellowSegmentA, YellowSegmentB, YellowSegmentC), MODE_COUNTERS_B(0), False, Null
		lampC.LightOff ModeTimeL
		lampC.LightOff MonsterHPL

		prepareSegments Array(RedSegmentA, RedSegmentB, RedSegmentC), MODE_COUNTERS_B(1), False, Null
		lampC.LightOff ShieldTimeL
		lampC.LightOff PlayerHPL
	ElseIf CURRENT_MODE < 1000 Then
		'Turn off all the segments for all (other) idle modes
		For Each segment In SegmentLights
			segment.state = 0
		Next
	Else
		' Bonus X
		For segment = 10 To 13 ' X
			BonusXSegments(segment).state = 1
		Next
		prepareSegments Array(BonusXSegments), currentPlayer("bonusX"), False, Null
		
		' Shield Time / Player HP
		If Clocks.data.Item("shield").timeLeft > 2000 Then	'2000 because we don't want to count the last 2 seconds as it's a grace period
			prepareSegments Array(RedSegmentA, RedSegmentB, RedSegmentC), ((Clocks.data.Item("shield").timeLeft - 2000) / 1000), True, SegmentRedDecimal
			lampC.LightOn ShieldTimeL
			lampC.LightOff PlayerHPL
		Else
			If Not dataGame.data.Item("gameDifficulty") = 0 Then
				prepareSegments Array(RedSegmentA, RedSegmentB, RedSegmentC), currentPlayer("HP"), False, SegmentRedDecimal
			Else
				prepareSegments Array(RedSegmentA, RedSegmentB, RedSegmentC), -1, False, SegmentRedDecimal
			End If
			lampC.LightOff ShieldTimeL
			lampC.LightOn PlayerHPL
		End If

		'Mode / Monster HP
		If Clocks.data.Item("mode").timeLeft > 0 Then
			prepareSegments Array(YellowSegmentA, YellowSegmentB, YellowSegmentC), (Clocks.data.Item("mode").timeLeft / 1000), True, SegmentYellowDecimal
			lampC.LightOn ModeTimeL
			lampC.LightOff MonsterHPL
		Else
			prepareSegments Array(YellowSegmentA, YellowSegmentB, YellowSegmentC), Null, True, SegmentYellowDecimal
			lampC.LightOff ModeTimeL
			lampC.LightOff MonsterHPL
		End If
	End If
	
End Sub

'---------------------------------------------------------------
' prepareSegments
' Displays a number on a segment display.
'
' PARAMETERS:
' collections (Array) - Array of VPX collections. Each
' collection is a digit of the segment display which contains
' each VPX light for that digit in order according to segDigit
' order. The collections array itself has collections in order
' from least significant digit to greatest.
' value (Number) - The value to display (null: turn off)
' showTenths (Boolean) - Whether to show the tenths place when
' value is a single or double and is below 100 but above 0.
' tenthsLight (Light) - The VPX light that shows the decimal.
' Pass null for segment displays without a decimal light.
'---------------------------------------------------------------
Sub prepareSegments(collections, byval value, showTenths, tenthsLight)
	Dim i
	Dim i2
	Dim i3
	Dim digit
	
	Dim lowerValue
	Dim upperValue
	Dim firstDigit
	firstDigit = False
	
	' Determine how many digits we have
	Dim digits
	digits = UBound(collections) + 1
	
	' Null value = turn off all the lights (unlike 0, which explicitly means show a 0)
	If IsNull(value) Then
		If Not IsNull(tenthsLight) Then tenthsLight.state = 0
		For Each i In collections
			For Each i2 In i
				i2.state = 0
			Next
		Next
		Exit Sub
	End If
	
	' Determine what to do with our decimal light and whether or not we want to show the tenths place
	If Not IsNull(tenthsLight) Then
		If value < 0.1 Or value >= (10 ^ (digits-1)) Or showTenths = False Then
			tenthsLight.state = 0
		Else
			tenthsLight.state = 1
			value = value * 10 ' Treat value as 10X what it actually is to make processing the numbers easier
		End If
	End If
	
	' We must round value downward to an integer to prevent weird issues with tens+ digits decreasing too soon
	value = Int(value)
	
	' If value is too big to display given the number of digits we have, or the value is negative, then show dashes across all digits
	If value >= (10 ^ digits) Or value < 0 Then
		If Not IsNull(tenthsLight) Then tenthsLight.state = 0
		For Each i In collections
			For i2 = 0 To 9
				i(i2).state = segDigit(10)(i2)
			Next
		Next
		Exit Sub
	End If
	
	' Process our numbers in reverse from greatest digit to least
	For i = (digits - 1) To 0 Step -1
		' Determine the lower and upper bound for this digit
		lowerValue = 10 ^ i
		If lowerValue = 1 Then lowerValue = 0 ' 10 ^ 0 = 1, but we want the lowerValue to be 0 in this case
		upperValue = 10 ^ (i + 1)
		
		' This digit should display a number if the value is within its range, any of the more significant digits are displaying a number, or we are showing the tenths place and this is the second digit (so we get a 0.X display) unless the value is 0
		If value >= lowerValue Or firstDigit = True Or (showTenths = True And i = 1 And value > 0) Then
			firstDigit = True
			
			' Determine the value this digit should show
			If lowerValue = 0 Then
				digit = (value Mod 10) ' Mod by 10 just in case; we don't want an overflow
			Else
				digit = Int(value / lowerValue)
			End If
			
			' Illuminate the segments appropriately to make out the number
			For i2 = 0 To 9
				collections(i)(i2).state = segDigit(digit)(i2)
			Next
			
			' Decrease the value so we can process the next digit down
			If lowerValue = 0 Then
				value = 0
			Else
				value = value - (digit * lowerValue)
			End If
		Else ' When a digit should not display a number, turn all the lights off
			For Each i2 In collections(i)
				i2.state = 0
			Next
		End If
	Next
End Sub

Sub segmentTimer_Timer()
	updateSegments()
End Sub

'*****************************************
'   END SEGMENT DISPLAYS
'*****************************************

'*****************************************
'   WHLP: HELPER FUNCTIONS
'*****************************************

Const ScoreLength = 13 'Maximum length of formatted score allowed in characters; anything higher is truncated. This includes commas inserted every third place, but does not include negative symbol.

'=====================================================
' FormatScore
' Formats a numerical score into a comma-delimited
' string. Accounts for overloading. In this case,
' scores > 9,999,999,999 are reduced to K, and scores
' above 999,999,999,999 are reduced to M (millions).
'=====================================================
Function FormatScore(ScoreVal)
	Dim ScoreBuf
	Dim Index
	Dim AddChar
	
	AddChar = ""
	
	ScoreBuf = CStr(FormatNumber(Abs(ScoreVal), 0, - 1, 0, - 1))
	If Len(ScoreBuf) > ScoreLength Then ' Too big? Truncate to thousands
		ScoreBuf = Left(ScoreBuf, Len(ScoreBuf) - 4)
		AddChar = "K"
		If Len(ScoreBuf) > ScoreLength Then ' I don't think anyone would ever score in the trillions. But just in case, truncate further to millions if so.
			ScoreBuf = Left(ScoreBuf, Len(ScoreBuf) - 4)
			AddChar = "M"
		End If
		' Realistically, no one would ever score above 999 trillion. That's ungodly insane especially since the game ends at some point anyway. So we won't add any further truncation.
	End If
	ScoreBuf = ScoreBuf & AddChar
	If ScoreVal < 0 Then ScoreBuf = "-" & ScoreBuf 'Add negative sign if this was a negative number
	FormatScore = ScoreBuf
End Function

'=====================================================
' WRandomize
' Initializes the tournament-friendly WRnd random
' number generator by generating seeds in the dataSeeds
' book keeper for each player. The seeds generated
' will be based on GLOBAL_SEED if not null, otherwise
' a rnd call.
'=====================================================
Sub WRandomize()
	' If we want a random seed, generate it, otherwise ensure the seed is a positive float
	If IsNull(GLOBAL_SEED) Then
		dataGame.data.item("seed") = CDbl(Rnd)
	ElseIf Abs(GLOBAL_SEED) >= 1 Then
		dataGame.data.item("seed") = CDbl(1 / Abs(GLOBAL_SEED))
	Else
		dataGame.data.item("seed") = CDbl(Abs(GLOBAL_SEED))
	End If

	If DEBUG_LOG_RANDOM_NUMBERS = True Then WriteToLog "Random numbers", "Game seed set to " & dataGame.data.item("seed")
	
	' Generate initial tournament-friendly seeds from our game seed; the seeds should be the same for every player
	Dim i
	Dim k
	For i = 0 To 3
		Rnd(-1)
		Randomize(dataGame.data.item("seed"))
		For Each k In dataSeeds(i).data.Keys
			dataSeeds(i).data.Item(k) = CDbl(Rnd)
		Next
	Next

	Randomize
End Sub

'=====================================================
' WRnd
' Generate a random number from the
' tournament-friendly dataSeeds seed of the current
' player. Returns a float between 0 and 1.
'
' PARAMETER
' key (String) - The dataSeeds seed to use
' increment (Boolean) - Whether to increment the seed
' number after use
'=====================================================
Function wrnd(key, increment)
	' Grab the current player
	Dim player
	player = dataGame.data.Item("playerUp") - 1

	' Coop mode: seeds for every player are in the player 0 slot
	If dataGame.data.Item("gameMode") = 1 Then player = 0
	
	' Grab our seed
	Dim seed
	seed = dataSeeds(player).data.Item(Key)
	
	' Generate our random number
	Rnd(-1)
	Randomize(seed)
	Dim rndNumber
	rndNumber = CDbl(Rnd)
	Randomize
	
	' Go to the next seed sequence if applicable. We "salt" the new seed using dataGame.data.item("seed") (GLOBAL_SEED / seed of the game) so that the random numbers are a little more random.
	If increment = True Then
		Dim initialSeed
		initialSeed = rndNumber - dataGame.data.item("seed")
		If initialSeed < 0 Then initialSeed = 1 + initialSeed ' + because initialSeed is negative
		dataSeeds(player).data.Item(Key) = initialSeed
	End If
	
	' Return our random number
	wrnd = rndNumber
	If DEBUG_LOG_RANDOM_NUMBERS = True Then WriteToLog "Random numbers", "Random number for " & key & " seed " & seed & " is " & wrnd
End Function

'=====================================================
' WShuffleArray
' Randomizes the items in an array using the
' provided seed.
'
' PARAMETERS
' arr (Array) - The array to shuffle
' seed (number) - The tournament-friendly seed number
' to use for randomizing the array.
'=====================================================
Function WShuffleArray(arr, seed)
    If Not IsArray(arr) Then
        WShuffleArray = Null ' Return Null if the input is not an array
        Exit Function
    End If

	'Initialize the random number generator with the given seed
	Rnd(-1)
	Randomize(seed)

	'Randomize the array
    Dim i, j, temp
    For i = UBound(arr) To LBound(arr) Step -1
        j = Int(rnd * (i + 1))
        temp = arr(i)
        arr(i) = arr(j)
        arr(j) = temp
    Next

	'Reset the random number generator
	Randomize

    WShuffleArray = arr
End Function

Function currentPlayer(Key) ' Retrieve a dataPlayer value for the player currently playing (use currentPlayerSet to set a value)
	Dim curPlayer
	curPlayer = dataGame.data.Item("playerUp") - 1

	'Coop mode; anything not pertaining to health or name is shared across all players and stored in the 0 slot
	If dataGame.data.Item("gameMode") = 1 And Not Key = "HP" And Not Key = "dead" And Not Key = "deathSaves" And Not Key = "name" Then curPlayer = 0
	
	' Bail on invalid player
	If curPlayer < 0 Or curPlayer > 3 Then
		currentPlayer = Null
		Exit Function
	End If
	
	currentPlayer = dataPlayer(curPlayer).data.Item(key)
End Function

Sub currentPlayerSet(key, value) ' Set a value for the player currently playing
	Dim curPlayer
	curPlayer = dataGame.data.Item("playerUp") - 1

	'Coop mode; anything not pertaining to health or name is shared across all players and stored in the 0 slot
	If dataGame.data.Item("gameMode") = 1 And Not key = "HP" And Not key = "dead" And Not Key = "deathSaves" And Not Key = "name" Then curPlayer = 0
	
	' Bail on invalid player
	If curPlayer < 0 Or curPlayer > 3 Then Exit Sub
	
	dataPlayer(curPlayer).data.Item(key) = value
End Sub

Function currentPlayerBonus(bonusType, multiplied) 'Return the value of the current player's end of game bonus, optionally with the multiplier applied. BonusType is the bonus to return (all: all of them)
	Dim i
	Dim totalBonus
	totalBonus = 0

	' CASTLE Bonus
	If bonusType = "all" Or bonusType = "CASTLE" Then
		Dim castleBonus
		castleBonus = 0
		i = currentPlayer("obj_CASTLE")
		Do While i > 0
			If i > 12 Then
				castleBonus = castleBonus + Scoring.bonus("CASTLE_phase4")
			ElseIf i = 12 Then
				castleBonus = castleBonus + Scoring.bonus("CASTLE_boss3")
			ElseIf i > 8 Then
				castleBonus = castleBonus + Scoring.bonus("CASTLE_phase3")
			ElseIf i = 8 Then
				castleBonus = castleBonus + Scoring.bonus("CASTLE_boss2")
			ElseIf i > 4 Then
				castleBonus = castleBonus + Scoring.bonus("CASTLE_phase2")
			ElseIf i = 4 Then
				castleBonus = castleBonus + Scoring.bonus("CASTLE_boss1")
			Else
				castleBonus = castleBonus + Scoring.bonus("CASTLE_phase1")
			End If
			
			i = i - 1
		Loop
		totalBonus = totalBonus + castleBonus
	End If
	
	'CHASER Bonus
	If bonusType = "all" Or bonusType = "CHASER" Then totalBonus = totalBonus + (currentPlayer("obj_CHASER") * Scoring.bonus("CHASER"))
	
	'DRAGON bonus
	If bonusType = "all" Or bonusType = "DRAGON" Then totalBonus = totalBonus + (currentPlayer("obj_DRAGON") * Scoring.bonus("DRAGON"))
	
	'VIKING Bonus
	If bonusType = "all" Or bonusType = "VIKING" Then totalBonus = totalBonus + (currentPlayer("obj_VIKING") * Scoring.bonus("VIKING"))
	
	'ESCAPE Bonus
	If bonusType = "all" Or bonusType = "ESCAPE" Then totalBonus = totalBonus + (currentPlayer("obj_ESCAPE") * Scoring.bonus("ESCAPE"))
	
	'SNIPER Bonus
	If bonusType = "all" Or bonusType = "SNIPER" Then totalBonus = totalBonus + (currentPlayer("obj_SNIPER") * Scoring.bonus("SNIPER"))
	
	'BLACKSMITH Bonus
	If bonusType = "all" Or bonusType = "BLACKSMITH" Then
		Dim blacksmithBonus
		blacksmithBonus = currentPlayer("AC") * Scoring.bonus("BLACKSMITH_AC")
		blacksmithBonus = blacksmithBonus + (currentPlayer("obj_BLACKSMITH") * Scoring.bonus("BLACKSMITH_bonus"))
		totalBonus = totalBonus + blacksmithBonus
	End If
	
	'FINAL BOSS Bonus
	If bonusType = "all" Or bonusType = "FINAL BOSS" Then
		Dim finalBossBonus
		finalBossBonus = 0
		If currentPlayer("obj_FINAL_BOSS") > 0 Then finalBossBonus = Scoring.bonus("FINAL_BOSS")
		totalBonus = totalBonus + finalBossBonus
	End If
	
	'? FINAL JUDGMENT Bonus
	If bonusType = "all" Or bonusType = "FINAL JUDGMENT" Then
		Dim finalJudgmentBonus
		finalJudgmentBonus = 0
		If currentPlayer("obj_FINAL_BOSS") > 1 Then finalJudgmentBonus = Scoring.bonus("FINAL_JUDGMENT")
		totalBonus = totalBonus + finalJudgmentBonus
	End If
	
	'Bonus X
	If multiplied = True Then totalBonus = totalBonus * currentPlayer("bonusX")

	currentPlayerBonus = totalBonus
End Function

Function currentModeHasInfiniteBallShield 'Whether the current mode has ball shield always on / untimed until the mode ends
	'We do not consider blacksmith kill as infinite because, although balls come back, there is a shield time, and after the shield time, drained balls steal AC.
	Select Case CURRENT_MODE
		Case MODE_GLITCH_WIZARD
			currentModeHasInfiniteBallShield = True

		Case Else
			currentModeHasInfiniteBallShield = False
	End Select
End Function

'*****************************************
'   WSOL: SOLENOIDS
'*****************************************
Const SOLENOID_SCOOP = 1
Const SOLENOID_VUK = 2
Const SOLENOID_HP_BANK = 3
Const SOLENOID_LEFT_SLINGSHOT = 4
Const SOLENOID_RIGHT_SLINGSHOT = 5
Const SOLENOID_LEFT_FLIPPER = 6

Sub SolScoop(Enabled)
	DMD.triggerSolenoid SOLENOID_SCOOP, enabled
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "Scoop " & enabled
	If Enabled Then
		Scoop.kickZ (9 - (18 * Rnd)), (10 + (Rnd * 3)), 20, 0
		If BallInScoop = True Then
			SoundSaucerKick 1, Scoop
		Else
			SoundSaucerKick 0, Scoop
		End If
		doSwitch SWITCH_SCOOP, False
		BallInScoop = False
		updateClocks
	End If
End Sub

Sub SolVUpKicker(Enabled)
	DMD.triggerSolenoid SOLENOID_VUK, enabled
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "VUK " & enabled
	If Enabled Then
		VerticalUpKicker.kickZ 0, (20 + (Rnd * 10)), 90, 85 ' TODO: Find a more realistic way to do this
		If BallInKicker = True Then
			SoundSaucerKick 1, VerticalUpKicker
		Else
			SoundSaucerKick 0, VerticalUpKicker
		End If
		BallInKicker = False
		updateClocks
	End If
End Sub

Sub SolHPBank(enabled)
	DMD.triggerSolenoid SOLENOID_HP_BANK, enabled
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "+HP Drop Target Bank " & enabled
	If enabled Then
		DTRaise SWITCH_DROPS_PLUS
		DTRaise SWITCH_DROPS_H
		DTRaise SWITCH_DROPS_P
	End If
End Sub

Sub SolLSling(enabled)
	DMD.triggerSolenoid SOLENOID_LEFT_SLINGSHOT, enabled
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "Left Slingshot " & enabled
	If enabled Then
		LSling.Visible = 0
		LSling1.Visible = 1
		sling2.rotx = 20
		LStep = 0
		RandomSoundSlingshotLeft Sling2
	ElseIf LSling1.Visible = True Then
		LeftSlingShot.TimerEnabled = 1
	End If
End Sub

Sub SolRSling(enabled)
	DMD.triggerSolenoid SOLENOID_RIGHT_SLINGSHOT, enabled
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "Right Slingshot " & enabled
	If enabled Then
		RSling.Visible = 0
		RSling1.Visible = 1
		sling1.rotx = 20
		RStep = 0
		RandomSoundSlingshotRight Sling1
	ElseIf RSling1.Visible = True Then
		RightSlingShot.TimerEnabled = 1
	End If
End Sub

Sub SolLFlipper(Enabled)
	DMD.triggerSolenoid SOLENOID_LEFT_FLIPPER, enabled
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "Left Flipper " & enabled
	If Enabled Then
		FlipperActivate leftflipper, 1
		LF.Fire  'leftflipper.rotatetoend
		
		If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
			RandomSoundReflipUpLeft LeftFlipper
		Else
			SoundFlipperUpAttackLeft LeftFlipper
			RandomSoundFlipperUpLeft LeftFlipper
		End If
	Else
		FlipperDeactivate leftflipper, 0
		LeftFlipper.RotateToStart
		If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
			RandomSoundFlipperDownLeft LeftFlipper
		End If
		FlipperLeftHitParm = FlipperUpSoundLevel
	End If
End Sub

Sub SolRFlipper(Enabled)
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "Right Flipper " & enabled
	If Enabled Then
		If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
			RandomSoundReflipUpRight RightFlipper
		Else
			SoundFlipperUpAttackRight RightFlipper
			RandomSoundFlipperUpRight RightFlipper
		End If
		FlipperActivate rightflipper, 1
		RF.Fire 'rightflipper.rotatetoend
	Else
		If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
			RandomSoundFlipperDownRight RightFlipper
		End If
		FlipperDeactivate rightflipper, 0
		RightFlipper.RotateToStart
		FlipperRightHitParm = FlipperUpSoundLevel
	End If
End Sub

Sub SolUFlipper(Enabled)
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "Upper Flipper " & enabled
	If Enabled Then
		If upperflipper.currentangle > upperflipper.endangle - ReflipAngle Then
			RandomSoundReflipUpRight UpperFlipper
		Else
			SoundFlipperUpAttackRight UpperFlipper
			RandomSoundFlipperUpRight UpperFlipper
		End If
		upperflipper.RotateToEnd
	Else
		If UpperFlipper.currentangle > UpperFlipper.startAngle + 5 Then
			RandomSoundFlipperDownRight UpperFlipper
		End If
		upperflipper.RotateToStart
		FlipperRightHitParm = FlipperUpSoundLevel
	End If
End Sub

Sub SolRelease(enabled)
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "Trough Release " & enabled
	If enabled Then
		If swTrough1.BallCntOver > 0 Then 
			ballsInTrough = ballsInTrough - 1
			If DEBUG_LOG_TROUGH = True Then WriteToLog "Trough", "Trough release activated. Balls in trough (presumed): " & ballsInTrough
		End If
		swTrough1.kick 90, (9 + (Rnd * 2))
		RandomSoundBallRelease swTrough1
	End If
End Sub

Sub SolVUKDiverter(Enabled)
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "VUK Diverter " & enabled
	If Enabled Then
		VUKDiverter.RotateToEnd
	Else
		VUKDiverter.RotateToStart
	End If
End Sub

Sub SolPegDrop(Index, Enabled)
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "Peg Drop " & Index & " " & enabled
	If Enabled Then
		If pegDrops(Index).Z < 0 Then
			pegDrops(Index).UserValue = "raise"
			RandomSoundDropTargetReset pegDrops(Index)
		End If
	Else
		If pegDrops(Index).Z >  - 53 Then
			pegDrops(Index).UserValue = "lower"
			SoundDropTargetDrop pegDrops(Index)
		End If
	End If
End Sub

Sub Flash1(Enabled)
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "Flasher 1 " & enabled
	If Enabled Then
		ObjTargetLevel(1) = 1
	Else
		ObjTargetLevel(1) = 0
	End If
	FlasherFlash1_Timer
	Sound_Flash_Relay enabled, Flasherbase1
End Sub

Sub Flash2(Enabled)
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "Flasher 2 " & enabled
	If Enabled Then
		ObjTargetLevel(2) = 1
	Else
		ObjTargetLevel(2) = 0
	End If
	FlasherFlash2_Timer
	Sound_Flash_Relay enabled, Flasherbase2
End Sub

Sub Flash3(Enabled)
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "Flasher 3 " & enabled
	If Enabled Then
		ObjTargetLevel(3) = 1
	Else
		ObjTargetLevel(3) = 0
	End If
	FlasherFlash3_Timer
	Sound_Flash_Relay enabled, Flasherbase3
End Sub

Sub Flash4(Enabled)
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "Flasher 4 " & enabled
	If Enabled Then
		ObjTargetLevel(4) = 1
	Else
		ObjTargetLevel(4) = 0
	End If
	FlasherFlash4_Timer
	Sound_Flash_Relay enabled, Flasherbase4
End Sub

Sub Flash5(Enabled)
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "Flasher 5 " & enabled
	If Enabled Then
		ObjTargetLevel(5) = 1
	Else
		ObjTargetLevel(5) = 0
	End If
	FlasherFlash5_Timer
	Sound_Flash_Relay enabled, Flasherbase5
End Sub

Sub Flash6(Enabled)
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "Flasher 6 " & enabled
	If Enabled Then
		ObjTargetLevel(6) = 1
	Else
		ObjTargetLevel(6) = 0
	End If
	FlasherFlash6_Timer
	Sound_Flash_Relay enabled, Flasherbase6
End Sub

Sub Flash7(Enabled)
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "Flasher 7 " & enabled
	If Enabled Then
		ObjTargetLevel(7) = 1
	Else
		ObjTargetLevel(7) = 0
	End If
	FlasherFlash7_Timer
	Sound_Flash_Relay enabled, Flasherbase7
End Sub

Sub Flash8(Enabled)
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "Flasher 8 " & enabled
	If Enabled Then
		ObjTargetLevel(8) = 1
	Else
		ObjTargetLevel(8) = 0
	End If
	FlasherFlash8_Timer
	Sound_Flash_Relay enabled, Flasherbase8
End Sub

Sub SolBumper1(Enabled)
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "Bumper 1 " & enabled
	If Enabled Then
		Bumper1.PlayHit()
		RandomSoundBumperTop Bumper1
	End If
End Sub

Sub SolBumper2(Enabled)
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "Bumper 2 " & enabled
	If Enabled Then
		Bumper2.PlayHit()
		RandomSoundBumperMiddle Bumper2
	End If
End Sub

Sub SolBumper3(Enabled)
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "Bumper 3 " & enabled
	If Enabled Then
		Bumper3.PlayHit()
		RandomSoundBumperBottom Bumper3
	End If
End Sub

Sub SolLBumpDiverter(Enabled)
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "Left Bumper Diverter " & enabled
	If Enabled Then
		LeftOrbitDiverter.Open = False
		PlaySoundAtLevelStatic ("Relay_On"), VolumeDial * 0.05, LeftOrbitDiverter
	Else
		LeftOrbitDiverter.Open = True
		PlaySoundAtLevelStatic ("Relay_Off"), VolumeDial * 0.05, LeftOrbitDiverter
	End If
End Sub

Sub SolRBumpDiverter(Enabled)
	If DEBUG_LOG_SOLENOIDS = True Then WriteToLog "Solenoid", "Right Bumper Diverter " & enabled
	If Enabled Then
		RightOrbitDiverter.Open = False
		PlaySoundAtLevelStatic ("Relay_On"), VolumeDial * 0.05, RightOrbitDiverter
	Else
		RightOrbitDiverter.Open = True
		PlaySoundAtLevelStatic ("Relay_Off"), VolumeDial * 0.05, RightOrbitDiverter
	End If
End Sub

'*****************************************
' WSWI:	SWITCH HANDLER
'*****************************************

Sub doSwitch(switchNum, enabled) 'Main switch controller and logic; every switch should always call this when activated and, if it supports deactivation, when deactivated
	Dim i

	If DEBUG_LOG_SWITCHES = True Then WriteToLog "Switch", "" & switchNum & " " & enabled

	'Reset ball search timer and allow mode timer to expire if ran out; no known missing balls
	If enabled = True Then 
		BALL_SEARCH_TIME = GameTime
		If Clocks.data.Item("mode").timeLeft = 1 Then Clocks.data.Item("mode").canExpire = True
	End If

	'Cases that should always run no matter what
	If Enabled = True Then
		Select Case switchNum
			Case SWITCH_DRAIN
				If BIPWhenDrained = 0 Then BIPWhenDrained = (BIP + BIS) 'We should not count balls in the subway as drained even though they do not count in BIP
				updateClocks
				troughQueue.Add "drainKick", "Drain.kick 57, 20 : doSwitch SWITCH_DRAIN, False", 10, 0, 5, 500, 0, False

			Case SWITCH_HOLE:
				BIS = BIS + 1 'Mark that the ball is now in the subway
				If DEBUG_LOG_TROUGH = True Then WriteToLog "Trough", "Ball entered the subway. " & BIS & " balls in subway; " & BIP & " balls in play."

			Case SWITCH_BALL_LAUNCH:
				impulseP.AddBall ActiveBall
				BIPL = BIPL + 1
				swTrough1.TimerEnabled = False
				updateClocks
				
				' Auto-Fire
				If BAutoPlunge = True Then
					troughQueue.Add "autoFire", "impulseP.AutoFire", 100, 0, 500, 100, 0, False
					If BQueue < 1 Then BAutoPlunge = False
				End If

			Case SWITCH_SUBWAY_GATE:
				ballCameFromSubway = True
		End Select
	ElseIf Enabled = False Then
		Select Case switchNum
			Case SWITCH_BALL_LAUNCH
				impulseP.RemoveBall ActiveBall
				BIPL = BIPL - 1
				updateClocks
				If swTrough1.TimerEnabled = False And BQueue > 0 Then swTrough1.TimerEnabled = True
		End Select
	End If

	' Key press switches
	If switchNum < 100 Then
		Select Case switchNum
			Case SWITCH_LEFT_FLIPPER_BUTTON
				If flippersReversed = True Then
					triggerRightFlipper enabled
				Else
					triggerLeftFlipper enabled
				End If

				Select Case CURRENT_MODE
					Case MODE_BLACKSMITH_WIZARD_SELECTION
						PlaySound "ui_action_button", 1, VOLUME_SFX
						startMode MODE_BLACKSMITH_WIZARD_INTRO_KILL
					Case MODE_HIGH_SCORE_ENTRY
						If enabled = True Then DMD.highScoreSequence_toggleChar -1
				End Select

			Case SWITCH_RIGHT_FLIPPER_BUTTON
				If flippersReversed = True Then
					triggerLeftFlipper enabled
				Else
					triggerRightFlipper enabled
				End If

				Select Case CURRENT_MODE
					Case MODE_BLACKSMITH_WIZARD_SELECTION
						PlaySound "ui_action_button", 1, VOLUME_SFX
						startMode MODE_BLACKSMITH_WIZARD_INTRO_SPARE
					Case MODE_HIGH_SCORE_ENTRY
						If enabled = True Then DMD.highScoreSequence_toggleChar 1
				End Select

			Case SWITCH_COIN_1, SWITCH_COIN_2
				'This game is forcefully free to play; show the message if a coin is inserted
				If Enabled = True And Not CURRENT_MODE = MODE_INITIALIZE And Not CURRENT_MODE = MODE_FAULT Then
					Select Case Int(Rnd * 3)
						Case 0
							PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
						Case 1
							PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
						Case 2
							PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
					End Select
					queue.Add "inserting_coin_2","DMD.sAttractCredits",1000,0,0,3000,0,True
				End If

			Case SWITCH_START_BUTTON
				If Enabled = True Then
					Select Case CURRENT_MODE
						Case MODE_ATTRACT, MODE_RESUME_GAME_PROMPT
							'Clear any resume game stuff
							'DMD.triggerVideo VIDEO_CLEAR
							StopSound "callout_nar_resume_game_prompt"
							queue.RemoveAll True

							startMode MODE_CHOOSE_DIFFICULTY
						Case MODE_CHOOSE_DIFFICULTY
							startMode MODE_START_GAME
						Case MODE_SKILLSHOT
							If GAME_OFFICIALLY_STARTED = False Then addPlayer
						Case MODE_HIGH_SCORE_ENTRY
							startMode MODE_HIGH_SCORE_FINAL
					End Select
				End If

			Case SWITCH_ACTION_BUTTON
				If Enabled = True Then
					Select Case CURRENT_MODE
						Case MODE_RESUME_GAME_PROMPT
							PlaySound "ui_action_button", 1, VOLUME_SFX

							'Reset queue / game
							LightsOff actionLights
							queue.RemoveAll True
							GAME_OFFICIALLY_STARTED = True
							Music.StopTrack 1000
							StopSound "callout_nar_resume_game_prompt"
							CURRENT_MODE = 0 'Prevent double-resuming

							'Indicate Scorbit is not supported in resumed games, and force Async in Scorbit to prevent stutter
							Scorbit.ForceAsynch True
							PlayCallout "callout_nar_scorbit_not_supported", 1, VOLUME_CALLOUTS
							'DMD.triggerVideo VIDEO_CLEAR
							queue.Add "resume_game_scorbit", "DMD.triggerVideo VIDEO_RESUME_GAME_SCORBIT", 2, 0, 0, 3500, 0, False
							'queue.Add "resume_game_scorbit_2", "DMD.triggerVideo VIDEO_CLEAR", 2, 0, 0, 100, 0, False

							'Perform operation depending on resume state
							Select Case dataGame.data.Item("resumeStatus")
								Case 0: 'Game stopped in the middle of play; current player takes drain penalty and go to next player
									queue.Add "resume_game_0_1", "DMD.triggerVideo VIDEO_RESUME_GAME_DRAIN_PENALTY : PlayCallout ""callout_nar_drain_penalty"", 1, VOLUME_CALLOUTS", 2, 0, 0, 7500, 0, False
									queue.Add "resume_game_0_2", "CURRENT_MODE = MODE_NORMAL : checkDrainedBalls True", 2, 0, 0, 100, 0, False

								Case 1: 'Ball was drained but the next turn did not start yet; go to next player without drain penalty
									queue.Add "resume_game_1_1", "DMD.triggerVideo VIDEO_RESUME_GAME_NEXT_PLAYER : PlayCallout ""callout_nar_next_player_turn"", 1, VOLUME_CALLOUTS", 2, 0, 0, 10000, 0, False
									queue.Add "resume_game_1_2", "CURRENT_MODE = MODE_NORMAL : checkDrainedBalls False", 2, 0, 0, 100, 0, False

								Case 2: 'Player did not yet attempt skillshit; resume current player in skillshot mode
									queue.Add "resume_game_2_1", "DMD.triggerVideo VIDEO_RESUME_GAME_CURRENT_PLAYER : PlayCallout ""callout_nar_resume_current_player_skillshot"", 1, VOLUME_CALLOUTS", 2, 0, 0, 8000, 0, False
									queue.Add "resume_game_2_2", "startMode MODE_SKILLSHOT", 2, 0, 0, 100, 0, False

								Case 3: 'Player did not yet launch their ball from the shooter lane (but was not a skillshot); resume current player in normal mode
									queue.Add "resume_game_3_1", "DMD.triggerVideo VIDEO_RESUME_GAME_CURRENT_PLAYER : PlayCallout ""callout_nar_resume_current_player"", 1, VOLUME_CALLOUTS", 2, 0, 0, 8000, 0, False
									queue.Add "resume_game_3_2", "startMode MODE_NORMAL", 2, 0, 0, 100, 0, False

								Case 4: 'Game crashed or exited on script error; resume current player in normal mode
									queue.Add "resume_game_4_1", "DMD.triggerVideo VIDEO_RESUME_GAME_CURRENT_PLAYER : PlayCallout ""callout_nar_resume_current_player_crash"", 1, VOLUME_CALLOUTS", 2, 0, 0, 8000, 0, False
									queue.Add "resume_game_4_2", "startMode MODE_NORMAL", 2, 0, 0, 100, 0, False
							End Select

						Case MODE_SKILLSHOT
							If dataGame.data.Item("gameDifficulty") = 0 And dataGame.data.Item("ball") > 1 Then 'Player is requesting to end a Zen game
								PlaySound "ui_action_button", 1, VOLUME_SFX
								startMode MODE_ZEN_END
							End If

							If GAME_OFFICIALLY_STARTED = False And dataGame.data.Item("numPlayers") > 1 And dataGame.data.Item("gameMode") = 0 And ALLOW_COOPERATIVE_PLAY = True Then 'Coop mode
								dataGame.data.Item("gameMode") = 1
								DMD.sShowScores
								updateActionButton

								PlaySound "ui_action_button", 1, VOLUME_SFX
								PlayCallout "callout_nar_coop_activated", 1, VOLUME_CALLOUTS

								DMD.jpFlush
								DMD.jpLocked = True
								DMD.jpDMD DMD.jpCL("CO-OP MODE"), DMD.jpCL("ACTIVATED"), "", eNone, eBlinkFast, eNone, 2500, True, ""

								'Stop Scorbit; it does not support coop mode right now
								Scorbit.StopSession2 0, 0, 0, 0, 0, True
								Scorbit.ForceAsynch True
							End If

						Case MODE_DEATH_SAVE
							If IsNull(MODE_VALUES) Then 'Action button was not pressed on this heart beat yet
								If MODE_COUNTERS_A = 1 Then 'Heart beat is active; score a CPR
									DMD.deathSave_CPR
								Else 'Bad timing; score a strike
									DMD.deathSave_strike
								End If
							End If

						Case MODE_DEATH_SAVE_INTRO
							If MODE_COUNTERS_A = 1 Then PlaySound "sfx_moo", 1, VOLUME_SFX, 0, 0.25  ';)

						Case MODE_HIGH_SCORE_ENTRY
							If enabled = True Then DMD.highScoreSequence_selectChar

						Case Else
							If currentPlayer("gems") >= 6 And CURRENT_MODE >= 1000 Then deployPowerup
					End Select
				End If
		End Select

		'No need to continue for button switches
		switchTime.Item(switchNum) = GameTime
		DMD.triggerSwitch switchNum, enabled
		Exit Sub
	End If
	
	' Double-trigger prevention for certain switches
	If Enabled = True And switchNum >= 400 And switchNum < 500 And switchTime.Exists(switchNum) And GameTime < (switchTime.Item(switchNum) + 250) Then Exit Sub 'TODO: remove time condition when triggers are functional
	If Enabled = False And switchTime.Exists(switchNum) Then switchTime.Remove switchNum

	'Send switch event through B2S
	DMD.triggerSwitch switchNum, enabled
	
	' Handler for non-active gameplay
	If CURRENT_MODE < 1000 Then
		Select Case switchNum
			Case SWITCH_SCOOP:
				If enabled = True Then Scoop.TimerEnabled = True
				
			Case SWITCH_VUK:
				If enabled = True Then queue.Add "verticalKick", "SolVUpKicker True", 1, 0, 0, 0, 0, False
		End Select
		
		' Ignore all switches beyond this point when game play is not active
		switchTime.Item(switchNum) = GameTime
		Exit Sub
	End If
	
	' Mode-based operations
	If enabled = True Then
		Select Case CURRENT_MODE
			Case MODE_SKILLSHOT
				If InArray(MODE_SHOTS_A, (SwitchNum - 210)) = True Then ' Skillshot!
					MODE_SHOTS_A = Array()
					
					' Award the player
					addScore Scoring.basic("skillshot"), "Skill Shot"

					' TODO: Advance letters
					
					' Showcase the skillshot
					queue.Add "skillshot_complete","DMD.skillshot True, LTrain(" & (SwitchNum - 210) & ")",2,0,0,4000,5000,False

					' Log it in Scorbit
					If ScorbitActive = 1 Then Scorbit.SetGameMode "CP:Skillshot"
					
					'Activate our debug mode (DEBUG_MODE should be 1000 for normal mode if not debugging a mode)
					If Not DEBUG_MODE = 1000 Then queue.RemoveAll False 'Fast forward the queue so display timings are accurate
					If Not DEBUG_MODE = 0 Then startMode DEBUG_MODE
					If DEBUG_MODE = 0 Then
						currentPlayerSet "spell_BLACK", 5
						currentPlayerSet "spell_SMITH", 5
						currentPlayerSet "AC", 19
						currentPlayerSet "bonusX", 9 
						currentPlayerSet "spell_BONUSX", 5

						startMode MODE_NORMAL

						checkBlacksmith
						spellBonusX
					End If
				ElseIf Not switchNum = SWITCH_BALL_LAUNCH And switchNum >= 100 Then ' Failed skillshot
					MODE_SHOTS_A = Array()
					
					' Failed skillshot screen
					queue.Add "skillshot_failed","DMD.skillshot False, Null",2,0,0,4000,5000,False
					
					'Activate our debug mode (DEBUG_MODE should be 1000 for normal mode if not debugging a mode)
					If Not DEBUG_MODE = 1000 Then queue.RemoveAll False 'Fast forward the queue so display timings are accurate
					If Not DEBUG_MODE = 0 Then startMode DEBUG_MODE
					If DEBUG_MODE = 0 Then
						currentPlayerSet "spell_BLACK", 5
						currentPlayerSet "spell_SMITH", 5
						currentPlayerSet "AC", 19
						currentPlayerSet "bonusX", 9 
						currentPlayerSet "spell_BONUSX", 5

						startMode MODE_NORMAL

						checkBlacksmith
						spellBonusX
					End If
				End If

			Case MODE_NORMAL
				Select Case SwitchNum
					Case SWITCH_HOLE:
						If currentPlayer("AC") >= 20 And currentPlayer("obj_BLACKSMITH") = 0 Then 'Blacksmith Mini-Wizard Start
							queue.Add "blacksmith_wizard_start", "startMode MODE_BLACKSMITH_WIZARD_INTRO", 1, 0, 0, 100, 0, False
						ElseIf currentPlayer("bonusX") = 9 And currentPlayer("spell_BONUSX") = 6 Then 'Glitch Mini-wizard Start
							queue.Add "glitch_wizard_start", "startMode MODE_GLITCH_WIZARD_INTRO", 1, 0, 0, 100, 0, False
						End If
				End Select

				'Combo shots
				If SwitchNum = SWITCH_CROSSBOW_CENTER_STANDUP Or _
				SwitchNum = SWITCH_CROSSBOW_LEFT_STANDUP Or _
				SwitchNum = SWITCH_CROSSBOW_RIGHT_STANDUP Or _
				SwitchNum = SWITCH_LEFT_HORSESHOE_COMPLETE Or _
				SwitchNum = SWITCH_LEFT_ORBIT_COMPLETE Or _
				SwitchNum = SWITCH_LEFT_RAMP_COMPLETE Or _
				SwitchNum = SWITCH_RIGHT_HORSESHOE_COMPLETE Or _
				SwitchNum = SWITCH_RIGHT_ORBIT_COMPLETE Or _
				SwitchNum = SWITCH_RIGHT_RAMP_COMPLETE Or _
				SwitchNum = SWITCH_SCOOP Then
					If Clocks.data.Item("combo").timeLeft <= 0 Then
						Clocks.data.Item("combo").timeLeft = 5000
						MODE_COUNTERS_A = 0 'Track combo level
						MODE_SHOTS_A = Array() 'Track shots made
					ElseIf InArray(MODE_SHOTS_A, SwitchNum) = False Or _
					((SwitchNum = SWITCH_CROSSBOW_CENTER_STANDUP Or SwitchNum = SWITCH_CROSSBOW_LEFT_STANDUP Or SwitchNum = SWITCH_CROSSBOW_RIGHT_STANDUP) And _
					InArray(MODE_SHOTS_A, SWITCH_CROSSBOW_CENTER_STANDUP) = False And InArray(MODE_SHOTS_A, SWITCH_CROSSBOW_LEFT_STANDUP) = False And InArray(MODE_SHOTS_A, SWITCH_CROSSBOW_RIGHT_STANDUP) = False) Then
						Clocks.data.Item("combo").timeLeft = 5000 'Fresh clock
						MODE_COUNTERS_A = MODE_COUNTERS_A + 1 'Increase combo level
						MODE_SHOTS_A = AppendArray(MODE_SHOTS_A, SwitchNum) 'Note the shot that was made as it cannot be shot again for this combo
						updateShotLights False

						If MODE_COUNTERS_A > 1 Then
							addScore (MODE_COUNTERS_A * Scoring.basic("combo")), "Combo level " & (MODE_COUNTERS_A - 1)
							PlaySound "sfx_combo", 1, VOLUME_SFX, 0, 0, (1000 * (MODE_COUNTERS_A - 2))
							queue.Add "combo_shot", "DMD.triggerVideo VIDEO_COMBO", 20, 0, 0, 3000, 5000, False
						End If
					End If
				End If

			Case MODE_BLACKSMITH_WIZARD_KILL
				If SwitchNum = SWITCH_RIGHT_RAMP_COMPLETE And MODE_COUNTERS_A = 0 Then 'Right ramp completed; horseshoe is ready
					MODE_COUNTERS_A = 1
					updateShotLights False
					'TODO: Sequence
				End If

			Case MODE_BLACKSMITH_WIZARD_SPARE
				Select Case SwitchNum
					Case SWITCH_HOLE:
						targetBIP BIP + 1, true 'Bring the ball back in play

						If MODE_COUNTERS_A > 0 Then 'Raid the Blacksmith Jackpot
							'Collect Jackpot
							addScore MODE_COUNTERS_A, "Raid the Blacksmith: Jackpot"
							MODE_VALUES = MODE_VALUES + MODE_COUNTERS_A

							'DMD
							PlaySound "sfx_wood_snap", 0, VOLUME_SFX
							queue.Add "blacksmith_raid_jackpot", "DMD.blacksmithWizardSpareJackpot " & MODE_COUNTERS_A, 100, 0, 0, 3000, 10000, True

							'Light Sequence
							StandardSeq.CenterX = Hole.X
							StandardSeq.CenterY = Hole.Y
							StartLightSequence StandardSeq, Array(dpurple, purple), 5, SeqClockRightOn, 25, 1
							queueB.Add "blacksmith_raid_jackpot_seq_stop", "StopLightSequence StandardSeq", 100, 2000, 0, 100, 0, True
							For i = 0 to 9
								flasherQueue.Add "blacksmith_raid_jackpot_" & i & "_a", "Flash6 True", 100, 0, 0, 150, 0, False
								flasherQueue.Add "blacksmith_raid_jackpot_" & i & "_b", "Flash6 False", 100, 0, 0, 150, 0, False
							Next

							'Reset blacksmith spelling progress and jackpot
							currentPlayerSet "spell_BLACK", 0
							currentPlayerSet "spell_SMITH", 0
							MODE_COUNTERS_A = 0
							updateBlacksmithLights
							checkHoleState

							'Update status
							DMD.gameStatus "JACKPOT: 0 / RAID TOTAL: " & FormatScore(MODE_VALUES)
						End If
				End Select
		End Select
	End If
	
	' Switch-based operations / basic scoring
	If enabled = True Then
		Select Case switchNum
			Case SWITCH_LEFT_RAMP_ENTER, SWITCH_RIGHT_RAMP_ENTER:
				addScore Scoring.basic("ramp.enter"), "Entered a ramp"
				
			Case SWITCH_LEFT_ORBIT_ENTER, SWITCH_RIGHT_ORBIT_ENTER:
				addScore Scoring.basic("orbit.enter"), "Entered an orbit"
				
			Case SWITCH_STANDUP_T, SWITCH_STANDUP_R, SWITCH_STANDUP_A, SWITCH_STANDUP_I, SWITCH_STANDUP_N:
				addScore Scoring.basic("trainStandup"), "Hit a TRAIN standup"
				If CURRENT_MODE < 2000 Then
					PlaySound "sfx_standup", 1, VOLUME_SFX, 0, 0.2
					If switchNum = SWITCH_STANDUP_T Then LampC.Pulse StandupsTL, 2
					If switchNum = SWITCH_STANDUP_R Then LampC.Pulse StandupsRL, 2
					If switchNum = SWITCH_STANDUP_A Then LampC.Pulse StandupsAL, 2
					If switchNum = SWITCH_STANDUP_I Then LampC.Pulse StandupsIL, 2
					If switchNum = SWITCH_STANDUP_N Then LampC.Pulse StandupsNL, 2
				End If
				
			Case SWITCH_CROSSBOW_LEFT_STANDUP, SWITCH_CROSSBOW_CENTER_STANDUP, SWITCH_CROSSBOW_RIGHT_STANDUP:
				addScore Scoring.basic("crossbow"), "Hit a Crossbow standup"
				If CURRENT_MODE < 2000 Then
					PlaySound "sfx_crossbow", 1, VOLUME_SFX, -0.2, 0.2
					LampC.Pulse CrossbowShotL, 2
					If switchNum = SWITCH_CROSSBOW_LEFT_STANDUP Then LampC.Pulse CrossbowLeftStandupL, 2
					If switchNum = SWITCH_CROSSBOW_CENTER_STANDUP Then LampC.Pulse CrossbowCenterStandupL, 2
					If switchNum = SWITCH_CROSSBOW_RIGHT_STANDUP Then LampC.Pulse CrossbowRightStandupL, 2
				End If
				checkGemShot 2
				
			Case SWITCH_DROPS_PLUS, SWITCH_DROPS_H, SWITCH_DROPS_P:
				addScore Scoring.basic("hpTarget"), "Hit a +HP Drop Target"
				If CURRENT_MODE < 2000 Then
					PlaySound "sfx_drop_target", 1, VOLUME_SFX, 0, 0.3
					If switchNum = SWITCH_DROPS_PLUS Then LampC.Pulse DropTargetsPlusL, 2
					If switchNum = SWITCH_DROPS_H Then LampC.Pulse DropTargetsHL, 2
					If switchNum = SWITCH_DROPS_P Then LampC.Pulse DropTargetsPL, 2
				End If

				updateHPLights

				If DTDropped(SWITCH_DROPS_PLUS) = True And DTDropped(SWITCH_DROPS_H) = True And DTDropped(SWITCH_DROPS_P) = True Then 'All targets dropped
					addScore Scoring.basic("hpTargets.complete"), "Completed +HP Drop Targets"
					Clocks.data.Item("hpTargets").timeLeft = 10000
					Dim hpCompleted: hpCompleted = Health.HP("heal.dropTargets.completed")
					adjustHP hpCompleted(dataGame.data.Item("gameDifficulty"))
					'TODO: SFX / PUP
				End If

			Case SWITCH_HP_DING_WALL
				addScore Scoring.basic("hpDingWall"), "Hit the +HP Ding Wall"
				If DTDropped(SWITCH_DROPS_PLUS) = True And DTDropped(SWITCH_DROPS_H) = True And DTDropped(SWITCH_DROPS_P) = True Then 'All HP targets dropped
					Dim hpBonus: hpBonus = Health.HP("heal.dropTargets.bonus")
					adjustHP hpBonus(dataGame.data.Item("gameDifficulty"))
					'TODO: SFX / PUP
				End If
				
			Case SWITCH_LEFT_HORSESHOE_ENTER, SWITCH_RIGHT_HORSESHOE_ENTER:
				addScore Scoring.basic("horseshoe.enter"), "Entered the horseshoe"
				
			Case SWITCH_SCOOP:
				addScore Scoring.basic("scoop"), "Entered the scoop"
				If BIP < 2 Then 'In multiball, we do not want to wait for the main queue to eject the ball, so use the secondary queue instead.
					queue.Add "scoopKick", "DMD.SolScoopWarning", 1, 1000, 0, -1, 0, False 'doNextItem is in the Scoop timer
				Else
					queueB.Add "scoopKick", "DMD.SolScoopWarning", 1, 1000, 0, -1, 0, False 'doNextItem is in the Scoop timer
				End If
				If CURRENT_MODE < 2000 Then
					PlaySound "sfx_scoop", 1, VOLUME_SFX, 0.5, 0.2
					LampC.Pulse ScoopShotL, 2
				End If
				checkGemShot 5
				
			'TODO
			Case SWITCH_UPPER_LEFT_INLANE
				addScore Scoring.basic("bumperInlane"), "Went through the left bumper inlane"
				If bumperInlanes(0) = False And (CurrentPlayer("bonusX") < 9 Or CurrentPlayer("spell_BONUSX") < 6 Or CURRENT_MODE = MODE_GLITCH_WIZARD) Then
					bumperInlanes(0) = True
					'If CurrentPlayer("bonusX") > 2 And Clocks.data.Item("bumperInlanes").timeLeft = 0 And Not CURRENT_MODE = MODE_GLITCH_WIZARD Then Clocks.data.Item("bumperInlanes").timeLeft = (90000 - (10000 * (CurrentPlayer("bonusX") - 3)))
				End If
				updateBumperInlanes

			Case SWITCH_UPPER_CENTER_INLANE
				addScore Scoring.basic("bumperInlane"), "Went through the center bumper inlane"
				If bumperInlanes(1) = False And (CurrentPlayer("bonusX") < 9 Or CurrentPlayer("spell_BONUSX") < 6 Or CURRENT_MODE = MODE_GLITCH_WIZARD) Then
					bumperInlanes(1) = True
					'If CurrentPlayer("bonusX") > 2 And Clocks.data.Item("bumperInlanes").timeLeft = 0 And Not CURRENT_MODE = MODE_GLITCH_WIZARD Then Clocks.data.Item("bumperInlanes").timeLeft = (90000 - (10000 * (CurrentPlayer("bonusX") - 3)))
				End If
				updateBumperInlanes

			Case SWITCH_UPPER_RIGHT_INLANE
				addScore Scoring.basic("bumperInlane"), "Went through the right bumper inlane"
				If bumperInlanes(2) = False And (CurrentPlayer("bonusX") < 9 Or CurrentPlayer("spell_BONUSX") < 6 Or CURRENT_MODE = MODE_GLITCH_WIZARD) Then
					bumperInlanes(2) = True
					'If CurrentPlayer("bonusX") > 2 And Clocks.data.Item("bumperInlanes").timeLeft = 0 And Not CURRENT_MODE = MODE_GLITCH_WIZARD Then Clocks.data.Item("bumperInlanes").timeLeft = (90000 - (10000 * (CurrentPlayer("bonusX") - 3)))
				End If
				updateBumperInlanes
				
			Case SWITCH_RED_BUMPER:
				addScore Scoring.basic("bumper"), "Hit the red bumper"
				If CURRENT_MODE >= 1000 Then PlaySound "sfx_bumper", 1, VOLUME_SFX, AudioPan(Bumper1), 0.25
				LampC.Pulse bumperlightcontrol1, 0
				
			Case SWITCH_GREEN_BUMPER:
				addScore Scoring.basic("bumper"), "Hit the green bumper"
				If CURRENT_MODE >= 1000 Then PlaySound "sfx_bumper", 1, VOLUME_SFX, AudioPan(Bumper2), 0.25
				LampC.Pulse bumperlightcontrol2, 0
				
			Case SWITCH_BLUE_BUMPER:
				addScore Scoring.basic("bumper"), "Hit the blue bumper"
				If CURRENT_MODE >= 1000 Then PlaySound "sfx_bumper", 1, VOLUME_SFX, AudioPan(Bumper3), 0.25
				LampC.Pulse bumperlightcontrol3, 0
				
			Case SWITCH_LEFT_SLINGSHOT:
				addScore Scoring.basic("slingshot"), "Hit the left slingshot"
				If CURRENT_MODE >= 1000 Then PlaySound "sfx_slingshot", 1, VOLUME_SFX, -0.5, 0.1
				
			Case SWITCH_RIGHT_SLINGSHOT:
				addScore Scoring.basic("slingshot"), "Hit the right slingshot"
				If CURRENT_MODE >= 1000 Then PlaySound "sfx_slingshot", 1, VOLUME_SFX, 0.5, 0.1
				
			Case SWITCH_LEFT_OUTLANE
				addScore Scoring.basic("outlane"), "Went down the left outlane"

				'Stop the ball shield from expiring for up to 5 seconds to allow an outlane ball to be shielded... but only if we are not within the grace period
				If Clocks.data.Item("shield").timeLeft > 2000 Then 
					Clocks.data.Item("shield").canExpire = False
					LeftOutlane.TimerEnabled = False
					LeftOutlane.TimerEnabled = True
				End If

				'Are we poisoned? Increase drain damage!
				If Clocks.data.Item("shield").timeLeft < 2000 And currentModeHasInfiniteBallShield = False And currentPlayer("AC") < 20 Then
					If currentPlayer("poisonL") = 1 Then
						PlaySound "sfx_poison", 1, VOLUME_SFX
						If dataGame.data.Item("gameDifficulty") < 2 Then 'Zen and Easy
							currentPlayerSet "drainDamage", currentPlayer("drainDamage") + 2
						ElseIf dataGame.data.Item("gameDifficulty") >= 4 Then 'Hard and Impossible
							currentPlayerSet "drainDamage", currentPlayer("drainDamage") + 4
						Else 'Normal
							currentPlayerSet "drainDamage", currentPlayer("drainDamage") + 3
						End If
						queue.Add "poisoned_outlane", "DMD.triggerVideo VIDEO_POISONED_OUTLANE", 11, 0, 0, 3000, 10000, False
						currentPlayerSet "poisonL", 0
						lampC.Pulse LeftOutlanePoisonL, 6
					Else
						currentPlayerSet "poisonL", 1
						lampC.Pulse LeftOutlanePoisonL, 3
					End If
					updatePoisonLights
				End If

			Case SWITCH_RIGHT_OUTLANE
				addScore Scoring.basic("outlane"), "Went down the right outlane"

				'Stop the ball shield from expiring for up to 5 seconds to allow an outlane ball to be shielded... but only if we are not within the grace period
				If Clocks.data.Item("shield").timeLeft > 2000 Then 
					Clocks.data.Item("shield").canExpire = False
					RightOutlane.TimerEnabled = False
					RightOutlane.TimerEnabled = True
				End If

				'Are we poisoned? Increase drain damage!
				If Clocks.data.Item("shield").timeLeft < 2000 And currentModeHasInfiniteBallShield = False And currentPlayer("AC") < 20 Then
					If currentPlayer("poisonR") = 1 Then
						PlaySound "sfx_poison", 1, VOLUME_SFX
						If dataGame.data.Item("gameDifficulty") < 2 Then 'Zen and Easy
							currentPlayerSet "drainDamage", currentPlayer("drainDamage") + 2
						ElseIf dataGame.data.Item("gameDifficulty") >= 4 Then 'Hard and Impossible
							currentPlayerSet "drainDamage", currentPlayer("drainDamage") + 4
						Else 'Normal
							currentPlayerSet "drainDamage", currentPlayer("drainDamage") + 3
						End If
						queue.Add "poisoned_outlane", "DMD.triggerVideo VIDEO_POISONED_OUTLANE", 11, 0, 0, 3000, 10000, False
						currentPlayerSet "poisonR", 0
						lampC.Pulse RightOutlanePoisonL, 6
					Else
						currentPlayerSet "poisonR", 1
						lampC.Pulse RightOutlanePoisonL, 3
					End If
					updatePoisonLights
				End If
				
			Case SWITCH_LEFT_INLANE:
				addScore Scoring.basic("inlane"), "Went through the left inlane"
				
				'Activate right orbit bumper diverter
				If Not CURRENT_MODE = MODE_GLITCH_WIZARD Then
					SolRBumpDiverter True
					Clocks.data.Item("rightBumperDiverter").timeLeft = 5000
					RightOrbitBumpersL.BlinkInterval = 250
					LampC.Blink RightOrbitBumpersL
				End If
				
			Case SWITCH_RIGHT_INLANE:
				addScore Scoring.basic("inlane"), "Went through the right inlane"
				
				'Activate left orbit bumper diverter
				If Not CURRENT_MODE = MODE_GLITCH_WIZARD Then
					SolLBumpDiverter True
					Clocks.data.Item("leftBumperDiverter").timeLeft = 5000
					LeftOrbitBumpersL.BlinkInterval = 250
					LampC.Blink LeftOrbitBumpersL
				End If
				
			Case SWITCH_SPINNER:
				addScore Scoring.basic("spinner"), "Spinner"
				
			Case SWITCH_LEFT_RAMP_COMPLETE:
				addScore Scoring.basic("ramp.complete"), "Went up the left ramp"
				If CURRENT_MODE < 2000 Then
					PlaySound "sfx_ramp_complete_1", 1, VOLUME_SFX, -0.5, 0.2
					LampC.Pulse LeftRampShotL, 2
				End If
				checkGemShot 0
				
			Case SWITCH_RIGHT_RAMP_COMPLETE:
				addScore Scoring.basic("ramp.complete"), "Went up the right ramp"
				If CURRENT_MODE < 2000 Then
					PlaySound "sfx_ramp_complete_2", 1, VOLUME_SFX, 0.2, 0.2
					LampC.Pulse RightRampShotL, 2
				End If
				checkGemShot 3
				
			Case SWITCH_LEFT_ORBIT_COMPLETE:
				addScore Scoring.basic("orbit.complete"), "Completed the left orbit"
				If CURRENT_MODE < 2000 Then
					PlaySound "sfx_orbit_complete", 1, VOLUME_SFX, -0.4, 0.2
					LampC.Pulse LeftOrbitShotL, 2
				End If
				checkGemShot 1
				
			Case SWITCH_RIGHT_ORBIT_COMPLETE:
				addScore Scoring.basic("orbit.complete"), "Completed the right orbit"
				If CURRENT_MODE < 2000 Then
					PlaySound "sfx_orbit_complete", 1, VOLUME_SFX, 0.4, 0.2
					LampC.Pulse RightOrbitShotL, 2
				End If
				checkGemShot 4
				
			Case SWITCH_LEFT_HORSESHOE_COMPLETE:
				addScore Scoring.basic("horseshoe.complete"), "Completed the left horseshoe"
				spellBlacksmith False
				If CURRENT_MODE < 2000 Then
					PlaySound "sfx_horseshoe", 1, VOLUME_SFX, 0, 0.2
					LampC.Pulse LeftHorseshoeShotL, 2
				End If
				
			Case SWITCH_RIGHT_HORSESHOE_COMPLETE:
				addScore Scoring.basic("horseshoe.complete"), "Completed the right horseshoe"
				spellBlacksmith True
				If CURRENT_MODE < 2000 Then
					PlaySound "sfx_horseshoe", 1, VOLUME_SFX, 0.2, 0.2
					LampC.Pulse RightHorseshoeShotL, 2
				End If
		End Select
	ElseIf Enabled = False Then
	End If

	' Log last time the switch was hit as now
	switchTime.Item(switchNum) = GameTime
End Sub

Sub triggerLeftFlipper(Enabled)
	Dim i
	Dim tempBumperInlanes(3)

	If Enabled = True Then
		LeftFlipperPressed = True

		' Choose difficulty
		If CURRENT_MODE = MODE_CHOOSE_DIFFICULTY And dataGame.data.Item("gameDifficulty") > 0 Then
			DMD.selectDifficulty (dataGame.data.Item("gameDifficulty") - 1)
		End If
	
		If CURRENT_MODE >= 1000 Then 
			SolLFlipper True

			'Alternate bumper inlanes to the left
			For i = 0 to 2
				tempBumperInlanes(i) = bumperInlanes(i)
			Next
			bumperInlanes(0) = tempBumperInlanes(1)
			bumperInlanes(1) = tempBumperInlanes(2)
			bumperInlanes(2) = tempBumperInlanes(0)
			updateBumperInlanes
		End If
	Else
		LeftFlipperPressed = False
		If CURRENT_MODE >= 1000 Then SolLFlipper False
	End If
End Sub

Sub triggerRightFlipper(Enabled)
	Dim i
	Dim tempBumperInlanes(3)

	If Enabled = True Then
		RightFlipperPressed = True

		' Choose difficulty
		If CURRENT_MODE = MODE_CHOOSE_DIFFICULTY And dataGame.data.Item("gameDifficulty") < 4 Then
			DMD.selectDifficulty (dataGame.data.Item("gameDifficulty") + 1)
		End If
		
		If CURRENT_MODE >= 1000 Then
			SolRFlipper True
			SolUFlipper True

			'Alternate bumper inlanes to the right
			For i = 0 to 2
				tempBumperInlanes(i) = bumperInlanes(i)
			Next
			bumperInlanes(0) = tempBumperInlanes(2)
			bumperInlanes(1) = tempBumperInlanes(0)
			bumperInlanes(2) = tempBumperInlanes(1)
			updateBumperInlanes
		End If
	Else
		RightFlipperPressed = False
		If CURRENT_MODE >= 1000 Then
			SolRFlipper False
			SolUFlipper False
		End If
	End If
End Sub

Sub reverseFlippers(Enabled)	'Activate or de-activate the reversing of the flippers
	flippersReversed = Enabled
	If Enabled = True Then
		SolLFlipper RightFlipperPressed
		SolRFlipper LeftFlipperPressed
		SolUFlipper LeftFlipperPressed
	Else
		SolLFlipper LeftFlipperPressed
		SolRFlipper RightFlipperPressed
		SolUFlipper RightFlipperPressed
	End If
End Sub

'*****************************************
'    INITIALIZATION
'*****************************************

Sub InitLampsNF()
	
	'Filtering (comment out to disable)
	Lampz.Filter = "LampFilter"    'Puts all lamp intensityscale output (no callbacks) through this function before updating
	
	'Lampz Assignments
	'  In a ROM based table, the lamp ID is used to set the state of the Lampz objects
	Dim l, idx
	idx = 0
	
	' add GI
	For Each l In GI
		Lampz.MassAssign(idx) = l
		idx = idx + 1
	Next
	
	' Add all the other lights individually
	For Each l In aLights
		Lampz.MassAssign(idx) = l
		idx = idx + 1
	Next
	
	'Adjust fading speeds (1 / full MS fading time)
	Dim x
	For x = 0 To 150
		Lampz.FadeSpeedUp(x) = 1 / 50
		Lampz.FadeSpeedDown(x) = 1 / 100
		Lampz.Modulate(x) = 1 / 100
	Next
	
	LampC.RegisterLights "Lampz"
	
	Lampz.Init
	Lampz.Update

	' Light sequence operations
	If DEBUG_COMPILE_LIGHT_SHOWS = True Then LampC.LoadLightShows
	lampC.CreateSeqRunner "accomplishment"
	lampC.CreateSeqRunner "attract"
	lampC.CreateSeqRunner "shots"
	lampC.CreateSeqRunner "GI"
End Sub

Sub Table1_Init()
	'Warn about having no displays when applicable
	If Table1.ShowDT = False And usePUP = False And UseFlexDMD = 0 Then msgBox "Warning! usePUP and UseFlexDMD are disabled in the table script, and you are not running in desktop mode. You will not have a display for Wallden."

	'Init displays
	DMD.Init
	DMD.sInitialization

	'Ball initializations need for physical trough and captive ball
	Set CBall1 = CaptiveBall.CreateSizedballWithMass(Ballsize / 2,Ballmass)
	CaptiveBall.kick 180, 1
	
	Set ETBall1 = swTrough1.CreateSizedballWithMass(Ballsize / 2,Ballmass)
	Set ETBall2 = swTrough2.CreateSizedballWithMass(Ballsize / 2,Ballmass)
	Set ETBall3 = swTrough3.CreateSizedballWithMass(Ballsize / 2,Ballmass)
	Set ETBall4 = swTrough4.CreateSizedballWithMass(Ballsize / 2,Ballmass)
	Set ETBall5 = swTrough5.CreateSizedballWithMass(Ballsize / 2,Ballmass)
	
	'*** Use gBOT in the script wherever BOT is normally used. Then there is no need for GetBalls calls ***
	gBOT = Array(CBall1, ETBall1, ETBall2, ETBall3, ETBall4, ETBall5)
	
	' Init flipper physics correction
	InitPolarity
	
	' Peg drops should be dropped by default (we set Z to 0 in VPX so physics work correctly)
	Dim pegDrop
	For Each pegDrop In pegDrops
		pegDrop.Z =  - 53
		pegDrop.collidable = 0
	Next
	
	' Initialize flashers
	InitFlasher 5, "green"
	InitFlasher 6, "purple"
	InitFlasher 3, "orange"
	InitFlasher 4, "yellow"
	InitFlasher 1, "red"
	InitFlasher 2, "blue"
	InitFlasher 7, "white"
	InitFlasher 8, "white"
	
	' initialize bumpers
	FlInitBumper 1, "red"
	FlInitBumper 2, "green"
	FlInitBumper 3, "blue"
	
	' Enable the frame timer
	FrameTimer.Enabled = True
	
	' Bumper diverters
	SolLBumpDiverter False
	SolRBumpDiverter False
	Clocks.data.Item("leftBumperDiverter").expiryCallback = "leftBumperDiverter_expired"
	Clocks.data.Item("rightBumperDiverter").expiryCallback = "rightBumperDiverter_expired"

	' HP targets
	Clocks.data.Item("hpTargets").expiryCallback = "resetHPTargets"

	' Bumper Inlanes
	Clocks.data.Item("bumperInlanes").expiryCallback = "expireBumperInlanes"

	' Mode timer
	Clocks.data.Item("mode").expiryCallback = "expireMode"

	' Combos
	Clocks.data.Item("combo").expiryCallback = "updateShotLights False"

	' Shield
	Clocks.data.Item("shield").expiryCallback = "shield_expired"

	' Impossible HP
	Clocks.data.Item("impossibleHP").expiryCallback = "impossibleHP_expired"
End Sub

' VPX was not letting me set coil ramp up in the editor. Force it in the script.
Sub LeftFlipper_Init()
	LeftFlipper.rampup = 2.5
End Sub
Sub RightFlipper_Init()
	RightFlipper.rampup = 2.5
End Sub

'*****************************************
'   EXIT
'*****************************************

Sub Table1_Exit()
	' Determine game resume status
	If Err.Number <> 0 Then 'Exiting on a script error; set status so player does not get a penalty on resume (resume current player no penalty on MODE_NORMAL)
		dataGame.data.Item("resumeStatus") = 4
	ElseIf CURRENT_MODE = MODE_SKILLSHOT And BIPL = 1 Then 'Player had not yet attempted the skillshot and the ball was still in the shooter lane (resume current player no penalty on MODE_SKILLSHOT)
		dataGame.data.Item("resumeStatus") = 2
	ElseIf BIP = 1 And BIPL = 1 Then 'Player's ball was in the shooter lane (resume current player no penalty on MODE_NORMAL)
		dataGame.data.Item("resumeStatus") = 3
	ElseIf CURRENT_MODE = MODE_DEATH_SAVE_INTRO Or CURRENT_MODE = MODE_DEATH_SAVE Or CURRENT_MODE = MODE_END_OF_TURN Or CURRENT_MODE = MODE_END_OF_GAME_BONUS Then 'Player drained; go to next player without penalty
		dataGame.data.Item("resumeStatus") = 1
	Else 'Current player should be penalized for terminating the game in the middle of play with a drain penalty
		dataGame.data.Item("resumeStatus") = 0
	End If

	saveData

	' Close DMDs
	DMD.close

	' Cancel a Scorbit session
	If ScorbitActive = 1 Then Scorbit.StopSession2 dataPlayer(0).data.Item("score"), dataPlayer(1).data.Item("score"), dataPlayer(2).data.Item("score"), dataPlayer(3).data.Item("score"), dataGame.data.Item("numPlayers"), true

	'Let the user know the game was saved when applicable (uncomment if you are interested in this; disabled as it may pose a problem for vcab users)
	'	If dataGame.data.Item("ball") > 0 And dataGame.data.Item("tournamentMode") = 0 Then MsgBox "Game progress saved to the VP Registry! Next time you load this table, you can resume your game by pressing the action button when prompted."
End Sub

'*****************************************
'    TABLE EVENTS
'*****************************************

Sub Table1_KeyDown(ByVal keycode)
	If keycode = PlungerKey Then
		Plunger.PullBack
		SoundPlungerPull()
	End If
	
	If keycode = LeftFlipperKey Then
		doSwitch SWITCH_LEFT_FLIPPER_BUTTON, True
	End If
	
	If keycode = RightFlipperKey Then
		doSwitch SWITCH_RIGHT_FLIPPER_BUTTON, True
	End If
	
	If keycode = LeftTiltKey Then
		DoNudge 90, 1.5
		SoundNudgeLeft()
	End If
	
	If keycode = RightTiltKey Then
		DoNudge 270, 1.5
		SoundNudgeRight()
	End If
	
	If keycode = CenterTiltKey Then
		DoNudge 0, 1.5
		SoundNudgeCenter()
	End If

	If keycode = MechanicalTilt Then
		CheckMechTilt
	End If
	
	If keycode = AddCreditKey Then
		doSwitch SWITCH_COIN_1, True
	End If

	If keycode = AddCreditKey2 Then
		doSwitch SWITCH_COIN_2, True
	End If
	
	If keycode = StartGameKey Then
		soundStartButton()
		doSwitch SWITCH_START_BUTTON, True
	End If

	If keycode = LockbarKey Then
		doSwitch SWITCH_ACTION_BUTTON, True
	End If
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If keycode = PlungerKey Then
		Plunger.Fire
		
		If BIPL = 1 Then
			SoundPlungerReleaseBall()                        'Plunger release sound when there is a ball in shooter lane
		Else
			SoundPlungerReleaseNoBall()                        'Plunger release sound when there is no ball in shooter lane
		End If
	End If
	
	If keycode = LeftFlipperKey Then
		doSwitch SWITCH_LEFT_FLIPPER_BUTTON, False
	End If
	
	If keycode = RightFlipperKey Then
		doSwitch SWITCH_RIGHT_FLIPPER_BUTTON, False
	End If

	If keycode = StartGameKey Then
		doSwitch SWITCH_START_BUTTON, False
	End If

	If keycode = LockbarKey Then
		doSwitch SWITCH_ACTION_BUTTON, False
	End If
End Sub

'*****************************************
'    COLLIDE EVENTS
'*****************************************

Sub LeftFlipper_Collide(parm)
	CheckLiveCatch ActiveBall, LeftFlipper, LFCount, parm
	LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
	CheckLiveCatch ActiveBall, RightFlipper, RFCount, parm
	RightFlipperCollide parm
End Sub

Sub UpperFlipper_Collide(parm)
	RightFlipperCollide parm
End Sub

'*****************************************
'    HIT / UNHIT EVENTS
'*****************************************

Sub Scoop_Hit()
	SoundSaucerLock
	BallInScoop = True
	updateClocks
	doSwitch SWITCH_SCOOP, True
End Sub

Sub VerticalUpKicker_Hit()
	SoundSaucerLock
	BallInKicker = True
	updateClocks
	doSwitch SWITCH_VUK, True
End Sub

Sub Drain_Hit()
	RandomSoundDrain Drain
	doSwitch SWITCH_DRAIN, True
End Sub

Sub Hole_Hit()
	doSwitch SWITCH_HOLE, True
End Sub

Sub TriggerLF_Hit()
	LF.Addball ActiveBall
End Sub
Sub TriggerLF_UnHit()
	LF.PolarityCorrect ActiveBall
End Sub
Sub TriggerRF_Hit()
	RF.Addball ActiveBall
End Sub
Sub TriggerRF_UnHit()
	RF.PolarityCorrect ActiveBall
End Sub

Sub HP1a_Hit
	DTHit SWITCH_DROPS_PLUS
End Sub

Sub HP2a_Hit
	DTHit SWITCH_DROPS_H
End Sub

Sub HP3a_Hit
	DTHit SWITCH_DROPS_P
End Sub

Sub Trigger1_Hit()
	doSwitch SWITCH_BALL_LAUNCH, True
End Sub

Sub Trigger1_Unhit()
	doSwitch SWITCH_BALL_LAUNCH, False
End Sub

Sub StandupT_Hit()
	STHit SWITCH_STANDUP_T
End Sub

Sub StandupR_Hit()
	STHit SWITCH_STANDUP_R
End Sub

Sub StandupA_Hit()
	STHit SWITCH_STANDUP_A
End Sub

Sub StandupI_Hit()
	STHit SWITCH_STANDUP_I
End Sub

Sub StandupN_Hit()
	STHit SWITCH_STANDUP_N
End Sub

Sub StandupCL_Hit()
	STHit SWITCH_CROSSBOW_LEFT_STANDUP
End Sub

Sub StandupCM_Hit()
	STHit SWITCH_CROSSBOW_CENTER_STANDUP
End Sub

Sub StandupCR_Hit()
	STHit SWITCH_CROSSBOW_RIGHT_STANDUP
End Sub

Sub Bumper1_Hit()
	SolBumper1 True
	doSwitch SWITCH_RED_BUMPER, True
End Sub

Sub Bumper2_Hit()
	SolBumper2 True
	doSwitch SWITCH_GREEN_BUMPER, True
End Sub

Sub Bumper3_Hit()
	SolBumper3 True
	doSwitch SWITCH_BLUE_BUMPER, True
End Sub

Sub BumperLeftInlane_Hit()
	doSwitch SWITCH_UPPER_LEFT_INLANE, True
End Sub

Sub BumperCenterInlane_Hit()
	doSwitch SWITCH_UPPER_CENTER_INLANE, True
End Sub

Sub BumperRightInlane_Hit()
	doSwitch SWITCH_UPPER_RIGHT_INLANE, True
End Sub

Sub BumperLeftInlane_Unhit()
	doSwitch SWITCH_UPPER_LEFT_INLANE, False
End Sub

Sub BumperCenterInlane_Unhit()
	doSwitch SWITCH_UPPER_CENTER_INLANE, False
End Sub

Sub BumperRightInlane_Unhit()
	doSwitch SWITCH_UPPER_RIGHT_INLANE, False
End Sub

Sub LeftRampA_Hit()
	If ActiveBall.VelY < 0 Then doSwitch SWITCH_LEFT_RAMP_ENTER, True
End Sub

Sub RightRampA_Hit()
	If ActiveBall.VelY < 0 Then doSwitch SWITCH_RIGHT_RAMP_ENTER, True
End Sub

Sub LeftOrbitA_Hit()
	If ActiveBall.VelY < 0 Then doSwitch SWITCH_LEFT_ORBIT_ENTER, True
End Sub

Sub RightOrbitA_Hit()
	If ActiveBall.VelY < 0 Then doSwitch SWITCH_RIGHT_ORBIT_ENTER, True
End Sub

Sub LeftOrbitB_Hit()
	If ActiveBall.VelX > 0 Then doSwitch SWITCH_LEFT_ORBIT_COMPLETE, True
End Sub

Sub RightOrbitB_Hit()
	If ActiveBall.VelX < 0 Then doSwitch SWITCH_RIGHT_ORBIT_COMPLETE, True
End Sub

Sub LeftHorseshoe_Hit()
	If ActiveBall.VelY < 0 Then doSwitch SWITCH_LEFT_HORSESHOE_ENTER, True
End Sub

Sub RightHorseshoe_Hit()
	If ActiveBall.VelY < 0 Then doSwitch SWITCH_RIGHT_HORSESHOE_ENTER, True
End Sub

Sub Spinner_Spin()
	SoundSpinner Spinner
	doSwitch SWITCH_SPINNER, True
End Sub

Sub LeftOutlane_Hit()
	doSwitch SWITCH_LEFT_OUTLANE, True
End Sub

Sub RightOutlane_Hit()
	doSwitch SWITCH_RIGHT_OUTLANE, True
End Sub

Sub LeftOutlane_Unhit()
	doSwitch SWITCH_LEFT_OUTLANE, False
End Sub

Sub RightOutlane_Unhit()
	doSwitch SWITCH_RIGHT_OUTLANE, False
End Sub

Sub LeftInlane_Hit()
	doSwitch SWITCH_LEFT_INLANE, True
End Sub

Sub RightInlane_Hit()
	doSwitch SWITCH_RIGHT_INLANE, True
End Sub

Sub LeftInlane_Unhit()
	doSwitch SWITCH_LEFT_INLANE, False
End Sub

Sub RightInlane_Unhit()
	doSwitch SWITCH_RIGHT_INLANE, False
End Sub

' Biff bars
Sub Primitive68_Hit()
	RandomSoundFlipperBallGuide
End Sub
Sub Primitive69_Hit()
	RandomSoundFlipperBallGuide
End Sub

' Ramp sounds
Sub LeftRampSoundPlastic_Hit()
	If ActiveBall.VelY < 0 Then WireRampOn True
End Sub
Sub LeftRampSoundPlastic_Unhit()
	If ActiveBall.VelY > 0 Then WireRampOff
End Sub
Sub RightRampSoundPlastic_Hit()
	If ActiveBall.VelY < 0 Then WireRampOn True
End Sub
Sub RightRampSoundPlastic_Unhit()
	If ActiveBall.VelY > 0 Then WireRampOff
End Sub
Sub LeftRampSoundOff_Unhit()
	WireRampOff
End Sub
Sub RightRampSoundOff_Unhit()
	WireRampOff
End Sub

Sub LeftRampComplete_Hit()
	If ActiveBall.VelY > 0 Then
		WireRampOn False
		doSwitch SWITCH_LEFT_RAMP_COMPLETE, True
	End If
End Sub
Sub RightRampComplete_Hit()
	If ActiveBall.VelY > 0 Then
		WireRampOn False
		doSwitch SWITCH_RIGHT_RAMP_COMPLETE, True
	End If
End Sub

Sub HorseshoeComplete_Hit()
	If ActiveBall.VelX >= 0 Then doSwitch SWITCH_LEFT_HORSESHOE_COMPLETE, True
	If ActiveBall.VelX < 0 Then doSwitch SWITCH_RIGHT_HORSESHOE_COMPLETE, True
End Sub

Sub pegDrops_Hit(Index)
	ExecuteGlobal "doSwitch SWITCH_POST_DROP_" & index & ", True"
End Sub

Sub HPDingWall_Hit()
	doSwitch SWITCH_HP_DING_WALL, True
End Sub

Sub LeftRampAOff_Unhit()
	doSwitch SWITCH_LEFT_RAMP_ENTER, False
End Sub

Sub LeftOrbitAOff_Unhit()
	doSwitch SWITCH_LEFT_ORBIT_ENTER, False
End Sub

Sub LeftOrbitBOff_Unhit()
	doSwitch SWITCH_LEFT_ORBIT_COMPLETE, False
End Sub

Sub RightOrbitBOff_Unhit()
	doSwitch SWITCH_RIGHT_ORBIT_COMPLETE, False
End Sub

Sub LeftHorseshoeOff_Unhit()
	doSwitch SWITCH_LEFT_HORSESHOE_ENTER, False
End Sub

Sub RightRampAOff_Unhit()
	doSwitch SWITCH_RIGHT_RAMP_ENTER, False
End Sub

Sub HorseshoeCompleteOff_Unhit()
	doSwitch SWITCH_LEFT_HORSESHOE_COMPLETE, False
	doSwitch SWITCH_RIGHT_HORSESHOE_COMPLETE, False
End Sub

Sub RightHorseshoeOff_Unhit()
	doSwitch SWITCH_RIGHT_HORSESHOE_ENTER, False
End Sub

Sub RightOrbitAOff_Unhit()
	doSwitch SWITCH_RIGHT_ORBIT_ENTER, False
End Sub

Sub Gate002_Hit()
	doSwitch SWITCH_SUBWAY_GATE, True
End Sub

'*****************************************
'   SLINGSHOTS
'*****************************************

Sub LeftSlingShot_Slingshot()
	SolLSling True
	SolLSling False
	doSwitch SWITCH_LEFT_SLINGSHOT, True
End Sub

Sub RightSlingShot_Slingshot()
	SolRSling True
	SolRSling False
	doSwitch SWITCH_RIGHT_SLINGSHOT, True
End Sub

'*****************************************
'    TIMER EVENTS
'*****************************************

' Update ball velocity calculations
Sub BallVelocityTimer_Timer()
	Cor.Update
	
	' Ball rolling sounds
	RollingUpdate
End Sub

' Update drop target animations
Sub DTAnim_Timer()
	DoDTAnim
	DoSTAnim
	
	' Peg Drop Animations
	Dim pegDrop
	For Each pegDrop In pegDrops
		If pegDrop.UserValue = "raise" Then
			pegDrop.collidable = 1
			pegDrop.Z = pegDrop.Z + 5.3
			If pegDrop.Z >= 15.9 Then ' allows for a nice pop-up animation when it exceeds 0 for a short moment
				pegDrop.Z = 0
				pegDrop.UserValue = ""
			End If
			
			' Pop up the ball if it's on top of the raising target
			Dim b
			For b = 0 To UBound(gBOT)
				If InRotRect(gBOT(b).x,gBOT(b).y,pegDrop.x, pegDrop.y, pegDrop.rotz, - 25, - 10,25, - 10,25,25, - 25,25) And gBOT(b).z < pegDrop.z + 53 + 25 Then
					gBOT(b).velz = 20
				End If
			Next
			
		ElseIf pegDrop.UserValue = "lower" Then
			pegDrop.Z = pegDrop.Z - 5.3
			If pegDrop.Z <= - 53 Then
				pegDrop.Z =  - 53
				pegDrop.UserValue = ""
				pegDrop.collidable = 0
			End If
		End If
	Next
End Sub

Sub ClockTimer_Timer()
	' Tick our queues
	queue.Tick
	queueB.Tick
	troughQueue.Tick
	flasherQueue.Tick
	postDropQueue.Tick
End Sub

Sub FrameTimer_Timer()
	If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
	FlipperVisualUpdate ' Update flipper shadows
	
	'Update Music
	Music.Tick
	Tracks.Tick

	' Tick our Clocks
	Clocks.Tick
End Sub

Sub FPS30Timer_Timer()
	Dim i
	'Update our lamp controllers in 30 FPS for better performance and smoother pulse animation
	lampC.Update
	Lampz.Update2
	
	'Update light intervals - Bumper orbit diversion
	LightIntervalByTime LeftOrbitBumpersL, Clocks.data.Item("leftBumperDiverter").timeLeft, 5000
	LightIntervalByTime RightOrbitBumpersL, Clocks.data.Item("rightBumperDiverter").timeLeft, 5000

	'Update light intervals - HP targets
	If Clocks.data.Item("hpTargets").timeLeft > 0 Then
		LightIntervalByTime DropTargetsPlusL, Clocks.data.Item("hpTargets").timeLeft, 10000
		LightIntervalByTime DropTargetsHL, Clocks.data.Item("hpTargets").timeLeft, 10000
		LightIntervalByTime DropTargetsPL, Clocks.data.Item("hpTargets").timeLeft, 10000
	End If

	'Update light intervals - Bumper Inlanes
	If Clocks.data.Item("bumperInlanes").timeLeft > 0 And Not CURRENT_MODE = MODE_GLITCH_WIZARD Then
		For i = 0 to 2
			If bumperInlanes(i) = True Then LightIntervalByTime lanternLights(i), Clocks.data.Item("bumperInlanes").timeLeft, 15000
		Next
	End If

	'Update light intervals - combos
	If CURRENT_MODE = MODE_NORMAL Then 'Must be a parent if to the MODE_COUNTERS_A check to prevent type errors.
		If Clocks.data.Item("combo").timeLeft > 0 And MODE_COUNTERS_A > 0 Then
			For Each i In shotLights
				LightIntervalByTime i, Clocks.data.Item("combo").timeLeft, 5000
			Next
		End If
	End If

	'Update light intervals - Shield
	If Clocks.data.Item("shield").timeLeft > 2000 Then
		LightIntervalByTime ShieldL, (Clocks.data.Item("shield").timeLeft - 2000), 10000
	ElseIf Clocks.data.Item("shield").timeLeft > 0 Then
		shield_lightExpired
	End If
End Sub

Sub ShakerMotorTimer_Timer()
	If CURRENT_MODE = MODE_FAULT Then Exit Sub
	
	Nudge (Rnd() * 360), 0.5
End Sub

Sub coinRejected_Timer()
	coinRejected.Text = ""
	coinRejected.timerEnabled = False
End Sub

Sub FlasherFlash1_Timer()
	FlashFlasher(1)
End Sub
Sub FlasherFlash2_Timer()
	FlashFlasher(2)
End Sub
Sub FlasherFlash3_Timer()
	FlashFlasher(3)
End Sub
Sub FlasherFlash4_Timer()
	FlashFlasher(4)
End Sub
Sub FlasherFlash5_Timer()
	FlashFlasher(5)
End Sub
Sub FlasherFlash6_Timer()
	FlashFlasher(6)
End Sub
Sub FlasherFlash7_Timer()
	FlashFlasher(7)
End Sub
Sub FlasherFlash8_Timer()
	FlashFlasher(8)
End Sub

' This trough timer checks for queued balls we want to release and releases them
Sub swTrough1_Timer()
	If CURRENT_MODE = MODE_FAULT Then 
		swTrough1.TimerEnabled = False
		BQueue = 0
		Exit Sub
	End If

	If BIP >= BQueue Then
		swTrough1.TimerEnabled = False
		BQueue = 0
		Exit Sub
	End If
	
	If BQueue > 0 And swTrough1.BallCntOver > 0 And BIPL = 0 And CURRENT_MODE >= 1000 Then
		solRelease True
		updateClocks
		swTrough1.TimerEnabled = False 'To prevent more balls from being ejected right away
		If BIP >= BQueue Then BQueue = 0 'No more balls to launch
	End If
End Sub

Sub Scoop_Timer()
	Scoop.timerEnabled = False
	SolScoop True
	If queue.qCurrentItem = "scoopKick" Then queue.DoNextItem
	If queueB.qCurrentItem = "scoopKick" Then queueB.DoNextItem
End Sub

' Controls when to trigger the ball search feature for a missing ball
Sub BallSearchTimer_Timer()
	If Not BALL_SEARCH_STAGE = 0 And ((GameTime - BALL_SEARCH_TIME) < 10000 Or Not BIPL = 0 Or (CURRENT_MODE < 800 Or (CURRENT_MODE >= 900 And CURRENT_MODE < 1000)) Or LeftFlipperPressed = True Or RightFlipperPressed = True) Then
		If DEBUG_LOG_BALL_SEARCH = True Then WriteToLog "Ball Search", "Ball search terminated (flipper pressed, activity detected, or gameplay mode does not warrant ball monitoring)."
		BALL_SEARCH_STAGE = 0
		updateClocks

		'TODO: ball search termination
	End If
	
	' Only trigger ball search when gameplay is active, no balls are in the shooter lane, queue is empty, and flippers are lowered.
	If (CURRENT_MODE >= 1000 Or (CURRENT_MODE >= 800 And CURRENT_MODE < 900)) And BIPL = 0 And Queue.qItems.Count = 0 And LeftFlipperPressed = False And RightFlipperPressed = False Then
		If (GameTime - BALL_SEARCH_TIME) >= 10000 And BALL_SEARCH_STAGE = 0 Then ' Do not trigger ball search until 10 seconds of inactivity
			DMD.searchForBall
		End If
	ElseIf BALL_SEARCH_STAGE = 0 Or Not BIPL = 0 Or (CURRENT_MODE < 800 Or (CURRENT_MODE >= 900 And CURRENT_MODE < 1000)) Or LeftFlipperPressed = True Or RightFlipperPressed = True Then
		BALL_SEARCH_TIME = GameTime
		BALL_SEARCH_STAGE = 0
	End If
End Sub

Sub LeftOutlane_Timer()
	LeftOutlane.TimerEnabled = False
	If RightOutlane.TimerEnabled = False Then Clocks.data.Item("shield").canExpire = True
End Sub

Sub RightOutlane_Timer()
	RightOutlane.TimerEnabled = False
	If LeftOutlane.TimerEnabled = False Then Clocks.data.Item("shield").canExpire = True
End Sub

Sub deathSaveTimer_Timer()
	deathSaveTimer.Enabled = False

	If MODE_COUNTERS_B(1) >= 3 Or MODE_COUNTERS_B(0) <= 0 Then Exit Sub 'Do not re-enable the timer / continue if we already finished

	Select Case MODE_COUNTERS_A
		Case 0 'Heart beat starts
			MODE_COUNTERS_A = 1
			LightsOn actionLights, Array(dwhite, white), 100

			LightsOff objectiveLights
			LampC.LightOnWithColor ObjectiveCastleL, Array(dwhite, white)
			LampC.LightOnWithColor ObjectiveDragonL, Array(dwhite, white)
			LampC.LightOnWithColor ObjectiveChaserL, Array(dwhite, white)
			LampC.LightOnWithColor ObjectiveEscapeL, Array(dwhite, white)
			LampC.LightOnWithColor ObjectiveVikingL, Array(dwhite, white)
			LampC.LightOnWithColor ObjectiveBlacksmithL, Array(dwhite, white)
			LampC.LightOnWithColor ObjectiveSniperL, Array(dwhite, white)

			If MODE_COUNTERS_D = True And MODE_COUNTERS_C > 0 Then 
				PlayCallout "callout_nar_" & (MODE_COUNTERS_C - 2), 1, VOLUME_CALLOUTS
				MODE_COUNTERS_C = MODE_COUNTERS_C - 1
				If MODE_COUNTERS_C = 0 Then 
					startMode MODE_DEATH_SAVE
				Else
					DMD.deathSave_countdown (MODE_COUNTERS_C - 1)
				End If
			End If

			If dataGame.data.Item("gameDifficulty") = 0 Then
				deathSaveTimer.Interval = 300
				PlaySound "sfx_heart_60", 1, VOLUME_SFX
			Else
				Select Case (currentPlayer("deathSaves") + ((dataGame.data.Item("gameDifficulty") - 1) * 2))
					Case 1:
						deathSaveTimer.Interval = 300
						PlaySound "sfx_heart_60", 1, VOLUME_SFX
					Case 2:
						deathSaveTimer.Interval = 233
						PlaySound "sfx_heart_80", 1, VOLUME_SFX
					Case 3:
						deathSaveTimer.Interval = 187
						PlaySound "sfx_heart_100", 1, VOLUME_SFX
					Case 4:
						deathSaveTimer.Interval = 150
						PlaySound "sfx_heart_120", 1, VOLUME_SFX
					Case 5:
						deathSaveTimer.Interval = 133
						PlaySound "sfx_heart_140", 1, VOLUME_SFX
				End Select
			End If

		Case 1 'Heart beat expired
			MODE_COUNTERS_A = 2
			LightsOff actionLights

			LightsOff objectiveLights
			LampC.LightOnWithColor ObjectiveFinalBossL, Array(dred, red)

			'Did the player not press the button in time? That's a strike!
			If IsNull(MODE_VALUES) And MODE_COUNTERS_C <= 0 Then DMD.deathSave_strike

			If dataGame.data.Item("gameDifficulty") = 0 Then
				deathSaveTimer.Interval = 500
			Else
				Select Case (currentPlayer("deathSaves") + ((dataGame.data.Item("gameDifficulty") - 1) * 2))
					Case 1:
						deathSaveTimer.Interval = 350
					Case 2:
						deathSaveTimer.Interval = 258
					Case 3:
						deathSaveTimer.Interval = 206
					Case 4:
						deathSaveTimer.Interval = 175
					Case 5:
						deathSaveTimer.Interval = 147
				End Select
			End If

		Case 2 'Reset strike grace
			MODE_COUNTERS_A = 0
			LightsOff actionLights

			MODE_VALUES = Null

			If dataGame.data.Item("gameDifficulty") = 0 Then
				deathSaveTimer.Interval = 500
			Else
				Select Case (currentPlayer("deathSaves") + ((dataGame.data.Item("gameDifficulty") - 1) * 2))
					Case 1:
						deathSaveTimer.Interval = 350
					Case 2:
						deathSaveTimer.Interval = 259
					Case 3:
						deathSaveTimer.Interval = 207
					Case 4:
						deathSaveTimer.Interval = 175
					Case 5:
						deathSaveTimer.Interval = 148
				End Select
			End If
	End Select

	deathSaveTimer.Enabled = True
End Sub

Sub DMDEffectTimer_Timer()
    DMDEffectTimer.Enabled = False
    DMD.jpProcessEffectOn()
End Sub

Sub DMDTimer_Timer()
    DMDTimer.Enabled = False
	DMD.jpAdvance
End Sub
'*****************************************
'    WROU: GENERAL WALLDEN ROUTINES
'*****************************************

' LEGACY
Sub DTAction(switchid)
	doSwitch switchid, True
End Sub

' LEGACY
Sub STAction(switchid)
	doSwitch switchid, True
End Sub

Sub targetBIP(ballsToBeInPlay, autoPlunge) 'Indicate we want to launch balls until we reach the specified number of balls in play.
	'Determine how many balls we need to queue; exit if there are none to queue
	If BQueue >= ballsToBeInPlay Then 
		If DEBUG_LOG_TROUGH = True Then WriteToLog "trough", "Requested " & ballsToBeInPlay & " balls in play; " & BIP & " balls in play now and " & BQueue & " balls were already queued. Did not queue any more balls."
		Exit Sub
	End If

	'Activate auto plunge? (Only set to true; don't set to false if autoPlunge is false because if BAutoPlunge is true, it should stay true until all balls fired by auto plunge)
	If autoPlunge = True Then BAutoPlunge = True
	If swTrough1.TimerEnabled = False And BIPL = 0 Then swTrough1.TimerEnabled = True

	If DEBUG_LOG_TROUGH = True Then WriteToLog "trough", "Requested " & ballsToBeInPlay & " balls in play; " & BIP & " balls in play now and " & BQueue & " balls were already queued; Auto plunge " & BAutoPlunge & ". Trough timer enabled " & swTrough1.TimerEnabled

	'Queue up necessary balls
	BQueue = ballsToBeInPlay
End Sub

Sub checkDrainedBalls(takeDamage) ' Call when a ball is drained either physically or virtually
	'Calculate number of balls that were drained
	Dim ballsDrained 
	ballsDrained = BIPWhenDrained - BIP

	Dim BIPAndQueue

	'We do not want to prematurely end modes when balls are still queued to launch
	If BIP > BQueue Then
		BIPAndQueue = BIP
	Else
		BIPAndQueue = BQueue
	End If

	If DEBUG_LOG_TROUGH = True Then WriteToLog "Trough", "Running drained ball checks; BIP when drained = " & BIPWhenDrained & "; BIP = " & BIP & "; Balls drained = " & ballsDrained & "; Effective balls in play = " & BIPAndQueue

	' Actions that should always occur
	updateClocks

	'Modes (multiballs) which have infinite shield
	Select Case CURRENT_MODE
		'5-ball multiballs where no damage is received
		Case MODE_GLITCH_WIZARD
			targetBIP 5, True
			Exit Sub
	End Select

	'Waiting on all balls to drain
	If CURRENT_MODE >= 800 And CURRENT_MODE < 900 Then
		If BIPAndQueue < 1 Then
			BAutoPlunge = False
			DMD.jpLocked = False
			DMD.jpFlush
			Select Case CURRENT_MODE
				'Modes with an outtro sequence
				Case MODE_GLITCH_WIZARD_END:
					queue.Add "drain_wait_done", "startMode MODE_GLITCH_WIZARD_OUTTRO", 1, 0, 1000, 100, 0, False

				'All other modes do not have an outtro; start normal game play
				Case Else
					queue.Add "drain_wait_done", "startMode MODE_NORMAL", 1, 0, 1000, 100, 0, False
			End Select
		End If
		Exit Sub
	End If
	
	' Ball saver intercept (do not proceed past the if condition if ball shield was active)
	If CURRENT_MODE >= 1000 And (Clocks.data.Item("shield").timeLeft > 0 Or (BIPAndQueue < 1 And DEBUG_INFINITE_SHIELD = True)) Then
		If BIPAndQueue < 1 Then 'Do not annoy players with ball shield screen / sfx in multiball
			PlaySound "sfx_ball_shield", 1, VOLUME_SFX
			queue.Add "drain_ballsave","DMD.triggerVideo VIDEO_BALL_SHIELD",1,0,0,3000,5000,True
		End If

		targetBIP BIPWhenDrained, True 'Get back to the number of BIP we had when the balls first drained

		'Reset the outlane timers which stop the shield from expiring when a ball hits the outlane
		If LeftOutlane.TimerEnabled = True Then
			LeftOutlane.TimerEnabled = False
		ElseIf RightOutlane.TimerEnabled = True Then
			RightOutlane.TimerEnabled = False
		End If
		If LeftOutlane.TimerEnabled = False And RightOutlane.TimerEnabled = False Then Clocks.data.Item("shield").canExpire = True

		Exit Sub
	End If

	'Any further non-gameplay modes should ignore ball drains; startMode will recover a ball if necessary
	If CURRENT_MODE < 1000 Then
		Exit Sub
	End If

	'Modes (multiballs) which do not have infinite shield but do other things when balls are drained
	Select Case CURRENT_MODE
		Case MODE_BLACKSMITH_WIZARD_KILL
			currentPlayerSet "AC", currentPlayer("AC") - ballsDrained 'Every drained ball = loss of AC

			If currentPlayer("AC") > 0 Then 'AC still left, so bring back the drained balls
				queue.Add "blacksmith_wizard_kill_lost_AC", "DMD.blacksmithWizardKillLostAC", 100, 0, 0, 5000, 5000, True
				targetBIP BIPWhenDrained, True
			Else 'No AC left; end the battle with a failure.
				startMode MODE_BLACKSMITH_WIZARD_KILL_FAIL
			End If

			Exit Sub

		Case MODE_BLACKSMITH_WIZARD_SPARE
			If ballsDrained >= 0 Then
				If BIPAndQueue = 1 Then
					Clocks.data.Item("mode").timeLeft = 10000
					PlayCallout "callout_nar_timerunningout", 1, VOLUME_CALLOUTS
					queue.Add "last_hurrah", "DMD.triggerVideo VIDEO_LAST_HURRAH", 2, 0, 0, 2500, 5000, True
				ElseIf BIPAndQueue < 1 Then 'Drained the ball before hurrah expired
					expireMode
				End If
			End If
	End Select
	
	'At this point, if this is the last ball in play, we are taking drain damage, activating death save, and/or moving to the next player when applicable
	If BIPAndQueue < 1 And ballsDrained >= 0 Then
		If takeDamage = True And Not CURRENT_MODE = MODE_ZEN_END Then 
			adjustHP -((currentPlayer("drainDamage")) * ballsDrained)
		Else
			adjustHP 0 'Take 0 damage so we can still trigger death saves etc when necessary
		End If
		
		' Only trigger the end-player sequence if not in death save (triggered by adjustHP); death save will trigger this accordingly after it is done
		If Not CURRENT_MODE = MODE_DEATH_SAVE_INTRO Then
			startMode MODE_END_OF_TURN
		End If
	End If
End Sub

' Called whenever the main queue is emptied
Sub queue_empty()
	Select Case CURRENT_MODE
		Case MODE_CHOOSE_DIFFICULTY
			' Go back to the difficulty selection screen
			'DMD.triggerVideo VIDEO_CLEAR
			DMD.selectDifficulty dataGame.data.Item("gameDifficulty")
			
		Case MODE_SKILLSHOT
			' Show the skillshot instructions
			'DMD.triggerVideo VIDEO_CLEAR
			If GAME_OFFICIALLY_STARTED = False Then
				DMD.triggerVideo VIDEO_SKILLSHOT_ADD_PLAYERS
			ElseIf dataGame.data.Item("gameDifficulty") = 0 And dataGame.data.Item("ball") > 1 Then 'Zen; player can end the game with action button
				DMD.triggerVideo VIDEO_SKILLSHOT_ZEN_END
			Else
				DMD.triggerVideo VIDEO_SKILLSHOT
			End If
			
		Case MODE_NORMAL
			'DMD.triggerVideo VIDEO_CLEAR
			If currentPlayer("AC") >= 20 And CURRENT_MODE = MODE_NORMAL And currentPlayer("obj_BLACKSMITH") = 0 Then
				DMD.triggerVideo VIDEO_BLACKSMITH_WIZARD_READY
			ElseIf currentPlayer("bonusX") = 9 And currentPlayer("spell_BONUSX") = 6 Then
				DMD.triggerVideo VIDEO_GLITCH_WIZARD_READY
			Else
				DMD.triggerVideo VIDEO_GAMEPLAY_NORMAL
			End If

		Case MODE_GLITCH_WIZARD
			DMD.triggerVideo VIDEO_GLITCH_WIZARD_SCREEN

		Case MODE_BLACKSMITH_WIZARD_KILL
			DMD.triggerVideo VIDEO_BLACKSMITH_WIZARD_KILL
	End Select

	'Sometimes all balls drained before we entered a "waiting for balls to drain" mode. Account for this.
	If CURRENT_MODE >= 800 And CURRENT_MODE < 900 And BIP < 1 Then checkDrainedBalls True
	
	checkIdleFlashers
End Sub

' Called whenever the flasher queue is emptied
Sub flasherQueue_empty()
	checkIdleFlashers
End Sub

'Check if any flashers should idly be flashing
Sub checkIdleFlashers()
	' Flashers for ready Mini-Wizards
	If CURRENT_MODE = MODE_NORMAL Then
		If currentPlayer("AC") >= 20 And currentPlayer("obj_BLACKSMITH") = 0 Then 'BLACKSMITH mini wizard is ready; idly flash the purple flasher
			flasherQueue.Add "blacksmith_wizard_ready_on", "Flash6 True", 1, 0, 0, 100, 0, False
			flasherQueue.Add "blacksmith_wizard_ready_off", "Flash6 False", 1, 0, 0, 900, 0, False
		ElseIf currentPlayer("bonusX") = 9 And currentPlayer("spell_BONUSX") = 6 Then 'Glitch Mini-wizard is ready; idly flash the green flasher
			flasherQueue.Add "glitch_wizard_ready_on", "Flash5 True", 1, 0, 0, 100, 0, False
			flasherQueue.Add "glitch_wizard_ready_off", "Flash5 False", 1, 0, 0, 900, 0, False
		End If
	End If
End Sub

Sub addPlayer() 'Add a new player to the game
	Dim currentPlayers
	currentPlayers = dataGame.data.Item("numPlayers")
	Dim gameDifficulty
	gameDifficulty = dataGame.data.Item("gameDifficulty")
	
	If currentPlayers >= 4 Then Exit Sub ' No more than 4 players!
	
	' Prepare the new player
	currentPlayers = currentPlayers + 1
	dataGame.data.Item("numPlayers") = currentPlayers
	
	' Play a callout when this is player 2-4
	If currentPlayers > 1 Then
		' TODO
		DMD.jpFlush
		DMD.jpLocked = True
		DMD.jpDMD DMD.jpCL("PLAYER " & currentPlayers), DMD.jpCL("Added"), "", eNone, eBlinkFast, eNone, 2500, True, ""
	End If

	' Coop mode callout for player 2
	If currentPlayers = 2 And ALLOW_COOPERATIVE_PLAY = True Then
		PlayCallout "callout_nar_toggle_coop", 1, VOLUME_CALLOUTS
		updateActionButton
	End If

	' Assign generic name
	dataPlayer(currentPlayers - 1).data.Item("name") = "PLAYER " & currentPlayers
	
	' Assign a color
	dataPlayer(currentPlayers - 1).data.Item("colorSlot") = currentPlayers
	
	' Assign starting HP, drain damage, and Armor Class
	Dim startingHP: startingHP = Health.HP("start")
	dataPlayer(currentPlayers - 1).data.Item("HP") = startingHP(gameDifficulty)
	Dim startingDrainDamage: startingDrainDamage = Health.drainDamage("start")
	dataPlayer(currentPlayers - 1).data.Item("drainDamage") = startingDrainDamage(gameDifficulty)
	Dim startingAC: startingAC = Health.AC("start")
	dataPlayer(currentPlayers - 1).data.Item("AC") = startingAC(gameDifficulty)
	
	' Assign the ones place in the score depending on game difficulty (don't use addScore because we do not want to trigger any scoring routines)
	If TOURNAMENT_MODE = 0 Then
		dataPlayer(currentPlayers - 1).data.Item("score") = gameDifficulty + 1
	Else
		dataPlayer(currentPlayers - 1).data.Item("score") = gameDifficulty + 5 'Tournament mode adds 5 to the score to separate it from non-tournament.
	End If
	
	' Update the display
	DMD.sUpdateScores True
End Sub

Sub addScore(score, scoreReason) 'Award points to the current player
	' Bail on inactive mode
	If CURRENT_MODE < 900 Then Exit Sub
	
	Dim curScore
	curScore = currentPlayer("score")
	
	' Add score to current player
	currentPlayerSet "score", curScore + CLng(score)
	
	' Update the display
	DMD.sUpdateScores False

	If DEBUG_LOG_SCORES = True Then WriteToLog "addScore", "Player " & dataGame.data.Item("playerUp") & " | " & FormatScore(score) & " points | " & scoreReason & " | score now " & FormatScore(currentPlayer("score"))
End Sub

Sub adjustHP(adjustment) 'Adjust the current player's HP
	Dim curHP
	curHP = currentPlayer("HP")
	
	' Adjust the player's HP
	currentPlayerSet "HP", curHP + CLng(adjustment)
	
	' Update the display
	DMD.sUpdateScores False

	' Update Scorbit
	ScorbitBuildGameModes
	
	If dataGame.data.Item("gameDifficulty") > 0 Then 'Do none of these routines in Zen difficulty; only the Action button ends the game
		If CLng(adjustment) < -1 Or BIP < 1 Then 'We lost HP (more than 1) or end of ball
			PlaySound "sfx_damage_received", 1, VOLUME_SFX
			If BIP < 2 And currentPlayer("HP") > 0 Then 'Do the damage sequence when not in multiball and when we are not dying
				queue.Add "drain_damage","DMD.damageReceived " & Abs(adjustment),2,0,0,2000,3000,False
			End If
		ElseIf CLng(adjustment) = -1 Then
			'TODO: Dragon mode
			PlaySound "sfx_blood_drip", 1, VOLUME_SFX, 0, 1000
		End If

		If currentPlayer("HP") <= 0 Then
			'Determine if the player has a death save, and activate it if so
			Select Case dataGame.data.Item("gameDifficulty")
				Case 1
					If currentPlayer("deathSaves") < 5 then 
						startMode MODE_DEATH_SAVE_INTRO
					Else
						startMode MODE_END_OF_TURN
					End If

				Case 2
					If currentPlayer("deathSaves") < 3 then
						startMode MODE_DEATH_SAVE_INTRO
					Else
						startMode MODE_END_OF_TURN
					End If

				Case 3
					If currentPlayer("deathSaves") < 1 then
						startMode MODE_DEATH_SAVE_INTRO
					Else
						startMode MODE_END_OF_TURN
					End If

				Case 4 'No death saves on impossible mode
					startMode MODE_END_OF_TURN
			End Select
		End If
	End If
End Sub

Sub EndGame() 'Helper method for quickly / formally marking the current game as over
	dataGame.data.Item("playerUp") = 0
	dataGame.data.Item("ball") = -1
	DMD.sUpdateScores True
End Sub

Sub updateClocks() 'Check when we need to pause or unpause certain clocks
	If CURRENT_MODE < 1000 Or BIP <= 0 Or CURRENT_MODE = MODE_SKILLSHOT Or Not BALL_SEARCH_STAGE = 0 Then 'Always pause when not in active game play, no balls are in play, ball search is active, or we are in skillshot mode
		Clocks.data.Item("shield").isPaused = True
		Clocks.data.Item("mode").isPaused = True
		Clocks.data.Item("hpTargets").isPaused = True
		Clocks.data.Item("bumperInlanes").isPaused = True
		Clocks.data.Item("impossibleHP").isPaused = True
		Exit Sub
	End If

	'Determine number of active balls in play
	Dim ballsActive
	ballsActive = BIP
	ballsActive = ballsActive - BIPL
	If BallInScoop = True Then ballsActive = ballsActive - 1
	If BallInKicker = True Then ballsActive = ballsActive - 1
	If BIPWhenDrained > 0 Then ballsActive = ballsActive - 1 'A drained ball might still be counted in BIP until the trough settles; count it as not active
	
	If ballsActive <= 0 Then 'Pause when all of the balls in play are not active
		Clocks.data.Item("shield").isPaused = True
		Clocks.data.Item("mode").isPaused = True
		Clocks.data.Item("hpTargets").isPaused = True
		Clocks.data.Item("bumperInlanes").isPaused = True
		Clocks.data.Item("combo").isPaused = True
		Clocks.data.Item("impossibleHP").isPaused = True
		Exit Sub
	End If

	'ImpossibleHP should only be unpaused if there is no shield time left
	If Clocks.data.Item("shield").timeLeft > 0 Then
		Clocks.data.Item("impossibleHP").isPaused = True
	Else
		Clocks.data.Item("impossibleHP").isPaused = False
	End If
	
	'At this point, timers should not be paused
	Clocks.data.Item("shield").isPaused = False
	Clocks.data.Item("mode").isPaused = False
	Clocks.data.Item("hpTargets").isPaused = False
	Clocks.data.Item("bumperInlanes").isPaused = False
	Clocks.data.Item("combo").isPaused = False
End Sub

Sub updateActionButton() 'Update the action button light To the state it should be
	If CURRENT_MODE = MODE_RESUME_GAME_PROMPT Then 
		LightsBlink actionLights, Array(dgreen, green), 250, 100
		Exit Sub
	End If

	'Coop mode
	If CURRENT_MODE = MODE_SKILLSHOT And GAME_OFFICIALLY_STARTED = False And ALLOW_COOPERATIVE_PLAY = True And dataGame.data.Item("numPlayers") > 1 And dataGame.data.Item("gameMode") = 0 Then
		LightsBlink actionLights, Array(dpurple, purple), 250, 100
		Exit Sub
	End If
	
	'We have a powerup waiting to be used
	If CURRENT_MODE >= 1000 And Not CURRENT_MODE = MODE_SKILLSHOT And currentPlayer("gems") > 5 Then
		LightsBlink actionLights, Array(dblue, blue), 500, 100
		Exit Sub
	End If

	'Action button on the start of a turn ends a Zen game if not the first ball / turn
	If CURRENT_MODE = MODE_SKILLSHOT And dataGame.data.Item("gameDifficulty") = 0 And dataGame.data.Item("ball") > 1 Then
		LightsBlink actionLights, Array(dred, red), 250, 100
		Exit Sub
	End If

	If CURRENT_MODE = MODE_HIGH_SCORE_ENTRY Then
		LightsBlink actionLights, Array(dgreen, green), 250, 100
		Exit Sub
	End If

	LightsOff actionLights
End Sub

Sub updateObjectiveLights()
	If CURRENT_MODE < 1000 Then Exit Sub

	LightsOff objectiveLights

	'Blacksmith
	If currentPlayer("obj_BLACKSMITH") > 0 Then	'AC might be <20 even when BLACKSMITH is complete as AC is lost in Blacksmith battles, but we want the objective light to be lit regardless
		LampC.LightLevel ObjectiveBlacksmithL, 100
		If currentPlayer("spare_BLACKSMITH") = 1 Then 'Spared / allied Blacksmith
			LampC.LightOnWithColor ObjectiveBlacksmithL, Array(dgreen, green)
		ElseIf currentPlayer("spare_BLACKSMITH") = 2 Then 'Failed Blacksmith kill
			LampC.LightOnWithColor ObjectiveBlacksmithL, Array(dpurple, purple)
		Else 'Successful Blacksmith kill
			LampC.LightOnWithColor ObjectiveBlacksmithL, Array(dred, red)
		End If
	Else 'Blacksmith was not complete yet; light objective diamond according to AC
		LampC.LightLevel ObjectiveBlacksmithL, 50
		Select Case currentPlayer("AC")
			Case 14, 15:
				LampC.LightOnWithColor ObjectiveBlacksmithL, Array(dorange, orange)
			Case 16, 17:
				LampC.LightOnWithColor ObjectiveBlacksmithL, Array(damber, amber)
			Case 18, 19:
				LampC.LightOnWithColor ObjectiveBlacksmithL, Array(dyellow, yellow)
			Case 20:
				LampC.LightOnWithColor ObjectiveBlacksmithL, Array(ddarkgreen, darkgreen)
		End Select
	End If
End Sub

Sub updateShotLights(calledFromStartMode)
	If calledFromStartMode = False Then LightsOff shotLights 'LightsOff should be called before newMode cases in startMode, and updateShotLights called after the cases.

	Select Case CURRENT_MODE
		Case MODE_BLACKSMITH_WIZARD_KILL
			LightsColor shotLights, Array(dpurple, purple)
			If MODE_COUNTERS_A = 0 Then 
				LampC.Blink RightRampShotL
				LampC.UpdateBlinkInterval RightRampShotL, 250
			Else
				If currentPlayer("spell_BLACK") < 5 Then LampC.Blink LeftHorseshoeShotL
				If currentPlayer("spell_SMITH") < 5 Then LampC.Blink RightHorseshoeShotL
				LampC.UpdateBlinkInterval LeftHorseshoeShotL, 250
				LampC.UpdateBlinkInterval RightHorseshoeShotL, 250
			End If

		Case MODE_NORMAL 'For Combo shots
			If Clocks.data.Item("combo").timeLeft > 0 And MODE_COUNTERS_A > 0 Then 'Combo is active

				'Determine which shots are enabled for the next combo
				If InArray(MODE_SHOTS_A, SWITCH_LEFT_RAMP_COMPLETE) = False Then LampC.Blink LeftRampShotL
				If InArray(MODE_SHOTS_A, SWITCH_LEFT_ORBIT_COMPLETE) = False Then LampC.Blink LeftOrbitShotL
				If InArray(MODE_SHOTS_A, SWITCH_CROSSBOW_CENTER_STANDUP) = False And InArray(MODE_SHOTS_A, SWITCH_CROSSBOW_LEFT_STANDUP) = False And InArray(MODE_SHOTS_A, SWITCH_CROSSBOW_RIGHT_STANDUP) = False Then LampC.Blink CrossbowShotL
				If InArray(MODE_SHOTS_A, SWITCH_LEFT_HORSESHOE_COMPLETE) = False Then LampC.Blink LeftHorseshoeShotL
				If InArray(MODE_SHOTS_A, SWITCH_RIGHT_RAMP_COMPLETE) = False Then LampC.Blink RightRampShotL
				If InArray(MODE_SHOTS_A, SWITCH_RIGHT_HORSESHOE_COMPLETE) = False Then LampC.Blink RightHorseshoeShotL
				If InArray(MODE_SHOTS_A, SWITCH_RIGHT_ORBIT_COMPLETE) = False Then LampC.Blink RightOrbitShotL
				If InArray(MODE_SHOTS_A, SWITCH_SCOOP) = False Then LampC.Blink ScoopShotL
			End If
	End Select
End Sub

'Check if the hole should be active for something
Sub checkHoleState()
	'TODO: Primitive for trap door

	Hole.Enabled = False

	Select Case CURRENT_MODE
		Case MODE_NORMAL
			LampC.lightOff HoleShotL
			If currentPlayer("AC") >= 20 And currentPlayer("obj_BLACKSMITH") = 0 Then 'BLACKSMITH wizard ready
				Hole.Enabled = True
				LampC.LightColor HoleShotL, Array(dpurple, purple)
				HoleShotL.BlinkPattern = "110"
				LampC.UpdateBlinkInterval HoleShotL, 100
				LampC.Blink HoleShotL
				Exit Sub
			End If
			If currentPlayer("bonusX") = 9 And currentPlayer("spell_BONUSX") = 6 Then 'Glitch wizard ready
				Hole.Enabled = True
				LampC.LightColor HoleShotL, Array(ddarkgreen, darkgreen)
				HoleShotL.BlinkPattern = "110"
				LampC.UpdateBlinkInterval HoleShotL, 100
				LampC.Blink HoleShotL
				Exit Sub
			End If

		Case MODE_GLITCH_WIZARD, MODE_GLITCH_WIZARD_INTRO:
			LampC.lightOff HoleShotL

		Case MODE_BLACKSMITH_WIZARD_SPARE
			If MODE_COUNTERS_A > 0 Then
				Hole.Enabled = True
				LampC.LightColor HoleShotL, Array(dpurple, purple)
				HoleShotL.BlinkPattern = "110"
				LampC.UpdateBlinkInterval HoleShotL, 100
				LampC.Blink HoleShotL
			Else
				LampC.lightOff HoleShotL
			End If
	End Select
End Sub

Sub addBallShield(mTime) 'Adds ball shield time to the clock (when currentModeHasInfiniteBallShield is false) and ensures a 2-second grace period
	'Exit if running a mode with unlimited shields (we should use separate logic outside this method if we want to instead add to mode time)
	If currentModeHasInfiniteBallShield = True Then Exit Sub

	If Clocks.data.Item("shield").timeLeft < 2000 Then Clocks.data.Item("shield").timeLeft = 2000	'Grace period
	Clocks.data.Item("shield").timeLeft = Clocks.data.Item("shield").timeLeft + mTime

	'Stop the slow loss of HP in impossible difficulty when shield is active
	If dataGame.data.Item("gameDifficulty") = 4 Then Clocks.data.Item("impossibleHP").timeLeft = 5000
	Clocks.data.Item("impossibleHP").isPaused = True

	'Prepare shield light
	Select Case CURRENT_MODE
		Case Else
			LampC.LightOnWithColor ShieldL, Array(dblue, blue)
			LampC.Blink ShieldL
	End Select
	LightIntervalByTime ShieldL, (Clocks.data.Item("shield").timeLeft - 2000), 10000

	updatePoisonLights
End Sub

Sub setBallShield(mTime) 'Explicitly set ball shield time (when currentModeHasInfiniteBallShield is false) and ensures a 2-second grace period
	'Exit if running a mode with unlimited shields (we should use separate logic outside this method if we want to instead add to mode time)
	If currentModeHasInfiniteBallShield = True Then Exit Sub

	Clocks.data.Item("shield").timeLeft = mTime + 2000

	'Stop the slow loss of HP in impossible difficulty when shield is active
	If dataGame.data.Item("gameDifficulty") = 4 Then Clocks.data.Item("impossibleHP").timeLeft = 5000
	Clocks.data.Item("impossibleHP").isPaused = True

	'Prepare shield light
	Select Case CURRENT_MODE
		Case Else
			LampC.LightOnWithColor ShieldL, Array(dblue, blue)
			LampC.Blink ShieldL
	End Select
	LightIntervalByTime ShieldL, (Clocks.data.Item("shield").timeLeft - 2000), 10000

	updatePoisonLights
End Sub

Sub PlayCallout(cName, cLoop, cVol)	'Stop the previously-active callout and play the indicated callout
	StopSound calloutActive
	PlaySound cName, cLoop, cVol
	calloutActive = cName
End Sub

Sub saveData 'Save data to vpReg
	Dim i
	For i = 0 To 4
		dataHighScores(i).SaveAll
	Next
	For i = 0 To 3
		dataPlayer(i).SaveAll
		dataSeeds(i).SaveAll
	Next
	dataGame.SaveAll
End Sub

Sub playNormalMusic()
	' Start music
	If dataGame.data.Item("gameDifficulty") = 0 Then
		Music.PlayTrack "music_zen_" & Int(Rnd * 3), Null, Null, Null, Null, 0, 0, 0, 0
	Else
		If dataGame.data.Item("ball") > 1 Then
			Music.PlayTrack "music_normal_" & Int(Rnd * 3), Null, Null, Null, Null, 0, 0, 0, 0
		Else
			Music.PlayTrack "music_normal_firstball", Null, Null, Null, Null, 0, 0, 0, 0
		End If
	End If
End Sub

'-----------------------------------------
'    WNUG: NUDGE ROUTINES
'-----------------------------------------
Sub DoNudge(aDir, aForce) 'Do a keyboard nudge
	aDir = aDir + (Rnd-0.5)*15*aForce : aForce = (0.6+Rnd*0.8)*aForce
	Nudge aDir, aForce
	CheckTilt aForce
End Sub

Sub CheckTilt(aForce)
	If CURRENT_MODE >= 1000 And BIPL = 0 Then
		If TiltDecreaseTimer.Enabled = True Then
			CheckMechTilt
		Else
			TiltDecreaseTimer.Enabled = True
		End If
		TiltDecreaseTimer.Enabled = True
	End If
End Sub

Sub TiltDecreaseTimer_Timer
	TiltDecreaseTimer.Enabled = False
End Sub

Sub CheckMechTilt
	If CURRENT_MODE >= 1000 And BIPL = 0 Then
		If TiltDebounceTimer.Enabled = False Then
			TiltPenalty
			TiltDebounceTimer.Enabled = True
		End If
	End If
End Sub

Sub TiltDebounceTimer_Timer
	TiltDebounceTimer.Enabled = False
End Sub

Sub TiltPenalty 'We do not have actual TILT in Wallden but rather penalties
	'Take HP damage
	adjustHP -(currentPlayer("drainDamage"))

	'Increase drain damage
	If dataGame.data.Item("gameDifficulty") < 2 Then 'Zen and Easy
		currentPlayerSet "drainDamage", currentPlayer("drainDamage") + 2
	ElseIf dataGame.data.Item("gameDifficulty") >= 4 Then 'Hard and Impossible
		currentPlayerSet "drainDamage", currentPlayer("drainDamage") + 4
	Else 'Normal
		currentPlayerSet "drainDamage", currentPlayer("drainDamage") + 3
	End If

	'Warn the player
	queue.Add "nudge_warning", "DMD.TiltWarning", 1000, 0, 0, 3000, 5000, True
End Sub

'-----------------------------------------
'    WCLK: CLOCK ROUTINES
'-----------------------------------------

Sub shield_expired()
	shield_lightExpired
	If dataGame.data.Item("gameDifficulty") = 4 Then Clocks.data.Item("impossibleHP").timeLeft = 5000
	updateClocks()
	updatePoisonLights()
End Sub

Sub shield_lightExpired()
	Select Case CURRENT_MODE
		Case MODE_BLACKSMITH_WIZARD_KILL
			LampC.LightOnWithColor ShieldL, Array(dyellow, yellow)
		Case Else
			LampC.LightOff ShieldL
	End Select
End Sub

Sub impossibleHP_expired()
	If Not dataGame.data.Item("gameDifficulty") = 4 Then Exit Sub

	If Clocks.data.Item("shield").timeLeft <= 0 Then adjustHP -1

	Clocks.data.Item("impossibleHP").timeLeft = 5000
End Sub

'-----------------------------------------
'    WMDE: MODES
'-----------------------------------------

Sub startMode(newMode) 'This is the main sub-routine for changing modes in the table
	Dim i

	If CURRENT_MODE = newMode Then Exit Sub ' Prevent starting the same mode twice.
	
	If DEBUG_LOG_MODES = True Then WriteToLog "modes", "Starting " & newMode
	
	' Disable Flippers, Bumpers, and Slingshots when mode is < 1000, else enable them.
	If newMode < 1000 Then
		SolLFlipper False
		SolRFlipper False
		SolUFlipper False
		Bumper1.HasHitEvent = False
		Bumper2.HasHitEvent = False
		' Bumper3.HasHitEvent = False (Ball can get stuck on bumper 3, so don't ever disable it)
		LeftSlingShot.HasHitEvent = False
		RightSlingShot.HasHitEvent = False
		BQueue = 0 'Reset ball queue for non-active modes as we do not want to fire any more balls
		If DEBUG_LOG_TROUGH = True Then WriteToLog "Trough", "Ball queue cleared; a non-active mode was started."
	Else
		saveData 'Save current game progress
		Bumper1.HasHitEvent = True
		Bumper2.HasHitEvent = True
		Bumper3.HasHitEvent = True
		LeftSlingShot.HasHitEvent = True
		RightSlingShot.HasHitEvent = True
	End If

	'Stop special mode shot sequences
	LampC.RemoveAllLightSeq "shots"
	
	' End tasks for current (previous) mode
	Select Case CURRENT_MODE
		Case MODE_SKILLSHOT
			LightsOff LTrain

			' Remove Scorbit QR Code
			ScorbitClaimQR False

			' Queue instruction callout if applicable
			If newMode = MODE_NORMAL Then
				If GAME_OFFICIALLY_STARTED = False Then
					queue.Add "callout_instructions","DMD.promptInstructions",1,0,0,100,0,False
				End If
			End If

			' Prohibit adding any more new players
			GAME_OFFICIALLY_STARTED = True
			
			' Unpause shield timer
			Clocks.data.Item("shield").isPaused = False

			' Advance the skillshot seed
			wrnd "skillshot", True

		Case MODE_BLACKSMITH_WIZARD_KILL
			'Reset AC to 5 if it is below 5 after the Blacksmith boss battle ends
			If currentPlayer("AC") < 5 Then currentPlayerSet "AC", 5
	End Select

	'Update current mode
	Dim prevMode 'The mode we were on prior to starting the new mode
	prevMode = CURRENT_MODE
	CURRENT_MODE = newMode

	'End / reset combo
	Clocks.data.Item("combo").timeLeft = 0
	
	'Run update routines that may change something because of the new mode (such as updating Lights)
	updateClocks
	updateActionButton
	ScorbitBuildGameModes
	updateBlacksmithLights
	updateObjectiveLights
	updatePowerupLights
	updateBumperInlanes
	updateHPLights
	checkHoleState
	updatePoisonLights
	LightsOff shotLights
	
	'Start tasks for the new mode
	Select Case newMode
		Case MODE_FAULT
			' Stop attract sequence and clear the queue
			DMD.resetTutorialSeq
			queue.RemoveAll True
			
			' Disable score overlay
			DMD.triggerOverlay OVERLAY_REMOVE
			
			' Go red
			DMD.triggerVideo VIDEO_CLEAR
			Music.StopTrack Null
			LightsOn coloredLights, Array(dred, red), 100
			LightsOn GI, Array(dred, red), 100
			
		Case MODE_ATTRACT
			If Not prevMode = MODE_GAME_OVER Then DMD.sBeginAttract 'Game over starts attract from the contributors section
			GAME_OFFICIALLY_STARTED = False
			DMD.gameStatus "FREE PLAY / PRESS START FOR NEW GAME"
			Music.PlayTrack "music_startup", Null, Null, Null, Null, 0, 0, 0, 0
			
		Case MODE_CHOOSE_DIFFICULTY
			'Reset game in case we are starting during the resume game prompt
			EndGame

			' Stop attract sequence and clear the queue
			DMD.resetTutorialSeq
			queue.RemoveAll True
			
			' Disable score overlay
			DMD.triggerOverlay OVERLAY_REMOVE
			
			' If we are forcing a difficulty, choose that difficulty and start the game immediately
			If Not TOURNAMENT_MODE = 0 Then
				dataGame.data.Item("gameDifficulty") = TOURNAMENT_MODE
				startMode MODE_START_GAME
				Exit Sub
			Else
				' Start on normal
				DMD.triggerVideo 33
				dataGame.data.Item("gameDifficulty") = 2
				StopLightSequence LightSeqAttract
				LightsOn coloredLights, Array(dyellow, yellow), 100
				LightsOn GI, Array(dyellow, yellow), 100
				
				' Play callout and start music
				Music.PlayTrack "music_choose_difficulty", Null, Null, Null, Null, 0, 0, 0, 0
				PlayCallout "callout_nar_choose_difficulty", 1, VOLUME_CALLOUTS
			End If
			
		Case MODE_START_GAME
			' Go black
			DMD.SetPage pDMD, pPageBlank
			'DMD.triggerVideo VIDEO_CLEAR
			Music.StopTrack Null
			LightsOff aLights
			LightsOff GI
			
			' Reset game and player stats (gameDifficulty, seed, and credits has to be re-loaded)
			wrandomize
			Dim gameDifficulty
			gameDifficulty = dataGame.data.Item("gameDifficulty")
			Dim gameSeed
			gameSeed = dataGame.data.Item("seed")
			dataGame.Reset = True
			dataGame.data.Item("ball") = 1
			dataGame.data.Item("gameDifficulty") = gameDifficulty
			dataGame.data.Item("seed") = gameSeed
			dataGame.data.Item("resumeStatus") = 4	'Set resume status to crash / error; it will be set appropriately when the table_exit routine is properly called.
			If gameDifficulty = 0 Or DEBUG_INFINITE_SHIELD = True Or Not DEBUG_MODE = MODE_NORMAL Then dataGame.data.Item("jester") = 1 ';)
			For i = 0 To 3
				dataPlayer(i).Reset = True
			Next
			dataGame.data.Item("tournamentMode") = TOURNAMENT_MODE
			
			'Add first player
			addPlayer
			
			' Play callout
			PlayCallout "callout_startgame_walls", 1, VOLUME_CALLOUTS
			
			' Play video intro
			DMD.triggerVideo VIDEO_START_GAME
			
			' Queue some awesomeness with our video
			queue.Add "start_game_1a","DMD.sInitialization_impact Array(dwhite, white) : DMD.sInitialization_pegDrops 0, 7, True",2,0,700,500,0,False
			queue.Add "start_game_1b","DMD.sInitialization_impactOff",2,0,0,500,0,False
			queue.Add "start_game_2a","DMD.sInitialization_impact Array(dwhite, white) : DMD.sInitialization_pegDrops 8, 14, True",2,0,0,500,0,False
			queue.Add "start_game_2b","DMD.sInitialization_impactOff",2,0,0,420,0,False
			queue.Add "start_game_3a","DMD.sInitialization_impact Array(dwhite, white) : DMD.sInitialization_pegDrops 15, 21, True",2,0,0,500,0,False
			queue.Add "start_game_3b","DMD.sInitialization_impactOff",2,0,0,500,0,False
			queue.Add "start_game_4a","DMD.sInitialization_impact Array(dwhite, white) : DMD.sInitialization_pegDrops 22, 29, True",2,0,0,500,0,False
			queue.Add "start_game_4b","DMD.sInitialization_impactOff : StartLightSequence LightSeqAttract, Null, 25, SeqRandom, 10, 1",2,0,0,630,0,False
			For i = 0 To 29
				queue.Add "start_game_5_" & i,"SolPegDrop " & i & ", False",2,0,0,100,0,False
			Next
			queue.Add "start_game_6a","StopLightSequence LightSeqAttract : SolVUKDiverter False",2,0,0,2500,0,False

			'Initialize first player
			DMD.startNextPlayer False

			'Start a Scorbit session but only if not cheating
			If DEBUG_INFINITE_SHIELD = False And DEBUG_MODE = MODE_NORMAL Then 
				Scorbit.StartSession
				If ScorbitActive = 1 Then Scorbit.SetGameMode "CP:Seed " & dataGame.data.Item("seed") 'Log the game seed in Scorbit. TODO: CP may be the wrong category
			Else
				Scorbit.ForceAsynch True
			End If
			
		Case MODE_MISSING_BALL
			'Remember our previous mode so we can return to it
			BALL_SEARCH_MODE_MEMORY = prevMode
			
			' Turn off the lights
			StopLightSequence StandardSeq
			LightsOff coloredLights
			
			'Clear the display
			DMD.triggerVideo VIDEO_CLEAR
			
			'Play callout and fade out music
			queue.Add "missing_ball_1","PlayCallout ""callout_lov_ball_search_fail"", 1, VOLUME_CALLOUTS : Music.StopTrack 5000",1,0,0,1000,0,False
			
			'L
			queue.Add "missing_ball_L1","SolPegDrop 8, True",1,0,0,1000,0,False
			queue.Add "missing_ball_L2","SolPegDrop 17, True",1,0,0,1000,0,False
			queue.Add "missing_ball_L3","SolPegDrop 27, True",1,0,0,1000,0,False
			queue.Add "missing_ball_L4","SolPegDrop 28, True",1,0,0,1000,0,False
			queue.Add "missing_ball_L5","SolPegDrop 29, True",1,0,0,1000,0,False
			
			'Start next sequence when the callout finishes
			DMD.gameStatus "GAME PAUSED"
			queue.Add "missing_ball_2","DMD.missingBallSummon",1,0,7000,100,0,False
			
		Case MODE_SKILLSHOT
			' Reset Lights
			LightsOff aLights
			LightsResetColor

			' GI
			LightsOn GI, Array(dpurple, purple), 100
			
			' Show score overlay
			DMD.SetPage pDMD, pPageDefault
			DMD.sCreditsGameplay
			DMD.sShowScores

			' Show Scorbit QR Code
			ScorbitClaimQR True
			
			' Display the mode
			If GAME_OFFICIALLY_STARTED = False Then
				DMD.triggerVideo VIDEO_SKILLSHOT_ADD_PLAYERS
			ElseIf dataGame.data.Item("gameDifficulty") = 0 And dataGame.data.Item("ball") > 1 Then 'Zen; player can end the game with action button
				DMD.triggerVideo VIDEO_SKILLSHOT_ZEN_END
			Else
				DMD.triggerVideo VIDEO_SKILLSHOT
			End If
			DMD.gameStatus "SKILLSHOT"

			playNormalMusic
			
			' Set up the shield clock
			Clocks.data.Item("shield").timeLeft = 0
			addBallShield currentPlayer("AC") * 1000
			Clocks.data.Item("shield").isPaused = True
			
			' Choose a standup for the skillshot
			Dim skillshottarget
			skillshottarget = Int(wrnd("skillshot", False) * 5)
			MODE_SHOTS_A = Array(skillshottarget) 'We do not advance the seed until after the skillshot in case the game is quit / resumed at this point
			LampC.LightColor LTrain(skillshottarget), Array(dpurple, purple)
			LTrain(skillshottarget).BlinkInterval = 200
			LTrain(skillshottarget).BlinkPattern = "110"
			LampC.Blink LTrain(skillshottarget)
			
			' Skillshot instructions
			PlayCallout "callout_nar_skillshot_instr", 1, VOLUME_CALLOUTS

			' Reset HP targets
			resetHPTargets

			' Reset Bumper Inlanes
			expireBumperInlanes
			
			' Launch a ball if necessary
			targetBIP 1, False
			
		Case MODE_NORMAL
			' Init for combo shots
			MODE_SHOTS_A = Array()
			MODE_COUNTERS_A = 0

			' GI
			LightsOn GI, Array(dbase, base), 100
			
			'Update lights
			LightsResetColor
			
			If Not prevMode = MODE_SKILLSHOT Then
				' Start music
				If dataGame.data.Item("gameDifficulty") = 0 Then
					Music.PlayTrack "music_zen_" & Int(Rnd * 3), Null, Null, Null, Null, 0, 0, 0, 0
				Else
					Music.PlayTrack "music_normal_" & Int(Rnd * 3), Null, Null, Null, Null, 0, 0, 0, 0
				End If
			End If

			DMD.gameStatus ""

			'Prompts for readied features
			If currentPlayer("AC") >= 20 And currentPlayer("obj_BLACKSMITH") = 0 Then 'Blacksmith Mini-Wizard is ready
				queue.Add "blacksmith_wizard_ready", "DMD.wizardReady_blacksmith", 1, 0, 0, 6000, 0, False
			ElseIf currentPlayer("bonusX") = 9 And currentPlayer("spell_BONUSX") = 6 Then 'Glitch Mini-wizard is ready
				queue.Add "glitch_wizard_ready", "DMD.wizardReady_glitch", 1, 0, 0, 6000, 0, False
			End If
			
		Case MODE_END_OF_GAME_BONUS
			' Perform end-of-game bonus sequence
			DMD.gameStatus "STAND BY FOR BONUS"
			queue.Add "youre_dead","DMD.playerIsDead",1,0,0,100,0,False
			
		Case MODE_END_OF_TURN
			DMD.gameStatus "BALL LOST"
			If prevMode = MODE_ZEN_END Then	'Force death and skip drain sequence if player requested to end a Zen game
				queue.Add "drain_turn_end","startMode MODE_END_OF_GAME_BONUS",3,0,100,0,0,False
			Else
				queue.Add "drain_turn_end","DMD.currentPlayerEnd",3,0,100,0,0,False
			End If

		Case MODE_GAME_OVER
			' End a Scorbit session
			If ScorbitActive = 1 Then Scorbit.StopSession dataPlayer(0).data.Item("score"), dataPlayer(1).data.Item("score"), dataPlayer(2).data.Item("score"), dataPlayer(3).data.Item("score"), dataGame.data.Item("numPlayers")

			'Formally end the game
			EndGame

			' Start game over sequence
			DMD.gameStatus "GAME OVER; ALL ADVENTURERS DEAD"

			'TODO: Check if any players completed Final Judgment and add sequence for that
			DMD.sAttractSequence_gameOverFailed

		Case MODE_RESUME_GAME_PROMPT
			' Show scores so players know what game is to be resumed
			DMD.sShowScores

			' Play music and prompt
			Music.PlayTrack "music_resume_game", Null, Null, Null, Null, 0, 0, 0, 0
			PlayCallout "callout_nar_resume_game_prompt", 1, VOLUME_CALLOUTS
			DMD.triggerVideo VIDEO_RESUME_GAME

			' Add queue for attract sequence if not resumed (also resets the game)
			queue.Add "resume_game_prompt_end","EndGame : startMode MODE_ATTRACT",1,0,22000,100,0,False

			DMD.jpFlush
			For i = 22 to 1 step -1
				DMD.jpDMD DMD.jpCL("Resume Game?"), DMD.jpCL("Press Action (" & i & ")"), "", eNone, eBlink, eNone, 1000, False, ""
			Next
			DMD.jpDMD DMD.jpCL("Resume Game?"), DMD.jpCL("Press Action (0)"), "", eBlink, eNone, eNone, 1000, True, ""

		Case MODE_ZEN_END
			'Prepare to end the game with a sequence
			Music.StopTrack 3000
			'TODO: Callout
			'DMD.triggerVideo VIDEO_CLEAR
			DMD.triggerVideo VIDEO_ZEN_END_WAIT
			DMD.gameStatus "STAND BY FOR BONUS"

			'Turn lights on a dim purple
			LightsOff coloredLights
			LightsOn coloredLights, Array(dpurple, purple), 25

			'Autofire plunger, and remove all queued balls
			bQueue = 0
			BAutoPlunge = False
			If BIPL > 0 Then troughQueue.Add "autoFire_zen", "impulseP.AutoFire", 100, 0, 500, 100, 0, False

			'Trigger drain if we already have no balls in play
			If BIP < 1 And BQueue < 1 Then checkDrainedBalls True

		Case MODE_DEATH_SAVE_INTRO
			MODE_COUNTERS_A = 0				'Used for heart beat
			MODE_COUNTERS_B = Array(15, 0) 	'Index 0 is CPRs left; index 1 is strikes
			MODE_COUNTERS_C = 5				'Countdown
			MODE_COUNTERS_D = False			'Not yet ready to count down
			MODE_VALUES = Null				'Track action button press per heart beat: Null = not pressed yet, 0 = pressed at the wrong time, 1 = pressed at the right time

			'Increment death save counter
			currentPlayerSet "deathSaves", currentPlayer("deathSaves") + 1

			'Prepare the table for the sequence
			LightsOff GI
			LightsOff aLights
			LightsOn coloredLights, Array(dred, red), 25
			DMD.gameStatus "DEATH SAVE"

			'Do some freaky stuff
			ShakerMotorTimer.Enabled = True
			'Music.trackManager.AddFade Music.nowPlaying, "volume", 0, 2000
			Music.trackManager.AddFade Music.nowPlaying, "pitch", -25000, 2500
			queueB.Add "death_save_glitch_3", "PlaySound ""sfx_scream"", 1, VOLUME_SFX : Music.PlayTrack ""music_death_save"", Null, Null, Null, Null, 0, 0, 0, 0", 100, 0, 2500, 0, 0, False
			For i = 0 to 8
				flasherQueue.Add "death_save_glitch_1_" & i, "Flash1 True : SolLFlipper True : SolRFlipper True : SolUFlipper False", 100, 2500, 0, 200, 0, False
				flasherQueue.Add "death_save_glitch_2_" & i, "Flash1 False : SolLFlipper False : SolRFlipper False : SolUFlipper True", 100, 2500, 0, 200, 0, False
			Next
			flasherQueue.Add "death_save_glitch_end", "SolUFlipper False : ShakerMotorTimer.Enabled = False", 100, 2500, 0, 200, 0, False

			'Queue up our intro sequence on a pre-queue delay so it is in sync with the music (when possible)
			queue.Add "death_save_intro", "DMD.deathSaveIntro", 2, 2500, 0, 8000, 0, False
			queue.Add "death_save_countdown_ready", "MODE_COUNTERS_D = True", 2, 2501, 0, 100, 0, False

			'Start our heart beat timer
			deathSaveTimer.Interval = 1000
			deathSaveTimer.Enabled = True

		Case MODE_DEATH_SAVE
			LightsOn objectiveLights, Array(dred, red), 100 'TODO: make a separate method for adjusting only the level
			LightsOff objectiveLights

			MODE_COUNTERS_B = Array(15, 0) 	'Index 0 is CPRs left; index 1 is strikes
			MODE_COUNTERS_C = 0				'Countdown
			MODE_COUNTERS_D = False			'Not yet ready to count down
			MODE_VALUES = Null				'Track action button press per heart beat: Null = not pressed yet, 0 = pressed at the wrong time, 1 = pressed at the right time

			'Load progress on display
			DMD.deathSave_countdown Null
			'DMD.triggerVideo VIDEO_CLEAR
			DMD.triggerVideo VIDEO_DEATH_SAVE_PROGRESS
			DMD.deathSave_progress MODE_COUNTERS_B(0), MODE_COUNTERS_B(1)

		Case MODE_GLITCH_WIZARD_INTRO
			expireBumperInlanes

			'Mark glitch wizard as having been played
			currentPlayerSet "spell_BONUSX", 7

			'Play the intro video (and hide the score overlay)
			DMD.triggerOverlay OVERLAY_REMOVE
			DMD.triggerVideo VIDEO_GLITCH_WIZARD_INTRO

			'Music
			Music.PlayTrack "music_glitch", Null, Null, Null, 500, 0, 0, 0, 0

			'GI / Lights / Shaker
			LightsOff GI
			StartLightSequence StandardSeq, Array(ddarkgreen, darkgreen), 25, seqRandom, 25, 1
			ShakerMotorTimer.Enabled = True

			'Stop flashing lights and shaking
			queue.Add "glitch_wizard_stop_shaking", "StopLightSequence StandardSeq : ShakerMotorTimer.Enabled = False", 1, 0, 3666, 0, 0, False

			'Delayed callout
			queue.Add "glitch_wizard_instructions", "PlayCallout ""callout_game_glitch_wizard"", 1, VOLUME_CALLOUTS", 1, 0, 5238, 0, 0, False

			'Start the wizard after the intro is done playing
			queue.Add "glitch_wizard_begin", "startMode MODE_GLITCH_WIZARD", 1, 0, 12429, 0, 0, False

		Case MODE_GLITCH_WIZARD:
			'Start with 1x multiplier
			MODE_COUNTERS_B = 1

			'Reset bumper inlanes
			expireBumperInlanes

			'Adjust lighting
			LightsOn GI, Array(ddarkgreen, darkgreen), 100

			'Queue balls for multiball
			targetBIP 5, True

			'Infinite ball shield
			LampC.LightOnWithColor ShieldL, Array(dblue, blue)

			'Set up the mode (add remaining shield time if applicable to the mode, and zero out shield time)
			Clocks.data.Item("mode").timeLeft = (1000 * 105) '105 seconds to coincide with the music.
			Clocks.data.Item("mode").timeLeft = Clocks.data.Item("mode").timeLeft + Clocks.data.Item("shield").timeLeft 'Add shield time to mode time
			Clocks.data.Item("shield").timeLeft = 0
			MODE_VALUES = 0 'Tally of total points scored in glitch wizard
			DMD.gameStatus "GLITCH TOTAL: 0"

			'Reset bumper diverters and permanently divert during wizard
			Clocks.data.Item("leftBumperDiverter").canExpire = True
			Clocks.data.Item("leftBumperDiverter").timeLeft = 0
			Clocks.data.Item("rightBumperDiverter").canExpire = True
			Clocks.data.Item("rightBumperDiverter").timeLeft = 0
			SolLBumpDiverter True
			SolRBumpDiverter True

			'Flash lights
			lampC.AddLightSeq "shots", lSeqGlitchWizard

			'Display scores again (and show glitch wizard screen)
			DMD.sShowScores
			DMD.triggerVideo VIDEO_GLITCH_WIZARD_SCREEN

			'Add into the secondary queue all of our glitches for the mode (these are in sync with the music!

			'First part; flashing
			For i = 0 To 7
				queueB.Add "glitch_wizard_stage_1a_" & i, "StartLightSequence StandardSeq, Array(ddarkgreen, darkgreen), 100, seqAllOn, 10, 1", 1, 2666 * i, 0, 100, 0, False
				queueB.Add "glitch_wizard_stage_1b_" & i, "StopLightSequence StandardSeq", 1, (2666 * i) + 333, 0, 100, 0, False
			Next

			'Freakout
			queueB.Add "glitch_wizard_stage_2_1", "ShakerMotorTimer.Enabled = True", 1, 21333, 0, 100, 0, False
			queueB.Add "glitch_wizard_stage_2_2", "ShakerMotorTimer.Enabled = False", 1, 24020, 0, 100, 0, False
			queueB.Add "glitch_wizard_stage_2_3", "SolRSling True : SolLSling True : Flash7 True : Flash8 True", 1, 25047, 0, 100, 0, False
			queueB.Add "glitch_wizard_stage_2_4", "SolRSling False : SolLSling False : Flash7 False : Flash8 False", 1, 25187, 0, 100, 0, False
			queueB.Add "glitch_wizard_stage_2_5", "SolRSling True : SolLSling True : Flash7 True : Flash8 True", 1, 25502, 0, 100, 0, False
			queueB.Add "glitch_wizard_stage_2_6", "SolRSling False : SolLSling False : Flash7 False : Flash8 False", 1, 25642, 0, 100, 0, False
			queueB.Add "glitch_wizard_stage_2_7", "SolRSling True : SolLSling True : Flash7 True : Flash8 True", 1, 25995, 0, 100, 0, False
			queueB.Add "glitch_wizard_stage_2_8", "SolRSling False : SolLSling False : Flash7 False : Flash8 False", 1, 26135, 0, 100, 0, False
			queueB.Add "glitch_wizard_stage_2_9", "ShakerMotorTimer.Enabled = True", 1, 26667, 0, 100, 0, False
			queueB.Add "glitch_wizard_stage_2_10", "ShakerMotorTimer.Enabled = False", 1, 27971, 0, 100, 0, False
			queueB.Add "glitch_wizard_stage_2_11", "SolRSling True : SolLSling True : Flash7 True : Flash8 True", 1, 29333, 0, 100, 0, False
			queueB.Add "glitch_wizard_stage_2_12", "SolRSling False : SolLSling False : Flash7 False : Flash8 False", 1, 29473, 0, 100, 0, False
			queueB.Add "glitch_wizard_stage_2_13", "SolRSling True : SolLSling True : Flash7 True : Flash8 True", 1, 29828, 0, 100, 0, False
			queueB.Add "glitch_wizard_stage_2_14", "SolRSling False : SolLSling False : Flash7 False : Flash8 False", 1, 29968, 0, 100, 0, False
			queueB.Add "glitch_wizard_stage_2_15", "SolRSling True : SolLSling True : Flash7 True : Flash8 True", 1, 30321, 0, 100, 0, False
			queueB.Add "glitch_wizard_stage_2_16", "SolRSling False : SolLSling False : Flash7 False : Flash8 False", 1, 30461, 0, 100, 0, False
			queueB.Add "glitch_wizard_stage_2_17", "SolRSling True : SolLSling True : Flash7 True : Flash8 True", 1, 30835, 0, 100, 0, False
			queueB.Add "glitch_wizard_stage_2_18", "SolRSling False : SolLSling False : Flash7 False : Flash8 False", 1, 30975, 0, 100, 0, False
			queueB.Add "glitch_wizard_stage_2_19", "SolBumper1 True", 1, 31333, 0, 100, 0, False
			queueB.Add "glitch_wizard_stage_2_20", "SolBumper2 True", 1, 31487, 0, 100, 0, False
			queueB.Add "glitch_wizard_stage_2_21", "SolBumper3 True", 1, 31666, 0, 100, 0, False
			queueB.Add "glitch_wizard_stage_2_22", "SolRSling True : SolLSling True : Flash7 True : Flash8 True", 1, 31704, 0, 100, 0, False
			queueB.Add "glitch_wizard_stage_2_23", "SolRSling False : SolLSling False : Flash7 False : Flash8 False", 1, 31875, 0, 100, 0, False

			'Periodic flipper pulses
			For i = 0 To 5
				queueB.Add "glitch_wizard_stage_3a_" & i, "LightsOn InlaneGI, Array(dred, red), 100", 1, 32000 + (5333 * i) + (666 * 4) + (166 * 0), 0, 100, 0, False
				queueB.Add "glitch_wizard_stage_3b_" & i, "LightsOn InlaneGI, Array(ddarkgreen, darkgreen), 100", 1, 32000 + (5333 * i) + (666 * 4) + (166 * 1), 0, 100, 0, False
				queueB.Add "glitch_wizard_stage_3c_" & i, "LightsOn InlaneGI, Array(dred, red), 100", 1, 32000 + (5333 * i) + (666 * 4) + (166 * 2), 0, 100, 0, False
				queueB.Add "glitch_wizard_stage_3d_" & i, "LightsOn InlaneGI, Array(ddarkgreen, darkgreen), 100", 1, 32000 + (5333 * i) + (666 * 4) + (166 * 3), 0, 100, 0, False
				queueB.Add "glitch_wizard_stage_3e_" & i, "SolLFlipper True : SolRFlipper True : SolUFlipper True", 1, 32000 + (5333 * i) + (666 * 6) + (166 * 0), 0, 100, 0, False
				queueB.Add "glitch_wizard_stage_3f_" & i, "SolLFlipper False : SolRFlipper False : SolUFlipper False", 1, 32000 + (5333 * i) + (666 * 6) + (166 * 1), 0, 100, 0, False
				queueB.Add "glitch_wizard_stage_3g_" & i, "SolLFlipper True : SolRFlipper True : SolUFlipper True", 1, 32000 + (5333 * i) + (666 * 6) + (166 * 2), 0, 100, 0, False
				queueB.Add "glitch_wizard_stage_3h_" & i, "SolLFlipper False : SolRFlipper False : SolUFlipper False", 1, 32000 + (5333 * i) + (666 * 6) + (166 * 3), 0, 100, 0, False
			Next

			'Double jackpot (on main queue)
			queue.Add "glitch_wizard_jackpot_doubled", "DMD.TriggerVideo VIDEO_GLITCH_JACKPOT_DOUBLED : MODE_COUNTERS_B = 2 : PlaySound ""ui_attention_low_short"", 1, (VOLUME_SFX * 0.7)", 1, 64000, 0, 4000, 0, False

			'Reverse flippers
			queueB.Add "glitch_wizard_stage_4a","reverseFlippers True : PlaySound ""sfx_glitch_matrix"", 1, VOLUME_SFX : lampC.AddLightSeq ""GI"", lSeqReversedFlippers", 1, 64000, 0, 100, 0, False
			queueB.Add "glitch_wizard_stage_4b","reverseFlippers False : lampC.RemoveAllLightSeq ""GI"" : StopSound ""sfx_glitch_matrix""", 1, 74666, 0, 100, 0, False

			'More flipper pulsing
			For i = 0 To 5
				queueB.Add "glitch_wizard_stage_5a_" & i, "Music.TrackManager.AddFade Music.nowPlaying, ""pan"", -0.5, 1333 : LightsOn LeftInlaneGI, Array(dred, red), 100", 1, 74667 + (5333 * i) + (666 * 0) + (166 * 0), 0, 100, 0, False
				queueB.Add "glitch_wizard_stage_5b_" & i, "LightsOn LeftInlaneGI, Array(ddarkgreen, darkgreen), 100", 1, 74667 + (5333 * i) + (666 * 0) + (166 * 1), 0, 100, 0, False
				queueB.Add "glitch_wizard_stage_5c_" & i, "LightsOn LeftInlaneGI, Array(dred, red), 100", 1, 74667 + (5333 * i) + (666 * 0) + (166 * 2), 0, 100, 0, False
				queueB.Add "glitch_wizard_stage_5d_" & i, "LightsOn LeftInlaneGI, Array(ddarkgreen, darkgreen), 100", 1, 74667 + (5333 * i) + (666 * 0) + (166 * 3), 0, 100, 0, False

				queueB.Add "glitch_wizard_stage_5e_" & i, "SolLFlipper True", 1, 74667 + (5333 * i) + (666 * 2) + (166 * 0), 0, 100, 0, False
				queueB.Add "glitch_wizard_stage_5f_" & i, "SolLFlipper False", 1, 74667 + (5333 * i) + (666 * 2) + (166 * 1), 0, 100, 0, False
				queueB.Add "glitch_wizard_stage_5g_" & i, "SolLFlipper True", 1, 74667 + (5333 * i) + (666 * 2) + (166 * 2), 0, 100, 0, False
				queueB.Add "glitch_wizard_stage_5h_" & i, "SolLFlipper False", 1, 74667 + (5333 * i) + (666 * 2) + (166 * 3), 0, 100, 0, False

				queueB.Add "glitch_wizard_stage_5i_" & i, "Music.TrackManager.AddFade Music.nowPlaying, ""pan"", 0.5, 1333 : LightsOn RightInlaneGI, Array(dred, red), 100", 1, 74667 + (5333 * i) + (666 * 4) + (166 * 0), 0, 100, 0, False
				queueB.Add "glitch_wizard_stage_5j_" & i, "LightsOn RightInlaneGI, Array(ddarkgreen, darkgreen), 100", 1, 74667 + (5333 * i) + (666 * 4) + (166 * 1), 0, 100, 0, False
				queueB.Add "glitch_wizard_stage_5k_" & i, "LightsOn RightInlaneGI, Array(dred, red), 100", 1, 74667 + (5333 * i) + (666 * 4) + (166 * 2), 0, 100, 0, False
				queueB.Add "glitch_wizard_stage_5l_" & i, "LightsOn RightInlaneGI, Array(ddarkgreen, darkgreen), 100", 1, 74667 + (5333 * i) + (666 * 4) + (166 * 3), 0, 100, 0, False

				queueB.Add "glitch_wizard_stage_5m_" & i, "SolRFlipper True", 1, 74667 + (5333 * i) + (666 * 6) + (166 * 0), 0, 100, 0, False
				queueB.Add "glitch_wizard_stage_5n_" & i, "SolRFlipper False", 1, 74667 + (5333 * i) + (666 * 6) + (166 * 1), 0, 100, 0, False
				queueB.Add "glitch_wizard_stage_5o_" & i, "SolRFlipper True", 1, 74667 + (5333 * i) + (666 * 6) + (166 * 2), 0, 100, 0, False
				queueB.Add "glitch_wizard_stage_5p_" & i, "SolRFlipper False", 1, 74667 + (5333 * i) + (666 * 6) + (166 * 3), 0, 100, 0, False
			Next

			'Reset pan effect
			queueB.Add "glitch_wizard_stage_6", "Music.TrackManager.AddFade Music.nowPlaying, ""pan"", 0, 1333", 1, 106667, 0, 100, 0, False

		Case MODE_GLITCH_WIZARD_END:
			'GI / Lights
			LightsOff GI
			LightsOff aLights
			DMD.gameStatus "STAND BY; DRAINING BALLS"
			Music.StopTrack 5000

			'Make sure total shows for at least 5 seconds even if balls drain before then
			queue.Add "glitch_wizard_total", "DMD.triggerVideo VIDEO_GLITCH_WIZARD_TOTAL", 1, 0, 0, 5000, 0, False

		Case MODE_GLITCH_WIZARD_OUTTRO:
			queue.Add "glitch_restart_0","DMD.triggerOverlay OVERLAY_REMOVE : DMD.triggerVideo VIDEO_GLITCH_OUTTRO",10,0,0,3500,0,False
			queue.Add "glitch_restart_1","DMD.triggerVideo VIDEO_RGB_TEST",10,0,0,2000,0,False
			queue.Add "glitch_restart_2","DMD.sInitializationStep 1",10,0,0,2500,0,False
			queue.Add "glitch_restart_3","StopLightSequence LightSeqAttract : LightsOff aLights : DMD.sShowScores : startMode MODE_NORMAL",10,0,0,100,0,False

		Case MODE_BLACKSMITH_WIZARD_INTRO:
			'Scorbit: Log that we started the Blacksmith wizard
			If ScorbitActive = 1 Then Scorbit.SetGameMode "CP:BLACKSMITH Mini-Wizard"

			'Mark objective as complete
			currentPlayerSet "obj_BLACKSMITH", 1

			'Queue up dialog and sequences
			If SKIP_DIALOG = False Then
				'GI / Lights
				LightsOff aLights
				LightsOn coloredLights, Array(dpurple, purple), 20
				LightsOn GI, Array(dpurple, purple), 20

				'Stop Music slowly
				Music.StopTrack 5000

				'Start the lore video
				DMD.triggerOverlay OVERLAY_REMOVE
				DMD.triggerVideo VIDEO_BLACKSMITH_LORE

				queue.Add "blacksmith_lore_1", "DMD.blacksmithLore 1", 1, 0, 3000, 6470, 0, False
				queue.Add "blacksmith_lore_2", "DMD.blacksmithLore 2", 1, 0, 0, 5000, 0, False
				queue.Add "blacksmith_lore_3", "DMD.blacksmithLore 3", 1, 0, 0, 3620, 0, False
				queue.Add "blacksmith_lore_4", "DMD.blacksmithLore 4", 1, 0, 0, 7655, 0, False
				queue.Add "blacksmith_lore_5", "DMD.blacksmithLore 5", 1, 0, 0, 7028, 0, False
				queue.Add "blacksmith_lore_6", "DMD.blacksmithLore 6", 1, 0, 0, 10015, 0, False
				queue.Add "blacksmith_lore_7", "DMD.blacksmithLore 7", 1, 0, 0, 8360, 0, False
			End If

			queue.Add "blacksmith_lore_select", "startMode MODE_BLACKSMITH_WIZARD_SELECTION", 1, 0, 0, 100, 0, False

		Case MODE_BLACKSMITH_WIZARD_SELECTION
			'GI / Lights
			LightsOn GI, Array(dpurple, purple), 100
			LightsOn coloredLights, Array(dpurple, purple), 100

			LightsOff objectiveLights
			LampC.AddShot "blacksmith_kill", ObjectiveBlacksmithL, Array(dred, red)
			LampC.AddShot "blacksmith_spare", ObjectiveBlacksmithL, Array(dwhite, white)
			LampC.Blink ObjectiveBlacksmithL
			LampC.UpdateBlinkInterval ObjectiveBlacksmithL, 250

			'Music
			Music.PlayTrack "music_tension", Null, Null, 0, Null, 0, 0, 0, 0

			'Callout
			PlayCallout "callout_nar_make_selection", 1, VOLUME_CALLOUTS

			'DMD
			DMD.triggerVideo VIDEO_BLACKSMITH_LORE_STILL
			DMD.triggerOverlay OVERLAY_KILL_OR_SPARE
			DMD.jpLocked = False
			DMD.jpFlush
			DMD.jpDMD DMD.jpCL("BLACKSMITH"), "", "d_E217_0", eBlink, eNone, eNone, 3000, False, ""

		Case MODE_BLACKSMITH_WIZARD_INTRO_KILL
			'Stop Music slowly
			Music.StopTrack 5000

			'GI / Lights
			LampC.RemoveShotsFromLight ObjectiveBlacksmithL
			LightsOn GI, Array(dpurple, purple), 25
			LightsOn coloredLights, Array(dpurple, purple), 25

			'DMD
			DMD.triggerOverlay OVERLAY_REMOVE
			DMD.triggerVideo VIDEO_BLACKSMITH_LORE_KILL
			DMD.jpLocked = False
			DMD.jpFlush
			DMD.jpDMD DMD.jpCL("BLACKSMITH"), "", "d_E218_0", eBlink, eNone, eNone, 5000, True, ""

			'Queue up dialog and sequences
			If SKIP_DIALOG = False Then
				queue.Add "blacksmith_lore_8", "DMD.blacksmithLore 8", 1, 0, 5000, 4165, 0, False
				queue.Add "blacksmith_lore_9", "DMD.blacksmithLore 9", 1, 0, 0, 5712, 0, False
				queue.Add "blacksmith_lore_10", "DMD.blacksmithLore 10", 1, 0, 0, 4392, 0, False
				queue.Add "blacksmith_lore_11", "DMD.blacksmithLore 11", 1, 0, 0, 4773, 0, False
				queue.Add "blacksmith_lore_start_kill", "DMD.wizardIntro_blacksmithKill", 1, 0, 0, 100, 0, False
			Else
				queue.Add "blacksmith_lore_start_kill", "DMD.wizardIntro_blacksmithKill", 1, 0, 5000, 100, 0, False
			End If

		Case MODE_BLACKSMITH_WIZARD_KILL
			MODE_COUNTERS_A = 0	'0 = Need right ramp shot; 1 = need horseshoe shot
			MODE_VALUES = 0		'Total points awarded

			'DMD
			DMD.triggerVideo VIDEO_BLACKSMITH_WIZARD_KILL
			DMD.gameStatus "BLACKSMITH HP: 100 / BATTLE TOTAL: 0"

			'GI
			LightsOn GI, Array(dpurple, purple), 100

			'Reset BLACKSMITH
			currentPlayerSet "spell_BLACK", 0
			currentPlayerSet "spell_SMITH", 0

			'Launch 2 balls to start with 20-second ball shield
			targetBIP 2, True
			addBallShield 20000

		Case MODE_BLACKSMITH_WIZARD_KILL_FAIL
			'Music / sounds
			Music.PlayTrack "music_blacksmith_kill_fail", Null, Null, Null, Null, 0, 0, 0, 0

			'GI
			LightsOn GI, Array(dred, red), 100
			LightsOn coloredLights, Array(dred, red), 100

			'Set progress to fail
			currentPlayerSet "spare_BLACKSMITH", 2

			'Reset AC to 5
			currentPlayerSet "AC", 5

			'DMD
			queue.Add "blacksmith_wizard_kill_fail_1", "DMD.triggerVideo VIDEO_BLACKSMITH_WIZARD_KILL_FAIL : DMD.blacksmithLore 12", 1, 0, 5000, 17000, 0, False
			queue.Add "blacksmith_wizard_kill_fail_2", "DMD.triggerVideo VIDEO_BLACKSMITH_WIZARD_KILL_FAIL_RUN : LightsOff GI : LightsOff aLights : LampC.LightOnWithColor ObjectiveBlacksmithL, Array(dpurple, purple)", 1, 0, 0, 5000, 0, False
			queue.Add "blacksmith_wizard_kill_fail_3", "DMD.triggerVideo VIDEO_BLACKSMITH_WIZARD_KILL_FAIL_TOTAL : LightsOff GI : LightsOff aLights", 1, 0, 0, 5000, 0, False
			queue.Add "blacksmith_wizard_kill_fail_4", "startMode MODE_BLACKSMITH_WIZARD_END", 1, 0, 0, 1000, 0, False

		Case MODE_BLACKSMITH_WIZARD_KILL_SUCCESS
			'Music / sounds
			Music.PlayTrack "music_blacksmith_kill_success", Null, Null, Null, Null, 0, 0, 0, 0

			'GI
			LightsOn GI, Array(dred, red), 100
			LightsOff coloredLights

			'Set progress to killed
			currentPlayerSet "spare_BLACKSMITH", 0

			'Reset ball shield and AC
			setBallShield 0
			if currentPlayer("AC") < 5 Then currentPlayerSet "AC", 5

			'DMD
			queue.Add "blacksmith_wizard_kill_success_1", "DMD.triggerVideo VIDEO_BLACKSMITH_WIZARD_KILL_SUCCESS", 300, 0, 0, 13270, 0, False
			queue.Add "blacksmith_wizard_kill_success_2", "Music.PlayTrack ""music_blacksmith_death"", Null, Null, Null, Null, 0, 0, 0, 0", 300, 0, 0, 2000, 0, False
			queue.Add "blacksmith_wizard_kill_success_3", "LightsOff coloredLights : LightsOff GI : DMD.blacksmithLore 13", 300, 0, 0, 25000, 0, False
			queue.Add "blacksmith_wizard_kill_success_4", "DMD.triggerVideo VIDEO_BLACKSMITH_WIZARD_KILL_SUCCESS_3 : LampC.LightOnWithColor ObjectiveBlacksmithL, Array(dred, red)", 300, 0, 0, 3000, 0, False
			queue.Add "blacksmith_wizard_kill_success_5", "DMD.triggerVideo VIDEO_BLACKSMITH_WIZARD_KILL_SUCCESS_TOTAL : LightsOff GI : LightsOff aLights", 1, 0, 0, 5000, 0, False
			queue.Add "blacksmith_wizard_kill_success_6", "startMode MODE_BLACKSMITH_WIZARD_END", 1, 0, 0, 1000, 0, False

			'Flasher show
			For i = 0 to 13
				flasherQueue.Add "blacksmith_wizard_kill_success_a_" & i, "Flash1 true", 300, i * (600), 0, 100, 0, False
				flasherQueue.Add "blacksmith_wizard_kill_success_b_" & i, "Flash1 false", 300, (i * (600)) + 75, 0, 100, 0, False
				flasherQueue.Add "blacksmith_wizard_kill_success_c_" & i, "Flash1 true", 300, (i * (600)) + 150, 0, 100, 0, False
				flasherQueue.Add "blacksmith_wizard_kill_success_d_" & i, "Flash1 false", 300, (i * (600)) + 225, 0, 100, 0, False
			Next

			'Light show
			queueB.Add "blacksmith_wizard_kill_success_seq_1", "StartLightSequence StandardSeq, Array(dred, red), 5, SeqRightOn, 1, 1", 300, (600 * 0), 0, 100, 0, False
			queueB.Add "blacksmith_wizard_kill_success_seq_2", "StartLightSequence StandardSeq, Array(dred, red), 5, SeqRightOff,0,1", 300, (600 * 1), 0, 100, 0, False
			queueB.Add "blacksmith_wizard_kill_success_seq_3", "StartLightSequence StandardSeq, Array(dred, red), 5, SeqLeftOn,1,1", 300, (600 * 2), 0, 100, 0, False
			queueB.Add "blacksmith_wizard_kill_success_seq_4", "StartLightSequence StandardSeq, Array(dred, red), 5, SeqLeftOff,0,1", 300, (600 * 3), 0, 100, 0, False
			queueB.Add "blacksmith_wizard_kill_success_seq_5", "StartLightSequence StandardSeq, Array(dred, red), 5, SeqUpOn,1,1", 300, (600 * 4), 0, 100, 0, False
			queueB.Add "blacksmith_wizard_kill_success_seq_6", "StartLightSequence StandardSeq, Array(dred, red), 5, SeqUpOff,0,1", 300, (600 * 5), 0, 100, 0, False
			queueB.Add "blacksmith_wizard_kill_success_seq_7", "StartLightSequence StandardSeq, Array(dred, red), 5, SeqDownOn,1,1", 300, (600 * 6), 0, 100, 0, False
			queueB.Add "blacksmith_wizard_kill_success_seq_8", "StartLightSequence StandardSeq, Array(dred, red), 5, SeqDownOff,0,1", 300, (600 * 7), 0, 100, 0, False
			queueB.Add "blacksmith_wizard_kill_success_seq_9", "StartLightSequence StandardSeq, Array(dred, red), 2, SeqDiagUpRightOn,1,1", 300, (600 * 8), 0, 100, 0, False
			queueB.Add "blacksmith_wizard_kill_success_seq_10", "StartLightSequence StandardSeq, Array(dred, red), 2, SeqDiagUpRightOff,0,1", 300, (600 * 9), 0, 100, 0, False
			queueB.Add "blacksmith_wizard_kill_success_seq_11", "StartLightSequence StandardSeq, Array(dred, red), 2, SeqDiagDownLeftOn,1,1", 300, (600 * 10), 0, 100, 0, False
			queueB.Add "blacksmith_wizard_kill_success_seq_12", "StartLightSequence StandardSeq, Array(dred, red), 2, SeqDiagDownLeftOff,0,1", 300, (600 * 11), 0, 100, 0, False
			queueB.Add "blacksmith_wizard_kill_success_seq_13", "StartLightSequence StandardSeq, Array(dred, red), 2, SeqDiagUpLeftOn,1,1", 300, (600 * 12), 0, 100, 0, False
			queueB.Add "blacksmith_wizard_kill_success_seq_14", "StartLightSequence StandardSeq, Array(dred, red), 2, SeqDiagUpLeftOff,0,1", 300, (600 * 13), 0, 100, 0, False
			queueB.Add "blacksmith_wizard_kill_success_seq_15", "StopLightSequence StandardSeq : LightsOn coloredLights, Array(dred, red), 100", 300, (600 * 14), 0, 100, 0, False
			queueB.Add "blacksmith_wizard_kill_success_seq_16", "LightsOff coloredLights", 300, (600 * 14.33), 0, 100, 0, False
			queueB.Add "blacksmith_wizard_kill_success_seq_17", "LightsOn coloredLights, Array(dred, red), 100", 300, (600 * 14.66), 0, 100, 0, False
			queueB.Add "blacksmith_wizard_kill_success_seq_18", "LightsOff coloredLights", 300, (600 * 15), 0, 100, 0, False
			queueB.Add "blacksmith_wizard_kill_success_seq_19", "LightsOn coloredLights, Array(dred, red), 100", 300, (600 * 15.33), 0, 100, 0, False
			queueB.Add "blacksmith_wizard_kill_success_seq_20", "LightsOff coloredLights", 300, (600 * 15.66), 0, 100, 0, False
			queueB.Add "blacksmith_wizard_kill_success_seq_22", "LightsOn coloredLights, Array(dred, red), 100", 300, (600 * 16), 0, 100, 0, False

		Case MODE_BLACKSMITH_WIZARD_END
			LightsOff aLights
			DMD.triggerVideo VIDEO_DRAINING_BALLS

		Case MODE_BLACKSMITH_WIZARD_INTRO_SPARE
			'Music
			Music.PlayTrack "music_blacksmith_spare_lore", Null, Null, 5000, 5000, 0, 0, 0, 0

			'GI / Lights
			LampC.RemoveShotsFromLight ObjectiveBlacksmithL
			LightsOn GI, Array(dpurple, purple), 25
			LightsOn coloredLights, Array(dwhite, white), 25

			'DMD
			DMD.triggerOverlay OVERLAY_REMOVE
			DMD.triggerVideo VIDEO_BLACKSMITH_LORE_SPARE
			DMD.jpLocked = False
			DMD.jpFlush
			DMD.jpDMD DMD.jpCL("BLACKSMITH"), "", "d_E219_0", eBlink, eNone, eNone, 5000, True, ""

			'Queue up dialog and sequences
			If SKIP_DIALOG = False Then
				queue.Add "blacksmith_lore_14", "DMD.blacksmithLore 14", 1, 0, 5000, 11500, 0, False
				queue.Add "blacksmith_lore_start_spare", "DMD.wizardIntro_blacksmithSpare", 1, 0, 0, 100, 0, False
			Else
				queue.Add "blacksmith_lore_start_spare", "DMD.wizardIntro_blacksmithSpare", 1, 0, 5000, 100, 0, False
			End If

		Case MODE_BLACKSMITH_WIZARD_SPARE
			MODE_COUNTERS_A = 0	'Current Jackpot
			MODE_VALUES = 0		'Total points awarded

			PlayCallout "callout_blacksmith_wizard_start", 0, VOLUME_CALLOUTS

			'DMD
			DMD.triggerVideo VIDEO_BLACKSMITH_WIZARD_SPARE
			DMD.gameStatus "RAID THE SHOP TOTAL: 0"

			'GI
			LightsOn GI, Array(dpurple, purple), 100

			'Reset BLACKSMITH
			currentPlayerSet "spell_BLACK", 0
			currentPlayerSet "spell_SMITH", 0

			'Launch 5 balls and award 60 seconds ball shield
			targetBIP 5, True
			addBallShield 60000

		Case MODE_HIGH_SCORE_ENTRY
			MODE_COUNTERS_A = " " 'Cursor character selection

			If Not currentPlayer("name") = "PLAYER " & dataGame.data.Item("playerUp") Then 'Name already defined, probably from Scorbit; just save the score
				startMode MODE_HIGH_SCORE_FINAL
				Exit Sub
			Else
				'DMD
				DMD.jpFlush
				DMD.jpLocked = True
				DMD.jpDMD DMD.jpCL(currentPlayer("name")), DMD.jpCL("High score!"), "", eNone, eBlink, eNone, 3000, False, ""
				DMD.jpDMD DMD.jpCL(currentPlayer("name")), DMD.jpCL("Enter your name"), "", eNone, eBlinkFast, eNone, 3000, True, ""
				currentPlayerSet "name", ""
				DMD.highScoreSequence_InsertName False
			End If

		Case MODE_HIGH_SCORE_FINAL
			DMD.highScoreSequence_final
			
		' TODO: Do not forget ScorbitBuildGameModes when adding new modes here!
	End Select

	updateShotLights True 'Must be updated after running above cases as this may depend on mode variables

	'Oops, there should always be at least 1 ball in play when game play is active. Auto-fire a ball into play if necessary.
	If newMode >= 1000 And BIP = 0 And BIPL = 0 And BQueue = 0 Then
		If DEBUG_LOG_MODES = True Then WriteToLog "modes", "No balls in play! Launching a ball."
		If newMode = MODE_NORMAL Or newMode = MODE_SKILLSHOT Then
			targetBIP 1, False 'No auto-plunge if activating normal mode (could be exiting a wizard and the player may want a rest) or skillshot

			'A ball shield should also be set in this instance
			Clocks.data.Item("shield").timeLeft = 0
			addBallShield currentPlayer("AC") * 1000
			Clocks.data.Item("shield").isPaused = True
		Else
			targetBIP 1, True
		End If
	End If

	'Prevent modes from expiring right away in case a ball gets stuck. Instead, when the time runs out, they will expire via doSwitch on the next switch activation.
	Clocks.data.Item("mode").canExpire = False

	If DEBUG_LOG_MODES = True Then WriteToLog "modes", "Started " & newMode
End Sub

Sub expireMode()
	PlaySound "ui_time_up", 1, VOLUME_SFX

	Clocks.data.Item("mode").timeLeft = 0

	Select Case CURRENT_MODE
		Case MODE_GLITCH_WIZARD
			startMode MODE_GLITCH_WIZARD_END
		Case MODE_BLACKSMITH_WIZARD_SPARE
			'DMD
			queue.Add "blacksmith_wizard_spare_total", "DMD.triggerVideo VIDEO_BLACKSMITH_WIZARD_SPARE_TOTAL", 2, 0, 0, 5000, 0, False

			'Reset BLACKSMITH
			currentPlayerSet "spell_BLACK", 0
			currentPlayerSet "spell_SMITH", 0

			startMode MODE_NORMAL
	End Select
End Sub

'-----------------------------------------
'    WGEM: GEMS AND POWERUPS
'-----------------------------------------

Sub updatePowerupLights() 'Check powerup progress and adjust lights accordingly
	' Also check action button light
	updateActionButton

	If CURRENT_MODE < 1000 Then
		LightsOff powerupLights
		Exit Sub
	End If

	'Ensure they are all the correct color
	LightsColor powerupLights, Array(dblue, blue)
	
	'Determine the order of powerup lights
	Dim lightOrder
	lightOrder = WShuffleArray(Array(0, 1, 2, 3, 4, 5), wrnd("gems", False))
	
	'Set light states accordingly
	Dim i
	For i = 0 To 5
		If currentPlayer("gems") > i Then
			LampC.LightLevel powerupLights(lightOrder(i)), 50
			LampC.LightOn powerupLights(lightOrder(i))
		ElseIf currentPlayer("gems") = i Then
			powerupLights(lightOrder(i)).BlinkPattern = "100"
			powerupLights(lightOrder(i)).BlinkInterval = 222
			LampC.LightLevel powerupLights(lightOrder(i)), 100
			LampC.Blink powerupLights(lightOrder(i))
		Else
			LampC.LightOff powerupLights(lightOrder(i))
		End If
	Next
End Sub

Sub checkGemShot(lightIndex) 'Check if a gem shot registers a collected gem, and collect it if so.
	If currentPlayer("gems") > 5 Then Exit Sub 	'Do we have a pending powerup to use? Nothing to do if so.
	If CURRENT_MODE < 1000 Then Exit Sub 'No collecting gems outside of active gameplay
	
	'Determine the order of powerup lights
	Dim lightOrder
	lightOrder = WShuffleArray(Array(0, 1, 2, 3, 4, 5), wrnd("gems", False))
	
	'Did we collect the gem?
	If lightOrder(currentPlayer("gems")) = lightIndex Then
		currentPlayerSet "gems", currentPlayer("gems") + 1
		updatePowerupLights
		LampC.Pulse powerupLights(lightIndex), 3
		
		If currentPlayer("gems") > 5 Then 'Powerup collected
			lampC.AddLightSeq "accomplishment", lSeqpowerupCollected
			PlaySound "sfx_powerup_collected", 1, VOLUME_SFX
			queue.Add "powerup_collected","DMD.triggerVideo VIDEO_POWERUP_COLLECTED : PlayCallout ""callout_nar_powerup_collected"", 1, VOLUME_CALLOUTS",5,0,0,5000,10000,False

			' Log it in Scorbit
			If ScorbitActive = 1 Then Scorbit.SetGameMode "CP:Powerup Collected"
		End If
	End If
End Sub

Sub deployPowerup() 'Use a powerup
	If currentPlayer("gems") < 6 Or CURRENT_MODE < 1000 Or CURRENT_MODE = MODE_SKILLSHOT Then Exit Sub

	'Reset gem shot count
	currentPlayerSet "gems", 0

	'Go to the next seed number for gem shots and update lights
	Dim lightOrder
	lightOrder = WShuffleArray(Array(0, 1, 2, 3, 4, 5), wrnd("gems", True))
	updatePowerupLights

	'Play sound effect
	PlaySound "sfx_powerup_deployed", 1, VOLUME_SFX

	'Play light sequence
	lampC.AddLightSeq "accomplishment", lSeqpowerupDeployed

	'Deploy powerup according to mode
	Select Case CURRENT_MODE
		Case MODE_NORMAL 'Advance letters
			'Display / sounds
			queue.Add "powerup_normal_1", "DMD.triggerVideo VIDEO_POWERUP_NORMAL : PlayCallout ""callout_nar_powerup_letters"", 1, VOLUME_CALLOUTS", 301, 0, 0, 2000, 0, True

			'Bonus X
			spellBonusX

			'Blacksmith (one on each side)
			spellBlacksmith False
			spellBlacksmith True

			'TODO: Finish

		Case MODE_GLITCH_WIZARD '3X glitch jackpot
			'3X Jackpot
			Dim jackpotAwarded
			jackpotAwarded = (currentPlayerBonus("all", False) * 3 * MODE_COUNTERS_B)
			addScore jackpotAwarded, "3X Glitch Jackpot"
			MODE_VALUES = MODE_VALUES + jackpotAwarded
			DMD.gameStatus "GLITCH TOTAL: " & FormatScore(MODE_VALUES)

			'Display / sounds
			queue.Add "powerup_glitch_1", "DMD.triggerVideo VIDEO_POWERUP_GLITCH", 301, 0, 0, 2000, 0, True
			queue.Add "powerup_glitch_2", "PlayCallout ""callout_nar_superjackpot"", 1, VOLUME_CALLOUTS : PlaySound ""sfx_glitch_6"", 1, VOLUME_SFX : DMD.glitchWizardJackpot " & jackpotAwarded, 300, 2000, 0, 5000, 0, True

		Case MODE_BLACKSMITH_WIZARD_KILL 'Spell up to 2 blacksmith letters, one on each side
			'Display / sounds
			queue.Add "powerup_blacksmith_wizard_kill_1", "DMD.triggerVideo VIDEO_POWERUP_BLACKSMITH_KILL", 301, 0, 0, 2000, 0, True

			'Spell BLACK and SMITH
			If currentPlayer("spell_BLACK") < 5 Then 
				currentPlayerSet "spell_BLACK", currentPlayer("spell_BLACK") + 1
				checkBlacksmith
			End If
			If currentPlayer("spell_SMITH") < 5 Then 
				currentPlayerSet "spell_SMITH", currentPlayer("spell_SMITH") + 1
				checkBlacksmith
			End If

		Case MODE_BLACKSMITH_WIZARD_SPARE 'Spell up to 2 blacksmith letters, one on each side
			'Display / sounds
			queue.Add "powerup_blacksmith_wizard_spare", "DMD.triggerVideo VIDEO_POWERUP_BLACKSMITH_SPARE", 301, 0, 0, 2000, 0, True

			'Spell BLACK and SMITH
			If currentPlayer("spell_BLACK") < 5 Then 
				currentPlayerSet "spell_BLACK", currentPlayer("spell_BLACK") + 1
				checkBlacksmith
			End If
			If currentPlayer("spell_SMITH") < 5 Then 
				currentPlayerSet "spell_SMITH", currentPlayer("spell_SMITH") + 1
				checkBlacksmith
			End If
			
	End Select
End Sub

'-----------------------------------------
'    WBUM: BUMPERS AND DIVERTERS
'-----------------------------------------

Sub leftBumperDiverter_expired()
	SolLBumpDiverter False
	LampC.LightOff LeftOrbitBumpersL
End Sub

Sub rightBumperDiverter_expired()
	SolRBumpDiverter False
	LampC.LightOff RightOrbitBumpersL
End Sub

'-----------------------------------------
'    WHPD: +HP DROP TARGETS
'-----------------------------------------

Sub resetHPTargets()
	'Raise the drop targets
	SolHPBank True

	'Reset the lights
	LightsOff hpLights
End Sub

Sub updateHPLights()
	If CURRENT_MODE < 1000 Then
		LightsOff hpLights
		Exit Sub
	End If

	If DTDropped(SWITCH_DROPS_PLUS) = True And DTDropped(SWITCH_DROPS_H) = True And DTDropped(SWITCH_DROPS_P) = True Then
		LightsBlink hpLights, Array(dred, red), 500, 100
	Else
		LightsOff hpLights
		If DTDropped(SWITCH_DROPS_PLUS) = True Then LampC.LightOn DropTargetsPlusL
		If DTDropped(SWITCH_DROPS_H) = True Then LampC.LightOn DropTargetsHL
		If DTDropped(SWITCH_DROPS_P) = True Then LampC.LightOn DropTargetsPL
	End If
End Sub

'-----------------------------------------
'    WBON: Bumper Inlanes / Bonus X
'-----------------------------------------

Sub updateBumperInlanes() 'Check for bonus X advancement and update light states of bumper inlanes
	Dim i

	'All inlanes lit; we should handle increase in the bonusX stage
	If bumperInlanes(0) = True And bumperInlanes(1) = True And bumperInlanes(2) = True Then
		'Turn off lanterns
		expireBumperInlanes
		For i = 0 to 2
			LampC.Pulse lanternLights(i), 3
		Next

		'Advance bonus X
		spellBonusX
	End If

	'Handle lantern states
	For i = 0 to 2
		If bumperInlanes(i) = True Then
			If Clocks.data.Item("bumperInlanes").timeLeft > 0 Then
				lanternLights(i).BlinkPattern = "10" 'interval is determined by time left on a timer
				LampC.Blink lanternLights(i)
			Else
				LampC.LightOn lanternLights(i)
			End If
		ElseIf CURRENT_MODE = MODE_GLITCH_WIZARD Then
			lanternLights(i).BlinkPattern = "0001"	'We want to indicate the player should shoot the lantern by flashing, but the lanterns should be off most of the time so the player knows which ones still need lit
			LampC.UpdateBlinkInterval lanternLights(i), 125
			LampC.Blink lanternLights(i)
		Else
			LampC.LightOff lanternLights(i)
		End If
	Next
End Sub

Sub expireBumperInlanes() 'Turn off all bumper inlanes and expire the timer
	Clocks.data.Item("bumperInlanes").timeLeft = 0

	Dim i
	For i = 0 to 2
		bumperInlanes(i) = False
	Next

	updateBumperInlanes
End Sub

Sub spellBonusX() 'Add a letter to Bonus X, or increase the Bonus X if we have all the letters
	Dim i

	If CURRENT_MODE = MODE_GLITCH_WIZARD Then 'Award the glitch jackpot
		'Sequence / Light show / Glitch
		StandardSeq.CenterX = 284
		StandardSeq.CenterY = 154
		StartLightSequence StandardSeq, Array(ddarkgreen, darkgreen), 10, SeqScrewRightOn, 90, 1
		For i = 0 to 9
			flasherQueue.Add "glitch_wizard_jackpot_on_" & i, "Flash5 True", 100, 0, 0, 200, 0, False
			flasherQueue.Add "glitch_wizard_jackpot_off_" & i, "Flash5 False", 100, 0, 0, 200, 0, False
		Next

		'Jackpot
		Dim jackpotAwarded
		jackpotAwarded = (currentPlayerBonus("all", False) * MODE_COUNTERS_B)
		addScore jackpotAwarded, "Glitch Jackpot"
		MODE_VALUES = MODE_VALUES + jackpotAwarded
		DMD.gameStatus "GLITCH TOTAL: " & FormatScore(MODE_VALUES)

		'Sounds / callout
		PlayCallout "callout_nar_jackpot", 1, VOLUME_CALLOUTS
		PlaySound "sfx_glitch_" & int(rnd*5)+1, 1, VOLUME_SFX

		'Display
		queue.Add "glitch_wizard_jackpot", "DMD.glitchWizardJackpot " & jackpotAwarded, 100, 0, 0, 5000, 5000, True
	ElseIf currentPlayer("bonusX") < 9 Or currentPlayer("spell_BONUSX") < 6 Then 'Bonus X has not yet been maxed out entirely
		'Increase number of letters spelled
		currentPlayerSet "spell_BONUSX", currentPlayer("spell_BONUSX") + 1

		'Award points
		If currentPlayer("spell_BONUSX") = 6 Then
			addScore Scoring.basic("bonusX"), "Bonus X"
		Else
			addScore Scoring.basic("spell_BONUSX"), "Spelled Bonus X letter"
		End If

		If CurrentPlayer("bonusX") < 9 Then 'Bonus X not yet reached maximum
			' Do something depending on the stage
			Select Case currentPlayer("spell_BONUSX")
				Case 1:
					queue.Add "bonusx_1", "DMD.triggerVideo VIDEO_BONUSX_1", 12, 0, 0, 3000, 5000, False
					PlaySound "sfx_bonusx_letter", 1, VOLUME_SFX
				Case 2:
					queue.Add "bonusx_2", "DMD.triggerVideo VIDEO_BONUSX_2", 12, 0, 0, 3000, 5000, False
					PlaySound "sfx_bonusx_letter", 1, VOLUME_SFX
				Case 3:
					queue.Add "bonusx_3", "DMD.triggerVideo VIDEO_BONUSX_3", 12, 0, 0, 3000, 5000, False
					PlaySound "sfx_bonusx_letter", 1, VOLUME_SFX
				Case 4:
					queue.Add "bonusx_4", "DMD.triggerVideo VIDEO_BONUSX_4", 12, 0, 0, 3000, 5000, False
					PlaySound "sfx_bonusx_letter", 1, VOLUME_SFX
				Case 5:
					queue.Add "bonusx_5", "DMD.triggerVideo VIDEO_BONUSX_5", 12, 0, 0, 3000, 5000, False
					PlaySound "sfx_bonusx_letter", 1, VOLUME_SFX
				Case 6: 'All letters collected; increase bonus multiplier
					currentPlayerSet "bonusX", currentPlayer("bonusX") + 1
					currentPlayerSet "spell_BONUSX", 0
					queue.Add "bonusx_6", "DMD.triggerVideo VIDEO_BONUSX_6 : PlayCallout ""callout_nar_bonus_multiplier"", 1, VOLUME_CALLOUTS", 20, 500, 0, 3000, 10000, False
					PlaySound "sfx_bonusx_completed", 1, VOLUME_SFX
					If ScorbitActive = 1 Then Scorbit.SetGameMode "CP:Bonus X"
			End Select
		ElseIf currentPlayer("bonusX") = 9 And currentPlayer("spell_BONUSX") < 7 Then 'Bonus X maxed but glitch mini-wizard not yet ready; progress towards glitch mini-wizard.
			Select Case currentPlayer("spell_BONUSX")
				Case 1:
					queue.Add "bonusx_1", "DMD.triggerVideo VIDEO_BONUSX_1 : DMD.glitchProgression 1", 12, 0, 0, 100, 5000, False
					PlaySound "sfx_bonusx_letter", 1, VOLUME_SFX
				Case 2:
					queue.Add "bonusx_2", "DMD.triggerVideo VIDEO_BONUSX_2 : DMD.glitchProgression 2", 12, 0, 0, 100, 5000, False
					PlaySound "sfx_bonusx_letter", 1, VOLUME_SFX
				Case 3:
					queue.Add "bonusx_3", "DMD.triggerVideo VIDEO_BONUSX_3 : DMD.glitchProgression 3", 12, 0, 0, 100, 5000, False
					PlaySound "sfx_bonusx_letter", 1, VOLUME_SFX
				Case 4:
					queue.Add "bonusx_4", "DMD.triggerVideo VIDEO_BONUSX_4 : DMD.glitchProgression 4", 12, 0, 0, 100, 5000, False
					PlaySound "sfx_bonusx_letter", 1, VOLUME_SFX
				Case 5:
					queue.Add "bonusx_5", "DMD.triggerVideo VIDEO_BONUSX_5 : DMD.glitchProgression 5", 12, 0, 0, 100, 5000, False
					PlaySound "sfx_bonusx_letter", 1, VOLUME_SFX
				Case 6: 'Glitch mini-wizard ready
					queue.Add "bonusx_6", "DMD.triggerVideo VIDEO_BONUSX_6 : DMD.glitchProgression 6", 20, 500, 0, 100, 0, False
					PlaySound "sfx_bonusx_completed", 1, VOLUME_SFX
					If ScorbitActive = 1 Then Scorbit.SetGameMode "CP:GLITCH wizard ready"
					checkHoleState
			End Select
		Else 'We maxed out our Bonus X so do nothing
			currentPlayerSet "spell_BONUSX", 7
		End If
	End If
End Sub

'-----------------------------------------
'    WBSM: Blacksmith
'-----------------------------------------

Sub updateBlacksmithLights() 'Update the horseshoe blacksmith light progress
	LightsOff blacksmithLights
	If CURRENT_MODE < 1000 Then Exit Sub

	Dim i
	Select Case CURRENT_MODE
		Case MODE_SKILLSHOT:
			'Do nothing; keep off

		Case MODE_BLACKSMITH_WIZARD_KILL
			For i = 0 To 4
				If currentPlayer("spell_BLACK") > i Then
					LampC.LightOn blacksmithLights(i)
				ElseIf currentPlayer("spell_BLACK") = i Then
					blacksmithLights(i).BlinkPattern = "10"
					blacksmithLights(i).BlinkInterval = 333
					If MODE_COUNTERS_A = 1 Then LampC.Blink blacksmithLights(i)
				Else
					LampC.LightOff blacksmithLights(i)
				End If
			Next
			For i = 0 To 4
				If currentPlayer("spell_SMITH") > i Then
					LampC.LightOn blacksmithLights(i + 5)
				ElseIf currentPlayer("spell_SMITH") = i Then
					blacksmithLights(i + 5).BlinkPattern = "10"
					blacksmithLights(i + 5).BlinkInterval = 333
					If MODE_COUNTERS_A = 1 Then LampC.Blink blacksmithLights(i + 5)
				Else
					LampC.LightOff blacksmithLights(i + 5)
				End If
			Next

		Case Else 'All other modes, show standard BLACKSMITH progress
			For i = 0 To 4
				If currentPlayer("spell_BLACK") > i Then
					LampC.LightOn blacksmithLights(i)
				ElseIf currentPlayer("spell_BLACK") = i Then
					blacksmithLights(i).BlinkPattern = "10"
					blacksmithLights(i).BlinkInterval = 333
					LampC.Blink blacksmithLights(i)
				Else
					LampC.LightOff blacksmithLights(i)
				End If
			Next
			For i = 0 To 4
				If currentPlayer("spell_SMITH") > i Then
					LampC.LightOn blacksmithLights(i + 5)
				ElseIf currentPlayer("spell_SMITH") = i Then
					blacksmithLights(i + 5).BlinkPattern = "10"
					blacksmithLights(i + 5).BlinkInterval = 333
					LampC.Blink blacksmithLights(i + 5)
				Else
					LampC.LightOff blacksmithLights(i + 5)
				End If
			Next
		End Select
End Sub

Sub spellBlacksmith(spellingSmith) 'Spell a letter of BLACKSMITH while running appropriate mode checks
	If CURRENT_MODE < 1000 Then Exit Sub 'No spelling in non-gameplay modes
	If CURRENT_MODE = MODE_SKILLSHOT Then Exit Sub 'No spelling in skillshot
	If CURRENT_MODE = MODE_BLACKSMITH_WIZARD_KILL Then 'Must be a parent if to MODE_COUNTERS_A check to prevent type errors.
		If MODE_COUNTERS_A = 0 Then Exit Sub 'Need to shoot ramp first
	End If

	If spellingSmith = False Then
		If currentPlayer("spell_BLACK") < 5 Then
			currentPlayerSet "spell_BLACK", currentPlayer("spell_BLACK") + 1
			checkBlackSmith
		End If
	Else
		If currentPlayer("spell_SMITH") < 5 Then
			currentPlayerSet "spell_SMITH", currentPlayer("spell_SMITH") + 1
			checkBlackSmith
		End If
	End If
End Sub

Sub checkBlacksmith() 'Check if BLACKSMITH is spelled and do stuff depending on mode; run this when and ONLY when a letter is spelled, and run for every letter spelled
	Dim i
	Dim blacksmithSpelled

	'Heal poisoned outlanes when we spell black or smith
	If currentPlayer("spell_BLACK") >= 5 Then
		currentPlayerSet "poisonL", 0
		updatePoisonLights
	End If
	If currentPlayer("spell_SMITH") >= 5 Then
		currentPlayerSet "poisonR", 0
		updatePoisonLights
	End If
	
	Select Case CURRENT_MODE
		Case MODE_SKILLSHOT:
			'Do nothing

		Case MODE_BLACKSMITH_WIZARD_KILL
			blacksmithSpelled = currentPlayer("spell_BLACK") + currentPlayer("spell_SMITH")

			addScore Scoring.blacksmith("wizard_kill_spell"), "Blacksmith boss battle: dealt damage"
			MODE_VALUES = MODE_VALUES + Scoring.blacksmith("wizard_kill_spell")

			'SFX
			StopSound "sfx_punch_shield"
			PlaySound "sfx_punch_shield", 1, VOLUME_SFX

			If blacksmithSpelled >= 10 Then 'Blacksmith killed
				addScore Scoring.blacksmith("wizard_kill_completed"), "Blacksmith boss battle: killed Blacksmith"
				startMode MODE_BLACKSMITH_WIZARD_KILL_SUCCESS
			Else
				'DMD Sequence
				DMD.gameStatus "BLACKSMITH HP: " & (100 - (10 * blacksmithSpelled)) & " / BATTLE TOTAL: " & FormatScore(MODE_VALUES)
				queue.Add "blacksmith_wizard_kill_damage", "DMD.blacksmithWizardKillDamage " & Scoring.blacksmith("wizard_kill_spell"), 100, 250, 0, 5000, 10000, True

				'Add 20 seconds of shield
				addBallShield 20000

				'Launch more balls when necessary to make it harder as the blacksmith weakens
				If blacksmithSpelled >= 3 Then targetBIP 3, True
				If blacksmithSpelled >= 6 Then targetBIP 4, True
				If blacksmithSpelled >= 8 Then targetBIP 5, True

				'Ramp needs to be shot again
				MODE_COUNTERS_A = 0
				updateShotLights False
			End If

		Case MODE_BLACKSMITH_WIZARD_SPARE
			blacksmithSpelled = currentPlayer("spell_BLACK") + currentPlayer("spell_SMITH")

			'SFX
			StopSound "sfx_punch_shield"
			PlaySound "sfx_punch_shield", 1, VOLUME_SFX

			If blacksmithSpelled >= 10 Then 'Super Jackpot
				'Collect Super Jackpot
				addScore Scoring.blacksmith("wizard_spare")(9), "Raid the Blacksmith: Super Jackpot"
				MODE_VALUES = MODE_VALUES + Scoring.blacksmith("wizard_spare")(9)

				'Light Sequence
				PlaySound "sfx_wood_destruction", 0, VOLUME_SFX
				StandardSeq.CenterX = HorseshoeComplete.X
				StandardSeq.CenterY = HorseshoeComplete.Y
				StartLightSequence StandardSeq, Array(dpurple, purple), 35, SeqCircleOutOn, 35, 1
				queueB.Add "blacksmith_raid_jackpot_seq_stop", "StopLightSequence StandardSeq", 200, 5000, 0, 100, 0, True
				For i = 0 to 9
					flasherQueue.Add "blacksmith_raid_jackpot_" & i & "_a", "Flash6 True", 100, 0, 0, 150, 0, False
					flasherQueue.Add "blacksmith_raid_jackpot_" & i & "_b", "Flash6 False", 100, 0, 0, 150, 0, False
				Next

				'DMD sequence
				queue.Add "blacksmith_raid_super_jackpot", "DMD.blacksmithWizardSpareSuperJackpot " & Scoring.blacksmith("wizard_spare")(9), 200, 0, 0, 5000, 10000, True

				'Reset blacksmith spelling progress and jackpot
				currentPlayerSet "spell_BLACK", 0
				currentPlayerSet "spell_SMITH", 0
				MODE_COUNTERS_A = 0

				'Update status
				DMD.gameStatus "JACKPOT: 0 / RAID TOTAL: " & FormatScore(MODE_VALUES)
			Else
				'Increment Jackpot
				MODE_COUNTERS_A = Scoring.blacksmith("wizard_spare")(blacksmithSpelled - 1)
				DMD.gameStatus "JACKPOT: " & FormatScore(MODE_COUNTERS_A) & " / RAID TOTAL: " & FormatScore(MODE_VALUES)
			End If

			checkHoleState

		Case Else 'All other modes, standard BLACKSMITH procedure (increase AC / activate ball shield when BLACKSMITH is spelled)
			If currentPlayer("spell_BLACK") >= 5 And currentPlayer("spell_SMITH") >= 5 Then 'We completed Blacksmith
				'Pulse Blacksmith lights
				For Each i in blacksmithLights
					LampC.Pulse i, 3
				Next

				'Pulse blacksmith objective light
				LampC.Pulse ObjectiveBlacksmithL, 3

				'Purple Flasher
				For i = 0 to 7
					flasherQueue.Add "blacksmith_on_" & i, "Flash6 True", 20, 0, 0, 100, 0, False
					flasherQueue.Add "blacksmith_off_" & i, "Flash6 False", 20, 0, 0, 100, 0, False
				Next

				'Reset blacksmith spelling progress
				currentPlayerSet "spell_BLACK", 0
				currentPlayerSet "spell_SMITH", 0

				'Random Callout + Sound Effect
				queue.Add "spell_blacksmith", "DMD.triggerVideo VIDEO_BLACKSMITH_SPELL : PlayCallout ""callout_blacksmith_spell_" & int(rnd * 10) & """, 1, VOLUME_CALLOUTS", 20, 0, 0, 8000, 10000, False
				PlaySound "sfx_blacksmith", 1, VOLUME_SFX

				'Upgrade AC if not maxed out and mini-wizard not started
				If currentPlayer("AC") < 20 And currentPlayer("obj_BLACKSMITH") = 0 Then
					currentPlayerSet "AC", currentPlayer("AC") + 1
					updateObjectiveLights

					'AC maxed out; Blacksmith mini-wizard is ready
					If currentPlayer("AC") >= 20 And CURRENT_MODE = MODE_NORMAL And currentPlayer("obj_BLACKSMITH") = 0 Then
						queue.Add "blacksmith_wizard_ready", "DMD.wizardReady_blacksmith", 9, 0, 0, 6000, 0, False
						checkHoleState
					End If
				End If

				'Depending on mode, add mode time or ball shield time according to the new AC
				If currentModeHasInfiniteBallShield = True And Clocks.data.Item("mode").timeLeft > 0 Then
					Clocks.data.Item("mode").timeLeft = Clocks.data.Item("mode").timeLeft + (currentPlayer("AC") * 1000)
				Else
					Select Case CURRENT_MODE
						Case Else
							addBallShield CurrentPlayer("AC") * 1000
					End Select
				End If

				'Scorbit
				If ScorbitActive = 1 Then Scorbit.SetGameMode "CP:Spelled BLACKSMITH"
			End If
	End Select

	updateBlacksmithLights
End Sub

'-----------------------------------------
'    WPOI: Poison
'-----------------------------------------

Sub updatePoisonLights()
	If currentPlayer("poisonL") = 0 Or currentPlayer("AC") >= 20 Then
		LampC.LightOff LeftOutlanePoisonL
	Else
		If Clocks.data.Item("shield").timeLeft > 2000 Or currentModeHasInfiniteBallShield = True Then
			LampC.LightOnWithColor LeftOutlanePoisonL, Array(dyellow, yellow)
		Else
			LampC.LightOnWithColor LeftOutlanePoisonL, Array(dred, red)
		End If
	End If
	If currentPlayer("poisonR") = 0 Or currentPlayer("AC") >= 20 Then
		LampC.LightOff RightOutlanePoisonL
	Else
		If Clocks.data.Item("shield").timeLeft > 2000 Or currentModeHasInfiniteBallShield = True Then
			LampC.LightOnWithColor RightOutlanePoisonL, Array(dyellow, yellow)
		Else
			LampC.LightOnWithColor RightOutlanePoisonL, Array(dred, red)
		End If
	End If
End Sub

'*****************************************
'   WVER: CHANGE LOG
'*****************************************

' 0.0.1-alpha 	Arelyel, DonkeyKlonk		Initial version with draft playfield layout
' 0.0.2-alpha 	Arelyel						Added nFozzy physics
' 0.0.3-alpha 	Arelyel 					Added Fleep sounds
' 0.0.4-alpha 	Arelyel 					Added PinUP, VPin Advanced Queuing System, and VPin BookKeeper; physics tweaks
' 0.0.5-alpha 	Arelyel, Flux 				Added Freezy Lampz, Flux light state controller, and post drop target primitives; attract sequence work; table layout color tweaks
' 0.0.6-alpha 	Arelyel 					Added Iakki, Apophis, and Wylte Dynamic Ball Shadows; Added Flupper Domes; modified script to be a bit more in line with VPW standards (still a WIP); removed core.vbs (unused and conflicting)
' 0.0.7-alpha 	Arelyel 					Added Flupper Bumpers; Removed Arelyel branding and changed to VPin Workshop; Re-arranged attract sequence; high scores actually on attract sequence; start of vpwConstants; 
'											add reset property to vpwBookkeeper
' 0.1.0-alpha 	Arelyel 					Beginnings of gameplay functionality / framework; Insert coin / choose your difficulty; special randomization that is tournament friendly; Fleep fixes; vpwBookkeeper re-write; 
'											decreased flipper spacing slightly
' 0.1.1-alpha 	Arelyel 					Start of segment displays; vpwClocks class; moved inlane pegs closer to plastics; GI intensity increased; increased slingshot strength; decreased bumper gate elasticity
' 0.1.2-alpha 	Arelyel 					Stable Diffusion artwork updates; changed start game intro sequence; added logic for draining and switching players; disabled override physics; added scoop ejection warning
' 0.1.3-alpha 	Arelyel 					Migrated all PUP sounds and music to VPX; added music fading; fixed random number generator; added captive ball to active balls
' 0.1.4-alpha 	Arelyel, Flux, Remdwaas		Added several 3D objects / toys; light state controller update; ball search function; table of contents / comment modifications
' 0.1.5-alpha 	Arelyel 					Re-organizing / renaming of the PUP class in preparation to support fallbacks to PUP
' 0.1.6-alpha 	Arelyel 					Added core.vbs to utilize a few of the classes as helpers; misc difficulty changes; incorporated most of remaining basic scoring; finished migrating PUP audio to VPX; basic SFX for shots; switches
' 0.1.7-alpha 	Arelyel 					End-of-game bonus; FormatScore bug (unnecessary left padding)
' 0.1.8-alpha 	Arelyel 					Tournament Scores on attract sequence; gem shots and powerup integration started; timer pausing / unpausing depending on if a ball is actively in play; moved lamp controllers and music to a 30FPS timer
' 0.1.9-alpha 	Arelyel 					bumper diverters / timer
' 0.1.10-alpha 	Arelyel 					Scorbit integration; skillshot bug (monitored the wrong standup for completed skillshot)
' 0.1.11-alpha 	Arelyel 					+HP Drop Targets and leaf switch; playfield cutouts for targets; auto-fire delay; switches for key presses; scoop/saucer slowed down; subway adjustment; flipper physics tweaks
' 0.1.12-alpha 	Arelyel, Flux 				Prevent ball shield from expiring for up to 5 seconds when outlane is hit; light state controller update; ball shield pause fixes; +HP on impossible now increases ball shield, not HP
' 0.1.13-alpha 	Arelyel 					DEBUG_INFINITE_SHIELD; DEBUG_COMPILE_LIGHT_SHOWS; game over sequence; bumper inlanes / bonus X; reduce load on DMD.sUpdateScores; start of custom light sequences
' 0.1.14-alpha 	Arelyel, Flux 				sling spin; resume game functionality; light state controller 0.7
' 0.1.15-alpha 	Arelyel 					Scorbit Dev Kit 1.0.2 update; Add Force Asynch to Scorbit to remove stutter when not using StartSession; Mystified end game sequence; Debug logging
' 0.1.16-alpha 	Arelyel 					Co-op mode; death save
' 0.2.0-alpha 	Arelyel 					VPW Beta Tester release
' 0.2.1-alpha 	Arelyel 					JPSalas' reduced DMD fallback w/ FlexDMD support (FlexDMD broken right now); Use constants for video triggers; music pitch shift when entering death save (WIP / buggy); better PUP tutorial screens; 
'											add player bug when game not started; ball search bugs; skillshot now prompts action button ends game in mystified; vpwMusic and vpwTracks update; more intuitive death save
' 0.2.2-alpha	Arelyel						Blacksmith spelling @ horseshoe
' 0.2.3-alpha	Arelyel, Flux				Blacksmith mini-wizard (ready); Glitch mini-wizard (ready); warnings for failed PinUp / FlexDMD loads; Action button lamps below flippers; vpwQueueManager 1.1.2; JP DMD flushing behavior adjustments; 
'											playfield indicator adjustments; LightIntervalByTime specify upper bounds milliseconds; added AC and drain damage stats to PUP and JP DMD; removed corrupt kickerT1 image; 
'											Bonus X max changed from 5 to 9 (and updated bonus values)
' 0.2.4-alpha	Arelyel						Glitch Mini-Wizard; bug fix with mode / shield pausing during multiball; modified queueBalls / targetBIP routine to take parameter for number of balls that should be in play (opposed to number of balls to queue);
'											added secondary queue and moved some non-flasher stuff that was using flasher queue to it;
' 0.2.5-alpha	Arelyel, Flux				Light State Controller bug fixes with VPX sequences; use secondary queue for scoop eject when in multiball; increase scoop eject warning delay to 1000 milliseconds; 
'											bumper and slingshot solenoid calls moved from switch routine to individual element events; ball search should monitor modes 800-899; trigger PUP event E0 to clear videos each time a video event is called;
'											vpwQueue preQueue error when adding an item that already exists in the actual queue; vpwBookKeeper replaced by vpwData; modifications to game resuming functionality; moo; blacksmith attracts
' 0.2.6-alpha	Arelyel						update bumper inlane lanterns and +HP lights on every mode change; reworked the trough system; powerups; update startMode routine to better handle lamp states;
'											Fixed random number generation (was using scientific notation in VPReg); added hidden 2-second grace to ball shield; 
'											use triggers for preventing gate double-hits (not implemented yet as triggers are not working); moved JP DMD to playfield; separate playfield image for lights to improve quality
' 0.2.7-alpha	Arelyel						Separate PlaySound (PlayCallout) routine for callouts; glitch mini-wizard glitch improvements; don't add shield time to mode time in Impossible; 
'											+HP should not add shield time when in a mode with unlimited shield; use constants for overlays; PinUp: use D triggers instead of E for triggerVideo;
' 0.2.8-alpha	Arelyel						Blacksmith mini-wizard (kill); skip dialog option; single digit score start at 1 instead of 0 for game difficulty; ball shield and mode timer graces for stuck balls; light insert image adjustments
' 0.2.9-alpha	Arelyel						Lost PinUP Pack; Improved trough system; Rename Mystified to Zen; changed behavior of impossible difficulty; fixed routine waiting for balls to drain; 
'											Light State Controller update to v 0.8.4; change +HP lights to red; added difficulty instructions; renamed queueBalls to targetBIP; updated vpw* libraries; 
'											removed all FreePD music due to unstable license issues with FreePD; added new music made by Arelyel; glitch mini-wizard overhaul; combo shots; added center post and shortened flippers;
'                                           adjusted table slope according to game difficulty (vpxLight) setting instead of using a static value;
' 0.3.0-alpha	Arelyel						Blacksmith (spare) / Raid the Shop Mini-Wizard; outlane poisons; DirectB2S; high score entry; nudge penalty;