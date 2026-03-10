# VPX Standalone Script Patch Overview

This file lists all issues found across the patch files in this repository,
ordered by frequency (most frequent issue first).
Each section lists the scripts affected by that issue.

**Total patch files analyzed:** 323  
**Total issue categories:** 50  
**Total issue occurrences across all scripts:** 349  

---

## Array-as-object refactored to Class

**Affected patches: 86**

Drop targets and standup targets were initialized as plain arrays (e.g. `DT1 = Array(prim, sec, prim, sw, 0)`) and accessed with double-index notation (`DTArray(i)(0)`). Wine's VBScript does not support this pattern. Fix: refactor to a proper `Class DropTarget` with named properties.

Affected scripts:

- [1342729923_RollerCoasterTycoon(Stern2002)1.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/1342729923_RollerCoasterTycoon%28Stern2002%291.3/1342729923_RollerCoasterTycoon%28Stern2002%291.3.vbs#L3168)
- [2001 (Gottlieb 1971) v0.99a.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/2001%20%28Gottlieb%201971%29%20v0.99a/2001%20%28Gottlieb%201971%29%20v0.99a.vbs#L5025)
- [AC-DC LUCI Premium VR (Stern 2013) v1.1.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/AC-DC%20LUCI%20Premium%20VR%20%28Stern%202013%29%20v1.1.3/AC-DC%20LUCI%20Premium%20VR%20%28Stern%202013%29%20v1.1.3.vbs#L330)
- [Alien Poker (Williams 1980).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/AlienPoker%28Williams1980%291.0/Alien%20Poker%20%28Williams%201980%29.vbs#L1935)
- [Apollo 13 (Sega 1995) w VR Room v2.1.1.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Apollo%2013%20%28Sega%201995%29%20w%20VR%20Room%20v2.1.1/Apollo%2013%20%28Sega%201995%29%20w%20VR%20Room%20v2.1.1.vbs#L2624)
- [Atlantis (Bally 1989) w VR Room v2.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0/Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0.vbs#L3029)
- [Barracora (Williams 1981) w VR Room v2.1.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Barracora%20%28Williams%201981%29%20w%20VR%20Room%20v2.1.3/Barracora%20%28Williams%201981%29%20w%20VR%20Room%20v2.1.3.vbs#L2022)
- [Batman Forever (Sega 1995) 1.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Batman%20Forever%20%28Sega%201995%29%201.3/Batman%20Forever%20%28Sega%201995%29%201.3.vbs#L3518)
- [Batman [The Dark Knight] (Stern 2008) 1.16.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Batman%20%5BThe%20Dark%20Knight%5D%20%28Stern%202008%29%201.16/Batman%20%5BThe%20Dark%20Knight%5D%20%28Stern%202008%29%201.16.vbs#L905)
- [Batman [The Dark Knight] (Stern 2008) v1.0.8.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Batman%20%5BThe%20Dark%20Knight%5D%20%28Stern%202008%29%20v1.0.8/Batman%20%5BThe%20Dark%20Knight%5D%20%28Stern%202008%29%20v1.0.8.vbs#L824)
- [Batman [The Dark Knight] (Stern 2008).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/BatmanTheDarkKnight%28Stern2008%29bordmod1.0.7/Batman%20%5BThe%20Dark%20Knight%5D%20%28Stern%202008%29.vbs#L823)
- [BeastieBoysv1.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/BeastieBoysv1.0/BeastieBoysv1.0.vbs#L100)
- [Black Knight (Williams 1980).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Black%20Knight%20%28Williams%201980%29/Black%20Knight%20%28Williams%201980%29.vbs#L896)
- [Black Knight 2000 (Williams 1989) w VR Room v2.0.2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2/Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2.vbs#L2687)
- [Blackout (Williams 1980) v1.0.1 - SBR34.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Blackout%20%28Williams%201980%29%20v1.0.1%20-%20SBR34/Blackout%20%28Williams%201980%29%20v1.0.1%20-%20SBR34.vbs#L1888)
- [Blood Machines 2.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Blood%20Machines%202.0/Blood%20Machines%202.0.vbs#L205)
- [Bounty Hunter (Gottlieb 1985).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Bounty%20Hunter%20%28Gottlieb%201985%29/Bounty%20Hunter%20%28Gottlieb%201985%29.vbs#L1420)
- [Cactus Canyon (Bally 1998) VPW 1.1.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Cactus%20Canyon%20%28Bally%201998%29%20VPW%201.1/Cactus%20Canyon%20%28Bally%201998%29%20VPW%201.1.vbs#L108)
- [Cactus Jacks (Gottlieb 1991) w VR Room v2.0.2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2/Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2.vbs#L1544)
- [Capt. Fantastic and The Brown Dirt Cowboy (Bally 1976) 2.0.2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Capt.%20Fantastic%20and%20The%20Brown%20Dirt%20Cowboy%20%28Bally%201976%29%202.0.2/Capt.%20Fantastic%20and%20The%20Brown%20Dirt%20Cowboy%20%28Bally%201976%29%202.0.2.vbs#L3871)
- [Centaur (Bally 1981).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Centaur%20%28Bally%201981%29/Centaur%20%28Bally%201981%29.vbs#L2069)
- [Centigrade 37 (Gottlieb 1977).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Centigrade%2037%20%28Gottlieb%201977%29/Centigrade%2037%20%28Gottlieb%201977%29.vbs#L3290)
- [Checkpoint (Data East 1991)2.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Checkpoint%20%28Data%20East%201991%292.0/Checkpoint%20%28Data%20East%201991%292.0.vbs#L2301)
- [Comet (Williams 1985) w VR Room v2.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Comet%20%28Williams%201985%29%20w%20VR%20Room%20v2.3/Comet%20%28Williams%201985%29%20w%20VR%20Room%20v2.3.vbs#L2233)
- [Cue Ball Wizard (Gottlieb 1992) v1.1.2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Cue%20Ball%20Wizard%20%28Gottlieb%201992%29%20v1.1.2/Cue%20Ball%20Wizard%20%28Gottlieb%201992%29%20v1.1.2.vbs#L1503)
- [Diner (Williams 1990) VPW Mod 1.0.2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Diner%20%28Williams%201990%29%20VPW%20Mod%201.0.2/Diner%20%28Williams%201990%29%20VPW%20Mod%201.0.2.vbs#L4176)
- [Doctor Who (Bally 1992) VPW Mod v1.1.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Doctor%20Who%20%28Bally%201992%29%20VPW%20Mod%20v1.1/Doctor%20Who%20%28Bally%201992%29%20VPW%20Mod%20v1.1.vbs#L1111)
- [Dracula (Stern 1979).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Dracula%20%28Stern%201979%292.1.1/Dracula%20%28Stern%201979%29.vbs#L60)
- [Elektra (Bally 1981) w VR Room v2.0.7.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Elektra%20%28Bally%201981%29%20w%20VR%20Room%20v2.0.7/Elektra%20%28Bally%201981%29%20w%20VR%20Room%20v2.0.7.vbs#L2205)
- [Elvis_MOD_2.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Elvis_MOD_2.0/Elvis_MOD_2.0.vbs#L626)
- [Family Guy 1.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Family%20Guy%201.0/Family%20Guy%201.0.vbs#L3797)
- [Flash (Williams 1979).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Flash%20%28Williams%201979%29/Flash%20%28Williams%201979%29.vbs#L2124)
- [Fog, The (Gottlieb 1979) v2.5 for 10.7.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Fog%2C%20The%20%28Gottlieb%201979%29%20v2.5%20for%2010.7/Fog%2C%20The%20%28Gottlieb%201979%29%20v2.5%20for%2010.7.vbs#L3380)
- [Galaxy (Stern 1980).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Galaxy%20%28Stern%201980%29/Galaxy%20%28Stern%201980%29.vbs#L3083)
- [Game of Thrones LE (Stern 2015) VPW v1.0.2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%20v1.0.2/Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%20v1.0.2.vbs#L8960)
- [Genie (Gottlieb 1979).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Genie%20%28Gottlieb%201979%29/Genie%20%28Gottlieb%201979%29.vbs#L408)
- [Gorgar_1.1.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Gorgar_1.1/Gorgar_1.1.vbs#L2747)
- [Halloween 1978-1981 (Original 2022) 1.03.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Halloween%201978-1981%20%28Original%202022%29%201.03/Halloween%201978-1981%20%28Original%202022%29%201.03.vbs#L3670)
- [Hang Glider (Bally 1976) VPW v1.2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Hang%20Glider%20%28Bally%201976%29%20VPW%20v1.2/Hang%20Glider%20%28Bally%201976%29%20VPW%20v1.2.vbs#L1420)
- [Harlem Globetrotters (Bally 1979) v1.14.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Harlem%20Globetrotters%20%28Bally%201979%29%20v1.14/Harlem%20Globetrotters%20%28Bally%201979%29%20v1.14.vbs#L3055)
- [Harley Davidson (Sega 1999) v1.12.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Harley%20Davidson%20%28Sega%201999%29%20v1.12/Harley%20Davidson%20%28Sega%201999%29%20v1.12.vbs#L618)
- [Heat Wave (Williams 1964).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/HeatWave%28Williams1964%291.0/Heat%20Wave%20%28Williams%201964%29.vbs#L1988)
- [Indiana Jones The Pinball Adventure (Williams 1993) VPWmod v1.1.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Indiana%20Jones%20The%20Pinball%20Adventure%20%28Williams%201993%29%20VPWmod%20v1.1/Indiana%20Jones%20The%20Pinball%20Adventure%20%28Williams%201993%29%20VPWmod%20v1.1.vbs#L4357)
- [Iron Maiden (Original 2022) VPW 1.0.12.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Iron%20Maiden%20%28Original%202022%29%20VPW%201.0.12/Iron%20Maiden%20%28Original%202022%29%20VPW%201.0.12.vbs#L7957)
- [Iron Man Vault Edition (Stern 2010) VPW v1.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Iron%20Man%20Vault%20Edition%20%28Stern%202010%29%20VPW%20v1.0/Iron%20Man%20Vault%20Edition%20%28Stern%202010%29%20VPW%20v1.0.vbs#L3280)
- [Jack-Bot (Williams 1995).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Jack-Bot%20%28Williams%201995%29/Jack-Bot%20%28Williams%201995%29.vbs#L3570)
- [Jive Time (Williams 1970) 2.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Jive%20Time%20%28Williams%201970%29%202.0/Jive%20Time%20%28Williams%201970%29%202.0.vbs#L777)
- [Joker Poker EM (Gottlieb 1978) 1.6.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Joker%20Poker%20EM%20%28Gottlieb%201978%29%201.6/Joker%20Poker%20EM%20%28Gottlieb%201978%29%201.6.vbs#L2948)
- [Judge Dredd (Bally 1993) VPW v1.1.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Judge%20Dredd%20%28Bally%201993%29%20VPW%20v1.1/Judge%20Dredd%20%28Bally%201993%29%20VPW%20v1.1.vbs#L3045)
- [Jungle Lord (Williams 1981) w VR Room v2.01.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Jungle%20Lord%20%28Williams%201981%29%20w%20VR%20Room%20v2.01/Jungle%20Lord%20%28Williams%201981%29%20w%20VR%20Room%20v2.01.vbs#L139)
- [Laser Cue (Williams 1984) w VR Room 2.0.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Laser%20Cue%20%28Williams%201984%29%20w%20VR%20Room%202.0.0/Laser%20Cue%20%28Williams%201984%29%20w%20VR%20Room%202.0.0.vbs#L1847)
- [Laser War (Data East 1987) w VR Room v2.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0/Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0.vbs#L1409)
- [Led Zeppelin Pinball 2.5.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Led%20Zeppelin%20Pinball%202.5/Led%20Zeppelin%20Pinball%202.5.vbs#L461)
- [Magic City (Williams 1967).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Magic%20City%20%28Williams%201967%291.0a/Magic%20City%20%28Williams%201967%29.vbs#L4325)
- [Maverick (Data East 1994) VPW v1.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Maverick%20%28Data%20East%201994%29%20VPW%20v1.3/Maverick%20%28Data%20East%201994%29%20VPW%20v1.3.vbs#L848)
- [Medusa (Bally 1981) w VR Room v2.0.1.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Medusa%20%28Bally%201981%29%20w%20VR%20Room%20v2.0.1/Medusa%20%28Bally%201981%29%20w%20VR%20Room%20v2.0.1.vbs#L2297)
- [Meteor (Stern 1979).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Meteor%20%28Stern%201979%29/Meteor%20%28Stern%201979%29.vbs#L1593)
- [Monster Bash (Williams 1998) VPWmod v1.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Monster%20Bash%20%28Williams%201998%29%20VPWmod%20v1.0/Monster%20Bash%20%28Williams%201998%29%20VPWmod%20v1.0.vbs#L2558)
- [MrDoom (Recel 1979)1.3.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/MrDoom%20%28Recel%201979%291.3.0/MrDoom%20%28Recel%201979%291.3.0.vbs#L1836)
- [MrEvil (Recel 1979)1.3.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/MrEvil%20%28Recel%201979%291.3.0/MrEvil%20%28Recel%201979%291.3.0.vbs#L1844)
- [Nine Ball (Stern 1980).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Nine%20Ball%20%28Stern%201980%29/Nine%20Ball%20%28Stern%201980%29.vbs#L1591)
- [O Brother Where Art Thou (Zoss 2021)1_6.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/O%20Brother%20Where%20Art%20Thou%20%28Zoss%202021%291_6/O%20Brother%20Where%20Art%20Thou%20%28Zoss%202021%291_6.vbs#L2520)
- [Pharaoh (Williams 1981) w VR Room v2.0.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Pharaoh%20%28Williams%201981%29%20w%20VR%20Room%20v2.0.3/Pharaoh%20%28Williams%201981%29%20w%20VR%20Room%20v2.0.3.vbs#L2385)
- [PinBlob (CLV 2024).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/PinBlob%20%28CLV%202024%29/PinBlob%20%28CLV%202024%29.vbs#L325)
- [PinBot (Williams 1986) 2.1.1.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/PinBot%20%28Williams%201986%29%202.1.1/PinBot%20%28Williams%201986%29%202.1.1.vbs#L3553)
- [Power Play (Bally 1977).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/1256692067_PowerPlay%28Bally1977%292.1/Power%20Play%20%28Bally%201977%29.vbs#L922)
- [Robocop (Data East 1989)_drakkon(mod_1.2).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Robocop%20%28Data%20East%201989%29_drakkon%28mod_1.2%29/Robocop%20%28Data%20East%201989%29_drakkon%28mod_1.2%29.vbs#L2457)
- [Seawitch (Stern 1980).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Seawitch%20%28Stern%201980%29/Seawitch%20%28Stern%201980%29.vbs#L1278)
- [Solar Fire (Williams 1981) w VR Room v2.0.5.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Solar%20Fire%20%28Williams%201981%29%20w%20VR%20Room%20v2.0.5/Solar%20Fire%20%28Williams%201981%29%20w%20VR%20Room%20v2.0.5.vbs#L2398)
- [Space Shuttle (Williams 1984).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Space%20Shuttle%20%28Williams%201984%292.0/Space%20Shuttle%20%28Williams%201984%29.vbs#L581)
- [Star Trek (Bally 1979) 2.1.1.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Star%20Trek%20%28Bally%201979%29%202.1.1/Star%20Trek%20%28Bally%201979%29%202.1.1.vbs#L731)
- [Star Wars (Data East 1992) VPW v1.2.2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Star%20Wars%20%28Data%20East%201992%29%20VPW%20v1.2.2/Star%20Wars%20%28Data%20East%201992%29%20VPW%20v1.2.2.vbs#L3651)
- [Stars (Stern 1978).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Stars%20%28Stern%201978%291.0.2/Stars%20%28Stern%201978%29.vbs#L1528)
- [Stellar Wars (Williams 1979).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Stellar%20Wars%20%28Williams%201979%29/Stellar%20Wars%20%28Williams%201979%29.vbs#L2163)
- [Stingray (Stern 1977).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Stingray%20%28Stern%201977%29%202.0/Stingray%20%28Stern%201977%29.vbs#L1696)
- [Strikes And Spares (Bally 1978) 2.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Strikes%20And%20Spares%20%28Bally%201978%29%202.0/Strikes%20And%20Spares%20%28Bally%201978%29%202.0.vbs#L2093)
- [Strip Joker Poker (Gottlieb 1978) 1.5.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Strip-Joker-Poker-%28Gottlieb-1978%29-1.5/Strip%20Joker%20Poker%20%28Gottlieb%201978%29%201.5.vbs#L2996)
- [Swords of Fury (Williams 1988).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Swords%20of%20Fury%20%28Williams%201988%29/Swords%20of%20Fury%20%28Williams%201988%29.vbs#L610)
- [TX-Sector (Gottlieb 1988) SG1bsoN Mod V1.1.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/TX-Sector%20%28Gottlieb%201988%29%20SG1bsoN%20Mod%20V1.1/TX-Sector%20%28Gottlieb%201988%29%20SG1bsoN%20Mod%20V1.1.vbs#L809)
- [Tales from the Crypt (Data East 1993) VPW v1.22.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Tales%20from%20the%20Crypt%20%28Data%20East%201993%29%20VPW%20v1.22/Tales%20from%20the%20Crypt%20%28Data%20East%201993%29%20VPW%20v1.22.vbs#L1528)
- [Transporter the Rescue (Midway 1989) VPW v1.05.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Transporter%20the%20Rescue%20%28Midway%201989%29%20VPW%20v1.05/Transporter%20the%20Rescue%20%28Midway%201989%29%20VPW%20v1.05.vbs#L3331)
- [Twilight Zone (Bally 1993) 2.3.6.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Twilight%20Zone%20%28Bally%201993%29%202.3.6/Twilight%20Zone%20%28Bally%201993%29%202.3.6.vbs#L2884)
- [Twister (Sega 1996) v2.0 w VR Room.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Twister%20%28Sega%201996%29%20v2.0%20w%20VR%20Room/Twister%20%28Sega%201996%29%20v2.0%20w%20VR%20Room.vbs#L2045)
- [Viking (Bally 1980).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Viking%20%28Bally%201980%29/Viking%20%28Bally%201980%29.vbs#L2124)
- [fireball II VPX.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/fireball%20II%20VPX/fireball%20II%20VPX.vbs#L38)

## Parenthesis on first argument of a procedure call not handled correctly by Wine VBScript

**Affected patches: 67**

When the first argument of a Sub/procedure call starts with `(`, Wine's VBScript parses it as a call with explicit parentheses and treats the rest of the expression as a separate statement. Any arithmetic that continues after the closing `)` is evaluated separately and discarded. Example: `AddScore (a+b)*c` is parsed as `AddScore(a+b)` then `*c` is discarded. The same rule applies inside `ExecuteGlobal` strings — `SetLamp (me.UserValue - INT(me.UserValue)) * 100` becomes `SetLamp (me.UserValue - INT(me.UserValue))` with the `* 100` thrown away. Fix: rearrange the expression so it does not start with `(`, or move the multiplier to the front, e.g. `AddScore (a+b)*c` → `AddScore c*(a+b)`, or wrap the entire expression in double parentheses.

Affected scripts:

- [2104398928_CARtoonsRC(Nailed2021)v1.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/2104398928_CARtoonsRC%28Nailed2021%29v1.3/2104398928_CARtoonsRC%28Nailed2021%29v1.3.vbs#L841)
- [A Charlie Brown Christmas feat. Vince Guaraldi (iDigStuff 2023).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/A%20Charlie%20Brown%20Christmas%20feat.%20Vince%20Guaraldi%20%28iDigStuff%202023%29%201.0/A%20Charlie%20Brown%20Christmas%20feat.%20Vince%20Guaraldi%20%28iDigStuff%202023%29.vbs#L1198)
- [Aladdin's Castle (Bally 1976) - DOZER - MJR_1.01.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Aladdin%27s%20Castle%20%28Bally%201976%29%20-%20DOZER%20-%20MJR_1.01/Aladdin%27s%20Castle%20%28Bally%201976%29%20-%20DOZER%20-%20MJR_1.01.vbs#L625)
- [Attack On Titan (cHuG_MOD_1.4).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Attack%20On%20Titan%20%28cHuG_MOD_1.4%29/Attack%20On%20Titan%20%28cHuG_MOD_1.4%29.vbs#L2475)
- [Ben-Hur (Staal 1977) V1.1.1 DT-FS-VR-MR Ext2k Conversion.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Ben-Hur%20%28Staal%201977%29%20V1.1.1%20DT-FS-VR-MR%20Ext2k%20Conversion/Ben-Hur%20%28Staal%201977%29%20V1.1.1%20DT-FS-VR-MR%20Ext2k%20Conversion.vbs#L1330)
- [Big Deal (Williams 1977)V2.1.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Big%20Deal%20%28Williams%201977%29V2.1/Big%20Deal%20%28Williams%201977%29V2.1.vbs#L1067)
- [Big Horse (Maresa 1975) v55_VPX8.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Big%20Horse%20%28Maresa%201975%29%20v55_VPX8/Big%20Horse%20%28Maresa%201975%29%20v55_VPX8.vbs#L1107)
- [Big Star (Williams 1972) v4.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Big%20Star%20%28Williams%201972%29%20v4.3/Big%20Star%20%28Williams%201972%29%20v4.3.vbs#L1113)
- [Black & Red (Inder 1975) v55_VPX8.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Black%20%26%20Red%20%28Inder%201975%29%20v55_VPX8/Black%20%26%20Red%20%28Inder%201975%29%20v55_VPX8.vbs#L931)
- [Bumper (Bill Port - 1977) v4.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Bumper%20%28Bill%20Port%20-%201977%29%20v4/Bumper%20%28Bill%20Port%20-%201977%29%20v4.vbs#L973)
- [Cannes (Segasa 1976)V1.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Cannes%20%28Segasa%201976%29V1.3/Cannes%20%28Segasa%201976%29V1.3.vbs#L1113)
- [DORAEMON.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/DORAEMON/DORAEMON.vbs#L1632)
- [DS.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Drunken%20Santa%20%28Original%202021%29/DS.vbs#L1724)
- [DUCKTALES.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Ducktales%20%28Orginal%202020%29/DUCKTALES.vbs#L1710)
- [FNAF.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/FNAF/FNAF.vbs#L1934)
- [Fifteen (Inder 1974) v4.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Fifteen%20%28Inder%201974%29%20v4/Fifteen%20%28Inder%201974%29%20v4.vbs#L899)
- [Fireball XL5.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Fireball%20XL5/Fireball%20XL5.vbs#L2424)
- [Game of Thrones LE (Stern 2015) VPW 1.2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%201.2/Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%201.2.vbs#L8915)
- [Game of Thrones LE (Stern 2015) VPW v1.0.2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%20v1.0.2/Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%20v1.0.2.vbs#L8960)
- [Gemini (Gottlieb 1978).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Gemini%20%28Gottlieb%201978%29/Gemini%20%28Gottlieb%201978%29.vbs#L1173)
- [Gun Men (Staal 1979).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Gun%20Men%20%28Staal%201979%29/Gun%20Men%20%28Staal%201979%29.vbs#L1226)
- [HALO.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/HALO/HALO.vbs#L1957)
- [Hellraiser 1.2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Hellraiser%201.2/Hellraiser%201.2.vbs#L2787)
- [Honey (Williams 1971)V1.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Honey%20%28Williams%201971%29V1.3/Honey%20%28Williams%201971%29V1.3.vbs#L1063)
- [Indiana Jones (Stern 2008)-Hanibal-2.6.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Indiana%20Jones%20%28Stern%202008%29-Hanibal-2.6/Indiana%20Jones%20%28Stern%202008%29-Hanibal-2.6.vbs#L2117)
- [Inhabiting Mars RC 4.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Inhabiting%20Mars%20RC%204/Inhabiting%20Mars%20RC%204.vbs#L2311)
- [Liberty Bell (Williams 1977) V1.01.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Liberty%20Bell%20%28Williams%201977%29%20V1.01/Liberty%20Bell%20%28Williams%201977%29%20V1.01.vbs#L404)
- [Lightning Ball (Gottlieb 1959).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Lightning%20Ball%20%28Gottlieb%201959%29/Lightning%20Ball%20%28Gottlieb%201959%29.vbs#L1216)
- [Luck Smile (Inder 1976) 1p v55_VPX8.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Luck%20Smile%20%28Inder%201976%29%201p%20v55_VPX8/Luck%20Smile%20%28Inder%201976%29%201p%20v55_VPX8.vbs#L897)
- [Luck Smile (Inder 1976) 4p v55_VPX8.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Luck%20Smile%20%28Inder%201976%29%204p%20v55_VPX8/Luck%20Smile%20%28Inder%201976%29%204p%20v55_VPX8.vbs#L898)
- [Mago de Oz - the pinball v4.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Mago%20de%20Oz%20-%20the%20pinball%20v4.3/Mago%20de%20Oz%20-%20the%20pinball%20v4.3.vbs#L2557)
- [Metal Slug_1.05.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Metal%20Slug_1.05/Metal%20Slug_1.05.vbs#L72)
- [Monaco (Segasa 1977)V1.4.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Monaco%20%28Segasa%201977%29V1.4/Monaco%20%28Segasa%201977%29V1.4.vbs#L1095)
- [Monkey Island VR Room.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Monkey%20Island%20VR%20Room/Monkey%20Island%20VR%20Room.vbs#L240)
- [Motley Crue SS.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Motley%20Crue%20SS/Motley%20Crue%20SS.vbs#L5229)
- [Munsters (Original 2020) 1.05.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Munsters%20%28Original%202020%29%201.05/Munsters%20%28Original%202020%29%201.05.vbs#L257)
- [PVM.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Putin%20Vodka%20Mania%20%28Original%202022%29/PVM.vbs#L2181)
- [Pat Hand (Williams 1975)V1.4.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Pat%20Hand%20%28Williams%201975%29V1.4/Pat%20Hand%20%28Williams%201975%29V1.4.vbs#L1081)
- [Punchout.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Punchout/Punchout.vbs#L2167)
- [Rancho (Williams 1976)V1.4.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Rancho%20%28Williams%201976%29V1.4/Rancho%20%28Williams%201976%29V1.4.vbs#L1076)
- [Rattlecan v1.5.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Rattlecan%20v1.5.3/Rattlecan%20v1.5.3.vbs#L604)
- [Riccione.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Riccione_V1.4/Riccione.vbs#L1128)
- [Route66.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Route66%280_2%29/Route66.vbs#L307)
- [Running Horse (Inder 1976) v55_VPX8.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Running%20Horse%20%28Inder%201976%29%20v55_VPX8/Running%20Horse%20%28Inder%201976%29%20v55_VPX8.vbs#L900)
- [SMGP.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/SMGP%20%28Super%20Mario%20Galaxy%20Pinball%29/SMGP.vbs#L1848)
- [SPP.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/SouthParkPinballv1.2/SPP.vbs#L2185)
- [Seven Winner (Inder 1973) v4.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Seven%20Winner%20%28Inder%201973%29%20v4/Seven%20Winner%20%28Inder%201973%29%20v4.vbs#L869)
- [Skylab (Williams 1974)V1.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Skylab%20%28Williams%201974%29V1.3/Skylab%20%28Williams%201974%29V1.3.vbs#L1134)
- [Sonic The Hedgehog (Brendan Bailey 2005) VPX MOD 1.33.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Sonic%20The%20Hedgehog%20%28Brendan%20Bailey%202005%29%20VPX%20MOD%201.33/Sonic%20The%20Hedgehog%20%28Brendan%20Bailey%202005%29%20VPX%20MOD%201.33.vbs#L156)
- [Sons Of Anarchy (Original 2019).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Sons%20Of%20Anarchy%20%28Original%202019%29/Sons%20Of%20Anarchy%20%28Original%202019%29.vbs#L51)
- [Stardust (Williams 1971) v4.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Stardust%20%28Williams%201971%29%20v4/Stardust%20%28Williams%201971%29%20v4.vbs#L933)
- [Streets of Rage (TBA 2018).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Streets%20of%20Rage%20%28TBA%202018%29/Streets%20of%20Rage%20%28TBA%202018%29.vbs#L72)
- [Super Star (Williams 1972) v4.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Super%20Star%20%28Williams%201972%29%20v4.3/Super%20Star%20%28Williams%201972%29%20v4.3.vbs#L1080)
- [The Grinch (Original 2022) pinballfan2018.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/The%20Grinch%20-%20Original%202022/The%20Grinch%20%28Original%202022%29%20pinballfan2018.vbs#L870)
- [TheATeam.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/TheATeam_2.0.7/TheATeam.vbs#L12146)
- [Van Halen 1.2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Van%20Halen%201.2/Van%20Halen%201.2.vbs#L4926)
- [indochinecentral.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/indochinecentral/indochinecentral.vbs#L2556)
- [indochinecentralPUP.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/indochinecentralPUP/indochinecentralPUP.vbs#L2556)
- [monkeyisland.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Monkeyislandv1.1/monkeyisland.vbs#L2441)
- [pizzatime-65.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/pizzatime-65/pizzatime-65.vbs#L48)
- [pulp_fiction.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Pulp%20Fiction%203.1/pulp_fiction.vbs#L2621)
- [speakeasy2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/speakeasy2/speakeasy2.vbs#L1775)
- [speakeasy4.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/speakeasy4/speakeasy4.vbs#L1775)
- [wackyraces.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/wackyraces/wackyraces.vbs#L2829)

## CreateObject / FileSystemObject / WshShell removed (not supported in Wine)

**Affected patches: 22**

`CreateObject("WScript.Shell")`, `CreateObject("Scripting.FileSystemObject")`, and `WshShell.RegWrite/RegRead` are Windows-only COM objects that are not available under Wine. Typically used for music folder scanning, registry access, or file system operations. These calls are removed or replaced with standalone-compatible equivalents.

Affected scripts:

- [ABBAv2.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/ABBAv2.0/ABBAv2.0.vbs#L47)
- [Blood Machines 2.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Blood%20Machines%202.0/Blood%20Machines%202.0.vbs#L205)
- [Bob Marley Mod.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Bob%20Marley%20Mod/Bob%20Marley%20Mod.vbs#L59)
- [DarkPrincess1.3.1.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/DarkPrincess1.3.1/DarkPrincess1.3.1.vbs#L161)
- [Darkest Dungeon pupevent_2.3c.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Darkest%20Dungeon%20pupevent_2.3c/Darkest%20Dungeon%20pupevent_2.3c.vbs#L175)
- [Diablo Pinball v4.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Diablo%20Pinball%20v4.3/Diablo%20Pinball%20v4.3.vbs#L2344)
- [Fire Action (Taito do Brasil - 1980) v4.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Fire%20Action%20%28Taito%20do%20Brasil%20-%201980%29%20v4/Fire%20Action%20%28Taito%20do%20Brasil%20-%201980%29%20v4.vbs#L1070)
- [Freddys Nightmares.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Freddys%20Nightmares/Freddys%20Nightmares.vbs#L110)
- [Ice Age Christmas (Original Balutito 2021) endeemillr mod.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Ice%20Age%20Christmas%20%28Original%20Balutito%202021%29%20endeemillr%20mod/Ice%20Age%20Christmas%20%28Original%20Balutito%202021%29%20endeemillr%20mod.vbs#L151)
- [Iron Maiden Virtual Time (Original 2020).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Iron%20Maiden%20Virtual%20Time%20%28Original%202020%29/Iron%20Maiden%20Virtual%20Time%20%28Original%202020%29.vbs#L408)
- [KISS (Stern 2015)_Bigus(MOD)1.1.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/KISS%20%28Stern%202015%29_Bigus%28MOD%291.1/KISS%20%28Stern%202015%29_Bigus%28MOD%291.1.vbs#L3642)
- [Lawman (Gottlieb 1971) 1.1.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Lawman%20%28Gottlieb%201971%29%201.1/Lawman%20%28Gottlieb%201971%29%201.1.vbs#L340)
- [Led Zeppelin Pinball 2.5.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Led%20Zeppelin%20Pinball%202.5/Led%20Zeppelin%20Pinball%202.5.vbs#L461)
- [Metal Slug_1.05.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Metal%20Slug_1.05/Metal%20Slug_1.05.vbs#L72)
- [Sonic The Hedgehog (Brendan Bailey 2005) VPX MOD 1.33.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Sonic%20The%20Hedgehog%20%28Brendan%20Bailey%202005%29%20VPX%20MOD%201.33/Sonic%20The%20Hedgehog%20%28Brendan%20Bailey%202005%29%20VPX%20MOD%201.33.vbs#L156)
- [Space Shuttle (Taito do Brasil - 1982) v4.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Space%20Shuttle%20%28Taito%20do%20Brasil%20-%201982%29%20v4/Space%20Shuttle%20%28Taito%20do%20Brasil%20-%201982%29%20v4.vbs#L962)
- [Spooky Wednesday Pro.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Spooky%20Wednesday%20Pro/Spooky%20Wednesday%20Pro.vbs#L46)
- [Spooky_Wednesday VPX 2024.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Spooky_Wednesday%20VPX%202024/Spooky_Wednesday%20VPX%202024.vbs#L46)
- [Student Prince (Williams 1968).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Student%20Prince%20%28Williams%201968%29/Student%20Prince%20%28Williams%201968%29.vbs#L12)
- [Three Angels (Original 2018) 1.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Three%20Angels%20%28Original%202018%29%201.3/Three%20Angels%20%28Original%202018%29%201.3.vbs#L291)
- [Thundercats Pinball v1.0.9.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Thundercats%20Pinball%20v1.0.9/Thundercats%20Pinball%20v1.0.9.vbs#L63)
- [Vortex (Taito do Brasil - 1981) v4.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4/Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4.vbs#L46)

## DMD standalone compatibility setup (ImplicitDMD_Init / UseVPMDMD)

**Affected patches: 22**

Scripts that use a separate DMD script (`.vbs.dmd`) need a `Sub ImplicitDMD_Init` entry point and/or a `Dim UseVPMDMD` variable so that VPX Standalone can render the DMD correctly. These additions enable the DMD display to work without VPinMAME on standalone builds.

Affected scripts:

- Avatar (Stern 2012) v1.12.vbs
- [BarbWire(Gottlieb1996)JoePicassoModv1.2.vbs.dmd](https://github.com/francisdb/vpx-standalone-scripts/blob/master/BarbWire%28Gottlieb1996%29JoePicassoModv1.2/BarbWire%28Gottlieb1996%29JoePicassoModv1.2.vbs.dmd#L1)
- [Blood Machines 2.0.vbs.dmd](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Blood%20Machines%202.0/Blood%20Machines%202.0.vbs.dmd#L100)
- [Bounty Hunter (Gottlieb 1985).vbs.dmd](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Bounty%20Hunter%20%28Gottlieb%201985%29/Bounty%20Hunter%20%28Gottlieb%201985%29.vbs.dmd#L19)
- [Cactus Canyon (Bally 1998) VPW 1.1.vbs.dmd](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Cactus%20Canyon%20%28Bally%201998%29%20VPW%201.1/Cactus%20Canyon%20%28Bally%201998%29%20VPW%201.1.vbs.dmd#L133)
- [Diner (Williams 1990) VPW Mod 1.0.2.vbs.dmd](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Diner%20%28Williams%201990%29%20VPW%20Mod%201.0.2/Diner%20%28Williams%201990%29%20VPW%20Mod%201.0.2.vbs.dmdcolored#L446)
- Futurama (Original 2024) v1.1.vbs
- [Game of Thrones LE (Stern 2015) VPW v1.0.2.vbs.dmd](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%20v1.0.2/Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%20v1.0.2.vbs.dmd#L163)
- [Goin' Nuts (Gottlieb 1983).vbs.dmd](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Goin%27%20Nuts%20%28Gottlieb%201983%29/Goin%27%20Nuts%20%28Gottlieb%201983%29.vbs.dmd#L29)
- [Hook (Data East 1992)_VPWmod_v1.0.vbs.dmd](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Hook%20%28Data%20East%201992%29_VPWmod_v1.0/Hook%20%28Data%20East%201992%29_VPWmod_v1.0.vbs.dmd#L91)
- [Pink Panther (Gottlieb 1981).vbs.dmd](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Pink%20Panther%20%28Gottlieb%201981%29/Pink%20Panther%20%28Gottlieb%201981%29.vbs.dmd#L8)
- [Scared Stiff (Bally 1996) v1.29.vbs.dmd](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Scared%20Stiff%20%28Bally%201996%29%20v1.29/Scared%20Stiff%20%28Bally%201996%29%20v1.29.vbs.dmd#L15)
- [Seawitch (Stern 1980).vbs.dmd](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Seawitch%20%28Stern%201980%29/Seawitch%20%28Stern%201980%29.vbs.dmd#L21)
- [South Park (Sega 1999) 1.3.vbs.dmd](https://github.com/francisdb/vpx-standalone-scripts/blob/master/South%20Park%20%28Sega%201999%29%201.3/South%20Park%20%28Sega%201999%29%201.3.vbs.dmd#L38)
- [Space Shuttle (Williams 1984).vbs.dmd](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Space%20Shuttle%20%28Williams%201984%292.0/Space%20Shuttle%20%28Williams%201984%29.vbs.dmd#L57)
- [Star Trek LE (Stern 2013) v1.09.vbs.dmd](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Star%20Trek%20LE%20%28Stern%202013%29%20v1.09/Star%20Trek%20LE%20%28Stern%202013%29%20v1.09.vbs.dmd#L91)
- [Taxi (Williams 1988) VPW v1.2.2.vbs.dmd](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Taxi%20%28Williams%201988%29%20VPW%20v1.2.2/Taxi%20%28Williams%201988%29%20VPW%20v1.2.2.vbs.dmd#L87)
- The Goonies Never Say Die Pinball (VPW 2021) v1.4.vbs
- [Three Angels (Original 2018) 1.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Three%20Angels%20%28Original%202018%29%201.3/Three%20Angels%20%28Original%202018%29%201.3.vbs#L291)
- X-Files (Sega 1997) v1.29.vbs
- [X-Files Hanibal 4k LED Edition.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/X-Files%20Hanibal%204k%20LED%20Edition/X-Files%20Hanibal%204k%20LED%20Edition.vbs#L15)
- [fireball II VPX.vbs.dmd](https://github.com/francisdb/vpx-standalone-scripts/blob/master/fireball%20II%20VPX/fireball%20II%20VPX.vbs.dmd#L38)

## Const declaration moved to different position (execution order fix)

**Affected patches: 16**

VBScript evaluates code sequentially; a `Const` must be declared before its first use. In some scripts the constant was declared *after* the code that referenced it. Fix: move the `Const` declaration to an earlier position in the file.

Affected scripts:

- [AC-DC Pro Vault-1.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/AC-DC%20ninuzzu%201.0%201.5/AC-DC%20Pro%20Vault-1.0.vbs#L89)
- [AC-DC Pro-1.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/AC-DC%20ninuzzu%201.0%201.5/AC-DC%20Pro-1.0.vbs#L89)
- [AC-DC_Back In Black LE-1.5.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/AC-DC%20ninuzzu%201.0%201.5/AC-DC_Back%20In%20Black%20LE-1.5.vbs#L98)
- [AC-DC_Let There Be Rock LE-1.5.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/AC-DC%20ninuzzu%201.0%201.5/AC-DC_Let%20There%20Be%20Rock%20LE-1.5.vbs#L98)
- [AC-DC_Luci-1.5.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/AC-DC%20ninuzzu%201.0%201.5/AC-DC_Luci-1.5.vbs#L98)
- [AC-DC_Premium-1.5.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/AC-DC%20ninuzzu%201.0%201.5/AC-DC_Premium-1.5.vbs#L98)
- [Aztec (Williams 1976) 1.3 Mod Citedor JPJ-ARNGRIM-CED Team PP.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Aztec%20%28Williams%201976%29%201.3%20Mod%20Citedor%20JPJ-ARNGRIM-CED%20Team%20PP/Aztec%20%28Williams%201976%29%201.3%20Mod%20Citedor%20JPJ-ARNGRIM-CED%20Team%20PP.vbs#L1)
- [Aztec High-Tapped (Williams 1976).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Aztec%20High-Tapped%20%28Williams%201976%29/Aztec%20High-Tapped%20%28Williams%201976%29.vbs#L1747)
- [Cybernaut (Bally 1985)_Bigus(MOD)1.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Cybernaut%20%28Bally%201985%29/Cybernaut%20%28Bally%201985%29_Bigus%28MOD%291.0.vbs#L40)
- [Cybernaut darkmod.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/cybernaut%20Bally/Cybernaut%20darkmod.vbs#L43)
- [Cybernaut.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/cybernaut%20Bally/Cybernaut.vbs#L55)
- [Fireball Classic (Bally 1984)1.2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Fireball%20Classic%20%28Bally%201984%291.2/Fireball%20Classic%20%28Bally%201984%291.2.vbs#L1)
- [Flash Gordon (Bally 1981) VPW Mod v3.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Flash%20Gordon%20%28Bally%201981%29%20VPW%20Mod%20v3.0/Flash%20Gordon%20%28Bally%201981%29%20VPW%20Mod%20v3.0.vbs#L74)
- [Harlem Globetrotters (Bally 1979) VPW 1.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Harlem%20Globetrotters%20%28Bally%201979%29%20VPW%201.0/Harlem%20Globetrotters%20%28Bally%201979%29%20VPW%201.0.vbs#L249)
- [High_Speed_(Williams 1986) v0.107a (MOD 3.0).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/High_Speed_%28Williams%201986%29%20v0.107a%20%28MOD%203.0%29/High_Speed_%28Williams%201986%29%20v0.107a%20%28MOD%203.0%29.vbs#L6)
- [Rollercoaster Tycoon (Stern 2002) VPW 2.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Rollercoaster%20Tycoon%20%28Stern%202002%29%20VPW%202.3/Rollercoaster%20Tycoon%20%28Stern%202002%29%20VPW%202.3.vbs#L53)

## NVramPatch calls removed (not supported in standalone)

**Affected patches: 13**

`NVramPatchLoad`, `NVramPatchExit`, and `NVramPatchKeyCheck` rely on a Windows-only NVram helper that is not available in standalone builds. Calls to these functions are commented out or removed.

Affected scripts:

- [ABBAv2.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/ABBAv2.0/ABBAv2.0.vbs#L47)
- [Bob Marley Mod.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Bob%20Marley%20Mod/Bob%20Marley%20Mod.vbs#L59)
- [Cosmic (Taito do Brasil - 1980) v4.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Cosmic%20%28Taito%20do%20Brasil%20-%201980%29%20v4/Cosmic%20%28Taito%20do%20Brasil%20-%201980%29%20v4.vbs#L56)
- [Last Starfighter, The (Taito, 1983) hybrid v1.04.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Last%20Starfighter%2C%20The%20%28Taito%2C%201983%29%20hybrid%20v1.04/Last%20Starfighter%2C%20The%20%28Taito%2C%201983%29%20hybrid%20v1.04.vbs#L104)
- [Mr Black (Taito do Brasil - 1984) v4.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Mr%20Black%20%28Taito%20do%20Brasil%20-%201984%29%20v4/Mr%20Black%20%28Taito%20do%20Brasil%20-%201984%29%20v4.vbs#L61)
- [Oba Oba (Taito do Brasil - 1979) v55_VPX8.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Oba%20Oba%20%28Taito%20do%20Brasil%20-%201979%29%20v55_VPX8/Oba%20Oba%20%28Taito%20do%20Brasil%20-%201979%29%20v55_VPX8.vbs#L55)
- [Rally (Taito do Brasil - 1980) v4.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Rally%20%28Taito%20do%20Brasil%20-%201980%29%20v4/Rally%20%28Taito%20do%20Brasil%20-%201980%29%20v4.vbs#L60)
- [Titan (Taito do Brasil - 1981) v4.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Titan%20%28Taito%20do%20Brasil%20-%201981%29%20v4/Titan%20%28Taito%20do%20Brasil%20-%201981%29%20v4.vbs#L57)
- [Topaz (Inder 1979).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Topaz%20%28Inder%201979%29/Topaz%20%28Inder%201979%29.vbs#L100)
- [Vortex (Taito do Brasil - 1981) v4.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4/Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4.vbs#L46)
- [Warriors The Full DMD 2.0 (Iceman 2023).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Warriors%20The%20Full%20DMD%202.0%20%28Iceman%202023%29/Warriors%20The%20Full%20DMD%202.0%20%28Iceman%202023%29.vbs#L14)
- [Wednesday (Netflix 2023).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Wednesday%20%28Netflix%202023%29/Wednesday%20%28Netflix%202023%29.vbs#L19)

## cGameName (ROM name) constant fix

**Affected patches: 10**

The `cGameName` constant sets the ROM name passed to VPinMAME. An incorrect ROM name prevents the game from loading. This fix corrects the ROM name string to match the actual ROM file name.

Affected scripts:

- [Bad Cats (Williams 1989) VPW 3.0.vbs.dmd](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Bad%20Cats%20%28Williams%201989%29%20VPW%203.0/Bad%20Cats%20%28Williams%201989%29%20VPW%203.0.vbs.dmdcolored#L384)
- [BeastieBoysv1.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/BeastieBoysv1.0/BeastieBoysv1.0.vbs#L100)
- [Death Proof Balutito V2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Death%20Proof%20Balutito%20V2/Death%20Proof%20Balutito%20V2.vbs#L207)
- [PinBlob (CLV 2024).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/PinBlob%20%28CLV%202024%29/PinBlob%20%28CLV%202024%29.vbs#L325)
- [Safe Cracker (Bally 1996) v1.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Safe%20Cracker%20%28Bally%201996%29%20v1.0/Safe%20Cracker%20%28Bally%201996%29%20v1.0.vbs#L54)
- [Simpsons Treehouse of Horror MOD v2023.3 (Simpsons Pinball Party, The (Stern 2003) VPW 2.0.3 base).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Simpsons%20Treehouse%20of%20Horror%20MOD%20v2023.3%20%28Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3%20base%29/Simpsons%20Treehouse%20of%20Horror%20MOD%20v2023.3%20%28Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3%20base%29.vbs#L95)
- [South Park - Halloween.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/South%20Park%20-%20Halloween/South%20Park%20-%20Halloween.vbs#L49)
- [Theatre of Magic (Bally 1995) 2.4.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Theatre%20of%20Magic%20%28Bally%201995%29%202.4/Theatre%20of%20Magic%20%28Bally%201995%29%202.4.vbs#L214)
- [Warriors The Full DMD 2.0 (Iceman 2023).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Warriors%20The%20Full%20DMD%202.0%20%28Iceman%202023%29/Warriors%20The%20Full%20DMD%202.0%20%28Iceman%202023%29.vbs#L14)
- [Wednesday (Netflix 2023).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Wednesday%20%28Netflix%202023%29/Wednesday%20%28Netflix%202023%29.vbs#L19)

## GetDMDColor function call removed (not available in standalone)

**Affected patches: 8**

`GetDMDColor` is a helper function that does not exist in all script environments. Calling it causes a runtime error on standalone. The call is removed or commented out.

Affected scripts:

- [DarkPrincess1.3.1.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/DarkPrincess1.3.1/DarkPrincess1.3.1.vbs#L161)
- [Metal Slug_1.05.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Metal%20Slug_1.05/Metal%20Slug_1.05.vbs#L72)
- [Motley Crue SS.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Motley%20Crue%20SS/Motley%20Crue%20SS.vbs#L5229)
- [Sons Of Anarchy (Original 2019).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Sons%20Of%20Anarchy%20%28Original%202019%29/Sons%20Of%20Anarchy%20%28Original%202019%29.vbs#L51)
- [Spooky Wednesday Pro.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Spooky%20Wednesday%20Pro/Spooky%20Wednesday%20Pro.vbs#L46)
- [Spooky_Wednesday VPX 2024.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Spooky_Wednesday%20VPX%202024/Spooky_Wednesday%20VPX%202024.vbs#L46)
- [Streets of Rage (TBA 2018).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Streets%20of%20Rage%20%28TBA%202018%29/Streets%20of%20Rage%20%28TBA%202018%29.vbs#L72)
- [Three Angels (Original 2018) LW.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Three%20Angels%20%28Original%202018%29%20LW/Three%20Angels%20%28Original%202018%29%20LW.vbs#L335)

## Execute/ExecuteGlobal with object Set in loop (replaced with IsObject check)

**Affected patches: 6**

`Execute "Set Lights(" & i & ") = L" & i` creates object references dynamically in a loop. In Wine's VBScript, this pattern fails when some loop variables are not objects. Fix: replace with an `If IsObject(eval("L" & i)) Then` guard before assigning.

Affected scripts:

- [1455577933_MedievalMadness_Upgrade(Real_Final).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/1455577933_MedievalMadness_Upgrade%28Real_Final%29/1455577933_MedievalMadness_Upgrade%28Real_Final%29.vbs#L2185)
- [Bram Stokers Dracula (1993).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Bram%20Stoker%27s%20Dracula%20%28Williams%201993%29/Bram%20Stokers%20Dracula%20%281993%29.vbs#L161)
- [Cirqus_Voltaire_Hanibal-Mod_3.7.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Cirqus_Voltaire_Hanibal-Mod_3.7/Cirqus_Voltaire_Hanibal-Mod_3.7.vbs#L1282)
- [VP10-Terminator 3 (Stern 2003) Hanibal v1.5.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Terminator%203%20%28Stern%202003%29%20Hanibal%20v1.5/VP10-Terminator%203%20%28Stern%202003%29%20Hanibal%20v1.5.vbs#L1236)
- [Vampirella.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Vampirella/Vampirella.vbs#L1)
- [Vampirella1 .3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Vampirella1%20.3/Vampirella1%20.3.vbs#L84)

## Nested Sub definition extracted to top level

**Affected patches: 6**

VBScript does not allow a `Sub` to be defined inside another `Sub`. Scripts that contained nested Sub definitions (e.g. `chilloutthemusic` or `turnitbackup` nested inside the PuPlayer initialisation block) fail to compile. Fix: move the inner Sub to the top level of the script.

Affected scripts:

- [1679379865_jurassicparklimitededition.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/1679379865_jurassicparklimitededition/1679379865_jurassicparklimitededition.vbs#L493)
- [Stranger Things (Original 2020) LW.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Stranger%20Things%20%28Original%202020%29%20LW%202.0.1/Stranger%20Things%20%28Original%202020%29%20LW.vbs#L361)
- [Stranger Things - SE 1.47_OSB.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Stranger%20Things%20-%20SE%201.47_OSB/Stranger%20Things%20-%20SE%201.47_OSB.vbs#L358)
- [Stranger Things 4 Premium.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Stranger%20Things%204%20Premium/Stranger%20Things%204%20Premium.vbs#L84)
- [The Beatles_007.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/The%20Beatles_007/The%20Beatles_007.vbs#L196)
- [tlk-0.35.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/tlk-0.35/tlk-0.35.vbs#L507)

## gBOT replaced with getballs() function call

**Affected patches: 6**

The `gBOT` global variable holds the current ball collection in older VPX scripts. In newer versions it is replaced by the `getballs()` function. Scripts using `gBOT` are updated to call `getballs()` instead.

Affected scripts:

- [Atlantis (Bally 1989) w VR Room v2.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0/Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0.vbs#L3029)
- [Black Knight 2000 (Williams 1989) w VR Room v2.0.2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2/Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2.vbs#L2687)
- [Cactus Jacks (Gottlieb 1991) w VR Room v2.0.2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2/Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2.vbs#L1544)
- [Iron Man Vault Edition (Stern 2010) VPW v1.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Iron%20Man%20Vault%20Edition%20%28Stern%202010%29%20VPW%20v1.0/Iron%20Man%20Vault%20Edition%20%28Stern%202010%29%20VPW%20v1.0.vbs#L3280)
- [Laser War (Data East 1987) w VR Room v2.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0/Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0.vbs#L1409)
- [X-Men LE (Stern 2012) VPW v1.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/X-Men%20LE%20%28Stern%202012%29%20VPW%20v1.0/X-Men%20LE%20%28Stern%202012%29%20VPW%20v1.0.vbs#L3214)

## File path case sensitivity fix (Linux filesystem)

**Affected patches: 6**

Linux file systems are case-sensitive, unlike Windows NTFS. File names embedded in scripts (images, videos, ROM file names) must match the actual file names exactly. Fix: correct the capitalisation of the file name string.

Affected scripts:

- [Beavis and Butt-head_Pinballed.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Beavis%20and%20Butt-head_Pinballed/Beavis%20and%20Butt-head_Pinballed.vbs#L51)
- [Blood Machines 2.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Blood%20Machines%202.0/Blood%20Machines%202.0.vbs#L205)
- [DarkPrincess1.3.1.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/DarkPrincess1.3.1/DarkPrincess1.3.1.vbs#L161)
- [StarWars BountyHunter 3.02.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/StarWars%20BountyHunter%203.02/StarWars%20BountyHunter%203.02.vbs#L55)
- [Thunderbirds original 2022 v1.0.2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Thunderbirds%20original%202022%20v1.0.2/Thunderbirds%20original%202022%20v1.0.2.vbs#L115)
- [UT99CTF_GE_2.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/UT99CTF_GE_2.3/UT99CTF_GE_2.3.vbs#L1203)

## Incorrect boolean expression (IsGIOn <> Not IsOff)

**Affected patches: 5**

`isGIOn <> Not IsOff` is logically incorrect — `Not IsOff` negates the variable, producing a boolean, and `<>` then compares it to `isGIOn`. Under Wine this expression evaluates incorrectly. Fix: use a simple assignment or comparison without the erroneous `<>` operator.

Affected scripts:

- [AceOfSpeed.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/AceOfSpeed/AceOfSpeed.vbs#L245)
- [Death Proof Balutito V2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Death%20Proof%20Balutito%20V2/Death%20Proof%20Balutito%20V2.vbs#L207)
- [Iron Maiden Virtual Time (Original 2020).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Iron%20Maiden%20Virtual%20Time%20%28Original%202020%29/Iron%20Maiden%20Virtual%20Time%20%28Original%202020%29.vbs#L408)
- [Laser War (Data East 1987) w VR Room v2.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0/Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0.vbs#L1409)
- [Mousin' Around! (Bally 1989) w VR Room v2.1.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Mousin%27%20Around%21%20%28Bally%201989%29%20w%20VR%20Room%20v2.1/Mousin%27%20Around%21%20%28Bally%201989%29%20w%20VR%20Room%20v2.1.vbs#L276)

## Configuration constant value changed for standalone compatibility

**Affected patches: 5**

Certain constants (e.g. `cController`, `PupScreenMiniGame`, `Mute_Sound_For_PuPPack`) have default values that assume a full Windows VPX environment. For standalone builds these values need to be adjusted — for example switching from VPinMAME controller to B2S server, or disabling features not available in standalone.

Affected scripts:

- [Bart VS the Space Mutants 1.1.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Bart%20VS%20the%20Space%20Mutants%201.1/Bart%20VS%20the%20Space%20Mutants%201.1.vbs#L33)
- [Centaur (Bally 1981).vbs.dmd](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Centaur%20%28Bally%201981%29/Centaur%20%28Bally%201981%29.vbs.dmd#L6)
- [Dream Daddy1.5.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Dream%20Daddy1.5/Dream%20Daddy1.5.vbs#L69)
- [Heavy Metal Meltdown (Bally 1987).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Heavy%20Metal%20Meltdown%20%28Bally%201987%29/Heavy%20Metal%20Meltdown%20%28Bally%201987%29.vbs#L12)
- [Megadeth (original).vbs.dmd](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Megadeth%20%28original%29/Megadeth%20%28original%29.vbs.dmd#L66)

## Sub/Function definition with inline code (no newline after definition)

**Affected patches: 5**

Placing code on the same line as the `Sub` or `Function` header (e.g. `Sub Sw18_Hit()Controller.Switch(18)=1`) is valid on Windows VBScript but fails to parse correctly under Wine. Fix: add a newline after the `()` so the body starts on the next line.

Affected scripts:

- [Bugs Bunny's Birthday Ball (Bally 1991)Rev2.3b.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Bugs%20Bunny%27s%20Birthday%20Ball%20%28Bally%201991%29Rev2.3b/Bugs%20Bunny%27s%20Birthday%20Ball%20%28Bally%201991%29Rev2.3b.vbs#L798)
- [Cue Ball Wizard (Gottlieb 1992) v1.1.2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Cue%20Ball%20Wizard%20%28Gottlieb%201992%29%20v1.1.2/Cue%20Ball%20Wizard%20%28Gottlieb%201992%29%20v1.1.2.vbs#L1503)
- [Haunted House (Gottlieb 1982) 1.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Haunted%20House%20%28Gottlieb%201982%29%201.0/Haunted%20House%20%28Gottlieb%201982%29%201.0.vbs#L2980)
- [Haunted House (Gottlieb 1982)_Bigus(MOD)1.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Haunted%20House%20%28Gottlieb%201982%29_Bigus%28MOD%291.0/Haunted%20House%20%28Gottlieb%201982%29_Bigus%28MOD%291.0.vbs#L2970)
- [Scared Stiff (Bally 1996) v1.29.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Scared%20Stiff%20%28Bally%201996%29%20v1.29/Scared%20Stiff%20%28Bally%201996%29%20v1.29.vbs#L466)

## For Each loop variable reused (causes error in Wine)

**Affected patches: 4**

Reusing the same loop variable in nested or back-to-back `For Each` loops causes a runtime error in Wine's VBScript. Fix: use a different variable name for each loop level.

Affected scripts:

- [Atlantis (Bally 1989) w VR Room v2.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0/Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0.vbs#L3029)
- [Black Knight 2000 (Williams 1989) w VR Room v2.0.2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2/Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2.vbs#L2687)
- [Cactus Jacks (Gottlieb 1991) w VR Room v2.0.2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2/Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2.vbs#L1544)
- [Laser War (Data East 1987) w VR Room v2.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0/Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0.vbs#L1409)

## UseVPMDMD variable renamed to UseVPMColoredDMD

**Affected patches: 4**

Scripts that use a colored DMD must use the variable name `UseVPMColoredDMD` (not `UseVPMDMD`) so that the VPX standalone renderer applies the correct colour palette.

Affected scripts:

- [Bad Cats (Williams 1989) VPW 3.0.vbs.dmd](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Bad%20Cats%20%28Williams%201989%29%20VPW%203.0/Bad%20Cats%20%28Williams%201989%29%20VPW%203.0.vbs.dmdcolored#L384)
- Champion Pub (Williams 1998) v1.43.vbs
- Cue Ball Wizard (Gottlieb 1992) v1.2.4.vbs
- Tron Legacy (Stern 2011) VPW Mod v1.1.vbs

## typename() casing fix (typename → TypeName)

**Affected patches: 4**

Wine's VBScript is stricter about built-in function capitalisation. `typename(x)` must be `TypeName(x)`.

Affected scripts:

- [Elvis_MOD_2.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Elvis_MOD_2.0/Elvis_MOD_2.0.vbs#L626)
- [Simpsons Pinball Party, The (Stern 2003) VPW 2.0.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3/Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3.vbs#L3332)
- [Simpsons Treehouse of Horror MOD v2023.3 (Simpsons Pinball Party, The (Stern 2003) VPW 2.0.3 base).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Simpsons%20Treehouse%20of%20Horror%20MOD%20v2023.3%20%28Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3%20base%29/Simpsons%20Treehouse%20of%20Horror%20MOD%20v2023.3%20%28Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3%20base%29.vbs#L95)
- [TX-Sector (Gottlieb 1988) SG1bsoN Mod V1.1.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/TX-Sector%20%28Gottlieb%201988%29%20SG1bsoN%20Mod%20V1.1/TX-Sector%20%28Gottlieb%201988%29%20SG1bsoN%20Mod%20V1.1.vbs#L809)

## Statement starting with colon (invalid VBScript syntax)

**Affected patches: 4**

Lines that begin with `:` followed by a statement (e.g. `        : Controller.B2SSetData 9, 0`) are not valid VBScript syntax — the colon is a statement separator, not a line prefix. Fix: remove the leading colon.

Affected scripts:

- [Junkyard Cats_1.07.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Junkyard%20Cats_1.07/Junkyard%20Cats_1.07.vbs#L508)
- [NBA (Stern 2009)_Bigus(MOD)1.4.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/NBA%20%28Stern%202009%29_Bigus%28MOD%291.4/NBA%20%28Stern%202009%29_Bigus%28MOD%291.4.vbs#L279)
- [Scarface - Balls and Power 1.2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Scarface%20-%20Balls%20and%20Power%201.2/Scarface%20-%20Balls%20and%20Power%201.2.vbs#L1000)
- [The Fifth Element 1.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/The%20Fifth%20Element%201.3/The%20Fifth%20Element%201.3.vbs#L478)

## UBound casing fix (Ubound → UBound)

**Affected patches: 3**

Wine's VBScript is stricter about built-in function capitalisation. `Ubound` must be `UBound`.

Affected scripts:

- [Monkey Island VR Room.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Monkey%20Island%20VR%20Room/Monkey%20Island%20VR%20Room.vbs#L240)
- [Power Play (Bally 1977).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/1256692067_PowerPlay%28Bally1977%292.1/Power%20Play%20%28Bally%201977%29.vbs#L922)
- [monkeyisland.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Monkeyislandv1.1/monkeyisland.vbs#L2441)

## Double-dot in string callback (e.g. "obj..Method")

**Affected patches: 3**

SolCallback strings like `"dtbank..SolHit"` contain a double dot, which is invalid. The double dot arises when the object name is accidentally repeated. Fix: correct to a single dot, e.g. `"dtbank.SolHit"`.

Affected scripts:

- [ABBAv2.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/ABBAv2.0/ABBAv2.0.vbs#L47)
- [Alfred Hitchcock's Psycho (TBA 2019).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Alfred%20Hitchcock%27s%20Psycho%20%28TBA%202019%29/Alfred%20Hitchcock%27s%20Psycho%20%28TBA%202019%29.vbs#L292)
- [Vortex (Taito do Brasil - 1981) v4.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4/Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4.vbs#L46)

## Hex literal with excess leading zeros (e.g. &H000000031)

**Affected patches: 3**

A hex literal with an odd or excessive number of digits (e.g. `&H000000031` — nine hex digits) is rejected by Wine's VBScript parser. Fix: reduce to the correct even number of digits (e.g. `&H00000031`).

Affected scripts:

- [Andromeda (Game Plan 1985) v4.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Andromeda%20%28Game%20Plan%201985%29%20v4/Andromeda%20%28Game%20Plan%201985%29%20v4.vbs#L922)
- [Grand Slam (Bally 1983) 2.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Grand%20Slam%20%28Bally%201983%29%202.3/Grand%20Slam%20%28Bally%201983%29%202.3.vbs#L791)
- [Topaz (Inder 1979).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Topaz%20%28Inder%201979%29/Topaz%20%28Inder%201979%29.vbs#L100)

## Me(Idx) collection indexer replaced with named collection

**Affected patches: 3**

Inside a VBScript `Class`, `Me(n)` is not a valid way to index a collection under Wine. Fix: use the named array/collection directly, e.g. `Spinner(Idx).IsDropped`.

Affected scripts:

- [Four Seasons (Gottlieb 1968).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Four%20Seasons%20%28Gottlieb%201968%29/Four%20Seasons%20%28Gottlieb%201968%29.vbs#L1214)
- [Four Seasons (Gottlieb 1968)_Teisen_MOD.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Four%20Seasons%20%28Gottlieb%201968%29_Teisen_MOD/Four%20Seasons%20%28Gottlieb%201968%29_Teisen_MOD.vbs#L1219)
- [Roller Coaster (Gottlieb 1971)x.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Roller%20Coaster%20%28Gottlieb%201971%29x/Roller%20Coaster%20%28Gottlieb%201971%29x.vbs#L1037)

## Trim.visible assignment removed (Trim is a VBScript reserved function)

**Affected patches: 3**

`Trim` is a built-in VBScript string function. Using it as an object name (e.g. `Trim.visible = 1`) is interpreted as a function call and raises a type error. Fix: comment out or rename the object.

Affected scripts:

- [Gladiators (Premier 1993) v1.1.1.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Gladiators%20%28Premier%201993%29%20v1.1.1/Gladiators%20%28Premier%201993%29%20v1.1.1.vbs#L344)
- [Surf'n Safari v1.3.4.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Surf%27n%20Safari%20v1.3.4/Surf%27n%20Safari%20v1.3.4.vbs#L117)
- [Wipe Out (Premier 1993) 1.1.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Wipe%20Out%20%28Premier%201993%29%201.1.0/Wipe%20Out%20%28Premier%201993%29%201.1.0.vbs#L258)

## One-liner If/Then with dangling Else (missing End If)

**Affected patches: 3**

A single-line `If x Then y: Else` with nothing after the `Else` is syntactically invalid — VBScript expects an `End If` or an inline statement after `Else`. Fix: remove the trailing `: Else` or add `End If`.

Affected scripts:

- [South Park (Sega 1999) 1.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/South%20Park%20%28Sega%201999%29%201.3/South%20Park%20%28Sega%201999%29%201.3.vbs#L1167)
- [South Park - Halloween.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/South%20Park%20-%20Halloween/South%20Park%20-%20Halloween.vbs#L49)
- [Starship Troopers (Sega 1997) VPW Mod v2.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Starship%20Troopers%20%28Sega%201997%29%20VPW%20Mod%20v2.0/Starship%20Troopers%20%28Sega%201997%29%20VPW%20Mod%20v2.0.vbs#L1324)

## Duplicate score function calls removed

**Affected patches: 2**

The same `AddScore` call was duplicated multiple times in the same event handler, causing the score to be added multiple times per activation. Fix: keep only one call.

Affected scripts:

- [2104398928_CARtoonsRC(Nailed2021)v1.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/2104398928_CARtoonsRC%28Nailed2021%29v1.3/2104398928_CARtoonsRC%28Nailed2021%29v1.3.vbs#L841)
- [Gemini (Gottlieb 1978).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Gemini%20%28Gottlieb%201978%29/Gemini%20%28Gottlieb%201978%29.vbs#L1173)

## 'default' reserved word conflict fixed (Const default = 0 added)

**Affected patches: 2**

Using `default` as a variable name conflicts with VBScript's reserved `Default` property keyword. Fix: declare `Const default = 0` at the top of the script to shadow the reserved word with a harmless constant.

Affected scripts:

- [Bud and Terence (Original 2024) v1.6.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Bud%20Spencer%20%26%20Terence%20Hill%20%28Original%202024%29/Bud%20and%20Terence%20%28Original%202024%29%20v1.6.vbs#L4386)
- [IT Pinball Madness (JP Salas,Joe Picasso)1.2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/IT%20Pinball%20Madness%20%28JP%20Salas%2CJoe%20Picasso%291.2/IT%20Pinball%20Madness%20%28JP%20Salas%2CJoe%20Picasso%291.2.vbs#L4870)

## BSize/BMass constant renamed to BallSize/BallMass

**Affected patches: 2**

`Const BSize` or `Const BMass` conflict with VPX's built-in `BallSize` and `BallMass` properties. Fix: rename the script constants to avoid the conflict.

Affected scripts:

- [Cyclone (Williams 1988) 2.0.1.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Cyclone%20%28Williams%201988%29%202.0.1/Cyclone%20%28Williams%201988%29%202.0.1.vbs#L5)
- [Space Jam (Sega 1996) 1.4 JPJ - Edizzle - TeamPP - JLou.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Space%20Jam%20%28Sega%201996%29%201.4%20JPJ%20-%20Edizzle%20-%20TeamPP%20-%20JLou/Space%20Jam%20%28Sega%201996%29%201.4%20JPJ%20-%20Edizzle%20-%20TeamPP%20-%20JLou.vbs#L7)

## OptionsLoad/OptionsEdit call removed (not available in standalone)

**Affected patches: 2**

`OptionsLoad` and `OptionsEdit` are helper functions provided by some VPX utilities but not available in standalone builds. Calling them causes a "Sub or Function not defined" error. Fix: remove the calls and inline the required initialisation.

Affected scripts:

- [Gremlins by Balutito 1.7.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Gremlins%20by%20Balutito%201.7/Gremlins%20by%20Balutito%201.7.vbs#L162)
- [Victory (Gottlieb 1987) 2.0.2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Victory%20%28Gottlieb%201987%29%202.0.2/Victory%20%28Gottlieb%201987%29%202.0.2.vbs#L179)

## VPinMAME Settings.Value() call fix

**Affected patches: 2**

`Controller.Settings.Value("key")` must be called without extra parentheses around the argument in VBScript when the return value is discarded. Fix: use `Controller.Settings.Value "key"` or restructure the call.

Affected scripts:

- [Phantom Of The Opera (Data East 1990) v1.24.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Phantom%20Of%20The%20Opera%20%28Data%20East%201990%29%20v1.24/Phantom%20Of%20The%20Opera%20%28Data%20East%201990%29%20v1.24.vbs#L316)
- [The Phantom Of The Opera (Data East 1990).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/The%20Phantom%20Of%20The%20Opera%20%28Data%20East%201990%29/The%20Phantom%20Of%20The%20Opera%20%28Data%20East%201990%29.vbs#L135)

## FlexDMD configuration added for standalone DMD display

**Affected patches: 2**

Tables using FlexDMD for their DMD require a `Const FlexDMDHighQuality` declaration and related setup code to display correctly in standalone mode. These lines are added when missing.

Affected scripts:

- The Dark Crystal (Original 2020) 1.2.vbs
- The Neverending Story (Original 2021) v1.vbs

## debug.print call removed or commented out

**Affected patches: 2**

`Debug.Print` outputs to the VBScript debug console, which is not available in standalone builds and can cause runtime errors. Fix: comment out or remove `Debug.Print` statements.

Affected scripts:

- [TimonPumbaa.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/TimonPumbaa/TimonPumbaa.vbs#L3548)
- [Wheel of Fortune (Stern 2007) 1.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Wheel%20of%20Fortune%20%28Stern%202007%29%201.0/Wheel%20of%20Fortune%20%28Stern%202007%29%201.0.vbs#L10)

## CDbl() replaced with CBool() (wrong type conversion)

**Affected patches: 1**

Affected scripts:

- [A-Go-Go (Williams 1966).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/A-Go-Go%20%28Williams%201966%29/A-Go-Go%20%28Williams%201966%29.vbs#L1293)

## DisplayTimer.Enabled = DesktopMode fix

**Affected patches: 1**

Setting `DisplayTimer.Enabled = DesktopMode` assigns a boolean to a timer's Enabled property. In VR mode `DesktopMode` is `False`, inadvertently disabling the timer. Fix: wrap in a conditional or evaluate correctly.

Affected scripts:

- [AceOfSpeed.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/AceOfSpeed/AceOfSpeed.vbs#L245)

## cvpmDictionary replaced with Scripting.Dictionary

**Affected patches: 1**

`cvpmDictionary` is an older VPM wrapper class. Scripts using `New cvpmDictionary` are updated to use `CreateObject("Scripting.Dictionary")` directly.

Affected scripts:

- [Batman66_1.1.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Batman66_1.1.0/Batman66_1.1.0.vbs#L93)

## LinkedTo property must be assigned an Array (not a single object)

**Affected patches: 1**

The `DropTarget.LinkedTo` property expects an Array of linked targets, not a bare object reference. Fix: wrap the value in `Array(...)`, e.g. `dtR.LinkedTo = Array(dtL)`.

Affected scripts:

- [Contact (Williams 1978)_Bigus(MOD)1.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Contact%20%28Williams%201978%29_Bigus%28MOD%291.0/Contact%20%28Williams%201978%29_Bigus%28MOD%291.0.vbs#L170)

## Const tnob value fix

**Affected patches: 1**

The `tnob` (total number of balls) constant had an incorrect value, causing ball tracking issues. Fix: set `Const tnob` to the correct number of simultaneous balls for the table.

Affected scripts:

- [GalaxyPlay_1_2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/GalaxyPlay_1_2/GalaxyPlay_1_2.vbs#L154)

## Code missing newlines between statements (all on one line)

**Affected patches: 1**

Multiple statements were concatenated onto a single line without newlines, making the script extremely long and causing parse errors in Wine. Fix: insert newlines to separate each statement.

Affected scripts:

- [Iron Maiden (Original 2022) VPW 1.0.12.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Iron%20Maiden%20%28Original%202022%29%20VPW%201.0.12/Iron%20Maiden%20%28Original%202022%29%20VPW%201.0.12.vbs#L7957)

## b2s.vbs GetTextFile replaced with inline B2S helper functions

**Affected patches: 1**

`ExecuteGlobal GetTextFile("b2s.vbs")` loads a shared B2S helper script at runtime. This approach fails in standalone because the helper file is not in the expected path. Fix: inline the needed B2S helper functions (`SetB2SData`, `StepB2SData`, etc.) directly into the script.

Affected scripts:

- [Jungle_Queen.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Jungle_Queen/Jungle_Queen.vbs#L15)

## Variable value or array size changed for standalone compatibility

**Affected patches: 1**

A variable value or array dimension is changed to a value more appropriate for standalone play — for example, increasing a `LampState` array size to accommodate all lamps, or adjusting a physics limit value.

Affected scripts:

- [Junk Yard (Williams 1996) v1.83.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Junk%20Yard%20%28Williams%201996%29%20v1.83/Junk%20Yard%20%28Williams%201996%29%20v1.83.vbs#L887)

## Controller.Pause missing assignment (should be = True/False)

**Affected patches: 1**

`Controller.Pause` is a property, not a method. Writing `Controller.Pause` without an assignment (`= True` or `= False`) does nothing useful and can cause a syntax error. Fix: use `Controller.Pause = True` / `Controller.Pause = False`.

Affected scripts:

- [Legend of Zelda v4.3.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Legend%20of%20Zelda%20v4.3/Legend%20of%20Zelda%20v4.3.vbs#L404)

## Duplicate InitPolarity call removed (second call crashes)

**Affected patches: 1**

Calling `InitPolarity` twice causes a crash. A duplicate call was present and is removed.

Affected scripts:

- [Party Animal (Bally 1987).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Party%20Animal%20%28Bally%201987%29/Party%20Animal%20%28Bally%201987%29.vbs#L223)

## .AddBall.0 invalid syntax fixed to .AddBall 0

**Affected patches: 1**

`bsTrough.AddBall.0` attempts to access property `0` of the return value of `AddBall`, which is invalid. Fix: use `bsTrough.AddBall 0` (passing `0` as an argument).

Affected scripts:

- [PinballMagic.v1.9.Hybrid.VPX8.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/PinballMagic.v1.9.Hybrid.VPX8/PinballMagic.v1.9.Hybrid.VPX8.vbs#L1448)

## SolCallback assignment commented out (callback not available in standalone)

**Affected patches: 1**

A `SolCallback` entry pointing to a helper function that does not exist in standalone is commented out to prevent a "Sub or Function not defined" error at runtime.

Affected scripts:

- [Playboy (Bally 1978).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Playboy%20%28Bally%201978%29/Playboy%20%28Bally%201978%29.vbs#L36)

## vpmShowDips call removed (not available in standalone)

**Affected patches: 1**

`vpmShowDips` displays the ROM DIP switch settings dialog. This function is not available in standalone builds. Fix: remove or comment out the call.

Affected scripts:

- [Playboy (Stern 2002) v1.1.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Playboy%28Stern2002%29v1.1/Playboy%20%28Stern%202002%29%20v1.1.vbs#L652)

## PlaySoundAt() dot instead of comma separating arguments

**Affected patches: 1**

`PlaySoundAt "name". Object` uses a dot instead of a comma between the sound name and the position object. Fix: replace the dot with a comma: `PlaySoundAt "name", Object`.

Affected scripts:

- [Ramones (HauntFreaks 2021).vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Ramones%20%28HauntFreaks%202021%29/Ramones%20%28HauntFreaks%202021%29.vbs#L802)

## Eval()/Dictionary item double-indexed result (Wine limitation)

**Affected patches: 1**

`Eval("name")(0)(step)` or `Dict.Item(key)(0)` chains two index operations on the returned object. Wine does not support chained indexing of Eval/Dictionary results. Fix: store the result in a temporary variable first, then index it.

Affected scripts:

- [Saving Wallden.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Saving%20Wallden/Saving%20Wallden.vbs#L9387)

## Object.state = off changed to = 0 (off evaluates to True in VBScript)

**Affected patches: 1**

In VBScript, the identifier `off` is not defined and evaluates to `Empty`, which coerces to `0` — but if `off` has been previously set to `True` in the script, assigning `state = off` sets state to `True` (on), not `False` (off). Fix: use the literal `0` instead.

Affected scripts:

- [SpaceRamp (SuperEd) v3.03b.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/SpaceRamp%20%28SuperEd%29%20v3.03b/SpaceRamp%20%28SuperEd%29%20v3.03b.vbs#L612)

## Constant renamed to avoid conflict or improve clarity

**Affected patches: 1**

A constant name conflicted with a built-in function or property name, or was renamed for clarity (e.g. `Const Offset` → `Const DigitsOffset` to avoid shadowing `Offset` elsewhere).

Affected scripts:

- [Taxi (Williams 1988) VPW v1.2.2.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/Taxi%20%28Williams%201988%29%20VPW%20v1.2.2/Taxi%20%28Williams%201988%29%20VPW%20v1.2.2.vbs#L1435)

## ActiveBall.id replaced with ActiveBall.Uservalue

**Affected patches: 1**

The `ActiveBall.id` property is not available in VPX Standalone. Scripts that use `.id` to track specific balls are updated to use `ActiveBall.Uservalue` instead.

Affected scripts:

- [The Rolling Stones (Stern 2011)_Bigus(MOD)3.0.vbs](https://github.com/francisdb/vpx-standalone-scripts/blob/master/The%20Rolling%20Stones%20%28Stern%202011%29_Bigus%28MOD%293.0/The%20Rolling%20Stones%20%28Stern%202011%29_Bigus%28MOD%293.0.vbs#L268)

