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

- 1342729923_RollerCoasterTycoon(Stern2002)1.3.vbs [[original]](1342729923_RollerCoasterTycoon%28Stern2002%291.3/1342729923_RollerCoasterTycoon%28Stern2002%291.3.vbs.original#L3168) [[patch]](1342729923_RollerCoasterTycoon%28Stern2002%291.3/1342729923_RollerCoasterTycoon%28Stern2002%291.3.vbs.patch) [[patched]](1342729923_RollerCoasterTycoon%28Stern2002%291.3/1342729923_RollerCoasterTycoon%28Stern2002%291.3.vbs)
- 2001 (Gottlieb 1971) v0.99a.vbs [[original]](2001%20%28Gottlieb%201971%29%20v0.99a/2001%20%28Gottlieb%201971%29%20v0.99a.vbs.original#L5025) [[patch]](2001%20%28Gottlieb%201971%29%20v0.99a/2001%20%28Gottlieb%201971%29%20v0.99a.vbs.patch) [[patched]](2001%20%28Gottlieb%201971%29%20v0.99a/2001%20%28Gottlieb%201971%29%20v0.99a.vbs)
- AC-DC LUCI Premium VR (Stern 2013) v1.1.3.vbs [[original]](AC-DC%20LUCI%20Premium%20VR%20%28Stern%202013%29%20v1.1.3/AC-DC%20LUCI%20Premium%20VR%20%28Stern%202013%29%20v1.1.3.vbs.original#L330) [[patch]](AC-DC%20LUCI%20Premium%20VR%20%28Stern%202013%29%20v1.1.3/AC-DC%20LUCI%20Premium%20VR%20%28Stern%202013%29%20v1.1.3.vbs.patch) [[patched]](AC-DC%20LUCI%20Premium%20VR%20%28Stern%202013%29%20v1.1.3/AC-DC%20LUCI%20Premium%20VR%20%28Stern%202013%29%20v1.1.3.vbs)
- Alien Poker (Williams 1980).vbs [[original]](AlienPoker%28Williams1980%291.0/Alien%20Poker%20%28Williams%201980%29.vbs.original#L1935) [[patch]](AlienPoker%28Williams1980%291.0/Alien%20Poker%20%28Williams%201980%29.vbs.patch) [[patched]](AlienPoker%28Williams1980%291.0/Alien%20Poker%20%28Williams%201980%29.vbs)
- Apollo 13 (Sega 1995) w VR Room v2.1.1.vbs [[original]](Apollo%2013%20%28Sega%201995%29%20w%20VR%20Room%20v2.1.1/Apollo%2013%20%28Sega%201995%29%20w%20VR%20Room%20v2.1.1.vbs.original#L2624) [[patch]](Apollo%2013%20%28Sega%201995%29%20w%20VR%20Room%20v2.1.1/Apollo%2013%20%28Sega%201995%29%20w%20VR%20Room%20v2.1.1.vbs.patch) [[patched]](Apollo%2013%20%28Sega%201995%29%20w%20VR%20Room%20v2.1.1/Apollo%2013%20%28Sega%201995%29%20w%20VR%20Room%20v2.1.1.vbs)
- Atlantis (Bally 1989) w VR Room v2.0.vbs [[original]](Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0/Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0.vbs.original#L3029) [[patch]](Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0/Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0.vbs.patch) [[patched]](Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0/Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0.vbs)
- Barracora (Williams 1981) w VR Room v2.1.3.vbs [[original]](Barracora%20%28Williams%201981%29%20w%20VR%20Room%20v2.1.3/Barracora%20%28Williams%201981%29%20w%20VR%20Room%20v2.1.3.vbs.original#L2022) [[patch]](Barracora%20%28Williams%201981%29%20w%20VR%20Room%20v2.1.3/Barracora%20%28Williams%201981%29%20w%20VR%20Room%20v2.1.3.vbs.patch) [[patched]](Barracora%20%28Williams%201981%29%20w%20VR%20Room%20v2.1.3/Barracora%20%28Williams%201981%29%20w%20VR%20Room%20v2.1.3.vbs)
- Batman Forever (Sega 1995) 1.3.vbs [[original]](Batman%20Forever%20%28Sega%201995%29%201.3/Batman%20Forever%20%28Sega%201995%29%201.3.vbs.original#L3518) [[patch]](Batman%20Forever%20%28Sega%201995%29%201.3/Batman%20Forever%20%28Sega%201995%29%201.3.vbs.patch) [[patched]](Batman%20Forever%20%28Sega%201995%29%201.3/Batman%20Forever%20%28Sega%201995%29%201.3.vbs)
- Batman [The Dark Knight] (Stern 2008) 1.16.vbs [[original]](Batman%20%5BThe%20Dark%20Knight%5D%20%28Stern%202008%29%201.16/Batman%20%5BThe%20Dark%20Knight%5D%20%28Stern%202008%29%201.16.vbs.original#L905) [[patch]](Batman%20%5BThe%20Dark%20Knight%5D%20%28Stern%202008%29%201.16/Batman%20%5BThe%20Dark%20Knight%5D%20%28Stern%202008%29%201.16.vbs.patch) [[patched]](Batman%20%5BThe%20Dark%20Knight%5D%20%28Stern%202008%29%201.16/Batman%20%5BThe%20Dark%20Knight%5D%20%28Stern%202008%29%201.16.vbs)
- Batman [The Dark Knight] (Stern 2008) v1.0.8.vbs [[original]](Batman%20%5BThe%20Dark%20Knight%5D%20%28Stern%202008%29%20v1.0.8/Batman%20%5BThe%20Dark%20Knight%5D%20%28Stern%202008%29%20v1.0.8.vbs.original#L824) [[patch]](Batman%20%5BThe%20Dark%20Knight%5D%20%28Stern%202008%29%20v1.0.8/Batman%20%5BThe%20Dark%20Knight%5D%20%28Stern%202008%29%20v1.0.8.vbs.patch) [[patched]](Batman%20%5BThe%20Dark%20Knight%5D%20%28Stern%202008%29%20v1.0.8/Batman%20%5BThe%20Dark%20Knight%5D%20%28Stern%202008%29%20v1.0.8.vbs)
- Batman [The Dark Knight] (Stern 2008).vbs [[original]](BatmanTheDarkKnight%28Stern2008%29bordmod1.0.7/Batman%20%5BThe%20Dark%20Knight%5D%20%28Stern%202008%29.vbs.original#L823) [[patch]](BatmanTheDarkKnight%28Stern2008%29bordmod1.0.7/Batman%20%5BThe%20Dark%20Knight%5D%20%28Stern%202008%29.vbs.patch) [[patched]](BatmanTheDarkKnight%28Stern2008%29bordmod1.0.7/Batman%20%5BThe%20Dark%20Knight%5D%20%28Stern%202008%29.vbs)
- BeastieBoysv1.0.vbs [[original]](BeastieBoysv1.0/BeastieBoysv1.0.vbs.original#L100) [[patch]](BeastieBoysv1.0/BeastieBoysv1.0.vbs.patch) [[patched]](BeastieBoysv1.0/BeastieBoysv1.0.vbs)
- Black Knight (Williams 1980).vbs [[original]](Black%20Knight%20%28Williams%201980%29/Black%20Knight%20%28Williams%201980%29.vbs.original#L896) [[patch]](Black%20Knight%20%28Williams%201980%29/Black%20Knight%20%28Williams%201980%29.vbs.patch) [[patched]](Black%20Knight%20%28Williams%201980%29/Black%20Knight%20%28Williams%201980%29.vbs)
- Black Knight 2000 (Williams 1989) w VR Room v2.0.2.vbs [[original]](Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2/Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2.vbs.original#L2687) [[patch]](Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2/Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2.vbs.patch) [[patched]](Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2/Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2.vbs)
- Blackout (Williams 1980) v1.0.1 - SBR34.vbs [[original]](Blackout%20%28Williams%201980%29%20v1.0.1%20-%20SBR34/Blackout%20%28Williams%201980%29%20v1.0.1%20-%20SBR34.vbs.original#L1888) [[patch]](Blackout%20%28Williams%201980%29%20v1.0.1%20-%20SBR34/Blackout%20%28Williams%201980%29%20v1.0.1%20-%20SBR34.vbs.patch) [[patched]](Blackout%20%28Williams%201980%29%20v1.0.1%20-%20SBR34/Blackout%20%28Williams%201980%29%20v1.0.1%20-%20SBR34.vbs)
- Blood Machines 2.0.vbs [[original]](Blood%20Machines%202.0/Blood%20Machines%202.0.vbs.original#L205) [[patch]](Blood%20Machines%202.0/Blood%20Machines%202.0.vbs.patch) [[patched]](Blood%20Machines%202.0/Blood%20Machines%202.0.vbs)
- Bounty Hunter (Gottlieb 1985).vbs [[original]](Bounty%20Hunter%20%28Gottlieb%201985%29/Bounty%20Hunter%20%28Gottlieb%201985%29.vbs.original#L1420) [[patch]](Bounty%20Hunter%20%28Gottlieb%201985%29/Bounty%20Hunter%20%28Gottlieb%201985%29.vbs.patch) [[patched]](Bounty%20Hunter%20%28Gottlieb%201985%29/Bounty%20Hunter%20%28Gottlieb%201985%29.vbs)
- Cactus Canyon (Bally 1998) VPW 1.1.vbs [[original]](Cactus%20Canyon%20%28Bally%201998%29%20VPW%201.1/Cactus%20Canyon%20%28Bally%201998%29%20VPW%201.1.vbs.original#L108) [[patch]](Cactus%20Canyon%20%28Bally%201998%29%20VPW%201.1/Cactus%20Canyon%20%28Bally%201998%29%20VPW%201.1.vbs.patch) [[patched]](Cactus%20Canyon%20%28Bally%201998%29%20VPW%201.1/Cactus%20Canyon%20%28Bally%201998%29%20VPW%201.1.vbs)
- Cactus Jacks (Gottlieb 1991) w VR Room v2.0.2.vbs [[original]](Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2/Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2.vbs.original#L1544) [[patch]](Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2/Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2.vbs.patch) [[patched]](Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2/Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2.vbs)
- Capt. Fantastic and The Brown Dirt Cowboy (Bally 1976) 2.0.2.vbs [[original]](Capt.%20Fantastic%20and%20The%20Brown%20Dirt%20Cowboy%20%28Bally%201976%29%202.0.2/Capt.%20Fantastic%20and%20The%20Brown%20Dirt%20Cowboy%20%28Bally%201976%29%202.0.2.vbs.original#L3871) [[patch]](Capt.%20Fantastic%20and%20The%20Brown%20Dirt%20Cowboy%20%28Bally%201976%29%202.0.2/Capt.%20Fantastic%20and%20The%20Brown%20Dirt%20Cowboy%20%28Bally%201976%29%202.0.2.vbs.patch) [[patched]](Capt.%20Fantastic%20and%20The%20Brown%20Dirt%20Cowboy%20%28Bally%201976%29%202.0.2/Capt.%20Fantastic%20and%20The%20Brown%20Dirt%20Cowboy%20%28Bally%201976%29%202.0.2.vbs)
- Centaur (Bally 1981).vbs [[original]](Centaur%20%28Bally%201981%29/Centaur%20%28Bally%201981%29.vbs.original#L2069) [[patch]](Centaur%20%28Bally%201981%29/Centaur%20%28Bally%201981%29.vbs.patch) [[patched]](Centaur%20%28Bally%201981%29/Centaur%20%28Bally%201981%29.vbs)
- Centigrade 37 (Gottlieb 1977).vbs [[original]](Centigrade%2037%20%28Gottlieb%201977%29/Centigrade%2037%20%28Gottlieb%201977%29.vbs.original#L3290) [[patch]](Centigrade%2037%20%28Gottlieb%201977%29/Centigrade%2037%20%28Gottlieb%201977%29.vbs.patch) [[patched]](Centigrade%2037%20%28Gottlieb%201977%29/Centigrade%2037%20%28Gottlieb%201977%29.vbs)
- Checkpoint (Data East 1991)2.0.vbs [[original]](Checkpoint%20%28Data%20East%201991%292.0/Checkpoint%20%28Data%20East%201991%292.0.vbs.original#L2301) [[patch]](Checkpoint%20%28Data%20East%201991%292.0/Checkpoint%20%28Data%20East%201991%292.0.vbs.patch) [[patched]](Checkpoint%20%28Data%20East%201991%292.0/Checkpoint%20%28Data%20East%201991%292.0.vbs)
- Comet (Williams 1985) w VR Room v2.3.vbs [[original]](Comet%20%28Williams%201985%29%20w%20VR%20Room%20v2.3/Comet%20%28Williams%201985%29%20w%20VR%20Room%20v2.3.vbs.original#L2233) [[patch]](Comet%20%28Williams%201985%29%20w%20VR%20Room%20v2.3/Comet%20%28Williams%201985%29%20w%20VR%20Room%20v2.3.vbs.patch) [[patched]](Comet%20%28Williams%201985%29%20w%20VR%20Room%20v2.3/Comet%20%28Williams%201985%29%20w%20VR%20Room%20v2.3.vbs)
- Cue Ball Wizard (Gottlieb 1992) v1.1.2.vbs [[original]](Cue%20Ball%20Wizard%20%28Gottlieb%201992%29%20v1.1.2/Cue%20Ball%20Wizard%20%28Gottlieb%201992%29%20v1.1.2.vbs.original#L1503) [[patch]](Cue%20Ball%20Wizard%20%28Gottlieb%201992%29%20v1.1.2/Cue%20Ball%20Wizard%20%28Gottlieb%201992%29%20v1.1.2.vbs.patch) [[patched]](Cue%20Ball%20Wizard%20%28Gottlieb%201992%29%20v1.1.2/Cue%20Ball%20Wizard%20%28Gottlieb%201992%29%20v1.1.2.vbs)
- Diner (Williams 1990) VPW Mod 1.0.2.vbs [[original]](Diner%20%28Williams%201990%29%20VPW%20Mod%201.0.2/Diner%20%28Williams%201990%29%20VPW%20Mod%201.0.2.vbs.original#L4176) [[patch]](Diner%20%28Williams%201990%29%20VPW%20Mod%201.0.2/Diner%20%28Williams%201990%29%20VPW%20Mod%201.0.2.vbs.patch) [[patched]](Diner%20%28Williams%201990%29%20VPW%20Mod%201.0.2/Diner%20%28Williams%201990%29%20VPW%20Mod%201.0.2.vbs)
- Doctor Who (Bally 1992) VPW Mod v1.1.vbs [[original]](Doctor%20Who%20%28Bally%201992%29%20VPW%20Mod%20v1.1/Doctor%20Who%20%28Bally%201992%29%20VPW%20Mod%20v1.1.vbs.original#L1111) [[patch]](Doctor%20Who%20%28Bally%201992%29%20VPW%20Mod%20v1.1/Doctor%20Who%20%28Bally%201992%29%20VPW%20Mod%20v1.1.vbs.patch) [[patched]](Doctor%20Who%20%28Bally%201992%29%20VPW%20Mod%20v1.1/Doctor%20Who%20%28Bally%201992%29%20VPW%20Mod%20v1.1.vbs)
- Dracula (Stern 1979).vbs [[original]](Dracula%20%28Stern%201979%292.1.1/Dracula%20%28Stern%201979%29.vbs.original#L60) [[patch]](Dracula%20%28Stern%201979%292.1.1/Dracula%20%28Stern%201979%29.vbs.patch) [[patched]](Dracula%20%28Stern%201979%292.1.1/Dracula%20%28Stern%201979%29.vbs)
- Elektra (Bally 1981) w VR Room v2.0.7.vbs [[original]](Elektra%20%28Bally%201981%29%20w%20VR%20Room%20v2.0.7/Elektra%20%28Bally%201981%29%20w%20VR%20Room%20v2.0.7.vbs.original#L2205) [[patch]](Elektra%20%28Bally%201981%29%20w%20VR%20Room%20v2.0.7/Elektra%20%28Bally%201981%29%20w%20VR%20Room%20v2.0.7.vbs.patch) [[patched]](Elektra%20%28Bally%201981%29%20w%20VR%20Room%20v2.0.7/Elektra%20%28Bally%201981%29%20w%20VR%20Room%20v2.0.7.vbs)
- Elvis_MOD_2.0.vbs [[original]](Elvis_MOD_2.0/Elvis_MOD_2.0.vbs.original#L626) [[patch]](Elvis_MOD_2.0/Elvis_MOD_2.0.vbs.patch) [[patched]](Elvis_MOD_2.0/Elvis_MOD_2.0.vbs)
- Family Guy 1.0.vbs [[original]](Family%20Guy%201.0/Family%20Guy%201.0.vbs.original#L3797) [[patch]](Family%20Guy%201.0/Family%20Guy%201.0.vbs.patch) [[patched]](Family%20Guy%201.0/Family%20Guy%201.0.vbs)
- Flash (Williams 1979).vbs [[original]](Flash%20%28Williams%201979%29/Flash%20%28Williams%201979%29.vbs.original#L2124) [[patch]](Flash%20%28Williams%201979%29/Flash%20%28Williams%201979%29.vbs.patch) [[patched]](Flash%20%28Williams%201979%29/Flash%20%28Williams%201979%29.vbs)
- Fog, The (Gottlieb 1979) v2.5 for 10.7.vbs [[original]](Fog%2C%20The%20%28Gottlieb%201979%29%20v2.5%20for%2010.7/Fog%2C%20The%20%28Gottlieb%201979%29%20v2.5%20for%2010.7.vbs.original#L3380) [[patch]](Fog%2C%20The%20%28Gottlieb%201979%29%20v2.5%20for%2010.7/Fog%2C%20The%20%28Gottlieb%201979%29%20v2.5%20for%2010.7.vbs.patch) [[patched]](Fog%2C%20The%20%28Gottlieb%201979%29%20v2.5%20for%2010.7/Fog%2C%20The%20%28Gottlieb%201979%29%20v2.5%20for%2010.7.vbs)
- Galaxy (Stern 1980).vbs [[original]](Galaxy%20%28Stern%201980%29/Galaxy%20%28Stern%201980%29.vbs.original#L3083) [[patch]](Galaxy%20%28Stern%201980%29/Galaxy%20%28Stern%201980%29.vbs.patch) [[patched]](Galaxy%20%28Stern%201980%29/Galaxy%20%28Stern%201980%29.vbs)
- Game of Thrones LE (Stern 2015) VPW v1.0.2.vbs [[original]](Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%20v1.0.2/Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%20v1.0.2.vbs.original#L8960) [[patch]](Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%20v1.0.2/Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%20v1.0.2.vbs.patch) [[patched]](Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%20v1.0.2/Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%20v1.0.2.vbs)
- Genie (Gottlieb 1979).vbs [[original]](Genie%20%28Gottlieb%201979%29/Genie%20%28Gottlieb%201979%29.vbs.original#L408) [[patch]](Genie%20%28Gottlieb%201979%29/Genie%20%28Gottlieb%201979%29.vbs.patch) [[patched]](Genie%20%28Gottlieb%201979%29/Genie%20%28Gottlieb%201979%29.vbs)
- Gorgar_1.1.vbs [[original]](Gorgar_1.1/Gorgar_1.1.vbs.original#L2747) [[patch]](Gorgar_1.1/Gorgar_1.1.vbs.patch) [[patched]](Gorgar_1.1/Gorgar_1.1.vbs)
- Halloween 1978-1981 (Original 2022) 1.03.vbs [[original]](Halloween%201978-1981%20%28Original%202022%29%201.03/Halloween%201978-1981%20%28Original%202022%29%201.03.vbs.original#L3670) [[patch]](Halloween%201978-1981%20%28Original%202022%29%201.03/Halloween%201978-1981%20%28Original%202022%29%201.03.vbs.patch) [[patched]](Halloween%201978-1981%20%28Original%202022%29%201.03/Halloween%201978-1981%20%28Original%202022%29%201.03.vbs)
- Hang Glider (Bally 1976) VPW v1.2.vbs [[original]](Hang%20Glider%20%28Bally%201976%29%20VPW%20v1.2/Hang%20Glider%20%28Bally%201976%29%20VPW%20v1.2.vbs.original#L1420) [[patch]](Hang%20Glider%20%28Bally%201976%29%20VPW%20v1.2/Hang%20Glider%20%28Bally%201976%29%20VPW%20v1.2.vbs.patch) [[patched]](Hang%20Glider%20%28Bally%201976%29%20VPW%20v1.2/Hang%20Glider%20%28Bally%201976%29%20VPW%20v1.2.vbs)
- Harlem Globetrotters (Bally 1979) v1.14.vbs [[original]](Harlem%20Globetrotters%20%28Bally%201979%29%20v1.14/Harlem%20Globetrotters%20%28Bally%201979%29%20v1.14.vbs.original#L3055) [[patch]](Harlem%20Globetrotters%20%28Bally%201979%29%20v1.14/Harlem%20Globetrotters%20%28Bally%201979%29%20v1.14.vbs.patch) [[patched]](Harlem%20Globetrotters%20%28Bally%201979%29%20v1.14/Harlem%20Globetrotters%20%28Bally%201979%29%20v1.14.vbs)
- Harley Davidson (Sega 1999) v1.12.vbs [[original]](Harley%20Davidson%20%28Sega%201999%29%20v1.12/Harley%20Davidson%20%28Sega%201999%29%20v1.12.vbs.original#L618) [[patch]](Harley%20Davidson%20%28Sega%201999%29%20v1.12/Harley%20Davidson%20%28Sega%201999%29%20v1.12.vbs.patch) [[patched]](Harley%20Davidson%20%28Sega%201999%29%20v1.12/Harley%20Davidson%20%28Sega%201999%29%20v1.12.vbs)
- Heat Wave (Williams 1964).vbs [[original]](HeatWave%28Williams1964%291.0/Heat%20Wave%20%28Williams%201964%29.vbs.original#L1988) [[patch]](HeatWave%28Williams1964%291.0/Heat%20Wave%20%28Williams%201964%29.vbs.patch) [[patched]](HeatWave%28Williams1964%291.0/Heat%20Wave%20%28Williams%201964%29.vbs)
- Indiana Jones The Pinball Adventure (Williams 1993) VPWmod v1.1.vbs [[original]](Indiana%20Jones%20The%20Pinball%20Adventure%20%28Williams%201993%29%20VPWmod%20v1.1/Indiana%20Jones%20The%20Pinball%20Adventure%20%28Williams%201993%29%20VPWmod%20v1.1.vbs.original#L4357) [[patch]](Indiana%20Jones%20The%20Pinball%20Adventure%20%28Williams%201993%29%20VPWmod%20v1.1/Indiana%20Jones%20The%20Pinball%20Adventure%20%28Williams%201993%29%20VPWmod%20v1.1.vbs.patch) [[patched]](Indiana%20Jones%20The%20Pinball%20Adventure%20%28Williams%201993%29%20VPWmod%20v1.1/Indiana%20Jones%20The%20Pinball%20Adventure%20%28Williams%201993%29%20VPWmod%20v1.1.vbs)
- Iron Maiden (Original 2022) VPW 1.0.12.vbs [[original]](Iron%20Maiden%20%28Original%202022%29%20VPW%201.0.12/Iron%20Maiden%20%28Original%202022%29%20VPW%201.0.12.vbs.original#L7957) [[patch]](Iron%20Maiden%20%28Original%202022%29%20VPW%201.0.12/Iron%20Maiden%20%28Original%202022%29%20VPW%201.0.12.vbs.patch) [[patched]](Iron%20Maiden%20%28Original%202022%29%20VPW%201.0.12/Iron%20Maiden%20%28Original%202022%29%20VPW%201.0.12.vbs)
- Iron Man Vault Edition (Stern 2010) VPW v1.0.vbs [[original]](Iron%20Man%20Vault%20Edition%20%28Stern%202010%29%20VPW%20v1.0/Iron%20Man%20Vault%20Edition%20%28Stern%202010%29%20VPW%20v1.0.vbs.original#L3280) [[patch]](Iron%20Man%20Vault%20Edition%20%28Stern%202010%29%20VPW%20v1.0/Iron%20Man%20Vault%20Edition%20%28Stern%202010%29%20VPW%20v1.0.vbs.patch) [[patched]](Iron%20Man%20Vault%20Edition%20%28Stern%202010%29%20VPW%20v1.0/Iron%20Man%20Vault%20Edition%20%28Stern%202010%29%20VPW%20v1.0.vbs)
- Jack-Bot (Williams 1995).vbs [[original]](Jack-Bot%20%28Williams%201995%29/Jack-Bot%20%28Williams%201995%29.vbs.original#L3570) [[patch]](Jack-Bot%20%28Williams%201995%29/Jack-Bot%20%28Williams%201995%29.vbs.patch) [[patched]](Jack-Bot%20%28Williams%201995%29/Jack-Bot%20%28Williams%201995%29.vbs)
- Jive Time (Williams 1970) 2.0.vbs [[original]](Jive%20Time%20%28Williams%201970%29%202.0/Jive%20Time%20%28Williams%201970%29%202.0.vbs.original#L777) [[patch]](Jive%20Time%20%28Williams%201970%29%202.0/Jive%20Time%20%28Williams%201970%29%202.0.vbs.patch) [[patched]](Jive%20Time%20%28Williams%201970%29%202.0/Jive%20Time%20%28Williams%201970%29%202.0.vbs)
- Joker Poker EM (Gottlieb 1978) 1.6.vbs [[original]](Joker%20Poker%20EM%20%28Gottlieb%201978%29%201.6/Joker%20Poker%20EM%20%28Gottlieb%201978%29%201.6.vbs.original#L2948) [[patch]](Joker%20Poker%20EM%20%28Gottlieb%201978%29%201.6/Joker%20Poker%20EM%20%28Gottlieb%201978%29%201.6.vbs.patch) [[patched]](Joker%20Poker%20EM%20%28Gottlieb%201978%29%201.6/Joker%20Poker%20EM%20%28Gottlieb%201978%29%201.6.vbs)
- Judge Dredd (Bally 1993) VPW v1.1.vbs [[original]](Judge%20Dredd%20%28Bally%201993%29%20VPW%20v1.1/Judge%20Dredd%20%28Bally%201993%29%20VPW%20v1.1.vbs.original#L3045) [[patch]](Judge%20Dredd%20%28Bally%201993%29%20VPW%20v1.1/Judge%20Dredd%20%28Bally%201993%29%20VPW%20v1.1.vbs.patch) [[patched]](Judge%20Dredd%20%28Bally%201993%29%20VPW%20v1.1/Judge%20Dredd%20%28Bally%201993%29%20VPW%20v1.1.vbs)
- Jungle Lord (Williams 1981) w VR Room v2.01.vbs [[original]](Jungle%20Lord%20%28Williams%201981%29%20w%20VR%20Room%20v2.01/Jungle%20Lord%20%28Williams%201981%29%20w%20VR%20Room%20v2.01.vbs.original#L139) [[patch]](Jungle%20Lord%20%28Williams%201981%29%20w%20VR%20Room%20v2.01/Jungle%20Lord%20%28Williams%201981%29%20w%20VR%20Room%20v2.01.vbs.patch) [[patched]](Jungle%20Lord%20%28Williams%201981%29%20w%20VR%20Room%20v2.01/Jungle%20Lord%20%28Williams%201981%29%20w%20VR%20Room%20v2.01.vbs)
- Laser Cue (Williams 1984) w VR Room 2.0.0.vbs [[original]](Laser%20Cue%20%28Williams%201984%29%20w%20VR%20Room%202.0.0/Laser%20Cue%20%28Williams%201984%29%20w%20VR%20Room%202.0.0.vbs.original#L1847) [[patch]](Laser%20Cue%20%28Williams%201984%29%20w%20VR%20Room%202.0.0/Laser%20Cue%20%28Williams%201984%29%20w%20VR%20Room%202.0.0.vbs.patch) [[patched]](Laser%20Cue%20%28Williams%201984%29%20w%20VR%20Room%202.0.0/Laser%20Cue%20%28Williams%201984%29%20w%20VR%20Room%202.0.0.vbs)
- Laser War (Data East 1987) w VR Room v2.0.vbs [[original]](Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0/Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0.vbs.original#L1409) [[patch]](Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0/Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0.vbs.patch) [[patched]](Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0/Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0.vbs)
- Led Zeppelin Pinball 2.5.vbs [[original]](Led%20Zeppelin%20Pinball%202.5/Led%20Zeppelin%20Pinball%202.5.vbs.original#L461) [[patch]](Led%20Zeppelin%20Pinball%202.5/Led%20Zeppelin%20Pinball%202.5.vbs.patch) [[patched]](Led%20Zeppelin%20Pinball%202.5/Led%20Zeppelin%20Pinball%202.5.vbs)
- Magic City (Williams 1967).vbs [[original]](Magic%20City%20%28Williams%201967%291.0a/Magic%20City%20%28Williams%201967%29.vbs.original#L4325) [[patch]](Magic%20City%20%28Williams%201967%291.0a/Magic%20City%20%28Williams%201967%29.vbs.patch) [[patched]](Magic%20City%20%28Williams%201967%291.0a/Magic%20City%20%28Williams%201967%29.vbs)
- Maverick (Data East 1994) VPW v1.3.vbs [[original]](Maverick%20%28Data%20East%201994%29%20VPW%20v1.3/Maverick%20%28Data%20East%201994%29%20VPW%20v1.3.vbs.original#L848) [[patch]](Maverick%20%28Data%20East%201994%29%20VPW%20v1.3/Maverick%20%28Data%20East%201994%29%20VPW%20v1.3.vbs.patch) [[patched]](Maverick%20%28Data%20East%201994%29%20VPW%20v1.3/Maverick%20%28Data%20East%201994%29%20VPW%20v1.3.vbs)
- Medusa (Bally 1981) w VR Room v2.0.1.vbs [[original]](Medusa%20%28Bally%201981%29%20w%20VR%20Room%20v2.0.1/Medusa%20%28Bally%201981%29%20w%20VR%20Room%20v2.0.1.vbs.original#L2297) [[patch]](Medusa%20%28Bally%201981%29%20w%20VR%20Room%20v2.0.1/Medusa%20%28Bally%201981%29%20w%20VR%20Room%20v2.0.1.vbs.patch) [[patched]](Medusa%20%28Bally%201981%29%20w%20VR%20Room%20v2.0.1/Medusa%20%28Bally%201981%29%20w%20VR%20Room%20v2.0.1.vbs)
- Meteor (Stern 1979).vbs [[original]](Meteor%20%28Stern%201979%29/Meteor%20%28Stern%201979%29.vbs.original#L1593) [[patch]](Meteor%20%28Stern%201979%29/Meteor%20%28Stern%201979%29.vbs.patch) [[patched]](Meteor%20%28Stern%201979%29/Meteor%20%28Stern%201979%29.vbs)
- Monster Bash (Williams 1998) VPWmod v1.0.vbs [[original]](Monster%20Bash%20%28Williams%201998%29%20VPWmod%20v1.0/Monster%20Bash%20%28Williams%201998%29%20VPWmod%20v1.0.vbs.original#L2558) [[patch]](Monster%20Bash%20%28Williams%201998%29%20VPWmod%20v1.0/Monster%20Bash%20%28Williams%201998%29%20VPWmod%20v1.0.vbs.patch) [[patched]](Monster%20Bash%20%28Williams%201998%29%20VPWmod%20v1.0/Monster%20Bash%20%28Williams%201998%29%20VPWmod%20v1.0.vbs)
- MrDoom (Recel 1979)1.3.0.vbs [[original]](MrDoom%20%28Recel%201979%291.3.0/MrDoom%20%28Recel%201979%291.3.0.vbs.original#L1836) [[patch]](MrDoom%20%28Recel%201979%291.3.0/MrDoom%20%28Recel%201979%291.3.0.vbs.patch) [[patched]](MrDoom%20%28Recel%201979%291.3.0/MrDoom%20%28Recel%201979%291.3.0.vbs)
- MrEvil (Recel 1979)1.3.0.vbs [[original]](MrEvil%20%28Recel%201979%291.3.0/MrEvil%20%28Recel%201979%291.3.0.vbs.original#L1844) [[patch]](MrEvil%20%28Recel%201979%291.3.0/MrEvil%20%28Recel%201979%291.3.0.vbs.patch) [[patched]](MrEvil%20%28Recel%201979%291.3.0/MrEvil%20%28Recel%201979%291.3.0.vbs)
- Nine Ball (Stern 1980).vbs [[original]](Nine%20Ball%20%28Stern%201980%29/Nine%20Ball%20%28Stern%201980%29.vbs.original#L1591) [[patch]](Nine%20Ball%20%28Stern%201980%29/Nine%20Ball%20%28Stern%201980%29.vbs.patch) [[patched]](Nine%20Ball%20%28Stern%201980%29/Nine%20Ball%20%28Stern%201980%29.vbs)
- O Brother Where Art Thou (Zoss 2021)1_6.vbs [[original]](O%20Brother%20Where%20Art%20Thou%20%28Zoss%202021%291_6/O%20Brother%20Where%20Art%20Thou%20%28Zoss%202021%291_6.vbs.original#L2520) [[patch]](O%20Brother%20Where%20Art%20Thou%20%28Zoss%202021%291_6/O%20Brother%20Where%20Art%20Thou%20%28Zoss%202021%291_6.vbs.patch) [[patched]](O%20Brother%20Where%20Art%20Thou%20%28Zoss%202021%291_6/O%20Brother%20Where%20Art%20Thou%20%28Zoss%202021%291_6.vbs)
- Pharaoh (Williams 1981) w VR Room v2.0.3.vbs [[original]](Pharaoh%20%28Williams%201981%29%20w%20VR%20Room%20v2.0.3/Pharaoh%20%28Williams%201981%29%20w%20VR%20Room%20v2.0.3.vbs.original#L2385) [[patch]](Pharaoh%20%28Williams%201981%29%20w%20VR%20Room%20v2.0.3/Pharaoh%20%28Williams%201981%29%20w%20VR%20Room%20v2.0.3.vbs.patch) [[patched]](Pharaoh%20%28Williams%201981%29%20w%20VR%20Room%20v2.0.3/Pharaoh%20%28Williams%201981%29%20w%20VR%20Room%20v2.0.3.vbs)
- PinBlob (CLV 2024).vbs [[original]](PinBlob%20%28CLV%202024%29/PinBlob%20%28CLV%202024%29.vbs.original#L325) [[patch]](PinBlob%20%28CLV%202024%29/PinBlob%20%28CLV%202024%29.vbs.patch) [[patched]](PinBlob%20%28CLV%202024%29/PinBlob%20%28CLV%202024%29.vbs)
- PinBot (Williams 1986) 2.1.1.vbs [[original]](PinBot%20%28Williams%201986%29%202.1.1/PinBot%20%28Williams%201986%29%202.1.1.vbs.original#L3553) [[patch]](PinBot%20%28Williams%201986%29%202.1.1/PinBot%20%28Williams%201986%29%202.1.1.vbs.patch) [[patched]](PinBot%20%28Williams%201986%29%202.1.1/PinBot%20%28Williams%201986%29%202.1.1.vbs)
- Power Play (Bally 1977).vbs [[original]](1256692067_PowerPlay%28Bally1977%292.1/Power%20Play%20%28Bally%201977%29.vbs.original#L922) [[patch]](1256692067_PowerPlay%28Bally1977%292.1/Power%20Play%20%28Bally%201977%29.vbs.patch) [[patched]](1256692067_PowerPlay%28Bally1977%292.1/Power%20Play%20%28Bally%201977%29.vbs)
- Robocop (Data East 1989)_drakkon(mod_1.2).vbs [[original]](Robocop%20%28Data%20East%201989%29_drakkon%28mod_1.2%29/Robocop%20%28Data%20East%201989%29_drakkon%28mod_1.2%29.vbs.original#L2457) [[patch]](Robocop%20%28Data%20East%201989%29_drakkon%28mod_1.2%29/Robocop%20%28Data%20East%201989%29_drakkon%28mod_1.2%29.vbs.patch) [[patched]](Robocop%20%28Data%20East%201989%29_drakkon%28mod_1.2%29/Robocop%20%28Data%20East%201989%29_drakkon%28mod_1.2%29.vbs)
- Seawitch (Stern 1980).vbs [[original]](Seawitch%20%28Stern%201980%29/Seawitch%20%28Stern%201980%29.vbs.original#L1278) [[patch]](Seawitch%20%28Stern%201980%29/Seawitch%20%28Stern%201980%29.vbs.patch) [[patched]](Seawitch%20%28Stern%201980%29/Seawitch%20%28Stern%201980%29.vbs)
- Solar Fire (Williams 1981) w VR Room v2.0.5.vbs [[original]](Solar%20Fire%20%28Williams%201981%29%20w%20VR%20Room%20v2.0.5/Solar%20Fire%20%28Williams%201981%29%20w%20VR%20Room%20v2.0.5.vbs.original#L2398) [[patch]](Solar%20Fire%20%28Williams%201981%29%20w%20VR%20Room%20v2.0.5/Solar%20Fire%20%28Williams%201981%29%20w%20VR%20Room%20v2.0.5.vbs.patch) [[patched]](Solar%20Fire%20%28Williams%201981%29%20w%20VR%20Room%20v2.0.5/Solar%20Fire%20%28Williams%201981%29%20w%20VR%20Room%20v2.0.5.vbs)
- Space Shuttle (Williams 1984).vbs [[original]](Space%20Shuttle%20%28Williams%201984%292.0/Space%20Shuttle%20%28Williams%201984%29.vbs.original#L581) [[patch]](Space%20Shuttle%20%28Williams%201984%292.0/Space%20Shuttle%20%28Williams%201984%29.vbs.patch) [[patched]](Space%20Shuttle%20%28Williams%201984%292.0/Space%20Shuttle%20%28Williams%201984%29.vbs)
- Star Trek (Bally 1979) 2.1.1.vbs [[original]](Star%20Trek%20%28Bally%201979%29%202.1.1/Star%20Trek%20%28Bally%201979%29%202.1.1.vbs.original#L731) [[patch]](Star%20Trek%20%28Bally%201979%29%202.1.1/Star%20Trek%20%28Bally%201979%29%202.1.1.vbs.patch) [[patched]](Star%20Trek%20%28Bally%201979%29%202.1.1/Star%20Trek%20%28Bally%201979%29%202.1.1.vbs)
- Star Wars (Data East 1992) VPW v1.2.2.vbs [[original]](Star%20Wars%20%28Data%20East%201992%29%20VPW%20v1.2.2/Star%20Wars%20%28Data%20East%201992%29%20VPW%20v1.2.2.vbs.original#L3651) [[patch]](Star%20Wars%20%28Data%20East%201992%29%20VPW%20v1.2.2/Star%20Wars%20%28Data%20East%201992%29%20VPW%20v1.2.2.vbs.patch) [[patched]](Star%20Wars%20%28Data%20East%201992%29%20VPW%20v1.2.2/Star%20Wars%20%28Data%20East%201992%29%20VPW%20v1.2.2.vbs)
- Stars (Stern 1978).vbs [[original]](Stars%20%28Stern%201978%291.0.2/Stars%20%28Stern%201978%29.vbs.original#L1528) [[patch]](Stars%20%28Stern%201978%291.0.2/Stars%20%28Stern%201978%29.vbs.patch) [[patched]](Stars%20%28Stern%201978%291.0.2/Stars%20%28Stern%201978%29.vbs)
- Stellar Wars (Williams 1979).vbs [[original]](Stellar%20Wars%20%28Williams%201979%29/Stellar%20Wars%20%28Williams%201979%29.vbs.original#L2163) [[patch]](Stellar%20Wars%20%28Williams%201979%29/Stellar%20Wars%20%28Williams%201979%29.vbs.patch) [[patched]](Stellar%20Wars%20%28Williams%201979%29/Stellar%20Wars%20%28Williams%201979%29.vbs)
- Stingray (Stern 1977).vbs [[original]](Stingray%20%28Stern%201977%29%202.0/Stingray%20%28Stern%201977%29.vbs.original#L1696) [[patch]](Stingray%20%28Stern%201977%29%202.0/Stingray%20%28Stern%201977%29.vbs.patch) [[patched]](Stingray%20%28Stern%201977%29%202.0/Stingray%20%28Stern%201977%29.vbs)
- Strikes And Spares (Bally 1978) 2.0.vbs [[original]](Strikes%20And%20Spares%20%28Bally%201978%29%202.0/Strikes%20And%20Spares%20%28Bally%201978%29%202.0.vbs.original#L2093) [[patch]](Strikes%20And%20Spares%20%28Bally%201978%29%202.0/Strikes%20And%20Spares%20%28Bally%201978%29%202.0.vbs.patch) [[patched]](Strikes%20And%20Spares%20%28Bally%201978%29%202.0/Strikes%20And%20Spares%20%28Bally%201978%29%202.0.vbs)
- Strip Joker Poker (Gottlieb 1978) 1.5.vbs [[original]](Strip-Joker-Poker-%28Gottlieb-1978%29-1.5/Strip%20Joker%20Poker%20%28Gottlieb%201978%29%201.5.vbs.original#L2996) [[patch]](Strip-Joker-Poker-%28Gottlieb-1978%29-1.5/Strip%20Joker%20Poker%20%28Gottlieb%201978%29%201.5.vbs.patch) [[patched]](Strip-Joker-Poker-%28Gottlieb-1978%29-1.5/Strip%20Joker%20Poker%20%28Gottlieb%201978%29%201.5.vbs)
- Swords of Fury (Williams 1988).vbs [[original]](Swords%20of%20Fury%20%28Williams%201988%29/Swords%20of%20Fury%20%28Williams%201988%29.vbs.original#L610) [[patch]](Swords%20of%20Fury%20%28Williams%201988%29/Swords%20of%20Fury%20%28Williams%201988%29.vbs.patch) [[patched]](Swords%20of%20Fury%20%28Williams%201988%29/Swords%20of%20Fury%20%28Williams%201988%29.vbs)
- TX-Sector (Gottlieb 1988) SG1bsoN Mod V1.1.vbs [[original]](TX-Sector%20%28Gottlieb%201988%29%20SG1bsoN%20Mod%20V1.1/TX-Sector%20%28Gottlieb%201988%29%20SG1bsoN%20Mod%20V1.1.vbs.original#L809) [[patch]](TX-Sector%20%28Gottlieb%201988%29%20SG1bsoN%20Mod%20V1.1/TX-Sector%20%28Gottlieb%201988%29%20SG1bsoN%20Mod%20V1.1.vbs.patch) [[patched]](TX-Sector%20%28Gottlieb%201988%29%20SG1bsoN%20Mod%20V1.1/TX-Sector%20%28Gottlieb%201988%29%20SG1bsoN%20Mod%20V1.1.vbs)
- Tales from the Crypt (Data East 1993) VPW v1.22.vbs [[original]](Tales%20from%20the%20Crypt%20%28Data%20East%201993%29%20VPW%20v1.22/Tales%20from%20the%20Crypt%20%28Data%20East%201993%29%20VPW%20v1.22.vbs.original#L1528) [[patch]](Tales%20from%20the%20Crypt%20%28Data%20East%201993%29%20VPW%20v1.22/Tales%20from%20the%20Crypt%20%28Data%20East%201993%29%20VPW%20v1.22.vbs.patch) [[patched]](Tales%20from%20the%20Crypt%20%28Data%20East%201993%29%20VPW%20v1.22/Tales%20from%20the%20Crypt%20%28Data%20East%201993%29%20VPW%20v1.22.vbs)
- Transporter the Rescue (Midway 1989) VPW v1.05.vbs [[original]](Transporter%20the%20Rescue%20%28Midway%201989%29%20VPW%20v1.05/Transporter%20the%20Rescue%20%28Midway%201989%29%20VPW%20v1.05.vbs.original#L3331) [[patch]](Transporter%20the%20Rescue%20%28Midway%201989%29%20VPW%20v1.05/Transporter%20the%20Rescue%20%28Midway%201989%29%20VPW%20v1.05.vbs.patch) [[patched]](Transporter%20the%20Rescue%20%28Midway%201989%29%20VPW%20v1.05/Transporter%20the%20Rescue%20%28Midway%201989%29%20VPW%20v1.05.vbs)
- Twilight Zone (Bally 1993) 2.3.6.vbs [[original]](Twilight%20Zone%20%28Bally%201993%29%202.3.6/Twilight%20Zone%20%28Bally%201993%29%202.3.6.vbs.original#L2884) [[patch]](Twilight%20Zone%20%28Bally%201993%29%202.3.6/Twilight%20Zone%20%28Bally%201993%29%202.3.6.vbs.patch) [[patched]](Twilight%20Zone%20%28Bally%201993%29%202.3.6/Twilight%20Zone%20%28Bally%201993%29%202.3.6.vbs)
- Twister (Sega 1996) v2.0 w VR Room.vbs [[original]](Twister%20%28Sega%201996%29%20v2.0%20w%20VR%20Room/Twister%20%28Sega%201996%29%20v2.0%20w%20VR%20Room.vbs.original#L2045) [[patch]](Twister%20%28Sega%201996%29%20v2.0%20w%20VR%20Room/Twister%20%28Sega%201996%29%20v2.0%20w%20VR%20Room.vbs.patch) [[patched]](Twister%20%28Sega%201996%29%20v2.0%20w%20VR%20Room/Twister%20%28Sega%201996%29%20v2.0%20w%20VR%20Room.vbs)
- Viking (Bally 1980).vbs [[original]](Viking%20%28Bally%201980%29/Viking%20%28Bally%201980%29.vbs.original#L2124) [[patch]](Viking%20%28Bally%201980%29/Viking%20%28Bally%201980%29.vbs.patch) [[patched]](Viking%20%28Bally%201980%29/Viking%20%28Bally%201980%29.vbs)
- fireball II VPX.vbs [[original]](fireball%20II%20VPX/fireball%20II%20VPX.vbs.original#L38) [[patch]](fireball%20II%20VPX/fireball%20II%20VPX.vbs.patch) [[patched]](fireball%20II%20VPX/fireball%20II%20VPX.vbs)

## Parenthesis on first argument of a procedure call not handled correctly by Wine VBScript

**Affected patches: 67**

When the first argument of a Sub/procedure call starts with `(`, Wine's VBScript parses it as a call with explicit parentheses and treats the rest of the expression as a separate statement. Any arithmetic that continues after the closing `)` is evaluated separately and discarded. Example: `AddScore (a+b)*c` is parsed as `AddScore(a+b)` then `*c` is discarded. The same rule applies inside `ExecuteGlobal` strings — `SetLamp (me.UserValue - INT(me.UserValue)) * 100` becomes `SetLamp (me.UserValue - INT(me.UserValue))` with the `* 100` thrown away. Fix: rearrange the expression so it does not start with `(`, or move the multiplier to the front, e.g. `AddScore (a+b)*c` → `AddScore c*(a+b)`, or wrap the entire expression in double parentheses.

Affected scripts:

- 2104398928_CARtoonsRC(Nailed2021)v1.3.vbs [[original]](2104398928_CARtoonsRC%28Nailed2021%29v1.3/2104398928_CARtoonsRC%28Nailed2021%29v1.3.vbs.original#L841) [[patch]](2104398928_CARtoonsRC%28Nailed2021%29v1.3/2104398928_CARtoonsRC%28Nailed2021%29v1.3.vbs.patch) [[patched]](2104398928_CARtoonsRC%28Nailed2021%29v1.3/2104398928_CARtoonsRC%28Nailed2021%29v1.3.vbs)
- A Charlie Brown Christmas feat. Vince Guaraldi (iDigStuff 2023).vbs [[original]](A%20Charlie%20Brown%20Christmas%20feat.%20Vince%20Guaraldi%20%28iDigStuff%202023%29%201.0/A%20Charlie%20Brown%20Christmas%20feat.%20Vince%20Guaraldi%20%28iDigStuff%202023%29.vbs.original#L1198) [[patch]](A%20Charlie%20Brown%20Christmas%20feat.%20Vince%20Guaraldi%20%28iDigStuff%202023%29%201.0/A%20Charlie%20Brown%20Christmas%20feat.%20Vince%20Guaraldi%20%28iDigStuff%202023%29.vbs.patch) [[patched]](A%20Charlie%20Brown%20Christmas%20feat.%20Vince%20Guaraldi%20%28iDigStuff%202023%29%201.0/A%20Charlie%20Brown%20Christmas%20feat.%20Vince%20Guaraldi%20%28iDigStuff%202023%29.vbs)
- Aladdin's Castle (Bally 1976) - DOZER - MJR_1.01.vbs [[original]](Aladdin%27s%20Castle%20%28Bally%201976%29%20-%20DOZER%20-%20MJR_1.01/Aladdin%27s%20Castle%20%28Bally%201976%29%20-%20DOZER%20-%20MJR_1.01.vbs.original#L625) [[patch]](Aladdin%27s%20Castle%20%28Bally%201976%29%20-%20DOZER%20-%20MJR_1.01/Aladdin%27s%20Castle%20%28Bally%201976%29%20-%20DOZER%20-%20MJR_1.01.vbs.patch) [[patched]](Aladdin%27s%20Castle%20%28Bally%201976%29%20-%20DOZER%20-%20MJR_1.01/Aladdin%27s%20Castle%20%28Bally%201976%29%20-%20DOZER%20-%20MJR_1.01.vbs)
- Attack On Titan (cHuG_MOD_1.4).vbs [[original]](Attack%20On%20Titan%20%28cHuG_MOD_1.4%29/Attack%20On%20Titan%20%28cHuG_MOD_1.4%29.vbs.original#L2475) [[patch]](Attack%20On%20Titan%20%28cHuG_MOD_1.4%29/Attack%20On%20Titan%20%28cHuG_MOD_1.4%29.vbs.patch) [[patched]](Attack%20On%20Titan%20%28cHuG_MOD_1.4%29/Attack%20On%20Titan%20%28cHuG_MOD_1.4%29.vbs)
- Ben-Hur (Staal 1977) V1.1.1 DT-FS-VR-MR Ext2k Conversion.vbs [[original]](Ben-Hur%20%28Staal%201977%29%20V1.1.1%20DT-FS-VR-MR%20Ext2k%20Conversion/Ben-Hur%20%28Staal%201977%29%20V1.1.1%20DT-FS-VR-MR%20Ext2k%20Conversion.vbs.original#L1330) [[patch]](Ben-Hur%20%28Staal%201977%29%20V1.1.1%20DT-FS-VR-MR%20Ext2k%20Conversion/Ben-Hur%20%28Staal%201977%29%20V1.1.1%20DT-FS-VR-MR%20Ext2k%20Conversion.vbs.patch) [[patched]](Ben-Hur%20%28Staal%201977%29%20V1.1.1%20DT-FS-VR-MR%20Ext2k%20Conversion/Ben-Hur%20%28Staal%201977%29%20V1.1.1%20DT-FS-VR-MR%20Ext2k%20Conversion.vbs)
- Big Deal (Williams 1977)V2.1.vbs [[original]](Big%20Deal%20%28Williams%201977%29V2.1/Big%20Deal%20%28Williams%201977%29V2.1.vbs.original#L1067) [[patch]](Big%20Deal%20%28Williams%201977%29V2.1/Big%20Deal%20%28Williams%201977%29V2.1.vbs.patch) [[patched]](Big%20Deal%20%28Williams%201977%29V2.1/Big%20Deal%20%28Williams%201977%29V2.1.vbs)
- Big Horse (Maresa 1975) v55_VPX8.vbs [[original]](Big%20Horse%20%28Maresa%201975%29%20v55_VPX8/Big%20Horse%20%28Maresa%201975%29%20v55_VPX8.vbs.original#L1107) [[patch]](Big%20Horse%20%28Maresa%201975%29%20v55_VPX8/Big%20Horse%20%28Maresa%201975%29%20v55_VPX8.vbs.patch) [[patched]](Big%20Horse%20%28Maresa%201975%29%20v55_VPX8/Big%20Horse%20%28Maresa%201975%29%20v55_VPX8.vbs)
- Big Star (Williams 1972) v4.3.vbs [[original]](Big%20Star%20%28Williams%201972%29%20v4.3/Big%20Star%20%28Williams%201972%29%20v4.3.vbs.original#L1113) [[patch]](Big%20Star%20%28Williams%201972%29%20v4.3/Big%20Star%20%28Williams%201972%29%20v4.3.vbs.patch) [[patched]](Big%20Star%20%28Williams%201972%29%20v4.3/Big%20Star%20%28Williams%201972%29%20v4.3.vbs)
- Black & Red (Inder 1975) v55_VPX8.vbs [[original]](Black%20%26%20Red%20%28Inder%201975%29%20v55_VPX8/Black%20%26%20Red%20%28Inder%201975%29%20v55_VPX8.vbs.original#L931) [[patch]](Black%20%26%20Red%20%28Inder%201975%29%20v55_VPX8/Black%20%26%20Red%20%28Inder%201975%29%20v55_VPX8.vbs.patch) [[patched]](Black%20%26%20Red%20%28Inder%201975%29%20v55_VPX8/Black%20%26%20Red%20%28Inder%201975%29%20v55_VPX8.vbs)
- Bumper (Bill Port - 1977) v4.vbs [[original]](Bumper%20%28Bill%20Port%20-%201977%29%20v4/Bumper%20%28Bill%20Port%20-%201977%29%20v4.vbs.original#L973) [[patch]](Bumper%20%28Bill%20Port%20-%201977%29%20v4/Bumper%20%28Bill%20Port%20-%201977%29%20v4.vbs.patch) [[patched]](Bumper%20%28Bill%20Port%20-%201977%29%20v4/Bumper%20%28Bill%20Port%20-%201977%29%20v4.vbs)
- Cannes (Segasa 1976)V1.3.vbs [[original]](Cannes%20%28Segasa%201976%29V1.3/Cannes%20%28Segasa%201976%29V1.3.vbs.original#L1113) [[patch]](Cannes%20%28Segasa%201976%29V1.3/Cannes%20%28Segasa%201976%29V1.3.vbs.patch) [[patched]](Cannes%20%28Segasa%201976%29V1.3/Cannes%20%28Segasa%201976%29V1.3.vbs)
- DORAEMON.vbs [[original]](DORAEMON/DORAEMON.vbs.original#L1632) [[patch]](DORAEMON/DORAEMON.vbs.patch) [[patched]](DORAEMON/DORAEMON.vbs)
- DS.vbs [[original]](Drunken%20Santa%20%28Original%202021%29/DS.vbs.original#L1724) [[patch]](Drunken%20Santa%20%28Original%202021%29/DS.vbs.patch) [[patched]](Drunken%20Santa%20%28Original%202021%29/DS.vbs)
- DUCKTALES.vbs [[original]](Ducktales%20%28Orginal%202020%29/DUCKTALES.vbs.original#L1710) [[patch]](Ducktales%20%28Orginal%202020%29/DUCKTALES.vbs.patch) [[patched]](Ducktales%20%28Orginal%202020%29/DUCKTALES.vbs)
- FNAF.vbs [[original]](FNAF/FNAF.vbs.original#L1934) [[patch]](FNAF/FNAF.vbs.patch) [[patched]](FNAF/FNAF.vbs)
- Fifteen (Inder 1974) v4.vbs [[original]](Fifteen%20%28Inder%201974%29%20v4/Fifteen%20%28Inder%201974%29%20v4.vbs.original#L899) [[patch]](Fifteen%20%28Inder%201974%29%20v4/Fifteen%20%28Inder%201974%29%20v4.vbs.patch) [[patched]](Fifteen%20%28Inder%201974%29%20v4/Fifteen%20%28Inder%201974%29%20v4.vbs)
- Fireball XL5.vbs [[original]](Fireball%20XL5/Fireball%20XL5.vbs.original#L2424) [[patch]](Fireball%20XL5/Fireball%20XL5.vbs.patch) [[patched]](Fireball%20XL5/Fireball%20XL5.vbs)
- Game of Thrones LE (Stern 2015) VPW 1.2.vbs [[original]](Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%201.2/Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%201.2.vbs.original#L8915) [[patch]](Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%201.2/Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%201.2.vbs.patch) [[patched]](Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%201.2/Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%201.2.vbs)
- Game of Thrones LE (Stern 2015) VPW v1.0.2.vbs [[original]](Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%20v1.0.2/Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%20v1.0.2.vbs.original#L8960) [[patch]](Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%20v1.0.2/Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%20v1.0.2.vbs.patch) [[patched]](Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%20v1.0.2/Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%20v1.0.2.vbs)
- Gemini (Gottlieb 1978).vbs [[original]](Gemini%20%28Gottlieb%201978%29/Gemini%20%28Gottlieb%201978%29.vbs.original#L1173) [[patch]](Gemini%20%28Gottlieb%201978%29/Gemini%20%28Gottlieb%201978%29.vbs.patch) [[patched]](Gemini%20%28Gottlieb%201978%29/Gemini%20%28Gottlieb%201978%29.vbs)
- Gun Men (Staal 1979).vbs [[original]](Gun%20Men%20%28Staal%201979%29/Gun%20Men%20%28Staal%201979%29.vbs.original#L1226) [[patch]](Gun%20Men%20%28Staal%201979%29/Gun%20Men%20%28Staal%201979%29.vbs.patch) [[patched]](Gun%20Men%20%28Staal%201979%29/Gun%20Men%20%28Staal%201979%29.vbs)
- HALO.vbs [[original]](HALO/HALO.vbs.original#L1957) [[patch]](HALO/HALO.vbs.patch) [[patched]](HALO/HALO.vbs)
- Hellraiser 1.2.vbs [[original]](Hellraiser%201.2/Hellraiser%201.2.vbs.original#L2787) [[patch]](Hellraiser%201.2/Hellraiser%201.2.vbs.patch) [[patched]](Hellraiser%201.2/Hellraiser%201.2.vbs)
- Honey (Williams 1971)V1.3.vbs [[original]](Honey%20%28Williams%201971%29V1.3/Honey%20%28Williams%201971%29V1.3.vbs.original#L1063) [[patch]](Honey%20%28Williams%201971%29V1.3/Honey%20%28Williams%201971%29V1.3.vbs.patch) [[patched]](Honey%20%28Williams%201971%29V1.3/Honey%20%28Williams%201971%29V1.3.vbs)
- Indiana Jones (Stern 2008)-Hanibal-2.6.vbs [[original]](Indiana%20Jones%20%28Stern%202008%29-Hanibal-2.6/Indiana%20Jones%20%28Stern%202008%29-Hanibal-2.6.vbs.original#L2117) [[patch]](Indiana%20Jones%20%28Stern%202008%29-Hanibal-2.6/Indiana%20Jones%20%28Stern%202008%29-Hanibal-2.6.vbs.patch) [[patched]](Indiana%20Jones%20%28Stern%202008%29-Hanibal-2.6/Indiana%20Jones%20%28Stern%202008%29-Hanibal-2.6.vbs)
- Inhabiting Mars RC 4.vbs [[original]](Inhabiting%20Mars%20RC%204/Inhabiting%20Mars%20RC%204.vbs.original#L2311) [[patch]](Inhabiting%20Mars%20RC%204/Inhabiting%20Mars%20RC%204.vbs.patch) [[patched]](Inhabiting%20Mars%20RC%204/Inhabiting%20Mars%20RC%204.vbs)
- Liberty Bell (Williams 1977) V1.01.vbs [[original]](Liberty%20Bell%20%28Williams%201977%29%20V1.01/Liberty%20Bell%20%28Williams%201977%29%20V1.01.vbs.original#L404) [[patch]](Liberty%20Bell%20%28Williams%201977%29%20V1.01/Liberty%20Bell%20%28Williams%201977%29%20V1.01.vbs.patch) [[patched]](Liberty%20Bell%20%28Williams%201977%29%20V1.01/Liberty%20Bell%20%28Williams%201977%29%20V1.01.vbs)
- Lightning Ball (Gottlieb 1959).vbs [[original]](Lightning%20Ball%20%28Gottlieb%201959%29/Lightning%20Ball%20%28Gottlieb%201959%29.vbs.original#L1216) [[patch]](Lightning%20Ball%20%28Gottlieb%201959%29/Lightning%20Ball%20%28Gottlieb%201959%29.vbs.patch) [[patched]](Lightning%20Ball%20%28Gottlieb%201959%29/Lightning%20Ball%20%28Gottlieb%201959%29.vbs)
- Luck Smile (Inder 1976) 1p v55_VPX8.vbs [[original]](Luck%20Smile%20%28Inder%201976%29%201p%20v55_VPX8/Luck%20Smile%20%28Inder%201976%29%201p%20v55_VPX8.vbs.original#L897) [[patch]](Luck%20Smile%20%28Inder%201976%29%201p%20v55_VPX8/Luck%20Smile%20%28Inder%201976%29%201p%20v55_VPX8.vbs.patch) [[patched]](Luck%20Smile%20%28Inder%201976%29%201p%20v55_VPX8/Luck%20Smile%20%28Inder%201976%29%201p%20v55_VPX8.vbs)
- Luck Smile (Inder 1976) 4p v55_VPX8.vbs [[original]](Luck%20Smile%20%28Inder%201976%29%204p%20v55_VPX8/Luck%20Smile%20%28Inder%201976%29%204p%20v55_VPX8.vbs.original#L898) [[patch]](Luck%20Smile%20%28Inder%201976%29%204p%20v55_VPX8/Luck%20Smile%20%28Inder%201976%29%204p%20v55_VPX8.vbs.patch) [[patched]](Luck%20Smile%20%28Inder%201976%29%204p%20v55_VPX8/Luck%20Smile%20%28Inder%201976%29%204p%20v55_VPX8.vbs)
- Mago de Oz - the pinball v4.3.vbs [[original]](Mago%20de%20Oz%20-%20the%20pinball%20v4.3/Mago%20de%20Oz%20-%20the%20pinball%20v4.3.vbs.original#L2557) [[patch]](Mago%20de%20Oz%20-%20the%20pinball%20v4.3/Mago%20de%20Oz%20-%20the%20pinball%20v4.3.vbs.patch) [[patched]](Mago%20de%20Oz%20-%20the%20pinball%20v4.3/Mago%20de%20Oz%20-%20the%20pinball%20v4.3.vbs)
- Metal Slug_1.05.vbs [[original]](Metal%20Slug_1.05/Metal%20Slug_1.05.vbs.original#L72) [[patch]](Metal%20Slug_1.05/Metal%20Slug_1.05.vbs.patch) [[patched]](Metal%20Slug_1.05/Metal%20Slug_1.05.vbs)
- Monaco (Segasa 1977)V1.4.vbs [[original]](Monaco%20%28Segasa%201977%29V1.4/Monaco%20%28Segasa%201977%29V1.4.vbs.original#L1095) [[patch]](Monaco%20%28Segasa%201977%29V1.4/Monaco%20%28Segasa%201977%29V1.4.vbs.patch) [[patched]](Monaco%20%28Segasa%201977%29V1.4/Monaco%20%28Segasa%201977%29V1.4.vbs)
- Monkey Island VR Room.vbs [[original]](Monkey%20Island%20VR%20Room/Monkey%20Island%20VR%20Room.vbs.original#L240) [[patch]](Monkey%20Island%20VR%20Room/Monkey%20Island%20VR%20Room.vbs.patch) [[patched]](Monkey%20Island%20VR%20Room/Monkey%20Island%20VR%20Room.vbs)
- Motley Crue SS.vbs [[original]](Motley%20Crue%20SS/Motley%20Crue%20SS.vbs.original#L5229) [[patch]](Motley%20Crue%20SS/Motley%20Crue%20SS.vbs.patch) [[patched]](Motley%20Crue%20SS/Motley%20Crue%20SS.vbs)
- Munsters (Original 2020) 1.05.vbs [[original]](Munsters%20%28Original%202020%29%201.05/Munsters%20%28Original%202020%29%201.05.vbs.original#L257) [[patch]](Munsters%20%28Original%202020%29%201.05/Munsters%20%28Original%202020%29%201.05.vbs.patch) [[patched]](Munsters%20%28Original%202020%29%201.05/Munsters%20%28Original%202020%29%201.05.vbs)
- PVM.vbs [[original]](Putin%20Vodka%20Mania%20%28Original%202022%29/PVM.vbs.original#L2181) [[patch]](Putin%20Vodka%20Mania%20%28Original%202022%29/PVM.vbs.patch) [[patched]](Putin%20Vodka%20Mania%20%28Original%202022%29/PVM.vbs)
- Pat Hand (Williams 1975)V1.4.vbs [[original]](Pat%20Hand%20%28Williams%201975%29V1.4/Pat%20Hand%20%28Williams%201975%29V1.4.vbs.original#L1081) [[patch]](Pat%20Hand%20%28Williams%201975%29V1.4/Pat%20Hand%20%28Williams%201975%29V1.4.vbs.patch) [[patched]](Pat%20Hand%20%28Williams%201975%29V1.4/Pat%20Hand%20%28Williams%201975%29V1.4.vbs)
- Punchout.vbs [[original]](Punchout/Punchout.vbs.original#L2167) [[patch]](Punchout/Punchout.vbs.patch) [[patched]](Punchout/Punchout.vbs)
- Rancho (Williams 1976)V1.4.vbs [[original]](Rancho%20%28Williams%201976%29V1.4/Rancho%20%28Williams%201976%29V1.4.vbs.original#L1076) [[patch]](Rancho%20%28Williams%201976%29V1.4/Rancho%20%28Williams%201976%29V1.4.vbs.patch) [[patched]](Rancho%20%28Williams%201976%29V1.4/Rancho%20%28Williams%201976%29V1.4.vbs)
- Rattlecan v1.5.3.vbs [[original]](Rattlecan%20v1.5.3/Rattlecan%20v1.5.3.vbs.original#L604) [[patch]](Rattlecan%20v1.5.3/Rattlecan%20v1.5.3.vbs.patch) [[patched]](Rattlecan%20v1.5.3/Rattlecan%20v1.5.3.vbs)
- Riccione.vbs [[original]](Riccione_V1.4/Riccione.vbs.original#L1128) [[patch]](Riccione_V1.4/Riccione.vbs.patch) [[patched]](Riccione_V1.4/Riccione.vbs)
- Route66.vbs [[original]](Route66%280_2%29/Route66.vbs.original#L307) [[patch]](Route66%280_2%29/Route66.vbs.patch) [[patched]](Route66%280_2%29/Route66.vbs)
- Running Horse (Inder 1976) v55_VPX8.vbs [[original]](Running%20Horse%20%28Inder%201976%29%20v55_VPX8/Running%20Horse%20%28Inder%201976%29%20v55_VPX8.vbs.original#L900) [[patch]](Running%20Horse%20%28Inder%201976%29%20v55_VPX8/Running%20Horse%20%28Inder%201976%29%20v55_VPX8.vbs.patch) [[patched]](Running%20Horse%20%28Inder%201976%29%20v55_VPX8/Running%20Horse%20%28Inder%201976%29%20v55_VPX8.vbs)
- SMGP.vbs [[original]](SMGP%20%28Super%20Mario%20Galaxy%20Pinball%29/SMGP.vbs.original#L1848) [[patch]](SMGP%20%28Super%20Mario%20Galaxy%20Pinball%29/SMGP.vbs.patch) [[patched]](SMGP%20%28Super%20Mario%20Galaxy%20Pinball%29/SMGP.vbs)
- SPP.vbs [[original]](SouthParkPinballv1.2/SPP.vbs.original#L2185) [[patch]](SouthParkPinballv1.2/SPP.vbs.patch) [[patched]](SouthParkPinballv1.2/SPP.vbs)
- Seven Winner (Inder 1973) v4.vbs [[original]](Seven%20Winner%20%28Inder%201973%29%20v4/Seven%20Winner%20%28Inder%201973%29%20v4.vbs.original#L869) [[patch]](Seven%20Winner%20%28Inder%201973%29%20v4/Seven%20Winner%20%28Inder%201973%29%20v4.vbs.patch) [[patched]](Seven%20Winner%20%28Inder%201973%29%20v4/Seven%20Winner%20%28Inder%201973%29%20v4.vbs)
- Skylab (Williams 1974)V1.3.vbs [[original]](Skylab%20%28Williams%201974%29V1.3/Skylab%20%28Williams%201974%29V1.3.vbs.original#L1134) [[patch]](Skylab%20%28Williams%201974%29V1.3/Skylab%20%28Williams%201974%29V1.3.vbs.patch) [[patched]](Skylab%20%28Williams%201974%29V1.3/Skylab%20%28Williams%201974%29V1.3.vbs)
- Sonic The Hedgehog (Brendan Bailey 2005) VPX MOD 1.33.vbs [[original]](Sonic%20The%20Hedgehog%20%28Brendan%20Bailey%202005%29%20VPX%20MOD%201.33/Sonic%20The%20Hedgehog%20%28Brendan%20Bailey%202005%29%20VPX%20MOD%201.33.vbs.original#L156) [[patch]](Sonic%20The%20Hedgehog%20%28Brendan%20Bailey%202005%29%20VPX%20MOD%201.33/Sonic%20The%20Hedgehog%20%28Brendan%20Bailey%202005%29%20VPX%20MOD%201.33.vbs.patch) [[patched]](Sonic%20The%20Hedgehog%20%28Brendan%20Bailey%202005%29%20VPX%20MOD%201.33/Sonic%20The%20Hedgehog%20%28Brendan%20Bailey%202005%29%20VPX%20MOD%201.33.vbs)
- Sons Of Anarchy (Original 2019).vbs [[original]](Sons%20Of%20Anarchy%20%28Original%202019%29/Sons%20Of%20Anarchy%20%28Original%202019%29.vbs.original#L51) [[patch]](Sons%20Of%20Anarchy%20%28Original%202019%29/Sons%20Of%20Anarchy%20%28Original%202019%29.vbs.patch) [[patched]](Sons%20Of%20Anarchy%20%28Original%202019%29/Sons%20Of%20Anarchy%20%28Original%202019%29.vbs)
- Stardust (Williams 1971) v4.vbs [[original]](Stardust%20%28Williams%201971%29%20v4/Stardust%20%28Williams%201971%29%20v4.vbs.original#L933) [[patch]](Stardust%20%28Williams%201971%29%20v4/Stardust%20%28Williams%201971%29%20v4.vbs.patch) [[patched]](Stardust%20%28Williams%201971%29%20v4/Stardust%20%28Williams%201971%29%20v4.vbs)
- Streets of Rage (TBA 2018).vbs [[original]](Streets%20of%20Rage%20%28TBA%202018%29/Streets%20of%20Rage%20%28TBA%202018%29.vbs.original#L72) [[patch]](Streets%20of%20Rage%20%28TBA%202018%29/Streets%20of%20Rage%20%28TBA%202018%29.vbs.patch) [[patched]](Streets%20of%20Rage%20%28TBA%202018%29/Streets%20of%20Rage%20%28TBA%202018%29.vbs)
- Super Star (Williams 1972) v4.3.vbs [[original]](Super%20Star%20%28Williams%201972%29%20v4.3/Super%20Star%20%28Williams%201972%29%20v4.3.vbs.original#L1080) [[patch]](Super%20Star%20%28Williams%201972%29%20v4.3/Super%20Star%20%28Williams%201972%29%20v4.3.vbs.patch) [[patched]](Super%20Star%20%28Williams%201972%29%20v4.3/Super%20Star%20%28Williams%201972%29%20v4.3.vbs)
- The Grinch (Original 2022) pinballfan2018.vbs [[original]](The%20Grinch%20-%20Original%202022/The%20Grinch%20%28Original%202022%29%20pinballfan2018.vbs.original#L870) [[patch]](The%20Grinch%20-%20Original%202022/The%20Grinch%20%28Original%202022%29%20pinballfan2018.vbs.patch) [[patched]](The%20Grinch%20-%20Original%202022/The%20Grinch%20%28Original%202022%29%20pinballfan2018.vbs)
- TheATeam.vbs [[original]](TheATeam_2.0.7/TheATeam.vbs.original#L12146) [[patch]](TheATeam_2.0.7/TheATeam.vbs.patch) [[patched]](TheATeam_2.0.7/TheATeam.vbs)
- Van Halen 1.2.vbs [[original]](Van%20Halen%201.2/Van%20Halen%201.2.vbs.original#L4926) [[patch]](Van%20Halen%201.2/Van%20Halen%201.2.vbs.patch) [[patched]](Van%20Halen%201.2/Van%20Halen%201.2.vbs)
- indochinecentral.vbs [[original]](indochinecentral/indochinecentral.vbs.original#L2556) [[patch]](indochinecentral/indochinecentral.vbs.patch) [[patched]](indochinecentral/indochinecentral.vbs)
- indochinecentralPUP.vbs [[original]](indochinecentralPUP/indochinecentralPUP.vbs.original#L2556) [[patch]](indochinecentralPUP/indochinecentralPUP.vbs.patch) [[patched]](indochinecentralPUP/indochinecentralPUP.vbs)
- monkeyisland.vbs [[original]](Monkeyislandv1.1/monkeyisland.vbs.original#L2441) [[patch]](Monkeyislandv1.1/monkeyisland.vbs.patch) [[patched]](Monkeyislandv1.1/monkeyisland.vbs)
- pizzatime-65.vbs [[original]](pizzatime-65/pizzatime-65.vbs.original#L48) [[patch]](pizzatime-65/pizzatime-65.vbs.patch) [[patched]](pizzatime-65/pizzatime-65.vbs)
- pulp_fiction.vbs [[original]](Pulp%20Fiction%203.1/pulp_fiction.vbs.original#L2621) [[patch]](Pulp%20Fiction%203.1/pulp_fiction.vbs.patch) [[patched]](Pulp%20Fiction%203.1/pulp_fiction.vbs)
- speakeasy2.vbs [[original]](speakeasy2/speakeasy2.vbs.original#L1775) [[patch]](speakeasy2/speakeasy2.vbs.patch) [[patched]](speakeasy2/speakeasy2.vbs)
- speakeasy4.vbs [[original]](speakeasy4/speakeasy4.vbs.original#L1775) [[patch]](speakeasy4/speakeasy4.vbs.patch) [[patched]](speakeasy4/speakeasy4.vbs)
- wackyraces.vbs [[original]](wackyraces/wackyraces.vbs.original#L2829) [[patch]](wackyraces/wackyraces.vbs.patch) [[patched]](wackyraces/wackyraces.vbs)

## CreateObject / FileSystemObject / WshShell removed (not supported in Wine)

**Affected patches: 22**

`CreateObject("WScript.Shell")`, `CreateObject("Scripting.FileSystemObject")`, and `WshShell.RegWrite/RegRead` are Windows-only COM objects that are not available under Wine. Typically used for music folder scanning, registry access, or file system operations. These calls are removed or replaced with standalone-compatible equivalents.

Affected scripts:

- ABBAv2.0.vbs [[original]](ABBAv2.0/ABBAv2.0.vbs.original#L47) [[patch]](ABBAv2.0/ABBAv2.0.vbs.patch) [[patched]](ABBAv2.0/ABBAv2.0.vbs)
- Blood Machines 2.0.vbs [[original]](Blood%20Machines%202.0/Blood%20Machines%202.0.vbs.original#L205) [[patch]](Blood%20Machines%202.0/Blood%20Machines%202.0.vbs.patch) [[patched]](Blood%20Machines%202.0/Blood%20Machines%202.0.vbs)
- Bob Marley Mod.vbs [[original]](Bob%20Marley%20Mod/Bob%20Marley%20Mod.vbs.original#L59) [[patch]](Bob%20Marley%20Mod/Bob%20Marley%20Mod.vbs.patch) [[patched]](Bob%20Marley%20Mod/Bob%20Marley%20Mod.vbs)
- DarkPrincess1.3.1.vbs [[original]](DarkPrincess1.3.1/DarkPrincess1.3.1.vbs.original#L161) [[patch]](DarkPrincess1.3.1/DarkPrincess1.3.1.vbs.patch) [[patched]](DarkPrincess1.3.1/DarkPrincess1.3.1.vbs)
- Darkest Dungeon pupevent_2.3c.vbs [[original]](Darkest%20Dungeon%20pupevent_2.3c/Darkest%20Dungeon%20pupevent_2.3c.vbs.original#L175) [[patch]](Darkest%20Dungeon%20pupevent_2.3c/Darkest%20Dungeon%20pupevent_2.3c.vbs.patch) [[patched]](Darkest%20Dungeon%20pupevent_2.3c/Darkest%20Dungeon%20pupevent_2.3c.vbs)
- Diablo Pinball v4.3.vbs [[original]](Diablo%20Pinball%20v4.3/Diablo%20Pinball%20v4.3.vbs.original#L2344) [[patch]](Diablo%20Pinball%20v4.3/Diablo%20Pinball%20v4.3.vbs.patch) [[patched]](Diablo%20Pinball%20v4.3/Diablo%20Pinball%20v4.3.vbs)
- Fire Action (Taito do Brasil - 1980) v4.vbs [[original]](Fire%20Action%20%28Taito%20do%20Brasil%20-%201980%29%20v4/Fire%20Action%20%28Taito%20do%20Brasil%20-%201980%29%20v4.vbs.original#L1070) [[patch]](Fire%20Action%20%28Taito%20do%20Brasil%20-%201980%29%20v4/Fire%20Action%20%28Taito%20do%20Brasil%20-%201980%29%20v4.vbs.patch) [[patched]](Fire%20Action%20%28Taito%20do%20Brasil%20-%201980%29%20v4/Fire%20Action%20%28Taito%20do%20Brasil%20-%201980%29%20v4.vbs)
- Freddys Nightmares.vbs [[original]](Freddys%20Nightmares/Freddys%20Nightmares.vbs.original#L110) [[patch]](Freddys%20Nightmares/Freddys%20Nightmares.vbs.patch) [[patched]](Freddys%20Nightmares/Freddys%20Nightmares.vbs)
- Ice Age Christmas (Original Balutito 2021) endeemillr mod.vbs [[original]](Ice%20Age%20Christmas%20%28Original%20Balutito%202021%29%20endeemillr%20mod/Ice%20Age%20Christmas%20%28Original%20Balutito%202021%29%20endeemillr%20mod.vbs.original#L151) [[patch]](Ice%20Age%20Christmas%20%28Original%20Balutito%202021%29%20endeemillr%20mod/Ice%20Age%20Christmas%20%28Original%20Balutito%202021%29%20endeemillr%20mod.vbs.patch) [[patched]](Ice%20Age%20Christmas%20%28Original%20Balutito%202021%29%20endeemillr%20mod/Ice%20Age%20Christmas%20%28Original%20Balutito%202021%29%20endeemillr%20mod.vbs)
- Iron Maiden Virtual Time (Original 2020).vbs [[original]](Iron%20Maiden%20Virtual%20Time%20%28Original%202020%29/Iron%20Maiden%20Virtual%20Time%20%28Original%202020%29.vbs.original#L408) [[patch]](Iron%20Maiden%20Virtual%20Time%20%28Original%202020%29/Iron%20Maiden%20Virtual%20Time%20%28Original%202020%29.vbs.patch) [[patched]](Iron%20Maiden%20Virtual%20Time%20%28Original%202020%29/Iron%20Maiden%20Virtual%20Time%20%28Original%202020%29.vbs)
- KISS (Stern 2015)_Bigus(MOD)1.1.vbs [[original]](KISS%20%28Stern%202015%29_Bigus%28MOD%291.1/KISS%20%28Stern%202015%29_Bigus%28MOD%291.1.vbs.original#L3642) [[patch]](KISS%20%28Stern%202015%29_Bigus%28MOD%291.1/KISS%20%28Stern%202015%29_Bigus%28MOD%291.1.vbs.patch) [[patched]](KISS%20%28Stern%202015%29_Bigus%28MOD%291.1/KISS%20%28Stern%202015%29_Bigus%28MOD%291.1.vbs)
- Lawman (Gottlieb 1971) 1.1.vbs [[original]](Lawman%20%28Gottlieb%201971%29%201.1/Lawman%20%28Gottlieb%201971%29%201.1.vbs.original#L340) [[patch]](Lawman%20%28Gottlieb%201971%29%201.1/Lawman%20%28Gottlieb%201971%29%201.1.vbs.patch) [[patched]](Lawman%20%28Gottlieb%201971%29%201.1/Lawman%20%28Gottlieb%201971%29%201.1.vbs)
- Led Zeppelin Pinball 2.5.vbs [[original]](Led%20Zeppelin%20Pinball%202.5/Led%20Zeppelin%20Pinball%202.5.vbs.original#L461) [[patch]](Led%20Zeppelin%20Pinball%202.5/Led%20Zeppelin%20Pinball%202.5.vbs.patch) [[patched]](Led%20Zeppelin%20Pinball%202.5/Led%20Zeppelin%20Pinball%202.5.vbs)
- Metal Slug_1.05.vbs [[original]](Metal%20Slug_1.05/Metal%20Slug_1.05.vbs.original#L72) [[patch]](Metal%20Slug_1.05/Metal%20Slug_1.05.vbs.patch) [[patched]](Metal%20Slug_1.05/Metal%20Slug_1.05.vbs)
- Sonic The Hedgehog (Brendan Bailey 2005) VPX MOD 1.33.vbs [[original]](Sonic%20The%20Hedgehog%20%28Brendan%20Bailey%202005%29%20VPX%20MOD%201.33/Sonic%20The%20Hedgehog%20%28Brendan%20Bailey%202005%29%20VPX%20MOD%201.33.vbs.original#L156) [[patch]](Sonic%20The%20Hedgehog%20%28Brendan%20Bailey%202005%29%20VPX%20MOD%201.33/Sonic%20The%20Hedgehog%20%28Brendan%20Bailey%202005%29%20VPX%20MOD%201.33.vbs.patch) [[patched]](Sonic%20The%20Hedgehog%20%28Brendan%20Bailey%202005%29%20VPX%20MOD%201.33/Sonic%20The%20Hedgehog%20%28Brendan%20Bailey%202005%29%20VPX%20MOD%201.33.vbs)
- Space Shuttle (Taito do Brasil - 1982) v4.vbs [[original]](Space%20Shuttle%20%28Taito%20do%20Brasil%20-%201982%29%20v4/Space%20Shuttle%20%28Taito%20do%20Brasil%20-%201982%29%20v4.vbs.original#L962) [[patch]](Space%20Shuttle%20%28Taito%20do%20Brasil%20-%201982%29%20v4/Space%20Shuttle%20%28Taito%20do%20Brasil%20-%201982%29%20v4.vbs.patch) [[patched]](Space%20Shuttle%20%28Taito%20do%20Brasil%20-%201982%29%20v4/Space%20Shuttle%20%28Taito%20do%20Brasil%20-%201982%29%20v4.vbs)
- Spooky Wednesday Pro.vbs [[original]](Spooky%20Wednesday%20Pro/Spooky%20Wednesday%20Pro.vbs.original#L46) [[patch]](Spooky%20Wednesday%20Pro/Spooky%20Wednesday%20Pro.vbs.patch) [[patched]](Spooky%20Wednesday%20Pro/Spooky%20Wednesday%20Pro.vbs)
- Spooky_Wednesday VPX 2024.vbs [[original]](Spooky_Wednesday%20VPX%202024/Spooky_Wednesday%20VPX%202024.vbs.original#L46) [[patch]](Spooky_Wednesday%20VPX%202024/Spooky_Wednesday%20VPX%202024.vbs.patch) [[patched]](Spooky_Wednesday%20VPX%202024/Spooky_Wednesday%20VPX%202024.vbs)
- Student Prince (Williams 1968).vbs [[original]](Student%20Prince%20%28Williams%201968%29/Student%20Prince%20%28Williams%201968%29.vbs.original#L12) [[patch]](Student%20Prince%20%28Williams%201968%29/Student%20Prince%20%28Williams%201968%29.vbs.patch) [[patched]](Student%20Prince%20%28Williams%201968%29/Student%20Prince%20%28Williams%201968%29.vbs)
- Three Angels (Original 2018) 1.3.vbs [[original]](Three%20Angels%20%28Original%202018%29%201.3/Three%20Angels%20%28Original%202018%29%201.3.vbs.original#L291) [[patch]](Three%20Angels%20%28Original%202018%29%201.3/Three%20Angels%20%28Original%202018%29%201.3.vbs.patch) [[patched]](Three%20Angels%20%28Original%202018%29%201.3/Three%20Angels%20%28Original%202018%29%201.3.vbs)
- Thundercats Pinball v1.0.9.vbs [[original]](Thundercats%20Pinball%20v1.0.9/Thundercats%20Pinball%20v1.0.9.vbs.original#L63) [[patch]](Thundercats%20Pinball%20v1.0.9/Thundercats%20Pinball%20v1.0.9.vbs.patch) [[patched]](Thundercats%20Pinball%20v1.0.9/Thundercats%20Pinball%20v1.0.9.vbs)
- Vortex (Taito do Brasil - 1981) v4.vbs [[original]](Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4/Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4.vbs.original#L46) [[patch]](Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4/Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4.vbs.patch) [[patched]](Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4/Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4.vbs)

## DMD standalone compatibility setup (ImplicitDMD_Init / UseVPMDMD)

**Affected patches: 22**

Scripts that use a separate DMD script (`.vbs.dmd`) need a `Sub ImplicitDMD_Init` entry point and/or a `Dim UseVPMDMD` variable so that VPX Standalone can render the DMD correctly. These additions enable the DMD display to work without VPinMAME on standalone builds.

Affected scripts:

- Avatar (Stern 2012) v1.12.vbs
- [BarbWire(Gottlieb1996)JoePicassoModv1.2.vbs.dmd](BarbWire%28Gottlieb1996%29JoePicassoModv1.2/BarbWire%28Gottlieb1996%29JoePicassoModv1.2.vbs.dmd#L1)
- [Blood Machines 2.0.vbs.dmd](Blood%20Machines%202.0/Blood%20Machines%202.0.vbs.dmd#L100)
- [Bounty Hunter (Gottlieb 1985).vbs.dmd](Bounty%20Hunter%20%28Gottlieb%201985%29/Bounty%20Hunter%20%28Gottlieb%201985%29.vbs.dmd#L19)
- [Cactus Canyon (Bally 1998) VPW 1.1.vbs.dmd](Cactus%20Canyon%20%28Bally%201998%29%20VPW%201.1/Cactus%20Canyon%20%28Bally%201998%29%20VPW%201.1.vbs.dmd#L133)
- [Diner (Williams 1990) VPW Mod 1.0.2.vbs.dmd](Diner%20%28Williams%201990%29%20VPW%20Mod%201.0.2/Diner%20%28Williams%201990%29%20VPW%20Mod%201.0.2.vbs.dmdcolored#L446)
- Futurama (Original 2024) v1.1.vbs
- [Game of Thrones LE (Stern 2015) VPW v1.0.2.vbs.dmd](Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%20v1.0.2/Game%20of%20Thrones%20LE%20%28Stern%202015%29%20VPW%20v1.0.2.vbs.dmd#L163)
- [Goin' Nuts (Gottlieb 1983).vbs.dmd](Goin%27%20Nuts%20%28Gottlieb%201983%29/Goin%27%20Nuts%20%28Gottlieb%201983%29.vbs.dmd#L29)
- [Hook (Data East 1992)_VPWmod_v1.0.vbs.dmd](Hook%20%28Data%20East%201992%29_VPWmod_v1.0/Hook%20%28Data%20East%201992%29_VPWmod_v1.0.vbs.dmd#L91)
- [Pink Panther (Gottlieb 1981).vbs.dmd](Pink%20Panther%20%28Gottlieb%201981%29/Pink%20Panther%20%28Gottlieb%201981%29.vbs.dmd#L8)
- [Scared Stiff (Bally 1996) v1.29.vbs.dmd](Scared%20Stiff%20%28Bally%201996%29%20v1.29/Scared%20Stiff%20%28Bally%201996%29%20v1.29.vbs.dmd#L15)
- [Seawitch (Stern 1980).vbs.dmd](Seawitch%20%28Stern%201980%29/Seawitch%20%28Stern%201980%29.vbs.dmd#L21)
- [South Park (Sega 1999) 1.3.vbs.dmd](South%20Park%20%28Sega%201999%29%201.3/South%20Park%20%28Sega%201999%29%201.3.vbs.dmd#L38)
- [Space Shuttle (Williams 1984).vbs.dmd](Space%20Shuttle%20%28Williams%201984%292.0/Space%20Shuttle%20%28Williams%201984%29.vbs.dmd#L57)
- [Star Trek LE (Stern 2013) v1.09.vbs.dmd](Star%20Trek%20LE%20%28Stern%202013%29%20v1.09/Star%20Trek%20LE%20%28Stern%202013%29%20v1.09.vbs.dmd#L91)
- [Taxi (Williams 1988) VPW v1.2.2.vbs.dmd](Taxi%20%28Williams%201988%29%20VPW%20v1.2.2/Taxi%20%28Williams%201988%29%20VPW%20v1.2.2.vbs.dmd#L87)
- The Goonies Never Say Die Pinball (VPW 2021) v1.4.vbs
- Three Angels (Original 2018) 1.3.vbs [[original]](Three%20Angels%20%28Original%202018%29%201.3/Three%20Angels%20%28Original%202018%29%201.3.vbs.original#L291) [[patch]](Three%20Angels%20%28Original%202018%29%201.3/Three%20Angels%20%28Original%202018%29%201.3.vbs.patch) [[patched]](Three%20Angels%20%28Original%202018%29%201.3/Three%20Angels%20%28Original%202018%29%201.3.vbs)
- X-Files (Sega 1997) v1.29.vbs
- X-Files Hanibal 4k LED Edition.vbs [[original]](X-Files%20Hanibal%204k%20LED%20Edition/X-Files%20Hanibal%204k%20LED%20Edition.vbs.original#L15) [[patch]](X-Files%20Hanibal%204k%20LED%20Edition/X-Files%20Hanibal%204k%20LED%20Edition.vbs.patch) [[patched]](X-Files%20Hanibal%204k%20LED%20Edition/X-Files%20Hanibal%204k%20LED%20Edition.vbs)
- [fireball II VPX.vbs.dmd](fireball%20II%20VPX/fireball%20II%20VPX.vbs.dmd#L38)

## Const declaration moved to different position (execution order fix)

**Affected patches: 16**

VBScript evaluates code sequentially; a `Const` must be declared before its first use. In some scripts the constant was declared *after* the code that referenced it. Fix: move the `Const` declaration to an earlier position in the file.

Affected scripts:

- AC-DC Pro Vault-1.0.vbs [[original]](AC-DC%20ninuzzu%201.0%201.5/AC-DC%20Pro%20Vault-1.0.vbs.original#L89) [[patch]](AC-DC%20ninuzzu%201.0%201.5/AC-DC%20Pro%20Vault-1.0.vbs.patch) [[patched]](AC-DC%20ninuzzu%201.0%201.5/AC-DC%20Pro%20Vault-1.0.vbs)
- AC-DC Pro-1.0.vbs [[original]](AC-DC%20ninuzzu%201.0%201.5/AC-DC%20Pro-1.0.vbs.original#L89) [[patch]](AC-DC%20ninuzzu%201.0%201.5/AC-DC%20Pro-1.0.vbs.patch) [[patched]](AC-DC%20ninuzzu%201.0%201.5/AC-DC%20Pro-1.0.vbs)
- AC-DC_Back In Black LE-1.5.vbs [[original]](AC-DC%20ninuzzu%201.0%201.5/AC-DC_Back%20In%20Black%20LE-1.5.vbs.original#L98) [[patch]](AC-DC%20ninuzzu%201.0%201.5/AC-DC_Back%20In%20Black%20LE-1.5.vbs.patch) [[patched]](AC-DC%20ninuzzu%201.0%201.5/AC-DC_Back%20In%20Black%20LE-1.5.vbs)
- AC-DC_Let There Be Rock LE-1.5.vbs [[original]](AC-DC%20ninuzzu%201.0%201.5/AC-DC_Let%20There%20Be%20Rock%20LE-1.5.vbs.original#L98) [[patch]](AC-DC%20ninuzzu%201.0%201.5/AC-DC_Let%20There%20Be%20Rock%20LE-1.5.vbs.patch) [[patched]](AC-DC%20ninuzzu%201.0%201.5/AC-DC_Let%20There%20Be%20Rock%20LE-1.5.vbs)
- AC-DC_Luci-1.5.vbs [[original]](AC-DC%20ninuzzu%201.0%201.5/AC-DC_Luci-1.5.vbs.original#L98) [[patch]](AC-DC%20ninuzzu%201.0%201.5/AC-DC_Luci-1.5.vbs.patch) [[patched]](AC-DC%20ninuzzu%201.0%201.5/AC-DC_Luci-1.5.vbs)
- AC-DC_Premium-1.5.vbs [[original]](AC-DC%20ninuzzu%201.0%201.5/AC-DC_Premium-1.5.vbs.original#L98) [[patch]](AC-DC%20ninuzzu%201.0%201.5/AC-DC_Premium-1.5.vbs.patch) [[patched]](AC-DC%20ninuzzu%201.0%201.5/AC-DC_Premium-1.5.vbs)
- Aztec (Williams 1976) 1.3 Mod Citedor JPJ-ARNGRIM-CED Team PP.vbs [[original]](Aztec%20%28Williams%201976%29%201.3%20Mod%20Citedor%20JPJ-ARNGRIM-CED%20Team%20PP/Aztec%20%28Williams%201976%29%201.3%20Mod%20Citedor%20JPJ-ARNGRIM-CED%20Team%20PP.vbs.original#L1) [[patch]](Aztec%20%28Williams%201976%29%201.3%20Mod%20Citedor%20JPJ-ARNGRIM-CED%20Team%20PP/Aztec%20%28Williams%201976%29%201.3%20Mod%20Citedor%20JPJ-ARNGRIM-CED%20Team%20PP.vbs.patch) [[patched]](Aztec%20%28Williams%201976%29%201.3%20Mod%20Citedor%20JPJ-ARNGRIM-CED%20Team%20PP/Aztec%20%28Williams%201976%29%201.3%20Mod%20Citedor%20JPJ-ARNGRIM-CED%20Team%20PP.vbs)
- Aztec High-Tapped (Williams 1976).vbs [[original]](Aztec%20High-Tapped%20%28Williams%201976%29/Aztec%20High-Tapped%20%28Williams%201976%29.vbs.original#L1747) [[patch]](Aztec%20High-Tapped%20%28Williams%201976%29/Aztec%20High-Tapped%20%28Williams%201976%29.vbs.patch) [[patched]](Aztec%20High-Tapped%20%28Williams%201976%29/Aztec%20High-Tapped%20%28Williams%201976%29.vbs)
- Cybernaut (Bally 1985)_Bigus(MOD)1.0.vbs [[original]](Cybernaut%20%28Bally%201985%29/Cybernaut%20%28Bally%201985%29_Bigus%28MOD%291.0.vbs.original#L40) [[patch]](Cybernaut%20%28Bally%201985%29/Cybernaut%20%28Bally%201985%29_Bigus%28MOD%291.0.vbs.patch) [[patched]](Cybernaut%20%28Bally%201985%29/Cybernaut%20%28Bally%201985%29_Bigus%28MOD%291.0.vbs)
- Cybernaut darkmod.vbs [[original]](cybernaut%20Bally/Cybernaut%20darkmod.vbs.original#L43) [[patch]](cybernaut%20Bally/Cybernaut%20darkmod.vbs.patch) [[patched]](cybernaut%20Bally/Cybernaut%20darkmod.vbs)
- Cybernaut.vbs [[original]](cybernaut%20Bally/Cybernaut.vbs.original#L55) [[patch]](cybernaut%20Bally/Cybernaut.vbs.patch) [[patched]](cybernaut%20Bally/Cybernaut.vbs)
- Fireball Classic (Bally 1984)1.2.vbs [[original]](Fireball%20Classic%20%28Bally%201984%291.2/Fireball%20Classic%20%28Bally%201984%291.2.vbs.original#L1) [[patch]](Fireball%20Classic%20%28Bally%201984%291.2/Fireball%20Classic%20%28Bally%201984%291.2.vbs.patch) [[patched]](Fireball%20Classic%20%28Bally%201984%291.2/Fireball%20Classic%20%28Bally%201984%291.2.vbs)
- Flash Gordon (Bally 1981) VPW Mod v3.0.vbs [[original]](Flash%20Gordon%20%28Bally%201981%29%20VPW%20Mod%20v3.0/Flash%20Gordon%20%28Bally%201981%29%20VPW%20Mod%20v3.0.vbs.original#L74) [[patch]](Flash%20Gordon%20%28Bally%201981%29%20VPW%20Mod%20v3.0/Flash%20Gordon%20%28Bally%201981%29%20VPW%20Mod%20v3.0.vbs.patch) [[patched]](Flash%20Gordon%20%28Bally%201981%29%20VPW%20Mod%20v3.0/Flash%20Gordon%20%28Bally%201981%29%20VPW%20Mod%20v3.0.vbs)
- Harlem Globetrotters (Bally 1979) VPW 1.0.vbs [[original]](Harlem%20Globetrotters%20%28Bally%201979%29%20VPW%201.0/Harlem%20Globetrotters%20%28Bally%201979%29%20VPW%201.0.vbs.original#L249) [[patch]](Harlem%20Globetrotters%20%28Bally%201979%29%20VPW%201.0/Harlem%20Globetrotters%20%28Bally%201979%29%20VPW%201.0.vbs.patch) [[patched]](Harlem%20Globetrotters%20%28Bally%201979%29%20VPW%201.0/Harlem%20Globetrotters%20%28Bally%201979%29%20VPW%201.0.vbs)
- High_Speed_(Williams 1986) v0.107a (MOD 3.0).vbs [[original]](High_Speed_%28Williams%201986%29%20v0.107a%20%28MOD%203.0%29/High_Speed_%28Williams%201986%29%20v0.107a%20%28MOD%203.0%29.vbs.original#L6) [[patch]](High_Speed_%28Williams%201986%29%20v0.107a%20%28MOD%203.0%29/High_Speed_%28Williams%201986%29%20v0.107a%20%28MOD%203.0%29.vbs.patch) [[patched]](High_Speed_%28Williams%201986%29%20v0.107a%20%28MOD%203.0%29/High_Speed_%28Williams%201986%29%20v0.107a%20%28MOD%203.0%29.vbs)
- Rollercoaster Tycoon (Stern 2002) VPW 2.3.vbs [[original]](Rollercoaster%20Tycoon%20%28Stern%202002%29%20VPW%202.3/Rollercoaster%20Tycoon%20%28Stern%202002%29%20VPW%202.3.vbs.original#L53) [[patch]](Rollercoaster%20Tycoon%20%28Stern%202002%29%20VPW%202.3/Rollercoaster%20Tycoon%20%28Stern%202002%29%20VPW%202.3.vbs.patch) [[patched]](Rollercoaster%20Tycoon%20%28Stern%202002%29%20VPW%202.3/Rollercoaster%20Tycoon%20%28Stern%202002%29%20VPW%202.3.vbs)

## NVramPatch calls removed (not supported in standalone)

**Affected patches: 13**

`NVramPatchLoad`, `NVramPatchExit`, and `NVramPatchKeyCheck` rely on a Windows-only NVram helper that is not available in standalone builds. Calls to these functions are commented out or removed.

Affected scripts:

- ABBAv2.0.vbs [[original]](ABBAv2.0/ABBAv2.0.vbs.original#L47) [[patch]](ABBAv2.0/ABBAv2.0.vbs.patch) [[patched]](ABBAv2.0/ABBAv2.0.vbs)
- Bob Marley Mod.vbs [[original]](Bob%20Marley%20Mod/Bob%20Marley%20Mod.vbs.original#L59) [[patch]](Bob%20Marley%20Mod/Bob%20Marley%20Mod.vbs.patch) [[patched]](Bob%20Marley%20Mod/Bob%20Marley%20Mod.vbs)
- Cosmic (Taito do Brasil - 1980) v4.vbs [[original]](Cosmic%20%28Taito%20do%20Brasil%20-%201980%29%20v4/Cosmic%20%28Taito%20do%20Brasil%20-%201980%29%20v4.vbs.original#L56) [[patch]](Cosmic%20%28Taito%20do%20Brasil%20-%201980%29%20v4/Cosmic%20%28Taito%20do%20Brasil%20-%201980%29%20v4.vbs.patch) [[patched]](Cosmic%20%28Taito%20do%20Brasil%20-%201980%29%20v4/Cosmic%20%28Taito%20do%20Brasil%20-%201980%29%20v4.vbs)
- Last Starfighter, The (Taito, 1983) hybrid v1.04.vbs [[original]](Last%20Starfighter%2C%20The%20%28Taito%2C%201983%29%20hybrid%20v1.04/Last%20Starfighter%2C%20The%20%28Taito%2C%201983%29%20hybrid%20v1.04.vbs.original#L104) [[patch]](Last%20Starfighter%2C%20The%20%28Taito%2C%201983%29%20hybrid%20v1.04/Last%20Starfighter%2C%20The%20%28Taito%2C%201983%29%20hybrid%20v1.04.vbs.patch) [[patched]](Last%20Starfighter%2C%20The%20%28Taito%2C%201983%29%20hybrid%20v1.04/Last%20Starfighter%2C%20The%20%28Taito%2C%201983%29%20hybrid%20v1.04.vbs)
- Mr Black (Taito do Brasil - 1984) v4.vbs [[original]](Mr%20Black%20%28Taito%20do%20Brasil%20-%201984%29%20v4/Mr%20Black%20%28Taito%20do%20Brasil%20-%201984%29%20v4.vbs.original#L61) [[patch]](Mr%20Black%20%28Taito%20do%20Brasil%20-%201984%29%20v4/Mr%20Black%20%28Taito%20do%20Brasil%20-%201984%29%20v4.vbs.patch) [[patched]](Mr%20Black%20%28Taito%20do%20Brasil%20-%201984%29%20v4/Mr%20Black%20%28Taito%20do%20Brasil%20-%201984%29%20v4.vbs)
- Oba Oba (Taito do Brasil - 1979) v55_VPX8.vbs [[original]](Oba%20Oba%20%28Taito%20do%20Brasil%20-%201979%29%20v55_VPX8/Oba%20Oba%20%28Taito%20do%20Brasil%20-%201979%29%20v55_VPX8.vbs.original#L55) [[patch]](Oba%20Oba%20%28Taito%20do%20Brasil%20-%201979%29%20v55_VPX8/Oba%20Oba%20%28Taito%20do%20Brasil%20-%201979%29%20v55_VPX8.vbs.patch) [[patched]](Oba%20Oba%20%28Taito%20do%20Brasil%20-%201979%29%20v55_VPX8/Oba%20Oba%20%28Taito%20do%20Brasil%20-%201979%29%20v55_VPX8.vbs)
- Rally (Taito do Brasil - 1980) v4.vbs [[original]](Rally%20%28Taito%20do%20Brasil%20-%201980%29%20v4/Rally%20%28Taito%20do%20Brasil%20-%201980%29%20v4.vbs.original#L60) [[patch]](Rally%20%28Taito%20do%20Brasil%20-%201980%29%20v4/Rally%20%28Taito%20do%20Brasil%20-%201980%29%20v4.vbs.patch) [[patched]](Rally%20%28Taito%20do%20Brasil%20-%201980%29%20v4/Rally%20%28Taito%20do%20Brasil%20-%201980%29%20v4.vbs)
- Titan (Taito do Brasil - 1981) v4.vbs [[original]](Titan%20%28Taito%20do%20Brasil%20-%201981%29%20v4/Titan%20%28Taito%20do%20Brasil%20-%201981%29%20v4.vbs.original#L57) [[patch]](Titan%20%28Taito%20do%20Brasil%20-%201981%29%20v4/Titan%20%28Taito%20do%20Brasil%20-%201981%29%20v4.vbs.patch) [[patched]](Titan%20%28Taito%20do%20Brasil%20-%201981%29%20v4/Titan%20%28Taito%20do%20Brasil%20-%201981%29%20v4.vbs)
- Topaz (Inder 1979).vbs [[original]](Topaz%20%28Inder%201979%29/Topaz%20%28Inder%201979%29.vbs.original#L100) [[patch]](Topaz%20%28Inder%201979%29/Topaz%20%28Inder%201979%29.vbs.patch) [[patched]](Topaz%20%28Inder%201979%29/Topaz%20%28Inder%201979%29.vbs)
- Vortex (Taito do Brasil - 1981) v4.vbs [[original]](Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4/Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4.vbs.original#L46) [[patch]](Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4/Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4.vbs.patch) [[patched]](Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4/Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4.vbs)
- Warriors The Full DMD 2.0 (Iceman 2023).vbs [[original]](Warriors%20The%20Full%20DMD%202.0%20%28Iceman%202023%29/Warriors%20The%20Full%20DMD%202.0%20%28Iceman%202023%29.vbs.original#L14) [[patch]](Warriors%20The%20Full%20DMD%202.0%20%28Iceman%202023%29/Warriors%20The%20Full%20DMD%202.0%20%28Iceman%202023%29.vbs.patch) [[patched]](Warriors%20The%20Full%20DMD%202.0%20%28Iceman%202023%29/Warriors%20The%20Full%20DMD%202.0%20%28Iceman%202023%29.vbs)
- Wednesday (Netflix 2023).vbs [[original]](Wednesday%20%28Netflix%202023%29/Wednesday%20%28Netflix%202023%29.vbs.original#L19) [[patch]](Wednesday%20%28Netflix%202023%29/Wednesday%20%28Netflix%202023%29.vbs.patch) [[patched]](Wednesday%20%28Netflix%202023%29/Wednesday%20%28Netflix%202023%29.vbs)

## cGameName (ROM name) constant fix

**Affected patches: 10**

The `cGameName` constant sets the ROM name passed to VPinMAME. An incorrect ROM name prevents the game from loading. This fix corrects the ROM name string to match the actual ROM file name.

Affected scripts:

- [Bad Cats (Williams 1989) VPW 3.0.vbs.dmd](Bad%20Cats%20%28Williams%201989%29%20VPW%203.0/Bad%20Cats%20%28Williams%201989%29%20VPW%203.0.vbs.dmdcolored#L384)
- BeastieBoysv1.0.vbs [[original]](BeastieBoysv1.0/BeastieBoysv1.0.vbs.original#L100) [[patch]](BeastieBoysv1.0/BeastieBoysv1.0.vbs.patch) [[patched]](BeastieBoysv1.0/BeastieBoysv1.0.vbs)
- Death Proof Balutito V2.vbs [[original]](Death%20Proof%20Balutito%20V2/Death%20Proof%20Balutito%20V2.vbs.original#L207) [[patch]](Death%20Proof%20Balutito%20V2/Death%20Proof%20Balutito%20V2.vbs.patch) [[patched]](Death%20Proof%20Balutito%20V2/Death%20Proof%20Balutito%20V2.vbs)
- PinBlob (CLV 2024).vbs [[original]](PinBlob%20%28CLV%202024%29/PinBlob%20%28CLV%202024%29.vbs.original#L325) [[patch]](PinBlob%20%28CLV%202024%29/PinBlob%20%28CLV%202024%29.vbs.patch) [[patched]](PinBlob%20%28CLV%202024%29/PinBlob%20%28CLV%202024%29.vbs)
- Safe Cracker (Bally 1996) v1.0.vbs [[original]](Safe%20Cracker%20%28Bally%201996%29%20v1.0/Safe%20Cracker%20%28Bally%201996%29%20v1.0.vbs.original#L54) [[patch]](Safe%20Cracker%20%28Bally%201996%29%20v1.0/Safe%20Cracker%20%28Bally%201996%29%20v1.0.vbs.patch) [[patched]](Safe%20Cracker%20%28Bally%201996%29%20v1.0/Safe%20Cracker%20%28Bally%201996%29%20v1.0.vbs)
- Simpsons Treehouse of Horror MOD v2023.3 (Simpsons Pinball Party, The (Stern 2003) VPW 2.0.3 base).vbs [[original]](Simpsons%20Treehouse%20of%20Horror%20MOD%20v2023.3%20%28Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3%20base%29/Simpsons%20Treehouse%20of%20Horror%20MOD%20v2023.3%20%28Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3%20base%29.vbs.original#L95) [[patch]](Simpsons%20Treehouse%20of%20Horror%20MOD%20v2023.3%20%28Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3%20base%29/Simpsons%20Treehouse%20of%20Horror%20MOD%20v2023.3%20%28Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3%20base%29.vbs.patch) [[patched]](Simpsons%20Treehouse%20of%20Horror%20MOD%20v2023.3%20%28Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3%20base%29/Simpsons%20Treehouse%20of%20Horror%20MOD%20v2023.3%20%28Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3%20base%29.vbs)
- South Park - Halloween.vbs [[original]](South%20Park%20-%20Halloween/South%20Park%20-%20Halloween.vbs.original#L49) [[patch]](South%20Park%20-%20Halloween/South%20Park%20-%20Halloween.vbs.patch) [[patched]](South%20Park%20-%20Halloween/South%20Park%20-%20Halloween.vbs)
- Theatre of Magic (Bally 1995) 2.4.vbs [[original]](Theatre%20of%20Magic%20%28Bally%201995%29%202.4/Theatre%20of%20Magic%20%28Bally%201995%29%202.4.vbs.original#L214) [[patch]](Theatre%20of%20Magic%20%28Bally%201995%29%202.4/Theatre%20of%20Magic%20%28Bally%201995%29%202.4.vbs.patch) [[patched]](Theatre%20of%20Magic%20%28Bally%201995%29%202.4/Theatre%20of%20Magic%20%28Bally%201995%29%202.4.vbs)
- Warriors The Full DMD 2.0 (Iceman 2023).vbs [[original]](Warriors%20The%20Full%20DMD%202.0%20%28Iceman%202023%29/Warriors%20The%20Full%20DMD%202.0%20%28Iceman%202023%29.vbs.original#L14) [[patch]](Warriors%20The%20Full%20DMD%202.0%20%28Iceman%202023%29/Warriors%20The%20Full%20DMD%202.0%20%28Iceman%202023%29.vbs.patch) [[patched]](Warriors%20The%20Full%20DMD%202.0%20%28Iceman%202023%29/Warriors%20The%20Full%20DMD%202.0%20%28Iceman%202023%29.vbs)
- Wednesday (Netflix 2023).vbs [[original]](Wednesday%20%28Netflix%202023%29/Wednesday%20%28Netflix%202023%29.vbs.original#L19) [[patch]](Wednesday%20%28Netflix%202023%29/Wednesday%20%28Netflix%202023%29.vbs.patch) [[patched]](Wednesday%20%28Netflix%202023%29/Wednesday%20%28Netflix%202023%29.vbs)

## GetDMDColor function call removed (not available in standalone)

**Affected patches: 8**

`GetDMDColor` is a helper function that does not exist in all script environments. Calling it causes a runtime error on standalone. The call is removed or commented out.

Affected scripts:

- DarkPrincess1.3.1.vbs [[original]](DarkPrincess1.3.1/DarkPrincess1.3.1.vbs.original#L161) [[patch]](DarkPrincess1.3.1/DarkPrincess1.3.1.vbs.patch) [[patched]](DarkPrincess1.3.1/DarkPrincess1.3.1.vbs)
- Metal Slug_1.05.vbs [[original]](Metal%20Slug_1.05/Metal%20Slug_1.05.vbs.original#L72) [[patch]](Metal%20Slug_1.05/Metal%20Slug_1.05.vbs.patch) [[patched]](Metal%20Slug_1.05/Metal%20Slug_1.05.vbs)
- Motley Crue SS.vbs [[original]](Motley%20Crue%20SS/Motley%20Crue%20SS.vbs.original#L5229) [[patch]](Motley%20Crue%20SS/Motley%20Crue%20SS.vbs.patch) [[patched]](Motley%20Crue%20SS/Motley%20Crue%20SS.vbs)
- Sons Of Anarchy (Original 2019).vbs [[original]](Sons%20Of%20Anarchy%20%28Original%202019%29/Sons%20Of%20Anarchy%20%28Original%202019%29.vbs.original#L51) [[patch]](Sons%20Of%20Anarchy%20%28Original%202019%29/Sons%20Of%20Anarchy%20%28Original%202019%29.vbs.patch) [[patched]](Sons%20Of%20Anarchy%20%28Original%202019%29/Sons%20Of%20Anarchy%20%28Original%202019%29.vbs)
- Spooky Wednesday Pro.vbs [[original]](Spooky%20Wednesday%20Pro/Spooky%20Wednesday%20Pro.vbs.original#L46) [[patch]](Spooky%20Wednesday%20Pro/Spooky%20Wednesday%20Pro.vbs.patch) [[patched]](Spooky%20Wednesday%20Pro/Spooky%20Wednesday%20Pro.vbs)
- Spooky_Wednesday VPX 2024.vbs [[original]](Spooky_Wednesday%20VPX%202024/Spooky_Wednesday%20VPX%202024.vbs.original#L46) [[patch]](Spooky_Wednesday%20VPX%202024/Spooky_Wednesday%20VPX%202024.vbs.patch) [[patched]](Spooky_Wednesday%20VPX%202024/Spooky_Wednesday%20VPX%202024.vbs)
- Streets of Rage (TBA 2018).vbs [[original]](Streets%20of%20Rage%20%28TBA%202018%29/Streets%20of%20Rage%20%28TBA%202018%29.vbs.original#L72) [[patch]](Streets%20of%20Rage%20%28TBA%202018%29/Streets%20of%20Rage%20%28TBA%202018%29.vbs.patch) [[patched]](Streets%20of%20Rage%20%28TBA%202018%29/Streets%20of%20Rage%20%28TBA%202018%29.vbs)
- Three Angels (Original 2018) LW.vbs [[original]](Three%20Angels%20%28Original%202018%29%20LW/Three%20Angels%20%28Original%202018%29%20LW.vbs.original#L335) [[patch]](Three%20Angels%20%28Original%202018%29%20LW/Three%20Angels%20%28Original%202018%29%20LW.vbs.patch) [[patched]](Three%20Angels%20%28Original%202018%29%20LW/Three%20Angels%20%28Original%202018%29%20LW.vbs)

## Execute/ExecuteGlobal with object Set in loop (replaced with IsObject check)

**Affected patches: 6**

`Execute "Set Lights(" & i & ") = L" & i` creates object references dynamically in a loop. In Wine's VBScript, this pattern fails when some loop variables are not objects. Fix: replace with an `If IsObject(eval("L" & i)) Then` guard before assigning.

Affected scripts:

- 1455577933_MedievalMadness_Upgrade(Real_Final).vbs [[original]](1455577933_MedievalMadness_Upgrade%28Real_Final%29/1455577933_MedievalMadness_Upgrade%28Real_Final%29.vbs.original#L2185) [[patch]](1455577933_MedievalMadness_Upgrade%28Real_Final%29/1455577933_MedievalMadness_Upgrade%28Real_Final%29.vbs.patch) [[patched]](1455577933_MedievalMadness_Upgrade%28Real_Final%29/1455577933_MedievalMadness_Upgrade%28Real_Final%29.vbs)
- Bram Stokers Dracula (1993).vbs [[original]](Bram%20Stoker%27s%20Dracula%20%28Williams%201993%29/Bram%20Stokers%20Dracula%20%281993%29.vbs.original#L161) [[patch]](Bram%20Stoker%27s%20Dracula%20%28Williams%201993%29/Bram%20Stokers%20Dracula%20%281993%29.vbs.patch) [[patched]](Bram%20Stoker%27s%20Dracula%20%28Williams%201993%29/Bram%20Stokers%20Dracula%20%281993%29.vbs)
- Cirqus_Voltaire_Hanibal-Mod_3.7.vbs [[original]](Cirqus_Voltaire_Hanibal-Mod_3.7/Cirqus_Voltaire_Hanibal-Mod_3.7.vbs.original#L1282) [[patch]](Cirqus_Voltaire_Hanibal-Mod_3.7/Cirqus_Voltaire_Hanibal-Mod_3.7.vbs.patch) [[patched]](Cirqus_Voltaire_Hanibal-Mod_3.7/Cirqus_Voltaire_Hanibal-Mod_3.7.vbs)
- VP10-Terminator 3 (Stern 2003) Hanibal v1.5.vbs [[original]](Terminator%203%20%28Stern%202003%29%20Hanibal%20v1.5/VP10-Terminator%203%20%28Stern%202003%29%20Hanibal%20v1.5.vbs.original#L1236) [[patch]](Terminator%203%20%28Stern%202003%29%20Hanibal%20v1.5/VP10-Terminator%203%20%28Stern%202003%29%20Hanibal%20v1.5.vbs.patch) [[patched]](Terminator%203%20%28Stern%202003%29%20Hanibal%20v1.5/VP10-Terminator%203%20%28Stern%202003%29%20Hanibal%20v1.5.vbs)
- Vampirella.vbs [[original]](Vampirella/Vampirella.vbs.original#L1) [[patch]](Vampirella/Vampirella.vbs.patch) [[patched]](Vampirella/Vampirella.vbs)
- Vampirella1 .3.vbs [[original]](Vampirella1%20.3/Vampirella1%20.3.vbs.original#L84) [[patch]](Vampirella1%20.3/Vampirella1%20.3.vbs.patch) [[patched]](Vampirella1%20.3/Vampirella1%20.3.vbs)

## Nested Sub definition extracted to top level

**Affected patches: 6**

VBScript does not allow a `Sub` to be defined inside another `Sub`. Scripts that contained nested Sub definitions (e.g. `chilloutthemusic` or `turnitbackup` nested inside the PuPlayer initialisation block) fail to compile. Fix: move the inner Sub to the top level of the script.

Affected scripts:

- 1679379865_jurassicparklimitededition.vbs [[original]](1679379865_jurassicparklimitededition/1679379865_jurassicparklimitededition.vbs.original#L493) [[patch]](1679379865_jurassicparklimitededition/1679379865_jurassicparklimitededition.vbs.patch) [[patched]](1679379865_jurassicparklimitededition/1679379865_jurassicparklimitededition.vbs)
- Stranger Things (Original 2020) LW.vbs [[original]](Stranger%20Things%20%28Original%202020%29%20LW%202.0.1/Stranger%20Things%20%28Original%202020%29%20LW.vbs.original#L361) [[patch]](Stranger%20Things%20%28Original%202020%29%20LW%202.0.1/Stranger%20Things%20%28Original%202020%29%20LW.vbs.patch) [[patched]](Stranger%20Things%20%28Original%202020%29%20LW%202.0.1/Stranger%20Things%20%28Original%202020%29%20LW.vbs)
- Stranger Things - SE 1.47_OSB.vbs [[original]](Stranger%20Things%20-%20SE%201.47_OSB/Stranger%20Things%20-%20SE%201.47_OSB.vbs.original#L358) [[patch]](Stranger%20Things%20-%20SE%201.47_OSB/Stranger%20Things%20-%20SE%201.47_OSB.vbs.patch) [[patched]](Stranger%20Things%20-%20SE%201.47_OSB/Stranger%20Things%20-%20SE%201.47_OSB.vbs)
- Stranger Things 4 Premium.vbs [[original]](Stranger%20Things%204%20Premium/Stranger%20Things%204%20Premium.vbs.original#L84) [[patch]](Stranger%20Things%204%20Premium/Stranger%20Things%204%20Premium.vbs.patch) [[patched]](Stranger%20Things%204%20Premium/Stranger%20Things%204%20Premium.vbs)
- The Beatles_007.vbs [[original]](The%20Beatles_007/The%20Beatles_007.vbs.original#L196) [[patch]](The%20Beatles_007/The%20Beatles_007.vbs.patch) [[patched]](The%20Beatles_007/The%20Beatles_007.vbs)
- tlk-0.35.vbs [[original]](tlk-0.35/tlk-0.35.vbs.original#L507) [[patch]](tlk-0.35/tlk-0.35.vbs.patch) [[patched]](tlk-0.35/tlk-0.35.vbs)

## gBOT replaced with getballs() function call

**Affected patches: 6**

The `gBOT` global variable holds the current ball collection in older VPX scripts. In newer versions it is replaced by the `getballs()` function. Scripts using `gBOT` are updated to call `getballs()` instead.

Affected scripts:

- Atlantis (Bally 1989) w VR Room v2.0.vbs [[original]](Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0/Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0.vbs.original#L3029) [[patch]](Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0/Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0.vbs.patch) [[patched]](Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0/Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0.vbs)
- Black Knight 2000 (Williams 1989) w VR Room v2.0.2.vbs [[original]](Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2/Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2.vbs.original#L2687) [[patch]](Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2/Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2.vbs.patch) [[patched]](Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2/Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2.vbs)
- Cactus Jacks (Gottlieb 1991) w VR Room v2.0.2.vbs [[original]](Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2/Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2.vbs.original#L1544) [[patch]](Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2/Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2.vbs.patch) [[patched]](Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2/Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2.vbs)
- Iron Man Vault Edition (Stern 2010) VPW v1.0.vbs [[original]](Iron%20Man%20Vault%20Edition%20%28Stern%202010%29%20VPW%20v1.0/Iron%20Man%20Vault%20Edition%20%28Stern%202010%29%20VPW%20v1.0.vbs.original#L3280) [[patch]](Iron%20Man%20Vault%20Edition%20%28Stern%202010%29%20VPW%20v1.0/Iron%20Man%20Vault%20Edition%20%28Stern%202010%29%20VPW%20v1.0.vbs.patch) [[patched]](Iron%20Man%20Vault%20Edition%20%28Stern%202010%29%20VPW%20v1.0/Iron%20Man%20Vault%20Edition%20%28Stern%202010%29%20VPW%20v1.0.vbs)
- Laser War (Data East 1987) w VR Room v2.0.vbs [[original]](Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0/Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0.vbs.original#L1409) [[patch]](Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0/Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0.vbs.patch) [[patched]](Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0/Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0.vbs)
- X-Men LE (Stern 2012) VPW v1.0.vbs [[original]](X-Men%20LE%20%28Stern%202012%29%20VPW%20v1.0/X-Men%20LE%20%28Stern%202012%29%20VPW%20v1.0.vbs.original#L3214) [[patch]](X-Men%20LE%20%28Stern%202012%29%20VPW%20v1.0/X-Men%20LE%20%28Stern%202012%29%20VPW%20v1.0.vbs.patch) [[patched]](X-Men%20LE%20%28Stern%202012%29%20VPW%20v1.0/X-Men%20LE%20%28Stern%202012%29%20VPW%20v1.0.vbs)

## File path case sensitivity fix (Linux filesystem)

**Affected patches: 6**

Linux file systems are case-sensitive, unlike Windows NTFS. File names embedded in scripts (images, videos, ROM file names) must match the actual file names exactly. Fix: correct the capitalisation of the file name string.

Affected scripts:

- Beavis and Butt-head_Pinballed.vbs [[original]](Beavis%20and%20Butt-head_Pinballed/Beavis%20and%20Butt-head_Pinballed.vbs.original#L51) [[patch]](Beavis%20and%20Butt-head_Pinballed/Beavis%20and%20Butt-head_Pinballed.vbs.patch) [[patched]](Beavis%20and%20Butt-head_Pinballed/Beavis%20and%20Butt-head_Pinballed.vbs)
- Blood Machines 2.0.vbs [[original]](Blood%20Machines%202.0/Blood%20Machines%202.0.vbs.original#L205) [[patch]](Blood%20Machines%202.0/Blood%20Machines%202.0.vbs.patch) [[patched]](Blood%20Machines%202.0/Blood%20Machines%202.0.vbs)
- DarkPrincess1.3.1.vbs [[original]](DarkPrincess1.3.1/DarkPrincess1.3.1.vbs.original#L161) [[patch]](DarkPrincess1.3.1/DarkPrincess1.3.1.vbs.patch) [[patched]](DarkPrincess1.3.1/DarkPrincess1.3.1.vbs)
- StarWars BountyHunter 3.02.vbs [[original]](StarWars%20BountyHunter%203.02/StarWars%20BountyHunter%203.02.vbs.original#L55) [[patch]](StarWars%20BountyHunter%203.02/StarWars%20BountyHunter%203.02.vbs.patch) [[patched]](StarWars%20BountyHunter%203.02/StarWars%20BountyHunter%203.02.vbs)
- Thunderbirds original 2022 v1.0.2.vbs [[original]](Thunderbirds%20original%202022%20v1.0.2/Thunderbirds%20original%202022%20v1.0.2.vbs.original#L115) [[patch]](Thunderbirds%20original%202022%20v1.0.2/Thunderbirds%20original%202022%20v1.0.2.vbs.patch) [[patched]](Thunderbirds%20original%202022%20v1.0.2/Thunderbirds%20original%202022%20v1.0.2.vbs)
- UT99CTF_GE_2.3.vbs [[original]](UT99CTF_GE_2.3/UT99CTF_GE_2.3.vbs.original#L1203) [[patch]](UT99CTF_GE_2.3/UT99CTF_GE_2.3.vbs.patch) [[patched]](UT99CTF_GE_2.3/UT99CTF_GE_2.3.vbs)

## Incorrect boolean expression (IsGIOn <> Not IsOff)

**Affected patches: 5**

`isGIOn <> Not IsOff` is logically incorrect — `Not IsOff` negates the variable, producing a boolean, and `<>` then compares it to `isGIOn`. Under Wine this expression evaluates incorrectly. Fix: use a simple assignment or comparison without the erroneous `<>` operator.

Affected scripts:

- AceOfSpeed.vbs [[original]](AceOfSpeed/AceOfSpeed.vbs.original#L245) [[patch]](AceOfSpeed/AceOfSpeed.vbs.patch) [[patched]](AceOfSpeed/AceOfSpeed.vbs)
- Death Proof Balutito V2.vbs [[original]](Death%20Proof%20Balutito%20V2/Death%20Proof%20Balutito%20V2.vbs.original#L207) [[patch]](Death%20Proof%20Balutito%20V2/Death%20Proof%20Balutito%20V2.vbs.patch) [[patched]](Death%20Proof%20Balutito%20V2/Death%20Proof%20Balutito%20V2.vbs)
- Iron Maiden Virtual Time (Original 2020).vbs [[original]](Iron%20Maiden%20Virtual%20Time%20%28Original%202020%29/Iron%20Maiden%20Virtual%20Time%20%28Original%202020%29.vbs.original#L408) [[patch]](Iron%20Maiden%20Virtual%20Time%20%28Original%202020%29/Iron%20Maiden%20Virtual%20Time%20%28Original%202020%29.vbs.patch) [[patched]](Iron%20Maiden%20Virtual%20Time%20%28Original%202020%29/Iron%20Maiden%20Virtual%20Time%20%28Original%202020%29.vbs)
- Laser War (Data East 1987) w VR Room v2.0.vbs [[original]](Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0/Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0.vbs.original#L1409) [[patch]](Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0/Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0.vbs.patch) [[patched]](Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0/Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0.vbs)
- Mousin' Around! (Bally 1989) w VR Room v2.1.vbs [[original]](Mousin%27%20Around%21%20%28Bally%201989%29%20w%20VR%20Room%20v2.1/Mousin%27%20Around%21%20%28Bally%201989%29%20w%20VR%20Room%20v2.1.vbs.original#L276) [[patch]](Mousin%27%20Around%21%20%28Bally%201989%29%20w%20VR%20Room%20v2.1/Mousin%27%20Around%21%20%28Bally%201989%29%20w%20VR%20Room%20v2.1.vbs.patch) [[patched]](Mousin%27%20Around%21%20%28Bally%201989%29%20w%20VR%20Room%20v2.1/Mousin%27%20Around%21%20%28Bally%201989%29%20w%20VR%20Room%20v2.1.vbs)

## Configuration constant value changed for standalone compatibility

**Affected patches: 5**

Certain constants (e.g. `cController`, `PupScreenMiniGame`, `Mute_Sound_For_PuPPack`) have default values that assume a full Windows VPX environment. For standalone builds these values need to be adjusted — for example switching from VPinMAME controller to B2S server, or disabling features not available in standalone.

Affected scripts:

- Bart VS the Space Mutants 1.1.vbs [[original]](Bart%20VS%20the%20Space%20Mutants%201.1/Bart%20VS%20the%20Space%20Mutants%201.1.vbs.original#L33) [[patch]](Bart%20VS%20the%20Space%20Mutants%201.1/Bart%20VS%20the%20Space%20Mutants%201.1.vbs.patch) [[patched]](Bart%20VS%20the%20Space%20Mutants%201.1/Bart%20VS%20the%20Space%20Mutants%201.1.vbs)
- [Centaur (Bally 1981).vbs.dmd](Centaur%20%28Bally%201981%29/Centaur%20%28Bally%201981%29.vbs.dmd#L6)
- Dream Daddy1.5.vbs [[original]](Dream%20Daddy1.5/Dream%20Daddy1.5.vbs.original#L69) [[patch]](Dream%20Daddy1.5/Dream%20Daddy1.5.vbs.patch) [[patched]](Dream%20Daddy1.5/Dream%20Daddy1.5.vbs)
- Heavy Metal Meltdown (Bally 1987).vbs [[original]](Heavy%20Metal%20Meltdown%20%28Bally%201987%29/Heavy%20Metal%20Meltdown%20%28Bally%201987%29.vbs.original#L12) [[patch]](Heavy%20Metal%20Meltdown%20%28Bally%201987%29/Heavy%20Metal%20Meltdown%20%28Bally%201987%29.vbs.patch) [[patched]](Heavy%20Metal%20Meltdown%20%28Bally%201987%29/Heavy%20Metal%20Meltdown%20%28Bally%201987%29.vbs)
- [Megadeth (original).vbs.dmd](Megadeth%20%28original%29/Megadeth%20%28original%29.vbs.dmd#L66)

## Sub/Function definition with inline code (no newline after definition)

**Affected patches: 5**

Placing code on the same line as the `Sub` or `Function` header (e.g. `Sub Sw18_Hit()Controller.Switch(18)=1`) is valid on Windows VBScript but fails to parse correctly under Wine. Fix: add a newline after the `()` so the body starts on the next line.

Affected scripts:

- Bugs Bunny's Birthday Ball (Bally 1991)Rev2.3b.vbs [[original]](Bugs%20Bunny%27s%20Birthday%20Ball%20%28Bally%201991%29Rev2.3b/Bugs%20Bunny%27s%20Birthday%20Ball%20%28Bally%201991%29Rev2.3b.vbs.original#L798) [[patch]](Bugs%20Bunny%27s%20Birthday%20Ball%20%28Bally%201991%29Rev2.3b/Bugs%20Bunny%27s%20Birthday%20Ball%20%28Bally%201991%29Rev2.3b.vbs.patch) [[patched]](Bugs%20Bunny%27s%20Birthday%20Ball%20%28Bally%201991%29Rev2.3b/Bugs%20Bunny%27s%20Birthday%20Ball%20%28Bally%201991%29Rev2.3b.vbs)
- Cue Ball Wizard (Gottlieb 1992) v1.1.2.vbs [[original]](Cue%20Ball%20Wizard%20%28Gottlieb%201992%29%20v1.1.2/Cue%20Ball%20Wizard%20%28Gottlieb%201992%29%20v1.1.2.vbs.original#L1503) [[patch]](Cue%20Ball%20Wizard%20%28Gottlieb%201992%29%20v1.1.2/Cue%20Ball%20Wizard%20%28Gottlieb%201992%29%20v1.1.2.vbs.patch) [[patched]](Cue%20Ball%20Wizard%20%28Gottlieb%201992%29%20v1.1.2/Cue%20Ball%20Wizard%20%28Gottlieb%201992%29%20v1.1.2.vbs)
- Haunted House (Gottlieb 1982) 1.0.vbs [[original]](Haunted%20House%20%28Gottlieb%201982%29%201.0/Haunted%20House%20%28Gottlieb%201982%29%201.0.vbs.original#L2980) [[patch]](Haunted%20House%20%28Gottlieb%201982%29%201.0/Haunted%20House%20%28Gottlieb%201982%29%201.0.vbs.patch) [[patched]](Haunted%20House%20%28Gottlieb%201982%29%201.0/Haunted%20House%20%28Gottlieb%201982%29%201.0.vbs)
- Haunted House (Gottlieb 1982)_Bigus(MOD)1.0.vbs [[original]](Haunted%20House%20%28Gottlieb%201982%29_Bigus%28MOD%291.0/Haunted%20House%20%28Gottlieb%201982%29_Bigus%28MOD%291.0.vbs.original#L2970) [[patch]](Haunted%20House%20%28Gottlieb%201982%29_Bigus%28MOD%291.0/Haunted%20House%20%28Gottlieb%201982%29_Bigus%28MOD%291.0.vbs.patch) [[patched]](Haunted%20House%20%28Gottlieb%201982%29_Bigus%28MOD%291.0/Haunted%20House%20%28Gottlieb%201982%29_Bigus%28MOD%291.0.vbs)
- Scared Stiff (Bally 1996) v1.29.vbs [[original]](Scared%20Stiff%20%28Bally%201996%29%20v1.29/Scared%20Stiff%20%28Bally%201996%29%20v1.29.vbs.original#L466) [[patch]](Scared%20Stiff%20%28Bally%201996%29%20v1.29/Scared%20Stiff%20%28Bally%201996%29%20v1.29.vbs.patch) [[patched]](Scared%20Stiff%20%28Bally%201996%29%20v1.29/Scared%20Stiff%20%28Bally%201996%29%20v1.29.vbs)

## For Each loop variable reused (causes error in Wine)

**Affected patches: 4**

Reusing the same loop variable in nested or back-to-back `For Each` loops causes a runtime error in Wine's VBScript. Fix: use a different variable name for each loop level.

Affected scripts:

- Atlantis (Bally 1989) w VR Room v2.0.vbs [[original]](Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0/Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0.vbs.original#L3029) [[patch]](Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0/Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0.vbs.patch) [[patched]](Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0/Atlantis%20%28Bally%201989%29%20w%20VR%20Room%20v2.0.vbs)
- Black Knight 2000 (Williams 1989) w VR Room v2.0.2.vbs [[original]](Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2/Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2.vbs.original#L2687) [[patch]](Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2/Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2.vbs.patch) [[patched]](Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2/Black%20Knight%202000%20%28Williams%201989%29%20w%20VR%20Room%20v2.0.2.vbs)
- Cactus Jacks (Gottlieb 1991) w VR Room v2.0.2.vbs [[original]](Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2/Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2.vbs.original#L1544) [[patch]](Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2/Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2.vbs.patch) [[patched]](Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2/Cactus%20Jacks%20%28Gottlieb%201991%29%20w%20VR%20Room%20v2.0.2.vbs)
- Laser War (Data East 1987) w VR Room v2.0.vbs [[original]](Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0/Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0.vbs.original#L1409) [[patch]](Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0/Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0.vbs.patch) [[patched]](Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0/Laser%20War%20%28Data%20East%201987%29%20w%20VR%20Room%20v2.0.vbs)

## UseVPMDMD variable renamed to UseVPMColoredDMD

**Affected patches: 4**

Scripts that use a colored DMD must use the variable name `UseVPMColoredDMD` (not `UseVPMDMD`) so that the VPX standalone renderer applies the correct colour palette.

Affected scripts:

- [Bad Cats (Williams 1989) VPW 3.0.vbs.dmd](Bad%20Cats%20%28Williams%201989%29%20VPW%203.0/Bad%20Cats%20%28Williams%201989%29%20VPW%203.0.vbs.dmdcolored#L384)
- Champion Pub (Williams 1998) v1.43.vbs
- Cue Ball Wizard (Gottlieb 1992) v1.2.4.vbs
- Tron Legacy (Stern 2011) VPW Mod v1.1.vbs

## typename() casing fix (typename → TypeName)

**Affected patches: 4**

Wine's VBScript is stricter about built-in function capitalisation. `typename(x)` must be `TypeName(x)`.

Affected scripts:

- Elvis_MOD_2.0.vbs [[original]](Elvis_MOD_2.0/Elvis_MOD_2.0.vbs.original#L626) [[patch]](Elvis_MOD_2.0/Elvis_MOD_2.0.vbs.patch) [[patched]](Elvis_MOD_2.0/Elvis_MOD_2.0.vbs)
- Simpsons Pinball Party, The (Stern 2003) VPW 2.0.3.vbs [[original]](Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3/Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3.vbs.original#L3332) [[patch]](Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3/Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3.vbs.patch) [[patched]](Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3/Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3.vbs)
- Simpsons Treehouse of Horror MOD v2023.3 (Simpsons Pinball Party, The (Stern 2003) VPW 2.0.3 base).vbs [[original]](Simpsons%20Treehouse%20of%20Horror%20MOD%20v2023.3%20%28Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3%20base%29/Simpsons%20Treehouse%20of%20Horror%20MOD%20v2023.3%20%28Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3%20base%29.vbs.original#L95) [[patch]](Simpsons%20Treehouse%20of%20Horror%20MOD%20v2023.3%20%28Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3%20base%29/Simpsons%20Treehouse%20of%20Horror%20MOD%20v2023.3%20%28Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3%20base%29.vbs.patch) [[patched]](Simpsons%20Treehouse%20of%20Horror%20MOD%20v2023.3%20%28Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3%20base%29/Simpsons%20Treehouse%20of%20Horror%20MOD%20v2023.3%20%28Simpsons%20Pinball%20Party%2C%20The%20%28Stern%202003%29%20VPW%202.0.3%20base%29.vbs)
- TX-Sector (Gottlieb 1988) SG1bsoN Mod V1.1.vbs [[original]](TX-Sector%20%28Gottlieb%201988%29%20SG1bsoN%20Mod%20V1.1/TX-Sector%20%28Gottlieb%201988%29%20SG1bsoN%20Mod%20V1.1.vbs.original#L809) [[patch]](TX-Sector%20%28Gottlieb%201988%29%20SG1bsoN%20Mod%20V1.1/TX-Sector%20%28Gottlieb%201988%29%20SG1bsoN%20Mod%20V1.1.vbs.patch) [[patched]](TX-Sector%20%28Gottlieb%201988%29%20SG1bsoN%20Mod%20V1.1/TX-Sector%20%28Gottlieb%201988%29%20SG1bsoN%20Mod%20V1.1.vbs)

## Statement starting with colon (invalid VBScript syntax)

**Affected patches: 4**

Lines that begin with `:` followed by a statement (e.g. `        : Controller.B2SSetData 9, 0`) are not valid VBScript syntax — the colon is a statement separator, not a line prefix. Fix: remove the leading colon.

Affected scripts:

- Junkyard Cats_1.07.vbs [[original]](Junkyard%20Cats_1.07/Junkyard%20Cats_1.07.vbs.original#L508) [[patch]](Junkyard%20Cats_1.07/Junkyard%20Cats_1.07.vbs.patch) [[patched]](Junkyard%20Cats_1.07/Junkyard%20Cats_1.07.vbs)
- NBA (Stern 2009)_Bigus(MOD)1.4.vbs [[original]](NBA%20%28Stern%202009%29_Bigus%28MOD%291.4/NBA%20%28Stern%202009%29_Bigus%28MOD%291.4.vbs.original#L279) [[patch]](NBA%20%28Stern%202009%29_Bigus%28MOD%291.4/NBA%20%28Stern%202009%29_Bigus%28MOD%291.4.vbs.patch) [[patched]](NBA%20%28Stern%202009%29_Bigus%28MOD%291.4/NBA%20%28Stern%202009%29_Bigus%28MOD%291.4.vbs)
- Scarface - Balls and Power 1.2.vbs [[original]](Scarface%20-%20Balls%20and%20Power%201.2/Scarface%20-%20Balls%20and%20Power%201.2.vbs.original#L1000) [[patch]](Scarface%20-%20Balls%20and%20Power%201.2/Scarface%20-%20Balls%20and%20Power%201.2.vbs.patch) [[patched]](Scarface%20-%20Balls%20and%20Power%201.2/Scarface%20-%20Balls%20and%20Power%201.2.vbs)
- The Fifth Element 1.3.vbs [[original]](The%20Fifth%20Element%201.3/The%20Fifth%20Element%201.3.vbs.original#L478) [[patch]](The%20Fifth%20Element%201.3/The%20Fifth%20Element%201.3.vbs.patch) [[patched]](The%20Fifth%20Element%201.3/The%20Fifth%20Element%201.3.vbs)

## UBound casing fix (Ubound → UBound)

**Affected patches: 3**

Wine's VBScript is stricter about built-in function capitalisation. `Ubound` must be `UBound`.

Affected scripts:

- Monkey Island VR Room.vbs [[original]](Monkey%20Island%20VR%20Room/Monkey%20Island%20VR%20Room.vbs.original#L240) [[patch]](Monkey%20Island%20VR%20Room/Monkey%20Island%20VR%20Room.vbs.patch) [[patched]](Monkey%20Island%20VR%20Room/Monkey%20Island%20VR%20Room.vbs)
- Power Play (Bally 1977).vbs [[original]](1256692067_PowerPlay%28Bally1977%292.1/Power%20Play%20%28Bally%201977%29.vbs.original#L922) [[patch]](1256692067_PowerPlay%28Bally1977%292.1/Power%20Play%20%28Bally%201977%29.vbs.patch) [[patched]](1256692067_PowerPlay%28Bally1977%292.1/Power%20Play%20%28Bally%201977%29.vbs)
- monkeyisland.vbs [[original]](Monkeyislandv1.1/monkeyisland.vbs.original#L2441) [[patch]](Monkeyislandv1.1/monkeyisland.vbs.patch) [[patched]](Monkeyislandv1.1/monkeyisland.vbs)

## Double-dot in string callback (e.g. "obj..Method")

**Affected patches: 3**

SolCallback strings like `"dtbank..SolHit"` contain a double dot, which is invalid. The double dot arises when the object name is accidentally repeated. Fix: correct to a single dot, e.g. `"dtbank.SolHit"`.

Affected scripts:

- ABBAv2.0.vbs [[original]](ABBAv2.0/ABBAv2.0.vbs.original#L47) [[patch]](ABBAv2.0/ABBAv2.0.vbs.patch) [[patched]](ABBAv2.0/ABBAv2.0.vbs)
- Alfred Hitchcock's Psycho (TBA 2019).vbs [[original]](Alfred%20Hitchcock%27s%20Psycho%20%28TBA%202019%29/Alfred%20Hitchcock%27s%20Psycho%20%28TBA%202019%29.vbs.original#L292) [[patch]](Alfred%20Hitchcock%27s%20Psycho%20%28TBA%202019%29/Alfred%20Hitchcock%27s%20Psycho%20%28TBA%202019%29.vbs.patch) [[patched]](Alfred%20Hitchcock%27s%20Psycho%20%28TBA%202019%29/Alfred%20Hitchcock%27s%20Psycho%20%28TBA%202019%29.vbs)
- Vortex (Taito do Brasil - 1981) v4.vbs [[original]](Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4/Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4.vbs.original#L46) [[patch]](Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4/Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4.vbs.patch) [[patched]](Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4/Vortex%20%28Taito%20do%20Brasil%20-%201981%29%20v4.vbs)

## Hex literal with excess leading zeros (e.g. &H000000031)

**Affected patches: 3**

A hex literal with an odd or excessive number of digits (e.g. `&H000000031` — nine hex digits) is rejected by Wine's VBScript parser. Fix: reduce to the correct even number of digits (e.g. `&H00000031`).

Affected scripts:

- Andromeda (Game Plan 1985) v4.vbs [[original]](Andromeda%20%28Game%20Plan%201985%29%20v4/Andromeda%20%28Game%20Plan%201985%29%20v4.vbs.original#L922) [[patch]](Andromeda%20%28Game%20Plan%201985%29%20v4/Andromeda%20%28Game%20Plan%201985%29%20v4.vbs.patch) [[patched]](Andromeda%20%28Game%20Plan%201985%29%20v4/Andromeda%20%28Game%20Plan%201985%29%20v4.vbs)
- Grand Slam (Bally 1983) 2.3.vbs [[original]](Grand%20Slam%20%28Bally%201983%29%202.3/Grand%20Slam%20%28Bally%201983%29%202.3.vbs.original#L791) [[patch]](Grand%20Slam%20%28Bally%201983%29%202.3/Grand%20Slam%20%28Bally%201983%29%202.3.vbs.patch) [[patched]](Grand%20Slam%20%28Bally%201983%29%202.3/Grand%20Slam%20%28Bally%201983%29%202.3.vbs)
- Topaz (Inder 1979).vbs [[original]](Topaz%20%28Inder%201979%29/Topaz%20%28Inder%201979%29.vbs.original#L100) [[patch]](Topaz%20%28Inder%201979%29/Topaz%20%28Inder%201979%29.vbs.patch) [[patched]](Topaz%20%28Inder%201979%29/Topaz%20%28Inder%201979%29.vbs)

## Me(Idx) collection indexer replaced with named collection

**Affected patches: 3**

Inside a VBScript `Class`, `Me(n)` is not a valid way to index a collection under Wine. Fix: use the named array/collection directly, e.g. `Spinner(Idx).IsDropped`.

Affected scripts:

- Four Seasons (Gottlieb 1968).vbs [[original]](Four%20Seasons%20%28Gottlieb%201968%29/Four%20Seasons%20%28Gottlieb%201968%29.vbs.original#L1214) [[patch]](Four%20Seasons%20%28Gottlieb%201968%29/Four%20Seasons%20%28Gottlieb%201968%29.vbs.patch) [[patched]](Four%20Seasons%20%28Gottlieb%201968%29/Four%20Seasons%20%28Gottlieb%201968%29.vbs)
- Four Seasons (Gottlieb 1968)_Teisen_MOD.vbs [[original]](Four%20Seasons%20%28Gottlieb%201968%29_Teisen_MOD/Four%20Seasons%20%28Gottlieb%201968%29_Teisen_MOD.vbs.original#L1219) [[patch]](Four%20Seasons%20%28Gottlieb%201968%29_Teisen_MOD/Four%20Seasons%20%28Gottlieb%201968%29_Teisen_MOD.vbs.patch) [[patched]](Four%20Seasons%20%28Gottlieb%201968%29_Teisen_MOD/Four%20Seasons%20%28Gottlieb%201968%29_Teisen_MOD.vbs)
- Roller Coaster (Gottlieb 1971)x.vbs [[original]](Roller%20Coaster%20%28Gottlieb%201971%29x/Roller%20Coaster%20%28Gottlieb%201971%29x.vbs.original#L1037) [[patch]](Roller%20Coaster%20%28Gottlieb%201971%29x/Roller%20Coaster%20%28Gottlieb%201971%29x.vbs.patch) [[patched]](Roller%20Coaster%20%28Gottlieb%201971%29x/Roller%20Coaster%20%28Gottlieb%201971%29x.vbs)

## Trim.visible assignment removed (Trim is a VBScript reserved function)

**Affected patches: 3**

`Trim` is a built-in VBScript string function. Using it as an object name (e.g. `Trim.visible = 1`) is interpreted as a function call and raises a type error. Fix: comment out or rename the object.

Affected scripts:

- Gladiators (Premier 1993) v1.1.1.vbs [[original]](Gladiators%20%28Premier%201993%29%20v1.1.1/Gladiators%20%28Premier%201993%29%20v1.1.1.vbs.original#L344) [[patch]](Gladiators%20%28Premier%201993%29%20v1.1.1/Gladiators%20%28Premier%201993%29%20v1.1.1.vbs.patch) [[patched]](Gladiators%20%28Premier%201993%29%20v1.1.1/Gladiators%20%28Premier%201993%29%20v1.1.1.vbs)
- Surf'n Safari v1.3.4.vbs [[original]](Surf%27n%20Safari%20v1.3.4/Surf%27n%20Safari%20v1.3.4.vbs.original#L117) [[patch]](Surf%27n%20Safari%20v1.3.4/Surf%27n%20Safari%20v1.3.4.vbs.patch) [[patched]](Surf%27n%20Safari%20v1.3.4/Surf%27n%20Safari%20v1.3.4.vbs)
- Wipe Out (Premier 1993) 1.1.0.vbs [[original]](Wipe%20Out%20%28Premier%201993%29%201.1.0/Wipe%20Out%20%28Premier%201993%29%201.1.0.vbs.original#L258) [[patch]](Wipe%20Out%20%28Premier%201993%29%201.1.0/Wipe%20Out%20%28Premier%201993%29%201.1.0.vbs.patch) [[patched]](Wipe%20Out%20%28Premier%201993%29%201.1.0/Wipe%20Out%20%28Premier%201993%29%201.1.0.vbs)

## One-liner If/Then with dangling Else (missing End If)

**Affected patches: 3**

A single-line `If x Then y: Else` with nothing after the `Else` is syntactically invalid — VBScript expects an `End If` or an inline statement after `Else`. Fix: remove the trailing `: Else` or add `End If`.

Affected scripts:

- South Park (Sega 1999) 1.3.vbs [[original]](South%20Park%20%28Sega%201999%29%201.3/South%20Park%20%28Sega%201999%29%201.3.vbs.original#L1167) [[patch]](South%20Park%20%28Sega%201999%29%201.3/South%20Park%20%28Sega%201999%29%201.3.vbs.patch) [[patched]](South%20Park%20%28Sega%201999%29%201.3/South%20Park%20%28Sega%201999%29%201.3.vbs)
- South Park - Halloween.vbs [[original]](South%20Park%20-%20Halloween/South%20Park%20-%20Halloween.vbs.original#L49) [[patch]](South%20Park%20-%20Halloween/South%20Park%20-%20Halloween.vbs.patch) [[patched]](South%20Park%20-%20Halloween/South%20Park%20-%20Halloween.vbs)
- Starship Troopers (Sega 1997) VPW Mod v2.0.vbs [[original]](Starship%20Troopers%20%28Sega%201997%29%20VPW%20Mod%20v2.0/Starship%20Troopers%20%28Sega%201997%29%20VPW%20Mod%20v2.0.vbs.original#L1324) [[patch]](Starship%20Troopers%20%28Sega%201997%29%20VPW%20Mod%20v2.0/Starship%20Troopers%20%28Sega%201997%29%20VPW%20Mod%20v2.0.vbs.patch) [[patched]](Starship%20Troopers%20%28Sega%201997%29%20VPW%20Mod%20v2.0/Starship%20Troopers%20%28Sega%201997%29%20VPW%20Mod%20v2.0.vbs)

## Duplicate score function calls removed

**Affected patches: 2**

The same `AddScore` call was duplicated multiple times in the same event handler, causing the score to be added multiple times per activation. Fix: keep only one call.

Affected scripts:

- 2104398928_CARtoonsRC(Nailed2021)v1.3.vbs [[original]](2104398928_CARtoonsRC%28Nailed2021%29v1.3/2104398928_CARtoonsRC%28Nailed2021%29v1.3.vbs.original#L841) [[patch]](2104398928_CARtoonsRC%28Nailed2021%29v1.3/2104398928_CARtoonsRC%28Nailed2021%29v1.3.vbs.patch) [[patched]](2104398928_CARtoonsRC%28Nailed2021%29v1.3/2104398928_CARtoonsRC%28Nailed2021%29v1.3.vbs)
- Gemini (Gottlieb 1978).vbs [[original]](Gemini%20%28Gottlieb%201978%29/Gemini%20%28Gottlieb%201978%29.vbs.original#L1173) [[patch]](Gemini%20%28Gottlieb%201978%29/Gemini%20%28Gottlieb%201978%29.vbs.patch) [[patched]](Gemini%20%28Gottlieb%201978%29/Gemini%20%28Gottlieb%201978%29.vbs)

## 'default' reserved word conflict fixed (Const default = 0 added)

**Affected patches: 2**

Using `default` as a variable name conflicts with VBScript's reserved `Default` property keyword. Fix: declare `Const default = 0` at the top of the script to shadow the reserved word with a harmless constant.

Affected scripts:

- Bud and Terence (Original 2024) v1.6.vbs [[original]](Bud%20Spencer%20%26%20Terence%20Hill%20%28Original%202024%29/Bud%20and%20Terence%20%28Original%202024%29%20v1.6.vbs.original#L4386) [[patch]](Bud%20Spencer%20%26%20Terence%20Hill%20%28Original%202024%29/Bud%20and%20Terence%20%28Original%202024%29%20v1.6.vbs.patch) [[patched]](Bud%20Spencer%20%26%20Terence%20Hill%20%28Original%202024%29/Bud%20and%20Terence%20%28Original%202024%29%20v1.6.vbs)
- IT Pinball Madness (JP Salas,Joe Picasso)1.2.vbs [[original]](IT%20Pinball%20Madness%20%28JP%20Salas%2CJoe%20Picasso%291.2/IT%20Pinball%20Madness%20%28JP%20Salas%2CJoe%20Picasso%291.2.vbs.original#L4870) [[patch]](IT%20Pinball%20Madness%20%28JP%20Salas%2CJoe%20Picasso%291.2/IT%20Pinball%20Madness%20%28JP%20Salas%2CJoe%20Picasso%291.2.vbs.patch) [[patched]](IT%20Pinball%20Madness%20%28JP%20Salas%2CJoe%20Picasso%291.2/IT%20Pinball%20Madness%20%28JP%20Salas%2CJoe%20Picasso%291.2.vbs)

## BSize/BMass constant renamed to BallSize/BallMass

**Affected patches: 2**

`Const BSize` or `Const BMass` conflict with VPX's built-in `BallSize` and `BallMass` properties. Fix: rename the script constants to avoid the conflict.

Affected scripts:

- Cyclone (Williams 1988) 2.0.1.vbs [[original]](Cyclone%20%28Williams%201988%29%202.0.1/Cyclone%20%28Williams%201988%29%202.0.1.vbs.original#L5) [[patch]](Cyclone%20%28Williams%201988%29%202.0.1/Cyclone%20%28Williams%201988%29%202.0.1.vbs.patch) [[patched]](Cyclone%20%28Williams%201988%29%202.0.1/Cyclone%20%28Williams%201988%29%202.0.1.vbs)
- Space Jam (Sega 1996) 1.4 JPJ - Edizzle - TeamPP - JLou.vbs [[original]](Space%20Jam%20%28Sega%201996%29%201.4%20JPJ%20-%20Edizzle%20-%20TeamPP%20-%20JLou/Space%20Jam%20%28Sega%201996%29%201.4%20JPJ%20-%20Edizzle%20-%20TeamPP%20-%20JLou.vbs.original#L7) [[patch]](Space%20Jam%20%28Sega%201996%29%201.4%20JPJ%20-%20Edizzle%20-%20TeamPP%20-%20JLou/Space%20Jam%20%28Sega%201996%29%201.4%20JPJ%20-%20Edizzle%20-%20TeamPP%20-%20JLou.vbs.patch) [[patched]](Space%20Jam%20%28Sega%201996%29%201.4%20JPJ%20-%20Edizzle%20-%20TeamPP%20-%20JLou/Space%20Jam%20%28Sega%201996%29%201.4%20JPJ%20-%20Edizzle%20-%20TeamPP%20-%20JLou.vbs)

## OptionsLoad/OptionsEdit call removed (not available in standalone)

**Affected patches: 2**

`OptionsLoad` and `OptionsEdit` are helper functions provided by some VPX utilities but not available in standalone builds. Calling them causes a "Sub or Function not defined" error. Fix: remove the calls and inline the required initialisation.

Affected scripts:

- Gremlins by Balutito 1.7.vbs [[original]](Gremlins%20by%20Balutito%201.7/Gremlins%20by%20Balutito%201.7.vbs.original#L162) [[patch]](Gremlins%20by%20Balutito%201.7/Gremlins%20by%20Balutito%201.7.vbs.patch) [[patched]](Gremlins%20by%20Balutito%201.7/Gremlins%20by%20Balutito%201.7.vbs)
- Victory (Gottlieb 1987) 2.0.2.vbs [[original]](Victory%20%28Gottlieb%201987%29%202.0.2/Victory%20%28Gottlieb%201987%29%202.0.2.vbs.original#L179) [[patch]](Victory%20%28Gottlieb%201987%29%202.0.2/Victory%20%28Gottlieb%201987%29%202.0.2.vbs.patch) [[patched]](Victory%20%28Gottlieb%201987%29%202.0.2/Victory%20%28Gottlieb%201987%29%202.0.2.vbs)

## VPinMAME Settings.Value() call fix

**Affected patches: 2**

`Controller.Settings.Value("key")` must be called without extra parentheses around the argument in VBScript when the return value is discarded. Fix: use `Controller.Settings.Value "key"` or restructure the call.

Affected scripts:

- Phantom Of The Opera (Data East 1990) v1.24.vbs [[original]](Phantom%20Of%20The%20Opera%20%28Data%20East%201990%29%20v1.24/Phantom%20Of%20The%20Opera%20%28Data%20East%201990%29%20v1.24.vbs.original#L316) [[patch]](Phantom%20Of%20The%20Opera%20%28Data%20East%201990%29%20v1.24/Phantom%20Of%20The%20Opera%20%28Data%20East%201990%29%20v1.24.vbs.patch) [[patched]](Phantom%20Of%20The%20Opera%20%28Data%20East%201990%29%20v1.24/Phantom%20Of%20The%20Opera%20%28Data%20East%201990%29%20v1.24.vbs)
- The Phantom Of The Opera (Data East 1990).vbs [[original]](The%20Phantom%20Of%20The%20Opera%20%28Data%20East%201990%29/The%20Phantom%20Of%20The%20Opera%20%28Data%20East%201990%29.vbs.original#L135) [[patch]](The%20Phantom%20Of%20The%20Opera%20%28Data%20East%201990%29/The%20Phantom%20Of%20The%20Opera%20%28Data%20East%201990%29.vbs.patch) [[patched]](The%20Phantom%20Of%20The%20Opera%20%28Data%20East%201990%29/The%20Phantom%20Of%20The%20Opera%20%28Data%20East%201990%29.vbs)

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

- TimonPumbaa.vbs [[original]](TimonPumbaa/TimonPumbaa.vbs.original#L3548) [[patch]](TimonPumbaa/TimonPumbaa.vbs.patch) [[patched]](TimonPumbaa/TimonPumbaa.vbs)
- Wheel of Fortune (Stern 2007) 1.0.vbs [[original]](Wheel%20of%20Fortune%20%28Stern%202007%29%201.0/Wheel%20of%20Fortune%20%28Stern%202007%29%201.0.vbs.original#L10) [[patch]](Wheel%20of%20Fortune%20%28Stern%202007%29%201.0/Wheel%20of%20Fortune%20%28Stern%202007%29%201.0.vbs.patch) [[patched]](Wheel%20of%20Fortune%20%28Stern%202007%29%201.0/Wheel%20of%20Fortune%20%28Stern%202007%29%201.0.vbs)

## CDbl() replaced with CBool() (wrong type conversion)

**Affected patches: 1**

Affected scripts:

- A-Go-Go (Williams 1966).vbs [[original]](A-Go-Go%20%28Williams%201966%29/A-Go-Go%20%28Williams%201966%29.vbs.original#L1293) [[patch]](A-Go-Go%20%28Williams%201966%29/A-Go-Go%20%28Williams%201966%29.vbs.patch) [[patched]](A-Go-Go%20%28Williams%201966%29/A-Go-Go%20%28Williams%201966%29.vbs)

## DisplayTimer.Enabled = DesktopMode fix

**Affected patches: 1**

Setting `DisplayTimer.Enabled = DesktopMode` assigns a boolean to a timer's Enabled property. In VR mode `DesktopMode` is `False`, inadvertently disabling the timer. Fix: wrap in a conditional or evaluate correctly.

Affected scripts:

- AceOfSpeed.vbs [[original]](AceOfSpeed/AceOfSpeed.vbs.original#L245) [[patch]](AceOfSpeed/AceOfSpeed.vbs.patch) [[patched]](AceOfSpeed/AceOfSpeed.vbs)

## cvpmDictionary replaced with Scripting.Dictionary

**Affected patches: 1**

`cvpmDictionary` is an older VPM wrapper class. Scripts using `New cvpmDictionary` are updated to use `CreateObject("Scripting.Dictionary")` directly.

Affected scripts:

- Batman66_1.1.0.vbs [[original]](Batman66_1.1.0/Batman66_1.1.0.vbs.original#L93) [[patch]](Batman66_1.1.0/Batman66_1.1.0.vbs.patch) [[patched]](Batman66_1.1.0/Batman66_1.1.0.vbs)

## LinkedTo property must be assigned an Array (not a single object)

**Affected patches: 1**

The `DropTarget.LinkedTo` property expects an Array of linked targets, not a bare object reference. Fix: wrap the value in `Array(...)`, e.g. `dtR.LinkedTo = Array(dtL)`.

Affected scripts:

- Contact (Williams 1978)_Bigus(MOD)1.0.vbs [[original]](Contact%20%28Williams%201978%29_Bigus%28MOD%291.0/Contact%20%28Williams%201978%29_Bigus%28MOD%291.0.vbs.original#L170) [[patch]](Contact%20%28Williams%201978%29_Bigus%28MOD%291.0/Contact%20%28Williams%201978%29_Bigus%28MOD%291.0.vbs.patch) [[patched]](Contact%20%28Williams%201978%29_Bigus%28MOD%291.0/Contact%20%28Williams%201978%29_Bigus%28MOD%291.0.vbs)

## Const tnob value fix

**Affected patches: 1**

The `tnob` (total number of balls) constant had an incorrect value, causing ball tracking issues. Fix: set `Const tnob` to the correct number of simultaneous balls for the table.

Affected scripts:

- GalaxyPlay_1_2.vbs [[original]](GalaxyPlay_1_2/GalaxyPlay_1_2.vbs.original#L154) [[patch]](GalaxyPlay_1_2/GalaxyPlay_1_2.vbs.patch) [[patched]](GalaxyPlay_1_2/GalaxyPlay_1_2.vbs)

## Code missing newlines between statements (all on one line)

**Affected patches: 1**

Multiple statements were concatenated onto a single line without newlines, making the script extremely long and causing parse errors in Wine. Fix: insert newlines to separate each statement.

Affected scripts:

- Iron Maiden (Original 2022) VPW 1.0.12.vbs [[original]](Iron%20Maiden%20%28Original%202022%29%20VPW%201.0.12/Iron%20Maiden%20%28Original%202022%29%20VPW%201.0.12.vbs.original#L7957) [[patch]](Iron%20Maiden%20%28Original%202022%29%20VPW%201.0.12/Iron%20Maiden%20%28Original%202022%29%20VPW%201.0.12.vbs.patch) [[patched]](Iron%20Maiden%20%28Original%202022%29%20VPW%201.0.12/Iron%20Maiden%20%28Original%202022%29%20VPW%201.0.12.vbs)

## b2s.vbs GetTextFile replaced with inline B2S helper functions

**Affected patches: 1**

`ExecuteGlobal GetTextFile("b2s.vbs")` loads a shared B2S helper script at runtime. This approach fails in standalone because the helper file is not in the expected path. Fix: inline the needed B2S helper functions (`SetB2SData`, `StepB2SData`, etc.) directly into the script.

Affected scripts:

- Jungle_Queen.vbs [[original]](Jungle_Queen/Jungle_Queen.vbs.original#L15) [[patch]](Jungle_Queen/Jungle_Queen.vbs.patch) [[patched]](Jungle_Queen/Jungle_Queen.vbs)

## Variable value or array size changed for standalone compatibility

**Affected patches: 1**

A variable value or array dimension is changed to a value more appropriate for standalone play — for example, increasing a `LampState` array size to accommodate all lamps, or adjusting a physics limit value.

Affected scripts:

- Junk Yard (Williams 1996) v1.83.vbs [[original]](Junk%20Yard%20%28Williams%201996%29%20v1.83/Junk%20Yard%20%28Williams%201996%29%20v1.83.vbs.original#L887) [[patch]](Junk%20Yard%20%28Williams%201996%29%20v1.83/Junk%20Yard%20%28Williams%201996%29%20v1.83.vbs.patch) [[patched]](Junk%20Yard%20%28Williams%201996%29%20v1.83/Junk%20Yard%20%28Williams%201996%29%20v1.83.vbs)

## Controller.Pause missing assignment (should be = True/False)

**Affected patches: 1**

`Controller.Pause` is a property, not a method. Writing `Controller.Pause` without an assignment (`= True` or `= False`) does nothing useful and can cause a syntax error. Fix: use `Controller.Pause = True` / `Controller.Pause = False`.

Affected scripts:

- Legend of Zelda v4.3.vbs [[original]](Legend%20of%20Zelda%20v4.3/Legend%20of%20Zelda%20v4.3.vbs.original#L404) [[patch]](Legend%20of%20Zelda%20v4.3/Legend%20of%20Zelda%20v4.3.vbs.patch) [[patched]](Legend%20of%20Zelda%20v4.3/Legend%20of%20Zelda%20v4.3.vbs)

## Duplicate InitPolarity call removed (second call crashes)

**Affected patches: 1**

Calling `InitPolarity` twice causes a crash. A duplicate call was present and is removed.

Affected scripts:

- Party Animal (Bally 1987).vbs [[original]](Party%20Animal%20%28Bally%201987%29/Party%20Animal%20%28Bally%201987%29.vbs.original#L223) [[patch]](Party%20Animal%20%28Bally%201987%29/Party%20Animal%20%28Bally%201987%29.vbs.patch) [[patched]](Party%20Animal%20%28Bally%201987%29/Party%20Animal%20%28Bally%201987%29.vbs)

## .AddBall.0 invalid syntax fixed to .AddBall 0

**Affected patches: 1**

`bsTrough.AddBall.0` attempts to access property `0` of the return value of `AddBall`, which is invalid. Fix: use `bsTrough.AddBall 0` (passing `0` as an argument).

Affected scripts:

- PinballMagic.v1.9.Hybrid.VPX8.vbs [[original]](PinballMagic.v1.9.Hybrid.VPX8/PinballMagic.v1.9.Hybrid.VPX8.vbs.original#L1448) [[patch]](PinballMagic.v1.9.Hybrid.VPX8/PinballMagic.v1.9.Hybrid.VPX8.vbs.patch) [[patched]](PinballMagic.v1.9.Hybrid.VPX8/PinballMagic.v1.9.Hybrid.VPX8.vbs)

## SolCallback assignment commented out (callback not available in standalone)

**Affected patches: 1**

A `SolCallback` entry pointing to a helper function that does not exist in standalone is commented out to prevent a "Sub or Function not defined" error at runtime.

Affected scripts:

- Playboy (Bally 1978).vbs [[original]](Playboy%20%28Bally%201978%29/Playboy%20%28Bally%201978%29.vbs.original#L36) [[patch]](Playboy%20%28Bally%201978%29/Playboy%20%28Bally%201978%29.vbs.patch) [[patched]](Playboy%20%28Bally%201978%29/Playboy%20%28Bally%201978%29.vbs)

## vpmShowDips call removed (not available in standalone)

**Affected patches: 1**

`vpmShowDips` displays the ROM DIP switch settings dialog. This function is not available in standalone builds. Fix: remove or comment out the call.

Affected scripts:

- Playboy (Stern 2002) v1.1.vbs [[original]](Playboy%28Stern2002%29v1.1/Playboy%20%28Stern%202002%29%20v1.1.vbs.original#L652) [[patch]](Playboy%28Stern2002%29v1.1/Playboy%20%28Stern%202002%29%20v1.1.vbs.patch) [[patched]](Playboy%28Stern2002%29v1.1/Playboy%20%28Stern%202002%29%20v1.1.vbs)

## PlaySoundAt() dot instead of comma separating arguments

**Affected patches: 1**

`PlaySoundAt "name". Object` uses a dot instead of a comma between the sound name and the position object. Fix: replace the dot with a comma: `PlaySoundAt "name", Object`.

Affected scripts:

- Ramones (HauntFreaks 2021).vbs [[original]](Ramones%20%28HauntFreaks%202021%29/Ramones%20%28HauntFreaks%202021%29.vbs.original#L802) [[patch]](Ramones%20%28HauntFreaks%202021%29/Ramones%20%28HauntFreaks%202021%29.vbs.patch) [[patched]](Ramones%20%28HauntFreaks%202021%29/Ramones%20%28HauntFreaks%202021%29.vbs)

## Eval()/Dictionary item double-indexed result (Wine limitation)

**Affected patches: 1**

`Eval("name")(0)(step)` or `Dict.Item(key)(0)` chains two index operations on the returned object. Wine does not support chained indexing of Eval/Dictionary results. Fix: store the result in a temporary variable first, then index it.

Affected scripts:

- Saving Wallden.vbs [[original]](Saving%20Wallden/Saving%20Wallden.vbs.original#L9387) [[patch]](Saving%20Wallden/Saving%20Wallden.vbs.patch) [[patched]](Saving%20Wallden/Saving%20Wallden.vbs)

## Object.state = off changed to = 0 (off evaluates to True in VBScript)

**Affected patches: 1**

In VBScript, the identifier `off` is not defined and evaluates to `Empty`, which coerces to `0` — but if `off` has been previously set to `True` in the script, assigning `state = off` sets state to `True` (on), not `False` (off). Fix: use the literal `0` instead.

Affected scripts:

- SpaceRamp (SuperEd) v3.03b.vbs [[original]](SpaceRamp%20%28SuperEd%29%20v3.03b/SpaceRamp%20%28SuperEd%29%20v3.03b.vbs.original#L612) [[patch]](SpaceRamp%20%28SuperEd%29%20v3.03b/SpaceRamp%20%28SuperEd%29%20v3.03b.vbs.patch) [[patched]](SpaceRamp%20%28SuperEd%29%20v3.03b/SpaceRamp%20%28SuperEd%29%20v3.03b.vbs)

## Constant renamed to avoid conflict or improve clarity

**Affected patches: 1**

A constant name conflicted with a built-in function or property name, or was renamed for clarity (e.g. `Const Offset` → `Const DigitsOffset` to avoid shadowing `Offset` elsewhere).

Affected scripts:

- Taxi (Williams 1988) VPW v1.2.2.vbs [[original]](Taxi%20%28Williams%201988%29%20VPW%20v1.2.2/Taxi%20%28Williams%201988%29%20VPW%20v1.2.2.vbs.original#L1435) [[patch]](Taxi%20%28Williams%201988%29%20VPW%20v1.2.2/Taxi%20%28Williams%201988%29%20VPW%20v1.2.2.vbs.patch) [[patched]](Taxi%20%28Williams%201988%29%20VPW%20v1.2.2/Taxi%20%28Williams%201988%29%20VPW%20v1.2.2.vbs)

## ActiveBall.id replaced with ActiveBall.Uservalue

**Affected patches: 1**

The `ActiveBall.id` property is not available in VPX Standalone. Scripts that use `.id` to track specific balls are updated to use `ActiveBall.Uservalue` instead.

Affected scripts:

- The Rolling Stones (Stern 2011)_Bigus(MOD)3.0.vbs [[original]](The%20Rolling%20Stones%20%28Stern%202011%29_Bigus%28MOD%293.0/The%20Rolling%20Stones%20%28Stern%202011%29_Bigus%28MOD%293.0.vbs.original#L268) [[patch]](The%20Rolling%20Stones%20%28Stern%202011%29_Bigus%28MOD%293.0/The%20Rolling%20Stones%20%28Stern%202011%29_Bigus%28MOD%293.0.vbs.patch) [[patched]](The%20Rolling%20Stones%20%28Stern%202011%29_Bigus%28MOD%293.0/The%20Rolling%20Stones%20%28Stern%202011%29_Bigus%28MOD%293.0.vbs)

