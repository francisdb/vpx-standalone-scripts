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

- 1342729923_RollerCoasterTycoon(Stern2002)1.3.vbs
- 2001 (Gottlieb 1971) v0.99a.vbs
- AC-DC LUCI Premium VR (Stern 2013) v1.1.3.vbs
- Alien Poker (Williams 1980).vbs
- Apollo 13 (Sega 1995) w VR Room v2.1.1.vbs
- Atlantis (Bally 1989) w VR Room v2.0.vbs
- Barracora (Williams 1981) w VR Room v2.1.3.vbs
- Batman Forever (Sega 1995) 1.3.vbs
- Batman [The Dark Knight] (Stern 2008) 1.16.vbs
- Batman [The Dark Knight] (Stern 2008) v1.0.8.vbs
- Batman [The Dark Knight] (Stern 2008).vbs
- BeastieBoysv1.0.vbs
- Black Knight (Williams 1980).vbs
- Black Knight 2000 (Williams 1989) w VR Room v2.0.2.vbs
- Blackout (Williams 1980) v1.0.1 - SBR34.vbs
- Blood Machines 2.0.vbs
- Bounty Hunter (Gottlieb 1985).vbs
- Cactus Canyon (Bally 1998) VPW 1.1.vbs
- Cactus Jacks (Gottlieb 1991) w VR Room v2.0.2.vbs
- Capt. Fantastic and The Brown Dirt Cowboy (Bally 1976) 2.0.2.vbs
- Centaur (Bally 1981).vbs
- Centigrade 37 (Gottlieb 1977).vbs
- Checkpoint (Data East 1991)2.0.vbs
- Comet (Williams 1985) w VR Room v2.3.vbs
- Cue Ball Wizard (Gottlieb 1992) v1.1.2.vbs
- Diner (Williams 1990) VPW Mod 1.0.2.vbs
- Doctor Who (Bally 1992) VPW Mod v1.1.vbs
- Dracula (Stern 1979).vbs
- Elektra (Bally 1981) w VR Room v2.0.7.vbs
- Elvis_MOD_2.0.vbs
- Family Guy 1.0.vbs
- Flash (Williams 1979).vbs
- Fog, The (Gottlieb 1979) v2.5 for 10.7.vbs
- Galaxy (Stern 1980).vbs
- Game of Thrones LE (Stern 2015) VPW v1.0.2.vbs
- Genie (Gottlieb 1979).vbs
- Gorgar_1.1.vbs
- Halloween 1978-1981 (Original 2022) 1.03.vbs
- Hang Glider (Bally 1976) VPW v1.2.vbs
- Harlem Globetrotters (Bally 1979) v1.14.vbs
- Harley Davidson (Sega 1999) v1.12.vbs
- Heat Wave (Williams 1964).vbs
- Indiana Jones The Pinball Adventure (Williams 1993) VPWmod v1.1.vbs
- Iron Maiden (Original 2022) VPW 1.0.12.vbs
- Iron Man Vault Edition (Stern 2010) VPW v1.0.vbs
- Jack-Bot (Williams 1995).vbs
- Jive Time (Williams 1970) 2.0.vbs
- Joker Poker EM (Gottlieb 1978) 1.6.vbs
- Judge Dredd (Bally 1993) VPW v1.1.vbs
- Jungle Lord (Williams 1981) w VR Room v2.01.vbs
- Laser Cue (Williams 1984) w VR Room 2.0.0.vbs
- Laser War (Data East 1987) w VR Room v2.0.vbs
- Led Zeppelin Pinball 2.5.vbs
- Magic City (Williams 1967).vbs
- Maverick (Data East 1994) VPW v1.3.vbs
- Medusa (Bally 1981) w VR Room v2.0.1.vbs
- Meteor (Stern 1979).vbs
- Monster Bash (Williams 1998) VPWmod v1.0.vbs
- MrDoom (Recel 1979)1.3.0.vbs
- MrEvil (Recel 1979)1.3.0.vbs
- Nine Ball (Stern 1980).vbs
- O Brother Where Art Thou (Zoss 2021)1_6.vbs
- Pharaoh (Williams 1981) w VR Room v2.0.3.vbs
- PinBlob (CLV 2024).vbs
- PinBot (Williams 1986) 2.1.1.vbs
- Power Play (Bally 1977).vbs
- Robocop (Data East 1989)_drakkon(mod_1.2).vbs
- Seawitch (Stern 1980).vbs
- Solar Fire (Williams 1981) w VR Room v2.0.5.vbs
- Space Shuttle (Williams 1984).vbs
- Star Trek (Bally 1979) 2.1.1.vbs
- Star Wars (Data East 1992) VPW v1.2.2.vbs
- Stars (Stern 1978).vbs
- Stellar Wars (Williams 1979).vbs
- Stingray (Stern 1977).vbs
- Strikes And Spares (Bally 1978) 2.0.vbs
- Strip Joker Poker (Gottlieb 1978) 1.5.vbs
- Swords of Fury (Williams 1988).vbs
- TX-Sector (Gottlieb 1988) SG1bsoN Mod V1.1.vbs
- Tales from the Crypt (Data East 1993) VPW v1.22.vbs
- Transporter the Rescue (Midway 1989) VPW v1.05.vbs
- Twilight Zone (Bally 1993) 2.3.6.vbs
- Twister (Sega 1996) v2.0 w VR Room.vbs
- Viking (Bally 1980).vbs
- fireball II VPX.vbs

## Parenthesis on first argument of a procedure call not handled correctly by Wine VBScript

**Affected patches: 67**

When the first argument of a Sub/procedure call starts with `(`, Wine's VBScript parses it as a call with explicit parentheses and treats the rest of the expression as a separate statement. Any arithmetic that continues after the closing `)` is evaluated separately and discarded. Example: `AddScore (a+b)*c` is parsed as `AddScore(a+b)` then `*c` is discarded. The same rule applies inside `ExecuteGlobal` strings — `SetLamp (me.UserValue - INT(me.UserValue)) * 100` becomes `SetLamp (me.UserValue - INT(me.UserValue))` with the `* 100` thrown away. Fix: rearrange the expression so it does not start with `(`, or move the multiplier to the front, e.g. `AddScore (a+b)*c` → `AddScore c*(a+b)`, or wrap the entire expression in double parentheses.

Affected scripts:

- 2104398928_CARtoonsRC(Nailed2021)v1.3.vbs
- A Charlie Brown Christmas feat. Vince Guaraldi (iDigStuff 2023).vbs
- Aladdin's Castle (Bally 1976) - DOZER - MJR_1.01.vbs
- Attack On Titan (cHuG_MOD_1.4).vbs
- Ben-Hur (Staal 1977) V1.1.1 DT-FS-VR-MR Ext2k Conversion.vbs
- Big Deal (Williams 1977)V2.1.vbs
- Big Horse (Maresa 1975) v55_VPX8.vbs
- Big Star (Williams 1972) v4.3.vbs
- Black & Red (Inder 1975) v55_VPX8.vbs
- Bumper (Bill Port - 1977) v4.vbs
- Cannes (Segasa 1976)V1.3.vbs
- DORAEMON.vbs
- DS.vbs
- DUCKTALES.vbs
- FNAF.vbs
- Fifteen (Inder 1974) v4.vbs
- Fireball XL5.vbs
- Game of Thrones LE (Stern 2015) VPW 1.2.vbs
- Game of Thrones LE (Stern 2015) VPW v1.0.2.vbs
- Gemini (Gottlieb 1978).vbs
- Gun Men (Staal 1979).vbs
- HALO.vbs
- Hellraiser 1.2.vbs
- Honey (Williams 1971)V1.3.vbs
- Indiana Jones (Stern 2008)-Hanibal-2.6.vbs
- Inhabiting Mars RC 4.vbs
- Liberty Bell (Williams 1977) V1.01.vbs
- Lightning Ball (Gottlieb 1959).vbs
- Luck Smile (Inder 1976) 1p v55_VPX8.vbs
- Luck Smile (Inder 1976) 4p v55_VPX8.vbs
- Mago de Oz - the pinball v4.3.vbs
- Metal Slug_1.05.vbs
- Monaco (Segasa 1977)V1.4.vbs
- Monkey Island VR Room.vbs
- Motley Crue SS.vbs
- Munsters (Original 2020) 1.05.vbs
- PVM.vbs
- Pat Hand (Williams 1975)V1.4.vbs
- Punchout.vbs
- Rancho (Williams 1976)V1.4.vbs
- Rattlecan v1.5.3.vbs
- Riccione.vbs
- Route66.vbs
- Running Horse (Inder 1976) v55_VPX8.vbs
- SMGP.vbs
- SPP.vbs
- Seven Winner (Inder 1973) v4.vbs
- Skylab (Williams 1974)V1.3.vbs
- Sonic The Hedgehog (Brendan Bailey 2005) VPX MOD 1.33.vbs
- Sons Of Anarchy (Original 2019).vbs
- Stardust (Williams 1971) v4.vbs
- Streets of Rage (TBA 2018).vbs
- Super Star (Williams 1972) v4.3.vbs
- The Grinch (Original 2022) pinballfan2018.vbs
- TheATeam.vbs
- Van Halen 1.2.vbs
- indochinecentral.vbs
- indochinecentralPUP.vbs
- monkeyisland.vbs
- pizzatime-65.vbs
- pulp_fiction.vbs
- speakeasy2.vbs
- speakeasy4.vbs
- wackyraces.vbs

## CreateObject / FileSystemObject / WshShell removed (not supported in Wine)

**Affected patches: 22**

`CreateObject("WScript.Shell")`, `CreateObject("Scripting.FileSystemObject")`, and `WshShell.RegWrite/RegRead` are Windows-only COM objects that are not available under Wine. Typically used for music folder scanning, registry access, or file system operations. These calls are removed or replaced with standalone-compatible equivalents.

Affected scripts:

- ABBAv2.0.vbs
- Blood Machines 2.0.vbs
- Bob Marley Mod.vbs
- DarkPrincess1.3.1.vbs
- Darkest Dungeon pupevent_2.3c.vbs
- Diablo Pinball v4.3.vbs
- Fire Action (Taito do Brasil - 1980) v4.vbs
- Freddys Nightmares.vbs
- Ice Age Christmas (Original Balutito 2021) endeemillr mod.vbs
- Iron Maiden Virtual Time (Original 2020).vbs
- KISS (Stern 2015)_Bigus(MOD)1.1.vbs
- Lawman (Gottlieb 1971) 1.1.vbs
- Led Zeppelin Pinball 2.5.vbs
- Metal Slug_1.05.vbs
- Sonic The Hedgehog (Brendan Bailey 2005) VPX MOD 1.33.vbs
- Space Shuttle (Taito do Brasil - 1982) v4.vbs
- Spooky Wednesday Pro.vbs
- Spooky_Wednesday VPX 2024.vbs
- Student Prince (Williams 1968).vbs
- Three Angels (Original 2018) 1.3.vbs
- Thundercats Pinball v1.0.9.vbs
- Vortex (Taito do Brasil - 1981) v4.vbs

## DMD standalone compatibility setup (ImplicitDMD_Init / UseVPMDMD)

**Affected patches: 22**

Scripts that use a separate DMD script (`.vbs.dmd`) need a `Sub ImplicitDMD_Init` entry point and/or a `Dim UseVPMDMD` variable so that VPX Standalone can render the DMD correctly. These additions enable the DMD display to work without VPinMAME on standalone builds.

Affected scripts:

- Avatar (Stern 2012) v1.12.vbs
- BarbWire(Gottlieb1996)JoePicassoModv1.2.vbs.dmd
- Blood Machines 2.0.vbs.dmd
- Bounty Hunter (Gottlieb 1985).vbs.dmd
- Cactus Canyon (Bally 1998) VPW 1.1.vbs.dmd
- Diner (Williams 1990) VPW Mod 1.0.2.vbs.dmd
- Futurama (Original 2024) v1.1.vbs
- Game of Thrones LE (Stern 2015) VPW v1.0.2.vbs.dmd
- Goin' Nuts (Gottlieb 1983).vbs.dmd
- Hook (Data East 1992)_VPWmod_v1.0.vbs.dmd
- Pink Panther (Gottlieb 1981).vbs.dmd
- Scared Stiff (Bally 1996) v1.29.vbs.dmd
- Seawitch (Stern 1980).vbs.dmd
- South Park (Sega 1999) 1.3.vbs.dmd
- Space Shuttle (Williams 1984).vbs.dmd
- Star Trek LE (Stern 2013) v1.09.vbs.dmd
- Taxi (Williams 1988) VPW v1.2.2.vbs.dmd
- The Goonies Never Say Die Pinball (VPW 2021) v1.4.vbs
- Three Angels (Original 2018) 1.3.vbs
- X-Files (Sega 1997) v1.29.vbs
- X-Files Hanibal 4k LED Edition.vbs
- fireball II VPX.vbs.dmd

## Const declaration moved to different position (execution order fix)

**Affected patches: 16**

VBScript evaluates code sequentially; a `Const` must be declared before its first use. In some scripts the constant was declared *after* the code that referenced it. Fix: move the `Const` declaration to an earlier position in the file.

Affected scripts:

- AC-DC Pro Vault-1.0.vbs
- AC-DC Pro-1.0.vbs
- AC-DC_Back In Black LE-1.5.vbs
- AC-DC_Let There Be Rock LE-1.5.vbs
- AC-DC_Luci-1.5.vbs
- AC-DC_Premium-1.5.vbs
- Aztec (Williams 1976) 1.3 Mod Citedor JPJ-ARNGRIM-CED Team PP.vbs
- Aztec High-Tapped (Williams 1976).vbs
- Cybernaut (Bally 1985)_Bigus(MOD)1.0.vbs
- Cybernaut darkmod.vbs
- Cybernaut.vbs
- Fireball Classic (Bally 1984)1.2.vbs
- Flash Gordon (Bally 1981) VPW Mod v3.0.vbs
- Harlem Globetrotters (Bally 1979) VPW 1.0.vbs
- High_Speed_(Williams 1986) v0.107a (MOD 3.0).vbs
- Rollercoaster Tycoon (Stern 2002) VPW 2.3.vbs

## NVramPatch calls removed (not supported in standalone)

**Affected patches: 13**

`NVramPatchLoad`, `NVramPatchExit`, and `NVramPatchKeyCheck` rely on a Windows-only NVram helper that is not available in standalone builds. Calls to these functions are commented out or removed.

Affected scripts:

- ABBAv2.0.vbs
- Bob Marley Mod.vbs
- Cosmic (Taito do Brasil - 1980) v4.vbs
- Last Starfighter, The (Taito, 1983) hybrid v1.04.vbs
- Mr Black (Taito do Brasil - 1984) v4.vbs
- Oba Oba (Taito do Brasil - 1979) v55_VPX8.vbs
- Rally (Taito do Brasil - 1980) v4.vbs
- Titan (Taito do Brasil - 1981) v4.vbs
- Topaz (Inder 1979).vbs
- Vortex (Taito do Brasil - 1981) v4.vbs
- Warriors The Full DMD 2.0 (Iceman 2023).vbs
- Wednesday (Netflix 2023).vbs

## cGameName (ROM name) constant fix

**Affected patches: 10**

The `cGameName` constant sets the ROM name passed to VPinMAME. An incorrect ROM name prevents the game from loading. This fix corrects the ROM name string to match the actual ROM file name.

Affected scripts:

- Bad Cats (Williams 1989) VPW 3.0.vbs.dmd
- BeastieBoysv1.0.vbs
- Death Proof Balutito V2.vbs
- PinBlob (CLV 2024).vbs
- Safe Cracker (Bally 1996) v1.0.vbs
- Simpsons Treehouse of Horror MOD v2023.3 (Simpsons Pinball Party, The (Stern 2003) VPW 2.0.3 base).vbs
- South Park - Halloween.vbs
- Theatre of Magic (Bally 1995) 2.4.vbs
- Warriors The Full DMD 2.0 (Iceman 2023).vbs
- Wednesday (Netflix 2023).vbs

## GetDMDColor function call removed (not available in standalone)

**Affected patches: 8**

`GetDMDColor` is a helper function that does not exist in all script environments. Calling it causes a runtime error on standalone. The call is removed or commented out.

Affected scripts:

- DarkPrincess1.3.1.vbs
- Metal Slug_1.05.vbs
- Motley Crue SS.vbs
- Sons Of Anarchy (Original 2019).vbs
- Spooky Wednesday Pro.vbs
- Spooky_Wednesday VPX 2024.vbs
- Streets of Rage (TBA 2018).vbs
- Three Angels (Original 2018) LW.vbs

## Execute/ExecuteGlobal with object Set in loop (replaced with IsObject check)

**Affected patches: 6**

`Execute "Set Lights(" & i & ") = L" & i` creates object references dynamically in a loop. In Wine's VBScript, this pattern fails when some loop variables are not objects. Fix: replace with an `If IsObject(eval("L" & i)) Then` guard before assigning.

Affected scripts:

- 1455577933_MedievalMadness_Upgrade(Real_Final).vbs
- Bram Stokers Dracula (1993).vbs
- Cirqus_Voltaire_Hanibal-Mod_3.7.vbs
- VP10-Terminator 3 (Stern 2003) Hanibal v1.5.vbs
- Vampirella.vbs
- Vampirella1 .3.vbs

## Nested Sub definition extracted to top level

**Affected patches: 6**

VBScript does not allow a `Sub` to be defined inside another `Sub`. Scripts that contained nested Sub definitions (e.g. `chilloutthemusic` or `turnitbackup` nested inside the PuPlayer initialisation block) fail to compile. Fix: move the inner Sub to the top level of the script.

Affected scripts:

- 1679379865_jurassicparklimitededition.vbs
- Stranger Things (Original 2020) LW.vbs
- Stranger Things - SE 1.47_OSB.vbs
- Stranger Things 4 Premium.vbs
- The Beatles_007.vbs
- tlk-0.35.vbs

## gBOT replaced with getballs() function call

**Affected patches: 6**

The `gBOT` global variable holds the current ball collection in older VPX scripts. In newer versions it is replaced by the `getballs()` function. Scripts using `gBOT` are updated to call `getballs()` instead.

Affected scripts:

- Atlantis (Bally 1989) w VR Room v2.0.vbs
- Black Knight 2000 (Williams 1989) w VR Room v2.0.2.vbs
- Cactus Jacks (Gottlieb 1991) w VR Room v2.0.2.vbs
- Iron Man Vault Edition (Stern 2010) VPW v1.0.vbs
- Laser War (Data East 1987) w VR Room v2.0.vbs
- X-Men LE (Stern 2012) VPW v1.0.vbs

## File path case sensitivity fix (Linux filesystem)

**Affected patches: 6**

Linux file systems are case-sensitive, unlike Windows NTFS. File names embedded in scripts (images, videos, ROM file names) must match the actual file names exactly. Fix: correct the capitalisation of the file name string.

Affected scripts:

- Beavis and Butt-head_Pinballed.vbs
- Blood Machines 2.0.vbs
- DarkPrincess1.3.1.vbs
- StarWars BountyHunter 3.02.vbs
- Thunderbirds original 2022 v1.0.2.vbs
- UT99CTF_GE_2.3.vbs

## Incorrect boolean expression (IsGIOn <> Not IsOff)

**Affected patches: 5**

`isGIOn <> Not IsOff` is logically incorrect — `Not IsOff` negates the variable, producing a boolean, and `<>` then compares it to `isGIOn`. Under Wine this expression evaluates incorrectly. Fix: use a simple assignment or comparison without the erroneous `<>` operator.

Affected scripts:

- AceOfSpeed.vbs
- Death Proof Balutito V2.vbs
- Iron Maiden Virtual Time (Original 2020).vbs
- Laser War (Data East 1987) w VR Room v2.0.vbs
- Mousin' Around! (Bally 1989) w VR Room v2.1.vbs

## Configuration constant value changed for standalone compatibility

**Affected patches: 5**

Certain constants (e.g. `cController`, `PupScreenMiniGame`, `Mute_Sound_For_PuPPack`) have default values that assume a full Windows VPX environment. For standalone builds these values need to be adjusted — for example switching from VPinMAME controller to B2S server, or disabling features not available in standalone.

Affected scripts:

- Bart VS the Space Mutants 1.1.vbs
- Centaur (Bally 1981).vbs.dmd
- Dream Daddy1.5.vbs
- Heavy Metal Meltdown (Bally 1987).vbs
- Megadeth (original).vbs.dmd

## Sub/Function definition with inline code (no newline after definition)

**Affected patches: 5**

Placing code on the same line as the `Sub` or `Function` header (e.g. `Sub Sw18_Hit()Controller.Switch(18)=1`) is valid on Windows VBScript but fails to parse correctly under Wine. Fix: add a newline after the `()` so the body starts on the next line.

Affected scripts:

- Bugs Bunny's Birthday Ball (Bally 1991)Rev2.3b.vbs
- Cue Ball Wizard (Gottlieb 1992) v1.1.2.vbs
- Haunted House (Gottlieb 1982) 1.0.vbs
- Haunted House (Gottlieb 1982)_Bigus(MOD)1.0.vbs
- Scared Stiff (Bally 1996) v1.29.vbs

## For Each loop variable reused (causes error in Wine)

**Affected patches: 4**

Reusing the same loop variable in nested or back-to-back `For Each` loops causes a runtime error in Wine's VBScript. Fix: use a different variable name for each loop level.

Affected scripts:

- Atlantis (Bally 1989) w VR Room v2.0.vbs
- Black Knight 2000 (Williams 1989) w VR Room v2.0.2.vbs
- Cactus Jacks (Gottlieb 1991) w VR Room v2.0.2.vbs
- Laser War (Data East 1987) w VR Room v2.0.vbs

## UseVPMDMD variable renamed to UseVPMColoredDMD

**Affected patches: 4**

Scripts that use a colored DMD must use the variable name `UseVPMColoredDMD` (not `UseVPMDMD`) so that the VPX standalone renderer applies the correct colour palette.

Affected scripts:

- Bad Cats (Williams 1989) VPW 3.0.vbs.dmd
- Champion Pub (Williams 1998) v1.43.vbs
- Cue Ball Wizard (Gottlieb 1992) v1.2.4.vbs
- Tron Legacy (Stern 2011) VPW Mod v1.1.vbs

## typename() casing fix (typename → TypeName)

**Affected patches: 4**

Wine's VBScript is stricter about built-in function capitalisation. `typename(x)` must be `TypeName(x)`.

Affected scripts:

- Elvis_MOD_2.0.vbs
- Simpsons Pinball Party, The (Stern 2003) VPW 2.0.3.vbs
- Simpsons Treehouse of Horror MOD v2023.3 (Simpsons Pinball Party, The (Stern 2003) VPW 2.0.3 base).vbs
- TX-Sector (Gottlieb 1988) SG1bsoN Mod V1.1.vbs

## Statement starting with colon (invalid VBScript syntax)

**Affected patches: 4**

Lines that begin with `:` followed by a statement (e.g. `        : Controller.B2SSetData 9, 0`) are not valid VBScript syntax — the colon is a statement separator, not a line prefix. Fix: remove the leading colon.

Affected scripts:

- Junkyard Cats_1.07.vbs
- NBA (Stern 2009)_Bigus(MOD)1.4.vbs
- Scarface - Balls and Power 1.2.vbs
- The Fifth Element 1.3.vbs

## UBound casing fix (Ubound → UBound)

**Affected patches: 3**

Wine's VBScript is stricter about built-in function capitalisation. `Ubound` must be `UBound`.

Affected scripts:

- Monkey Island VR Room.vbs
- Power Play (Bally 1977).vbs
- monkeyisland.vbs

## Double-dot in string callback (e.g. "obj..Method")

**Affected patches: 3**

SolCallback strings like `"dtbank..SolHit"` contain a double dot, which is invalid. The double dot arises when the object name is accidentally repeated. Fix: correct to a single dot, e.g. `"dtbank.SolHit"`.

Affected scripts:

- ABBAv2.0.vbs
- Alfred Hitchcock's Psycho (TBA 2019).vbs
- Vortex (Taito do Brasil - 1981) v4.vbs

## Hex literal with excess leading zeros (e.g. &H000000031)

**Affected patches: 3**

A hex literal with an odd or excessive number of digits (e.g. `&H000000031` — nine hex digits) is rejected by Wine's VBScript parser. Fix: reduce to the correct even number of digits (e.g. `&H00000031`).

Affected scripts:

- Andromeda (Game Plan 1985) v4.vbs
- Grand Slam (Bally 1983) 2.3.vbs
- Topaz (Inder 1979).vbs

## Me(Idx) collection indexer replaced with named collection

**Affected patches: 3**

Inside a VBScript `Class`, `Me(n)` is not a valid way to index a collection under Wine. Fix: use the named array/collection directly, e.g. `Spinner(Idx).IsDropped`.

Affected scripts:

- Four Seasons (Gottlieb 1968).vbs
- Four Seasons (Gottlieb 1968)_Teisen_MOD.vbs
- Roller Coaster (Gottlieb 1971)x.vbs

## Trim.visible assignment removed (Trim is a VBScript reserved function)

**Affected patches: 3**

`Trim` is a built-in VBScript string function. Using it as an object name (e.g. `Trim.visible = 1`) is interpreted as a function call and raises a type error. Fix: comment out or rename the object.

Affected scripts:

- Gladiators (Premier 1993) v1.1.1.vbs
- Surf'n Safari v1.3.4.vbs
- Wipe Out (Premier 1993) 1.1.0.vbs

## One-liner If/Then with dangling Else (missing End If)

**Affected patches: 3**

A single-line `If x Then y: Else` with nothing after the `Else` is syntactically invalid — VBScript expects an `End If` or an inline statement after `Else`. Fix: remove the trailing `: Else` or add `End If`.

Affected scripts:

- South Park (Sega 1999) 1.3.vbs
- South Park - Halloween.vbs
- Starship Troopers (Sega 1997) VPW Mod v2.0.vbs

## Duplicate score function calls removed

**Affected patches: 2**

The same `AddScore` call was duplicated multiple times in the same event handler, causing the score to be added multiple times per activation. Fix: keep only one call.

Affected scripts:

- 2104398928_CARtoonsRC(Nailed2021)v1.3.vbs
- Gemini (Gottlieb 1978).vbs

## 'default' reserved word conflict fixed (Const default = 0 added)

**Affected patches: 2**

Using `default` as a variable name conflicts with VBScript's reserved `Default` property keyword. Fix: declare `Const default = 0` at the top of the script to shadow the reserved word with a harmless constant.

Affected scripts:

- Bud and Terence (Original 2024) v1.6.vbs
- IT Pinball Madness (JP Salas,Joe Picasso)1.2.vbs

## BSize/BMass constant renamed to BallSize/BallMass

**Affected patches: 2**

`Const BSize` or `Const BMass` conflict with VPX's built-in `BallSize` and `BallMass` properties. Fix: rename the script constants to avoid the conflict.

Affected scripts:

- Cyclone (Williams 1988) 2.0.1.vbs
- Space Jam (Sega 1996) 1.4 JPJ - Edizzle - TeamPP - JLou.vbs

## OptionsLoad/OptionsEdit call removed (not available in standalone)

**Affected patches: 2**

`OptionsLoad` and `OptionsEdit` are helper functions provided by some VPX utilities but not available in standalone builds. Calling them causes a "Sub or Function not defined" error. Fix: remove the calls and inline the required initialisation.

Affected scripts:

- Gremlins by Balutito 1.7.vbs
- Victory (Gottlieb 1987) 2.0.2.vbs

## VPinMAME Settings.Value() call fix

**Affected patches: 2**

`Controller.Settings.Value("key")` must be called without extra parentheses around the argument in VBScript when the return value is discarded. Fix: use `Controller.Settings.Value "key"` or restructure the call.

Affected scripts:

- Phantom Of The Opera (Data East 1990) v1.24.vbs
- The Phantom Of The Opera (Data East 1990).vbs

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

- TimonPumbaa.vbs
- Wheel of Fortune (Stern 2007) 1.0.vbs

## CDbl() replaced with CBool() (wrong type conversion)

**Affected patches: 1**

Affected scripts:

- A-Go-Go (Williams 1966).vbs

## DisplayTimer.Enabled = DesktopMode fix

**Affected patches: 1**

Setting `DisplayTimer.Enabled = DesktopMode` assigns a boolean to a timer's Enabled property. In VR mode `DesktopMode` is `False`, inadvertently disabling the timer. Fix: wrap in a conditional or evaluate correctly.

Affected scripts:

- AceOfSpeed.vbs

## cvpmDictionary replaced with Scripting.Dictionary

**Affected patches: 1**

`cvpmDictionary` is an older VPM wrapper class. Scripts using `New cvpmDictionary` are updated to use `CreateObject("Scripting.Dictionary")` directly.

Affected scripts:

- Batman66_1.1.0.vbs

## LinkedTo property must be assigned an Array (not a single object)

**Affected patches: 1**

The `DropTarget.LinkedTo` property expects an Array of linked targets, not a bare object reference. Fix: wrap the value in `Array(...)`, e.g. `dtR.LinkedTo = Array(dtL)`.

Affected scripts:

- Contact (Williams 1978)_Bigus(MOD)1.0.vbs

## Const tnob value fix

**Affected patches: 1**

The `tnob` (total number of balls) constant had an incorrect value, causing ball tracking issues. Fix: set `Const tnob` to the correct number of simultaneous balls for the table.

Affected scripts:

- GalaxyPlay_1_2.vbs

## Code missing newlines between statements (all on one line)

**Affected patches: 1**

Multiple statements were concatenated onto a single line without newlines, making the script extremely long and causing parse errors in Wine. Fix: insert newlines to separate each statement.

Affected scripts:

- Iron Maiden (Original 2022) VPW 1.0.12.vbs

## b2s.vbs GetTextFile replaced with inline B2S helper functions

**Affected patches: 1**

`ExecuteGlobal GetTextFile("b2s.vbs")` loads a shared B2S helper script at runtime. This approach fails in standalone because the helper file is not in the expected path. Fix: inline the needed B2S helper functions (`SetB2SData`, `StepB2SData`, etc.) directly into the script.

Affected scripts:

- Jungle_Queen.vbs

## Variable value or array size changed for standalone compatibility

**Affected patches: 1**

A variable value or array dimension is changed to a value more appropriate for standalone play — for example, increasing a `LampState` array size to accommodate all lamps, or adjusting a physics limit value.

Affected scripts:

- Junk Yard (Williams 1996) v1.83.vbs

## Controller.Pause missing assignment (should be = True/False)

**Affected patches: 1**

`Controller.Pause` is a property, not a method. Writing `Controller.Pause` without an assignment (`= True` or `= False`) does nothing useful and can cause a syntax error. Fix: use `Controller.Pause = True` / `Controller.Pause = False`.

Affected scripts:

- Legend of Zelda v4.3.vbs

## Duplicate InitPolarity call removed (second call crashes)

**Affected patches: 1**

Calling `InitPolarity` twice causes a crash. A duplicate call was present and is removed.

Affected scripts:

- Party Animal (Bally 1987).vbs

## .AddBall.0 invalid syntax fixed to .AddBall 0

**Affected patches: 1**

`bsTrough.AddBall.0` attempts to access property `0` of the return value of `AddBall`, which is invalid. Fix: use `bsTrough.AddBall 0` (passing `0` as an argument).

Affected scripts:

- PinballMagic.v1.9.Hybrid.VPX8.vbs

## SolCallback assignment commented out (callback not available in standalone)

**Affected patches: 1**

A `SolCallback` entry pointing to a helper function that does not exist in standalone is commented out to prevent a "Sub or Function not defined" error at runtime.

Affected scripts:

- Playboy (Bally 1978).vbs

## vpmShowDips call removed (not available in standalone)

**Affected patches: 1**

`vpmShowDips` displays the ROM DIP switch settings dialog. This function is not available in standalone builds. Fix: remove or comment out the call.

Affected scripts:

- Playboy (Stern 2002) v1.1.vbs

## PlaySoundAt() dot instead of comma separating arguments

**Affected patches: 1**

`PlaySoundAt "name". Object` uses a dot instead of a comma between the sound name and the position object. Fix: replace the dot with a comma: `PlaySoundAt "name", Object`.

Affected scripts:

- Ramones (HauntFreaks 2021).vbs

## Eval()/Dictionary item double-indexed result (Wine limitation)

**Affected patches: 1**

`Eval("name")(0)(step)` or `Dict.Item(key)(0)` chains two index operations on the returned object. Wine does not support chained indexing of Eval/Dictionary results. Fix: store the result in a temporary variable first, then index it.

Affected scripts:

- Saving Wallden.vbs

## Object.state = off changed to = 0 (off evaluates to True in VBScript)

**Affected patches: 1**

In VBScript, the identifier `off` is not defined and evaluates to `Empty`, which coerces to `0` — but if `off` has been previously set to `True` in the script, assigning `state = off` sets state to `True` (on), not `False` (off). Fix: use the literal `0` instead.

Affected scripts:

- SpaceRamp (SuperEd) v3.03b.vbs

## Constant renamed to avoid conflict or improve clarity

**Affected patches: 1**

A constant name conflicted with a built-in function or property name, or was renamed for clarity (e.g. `Const Offset` → `Const DigitsOffset` to avoid shadowing `Offset` elsewhere).

Affected scripts:

- Taxi (Williams 1988) VPW v1.2.2.vbs

## ActiveBall.id replaced with ActiveBall.Uservalue

**Affected patches: 1**

The `ActiveBall.id` property is not available in VPX Standalone. Scripts that use `.id` to track specific balls are updated to use `ActiveBall.Uservalue` instead.

Affected scripts:

- The Rolling Stones (Stern 2011)_Bigus(MOD)3.0.vbs

