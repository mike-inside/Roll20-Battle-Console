#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetBatchLines -1
SetTitleMatchMode, 3
#SingleInstance Force
; Register a function to be called on exit:
OnExit("ExitFunc")

;-------------------------------------------------------------------------------
; User Settings
;-------------------------------------------------------------------------------
consoleName := "Bo's Battle Console"
roll20Games := "test,Method Style D&D 5e" ;comma-separated list of Roll20 games you might use this app in
browsers := "chrome.exe,firefox.exe" ;comma-separated list of browsers you use
; The browser you want to open new windows in :
browserPath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --allow-file-access-from-files --new-window
;browserPath = "C:\Program Files\Firefox Developer Edition\firefox.exe" -new-window
gArrange := 0 ; do you want this app to automatically arrange windows by default (1 for true, 0 for false)

;-------------------------------------------------------------------------------
; Character Stats
;-------------------------------------------------------------------------------
charName := "Bo"
mods := {str: 5, dex: 5, con: 3, int: 4, wis: 2, cha: 1}
prof := 4 ; proficiency bonus
rageDamage := 2
sneakAttackDamage := "5d6"
baseAC := 20
offhandWeapon := "Drow Shortblade"

;-------------------------------------------------------------------------------
; Initialise global variables
;-------------------------------------------------------------------------------
commandBuffer := ""
activatePrimed := true
; commandTime := A_TickCount - 5000 ; not used, was going to put a delay in how often commands could fire
roll20WindowSuffix := " | Roll20" ; Should not need to be changed unless Roll20 changes the way their website works
w := {} ; array of objects containing window information
oncePerTurnAbility := {} ; used to store the names of weapons that have used their once per turn abilities, reset when beginning new round
gMute := 0

;-------------------------------------------------------------------------------
; Set Window data
;-------------------------------------------------------------------------------
; pass variables containing Window Title, Title Match Type, URL (for automatic opening), Exe 
;  (ignore exe if it is a browser window, or pass true if the window should be automatically opened)
; Match 1 = starts with
; Match 2 = contains
; Match 3 = exact
w.console := NewWindow(consoleName, 3, false, "AutoHotKey.exe")
w.weapon :=  NewWindow("Arsenal Weapon Notes", 1, "https://docs.google.com/document/d/18c5vmZ7usGORrVS3i88AXlfV1MIo0DIsa0GOvBxjC00/edit#", true)
w.tool :=    NewWindow("5etools", 2, "file:///E:/Share/DnD/5eTools.1.73.5/5etools.html", true)
w.tool2 :=   NewWindow("D&D Inside", 1, "http://www.mikeinside.com/")
w.drive :=   NewWindow("D&D - Google Drive", 1, "https://drive.google.com/drive/folders/1oZsTNdPkUpMvfhxRKqOWE2qzny3qp7iU")
w.char :=    NewWindow(charName, 1)
w.game :=    NewWindow(roll20Game, 1, "https://app.roll20.net/campaigns/search/", true)

;-------------------------------------------------------------------------------
; Things to say when Arsenal transforms
;-------------------------------------------------------------------------------
arsenalChange := ["transmogrify!"
    , "metamorphose!"
    , "evolve!"
    , "alter yourself!"
    , "modify!"
    , "transfigure!"
    , "choose a new form!"
    , "warp!"
    , "reform!"
    , "let's shift gears!"
    , "convert!"
    , "mutate!"
    , "transmute!"
    , "we have to adapt!"
    , "modulate!"
    , "refashion!"
    , "give yourself a make over!"
    , "innovate!"
    , "renew thyself!"
    , "metamorphosize!"
    , "gimme some variety!"
    , "transform!"
    , "recast your mold!"
    , "power morph!"
    , "make some tweaks!"
    , "revamp!"
    , "variegate!"
    , "turn over a new leaf!"
    , "switch!"
    , "reshape!"
    , "regenerate!"
    , "I feel like a change!"]

;-------------------------------------------------------------------------------
; Set Icon
;-------------------------------------------------------------------------------

I_Icon = .\ico\bugbear.ico
ICON [I_Icon]                        ;Changes a compiled script's icon (.exe)
if I_Icon <>
IfExist, %I_Icon%
	Menu, Tray, Icon, %I_Icon%   ;Changes menu tray icon 

;-------------------------------------------------------------------------------
; Check if Roll20 is currently open. Open it if the user requests it
;-------------------------------------------------------------------------------

FindWindowIDs()

if (!w.game.id) {
    result := MsgBoxEx("Would you like to open it in a browser, continue without it, or exit this app?", "Roll20 Game Not Found", "Open|Continue|Exit", 2)

    If (result = "Open") {
        OpenWindows()
    } Else If (result = "Continue") {
        gMute := 1
    } Else { ; (Result == "Exit" or Result == "Cancel") 
        ExitApp
    } 
}

;-------------------------------------------------------------------------------
; Import CSV data
;-------------------------------------------------------------------------------

LoadCSV(name, sort := "")
{
	;This makes sure the settings and data files exist. If there is a new xlsx version it automatically converts it to csv so it can be loaded
	MakeCSV(name, sort)
	;Import CSV file and use the data to create an object 
	return ObjCSV_CSV2Collection(".\csv\" . name . ".csv", "") 
}

weaponObj := LoadCSV("Arsenal", "B2")
chargeObj := LoadCSV("charge") 
settingsObj := LoadCSV("settings")

;flatten the settings object into a single array of key:value pairs
settings := {}
for index, line in settingsObj
{
	settings[line.Key] := line.Value
}
; Make sure we have default settings if the settings csv is not loaded or corrupt
if (!settings.changeCounter) {
	Random, changeCounter, 0, % arsenalChange.Length() + 5
	settings.changeCounter := changeCounter
}
if (!settings.currentWeapon) {
	settings.currentWeapon := "Vorpal Scimitar"
}

ExitFunc(exitReason,exitCode)
{
	global settings, chargeObj
	settings.currentWeapon := GetCurrentWeapon().Weapon
	
	; Regenerate the settings object
	settingsObj := {}
	for key, value in settings
	{
		settingsObj.Push({Key: key, Value: value})
	}
    ObjCSV_Collection2CSV( settingsObj, ".\csv\settings.csv", 1, "Key, Value", 0, 1 ) 
    ObjCSV_Collection2CSV( chargeObj, ".\csv\charge.csv", 1, "Weapon, Charge, TotalCharge, RechargeDie, RechargeMod, Slot1, Slot2, Slot3, Slot4, Slot5", 0, 1 ) 
	
    ; Do not call ExitApp -- that would prevent other OnExit functions from being called.
}

;-------------------------------------------------------------------------------
; Create weapon reference objects from CSV data
;-------------------------------------------------------------------------------

count := 0
a := [] ;Simple array of objects containing weapon info
arsenalKey:= {} ;Associative array linking weapon names (key) to their position in the master weapon list (value)
Loop, % weaponObj.Length() 
{
    if (weaponObj[A_Index].Disable != 1) {
		if (weaponObj[A_Index].Weapon = offhandWeapon){
            offhand := A_Index
		} else if (weaponObj[A_Index].Disable != 2) {
            count++
            weaponlist .= weaponObj[A_Index].Weapon "|"
            a.Push(weaponObj[A_Index])
            arsenalKey[weaponObj[A_Index].Weapon] := count
        }
    }
}
if (offhand) {
    count ++
    a.Push(weaponObj[offhand])
    arsenalKey[weaponObj[offhand].Weapon] := count
    offhand := count
}

;-------------------------------------------------------------------------------
; Create GUI
;-------------------------------------------------------------------------------

; ARSENAL Current Form
Gui, Add, GroupBox, x24 y2 w516 h109, Arsenal Current Form
Gui, Add, Text, x32 y20 w478 h31 +Center vgDescription, 
Gui, Add, Text, x32 y47 w478 h23 +0x200 +Center vgCurrentInfo, 
Gui, Font, s11
Gui, Add, Combobox, x127 y70 w225 vgCurrent gAutoComplete, % StrReplace(weaponlist, settings.currentWeapon . "|", settings.currentWeapon . "||")
Gui, Font
Gui, Add, CheckBox, x32 y70 w93 h23 cBlack +Hidden vgVersatile gVersatile, Versatile
Gui, Add, CheckBox, x368 y70 w168 h23 cBlack +Hidden vgCondition gCondition, Condition
Gui, Add, Button, x368 y70 w160 h28 +Hidden vgSpecial gSpecial, Special

; ARSENAL New Form
Gui, Add, GroupBox, x24 y118 w516 h88, Arsenal New Form
Gui, Font, s11
Gui, Add, Combobox, x127 y168 w225 vgNew gAutoComplete, %weaponlist%
Gui, Font
Gui, Add, Button, x32 y166 w80 h28 vgFilter gFilter, Filter
;Gui Add, Text, x32 y138 w328 h23 +0x200 +Center +Border vnewInfo, Weapon Info
Gui, Add, Text, x32 y138 w328 h23 +0x200 vgNewInfo, 
Gui, Add, CheckBox, x374 y138 w148 h23 cBlack vgArsenalTransformed, Arsenal Transformed
Gui, Add, Button, x368 y166 w160 h28 vgTransformArsenal gTransformArsenal, Transform Arsenal

; BATTLE PHASE BUTTONS
Gui, Add, Button, x34 y216 w100 h28 gLongRest, Long Rest
Gui, Add, Button, x160 y214 w238 h32 gBeginNewRound, Begin New Round
Gui, Add, Button, x424 y216 w100 h28 gEndTurn, End Turn

; BATTLE STATE
Gui, Add, GroupBox, x24 y256 w152 h313 , Battle State
Gui, Add, CheckBox, x38 y280 w120 h23 cBlack vgRageActive , Rage Active
Gui, Add, CheckBox, x38 y312 w130 h23 cBlack vgSneakAttackMade , Sneak Attack Made
Gui, Add, Radio, x38 y352 w120 h23 cBlack vgAdvantage, Advantage
Gui, Add, Radio, x38 y376 w120 h23 cBlack +Checked, None
Gui, Add, Radio, x38 y400 w120 h23 cBlack , Disadvantage
Gui, Add, CheckBox, x38 y440 w120 h23 cBlack vgAllySupport, Ally Support
Gui, Add, CheckBox, x38 y472 w120 h23 cBlack vgBugbearSurprise gBugbearSurprise, Bugbear Surprise
Gui, Add, CheckBox, x38 y504 w120 h23 cBlack vgAssassinate gAssassinate, Surprise Assassinate
Gui, Add, CheckBox, x38 y536 w120 h23 cBlack vgPoisonedBlade, Poisoned Blade

; ACTIONS
Gui, Add, GroupBox, x192 y256 w200 h360 , Action
Gui, Add, Radio, x208 y280 w120 h23 vgAction +Checked, No action made
Gui, Add, Radio, x208 y304 w120 h23 , First Attack
Gui, Add, Radio, x208 y328 w168 h23 , Second Attack or Full Action
Gui, Add, Button, x208 y368 w168 h28 vgAttack gAttack, ATTACK!
Gui, Add, Button, x208 y408 w168 h28 vgImprovisedAttack gImprovisedAttack, Improvised Attack
Gui, Add, Button, x208 y448 w80 h28 vgGrapple gGrapple, Grapple
Gui, Add, Button, x296 y448 w80 h28 vgShove gShove, Shove
Gui, Add, Button, x208 y488 w80 h28 vgShield gShield, Shield
Gui, Add, Button, x296 y488 w80 h28 vgDodge gDodge, Dodge
Gui, Add, Button, x208 y528 w80 h28 vgDisengage gDisengage, Disengage
Gui, Add, Button, x208 y568 w80 h28 vgDash gDash, Dash
Gui, Add, Button, x296 y528 w80 h28 vgHide gHide, Hide
Gui, Add, Button, x296 y568 w80 h28 vgOtherAction gOtherAction, Other Action

; BONUS ACTIONS
Gui, Add, GroupBox, x408 y256 w132 h187 , Bonus Action
Gui, Add, CheckBox, x424 y277 w115 h23 cBlack vgBonusActionUsed, Bonus Action Used
Gui, Add, Button, x424 y307 w100 h28 vgStartRage gStartRage, Start Rage
Gui, Add, Button, x424 y350 w100 h28 vgOffhandAttack gOffhandAttack, Offhand Attack
Gui, Add, Button, x424 y393 w100 h28 vgCunningAction gCunningAction, Cunning Action

; LEFT
Gui, Add, GroupBox, x24 y586 w107 h184 , Left Hand
Gui, Add, Radio, x36 y656 w75 h23 vgLeft +Checked, Empty
Gui, Add, Radio, x36 y680 w92 h23 , Two-handed
Gui, Add, Radio, x36 y704 w75 h22 , Grapple 1
Gui, Add, Radio, x36 y728 w75 h23 , Other
Gui, Add, Radio, x36 y608 w75 h23 , Shortblade
Gui, Add, Radio, x36 y632 w75 h23 , Shield

; RIGHT
Gui, Add, GroupBox, x143 y632 w103 h138 , Right Hand
Gui, Add, Radio, x152 y656 w75 h23 vgRight, Empty
Gui, Add, Radio, x152 y680 w75 h23 +Checked, Arsenal
Gui, Add, Radio, x152 y704 w75 h23 , Grapple 2
Gui, Add, Radio, x152 y728 w75 h23 , Other

; REACTIONS
Gui, Add, GroupBox, x262 y632 w129 h138 , Reaction
Gui, Add, CheckBox, x280 y653 w95 h23 cBlack vgReactionUsed, Reaction Used
Gui, Add, Button, x275 y685 w100 h28 vgUncannyDodge gUncannyDodge, Uncanny Dodge
Gui, Add, Button, x275 y725 w100 h28 vgOpportunityAttack gOpportunityAttack, Opportunity Attack

; FREE ACTIONS
Gui, Add, GroupBox, x408 y463 w132 h242 , Free Actions
Gui, Add, CheckBox, x424 y484 w111 h23 cBlack vgObjectInteract, Object Interact
Gui, Add, Button, x424 y514 w100 h28 vgStowArsenal gStowArsenal, Stow Arsenal
Gui, Add, Button, x424 y546 w100 h28 vgDropArsenal gDropArsenal, Drop Arsenal
Gui, Add, Button, x424 y587 w100 h28 vgStowShortblade gStowShortblade, Stow Shortblade
Gui, Add, Button, x424 y619 w100 h28 vgDropShortblade gDropShortblade, Drop Shortblade
Gui, Add, Button, x424 y660 w100 h28 vgRecklessAttack gRecklessAttack, Reckless Attack

;Gui, Add, Button, x424 y725 w100 h28 gLongRest, Long Rest
Gui, Add, CheckBox, x424 y715 w140 h23 cBlack vgArrange gArrange +Checked%gArrange%, Auto-arrange
Gui, Add, CheckBox, x424 y740 w100 h23 cBlack vgMute +Checked%gMute%, Mute Output

Gui, Add, StatusBar, vgStatus, AC: 20
Gui, Show, w560 h815, % consoleName
;OnMessage(0x201, "WM_LBUTTONDOWN")
OnMessage(0x06, "WM_ACTIVATE")
OnMessage(0x111, "WM_COMMAND")

InitChargeWeapons() 
WM_COMMAND()
WM_ACTIVATE(1)
; Initialisation of app finished. Wait for user input into GUI
Return
 
 

;-------------------------------------------------------------------------------
; ADMIN FUNCTIONS
;-------------------------------------------------------------------------------

GuiEscape:
GuiClose:
    ExitApp

AutoComplete:
    CbAutoComplete() ;AutoComplete Combobox Library 
return

Arrange:
if (gArrange)
    activatePrimed := true
    WM_ACTIVATE(1) 
return

BeginNewRound:
	ResetGUI()
    hand := ""
    if (gLeft = 2) {
        hand .= "gripping Arsenal with both hands "
    } else {
        if (gLeft = 5) {
            hand .= "holding his shortblade "
        } else if (gLeft = 6) {
            hand .= "holding a shield "
        } if (gLeft = 3 or gRight = 3) {
            if (hand)
                hand .= "and "
            if (gLeft = 3 and gRight = 3) {
                hand .= "grappling two creatures "
            } else {
                hand .= "grappling a creature "
            }
        }
        if (gRight = 2) {
            if (hand)
                hand .= "and "
            if (!hand or gLeft = 3)
                hand .= "holding "
            hand .= "Arsenal "
        }
    }
    if (gRight == 2) {
        hand .= "in ````" . a[arsenalKey[gCurrent]].Weapon . "```` form "
    }
    hand := (hand) ? ", " . TrimEnd(hand) : ""
	
	status := (gRight = 2) ? TrimEnd(StrReplace(StatusText(false), "   ", ", "), -2) : ""
	status := (status) ? " ````" . status ".````" : ""
	
	SendCommand("/me begins new round" . hand . "." . status)
    
    WM_COMMAND() 
return

EndTurn:
	SendCommand("/me ends his turn")
return

LongRest:
	result := MsgBoxEx("A long rest will recharge your weapons and reset your 1/day abilities, do you wish to continue?", "Long Rest", "Continue*|Cancel", [1, ".\ico\bedroll.ico"], "", "", WinExist("A"))
	if (result = "Continue") {
		SendCommand("/me takes a long rest")
		RechargeWeapons()
		ResetGUI()
		WM_COMMAND() 
	}
return

ResetGUI() {
	global
    GuiControl, 1: , gArsenalTransformed, 0
    GuiControl, 1: , gAction, 1
    GuiControl, 1: , gBonusActionUsed, 0
    GuiControl, 1: , gReactionUsed, 0
    GuiControl, 1: , gObjectInteract, 0
    GuiControl, 1: , gSneakAttackMade, 0
    GuiControl, 1: , gBugbearSurprise, 0
    GuiControl, 1: , gAssassinate, 0
    GuiControl, 1: , gRecklessAttack, Reckless Attack
    GuiControl, 1:Enable, gRecklessAttack
	
	oncePerTurnAbility := {}
	if (InStr(GetCurrentWeapon().Weapon, "Defender"))
		GetCurrentWeapon().AC := 0
}

Filter:
result := MsgBoxEx("Choose a filter", "Filter Weapons", "No Filter*|Finesse|Non-Finesse|Ranged|Light", 0, "", "", WinExist("A"), 0, "s9 c0x000000", "Segoe UI")
if (result = "No Filter" or result = "Cancel") {
    result := "Filter"
}
GuiControl, 1:Text, gFilter, % result

weaponlist := "|"
Loop, % a.Length() 
{
    If (result = "Filter"
    or (result = "Finesse" and a[A_Index].Sneak >= 1)
    or (result = "Non-Finesse" and !a[A_Index].Sneak)
    or (result = "Ranged" and a[A_Index].Range)
    or (result = "Light" and a[A_Index].Light)) 
    {
		if (!a[A_Index].Disable)
			weaponlist .= a[A_Index].Weapon "|"
    }
    GuiControl, 1:, gNew, % weaponlist
    GuiControl, 1:Focus, gNew
}
return

InitChargeWeapons() {
	global a
	Loop, % a.Length() 
	{
		i := A_Index
		if (a[i].TotalCharge) {
			SetTotalCharge(a[i].Weapon, a[i].TotalCharge, a[i].RechargeDie, a[i].RechargeMod)
		}
	}
}
RechargeWeapons() {
	global chargeObj
	Loop, % chargeObj.Length() 
	{
		i := A_Index
		if (chargeObj[i].TotalCharge) {
			if chargeObj[i].Charge < 0
				chargeObj[i].Charge := 0
				
			chargeObj[i].Charge += Random(1, chargeObj[i].RechargeDie) + chargeObj[i].RechargeMod
			
			if chargeObj[i].Charge > chargeObj[i].TotalCharge
				chargeObj[i].Charge := chargeObj[i].TotalCharge
		}
		if chargeObj[i].Slot1
			chargeObj[i].Slot1 := ""
		if chargeObj[i].Slot2
			chargeObj[i].Slot2 := ""
		if chargeObj[i].Slot3
			chargeObj[i].Slot3 := ""
		if chargeObj[i].Slot4
			chargeObj[i].Slot4 := ""
		if chargeObj[i].Slot5
			chargeObj[i].Slot5 := ""
	}
}
SetTotalCharge(name, total, rechargeDie, rechargeMod := 0){
	weapon := GetChargeWeapon(name, true)
	weapon.TotalCharge := total
	weapon.RechargeDie := rechargeDie
	weapon.RechargeMod := rechargeMod
	if (!weapon.Charge)
		weapon.Charge := total
	return
}
; Finds a weapon by name in the charge object, returns false if not found
; If create parameter is set to true, then if the charge name is not found it is created as a new entry and returned
GetChargeWeapon(name, create := false){
	global chargeObj
	for index, line in chargeObj
	{
		if (line.Weapon = name) {
			return line
		}
	}
	if create
		return chargeObj[chargeObj.Push({Weapon: name})]
	else
		return false
}

Random(min, max){
	Random, out, min, max
	return out
}

;-------------------------------------------------------------------------------
; WEAPON SPECIAL ABILITY BUTTONS
;-------------------------------------------------------------------------------

Special:
i := arsenalKey[gCurrent]
chargeWeapon := GetChargeWeapon(a[i].Weapon)

if (a[i].Spells) {
	buttons := a[i].Spells . "|"
} else if (a[i].Special) {
	buttons := a[i].Special . "|"
} else {
	buttons := ""
}
if (a[i].TotalCharge) {
	buttons .= "Recharge|"
}

dailyPowersUsed := ""
Loop, 5  {
	if (chargeWeapon["Slot" . A_Index])
		dailyPowersUsed .= "* " . chargeWeapon["Slot" . A_Index] . "`n"
}

result := MsgBoxEx(StrReplace(a[i].Detail,"``n","`n") . (dailyPowersUsed ? "`n`nDaily Powers Used:`n" . dailyPowersUsed : "")
, a[i].Weapon . " - Special Abilities"
, buttons . "Share to Chat|Cancel*"
, [1, ".\ico\weaponmagic.ico"], "", "", WinExist("A"))
weaponHeading := a[i].Weapon . " Special Ability"

if (result = "Share to Chat") {
	SendDescription(a[i].Detail, a[i].Weapon, WeaponText(i)) ;adv = always 
	
} else if (result = "Recharge") {
	Gui, +OwnDialogs
	InputBox, userInput, Recharge, % "Current charge: " . chargeWeapon.Charge . "/" . a[i].TotalCharge . "`nEnter new charge level" , , 350, 150
	if userInput is Integer
		chargeWeapon.Charge := userInput

} else if (result = "Reroll") {
	if (ExpendDailyUse(a[i].Weapon, result, 1)) {
		Gui, +OwnDialogs ;Add this line just before input box to make it modal
		InputBox, userInput, Luck Tuck Reroll, Please enter the modifier for the reroll you are making, , 350, 150
		if (!ErrorLevel){
			SendCommand(GenerateAttack(weaponHeading
			, AttackRoll(20, userInput, "Mod", 0, 0, "always")
			, false ; Damage 1
			, false ; Damage 2
			, false ; Save
			, "If the sword is on your person, you can call on its luck (no action required) to reroll one attack roll, ability check, or saving throw you dislike once per day. You must use the second roll."
			, ""
			, "1/day"))
		}
	}
	
} else if (result = "Form Blade") {
	if (ExpendBonusAction(result)) {
		SendDescription("While grasping the hilt, you can use a bonus action to cause a blade of pure radiance to spring into existence, bathing a 15 ft radius in pure daylight and dim light for an additional 15ft. While the blade persists, you can use an action to expand or reduce its radius of bright and dim light by 5 feet each, to a maximum of 30 feet each or a minimum of 10 feet each.", result, weaponHeading)
	}
	
} else if (result = "Extinguish Flames") {
	SendDescription("When you draw this weapon, you can extinguish all nonmagical flames within 30 feet of you. This property can be used no more than once per hour.", result, weaponHeading)

} else if (result = "Compass") {
if (ExpendBonusAction(result)) {
		SendDescription("While resting the dagger on your palm, the blade spins to point north.", result, weaponHeading)
	}
	
} else if (result = "Dimension Door") {
	if (ExpendBonusAction(result) and ExpendDailyUse(a[i].Weapon, result, 1)) {
		roll = @{%charName%|wtype}&{template:spell} {{level=conjuration 4}}  {{name=Dimension Door}} {{castingtime=1 bonus action}} {{range=500 feet}} {{target=See text}} {{v=1}} 0 0 {{material=}} {{duration=Instantaneous}} {{description=You teleport yourself from your current location to any other spot within range. You arrive at exactly the spot desired. It can be a place you can see, one you can visualize, or one you can describe by stating distance and direction, such as “200 feet straight downward” or “upward to the northwest at a 45-degree angle, 300 feet.” You can bring along objects as long as their weight doesn’t exceed what you can carry. You can also bring one willing creature of your size or smaller who is carrying gear up to its carrying capacity. The creature must be within 5 feet of you when you cast this spell. If you would arrive in a place already occupied by an object or a creature, you and any creature traveling with you each take 4d6 force damage, and the spell fails to teleport you.}} {{athigherlevels=}} 0 {{innate=1/day}} 0 @{%charName%|charname_output}
		SendCommand(roll)
	}
		
} else if (result = "Spider Compulsion") {
	if (ExpendBonusAction(result) and ExpendDailyUse(a[i].Weapon, result, 2)) {
		roll = @{%charName%|wtype}&{template:spell} {{level=enchantment 4}}  {{name=Spider Compulsion}} {{castingtime=1 bonus action}} {{range=90 ft}} {{target=}} {{v=1}} {{s=1}} 0 {{material=}} {{duration=Up to 1 minute}} {{description=Spiders of the type Beast of your choice that you can see within range and that can hear you must make a DC15 Wisdom saving throw. A target automatically succeeds on this saving throw if it can't be charmed. On a failed save, a target is affected by this spell. Until the spell ends, you can use a bonus action on each of your turns to designate a direction that is horizontal to you. Each affected target must use as much of its movement as possible to move in that direction on its next turn. It can take its action before it moves. After moving in this way, it can make another Wisdom saving to try to end the effect. A target isn't compelled to move into an obviously deadly hazard, such as a fire or pit, but it will provoke opportunity attacks to move in the designated direction.}} {{athigherlevels=}} 0 {{innate=1/day}} {{concentration=1}} @{%charName%|charname_output}
		SendCommand(roll)
	}		
		
} else if (result = "Conjure Earth Elemental") {
	if (ExpendAction(result) and ExpendDailyUse(a[i].Weapon, result, 1)) {
		roll = 
		(
		@{%charName%|wtype}&{template:spell} {{level=conjuration 5}} {{name=Conjure Earth Elemental}} {{castingtime=1 action}} {{range=90 feet}} {{target=A 10-foot cube within range}} {{v=1}} {{s=1}} 0 {{material=}} {{duration=Up to 1 hour}} {{description=You call forth an elemental servant. Choose an area of earth that fills a 10-foot cube within range. An earth elemental appears in an unoccupied space within 10 feet of it, rising up from the ground. The elemental disappears when it drops to 0 hit points or when the spell ends.
		The elemental is friendly to you and your companions for the duration. Roll initiative for the elemental, which has its own turns. It obeys any verbal commands that you issue to it (no action required by you). If you don't issue any commands to the elemental, it defends itself from hostile creatures but otherwise takes no actions.
		If your concentration is broken, the elemental doesn't disappear. Instead, you lose control of the elemental, it becomes hostile toward you and your companions, and it might attack. An uncontrolled elemental can't be dismissed by you, and it disappears 1 hour after you summoned it. The GM has the elemental’s statistics.}} {{athigherlevels=}} 0 {{innate=1/day}} {{concentration=1}} @{%charName%|charname_output}
		)
		SendCommand(roll)
	}

} else if (result = "Fabricate") {
	if (ExpendAction(result)) {
		if (!chargeWeapon.Slot2 or ExpendDailyUse(a[i].Weapon, result, 2)) {
			rollResult := Random(1, 6)
			roll = 
			(
			@{%charName%|wtype}&{template:spell} {{level=transmutation 4}}  {{name=Fabricate}} {{castingtime=1 action}} {{range=120 feet}} {{target=Raw materials that you can see within range}} {{v=1}} {{s=1}} 0 {{material=}} {{duration=Instantaneous}} {{description=You convert raw materials into products of the same material. For example, you can fabricate a wooden bridge from a clump of trees, a rope from a patch of hemp, and clothes from flax or wool. Choose raw materials that you can see within range. You can fabricate a Large or smaller object (contained within a 10-foot cube, or eight connected 5-foot cubes), given a sufficient quantity of raw material. If you are working with metal, stone, or another mineral substance, however, the fabricated object can be no larger than Medium (contained within a single 5-foot cube). The quality of objects made by the spell is commensurate with the quality of the raw materials. Creatures or magic items can’t be created or transmuted by this spell. You also can’t use it to create items that ordinarily require a high degree of craftsmanship, such as jewelry, weapons, glass, or armor, unless you have proficiency with the type of artisan’s tools used to craft such objects.}} {{athigherlevels=}} 0 {{innate=
			On a roll of 1-5 on a d6, you can't cast it again until the next dawn.
			Roll result: ````%rollResult%````}} 0 @{%charName%|charname_output}
			)
			SendCommand(roll)
			if(rollResult < 6) {
				ExpendDailyUse(a[i].Weapon, result, 2, false)
			}			
		}
	}
		
} else if (result = "Travel the Depths" or result = "Teleport") {
	if (ExpendAction(result) and ((result = "Travel the Depths" and ExpendDailyUse(a[i].Weapon, result, 3)) or ExpendCharge(i, 3))) {
		if (result = "Travel the Depths") {
			roll = You can use an action to touch the axe to a fixed piece of dwarven stonework and cast the teleport spell from the axe. If your intended destination is underground, there is no chance of a mishap or arriving somewhere unexpected.`n 
			innate = 1/3days
		} else {
			roll = 
			innate := "`n3 charges used, " . chargeWeapon.Charge . " charges remaining"
		}
		roll = 
		(
		@{%charName%|wtype}&{template:spell} {{level=conjuration}}  {{name=%result%}} {{castingtime=1 action}} {{range=10 feet}} {{target=You and up to eight willing creatures that you can see within range, or a single object that you can see within range}} {{v=1}} 0 0 {{material=}} {{duration=Instantaneous}} {{description=%roll%This spell instantly transports you and up to eight willing creatures of your choice that you can see within range, or a single object that you can see within range, to a destination you select. If you target an object, it must be able to fit entirely inside a 10-foot cube, and it can’t be held or carried by an unwilling creature. The destination you choose must be known to you, and it must be on the same plane of existence as you. See page 281 of PHB for full details.}} {{athigherlevels=}} 0 {{innate=%innate%}} 0 @{%charName%|charname_output}
		)
		SendCommand(roll)
	}

} else if (result = "Thunder" and  ExpendDailyUse(a[i].Weapon, result, 2)) {
	desc = When you hit with a melee attack using the staff, you can cause the staff to emit a crack of thunder, audible out to 300 feet. The target you hit must succeed on a DC 17 Constitution saving throw or become stunned until the end of your next turn.		
	SendCommand(GenerateAttack(result
	, false ; Attack Roll
	, false ; Damage 1
	, false ; Damage 2
	, SaveRoll("Constitution", "Stunned on fail", "17") ; Save
	, desc ;desc
	, "Melee" ;range
	, "" ;spell
	, "1/day"))
	
} else if (result = "Lightning Strike") {
	if (ExpendAction(result) and ExpendDailyUse(a[i].Weapon, result, 3)) {
		desc = You can use an action to cause a bolt of lightning to leap from the staff's tip in a line that is 5 feet wide and 120 feet long. Each creature in that line must make a DC 17 Dexterity saving throw, taking 9d6 lightning damage on a failed save, or half as much damage on a successful one.	
		SendCommand(GenerateAttack(result
		, false ; Attack Roll
		, DamagePrepare(AddDamage("9d6", "lightning", "lightning")) ; Damage 1
		, false ; Damage 2
		, SaveRoll("Dexterity", "Half damage on success", "17") ; Save
		, desc ;desc
		, "5ft line, 120ft long" ;range
		, "" ;spell
		, "1/day"))
	}
} else if (result = "Thunderclap") {
	if (ExpendAction(result) and ExpendDailyUse(a[i].Weapon, result, 4)) {
		desc = You can use an action to cause the staff to issue a deafening thunderclap, audible out to 600 feet. Each creature within 60 feet of you (not including you) must make a DC 17 Constitution saving throw. On a failed save, a creature takes 2d6 thunder damage and becomes deafened for 1 minute. On a successful save, a creature takes half damage and isn't deafened.
		SendCommand(GenerateAttack(result
		, false ; Attack Roll
		, DamagePrepare(AddDamage("2d6", "thunder", "thunder")) ; Damage 1
		, false ; Damage 2
		, SaveRoll("Constitution", "Deafened for 1 minute on fail`nHalf damage on success", "17") ; Save
		, desc ;desc
		, "60ft" ;range
		, "" ;spell
		, "1/day"))
	}
} else if (result = "Thunder and Lightning") {
	if (ExpendAction("Thunderbolts and Lightning Very Very Frightening") and ExpendDailyUse(a[i].Weapon, result, 5)) {
		desc = The staff issues a deafening thunderclap, audible out to 600 feet. Each creature within 60 feet of you (not including you) must make a DC 17 Constitution saving throw. On a failed save, a creature takes 2d6 thunder damage and becomes deafened for 1 minute. On a successful save, a creature takes half damage and isn't deafened.
		BufferCommand("Galileo Figaro Magnifico!")
		BufferCommand(GenerateAttack("Thunderbolt"
		, false ; Attack Roll
		, DamagePrepare(AddDamage("2d6", "thunder", "thunder")) ; Damage 1
		, false ; Damage 2
		, SaveRoll("Constitution", "Deafened for 1 minute on fail,`nhalf damage on success", "17") ; Save
		, desc ;desc
		, "60ft" ;range
		, "" ;spell
		, "" )) ;innate
		desc = A bolt of lightning to leaps from the staff's tip in a line that is 5 feet wide and 120 feet long. Each creature in that line must make a DC 17 Dexterity saving throw, taking 9d6 lightning damage on a failed save, or half as much damage on a successful one.	
		SendCommand(GenerateAttack("Lightning Strike"
		, false ; Attack Roll
		, DamagePrepare(AddDamage("9d6", "lightning", "lightning")) ; Damage 1
		, false ; Damage 2
		, SaveRoll("Dexterity", "Half damage on success", "17") ; Save
		, desc ;desc
		, "5ft line, 120ft long" ;range
		, "" ;spell
		, "" )) ;innate
	}

} else if (result = "Shatter") {
	if (ExpendAction(result) and ExpendCharge(i)) {
		desc = A sudden loud ringing noise, painfully intense, erupts from a point of your choice within range. Each creature in a 10-foot-radius sphere centered on that point must make a Constitution saving throw. A creature takes 3d8 thunder damage on a failed save, or half as much damage on a successful one. A creature made of inorganic material such as stone, crystal, or metal has disadvantage on this saving throw. A nonmagical object that isn't being worn or carried also takes the damage if it's in the spell's area
		
		SendCommand(GenerateAttack(result
		, false ; Attack Roll
		, DamagePrepare(AddDamage("3d8", "thunder", "thunder")) ; Damage 1
		, false ; Damage 2
		, SaveRoll("Constitution", "Half damage on success", "17") ; Save
		, desc ;desc
		, "60ft" ;range
		, "2" ;spell
		, "`n" . chargeWeapon.Charge . " charges remaining"))
	}
	
} else if (result = "Determine Distance to Surface") {
	if (ExpendAction(result) and ExpendCharge(i, 1)) {
		SendDescription("If you are underground or underwater, you can determine the distance to the surface.", result, weaponHeading . "*`n**````1 charge used, " . chargeWeapon.Charge . " charges remaining````*")
	}
	
} else if (result = "Sending") {
	if (ExpendAction(result) and ExpendCharge(i, 2)) {
		innate := "`n2 charges used, " . chargeWeapon.Charge . " charges remaining"
		roll = 
		(
		@{%charName%|wtype}&{template:spell} {{level=evocation 3}}  {{name=Sending}} {{castingtime=1 action}} {{range=Unlimited}} {{target=A creature with which you are familiar}} {{v=1}} {{s=1}} {{m=1}} {{material=A short piece of fine copper wire}} {{duration=1 round}} {{description=You send a short message of twenty-five words or less to a creature with which you are familiar. The creature hears the message in its mind, recognizes you as the sender if it knows you, and can answer in a like manner immediately. The spell enables creatures with Intelligence scores of at least 1 to understand the meaning of your message. You can send the message across any distance and even to other planes of existence, but if the target is on a different plane than you, there is a 5 percent chance that the message doesn’t arrive.}} {{athigherlevels=}} 0 {{innate=%innate%}} 0 @{%charName%|charname_output}
		)
		SendCommand(roll)
	}
	
} else if (result = "Frighten") {
	if (ExpendAction(result) and ExpendCharge(i)) {
		desc = You release a wave of terror. Each creature of your choice in a 30-foot radius extending from you must succeed on a DC 15 Wisdom saving throw or become frightened of you for 1 minute. While it is frightened in this way, a creature must spend its turns trying to move as far away from you as it can, and it can't willingly move to a space within 30 feet of you. It also can't take reactions. For its action it can use only the Dash action or try to escape from an effect that prevents it from moving. If it has nowhere it can move, the creature can use the Dodge action. At the end of each of its turns, a creature can repeat the saving throw, ending the effect on itself on a success.
		
		SendCommand(GenerateAttack(result
		, false ; Attack Roll
		, false ; Damage 1
		, false ; Damage 2
		, SaveRoll("Wisdom", "Frighten", "15") ; Save
		, desc ;desc
		, "30ft" ;range
		, "" ;spell
		, "`n" . chargeWeapon.Charge . " charges remaining"))
	}
	
} else if (result = "Dominate Sea Beast") {
	if (ExpendAction(result) and ExpendCharge(i)) {
		innate := chargeWeapon.Charge . " charges remaining"
		roll = 
		(
		@{%charName%|wtype}&{template:spell} {{level=enchantment 4}} {{name=Dominate Sea Beast}} {{castingtime=1 action}} {{range=60 feet}} {{target=A beast that you can see within range}} {{v=1}} {{s=1}} 0 {{material=}} {{duration=Up to 1 minute}} {{description=You attempt to beguile a beast that has an innate swimming speed and you can see within range. It must succeed on a 
		**````DC15 Wisdom saving throw````** or be charmed by you for the duration. If you or creatures that are friendly to you are fighting it, it has advantage on the saving throw. While the beast is charmed, you have a telepathic link with it as long as the two of you are on the same plane of existence. You can use this telepathic link to issue commands to the creature while you are conscious (no action required), which it does its best to obey. You can specify a simple and general course of action, such as “Attack that creature,” “Run over there,” or “Fetch that object.” If the creature completes the order and doesn’t receive further direction from you, it defends and preserves itself to the best of its ability. You can use your action to take total and precise control of the target. Until the end of your next turn, the creature takes only the actions you choose, and doesn’t do anything that you don’t allow it to do. During this time, you can also cause the creature to use a reaction, but this requires you to use your own reaction as well. Each time the target takes damage, it makes a new Wisdom saving throw against the spell. If the saving throw succeeds, the spell ends.}} {{athigherlevels=}} 0 {{innate=`n%innate%}} {{concentration=1}} @{%charName%|charname_output}
		)
		SendCommand(roll)
	}

} else if (result = "Explosion") {
	if (ExpendAction(result)) {
		desc = You hurl the weapon up to 120 feet to a point you can see. When it reaches that point, the weapon vanishes in an explosion, and each creature in a 20-foot-radius sphere centered on that point must make a DC 15 Dexterity saving throw, taking 6d6 fire damage on a failed save, or half as much damage on a successful one. You can use another action to cause the weapon to reappear in your empty hand. 
		
		SendCommand(GenerateAttack(result
		, false ; Attack Roll
		, DamagePrepare(AddDamage("6d6", "fire", "fire")) ; Damage 1
		, false ; Damage 2
		, SaveRoll("Dexterity", "Half damage on success", "15") ; Save
		, desc ;desc
		, "120ft" ;range
		, "" ;spell
		, "1/short rest"))
	}

} else if (RegExMatch(result, "Dominate (.*) Elemental", match) and ExpendAction(result) and ExpendDailyUse(a[i].Weapon, result, 1)) {
	DominateElemental(match1)
}
commandBuffer := ""
WM_COMMAND()
return

DominateElemental(type){

	desc = 
	(
	**Dominate %type% Elemental**
	Casting Time: 1 action
	Range: 60 feet
	Components: V, S
	Duration: Concentration, up to 1 hour	
	
	You attempt to beguile a %type% Elemental that you can see within range. It must succeed on a DC17 Wisdom saving throw or be charmed by you for the duration. If you or creatures that are friendly to you are fighting it, it has advantage on the saving throw.

	While the Elemental is charmed, you have a telepathic link with it as long as the two of you are on the same plane of existence. You can use this telepathic link to issue commands to the Elemental while you are conscious (no action required), which it does its best to obey. You can specify a simple and general course of action, such as "Attack that creature," "Run over there," or "Fetch that object." If the Elemental completes the order and doesn't receive further direction from you, it defends and preserves itself to the best of its ability.

	You can use your action to take total and precise control of the target. Until the end of your next turn, the Elemental takes only the actions you choose, and doesn't do anything that you don't allow it to do. During this time, you can also cause the Elemental to use a reaction, but this requires you to use your own reaction as well.

	Each time the target takes damage, it makes a new Wisdom saving throw against the spell. If the saving throw succeeds, the spell ends.
	)
	
	name := "Dominate " . type . " Elemental"
	
	SendCommand(GenerateAttack(name
		, false ; Attack Roll
		, false ; Damage 1
		, false ; Damage 2
		, SaveRoll("Wisdom", name, "17") ; Save
		, desc
		, ""
		, ""
		, "1/day"))
}

;-------------------------------------------------------------------------------
; WEAPON SPECIAL ABILITY HELPER FUNCTIONS
;-------------------------------------------------------------------------------

ExpendBonusAction(name){
	global gBonusActionUsed
	if (gBonusActionUsed) {
		Gui +OwnDialogs
		MsgBox 0x10, Error, You require a free bonus action to use %name%
		return false
	} else {
		GuiControl, 1:, gBonusActionUsed, 1
		gBonusActionUsed := 1
		BufferCommand("/me uses a bonus action to activate ````" . name . "````")
		return true
	}
}

ExpendAction(name){
	global gAction
	if (gAction > 1) {
		Gui +OwnDialogs
		MsgBox 0x10, Error, You require a free action to use %name%
		return false
	} else {
		GuiControl, 1: , Second Attack or Full Action, 1
        gAction = 3
		BufferCommand("/me uses an action to activate ````" . name . "````")
		return true
	}
}
ExpendCharge(index, spend := 1){
	global a
	chargeWeapon := GetChargeWeapon(a[index].Weapon, true)
	if (chargeWeapon.TotalCharge) {
		if (chargeWeapon.Charge - spend < 0){
			result := MsgBoxEx("This ability costs " . spend . " charges, but you only have " . chargeWeapon.Charge . " remaining. Do you wish to continue anyway or cancel?", "Not Enough Charges Remaining", "Continue|Cancel*", 2, "", "", WinExist("A"))
			
			if (result = "Continue") {
				chargeWeapon.Charge -= spend
				return true
			} else {
				return false
			}
		}
		chargeWeapon.Charge -= spend
		return true
	} else {
		MsgBox 0x10, Error, Charge level not found
		return false
	}
}
ExpendDailyUse(name, ability := true, slot := 1, notify := true){
	;global a
	chargeWeapon := GetChargeWeapon(name, true) 
	if (chargeWeapon[("Slot" . slot)] and notify) {
		
		result := MsgBoxEx("This ability has already been used and is limited to once per rest. Do you wish to continue anyway or cancel?", "Daily Ability", "Continue|Cancel*", 2, "", "", WinExist("A"))
		if (result != "Continue") {
			return false
		}
	}
	chargeWeapon[("Slot" . slot)] := ability
	/*for k, v in chargeWeapon
		str .= k . ": " . v . "`n"
	MsgBox % str
	*/
	return true
}

SendDescription(description, title, subtitle){
	str := "&{template:desc} {{desc="
	if title
		str .= "**"  . title "**`n"
	if subtitle
		str .= "*"  . subtitle "*`n"
	if title or subtitle
		str .= "***~~~~~~~~~~~~~~~~~~~~~~~***`n"
	str .= StrReplace(description, "``n", "`n") . "}}"
	
	SendCommand(str)
}

Condition:
GuiBatch("|gCondition", "cBlue")			
if (gCondition and a[arsenalKey[gCurrent]].Condition = "sworn enemy") {
	gAdvantage := 1
	GuiControl, 1:, gAdvantage, 1
}
WM_COMMAND()
return

Versatile:
GuiControl, , Two-handed, % !gVersatile ;opposite because WM_COMMAND fires first and changes it
if (gVersatile) 
    GuiControl, , Two-handed, 1
else
    GuiControl, , gLeft, 1
WM_COMMAND()
return

BugbearSurprise:
if (gBugbearSurprise) ;opposite because WM_COMMAND fires first and changes it
    GuiControl, , gAssassinate, 1
WM_COMMAND() 
Assassinate:
if (gAssassinate) ;opposite because WM_COMMAND fires first and changes it
    GuiControl, , gAdvantage, 1
WM_COMMAND() 
return

;--------------------------------------------------------
; CHANGING WEAPONS
;--------------------------------------------------------
   
TransformArsenal:
if (InStr(a[arsenalKey[gNew]].Weapon, "Defender"))
	a[arsenalKey[gNew]].AC := 0
changeAC := a[arsenalKey[gNew]].AC != a[arsenalKey[gCurrent]].AC
	
GuiControl, Choose, gCurrent, % arsenalKey[gNew]
gCurrent := gNew
GuiControl, Choose, gNew, 0
GuiControl, , gArsenalTransformed, 1
if ((a[arsenalKey[gNew]].Twohanded or a[arsenalKey[gNew]].Versatile) and gLeft < 3) {
    GuiControl, , Two-handed, 1
    GuiControl, , gVersatile, 1
} else {
    GuiControl, , gVersatile, 0
}
GuiControl, , gCondition, 0
if (a[arsenalKey[gNew]].Condition) {
	GuiBatch("gCondition", "cBlue")
} else {
	GuiBatch("|gCondition", "cBlue")
}

settings.changeCounter := Mod(settings.changeCounter, arsenalChange.Length()) + 1
BufferCommand("***Arsenal, " . arsenalChange[settings.changeCounter] . "***")
WeaponPanel(arsenalKey[gCurrent])
if (changeAC)
    BufferCommand(GetAcText())
SendCommand("")

WM_COMMAND() 
return

; returns comma-separated string of weapon properties
WeaponText(index, versatileShow := true) {
    global
    weaponText := a[index].Damage 
    if a[index].Modifier
        weaponText .= "+" . a[index].Modifier
    weaponText .= " " . a[index].Type
    if (!a[index].Condition and a[index].Damage2)
        weaponText .= " + " . a[index].Damage2 . " " . a[index].Type2
    weaponText .= ", "
    if a[index].TwoHanded
        weaponText .= "two-handed, "
    else if a[index].Versatile and versatileShow
        weaponText .= "versatile (" . a[index].Versatile . "), "
    if a[index].Light = 1
        weaponText .= "light, "
    if a[index].Sneak = 1
        weaponText .= "finesse, "
    else if a[index].Sneak = 2
        weaponText .= "ranged (" . a[index].Range . "), "
    if (a[index].Range and a[index].Sneak != 2)
        weaponText .= "thrown (" . a[index].Range . "), "
    return TrimEnd(weaponText, -2)
}

; created formatted weapon property box for Roll20
WeaponPanel(i){
    global a
    BufferCommand("@{Bo|wtype}&{template:traits} @{Bo|charname_output} {{name=" . a[i].Weapon . "}} {{source= " . a[i].Description . "}} {{description=" . WeaponText(i) . "}}")
}

;--------------------------------------------------------
; ATTACK BUTTONS
;--------------------------------------------------------

OpportunityAttack:
if (gRight != 2 and gLeft != 5) {
    DoAttack(true, false, true) ; Make with teeth and claws
} else {
    arsenal := (gRight == 2) ? "Arsenal|" : ""
    shortBlade := (gLeft == 5) ? "|Shortblade" : ""
    result := MsgBoxEx("What weapon do you use?", "Opportunity Attack!", arsenal . "Improvised" . shortBlade, [1, ".\ico\opportunity.ico"], "", "", WinExist("A"))
    if (result == "Arsenal") {
        DoAttack(false, false, true)
    } else if (result == "Shortblade") {
        DoAttack(false, true, true)
    } else if (result == "Improvised")  {
        DoAttack(true, false, true)
    } else {
        return
    }
}
GuiControl, , gReactionUsed, 1
gReactionUsed := 1
WM_COMMAND()
return

OffhandAttack:
if (gLeft = 1)
    StowDrawShortblade()
GuiControl, 1:, gBonusActionUsed, 1
gBonusActionUsed := 1
DoAttack(false, true)
return

ImprovisedAttack:
DoAttack(true)
return

Attack:
DoAttack()
return

;--------------------------------------------------------
; WEAPON ATTACK GENERATOR
;--------------------------------------------------------

DoAttack(improvised := false, offhandAttack := false, reaction := false, doubleAttack := false) {
    global
	description := ""
    damage := DamageRoll()
    damage2 := false
	save := false
    
    if (reaction) {
        BufferCommand("/me makes an opportunity attack")
    } else if (offhandAttack) {
        BufferCommand("/me makes an offhand attack")
    } else if (doubleAttack) {
        BufferCommand("/me uses the other end of the blade to attack again")
    } else if (gAction = 1) {
        BufferCommand("/me makes his first attack")
    } else {
        BufferCommand("/me makes his second attack")
    }

    i := arsenalKey[gCurrent]

    if (improvised) {
        if (gRight = 2) {
            weaponName := "Improvised Attack"
            magicMod := a[i].Modifier ? a[i].Modifier : 0
            type := "bludgeoning"
        } else {
            weaponName := "Teeth and Claws"
            magicMod := 0
            type := "slashing"
        }
        ability := "str"
        range := ""
        sneak := false
        AddDamage("1d4", type, type, damage)
        
    } else {
        if (offhandAttack) {
			i := offhand
		}
        weaponName := a[i].Weapon
        magicMod := a[i].Modifier ? a[i].Modifier : 0
        ability := a[i].Sneak = 2 ? "dex" : "str" ; Ranged weapons must use dexterity, everything else can use strength
        range := a[i].Range
        sneak := a[i].Sneak and !gSneakAttackMade and (gAdvantage = 1 or (gAllySupport and gAdvantage < 3))
        
        ; weapon damage
        if (a[i].Versatile and gLeft = 2) {
            AddDamage(a[i].Versatile, a[i].Type, a[i].Type, damage)
		} else if (doubleAttack) { ; The second attack of a Vorpal Double-Bladed Scimitar does less damage
            AddDamage("1d4", a[i].Type, a[i].Type, damage)
        } else {
            AddDamage(a[i].Damage, a[i].Type, a[i].Type, damage)
        }
        if (!offhandAttack and gPoisonedBlade) {
            AddDamage("1d6", "poison", "poison", damage)
        }
    }
    abilityMod := mods[ability]
    cs := gAssassinate ? 2 : 20
    vantage := (gAdvantage = 1) ? "advantage" : (gAdvantage = 2) ? "normal" : "disadvantage"
	
	
	; Some weapons can include extra damage/saving throws that is best included with the attack rather than making a seperate button for it like most weapon abilities
	if (!improvised and !offhandAttack) {
		
		if (a[i].Weapon = "Mace of Disruption" and gCondition) {
			description := "If a fiend or an undead has 25 hit points or fewer after being hit, it must make the saving throw above."
			save := SaveRoll("Wisdom", "Destroyed on fail.`nFrightened until the end of your`nnext turn on success", "15") 
			
		} else if (a[i].Weapon = "Javelin of Lightning" and gCondition and ExpendDailyUse(a[i].Weapon, a[i].Condition, 1)) {
			description := "It transforms into a bolt of lightning, forming a line 5 feet wide that extends out from you to a target within 120 feet. Each creature in the line excluding you and the target must make a DC 13 Dexterity saving throw, taking 4d6 lightning damage on a failed save, and half as much damage on a successful one. Make a ranged weapon attack against the target. On a hit, the target takes damage from the javelin plus 4d6 lightning damage."
			damage2 := DamagePrepare(AddDamage("4d6", "lightning", "lightning", damage2))
			save := SaveRoll("Dexterity", "Half damage on success", "13") 
			gCondition := 0
			GuiControl, 1: , gCondition, 0
			
		} else if (a[i].Weapon = "Rapier of Wounding" and gCondition and !oncePerTurnAbility[a[i].Weapon]) {
			oncePerTurnAbility[a[i].Weapon] := true
			static wounds := 0
			wounds++
			Gui, +OwnDialogs
			InputBox, userInput, Wounds, % "Enter how many wounds have been accumulated:" , , 350, 150, , , , , % wounds
			if userInput is Integer
				wounds := userInput
			description := "Hit points lost to this weapon's damage can be regained only through a short or long rest, rather than by regeneration, magic, or any other means. Once per turn, when you hit a creature with an attack using this magic weapon, you can wound the target. At the start of each of the wounded creature's turns, it takes 1d4necrotic damage for each time you've wounded it, and it can then make a DC 15 Constitution saving throw, ending the effect of all such wounds on itself on a success. Alternatively, the wounded creature, or a creature within 5 feet of it, can use an action to make a DC 15 Wisdom (Medicine) check, ending the effect of such wounds on it on a success."
			damage2 := DamagePrepare(AddDamage(wounds . "d4", "necrotic", "necrotic", damage2))
			save := SaveRoll("Constitution", "Wounds close on success", "15")
			
			
		} else if (InStr(a[i].Weapon, "Defender") and !oncePerTurnAbility[a[i].Weapon]) {
			oncePerTurnAbility[a[i].Weapon] := true
			Gui, +OwnDialogs
			InputBox, userInput, Defender Ability, % "Enter how much of the +3 weapon modifer you would like to convert to AC:" , , 350, 150, , , , , % 3
			if (userInput is Integer and !ErrorLevel)
				a[i].AC := Min(Max(userInput, 0), 3)
			
		} else if (a[i].Weapon = "Sword of the Paruns" and !oncePerTurnAbility[a[i].Weapon]) {
			oncePerTurnAbility[a[i].Weapon] := true
			description := "Immediately after you use the Attack action to attack with the sword, you can enable one creature within 60 feet of you to use its reaction to make one weapon attack."
			
		} else if (a[i].Weapon = "Staff of Thunder and Lightning" and gCondition) {
			ExpendDailyUse(a[i].Weapon, "Lightning Zap", 1, false)
			AddDamage("2d6", "lightning", "lightning", damage)
			gCondition := 0
			GuiControl, 1: , gCondition, 0
		}
		
		if (InStr(a[i].Weapon, "Defender") and a[i].AC) {
			description := "AC has been boosted to **" . GetAC() . "** until the start of the next turn"
			magicMod -= a[i].AC
		}
	}
	
	; calculate all damage types
	
	if (!offhandAttack)
		AddDamage(mods[ability], ability, "", damage)
		
    if (magicMod)
        AddDamage(magicMod, "magic", "", damage)
    if (!improvised and a[i].Damage2 and (!a[i].Condition or (a[i].Condition and gCondition))) {
        AddDamage(a[i].Damage2, a[i].Type2, (a[i].Type = a[i].Type2) ? "" : a[i].Type2, damage)
    }
    if (gRageActive and (a[i].Sneak != 2 or improvised) and (!gCondition or a[i].Condition != "thrown")) {
        AddDamage(rageDamage, "rage", "", damage)
    }
    if (sneak) {
        AddDamage(sneakAttackDamage, "snek attac", "snek", damage)
    }
    if (gBugbearSurprise) {
        AddDamage("2d6", "bugbear surprise", "surprise", damage)
    }
    if (InStr(a[i].Weapon, "Vorpal")) {
		if (gAssassinate) {
			description .= "Decapitation on Natural 20! `n(+[[12d8]] damage if head immune)"
		} else {
			damage.crit .= "12d8[DECAPITATION]+"
			range := "Decapitation on Nat20"
		}
    } else if (InStr(a[i].Weapon, "Dwarvish")) {
		limbSever := (Random(1, 20) >= 20)
		if (gAssassinate and !limbSever) {
			description .= "+14 damage on a Natural 20!"
		} else if (gAssassinate) {
			description .= "Perfect Strike chance active! +14 damage and severed limb on a Natural 20!"
		} else if (!limbSever) {
			damage.crit .= "14[sharp blade]+"
		} else { 
			description .= "Perfect Strike chance active! Weapon will sever enemy's limb on a Natural 20!"
			damage.crit .= "14[sharp blade]+"
		}
	}
	
	roll := GenerateAttack(weaponName
	, AttackRoll(cs, abilityMod, ability, prof, magicMod, vantage)
	, DamagePrepare(damage)
	, damage2 ; Damage 2
	, save ; Save
	, description
	, range) ;attackMod

    SendCommand(roll)
    if (improvised and !gBonusActionUsed and !reaction) {
        result := MsgBoxEx("If the attack hit, you may use your bonus action to make a grapple attempt. Do you?", "Tavern Brawling!", "Yes|No", [1, ".\ico\grapple.ico"], "", "", WinExist("A"))
        if (result == "Yes") {
            GuiControl, 1:, gBonusActionUsed, 1
			gBonusActionUsed := 1
            Grapple(false, true)
        }
    }
    if (a[i].Sneak and (!gSneakAttackMade or reaction) and !improvised and (gAdvantage = 1 or (gAllySupport and gAdvantage < 3))) {
        result := MsgBoxEx("Did the attack hit?", "Sneak Attack attempt made!", "Yes|No", [1, ".\ico\sneakattack.ico"], "", "", WinExist("A"))
        if (result == "Yes") {
            GuiControl, 1:, gSneakAttackMade, 1
			gSneakAttackMade := 1
            GuiControl, 1:, gBugbearSurprise, 0
			gBugbearSurprise := 0
        }
    } else if (gBugbearSurprise) {
		result := MsgBoxEx("Did the attack hit?", "Bugbear Surprise!", "Yes|No", [1, ".\ico\bugbear.ico"], "", "", WinExist("A"))
        if (result == "Yes") {
            GuiControl, 1:, gBugbearSurprise, 0
			gBugbearSurprise := 0
        }
	}
	if (a[i].Weapon = "Vorpal Double-Bladed Scimitar" and !gBonusActionUsed and !reaction and !offhandAttack) {
		result := MsgBoxEx("Would you like to use a bonus action to attack with the second blade?", "Vorpal Double-Bladed Scimitar", "Yes|No", [1, ".\ico\sneakattack.ico"], "", "", WinExist("A"))
        if (result == "Yes") {
            GuiControl, 1:, gBonusActionUsed, 1
			gBonusActionUsed := 1
			DoAttack(false, false, false, true)
        }
	}
	
    if (!offhandAttack and !reaction and !doubleAttack) {
        AdvanceAttack()
    }
    WM_COMMAND()
	
    return
}

;--------------------------------------------------------
; ATTACK HELPER FUNCTIONS
;--------------------------------------------------------

; Creates a simple object holding Attack information, that will be sent to GenerateAttack
AttackRoll(cs, abilityMod, ability, prof, magicMod, vantage){
	attack := []
	attack.cs := cs 
	attack.abilityMod := abilityMod
	attack.ability := ability
	attack.prof := prof
	attack.magicMod := magicMod
	attack.vantage := vantage
	return attack
}
; Creates a simple object holding saving throw information, that will be sent to GenerateAttack
SaveRoll(attr, desc, dc){
	save := []
	save.attr := attr
	save.desc := desc
	save.dc := dc
	return save
}
; Creates a simple object holding damage information, that will be populated by AddDamage and sent to GenerateAttack
DamageRoll(dice := "", crit := "", desc := ""){
	damage := []
	damage.dice := dice
	damage.crit := crit
	damage.desc := desc
	return damage
}
; Creates strings in Roll20 format for damage dice and crits and adds them to the damage object
; Can be called multiple times to add different types of damage for an attack
AddDamage(die, type, shortType := "", damage := false) {
	if (!IsObject(damage))
		damage := DamageRoll()
	
    damage.dice .= die . "[" . type . "]+"
    
    if die is Integer
        damage.crit .= ""
    else
        damage.crit .= die . "[" . type . "]+"
        
    if (shortType) {
        damage.desc .= Format("{:T}", shortType) . ", "
    }
	
	return damage
}
; Trims excess characters from the damage strings that were previously generated by AddDamage, use just before sending the damage object to GenerateAttack
DamagePrepare(damage) {
    damage.dice := TrimEnd(damage.dice)
    damage.crit := TrimEnd(damage.crit)
    damage.desc := TrimEnd(damage.desc, -2)
	return damage
}

;--------------------------------------------------------
; GENERIC ATTACK/DAMAGE/SAVE GENERATOR FOR ROLL20
;--------------------------------------------------------

; Combines attack, damage, savingthrow and other information into a single string in Roll20 format using the atkdmg template
GenerateAttack(weaponName
	, attack := false
	, damage1 := false
	, damage2 := false
	, save := false
	, description := ""
	, range := "" 
	, spell := ""
	, innate := ""
	, attackMod := false) {
	
	global charName
	
	if (attackMod is Integer){
		attackMod := "+" . attackMod
	} else if (attack) {
		attackMod := "+" . (attack.prof + attack.abilityMod + attack.magicMod)
	} else {
		attackMod := ""
	}
	
	if (IsObject(attack)) {
		cs := attack.cs
		abilityMod := attack.abilityMod
		ability := attack.ability
		prof := attack.prof
		magicText := (attack.magicMod) ? " + " . attack.magicMod . " [magic]" : ""
		vantage := attack.vantage
		
		attackText={{r1=[[@{%charName%|d20}cs>%cs% + %abilityMod%[%ability%] + %prof%[prof]%magicText%]]}} {{%vantage%=1}} {{r2=[[1d20cs>%cs% + %abilityMod%[%ability%] + %prof%[prof]%magicText%]]}} {{attack=1}} 
		
	} else {
		attackText={{r1=[[0d20cs>20]]}} {{r2=[[0d20cs>20]]}} 0 
	}
	
	if (IsObject(damage1)) {
		dice := damage1.dice
		crit := damage1.crit
		desc := damage1.desc
		
		damage1Text = {{damage=1}} {{dmg1flag=1}} {{dmg1=[[%dice%]]}} {{dmg1type=%desc%}} {{crit1=[[%crit%]]}}
		
	} else {
		damage1Text = 0 {{dmg1=[[0]]}} {{dmg1type=}} {{crit1=[[0[CRIT]]]}}
	}

	if (IsObject(damage2)) {
		dice := damage2.dice
		desc := damage2.desc
		crit := damage2.crit
		
		damage2Text = {{damage=1}} {{dmg2flag=1}} {{dmg2=[[%dice%]]}} {{dmg2type=%desc%}} {{crit2=[[%crit%]]}}
		
	} else {
		damage2Text = 0 {{dmg2=[[0]]}} {{dmg2type=}} {{crit2=[[0[CRIT]]]}}
	}
		
	if (IsObject(save)) {
		saveAttr := save.attr
		saveDesc := save.desc
		saveDC :=  save.dc
		
		saveText = {{save=1}} {{saveattr=%saveAttr%}} {{savedesc=%saveDesc%}} {{savedc=[[%saveDC%[SAVE]]]}}
	
	} else {
		saveText = 0
	}
	
	spellText = {{spelllevel=%spell%}} {{innate=%innate%}}

	
		 roll=@{%charName%|wtype}&{template:atkdmg} {{mod=%attackMod%}} {{rname=%weaponName%}} %attackText% {{range=%range%}} %damage1Text% %damage2Text% %saveText% %spellText% {{desc=%description%}} {{globalattack=@{%charName%|global_attack_mod}}} {{globaldamage=}} {{globaldamagecrit=}} {{globaldamagetype=}} ammo= @{%charName%|charname_output}

	return roll
}


AdvanceAttack(){
    global gAction
    if (gAction == 1) {
        GuiControl, 1: , First Attack, 1
        gAction = 2
    } else {
        GuiControl, 1: , Second Attack or Full Action, 1
        gAction = 3
    }
}

;--------------------------------------------------------
; GRAPPLES AND SHOVES
;--------------------------------------------------------

Shove:
Grapple(true)
AdvanceAttack()
WM_COMMAND()
return

Grapple:
Grapple(false)
AdvanceAttack()
WM_COMMAND()
return

Grapple(shove := false, bonus := false, result := false){
    global gLeft, gRageActive, gAction, gLeft, gRight
    ;vantage := (gAdvantage = 1) ? "advantage" : (gAdvantage = 2) ? "normal" : "disadvantage"
    vantage := (gRageActive = 1) ? "advantage" : "normal"
    
    if (shove) {
        grappleType := "shove"
    } else {
        grappleType := "grapple"
    }
    
    if (bonus) {
        BufferCommand("/me uses tavern brawling to make a " . grappleType . " attempt")
    } else if (gAction = 1) {
        BufferCommand("/me uses his first attack to attempt a " . grappleType)
    } else {
        BufferCommand("/me uses his second attack to attempt a " . grappleType)
    }
    
    roll=@{Bo|wtype}&{template:simple} {{rname=^{athletics-u}}} {{mod=@{Bo|athletics_bonus}}} {{r1=[[@{Bo|d20}+@{Bo|athletics_bonus}@{Bo|pbd_safe}]]}} {{%vantage%=1}} {{r2=[[@{Bo|d20}+@{Bo|athletics_bonus}@{Bo|pbd_safe}]]}} {{global=@{Bo|global_skill_mod}}} @{Bo|charname_output}
    SendCommand(roll)
    
    if(!result) {
        result := MsgBoxEx("Was the " . grappleType . " successful?", Format("{:T}", grappleType) . " attempt made!", "Yes|No", [1, ".\ico\grapple.ico"], "", "", WinExist("A"))
    }
    if (result == "Yes") {
        SendCommand("/me " . grappleType . " success!")
        if (gLeft <= 2 and !shove) {
            GuiControl, 1: , Grapple 1, 1
            gLeft = 3
        } else if (!shove) {
            GuiControl, 1: , Grapple 2, 1
            gRight = 3
        }
    } else if (result == "No") {
        SendCommand("/me " . grappleType . " fail!")
    }
    return
}

;--------------------------------------------------------
; DISENGAGE DASH DODGE HIDE / CUNNING ACTIONS
;--------------------------------------------------------

OtherAction:
Gui, +OwnDialogs ;Add this line just before input box to make it modal
InputBox, userInput, Other Action, Please enter the name of the action you are taking`nEg. Help or Ready, , 350, 150
if (!ErrorLevel){
    OtherAction(userInput)
}
return
Disengage:
OtherAction("Disengage")
return
Dodge:
OtherAction("Dodge")
return
Dash:
OtherAction("Dash")
return
Hide:
OtherAction("Hide")
return

CunningAction:
result := MsgBoxEx("Choose an option", "Cunning Action", "Disengage|Dash|Hide|Cancel*", 0, "", "", WinExist("A"), 0, "s9 c0x000000", "Segoe UI")
if (result != "Cancel") {
    OtherAction(result, true)
}
return

OtherAction(type, bonus := false) {
	global gBonusActionUsed, oncePerTurnAbility
    if (bonus){
        gBonusActionUsed = 1
        GuiControl, 1:, gBonusActionUsed, 1
        BufferCommand("/me uses a cunning action to " . type)
    } else {
        GuiControl, , Second Attack or Full Action, 1
        BufferCommand("/me takes the " . type . " action")
    }
	if (type = "Hide") {
		roll=@{Bo|wtype}&{template:simple} {{rname=^{stealth-u}}} {{mod=@{Bo|stealth_bonus}}} {{r1=[[@{Bo|d20}+@{Bo|stealth_bonus}@{Bo|pbd_safe}]]}} {{always=1}} {{r2=[[@{Bo|d20}+@{Bo|stealth_bonus}@{Bo|pbd_safe}]]}} {{global=@{Bo|global_skill_mod}}} @{Bo|charname_output}
		BufferCommand(roll)
	}
	if ((type = "Dash" or type = "Dodge") and GetCurrentWeapon().Weapon = "Sword of the Paruns" and !oncePerTurnAbility[GetCurrentWeapon().Weapon]) {
		oncePerTurnAbility[GetCurrentWeapon().Weapon] := true
		if (type = "Dash") {
			SendDescription("  Immediately after you take the Dash action, you can enable one creature within 60 feet of you to use its reaction to move up to its speed.", "Teamwork Dash", "Sword of the Paruns Special Ability")
		} else {
			SendDescription("  Immediately after you take the Dodge action, you can enable one creature within 60 feet of you to use its reaction to gain the benefits of the Dodge action.", "Teamwork Dodge", "Sword of the Paruns Special Ability")
		}
	} else {
		SendCommand("")
	}
    WM_COMMAND()
}

GetCurrentWeapon(){
	global a, arsenalKey, gCurrent
	return a[arsenalKey[gCurrent]]
}

;--------------------------------------------------------
; MINOR ACTIONS
;--------------------------------------------------------

Shield:
if (gLeft = 6) {
    BufferCommand("/me puts away his shield")
    GuiControl, , gLeft, 1
} else { 
    BufferCommand("/me equips his shield")
    GuiControl, , Shield, 1
}
GuiControl, , Second Attack or Full Action, 1
WM_COMMAND()
SendCommand(GetAcText())
return

StartRage:
GuiControl, , gRageActive, % !gRageActive
GuiControl, , gBonusActionUsed, 1
if (gRageActive = 0) {
    SendCommand("**WAAAGH!**`n/me enters a rage")
} else {
    SendCommand("/me calms down")
}
WM_COMMAND()
return

UncannyDodge:
GuiControl, , gReactionUsed, 1
SendCommand("/me uncannily dodges")
WM_COMMAND()
return

StowArsenal:
if (gRight = 2) {
    GuiControl, , gRight, 1
    SendCommand("/me puts Arsenal away")
} else {
    GuiControl, , Arsenal, 1
    SendCommand("/me draws Arsenal")
}
GuiControl, , gObjectInteract, 1
WM_COMMAND()
return

DropArsenal:
GuiControl, , gRight, 1
WM_COMMAND()
return

StowShortblade:
StowDrawShortblade()
SendCommand("")
return

StowDrawShortblade(){
    global
    if (gLeft = 5) {
        GuiControl, , gLeft, 1
        SendCommand("/me stows his shortblade")
    } else {
        GuiControl, , Shortblade, 1
        BufferCommand("/me draws his shortblade")
        WeaponPanel(offhand)
    }
    GuiControl, , gObjectInteract, 1
    WM_COMMAND()
    return
}

DropShortblade:
GuiControl, , gLeft, 1
WM_COMMAND()
return

RecklessAttack:
GuiControl, , gRecklessAttack, Reckless Active!
GuiControl, Disable, gRecklessAttack
GuiControl, , gAdvantage, 1
SendCommand("/me attacks recklessly!")
WM_COMMAND()
return


;-------------------------------------------------------------------------------
; WM COMMAND
; runs every time the GUI is interacted with to update what gui elements are active
;-------------------------------------------------------------------------------

WM_COMMAND() { 
    global
    /*if (A_TickCount - commandTime <= 100) ; elapsedTime since this function was last run
        Return
    commandTime := A_TickCount
    */
    Gui, Submit, NoHide
            
    GuiControl, Text, gNewInfo, % WeaponText(arsenalKey[gNew])
    
    i := arsenalKey[gCurrent] ; index of current weapon in "a" arsenal array
	
    if a[i].Versatile {
        GuiControl, Text, gVersatile, % "Versatile (" . a[i].Versatile . ")"
        GuiControl, Show, gVersatile
    } else {
        GuiControl, Hide, gVersatile
    }
    if a[i].Condition {
		condition := " " . a[i].Condition
		condition .= a[i].Damage2 ? " (" . a[i].Damage2 . " " . a[i].Type2 . ")" : ""
        GuiControl, Text, gCondition, %condition%
        GuiControl, Show, gCondition
    } else {
        GuiControl, Hide, gCondition
    }
    if (a[i].Special or a[i].Detail) {
		if a[i].Special
			GuiControl, Text, gSpecial, % a[i].Special
		else
			GuiControl, Text, gSpecial, Info
        GuiControl, Show, gSpecial
    } else {
        GuiControl, Hide, gSpecial
    }
	if ((a[i].Special or a[i].Detail) and a[i].Condition) {
		GuiControl, Move, gCondition, y47
		;MsgBox % "*" . a[i].Special . " *" . a[i].Condition
		GuiControl, 1:Move, gCurrentInfo, w330
	} else {
		GuiControl, Move, gCondition, y70
		GuiControl, Move, gCurrentInfo, w478
	}
   
    GuiControl, Text, gDescription, % a[i].Description
    GuiControl, Text, gCurrentInfo, % WeaponText(i, false)
    
; ARSENAL
    if (gArsenalTransformed or !arsenalKey[gNew]) {
        GuiBatch("|gTransformArsenal")
    } else {
        GuiBatch("gTransformArsenal")
    }
	    
    if (gLeft = 6)
        GuiControl, Text, gShield, Doff Shield
    else 
        GuiControl, Text, gShield, Don Shield
		
	
    GuiControl, Text, gStatus, % StatusText() 

; BOLD SUGGESTED ACTIONS

    if (gAction < 3 and a[i].TwoHanded and gLeft != 2)
        GuiBatch("Two-handed", "Bold")
    else
        GuiBatch("|Two-handed", "Bold")
    
    if (a[i].Sneak and !gSneakAttackMade and (gAdvantage = 1 or (gAllySupport and gAdvantage < 3))) {
        GuiBatch("gAttack", "Bold")
		GuiControl, Text, gAttack, SNEK ATTAC!
        GuiBatch("|gSneakAttackMade,gAdvantage,gAllySupport,gRecklessAttack", "cRed")
    } else {
	
		if (gAction < 3 and !gSneakAttackMade and gAdvantage != 1 and !gAllySupport and a[i].Sneak) {
			GuiBatch("gRecklessAttack", "Bold", (gAction == 1))
			
			GuiBatch("gAdvantage,gAllySupport|gSneakAttackMade", "cBlue")
		} else {
			GuiBatch("|gAdvantage,gAllySupport,gRecklessAttack", "Bold")
		
			if (!gSneakAttackMade) {
				GuiBatch("gSneakAttackMade", "cRed")
			} else {
				GuiBatch("|gSneakAttackMade", "cRed")
			}
		}
        GuiBatch("|gAttack", "Bold")
		GuiControl, Text, gAttack, ATTACK!
    }
	if (gAdvantage = 1) {
		GuiBatch("gAdvantage", "cGreen")
	}
	GuiBatch("Disadvantage", "cRed", (gAdvantage = 3))
	
	; Remind user to keep advantage turned on when attacking a sworn enemy with Oathbow
	if (gCondition and a[i].Condition = "sworn enemy") {
		GuiControl, Text, gAdvantage, Sworn Enemy
		if (gAdvantage = 1) {
			GuiBatch("gAdvantage", "cGreen Bold")
		} else {
			GuiBatch("gAdvantage", "cBlue Bold")
		}
	} else {
		GuiControl, Text, gAdvantage, Advantage
	}
        
    if (!gBonusActionUsed && !gRageActive)
        GuiBatch("gStartRage", "Bold")
    else
        GuiBatch("|gStartRage", "Bold")
		
	if (gAction > 2) {
		GuiBatch("gBonusActionUsed", "cBlue", !gBonusActionUsed)
		GuiBatch("gObjectInteract", "cBlue", !gObjectInteract)
		GuiBatch("gArsenalTransformed", "cBlue", !gArsenalTransformed)
	} else {
		GuiBatch("|gBonusActionUsed,gObjectInteract,gArsenalTransformed", "cBlue")
	}
	 
    
; BONUS ACTIONS
    if (gRageActive) {
        GuiControl, Text, gStartRage, End Rage
    } else {
        GuiControl, Text, gStartRage, Start Rage
    }
    if gBonusActionUsed {
        GuiBatch("|gStartRage,gCunningAction")
    } else {
        GuiBatch("gStartRage,gCunningAction")
    }
    /*
    1 Empty
    2 Arsenal
    3 Grapple
    4 Other
    5 Shortblade
    6 Shield
    */
    if (gAction > 1 and gBonusActionUsed = 0 and a[i].Light = 1 and (gLeft = 5 or (gLeft = 1 and gObjectInteract = 0))) {
        GuiBatch("gOffhandAttack")
    } else {
        GuiBatch("|gOffhandAttack")        
    }
; HANDS
    if (gRight != 2 or (!a[i].Twohanded and !a[i].Versatile)) {
        GuiControl, Disable, Two-handed
        if(gLeft = 2) {
            GuiControl, , gLeft, 1
        }
    } else {
        GuiControl, Enable, Two-handed
    }
    if (gLeft = 2) {
        GuiControl, , gVersatile, 1
    } else {
        GuiControl, , gVersatile, 0
    }
    if (gLeft > 2) {
        GuiControl, Disable, gVersatile
    } else {
        GuiControl, Enable, gVersatile
    }
    
    
; REACTIONS
    if gReactionUsed {
        GuiBatch("|gUncannyDodge,gOpportunityAttack")
    } else {
        GuiBatch("gUncannyDodge,gOpportunityAttack")
    }    
    
; FREE ACTIONS
    if (gRight = 2) {
        GuiControl, Text, gStowArsenal, Stow Arsenal
        GuiBatch("gDropArsenal")
    } else {
        GuiControl, Text, gStowArsenal, Draw Arsenal
        GuiBatch("|gDropArsenal")
    }
    if (gLeft = 5) {
        GuiControl, Text, gStowShortblade, Stow Shortblade
        GuiBatch("gDropShortblade")
    } else {
        GuiControl, Text, gStowShortblade, Draw Shortblade
        GuiBatch("|gDropShortblade")
    }
    if (!gObjectInteract and (gLeft = 1 or gLeft = 5)) {
        GuiBatch("gStowShortblade")
    } else {
        GuiBatch("|gStowShortblade")
    }
    if (!gObjectInteract) {
        GuiBatch("gStowArsenal")
    } else {
        GuiBatch("|gStowArsenal")
    }
; ACTIONS
    if (gRight = 2 and gAction <= 2 and a[i].Damage and (gLeft = 2 or !a[i].Twohanded)) {
        GuiBatch("gAttack")
        ;MsgBox YES
    } else {
        ;MsgBox no attack
        GuiBatch("|gAttack")        
    }
    if (gLeft = 1 or gRight = 1) {
        GuiBatch("gGrapple")
    } else {
        GuiBatch("|gGrapple")
    }
    if (gLeft = 1 or gLeft = 6) {
        GuiBatch("gShield")
    } else {
        GuiBatch("|gShield")
    }
    if (gRight = 2) {
        GuiControl, Text, gImprovisedAttack, Improvised Attack
    } else {
        GuiControl, Text, gImprovisedAttack, Claws and Teeth
    }
    if (gAction = 1) {
        GuiBatch("gImprovisedAttack,gShove,gDisengage,gDodge,gDash,gHide,gOtherAction")
    } else if (gAction = 2) {
        GuiBatch("gImprovisedAttack,gShove|gShield,gDisengage,gDodge,gDash,gHide,gOtherAction,gRecklessAttack")
    } else {
        GuiBatch("|gAttack,gGrapple,gImprovisedAttack,gShove,gShield,gDisengage,gDodge,gDash,gHide,gOtherAction,gRecklessAttack")
    }
}

;-------------------------------------------------------------------------------
; FUNCTIONS
;-------------------------------------------------------------------------------

/*w.console := NewWindow(consoleName, 3, false, "AutoHotKey.exe")
w.weapon :=  NewWindow("Arsenal Weapon Notes", 1, "https://docs.google.com/document/d/18c5vmZ7usGORrVS3i88AXlfV1MIo0DIsa0GOvBxjC00/edit#", true)
w.tool :=    NewWindow("5etools", 2, "file:///E:/Share/DnD/5eTools.1.73.5/5etools.html", true)
w.tool2 :=   NewWindow("Character Feats", 1, "http://www.mikeinside.com/characters/feats/")
w.drive :=   NewWindow("D&D - Google Drive", 1, "https://drive.google.com/drive/folders/1oZsTNdPkUpMvfhxRKqOWE2qzny3qp7iU")
w.char :=    NewWindow(charName, 1)
w.game :=    NewWindow(roll20Game, 1, "https://app.roll20.net/campaigns/search/", true)
*/
WM_ACTIVATE(wparam) {
    global w, activatePrimed, gArrange
    FindWindowIDs()
    
    charW := w.char.id ? 0 : 798 ; extra width if no character window
    WinGet, consoleMin, MinMax, % w.console.id
    consoleW := (consoleMin == -1) ? 562 : 0 ; extra width if console window is minimised
    
    ;if console window is deactivated and new window does not belong to program, set this function to active
    if (wparam = 0) { 
        WinGet, exe, ProcessName, A
        WinGetTitle, title, A
        if (exe != "AutoHotKey.exe" and !InStr(title, "Roll20" )) {
            activatePrimed := true
        }
    }
    
    if (wparam > 0 and activatePrimed and gArrange) {
        activatePrimed := false
        
        WinPos(w.console.id, 1217 + charW, 539) ;, width 566, height 844
        WinPos(w.char.id, 1772, 0, 798, 1386)
        
        toolId := w.tool.id ? w.tool.id : WinId(w.tool2)       
        
        if (w.weapon.id and toolId and !consoleW) {
            WinPos(w.weapon.id, -11, 0, 1239, 548)
            WinPos(toolId, 1212, 0, 576 + charW, 548)
            WinPos(w.game.id, -11, 540, 1239  + charW, 854)
            
        } else if (toolId and !consoleW) {
            WinPos(toolId, 1212, 0, 576  + charW, 548)
            WinPos(w.game.id, -11, 0, 1239  + charW + consoleW, 1394)
            ;WinPos(w.tool.id, -11, 0, 1771  + charW, 548)
            ;WinPos(w.game.id, -11, 0, 1239 + charW, 854)
            
        } else if (w.weapon.id and !consoleW) {
            WinPos(w.weapon.id, -11, 0, 1771 + charW, 548)
            WinPos(w.game.id, -11, 540, 1239 + charW + consoleW, 854)
            
        } else {
            WinPos(w.game.id, -11, 0, 1239 + charW + consoleW, 1394)
        }
    }
}

; rearrange windows to take advantage of extra space when console is minimised (good when out of combat)
GuiSize:
if (A_EventInfo < 2 and gArrange) {
    activatePrimed := true
    WM_ACTIVATE(1)
}
return

WinPos(id, x, y, w := false, h := false) {
	global gArrange
    if (id) {
	
        WinSet, AlwaysOnTop, On, %id%
        if (w and h) {
            WinRestore, %id%
			if (!GetKeyState("Shift"))
				WinMove, %id%, , %x%, %y%, %w%, %h%
        } else if (!GetKeyState("Shift")) {
			WinMove, %id%, , %x%, %y%
        } else {
			gArrange := 0
			GuiControl, , gArrange, 0
		}
        WinSet, AlwaysOnTop, Off, %id%
    }
}

WinID(winObj) {
    global browsers, roll20Games, roll20WindowSuffix, gMute
    
    if (winObj.match) {
        SetTitleMatchMode, % winObj.match
    }
    if (winObj.exe == false or winObj.exe == true) {  ; check browsers for window name
        Loop, Parse, browsers, `,
        {
            browser := A_LoopField
            if (winObj.title == "") { ; check list of Roll20 games
                Loop, Parse, roll20Games, `,
                {
                    test := A_LoopField . roll20WindowSuffix
                    WinGet, winId, ID, % test . " ahk_exe " . browser
                    if (winId) {
                        winObj.title := test
                        Break
                    }
                }
                ;gMute := winId ? 0 : 1 ;Mute output if no game found - not strictly necessary as console will not send commands if it cannot find game window
                
            } else {
                WinGet, winId, ID, % winObj.title . " ahk_exe " . browser
            }
        
            if (winId) {
                Break
            }
        }
    } else {
        WinGet, winId, ID, % winObj.title . " ahk_exe " . winObj.exe
    }
    SetTitleMatchMode, 3
    winObj.id := winId ? "ahk_id " . winId : false
    return winObj.id
}

GuiBatch(str, command := "Enable", active := true) {
    ; Initialize counter to keep track of our position in the string.
    position := 0
    
    Loop, Parse, str, `,|
    {
        if (command = "Enable") {
            type := active ? "Enable" : "Disable" 
            if StrLen(A_LoopField) > 0 
                GuiControl, 1:%type%, %A_LoopField%
        } else {
            if active
                Gui, Font, %command%  
            GuiControl, Font, %A_LoopField%
            Gui, Font,
        }
        ; Calculate the position of the delimiter at the end of this field.
        position += StrLen(A_LoopField) + 1
        ; Retrieve the delimiter found by the parsing loop.
        delimiter := SubStr(str, position, 1)
        if (delimiter = "|")
            active := !active 
    }

}

GetHP(conMod := 0) {
    global mods
    constitution := mods["con"]
    hpBarb = 7
    hpRogue = 5
    hpMystic = 6
    lvlBarb = 6
    lvlRogue = 5
    lvlMystic = 0
    hitpoints := (hpBarb - 2) + (hpBarb * lvlBarb) + (hpRogue * lvlRogue) + (hpMystic * lvlMystic) + ((constitution + conMod) * (lvlBarb + lvlRogue + lvlMystic))
    return hitpoints
}
GetAC()
{
    global baseAC, gLeft, a, arsenalKey, gCurrent
    AC := a[arsenalKey[gCurrent]].AC 
    return baseAC + (gLeft = 6 ? 2 : 0) + (AC ? AC : 0 )
}
GetAcText(){
    return "/me AC is now " . GetAC()
}
StatusText(verbose := true) {
	global gRageActive, baseAC
	a := GetCurrentWeapon()
	
	resist := ""
	if (gRageActive) {
		if (InStr(a.Resist, "Psychic")) {
			resist .= "Resistance: All damage   "
		} else {
			resist .= "Rage active   "
		}
	} else if (a.Resist) {
		resist .= "Resistance: " . a.Resist . "   "
	}
	
	ac := GetAC()
	if (ac != baseAC or verbose) {
		ac := "AC: " . ac . "   "
	} else {
		ac := ""
	}
	
	extra := ""
	if (a.Extra and (verbose or (!verbose and !a.Extra2))) {
		extra := a.Extra . "   "
	} else if (!verbose and a.Extra2) {
		extra := (a.Extra2 = "-") ? "" : a.Extra2 . "   "
	}
	
	
	return ac
	. ((verbose) ? "MaxHP: " . GetHP(a.Weapon = "Axe of the Dwarvish Lords" ? 1 : 0)  . "   " : "")
	. resist 
	. extra
	. ((a.TotalCharge) ? "Charges: " . GetChargeWeapon(a.Weapon).Charge . "/" . a.TotalCharge  . "   " : "")
}
TrimEnd(str, charCount := -1){
    if str
        str := SubStr(str, 1, charCount)
    return str
}



NewWindow(title := "", match := 2, url := false, exe := false, id := false){
    window := {}
    window.title := title
    window.exe := exe
    window.match := match
    window.url := url
    window.id := id
    return window
}
OpenWindows(){
    global w, browserPath
    for k in w {
        if (w[k].exe == true and w[k].url) {
            if (!WinID(w[k])) {
                url := w[k].url
                Run, %browserPath% %url%
            }
        }
        
    }
}
FindWindowIDs(){
    global w
    for k in w {
        WinID(w[k])
    }
}


BufferCommand(str) {
    global commandBuffer .= str . "`n"
}
SendCommand(str){
    global w, gMute, commandBuffer
    str := commandBuffer . str
    commandBuffer := ""
    if (!gMute and w.game.id and !GetKeyState("Shift")) {
        WinSet, AlwaysOnTop, On, % w.console.id
        WinActivate, % w.game.id
        
        CoordMode, Mouse, Client
        WinGetPos, x, y, width, height, % w.game.id
        MouseGetPos, mouseOldX, mouseOldY
        mouseNewX := width - 45
        mouseNewY := height - 110
        Click, %mouseNewX%, %mouseNewY%  
        Click, %mouseOldX%, %mouseOldY%, 0
        
        clipSaved := clipboardAll   ; Save the entire clipboard.
        clipboard := str
        
        Send ^v
        Send {Enter}
        Sleep, 100
        
        clipboard := clipSaved   ; Restore the original clipboard. Note the use of Clipboard (not ClipboardAll).
        clipSaved =   ; Free the memory in case the clipboard was very large.
        
        WinSet, AlwaysOnTop, Off, % w.console.id
        WinActivate, % w.console.id
    }
}