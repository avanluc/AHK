/*
File structure & Info
Settings
	-ahk script setting
	-other settings
	-app method settings
Hotkeys for changing settings
	-WinKey + WheelUp
	-WinKey + WheelDown
	-WinKey + WheelRight
	-WinKey + WheelLeft
Main HotKeys
	-WheelUp
	-WheelDown
	-WheelRight
	-WheelLeft
Subs used by the main hotkeys
	-Set_vMethod
	-Set_hMethod
	-Get_SB_ID
Subs use by the setting hotkeys
	-Get_TApp_vMethod
	-Get_TApp_hMethod
	-Set_TApp_vMethod
	-Set_TApp_hMethod
	-SAVEtoFILE
*/

#SingleInstance force
#NoEnv
#InstallMouseHook
Critical
SendMode Input
CoordMode, Mouse, Screen
CoordMode, ToolTip, Screen

SAVEdelay := -3000		;set the delay to save setting to file
methods := 6			;set to method count except the OFF "method"; used in loops by settings hotkeys

/*
the section tag bellow allows iniWrite to opperate on this part of the file; variables here are lists of apps for each method, separated for vertical and horizontal
[AppMethodSettings]
*/
;Method1 - the default method - uses PostMessage to send WM_MOUSEWHEEL messages to the target control
;vMethod1_AppList =		;not needed since method1 is the default method
;hMethod1_AppList =		;not needed since method1 is the default method

; a list of control classes that are to use the default method regardless of the application they belong to; "Overwrite to Default Control Class List"
OtD_ConCList = SysTreeView SysListView

;Method2 - similar to method1, but sends messages to the window under the mouse rather than directly to the control... works for firefox
vMethod2_AppList = firefox.exe YahooMessenger.exe
hMethod2_AppList = firefox.exe

;Method3 - makes use of ControlClick to send WHEEL* clicks to the target control
vMethod3_AppList =            
hMethod3_AppList =              

;Method4 - sends WM_vSCROLL/WM_HSCROLL messages
vMethod4_AppList =             
hMethod4_AppList =              

;Method5 - chooses a likely ScrollBar controll and clicks it's buttons
vMethod5_AppList =               
hMethod5_AppList = AcroRd32.exe NitroPDFReader.exe Explorer.EXE

;Method6 - same as method 5, but looks for "NetUIHWND" instead of "ScrollBar" - works for office 2007 apps, with some side effects depending on mouse position, because NetUIHWND elements are not meant to be only scrollbars
vMethod6_AppList =             
hMethod6_AppList = WINWORD.EXE EXCEL.EXE POWERPNT.EXE

;OFF - sends a WHEEL* via the Click command to the active window - should simulate normal behaviour
vMethodOff_AppList =                          
hMethodOff_AppList =
/*
[Rest of Script]
this tag is just here to mark the end of the above section
*/

;Hotkeys for changing the settings above
#WheelUp::
gosub, Get_TApp_vMethod			; a sub to get the current setings for the app from the settings section above
if TApp_vMethod_ex = Off		; decides the change in setting
	TApp_vMethod_new := 1
else if TApp_vMethod_ex = %methods%
	TApp_vMethod_new := "Off"
else TApp_vMethod_new := TApp_vMethod_ex + 1
gosub, Set_TApp_vMethod			; a sub to apply the change decided
return

#WheelDown::
gosub, Get_TApp_vMethod
if TApp_vMethod_ex = Off
	TApp_vMethod_new := methods
else if TApp_vMethod_ex = 1
	TApp_vMethod_new := "Off"
else TApp_vMethod_new := TApp_vMethod_ex - 1
gosub, Set_TApp_vMethod
return

#WheelRight::
gosub, Get_TApp_hMethod
if TApp_hMethod_ex = Off
	TApp_hMethod_new := 1
else if TApp_hMethod_ex = %methods%
	TApp_hMethod_new := "Off"
else TApp_hMethod_new := TApp_hMethod_ex + 1
gosub, Set_TApp_hMethod
return

#WheelLeft::
gosub, Get_TApp_hMethod
if TApp_hMethod_ex = Off
	TApp_hMethod_new := methods
else if TApp_hMethod_ex = 1
	TApp_hMethod_new := "Off"
else TApp_hMethod_new := TApp_hMethod_ex - 1
gosub, Set_TApp_hMethod
return

;Main Hotkeys ;Commands issued here will run each time the hotkey is called
$WheelUp::
MouseGetPos, mX, mY, TWin, TCon, 2		; get the mouse position and the IDs of the window and the control under it; 2 at the end --> TCon is a HWND
if (TWin <> %exWin4v%) or (TCon <> %exCon4v%)	; check if what's under the mouse has changed since this hotkey was last called to minimize redundant opperations
	gosub, Set_vMethod			; if it has, run this sub to decide what method to use, else use the same method as before
if (vMethod = 0)			;apply the chosen method (see setting section for comments)
	Click, WheelUp
else if (vMethod = 1)
	PostMessage, 0x20A, 120 << 16, (mY << 16) | mX,, ahk_id%TCon%
else if (vMethod = 2)
	PostMessage, 0x20A, 120 << 16, (mY << 16) | mX,, ahk_id%TWin%
else if (vMethod = 3)
	ControlClick,, ahk_id%TCon%,, WheelUp, 1, NA
else if (vMethod = 4)
	PostMessage, 0x115, 0, 0,, ahk_id%TCon%
else if (vMethod = 5) or (hMethod = 6)
	{
	PostMessage, 0x201,, (5 << 16) | 5,, ahk_id%VBarID%
	PostMessage, 0x202,, (5 << 16) | 5,, ahk_id%VBarID%
	;ToolTip, click, (TWinX + VBarX + 5), (TWinY + VBarY + 5)		; enable in order to see where the clicks are being sent in methods 5 and 6; for debuging purposes
	}
return

$WheelDown::
MouseGetPos, mX, mY, TWin, TCon, 2
if (TWin <> %exWin4v%) or (TCon <> %exCon4v%)
	gosub, Set_vMethod
if (vMethod = 0)
	Click, WheelDown
else if (vMethod = 1)
	PostMessage, 0x20A, -120 << 16, (mY << 16) | mX,, ahk_id%TCon%
else if (vMethod = 2)
	PostMessage, 0x20A, -120 << 16, (mY << 16) | mX,, ahk_id%TWin%
else if (vMethod = 3)
	ControlClick,, ahk_id%TCon%,, WheelDown, 1, NA
else if (vMethod = 4)
	PostMessage, 0x115, 1, 0,, ahk_id%TCon%
else if (vMethod = 5) or (hMethod = 6)
	{
	PostMessage, 0x201,, ((VBarH - 5) << 16) | 5,, ahk_id%VBarID%
	PostMessage, 0x202,, ((VBarH - 5) << 16) | 5,, ahk_id%VBarID%
	;ToolTip, click, (TWinX + VBarX + 5), (TWinY + VBarY + VBarH - 5)
	}
return

$WheelLeft::
MouseGetPos, mX, mY, TWin, TCon, 2
if (TWin <> %exWin4h%) or (TCon <> %exCon4h%)
	gosub, Set_hMethod
if (hMethod = 0)
	Click, WheelLeft
else if (hMethod = 1)
	PostMessage, 0x20E, -120 << 16, (mY << 16) | mX,, ahk_id%TCon%
else if (hMethod = 2)
	PostMessage, 0x20E, -120 << 16, (mY << 16) | mX,, ahk_id%TWin%
else if (hMethod = 3)
	ControlClick,, ahk_id%TCon%,, WheelLeft, 1, NA
else if (hMethod = 4)
	PostMessage, 0x114, 0, 0,, ahk_id%TCon%
else if (hMethod = 5)  or (hMethod = 6)
	{
	PostMessage, 0x201,, (5 << 16) | 5,, ahk_id%HBarID%
	PostMessage, 0x202,, (5 << 16) | 5,, ahk_id%HBarID%
	;ToolTip, click, (TWinX + HBarX + 5), (TWinY + HBarY + 5)
	}
return

$WheelRight::
MouseGetPos, mX, mY, TWin, TCon, 2
if (TWin <> %exWin4h%) or (TCon <> %exCon4h%)
	gosub, Set_hMethod
if (hMethod = 0)
	Click, WheelRight
else if (hMethod = 1)
	PostMessage, 0x20E, 120 << 16, (mY << 16) | mX,, ahk_id%TCon%
else if (hMethod = 2)
	PostMessage, 0x20E, 120 << 16, (mY << 16) | mX,, ahk_id%TWin%
else if (hMethod = 3)
	ControlClick,, ahk_id%TCon%,, WheelRight, 1, NA
else if (hMethod = 4)
	PostMessage, 0x114, 1, 0,, ahk_id%TCon%
else if (hMethod = 5) or (hMethod = 6)
	{
	PostMessage, 0x201,, (5 << 16) | (HBarW - 5),, ahk_id%HBarID%
	PostMessage, 0x202,, (5 << 16) | (HBarW - 5),, ahk_id%HBarID%
	;ToolTip, click, (TWinX + HBarX + HBarW - 5), (TWinY + HBarY + 5)
	}
return

;Subs to choose the method to be used by the main hotkeys ;commands issued here will run only if the control or window under the mouse have changed
Set_vMethod:				;sub to set the vertical method
exWin4v := TWin, exCon4v := TCon	;save the curent TWin and TCon for reference the next time Hotkeys are called, separated for v and h; also reset Control Class Override to false
MouseGetPos,,,, TConC			;get the ClassNN ofthe control under the mouse ;I haven't found a way to get it directly from its HWND, already saved as TCon
if (TConC <> "")		; skip this loop if no ClassNN was found above
	{
	loop, parse, OtD_ConCList, %A_Space%		;parse the Override to Default Control Class List
		{
		if inStr(TConC, A_LoopField)		;check if TCon belongs to any of the classes listed in OtD_ConCList
			{
			vMethod := 1			;set method to default
			return				;and return to caller
			}
		}
	}
WinGet, TApp, ProcessName, ahk_id%TWin%		;if the sub didn't retun above, get the name of the application to which TWin belongs
if inStr(vMethod2_AppList, TApp)		;check if the app targeted appears on any setting list and run additional commands required by the method
	vMethod := 2
else if inStr(vMethod3_AppList, TApp)
	vMethod := 3
else if inStr(vMethod4_AppList, TApp)
	vMethod := 4
else if inStr(vMethod5_AppList, TApp)
	{
	vMethod := 5, VorH := "V", Look4 := "ScrollBar"
	gosub, Get_SB_ID		; a sub to find the most likely control that acts as a scrollbar by name(Look4) and direction(VorH)
	}
else if inStr(vMethod6_AppList, TApp)
	{
	vMethod := 6, VorH := "V", Look4 := "NetUIHWND"
	gosub, Get_SB_ID
	}
else if inStr(vMethodOff_AppList, TApp)
	vMethod := 0
else 
	vMethod := 1
return


Set_hMethod:			;sub to set the horizontal method
exWin4h := TWin, exCon4h := TCon
MouseGetPos,,,, TConC
if (TConC <> "")
	{
	loop, parse, OtD_ConCList, %A_Space%
		{
		if inStr(TConC, A_LoopField)
			{
			hMethod := 1
			return
			}
		}
	}
WinGet, TApp, ProcessName, ahk_id%TWin%
if inStr(hMethod2_AppList, TApp)
	hMethod := 2
else if inStr(hMethod3_AppList, TApp)
	hMethod := 3
else if inStr(hMethod4_AppList, TApp)
	hMethod := 4
else if inStr(hMethod5_AppList, TApp)
	{
	hMethod := 5, VorH := "H", Look4 := "ScrollBar"
	gosub, Get_SB_ID
	}
else if inStr(hMethod6_AppList, TApp)
	{
	hMethod := 6, VorH := "H", Look4 := "NetUIHWND"
	gosub, Get_SB_ID
	}
else if inStr(vMethodOff_AppList, TApp)
	hMethod := 0
else 
	hMethod := 1
return


Get_SB_ID:			; a sub that looks for scroll bars, either vertical or horizontal
%VorH%BarID := ""		; clears any previously stored BarID for the desired direction
if inStr(TConC, "ScrollBar")						; check if the control targeted isn't a scrollbar itself
	{
	ControlGetPos, SBarX, SBarY, SBarW, SBarH,, ahk_id%TCon%	;if it is, get its pos and dimentions
	if (VorH = "V") and (SBarH > SBarW)				;and check if it's of the desired type, V or H
		{
		ControlGet, VBarID, Hwnd,,, ahk_id%TCon%					;if yes, stores its ID and info as the desired scroll bar
		VBarC := TConC, VBarX := SBarX, VBarY := SBarY, VBarW := SBarW, VBarH := SBarH
		return										;and returns
		}
	else if (VorH = "H") and (SBarW > SBarH)
		{
		ControlGet, HBarID, Hwnd,,, ahk_id%TCon%
		HBarC := TConC, HBarX := SBarX, HBarY := SBarY, HBarW := SBarW, HBarH := SBarH
		return
		}
	else TSBarH := SBarH, TSBarW := SBarW			;if the control is a scrollbar, but not of the desired type, asume it's the other scrollbar for the target surface and store it's dimentions to expand the search area
	}
else  TSBarH := 0, TSBarW := 0				;if the controll under the mouse is not a scrollbar, clear theese vars
WinGet, TWinConList, ControlList, ahk_id%TWin%		;if the sub didn't return above, get a list of all the controlls in TWin
WinGetPos, TWinX, TWinY,,, ahk_id%TWin%			;get the position of the target window (Twin)
loop, Parse, TWinConList, `n				;parse said list
	{
	if inStr(A_LoopField, Look4)			;and look for the name stored in Look4, either "ScrollBar"(method5) or "NetUIHWND"(method)6
		{
		ControlGet, vis, Visible,, %A_LoopField%, ahk_id%TWin%		;find out if that control is visible
		if vis = 1
			{
			ControlGetPos, SBarX, SBarY, SBarW, SBarH,  %A_LoopField%, ahk_id%TWin%							;if it's visible, get the position and size of the found scrollbar, relative to the window
			if (VorH = "V") and (SBarH > SBarW) and (mX < (TWinX + SBarX + SBarW)) and (mY < (TWinY + SBarY + SBarH + TSBarH))	;check its type and position relative to the mouse
				{
				if (VBarID = "") or (SBarX < VBarX) or ((TWinY + VBarY) > (TWinY + SBarY + SBarH))				;if it cleared, compare it to any previously stored scrollbar to see if it's better suited
					{
					ControlGet, VBarID, Hwnd,, %A_LoopField%, ahk_id%TWin%							;if it's the only candidate, or better than the previous, save it as the desired scrollbar
					VBarC := A_LoopField, VBarX := SBarX, VBarY := SBarY, VBarW := SBarW, VBarH := SBarH
					}
				}
			else if (VorH = "H") and (SBarW > SBarH) and (mY < (TWinY + SBarY + SBarH)) and (mX < (TWinX + SBarX + SBarW + TSBarW))
				{
				if (HBarID = "") or (SBarY < HBarY) or ((TWinX + HBarX) > (TWinX + SBarX + SBarW))
					{
					ControlGet, HBarID, Hwnd,, %A_LoopField%, ahk_id%TWin%
					HBarC := A_LoopField, HBarX := SBarX, HBarY := SBarY, HBarW := SBarW, HBarH := SBarH
					}
				}
			}

		}
	}
return


;Subs used by settings hotkeys
Get_TApp_vMethod:				;a sub to get the current vMethod setting for the target app
MouseGetPos,,, TWin				;get the ID of the window under the mouse
WinGet, TApp, ProcessName, ahk_id%TWin%		;and find what application it belongs to
TApp_vMethod_ex := ""			;reset this; this is what the sub will find
If inStr(vMethodOff_AppList, TApp)	;check if the app is on the Off list
	{
	TApp_vMethod_ex := "Off"						;if it is, record it
	StringReplace, vMethodOff_AppList, vMethodOff_AppList, %TApp%,, All	;and remove it from the list
	}
Loop				;a loop to check the othe lists
	{
	if (A_Index <> 1) and inStr(vMethod%A_Index%_AppList, TApp)	; same as above for the other lists, except for 1 - the default, which has no list
		{
		TApp_vMethod_ex := A_Index
		StringReplace, vMethod%A_Index%_AppList, vMethod%A_Index%_AppList, %TApp%,, All
		}
	if A_Index = %methods%			;when the loop reaches this, it has checked all lists
		{
		if not TApp_vMethod_ex		;if the app wasn't on any list, set to default
			TApp_vMethod_ex := 1
		if TApp_vMethod_ex = 1			;show the found method in a tooltip (probably for too short a time to be seen
			ToolTip, default
		else, ToolTip, %TApp_vMethod_ex%
		break
		}
	}
return

Get_TApp_hMethod:		;a sub to get the current hMethod setting for the target app(see above for details)
MouseGetPos,,, TWin
WinGet, TApp, ProcessName, ahk_id%TWin%
TApp_hMethod_ex := ""
If inStr(hMethodOff_AppList, TApp)
	{
	TApp_hMethod_ex := "Off"
	StringReplace, hMethodOff_AppList, hMethodOff_AppList, %TApp%,, All
	}
Loop
	{
	if (A_Index <> 1) and inStr(hMethod%A_Index%_AppList, TApp)
		{
		TApp_hMethod_ex := A_Index
		StringReplace, hMethod%A_Index%_AppList, hMethod%A_Index%_AppList, %TApp%,, All
		}
	if A_Index = %methods%
		{
		if not TApp_hMethod_ex
			TApp_hMethod_ex := 1
		if TApp_hMethod_ex = 1
			ToolTip, default
		else, ToolTip, %TApp_hMethod_ex%
		break
		}
	}
return

Set_TApp_vMethod:		;a sub to set the new vMethod setting for the target app
if TApp_vMethod_new = "Off"						;add the app name to the apropiate list
	vMethodOff_AppList := vMethodOff_AppList . A_Space . TApp
else if TApp_vMethod_new <> 1
	vMethod%TApp_vMethod_new%_AppList := vMethod%TApp_vMethod_new%_AppList . A_Space . TApp
if TApp_vMethod_new = 1			;refresh the tooltip with the new method
	ToolTip, default
else ToolTip, %TApp_vMethod_new%
exWin4v := "",exCon4v := ""		;clear theese so the main hotkey perceive the change
SetTimer, SAVEtoFILE, %SAVEdelay%	;set a timer to save the changes in the script file and to clear the tooltip
return

Set_TApp_hMethod:		;a sub to set the new hMethod setting for the target app
if TApp_hMethod_new = "Off"
	hMethodOff_AppList := hMethodOff_AppList . A_Space . TApp
else if TApp_hMethod_new <> 1
	hMethod%TApp_hMethod_new%_AppList := hMethod%TApp_hMethod_new%_AppList . A_Space . TApp
if TApp_hMethod_new = 1
	ToolTip, default
else ToolTip, %TApp_hMethod_new%
exWin4h := "",exCon4h := ""
SetTimer, SAVEtoFILE, %SAVEdelay%
return

SAVEtoFILE:		;a sub to write settings to the file and clear the tooltips
Loop %methods%
	{
	if (A_Index <> 1)		;except 1 because it's the default method and needs no list; Off List is executed bellow the loop
		{			;see the commands for the Off list bellow for details
		vMethod%A_Index%_AppList := RegExReplace(vMethod%A_Index%_AppList, "^\s+|\s+$")
		hMethod%A_Index%_AppList := RegExReplace(hMethod%A_Index%_AppList, "^\s+|\s+$")
		StringReplace, vMethod%TApp_vMethod_new%_AppList, vMethod%TApp_vMethod_new%_AppList, %A_Space%%A_Space%,%A_Space%, All
		StringReplace, hMethod%TApp_hMethod_new%_AppList, hMethod%TApp_hMethod_new%_AppList, %A_Space%%A_Space%,%A_Space%, All
		IniWrite, % A_Space . vMethod%A_Index%_AppList, %A_ScriptName%, AppMethodSettings, vMethod%A_Index%_AppList
		IniWrite, % A_Space . hMethod%A_Index%_AppList, %A_ScriptName%, AppMethodSettings, hMethod%A_Index%_AppList
		}
	}
vMethodOff_AppList := RegExReplace(vMethodOff_AppList, "^\s+|\s+$")					;clear any blanks at the begginig or end of the list
vMethodOff_AppList := RegExReplace(vMethodOff_AppList, "^\s+|\s+$")
StringReplace, vMethodOff_AppList, vMethodOff_AppList, %A_Space%%A_Space%,%A_Space%, All		;clear doublespaces inside the list
StringReplace, vMethodOff_AppList, vMethodOff_AppList, %A_Space%%A_Space%,%A_Space%, All
IniWrite, % A_Space . vMethodOff_AppList, %A_ScriptName%, AppMethodSettings, vMethodOff_AppList		;write the list to file as ini keys in the [AppMethodSettings] section at the top of the script
IniWrite, % A_Space . vMethodOff_AppList, %A_ScriptName%, AppMethodSettings, vMethodOff_AppList
ToolTip			;clear the tooltip
return