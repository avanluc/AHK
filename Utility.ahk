#NoEnv
#MaxHotkeysPerInterval 99000000
#HotkeyInterval 99000000
#KeyHistory 0
ListLines Off
SetKeyDelay, -1, -1
SetMouseDelay, -1
SetDefaultMouseSpeed, 0
SetWinDelay, -1
SetControlDelay, -1
SendMode Input
SetWorkingDir %A_Temp%

;Global variables
RIL_DIR_SV := "D:\Rilasci\Sviluppo"
RIL_DIR_US := "D:\Rilasci\Utente"
RIL_BACKUP := "D:\Temp\Backup ultimo rilascio"
CURR_DEV   := 1


#include %A_ScriptDir%\lib\VA.ahk


VA_SetDefaultEndpoint("playback:" CURR_DEV, 0)
VA_SetDefaultEndpoint("playback:" CURR_DEV, 1)


;-----------------------------------------------------------------------------
; Switch del canale audio
#h::
	CURR_DEV := Mod(CURR_DEV,2) + 1
	VA_SetDefaultEndpoint("playback:" CURR_DEV, 0)
	VA_SetDefaultEndpoint("playback:" CURR_DEV, 1)
	return


;-----------------------------------------------------------------------------
; Svuotamento cartelle locali dei rilasci
!e::
	If !IsEmpty(RIL_DIR_US)
	{
		Cliente = ""
		NomeFile = ""
		Elenco := Object()
		;loop tra i file della cartella utente
		Loop, Files, %RIL_DIR_US%\*.*
		{
			;gestione rilasci zpf
			If (A_LoopFileExt = "ZPF")
			{
				delPos = 0
				StringGetPos, delPos, A_LoopFileName, _
				Cliente := SubStr(A_LoopFileName, 1, delPos)
				IfNotExist, %RIL_BACKUP%\%Cliente%\
					FileCreateDir, %RIL_BACKUP%\%Cliente%\
				FileMove, %RIL_DIR_US%\%A_LoopFileName%, %RIL_BACKUP%\%Cliente%\%A_LoopFileName%, 1
				Elenco.Insert(A_LoopFileName)
				Sleep, 200
			}
			
			;gestione rilasci zip
			If (A_LoopFileExt = "zip")
			{
				Loop, %RIL_DIR_SV%\*.INF, 1, 1
				{
					delPos = 0
					StringGetPos, delPos, A_LoopFileName, _
					Cliente := SubStr(A_LoopFileName, 1, delPos)
					StringGetPos, delPos, A_LoopFileName, .inf
					NomeFile := SubStr(A_LoopFileName, 1, delPos) ".zip"
					Elenco.Insert(NomeFile)
				}
				FileMove, %RIL_DIR_US%\%A_LoopFileName%, %RIL_BACKUP%\%Cliente%\%NomeFile%, 1
				Sleep, 200
			}
		}
		
		;Se la cartella è vuota il backup è andato a buon fine
		If IsEmpty(RIL_DIR_US)
		{
			FileRemoveDir, %RIL_DIR_SV%, 1
			FileCreateDir, %RIL_DIR_SV%
			str := ""
			for index, element in Elenco
				str := str "`n" element
			TrayTip, Backup eseguito, %str%, 1.5, 1
		}
		Else
		{
			TrayTip, Backup non eseguito, Errore nella gestione dei rilasci., 1.5, 3
			Return
		}
	}
	Else If !IsEmpty(RIL_DIR_SV)
	{
		FileRemoveDir, %RIL_DIR_SV%, 1
		FileCreateDir, %RIL_DIR_SV%
		TrayTip, No file to backup, Cartella sviluppo svuotata., 1.5, 1
	}
	Else
		TrayTip, Cartelle vuote, Cartelle di rilascio già vuote., 1.5, 1
		
	;inserisci nella clipboard il path per la cartella user
	clipboard = %RIL_DIR_US%\
	Return

	
IsEmpty(Dir){
	Loop %Dir%\*.*, 0, 1
		return 0
	return 1
}


;-----------------------------------------------------------------------------
; Utility menu
^+u::
	Width := 300
	Height := 400
	Gui, Destroy
	Gui, +AlwaysOnTop -MinimizeBox -MaximizeBox
	Gui, Margin, 20,20
	Gui, Font, w600 s11, Arial
	Gui, Add, Text,, Utility Menù:
	Gui, Font, w400 s10, Arial
	Gui, Add, Button, % "w150 x"(Width/2)-75 " gJVMRunPar", JVM Run Params
	Gui, Add, Button, % "w150 x"(Width/2)-75 " gJVMDebugPar", JVM Debug Params
	Gui, Add, Button, % "w150 x"(Width/2)-75 " gRilasciPath", Path cartella rilasci
	Gui, Font, w600
	Gui, Add, Text, x20, Path ambienti master:
	Gui, Font, w400
	Gui, Add, Button, % "w40 x"(Width/2)-100 " gMasterPath", AHR
	Gui, Add, Button, % "w40 x+40 gMasterPath", AHE
	Gui, Add, Button, % "w40 x+40 gMasterPath", IP
	Gui, Font, w600
	Gui, Add, Text, x20, Path ambienti work:
	Gui, Font, w400
	Gui, Add, Button, % "w40 x"(Width/2)-100 " gWorkPath", AHR
	Gui, Add, Button, % "w40 x+40 gWorkPath", AHE
	Gui, Add, Button, % "w40 x+40 gWorkPath", IP
	GuiControl,, JVMRunPar
	GuiControl,, JVMDebugPar
	GuiControl,, RilasciPath
	GuiControl,, MasterPath
	GuiControl,, WorkPath
	Gui, Show, % "w" Width " h" Height " xCenter yCenter"
	Return
	
JVMRunPar:
	clipboard = -Xdebug -Xms256m -Xmx1024m -XX:MaxPermSize=512m -Dfile.encoding=UTF-8
	Return
	
JVMDebugPar:
	clipboard = -Xlint:-unchecked -source 1.5 -target 1.5
	Return
	
RilasciPath:
	clipboard = %RIL_DIR2%
	Return
	
MasterPath:
	GuiControlGet, var,, % A_GuiControl
	clipboard = D:\Ambienti\Master\%var%\
	Return
	
WorkPath:
	GuiControlGet, var,, % A_GuiControl
	clipboard = D:\Ambienti\Work\%var%\
	Return

GuiMove:
   PostMessage, 0xA1, 2
	return

Close:
	Gui, Destroy
	return

GuiEscape:
	Gui, Destroy
	return


;-----------------------------------------------------------------------------
^+!c::
	Width := 150
	Gui, Destroy
	Gui, +AlwaysOnTop
	;WinSet, Transparent, 250
	Gui, Color, 808080
	Gui, Margin, 0, 0
	Gui, Font, s11 cDCDCDC Bold
	Gui, Add, Progress, % "x-1 y-1 w" (Width+2) " h31 Background404040 Disabled hwndHPROG"
	Control, ExStyle, -0x20000, , ahk_id %HPROG% ; propably only needed on Win XP
	Gui, Add, Text, % "x0 y0 w" Width " h30 BackgroundTrans Center 0x200 gGuiMove vCaption", CodePainter:
	Gui, Font, s10
	Gui, Add, Text, % "x60 y+10 w" (Width-14) "r1 +0x4000 vTX1 gLaunchCodePainter", 55
	Gui, Add, Text, % "x60 y+10 w" (Width-14) "r1 +0x4000 vTX3 gLaunchCodePainter", 56
	Gui, Add, Text, % "x60 y+10 w" (Width-14) "r1 +0x4000 vTX2 gLaunchCodePainter", 59
	Gui, Add, Text, % "x60 y+10 w" (Width-14) "r1 +0x4000 vTX4 gLaunchCodePainter", 60
	;Gui, Add, Text, % "x50 y+10 w" (Width-14) "r1 +0x4000 vTX5 gClose", Close
	Gui, Add, Text, % "x60 y+10 w" (Width-14) "h5 vP"
	GuiControlGet, P, Pos
	H := PY + PH
	Gui, -Caption
	WinSet, Region, 0-0 w%Width% h%H% r6-6
	Gui, Show, % "w" Width " NA"
	WinSet AlwaysOnTop
	GuiControl, Focus, 55
	return

LaunchCodePainter:
	GuiControlGet, var,, % A_GuiControl
	run, C:\CP_local\CpwR%var%\cpl_host.exe cprfrontend
	Gui, Destroy
	Return

;-----------------------------------------------------------------------------
^+v::
	Width := 250
	Gui, Destroy
	Gui, +AlwaysOnTop
	;WinSet, Transparent, 250
	Gui, Color, 808080
	Gui, Margin, 0, 0
	Gui, Font, s11 cDCDCDC Bold
	Gui, Add, Progress, % "x-1 y-1 w" (Width+2) " h31 Background404040 Disabled hwndHPROG"
	Control, ExStyle, -0x20000, , ahk_id %HPROG% ; propably only needed on Win XP
	Gui, Add, Text, % "x0 y0 w" Width " h30 BackgroundTrans Center 0x200 gGuiMove vCaption", Lancio validazione:
	Gui, Font, s10
	Gui, Add, Text, % "x90 y+10 w" (Width-14) "r1 +0x4000 vTX1 gValidateCodePainter", CodePainter
	Gui, Add, Text, % "x90 y+10 w" (Width-14) "r1 +0x4000 vTX2 gValidateSitePainter", SitePainter
	;Gui, Add, Text, % "x105 y+10 w" (Width-14) "r1 +0x4000 vTX3 gClose", Close
	Gui, Add, Text, % "x95 y+10 w" (Width-14) "h5 vP"
	GuiControlGet, P, Pos
	H := PY + PH
	Gui, -Caption
	WinSet, Region, 0-0 w%Width% h%H% r6-6
	Gui, Show, % "w" Width " NA"
	WinSet AlwaysOnTop
	GuiControl, Focus, SitePainter
	return

ValidateSitePainter:
	Gui, Destroy
	sleep, 100
	run "D:\Scripts\RemoteValidate.bat" srv-ip-build NBAVANZINISPI
	Return

ValidateCodePainter:
	Gui, Destroy
	sleep, 100
	run "D:\Scripts\RemoteValidate.bat" pc-muletto NBAVANZINICPR
	Return



;-----------------------------------------------------------------------------
; Zoom on Foxit reader like Adobe
SetTitleMatchMode RegEx
#IfWinActive, ^(Foxit Reader \d.\d|.*? - Foxit Reader \d.\d - \[.*?\])$
{
	^WheelUp::Send, ^{NumpadAdd}
	^WheelDown::Send, ^{NumpadSub}
}
