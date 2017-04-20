#NoEnv
SendMode Input 
SetWorkingDir %A_ScriptDir% 
#include lib/Acc.ahk


Excel_Get(WinTitle="ahk_class XLMAIN", Excel7#=1) {
  WinGetClass, WinClass, %WinTitle%
  if (WinClass == "XLMAIN") {
    ControlGet, hwnd, hwnd, , Excel7%Excel7#%, %WinTitle%
    if Not ErrorLevel {
      window := Acc_ObjectFromWindow(hwnd, -16)
      if ComObjType(window) = 9
        while Not xl
          try xl := window.application
          catch e
            if SubStr(e.message,1,10) = "0x80010001"
              ControlSend, Excel7%Excel7#%, {Esc}, %WinTitle%
            else
              return "Error accessing the application object."
    }
  }
  return xl
}


Gui, Destroy
Gui, +AlwaysOnTop -SysMenu -Caption +Border
Gui, Margin, 20,20
Gui, Font, w400 s10, Arial
Gui, Add, Button, % "w135 h50 x20 gCopiaFormula", Copia Formula Excel
Gui, Add, Button, % "w135 h50 x+30 gSostituzioneQuery", Sostituzione query
GuiControl,, CopiaFormula
GuiControl,, SostituzioneQuery
Gui, Show, xCenter yCenter

GuiMove:
    PostMessage, 0xA1, 2
	return

Close:
	Gui, Destroy
	ExitApp

GuiEscape:
	Gui, Destroy
	ExitApp
    
SostituzioneQuery:
    Gui, Destroy
    try
      oWorkbook := Excel_Get() ; try to access active Workbook object
    catch
    ExitApp 
    InputBox, query, Inserire nome query, , ,375, 100
    if ErrorLevel
      ExitApp
	current  := oWorkbook.ActiveCell
    SQLquery := oWorkbook.ActiveCell.Value
    oWorkbook.ActiveCell.Value := query
	oWorkbook.ActiveCell.Font.Bold := true
    oWorkbook.Range("G2").Select
    
    While oWorkbook.ActiveCell.Value{
      if (oWorkbook.ActiveCell.Value = SQLquery){
        oWorkbook.ActiveCell.Value := query
		oWorkbook.ActiveCell.Font.Bold := true
		}
      oWorkbook.ActiveCell.Offset(1,0).select
    }
    
    oWorkbook.Range("J2").Select
    While oWorkbook.ActiveCell.Value{
      if (oWorkbook.ActiveCell.Value = SQLquery){
        oWorkbook.ActiveCell.Value := query
		oWorkbook.ActiveCell.Font.Bold := true
		}
      oWorkbook.ActiveCell.Offset(1,0).select
    }
	current.select
    ExitApp
    
CopiaFormula:
    Gui, Destroy
	try
      oWorkbook := Excel_Get() ; try to access active Workbook object
    catch
    ExitApp 
    
    MasterForm := oWorkbook.ActiveCell.Formula
    MasterInt  := SubStr(MasterForm, 1, InStr(MasterForm, ";"))
    oWorkbook.Selection.Formula := MasterForm
    
    While oWorkbook.ActiveCell.Formula
    {
      Form := oWorkbook.ActiveCell.Formula
      oWorkbook.ActiveCell.Formula := MasterInt SubStr(Form, InStr(Form, ";")+1)  
      oWorkbook.ActiveCell.Offset(0,1).select
    }
	ExitApp
