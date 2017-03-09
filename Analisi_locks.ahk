/*
* Analisi_locks.ahk
* Avanzini Luca - 07/03/2017
* v1.04
*/

#NoEnv
;#Warn
SendMode Input 
SetWorkingDir %A_ScriptDir% 


/*
* VARIABILI GLOBALI
*/
		   TEMP_DIR := "D:\Temp\PEDROLLO"
		  TEMP_FILE := "D:\Temp\PEDROLLO\temporary_trace.xml"

	EVENT_BEGIN_STR := "<Event"
	  EVENT_END_STR := "</Event>"
	 
	 LOCK_BEGIN_STR := "<blocked-process-report"
	   LOCK_END_STR := "</blocked-process-report>"

 DEADLOCK_BEGIN_STR := "<deadlock-list>"
   DEADLOCK_END_STR := "</deadlock-list>"

	 PROC_BEGIN_STR := "<blocked-process>"
	   PROC_END_STR := "</blocked-process>"
	PROC2_BEGIN_STR := "<blocking-process>"
	  PROC2_END_STR := "</blocking-process>"
	  
	   ID_BEGIN_STR := "<process id="""
		 ID_END_STR := """"

	QUERY_BEGIN_STR := "<inputbuf>"
	  QUERY_END_STR := "</inputbuf>"
	  
 DURATION_BEGIN_STR := "<Column id=""13"" name=""Duration"">"
	START_BEGIN_STR := "<Column id=""14"" name=""StartTime"">"
	 COLUMN_END_STR := "</Column>"
	 
	LOGIN_BEGIN_STR := "loginname="""
	  LOGIN_END_STR := """ isolationlevel"

		 EventCount := 0
		  LockCount := 0
	  DeadlockCount := 0

			 Result := []
		  AllResult := []


		  
/*
* Auto-Parser for XML / HTML | By Skan | v1.0-196c 21-Nov-2009 www.autohotkey.com/forum/topic51354.html
*/			 
StrX(H,BS="",BO=0,BT=1,ES="",EO=0,ET=1,ByRef N="") { 
	Return SubStr(H,P:=(((Z:=StrLen(ES))+(X:=StrLen(H))+StrLen(BS)-Z-X)?((T:=InStr(H,BS,0,((BO
		<0)?(1):(BO))))?(T+BT):(X+1)):(1)),(N:=P+((Z)?((T:=InStr(H,ES,0,((EO)?(P+1):(0))))?(T-P+Z
		+(0-ET)):(X+P)):(X)))-P) 
}


/*
* Cast formatted string to int
*/
intParse(str) {
	str = %str%
	Loop, Parse, str
		If A_LoopField in 0,1,2,3,4,5,6,7,8,9,.,+,-
			int = %int%%A_LoopField%
	Return, int + 0.0
}


/*
* Sort multidimensional array
*/
SortArray2DByElement(ByRef Array, Element) {
   Static Delim := Chr(1)
   SortVar := ""
   For K, V In Array
      SortVar .= V[Element] . Delim . K . "`n"
   Sort, SortVar
   NewArray := {}
   Loop, Parse, SortVar, `n
   {
      If (A_LoopField) {
         StringSplit, Part, A_LoopField, %Delim%
         NewArray.Insert(Array[Part2])
      }
   }
   Array := NewArray
}


/*
* Export array Data in Excel
*/
ExportInExcel(Data, Intestazioni){
	; Crea nuovo foglio excel
	oExcel := ComObjCreate("Excel.Application")
	oExcel.Workbooks.Add 
	oExcel.Range("A1").Select

	; Valorizza la riga di intestazione
	for i, desc in Intestazioni
		oExcel.ActiveCell.Offset(0,i-1).Value := desc
	oExcel.Range("A1:I1").Interior.ColorIndex := 1
	oExcel.Range("A1:I1").Font.ColorIndex := 2
	oExcel.Range("A2").Select

	; Valorizza le celle
	for i, row in Data{
		for j, col in row{
			oExcel.ActiveCell.Offset( i-1,j-1).Value := col
			if (j=7 or j=9)  and (Not(InStr(col, "update") or InStr(col, "insert")))
				oExcel.ActiveCell.Offset( i-1,j-1).Font.Color := -16776961
		}
		_tmp := (i * 100)/Data.Length()
		GuiControl,,ExcelProgress,%_tmp%
	}

	; Rimuovi la colonna con lo StartTime
	oExcel.Columns(10).EntireColumn.Delete
	oExcel.Columns(2).EntireColumn.Delete
	
	; Abilita l'autofit
	oExcel.Range("A1:H" Data.Length()).Select
	oExcel.Selection.Columns.AutoFit
	oExcel.Range("A2").Select
	oExcel.Visible := 1
}
	
	

; Selezione del file di input
FileSelectFile, inputFilePath, 3, %TEMP_DIR%,Selezionare il trace file, *.xml
If (inputFilePath = "")
	ExitApp
FileCopy, %inputFilePath%, %TEMP_FILE%, 1
FileRead, inputFile, %TEMP_FILE%

; Modifica del working file per risolvere i riferimenti HTML
inputFile := StrReplace(inputFile, "&lt;", "<")
inputFile := StrReplace(inputFile, "&gt;", ">")
inputFile := StrReplace(inputFile, "&amp;apos;", "'")
inputFile := StrReplace(inputFile, "&amp;quot;", """")

; Recupero degli eventi
inputFile := StrX(inputFile, "<Events>", 1, StrLen("<Events>"), "</Events>", 0, StrLen("</Events>"), "")

; Feedback grafico
Gui, +AlwaysOnTop -MinimizeBox -MaximizeBox
Gui, Margin, 20,20
Gui, Font, w400 s11, Verdana
Gui, Add, Text, +Center, Elaborazione eventi
Gui, Add, Progress, w300 h20 -0x00000001 vElabProgress
Gui, Add, Text, +Center, Ordinamento eventi
Gui, Add, Progress, w300 h20 -0x00000001 vSortProgress
Gui, Add, Text, +Center, Calcolo durata eventi
Gui, Add, Progress, w300 h20 -0x00000001 vCalcProgress
Gui, Add, Text, +Center, Creazione file Excel
Gui, Add, Progress, w300 h20 -0x00000001 vExcelProgress
Gui, Show, xCenter yCenter

; Lettura degli eventi
While _event := StrX(inputFile, EVENT_BEGIN_STR, N, 0, EVENT_END_STR, 1, 0, N)
{
	; GESTIONE LOCK
	If ((_report := StrX(_event, LOCK_BEGIN_STR, 1, 0, LOCK_END_STR, 1, 0, "")) != "")
	{
		; Lettura dati
		_blocked  := StrX(_report, PROC_BEGIN_STR, 1, 0, PROC_END_STR, 1, 0, "")
		_client1  := StrX(_blocked, "clientapp=""", 1, StrLen("clientapp="""), """", 1, StrLen(""""), "")
		_login1   := StrX(_blocked, LOGIN_BEGIN_STR, 1, StrLen(LOGIN_BEGIN_STR), LOGIN_END_STR, 1, StrLen(LOGIN_END_STR), "")
		_query1   := StrX(_blocked, QUERY_BEGIN_STR, 1, StrLen(QUERY_BEGIN_STR), QUERY_END_STR, 1, StrLen(QUERY_END_STR), "")
		_processId:= StrX(_blocked, ID_BEGIN_STR, 1, StrLen(ID_BEGIN_STR), ID_END_STR, 1, StrLen(ID_END_STR), "")
		_ownerId  := StrX(_blocked, "ownerId=""", 1, StrLen("ownerId="""), """", 1, StrLen(""""), "")
		_desc     := StrX(_blocked, "XDES=""", 1, StrLen("XDES="""), """", 1, StrLen(""""), "")
		
		_blocking := StrX(_report, PROC2_BEGIN_STR, 1, 0, PROC2_END_STR, 1, 0, "")
		_client2  := StrX(_blocking, "clientapp=""", 1, StrLen("clientapp="""), """", 1, StrLen(""""), "")
		_login2   := StrX(_blocking, LOGIN_BEGIN_STR, 1, StrLen(LOGIN_BEGIN_STR), LOGIN_END_STR, 1, StrLen(LOGIN_END_STR), "")
		_query2   := StrX(_blocking, QUERY_BEGIN_STR, 1, StrLen(QUERY_BEGIN_STR), QUERY_END_STR, 1, StrLen(QUERY_END_STR), "")
				
		_duration := StrX(_event, DURATION_BEGIN_STR, 1, StrLen(DURATION_BEGIN_STR), COLUMN_END_STR, 1, StrLen(COLUMN_END_STR), "")
		_startTime:= StrX(_event, START_BEGIN_STR, 1, StrLen(START_BEGIN_STR), COLUMN_END_STR, 1, StrLen(COLUMN_END_STR), "")

		; Elaborazione dati
		If (_login1 = "")
			_login1 := SubStr(_client1, 67, StrLen(SubStr(_client1, 67))-10)
		If (_login2 = "")
			_login2 := SubStr(_client2, 67, StrLen(SubStr(_client2, 67))-10)
		_login1   := StrReplace(_login1, "PEDROLLOSPA\", "")
		_login2   := StrReplace(_login2, "PEDROLLOSPA\", "")
		_duration := Floor(intParse(_duration)/1000000)
		_date     := SubStr(_startTime,1,10)
		_time     := SubStr(SubStr(_startTime, 1, StrLen(_startTime)-6),12)
		_startTime:= SubStr(_startTime,1,10) " " SubStr(SubStr(_startTime, 1, StrLen(_startTime)-6),12)
		_id 	  := _processId _ownerId _desc
		_query1 := StrReplace(StrReplace(_query1, "&amp;gt;", ">"), "&amp;lt;", "<")
		_query2 := StrReplace(StrReplace(_query2, "&amp;gt;", ">"), "&amp;lt;", "<")
		
		; Scrittura dati
		AllResult.Push(["lock", _startTime, _date, _time, _duration, _login1, _query1, _login2, _query2, _id])
		if(_duration < 10)
		{
			LockCount := LockCount + 1	
			Result.Push(["lock", _startTime, _date, _time, _duration, _login1, _query1, _login2, _query2, _id])
		}
	}
	; GESTIONE DEADLOCK
	Else If ((_report := StrX(_event, DEADLOCK_BEGIN_STR, 1, 0, DEADLOCK_END_STR, 1, 0, "")) != "")
	{
		; da implementare
		DeadlockCount := DeadlockCount + 1
	}
	EventCount := EventCount + 1
	GuiControl,,ElabProgress,%EventCount%
}
GuiControl,,ElabProgress,100

If FileExist(TEMP_FILE)
	FileDelete, %TEMP_FILE%

; Ordina gli array in base allo start time
GuiControl,,SortProgress,0
SortArray2DByElement(Result, 2)
GuiControl,,SortProgress,50
SortArray2DByElement(AllResult, 2)
GuiControl,,SortProgress,100

max := 0
prevId := AllResult[1][10]
for i, row in AllResult{
	if( row[10] != prevId ){
		for _i, _row in Result{
			if( _row[10] = prevId ){
				_row[5] := max
				break
			}		
		}
		max := 0
		prevId := row[10]
	}
	if(row[5] > max)
		max := row[5]
	_tmp := (i * 100)/AllResult.Length()
	GuiControl,,CalcProgress,%_tmp%
}


; Export dati in excel
Intest := ["TIPO", "START TIME", "DATA", "TIME", "DURATION", "LOGIN1", "QUERY1", "LOGIN2", "QUERY2", "ID"]
ExportInExcel(Result, Intest)

Gui, Destroy
MsgBox, Elaborazione completata `n`n- Lock : %LockCount% `n- Deadlock : %DeadlockCount%

ExitApp

