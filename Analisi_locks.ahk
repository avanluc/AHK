/*
* Analisi_locks.ahk
* Avanzini Luca - 07/03/2017
* v1.05
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

		   EventTot := 0
		 EventCount := 0
		  LockCount := 0
	  DeadlockCount := 0

			 Result := []
		  AllResult := []
		 ResultDead := []


		  
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
ExportInExcel(Data, Intestazioni, ByRef oExcel=0){

	if (Data.Length() = 0)
		return oExcel
	
	if (oExcel = 0){
		; Crea nuovo foglio excel
		oExcel := ComObjCreate("Excel.Application")
		oExcel.Workbooks.Add 
	}
	else
		oExcel.WorkSheets.Add()

	; Valorizza la riga di intestazione
	oExcel.Range("A1").Select
	for i, desc in Intestazioni
		oExcel.ActiveCell.Offset(0,i-1).Value := desc
	oExcel.Range("A1:I1").Interior.ColorIndex := 1
	oExcel.Range("A1:I1").Font.ColorIndex := 2
	oExcel.Range("A2").Select

	; Valorizza le celle
	for i, row in Data{
		for j, col in row{
			oExcel.ActiveCell.Offset( i-1,j-1).Value := col
			if (j=7 or j=9) and (Not(InStr(col, "update") or InStr(col, "insert")))
				oExcel.ActiveCell.Offset( i-1,j-1).Font.Color := -16776961
		}
		GuiControl,,ProgressText, % "Compilazione righe Excel  ( " i " / " Data.Length() " )"
		GuiControl,,ProgressStatus,% (i * 100)/Data.Length()
	}

	; Rimuovi la colonna con lo StartTime
	oExcel.Columns(10).EntireColumn.Delete
	oExcel.Columns(2).EntireColumn.Delete
	
	; Abilita l'autofit
	oExcel.Range("A1:H" Data.Length()).Select
	oExcel.Selection.Columns.AutoFit
	oExcel.ActiveSheet.Name := "Analisi " Data[1][1]
	oExcel.Range("A2").Select
	
	return oExcel
}
	

/*
* Check if a process is an Exchange Event
*/
RepeatedProcess(Array, query){
	for i, row in Array
		if (Array[4] = query)
			return true
	return false
}


/*
* Lettura ed elaborazione dati di un lock
*/
AnalyzeLock(Event, Report, ByRef Result, ByRef AllResult, ByRef Counter ){
	
	; Lettura dati
	blocked  := StrX(Report, PROC_BEGIN_STR, 1, 0, PROC_END_STR, 1, 0, "")
	client1  := StrX(blocked, "clientapp=""", 1, StrLen("clientapp="""), """", 1, StrLen(""""), "")
	login1   := StrX(blocked, LOGIN_BEGIN_STR, 1, StrLen(LOGIN_BEGIN_STR), LOGIN_END_STR, 1, StrLen(LOGIN_END_STR), "")
	query1   := StrX(blocked, QUERY_BEGIN_STR, 1, StrLen(QUERY_BEGIN_STR), QUERY_END_STR, 1, StrLen(QUERY_END_STR), "")
	processId:= StrX(blocked, ID_BEGIN_STR, 1, StrLen(ID_BEGIN_STR), ID_END_STR, 1, StrLen(ID_END_STR), "")
	ownerId  := StrX(blocked, "ownerId=""", 1, StrLen("ownerId="""), """", 1, StrLen(""""), "")
	desc     := StrX(blocked, "XDES=""", 1, StrLen("XDES="""), """", 1, StrLen(""""), "")
	
	blocking := StrX(Report, PROC2_BEGIN_STR, 1, 0, PROC2_END_STR, 1, 0, "")
	client2  := StrX(blocking, "clientapp=""", 1, StrLen("clientapp="""), """", 1, StrLen(""""), "")
	login2   := StrX(blocking, LOGIN_BEGIN_STR, 1, StrLen(LOGIN_BEGIN_STR), LOGIN_END_STR, 1, StrLen(LOGIN_END_STR), "")
	query2   := StrX(blocking, QUERY_BEGIN_STR, 1, StrLen(QUERY_BEGIN_STR), QUERY_END_STR, 1, StrLen(QUERY_END_STR), "")
			
	duration := StrX(Event, DURATION_BEGIN_STR, 1, StrLen(DURATION_BEGIN_STR), COLUMN_END_STR, 1, StrLen(COLUMN_END_STR), "")
	startTime:= StrX(Event, START_BEGIN_STR, 1, StrLen(START_BEGIN_STR), COLUMN_END_STR, 1, StrLen(COLUMN_END_STR), "")

	; Elaborazione dati
	if (login1 = "")
		login1 := SubStr(client1, 67, StrLen(SubStr(client1, 67))-10)
	if (login2 = "")
		login2 := SubStr(client2, 67, StrLen(SubStr(client2, 67))-10)
	login1   := StrReplace(login1, "PEDROLLOSPA\", "")
	login2   := StrReplace(login2, "PEDROLLOSPA\", "")
	duration := Floor(intParse(duration)/1000000)
	date     := SubStr(startTime,1,10)
	time     := SubStr(SubStr(startTime, 1, StrLen(startTime)-6),12)
	startTime:= SubStr(startTime,1,10) " " SubStr(SubStr(startTime, 1, StrLen(startTime)-6),12)
	id 	     := processId ownerId desc
	
	; Scrittura dati
	AllResult.Push(["lock", startTime, date, time, duration, login1, query1, login2, query2, id])
	if(duration < 10)
	{
		Counter := Counter + 1	
		Result.Push(["lock", startTime, date, time, duration, login1, query1, login2, query2, id])
	}
}

; Selezione del file di input
FileSelectFile, inputFilePath, 3, %TEMP_DIR%,Selezionare il trace file, *.xml
if (inputFilePath = "")
	ExitApp
FileCopy, %inputFilePath%, %TEMP_FILE%, 1
FileRead, inputFile, %TEMP_FILE%

; Modifica del working file per risolvere i riferimenti HTML
inputFile := StrReplace(inputFile, "&amp;", "&")
inputFile := StrReplace(inputFile, "&lt;", "<")
inputFile := StrReplace(inputFile, "&gt;", ">")
inputFile := StrReplace(inputFile, "&apos;", "'")
inputFile := StrReplace(inputFile, "&quot;", """")

; Recupero degli eventi
inputFile := StrX(inputFile, "<Events>", 1, StrLen("<Events>"), "</Events>", 0, StrLen("</Events>"), "")

; Feedback grafico
Gui, +AlwaysOnTop -MinimizeBox -MaximizeBox
Gui, Margin, 20,20
Gui, Add, Text, w300 +Center vProgressText
Gui, Add, Progress, w300 h20 +Smooth BackgroundC9C9C9 vProgressStatus
Gui, Show, xCenter yCenter

StringReplace, inputFile, inputFile, %EVENT_BEGIN_STR%, %EVENT_BEGIN_STR%, UseErrorLevel
EventTot := ErrorLevel

; Lettura degli eventi
while _event := StrX(inputFile, EVENT_BEGIN_STR, N, 0, EVENT_END_STR, 1, 0, N)
{
	; GESTIONE LOCK
	if ((_report := StrX(_event, LOCK_BEGIN_STR, 1, 0, LOCK_END_STR, 1, 0, "")) != "")
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
		if (_login1 = "")
			_login1 := SubStr(_client1, 67, StrLen(SubStr(_client1, 67))-10)
		if (_login2 = "")
			_login2 := SubStr(_client2, 67, StrLen(SubStr(_client2, 67))-10)
		_login1   := StrReplace(_login1, "PEDROLLOSPA\", "")
		_login2   := StrReplace(_login2, "PEDROLLOSPA\", "")
		_duration := Floor(intParse(_duration)/1000000)
		_date     := SubStr(_startTime,1,10)
		_time     := SubStr(SubStr(_startTime, 1, StrLen(_startTime)-6),12)
		_startTime:= SubStr(_startTime,1,10) " " SubStr(SubStr(_startTime, 1, StrLen(_startTime)-6),12)
		_id 	  := _processId _ownerId _desc
		
		; Scrittura dati
		AllResult.Push(["lock", _startTime, _date, _time, _duration, _login1, _query1, _login2, _query2, _id])
		if(_duration < 10)
		{
			LockCount := LockCount + 1	
			Result.Push(["lock", _startTime, _date, _time, _duration, _login1, _query1, _login2, _query2, _id])
		}
		
		;AnalyzeLock(_event, _report, Result, AllResult, LockCount)
	}
	; GESTIONE DEADLOCK
	else if ((_report := StrX(_event, DEADLOCK_BEGIN_STR, 1, 0, DEADLOCK_END_STR, 1, 0, "")) != "")
	{
		_TempProcess := []
		_victim  := StrX(_report, "<deadlock victim=""", 1, StrLen("<deadlock victim="""), """", 1, StrLen(""""), "")
		_ExcEvt := false
		_N := 0
		
		; Ciclo sui processi coinvolti nel deadlock
		
		while _process := StrX(_report, "<process  id", _N, 0, "</process>", 1, 0, _N)
		{
			; Lettura dati
			_processId:= StrX(_process, ID_BEGIN_STR, 1, StrLen(ID_BEGIN_STR), ID_END_STR, 1, StrLen(ID_END_STR), "")
			_client   := StrX(_process, "clientapp=""", 1, StrLen("clientapp="""), """", 1, StrLen(""""), "")
			_login    := StrX(_process, LOGIN_BEGIN_STR, 1, StrLen(LOGIN_BEGIN_STR), LOGIN_END_STR, 1, StrLen(LOGIN_END_STR), "")
			_query    := StrX(_process, QUERY_BEGIN_STR, 1, StrLen(QUERY_BEGIN_STR), QUERY_END_STR, 1, StrLen(QUERY_END_STR), "")
			_ownerId  := StrX(_process, "ownerId=""", 1, StrLen("ownerId="""), """", 1, StrLen(""""), "")
			_desc     := StrX(_process, "XDES=""", 1, StrLen("XDES="""), """", 1, StrLen(""""), "")
			
			; Elaborazione dati
			if (_login = "")
				_login := SubStr(_client, 67, StrLen(SubStr(_client, 67))-10)
			_login := StrReplace(_login, "PEDROLLOSPA\", "")
			_id    := _processId _ownerId _desc
			
			if( Not(RepeatedProcess(_TempProcess, _query)) )
				_TempProcess.Push([_processId, _id, _login, _query])
			else 
				_ExcEvt := true
			__A := SubStr(_report,_N)
		}
		
		_startTime:= StrX(_event, START_BEGIN_STR, 1, StrLen(START_BEGIN_STR), COLUMN_END_STR, 1, StrLen(COLUMN_END_STR), "")
		
		_date     := SubStr(_startTime,1,10)
		_time     := SubStr(SubStr(_startTime, 1, StrLen(_startTime)-6),12)
		_startTime:= SubStr(_startTime,1,10) " " SubStr(SubStr(_startTime, 1, StrLen(_startTime)-6),12)
		_id 	  := _processId _ownerId _desc

		_Appo := ["deadlock", _startTime, _date, _time]
		for i, row in _TempProcess {
			_Appo.Push(row[3])
			_Appo.Push(row[4])
		}
		_Appo.Push(_ExcEvt)
		ResultDead.Push(_Appo)
		
		DeadlockCount := DeadlockCount + 1
	}
	
	EventCount := EventCount + 1
	GuiControl,,ProgressText, % "Elaborazione eventi  ( " EventCount " / " EventTot " )"
	GuiControl,,ProgressStatus,% (EventCount*100)/EventTot
}

if FileExist(TEMP_FILE)
	FileDelete, %TEMP_FILE%

; Ordina gli array in base allo start time
SortArray2DByElement(Result, 2)
SortArray2DByElement(AllResult, 2)

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
	GuiControl,,ProgressText, % "Calcolo durata eventi  ( " i " / " AllResult.Length() " )"
	GuiControl,,ProgressStatus,% (i * 100)/AllResult.Length()
}


; Export dati in excel
Intest := ["TIPO", "START TIME", "DATA", "TIME", "DURATION", "LOGIN1", "QUERY1", "LOGIN2", "QUERY2", "ID"]
oExcel := ExportInExcel(Result, Intest)
Intest := ["TIPO", "START TIME", "DATA", "TIME", "LOGIN1", "QUERY1", "LOGIN2", "QUERY2", "ID"]
oExcel := ExportInExcel(ResultDead, Intest, oExcel)
oExcel.Visible := 1

Gui, Destroy
MsgBox % "Elaborazione completata `n`n- Lock : " LockCount "`n- Deadlock : " DeadlockCount

ExitApp

