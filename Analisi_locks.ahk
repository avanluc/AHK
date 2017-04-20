/*
* AutoHotkey Version: AutoHotkey 1.1
* Author: 			  Avanzini Luca
* Description:		  Extract information about locks and deadlocks from a SQL trace file
* Last Modification:  13/04/2017
* Version:  		  v1.2
*/

#NoEnv
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
	oExcel.Range("A1").EntireRow.Interior.ColorIndex := 1
	oExcel.Range("A1").EntireRow.Font.ColorIndex := 2
	oExcel.Range("A2").Select

	; Valorizza le celle
	for i, row in Data{
		for j, col in row{
			oExcel.ActiveCell.Offset( i-1,j-1).Value := col
			if (j=8 or j=11) and (InStr(col, "select")) and (Not(InStr(col, "@P1")))
				oExcel.ActiveCell.Offset( i-1,j-1).Font.Color := -16776961
		}
		GuiControl,,ProgressText, % "Compilazione righe Excel  ( " i " / " Data.Length() " )"
		GuiControl,,ProgressStatus,% (i * 100)/Data.Length()
	}

	; Rimuovi la colonna con lo StartTime e l'identificativo
	if(Data[1][1] = "lock")
		oExcel.Columns(12).EntireColumn.Delete
	oExcel.Columns(2).EntireColumn.Delete
	
	; Abilita l'autofit e setta MAX WIDTH
	MAX_WIDTH := 70
	oExcel.Columns.AutoFit
	for i in Data[1]
		if(oExcel.Columns(i).ColumnWidth > MAX_WIDTH)
			oExcel.Columns(i).ColumnWidth := MAX_WIDTH
		
	oExcel.ActiveSheet.Name := "Analisi " Data[1][1]
	oExcel.Range("A2").Select
	
	return oExcel
}
	

/*
* Check if a process is an Exchange Event
*/
RepeatedProcess(Array, query){
	for i, row in Array
		if (row[4] = query)
			return true
	return false
}


/*
* Evaluate each lock maximum duration
*/
EvaluateDuration(ByRef AllData, ByRef Data){
	max := 0
	prevId := AllData[1][12]
	for i, row in AllData{
		if( row[12] != prevId ){
			for _i, _row in Data{
				if( _row[12] = prevId ){
					_row[5] := max
					break
				}		
			}
			max := 0
			prevId := row[12]
		}
		if(row[5] > max)
			max := row[5]
		GuiControl,,ProgressText, % "Calcolo durata eventi  ( " i " / " AllData.Length() " )"
		GuiControl,,ProgressStatus,% (i * 100)/AllData.Length()
	}
}


/*
* Analyze a deadlock resources and return a record with the information extracted
*/
AnalyzeResource(tipo, lockText, _TempProcess){
	if (tipo = "pagelock")
		__strId := "pageid="""
	else if (tipo = "keylock")
		__strId := "indexname="""
	else if (tipo = "objectlock")
		__strId := "NULL_STRING"
	else
		return []
	
	_Owners 	:= []
	_Waiters 	:= []
	_pageId  	:= StrX(lockText, __strId, 1, StrLen(__strId), """", 1, StrLen(""""), "")
	_objId  	:= StrX(lockText, "objectname=""", 1, StrLen("objectname="""), """", 1, StrLen(""""), "")
	_objId 		:= StrReplace(_objId, "AHE_PEDR.dbo.", "") " "
	_ownerList  := StrX(lockText, "<owner-list>", 1, StrLen("<owner-list>"), "</owner-list>", 1, StrLen("</owner-list>"), "")
	_waiterList := StrX(lockText, "<waiter-list>", 1, StrLen("<waiter-list>"), "</waiter-list>", 1, StrLen("</waiter-list>"), "")
	K1 = 0
	while _owner := StrX(_ownerList, "<owner", K1, 0, ">", 1, 0, K1)
	{
		_ownerId   := StrX(_owner, "<owner id=""", 1, StrLen("<owner id="""), """", 1, StrLen(""""), "")
		_ownerMode := StrX(_owner, "mode=""", 1, StrLen("mode="""), """", 1, StrLen(""""), "")
		for i, row in _TempProcess
			if(_ownerId = row[1]){
				_Owners.push([i, _ownerMode, row[4]])
				break
			}
	}
	K2 = 0
	while _waiter := StrX(_waiterList, "<waiter", K2, 0, ">", 1, 0, K2)
	{
		_waiterId   := StrX(_waiter, "<waiter id=""", 1, StrLen("<waiter id="""), """", 1, StrLen(""""), "")
		_waiterMode := StrX(_waiter, "mode=""", 1, StrLen("mode="""), """", 1, StrLen(""""), "")
		for i, row in _TempProcess
			if(_waiterId = row[1]){
				_Waiters.push([i, _ownerMode, row[4]])
				break
			}
	}
	return [tipo, _pageId, _objId, _Owners[1][1], _Owners[1][2], _Waiters[1][1], _Waiters[1][2]]
}


CheckClient(client){
	if (client = "")
		return " "
	else if (InStr(client, "ENTERPRISE"))
		return "AHE"
	else if (InStr(client, "SQL Server"))
		return "SQL Server Management Studio"
	else 
		return client
}

/*
* Inizio Esecuzione
*/
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
Gui, +AlwaysOnTop -SysMenu -Caption +Border
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
		_client1  := CheckClient(_client1)
		_client2  := CheckClient(_client2)
		_login1   := StrReplace(_login1, "PEDROLLOSPA\", "") " "
		_login2   := StrReplace(_login2, "PEDROLLOSPA\", "") " "
		_login1   := StrReplace(_login1, "LINZELECTRIC\", "") " "
		_login2   := StrReplace(_login2, "LINZELECTRIC\", "") " "
		_duration := Floor(intParse(_duration)/1000000)
		_date     := SubStr(_startTime,1,10)
		_time     := SubStr(SubStr(_startTime, 1, StrLen(_startTime)-6),12)
		_startTime:= SubStr(_startTime,1,10) " " SubStr(SubStr(_startTime, 1, StrLen(_startTime)-6),12)
		_id 	  := _processId _ownerId _desc
		_query1   := RegExReplace(RegExReplace(_query1, "^`r`n[\t]+", ""), "`n", "")
		_query2   := RegExReplace(RegExReplace(_query2, "^`r`n[\t]+", ""), "`n", "")
		
		; Scrittura dati
		AllResult.Push(["lock", _startTime, _date, _time, _duration, _client1, _login1, _query1, _client2, _login2, _query2, _id])
		if(_duration < 10){
			LockCount := LockCount + 1	
			Result.Push(["lock", _startTime, _date, _time, _duration, _client1, _login1, _query1, _client2, _login2, _query2, _id, " "])
		}
	}
	; GESTIONE DEADLOCK
	else if ((_report := StrX(_event, DEADLOCK_BEGIN_STR, 1, 0, DEADLOCK_END_STR, 1, 0, "")) != "")
	{
		_TempProcess := []
		_victim  := StrX(_report, "<deadlock victim=""", 1, StrLen("<deadlock victim="""), """", 1, StrLen(""""), "")
		_ExcEvt := 0
		_N := 1
		
		; Ciclo sui processi coinvolti nel deadlock
		while _process := StrX(_report, "<process id", _N, 0, "</process>", 1, 0, _N)
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
			_login := StrReplace(_login, "PEDROLLOSPA\", "") " "
			_login := StrReplace(_login, "LINZELECTRIC\", "") " "
			_client := CheckClient(_client)
			_query := RegExReplace(RegExReplace(_query, "^`r`n[\t]+", ""), "`n", "")
			_id    := _processId _ownerId _desc
			
			if( Not(RepeatedProcess(_TempProcess, _query)) )
				_TempProcess.Push([_processId, _id, _client, _login, _query])
			else 
				_ExcEvt := _ExcEvt + 1 
		}
		; Lettura dati
		_startTime := StrX(_event, START_BEGIN_STR, 1, StrLen(START_BEGIN_STR), COLUMN_END_STR, 1, StrLen(COLUMN_END_STR), "")
		
		
		; Analisi delle risorse coinvolte nel deadlock
		/*
		_Resources := []
		__N := 0
		while _pagelock := StrX(_report, "<pagelock", __N, 0, "</pagelock>", 1, 0, __N)
			_Resources.Push(AnalyzeResource("pagelock", _pagelock, _TempProcess))
		__N := 0
		while _keylock := StrX(_report, "<pagelock", __N, 0, "</pagelock>", 1, 0, __N)
			_Resources.Push(AnalyzeResource("keylock", _pagelock, _TempProcess))
		__N := 0
		while _objectlock := StrX(_report, "<objectlock", __N, 0, "</objectlock>", 1, 0, __N)
			_Resources.Push(AnalyzeResource("objectlock", _pagelock, _TempProcess))
		
		if (_Resources.Length() = 3)
			
		else if (_Resources.Length() = 2)
			
		*/
		
		; Elaborazione dati
		_date     := SubStr(_startTime,1,10)
		_time     := SubStr(SubStr(_startTime, 1, StrLen(_startTime)-6),12)
		_startTime:= SubStr(_startTime,1,10) " " SubStr(SubStr(_startTime, 1, StrLen(_startTime)-6),12)
		_id 	  := _processId _ownerId _desc
		
		; Scrittura dati
		_Appo := ["deadlock", _startTime, _date, _time, _ExcEvt]
		for i, row in _TempProcess {
			_Appo.Push(row[3])
			_Appo.Push(row[4])
			_Appo.Push(row[5])
		}
		_Appo.Push(" ")
		ResultDead.Push(_Appo)
		DeadlockCount := DeadlockCount + 1
	}
	
	EventCount := EventCount + 1
	GuiControl,,ProgressText, % "Elaborazione eventi  ( " EventCount " / " EventTot " )"
	GuiControl,,ProgressStatus,% (EventCount*100)/EventTot
}
AllResult.Push(["lock", _startTime, _date, _time, _duration, _login1, _query1, _login2, _query2, ""])

if FileExist(TEMP_FILE)
	FileDelete, %TEMP_FILE%

; Ordina gli array in base allo start time e calcola la durata dei lock
SortArray2DByElement(Result, 2)
SortArray2DByElement(AllResult, 2)
SortArray2DByElement(ResultDead, 2)
EvaluateDuration(AllResult, Result)
GuiControl,,ProgressText, % "Creazione file Excel.."

; Export dati in excel
Intest := ["TIPO", "START TIME", "DATA", "TIME", "DURATA (s)", "CLIENT1", "LOGIN1", "QUERY1", "CLIENT2", "LOGIN2", "QUERY2", "ID"]
oExcel := ExportInExcel(Result, Intest)
Intest := ["TIPO", "START TIME", "DATA", "TIME", "EXCHANGE EVENT", "CLIENT1", "LOGIN1", "QUERY1 (vittima)", "CLIENT2", "LOGIN2", "QUERY2", "LOGIN3", "QUERY3"]
oExcel := ExportInExcel(ResultDead, Intest, oExcel)
oExcel.Visible := 1

Gui, Destroy
MsgBox, , Risultati, % "Elaborazione completata `n`n- Lock : " LockCount "`n- Deadlock : " DeadlockCount

ExitApp

