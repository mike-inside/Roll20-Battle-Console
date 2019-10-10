#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
;SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
;SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

MakeCSV(filename := "Arsenal", sort := false, folder := false) {
	if !folder
		folder := A_ScriptDir . "\csv\"
	pathIn :=  folder . filename . ".xlsx"
	pathOut := folder . filename . ".csv"
	existIn := FileExist( pathIn )
	existOut := FileExist( pathOut )
	if (!existIn and !existOut) {
		MsgBox, % filename . " file does not exist`n`nProgram will now exit"
		ExitApp
	} else if (existIn) {
		dateX := 0
		if (existOut) {
			FileGetTime, dateX, % pathIn
			FileGetTime, dateC, % pathOut
			EnvSub, dateX, %dateC%
		}
		if (dateX >= 0) 
		{
			if existOut
				FileDelete, % pathOut
			xl := ComObjCreate("Excel.Application").Workbooks.Open(pathIn)
			if sort
				xl.ActiveSheet.Range(StrReplace(xl.ActiveSheet.UsedRange.Address, "$1:", "$2:")).Sort(xl.ActiveSheet.Range(sort),1 )
			xl.SaveAs(pathOut, 6) 	; type 6 is CSV c.f. https://msdn.microsoft.com/en-us/library/office/ff198017.aspx
			xl.close(0) 
		}
	}
}