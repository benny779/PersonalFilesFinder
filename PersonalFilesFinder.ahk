#SingleInstance force



if !FileExist(A_ScriptDir "\ComputerList.txt")
{
	MsgBox, 16, , Error getting list of computers.`nPlease check the "ComputerList.txt" file.
	ExitApp
}


Loop, Read, % A_ScriptDir "\ComputerList.txt"
{
	if A_LoopReadLine
	{
		computer := A_LoopReadLine
		computernameesc := Escape_Characters(computer)
		if !%computernameesc%_scanned
		{
			if ComputerReturnPing(computer)
			{
				OutputDebug % "-- " computer
				MasterCount++
				
				vPath := "\\" computer "\" SubStr(systemdrive,1,1) "$\users\"
				vPathFolders := "Desktop,Documents,Downloads,Pictures"
				vPathSearchType := "FR"
				gosub, Loop

				vPath := "\\" computer "\" SubStr(systemdrive,1,1) "$"
				vPathFound := vPath
				vPathFolders := false
				vPathSearchType := "F"
				aaa := 5
				gosub, Loop
				aaa := 0
				%computernameesc%_scanned = 1
			}
		}
	}
}

MsgBox % "END SCANNING " MasterCount " COMPUTERS"
ExitApp

Loop:
if aaa = 5
	goto skip

Loop, Files, % vPath "*", D
{
	if A_LoopFileName not in Administrator,Public,user,Default,All Users,Default User,DefaultAppPool
	{
		vPathFound := vPath . A_LoopFileName "\"
		gosub skip
	}
}
return

skip:
		Loop, Parse, vPathFolders, `,
		{
			CurrPath := vPathFound . (!A_LoopField?"":A_LoopField) "\*"
			Loop, Files, % CurrPath, % vPathSearchType
			{
				if A_LoopFileExt in csv,xls,xlsx,doc,docx,txt,pdf
				{
					CurrFile := A_LoopFilePath
					Loop, Parse, % GetNumbers(CurrFile), |
					{
						if IDValidator(A_LoopField) ;valid ID
							Log(computer,CurrFile,A_LoopField)
					}
				}
			}
		}
return


	



esc::ExitApp

Log(computer,file,ID) {
	OutputDebug % "!!! " computer " - " file " - " ID
	FileAppend, % computer "," file "," ID "`n", % A_ScriptDir "\log.csv", UTF-8
}

ComputerReturnPing(computer:="localhost")
{
	for Item in ComObjGet("winmgmts:").ExecQuery("Select StatusCode From Win32_PingStatus where Address = '" computer "'")
		if Item.StatusCode = 0
			return true
		return false
}

Escape_Characters(Var)
{
   StringReplace, Var, Var, ., _, All
   StringReplace, Var, Var, -, _, All
   
   StringReplace, Var, Var, (, , All
   StringReplace, Var, Var, ), , All
   StringReplace, Var, Var, :, , All
   StringReplace, Var, Var, %A_Space%,_, All
   return Var
}

GetNumbers(str) {
	pos:=1
	while pos := RegExMatch(str, "(\d+)", Match, Pos + StrLen(Match))
		res .= (!res ? "" : "|") Match
	return res
}

IDValidator(ID)
{
	ID := Round(ID) ;Remove 0 from start
	
	;~ Check input requirements
	if ID is not digit
		return 0
	if StrLen(ID) > 9
		return 0
	if StrLen(ID) < 7
		return 0

	ID := Format("{:09}", ID)
	
	Loop, Parse, ID
	{
		ID = % A_LoopField
		If (Mod(A_Index, 2) = 0)
			ID := ID * 2
		VAR += (ID > 9) ? ID - 9 : ID
	}
	
	return SubStr(VAR, 2, 1) = 0 ? 1 : 0
}