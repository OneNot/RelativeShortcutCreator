;get target
try {
	FileSelectFile, files, M3,,Select shortcut target
} catch e {
	MsgBox Something went wrong: %e%
}
if (files = "")
{
    Exit
}

TargetPath := ""
Loop, parse, files, `n
{
    if (A_Index = 1)
    {
		TargetPath := GetRelativePath(A_LoopField)
	}
    else
    {
		;Create target from path + current iteration file
		Target := TargetPath . A_LoopField
		;trim trailing \ and add quotes
		Target := RTrim(Target, "\")
		Target :=  """" . Target . """"
		
		ShortcutName := A_LoopField . " - Relative Shortcut.lnk"
		HoverMsg := RTrim(ShortcutName, ".lnk")

		;create shortcut
		Explorer := "%windir%\explorer.exe"
		FileCreateShortcut,%Explorer%, %ShortcutName%,,%Target%,%HoverMsg%,%windir%\explorer.exe,,,1
    }
}


GetRelativePath(path)
{
	;init
	RelPath := ""
	WorkingDirSplit := []
	TargetDirSplit := []
	WorkingDirSplit := StrSplit(A_WorkingDir, "\")
	TargetDirSplit := StrSplit(path, "\")
	
	;Get number of same parts in working dir and target file dir
	NumOfSameParts := 0
	AfterSameParts := ""
	Loop
	{
		if(A_Index > WorkingDirSplit.Count() or A_Index > TargetDirSplit.Count())
			break
	
		if(WorkingDirSplit[A_Index] = TargetDirSplit[A_Index])
		{
			NumOfSameParts += 1
			AfterSameParts := 
		}
		else
			break
	}
	if(NumOfSameParts = 0) 
	{
		MsgBox, Error: Cannot create a relative path shortcut to another drive
		Exit
	}
	
	;How far is the last same part from working dir
	HowFarFromWorkDir := WorkingDirSplit.Count() - NumOfSameParts
	
	;set the number of steps back (..\) to get to the last same part
	Loop, %HowFarFromWorkDir%
	{
		RelPath := RelPath . "..\"
	}
	;if taget file is in working dir
	if(HowFarFromWorkDir = 0) 
	{
		RelPath := ".\"
	}
	
	;Add the rest of the path
	Loop
	{
		if(NumOfSameParts + A_Index > TargetDirSplit.Count())
			break
		else if(StrLen(TargetDirSplit[NumOfSameParts + A_Index]) > 0)
		{
			RelPath := RelPath . TargetDirSplit[NumOfSameParts + A_Index] . "\"
		}
	}
		
	return RelPath
}
