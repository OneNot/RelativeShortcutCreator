;get target
try {
	FileSelectFile, TestFile, 3,,Select shortcut target
} catch e {
	MsgBox Something went wrong: %e%
}

;make some variables from target and working dir
WorkingDirSplit := []
TargetDirSplit := []
WorkingDirSplit := StrSplit(A_WorkingDir, "\")
TargetDirSplit := StrSplit(TestFile, "\")
WorkingDirSplitLength := WorkingDirSplit.MaxIndex()
TargetDirSplitLength := TargetDirSplit.MaxIndex()

;initialization
NumOfSameparts := 0
ArrayCount := 1

;calculate the number of same parts in paths 
;example: "C:\test\test2\test3\" AND "C:\test\test2\something.exe" would have 3 same parts: "C:", "test" and "test2"
Loop
{
	if(ArrayCount = (WorkingDirSplitLength) or ArrayCount = (WorkingDirSplitLength)) 
	{
		NumOfSameparts := ArrayCount
		break
	}

	if(WorkingDirSplit[ArrayCount] = TargetDirSplit[ArrayCount])
		NumOfSameparts += 1
	else
		break
		
	ArrayCount += 1
}

if(NumOfSameparts = 0) 
{
	MsgBox, Error: Cannot create a relative path shortcut to another drive
	Exit
}

;How far is the last same part from working dir
HowFarFromWorkDir := WorkingDirSplitLength - NumOfSameparts

;set the number of steps back (..\) to get to the last same part
Target := ""
Loop, %HowFarFromWorkDir%
{
	Target := Target . "..\"
}

;or if none are needed set it to "current folder" (.\)
if(HowFarFromWorkDir = 0) 
{
	Target := ".\"
}

;How far is the last same part from target dir (-1 because the last part is the filename)
HowFarFromTargetDir := TargetDirSplitLength - NumOfSameparts - 1

;add in the rest of the path after last same part
ArrayCount := TargetDirSplitLength - HowFarFromTargetDir
Loop
{
	if(ArrayCount = TargetDirSplitLength+1)
		break
	else
		Target := Target . TargetDirSplit[ArrayCount] . "\"
	
	ArrayCount += 1
}

;trim trailing \ and add quotes
Target := RTrim(Target, "\")
Target :=  """" . Target . """"

ShortcutName := TargetDirSplit[TargetDirSplit.MaxIndex()] . " - Relative Shortcut.lnk"
HoverMsg := RTrim(ShortcutName, ".lnk")

;create shortcut
Explorer := "%windir%\explorer.exe"
FileCreateShortcut,%Explorer%, %ShortcutName%,,%Target%,%HoverMsg%,%windir%\explorer.exe,,,1