#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=AcLauncher2000.ico
#AutoIt3Wrapper_Outfile=AcLauncher2000.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Comment=AutoCAD DWG Launcher
#AutoIt3Wrapper_Res_Description=AutoCAD DWG Launcher
#AutoIt3Wrapper_Res_Fileversion=1.0.0.16
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductVersion=1.0.0
#AutoIt3Wrapper_Res_LegalCopyright=©V@no 2018
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_Res_Field=ProductName|AutoCAD2000 Launcher
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Ignore_Funcs=reg_OnObjectReady
#Au3Stripper_Parameters=/pe /rm /rslm /pe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

Global Const $VERSION = "1.0.0.16"
;registry key were we'll temporary store list of files to open in Acad
Global $regHive = "HKEY_USERS"
Global $regKey = _GetSID() & "\Software\Autodesk\AutoCAD"
Global $regKeyFull = $regHive & "\" & $regKey
Global $regName = "AcLauncher2000"
Global $errObj = ObjEvent("AutoIt.Error", "Err")
Global $self = @Compiled ? @ScriptName : @AutoItExe
Global $selfQuery = @Compiled ? "" : '/AutoIt3ExecuteScript "' & @ScriptName & '"'
Global $obj, $oError, $notify = False, $firstRun = True, $delay = 300

$selfPID = ProcessExists($self)
If $CmdLine[0] Then
	Local $s = RegRead($regKeyFull, $regName)
	For $i = 1 To $CmdLine[0]
		$s &= ($s = "" ? "" : "|") & '"' & $CmdLine[$i] & '"'
	Next
	;store everything from command line into registry
	RegWrite($regKeyFull, $regName, "REG_SZ", $s)
EndIf

If IsAdmin() Then
	;make sure only one instance of this launcher started
	If $selfPID And $selfPID <> @AutoItPID Then quit()
Else
	If (Not $selfPID Or $selfPID = @AutoItPID) Then
		;request administrative privileges
		ShellExecute($self, $selfQuery, @ScriptDir, "runas")
	EndIf
	quit(False)
EndIf

$acadPID = ProcessExists("acad.exe")
;~ Credits must go to XxXFaNtA =)
Global $o_WMI, $o_Sink
;registry changes notificaiton
$o_Sink = ObjCreate("WbemScripting.SWbemSink")
ObjEvent($o_Sink , "reg_")
$o_WMI = ObjGet('winmgmts:\\' & @ComputerName & '\root\default')
If Not @error Then
	;registering our registry key changes notification
    $o_WMI.ExecNotificationQueryAsync($o_Sink,	'Select * FROM RegistryValueChangeEvent WHERE Hive="' & $regHive & _
												'" AND KeyPath="' & StringReplace($regKey, "\", "\\") & _
												'" AND ValueName="' & $regName & '"', Default, Default, Default)
	If Not @error Then
		$notify = True
		$delay = 100
	EndIf
EndIf

If Not $acadPID Then
	;acad.exe must be running in order to get ActiveX object
	$acadPID = ShellExecute("acad.exe", StringReplace(RegRead($regKeyFull, $regName), "|", " "))
	$firstRun = False
	RegDelete($regKeyFull, $regName)
EndIf

;loop while acad.exe is running, since we are running as elevated, no more UAC prompts shall be displayed
While (1)
	If $acadPID Then
		If IsObj($obj) Then
			;we only need fire registry check if it hasn't been registered or on first launch with acad.exe already running
			If Not $notify Or $firstRun Then reg_OnObjectReady()
			;with registry notificaiton we don't need loop anymore
			If $notify Then ExitLoop
		Else
			;wait for the ActiveX object while acad.exe is still initializing
			$obj = ObjGet("", "AutoCAD.Application")
		EndIf
	EndIf
	;acad.exe no longer running? exit then
	If Not $acadPID Or ($acadPID And Not ProcessExists($acadPID)) Then quit()
	Sleep($delay)
WEnd

ProcessWaitClose($acadPID);
quit(True)

Func reg_OnObjectReady($objLatestEvent = False, $objAsyncContext = False)
	;can't go further until AutoCAD object available
	If Not IsObj($obj) Then Return
	$firstRun = False
	Local $dwg = RegRead($regKeyFull, $regName)
	;are there any new file paths available?
	If Not $dwg Then Return
	;delete registry data after first use
	RegDelete($regKeyFull, $regName)
	$dwg = StringSplit($dwg, "|")
	For $i = 1 To $dwg[0]
		;open drawings, one-by-one
		Local $o = $obj.Documents.Open($dwg[$i])
		If IsObj($o) Then
			;bring AutoCAD window to front
			WinActivate("[REGEXPCLASS:Afx:400000:8:10005:0]")
		Else
			MsgBox(0, "Error", $oError)
		EndIf
	Next
EndFunc   ;==>registry_OnObjectReadyEndFunc

Func _GetSID($s_User = @UserName, $s_Domain = @LogonDomain)
    Local $o_WMI = ObjGet("winmgmts:\\.\root\cimv2")
    If Not IsObj($o_WMI) Then Return SetError(1, 0, -1)

    Local $s_SID = $o_WMI.Get("Win32_UserAccount.Name='" & $s_User & "',Domain='" & $s_Domain & "'")

    Return $s_SID.SID
EndFunc

Func Err($e)
	$oError =	"err.number is:	" & "0x" & Hex($e.number) & @CRLF & _
				"err.windescript.:	" & $e.windescription & ($e.windescription ? "" : @CRLF) & _
				"err.description:	" & $e.description & @CRLF & _
				"err.source is:	" & $e.source & @CRLF & _
				"err.helpfile is:	" & $e.helpfile & @CRLF & _
				"err.helpcontext is:	" & $e.helpcontext & @CRLF & _
				"err.lastdllerror is:	" & $e.lastdllerror & @CRLF & _
				"err.scriptline is:	" & $e.scriptline & @CRLF & _
				"err.retcode is:	" & "0x" & Hex($e.retcode) & @CRLF
EndFunc   ;==>Err

Func msg($t)
	ConsoleWrite($t & @cr)
EndFunc

Func quit($clearReg = True)
	If $clearReg Then RegDelete($regKeyFull, $regName)
	Exit
EndFunc   ;==>quit
