#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=AcLauncher2000.ico
#AutoIt3Wrapper_Outfile=AcLauncher2000.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Comment=AutoCAD DWG Launcher
#AutoIt3Wrapper_Res_Description=AutoCAD DWG Launcher
#AutoIt3Wrapper_Res_Fileversion=1.1.0.1
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductVersion=1.1.0
#AutoIt3Wrapper_Res_LegalCopyright=©V@no 2018
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_Res_Field=ProductName|AutoCAD2000 Launcher
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Ignore_Funcs=reg_OnObjectReady
#Au3Stripper_Parameters=/pe /rm /rslm /pe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <DDEML.au3>
#include <DDEMLClient.au3>

Global Const $VERSION = "1.1.0.1"
;registry key were we'll temporary store list of files to open in Acad
Global $regHive = "HKEY_USERS"
Global $regKey = _GetSID() & "\Software\Autodesk\AutoCAD"
Global $regKeyFull = $regHive & "\" & $regKey
Global $regName = "AcLauncher2000"
Global $self = @Compiled ? @ScriptName : @AutoItExe
Global $selfQuery = @Compiled ? "" : '/AutoIt3ExecuteScript "' & @ScriptName & '"'
Global $obj, $notify = False, $firstRun = True, $delay = 300

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

; getting DDE command from registry
Global $drawing = RegRead("HKEY_CLASSES_ROOT\AutoCAD.Drawing\CurVer", "")
Global $dde = RegRead("HKEY_CLASSES_ROOT\" & $drawing & "\shell\print\ddeexec\application", "")
Global $ddeexec = RegRead("HKEY_CLASSES_ROOT\" & $drawing & "\shell\open\ddeexec", "")
If Not $dde Then $dde = "AutoCAD.r15.DDE"
If Not $ddeexec Then $ddeexec = '[open("%1")]'

;loop while acad.exe is running, since we are running as elevated, no more UAC prompts shall be displayed
While (1)
	If $acadPID Then
		;we only need fire registry check if it hasn't been registered or on first launch with acad.exe already running
		If Not $notify Or $firstRun Then reg_OnObjectReady()
		;with registry notificaiton we don't need loop anymore
		If $notify Then ExitLoop
	EndIf
	;acad.exe no longer running? exit then
	If Not $acadPID Or ($acadPID And Not ProcessExists($acadPID)) Then quit()
	Sleep($delay)
WEnd

ProcessWaitClose($acadPID);
quit(True)

Func reg_OnObjectReady($objLatestEvent = False, $objAsyncContext = False)
	;can't go further until AutoCAD object available
	$firstRun = False
	Local $dwg = RegRead($regKeyFull, $regName)
	;are there any new file paths available?
	If Not $dwg Then Return
	;delete registry data after first use
	RegDelete($regKeyFull, $regName)
	$dwg = StringSplit($dwg, "|")
	For $i = 1 To $dwg[0]
		Local $command = StringReplace($ddeexec, "%1", StringRegExpReplace($dwg[$i], '^"|"$', ""))
		;open drawings, one-by-one
		$hData = _DDEMLClient_Execute($dde, "system", $command, $CF_UNICODETEXT)
		Local $er = @error, $ere = @extended
		;bring AutoCAD window to front
		WinActivate(WinGetHandle("[REGEXPCLASS:Afx:400000:8:10005:0]"))
		If Not $hData Then
			MsgBox(16+262144+4096, "Error", "Error: " & $er & @CR & "Extended: " & $ere & @CR & "DDE: " & $dde & @CR & "DDEEXEC: " & $command)
		EndIf
	Next
EndFunc   ;==>registry_OnObjectReadyEndFunc

Func _GetSID($s_User = @UserName, $s_Domain = @LogonDomain)
    Local $o_WMI = ObjGet("winmgmts:\\.\root\cimv2")
    If Not IsObj($o_WMI) Then Return SetError(1, 0, -1)

    Local $s_SID = $o_WMI.Get("Win32_UserAccount.Name='" & $s_User & "',Domain='" & $s_Domain & "'")

    Return $s_SID.SID
EndFunc
Func msg($t)
	ConsoleWrite($t & @cr)
EndFunc

Func quit($clearReg = True)
	If $clearReg Then RegDelete($regKeyFull, $regName)
	Exit
EndFunc   ;==>quit
