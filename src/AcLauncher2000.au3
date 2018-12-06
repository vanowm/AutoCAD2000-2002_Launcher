#NoTrayIcon
#include <Array.au3>
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=acad_1.ico
#AutoIt3Wrapper_Outfile=AcLauncher2000.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Comment=AutoCAD DWG Launcher
#AutoIt3Wrapper_Res_Description=AutoCAD DWG Launcher
#AutoIt3Wrapper_Res_Fileversion=1.0.0.10
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductVersion=1.0.0
#AutoIt3Wrapper_Res_LegalCopyright=©V@no 2018
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_Res_Field=ProductName|AutoCAD2000 Launcher
#AutoIt3Wrapper_Run_Au3Stripper=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

Global Const $VERSION = "1.0.0.10"
;registry key were we'll temporary store list of files to open in Acad
Global $regKey = "HKCU\Software\Autodesk\AutoCAD"
Global $regName = "AcLauncher2000"
Global $errObj = ObjEvent("AutoIt.Error", "Err")
Global $self = @Compiled ? @ScriptName : @AutoItExe
Global $obj, $oError
Dim $dwg

$selfPID = ProcessExists($self)
If $CmdLine[0] Then
	Local $s = RegRead($regKey, $regName)
	For $i = 1 To $CmdLine[0]
		$s &= ($s = "" ? "" : "|") & '"' & $CmdLine[$i] & '"'
	Next
	;store everything from command line into registry
	RegWrite($regKey, $regName, "REG_SZ", $s)
EndIf

If IsAdmin() Then
	;make sure only one instance of this launcher started
	If $selfPID And $selfPID <> @AutoItPID Then quit()
Else
	If (Not $selfPID Or $selfPID = @AutoItPID) Then
		;request administrative privileges
		ShellExecute($self, "", @ScriptDir, "runas")
	EndIf
	quit(False)
EndIf

$acadPID = ProcessExists("acad.exe")
If Not $acadPID Then
	;acad.exe must be running in order to get ActiveX object
	$acadPID = ShellExecute("acad.exe", StringReplace(RegRead($regKey, $regName), "|", " "))
	RegDelete($regKey, $regName)
EndIf
;loop while acad.exe is running, since we are running as elevated, no more UAC prompts will be displayed
While (1)
	If $acadPID Then
		If IsObj($obj) Then
			$dwg = RegRead($regKey, $regName)
			;are there any new file paths available?
			If $dwg Then
				;delete registry data after first use
				RegDelete($regKey, $regName)
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
			EndIf
		Else
			;wait for the ActiveX object while acad.exe is still initializing
			$obj = ObjGet("", "AutoCAD.Application")
		EndIf
	EndIf
	;acad.exe no longer running? exit then
	If Not $acadPID Or ($acadPID And Not ProcessExists($acadPID)) Then quit()
	Sleep(300)
WEnd

Func Err($e)
	$oError = "err.number is:	" & "0x" & Hex($e.number) & @CRLF & _
			"err.windescript.:	" & $e.windescription & ($e.windescription ? "" : @CRLF) & _
			"err.description:	" & $e.description & @CRLF & _
			"err.source is:	" & $e.source & @CRLF & _
			"err.helpfile is:	" & $e.helpfile & @CRLF & _
			"err.helpcontext is:	" & $e.helpcontext & @CRLF & _
			"err.lastdllerror is:	" & $e.lastdllerror & @CRLF & _
			"err.scriptline is:	" & $e.scriptline & @CRLF & _
			"err.retcode is:	" & "0x" & Hex($e.retcode) & @CRLF
EndFunc   ;==>Err

Func quit($clearReg = True)
	If $clearReg Then RegDelete($regKey, $regName)
	Exit
EndFunc   ;==>quit
