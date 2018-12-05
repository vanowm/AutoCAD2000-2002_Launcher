#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=acad_1.ico
#AutoIt3Wrapper_Outfile=AcLauncher2000.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Comment=AutoCAD DWG Launcher
#AutoIt3Wrapper_Res_Description=AutoCAD DWG Launcher
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductVersion=1.0.0
#AutoIt3Wrapper_Res_LegalCopyright=©V@no 2018
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_Res_Field=ProductName|AutoCAD2000 Launcher
#AutoIt3Wrapper_Run_Au3Stripper=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

Global Const $VERSION = "1.0.0.0"
;registry key were we'll temporary store list of files to open in Acad
Global $reglist = "HKCU\Software\Autodesk\AutoCAD"
Global $obj
ObjEvent("AutoIt.Error", "Err")
Global $self = @Compiled ? @ScriptName : @AutoItExe
Dim $dwg

$pid = ProcessExists($self)
If $CmdLine[0] Then
	Local $s = ""
	For $i = 1 To $CmdLine[0]
		$s &= ($s = "" ? "" : "|") & $CmdLine[$i]
	Next
	;store everything from command line into registry
	RegWrite($reglist, "launcher", "REG_SZ", $s)
EndIf
If Not IsAdmin() Then
	If (Not $pid Or $pid = @AutoItPID) Then
		;request administrative privileges
		ShellExecute($self, $CmdLineRaw, "", "runas")
	EndIf
	quit()
Else
	;make sure only one instance of this launcher started
	If $pid And $pid <> @AutoItPID Then quit()
EndIf

$acadPid = ProcessExists("acad.exe")
If Not $acadPid Then
	;we need acad.exe running in order to get ActiveX object
	$acadPid = ShellExecute("acad.exe")
EndIf
;we'll loop here while acad.exe is running, since we are running with elevated
While (1)
	If $acadPid Then
		If Not IsObj($obj) Then
			;we'll wait for the ActiveX object while acad.exe is still initializing
			$obj = ObjGet("", "AutoCAD.Application")
		Else
			$dwg = RegRead($reglist, "launcher")
			;are there any new file paths available?
			If $dwg Then
				;delete registry data after first use
				RegDelete($reglist, "launcher")
				$dwg = StringSplit($dwg, "|")
				For $i = 1 To $dwg[0]
					;open drawings, one-by-one
					$obj.Documents.open($dwg[$i])
				Next
			EndIf
		EndIf
	EndIf
	;acad.exe no longer running? exit then
	If Not $acadPid Or ($acadPid And Not ProcessExists($acadPid)) Then quit()
	Sleep(300)
WEnd

Func Err()
EndFunc   ;==>MyErrFunc

Func quit()
	Exit
EndFunc   ;==>quit

