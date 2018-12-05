# AutoCAD2000-2002_Launcher
A launcher for AutoCAD 2000/2002 that allows open multiple drawings in a single window on Windows7/10 with UAC enabled or if MS Office not installed


AutoCAD 2000/2002 requires administrative privileges, therefor with UAC enabled it opens separate window each time a new .dwg file is opened from explorer.

This launcher fixes it by adding requested .dwg files via ActiveX interface instead of relying on system

How to use:
copy AcLauncher2000.exe into same directory where acad.exe is
launch .dwg files via right click -> "Open With"
(You can make it permanently associated with the launcher by enabling "Always use this app to open .dwg files" in the dialog)