#Persistent
#NoEnv
;accepted file extensions, separate those using |. E.g. txt|exe|ahk
FileExtensions=zip|rar
WatchDirectory=C:\Users\tristanhaffenden\Downloads

OnMessage(0x4e,"WM_NOTIFY") ;Will make LinkClick and ToolTipClose possible
Loop,Parse,WatchDirectory,|
	WatchDirectory(A_LoopField,1)
SetTimer,Report,50 ;comment this and uncomment below if you like to have a hotkey to get latest changes
;#q::GoTo, Report 
Return

Report:
WatchDirectory("Report")
Return

Report(action,folder,file){ 		;#1:="New File", #2:="Deleted", #3:="Modified", #4:="Renamed From", #5:="Renamed To"
	global FileExtensions, Changes
	static #1:="New File", #2:="Deleted", #3:="Modified"
	If action not in 1,2,3			;report only if new file is created or file is modified
		Return
	path:=folder . (SubStr(folder,0)="" ? "" : "") . file
	If InStr(FileExist(path),"D") 	;return if path is folder
		Return
	SplitPath,path,,,EXT
	Loop,Parse,FileExtensions,|
		found+= (Ext=A_LoopField ? 1 : 0)
	If (!found and FileExtensions)
		Return
	Thread,NoTimers
	Changes++
	File%changes%:=path
	Loop,%path%
		FormatTime,time,%A_LoopFileTimeModified%,HH:mm
	If !time
		FormatTime,time,%A_Now%,HH:mm
	text.= "Path: " . path . "`n`n<a 1" . path . ">Open</a> - <a 2" . path . ">Explore</a> - <a 3" . path . ">Copy</a>`n`n<a>ExitApp</a>"
	ToolTip(Changes,text,"File " . file . " was " . #%action% " at " . time,"O1 C1 L1 P1 I1 xTrayIcon yTrayIcon")
}

WM_NOTIFY(wParam, lParam){
   ToolTip("",lParam,"")
}

ToolTip:
If ErrorLevel=ExitApp
	Goto, ExitApp
action:=SubStr(ErrorLevel,1,1)
path:=SubStr(ErrorLevel,2)
If action=1
	Run % path
else if action=2
	Run,% "explorer.exe /e`, /n`, /select`," . path
else if action=3
	Clipboard:=path
Return

ToolTipClose:
Return

ExitApp:
WatchDirectory()
ExitApp











;Function WatchDirectory()
;
;Parameters
;		WatchFolder			- Specify a valid path to watch for changes in.
;							- can be directory or drive (e.g. c:\ or c:\Temp) 
;							- can be network path e.g. \\192.168.2.101\Shared)
;							- can include last backslash. e.g. C:\Temp\ (will be reported same form)
;
;		WatchSubDirs		- Specify whether to search in subfolders
;
;StopWatching	-	THIS SHOULD BE DONE BEFORE EXITING SCRIPT AT LEAST (OnExit)
;		Call WatchDirectory() without parameters to stop watching all directories
;
;ReportChanges
;		Call WatchDirectory("ReportingFunctionName") to process registered changes.
;		Syntax of ReportingFunctionName(Action,Folder,File)
;
;		Example
/*
		#Persistent
		OnExit,Exit
		WatchDirectory("C:\Windows",1)
		SetTimer,WatchFolder,100
		Return
		WatchFolder:
			WatchDirectory("RegisterChanges")
		Return
		RegisterChanges(action,folder,file){
			static 
			#1:="New File", #2:="Deleted", #3:="Modified", #4:="Renamed From", #5:="Renamed To"
			ToolTip % #%Action% "`n" folder . (SubStr(folder,0)="" ? "" : "") . file
		}	
		Exit:
			WatchDirectory()
		ExitApp
*/

WatchDirectory(WatchFolder="", WatchSubDirs=true)
{
	static
	local hDir, hEvent, r, Action, FileNameLen, pFileName, Restart, CurrentFolder, PointerFNI, _SizeOf_FNI_=65536
	nReadLen := 0
	If !(WatchFolder){
		Gosub, StopWatchingDirectories
	} else if IsFunc(WatchFolder) {
		r := DllCall("MsgWaitForMultipleObjectsEx", UInt, DirIdx, UInt, &DirEvents, UInt, -1, UInt, 0x4FF, UInt, 0x6) ;Timeout=-1
		if !(r >= 0 && r < DirIdx)
			Return
		r += 1
		CurrentFolder := Dir%r%Path
		PointerFNI := &Dir%r%FNI
		DllCall( "GetOverlappedResult", UInt, hDir, UInt, &Dir%r%Overlapped, UIntP, nReadLen, Int, true )
		Loop {
			pNext   	:= NumGet( PointerFNI + 0  )
			Action      := NumGet( PointerFNI + 4  )
			FileNameLen := NumGet( PointerFNI + 8  )
			pFileName :=       ( PointerFNI + 12 )
			If (Action < 0x6){
				VarSetCapacity( FileNameANSI, FileNameLen )
				DllCall( "WideCharToMultiByte",UInt,0,UInt,0,UInt,pFileName,UInt,FileNameLen,Str,FileNameANSI,UInt,FileNameLen,UInt,0,UInt,0)
				%WatchFolder%(Action,CurrentFolder,SubStr( FileNameANSI, 1, FileNameLen/2 ))
			}
			If (!pNext or pNext = 4129024)
				Break
			Else
				PointerFNI := (PointerFNI + pNext)
		}
		DllCall( "ResetEvent", UInt,NumGet( Dir%r%Overlapped, 16 ) )
		Gosub, ReadDirectoryChanges
		return r
	} else {
		Loop % (DirIdx) {
			If InStr(WatchFolder, Dir%A_Index%Path){
				If (Dir%A_Index%Subdirs)
					Return
			} else if InStr(Dir%A_Index%Path, WatchFolder) {
				If (WatchSubDirs){
					DllCall( "CloseHandle", UInt,Dir%A_Index% )
					DllCall( "CloseHandle", UInt,NumGet(Dir%A_Index%Overlapped, 16) )
					Restart := DirIdx, DirIdx := A_Index
				}
			}
		}
		If !Restart
			DirIdx += 1
		r:=DirIdx
		hDir := DllCall( "CreateFile"
					 , Str  , WatchFolder
					 , UInt , ( FILE_LIST_DIRECTORY := 0x1 )
					 , UInt , ( FILE_SHARE_READ     := 0x1 )
							| ( FILE_SHARE_WRITE    := 0x2 )
							| ( FILE_SHARE_DELETE   := 0x4 )
					 , UInt , 0
					 , UInt , ( OPEN_EXISTING := 0x3 )
					 , UInt , ( FILE_FLAG_BACKUP_SEMANTICS := 0x2000000  )
							| ( FILE_FLAG_OVERLAPPED       := 0x40000000 )
					 , UInt , 0 )
		Dir%r%         := hDir
		Dir%r%Path     := WatchFolder
		Dir%r%Subdirs  := WatchSubDirs
		VarSetCapacity( Dir%r%FNI, _SizeOf_FNI_ )
		VarSetCapacity( Dir%r%Overlapped, 20, 0 )
		DllCall( "CloseHandle", UInt,hEvent )
		hEvent := DllCall( "CreateEvent", UInt,0, Int,true, Int,false, UInt,0 )
		NumPut( hEvent, Dir%r%Overlapped, 16 )
		if ( VarSetCapacity(DirEvents) < DirIdx*4 and VarSetCapacity(DirEvents, DirIdx*4 + 60))
			Loop %DirIdx%
			{
				If (SubStr(Dir%A_Index%Path,1,1)!="-"){
					action++
					NumPut( NumGet( Dir%action%Overlapped, 16 ), DirEvents, action*4-4 )
				}
			}
		NumPut( hEvent, DirEvents, DirIdx*4-4)
		Gosub, ReadDirectoryChanges
		If Restart
			DirIdx = %Restart%
	}
	Return
	StopWatchingDirectories:
		Loop % (DirIdx) {
			DllCall( "CloseHandle", UInt,Dir%A_Index% )
			DllCall( "CloseHandle", UInt,NumGet(Dir%A_Index%Overlapped, 16) )
			Dir%A_Index%=
			Dir%A_Index%Path=
			Dir%A_Index%Subdirs=
			Dir%A_Index%FNI=
			DllCall( "CloseHandle", UInt, NumGet(Dir%A_Index%Overlapped,16) )
			VarSetCapacity(Dir%A_Index%Overlapped,0)
		}
		DirIdx=
		VarSetCapacity(DirEvents,0)
	Return
	ReadDirectoryChanges:
		DllCall( "ReadDirectoryChangesW"
			, UInt , Dir%r%
			, UInt , &Dir%r%FNI
			, UInt , _SizeOf_FNI_
			, UInt , Dir%r%SubDirs
			, UInt , ( FILE_NOTIFY_CHANGE_FILE_NAME   := 0x1   )
				   | ( FILE_NOTIFY_CHANGE_DIR_NAME    := 0x2   )
				   | ( FILE_NOTIFY_CHANGE_ATTRIBUTES  := 0x4   )
				   | ( FILE_NOTIFY_CHANGE_SIZE        := 0x8   )
				   | ( FILE_NOTIFY_CHANGE_LAST_WRITE  := 0x10  )
				   | ( FILE_NOTIFY_CHANGE_CREATION    := 0x40  )
				   | ( FILE_NOTIFY_CHANGE_SECURITY    := 0x100 )
			, UInt , 0
			, UInt , &Dir%r%Overlapped
			, UInt , 0  )
	Return
}

;ToolTip(1,"hallo","test","i1 q1 e1.1.1.5 C1 O1 XTrayIcon Y20")
/*
ToolTip() by HotKeyIt http://www.autohotkey.com/forum/viewtopic.php?t=40165

Syntax: ToolTip(Number,Text,Title,Options)

Return Value: ToolTip returns hWnd of the ToolTip

|-------------------------------------------------------------------------------------------------------------------|
|			Options can include any of following parameters separated by space
|-------------------------------------------------------------------------------------------------------------------|
| Option		|	   Meaning
|-------------------------------------------------------------------------------------------------------------------|
| A			|	Aim ConrolId or ClassNN (Button1, Edit2, ListBox1, SysListView321...)
|			|	- using this, ToolTip will be shown when you point mouse on a control
|			|	- D (delay) can be used to change how long ToolTip is shown
|			|	- W (wait) can wait for specified seconds before ToolTip will be shown
|			|	- Some controls like Static require a subroutine to have a ToolTip!!!
|-------------------------------------------------------------------------------------------------------------------|
| B and F		|	Specify here the color for ToolTip in 6-digit hexadecimal RGB code
|			|	- B = Background color, F = Foreground color (text color)
|			|	- this can be 0x00FF00 or 00FF00 or Blue, Lime, Black, White...
|-------------------------------------------------------------------------------------------------------------------|
| C			|	Close button for ToolTip/BalloonTip. See ToolTip actions how to use it
|-------------------------------------------------------------------------------------------------------------------|
| D			|	Delay. This option will determine how long ToolTip should be shown.30 sec. is maximum
|			|	- this option is also available when assigning the ToolTip to a control.
|-------------------------------------------------------------------------------------------------------------------|
| E			|	Edges for ToolTip, Use this to set margin of ToolTip window (space between text and border)
|			|	- Supply Etop.left.bottom.right in pixels, for example: E10.0.10.5
|-------------------------------------------------------------------------------------------------------------------|
| G			|	Execute one or more internal Labels of ToolTip function only.
|			|	For example:
|			|	- Track the position only, use ToolTip(1,"","","Xcaret Ycaret gTTM_TRACKPOSITION")
|			|		- When X+Y are empty (= display near mouse position) you can use TTM_UPDATE
|			|	- Update text only, use ToolTip(1,"text","","G1"). Note specify L1 if links are used.
|			|	- Update title only, use ToolTip(1,"","Title","G1")
|			|	- Hide ToolTip, use ToolTip(1,"","","gTTM_POP")
|			|		- To show ToolTip again use ToolTip(1,"","","gTTM_TRACKPOSITION.TTM_TRACKACTIVATE")
|			|	- Update background color + text color, specify . between gLabels to execute several:
|			|		- ToolTip(1,"","","BBlue FWhite gTTM_SETTIPBKCOLOR.TTM_SETTIPTEXTCOLOR")
|			|	- Following labels can be used: TTM_SETTITLEA + TTM_SETTITLEW (title+I), TTM_POPUP, TTM_POP
|			|	  TTM_SETTIPBKCOLOR (B), TTM_SETTIPTEXTCOLOR (F), TTM_TRACKPOSITION (N+X+Y),
|			|	  TTM_SETMAXTIPWIDTH (R), TTM_SETMARGIN (E), TT_SETTOOLINFO (text+A+P+N+X+Y+S+L)
|			|	  TTM_SETWINDOWTHEME (Q)
|-------------------------------------------------------------------------------------------------------------------|
| H			|	Hide ToolTip after a link is clicked.See L option
|-------------------------------------------------------------------------------------------------------------------|
| I			|	Icon 1-3, e.g. I1. If this option is missing no Icon will be used (same as I0)
|			|	- 1 = Info, 2 = Warning, 3 = Error, > 3 is meant to be a hIcon (handle to an Icon)
|			|	Use Included MI_ExtractIcon and GetAssociatedIcon functions to get hIcon
|-------------------------------------------------------------------------------------------------------------------|
| L			|	Links for ToolTips. See ToolTip actions how Links for ToolTip work.
|-------------------------------------------------------------------------------------------------------------------|
| M			|	Mouse click-trough. So a click will be forwarded to the window underneath ToolTip
|-------------------------------------------------------------------------------------------------------------------|
| N			|	Do NOT activate ToolTip (N1), To activate (show) call ToolTip(1,"","","gTTM_TRACKACTIVATE")
|-------------------------------------------------------------------------------------------------------------------|
| O			|	Oval ToolTip (BalloonTip). Specify O1 to use a BalloonTip instead of ToolTip.
|-------------------------------------------------------------------------------------------------------------------|
| P			|	Parent window hWnd or GUI number. This will assign a ToolTip to a window.
|			|	- Reqiered to assign ToolTip to controls and actions.
|-------------------------------------------------------------------------------------------------------------------|
| Q			|	Quench Style/Theme. Use this to disable Theme of ToolTip.
|			|	Using this option you can have for example colored ToolTips in Vista.
|-------------------------------------------------------------------------------------------------------------------|
| R			|	Restrict width. This will restrict the width of the ToolTip.
|			|	So if Text is to long it will be shown in several lines
|-------------------------------------------------------------------------------------------------------------------|
| S			|	Show at coordinates regardless of position. Specify S1 to use that feature
|			|	- normally it is fed automaticaly to show on screen
|-------------------------------------------------------------------------------------------------------------------|
| T			|	Transparency. This option will apply Transparency to a ToolTip.
|			|	- this option is not available to ToolTips assigned to a control.
|-------------------------------------------------------------------------------------------------------------------|
| V			|	Visible: even when the parent window is not active, a control-ToolTip will be shown
|-------------------------------------------------------------------------------------------------------------------|
| W			|	Wait time in seconds (max 30) before ToolTip pops up when pointing on one of controls.
|-------------------------------------------------------------------------------------------------------------------|
| X and Y		|	Coordinates where ToolTip should be displayed, e.g. X100 Y200
|			|	- leave empty to display ToolTip near mouse
|			|	- you can specify Xcaret Ycaret to display at caret coordinates
|-------------------------------------------------------------------------------------------------------------------|
|
| 			To destroy a ToolTip use ToolTip(Number), to destroy all ToolTip()
|
|-------------------------------------------------------------------------------------------------------------------|
|					ToolTip Actions (NOTE, OPTION P MUST BE PRESENT TO USE THAT FEATURE)
|-------------------------------------------------------------------------------------------------------------------|
|		Assigning an action to a ToolTip to works using OnMessage(0x4e,"Function") - WM_NOTIFY
|		Parameter/option P must be present so ToolTip will forward messages to script
|		All you need to do inside this OnMessage function is to include:
|			- If wParam=0
|				ToolTip("",lParam[,Label])
|
|		Additionally you need to have one or more of following labels in your script
|		- ToolTip: when clicking a link
|		- ToolTipClose: when closing ToolTip
|			- You can also have a diferent label for one or all ToolTips
|			- Therefore enter the number of ToolTip in front of the label
|				- e.g. 99ToolTip: or 1ToolTipClose:
|
|		- Those labels names can be customized as well
|			- e.g. ToolTip("",lParam,"MyTip") will use MyTip: and MyTipClose:
|			- you can enter the number of ToolTip in front of that label as well.
|
|		- Links have following syntax:
|			- <a>Link</a> or <a link>LinkName</a>
|			- When a Link is clicked, ToolTip() will jump to the label
|				- Variable ErrorLevel will contain clicked link 
|
|			- So when only LinkName is given, e.g. <a>AutoHotkey</a> Errorlevel will be AutoHotkey
|			- When using Link is given as well, e.g. <a http://www.autohotkey.com>AutoHotkey</a>
|				- Errorlevel will be set to http://www.autohotkey.com
|
|-------------------------------------------------------------------------------------------------------------------|
|		Please note some options like Close Button and Links will require Win2000++ (+version 6.0 of comctl32.dll)
|  		AutoHotKey Version 1.0.48++ is required due to "assume static mode"
|  		If you use 1 ToolTip for several controls, the only difference between those can be the text. 
|  			- Rest, like Title, color and so on, will be valid globally 
|-------------------------------------------------------------------------------------------------------------------|
|		Example for LinkClick and ToolTip close!
|-------------------------------------------------------------------------------------------------------------------|
			OnMessage(0x4e,"WM_NOTIFY") ;Will make LinkClick and ToolTipClose possible
			WM_NOTIFY(wParam, lParam, msg, hWnd)
			{
			   ToolTip("",lParam,"")
			}
			Sleep, 10
			ToolTip(1,"<a>Click</a>`n<a>Onother one</a>`n"
			. "<a This link is different`nit uses different text>Different</a>`n"
			. "<a>ExitApp</a>","ClickMe","L1 P99 C1")
			Return
			ToolTip:
			link:=ErrorLevel
			SetTimer, MsgBox, -10
			Return

			ToolTipClose:
			ExitApp

			MsgBox:
			If Link=ExitApp
				ExitApp
			MsgBox % Link
			Return
|-------------------------------------------------------------------------------------------------------------------|

ToolTip() von HotKeyIt http://de.autohotkey.com/forum/viewtopic.php?t=4556

Syntax: ToolTip([Nummer,Text,Titel,Optionen])

Return Wert: ToolTip gibt die hWnd des ToolTips zurück

|-------------------------------------------------------------------------------------------------------------------|
|			Folgende Optionen sind möglich
|-------------------------------------------------------------------------------------------------------------------|
| Option		|	   Bedeutung                                                                          |
|-------------------------------------------------------------------------------------------------------------------|
| A			|	Aim ConrolId or ClassNN (Button1, Edit2, ListBox1, SysListView321...)
|			|	- mit dieser Option wird ToolTip angezeigt wenn Maus auf ein Control zeigt
|			|	- D (delay/Verzögerung) kann benutzt werden um die Anzeigedauer zu verlängern
|			|	- W (Wartezeit) kann benutzt werden um die Anzeige des ToolTips zu verzögern
|			|	- Manche Controls wie Static brauchen ein gLabel um ToolTips für Controls anzuzeigen
|-------------------------------------------------------------------------------------------------------------------|
| B and F		|	Gebe hier die Farbe für ToolTip in 6-digit hexadecimal RGB Code
|			|	- B = Hintergrund Farbe, F = Vordergrund Farbe (Text Farbe)
|			|	- das kann in folgender Form erfolgen: 0x00FF00 or 00FF00 or Blue, Green, Lime...
|-------------------------------------------------------------------------------------------------------------------|
| C			|	Button Schließen for ToolTip/BalloonTip. Näheres findest du bei ToolTip Aktionen
|-------------------------------------------------------------------------------------------------------------------|
| D			|	Verzögerung, gibt an wie lange ToolTip angezeigt wird. Maximum ist 30 Sek
|			|	- diese option gilt auch für ToolTips für Controls
|-------------------------------------------------------------------------------------------------------------------|
| E			|	Ecken vom ToolTip. Diese Option ändert den Abstand zwischen dem Rand und Text.
|			|	- Gebe an Eoben.links.unten.rechts in Pixel, zum Beispiel: E10.0.10.5
|-------------------------------------------------------------------------------------------------------------------|
| G			|	Führe nur das entsprechende interne Label von ToolTip() aus.
|			|	Zum Beispiel:
|			|	- Nur die Position aktualisieren, ToolTip(1,"","","xCaret Ycaret gTTM_TRACKPOSITION")
|			|		- Wenn X+Y leer sind (= Zeige bei Maus position), reicht TTM_UPDATE.
|			|	- Nur Text aktualisieren, ToolTip(1,"text","","G1"). Gebe L1 ein, wenn Text mit Links.
|			|	- Nur Titel aktualisieren, ToolTip(1,"","Title","G1")
|			|	- ToolTip Verstecken, ToolTip(1,"","","gTTM_POP")
|			|		- ToolTip wieder anzeigen ToolTip(1,"","","gTTM_TRACKPOSITION.TTM_TRACKACTIVATE")
|			|	- Nur Hintergrund und Text Farbe aktualisieren. Gebe . zwischen den gLabels ein:
|			|		- ToolTip(1,"","","BBlue FWhite gTTM_SETTIPBKCOLOR.TTM_SETTIPTEXTCOLOR")
|			|	- Folgende interne Labels können benutzt werden: TTM_SETTITLEA + TTM_SETTITLEW (Titel + I)
|			|	  TTM_SETTIPBKCOLOR (B), TTM_SETTIPTEXTCOLOR (F), TTM_TRACKPOSITION (N+X+Y), TTM_POPUP, TTM_POP,
|			|	  TTM_SETMAXTIPWIDTH (R), TTM_SETMARGIN (E), TT_SETTOOLINFO (text+A+P+N+X+Y+S+L)
|-------------------------------------------------------------------------------------------------------------------|
| H			|	Verstecke ToolTip wenn ein Link aktiviert wird
|-------------------------------------------------------------------------------------------------------------------|
| I			|	Icon 1-3, z.B. I1. Ohne diese Option wird kein Icon angezeigt, ebenso wie I0
|			|	- 1 = Info, 2 = Warnung, 3 = Fehler, > 3 ist ein hIcon (Handle zu einem Icon)
|-------------------------------------------------------------------------------------------------------------------|
| L			|	Links für ToolTips. Näheres findest du bei ToolTip Aktionen
|-------------------------------------------------------------------------------------------------------------------|
| M			|	Maus click-durch. Ein Mausklick wird an das darunter liegende Fenster weitergegeben.
|-------------------------------------------------------------------------------------------------------------------|
| N			|	N1 = erstellen ohne zu aktivieren, ToolTip(1,"","","gTTM_TRACKACTIVATE") aktiviert!
|-------------------------------------------------------------------------------------------------------------------|
| O			|	Ovaler ToolTip (BalloonTip). Gebe O1 an um diesen zu aktivieren.
|-------------------------------------------------------------------------------------------------------------------|
| P			|	Parent Fenster hWnd oder GUI Nummer. Das ordnet ein ToolTip einem Fenster zu.
|			|	- Wird benötigt um ToolTip einem Control zuzuweisen und um Aktionen durchzuführen.
|-------------------------------------------------------------------------------------------------------------------|
| Q			|	Deaktiviere Style/Theme (Q1). Windows Classic Theme wird aktiviert.
|			|	Mit dieser Option kann man z.B. normale ToolTips in Vista aktivieren und Option Farben nutzen.
|-------------------------------------------------------------------------------------------------------------------|
| R			|	Beschränke die Breite des ToolTips. Zeilenumbruch für Text wird automatisch eingefügt
|-------------------------------------------------------------------------------------------------------------------|
| S			|	Zeige Tooltip an den angegebenen Koordinaten.
|			|	- normalerweise wird ToolTip automatisch angepasst so dass er immer sichtbar ist
|-------------------------------------------------------------------------------------------------------------------|
| T			|	Transparenz.
|			|	- diese Option gilt nicht in Verbindung mit Controls.
|-------------------------------------------------------------------------------------------------------------------|
| V			|	Immer Sichtbar: sogar wenn das Hauptfenster inaktiv ist wird Control-ToolTip angezeigt
|-------------------------------------------------------------------------------------------------------------------|
| W			|	Wartezeit befor ein ToolTip beim control angezeigt wird. Max. 30 sec.
|-------------------------------------------------------------------------------------------------------------------|
| X and Y		|	Koordinaten wo ein ToolTip angezeigt werden soll. z.B. X100 Y200
|			|	- lass diese weg um ToolTip neben Maus anzuzeigen
|			|	- man kann auch Xcaret Ycaret benutzen um bei Caret Position anzuzeigen
|-------------------------------------------------------------------------------------------------------------------|
|-------------------------------------------------------------------------------------------------------------------|
|					ToolTip Aktionen (MERKE, OPTION P MUSS ANGEGEBEN WERDEN)
|-------------------------------------------------------------------------------------------------------------------|
|-------------------------------------------------------------------------------------------------------------------|
|		Eine Aktion wird über OnMessage(0x4e,"Function") - WM_NOTIFY möglich.
|		Parameter P muss angegeben werden um die Messages an das Script weiterzuleiten.
|		Alles was in der OnMessage Funktion angegeben werden muss ist:
|			- If wParam=0
|				ToolTip("",lParam[,Label])
|
|		Zusätzlich können folgende Labels deklariert und benutzt werden.
|		- ToolTip: beim Klick auf ein Link
|		- ToolTipClose: bei schließen des ToolTips
|			- Man kann auch verschiedene Labels für jeden ToolTip habne
|			- Hierfür einfach die Nummmer vor dem Label deklarieren
|				- z.B. 99ToolTip: oder 1ToolTipClose:
|
|		- Namen der Labels können auch angepasst werden
|			- z.B. ToolTip("",lParam,"MyTip") wird zum Label MyTip: and MyTipClose: springen.
|			- man kann auch bei diesen Labels die Nummer davor setzen.
|
|		- Links haben folgenden Syntax:
|			- <a>Link</a> oderr <a link>LinkName</a>
|			- wenn ein Klick auf ein Link erfolgt, springt ToolTip() zu dem Label
|				- Variable ErrorLevel wird den Link beinhalten
|
|			- Wenn nur ein LinkName angegeben wird, z.B. <a>AutoHotkey</a> Errorlevel wird AutoHotkey
|			- Wenn auch ein Link angegeben wird, z.B. <a http://www.autohotkey.com>AutoHotkey</a>
|				- Errorlevel wird http://www.autohotkey.com
|
|		!!! DO NOT PERFORM ANY ACTION TO THE CURRENT TOOLTIP NUMBER, OnMessage must finish first !!!
|-------------------------------------------------------------------------------------------------------------------|
|		Merke einige Options wie Close Button benötigen Win2000++ (version 6.0 of comctl32.dll)
|  		AutoHotKey Version 1.0.48++ wird benötigt wegen "assume static mode"
|  		Wenn du einen ToolTip für viele Controls benutzt kann nur der Text unterschiedlich sein
|  			- Der Rest, wie Titel, Farbe and so weiter, gilt global für diesen ToolTip
|-------------------------------------------------------------------------------------------------------------------|
|		Beispiel für LinkClick and ToolTip close!
|-------------------------------------------------------------------------------------------------------------------|
			OnMessage(0x4e,"WM_NOTIFY") ;Macht LinkClick und ToolTipClose möglich
			WM_NOTIFY(wParam, lParam, msg, hWnd)
			{
			   ToolTip("",lParam,"")
			}
			Sleep, 10
			ToolTip(1,"<a>Klick</a>`t<a>Weiterer Click</a>`n"
			. "<a Dieser link hat anderen Text>Anders</a>`n"
			. "<a>Programm schließen</a>","ClickMich","L1 P99 C1")
			Return
			ToolTip:
			link:=ErrorLevel
			SetTimer, MsgBox, -10
			Return

			ToolTipClose:
			ExitApp

			MsgBox:
			If Link=Programm schließen
				ExitApp
			MsgBox % Link
			Return
|-------------------------------------------------------------------------------------------------------------------|
*/

ToolTip(ID="", text="", title="",options=""){
	;____  Assume Static Mode for internal variables and structures  ____
	
	static
	;________________________  ToolTip Messages  ________________________
	
	static TTM_POPUP:=0x422,   		TTM_ADDTOOL:=0x404,     	TTM_UPDATETIPTEXT:=0x40c
	,TTM_POP:=0x41C,     		TTM_DELTOOL:=0x405,     	TTM_GETBUBBLESIZE:=0x41e
	,TTM_UPDATE:=0x41D,  		TTM_SETTOOLINFO:=0x409,		TTN_FIRST:=0xfffffdf8
	,TTM_TRACKPOSITION:=0x412, 	TTM_SETTIPBKCOLOR:=0x413,	TTM_SETTIPTEXTCOLOR:=0x414
	,TTM_SETTITLEA:=0x420,		TTM_SETTITLEW:=0x421,		TTM_SETMARGIN:=0x41a
	,TTM_SETWINDOWTHEME:=0x200b,	TTM_SETMAXTIPWIDTH:=0x418
	
	;_______________Remote Buffer Messages for TrayIcon pos______________
	
	;MEM_COMMIT:=0x1000, 		PAGE_READWRITE:=4, 			MEM_RELEASE:=0x8000
	
	;________________________  ToolTip colors  ________________________
	
	,Black:=0x000000,    Green:=0x008000,		Silver:=0xC0C0C0
	,Lime:=0x00FF00,		Gray:=0x808080,    		Olive:=0x808000
	,White:=0xFFFFFF,    Yellow:=0xFFFF00,		Maroon:=0x800000
    ,Navy:=0x000080,		Red:=0xFF0000,    		Blue:=0x0000FF
	,Purple:=0x800080,   Teal:=0x008080,			Fuchsia:=0xFF00FF
    ,Aqua:=0x00FFFF

	;________________________  Local variables for options ________________________
	
	local option,a,b,c,d,e,f,g,h,i,k,l,m,n,o,p,q,r,s,t,v,w,x,y,xc,yc,xw,yw,update,RECT

	If ((#_DetectHiddenWindows:=A_DetectHiddenWindows)="Off")
		DetectHiddenWindows, On
	
	;____________________________  Delete all ToolTips or return link _____________

	If !ID
	{
		If text
			If text is Xdigit
				GoTo, TTN_LINKCLICK
		Loop, Parse, hWndArray, % Chr(2) ;Destroy all ToolTip Windows
		{
			If WinExist("ahk_id " . A_LoopField)
				DllCall("DestroyWindow","Uint",A_LoopField)
			hWndArray%A_LoopField%=
		}
		hWndArray=
		Loop, Parse, idArray, % Chr(2) ;Destroy all ToolTip Structures
		{
			TT_ID:=A_LoopField
			If TT_ALL_%TT_ID%
				Gosub, TT_DESTROY
		}
		idArray=
		DetectHiddenWindows,%#_DetectHiddenWindows%
		Return
	}
	
	TT_ID:=ID
	TT_HWND:=TT_HWND_%TT_ID%
	
	;___________________  Load Options Variables and Structures ___________________
	
	If (options){
		Loop,Parse,options,%A_Space%
			If (option:= SubStr(A_LoopField,1,1))
				%option%:= SubStr(A_LoopField,2)
	}
	If (G){
		If (Title!=""){
			Gosub, TTM_SETTITLEA
			Gosub, TTM_UPDATE
		}
		If (Text!=""){
			If (InStr(text,"<a") and L){
				TOOLTEXT_%TT_ID%:=text
				text:=RegExReplace(text,"<a\K[^<]*?>",">")
			} else
				TOOLTEXT_%TT_ID%:=
			NumPut(&text,TOOLINFO_%TT_ID%,36)
			Gosub, TTM_UPDATETIPTEXT
		}
		Loop, Parse,G,.
			If IsLabel(A_LoopField)
				Gosub, %A_LoopField%
		DetectHiddenWindows,%#_DetectHiddenWindows%
		Return
	}
	;__________________________  Save TOOLINFO Structures _________________________
	
	If P {
		If (p<100 and !WinExist("ahk_id " p)){
			Gui,%p%:+LastFound
			P:=WinExist()
		}
		If !InStr(TT_ALL_%TT_ID%,Chr(2) . Abs(P) . Chr(2))
			TT_ALL_%TT_ID% .= Chr(2) . Abs(P) . Chr(2)
	} 
	If !InStr(TT_ALL_%TT_ID%,Chr(2) . ID . Chr(2))
		TT_ALL_%TT_ID% .= Chr(2) . ID . Chr(2)
	If H
		TT_HIDE_%TT_ID%:=1
	;__________________________  Create ToolTip Window  __________________________
	
	If (!TT_HWND and text)
	{
		TT_HWND := DllCall("CreateWindowEx", "Uint", 0x8, "str", "tooltips_class32", "str", "", "Uint", 0x02 + (v ? 0x1 : 0) + (l ? 0x100 : 0) + (C ? 0x80 : 0)+(O ? 0x40 : 0), "int", 0x80000000, "int", 0x80000000, "int", 0x80000000, "int", 0x80000000, "Uint", P ? P : 0, "Uint", 0, "Uint", 0, "Uint", 0)
		TT_HWND_%TT_ID%:=TT_HWND
		hWndArray.=(hWndArray ? Chr(2) : "") . TT_HWND
		idArray.=(idArray ? Chr(2) : "") . TT_ID
		Gosub, TTM_SETMAXTIPWIDTH
		DllCall("SendMessage", "Uint", TT_HWND, "Uint", 0x403, "Uint", 2, "Uint", (D ? D*1000 : -1)) ;TTDT_AUTOPOP
		DllCall("SendMessage", "Uint", TT_HWND, "Uint", 0x403, "Uint", 3, "Uint", (W ? W*1000 : -1)) ;TTDT_INITIAL
		DllCall("SendMessage", "Uint", TT_HWND, "Uint", 0x403, "Uint", 1, "Uint", (W ? W*1000 : -1)) ;TTDT_RESHOW
	} else if (!text and !options){
		DllCall("DestroyWindow","Uint",TT_HWND)
		TT_HWND_%TT_ID%=
		Gosub, TT_DESTROY
		DetectHiddenWindows,%#_DetectHiddenWindows%
		Return
	}
	
	;______________________  Create TOOLINFO Structure  ______________________
	
	Gosub, TT_SETTOOLINFO

	If (Q!="")
		Gosub, TTM_SETWINDOWTHEME
	If (E!="")
		Gosub, TTM_SETMARGIN
	If (F!="")
		Gosub, TTM_SETTIPTEXTCOLOR
	If (B!="")
		Gosub, TTM_SETTIPBKCOLOR
	If (title!="")
		Gosub, TTM_SETTITLEA
	
	If (!A){
		Gosub, TTM_UPDATETIPTEXT
		Gosub, TTM_UPDATE
		If D {
			A_Timer := A_TickCount, D *= 1000
			Gosub, TTM_TRACKPOSITION
			Gosub, TTM_TRACKACTIVATE
			Loop
			{
				Gosub, TTM_TRACKPOSITION
				If (A_TickCount - A_Timer > D)
					Break
			}
			Gosub, TT_DESTROY
			DllCall("DestroyWindow","Uint",TT_HWND)
			TT_HWND_%TT_ID%=
		} else {
			Gosub, TTM_TRACKPOSITION
			Gosub, TTM_TRACKACTIVATE
			If T
				WinSet,Transparent,%T%,ahk_id %TT_HWND%
			If M
				WinSet,ExStyle,^0x20,ahk_id %TT_HWND%
		}
	}

	;________  Restore DetectHiddenWindows and return HWND of ToolTip  ________

	DetectHiddenWindows, %#_DetectHiddenWindows%
	Return TT_HWND
	
	;________________________  Internal Labels  ________________________
	
	TTM_POP: 	;Hide ToolTip
	TTM_POPUP: 	;Causes the ToolTip to display at the coordinates of the last mouse message.
	TTM_UPDATE: ;Forces the current tool to be redrawn.
		DllCall("SendMessage", "Uint", TT_HWND, "Uint", %A_ThisLabel%, "Uint", 0, "Uint", 0)
	Return
	TTM_TRACKACTIVATE: ;Activates or deactivates a tracking ToolTip.
	DllCall("SendMessage", "Uint", TT_HWND, "Uint", 0x411, "Uint", (N ? 0 : 1), "Uint", &TOOLINFO_%ID%)
	Return
	TTM_UPDATETIPTEXT:
	TTM_GETBUBBLESIZE:
	TTM_ADDTOOL:
	TTM_DELTOOL:
	TTM_SETTOOLINFO:
		DllCall("SendMessage", "Uint", TT_HWND, "Uint", %A_ThisLabel%, "Uint", 0, "Uint", &TOOLINFO_%ID%)
	Return
	TTM_SETTITLEA:
	TTM_SETTITLEW:
		title := (StrLen(title) < 96) ? title : ("…" . SubStr(title, -97))
		DllCall("SendMessage", "Uint", TT_HWND, "Uint", %A_ThisLabel%, "Uint", I, "Uint", &Title)
	Return
	TTM_SETWINDOWTHEME:
		If Q
			DllCall("uxtheme\SetWindowTheme", "Uint", TT_HWND, "Uint", 0, "UintP", 0)
		else
			DllCall("SendMessage", "Uint", TT_HWND, "Uint", %A_ThisLabel%, "Uint", 0, "Uint", &Q)
	Return
	TTM_SETMAXTIPWIDTH:
		DllCall("SendMessage", "Uint", TT_HWND, "Uint", %A_ThisLabel%, "Uint", 0, "Uint", R ? R : A_ScreenWidth)
	Return
	TTM_TRACKPOSITION:
		VarSetCapacity(xc, 20, 0), xc := Chr(20)
		DllCall("GetCursorInfo", "Uint", &xc)
		yc := NumGet(xc,16), xc := NumGet(xc,12)
		xc+=15,yc+=15
		If (x="caret" or y="caret"){
			WinGetPos,xw,yw,,,A
			If x=caret
			{
				SysGet,xl,76
				SysGet,xr,78
				xc:=xw+A_CaretX +1
				xc:=(xl>xc ? xl : (xr<xc ? xr : xc))
			}
			If (y="caret"){
				SysGet,yl,77
				SysGet,yr,79
				yc:=yw+A_CaretY+15
				yc:=(yl>yc ? yl : (yr<yc ? yr : yc))
			}
		} else if (x="TrayIcon" or y="TrayIcon"){
			Process, Exist
			PID:=ErrorLevel
			hWndTray:=WinExist("ahk_class Shell_TrayWnd")
			ControlGet,hWndToolBar,Hwnd,,ToolbarWindow321,ahk_id %hWndTray%
			RemoteBuf_Open(TrayH,hWndToolBar,20)
			DataH:=NumGet(TrayH,0)
			SendMessage, 0x418,0,0,,ahk_id %hWndToolBar%
			Loop % ErrorLevel
			{
				SendMessage,0x417,A_Index-1,RemoteBuf_Get(TrayH),,ahk_id %hWndToolBar%
				RemoteBuf_Read(TrayH,lpData,20)
				VarSetCapacity(dwExtraData,8)
				pwData:=NumGet(lpData,12)
				DllCall( "ReadProcessMemory", "uint", DataH, "uint", pwData, "uint", &dwExtraData, "uint", 8, "uint", 0 )
				BWID:=NumGet(dwExtraData,0)
				WinGet,BWPID,PID, ahk_id %BWID%
				If (BWPID!=PID)
					continue
				SendMessage, 0x41d,A_Index-1,RemoteBuf_Get(TrayH),,ahk_id %hWndToolBar%
				RemoteBuf_Read(TrayH,rcPosition,20)
				If (NumGet(lpData,8)>7){
					ControlGetPos,xc,yc,xw,yw,Button2,ahk_id %hWndTray%
					xc+=xw/2
					yc+=yw/2
				} else {
					ControlGetPos,xc,yc,,,ToolbarWindow321,ahk_id %hWndTray%
					halfsize:=NumGet(rcPosition,12)/2
					xc+=NumGet(rcPosition,0)+ halfsize
					yc+=NumGet(rcPosition,4)+ (halfsize/2)
				}
			}
			RemoteBuf_close(TrayH)
		}
		If (!x and !y)
			Gosub, TTM_UPDATE
		else if !WinActive("ahk_id " . TT_HWND)
			DllCall("SendMessage", "Uint", TT_HWND, "Uint", %A_ThisLabel%, "Uint", 0, "Uint", (x<9999999 ? x : xc & 0xFFFF)|(y<9999999 ? y : yc & 0xFFFF)<<16)
	Return
	TTM_SETTIPBKCOLOR:
		If B is alpha
			If (%b%)
				B:=%b%
		B := (StrLen(B) < 8 ? "0x" : "") . B
		B := ((B&255)<<16)+(((B>>8)&255)<<8)+(B>>16) ; rgb -> bgr
		DllCall("SendMessage", "Uint", TT_HWND, "Uint", %A_ThisLabel%, "Uint", B, "Uint", 0)
	Return
	TTM_SETTIPTEXTCOLOR:
		If F is alpha
			If (%F%)
				F:=%f%
		F := (StrLen(F) < 8 ? "0x" : "") . F
		F := ((F&255)<<16)+(((F>>8)&255)<<8)+(F>>16) ; rgb -> bgr
		DllCall("SendMessage", "Uint", TT_HWND, "Uint", %A_ThisLabel%, "Uint",F & 0xFFFFFF, "Uint", 0)
	Return
	TTM_SETMARGIN:
		VarSetCapacity(RECT,16)
		Loop,Parse,E,.
			NumPut(A_LoopField,RECT,(A_Index-1)*4)
		DllCall("SendMessage", "Uint", TT_HWND, "Uint", %A_ThisLabel%, "Uint", 0, "Uint", &RECT)
	Return
	TT_SETTOOLINFO:
		If A {
			If A is not Xdigit
				ControlGet,A,Hwnd,,%A%,ahk_id %P%
			ID :=Abs(A)
			If !InStr(TT_ALL_%TT_ID%,Chr(2) . ID . Chr(2))
				TT_ALL_%TT_ID% .= Chr(2) . ID . Chr(2) . ID+Abs(P) . Chr(2)
			If !TOOLINFO_%ID%
				VarSetCapacity(TOOLINFO_%ID%, 40, 0),TOOLINFO_%ID%:=Chr(40)
			else
				Gosub, TTM_DELTOOL
			Numput((N ? 0 : 1)|16,TOOLINFO_%ID%,4),Numput(P,TOOLINFO_%ID%,8),Numput(ID,TOOLINFO_%ID%,12)
			If (text!="")
				NumPut(&text,TOOLINFO_%ID%,36)
			Gosub, TTM_ADDTOOL
			ID :=ID+Abs(P)
			If !TOOLINFO_%ID%
			{
				VarSetCapacity(TOOLINFO_%ID%, 40, 0),TOOLINFO_%ID%:=Chr(40)
				Numput(0|16,TOOLINFO_%ID%,4), Numput(P,TOOLINFO_%ID%,8), Numput(P,TOOLINFO_%ID%,12)
			}
			Gosub, TTM_ADDTOOL
			ID :=Abs(A)
		} else {
			If !TOOLINFO_%ID%
				VarSetCapacity(TOOLINFO_%ID%, 40, 0),TOOLINFO_%ID%:=Chr(40)
			else update:=True
			If (text!=""){
				If InStr(text,"<a"){
					TOOLTEXT_%ID%:=text
					text:=RegExReplace(text,"<a\K[^<]*?>",">")
				} else
					TOOLTEXT_%ID%:=
				NumPut(&text,TOOLINFO_%ID%,36)
			}
			NumPut((!(x . y) ? 0 : 0x20)|(S ? 0x80 : 0)|(L ? 0x1000 : 0),TOOLINFO_%ID%,4), Numput(P,TOOLINFO_%ID%,8), Numput(P,TOOLINFO_%ID%,12)
			Gosub, TTM_ADDTOOL
		}
	Return
	TTN_LINKCLICK:
		Loop 4
			m += *(text + 8 + A_Index-1) << 8*(A_Index-1)
		If !(TTN_FIRST-2=m or TTN_FIRST-3=m)
			Return
		Loop 4
			p += *(text + 0 + A_Index-1) << 8*(A_Index-1)
		If (TTN_FIRST-3=m)
			Loop 4
				option += *(text + 16 + A_Index-1) << 8*(A_Index-1)
		Loop,Parse,hWndArray,% Chr(2)
			If (p=A_LoopField and i:=A_Index)
				break
		Loop,Parse,idArray,% Chr(2)
		{
			If (i=A_Index){
				text:=TOOLTEXT_%A_LoopField%
				If (TTN_FIRST-2=m){
					If Title
					{
						If IsLabel(A_LoopField . title . "Close")
							Gosub % A_LoopField . title . "Close"
						else If IsLabel(title . "Close")
							Gosub % title . "Close"
					} else {
						If IsLabel(A_LoopField . A_ThisFunc . "Close")
							Gosub % A_LoopField . A_ThisFunc . "Close"
						else If IsLabel(A_ThisFunc . "Close")
							Gosub % A_ThisFunc . "Close"
					}
				} else If (InStr(TOOLTEXT_%A_LoopField%,"<a")){
					Loop % option+1
						StringTrimLeft,text,text,% InStr(text,"<a")+1
					If TT_HIDE_%A_LoopField%
						ToolTip(A_LoopField,"","","gTTM_POP")
					If ((a:=A_AutoTrim)="Off")
						AutoTrim, On
					ErrorLevel:=SubStr(text,1,InStr(text,">")-1)
					StringTrimLeft,text,text,% InStr(text,">")
					text:=SubStr(text,1,InStr(text,"</a>")-1)
					If !ErrorLevel
						ErrorLevel:=text
					ErrorLevel=%ErrorLevel%
					AutoTrim, %a%
					If Title
					{
						If IsFunc(f:=(A_LoopField . title))
							%f%(ErrorLevel)
						else if IsLabel(A_LoopField . title)
							Gosub % A_LoopField . title
						else if IsFunc(title)
							%title%(ErrorLevel)
						else If IsLabel(title)
							Gosub, %title%
					} else {
						if IsFunc(f:=(A_LoopField . A_ThisFunc))
							%f%(ErrorLevel)
						else If IsLabel(A_LoopField . A_ThisFunc)
							Gosub % A_LoopField . A_ThisFunc
						else If IsLabel(A_ThisFunc)
							Gosub % A_ThisFunc
					}
				}
				break
			}
		}
	Return
	TT_DESTROY:
		Loop, Parse, TT_ALL_%TT_ID%,% Chr(2)
			If A_LoopField
			{
				ID:=A_LoopField
				Gosub, TTM_DELTOOL
				TOOLINFO_%A_LoopField%:="", TT_HWND_%A_LoopField%:="", TOOLTEXT_%A_LoopField%:="", TT_HIDE_%A_LoopField%:=""
			}
		TT_ALL_%TT_ID%=
	Return
}

MI_ExtractIcon(Filename, IconNumber, IconSize)
{
	If A_OSVersion in WIN_VISTA,WIN_2003,WIN_XP,WIN_2000
	{
	  DllCall("PrivateExtractIcons", "Str", Filename, "Int", IconNumber-1, "Int", IconSize, "Int", IconSize, "UInt*", hIcon, "UInt*", 0, "UInt", 1, "UInt", 0, "Int")
		If !ErrorLevel
		Return hIcon
	}
	If DllCall("shell32.dll\ExtractIconExA", "Str", Filename, "Int", IconNumber-1, "UInt*", hIcon, "UInt*", hIcon_Small, "UInt", 1)
	{
		SysGet, SmallIconSize, 49
		
		If (IconSize <= SmallIconSize) {
		 DllCall("DeStroyIcon", "UInt", hIcon)
		 hIcon := hIcon_Small
		}
	  Else
		DllCall("DeStroyIcon", "UInt", hIcon_Small)
		
		If (hIcon && IconSize)
			hIcon := DllCall("CopyImage", "UInt", hIcon, "UInt", 1, "Int", IconSize, "Int", IconSize, "UInt", 4|8)
	}
	Return, hIcon ? hIcon : 0
}
GetAssociatedIcon(File){
	static 
	sfi_size:=352
	local Ext,Fileto,FileIcon,FileIcon#
	If not sfi
		VarSetCapacity(sfi, sfi_size)
	SplitPath, File,,, Ext
	if Ext in EXE,ICO,ANI,CUR,LNK
	{
		If ext=LNK
		{
			FileGetShortcut,%File%,Fileto,,,,FileIcon,FileIcon#
			File:=!FileIcon ? FileTo : FileIcon
		}
		SplitPath, File,,, Ext
		If !(hIcon%Ext%:=MI_ExtractIcon(InStr(File,"`n") ? SubStr(file,1,InStr(file,"`n")-1) : file,FileIcon# ? FileIcon# : 1,32))
			hIcon%Ext%:=#_hIcon_3
	}
	else If ((!Ext and !#_hIcon) or !InStr(#_hIcons,"|" . Ext . "|")){
		If DllCall("Shell32\SHGetFileInfoA", "str", File, "uint", 0, "str", sfi, "uint", sfi_size, "uint", 0x101){
			Loop 4
				hIcon%Ext% += *(&sfi + A_Index-1) << 8*(A_Index-1)
		}
		hIcons.= "|" . Ext . "|"
	}
	return hIcon%Ext%
}
;	Title:	Remote Buffer
;			*Read and write process memory*
;

/*-------------------------------------------------------------------------------
	Function: Open
			  Open remote buffer

	Parameters:
			H		- Reference to variable to receive remote buffer handle
			hwnd    - HWND of the window that belongs to the process
			size    - Size of the buffer

	Returns:
			Error message on failure
 */
RemoteBuf_Open(ByRef H, hwnd, size) {
	static MEM_COMMIT=0x1000, PAGE_READWRITE=4

	WinGet, pid, PID, ahk_id %hwnd%
	hProc   := DllCall( "OpenProcess", "uint", 0x38, "int", 0, "uint", pid) ;0x38 = PROCESS_VM_OPERATION | PROCESS_VM_READ | PROCESS_VM_WRITE
	IfEqual, hProc,0, return A_ThisFunc ">   Unable to open process (" A_LastError ")"
      
	bufAdr  := DllCall( "VirtualAllocEx", "uint", hProc, "uint", 0, "uint", size, "uint", MEM_COMMIT, "uint", PAGE_READWRITE)
	IfEqual, bufAdr,0, return A_ThisFunc ">   Unable to allocate memory (" A_LastError ")"

	; Buffer handle structure:
	 ;	@0: hProc
	 ;	@4: size
	 ;	@8: bufAdr
	VarSetCapacity(H, 12, 0 )
	NumPut( hProc,	H, 0) 
	NumPut( size,	H, 4)
	NumPut( bufAdr, H, 8)
}

/*----------------------------------------------------
	Function: Close
			  Close the remote buffer

	Parameters:
			  H - Remote buffer handle
 */
RemoteBuf_Close(ByRef H) {
	static MEM_RELEASE = 0x8000
	
	handle := NumGet(H, 0)
	IfEqual, handle, 0, return A_ThisFunc ">   Invalid remote buffer handle"
	adr    := NumGet(H, 8)

	r := DllCall( "VirtualFreeEx", "uint", handle, "uint", adr, "uint", 0, "uint", MEM_RELEASE)
	ifEqual, r, 0, return A_ThisFunc ">   Unable to free memory (" A_LastError ")"
	DllCall( "CloseHandle", "uint", handle )
	VarSetCapacity(H, 0 )
}

/*----------------------------------------------------
	Function:   Read 
				Read from the remote buffer into local buffer

	Parameters: 
         H			- Remote buffer handle
         pLocal		- Reference to the local buffer
         pSize		- Size of the local buffer
         pOffset	- Optional reading offset, by default 0

Returns:
         TRUE on success or FALSE on failure. ErrorMessage on bad remote buffer handle
 */
RemoteBuf_Read(ByRef H, ByRef pLocal, pSize, pOffset = 0){
	handle := NumGet( H, 0),   size:= NumGet( H, 4),   adr := NumGet( H, 8)
	IfEqual, handle, 0, return A_ThisFunc ">   Invalid remote buffer handle"	
	IfGreaterOrEqual, offset, %size%, return A_ThisFunc ">   Offset is bigger then size"

	VarSetCapacity( pLocal, pSize )
	return DllCall( "ReadProcessMemory", "uint", handle, "uint", adr + pOffset, "uint", &pLocal, "uint", size, "uint", 0 ), VarSetCapacity(pLocal, -1)
}

/*----------------------------------------------------
	Function:   Write 
				Write local buffer into remote buffer

	Parameters: 
         H			- Remote buffer handle
         pLocal		- Reference to the local buffer
         pSize		- Size of the local buffer
         pOffset	- Optional writting offset, by default 0

	Returns:
         TRUE on success or FALSE on failure. ErrorMessage on bad remote buffer handle
 */

RemoteBuf_Write(Byref H, byref pLocal, pSize, pOffset=0) {
	handle:= NumGet( H, 0),   size := NumGet( H, 4),   adr := NumGet( H, 8)
	IfEqual, handle, 0, return A_ThisFunc ">   Invalid remote buffer handle"	
	IfGreaterOrEqual, offset, %size%, return A_ThisFunc ">   Offset is bigger then size"

	return DllCall( "WriteProcessMemory", "uint", handle,"uint", adr + pOffset,"uint", &pLocal,"uint", pSize, "uint", 0 )
}

/*----------------------------------------------------
	Function:   Get
				Get address or size of the remote buffer

	Parameters: 
         H		- Remote buffer handle
         pQ     - Query parameter: set to "adr" to get address (default), to "size" to get the size or to "handle" to get Windows API handle of the remote buffer.

	Returns:
         Address or size of the remote buffer
 */
RemoteBuf_Get(ByRef H, pQ="adr") {
	return pQ = "adr" ? NumGet(H, 8) : pQ = "size" ? NumGet(H, 4) : NumGet(H)
}

/*---------------------------------------------------------------------------------------
Group: Example
(start code)
	;get the handle of the Explorer window
	   WinGet, hw, ID, ahk_class ExploreWClass

	;open two buffers
	   RemoteBuf_Open( hBuf1, hw, 128 ) 		
	   RemoteBuf_Open( hBuf2, hw, 16  ) 

	;write something
	   str := "1234" 
	   RemoteBuf_Write( hBuf1, str, strlen(str) ) 

	   str := "_5678" 
	   RemoteBuf_Write( hBuf1, str, strlen(str), 4) 

	   str := "_testing" 
	   RemoteBuf_Write( hBuf2, str, strlen(str)) 


	;read 
	   RemoteBuf_Read( hBuf1, str, 10 ) 
	   out = %str% 
	   RemoteBuf_Read( hBuf2, str, 10 ) 
	   out = %out%%str% 

	   MsgBox %out% 

	;close 
	   RemoteBuf_Close( hBuf1 ) 
	   RemoteBuf_Close( hBuf2 ) 
(end code)
 */

/*-------------------------------------------------------------------------------------------------------------------
	Group: About
	o Ver 2.0 by majkinetor. See http://www.autohotkey.com/forum/topic12251.html
	o Code updates by infogulch
	o Licenced under Creative Commons Attribution-Noncommercial <http://creativecommons.org/licenses/by-nc/3.0/>.  
 */