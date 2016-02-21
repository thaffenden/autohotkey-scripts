; Youtube
#F2:: WinActivate( "Youtube", "https://www.youtube.com/" )

#SingleInstance force
SetTitleMatchMode, 2 

WinActivate( TheWindowTitle, TheProgramTitle )
{
	SetTitleMatchMode, 2
	IfWinExist, - Google Chrome
	{
	   	WinActivate ; use the window found above
   
		;Click in the address bar
		MouseClick, left, 196, 59
      
		WinGetTitle, Title, A  ;get active window title
		OrigTitle := Title 		
	
		Loop
		{
			if(InStr(Title, "Youtube")>0)
			{
				;Found
				break ; Terminate the loop
			}
			Send ^{Tab}
			Sleep, 50
			WinGetTitle, Title, A  ;get active window title
			if(Title = OrigTitle ) 
			{
				;Match to original title, assume all seen
				break			
			}
   		} 
	} 

	;Was the tab found?
	WinGetTitle, Title, A  ;get active window title
	if(InStr(Title, "Youtube")>0)
	{
		;Found
	} else {
		Run, "https://www.youtube.com/"
	}
}
return