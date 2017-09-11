/*
	Function: hotstrings
		Dynamically adds regular expression hotstrings.

	Parameters:
		c - regular expression hotstring
		a - (optional) text to replace hotstring with or a label to goto, 
			leave blank to remove hotstring definition from triggering an action

	Examples:
> hotstrings("(B|b)tw\s", "%$1%y the way") ; type 'btw' followed by space, tab or return
> hotstrings("i)omg", "oh my god{!}") ; type 'OMG' in any case, upper, lower or mixed
> hotstrings("\bcolou?r", "rgb(128, 255, 0);") ; '\b' prevents matching with anything before the word, e.g. 'multicololoured'

	License:
		- Version 2.56 <http://www.autohotkey.net/~polyethene/#hotstrings>
		- Dedicated to the public domain <http://creativecommons.org/licenses/publicdomain/>
*/
hotstrings(k, a = "")
{
	static z, m = "*~$", s, t, w = 2000
	global $
	If z = ; init
	{
		Loop, 94
		{
			c := Chr(A_Index + 32)
			If A_Index not between 33 and 58
				Hotkey, %m%%c%, __hs
		}
		e = BS|Space|Enter|Return|Tab
		Loop, Parse, e, |
			Hotkey, %m%%A_LoopField%, __hs
		z = 1
	}
	If (a == "" and k == "") ; poll
	{
		StringTrimLeft, q, A_ThisHotkey, StrLen(m)
		If q = BS
		{
			If (SubStr(s, 0) != "}")
				StringTrimRight, s, s, 1
		}
		Else
		{
			If q = Space
				q := " "
			Else If q = Tab
				q := "`t"
			Else If q in Enter,Return
				q := "`n"
			Else If (StrLen(q) != 1)
				q = {%q%}
			Else If (GetKeyState("Shift") ^ GetKeyState("CapsLock", "T"))
				StringUpper, q, q
			s .= q
		}
		Loop, Parse, t, `n ; check
		{
			StringSplit, x, A_LoopField, `r
			If (RegExMatch(s, x1 . "$", $)) ; match
			{
				StringLen, l, $
				StringTrimRight, s, s, l
				SendInput, {BS %l%}
				If (IsLabel(x2))
					Gosub, %x2%
				Else
				{
					Transform, x0, Deref, %x2%
					SendInput, %x0%
				}
			}
		}
		If (StrLen(s) > w)
			StringTrimLeft, s, s, w // 2
	}
	Else ; assert
	{
		StringReplace, k, k, `n, \n, All ; normalize
		StringReplace, k, k, `r, \r, All
		Loop, Parse, t, `n
		{
			l = %A_LoopField%
			If (SubStr(l, 1, InStr(l, "`r") - 1) == k)
				StringReplace, t, t, `n%l%
		}
		If a !=
			t = %t%`n%k%`r%a%
	}
	Return
	__hs: ; event
	hotstrings("", "")
	Return
}

#Numpad2::
IfWinActive Spotify
{
	WinMinimize
}
else
{
	IfWinExist Spotify
	{
		DetectHiddenWindows, On 
		WinActivate
		DetectHiddenWindows, Off
	}
	else
	{
		Run C:\Program Files\Spotify\Spotify.exe
	}
}
return


SendToSpotify(key)
{
	DetectHiddenWindows, On 
	IfWinExist Spotify
	{
		ControlSend, ahk_parent, %key%, ahk_class SpotifyMainWindow 
	}
	DetectHiddenWindows, Off
	return
}


;;; FUNCTIONS
SendToGoogleMusic(action)
{

;;; SETTINGS
#SingleInstance, Force
SetTitleMatchMode, 2
SetTitleMatchMode, fast
DetectHiddenWindows, On

   IfWinExist, Google Play
    {
        WinGet, current_active_window, ID, A
        WinActivate, Google Play
        
        if (action = "prev")
            Send, {Left}
        else if (action = "next")
            Send, {Right}
        else if (action = "playpause")
            Send, {Space}
        else
            MsgBox % "Invalid action: " . action
	 
	WinActivate, ahk_id %current_active_window%
    }
    IfWinExist, Grooveshark
    {
        WinGet, current_active_window, ID, A
        WinActivate, Grooveshark
        
        if (action = "prev")
            Send, {Left}
        else if (action = "next")
            Send, {Right}
        else if (action = "playpause")
            Send, {Space}
        else
            MsgBox % "Invalid action: " . action
	 
	WinActivate, ahk_id %current_active_window%
    }
    return
}


#Numpad8::
key = {Space}
;SendToGoogleMusic("playpause")
;SendToSpotify(key) ; Not working as of spring 2015 for some reason;. Workaround: use media play/pause key. Drawback: will not work if another media application open that overrides.
Send {Media_Play_Pause}
return

#Numpad4::
key = ^{Left}
;SendToGoogleMusic("prev")
SendToSpotify(key)
return

#Numpad6::
;SendToGoogleMusic("next")
key = ^{Right}
SendToSpotify(key)
return


#NumpadAdd::
Send {Volume_Up}
return

#NumpadSub::
Send {Volume_Down}
return

#NumpadMult::
Send {Volume_Mute}
return

#n:: Run, "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE" /c ipm.note

#o:: Run, "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE" /recycle

#c:: Run, "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" about:blank chrome://newtab 

#t:: Run, "C:\Users\efillan\Desktop\nice2have\Tider2.xlsm"

#w::
DetectHiddenWindows, On
IfWinActive, Microsoft Word
{
	Send {Ctrl down}{F6}{Ctrl up}
	return
}
SetTitleMatchMode 2
WinActivate, Word
DetectHiddenWindows, Off
return

#x:: 
SetTitleMatchMode 2
IfWinActive, Microsoft Excel
{
	Send ^{Tab}
	return
}
DetectHiddenWindows, On
WinActivate, Microsoft Excel
DetectHiddenWindows, Off
return

;#f:: 
;SetTitleMatchMode 2
;DetectHiddenWindows, On
;WinActivate, Mozilla Firefox
;DetectHiddenWindows, Off
;return

#i:: 
SetTitleMatchMode 2
DetectHiddenWindows, On
WinActivate, Windows Internet Explorer
DetectHiddenWindows, Off
return

;SC029::Send {Backspace} ; backtick key

;§::Backspace

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;   These two commands sorts by outlook by Received or From by pressing Alt+1 or Alt+2      ;;
;; 2012-11-20: Conflicts with WMII, remove. Also, it's not hard to type Alt,v,a,d manually.  ;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;


;#!1::
;Send {Alt}vad
;return

;#!2::
;Send {Alt}vaf
;return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;



Capslock::Ctrl
+Capslock::Capslock




::$today::
::¤today::
formattime, var, , yyyy-MM-dd
  send %var%
return

:Os:dallasstart::
SendInput ./lts_system_setup -deinstall; sleep 60; ./lts_system_setup -install; sleep 60; lts_control start -c /lab/cgw/dallas/cgwtesttool-
return

:Os:brfl::
SendInput Best regards,{Enter}Filip Lange
return

:Os:mvhfl::
SendInput Med vänlig hälsning,{Enter}Filip Lange
return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; CLEARQUEST TR FINDER ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

#y::
InputBox text, ClearQuest TR finder, Syntax:`nGGSN TRs: "GGSN00012345" or "12345"`nCPG TRs: "CGW00012345"`nCRs: "GSNCR00012345"
if ErrorLevel
	return
text = %text% ; trim whitespaces
StringMid, stringstart, text, 1, 3

; Default is GGSN
nodetype = GGSN
database = cqggsn
recordtype = Defect

if stringstart = GGS
{
	; do nothing
}
else if stringstart = CGW
{
	nodetype = CGW
}
else if stringstart = GSN
{
	database = GSNCR
	nodetype = GSNCR
	recordtype = CR
}
else
{
	StringLen, input_string_length, text
	tr_number_template = GGSN00000000
	StringLeft, tr_number_template_short, tr_number_template, 12-input_string_length 
	text = %tr_number_template_short%%text% 	
}
Run, https://pdupc-clearquest.mo.sw.ericsson.se/cqweb/restapi/%database%/%nodetype%/RECORD/%text%?format=HTML&noframes=true&recordType=%recordtype%
return

; Examples
; https://pdupc-clearquest.mo.sw.ericsson.se/cqweb/restapi/cqggsn/CGW/RECORD/CGW00017158?format=HTML&noframes=true&recordType=Defect
; https://pdupc-clearquest.mo.sw.ericsson.se/cqweb/restapi/GSNCR/GSNCR/RECORD/GSNCR00006279?format=HTML&noframes=true&recordType=CR

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; PDC Web Test Case Navigator ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

#u::
InputBox text, PDC Web Test Case Navigator, Syntax: "TC1512.2" or "1512.2"
if ErrorLevel
	return
text = %text% ; trim whitespaces
StringMid, stringstart, text, 1, 1
if stringstart is integer
text = TC%text% 

Run, http://selnpcweb01.seln.ete.ericsson.se/epgtest/?to=all&tc=%text%&team=LSV&build=all&view=graph&official=false&valid=false

return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;


; Note: You can optionally release Capslock or the middle mouse button after
; pressing down the mouse button rather than holding it down the whole time.
; This script requires v1.0.25+.

~MButton & LButton::
CoordMode, Mouse  ; Switch to screen/absolute coordinates.
MouseGetPos, EWD_MouseStartX, EWD_MouseStartY, EWD_MouseWin
WinGetPos, EWD_OriginalPosX, EWD_OriginalPosY,,, ahk_id %EWD_MouseWin%
WinGet, EWD_WinState, MinMax, ahk_id %EWD_MouseWin% 
if EWD_WinState = 0  ; Only if the window isn't maximized 
    SetTimer, EWD_WatchMouse, 10 ; Track the mouse as the user drags it.
return

EWD_WatchMouse:
GetKeyState, EWD_LButtonState, LButton, P
if EWD_LButtonState = U  ; Button has been released, so drag is complete.
{
    SetTimer, EWD_WatchMouse, off
    return
}
GetKeyState, EWD_EscapeState, Escape, P
if EWD_EscapeState = D  ; Escape has been pressed, so drag is cancelled.
{
    SetTimer, EWD_WatchMouse, off
    WinMove, ahk_id %EWD_MouseWin%,, %EWD_OriginalPosX%, %EWD_OriginalPosY%
    return
}
; Otherwise, reposition the window to match the change in mouse coordinates
; caused by the user having dragged the mouse:
CoordMode, Mouse
MouseGetPos, EWD_MouseX, EWD_MouseY
WinGetPos, EWD_WinX, EWD_WinY,,, ahk_id %EWD_MouseWin%
SetWinDelay, -1   ; Makes the below move faster/smoother.
WinMove, ahk_id %EWD_MouseWin%,, EWD_WinX + EWD_MouseX - EWD_MouseStartX, EWD_WinY + EWD_MouseY - EWD_MouseStartY
EWD_MouseStartX := EWD_MouseX  ; Update for the next timer-call to this subroutine.
EWD_MouseStartY := EWD_MouseY
return

;
; AutoHotkey Version: 1.x
; Language:       English
; Platform:       Win9x/NT
; Written by Philipp Otto, Germany
;
; Script Function:
;	Template script (you can customize this template by editing "ShellNew\Template.ahk" in your Windows folder)
;

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

SetCapsLockState, AlwaysOff

;CapsLock & k::
;       if getkeystate("alt") = 0
;               Send,{Up}
;       else
;               Send,+{Up}
;return
;
;CapsLock & l::
;       if getkeystate("alt") = 0
;               Send,{Right}
;       else
;               Send,+{Right}
;return
;
;CapsLock & h::
;       if getkeystate("alt") = 0
;               Send,{Left}
;       else
;               Send,+{Left}
;return
;
;CapsLock & j::
;       if getkeystate("alt") = 0
;               Send,{Down}
;       else
;               Send,+{Down}
;return
;
;CapsLock & ,::
;       if getkeystate("alt") = 0
;               Send,{PgDn}
;       else
;               Send,+{PgDn}
;return
;
;CapsLock & i::
;       if getkeystate("alt") = 0
;               Send,{PgUp}
;       else
;               Send,+{PgUp}
;return
;
;CapsLock & u::
;       if getkeystate("alt") = 0
;               Send,^{Left}
;       else
;               Send,+^{Left}
;return
;
;CapsLock & o::
;       if getkeystate("alt") = 0
;               Send,^{Right}
;       else
;               Send,+^{Right}
;return
;
;CapsLock & m::
;       if getkeystate("alt") = 0
;               Send,{Home}
;       else
;               Send,+{Home}
;return
;
;
;CapsLock & .::
;       if getkeystate("alt") = 0
;               Send,{End}
;       else
;               Send,+{End}
;return


;remap Caps Lock + [ to escape for Vim
;CapsLock & [::
;       Send,{End}
;return


;       CapsLock & �::                                  ;has to be changed (depending on the keyboard-layout)
;               if getkeystate("alt") = 0
;                       Send,^{Right}
;               else
;                       Send,+^{Right}
;       return

;CapsLock & BS::Send,{Del}

;Prevents CapsState-Shifting
;CapsLock & Space::Send,{Space}

;*Capslock::SetCapsLockState, AlwaysOff
;+Capslock::SetCapsLockState, On


;;; Paste and then increment first number in clipboard by 1. Keeps leading zeros. Goes from 0099 to 0100. ;;;

#v::
Send,^v
RegExMatch(clipboard,"(\d+)",Number)
RegExMatch(Number,"(^0+)",leadingZeros)
IncrementedNumber := Number+1
NewNumber = %leadingZeros%%IncrementedNumber%
if (StrLen(NewNumber) > StrLen(Number)) {
	NewNumber := SubStr(NewNumber, 2)
}
Clipboard:=RegExReplace(clipboard,Number,NewNumber)

return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;


#g:: 

dropFolder1 = E:\Dropbox
dropFolder2 = C:\Users\efillan\Dropbox

IfExist %DropFolder1%
  dropfolder = %dropfolder1%
else
  dropfolder = %dropfolder2%

Run, C:\Users\efillan\AppData\Local\Continuum\Anaconda2\pythonw.exe %dropfolder%\scripts\navigate_excel_sheets.py

return
