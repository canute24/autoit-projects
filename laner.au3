; Name: Laner
; Description: Creates a hotkey entry to annouce role in LoL normal queue

#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=D:\Riot Games\League of Legends\RADS\system\lol.ico
#AutoIt3Wrapper_Compression=4
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

HotKeySet("!j", "Top")
HotKeySet("!k", "Mid")
HotKeySet("!l", "ADC")
HotKeySet("!;", "JG")

While (1)

	Sleep(10000)

WEnd

Func Top()

	WinWaitActive("League of Legends")
	MouseClick("left", 248, 679, 1)
	Send("top{ENTER}")

EndFunc

Func Mid()

	WinWaitActive("League of Legends")
	MouseClick("left", 248, 679, 1)
	Send("mid{ENTER}")

EndFunc

Func ADC()

	WinWaitActive("League of Legends")
	MouseClick("left", 248, 679, 1)
	Send("adc{ENTER}")

EndFunc

Func JG()

	WinWaitActive("League of Legends")
	MouseClick("left", 248, 679, 1)
	Send("jg{ENTER}")

EndFunc
