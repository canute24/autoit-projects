; Name: File Downloader
; Description: Gets a list of links for a file and downloads them with IE using OLE

#include<IE.au3>
#include<Date.au3>

OnAutoItExitRegister ("Cleanup")
Local $bIECreated = 0

If StringLen(@ScriptDir) = 3 Then
	FileWriteLine("links.txt", "Cant run from drive root. Put in folder")
	Exit
EndIf

Local $oIE = _IECreate(@ScriptDir & "\download.htm", 0, 0)
Local $bIECreated = Not IsNumber(@error)
Local $sSearchStr = FileReadLine("settings.ini")
Local $oLinks = _IELinkGetCollection($oIE)
Local $hDling
Local $sFilename

If DirGetSize("Downloads") = -1 Then DirCreate("Downloads")

For $i in $oLinks
	If StringRegExp($i.href, $sSearchStr) Then
		$sFilename = StringTrimLeft($i.href, StringLen($i.href) - StringInStr(StringReverse($i.href), "/") + 1)
		If FileExists("Downloads\" & $sFilename) Then $sFilename = StringReplace(_NowTime(5), ":", "") & $sFilename
				$hDling = InetGet($i.href, @ScriptDir & "\Downloads\" & $sFilename)
		FileWriteLine("links.txt", $i.href & " -> " & @ScriptDir & "\Downloads\" & $sFilename)
	EndIf
Next

Func Cleanup()
	If $bIECreated Then _IEQuit($oIE)
EndFunc
