; Name: UNCT Downloader
; Description: Downloads list of HSCodes from UNCT website with IE using OLE

; P.S: SOME OPTION SELECTIONS ARE HARDCODED
#include <IE.au3>
#include <Debug.au3>

_DebugSetup("IE Debug", 1, 4, "UNCTLog.txt", 1)

Global $iWaiting = 0

;~ AdlibRegister("Reboot", 1000)

Local $sHSCodeList = FileReadToArray("HSCodeList.txt")

For $iListIndex = 0 To UBound($sHSCodeList) - 1
	Local $sHSCode = $sHSCodeList[$iListIndex]
	Downloader()
Next

MsgBox(0, "Completed", "Click OK to close IE")

Func Downloader()
	Local $oIE = _IECreate("https://comtrade.un.org/db/dqbasicquery.aspx")
	Local $oForm = _IEGetObjByName($oIE, "comtrade")
	Local $oHSCode = _IEFormElementGetObjByName($oForm, "plCmdTree:tbInput")
	Send("#{UP}")
	_IEFormElementSetValue ($oHSCode, $sHSCode)

	Local $oSubmit = _IEGetObjById($oIE, "plCmdTree_btSearch")
	_IEAction($oSubmit, "click")
	_IELoadWait($oIE)

	_IENavigate($oIE, "javascript:__doPostBack('plCmdTree$cmdTree','onexpand,0.0')")
	_IELoadWait($oIE)
	_IENavigate($oIE, "javascript:__doPostBack('plCmdTree$cmdTree','onexpand,0.0.0')")
	_IELoadWait($oIE)
	_IENavigate($oIE, "javascript:__doPostBack('plCmdTree$cmdTree','oncheck,0.0.0.0')")
	_IELoadWait($oIE)
	Local $oAdd = _IEGetObjById($oIE, "plCmdTree_btAdd")
	_IEAction($oAdd, "click")
	_IELoadWait($oIE)

	_IELinkClickByText($oIE, "Reporters")
	Local $oList = _IEGetObjByName($oIE, "plSltReporters_lstSO")
	_IEFormElementOptionSelect($oList, "699")
	$oAdd = _IEGetObjById($oIE, "plSltReporters_btAdd")
	_IEAction($oAdd, "click")
	_IELoadWait($oIE)

	_IELinkClickByText($oIE, "Partners")
	Local $oList = _IEGetObjByName($oIE, "plSltPartners_lstSO")
	_IEFormElementOptionSelect($oList, "all")
	$oAdd = _IEGetObjById($oIE, "plSltPartners_btAdd")
	_IEAction($oAdd, "click")
	_IELoadWait($oIE)

	_IELinkClickByText($oIE, "Years")
	Local $oList = _IEGetObjByName($oIE, "plSltYears_lstSO")
	_IEFormElementOptionSelect($oList, "2016")
	$oAdd = _IEGetObjById($oIE, "plSltYears_btAdd")
	_IEAction($oAdd, "click")
	_IELoadWait($oIE)

	Local $oCheckbx = _IEGetObjById($oIE, "cbRgI")
	_IEAction($oCheckbx, "click")
	Local $oCheckbx = _IEGetObjById($oIE, "cbRgE")
	_IEAction($oCheckbx, "click")
	_IELoadWait($oIE)

	Local $oSubmit = _IEGetObjById($oIE, "btSubmit")
	_IEAction($oSubmit, "click")
	_IELoadWait($oIE)

	If Not _IELinkClickByText($oIE, "Direct Download") Then
;~ 		MsgBox(0, "No data", "No Records found for the selection")
		_DebugReport("No Records found for HSCode: " & $sHSCode)
		$iWaiting = 0
		_IEQuit($oIE)
		Return
	EndIf

	If ControlGetHandle("Message from webpage", "", "") Then
		ControlClick("","","[CLASS:Button; INSTANCE:1]")
		_IELoadWait($oIE)
	EndIf

	Sleep(2002)
	Send("!a")

	;~ MouseClick("left", 598, 555, 1)
	;~ ConsoleWrite("Clicked Save As" & @CRLF)

	ConsoleWrite("Waiting for Save As" & $sHSCode & @CRLF)
	WinWaitActive("Save As")
	ControlSend("", "", 1001, @ScriptDir & "\" & $sHSCode & ".csv")
	ControlClick("", "", "[CLASS:Button; INSTANCE:1; TEXT:&Save]")
	_IEQuit($oIE)

	$iWaiting = 0

EndFunc

Func Reboot()

	$iWaiting += 1
	ConsoleWrite($iWaiting & " ")

;~ 	If ControlGetHandle("Confirm Save As", "", "") Then
;~ 		ControlClick("","","[CLASS:Button; INSTANCE:1]")
;~ 	EndIf

	If @error = 3 Or @error = 7 Or $iWaiting > 120 Then	ShellExecute("Terminator.bat")

EndFunc
