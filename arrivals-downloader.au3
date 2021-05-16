; Name: Arrivals Downloader
; Description: A bot to automate downloading arrivals from old agmark website with IE using OLE

; P.S: Supporting Terminator.bat is not included but is simple to make
; P.S: SOME SELECTIONS ARE HARDCODED
; Manual changes required in code for:
; Year at: 28, 58, 111, 119

#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Change2CUI=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <IE.au3>
#include <Debug.au3>

_DebugSetup("IE Debug", 1, 4, "AppLog.txt", 1)

Global $iWaiting

AdlibRegister("Reboot", 1000)

If Not $cmdline[0] Then
	If MsgBox(1, "Warning", "This programs terminates all open instances of Excel and IE on " & _
		"accounting an error. Close all open programs before continuing." & @CRLF & @CRLF & _
		"Start Execution?") = 2 Then Exit
EndIf

Local $sStateIndex = IniRead("Status.ini", "Last Session", "State Index", "0")
Local $iCropIndex = IniRead("Status.ini", "Last Session", "Crop Index", "0")
Local $sMonthName = IniRead("Status.ini", "Last Session", "Month Name", "")
Local $sCropList = FileReadToArray("App Crops.txt")
Local $sStateList = FileReadToArray("App States.txt")

For $iState = $sStateIndex To UBound($sStateList) - 1
	IniWrite("Status.ini", "Last Session", "State Index", $iState)
	Local $sStateName = $sStateList[$iState]

	For $iCrop = $iCropIndex To UBound($sCropList) - 1
		IniWrite("Status.ini", "Last Session", "Crop Index", $iCrop)
		Local $sCropName = $sCropList[$iCrop]
		Global $oIE = _IECreate("http://agmarkweb.dacnet.nic.in/SA_Month_ArrMar.aspx",0,1,1)

		Local $oForm = _IEFormGetObjByName($oIE, "Form1")
		Local $oCommList = _IEFormElementGetObjByName($oForm, "Commodit_list")
		_IEFormElementOptionSelect($oCommList, $sCropName, 1, "byText")
		_IELoadWait($oIE)
		Sleep(1001)

		$oForm = _IEFormGetObjByName($oIE, "Form1")
		Local $oStateList = _IEFormElementGetObjByName($oForm, "State_list")
		_IEFormElementOptionSelect($oStateList, $sStateName , 1, "byValue")
		_IELoadWait($oIE)
		Sleep(1001)

		$oForm = _IEFormGetObjByName($oIE, "Form1")
		Local $oYearList = _IEFormElementGetObjByName($oForm, "Yea_list")
		_IEFormElementOptionSelect($oYearList, "2014", 1, "byValue")
		_IELoadWait($oIE)
		Sleep(1001)

		$oForm = _IEFormGetObjByName($oIE, "Form1")
		Local $oMonthList = _IEFormElementGetObjByName($oForm, "Mont_list")
		Local $sMonthLine = _IEPropertyGet($oMonthList, "innertext")
		Local $sMonthList = StringSplit(StringMid($sMonthLine, 15), " ", 2)

		_DebugReportVar("State: ", $sStateName)
		_DebugReportVar("Crop: ", $sCropName)
		_DebugReportVar("Months Found: ", $sMonthList)

		_IEQuit($oIE)

		For $sMonth in $sMonthList
			If $sMonthList[0] = "" Then ExitLoop
			If ($sMonthName <> "" And $sMonth <> $sMonthName) Then ContinueLoop
			IniWrite("Status.ini", "Last Session", "Month Name", $sMonth)
			$sMonthName = ""

			$oIE = _IECreate("http://agmarkweb.dacnet.nic.in/SA_Month_ArrMar.aspx",0,1,1)
			$oForm = _IEFormGetObjByName($oIE, "Form1")
			$oCommList = _IEFormElementGetObjByName($oForm, "Commodit_list")
			_IEFormElementOptionSelect($oCommList, $sCropName, 1, "byText")
			_IELoadWait($oIE)
			Sleep(1001)

			$oForm = _IEFormGetObjByName($oIE, "Form1")
			$oStateList = _IEFormElementGetObjByName($oForm, "State_list")
			_IEFormElementOptionSelect($oStateList, $sStateName , 1, "byValue")
			_IELoadWait($oIE)
			Sleep(1001)

			$oForm = _IEFormGetObjByName($oIE, "Form1")
			$oYearList = _IEFormElementGetObjByName($oForm, "Yea_list")
			_IEFormElementOptionSelect($oYearList, "2014", 1, "byValue")
			_IELoadWait($oIE)
			Sleep(1001)

			$oForm = _IEFormGetObjByName($oIE, "Form1")
			$oMonthList = _IEFormElementGetObjByName($oForm, "Mont_list")
			_DebugReportVar("Month: ", $sMonth)
			_IEFormElementOptionSelect($oMonthList, $sMonth, 1, "byText")
			_IELoadWait($oIE)
			Sleep(1001)

			Local $oSubmit = _IEGetObjById($oIE, "But_submit")
			_IEAction($oSubmit, "click")
			_IELoadWait($oIE)
			Sleep(1001)

			Send("^s")

			WinWaitActive("Save Webpage")
			ControlSend("", "", 1001, @ScriptDir & "\" & $sStateName & "-" & $sCropName & "-" & $sMonth & ".xls")
			ControlClick("", "", "[CLASS:Button; INSTANCE:1; TEXT:&Save]")
			Sleep(2000)

			If ControlGetHandle("Confirm Save As", "", "") Then
				ControlClick("","","[CLASS:Button; INSTANCE:1]")
			EndIf

			If ControlGetHandle("Error Saving Webpage", "", "") Then
				ControlClick("Error Saving Webpage", "OK", "")
				Send("^s")
				WinWaitActive("Save Webpage")
				ControlSend("", "", 1001, @ScriptDir & "\" & $sStateName & "-" & $sCropName & "-" & $sMonth & ".xls")
				ControlClick("", "", "[CLASS:Button; INSTANCE:1; TEXT:&Save]")
				Sleep(2000)
			EndIf

			While _IEPropertyGet($oIE, "busy")
				Sleep(2000)
			WEnd
			Sleep(4000)
			_IEQuit($oIE)

			Local $sFileLoc = @ScriptDir & "\" & $sStateName & "-" & $sCropName _
				& "-" & $sMonth & ".xls"

			If	FileExists($sFileLoc) Then
				ShellExecute($sFileLoc)
				$sFileLoc = ""
			Else
				ShellExecute("Terminator.bat")
				Exit
			EndIf

			WinWaitActive("Microsoft Office Excel")
			ControlClick("", "", "[CLASS:Button; INSTANCE:1]")

			Sleep(5000)

			$oExcel = ObjGet("", "Excel.Application")
			$oExcel.ActiveCell.Offset(1,1).Select
			$oExcel.ActiveCell.CurrentRegion.Activate

			Local $iStates = $oExcel.Selection.Rows.Count - 5
			Local $hDataFile = FileOpen("App.csv", 1)

			For $i = 1 To $iStates
				If $oExcel.ActiveCell.Offset(1+$i,1).Value <> "" Then
					Local $iQtyDays = StringSplit($oExcel.ActiveCell.Offset(1+$i,1).Value, " ")
					$iQtyDays[2] = StringMid($iQtyDays[2], 2, StringLen($iQtyDays[2])-2)
					FileWriteLine($hDataFile, $sStateName & "," & _
						"2014," & $sCropName & "," & $sMonth & "," & _
						$oExcel.ActiveCell.Offset(1+$i,0).Value & "," & _
						$iQtyDays[1] * $iQtyDays[2] & "," & $iQtyDays[2])
				EndIf

				If $oExcel.ActiveCell.Offset(1+$i,3).Value <> "" Then
					$iQtyDays = StringSplit($oExcel.ActiveCell.Offset(1+$i,3).Value, " ")
					$iQtyDays[2] = StringMid($iQtyDays[2], 2, StringLen($iQtyDays[2])-2)
					FileWriteLine($hDataFile, $sStateName & "," & _
						"2013," & $sCropName & "," & $sMonth & "," & _
						$oExcel.ActiveCell.Offset(1+$i,0).Value & "," & _
						$iQtyDays[1] * $iQtyDays[2] & "," & $iQtyDays[2])
				EndIf
			Next

			IniDelete("Status.ini", "Last Session", "Month Name")
			FileClose($hDataFile)

;~ 			$oExcel.ActiveWorkbook.Saved = 1
;~ 			$oExcel.Activeworkbook.Close

			If ControlGetHandle("", "", "[CLASS:NetUIHWND; INSTANCE:1]") Then
				WinActivate("[CLASS:NetUIHWND; INSTANCE:1]")
				ControlClick("", "", "[CLASS:NetUIHWND; INSTANCE:1]", "left" , 1, 175, 66)
			EndIf
			$iWaiting = 0
			Send("#m")
		Next
		$sMonthName = ""
		WinKill("[CLASS:XLMAIN]")
	Next
	$iCropIndex = 0
Next

Func Reboot()

	$iWaiting += 1
	ConsoleWrite($iWaiting & " ")

	If ControlGetHandle("Confirm Save As", "", "") Then
		ControlClick("","","[CLASS:Button; INSTANCE:1]")
	EndIf

	If @error = 3 Or @error = 7 Or ControlGetHandle("Message from webpage", "", "") Or $iWaiting > 200 Then
		Sleep(10000)
		ShellExecute("Terminator.bat")
	EndIf

EndFunc

FileDelete("Status.ini")

MsgBox(0, "Mission Accomplished!", "Completed")

ShellExecute("AppLog.txt")
