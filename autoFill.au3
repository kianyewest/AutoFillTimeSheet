#include <IE.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <File.au3>

;Set up Variables
Local $dayIndex = @WDAY+2
IF $dayIndex > 7 THEN $dayIndex=0 EndIf

;Get details

Local $aRetArray, $sFilePath =  "entries.csv"
    ; Re-read it - without count
    _FileReadToArray($sFilePath, $aRetArray, $FRTA_NOCOUNT, ",")
    ;_ArrayDisplay($aRetArray, "1D array - no count", Default, 8,",")

	  Local $sDescription ="Description"
	  Local $sDuration ="Duration"
	  Local $sDescriptionIndex =0
	  Local $sDurationIndex =0

	  For $i = 0 to UBound($aRetArray,$UBOUND_COLUMNS) -1
		 If StringRegExp($aRetArray[0][$i], $sDescription) Then
			$sDescriptionIndex=$i
			;MsgBox($MB_SYSTEMMODAL, "Title", "This message box will timeout after 10 seconds or select the OK button.", 10)
		 endif
		 If StringRegExp($aRetArray[0][$i], $sDuration) Then
			$sDurationIndex=$i
			;MsgBox($MB_SYSTEMMODAL, "Title", "This message box will timeout after 10 seconds or select the OK button.", 10)
		 endif
	  Next

;Do browser Stuff




;Open Page
Local $oIE = _IECreate("https://timetorque.datatorque.com/timesheet.asp")

;Login
Local $oDiv = _IEGetObjById($oIE, "cmdSubmit")
_IEAction($oDiv, "click")

;Go to timesheet page
_IENavigate ($oIE,"https://timetorque.datatorque.com/timesheet.asp")

;Change Day
Local $oInputs = _IETagNameGetCollection($oIE, "td")
Local $sTxt = ""
Local $days = [7]
For $oInput In $oInputs
    $sTxt &= $oInput.id & @CRLF & "h"
	If  $oInput.id == "mnu" Then
    _ArrayAdd($days, $oInput+"hm")
   EndIf
Next
_IEAction($days[$dayIndex], "click")

;Loop over rows, skip first as header
	  For $i = 1 to UBound($aRetArray,$UBOUND_ROWS) -1
		 $hours = convertToHours($aRetArray[$i][$sDurationIndex])
		 SaveRow($aRetArray[$i][$sDescriptionIndex],$hours)
	  Next

Func SaveRow($comment,$hours)
   sleep(1000)
   ;Input Info
   Send("FTM")
   ;sleep(1000)
   Send("{TAB}")
   ;sleep(1000)
   Send("Dev")
   ;sleep(1000)
   Send("{TAB}")
   ;sleep(1000)
   Send($hours)
   ;sleep(1000)
   Send("{TAB}")
   ;sleep(1000)
   Send($comment)

   sleep(1000)

   ;Get Save Button
   Local $oInputs = _IETagNameGetCollection($oIE, "input")

   For $oInput In $oInputs
	   If  $oInput.name == "Save" Then
		 ;Why not have a variable and save $oInput? Cause I tired and it didnt work :(
		 _IEAction($oInput, "click")
	  EndIf
   Next

EndFunc


Func convertToHours($preSplitTime)
   Local $Time = StringSplit($preSplitTime,":",$STR_NOCOUNT )
		 ;_ArrayDisplay($Time, "$Time", Default, 8,",")
		 Local $TimeInHours = $Time[0]+($Time[1]/60)
		 Local $RoundedTimeInHours = Round($TimeInHours, 2) ;
		 IF $RoundedTimeInHours < 0.25 Then
			$RoundedTimeInHours = 0.25
		 EndIf
		 Return $RoundedTimeInHours
EndFunc