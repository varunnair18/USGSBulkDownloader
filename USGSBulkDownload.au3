#include <Constants.au3>
#include <AutoItConstants.au3>
#include <IE.au3>
#include <Array.au3>
#pragma compile(Console, True)

;
; AutoIt Version: 3.0
; Language:       English
; Platform:       Win9x/NT
; Author:         Varun Nair
;
; Script Function:
;   Given coordindates and a timeframe, downloads all orthographic high-resolution imagery from USGS in Internet Explorer.
;

Global $oIE = _IECreate("https://earthexplorer.usgs.gov/")

; Add Coords Here
; Coordinates entered here are for: Clyde, NC
Global $coords[4][6] = [[35, 33, 23, 82, 51, 28], [35,33,25,82,50,16], [35,30,49,82,50,08],[35,30,49,82,51,21]]
; Maximize Window
$HWND = _IEPropertyGet($oIE, "hwnd")
WinSetState($HWND, "", @SW_MAXIMIZE)

; Calls to all functions
Call ("addCoords", 4)
Call ("addDate", "03/19/2010", "03/19/2010")
Call ("maxResultOptions")
Call ("selectDataSet")
Call ("downloadResults")
Call ("seeNextPage")

; Adds coordinates to create a polygon from which imagery will be downloaded.
Func addCoords ($numOfCoords)

   ; Clear all exisiting coordinates
   $AddCoordinate = _IEGetObjById($oIE, "coordEntryClear")
   _IEAction($AddCoordinate, "click")

   ; Add as many coordinates as given by parameter $numOfCoords
   For $counter = 1 to $numOfCoords

   ; Get Add Coordinate Button Object and Click on it
   $AddCoordinate = _IEGetObjById($oIE, "coordEntryAdd")
   _IEAction($AddCoordinate, "click")

   ; Get Add Coordinate Popup and Populate It
   $popup = _IEGetObjById($oIE, "coordEntryDialogArea")
   Local $oElements = _IETagNameAllGetCollection($popup)
	  For $oElement In $oElements
		 If $oElement.id == "degreesLat" Then
			_IEPropertySet($oElement, "innertext", String($coords[$counter-1][0]))
		 EndIf
		 If $oElement.id == "minutesLat" Then
			_IEPropertySet($oElement, "innertext", String($coords[$counter-1][1]))
		 EndIf
		 If $oElement.id == "secondsLat" Then
			_IEPropertySet($oElement, "innertext", String($coords[$counter-1][2]))
		 EndIf
		 If $oElement.id == "degreesLng" Then
			_IEPropertySet($oElement, "innertext", String($coords[$counter-1][3]))
		 EndIf
		 If $oElement.id == "minutesLng" Then
			_IEPropertySet($oElement, "innertext", String($coords[$counter-1][4]))
		 EndIf
		 If $oElement.id == "secondsLng" Then
			_IEPropertySet($oElement, "innertext", String($coords[$counter-1][5]))
		 EndIf

	  ; IMPORTANT NOTE: This script currently only works with coordinates in the Northwestern Hemisphere.
	  Next

	  ; Click Add Button
	  Local $oButtons = _IETagNameGetCollection($oIE, "button")
		 For $oButton In $oButtons
			If $oButton.innerText == "Add" Then _IEAction($oButton, "click")
		 Next
	  Next
EndFunc

Func addDate ($startDate, $endDate)

   $startDateInput = _IEGetObjByName($oIE, "start_linked")
   _IEAction($startDateInput, "focus")
   _IEPropertySet($startDateInput, "innerText", $startDate)

   $startDateInput = _IEGetObjByName($oIE, "end_linked")
   _IEAction($startDateInput, "focus")
   _IEPropertySet($startDateInput, "innerText", $endDate)
EndFunc

Func maxResultOptions ()
   ; Click on Result Options Tab
   $resultOptionsTab = _IEGetObjById($oIE, "tabResultOptions")
   _IEAction($resultOptionsTab, "click")

   ;Select Max Number of Results (25,000)
   $resultOptionsSelector = _IEGetObjById($oIE, "numReturnSelect")
   _IEAction($resultOptionsSelector, "focus")
   Send("{DOWN 6}")

   ; Click Data Sets Button
   $dataSetsTab = _IEGetObjById($oIE, "tab2")
   _IEAction($dataSetsTab, "click")


      Sleep(3000)
EndFunc

Func selectDataSet()
   ; Click on High Resolution Orthoimagery ("coll
   $dataSetsTab = _IEGetObjById($oIE, "collLabel_3411")
   _IEAction($dataSetsTab, "click")

      ;Click on Results Tab
   $dataSetsTab = _IEGetObjById($oIE, "tab4")
   _IEAction($dataSetsTab, "click")
   Sleep(2000)

   EndFunc


Func downloadResults()

   ; Click on Download Link for All Images on the Page
   Local $oInputs = _IETagNameGetCollection($oIE, "a")

   ; Creating temp variable that will allow us to grab every third string returned from the items with a class of "metadata".
   Local $oMod = 2
   Local $entityIds[10] = []
   For $oInput In $oInputs
	  If $oInput.classname == "metadata" Then
		 ; Add image entity ID to array to navigate to
		 If Mod($oMod, 3) == 0 Then _ArrayAdd($entityIds, $oInput.innertext)
		 $oMod = $oMod + 1
	  EndIf
   Next
   _ArrayDelete($entityIds, "0-9")

   For $numOfResults = 1 to UBound($entityIds)
	  ; Opening default browser to download all of the images (Have your default browser Google Chrome so that downloads automatically begin
	  ShellExecute("https://earthexplorer.usgs.gov/download/3411/" & $entityIds[$numOfResults-1] & "/STANDARD/EE")
	  Sleep(100)
   Next
   Local $entityIds[10] = []

   EndFunc

Func seeNextPage()

   For $numOfPage = 2 to 100
	  $id = String($numOfPage) & "_3411"
	  $nextLink = _IEGetObjById($oIE, $id)
	  _IEAction($nextLink, "click")
	  If @error Then
		 ExitLoop
	  Else
	  Sleep(2000)
	  Call("downloadResults")
	  Sleep(2000)
	  EndIf
   Next

   EndFunc

Func _IEGetObjByClass(ByRef $o_object, $s_Class, $i_index = 0)
    If Not IsObj($o_object) Then
        __IEErrorNotify("Error", "_IEGetObjByClass", "$_IEStatus_InvalidDataType")
        SetError($_IEStatus_InvalidDataType, 1)
        Return 0
    EndIf
    ;
    If Not __IEIsObjType($o_object, "browserdom") Then
        __IEErrorNotify("Error", "_IEGetObjByClass", "$_IEStatus_InvalidObjectType")
        SetError($_IEStatus_InvalidObjectType, 1)
        Return 0
    EndIf
    ;
    Local $i_found = 0
    ;
    $o_tags = _IETagNameAllGetCollection($o_object)
    For $o_tag In $o_tags
        If String($o_tag.className) = $s_Class Then
            If ($i_found = $i_index) Then
                SetError($_IEStatus_Success)
                Return $o_tag
            Else
                $i_found += 1
            EndIf
        EndIf
    Next
    ;
    __IEErrorNotify("Warning", "_IEGetObjByClass", "$_IEStatus_NoMatch", $s_Class)
    SetError($_IEStatus_NoMatch, 2)
    Return 0
EndFunc   ;==>_IEGetObjByClass

