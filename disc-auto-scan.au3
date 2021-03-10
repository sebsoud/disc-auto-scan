#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Version=Beta
#AutoIt3Wrapper_Icon=disc-auto-scan.ico
#AutoIt3Wrapper_Res_Fileversion=0.2.0.0
#AutoIt3Wrapper_Res_ProductVersion=0.2
#AutoIt3Wrapper_Res_LegalCopyright=Sébastien Soudan 2021
#AutoIt3Wrapper_Res_Language=1036
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****


; by seb

#include <WindowsConstants.au3>
#include <ColorConstants.au3>
#include <Timers.au3>
#include <array.au3>
#include <WinAPISysWin.au3>
#include <WinAPIGdi.au3>
#include <Misc.au3>
#include <Date.au3>
#include <GDIPlus.au3>
#include <ClipBoard.au3>
#include <GuiEdit.au3>
#include <GuiMenu.au3>

#include <WinAPIsysinfoConstants.au3>

#include <GUIConstantsEx.au3>
#include <StringConstants.au3>
#include <ProgressConstants.au3>
#include <MenuConstants.au3>

#include <GuiComboBox.au3>
#include <GuiButton.au3>
#include <GuiTab.au3>

#include <MsgBoxConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <ButtonConstants.au3>
#include <StaticConstants.au3>

#include "drivestray.au3" ; special functions For drives tray


Const $PRF_CHECKVISIBLE = 0x1; Draws the window only if it is visible
Const $PRF_ERASEBKGND = 0x8 ; Erases the background before drawing the window
Const $PRF_CLIENT = 0x4 ; Draw the window's client area.
Const $PRF_NONCLIENT = 0x2 ; Draw the window's Title area.
Const $PRF_CHILDREN = 0x10; Draw all visible child windows.
Const $PRF_OWNED = 0x20 ; Draw all owned windows.

;Global Const $COLOR_SCAN_RUNNING = 0xC7E9FF
Global Const $COLOR_SCAN_RUNNING = 0xB1FFFC
;Global Const $COLOR_NO_SCAN_RUNNING = 0xFFFCF2
Global Const $COLOR_NO_SCAN_RUNNING = 0xFDFCFC


Global Const $COLOR_APP_READY = 0x018501
Global Const $COLOR_APP_BUSY = 0xEE0202

Global Const $sDiscAutoScanRelease = "Disc Auto Scan 1.0"

Const $DBT_DEVICEARRIVAL = "0x00008000"
Const $DBT_DEVICEREMOVECOMPLETE = "0x00008004"


; VSO Inspector window state : values for $aVSOWinStates
Global Const $VSO_WIN_NO_WINDOW = 0
Global Const $VSO_WIN_IDLE = 1
Global Const $VSO_WIN_SCANNING = 2

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; resources values for VSO Inspector window
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

Global Const $VSOScanButton = "TButton4"
Global Const $VSOPropPageControl = "TPageControl1"
Global Const $VSODrivesComboBox = "TComboBox1"
Global Const $VSOFilesTestCheckBox = "TJvCheckBox1"
Global Const $VSOSaveReportButton = "TButton2"

; text messages
Global Const $sReadyForDisc_DetailsFR = "Vous pouvez insérer UN disque puis ATTENDRE que le scan ait été lancé"
Global Const $sReadyForDisc_DetailsEN = "You can insert ONE disc then WAIT that scan is launched"

Global Const $sReadyForDiscFR = "PRET à traiter UNE insertion de disque"
Global Const $sReadyForDiscEN = "READY for processing ONE disc insertion"

Global Const $sBusyTreatingDiscFR = "OCCUPE, lancement scan en cours..."
Global Const $sBusyTreatingDiscEN = "BUSY, scan launching..."

Global Const $sScanLaunchedFR = "-     Scan en cours... Vous pouvez reprendre la main sur Windows"
Global Const $sScanLaunchedEN = "-     Scan on going... You can again work in Windows"


Global $bShowAllVSoWindows = True ; if True, next call to DoVSOWinMinimizeAll() will show all windows. If False, it will minimize them
Global Const $sShowAllVSO_FR = "Montrer tous VSO Inspector"
Global Const $sMinimizeAllVSO_FR = "Minimiser tous VSO Inspector"
Global Const $sShowAllVSO_EN = "Show all VSO Inspector"
Global Const $sMinimizeAllVSO_EN = "Minimize all VSO Inspector"

;;;;;
Global Const $VSOWInPercentsStatusWidth = 198 ; pixels width for percentages status (TPanel1 Control)
Global Const $PercentsStatusWidth = 170 ; pixels width for percentages status control, in this app, for each drive


Global Const $DetailedVSOStatusUpdateInterval = 3 * 1000 ; to add in .ini?


$sIniSection="Configuration"


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; read values from ini file and registry (configuration)
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

Global Const $sIniFile = ".\disc-auto-scan.ini"
Global $DriveLetter = ""

; language (EN or FR)
Global $sLang = IniRead($sIniFile, $sIniSection, "Lang", "EN")

; log debug info
$sDebugInfo = IniRead($sIniFile, $sIniSection, "DebugInfo", "N")
Global $bDebugInfo = ($sDebugInfo == "Y")

Global $sVSOLocaleFile = RegRead("HKEY_CURRENT_USER\SOFTWARE\Vso\Inspector", "locale_file")
If ($sVSOLocaleFile == "") Then ; registry can be empty if VSO Inspector installed but user didn't manually choose a language in menu
	$sVSOLocaleFile = "INS_original.ini" ; default texts
EndIf

; global drives list
Global $sDriveListWithSpaces = IniRead($sIniFile, $sIniSection, "DriveList", "D")
Global $sDriveList = ""

; sound to be played at end/cancel of scan
Global $sEndScanNotification = IniRead($sIniFile, $sIniSection, "EndScanNotification", ".\tada.wav")

; VSO configuration
;;;;;;;;;;;;;;;;;;;;
; if speficied in .ini (not empty), path and exe are used (priority)
Global $VSOInspectorPath = IniRead($sIniFile, $sIniSection, "VSOInspectorPath", "")
Global $VSOInspectorApp = IniRead($sIniFile, $sIniSection, "VSOInspectorApp", "")
Global $VSOInspectorCmd = $VSOInspectorPath & $VSOInspectorApp

; if empty in .ini then we search in registry
If ($VSOInspectorPath == "" Or $VSOInspectorApp == "") Then
	; to get VSO path: use shortcut created by installer, in start menu
	; located by default in: C:\ProgramData\Microsoft\Windows\Start Menu\Programs\vso\tools\VSO Inspector
	Local $sShortcutPath = @ProgramsCommonDir & '\vso\tools\VSO Inspector\VSO Inspector.lnk'
	Local $aShortcutDetails = FileGetShortcut($sShortcutPath)
	If Not @error Then
		$VSOInspectorCmd = $aShortcutDetails[0] ; full path shortcut target (including exe)
		$VSOInspectorPath = $aShortcutDetails[1] ; working directory
	Endif

EndIf

; check configuration
If (Not FileExists($VSOInspectorCmd)) Then
	; configuration error
	If ($sLang == "Fr") Then
		$sErrorMsg =  "ERROR, les paramètres pour lancer VSO Inspector sont invalides. " & @LF & "Vous devez spécifier correctement VSOInspectorPath= et VSOInspectorApp= dans disc-auto-scan.ini"
	Else
		$sErrorMsg =  "ERROR, parameters for launching VSO Inspector are invalid. " & @LF & "You must correctly specify VSOInspectorPath= and VSOInspectorApp= in disc-auto-scan.ini"
	EndIf
	MsgBox($MB_SYSTEMMODAL, "Configuration", $sErrorMsg)
	Exit
EndIf


; increase of VSOInspector window width at opening (pixels, max 200. 0 is possible)
Global $iVSOInspectorWindowWidthIncrease = 0 ;
Global Const $sVSOInspectorWindowWidthIncrease = IniRead($sIniFile, $sIniSection, "VSOInspectorWindowWidthIncrease", "0")
If (StringIsInt($sVSOInspectorWindowWidthIncrease)) Then
	$iVSOInspectorWindowWidthIncrease = Number($sVSOInspectorWindowWidthIncrease)
	If ($iVSOInspectorWindowWidthIncrease > 200) Then
		$iVSOInspectorWindowWidthIncrease = 200
	EndIf
EndIf

; title of Inspector window. Used for window recognition
Global Const $sVSOInspectorWindow = "VSO Inspector"

; text of "save scan results" VSO Inspector window. Used for window recognition
Global Const $SaveScanResultsTextFR = "Enregistre les résultats du scan..."
Global Const $SaveScanResultsTextEN = "Save scan results..."


; MPC-HC configuration
;;;;;;;;;;;;;;;;;;;;;;;
; if speficied in .ini (not empty), path and exe are used (priority)
Global $sMPCHC = IniRead($sIniFile, $sIniSection, "MPCHC", "")

; if empty in .ini then we search in registry
If ($sMPCHC == "") Then
	; we check entries corresponding to x64 and x32 version

	; use shortcut created by installer, in start menu
	; location by default is: C:\ProgramData\Microsoft\Windows\Start Menu\Programs\MPC-HC x64\MPC-HC x64.lnk for x64
	; or C:\ProgramData\Microsoft\Windows\Start Menu\Programs\MPC-HC\MPC-HC.lnk for x86
	Local $sShortcutPath = @ProgramsCommonDir & '\MPC-HC x64\MPC-HC x64.lnk' ; try x64 version first
	Local $aShortcutDetails = FileGetShortcut($sShortcutPath)
	If Not @error Then
		$sMPCHC = $aShortcutDetails[0] ; full path shortcut target (including exe)
	Else
		$sShortcutPath = @ProgramsCommonDir & '\MPC-HC\MPC-HC.lnk' ; try x86 version
		$aShortcutDetails = FileGetShortcut($sShortcutPath)
		If Not @error Then
			$sMPCHC = $aShortcutDetails[0] ; full path shortcut target (including exe)
		EndIf
	Endif

EndIf

; check configuration
If (Not FileExists($sMPCHC)) Then
	; configuration error
	If ($sLang == "Fr") Then
		$sErrorMsg =  "ERROR, les paramètres pour lancer MPC-HC sont invalides. " & @LF & "Vous devez spécifier correctement MPCHC= dans disc-auto-scan.ini"
	Else
		$sErrorMsg =  "ERROR, parameters for launching MPC-HC are invalid. " & @LF & "You must correctly specify MPCHC= in disc-auto-scan.ini"
	EndIf
	MsgBox($MB_SYSTEMMODAL, "Configuration", $sErrorMsg)
	Exit
EndIf


;;;;;;;;;;;;;;;;;;;;;;
; beginning of program
;;;;;;;;;;;;;;;;;;;;;;

; initialization for tray functions, see drivestray.au3
InitForTrayFunctions()


$sDriveList = StringStripWS($sDriveListWithSpaces, $STR_STRIPALL) ; remove white spaces
$sDriveList = GetValidDrives($sDriveList) ; check drives are cd/dvd
Local $iDrivesNumber = StringLen($sDriveList)


; initial window position. If possible, put at bottom of screen, so for Fullhd or greater resolution, the VSO Inspector windows will be rather at bottom of screen
Global $iDiscAutoScanWinPosX = 50
Global $iDiscAutoScanWinPosY = 100
Global $iDiscAutoScanWinWidth = 652+ $PercentsStatusWidth + 10


; Y coordinate for status controls (used below several times)
$iStatusY = 3 + ($iDrivesNumber * 40) + 10 ; next Y coordinate. Take into consideration controls created for each drive (dynamic)

; special computation for window height, according to drives number
Global $iDiscAutoScanWinHeight = $iStatusY + 150
If ($bDebugInfo) Then
	$iDiscAutoScanWinHeight += 100 ; increase since log window will have much more many lines
EndIf

If (@DesktopHeight > 300 + $iDiscAutoScanWinHeight) Then  ; if "big" screen, move window at bottom of screen
	$iDiscAutoScanWinPosY = @DesktopHeight - $iDiscAutoScanWinHeight - 50
EndIf


; window creation
;;;;;;;;;;;;;;;;;
Global $hDiscAutoScanGUI = GUICreate("Disc auto scan", $iDiscAutoScanWinWidth, $iDiscAutoScanWinHeight, $iDiscAutoScanWinPosX, $iDiscAutoScanWinPosY)


; controls creation
; -----------------

;$idExit = GUICtrlCreateButton($sLang == "FR" ? "Quitter" : "Exit", $iDiscAutoScanWinWidth - 70 , 10, 60, 25)
$idExit = -1

Global $iParametersButtonId = GUICtrlCreateButton("", $iDiscAutoScanWinWidth - 71 , 9, 29, 29, $BS_ICON)
GUICtrlSetImage($iParametersButtonId, "icons\25\parameters-25.ico")

Global $iHelpButtonId = GUICtrlCreateButton("", $iDiscAutoScanWinWidth - 37 , 9, 29, 29, $BS_ICON)
GUICtrlSetImage($iHelpButtonId, ".\icons\25\help-25.ico")



;;;;;;;;;;;;;;;;;;
; DYNAMIC section: sets of controls created, one set per drive. We have global arrays for memorizing needed data
Global $aVSOWinHandles[$iDrivesNumber]  ; array of the VSO Inspector windows handles, same order as $sDriveList
Global $aVSOWinStates[$iDrivesNumber]  ; array of states of VSO Inspector windows. Updated by UpdateVSOWinStatus()
; initialization
For $i = 0 To $iDrivesNumber - 1
	$aVSOWinStates[$i] = $VSO_WIN_NO_WINDOW
Next

; declaration of arrays for the dynamic buttons
Global $aVSOWinMinimizeButtonsIds[$iDrivesNumber] ; array containing the buttons ids, same order as $sDriveList
Global $aVSOWinCloseButtonsIds[$iDrivesNumber] ; array containing the buttons ids, same order as $sDriveList
Global $aVSOWinEndSaveButtonsIds[$iDrivesNumber] ; array containing the buttons ids, same order as $sDriveList
Global $aTrayButtonsIds[$iDrivesNumber] ; array containing the buttons ids, same order as $sDriveList
Global $aDiscSizeLabelIds[$iDrivesNumber] ; array containing the label ids, same order as $sDriveList
Global $aDiscLabelLabelIds[$iDrivesNumber] ; array containing the label ids, same order as $sDriveList
Global $aProgressBarIds[$iDrivesNumber] ; array containing the progress bar ids, same order as $sDriveList
Global $aPercentsStatusLabelIds[$iDrivesNumber] ; array containing the labels controls ids, same order as $sDriveList
Global $aPercentsStatusLabelIds[$iDrivesNumber] ; array containing the labels controls ids, same order as $sDriveList

Global $aCurrentSpeedIds[$iDrivesNumber] ; array containing the current speed controls ids, same order as $sDriveList

; call function which creates all the sets of buttons
CreateButtonsSets($sDriveList, $aVSOWinHandles, $aVSOWinMinimizeButtonsIds, $aVSOWinCloseButtonsIds, _
		$aVSOWinEndSaveButtonsIds, $aTrayButtonsIds, $aProgressBarIds, $aPercentsStatusLabelIds, $aCurrentSpeedIds)
;;;;;;;;;;;;;;;;;;;;


; button removed on 210214: not compatible with percentages statuses updates, because for that to work VSO windows must not be minimized
; as DoVSOWinMinimizeAll()  was needing to set WinSetOnTop, it's the mess with other windows (z-order Windows management)
; display/minimize all button: below drives buttons sets
;$idDisplayAllButton = GUICtrlCreateButton("", 10, $iStatusY, 150, 27)
;UpdateShowAllVSoWindowsButtonLabel($idDisplayAllButton)

; app status:
GUICtrlCreateGroup("", 10, $iStatusY - 5, 280, 32)
Global Const $IdAppStatus = GUICtrlCreateLabel("", 22, $iStatusY + 5, 280-40, 20, $SS_LEFTNOWORDWRAP )

; app status detail:
GUICtrlCreateGroup("", 300, $iStatusY - 5, 380, 32)
Global Const $IdAppStatusDetail = GUICtrlCreateLabel("", 305, $iStatusY + 5, 380-10, 20, $SS_LEFTNOWORDWRAP )

; app label (incl release)
If ($iDrivesNumber > 0) Then ; handle special case where no drive is configured: avoid overlap with buttons
	$IdAppLabel = GUICtrlCreateLabel($sDiscAutoScanRelease, $iDiscAutoScanWinWidth - 130, $iStatusY + 5, 118, 15, $SS_RIGHT)
EndIf

; edit control for scan ended
Local $Height = $iDiscAutoScanWinHeight - $iStatusY - 45
Local $sLogScanEndedGroup = "Scans ended/canceled" ; defaut: EN language
If ($sLang == "FR") Then
	$sLogScanEndedGroup = "Scans terminés/annulés"
EndIf
GUICtrlCreateGroup($sLogScanEndedGroup, 10, $iStatusY + 35, 280, $Height)
Global $iLogScanEndedEdit = GUICtrlCreateEdit("", 13, $iStatusY + 50, 280 - 6, $Height -19 , BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $WS_VSCROLL, $ES_READONLY, $ES_MULTILINE))


; edit control for log
Local $Height = $iDiscAutoScanWinHeight - $iStatusY - 45
GUICtrlCreateGroup("Log", 300, $iStatusY + 35, $iDiscAutoScanWinWidth -300 -10, $Height )
Global $idLogEdit = GUICtrlCreateEdit("", 303, $iStatusY + 50, $iDiscAutoScanWinWidth - 300 - 16, $Height -19 , BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $WS_VSCROLL, $ES_READONLY, $ES_MULTILINE))


;;;;;;;;;;;;;;;;;

Global Const $sCancelScanText = GetLocalForVSOInsp($sVSOLocaleFile, "0015_INCODE", "Cancel")
Global Const $sSaveScanResultsWinText = GetLocalForVSOInsp($sVSOLocaleFile, "0049_INCODE", "Save scan results...")


;;;;;;;;;;;;;
; inform user that program is ready for new disc
If ($bDebugInfo) Then
	AddTrace($sReadyForDisc_DetailsFR, $sReadyForDisc_DetailsEN)
EndIf

SetAppStatusReady()

; global variables
;;;;;;;;;;;;;;;;;;;
Global $bScanToLaunch = False

_GDIPlus_StartUp() ; initialize library

; ---------------------------------

; FOR DEBUG ONLY
; FOR DEBUG ONLY
; FORCE direct test (avoiding to have to insert a disc for detection)
;PerformScan("E")
;PerformScan("G")
;PerformScan("F")

;_GDIPlus_ShutDown()
;Exit

;---------------------------------


; 210303: no more menu items added, directly managed with 2 buttons
#comments-start

; manage App menu
Global Enum $iOptionsMenuId = 1000, $iOptionsHelpId

Local $hMenu = _GUICtrlMenu_GetSystemMenu($hDiscAutoScanGUI)

; we insert new menu items after separator
Local $iCount = _GUICtrlMenu_GetItemCount($hMenu)
For $iMenuIndex = 0 To $iCount - 1
	If (_GUICtrlMenu_GetItemText($hMenu, $iMenuIndex) == "") Then
		ExitLoop
	EndIf
;        AddDebugTrace("Item " & $iI & " text ......: " & )
Next

_GUICtrlMenu_InsertMenuItem ($hMenu, $iMenuIndex, "Options", $iOptionsMenuId)
_GUICtrlMenu_InsertMenuItem ($hMenu, $iMenuIndex+1, $sLang == "FR" ? "Aide" : "Help", $iOptionsHelpId)
#comments-end

; options dialog
Global $hOptions = -1
Global $idOptionsOk = -1
Global $idOptionsCancel = -1
Global $aOptionsDialogCtrlIds[20][2] ; array containing all the control ids of the options dialog. each element is a duo <String, controlid>
Global $iOptionsDialogCtrlIdsNumber = 0

; help dialog
Global $hHelp = -1
Global $idHelpClose = -1


; needed for menu commands management
;$bOk= GUIRegisterMsg($WM_SYSCOMMAND , "WM_SYSCOMMAND_Handler")

; register messages handlers
GUIRegisterMsg($WM_DEVICECHANGE, "DeviceChange")

; timer for status update, every 1 second
Global $iPreviousTime = 0 ; memorize previous time
_Timer_SetTimer($hDiscAutoScanGUI, 1000, "UpdateVSOWinStatus")

GUISetState() ; display the GUI


;;;;;;;;;;;
; main loop

Do
	$GuiMsg = GUIGetMsg()

	If ($bScanToLaunch) Then
		;		WinSetState($hDiscAutoScanGUI, "", @SW_DISABLE ) ; prevent user click
		PerformScan($DriveLetter)
		;		WinSetState($hDiscAutoScanGUI, "", @SW_ENABLE )
		_WinAPI_RedrawWindow($hDiscAutoScanGUI) ; must force redraw, because some paint message may have been ignored
		; and if this gui window was in foreground, some part may not be anymore in foreground...
		$bScanToLaunch = False
	EndIf


	; treat click on VSO Inspector windows "restore/minimize" buttons
	$iButtonIndex = _ArraySearch($aVSOWinMinimizeButtonsIds, $GuiMsg)
	If ($iButtonIndex > -1) Then
		DoVSOWinMinimize($iButtonIndex)
	EndIf ; If ($iButtonIndex > -1)

	; treat click on VSO Inspector windows  "close" buttons
	If ($iButtonIndex == -1) Then    ; button was not identified yet
		$iButtonIndex = _ArraySearch($aVSOWinCloseButtonsIds, $GuiMsg)
		If ($iButtonIndex > -1) Then
			DoVSOWinClose($iButtonIndex)
		EndIf ; If ($iButtonIndex > -1)
	EndIf ; If ($iButtonIndex == -1

	; treat click on VSO Inspector windows  "End scan and save" buttons
	If ($iButtonIndex == -1) Then    ; button was not identified yet
		$iButtonIndex = _ArraySearch($aVSOWinEndSaveButtonsIds, $GuiMsg)
		If ($iButtonIndex > -1) Then
			DoVSOWinEndSave($iButtonIndex)
		EndIf ; If ($iButtonIndex > -1)
	EndIf ; If ($iButtonIndex == -1

	; treat click on drives tray buttons (open/close tray)
	If ($iButtonIndex == -1) Then    ; button was not identified yet
		$iButtonIndex = _ArraySearch($aTrayButtonsIds, $GuiMsg)
		If ($iButtonIndex > -1) Then
			DoTrayOpenClose($iButtonIndex)
		EndIf ; If ($iButtonIndex > -1)
	EndIf ; If ($iButtonIndex == -1)

	; treak click on options button
	If ($GuiMsg == $idOptionsOk Or $GuiMsg == $idOptionsCancel) Then
		$bCloseOptionsDialog = True ; by default, for cancel button
		If ($GuiMsg == $idOptionsOk) Then
			$bCloseOptionsDialog = ValidateOptionsDialog($aOptionsDialogCtrlIds, $iOptionsDialogCtrlIdsNumber)
		EndIf
		If $bCloseOptionsDialog Then

			GUIDelete($hOptions) ; close options window

			GUISetState (@SW_ENABLE, $hDiscAutoScanGUI) ; re-enable main window
			WinActivate($hDiscAutoScanGUI)
		EndIf
	ElseIf ($GuiMsg == $iParametersButtonId) Then
		; options dialog
		$hOptions = CreateOptionsDialog($idOptionsOk, $idOptionsCancel, $aOptionsDialogCtrlIds, $iOptionsDialogCtrlIdsNumber)

		GUISetState(@SW_SHOW, $hOptions) ; show options window
		GUISetState (@SW_DISABLE, $hDiscAutoScanGUI) ; disable main window

	ElseIf ($GuiMsg == $iHelpButtonId) Then
		; help dialog
		; options dialog
		$hHelp = CreateHelpDialog()

		GUISetState(@SW_SHOW, $hHelp) ; show options window

	ElseIf ($GuiMsg == $idHelpClose) Then
		GUIDelete($hHelp) ; close help window
		$hHelp = -1
	Else
		; treat messages linked to options dialog controls
		TreatOptionsDialogActions($GuiMsg, $aOptionsDialogCtrlIds, $iOptionsDialogCtrlIdsNumber)
	EndIf

Until ($GuiMsg == $GUI_EVENT_CLOSE Or $GuiMsg == $idExit)

_GDIPlus_ShutDown() ; close library

Exit

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;; FUNCTIONS ;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;



; 210303: no more menu items added, directly managed with 2 buttons
#comments-start

; treat menu commands
Func WM_SYSCOMMAND_Handler($hWnd, $iMsg, $wParam, $lParam)
	Local $iID = _WinAPI_LoWord($wParam)
    Switch $iID
        Case $iOptionsMenuId
			$hOptions = CreateOptionsDialog($idOptionsOk, $idOptionsCancel, $aOptionsDialogCtrlIds, $iOptionsDialogCtrlIdsNumber)

			GUISetState(@SW_SHOW, $hOptions) ; show options window
			GUISetState (@SW_DISABLE, $hDiscAutoScanGUI) ; disable main window
        Case $iOptionsHelpId
    EndSwitch
    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_COMMAND
#comments-end



; update $idDisplayAllButton label
Func UpdateShowAllVSoWindowsButtonLabel($idDisplayAllButton)
	Local $sButtonLabel = ""

	If ($sLang == "FR") Then
		$sButtonLabel = $bShowAllVSoWindows ? $sShowAllVSO_FR : $sMinimizeAllVSO_FR
	Else
		$sButtonLabel = $bShowAllVSoWindows ? $sShowAllVSO_EN : $sMinimizeAllVSO_EN
	EndIf

	; change label
	GUICtrlSetData($idDisplayAllButton, $sButtonLabel)
EndFunc   ;==>UpdateShowAllVSoWindowsButtonLabel



; Function: adding trace in window
Func AddTrace($FrMsg, $EnMsg)

	Local $iText = $sLang == "FR" ? $FrMsg : $EnMsg

	; first, ensure caret is at end of edit
	Local $hWnd = ControlGetHandle($hDiscAutoScanGUI, "", $idLogEdit)
    Local $hTextLengh
    $hTextLengh = _SendMessage ($hWnd, $WM_GETTEXTLENGTH, 0, 0);

    _SendMessage ($hWnd, $EM_SETSEL, $hTextLengh, $hTextLengh);

	GUICtrlSetData($idLogEdit,$iText & @CRLF, 1);

EndFunc   ;==>AddTrace


; Function: adding trace in window
Func AddScanEndedTrace($ScanEndedTrace)

	; first, ensure caret is at end of edit
	Local $hWnd = ControlGetHandle($hDiscAutoScanGUI, "", $idLogEdit)
    Local $hTextLengh
    $hTextLengh = _SendMessage ($hWnd, $WM_GETTEXTLENGTH, 0, 0);

    _SendMessage ($hWnd, $EM_SETSEL, $hTextLengh, $hTextLengh);

	GUICtrlSetData($iLogScanEndedEdit,$ScanEndedTrace & @CRLF, 1);
EndFunc


; helper function, retrieve VSO Inspector window handle, after checking that it is still valid
; if handle is not valid, it returns 0 instead of handle, AND, if needed, $aVSOWinHandles[$iDriveIndex] is set to 0 for subsequent uses
Func GetValidVSOWinHandle($iDriveIndex)
	$hVSOWinHandle = 0

	If ($iDriveIndex >= 0 AND $iDriveIndex < UBound($aVSOWinHandles)) Then ; check $iDriveIndex for safety...
		$hVSOWinHandle = $aVSOWinHandles[$iDriveIndex]
	EndIf

	If ($hVSOWinHandle > 0) Then
		; check that the window still exist
		$WinWaitDelay = AutoItSetOption ("WinWaitDelay", 0) ; disable delay so there's no flickering
		Local $iWinState = WinGetState($hVSOWinHandle)
		If ($iWinState == 0) Then
			; window doesn't exist anymore -> reset handle to 0
			$aVSOWinHandles[$iDriveIndex] = 0
			$hVSOWinHandle = 0
		EndIf
		AutoItSetOption($WinWaitDelay)
	EndIf

	Return $hVSOWinHandle
EndFunc   ;==>GetValidVSOWinHandle

; show/minimize all VSO Inspector windows
Func DoVSOWinMinimizeAll()
	; disable temporary animation, so it is quicker
	Local $bMinAnimation = GetMinAnimate()
	SetMinAnimate(False)

	$WinWaitDelay = AutoItSetOption ("WinWaitDelay", 0) ; disable delay so there's no flickering

	$iIndexMax = UBound($aVSOWinHandles) - 1
	Local $iButtonIndex = 0
	For $iButtonIndex = 0 To $iIndexMax ; loop on all VSO Insp windows
		Local $hVSOWinHandle = GetValidVSOWinHandle($iButtonIndex)
		If ($hVSOWinHandle > 0) Then
			; NOTE: GetValidVSOWinHandle() guarantees that WinGetState will succeed
			Local $iWinState = WinGetState($hVSOWinHandle)
			Local $bIsWinMinimized = (BitAND($iWinState, $WIN_STATE_MINIMIZED) == $WIN_STATE_MINIMIZED)

			If ($bShowAllVSoWindows) Then
				; show window if not yet
				If $bIsWinMinimized Then
					WinSetState($hVSOWinHandle, "", @SW_RESTORE)
				EndIf
				WinSetOnTop($hVSOWinHandle, "", $WINDOWS_ONTOP)
			Else
				; minimize window if not yet
				If ($bIsWinMinimized == False) Then
					WinSetState($hVSOWinHandle, "", @SW_MINIMIZE)
				EndIf
				WinSetOnTop($hVSOWinHandle, "", $WINDOWS_NOONTOP)
			EndIf
		EndIf ; If ($hVSOWinHandle > 0)
	Next ; For $iButtonIndex = 0 To $iIndexMax -1

	; restore values
	SetMinAnimate($bMinAnimation)
	AutoItSetOption ("WinWaitDelay", $WinWaitDelay) ;

	; change behaviour for next call (global variable)
	$bShowAllVSoWindows = Not ($bShowAllVSoWindows)
	; update button label
;	UpdateShowAllVSoWindowsButtonLabel($idDisplayAllButton)
EndFunc   ;==>DoVSOWinMinimizeAll

Func DoVSOWinMinimize($iButtonIndex)
	Local $hVSOWinHandle = GetValidVSOWinHandle($iButtonIndex)
	If ($hVSOWinHandle > 0) Then
		; NOTE: GetValidVSOWinHandle() guarantees that WinGetState will succeed
		Local $iWinState = WinGetState($hVSOWinHandle)
		If (BitAND($iWinState, $WIN_STATE_MINIMIZED) == $WIN_STATE_MINIMIZED) Then
			WinSetState($hVSOWinHandle, "", @SW_RESTORE)
		Else
			WinSetState($hVSOWinHandle, "", @SW_MINIMIZE)
		EndIf
	EndIf ; If ($hVSOWinHandle > 0)
EndFunc   ;==>DoVSOWinMinimize

Func DoVSOWinClose($iButtonIndex)
	; close window
	$hVSOWinHandle = GetValidVSOWinHandle($iButtonIndex)
	If ($hVSOWinHandle > 0) Then
		; if window is minimized, restore it first
		Local $iWinState = WinGetState($hVSOWinHandle)
		If (BitAND($iWinState, $WIN_STATE_MINIMIZED) == $WIN_STATE_MINIMIZED) Then
			WinSetState($hVSOWinHandle, "", @SW_RESTORE)
		EndIf

		; to be sure the window confirmation is on top...
		WinActivate ($hVSOWinHandle)

		If ($aVSOWinStates[$iButtonIndex] == $VSO_WIN_SCANNING) Then
		EndIf

		If (WinClose($hVSOWinHandle) == 0) Then
			; don't update variables here, it will be done by UpdateVSOWinStatus()
		Else
			;; log error to add
		EndIf
	EndIf ; If ($hVSOWinHandle > 0)
EndFunc   ;==>DoVSOWinClose

Func DoVSOWinEndSave($iButtonIndex)
	Local $hVSOWinHandle = GetValidVSOWinHandle($iButtonIndex)

	If ($hVSOWinHandle > 0) Then
		; first, show window in case it's minimized
		; NOTE: GetValidVSOWinHandle() guarantees that WinGetState will succeed
		Local $iWinState = WinGetState($hVSOWinHandle)
		If (BitAND($iWinState, $WIN_STATE_MINIMIZED) == $WIN_STATE_MINIMIZED) Then
			WinSetState($hVSOWinHandle, "", @SW_RESTORE)
			Sleep(100)
		EndIf

		; now, check if scan is running. If yes, cancel it
		$sScanButtonText = ControlGetText($hVSOWinHandle, "", $VSOScanButton)
		$bIsScanRunning = ($sScanButtonText == $sCancelScanText)
		If ($bIsScanRunning) Then
			; ask for end of scan
			ControlClick($hVSOWinHandle, "", $VSOScanButton)

			; it may take some time to stop scan, in case media has errors and drive is spinning slowly
			Local $iTotalWait = 0
			While 1
				Sleep(500)
				$iTotalWait += 500
				$sScanButtonText = ControlGetText($hVSOWinHandle, "", $VSOScanButton)
				$bIsScanRunning = ($sScanButtonText == $sCancelScanText)
				If ($bIsScanRunning == False) Then
					ExitLoop
				EndIf
				If ($iTotalWait >= 20000) Then ; wait for 20 seconds maximum (avoid infinite loop)
					ExitLoop
				EndIf
			WEnd
		EndIf

		; now save report ($VSOSaveReport button)
		; check that button is available
		$isSaveEnabled = ControlCommand($hVSOWinHandle, "", $VSOSaveReportButton, "IsEnabled")
		If ($isSaveEnabled) Then
			; save report
			ControlClick($hVSOWinHandle, "", $VSOSaveReportButton)

			If (WinWaitActive("[CLASS:#32770]", "", 5)) Then ; save as window, max 5 secs
				; send  report default filename
				$aDriveArray = StringSplit($sDriveList, "")
				Local $DriveLetter = $aDriveArray[$iButtonIndex + 1]

				Local $sDiscLabel = DriveGetLabel($DriveLetter & ":\")

				Local $fDiscSizeInGb = DriveSpaceTotal($DriveLetter & ":\") / 1024
				Local $sDiscSizeInGb = StringFormat("%.2f", $fDiscSizeInGb)

				Local $sDefaultFileName = $sDiscLabel & "-" & $sDiscSizeInGb & ".txt"

				Send($sDefaultFileName)
			EndIf
		Else
			; log error to add: cannot save report and close window
		EndIf ; If ($isSaveEnabled) Then
	EndIf ; If ($hVSOWinHandle == 0)
EndFunc   ;==>DoVSOWinEndSave

Func DoTrayOpenClose($iButtonIndex)
	$aDriveArray = StringSplit($sDriveList, "")
	Local $DriveTrayLetter = $aDriveArray[$iButtonIndex + 1]

	If ($aVSOWinStates[$iButtonIndex] == $VSO_WIN_SCANNING) Then
		; drive is locked during scanning -> disable open
		Return
	EndIf

	Local $DriveTrayError = ""
	Local $bIsDriveTrayOpen = IsDriveTrayOpen($DriveTrayLetter, $DriveTrayError)

	If ($DriveTrayError <> "") Then
		AddDebugTrace("ERROR: IsDriveTrayOpen([" & $DriveTrayLetter & "]) returned error:" &$DriveTrayError)
	EndIf

	If ($bIsDriveTrayOpen) Then
		CDTray($DriveTrayLetter & ":", $CDTRAY_CLOSED)
	Else
		CDTray($DriveTrayLetter & ":", $CDTRAY_OPEN)
	EndIf
EndFunc   ;==>DoTrayOpenClose

; Function: handler of $WM_DEVICECHANGE message (media inserted in drive notification)
Func DeviceChange($hWndGUI, $MsgID, $WParam, $LParam)
	If (($WParam == $DBT_DEVICEARRIVAL) Or ($WParam == $DBT_DEVICEREMOVECOMPLETE)) Then ; disc inserted or removed
		; must check if it's in a drive which is configured
		; Create a struct from $lParam which contains a pointer to a Windows-created struct.

		Local Const $tagDEV_BROADCAST_VOLUME = "dword dbcv_size; dword dbcv_devicetype; dword dbcv_reserved; dword dbcv_unitmask; word dbcv_flags"
		Local Const $DEV_BROADCAST_VOLUME = DllStructCreate($tagDEV_BROADCAST_VOLUME, $LParam)

		Local Const $DeviceType = DllStructGetData($DEV_BROADCAST_VOLUME, "dbcv_devicetype")
		Local Const $UnitMask = DllStructGetData($DEV_BROADCAST_VOLUME, "dbcv_unitmask")

		$DriveLetter = GetDriveLetterFromUnitMask($UnitMask) ; global variable

		Local $iDriveIndex = StringInStr($sDriveList, $DriveLetter)
		If ($iDriveIndex > 0) Then
			; Ok, it's a managed drive
			;			WinSetState($hDiscAutoScanGUI, "", @SW_DISABLE ) ; prevent user click

			If ($WParam == $DBT_DEVICEARRIVAL) Then
				$sDiscLabel = DriveGetLabel($DriveLetter & ":\")

				$fDiscSizeInGb = DriveSpaceTotal($DriveLetter & ":\") / 1024
				$sDiscSizeInGb = StringFormat("%.2f", $fDiscSizeInGb)
				$sDiscSizeInGb &= ($sLang == "FR") ? " Go" : " Gb"
				If ($bScanToLaunch) Then
					; scan of previous disc not finished! -> log warning message
					AddTrace("ATTENTION: insertion de disque [" & $sDiscLabel & "][" & $sDiscSizeInGb & "] dans le lecteur [" & $DriveLetter & "] TROP TOT, ne sera pas traité", _
							"WARNING: disc [" & $sDiscLabel & "][" & $sDiscSizeInGb & "] inserted TOO EARLY in drive [" & $DriveLetter & "], will not be treated")
					;							WinSetState($hDiscAutoScanGUI, "", @SW_ENABLE )
					Return
				Else
					$sTraceFR = "["  & $DriveLetter & "][" & $sDiscLabel & "][" & $sDiscSizeInGb & "] inséré, veuillez patienter..."
					$sTraceEN = "["  & $DriveLetter & "][" & $sDiscLabel & "][" & $sDiscSizeInGb & "] inserted, please wait..."
					; update status label
					SetAppStatusBusy($sTraceFR, $sTraceEN)

					If ($bDebugInfo) Then ; 210215 now that $IdAppStatusDetail does exist, insertion message goes in this control only, if debug trace is not active
						AddTrace("-> " & $sTraceFR, "-> " & $sTraceEN)
					EndIf

					; update disc infos
					GUICtrlSetData($aDiscSizeLabelIds[$iDriveIndex - 1], $sDiscSizeInGb)
					GUICtrlSetData($aDiscLabelLabelIds[$iDriveIndex - 1], $sDiscLabel)

				EndIf
				$bScanToLaunch = True
			Else
				; here we know $WParam == DBT_DEVICEREMOVECOMPLETE

				; empty disc informations
				GUICtrlSetData($aDiscSizeLabelIds[$iDriveIndex - 1], "")
				GUICtrlSetData($aDiscLabelLabelIds[$iDriveIndex - 1], "")

			EndIf ; If ($WParam == $DBT_DEVICEARRIVAL)

			;			WinSetState($hDiscAutoScanGUI, "", @SW_ENABLE )
		EndIf ; If ($iDriveIndex > 0) Then
	EndIf
EndFunc   ;==>DeviceChange

Func GetDriveLetterFromUnitMask($UnitMask)

	Local Const $Drives[26] = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
	Local $Count = 1
	Local $Pom = $UnitMask / 2

	While $Pom <> 0
		; $Pom = BitShift($Pom, 1)
		$Pom = Int($Pom / 2)
		$Count += 1
	WEnd

	If $Count >= 1 And $Count <= 26 Then
		Return $Drives[$Count - 1]
	Else
		Return SetError(-1, 0, '?')
	EndIf
EndFunc   ;==>GetDriveLetterFromUnitMask

; -----------------------------------


Func PerformScan($DriveLetter)
	SetAppStatusBusy()

	; wait enough time that disc is recognized by windows and check file system
	$iTotalWait = 0
	While 1
		Sleep(500)
		$iTotalWait += 500
		$sDriveFS = DriveGetFileSystem($DriveLetter & ":")
		If ($sDriveFS == $DT_CDFS Or $sDriveFS == $DT_UDF) Then  ; cd, dvd or bluray    NOTE:  $DT_UNDEFINED is returned if no disc or raw file system
			ExitLoop
		EndIf
		If ($iTotalWait >= 15000) Then ; wait for 15 seconds maximum (avoid infinite loop)
			ExitLoop
		EndIf
	WEnd

	; check if it's video DVD. In this case, MPC-HC must be launched to be sure CSS protection is removed
	Local $bIsVideoDVD = FileExists ($DriveLetter & ":\VIDEO_TS\VIDEO_TS.IFO")

	If ($bDebugInfo) Then
		AddTrace("       Debug: DVD video=" & $bIsVideoDVD, "       Debug: video DVD=" & $bIsVideoDVD)
	EndIf

	If ($bIsVideoDVD) Then
		; launch MPC-HC, so dvd is accessible for VSO Inspector if protected (CSS protection)
		;;;;;;;;;;;;;;;;
		Local $sMPCHCCmd = """" & $sMPCHC & """ " & $DriveLetter & ": /open /minimized /new"

		If ($bDebugInfo) Then
			AddTrace("       Debug: lancement de " & $sMPCHCCmd, "       Debug: Launching " & $sMPCHCCmd)
		EndIf

		Local $sMPCHCPid = Run($sMPCHCCmd)

		If ($sMPCHCPid == 0) Then
			AddTrace("ERREUR: Lancement de " & $sMPCHCCmd & " terminé avec @error=" & @error, _
					"ERROR: Launching " & $sMPCHCCmd & " terminated with @error=" & @error)
			SetAppStatusReady()
			Return
		EndIf

		Local $sMPCHCWindow = IniRead($sIniFile, $sIniSection, "MPCHCWindow", "Media Player Classic Home Cinema")

		If ($bDebugInfo) Then
			AddTrace("       Debug: attente fenêtre: " & $sMPCHCWindow, "       Debug: waiting for window: " & $sMPCHCWindow)
		EndIf

		Local $hMPCHC = WinWait("[CLASS:" & $sMPCHCWindow & "]", "", 10)

		; now wait that disc is opened (max 40 seconds, it can be quite slow on usb drives)
		; for this, we check the MPC-HC status evolution
		; it evolves from "" to "opening" to "stopped"
		; we cannot test directly hard-coded text because MPC-HC is multi-languages
		Local $sMPCStatusControl = IniRead($sIniFile, $sIniSection, "MPCStatusControl", "[CLASS:Static; INSTANCE:3]")

		$sMPCPreviousStatusText = ""
		$iTotalWait = 0
		While 1
			Sleep(1000)
			$iTotalWait += 1000

			$sMPCStatusText = ControlGetText($hMPCHC, "", $sMPCStatusControl) ; check current MPC-HC status

			If ($bDebugInfo) Then
				AddTrace("       Debug: état MPC-HC=" & $sMPCStatusText, "       Debug: MPC-HC status=" & $sMPCStatusText)
			EndIf

			If ($sMPCStatusText <> "") Then
				If ($sMPCStatusText <> $sMPCPreviousStatusText And $sMPCPreviousStatusText <> "") Then
					; ok, we can consider that disc is opened now
					ExitLoop
				EndIf
				$sMPCPreviousStatusText = $sMPCStatusText ; update
			EndIf
			If ($iTotalWait >= 40000) Then ; wait for 40 seconds maximum (avoid infinite loop)
				ExitLoop
			EndIf
		WEnd

		If ($bDebugInfo) Then
			AddTrace("       Debug: Fermeture de " & $sMPCHC, "       Debug: Closing " & $sMPCHC)
		EndIf

		; close MPC-HC
		ProcessClose($sMPCHCPid)
		; wait that it's closed
		ProcessWaitClose($sMPCHCPid, 10)

		If ($bDebugInfo) Then
			AddTrace("       Debug: " & $sMPCHC & " terminé", "       Debug: " & $sMPCHC & " terminated")
		EndIf
	EndIf ; If (bIsVideoDVD) Then

	; launch VSOInspector
	;;;;;;;;;;;;;;;;;;;;;;

	If ($bDebugInfo) Then
		AddTrace("       Debug: lancement de " & $VSOInspectorCmd, "       Debug: Launching " & $VSOInspectorCmd)
	EndIf
	Local $VSOInspectorPid = Run($VSOInspectorCmd, $VSOInspectorPath)

	If ($VSOInspectorPid == 0) Then
		AddTrace("ERREUR: Lancement de " & $VSOInspectorCmd & " terminé avec @error=" & @error, _
				"ERROR: Launching " & $VSOInspectorCmd & " terminated with @error=" & @error)
		SetAppStatusReady()
		Return
	EndIf

	Sleep(1000)


	; get handle of window
	Local $aVSOWindows = WinList($sVSOInspectorWindow)
	$iMax = UBound($aVSOWindows) ; get array size

	Local $hVSOWnd = 0
	For $i = 1 To $aVSOWindows[0][0]
		If $aVSOWindows[$i][0] <> "" Then
			$iPID2 = WinGetProcess($aVSOWindows[$i][1])
			If $iPID2 = $VSOInspectorPid Then
				$hVSOWnd = $aVSOWindows[$i][1]
				ExitLoop
			EndIf
		EndIf
	Next
	If $hVSOWnd == 0 Then
		AddTrace("ERREUR: Impossible de trouver la fenêtre avec le titre " & $sVSOInspectorWindow, _
				"ERROR: could not find window with title " & $sVSOInspectorWindow)
		SetAppStatusReady()
		Return
	EndIf

	; update window handle reference
	$aDriveArray = StringSplit($sDriveList, "")
	$iButtonIndex = _ArraySearch($aDriveArray, $DriveLetter) - 1
	If ($iButtonIndex > -1) Then
		$aVSOWinHandles[$iButtonIndex] = $hVSOWnd
	EndIf

	; move window, so if we have several drives the windows are spread on screen
	SpreadVSOWindow($hVSOWnd, $DriveLetter)

	; now activate the proper page for scan, in VSO
	; get handle for property pages control
	Local $hPPControl = ControlGetHandle($hVSOWnd, "", $VSOPropPageControl)

	If ($hPPControl == 0) Then
		AddTrace("ERREUR: Impossible de trouver le contrôle d'onglets " & $VSOPropPageControl & " @error=" & @error, _
				"ERROR: could not find Property page " & $VSOPropPageControl & " @error=" & @error)
		SetAppStatusReady()
		Return
	EndIf

	; tab index (start at zero!) for scan action (VSO Inspector)
	Local $ScanTabIndex = 2
	_GUICtrlTab_ClickTab($hPPControl, $ScanTabIndex)

	; choose proper drive in combo box
	Local $hDriveCombo = ControlGetHandle($hVSOWnd, "", $VSODrivesComboBox)

	If ($hPPControl == 0) Then
		AddTrace("ERREUR: Impossible de trouver le contrôle ComboBox " & $VSODrivesComboBox & " @error=" & @error, _
				"ERROR: could not find Combobox " & $VSODrivesComboBox & " @error=" & @error)
		SetAppStatusReady()
		Return
	EndIf

	Sleep(500)

	Opt("GUIDataSeparatorChar", ",") ; set seperator char to char we want to use
	Local $aComboList = StringSplit(_GUICtrlComboBox_GetList($hDriveCombo), ",")

	Local $ComboSearchedStr = "[" & $DriveLetter & "]" ; string that we search
	Local $ComboStrToSelect = ""
	Local $DriveIndex = -1 ;
	For $x = 1 To $aComboList[0]
		If StringInStr($aComboList[$x], $ComboSearchedStr) Then
			$ComboStrToSelect = $aComboList[$x]
			$DriveIndex = $x - 1
			ExitLoop
		EndIf
	Next

	If ($ComboStrToSelect <> "") Then

		;_GUICtrlComboBox_SetCurSel($hDriveCombo,$DriveIndex)
		_GUICtrlComboBox_SelectString($hDriveCombo, $ComboStrToSelect)
		Sleep(1000)

		ControlCommand($hVSOWnd, "", $VSODrivesComboBox, "SendCommandID", BitShift($CBN_SELCHANGE, -16))

		Sleep(3000)

		; it may take some time that $VSOScanButton is available, in case media was just inserted
		$iTotalWait = 0
		While 1
			; check that button is available
			$isScanEnabled = (ControlCommand($hVSOWnd, "", $VSOScanButton, "IsEnabled") == 1)
			If ($isScanEnabled) Then
				ExitLoop
			EndIf
			If ($iTotalWait >= 10000) Then ; wait for 10 seconds maximum (avoid infinite loop)
				ExitLoop
			EndIf
			Sleep(500)
			$iTotalWait += 500
		WEnd
	Else
		; drive not found!!!
		AddTrace("ERREUR: Impossible de trouver la chaîne " & $ComboSearchedStr & " dans la Combobox des lecteurs", _
				"ERROR: could not find drive string " & $ComboSearchedStr & " in drives Combobox")
		SetAppStatusReady()
		Return
	EndIf

	If ($bDebugInfo) Then
		AddTrace("       Debug: Drive " & $DriveLetter & " sélectionné dans VSO Inspector", "       Debug: Drive " & $DriveLetter & " selected in VSO Inspector")
	EndIf

	; disable File Test check box
	Local $hFileTestControl = ControlGetHandle($hVSOWnd, "", $VSOFilesTestCheckBox)
	Local $nFileTestControlId = _WinAPI_GetDlgCtrlID($hFileTestControl)

	ControlCommand($hVSOWnd, "", $VSOFilesTestCheckBox, "UnCheck")

	Sleep(200)

	; launch scan
	ControlClick($hVSOWnd, "", $VSOScanButton)

	If ($bDebugInfo) Then
		AddTrace("       Debug: scan lancé dans VSO Inspector", "       Debug: scan launched in VSO Inspector")
	EndIf

	Sleep(1000)

	; minimize window
	; 210214 : minimization removed, because percentages statuses updates don't work if window is minimized
;	WinSetState($hVSOWnd, "", @SW_MINIMIZE)

	If ($bDebugInfo) Then
		AddTrace($sScanLaunchedFR, $sScanLaunchedEN)
	EndIf

	; inform user that program is ready for new disc
	If ($bDebugInfo) Then
		AddTrace("", "")
		AddTrace($sReadyForDisc_DetailsFR, $sReadyForDisc_DetailsEN)
	EndIf

	SetAppStatusReady()

EndFunc   ;==>PerformScan


; this function moves the VSO Inspector window, so in case of multiple drives, the VSO Inspector windows are spread on screen insted of stacked
Func SpreadVSOWindow($hVSOWnd, $DriveLetter)
	Local $aPos = WinGetPos($hVSOWnd)
	;AddDebugTrace("    VSO window: " & $aPos[2] & "|" & $aPos[3])

	Local $iDrivePositionInList = StringInStr($sDriveList, $DriveLetter)
	Local $iDrivesNumber = StringLen($sDriveList)

	;AddDebugTrace( $DriveLetter)
	;AddDebugTrace("    $sDriveList= " & $sDriveList)
	;AddDebugTrace("    $iDrivePositionInList= " & $iDrivePositionInList)


	Local $iVSOWindowWidth = $aPos[2]      ; width of VSO Inspector window
	Local $iVSOWindowHeight = $aPos[3] ; height of VSO Inspector window

	; increase window width if required
	If ($iVSOInspectorWindowWidthIncrease > 0) Then
		$iVSOWindowWidth += $iVSOInspectorWindowWidthIncrease
	EndIf

	; we start VSO Inspector window spread from this coordinate:
	Local $iStartX = 50
	Local $iStartY = 20


	; screen resolution is @DeskTopWidth * @DeskTopHeight
	; number of VSO Inspector windows per row which can be displayed
	Local $iVSOWindowsPerRow = Floor((@DesktopWidth - $iStartX) / $iVSOWindowWidth)

	; number of VSO Inspector windows which can be displayed vertically (number of rows)
	Local $iVSOWindowsMaxRows = Floor((@DesktopHeight - $iStartY) / $iVSOWindowHeight)

	;AddDebugTrace("    $iVSOWindowsPerRow= " & $iVSOWindowsPerRow)
	;AddDebugTrace("    $iVSOWindowsMaxRows= " & $iVSOWindowsMaxRows)

	; on the screen we can display at same time maximum $iVSOWindowsPerScreen VSO Inspector windows
	; if more windows must be displayed (more drives that what the desktop supports) then we start again from ($iStartX, $iStartY)
	Local $iVSOWindowsPerScreen = $iVSOWindowsPerRow * $iVSOWindowsMaxRows
	;AddDebugTrace("    $iVSOWindowsPerScreen= " & $iVSOWindowsPerScreen)

	; so we compute X and Y indexes (starting at 1) of position on screen, in a matrix of ($iVSOWindowsPerRow * $iVSOWindowsMaxRows) windows
	; total number of VSO Screens it $iDrivesNumber

	Local $iWindowIndex = Mod($iDrivePositionInList - 1, $iVSOWindowsPerScreen) + 1

	Local $iXIndex = Mod($iWindowIndex - 1, $iVSOWindowsPerRow) + 1
	Local $iYIndex = 0
	If ($iVSOWindowsPerRow == 1) Then ; special case
		$iYIndex = $iWindowIndex
	Else
		$iYIndex = Floor(($iWindowIndex - 1) / $iVSOWindowsPerRow) + 1
		If ($iYIndex == 0) Then ;special case for first row
			$iYIndex = 1
		EndIf
	EndIf

	;AddDebugTrace("    $iWindowIndex= " & $iWindowIndex)
	;AddDebugTrace("    $iXIndex= " & $iXIndex)
	;AddDebugTrace("    $iYIndex= " & $iYIndex)

	Local $iVSOWinPositionX = $iStartX + ($iXIndex - 1) * $iVSOWindowWidth
	Local $iVSOWinPositionY = $iStartY + ($iYIndex - 1) * $iVSOWindowHeight

	;AddDebugTrace("    $iVSOWinPositionX= " & $iVSOWinPositionX)
	;AddDebugTrace("    $iVSOWinPositionY= " & $iVSOWinPositionY)

	; move the window
	WinMove($hVSOWnd, "", $iVSOWinPositionX, $iVSOWinPositionY, $iVSOWindowWidth, $iVSOWindowHeight)
EndFunc   ;==>SpreadVSOWindow



Func AddDebugTrace($DebugMsg)
	AddTrace($DebugMsg, $DebugMsg)
EndFunc   ;==>AddDebugTrace


; Func CreateButtonsSets
; dynamic creation of buttons which will allow direct access to VSO Inspector windows for a corresponding drive
; for each drive, we have these buttons:
; restore/minimize the VSO window,
; close VSO window
; cancel scan if running, then save report, then close VSO window
; drives trays (open/close switch)

; $sDriveList: (input) : string containing drives list, specified by windows letter
; $aVSOWinHandles : ByRef (output) parameter : array of the windows handles, same order as $sDriveList
; $aVSOWinMinimizeButtonsIds : ByRef (output) parameter : array containing the buttons ids, same order as $sDriveList. These buttons allow to restore/minimize the VSO window
; $aVSOWinCloseButtonsIds : ByRef (output) parameter : array containing the buttons ids, same order as $sDriveList. These buttons allow to close VSO window
; $aTrayButtonsIds : ByRef (output) parameter : array containing the buttons ids, same order as $sDriveList. These buttons allow to open/close drives trays
; $aVSOWinEndSaveButtonsIds : ByRef (output) parameter : array containing the buttons ids, same order as $sDriveList. These buttons allow to cancel scan in VSO window, save report and close window
; $aDiscSizeLabelIds : ByRef (output) parameter : array containing the labels ids, same order as $sDriveList. These labels show the inserted disc size in Gb
; $aDiscLabelLabelIds : ByRef (output) parameter : array containing the labels ids, same order as $sDriveList. These labels show the inserted disc label
; $aProgressBarIds : ByRef (output) parameter : array containing the labels ids, same order as $sDriveList. progress bar of scan
; $aPercentsStatusLabelIds : ByRef (output) parameter : array containing the label ids, same order as $sDriveList. percentages status from VSO win
; $aCurrentSpeedIds
Func CreateButtonsSets($sDriveList, ByRef $aVSOWinHandles, ByRef $aVSOWinMinimizeButtonsIds, ByRef $aVSOWinCloseButtonsIds, _
		ByRef $aVSOWinEndSaveButtonsIds, ByRef $aTrayButtonsIds, ByRef $aProgressBarIds, ByRef $aPercentsStatusLabelIds, ByRef $aCurrentSpeedIds)

	Local $iDrivesNumber = StringLen($sDriveList)

	; initialization: top left of window
	Local $iControlX = 10
	Local $iControlY = 3

	; increment between buttons set
	Local $iButtonsSetsYIncrement = 40
	Local $iGroupWidth = 563 + $PercentsStatusWidth + 10

	Local $aDriveArray = StringSplit($sDriveList, "")

	For $i = 0 To ($iDrivesNumber - 1)
		Local $ButtonX = $iControlX + 10 ; reset for each drive letter
		Local $ButtonY = $iControlY + 8
		Local $DriveLetter = $aDriveArray[$i + 1]

		; group control
		GUICtrlCreateGroup("", $iControlX, $iControlY, $iGroupWidth, 36)

		; create restore/minimize VSO Inspector window button
		; each button is 25*25 pixels, with 10 pixels of space between them
		$idButton = GUICtrlCreateButton($DriveLetter, $ButtonX, $ButtonY, 25, 25)
		$aVSOWinMinimizeButtonsIds[$i] = $idButton

		$ButtonX += 35

		; create VSO Inspector window close button
		$idButton = GUICtrlCreateButton($DriveLetter, $ButtonX, $ButtonY, 25, 25, $BS_ICON)
		GUICtrlSetImage($idButton, ".\icons\25\close-25.ico")
		$aVSOWinCloseButtonsIds[$i] = $idButton

		$ButtonX += 35

		; create VSO Inspector scan cancel button
		$idButton = GUICtrlCreateButton($DriveLetter, $ButtonX, $ButtonY, 25, 25, $BS_ICON)
		GUICtrlSetImage($idButton, ".\icons\25\save-as-25.ico")
		$aVSOWinEndSaveButtonsIds[$i] = $idButton

		$ButtonX += 35
		$ButtonX += 5 ; a bit more space for drive tray buttons

		; create drive tray open/close button
		Local $idButton = GUICtrlCreateButton("", $ButtonX, $ButtonY, 25, 25, $BS_ICON)
		GUICtrlSetImage($idButton, ".\icons\25\eject-25.ico")
		$aTrayButtonsIds[$i] = $idButton

		$ButtonX += 50

		; create drive size and label labels
		$aDiscSizeLabelIds[$i] = GUICtrlCreateLabel("", $ButtonX, $ButtonY + 4, 45, 21)

		$ButtonX += 50

		$aDiscLabelLabelIds[$i] = GUICtrlCreateLabel("", $ButtonX, $ButtonY + 4, 180, 21)

		$ButtonX += 180

		; initialization of disc informations, in case there's already disc in drive when app is launched
		If (DriveStatus($DriveLetter & ":") == $DS_READY) Then
			$fDiscSizeInGb = DriveSpaceTotal($DriveLetter & ":\") / 1024
			$sDiscSizeInGb = StringFormat("%.2f", $fDiscSizeInGb)
			$sDiscSizeInGb &= ($sLang == "FR") ? " Go" : " Gb"
			$sDiscLabel = DriveGetLabel($DriveLetter & ":\")

			GUICtrlSetData($aDiscSizeLabelIds[$i], $sDiscSizeInGb)
			GUICtrlSetData($aDiscLabelLabelIds[$i], $sDiscLabel)
		EndIf


		; create progress bar corresponding to VSO win
		$aProgressBarIds[$i] = GUICtrlCreateProgress($ButtonX, $ButtonY + 2, 123, 21)

		$ButtonX += 126

		; control for current speed
		$aCurrentSpeedIds[$i] = GUICtrlCreateLabel("", $ButtonX, $ButtonY + 4, 41, 17)

		$ButtonX += 44

		;  create label for scanned percentages status (ok, warnings, errors)
		$aPercentsStatusLabelIds[$i] = GUICtrlCreateLabel("",$ButtonX, $ButtonY + 2, $PercentsStatusWidth ,21)


		$aVSOWinHandles[$i] = 0 ; init to avoid invalid window handle access

		$iControlY += $iButtonsSetsYIncrement ; for next loop (next drive)
	Next

EndFunc   ;==>CreateButtonsSets



; function (called with timer) which update status for VSO Inspector last launched windows
Func UpdateVSOWinStatus($hWnd, $iMsg, $iIDTimer, $iTime)

	Local $aDriveArray = StringSplit($sDriveList, "")
	Local $iDrivesNumber = StringLen($sDriveList)

	; Optimization: for VSO progress bar and percents status, we don't update every second, but every 4 seconds
	Local $iDiffTime = $iTime - $iPreviousTime
	Local $bUpdateProgressAndPercents = ($iDiffTime >= $DetailedVSOStatusUpdateInterval)

	; check all drives
	For $iDriveIndex = 0 To ($iDrivesNumber - 1)
		$hVSOWinHandle = GetValidVSOWinHandle($iDriveIndex)
		If ($hVSOWinHandle > 0) Then
			$sScanButtonText = ControlGetText($hVSOWinHandle, "", $VSOScanButton)
			$bIsScanRunning = ($sScanButtonText = $sCancelScanText)

			If ($bIsScanRunning) Then
				If ($aVSOWinStates[$iDriveIndex] <> $VSO_WIN_SCANNING) Then
					GUICtrlSetBkColor($aVSOWinMinimizeButtonsIds[$iDriveIndex], $COLOR_SCAN_RUNNING) ; set button color according to scan state
				EndIf

				$aVSOWinStates[$iDriveIndex] = $VSO_WIN_SCANNING

				If ($bUpdateProgressAndPercents) Then
					; we update now

					; progress bar
					$hProgBarHdl = ControlGetHandle($hVSOWinHandle, "", "TProgressBar1")
					If ($hProgBarHdl > 0) Then
						Local $iPBPos = ProgressBarCtrlGetPos($hProgBarHdl)
						Local $iTotalPB = ProgressBarCtrlGetRange($hProgBarHdl)
						Local $fPercentProgress = ($iTotalPB <> 0) ? ($iPBPos / $iTotalPB * 100) : 0 ; progress bar in %
						Local $iCurrentPercent = Int($fPercentProgress)
						If ($iPBPos == $iTotalPB) Then
							$iCurrentPercent = 100 ; to be sure  there's no rounding issue...
						EndIf
						Local $State = _SendMessage($hVSOWinHandle, $PBM_GETSTATE, 0, 0)
						GUICtrlSetData($aProgressBarIds[$iDriveIndex], $iCurrentPercent)
					EndIf

					; update percents statuses
					UpdateVSOStatuses($hVSOWinHandle, $iDriveIndex)
				EndIf ; If ($bUpdateProgressAndPercents)

			Else ; If ($bIsScanRunning) Then
				; if scan is NOT running but was running just before, we play sound to notify the user
				If ($aVSOWinStates[$iDriveIndex] == $VSO_WIN_SCANNING) Then ; a scan was running
					If ($sEndScanNotification <> "") Then
						SoundPlay($sEndScanNotification, 0)

						Local $sDiscSizeInGb = ControlGetText($hDiscAutoScanGUI, "", $aDiscSizeLabelIds[$iDriveIndex])
						Local $sDiscLabel = ControlGetText($hDiscAutoScanGUI, "", $aDiscLabelLabelIds[$iDriveIndex])

						Local $DriveLetter = $aDriveArray[$iDriveIndex+1]
						Local $ScanEndedTrace = "["  & $DriveLetter & "][" & $sDiscLabel & "][" & $sDiscSizeInGb & "]"
						AddScanEndedTrace($ScanEndedTrace)
						If ($bDebugInfo) Then
							Local $sTraceFR = "* Scan terminé ou annulé pour " & $ScanEndedTrace
							Local $sTraceEN = "* Scan ended or canceled for " & $ScanEndedTrace
							AddTrace($sTraceFR, $sTraceEN)
						EndIf

					EndIf
					GUICtrlSetBkColor($aVSOWinMinimizeButtonsIds[$iDriveIndex], $COLOR_NO_SCAN_RUNNING) ; set button color according to scan state

					; reset progress bar to zero: no scanning running
					GUICtrlSetData($aProgressBarIds[$iDriveIndex], 0)


					; update percentages to last infos
					UpdateVSOStatuses($hVSOWinHandle, $iDriveIndex)

					; reset drive speed info
					GUICtrlSetStyle($aCurrentSpeedIds[$iDriveIndex], 0) ; resetting style forces redraw with empty label

				EndIf
				$aVSOWinStates[$iDriveIndex] = $VSO_WIN_IDLE

			EndIf ; If ($bIsScanRunning) Then

		Else ; If ($hVSOWinHandle > 0) Then
			; here $hVSOWinHandle == 0 (no VSO Inspector window)
			If ($aVSOWinStates[$iDriveIndex] <> $VSO_WIN_NO_WINDOW) Then        ; window was closed: must update all states

				GUICtrlSetStyle($aVSOWinMinimizeButtonsIds[$iDriveIndex], 0)   ; reset to default look, which says there's not VSO Inspector window opened

				EmptyStatusesFromVSOWin($iDriveIndex) ; reset to empty

				$aVSOWinStates[$iDriveIndex] = $VSO_WIN_NO_WINDOW
			EndIf
		EndIf ; If ($hVSOWinHandle > 0) Then

	Next ;For $iDriveIndex = 0 To ($iDrivesNumber - 1)

	If ($bUpdateProgressAndPercents) Then
		$iPreviousTime = $iTime ; update for next calls
	EndIf


EndFunc   ;==>UpdateVSOWinStatus

Func ShowHideLog()
	$pos = WinGetPos($hDiscAutoScanGUI)
	WinMove($hDiscAutoScanGUI, "", $pos[0], $pos[1], $pos[2] + 10, $pos[3])

	;ControlHide ( $hDiscAutoScanGUI, "", $idLogEdit )
	;ControlShow ( $hDiscAutoScanGUI, "", $idLogEdit )
EndFunc   ;==>ShowHideLog


; get key from locale ini file of VSO Inspector, according to selected language
Func GetLocalForVSOInsp($sVSOLocaleFile, $Section, $Default)

	$Value = IniRead(@ProgramFilesDir & "\vso\tools\Lang\" & $sVSOLocaleFile, $Section, "locale", "")

	; if $Value is empty -> use "original" key instead of "locale"

	If ($Value == "") Then
		$Value = IniRead(@ProgramFilesDir & "\vso\tools\Lang\" & $sVSOLocaleFile, $Section, "original", $Default)
	EndIf

	;AddDebugTrace($Section & "=" & $Value)

	Return $Value
EndFunc   ;==>GetLocalForVSOInsp

; --

Func SetAppStatusBusy($sDetailMsgFR="", $sDetailMsgEN="")
	If ($sLang == "FR") Then
		GUICtrlSetData($IdAppStatus, $sBusyTreatingDiscFR)
		IF ($sDetailMsgFR <> "") Then
			GUICtrlSetData ($IdAppStatusDetail, $sDetailMsgFR)
		EndIf
	Else
		GUICtrlSetData($IdAppStatus, $sBusyTreatingDiscEN)
		IF ($sDetailMsgEN <> "") Then
			GUICtrlSetData ($IdAppStatusDetail, $sDetailMsgEN)
		EndIf
	EndIf
	GUICtrlSetColor($IdAppStatus, $COLOR_APP_BUSY)
EndFunc   ;==>SetAppStatusBusy


Func SetAppStatusReady()
	If ($sLang == "FR") Then
		GUICtrlSetData($IdAppStatus, $sReadyForDiscFR)
		GUICtrlSetData($IdAppStatusDetail, $sReadyForDisc_DetailsFR)
	Else
		GUICtrlSetData($IdAppStatus, $sReadyForDiscEN)
		GUICtrlSetData($IdAppStatusDetail, $sReadyForDisc_DetailsEN)
	EndIf
	GUICtrlSetColor($IdAppStatus, $COLOR_APP_READY)

EndFunc   ;==>SetAppStatusReady


; check drives are cd/dvd
; $sDriveList drives characters, no space char
; return the same string, but if drive is not cd/dvd then it is removed of the string
Func GetValidDrives($sDriveList)
	Local $sRetDriveList = ""

	Local $aDriveArray = StringSplit($sDriveList, "")
	Local $iDrivesNumber = UBound($aDriveArray)
	; check all drives
	For $i = 1 To ($aDriveArray[0])
		Local $DriveLetter = $aDriveArray[$i]
		$DriveType = DriveGetType($DriveLetter & ":")
		If ($DriveType == $DT_CDROM) Then
			$sRetDriveList &= $DriveLetter
		Else
			If ($sLang == "FR") Then
				MsgBox($IDOK, "ERREUR", "le drive [" & $DriveLetter & "] n'est pas de type cd/dvd/bluray. Ce drive est ignoré")
			Else
				MsgBox($IDOK, "ERROR", "the drive [" & $DriveLetter & "] is not of type cd/dvd/bluray. This drive is ignored")
			EndIf
		EndIf ; If ($DriveType == $DT_CDROM) Then
	Next ; For $i = 0 To ($iDrivesNumber - 1)

	If (StringLen($sRetDriveList) == 0) Then
		If ($sLang == "FR") Then
			MsgBox($IDOK, "ERREUR", "Il n'y a aucun drive valide configuré. Veuillez en paramétrer au moins un")
		Else
			MsgBox($IDOK, "ERROR", "There is no valid configured drive. Please configure at least one")
		EndIf
	EndIf

	Return ($sRetDriveList)
EndFunc   ;==>GetValidDrives




Func ProgressBarCtrlGetPos($h_wnd)
	;GUICtrlSendMsg
	Return _SendMessage($h_wnd, $PBM_GETPOS, 0, 0)
EndFunc   ;==>ProgressBarCtrlGetPos

Func ProgressBarCtrlGetRange($h_wnd)
	Return _SendMessage($h_wnd, $PBM_GETRANGE, False, 0)
EndFunc   ;==>ProgressBarCtrlGetRange



; CODING IDEA from https://www.autoitscript.com/forum/topic/71629-get-control-screenshot-as-bitmap/
Func UpdateVSOStatuses($hVSOWnd, $aDriveIndex)

	$hVSOWnd = GetValidVSOWinHandle($aDriveIndex) ; to be sure window still exists

	If ($hVSOWnd <=0 ) Then
		Return
	EndIf

	; _WinAPI_BitBlt doesn't work if window is minimize
	; solution in this case: from https://www.codeproject.com/articles/20651/capturing-minimized-window-a-kid-s-trick
	; but it creates flickering. when minimizing again the window, Windows automatically activates the next window in Z order.
	; I tried to temporary set originaly activated window at "top most", but then flickering is even worse (this concerned window is flickering, instead of disc-auto-scan window)
	; For this reason, percentages statuses are NOT updated if window is minimized
	Local $iWinState = WinGetState($hVSOWnd)
	Local $bIsWinMinimized = (BitAND($iWinState, $WIN_STATE_MINIMIZED) == $WIN_STATE_MINIMIZED)

	If ($bIsWinMinimized) Then
		EmptyCopiedStatusesFromVSOWin($aDriveIndex) ; reset to empty images
		Return
	EndIf

	Local $PictId = $aPercentsStatusLabelIds[$aDriveIndex]

	If ($PictId <= 0) Then
		Return
	EndIf

	Local $hWnd = ControlGetHandle($hVSOWnd,"","TPanel1")

	If ($hWnd == 0) Then
		AddDebugTrace("UpdateVSOPercentsStatus error: control for TPanel1 not found.")
		Return
	EndIf

	Local $hDC = _WinAPI_GetDC($hWnd)
	Local $memDC = _WinAPI_CreateCompatibleDC($hDC)
	Local $memBmp = _WinAPI_CreateCompatibleBitmap($hDC, $VSOWInPercentsStatusWidth, 21)
	_WinAPI_SelectObject ($memDC, $memBmp)

#comments-start
	Local $iVSOWndExStyle = 0
	Local $bMinAnimation = GetMinAnimate()

	Local $WinWaitDelay = 0

	; was needed for minimized windows
	Local $ActiveWindowHandle = WinGetHandle("[ACTIVE]") ;

	; OLD CODE, doesn't work -> disabled if minimized window
	If ($bIsWinMinimized) Then
		SetMinAnimate(False) ; disable minimize/maximize animation temporary

		$iVSOWndExStyle = _WinAPI_GetWindowLong($hVSOWnd, $GWL_EXSTYLE) ; memorize extended style

		_WinAPI_SetWindowLong($hVSOWnd, $GWL_EXSTYLE, BitOR($iVSOWndExStyle, $WS_EX_LAYERED)) ; temporary add WS_EX_LAYERED

		Local $iTranscolor, $iAlpha
		Local $iInfo = _WinAPI_GetLayeredWindowAttributes($hVSOWnd, $iTranscolor, $iAlpha)
		If ($iInfo <> $LWA_ALPHA OR $iTranscolor <> 0 OR $iAlpha <> 1) Then
			_WinAPI_SetLayeredWindowAttributes($hVSOWnd, 0, 1, $LWA_ALPHA) ; make window transparent
		EndIf

		$WinWaitDelay = AutoItSetOption ("WinWaitDelay", 0) ; disable delay so there's no flickering

		;solution 1: flickering because window is activated
		_WinAPI_ShowWindow($hVSOWnd,@SW_SHOWNOACTIVATE)


;		If ($ActiveWindowHandle <> $hDiscAutoScanGUI) Then
;			WinActivate($ActiveWindowHandle) ; reactivate window which was active
;		EndIf

		Sleep(50) ; wait that window was painted so _WinAPI_BitBlt can work
	EndIf
#comments-end

	; copy pixels zone of media errors
	_WinAPI_BitBlt($memDC, 0, 0,  $VSOWInPercentsStatusWidth, 21, $hDC, 7, 15, $SRCCOPY)

#comments-start
	; OLD CODE, doesn't work -> disabled if minimized window
	If ($bIsWinMinimized) Then
		_WinAPI_ShowWindow ($hVSOWnd, @SW_MINIMIZE )

		;WinSetState($hVSOWnd, "", @SW_MINIMIZE) ; minimize window

		If ($ActiveWindowHandle <> $hDiscAutoScanGUI) Then
			WinActivate($ActiveWindowHandle) ; reactivate window which was active because when mininizing, Windows activates according to z-order
		EndIf

		AutoItSetOption("WinWaitDelay", $WinWaitDelay) ; restore value

		_WinAPI_SetWindowLong($hVSOWnd, $GWL_EXSTYLE, $iVSOWndExStyle) ; restore extended style
		SetMinAnimate($bMinAnimation) ; restore animation flag
	EndIf
#comments-end

	; now update directly percentages control
	Local $PictHandle = GUICtrlGetHandle($PictId)
	Local $hPictDC = _WinAPI_GetDC($PictHandle)

	Local $iPercentBoxWidth = 56


	; DOCUMENTATON: _WinAPI_BitBlt($hDCDest, $iXDest, $iYDest, $iWidth, $iHeight, $hDCSrc, $iXSrc, $iYSrc, $SRCCOPY)

	_WinAPI_BitBlt($hPictDC, 0, 0,  $iPercentBoxWidth, 21, $memDC, 0, 0, $SRCCOPY) ; copy 1st box
	_WinAPI_BitBlt($hPictDC, $iPercentBoxWidth + 1, 0,  $iPercentBoxWidth, 21, $memDC, 71, 0, $SRCCOPY) ; copy 2nd box
	_WinAPI_BitBlt($hPictDC, 2 * ($iPercentBoxWidth + 1), 0,  $iPercentBoxWidth, 21, $memDC, 142, 0, $SRCCOPY) ; copy 3rd box

	; cleanup memory
	_WinAPI_ReleaseDC($PictHandle, $hPictDC)
	_WinAPI_ReleaseDC($hWnd, $hDC)
	_WinAPI_DeleteDC($memDC)
	_WinAPI_DeleteObject ($memBmp)


	; NOW update also drive current speed

	; rectangle coordinates for "snapshot"
	Local $iDriveSpeedXStart = 89
	Local $iDriveSpeedYStart = 85
	Local $iDriveSpeedXEnd = 129
	Local $iDriveSpeedYEnd = 101

	Local $hWnd = ControlGetHandle($hVSOWnd,"","TPageControl1")

	$hDC = _WinAPI_GetDC($hWnd)

	; now update directly percentages control
	Local $PictHandle = GUICtrlGetHandle($aCurrentSpeedIds[$aDriveIndex])
	Local $hPictDC = _WinAPI_GetDC($PictHandle)

	; copy pixels zone of drive speed
	; DOCUMENTATON: _WinAPI_BitBlt($hDCDest, $iXDest, $iYDest, $iWidth, $iHeight, $hDCSrc, $iXSrc, $iYSrc, $SRCCOPY)
	_WinAPI_BitBlt($hPictDC, 0, 0, $iDriveSpeedXEnd - $iDriveSpeedXStart + 1,$iDriveSpeedYEnd - $iDriveSpeedYStart + 1, $hDC, $iDriveSpeedXStart, $iDriveSpeedYStart, $SRCCOPY)

	; cleanup memory
	_WinAPI_ReleaseDC($PictHandle, $hPictDC)
	_WinAPI_ReleaseDC($hWnd, $hDC)


	_WinAPI_GetLastError()

EndFunc


; reset the controls related to infos from VSO Inspector win
Func EmptyStatusesFromVSOWin($iDriveIndex)
	If ($iDriveIndex >= 0 And $iDriveIndex < UBound($aProgressBarIds)) Then

		; reset progress bar to zero
		GUICtrlSetData($aProgressBarIds[$iDriveIndex], 0)

		GUICtrlSetStyle($aCurrentSpeedIds[$iDriveIndex], 0) ; resetting style forces redraw with empty label

		GUICtrlSetStyle($aPercentsStatusLabelIds[$iDriveIndex], 0) ; resetting style forces redraw with empty label

	EndIf
EndFunc

; empty statuses which are directly copied with _WinAPI_BitBlt. These statuses cannot be updated if window is minimized
Func EmptyCopiedStatusesFromVSOWin($iDriveIndex)
	If ($iDriveIndex >= 0 And $iDriveIndex < UBound($aPercentsStatusLabelIds)) Then

		GUICtrlSetStyle($aCurrentSpeedIds[$iDriveIndex], 0) ; resetting style forces redraw with empty label

		GUICtrlSetStyle($aPercentsStatusLabelIds[$iDriveIndex], 0) ; resetting style forces redraw with empty label

	EndIf
EndFunc

Func SetMinAnimate($bBoolean = True)
    Local Const $tagANIMATIONINFO = "uint cbSize;int iMinAnimate"
    Local $sStruct = DllStructCreate($tagANIMATIONINFO)
    DllStructSetData($sStruct, "iMinAnimate", $bBoolean)
    DllStructSetData($sStruct, "cbSize", DllStructGetSize($sStruct))
    $aReturn = DllCall('user32.dll', 'int', 'SystemParametersInfo', 'uint', $SPI_SETANIMATION, 'int', DllStructGetSize($sStruct), 'ptr', DllStructGetPtr($sStruct), 'uint', 0)
    If IsArray($aReturn) Then Return 1
    Return 0
EndFunc   ;==>_SetMinAnimate



Func GetMinAnimate()
    Local Const $tagANIMATIONINFO = "uint cbSize;int iMinAnimate"
    Local $sStruct = DllStructCreate($tagANIMATIONINFO)
    DllStructSetData($sStruct, "cbSize", DllStructGetSize($sStruct))
    $aReturn = DllCall('user32.dll', 'int', 'SystemParametersInfo', 'uint', $SPI_GETANIMATION, 'int', DllStructGetSize($sStruct), 'ptr', DllStructGetPtr($sStruct), 'uint', 0)
    If IsArray($aReturn) Then Return $sStruct.iMinAnimate
    Return 0
EndFunc   ;==>_GetMinAnimate


Func _GetBrushColor($hBrush)
    ;https://devblogs.microsoft.com/oldnewthing/20190802-00/?p=102747
    Local $iErr
    Local Const $BS_SOLID = 0x00000000
    Local Const $tagLogBrush = "uint lbStyle; dword lbColor; ulong_ptr lbHatch"
    Local $tLogBrush = DllStructCreate($tagLogBrush)
    Local $iSzLogBrush = DllStructGetSize($tLogBrush)
    If _WinAPI_GetObject($hBrush, $iSzLogBrush, DllStructGetPtr($tLogBrush)) <> $iSzLogBrush Then
        $iErr = 0x1
    ElseIf DllStructGetData($tLogBrush, "lbStyle") <> $BS_SOLID Then
        $iErr = 0x2
    Else
        Return DllStructGetData($tLogBrush, "lbColor")
    EndIf
    Return SetError($iErr, 0, 0xFFFFFFFF) ;CLR_NONE
EndFunc

Func GUICtrlGetBkColor($hWnd)
    If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
    Local $hWnd_Main = _WinAPI_GetParent ($hWnd)
    Local $hDC = _WinAPI_GetDC($hWnd)
    Local $hBrush = _SendMessage($hWnd_Main, $WM_CTLCOLORSTATIC, $hdc, $hWnd)
    Local $iColor = _GetBrushColor($hBrush)
    _WinAPI_ReleaseDC($hWnd, $hDC)
    Return $iColor
EndFunc

; creation of the "options" dialog window
Func CreateOptionsDialog(ByRef $idOptionsOk, ByRef $idOptionsCancel, ByRef $aOptionsDialogCtrlIds, ByRef $iOptionsDialogCtrlIdsNumber)

	$hOptions = GUICreate("Options", 600, 300, -1, -1,	 BitOR($WS_CAPTION, $DS_MODALFRAME,  $DS_SETFOREGROUND), -1, $hDiscAutoScanGUI )

	Local $iControlY = 10
	Local $LabelsWidth = 115
	Local $OptionsControlX = 130
	Local $iCtrlIdIndex = 0

	Opt("GUIDataSeparatorChar","|")

	; "Lang" parameter
	;;;;;;;;;;;;;;;;;;
	Local $sLangParam = IniRead($sIniFile, $sIniSection, "Lang", "EN")

	GUICtrlCreateLabel($sLang == "FR" ? "Langue" : "Language", 15, $iControlY + 4, $LabelsWidth, 17)
	GUICtrlSetTip(-1, $sLang == "FR" ? "Langue de l'application, FR pour français, EN pour anglais" : "Application language, FR for french, EN for english")

	$iLangParamOptionCtrlId = GUICtrlCreateCombo("", $OptionsControlX, $iControlY, 40, 30, BitOR( $CBS_DROPDOWNLIST , $WS_VSCROLL))
	GUICtrlSetData($iLangParamOptionCtrlId, "FR|EN", $sLangParam)

	$aOptionsDialogCtrlIds[$iCtrlIdIndex][0] = "iLangParamOptionCtrlId"
	$aOptionsDialogCtrlIds[$iCtrlIdIndex][1] = $iLangParamOptionCtrlId
	$iCtrlIdIndex += 1
	$iControlY += 35 ; for next control

	; "DriveList" parameter
	;;;;;;;;;;;;;;;;;;;;;;;
	Local $sDriveListWithSpacesParam = IniRead($sIniFile, $sIniSection, "DriveList", "D")

	GUICtrlCreateLabel($sLang == "FR" ? "Drives configurés" : "Configured drives", 15, $iControlY+2, $LabelsWidth, 17)
	GUICtrlSetTip(-1, $sLang == "FR" ? "Lettres des lecteurs configurés pour scan. Séparer avec caractère espace" : "Letters of drives configured for scanning. Separate with space character")

	$iDriveListOptionCtrlId = GUICtrlCreateEdit($sDriveListWithSpacesParam, $OptionsControlX, $iControlY, 80, 17, 0)

	$aOptionsDialogCtrlIds[$iCtrlIdIndex][0] = "iDriveListOptionCtrlId"
	$aOptionsDialogCtrlIds[$iCtrlIdIndex][1] = $iDriveListOptionCtrlId

	$iCtrlIdIndex += 1
	$iControlY += 35 ; for next control

	; "EndScanNotification" parameter
	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
	Local $sEndScanNotificationParam = IniRead($sIniFile, $sIniSection, "EndScanNotification", ".\tada.wav")

	GUICtrlCreateLabel($sLang == "FR" ? "Son fin de scan" : "End of scan sound", 15, $iControlY+2, $LabelsWidth, 17)
	GUICtrlSetTip(-1, $sLang == "FR" ? "Son joué lorsque fin de scan détectée. Doit être fichier de type .wav. Peut être vide si pas de son souhaité" _
									: "Sound played when end of scan detected. Must be .wav file type. Can be empty if no sound wanted")

	$iEndScanNotificationOptionCtrlId = GUICtrlCreateEdit($sEndScanNotificationParam, $OptionsControlX, $iControlY, 410, 17, 0)

	$iEndScanNotificationOptionButtonCtrlId = GUICtrlCreateButton("...", $OptionsControlX + 420, $iControlY-1, 20, 19)

	; we add in $aOptionsDialogCtrlIds both edit and button controls
	$aOptionsDialogCtrlIds[$iCtrlIdIndex][0] = "iEndScanNotificationOptionCtrlId"
	$aOptionsDialogCtrlIds[$iCtrlIdIndex][1] = $iEndScanNotificationOptionCtrlId
	$iCtrlIdIndex += 1
	$aOptionsDialogCtrlIds[$iCtrlIdIndex][0] = "$EndScanNotificationOptionButtonCtrlId"
	$aOptionsDialogCtrlIds[$iCtrlIdIndex][1] = $iEndScanNotificationOptionButtonCtrlId
	$iCtrlIdIndex += 1
	$iControlY += 35 ; for next control

	; "DebugInfo" parameter
	;;;;;;;;;;;;;;;;;;;;;;;

	Local $sDebugInfoParam = IniRead($sIniFile, $sIniSection, "DebugInfo", "N")

	GUICtrlCreateLabel("Debug", 15, $iControlY + 4, $LabelsWidth, 17)
	GUICtrlSetTip(-1, $sLang == "FR" ? "Mode debug. Y pour oui, N pour non" : "Debug mode. Y or N")


	$iDebugInfoOptionCtrlId = GUICtrlCreateCombo("", $OptionsControlX, $iControlY, 40, 30, BitOR( $CBS_DROPDOWNLIST , $CBS_AUTOHSCROLL, $WS_VSCROLL))
	GUICtrlSetData($iDebugInfoOptionCtrlId, "Y|N", $sDebugInfoParam)

	$aOptionsDialogCtrlIds[$iCtrlIdIndex][0] = "iDebugInfoOptionCtrlId"
	$aOptionsDialogCtrlIds[$iCtrlIdIndex][1] = $iDebugInfoOptionCtrlId
	$iCtrlIdIndex += 1
	$iControlY += 55 ; for next control

	; "VSOInspectorWindowWidthIncrease" parameter
	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
	Local $sVSOInspectorWindowWidthIncreaseParam = IniRead($sIniFile, $sIniSection, "VSOInspectorWindowWidthIncrease", "0")

	GUICtrlCreateLabel($sLang == "FR" ? "Augment. fenêtre VSO" : "VSOInspectorWindowWidthIncrease", 15, $iControlY+2, $LabelsWidth, 17)
	GUICtrlSetTip(-1, $sLang == "FR" ? "Augmentation de la largeur de la fenêtre VSO Inspector, en pixels, lors de son lancement" : "VSO Inspector window width increase, in pixels, when it is launched")

	$iVSOInspectorWindowWidthIncreaseParamCtrlId = GUICtrlCreateEdit($sVSOInspectorWindowWidthIncreaseParam, $OptionsControlX, $iControlY, 40, 17, 0)

	$aOptionsDialogCtrlIds[$iCtrlIdIndex][0] = "iVSOInspectorWindowWidthIncreaseParamCtrlId"
	$aOptionsDialogCtrlIds[$iCtrlIdIndex][1] = $iVSOInspectorWindowWidthIncreaseParamCtrlId
	$iCtrlIdIndex += 1
	$iControlY += 35 ; for next control

	; "MPCHC" parameter
	;;;;;;;;;;;;;;;;;;;;;;;
	Local $sMPCHCParam = IniRead($sIniFile, $sIniSection, "MPCHC", "")

	GUICtrlCreateLabel($sLang == "FR" ? "Executable MPC-HC" : "MPC-HC executable", 15, $iControlY+2, $LabelsWidth, 17)
	GUICtrlSetTip(-1, $sLang == "FR" ? "Executable MPC-HC avec chemin complet. Si vide, l'application détecte via entrée dans le menu démarrer de Windows" _
									: "MPC-HC executable with path. If empty, application detects with entry in Windows start menu")

	$iMPCHCParamCtrlId = GUICtrlCreateEdit($sMPCHCParam, $OptionsControlX, $iControlY, 410, 17, 0)

	$iMPCHCParamButtonCtrlId = GUICtrlCreateButton("...", $OptionsControlX + 420, $iControlY-1, 20, 19)

	; we add in $aOptionsDialogCtrlIds both edit and button controls
	$aOptionsDialogCtrlIds[$iCtrlIdIndex][0] = "iMPCHCParamCtrlId"
	$aOptionsDialogCtrlIds[$iCtrlIdIndex][1] = $iMPCHCParamCtrlId
	$iCtrlIdIndex += 1
	$aOptionsDialogCtrlIds[$iCtrlIdIndex][0] = "iMPCHCParamButtonCtrlId"
	$aOptionsDialogCtrlIds[$iCtrlIdIndex][1] = $iMPCHCParamButtonCtrlId
	$iCtrlIdIndex += 1

	$iControlY += 35 ; for next control


	$iControlY += 15 ; more space before ok/cancel button

	$idOptionsOk = GUICtrlCreateButton("OK", (600/2)-50, $iControlY, 45, 25)
	$idOptionsCancel = GUICtrlCreateButton($sLang == "FR" ? "Annuler" : "Cancel", (600/2)+10, $iControlY, 45, 25)

	$iOptionsDialogCtrlIdsNumber = $iCtrlIdIndex

	return $hOptions
EndFunc

; treat messages linked to options dialog controls
Func TreatOptionsDialogActions($GuiMsg, ByRef $aOptionsDialogCtrlIds, ByRef $iOptionsDialogCtrlIdsNumber)

	If ($iOptionsDialogCtrlIdsNumber == 0) Then
		Return
	EndIf

	; check if $GuiMsg is control id of options dialog
	; we search in dimension 1 and not 0
	$iOptionCtrlIndex = _ArraySearch($aOptionsDialogCtrlIds, $GuiMsg, 0, $iOptionsDialogCtrlIdsNumber-1, 0, 0, 1, 1)

	If ($iOptionCtrlIndex == -1) Then
		Return
	EndIf

	$sOptionsDialogParam = $aOptionsDialogCtrlIds[$iOptionCtrlIndex][0]

	; treat buttons click -> open file choice dialog
	Switch $sOptionsDialogParam
		Case "$EndScanNotificationOptionButtonCtrlId"
			$sMessage = ($sLang == "FR") ? "Choix de fichier" : "File Choice"
			$sFilter = ($sLang == "FR") ? "Son wav (*.wav)" : "Wav sound (*.wav)"
			Local $sFileOpenDialog = FileOpenDialog($sMessage, ".\", $sFilter, $FD_FILEMUSTEXIST)

			; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
			FileChangeDir(@ScriptDir)

			If (@error == 0 And $sFileOpenDialog <> "") Then
				; edit control is previous array element (just before current button)
				$sOptionsDialogParamCtrlId = $aOptionsDialogCtrlIds[$iOptionCtrlIndex-1][1]

				GUICtrlSetData($sOptionsDialogParamCtrlId, $sFileOpenDialog)
			EndIf

		Case "iMPCHCParamButtonCtrlId"
			$sMessage = ($sLang == "FR") ? "Choix de fichier" : "File Choice"
			$sFilter = ($sLang == "FR") ? "Executable MPC-HC (*.exe)" : "MPC-HC executable (*.exe)"
			Local $sFileOpenDialog = FileOpenDialog($sMessage, ".\", $sFilter, $FD_FILEMUSTEXIST)

			; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
			FileChangeDir(@ScriptDir)

			If (@error == 0 And $sFileOpenDialog <> "") Then
				; edit control is previous array element (just before current button)
				$sOptionsDialogParamCtrlId = $aOptionsDialogCtrlIds[$iOptionCtrlIndex-1][1]

				GUICtrlSetData($sOptionsDialogParamCtrlId, $sFileOpenDialog)
			EndIf


;		Case Else
	EndSwitch

EndFunc

; validation of options dialog
; step 1 check that all values are valid
; step 2 if ok, then save/update values in .ini file
Func ValidateOptionsDialog(ByRef $aOptionsDialogCtrlIds, ByRef $iOptionsDialogCtrlIdsNumber)

	Local $bAreOptionsValidated = False

	; step 1: read all values from dialog controls
	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

	Local $iOptionCtrlIndex = -1
	Local $sParamValue = ""

	$iOptionCtrlIndex = _ArraySearch($aOptionsDialogCtrlIds, "iLangParamOptionCtrlId", 0, $iOptionsDialogCtrlIdsNumber-1, 0, 0, 1, 0)
	Local $sLangParam = GUICtrlRead($aOptionsDialogCtrlIds[$iOptionCtrlIndex][1])

	$iOptionCtrlIndex = _ArraySearch($aOptionsDialogCtrlIds, "iDriveListOptionCtrlId", 0, $iOptionsDialogCtrlIdsNumber-1, 0, 0, 1, 0)
	Local $sDriveListWithSpacesParam = GUICtrlRead($aOptionsDialogCtrlIds[$iOptionCtrlIndex][1])

	$iOptionCtrlIndex = _ArraySearch($aOptionsDialogCtrlIds, "iEndScanNotificationOptionCtrlId", 0, $iOptionsDialogCtrlIdsNumber-1, 0, 0, 1, 0)
	Local $sEndScanNotificationParam = GUICtrlRead($aOptionsDialogCtrlIds[$iOptionCtrlIndex][1])

	$iOptionCtrlIndex = _ArraySearch($aOptionsDialogCtrlIds, "iDebugInfoOptionCtrlId", 0, $iOptionsDialogCtrlIdsNumber-1, 0, 0, 1, 0)
	Local $sDebugInfoParam = GUICtrlRead($aOptionsDialogCtrlIds[$iOptionCtrlIndex][1])

	$iOptionCtrlIndex = _ArraySearch($aOptionsDialogCtrlIds, "iVSOInspectorWindowWidthIncreaseParamCtrlId", 0, $iOptionsDialogCtrlIdsNumber-1, 0, 0, 1, 0)
	Local $sVSOInspectorWindowWidthIncreaseParam = GUICtrlRead($aOptionsDialogCtrlIds[$iOptionCtrlIndex][1])

	$iOptionCtrlIndex = _ArraySearch($aOptionsDialogCtrlIds, "iMPCHCParamCtrlId", 0, $iOptionsDialogCtrlIdsNumber-1, 0, 0, 1, 0)
	Local $sMPCHCParam = GUICtrlRead($aOptionsDialogCtrlIds[$iOptionCtrlIndex][1])


	; step 2: check values are valid
	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
	Local $ValidationErrorMsg = ""

	; check $sDriveListWithSpacesParam
	Local $sDriveListParam = StringStripWS($sDriveListWithSpacesParam, $STR_STRIPALL) ; remove white spaces

	Local $aDriveArray = StringSplit($sDriveListParam, "")
	Local $iDrivesNumber = UBound($aDriveArray)
	Local $sRetDriveList = ""
	; check all drives
	For $i = 1 To ($aDriveArray[0])
		Local $DriveLetter = $aDriveArray[$i]
		$DriveType = DriveGetType($DriveLetter & ":")
		If ($DriveType == $DT_CDROM) Then
			$sRetDriveList &= $DriveLetter
		Else
			If ($sLang == "FR") Then
				$ValidationErrorMsg &= "ERREUR: le drive [" & $DriveLetter & "] n'est pas de type cd/dvd/bluray." & @LF
			Else
				$ValidationErrorMsg &= "ERROR: the drive [" & $DriveLetter & "] is not of type cd/dvd/bluray." & @LF
			EndIf
		EndIf ; If ($DriveType == $DT_CDROM) Then
	Next ; For $i = 0 To ($iDrivesNumber - 1)

	If (StringLen($sRetDriveList) == 0) Then
		If ($sLang == "FR") Then
			$ValidationErrorMsg &= "ERREUR: Il n'y a aucun drive valide configuré." & @LF
		Else
			$ValidationErrorMsg &= "ERROR: there is no valid configured drive." & @LF
		EndIf
	EndIf

	; check $sEndScanNotificationParam
	If ($sEndScanNotificationParam <> "") Then
		If (Not FileExists($sEndScanNotificationParam)) Then
			If ($sLang == "FR") Then
				$ValidationErrorMsg &= "ERREUR: Son fin de scan: le fichier n'existe pas: " & $sEndScanNotificationParam & @LF
			Else
				$ValidationErrorMsg &= "ERREUR: End of scan sound: the file doesn't exist: " & $sEndScanNotificationParam & @LF
			EndIf
		EndIf
	EndIf

	; check $sVSOInspectorWindowWidthIncreaseParam
	Local $bParamIsValid = True
	If (Not StringIsDigit($sVSOInspectorWindowWidthIncreaseParam)) Then
		$bParamIsValid = False
	Else
		If ($sVSOInspectorWindowWidthIncreaseParam < 0 Or $sVSOInspectorWindowWidthIncreaseParam>200) Then
			$bParamIsValid = False
		EndIf
	EndIf

	If $bParamIsValid == False Then
		If ($sLang == "FR") Then
			$ValidationErrorMsg &= "ERREUR: Augment. fenêtre VSO: la veleur doit être comprise entre 0 et 200 pixels"  & @LF
		Else
			$ValidationErrorMsg &= "ERREUR: VSO window increase: value must be between 0 and 200 pixels" & @LF
		EndIf
	EndIf


	; check $sMPCHCParam
	If ($sMPCHCParam <> "") Then
		If (Not FileExists($sMPCHCParam)) Then
			If ($sLang == "FR") Then
				$ValidationErrorMsg &= "ERREUR: Executable MPC-HC: le fichier n'existe pas: " & $sMPCHCParam & @LF
			Else
				$ValidationErrorMsg &= "ERREUR:MPC-HC executable: the file doesn't exist: " & $sMPCHCParam & @LF
			EndIf
		EndIf
	EndIf

	; if error in parameter, display message box and exit function
	If ($ValidationErrorMsg <> "") Then
		MsgBox(BitOR($MB_OK, $MB_ICONERROR, $MB_APPLMODAL), ($sLang == "FR") ? "Erreur d'options" : "Options error", $ValidationErrorMsg)
		Return
	EndIf


	; step 3: update variables and .ini file

	Local $bParamsNeedRestart = False ; does the paramaters change needs app restart?

	If ($sLangParam <> $sLang) Then
		$sLang = $sLangParam
		$bParamsNeedRestart = True
		IniWrite($sIniFile, $sIniSection, "Lang", $sLangParam)
	EndIf

	If ($sDriveListWithSpacesParam <> $sDriveListWithSpaces) Then
		; we don't update variables in memory now, will be done at restart of app
		$bParamsNeedRestart = True
		IniWrite($sIniFile, $sIniSection, "DriveList", $sDriveListWithSpacesParam)
	EndIf

	If ($sEndScanNotificationParam <> $sEndScanNotification) Then
		$sEndScanNotification = $sEndScanNotificationParam
		IniWrite($sIniFile, $sIniSection, "EndScanNotification", $sEndScanNotificationParam)
	EndIf

	If ($sDebugInfoParam <> $sDebugInfo) Then
		$bDebugInfo = ($sDebugInfoParam == "Y")
		IniWrite($sIniFile, $sIniSection, "DebugInfo", $sDebugInfoParam)
	EndIf

	Local $sVSOInspectorWindowWidthIncreaseIni = IniRead($sIniFile, $sIniSection, "VSOInspectorWindowWidthIncrease", "0") ; must read from ini file, because global variable may be computed differently from ini value
	If ($sVSOInspectorWindowWidthIncreaseParam <> $sVSOInspectorWindowWidthIncreaseIni) Then
		$bParamsNeedRestart = True
		IniWrite($sIniFile, $sIniSection, "VSOInspectorWindowWidthIncrease", $sVSOInspectorWindowWidthIncreaseParam)
	EndIf

	If ($sMPCHCParam <> "") Then
		If ($sMPCHCParam <> $sMPCHC) Then
			$sMPCHC = $sMPCHCParam
			IniWrite($sIniFile, $sIniSection, "MPCHC", $sMPCHCParam)
		EndIf
	EndIf


	If ($bParamsNeedRestart) Then
		If ($sLang == "FR") Then
			MsgBox(BitOR($MB_OK, $MB_ICONWARNING, $MB_APPLMODAL), "Information", "L'application doit être redémarrée pour que les nouveaux paramètres prennent effet")
		Else
			MsgBox(BitOR($MB_OK, $MB_ICONWARNING, $MB_APPLMODAL), "Information", "Application needs to be restart to apply new parameters")
		EndIf
	EndIf

	$bAreOptionsValidated = True

	Return $bAreOptionsValidated

EndFunc


; creation of the "options" dialog window
Func CreateHelpDialog()

	$hHelp = GUICreate(($sLang == "FR") ? "Aide" : "Help" , 950, 575, -1, -1,BitOR($WS_CAPTION,  $WS_POPUP, $DS_SETFOREGROUND) , -1, $hDiscAutoScanGUI )	;BitOR($WS_CAPTION,  $WS_POPUP, $DS_SETFOREGROUND), -1, $hDiscAutoScanGUI )

	Local Const $sFontName = "Segoe UI"
	Local Const $iFontSize = 8.5

	Local $iControlY = 10
	Local $HelpText

	If ($sLang == "FR") Then
		$HelpText = "Disc-auto-scan est un logiciel permettant de simplifier la vérification de l'intégrité physique de disques: cds, dvds et blurays."
	Else
		$HelpText = "Disc-auto-scan is a software which allows simplification of discs physical integrity check: cds, dvds and blurays."
	EndIf
	GUICtrlCreateLabel($HelpText, 15, $iControlY+2, 925, 17)
	GUICtrlSetFont(-1, $iFontSize, 0, 0,$sFontName) ; change font
	$iControlY += 22 ; for next control

	If ($sLang == "FR") Then
		$HelpText = 'A cette fin, le logiciel gratuit "VSO Inspector" (développé par VSO Software) est utilisé et piloté par Disc-auto-scan: '
	Else
		$HelpText = 'For this, the "VSO Inspector" freeware (developpped by VSO Software) is used and driven by Disc-auto-scan: '
	EndIf
	GUICtrlCreateLabel($HelpText, 15, $iControlY+2, 925, 17)
	GUICtrlSetFont(-1, $iFontSize, 0, 0,$sFontName) ; change font
	$iControlY += 18 ; for next control

	If ($sLang == "FR") Then
		$HelpText = '   à chaque disque inséré, "VSO Inspector" est automatiquement lancé ainsi que le scan du disque par ce dernier logiciel.'
	Else
		$HelpText = '   for each inserted disc, "VSO Inspector" is automatically launched, as well as the disc scan by this latest software.'
	EndIf
	GUICtrlCreateLabel($HelpText, 15, $iControlY+2, 850, 17)
	GUICtrlSetFont(-1, $iFontSize, 0, 0,$sFontName) ; change font



	$iControlY += 40 ; for next control

	If ($sLang == "FR") Then
		$HelpText = "Pour configurer Disc-auto-scan, veuillez ouvrir la boite des options en cliquant, dans la fenêtre principale, sur le bouton:"
		GUICtrlCreateLabel($HelpText, 15, $iControlY+2, 630, 17)
		GUICtrlSetFont(-1, $iFontSize, 0, 0,$sFontName) ; change font
		GUICtrlCreateIcon("icons\25\parameters-25.ico", -1, 635 , $iControlY-3, 25, 25)
	Else
		$HelpText = "To configure Disc-auto-scan, please open the options dialog box by clicking, in the main window, on the button:"
		GUICtrlCreateLabel($HelpText, 15, $iControlY+2, 590, 17)
		GUICtrlSetFont(-1, $iFontSize, 0, 0,$sFontName) ; change font
		GUICtrlCreateIcon("icons\25\parameters-25.ico", -1, 598 , $iControlY-3, 25, 25)
	EndIf


	$iControlY += 22 ; for next control

	If ($sLang == "FR") Then
		$HelpText = "  Dans la boite des options, vous pourrez avoir une description détaillée de chaque option, en passant la souris sur le texte de l'option."
	Else
		$HelpText = "  In the options dialog, you can get a detailed description of each option, by hovering mouse on the option text."
	EndIf
	GUICtrlCreateLabel($HelpText, 15, $iControlY+2, 775, 17)
	GUICtrlSetFont(-1, $iFontSize, 0, 0,$sFontName) ; change font

	$iControlY += 22 ; for next control

	If ($sLang == "FR") Then
		$HelpText = "  L'option la plus importante est d'indiquer à Disc-auto-scan la liste du/des drive(s) (lecteur ou graveur) à utiliser. Ceci est spécifié avec la lettre correspondante de Windows."
	Else
		$HelpText = "  The most important option is to indicate Disc-auto-scan the list of drive(s) (reader or burner) to be used. This is specified with the corresponding Windows letter."
	EndIf
	GUICtrlCreateLabel($HelpText, 15, $iControlY+2, 925, 17)
	GUICtrlSetFont(-1, $iFontSize, 0, 0,$sFontName) ; change font

	$iControlY += 40 ; for next control

	If ($sLang == "FR") Then
		$HelpText = "Une fois paramétré, voici comment utiliser Disc-auto-scan: "
	Else
		$HelpText = "When configured, here is how to use Disc-auto-scan: "
	EndIf
	GUICtrlCreateLabel($HelpText, 15, $iControlY+2, 925, 17)
	GUICtrlSetFont(-1, $iFontSize, 0, 0,$sFontName) ; change font

	$iControlY += 22 ; for next control

	If ($sLang == "FR") Then
		$HelpText = "Les drives (si plusieurs) sont organisés par lignes (une ligne par drive). En dessous de ces lignes, Disc-auto-scan indique les informations d'état et de disques traités."
	Else
		$HelpText = "The drives (if several) are organized by rows (one row per drive). Below these rows, Disc-auto-scan indicates status informations and treated discs."
	EndIf
	GUICtrlCreateLabel($HelpText, 15, $iControlY+2, 925, 17)
	GUICtrlSetFont(-1, $iFontSize, 0, 0,$sFontName) ; change font

	$iControlY += 32 ; for next control

	If ($sLang == "FR") Then
		$HelpText = "Pour chaque drive, des boutons correspondent aux actions relatives à ce drive:"
	Else
		$HelpText = "For each drive, buttons are corresponding to actions related to this drive:"
	EndIf
	GUICtrlCreateLabel($HelpText, 15, $iControlY+2, 925, 17)
	GUICtrlSetFont(-1, $iFontSize, 0, 0,$sFontName) ; change font

	$iControlY += 22 ; for next control

	If ($sLang == "FR") Then
		$HelpText = "  Le premier bouton, avec la lettre du drive, permet d'afficher ou minimiser la fenêtre correspondante de VSO Inspector (dernier scan lancé)."
	Else
		$HelpText = "  The first button, with the drive letter, allows to display or minimize the corresponding VSO Inspector window (last launched scan)."
	EndIf
	GUICtrlCreateLabel($HelpText, 15, $iControlY+2, 925, 17)
	GUICtrlSetFont(-1, $iFontSize, 0, 0,$sFontName) ; change font

	$iControlY += 18 ; for next control

	If ($sLang == "FR") Then
		$HelpText = "     Le fond de ce premier bouton est: gris si pas de fenêtre VSO Inspector ouverte, bleu ciel si le scan est en cours, et blanc si le scan est terminé."
	Else
		$HelpText = "     The background of this first button is: gray if no opened VSO Inspector window, cyan if scan is running, and white if scan is ended."
	EndIf
	GUICtrlCreateLabel($HelpText, 15, $iControlY+2, 925, 17)
	GUICtrlSetFont(-1, $iFontSize, 0, 0,$sFontName) ; change font

	$iControlY += 26 ; for next control


	GUICtrlCreateIcon("icons\25\close-25.ico", -1, 20 , $iControlY-3, 25, 25)
	If ($sLang == "FR") Then
		$HelpText = "permet de fermer la fenêtre VSO Inspector."
	Else
		$HelpText = "allows to close the VSO Inspector window."
	EndIf
	GUICtrlCreateLabel($HelpText, 47, $iControlY+2, 925, 17)
	GUICtrlSetFont(-1, $iFontSize, 0, 0,$sFontName) ; change font

	$iControlY += 25 ; for next control

	GUICtrlCreateIcon("icons\25\save-as-25.ico", -1, 20 , $iControlY-3, 25, 25)
	If ($sLang == "FR") Then
		$HelpText = "permet d'enregistrer le log du scan de VSO Inspector. Si un scan est encore en cours, il est d'abord terminé (annulation). Le nom du fichier est pré-rempli au nom du disque."
	Else
		$HelpText = "allows to save the VSO Inspector scan log. If a scan is running, it is first terminated (cancelled). The filename is pre-filled with the disc name."
	EndIf
	GUICtrlCreateLabel($HelpText, 47, $iControlY+2, 925, 17)
	GUICtrlSetFont(-1, $iFontSize, 0, 0,$sFontName) ; change font

	$iControlY += 25 ; for next control

	GUICtrlCreateIcon("icons\25\eject-25.ico", -1, 20 , $iControlY-3, 25, 25)
	If ($sLang == "FR") Then
		$HelpText = "pilote le tiroir du drive (éjection et fermeture)."
	Else
		$HelpText = "pilots the drive tray (ejection and closure)."
	EndIf
	GUICtrlCreateLabel($HelpText, 47, $iControlY+2, 925, 17)
	GUICtrlSetFont(-1, $iFontSize, 0, 0,$sFontName) ; change font

	$iControlY += 25 ; for next control

	If ($sLang == "FR") Then
		$HelpText = "  La suite de la ligne pour le drive indique les informations suivantes: taille et nom du disque (nom donné par l'éditeur), "
	Else
		$HelpText = "  The rest of the line for the drive indicates the following informations: size and name of disc (name given by the editor), "
	EndIf
	GUICtrlCreateLabel($HelpText, 15, $iControlY+2, 925, 17)
	GUICtrlSetFont(-1, $iFontSize, 0, 0,$sFontName) ; change font

	$iControlY += 18 ; for next control

	If ($sLang == "FR") Then
		$HelpText = "       puis l'état du scan de VSO Inspector: barre de progrès, vitesse actuelle de lecture, et pourcentages pour les secteurs lus: Bon, Problème, et Erreur."
	Else
		$HelpText = "       and then the VSO Inspector scan state: progress bar, current reading speed, and percentages for read sectors: Good, Problem, and Error."
	EndIf
	GUICtrlCreateLabel($HelpText, 15, $iControlY+2, 925, 17)
	GUICtrlSetFont(-1, $iFontSize, 0, 0,$sFontName) ; change font

	$iControlY += 18 ; for next control

	If ($sLang == "FR") Then
		$HelpText = "  NB: pour que ces états puissent être techniquement repris de la fenêtre VSO Inspector, cette dernière ne doit pas être minimisée dans Windows."
	Else
		$HelpText = "  NOTE: in order that these states can be technically retrieved from the VSO Inspector window, this one must not be minimized in Windows."
	EndIf
	GUICtrlCreateLabel($HelpText, 15, $iControlY+2, 925, 17)
	GUICtrlSetFont(-1, $iFontSize, 0, 0,$sFontName) ; change font

	$iControlY += 40 ; for next control

	If ($sLang == "FR") Then
		$HelpText = "IMPORTANT: A chaque fois que vous insérez un disque, vous devez attendre (i.e. ne plus agir dans Windows) que le disque soit détecté, puis que le scan soit lancé. "
	Else
		$HelpText = "IMPORTANT: At each time you insert a disc, you must wait (mean. do not interact with Windows) that the disc is detected, and then the scan is launched. "
	EndIf
	GUICtrlCreateLabel($HelpText, 15, $iControlY+2, 925, 17)
	GUICtrlSetFont(-1, $iFontSize, 0, 0,$sFontName) ; change font

	$iControlY += 18 ; for next control

	If ($sLang == "FR") Then
		$HelpText = "    Seulement ensuite vous pouvez reprendre le contrôle avec Windows (vous pouvez afficher vos fenêtres de travail par dessus celles de VSO Inspector voire Disc-auto-scan)"
	Else
		$HelpText = "    Only after that you can take back control of Windows (you can display your working windows/applications on top of VSO Inspector or even Disc-auto-scan windows)"
	EndIf
	GUICtrlCreateLabel($HelpText, 15, $iControlY+2, 925, 17)
	GUICtrlSetFont(-1, $iFontSize, 0, 0,$sFontName) ; change font

	$iControlY += 22 ; for next control

	If ($sLang == "FR") Then
		$HelpText = "En fin de scan, pour les blurays veuillez contrôler d'office manuellement le résultat dans la fenêtre VSO Inspector correspondante."
	Else
		$HelpText = "At end of scan, for blurays please always control manually the result in the corresponding VSO Inspector window."
	EndIf
	GUICtrlCreateLabel($HelpText, 15, $iControlY+2, 925, 17)
	GUICtrlSetFont(-1, $iFontSize, 0, 0,$sFontName) ; change font

	$iControlY += 18 ; for next control

	If ($sLang == "FR") Then
		$HelpText = "    Du fait du grand nombre de secteurs des disques blurays, le pourcentage 'Problème' peut rester à zéro malgré un ou quelques secteur(s) non fiable(s)."
	Else
		$HelpText = "    Because bluray discs have a large number of sectors, the 'Problem' percentage may stay at zero in spite of one or a few unreliable sector(s)."
	EndIf
	GUICtrlCreateLabel($HelpText, 15, $iControlY+2, 925, 17)
	GUICtrlSetFont(-1, $iFontSize, 0, 0,$sFontName) ; change font

	$iControlY += 22 ; for next control


	$iControlY += 10 ; more space before close button

	$idHelpClose = GUICtrlCreateButton(($sLang == "FR") ? "Fermer" : "Close", (950/2)-22, $iControlY, 45, 25)

	return $hHelp
EndFunc