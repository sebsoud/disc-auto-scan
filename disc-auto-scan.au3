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
#include <GUIConstantsEx.au3>
#include <StringConstants.au3>

#include <GuiComboBox.au3>
#include <GuiButton.au3>
#include <GuiTab.au3>

#include <MsgBoxConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <ButtonConstants.au3>
#include <StaticConstants.au3>

#include "drivestray.au3" ; special functions for drives tray


;Global Const $COLOR_SCAN_RUNNING = 0xC7E9FF
Global Const $COLOR_SCAN_RUNNING = 0xB1FFFC
;Global Const $COLOR_NO_SCAN_RUNNING = 0xFFFCF2
Global Const $COLOR_NO_SCAN_RUNNING = 0xFDFCFC


Global Const $COLOR_APP_READY = 0x018501
Global Const $COLOR_APP_BUSY = 0xEE0202

Global const $sDiscAutoScanRelease = "Disc Auto Scan 0.2"

const $DBT_DEVICEARRIVAL = "0x00008000"


; VSO Inspector window state : values for $aVSOWinStates
Global const $VSO_WIN_NO_WINDOW 	= 0
Global const $VSO_WIN_IDLE 			= 1
Global const $VSO_WIN_SCANNING		= 2

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; resources values for VSO Inspector window
Global const $VSOScanButton = "TButton4"
Global const $VSOPropPageControl = "TPageControl1"
Global const $VSODrivesComboBox = "TComboBox1"
Global const $VSOFilesTestCheckBox = "TJvCheckBox1"
Global const $VSOSaveReportButton = "TButton2"

; text messages
Global const $sReadyForDiscFR_Long = "-     PRET, vous pouvez insérer UN disque puis ATTENDRE que le scan ait été lancé"
Global const $sReadyForDiscEN_Long = "-     READY, you can insert ONE disc then WAIT that scan is launched"

Global const $sReadyForDiscFR_Short = "PRET à traiter UN disque"
Global const $sReadyForDiscEN_Short = "READY for processing ONE disc"

Global const $sBusyTreatingDiscFR = "OCCUPE, disque en traitement..."
Global const $sBusyTreatingDiscEN = "BUSY, processing disc..."

Global const $sScanLaunchedFR = "-     Scan en cours... Vous pouvez reprendre la main sur Windows"
Global const $sScanLaunchedEN = "-     Scan on going... You can again work in Windows"


Global $bShowAllVSoWindows = True ; if True, next call to DoVSOWinMinimizeAll() will show all windows. If False, it will minimize them
Global const $sShowAllVSO_FR = "Montrer tous VSO Inspector"
Global const $sMinimizeAllVSO_FR = "Minimiser tous VSO Inspector"
Global const $sShowAllVSO_EN = "Show all VSO Inspector"
Global const $sMinimizeAllVSO_EN = "Minimize all VSO Inspector"



;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; read values from ini file (configuration)

Global const $sIniFile = ".\disc-auto-scan.ini"
Global $DriveLetter = ""

; language (EN or FR)
Global const $sLang = IniRead($sIniFile, "Configuration", "Lang", "EN")

; log debug info
$sDebugInfo = IniRead($sIniFile, "Configuration", "DebugInfo", "N")
Global $bDebugInfo = ($sDebugInfo == "Y")

Global $sVSOLocaleFile = RegRead("HKEY_CURRENT_USER\SOFTWARE\Vso\Inspector", "locale_file")
If ($sVSOLocaleFile == "") Then ; registry can be empty if VSO Inspector installed but user didn't manually choose a language in menu
	$sVSOLocaleFile = "INS_original.ini" ; default texts
EndIf

Global const $sEndScanNotification = IniRead($sIniFile, "Configuration", "EndScanNotification", ".\tada.wav")


Global $sDriveListWithSpaces = IniRead($sIniFile, "Configuration", "DriveList", "D")
Global $sDriveList = ""

; increase of VSOInspector window width at opening (pixels, max 200. 0 is possible)
Global $iVSOInspectorWindowWidthIncrease = 0;
Global const $sVSOInspectorWindowWidthIncrease = IniRead($sIniFile, "Configuration", "VSOInspectorWindowWidthIncrease", "0")
If (StringIsInt($sVSOInspectorWindowWidthIncrease)) Then
	$iVSOInspectorWindowWidthIncrease = Number($sVSOInspectorWindowWidthIncrease)
	If ($iVSOInspectorWindowWidthIncrease > 200) Then
		$iVSOInspectorWindowWidthIncrease = 200
	EndIf
EndIf

Global const $SaveScanResultsTextFR = IniRead($sIniFile, "Configuration", "SaveScanResultsTextFR", "Annuler")
Global const $SaveScanResultsTextEN = IniRead($sIniFile, "Configuration", "SaveScanResultsTextEN", "Save scan results...")

;;;;;;;;;;;;;;;;;;;;;;
; beginning of program
;;;;;;;;;;;;;;;;;;;;;;

; initialization for tray functions, see drivestray.au3
InitForTrayFunctions()

; initial window position. If possible, put at bottom of screen, so for Fullhd or greater resolution, the VSO Inspector windows will be rather on top of screen
Global $iDiscAutoScanWinPosX = 50
Global $iDiscAutoScanWinPosY = 100
Global $iDiscAutoScanWinWidth = 700
Global $iDiscAutoScanWinHeight = 480
If (@DeskTopHeight > 100 + $iDiscAutoScanWinHeight) Then
 	$iDiscAutoScanWinPosY = @DeskTopHeight - $iDiscAutoScanWinHeight - 100
EndIf

$sDriveList = StringStripWS ($sDriveListWithSpaces, $STR_STRIPALL); remove white spaces
$sDriveList = GetValidDrives($sDriveList) ; check drives are cd/dvd
Local $iDrivesNumber = StringLen ($sDriveList)

; increase width if many drives
If ($iDrivesNumber > 3 AND @DesktopWidth > 850) Then
	$iDiscAutoScanWinWidth += 120
EndIf

; window creation
Global $hDiscAutoScanGUI = GUICreate("Disc auto scan", $iDiscAutoScanWinWidth, $iDiscAutoScanWinHeight, $iDiscAutoScanWinPosX, $iDiscAutoScanWinPosY)

; controls creation
; -----------------

; compute area for buttons, on right of log edit control

; 100=space for legend   5=space on right of group box, 10=space on right of groupbox before end of window
Global $iButtonsAreaX = 100 + $iDrivesNumber * 35 + 5 + 10
If ($iButtonsAreaX < 180) Then
	$iButtonsAreaX = 180 ; minimum space at right of log edit control
EndIf

; display/minimize all button
$idDisplayAllButton = GUICtrlCreateButton("", $iDiscAutoScanWinWidth - $iButtonsAreaX + 10, 10, $iButtonsAreaX - 20 , 25)
UpdateShowAllVSoWindowsButtonLabel($idDisplayAllButton)

; edit control for log
Local $Height = $iDiscAutoScanWinHeight - 20
;Local $Height = $iDiscAutoScanWinHeight - 120
Global $idLogEdit = GUICtrlCreateEdit("", 10, 10, $iDiscAutoScanWinWidth - 10 - $iButtonsAreaX, $Height , BITOR($ES_AUTOVSCROLL , $WS_VSCROLL , $WS_HSCROLL, $ES_READONLY , $ES_MULTILINE))

;AddTrace("  Lecteurs configuré pour autoscan: " & $sDriveListWithSpaces, "  Drives configured for autoscan: " & $sDriveListWithSpaces)

; creation of buttons which will allow direct access to VSO Inspector windows for a corresponding drive
; 4 sets of buttons:
; 1st set allows to restore/minimize the VSO window,
; 2nd set allows to close VSO window
; 3rd set cancelling scan if running, then save report, then close VSO window
; 4th set pilot the drives trays

Global $aVSOWinHandles[$iDrivesNumber]  ; array of the VSO Inspector windows handles, same order as $sDriveList
Global $aVSOWinStates[$iDrivesNumber]  ; array of states of VSO Inspector windows. Updated by UpdateVSOWinStatus()
; initialization
For $i = 0 To $iDrivesNumber - 1
	$aVSOWinStates[$i] = $VSO_WIN_NO_WINDOW
Next


Global $aVSOWinMinimizeButtonsIds[$iDrivesNumber] ; array containing the buttons ids, same order as $sDriveList
Global $aVSOWinCloseButtonsIds[$iDrivesNumber] ; array containing the buttons ids, same order as $sDriveList
Global $aVSOWinEndSaveButtonsIds[$iDrivesNumber] ; array containing the buttons ids, same order as $sDriveList
Global $aTrayButtonsIds[$iDrivesNumber] ; array containing the buttons ids, same order as $sDriveList
CreateButtonsSets($sDriveList, $aVSOWinMinimizeButtonsIds, $aVSOWinCloseButtonsIds , $aVSOWinEndSaveButtonsIds, $aTrayButtonsIds, $aVSOWinHandles)

; app label (incl release)
$IdAppLabel = GUICtrlCreateLabel ($sDiscAutoScanRelease, $iDiscAutoScanWinWidth - $iButtonsAreaX + 10, $iDiscAutoScanWinHeight - 30)

; exit button creation
$idExit=0
If ($sLang == "FR") Then
	$idExit = GUICtrlCreateButton("Quitter", $iDiscAutoScanWinWidth - 60 , $iDiscAutoScanWinHeight - 35, 50, 25)
Else
	$idExit = GUICtrlCreateButton("Exit", $iDiscAutoScanWinWidth - 60 , $iDiscAutoScanWinHeight - 35, 50, 25)
EndIf

$iStatusY = 44 + (4*40) + 20
If ($sLang == "FR") Then
	GUICtrlCreateGroup ("Etat Disc auto scan", $iDiscAutoScanWinWidth - $iButtonsAreaX + 10, 44 + (4*40) + 20 , $iButtonsAreaX - 20, 40)
Else
	GUICtrlCreateGroup ("Disc auto scan Status", $iDiscAutoScanWinWidth - $iButtonsAreaX + 10, 44 + (4*40) + 20 , $iButtonsAreaX - 20, 40)
EndIf
Global const $IdAppStatus = GUICtrlCreateLabel ("", $iDiscAutoScanWinWidth - $iButtonsAreaX + 20, $iStatusY + 16, $iButtonsAreaX - 35 ,22 )

;;;;;;;;;;;;;;;;;
GUISetState() ; display the GUI

Global const $sCancelScanText = GetLocalForVSOInsp($sVSOLocaleFile, "0015_INCODE", "Cancel")
Global const $sSaveScanResultsWinText = GetLocalForVSOInsp($sVSOLocaleFile, "0049_INCODE", "Save scan results...")


; register messages handlers
GUIRegisterMsg($WM_DEVICECHANGE , "DeviceChange")

; timer for status update, every 1 second
_Timer_SetTimer($hDiscAutoScanGUI, 1000, "UpdateVSOWinStatus")

;;;;;;;;;;;;;
; inform user that program is ready for new disc
AddTrace($sReadyForDiscFR_Long, $sReadyForDiscEN_Long)
SetAppStatusReady()

; global variables
;;;;;;;;;;;;;;;;;;;
Global $bScanToLaunch = False


; ---------------------------------

; FOR DEBUG ONLY
; FOR DEBUG ONLY
; FORCE direct test (avoiding to have to insert a disc for detection)
;PerformScan("E")
;PerformScan("G")
;PerformScan("F")
;Exit

;---------------------------------


;;;;;;;;;;;
; main loop

Do
    $GuiMsg = GUIGetMsg()

	If ($bScanToLaunch) Then
		PerformScan($DriveLetter)
		$bScanToLaunch = False
	Endif

	; treat click on "show/minimize all" button
	If ($GuiMsg == $idDisplayAllButton) Then
		DoVSOWinMinimizeAll()
		ContinueLoop
	EndIf

	; treat click on VSO Inspector windows "restore/minimize" buttons

	$iButtonIndex = _ArraySearch($aVSOWinMinimizeButtonsIds, $GuiMsg)
	If ($iButtonIndex > -1) Then
		DoVSOWinMinimize($iButtonIndex)
	EndIf ; If ($iButtonIndex > -1)

	; treat click on VSO Inspector windows  "close" buttons
	If ($iButtonIndex == -1) Then	; button was not identified yet
		$iButtonIndex = _ArraySearch($aVSOWinCloseButtonsIds, $GuiMsg)
		If ($iButtonIndex > -1) Then
			DoVSOWinClose($iButtonIndex)
		EndIf ; If ($iButtonIndex > -1)
	EndIf ; If ($iButtonIndex == -1

	; treat click on VSO Inspector windows  "End scan and save" buttons
	If ($iButtonIndex == -1) Then	; button was not identified yet
		$iButtonIndex = _ArraySearch($aVSOWinEndSaveButtonsIds, $GuiMsg)
		If ($iButtonIndex > -1) Then
			DoVSOWinEndSave($iButtonIndex)
		EndIf ; If ($iButtonIndex > -1)
	EndIf ; If ($iButtonIndex == -1

	; treat click on drives tray buttons (open/close tray)
	If ($iButtonIndex == -1) Then	; button was not identified yet
		$iButtonIndex = _ArraySearch($aTrayButtonsIds, $GuiMsg)
		If ($iButtonIndex > -1) Then
			DoTrayOpenClose($iButtonIndex)
		Endif ; If ($iButtonIndex > -1)
	EndIf ; If ($iButtonIndex == -1)

Until $GuiMsg = $GUI_EVENT_CLOSE Or $GuiMsg = $idExit

Exit


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;; FUNCTIONS ;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;


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
EndFunc



; Function: adding trace in window
Func AddTrace($FrMsg, $EnMsg)
	If $sLang == "FR" Then
		GUICtrlSetData($idLogEdit, $FrMsg & @CRLF, 1)
	Else
		GUICtrlSetData($idLogEdit, $EnMsg & @CRLF, 1)
	EndIf
EndFunc

; helper function, retrieve VSO Inspector window handle, after checking that it is still valid
; if handle is not valid, it returns 0 instead of handle, AND, if needed, $aVSOWinHandles[$iButtonIndex] is set to 0 for subsequent uses
Func GetValidVSOWinHandle($iButtonIndex)
	$hVSOWinHandle = 0

	If ($iButtonIndex <= UBound($aVSOWinHandles) - 1) Then ; check $iButtonIndex for safety...
		$hVSOWinHandle = $aVSOWinHandles[$iButtonIndex]
	Endif

	If ($hVSOWinHandle > 0) Then
		; check that the window still exist
		Local $iWinState = WinGetState($hVSOWinHandle)
		If ($iWinState == 0) Then
			; window doesn't exist anymore -> reset handle to 0
			$aVSOWinHandles[$iButtonIndex] = 0
		EndIf
	Endif

	Return $hVSOWinHandle
EndFunc

; show/minimize all VSO Inspector windows
Func DoVSOWinMinimizeAll()
	$iIndexMax = UBound($aVSOWinHandles) - 1
	Local $iButtonIndex = 0
	For $iButtonIndex = 0 To $iIndexMax ; loop on all VSO Insp windows
		Local $hVSOWinHandle = GetValidVSOWinHandle($iButtonIndex)
		If ($hVSOWinHandle > 0) Then
			; NOTE: GetValidVSOWinHandle() guarantees that WinGetState will succeed
			Local $iWinState = WinGetState($hVSOWinHandle)
			Local $bIsWinMinimized = (BitAND($iWinState, $WIN_STATE_MINIMIZED) <> 0)

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
		Endif ; If ($hVSOWinHandle > 0)
	Next ; For $iButtonIndex = 0 To $iIndexMax -1

	; change behaviour for next call (global variable)
	$bShowAllVSoWindows = Not($bShowAllVSoWindows)
	; update button label
	UpdateShowAllVSoWindowsButtonLabel($idDisplayAllButton)
EndFunc

Func DoVSOWinMinimize($iButtonIndex)
	Local $hVSOWinHandle = GetValidVSOWinHandle($iButtonIndex)
	If ($hVSOWinHandle > 0) Then
		; NOTE: GetValidVSOWinHandle() guarantees that WinGetState will succeed
		Local $iWinState = WinGetState($hVSOWinHandle)
		If (BitAND($iWinState, $WIN_STATE_MINIMIZED) <> 0) Then
			WinSetState($hVSOWinHandle, "", @SW_RESTORE)
		Else
			WinSetState($hVSOWinHandle, "", @SW_MINIMIZE)
		EndIf
	Endif ; If ($hVSOWinHandle > 0)
EndFunc

Func DoVSOWinClose($iButtonIndex)
	; close window
	$hVSOWinHandle = GetValidVSOWinHandle($iButtonIndex)
	If ($hVSOWinHandle > 0) Then
		If (WinClose($hVSOWinHandle) == 0) Then
;; log error to add
		EndIf
	EndIf ; If ($hVSOWinHandle > 0)
EndFunc

Func DoVSOWinEndSave($iButtonIndex)
	$hVSOWinHandle = GetValidVSOWinHandle($iButtonIndex)

	If ($hVSOWinHandle > 0) Then
		; first, show window in case it's minimized
		; NOTE: GetValidVSOWinHandle() guarantees that WinGetState will succeed
		Local $iWinState = WinGetState($hVSOWinHandle)
		If BitAND($iWinState, $WIN_STATE_MINIMIZED) Then
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
			$iTotalWait = 0
			While 1
				Sleep(500)
				$iTotalWait += 500
				$sScanButtonText = ControlGetText($hVSOWinHandle, "", $VSOScanButton)
				$bIsScanRunning = ($sScanButtonText == $sCancelScanText)
				If ($bIsScanRunning == false) Then
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
				Local $DriveTrayLetter = $aDriveArray[$iButtonIndex + 1]

;				Sleep(1000) ; wait that disc is available since is was recently closed
				$sDiscLabel = DriveGetLabel($DriveTrayLetter & ":\")

				$fDiscSizeInGb = DriveSpaceTotal($DriveLetter & ":\") / 1024
				$sDiscSizeInGb = StringFormat("%.2f", $fDiscSizeInGb)

				$sDefaultFileName = $sDiscLabel & "-" & $sDiscSizeInGb & ".txt"

				Send ($sDefaultFileName)
			Endif
		Else
; log error to add: cannot save report and close window
		Endif ; If ($isSaveEnabled) Then
	EndIf ; If ($hVSOWinHandle == 0)
EndFunc

Func DoTrayOpenClose($iButtonIndex)
	$aDriveArray = StringSplit($sDriveList, "")
	Local $DriveTrayLetter = $aDriveArray[$iButtonIndex + 1]

	If (IsDriveTrayOpen($DriveTrayLetter)) Then
		CDTray($DriveTrayLetter & ":", $CDTRAY_CLOSED)
	Else
		CDTray($DriveTrayLetter & ":", $CDTRAY_OPEN)
	EndIf
EndFunc

; Function: handler of $WM_DEVICECHANGE message (media inserted in drive notification)
Func DeviceChange($hWndGUI, $MsgID, $WParam, $LParam)
    If $WParam == $DBT_DEVICEARRIVAL Then

        ; Create a struct from $lParam which contains a pointer to a Windows-created struct.

        Local Const $tagDEV_BROADCAST_VOLUME = "dword dbcv_size; dword dbcv_devicetype; dword dbcv_reserved; dword dbcv_unitmask; word dbcv_flags"
        Local Const $DEV_BROADCAST_VOLUME = DllStructCreate($tagDEV_BROADCAST_VOLUME, $lParam)

        Local Const $DeviceType = DllStructGetData($DEV_BROADCAST_VOLUME, "dbcv_devicetype")
		Local Const $UnitMask = DllStructGetData($DEV_BROADCAST_VOLUME, "dbcv_unitmask")

		$DriveLetter = GetDriveLetterFromUnitMask($UnitMask)

		If StringInStr ($sDriveList, $DriveLetter) Then
			$sDiscLabel = DriveGetLabel($DriveLetter & ":\")
			$fDiscSizeInGb = DriveSpaceTotal($DriveLetter & ":\") / 1024
			$sDiscSizeInGb = StringFormat("%.2f", $fDiscSizeInGb)
			If ($bScanToLaunch) Then
				; scan of previous disc not finished! -> log warning message
				AddTrace("ATTENTION: insertion de disque [" & $sDiscLabel & "] dans le lecteur [" & $DriveLetter & "] TROP TOT, ne sera pas traité" , "WARNING: disc [" & $sDiscLabel & "] inserted TOO EARLY in drive [" & $DriveLetter & "], will not be treated")
				Return
			Else
				$sTraceFR = " [" & $sDiscLabel & "][" & $sDiscSizeInGb & " Go] inséré dans lecteur [" & $DriveLetter & "], veuillez patienter..."
				$sTraceEN = " [" & $sDiscLabel & "][" & $sDiscSizeInGb & " Gb] inserted in drive [" & $DriveLetter & "], please wait..."
				AddTrace($sTraceFR , $sTraceEN)
				; update status label
				SetAppStatusBusy()
			EndIf
			$bScanToLaunch = True
		EndIf
    EndIf
EndFunc

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
EndFunc ;==>GetDriveLetterFromUnitMask

; -----------------------------------


Func PerformScan($DriveLetter)
	SetAppStatusBusy()

	; wait enough time that disc is recognized by windows
	$iTotalWait = 0
	While 1
		Sleep(500)
		$iTotalWait += 500
		$sDriveFS = DriveGetFileSystem ($DriveLetter & ":\")
		If ($sDriveFS == $DT_UDF) Then  ; dvd or bluray    NOTE:  $DT_UNDEFINED is returned if no disc or raw file system
			ExitLoop
		EndIf
		If ($iTotalWait >= 15000) Then ; wait for 15 seconds maximum (avoid infinite loop)
			ExitLoop
		EndIf
	WEnd

	; launch MPC-HC, so dvd is accessible for VSO Inspector if protected (CSS protection)
	;;;;;;;;;;;;;;;;
	Local $sMPCHC = IniRead($sIniFile, "Configuration", "MPCHC", "C:\Program Files\MPC-HC\mpc-hc64.exe")
	Local $sMPCHCCmd = """" & $sMPCHC &  """ " & $DriveLetter & ": /open /minimized /new"

	If ($bDebugInfo) Then
		AddTrace("       Debug: lancement de " & $sMPCHCCmd , "       Debug: Launching " & $sMPCHCCmd)
	Endif

	Local $sMPCHCPid = Run($sMPCHCCmd )

	If ($sMPCHCPid == 0) Then
		AddTrace("ERREUR: Lancement de " & $sMPCHCCmd & " terminé avec @error=" & @error, "ERROR: Launching " & $sMPCHCCmd & " terminated with @error=" & @error)
		SetAppStatusReady()
		Return
	EndIf

	Local $sMPCHCWindow = IniRead($sIniFile, "Configuration", "MPCHCWindow", "Media Player Classic Home Cinema")

	If ($bDebugInfo) Then
		AddTrace("       Debug: attente fenêtre: " & $sMPCHCWindow , "       Debug: waiting for window: " & $sMPCHCWindow)
	Endif

	Local $hMPCHC = WinWait("[CLASS:"& $sMPCHCWindow &"]", "", 10)

	; now wait that disc is opened (max 40 seconds, it can be quite slow on usb drives)
	; for this, we check the MPC-HC status evolution
	; it evolves from "" to "opening" to "stopped"
	; we cannot test directly hard-coded text because MPC-HC is multi-languages
	Local $sMPCStatusControl = IniRead($sIniFile, "Configuration", "MPCStatusControl", "[CLASS:Static; INSTANCE:3]")

	$sMPCPreviousStatusText = ""
	$iTotalWait = 0
	While 1
		Sleep(1000)
		$iTotalWait += 1000

		$sMPCStatusText = ControlGetText($hMPCHC, "", $sMPCStatusControl) ; check current MPC-HC status

		If ($bDebugInfo) Then
			AddTrace("       Debug: état MPC-HC=" & $sMPCStatusText , "       Debug: MPC-HC status=" & $sMPCStatusText)
		Endif

		If ($sMPCStatusText <> "") Then
			If ($sMPCStatusText <> $sMPCPreviousStatusText AND $sMPCPreviousStatusText <> "") Then
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
	Endif

	; close MPC-HC
	ProcessClose($sMPCHCPid)
	; wait that it's closed
	ProcessWaitClose($sMPCHCPid, 10)

	If ($bDebugInfo) Then
		AddTrace("       Debug: " & $sMPCHC & " terminé", "       Debug: " & $sMPCHC & " terminated")
	Endif

	; launch VSOInspector
	;;;;;;;;;;;;;;;;;;;;;;
	Local $VSOInspectorPath = IniRead($sIniFile, "Configuration", "VSOInspectorPath", "C:\Program Files (x86)\vso\tools\")
	Local $VSOInspectorApp = IniRead($sIniFile, "Configuration", "VSOInspectorApp", "Inspector.exe")
	local $VSOInspectorCmd = $VSOInspectorPath & $VSOInspectorApp

	If ($bDebugInfo) Then
		AddTrace("       Debug: lancement de " & $VSOInspectorCmd , "       Debug: Launching " & $VSOInspectorCmd)
	Endif
	Local $VSOInspectorPid = Run($VSOInspectorCmd, $VSOInspectorPath)

	If ($VSOInspectorPid == 0) Then
		AddTrace("ERREUR: Lancement de " & $VSOInspectorCmd & " terminé avec @error=" & @error, "ERROR: Launching " & $VSOInspectorCmd & " terminated with @error=" & @error)
		SetAppStatusReady()
		Return
	EndIf

	Sleep(1000)

	Local $sVSOInspectorWindow = IniRead($sIniFile, "Configuration", "VSOInspectorWindow", "VSO Inspector")

	; get handle of window
	Local $aVSOWindows = WinList ($sVSOInspectorWindow)
	$iMax = UBound($aVSOWindows); get array size

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
		AddTrace("ERREUR: Impossible de trouver la fenêtre avec le titre " & $sVSOInspectorWindow, "ERROR: could not find window with title " & $sVSOInspectorWindow)
		SetAppStatusReady()
		Return
	EndIf

	; update window handle reference
	$aDriveArray = StringSplit($sDriveList, "")
	$iButtonIndex = _ArraySearch($aDriveArray, $DriveLetter) - 1
	If ($iButtonIndex > -1) Then
		$aVSOWinHandles[$iButtonIndex] = $hVSOWnd
	Endif

	; move window, so if we have several drives the windows are spread on screen
	SpreadVSOWindow($hVSOWnd, $DriveLetter)

	; now activate the proper page for scan, in VSO
	; get handle for property pages control
	Local $hPPControl = ControlGetHandle($hVSOWnd, "", $VSOPropPageControl)

	If ($hPPControl == 0) Then
		AddTrace("ERREUR: Impossible de trouver le contrôle d'onglets " & $VSOPropPageControl & " @error=" & @error, "ERROR: could not find Property page " & $VSOPropPageControl & " @error=" & @error)
		SetAppStatusReady()
		Return
	EndIf

	Local $ScanTabIndex = IniRead($sIniFile, "Configuration", "ScanTabIndex", "2")
    _GUICtrlTab_ClickTab($hPPControl, $ScanTabIndex)

	; choose proper drive in combo box
	Local $hDriveCombo = ControlGetHandle($hVSOWnd, "", $VSODrivesComboBox)

	If ($hPPControl == 0) Then
		AddTrace("ERREUR: Impossible de trouver le contrôle ComboBox " & $VSODrivesComboBox & " @error=" & @error, "ERROR: could not find Combobox " & $VSODrivesComboBox & " @error=" & @error)
		SetAppStatusReady()
		Return
	EndIf


    Opt("GUIDataSeparatorChar", ",") ; set seperator char to char we want to use
    Local $aComboList = StringSplit(_GUICtrlComboBox_GetList($hDriveCombo), ",")

	Local $ComboSearchedStr = "[" & $DriveLetter & "]" ; string that we search
	Local $ComboStrToSelect = ""
	Local $DriveIndex = -1;
	For $x = 1 To $aComboList[0]
		If StringInStr ($aComboList[$x], $ComboSearchedStr) Then
			$ComboStrToSelect = $aComboList[$x]
			$DriveIndex = $x-1
			ExitLoop
		EndIf
    Next

	If ($ComboStrToSelect <> "") Then

		;_GUICtrlComboBox_SetCurSel($hDriveCombo,$DriveIndex)
		_GUICtrlComboBox_SelectString($hDriveCombo, $ComboStrToSelect)
		sleep(1000)

		ControlCommand($hVSOWnd, "", $VSODrivesComboBox, "SendCommandID", BitShift($CBN_SELCHANGE, -16))

		sleep(3000)

		; it may take some time that $VSOScanButton is available, in case media was just inserted
		$iTotalWait = 0
		While 1
			; check that button is available
			$isScanEnabled = (ControlCommand($hVSOWnd, "", $VSOScanButton, "IsEnabled") == 1)
			if ($isScanEnabled) Then
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
		AddTrace("ERREUR: Impossible de trouver la chaîne " & $ComboSearchedStr & " dans la Combobox des lecteurs" , "ERROR: could not find drive string " & $ComboSearchedStr & " in drives Combobox")
		SetAppStatusReady()
		Return
	EndIf

	If ($bDebugInfo) Then
		AddTrace("       Debug: Drive " & $DriveLetter & " sélectionné dans VSO Inspector", "       Debug: Drive " & $DriveLetter & " selected in VSO Inspector")
	Endif

	; disable File Test check box
	Local $hFileTestControl = ControlGetHandle($hVSOWnd, "", $VSOFilesTestCheckBox)
	Local $nFileTestControlId = _WinAPI_GetDlgCtrlID ($hFileTestControl)

	ControlCommand ($hVSOWnd, "", $VSOFilesTestCheckBox, "UnCheck")

	Sleep(200)

	; launch scan
	ControlClick($hVSOWnd, "", $VSOScanButton)

	If ($bDebugInfo) Then
		AddTrace("       Debug: scan lancé dans VSO Inspector", "       Debug: scan launched in VSO Inspector")
	Endif

	Sleep(1000)

	; minimize window
	WinSetState($hVSOWnd, "", @SW_MINIMIZE )

	AddTrace($sScanLaunchedFR, $sScanLaunchedEN)

	; inform user that program is ready for new disc
	AddTrace("", "")
	AddTrace($sReadyForDiscFR_Long, $sReadyForDiscEN_Long)

	SetAppStatusReady()

EndFunc ;==>PerformScan


; this function moves the VSO Inspector window, so in case of multiple drives, the VSO Inspector windows are spread on screen insted of stacked
Func SpreadVSOWindow($hVSOWnd, $DriveLetter)
	Local $aPos = WinGetPos($hVSOWnd)
;AddDebugTrace("    VSO window: " & $aPos[2] & "|" & $aPos[3])

	Local $iDrivePositionInList = StringInStr ($sDriveList, $DriveLetter)
	Local $iDrivesNumber = StringLen ($sDriveList)

;AddDebugTrace( $DriveLetter)
;AddDebugTrace("    $sDriveList= " & $sDriveList)
;AddDebugTrace("    $iDrivePositionInList= " & $iDrivePositionInList)


	Local $iVSOWindowWidth = $aPos[2]  	; width of VSO Inspector window
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
	Local $iVSOWindowsPerRow = Floor ( (@DeskTopWidth-$iStartX) / $iVSOWindowWidth)

	; number of VSO Inspector windows which can be displayed vertically (number of rows)
	Local $iVSOWindowsMaxRows = Floor ( (@DeskTopHeight-$iStartY) / $iVSOWindowHeight)

;AddDebugTrace("    $iVSOWindowsPerRow= " & $iVSOWindowsPerRow)
;AddDebugTrace("    $iVSOWindowsMaxRows= " & $iVSOWindowsMaxRows)

	; on the screen we can display at same time maximum $iVSOWindowsPerScreen VSO Inspector windows
	; if more windows must be displayed (more drives that what the desktop supports) then we start again from ($iStartX, $iStartY)
	Local $iVSOWindowsPerScreen = $iVSOWindowsPerRow * $iVSOWindowsMaxRows
;AddDebugTrace("    $iVSOWindowsPerScreen= " & $iVSOWindowsPerScreen)

	; so we compute X and Y indexes (starting at 1) of position on screen, in a matrix of ($iVSOWindowsPerRow * $iVSOWindowsMaxRows) windows
	; total number of VSO Screens it $iDrivesNumber

	Local $iWindowIndex = Mod($iDrivePositionInList-1 , $iVSOWindowsPerScreen) + 1

	Local $iXIndex = Mod ($iWindowIndex-1 , $iVSOWindowsPerRow) + 1
	Local $iYIndex = 0
	If ($iVSOWindowsPerRow == 1) Then ; special case
		$iYIndex = $iWindowIndex
	Else
		$iYIndex = Floor (($iWindowIndex -1) / $iVSOWindowsPerRow) + 1
		If ($iYIndex == 0) Then ;special case for first row
			$iYIndex = 1
		EndIf
	Endif

;AddDebugTrace("    $iWindowIndex= " & $iWindowIndex)
;AddDebugTrace("    $iXIndex= " & $iXIndex)
;AddDebugTrace("    $iYIndex= " & $iYIndex)

 	Local $iVSOWinPositionX = $iStartX + ($iXIndex - 1) * $iVSOWindowWidth
	Local $iVSOWinPositionY = $iStartY + ($iYIndex - 1) * $iVSOWindowHeight

;AddDebugTrace("    $iVSOWinPositionX= " & $iVSOWinPositionX)
;AddDebugTrace("    $iVSOWinPositionY= " & $iVSOWinPositionY)

	; move the window
	WinMove($hVSOWnd, "", $iVSOWinPositionX, $iVSOWinPositionY, $iVSOWindowWidth, $iVSOWindowHeight)
EndFunc



Func AddDebugTrace($DebugMsg)
	AddTrace($DebugMsg, $DebugMsg)
EndFunc


; creation of sets of buttons, for drive to monitor, for direct access to the last corresponding created VSO Inspector window
; $aVSOWinMinimizeButtonsIds : ByRef (output) parameter : array containing the buttons ids, same order as $sDriveList. These buttons allow to restore/minimize the VSO window
; $aVSOWinCloseButtonsIds : ByRef (output) parameter : array containing the buttons ids, same order as $sDriveList. These buttons allow to close VSO window
; $aTrayButtonsIds : ByRef (output) parameter : array containing the buttons ids, same order as $sDriveList. These buttons allow to open/close drives trays
; $aVSOWinEndSaveButtonsIds : ByRef (output) parameter : array containing the buttons ids, same order as $sDriveList. These buttons allow to cancel scan in VSO window, save report and close window
; $aVSOWinHandles : ByRef (output) parameter : array of the windows handles, same order as $sDriveList
Func CreateButtonsSets($sDriveList, ByRef $aVSOWinMinimizeButtonsIds, ByRef $aVSOWinCloseButtonsIds, ByRef $aVSOWinEndSaveButtonsIds, ByRef $aTrayButtonsIds, ByRef $aVSOWinHandles)

	Local $iDrivesNumber = StringLen ($sDriveList)

	; $iButtonsAreaX was computed at start of program (Global variable)
	$iControlX = $iDiscAutoScanWinWidth - $iButtonsAreaX + 10
	$iControlY = 4 + 40

	; increment between buttons set
	$iButtonsSetsYIncrement = 40

	$iGroupWidth = $iButtonsAreaX - 20

	GUICtrlCreateGroup ("", $iControlX, $iControlY , $iGroupWidth, 36)
	If $sLang == "FR" Then
		GUICtrlCreateLabel ( "Afficher VSO Insp", $iControlX + 7, $iControlY + 13)
	Else
		GUICtrlCreateLabel ( "Display VSO Insp", $iControlX + 7, $iControlY + 13)
	Endif

	$iControlY += $iButtonsSetsYIncrement

	GUICtrlCreateGroup ("", $iControlX, $iControlY, $iGroupWidth, 36)
	If $sLang == "FR" Then
		GUICtrlCreateLabel ( "Fermer VSO Insp ", $iControlX + 7, $iControlY + 13)
	Else
		GUICtrlCreateLabel ( "Close VSO Insp", $iControlX + 7, $iControlY + 13)
	Endif

	$iControlY += $iButtonsSetsYIncrement

	GUICtrlCreateGroup ("", $iControlX, $iControlY, $iGroupWidth, 36)
	If $sLang == "FR" Then
		GUICtrlCreateLabel ( "Fin scan + enreg", $iControlX + 7, $iControlY + 13)
	Else
		GUICtrlCreateLabel ( "End scan + save", $iControlX + 7, $iControlY + 13)
	Endif

	$iControlY += $iButtonsSetsYIncrement
	$iControlY += 10 ; a bit more space for drive tray buttons


	GUICtrlCreateGroup ("", $iControlX, $iControlY, $iGroupWidth, 36)
	If $sLang == "FR" Then
		GUICtrlCreateLabel ( "Tiroirs drives", $iControlX + 7, $iControlY + 13)
	Else
		GUICtrlCreateLabel ( "Drives trays", $iControlX + 7, $iControlY + 13)
	Endif

	$ButtonX = $iDiscAutoScanWinWidth - $iButtonsAreaX + 110 ; initialization for first button. Buttons created from left to right

	$aDriveArray = StringSplit($sDriveList, "")

	For $i = 0 To ($iDrivesNumber - 1)
		$ButtonY = 12 + 40 ; reset for each drive letter
		Local $DriveLetter = $aDriveArray[$i + 1]

		; create restore/minimize VSO Inspector window button
		; each button is 25*25 pixels, with 10 pixels of space between them
		$idButton = GUICtrlCreateButton($DriveLetter, $ButtonX, $ButtonY, 25, 25)
		$aVSOWinMinimizeButtonsIds[$i] = $idButton

		$ButtonY += $iButtonsSetsYIncrement

		; create VSO Inspector window close button
		$idButton = GUICtrlCreateButton($DriveLetter, $ButtonX, $ButtonY, 25, 25)
		$aVSOWinCloseButtonsIds[$i] = $idButton

		$ButtonY += $iButtonsSetsYIncrement

		; create VSO Inspector scan cancel button
		$idButton = GUICtrlCreateButton($DriveLetter, $ButtonX, $ButtonY, 25, 25)
		$aVSOWinEndSaveButtonsIds[$i] = $idButton

		$ButtonY += $iButtonsSetsYIncrement
		$ButtonY += 10 ; a bit more space for drive tray buttons

		; create drive tray open/close button
		$idButton = GUICtrlCreateButton($DriveLetter, $ButtonX, $ButtonY, 25, 25)
; test for icon
		;$idButton = GUICtrlCreateButton("", $ButtonX, $ButtonY, 25, 25, $BS_ICON)
		;GUICtrlSetImage($idButton, ".\icons\25\ejecter-25.ico")

		$aTrayButtonsIds[$i] = $idButton

		$aVSOWinHandles[$i] = 0 ; init to avoid invalid window handle access

		; increment for next drive
		$ButtonX += 35
	Next

EndFunc

; function (called with timer) which update status for VSO Inspector last launched windows
Func UpdateVSOWinStatus($hWnd, $iMsg, $iIDTimer, $iTime)
	$aDriveArray = StringSplit($sDriveList, "")
	; check all drives
	For $i = 0 To ($iDrivesNumber - 1)
		$hVSOWinHandle = GetValidVSOWinHandle ($i)
		If ($hVSOWinHandle > 0) Then
			$sScanButtonText = ControlGetText($hVSOWinHandle, "", $VSOScanButton)
			$bIsScanRunning = ($sScanButtonText = $sCancelScanText)

			If ($bIsScanRunning) Then
				$aVSOWinStates[$i] = $VSO_WIN_SCANNING
			Else
				; if scan is NOT running but was running just before, we play sound to notify the user
				If ($aVSOWinStates[$i] == $VSO_WIN_SCANNING) Then ; un scan était en cours et s'est arrêté
					If ($sEndScanNotification <> "") Then
						SoundPlay($sEndScanNotification, 0)
					EndIf
				EndIf
				$aVSOWinStates[$i] = $VSO_WIN_IDLE
			EndIf

			; set button color according to scan state
			GUICtrlSetBkColor ($aVSOWinMinimizeButtonsIds[$i], $bIsScanRunning ? $COLOR_SCAN_RUNNING : $COLOR_NO_SCAN_RUNNING)
		Else
			; here $hVSOWinHandle == 0 (no VSO Inspector window)
			If ($aVSOWinStates[$i] <> $VSO_WIN_NO_WINDOW) Then        ; must update state
				GUICtrlSetStyle($aVSOWinMinimizeButtonsIds[$i], 0)   ; default look, which says there's not VSO Inspector window opened
				$aVSOWinStates[$i] = $VSO_WIN_NO_WINDOW
			EndIf
		EndIf ; If ($hVSOWinHandle > 0) Then

	Next ;For $i = 0 To ($iDrivesNumber - 1)

EndFunc

Func ShowHideLog()
      $pos = WinGetPos($hDiscAutoScanGUI)
      WinMove($hDiscAutoScanGUI,"", $pos[0], $pos[1], $pos[2]+10, $pos[3])

	  ;ControlHide ( $hDiscAutoScanGUI, "", $idLogEdit )
	  ;ControlShow ( $hDiscAutoScanGUI, "", $idLogEdit )
EndFunc


; get key from locale ini file of VSO Inspector, according to selected language
Func GetLocalForVSOInsp($sVSOLocaleFile, $Section, $Default)

	$Value = IniRead(@ProgramFilesDir & "\vso\tools\Lang\" & $sVSOLocaleFile, $Section, "locale", "")

	; if $Value is empty -> use "original" key instead of "locale"

	If ($Value == "") Then
		$Value = IniRead(@ProgramFilesDir & "\vso\tools\Lang\" & $sVSOLocaleFile, $Section, "original", $Default)
	EndIf

;AddDebugTrace($Section & "=" & $Value)

	Return $Value
EndFunc

; --

Func SetAppStatusBusy()
	If ($sLang == "FR" ) Then
		GUICtrlSetData ($IdAppStatus, $sBusyTreatingDiscFR)
	Else
		GUICtrlSetData ($IdAppStatus, $sBusyTreatingDiscEN)
	EndIf
	GUICtrlSetColor ($IdAppStatus, $COLOR_APP_BUSY)
EndFunc


Func SetAppStatusReady()
	If ($sLang == "FR" ) Then
		GUICtrlSetData ($IdAppStatus, $sReadyForDiscFR_Short)
	Else
		GUICtrlSetData ($IdAppStatus, $sReadyForDiscEN_Short)
	EndIf
	GUICtrlSetColor ($IdAppStatus, $COLOR_APP_READY)
EndFunc


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

	Return($sRetDriveList)
EndFunc
