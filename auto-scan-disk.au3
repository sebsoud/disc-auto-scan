#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=auto-scan-disk.ico
#AutoIt3Wrapper_Res_Fileversion=0.1
#AutoIt3Wrapper_Res_Language=1036
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=auto_scan_disk.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=auto_scan_disk_nMu_icon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
; by seb

#include <WindowsConstants.au3>
#include <Array.au3> ; Required for _ArrayDisplay() only.
#include <GuiTab.au3>
#include <WinAPISysWin.au3>
#include <GuiComboBox.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>


$DBT_DEVICEARRIVAL = "0x00008000"

$sIniFile = ".\auto-scan-disk.ini"

; language (EN or FR)
$sLang = IniRead($sIniFile, "Configuration", "Lang", "EN")

; log debug info
$sDebugInfo = IniRead($sIniFile, "Configuration", "DebugInfo", "N")
$bDebugInfo = ($sDebugInfo == "Y")

$idExit=0

GUICreate("Auto scan disk", 700, 400)

If ($sLang == "FR") Then
	$idExit = GUICtrlCreateButton("Quitter", 640, 375, 50, 20)
Else
	$idExit = GUICtrlCreateButton("Exit", 640, 375, 50, 20)
EndIf
$idLogEdit = GUICtrlCreateEdit("", 10, 10, 680, 360, $ES_AUTOVSCROLL + $WS_VSCROLL + $ES_READONLY + $ES_MULTILINE)

GUISetState() ; display the GUI

;GUICtrlSetData($idLogEdit, "log" & @CRLF, 1)




; ---------------------------------

; direct test (avoiding to have to insert a disc)
;PerformScan("G")
;Exit

;---------------------------------



$sDriveList = IniRead($sIniFile, "Configuration", "DriveList", "D")

AddTrace("  Lecteurs configuré pour autoscan: " & $sDriveList, "  Drives configured for autoscan: " & $sDriveList)

GUIRegisterMsg($WM_DEVICECHANGE , "DeviceChange")

AddTrace("", "")
AddTrace("-PRET, vous pouvez insérer UN disque puis ATTENDRE", "-READY, you can insert ONE disc then WAIT")


Do
    $GuiMsg = GUIGetMsg()
Until $GuiMsg = $GUI_EVENT_CLOSE Or $GuiMsg = $idExit


Exit

; adding trace in window
Func AddTrace($FrMsg, $EnMsg)
	If $sLang == "FR" Then
		GUICtrlSetData($idLogEdit, $FrMsg & @CRLF, 1)
	Else
		GUICtrlSetData($idLogEdit, $EnMsg & @CRLF, 1)
	EndIf
EndFunc


; handler of $WM_DEVICECHANGE
Func DeviceChange($hWndGUI, $MsgID, $WParam, $LParam)
    If $WParam == $DBT_DEVICEARRIVAL Then
        ; Create a struct from $lParam which contains a pointer to a Windows-created struct.

        Local Const $tagDEV_BROADCAST_VOLUME = "dword dbcv_size; dword dbcv_devicetype; dword dbcv_reserved; dword dbcv_unitmask; word dbcv_flags"
        Local Const $DEV_BROADCAST_VOLUME = DllStructCreate($tagDEV_BROADCAST_VOLUME, $lParam)

        Local Const $DeviceType = DllStructGetData($DEV_BROADCAST_VOLUME, "dbcv_devicetype")
		Local Const $UnitMask = DllStructGetData($DEV_BROADCAST_VOLUME, "dbcv_unitmask")

		Local Const $DriveLetter = GetDriveLetterFromUnitMask($UnitMask)

		if StringInStr ($sDriveList, $DriveLetter) Then
			AddTrace("  Disque inséré dans lecteur " & $DriveLetter, "  Disc inserted in drive " & $DriveLetter)

			; wait enough time that disc is recognized by windows
			Sleep(7000)

			PerformScan($DriveLetter)
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

	; launch MPC-HC, so dvd is accessible if protected
	Local $sMPCHC = IniRead($sIniFile, "Configuration", "MPCHC", "C:\Program Files\MPC-HC\mpc-hc64.exe")
	Local $sMPCHCCmd = """" & $sMPCHC &  """ " & $DriveLetter & ": /open /minimized /new"

	If ($bDebugInfo) Then
		AddTrace("  Debug: lancement de " & $sMPCHCCmd , "  Debug: Launching " & $sMPCHCCmd)
	Endif

	Local $sMPCHCPid = Run($sMPCHCCmd )

	If ($sMPCHCPid == 0) Then
		AddTrace("ERREUR: Lancement de " & $sMPCHCCmd & " terminé avec @error=" & @error, "ERROR: Launching " & $sMPCHCCmd & " terminated with @error=" & @error)
		Return
	EndIf

	Local $sMPCHCWindow = IniRead($sIniFile, "Configuration", "MPCHCWindow", "Media Player Classic Home Cinema")

	Local $hMPCHC = WinWait("[CLASS:"& $sMPCHCWindow &"]", "", 10)

	; wait for a few seconds that dvd movie is launched
	sleep (5000)

	If ($bDebugInfo) Then
		AddTrace("  Debug: Fermeture de " & $sMPCHC, "  Debug: Closing " & $sMPCHC)
	Endif
	ProcessClose($sMPCHCPid)
	Sleep(1000)

	; launch VSOInspector
	Local $sVSOInspector = IniRead($sIniFile, "Configuration", "VSOInspector", "C:\Program Files (x86)\vso\tools\Inspector.exe")
	If ($bDebugInfo) Then
		AddTrace("  Debug: lancement de " & $sVSOInspector , "  Debug: Launching " & $sVSOInspector)
	Endif
	Local $VSOInspectorPid = Run($sVSOInspector)

	If ($VSOInspectorPid == 0) Then
		AddTrace("ERREUR: Lancement de " & $sVSOInspector & " terminé avec @error=" & @error, "ERROR: Launching " & $sVSOInspector & " terminated with @error=" & @error)
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
		Return
	EndIf


	; get handle for property pages control
	Local $hPPControl = ControlGetHandle($hVSOWnd, "", "TPageControl1")

	If ($hPPControl == 0) Then
		AddTrace("ERREUR: Impossible de trouver le contrôle d'onglets TPageControl1 @error=" & @error, "ERROR: could not find Property page TPageControl1 @error=" & @error)
		Return
	EndIf


	Local $ScanTabIndex = IniRead($sIniFile, "Configuration", "ScanTabIndex", "2")
    _GUICtrlTab_ClickTab($hPPControl, $ScanTabIndex)

	; choose proper drive in combo box
	Local $hDriveCombo = ControlGetHandle($hVSOWnd, "", "TComboBox1")

	If ($hPPControl == 0) Then
		AddTrace("ERREUR: Impossible de trouver le contrôle ComboBox TComboBox1 @error=" & @error, "ERROR: could not find Combobox TComboBox @error=" & @error)
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

;AddTrace("  $DriveIndex= " & $DriveIndex)
;AddTrace("  $ComboStrToSelect= " & $ComboStrToSelect)

	If ($ComboStrToSelect <> "") Then

		;_GUICtrlComboBox_SetCurSel($hDriveCombo,$DriveIndex)
		_GUICtrlComboBox_SelectString($hDriveCombo, $ComboStrToSelect)
		sleep(1000)

		ControlCommand($hVSOWnd, "", "TComboBox1", "SendCommandID", BitShift($CBN_SELCHANGE, -16))

		sleep(3000)

	Else
		; drive not found!!!
		AddTrace("ERREUR: Impossible de trouver la chaîne " & $ComboSearchedStr & " dans la Combobox des lecteurs" , "ERROR: could not find drive string " & $ComboSearchedStr & " in drives Combobox")
		Return
	EndIf


	; disable File Test check box
	Local $hFileTestControl = ControlGetHandle($hVSOWnd, "", "TJvCheckBox1")
	Local $nFileTestControlId = _WinAPI_GetDlgCtrlID ($hFileTestControl)


	ControlCommand ($hVSOWnd, "", "TJvCheckBox1", "UnCheck")

	Sleep(200)

	; launch scan
	ControlClick($hVSOWnd, "", "TButton4")

	Sleep(2000)

	; minimize window
	WinSetState($hVSOWnd, "", @SW_MINIMIZE )

	AddTrace("-Scan en cours... Vous pourrez fermer la fenêtre VSO Inspector lorsque scan terminé", "-Scan on going... You can close the VSO Inspector window when finished")

	AddTrace("", "")
	AddTrace("-PRET, vous pouvez insérer UN disque puis ATTENDRE", "-READY, you can insert ONE disc then WAIT")

EndFunc ;==>PerformScan

; --