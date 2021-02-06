Global Const $WS_VSCROLL = 0x00200000
Global Const $WM_DEVICECHANGE = 0x0219
Global Const $MB_SYSTEMMODAL = 4096
Global Const $_ARRAYCONSTANT_SORTINFOSIZE = 11
Global $__g_aArrayDisplay_SortInfo[$_ARRAYCONSTANT_SORTINFOSIZE]
Global Const $_ARRAYCONSTANT_tagLVITEM = "struct;uint Mask;int Item;int SubItem;uint State;uint StateMask;ptr Text;int TextMax;int Image;lparam Param;" & "int Indent;int GroupID;uint Columns;ptr pColumns;ptr piColFmt;int iGroup;endstruct"
#Au3Stripper_Ignore_Funcs=__ArrayDisplay_SortCallBack
Func __ArrayDisplay_SortCallBack($nItem1, $nItem2, $hWnd)
If $__g_aArrayDisplay_SortInfo[3] = $__g_aArrayDisplay_SortInfo[4] Then
If Not $__g_aArrayDisplay_SortInfo[7] Then
$__g_aArrayDisplay_SortInfo[5] *= -1
$__g_aArrayDisplay_SortInfo[7] = 1
EndIf
Else
$__g_aArrayDisplay_SortInfo[7] = 1
EndIf
$__g_aArrayDisplay_SortInfo[6] = $__g_aArrayDisplay_SortInfo[3]
Local $sVal1 = __ArrayDisplay_GetItemText($hWnd, $nItem1, $__g_aArrayDisplay_SortInfo[3])
Local $sVal2 = __ArrayDisplay_GetItemText($hWnd, $nItem2, $__g_aArrayDisplay_SortInfo[3])
If $__g_aArrayDisplay_SortInfo[8] = 1 Then
If(StringIsFloat($sVal1) Or StringIsInt($sVal1)) Then $sVal1 = Number($sVal1)
If(StringIsFloat($sVal2) Or StringIsInt($sVal2)) Then $sVal2 = Number($sVal2)
EndIf
Local $nResult
If $__g_aArrayDisplay_SortInfo[8] < 2 Then
$nResult = 0
If $sVal1 < $sVal2 Then
$nResult = -1
ElseIf $sVal1 > $sVal2 Then
$nResult = 1
EndIf
Else
$nResult = DllCall('shlwapi.dll', 'int', 'StrCmpLogicalW', 'wstr', $sVal1, 'wstr', $sVal2)[0]
EndIf
$nResult = $nResult * $__g_aArrayDisplay_SortInfo[5]
Return $nResult
EndFunc
Func __ArrayDisplay_GetItemText($hWnd, $iIndex, $iSubItem = 0)
Local $tBuffer = DllStructCreate("wchar Text[4096]")
Local $pBuffer = DllStructGetPtr($tBuffer)
Local $tItem = DllStructCreate($_ARRAYCONSTANT_tagLVITEM)
DllStructSetData($tItem, "SubItem", $iSubItem)
DllStructSetData($tItem, "TextMax", 4096)
DllStructSetData($tItem, "Text", $pBuffer)
If IsHWnd($hWnd) Then
DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $hWnd, "uint", 0x1073, "wparam", $iIndex, "struct*", $tItem)
Else
Local $pItem = DllStructGetPtr($tItem)
GUICtrlSendMsg($hWnd, 0x1073, $iIndex, $pItem)
EndIf
Return DllStructGetData($tBuffer, "Text")
EndFunc
Global Const $MEM_COMMIT = 0x00001000
Global Const $MEM_RESERVE = 0x00002000
Global Const $PAGE_READWRITE = 0x00000004
Global Const $MEM_RELEASE = 0x00008000
Global Const $PROCESS_VM_OPERATION = 0x00000008
Global Const $PROCESS_VM_READ = 0x00000010
Global Const $PROCESS_VM_WRITE = 0x00000020
Global Const $SE_PRIVILEGE_ENABLED = 0x00000002
Global Enum $SECURITYANONYMOUS = 0, $SECURITYIDENTIFICATION, $SECURITYIMPERSONATION, $SECURITYDELEGATION
Global Const $TOKEN_QUERY = 0x00000008
Global Const $TOKEN_ADJUST_PRIVILEGES = 0x00000020
Func _WinAPI_GetLastError(Const $_iCurrentError = @error, Const $_iCurrentExtended = @extended)
Local $aResult = DllCall("kernel32.dll", "dword", "GetLastError")
Return SetError($_iCurrentError, $_iCurrentExtended, $aResult[0])
EndFunc
Func _Security__AdjustTokenPrivileges($hToken, $bDisableAll, $tNewState, $iBufferLen, $tPrevState = 0, $pRequired = 0)
Local $aCall = DllCall("advapi32.dll", "bool", "AdjustTokenPrivileges", "handle", $hToken, "bool", $bDisableAll, "struct*", $tNewState, "dword", $iBufferLen, "struct*", $tPrevState, "struct*", $pRequired)
If @error Then Return SetError(@error, @extended, False)
Return Not($aCall[0] = 0)
EndFunc
Func _Security__ImpersonateSelf($iLevel = $SECURITYIMPERSONATION)
Local $aCall = DllCall("advapi32.dll", "bool", "ImpersonateSelf", "int", $iLevel)
If @error Then Return SetError(@error, @extended, False)
Return Not($aCall[0] = 0)
EndFunc
Func _Security__LookupPrivilegeValue($sSystem, $sName)
Local $aCall = DllCall("advapi32.dll", "bool", "LookupPrivilegeValueW", "wstr", $sSystem, "wstr", $sName, "int64*", 0)
If @error Or Not $aCall[0] Then Return SetError(@error, @extended, 0)
Return $aCall[3]
EndFunc
Func _Security__OpenThreadToken($iAccess, $hThread = 0, $bOpenAsSelf = False)
If $hThread = 0 Then
Local $aResult = DllCall("kernel32.dll", "handle", "GetCurrentThread")
If @error Then Return SetError(@error + 10, @extended, 0)
$hThread = $aResult[0]
EndIf
Local $aCall = DllCall("advapi32.dll", "bool", "OpenThreadToken", "handle", $hThread, "dword", $iAccess, "bool", $bOpenAsSelf, "handle*", 0)
If @error Or Not $aCall[0] Then Return SetError(@error, @extended, 0)
Return $aCall[4]
EndFunc
Func _Security__OpenThreadTokenEx($iAccess, $hThread = 0, $bOpenAsSelf = False)
Local $hToken = _Security__OpenThreadToken($iAccess, $hThread, $bOpenAsSelf)
If $hToken = 0 Then
Local Const $ERROR_NO_TOKEN = 1008
If _WinAPI_GetLastError() <> $ERROR_NO_TOKEN Then Return SetError(20, _WinAPI_GetLastError(), 0)
If Not _Security__ImpersonateSelf() Then Return SetError(@error + 10, _WinAPI_GetLastError(), 0)
$hToken = _Security__OpenThreadToken($iAccess, $hThread, $bOpenAsSelf)
If $hToken = 0 Then Return SetError(@error, _WinAPI_GetLastError(), 0)
EndIf
Return $hToken
EndFunc
Func _Security__SetPrivilege($hToken, $sPrivilege, $bEnable)
Local $iLUID = _Security__LookupPrivilegeValue("", $sPrivilege)
If $iLUID = 0 Then Return SetError(@error + 10, @extended, False)
Local Const $tagTOKEN_PRIVILEGES = "dword Count;align 4;int64 LUID;dword Attributes"
Local $tCurrState = DllStructCreate($tagTOKEN_PRIVILEGES)
Local $iCurrState = DllStructGetSize($tCurrState)
Local $tPrevState = DllStructCreate($tagTOKEN_PRIVILEGES)
Local $iPrevState = DllStructGetSize($tPrevState)
Local $tRequired = DllStructCreate("int Data")
DllStructSetData($tCurrState, "Count", 1)
DllStructSetData($tCurrState, "LUID", $iLUID)
If Not _Security__AdjustTokenPrivileges($hToken, False, $tCurrState, $iCurrState, $tPrevState, $tRequired) Then Return SetError(2, @error, False)
DllStructSetData($tPrevState, "Count", 1)
DllStructSetData($tPrevState, "LUID", $iLUID)
Local $iAttributes = DllStructGetData($tPrevState, "Attributes")
If $bEnable Then
$iAttributes = BitOR($iAttributes, $SE_PRIVILEGE_ENABLED)
Else
$iAttributes = BitAND($iAttributes, BitNOT($SE_PRIVILEGE_ENABLED))
EndIf
DllStructSetData($tPrevState, "Attributes", $iAttributes)
If Not _Security__AdjustTokenPrivileges($hToken, False, $tPrevState, $iPrevState, $tCurrState, $tRequired) Then Return SetError(3, @error, False)
Return True
EndFunc
Global Const $tagPOINT = "struct;long X;long Y;endstruct"
Global Const $tagRECT = "struct;long Left;long Top;long Right;long Bottom;endstruct"
Global Const $tagSECURITY_ATTRIBUTES = "dword Length;ptr Descriptor;bool InheritHandle"
Global Const $tagMEMMAP = "handle hProc;ulong_ptr Size;ptr Mem"
Func _MemFree(ByRef $tMemMap)
Local $pMemory = DllStructGetData($tMemMap, "Mem")
Local $hProcess = DllStructGetData($tMemMap, "hProc")
Local $bResult = _MemVirtualFreeEx($hProcess, $pMemory, 0, $MEM_RELEASE)
DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hProcess)
If @error Then Return SetError(@error, @extended, False)
Return $bResult
EndFunc
Func _MemInit($hWnd, $iSize, ByRef $tMemMap)
Local $aResult = DllCall("user32.dll", "dword", "GetWindowThreadProcessId", "hwnd", $hWnd, "dword*", 0)
If @error Then Return SetError(@error + 10, @extended, 0)
Local $iProcessID = $aResult[2]
If $iProcessID = 0 Then Return SetError(1, 0, 0)
Local $iAccess = BitOR($PROCESS_VM_OPERATION, $PROCESS_VM_READ, $PROCESS_VM_WRITE)
Local $hProcess = __Mem_OpenProcess($iAccess, False, $iProcessID, True)
Local $iAlloc = BitOR($MEM_RESERVE, $MEM_COMMIT)
Local $pMemory = _MemVirtualAllocEx($hProcess, 0, $iSize, $iAlloc, $PAGE_READWRITE)
If $pMemory = 0 Then Return SetError(2, 0, 0)
$tMemMap = DllStructCreate($tagMEMMAP)
DllStructSetData($tMemMap, "hProc", $hProcess)
DllStructSetData($tMemMap, "Size", $iSize)
DllStructSetData($tMemMap, "Mem", $pMemory)
Return $pMemory
EndFunc
Func _MemRead(ByRef $tMemMap, $pSrce, $pDest, $iSize)
Local $aResult = DllCall("kernel32.dll", "bool", "ReadProcessMemory", "handle", DllStructGetData($tMemMap, "hProc"), "ptr", $pSrce, "struct*", $pDest, "ulong_ptr", $iSize, "ulong_ptr*", 0)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _MemVirtualAllocEx($hProcess, $pAddress, $iSize, $iAllocation, $iProtect)
Local $aResult = DllCall("kernel32.dll", "ptr", "VirtualAllocEx", "handle", $hProcess, "ptr", $pAddress, "ulong_ptr", $iSize, "dword", $iAllocation, "dword", $iProtect)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _MemVirtualFreeEx($hProcess, $pAddress, $iSize, $iFreeType)
Local $aResult = DllCall("kernel32.dll", "bool", "VirtualFreeEx", "handle", $hProcess, "ptr", $pAddress, "ulong_ptr", $iSize, "dword", $iFreeType)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func __Mem_OpenProcess($iAccess, $bInherit, $iPID, $bDebugPriv = False)
Local $aResult = DllCall("kernel32.dll", "handle", "OpenProcess", "dword", $iAccess, "bool", $bInherit, "dword", $iPID)
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] Then Return $aResult[0]
If Not $bDebugPriv Then Return SetError(100, 0, 0)
Local $hToken = _Security__OpenThreadTokenEx(BitOR($TOKEN_ADJUST_PRIVILEGES, $TOKEN_QUERY))
If @error Then Return SetError(@error + 10, @extended, 0)
_Security__SetPrivilege($hToken, "SeDebugPrivilege", True)
Local $iError = @error
Local $iExtended = @extended
Local $iRet = 0
If Not @error Then
$aResult = DllCall("kernel32.dll", "handle", "OpenProcess", "dword", $iAccess, "bool", $bInherit, "dword", $iPID)
$iError = @error
$iExtended = @extended
If $aResult[0] Then $iRet = $aResult[0]
_Security__SetPrivilege($hToken, "SeDebugPrivilege", False)
If @error Then
$iError = @error + 20
$iExtended = @extended
EndIf
Else
$iError = @error + 30
EndIf
DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hToken)
Return SetError($iError, $iExtended, $iRet)
EndFunc
Func _SendMessage($hWnd, $iMsg, $wParam = 0, $lParam = 0, $iReturn = 0, $wParamType = "wparam", $lParamType = "lparam", $sReturnType = "lresult")
Local $aResult = DllCall("user32.dll", $sReturnType, "SendMessageW", "hwnd", $hWnd, "uint", $iMsg, $wParamType, $wParam, $lParamType, $lParam)
If @error Then Return SetError(@error, @extended, "")
If $iReturn >= 0 And $iReturn <= 4 Then Return $aResult[$iReturn]
Return $aResult
EndFunc
Global Const $TCM_FIRST = 0x1300
Global Const $TCM_GETITEMRECT =($TCM_FIRST + 10)
Global Const $tagOSVERSIONINFO = 'struct;dword OSVersionInfoSize;dword MajorVersion;dword MinorVersion;dword BuildNumber;dword PlatformId;wchar CSDVersion[128];endstruct'
Global Const $__WINVER = __WINVER()
Func _WinAPI_GetDlgCtrlID($hWnd)
Local $aResult = DllCall("user32.dll", "int", "GetDlgCtrlID", "hwnd", $hWnd)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func __WINVER()
Local $tOSVI = DllStructCreate($tagOSVERSIONINFO)
DllStructSetData($tOSVI, 1, DllStructGetSize($tOSVI))
Local $aRet = DllCall('kernel32.dll', 'bool', 'GetVersionExW', 'struct*', $tOSVI)
If @error Or Not $aRet[0] Then Return SetError(@error, @extended, 0)
Return BitOR(BitShift(DllStructGetData($tOSVI, 2), -8), DllStructGetData($tOSVI, 3))
EndFunc
Func _WinAPI_ClientToScreen($hWnd, ByRef $tPoint)
Local $aRet = DllCall("user32.dll", "bool", "ClientToScreen", "hwnd", $hWnd, "struct*", $tPoint)
If @error Or Not $aRet[0] Then Return SetError(@error + 10, @extended, 0)
Return $tPoint
EndFunc
Func _WinAPI_GetXYFromPoint(ByRef $tPoint, ByRef $iX, ByRef $iY)
$iX = DllStructGetData($tPoint, "X")
$iY = DllStructGetData($tPoint, "Y")
EndFunc
Func _WinAPI_PointFromRect(ByRef $tRECT, $bCenter = True)
Local $iX1 = DllStructGetData($tRECT, "Left")
Local $iY1 = DllStructGetData($tRECT, "Top")
Local $iX2 = DllStructGetData($tRECT, "Right")
Local $iY2 = DllStructGetData($tRECT, "Bottom")
If $bCenter Then
$iX1 = $iX1 +(($iX2 - $iX1) / 2)
$iY1 = $iY1 +(($iY2 - $iY1) / 2)
EndIf
Local $tPoint = DllStructCreate($tagPOINT)
DllStructSetData($tPoint, "X", $iX1)
DllStructSetData($tPoint, "Y", $iY1)
Return $tPoint
EndFunc
Global $__g_aInProcess_WinAPI[64][2] = [[0, 0]]
Func _WinAPI_GetParent($hWnd)
Local $aResult = DllCall("user32.dll", "hwnd", "GetParent", "hwnd", $hWnd)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetWindowThreadProcessId($hWnd, ByRef $iPID)
Local $aResult = DllCall("user32.dll", "dword", "GetWindowThreadProcessId", "hwnd", $hWnd, "dword*", 0)
If @error Then Return SetError(@error, @extended, 0)
$iPID = $aResult[2]
Return $aResult[0]
EndFunc
Func _WinAPI_InProcess($hWnd, ByRef $hLastWnd)
If $hWnd = $hLastWnd Then Return True
For $iI = $__g_aInProcess_WinAPI[0][0] To 1 Step -1
If $hWnd = $__g_aInProcess_WinAPI[$iI][0] Then
If $__g_aInProcess_WinAPI[$iI][1] Then
$hLastWnd = $hWnd
Return True
Else
Return False
EndIf
EndIf
Next
Local $iPID
_WinAPI_GetWindowThreadProcessId($hWnd, $iPID)
Local $iCount = $__g_aInProcess_WinAPI[0][0] + 1
If $iCount >= 64 Then $iCount = 1
$__g_aInProcess_WinAPI[0][0] = $iCount
$__g_aInProcess_WinAPI[$iCount][0] = $hWnd
$__g_aInProcess_WinAPI[$iCount][1] =($iPID = @AutoItPID)
Return $__g_aInProcess_WinAPI[$iCount][1]
EndFunc
Global $__g_hTabLastWnd
Func _GUICtrlTab_ClickTab($hWnd, $iIndex, $sButton = "left", $bMove = False, $iClicks = 1, $iSpeed = 1)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Local $iX, $iY
If Not $bMove Then
Local $hWinParent = _WinAPI_GetParent($hWnd)
Local $avTabPos = _GUICtrlTab_GetItemRect($hWnd, $iIndex)
$iX = $avTabPos[0] +(($avTabPos[2] - $avTabPos[0]) / 2)
$iY = $avTabPos[1] +(($avTabPos[3] - $avTabPos[1]) / 2)
ControlClick($hWinParent, "", $hWnd, $sButton, $iClicks, $iX, $iY)
Else
Local $tRECT = _GUICtrlTab_GetItemRectEx($hWnd, $iIndex)
Local $tPoint = _WinAPI_PointFromRect($tRECT, True)
$tPoint = _WinAPI_ClientToScreen($hWnd, $tPoint)
_WinAPI_GetXYFromPoint($tPoint, $iX, $iY)
Local $iMode = Opt("MouseCoordMode", 1)
MouseClick($sButton, $iX, $iY, $iClicks, $iSpeed)
Opt("MouseCoordMode", $iMode)
EndIf
EndFunc
Func _GUICtrlTab_GetItemRect($hWnd, $iIndex)
Local $aRect[4]
Local $tRECT = _GUICtrlTab_GetItemRectEx($hWnd, $iIndex)
$aRect[0] = DllStructGetData($tRECT, "Left")
$aRect[1] = DllStructGetData($tRECT, "Top")
$aRect[2] = DllStructGetData($tRECT, "Right")
$aRect[3] = DllStructGetData($tRECT, "Bottom")
Return $aRect
EndFunc
Func _GUICtrlTab_GetItemRectEx($hWnd, $iIndex)
Local $tRECT = DllStructCreate($tagRECT)
If IsHWnd($hWnd) Then
If _WinAPI_InProcess($hWnd, $__g_hTabLastWnd) Then
_SendMessage($hWnd, $TCM_GETITEMRECT, $iIndex, $tRECT, 0, "wparam", "struct*")
Else
Local $iRect = DllStructGetSize($tRECT)
Local $tMemMap
Local $pMemory = _MemInit($hWnd, $iRect, $tMemMap)
_SendMessage($hWnd, $TCM_GETITEMRECT, $iIndex, $pMemory, 0, "wparam", "ptr")
_MemRead($tMemMap, $pMemory, $tRECT, $iRect)
_MemFree($tMemMap)
EndIf
Else
GUICtrlSendMsg($hWnd, $TCM_GETITEMRECT, $iIndex, DllStructGetPtr($tRECT))
EndIf
Return $tRECT
EndFunc
Func _Singleton($sOccurrenceName, $iFlag = 0)
Local Const $ERROR_ALREADY_EXISTS = 183
Local Const $SECURITY_DESCRIPTOR_REVISION = 1
Local $tSecurityAttributes = 0
If BitAND($iFlag, 2) Then
Local $tSecurityDescriptor = DllStructCreate("byte;byte;word;ptr[4]")
Local $aRet = DllCall("advapi32.dll", "bool", "InitializeSecurityDescriptor", "struct*", $tSecurityDescriptor, "dword", $SECURITY_DESCRIPTOR_REVISION)
If @error Then Return SetError(@error, @extended, 0)
If $aRet[0] Then
$aRet = DllCall("advapi32.dll", "bool", "SetSecurityDescriptorDacl", "struct*", $tSecurityDescriptor, "bool", 1, "ptr", 0, "bool", 0)
If @error Then Return SetError(@error, @extended, 0)
If $aRet[0] Then
$tSecurityAttributes = DllStructCreate($tagSECURITY_ATTRIBUTES)
DllStructSetData($tSecurityAttributes, 1, DllStructGetSize($tSecurityAttributes))
DllStructSetData($tSecurityAttributes, 2, DllStructGetPtr($tSecurityDescriptor))
DllStructSetData($tSecurityAttributes, 3, 0)
EndIf
EndIf
EndIf
Local $aHandle = DllCall("kernel32.dll", "handle", "CreateMutexW", "struct*", $tSecurityAttributes, "bool", 1, "wstr", $sOccurrenceName)
If @error Then Return SetError(@error, @extended, 0)
Local $aLastError = DllCall("kernel32.dll", "dword", "GetLastError")
If @error Then Return SetError(@error, @extended, 0)
If $aLastError[0] = $ERROR_ALREADY_EXISTS Then
If BitAND($iFlag, 1) Then
DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $aHandle[0])
If @error Then Return SetError(@error, @extended, 0)
Return SetError($aLastError[0], $aLastError[0], 0)
Else
Exit -1
EndIf
EndIf
Return $aHandle[0]
EndFunc
Global Const $GUI_EVENT_CLOSE = -3
Global Const $CB_ERR = -1
Global Const $CB_GETCOUNT = 0x146
Global Const $CB_GETLBTEXT = 0x148
Global Const $CB_GETLBTEXTLEN = 0x149
Global Const $CB_SELECTSTRING = 0x14D
Global Const $CBN_SELCHANGE = 1
Func _GUICtrlComboBox_GetCount($hWnd)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Return _SendMessage($hWnd, $CB_GETCOUNT)
EndFunc
Func _GUICtrlComboBox_GetLBText($hWnd, $iIndex, ByRef $sText)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Local $iLen = _GUICtrlComboBox_GetLBTextLen($hWnd, $iIndex)
Local $tBuffer = DllStructCreate("wchar Text[" & $iLen + 1 & "]")
Local $iRet = _SendMessage($hWnd, $CB_GETLBTEXT, $iIndex, $tBuffer, 0, "wparam", "struct*")
If($iRet == $CB_ERR) Then Return SetError($CB_ERR, $CB_ERR, $CB_ERR)
$sText = DllStructGetData($tBuffer, "Text")
Return $iRet
EndFunc
Func _GUICtrlComboBox_GetLBTextLen($hWnd, $iIndex)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Return _SendMessage($hWnd, $CB_GETLBTEXTLEN, $iIndex)
EndFunc
Func _GUICtrlComboBox_GetList($hWnd)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Local $sDelimiter = Opt("GUIDataSeparatorChar")
Local $sResult = "", $sItem
For $i = 0 To _GUICtrlComboBox_GetCount($hWnd) - 1
_GUICtrlComboBox_GetLBText($hWnd, $i, $sItem)
$sResult &= $sItem & $sDelimiter
Next
Return StringTrimRight($sResult, StringLen($sDelimiter))
EndFunc
Func _GUICtrlComboBox_SelectString($hWnd, $sText, $iIndex = -1)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Return _SendMessage($hWnd, $CB_SELECTSTRING, $iIndex, $sText, 0, "wparam", "wstr")
EndFunc
Global Const $ES_MULTILINE = 4
Global Const $ES_AUTOVSCROLL = 64
Global Const $ES_READONLY = 2048
$DBT_DEVICEARRIVAL = "0x00008000"
$sIniFile = ".\auto-scan-disk.ini"
$sLang = IniRead($sIniFile, "Configuration", "Lang", "EN")
$sDebugInfo = IniRead($sIniFile, "Configuration", "DebugInfo", "N")
$bDebugInfo =($sDebugInfo == "Y")
If _Singleton("Autoscandisk", 1) = 0 Then
If($sLang == "FR") Then
MsgBox($MB_SYSTEMMODAL, "Error", "Une instance de auto-scan-disk est déjà en cours d'éxécution")
Else
MsgBox($MB_SYSTEMMODAL, "Error", "An occurrence of auto-scan-disk is already running")
EndIf
Exit
EndIf
$idExit=0
GUICreate("Auto scan disk", 700, 400)
If($sLang == "FR") Then
$idExit = GUICtrlCreateButton("Quitter", 640, 375, 50, 20)
Else
$idExit = GUICtrlCreateButton("Exit", 640, 375, 50, 20)
EndIf
$idLogEdit = GUICtrlCreateEdit("", 10, 10, 680, 360, $ES_AUTOVSCROLL + $WS_VSCROLL + $ES_READONLY + $ES_MULTILINE)
GUISetState()
$sDriveList = IniRead($sIniFile, "Configuration", "DriveList", "D")
AddTrace("  Lecteurs configuré pour autoscan: " & $sDriveList, "  Drives configured for autoscan: " & $sDriveList)
GUIRegisterMsg($WM_DEVICECHANGE , "DeviceChange")
AddTrace("", "")
AddTrace("-PRET, vous pouvez insérer UN disque puis ATTENDRE", "-READY, you can insert ONE disc then WAIT")
Do
$GuiMsg = GUIGetMsg()
Until $GuiMsg = $GUI_EVENT_CLOSE Or $GuiMsg = $idExit
Exit
Func AddTrace($FrMsg, $EnMsg)
If $sLang == "FR" Then
GUICtrlSetData($idLogEdit, $FrMsg & @CRLF, 1)
Else
GUICtrlSetData($idLogEdit, $EnMsg & @CRLF, 1)
EndIf
EndFunc
Func DeviceChange($hWndGUI, $MsgID, $WParam, $LParam)
If $WParam == $DBT_DEVICEARRIVAL Then
Local Const $tagDEV_BROADCAST_VOLUME = "dword dbcv_size; dword dbcv_devicetype; dword dbcv_reserved; dword dbcv_unitmask; word dbcv_flags"
Local Const $DEV_BROADCAST_VOLUME = DllStructCreate($tagDEV_BROADCAST_VOLUME, $lParam)
Local Const $DeviceType = DllStructGetData($DEV_BROADCAST_VOLUME, "dbcv_devicetype")
Local Const $UnitMask = DllStructGetData($DEV_BROADCAST_VOLUME, "dbcv_unitmask")
Local Const $DriveLetter = GetDriveLetterFromUnitMask($UnitMask)
if StringInStr($sDriveList, $DriveLetter) Then
AddTrace("  Disque inséré dans lecteur " & $DriveLetter, "  Disc inserted in drive " & $DriveLetter)
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
$Pom = Int($Pom / 2)
$Count += 1
WEnd
If $Count >= 1 And $Count <= 26 Then
Return $Drives[$Count - 1]
Else
Return SetError(-1, 0, '?')
EndIf
EndFunc
Func PerformScan($DriveLetter)
Local $sMPCHC = IniRead($sIniFile, "Configuration", "MPCHC", "C:\Program Files\MPC-HC\mpc-hc64.exe")
Local $sMPCHCCmd = """" & $sMPCHC & """ " & $DriveLetter & ": /open /minimized /new"
If($bDebugInfo) Then
AddTrace("  Debug: lancement de " & $sMPCHCCmd , "  Debug: Launching " & $sMPCHCCmd)
Endif
Local $sMPCHCPid = Run($sMPCHCCmd )
If($sMPCHCPid == 0) Then
AddTrace("ERREUR: Lancement de " & $sMPCHCCmd & " terminé avec @error=" & @error, "ERROR: Launching " & $sMPCHCCmd & " terminated with @error=" & @error)
Return
EndIf
Local $sMPCHCWindow = IniRead($sIniFile, "Configuration", "MPCHCWindow", "Media Player Classic Home Cinema")
Local $hMPCHC = WinWait("[CLASS:"& $sMPCHCWindow &"]", "", 10)
sleep(5000)
If($bDebugInfo) Then
AddTrace("  Debug: Fermeture de " & $sMPCHC, "  Debug: Closing " & $sMPCHC)
Endif
ProcessClose($sMPCHCPid)
Sleep(1000)
Local $sVSOInspector = IniRead($sIniFile, "Configuration", "VSOInspector", "C:\Program Files (x86)\vso\tools\Inspector.exe")
If($bDebugInfo) Then
AddTrace("  Debug: lancement de " & $sVSOInspector , "  Debug: Launching " & $sVSOInspector)
Endif
Local $VSOInspectorPid = Run($sVSOInspector)
If($VSOInspectorPid == 0) Then
AddTrace("ERREUR: Lancement de " & $sVSOInspector & " terminé avec @error=" & @error, "ERROR: Launching " & $sVSOInspector & " terminated with @error=" & @error)
Return
EndIf
Sleep(1000)
Local $sVSOInspectorWindow = IniRead($sIniFile, "Configuration", "VSOInspectorWindow", "VSO Inspector")
Local $aVSOWindows = WinList($sVSOInspectorWindow)
$iMax = UBound($aVSOWindows)
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
Local $hPPControl = ControlGetHandle($hVSOWnd, "", "TPageControl1")
If($hPPControl == 0) Then
AddTrace("ERREUR: Impossible de trouver le contrôle d'onglets TPageControl1 @error=" & @error, "ERROR: could not find Property page TPageControl1 @error=" & @error)
Return
EndIf
Local $ScanTabIndex = IniRead($sIniFile, "Configuration", "ScanTabIndex", "2")
_GUICtrlTab_ClickTab($hPPControl, $ScanTabIndex)
Local $hDriveCombo = ControlGetHandle($hVSOWnd, "", "TComboBox1")
If($hPPControl == 0) Then
AddTrace("ERREUR: Impossible de trouver le contrôle ComboBox TComboBox1 @error=" & @error, "ERROR: could not find Combobox TComboBox @error=" & @error)
Return
EndIf
Opt("GUIDataSeparatorChar", ",")
Local $aComboList = StringSplit(_GUICtrlComboBox_GetList($hDriveCombo), ",")
Local $ComboSearchedStr = "[" & $DriveLetter & "]"
Local $ComboStrToSelect = ""
Local $DriveIndex = -1
For $x = 1 To $aComboList[0]
If StringInStr($aComboList[$x], $ComboSearchedStr) Then
$ComboStrToSelect = $aComboList[$x]
$DriveIndex = $x-1
ExitLoop
EndIf
Next
If($ComboStrToSelect <> "") Then
_GUICtrlComboBox_SelectString($hDriveCombo, $ComboStrToSelect)
sleep(1000)
ControlCommand($hVSOWnd, "", "TComboBox1", "SendCommandID", BitShift($CBN_SELCHANGE, -16))
sleep(3000)
Else
AddTrace("ERREUR: Impossible de trouver la chaîne " & $ComboSearchedStr & " dans la Combobox des lecteurs" , "ERROR: could not find drive string " & $ComboSearchedStr & " in drives Combobox")
Return
EndIf
Local $hFileTestControl = ControlGetHandle($hVSOWnd, "", "TJvCheckBox1")
Local $nFileTestControlId = _WinAPI_GetDlgCtrlID($hFileTestControl)
ControlCommand($hVSOWnd, "", "TJvCheckBox1", "UnCheck")
Sleep(200)
ControlClick($hVSOWnd, "", "TButton4")
Sleep(2000)
WinSetState($hVSOWnd, "", @SW_MINIMIZE )
AddTrace("-Scan en cours... Vous pourrez fermer la fenêtre VSO Inspector lorsque scan terminé", "-Scan on going... You can close the VSO Inspector window when finished")
AddTrace("", "")
AddTrace("-PRET, vous pouvez insérer UN disque puis ATTENDRE", "-READY, you can insert ONE disc then WAIT")
EndFunc
