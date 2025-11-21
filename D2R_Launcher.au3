#RequireAdmin
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <WinAPIFiles.au3>

; Global Constants
Global Const $INI_FILE = @ScriptDir & "\D2R_Launcher.ini"
Global Const $TITLE = "D2R Multi-Launcher"

; Global Variables
Global $g_hGui, $g_hListView
Global $g_idBtnLaunch, $g_idBtnAdd, $g_idBtnEdit, $g_idBtnDelete, $g_idBtnSettings
Global $g_sHandle64Path, $g_sD2RFolder, $g_sBnetConfigPath
Global $g_sLogFile

; Load Settings
LoadSettings()

; Init logging
InitLogging()

; Create GUI
CreateMainGUI()

; Main Loop
While 1
    Switch GUIGetMsg()
        Case $GUI_EVENT_CLOSE
            ExitLoop
        Case $g_idBtnLaunch
            LaunchSelectedAccount()
        Case $g_idBtnAdd
            AddAccount()
        Case $g_idBtnEdit
            EditAccount()
        Case $g_idBtnDelete
            DeleteAccount()
        Case $g_idBtnSettings
            ShowSettings()
    EndSwitch
WEnd

; Functions

Func CreateMainGUI()
    $g_hGui = GUICreate($TITLE, 600, 400)

    $g_hListView = GUICtrlCreateListView("Profile Name|Account Name|Server|Mod", 10, 10, 580, 300)
    _GUICtrlListView_SetColumnWidth($g_hListView, 0, 150)
    _GUICtrlListView_SetColumnWidth($g_hListView, 1, 200)
    _GUICtrlListView_SetColumnWidth($g_hListView, 2, 100)
    _GUICtrlListView_SetColumnWidth($g_hListView, 3, 100)

    $g_idBtnLaunch = GUICtrlCreateButton("Launch", 10, 320, 100, 30)
    $g_idBtnAdd = GUICtrlCreateButton("Add Account", 120, 320, 100, 30)
    $g_idBtnEdit = GUICtrlCreateButton("Edit Account", 230, 320, 100, 30)
    $g_idBtnDelete = GUICtrlCreateButton("Delete Account", 340, 320, 100, 30)
    $g_idBtnSettings = GUICtrlCreateButton("Settings", 490, 320, 100, 30)

    RefreshAccountList()

    GUISetState(@SW_SHOW, $g_hGui)
EndFunc

Func InitLogging()
    $g_sLogFile = @ScriptDir & "\D2R_launcher.log"
    ; Create/clear log on start
    Local $h = FileOpen($g_sLogFile, $FO_OVERWRITE)
    If $h <> -1 Then
        FileWriteLine($h, @MON & " " & @MDAY & " " & @YEAR & " - Launcher started")
        FileClose($h)
    EndIf
EndFunc

Func WriteLog($sMsg)
    Local $h = FileOpen($g_sLogFile, $FO_APPEND)
    If $h <> -1 Then
        FileWriteLine($h, @HOUR & ":" & @MIN & ":" & @SEC & " - " & $sMsg)
        FileClose($h)
    EndIf
EndFunc

Func LoadSettings()
    $g_sHandle64Path = IniRead($INI_FILE, "Settings", "Handle64Path", "C:\Program Files (x86)\Diablo II Resurrected")
    $g_sD2RFolder = IniRead($INI_FILE, "Settings", "D2RFolder", "C:\Program Files (x86)\Diablo II Resurrected")
    $g_sBnetConfigPath = IniRead($INI_FILE, "Settings", "BnetConfigPath", "C:\Users\" & @UserName & "\AppData\Roaming\Battle.net\Battle.net.config")

    ; If the script is placed in the D2R installation folder, prefer that folder (mirrors PS script behavior)
    If FileExists(@ScriptDir & "\D2R.exe") Then
        $g_sD2RFolder = @ScriptDir
        WriteLog('Auto-detected D2RFolder = ' & $g_sD2RFolder)
    EndIf
EndFunc

Func SaveSettings()
    IniWrite($INI_FILE, "Settings", "Handle64Path", $g_sHandle64Path)
    IniWrite($INI_FILE, "Settings", "D2RFolder", $g_sD2RFolder)
    IniWrite($INI_FILE, "Settings", "BnetConfigPath", $g_sBnetConfigPath)
EndFunc

Func RefreshAccountList()
    _GUICtrlListView_DeleteAllItems($g_hListView)
    Local $aSections = IniReadSectionNames($INI_FILE)
    If @error Then Return

    For $i = 1 To $aSections[0]
        If $aSections[$i] <> "Settings" Then
            Local $sProfile = $aSections[$i]
            Local $sAccount = IniRead($INI_FILE, $sProfile, "Account", "")
            Local $sServer = IniRead($INI_FILE, $sProfile, "Server", "")
            Local $sMod = IniRead($INI_FILE, $sProfile, "Mod", "")
            GUICtrlCreateListViewItem($sProfile & "|" & $sAccount & "|" & $sServer & "|" & $sMod, $g_hListView)
        EndIf
    Next
EndFunc

Func ShowSettings()
    Local $hSettingsGui = GUICreate("Global Settings", 500, 200, -1, -1, -1, -1, $g_hGui)
    
    GUICtrlCreateLabel("Handle64 Path:", 10, 20, 100, 20)
    Local $idInputHandle = GUICtrlCreateInput($g_sHandle64Path, 120, 18, 300, 20)
    Local $idBtnBrowseHandle = GUICtrlCreateButton("...", 430, 18, 30, 20)

    GUICtrlCreateLabel("D2R Folder:", 10, 50, 100, 20)
    Local $idInputD2R = GUICtrlCreateInput($g_sD2RFolder, 120, 48, 300, 20)
    Local $idBtnBrowseD2R = GUICtrlCreateButton("...", 430, 48, 30, 20)

    GUICtrlCreateLabel("Bnet Config:", 10, 80, 100, 20)
    Local $idInputConfig = GUICtrlCreateInput($g_sBnetConfigPath, 120, 78, 300, 20)
    Local $idBtnBrowseConfig = GUICtrlCreateButton("...", 430, 78, 30, 20)

    Local $idBtnSave = GUICtrlCreateButton("Save", 150, 130, 80, 30)
    Local $idBtnCancel = GUICtrlCreateButton("Cancel", 250, 130, 80, 30)

    GUISetState(@SW_SHOW, $hSettingsGui)

    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE, $idBtnCancel
                GUIDelete($hSettingsGui)
                ExitLoop
            Case $idBtnSave
                $g_sHandle64Path = GUICtrlRead($idInputHandle)
                $g_sD2RFolder = GUICtrlRead($idInputD2R)
                $g_sBnetConfigPath = GUICtrlRead($idInputConfig)
                SaveSettings()
                GUIDelete($hSettingsGui)
                ExitLoop
            Case $idBtnBrowseHandle
                Local $sFile = FileOpenDialog("Select Handle64 Folder", "", "Executables (*.exe)|All (*.*)")
                If @error Then ContinueLoop
                GUICtrlSetData($idInputHandle, StringLeft($sFile, StringInStr($sFile, "\", 0, -1) - 1)) ; Store folder, logic assumes handle64.exe is in it or user selects file? User said "Handle64 location". Let's assume folder.
            Case $idBtnBrowseD2R
                Local $sFolder = FileSelectFolder("Select D2R Folder", "")
                If @error Then ContinueLoop
                GUICtrlSetData($idInputD2R, $sFolder)
            Case $idBtnBrowseConfig
                Local $sFile = FileOpenDialog("Select Battle.net.config", "", "Config (*.config)|All (*.*)")
                If @error Then ContinueLoop
                GUICtrlSetData($idInputConfig, $sFile)
        EndSwitch
    WEnd
EndFunc

Func AddAccount()
    ShowAccountDialog("", "", "", "", "")
EndFunc

Func EditAccount()
    Local $idItem = GUICtrlRead($g_hListView)
    If $idItem = 0 Then
        MsgBox($MB_ICONWARNING, "Warning", "Please select an account to edit.")
        Return
    EndIf
    
    Local $sText = GUICtrlRead($idItem)
    Local $aData = StringSplit($sText, "|")
    Local $sProfile = $aData[1]
    
    Local $sAccount = IniRead($INI_FILE, $sProfile, "Account", "")
    Local $sPassword = IniRead($INI_FILE, $sProfile, "Password", "")
    Local $sServer = IniRead($INI_FILE, $sProfile, "Server", "")
    Local $sMod = IniRead($INI_FILE, $sProfile, "Mod", "")
    
    ShowAccountDialog($sProfile, $sAccount, $sPassword, $sServer, $sMod, True)
EndFunc

Func DeleteAccount()
    Local $idItem = GUICtrlRead($g_hListView)
    If $idItem = 0 Then
        MsgBox($MB_ICONWARNING, "Warning", "Please select an account to delete.")
        Return
    EndIf
    
    Local $sText = GUICtrlRead($idItem)
    Local $aData = StringSplit($sText, "|")
    Local $sProfile = $aData[1]
    
    If MsgBox($MB_YESNO + $MB_ICONQUESTION, "Confirm", "Are you sure you want to delete profile '" & $sProfile & "'?") = $IDYES Then
        IniDelete($INI_FILE, $sProfile)
        RefreshAccountList()
    EndIf
EndFunc

Func ShowAccountDialog($sProfile, $sAccount, $sPassword, $sServer, $sMod, $bEdit = False)
    Local $hAccGui = GUICreate($bEdit ? "Edit Account" : "Add Account", 400, 250, -1, -1, -1, -1, $g_hGui)

    GUICtrlCreateLabel("Profile Name:", 10, 20, 100, 20)
    Local $idInputProfile = GUICtrlCreateInput($sProfile, 120, 18, 250, 20)
    If $bEdit Then GUICtrlSetState($idInputProfile, $GUI_DISABLE)

    GUICtrlCreateLabel("Account (Email):", 10, 50, 100, 20)
    Local $idInputAccount = GUICtrlCreateInput($sAccount, 120, 48, 250, 20)

    GUICtrlCreateLabel("Password:", 10, 80, 100, 20)
    Local $idInputPassword = GUICtrlCreateInput($sPassword, 120, 78, 250, 20, 0x0020) ; Password style

    GUICtrlCreateLabel("Server (Region):", 10, 110, 100, 20)
    Local $idComboServer = GUICtrlCreateCombo($sServer, 120, 108, 250, 20)
    GUICtrlSetData($idComboServer, "us|eu|asia", "asia")

    GUICtrlCreateLabel("Mod / Args (Optional):", 10, 140, 150, 20)
    Local $idInputMod = GUICtrlCreateInput($sMod, 120, 138, 250, 20)

    Local $idBtnSave = GUICtrlCreateButton("Save", 100, 190, 80, 30)
    Local $idBtnCancel = GUICtrlCreateButton("Cancel", 220, 190, 80, 30)

    GUISetState(@SW_SHOW, $hAccGui)

    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE, $idBtnCancel
                GUIDelete($hAccGui)
                ExitLoop
            Case $idBtnSave
                Local $sNewProfile = GUICtrlRead($idInputProfile)
                If $sNewProfile = "" Then
                    MsgBox($MB_ICONERROR, "Error", "Profile Name cannot be empty.")
                    ContinueLoop
                EndIf
                
                IniWrite($INI_FILE, $sNewProfile, "Account", GUICtrlRead($idInputAccount))
                IniWrite($INI_FILE, $sNewProfile, "Password", GUICtrlRead($idInputPassword))
                IniWrite($INI_FILE, $sNewProfile, "Server", GUICtrlRead($idComboServer))
                IniWrite($INI_FILE, $sNewProfile, "Mod", GUICtrlRead($idInputMod))
                
                RefreshAccountList()
                GUIDelete($hAccGui)
                ExitLoop
        EndSwitch
    WEnd
EndFunc

Func LaunchSelectedAccount()
    Local $idItem = GUICtrlRead($g_hListView)
    If $idItem = 0 Then
        MsgBox($MB_ICONWARNING, "Warning", "Please select an account to launch.")
        Return
    EndIf

    Local $sText = GUICtrlRead($idItem)
    Local $aData = StringSplit($sText, "|")
    Local $sProfile = $aData[1]

    Local $sAccount = IniRead($INI_FILE, $sProfile, "Account", "")
    Local $sPassword = IniRead($INI_FILE, $sProfile, "Password", "")
    Local $sServer = IniRead($INI_FILE, $sProfile, "Server", "asia")
    Local $sMod = IniRead($INI_FILE, $sProfile, "Mod", "")

    ; 1. Map Server Region
    Local $sRegionPrefix = $sServer
    If $sServer = "asia" Then $sRegionPrefix = "kr"
    
    ; 2. Update Battle.net Config
    UpdateBnetConfig($sAccount)

    ; 3. Launch D2R
    Local $sD2RExe = $g_sD2RFolder & "\D2R.exe"
    If Not FileExists($sD2RExe) Then
        MsgBox($MB_ICONERROR, "Error", "D2R.exe not found at: " & $sD2RExe)
        Return
    EndIf

    ; Build parameter string with proper quoting
    Local $sParams = StringFormat('-username "%s" -password "%s" -address "%s.actual.battle.net"', $sAccount, $sPassword, $sRegionPrefix)

    ; Handle Mod / Extra Args
    If $sMod = "" Then
        $sParams &= ' -ns -w -txt'
    Else
        $sParams &= ' ' & $sMod
    EndIf

    ; Launch with Run
    Local $iPID = Run('"' & $sD2RExe & '" ' & $sParams, $g_sD2RFolder, @SW_SHOW)
    If $iPID = 0 Then
        MsgBox($MB_ICONERROR, "Error", "Failed to launch D2R.exe")
        Return
    EndIf
    WriteLog('Launched: "' & $sD2RExe & '" ' & $sParams & '  (PID=' & $iPID & ')')

    ; 4. Post-Launch Handle Kill (Wait for handle to be created)
    ; Advice suggests waiting for the handle to appear.
    WriteLog('Waiting 5 seconds for handle creation...')
    Sleep(5000) 
    KillD2RHandleForPID($iPID)

    ; 5. Rename Window
    Local $hWindow = 0
    Local $iTimeout = 30 ; seconds
    Local $iTimer = TimerInit()
    
    While TimerDiff($iTimer) < $iTimeout * 1000
        Local $aList = WinList("Diablo II: Resurrected")
        For $i = 1 To $aList[0][0]
            Local $iWinPID = 0
            GetWindowThreadProcessId($aList[$i][1], $iWinPID)
            If $iWinPID = $iPID Then
                $hWindow = $aList[$i][1]
                ExitLoop 2
            EndIf
        Next
        Sleep(500)
    WEnd

    If $hWindow Then
        WinSetTitle($hWindow, "", $sProfile)
    EndIf

EndFunc

Func KillD2RHandleForPID($iPID)
    Local $sHandleExe = $g_sHandle64Path & "\handle64.exe"
    If Not FileExists($sHandleExe) Then
        $sHandleExe = $g_sHandle64Path
        If Not StringInStr($sHandleExe, "handle64.exe") Then $sHandleExe &= "\handle64.exe"
        If Not FileExists($sHandleExe) Then
            WriteLog("Error: handle64.exe not found at " & $sHandleExe)
            Return
        EndIf
    EndIf

    ; Run handle64 to find handles for specific PID
    WriteLog('Running handle64 for PID ' & $iPID)
    Local $iH64PID = Run('"' & $sHandleExe & '" -accepteula -a -p ' & $iPID, "", @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
    ProcessWaitClose($iH64PID)
    Local $sOutput = StdoutRead($iH64PID)
    
    ; Parse output for "Check For Other Instances"
    Local $aLines = StringSplit($sOutput, @CRLF, 1)
    For $i = 1 To $aLines[0]
        Local $sLine = $aLines[$i]
        If StringInStr($sLine, "Check For Other Instances") Then
            ; Extract Handle ID (Hex)
            ; Format:   12C: Event  \Sessions\...\DiabloII Check For Other Instances
            Local $sHandleMatch = StringRegExp($sLine, '\s*([0-9A-Fa-f]+):\s+Event.*DiabloII Check For Other Instances', 3)
            If IsArray($sHandleMatch) Then 
                Local $sHandleID = $sHandleMatch[0]
                WriteLog('Found Handle ' & $sHandleID & ' for PID ' & $iPID & '. Closing...')
                RunWait('"' & $sHandleExe & '" -p ' & $iPID & ' -c ' & $sHandleID & ' -y', "", @SW_HIDE)
            EndIf
        EndIf
    Next
EndFunc

Func UpdateBnetConfig($sEmail)
    If Not FileExists($g_sBnetConfigPath) Then
        MsgBox($MB_ICONERROR, "Error", "Battle.net.config not found at: " & $g_sBnetConfigPath)
        Return
    EndIf

    Local $sContent = FileRead($g_sBnetConfigPath)
    ; Replace "SavedAccountNames": "..." with "SavedAccountNames": "email"
    ; PS: "SavedAccountNames`": `".+@.+`","
    
    ; AutoIt Regex Replace
    Local $sNewContent = StringRegExpReplace($sContent, '"SavedAccountNames":\s*".+?",', '"SavedAccountNames": "' & $sEmail & '",')
    
    If @extended > 0 Then
        Local $hFile = FileOpen($g_sBnetConfigPath, $FO_OVERWRITE)
        FileWrite($hFile, $sNewContent)
        FileClose($hFile)
        WriteLog('Updated Battle.net.config SavedAccountNames -> ' & $sEmail)
    Else
        WriteLog('Warning: SavedAccountNames pattern not found in Battle.net.config')
    EndIf
EndFunc

; Helper to get PID from Window Handle
Func GetWindowThreadProcessId($hWnd, ByRef $iPID)
    Local $aRet = DllCall("user32.dll", "int", "GetWindowThreadProcessId", "hwnd", $hWnd, "int*", 0)
    If IsArray($aRet) Then
        $iPID = $aRet[2]
        Return $aRet[0]
    EndIf
    Return 0
EndFunc
