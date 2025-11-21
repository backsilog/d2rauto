#Requires AutoHotkey v2.0
#SingleInstance Force

if not A_IsAdmin {
    try {
        Run '*RunAs "' A_ScriptFullPath '"'
    } catch {
        MsgBox("This script requires admin privileges to manage D2R handles.", "Error", "Iconx")
    }
    ExitApp
}

; Global Constants
global INI_FILE := A_ScriptDir "\D2R_Launcher.ini"
global TITLE := "D2R Multi-Launcher"
global LOG_FILE := A_ScriptDir "\D2R_launcher.log"

; Global Variables
global g_MainGui := ""
global g_ListView := ""
global g_Handle64Path := ""
global g_D2RFolder := ""
global g_BnetConfigPath := ""

; Load Settings
LoadSettings()

; Init logging
InitLogging()

; Create GUI
CreateMainGUI()

; Functions

CreateMainGUI() {
    global g_MainGui, g_ListView

    g_MainGui := Gui(, TITLE)
    g_MainGui.OnEvent("Close", (*) => ExitApp())

    g_ListView := g_MainGui.Add("ListView", "w580 h300 Grid", ["Profile Name", "Account Name", "Server", "Mod"])
    g_ListView.ModifyCol(1, 150)
    g_ListView.ModifyCol(2, 200)
    g_ListView.ModifyCol(3, 100)
    g_ListView.ModifyCol(4, 100)

    btnLaunch := g_MainGui.Add("Button", "x10 y320 w100 h30", "Launch")
    btnLaunch.OnEvent("Click", LaunchSelectedAccount)

    btnAdd := g_MainGui.Add("Button", "x120 y320 w100 h30", "Add Account")
    btnAdd.OnEvent("Click", AddAccount)

    btnEdit := g_MainGui.Add("Button", "x230 y320 w100 h30", "Edit Account")
    btnEdit.OnEvent("Click", EditAccount)

    btnDelete := g_MainGui.Add("Button", "x340 y320 w100 h30", "Delete Account")
    btnDelete.OnEvent("Click", DeleteAccount)

    btnSettings := g_MainGui.Add("Button", "x490 y320 w100 h30", "Settings")
    btnSettings.OnEvent("Click", ShowSettings)

    RefreshAccountList()

    g_MainGui.Show("w600 h400")
}

InitLogging() {
    try {
        FileAppend(FormatTime(, "MM/dd/yyyy") " - Launcher started`n", LOG_FILE)
    }
}

WriteLog(sMsg) {
    try {
        FileAppend(FormatTime(, "HH:mm:ss") " - " sMsg "`n", LOG_FILE)
    }
}

LoadSettings() {
    global g_Handle64Path, g_D2RFolder, g_BnetConfigPath

    g_Handle64Path := IniRead(INI_FILE, "Settings", "Handle64Path", "C:\Program Files (x86)\Diablo II Resurrected")
    g_D2RFolder := IniRead(INI_FILE, "Settings", "D2RFolder", "C:\Program Files (x86)\Diablo II Resurrected")
    g_BnetConfigPath := IniRead(INI_FILE, "Settings", "BnetConfigPath", "C:\Users\" A_UserName "\AppData\Roaming\Battle.net\Battle.net.config"
    )

    ; If the script is placed in the D2R installation folder, prefer that folder
    if FileExist(A_ScriptDir "\D2R.exe") {
        g_D2RFolder := A_ScriptDir
        WriteLog("Auto-detected D2RFolder = " g_D2RFolder)
    }
}

SaveSettings() {
    IniWrite(g_Handle64Path, INI_FILE, "Settings", "Handle64Path")
    IniWrite(g_D2RFolder, INI_FILE, "Settings", "D2RFolder")
    IniWrite(g_BnetConfigPath, INI_FILE, "Settings", "BnetConfigPath")
}

RefreshAccountList() {
    g_ListView.Delete()

    try {
        fileContent := FileRead(INI_FILE)
        loop parse, fileContent, "`n", "`r" {
            if RegExMatch(A_LoopField, "^\[(.*)\]$", &match) {
                sectionName := match[1]
                if (sectionName != "Settings") {
                    sAccount := IniRead(INI_FILE, sectionName, "Account", "")
                    sServer := IniRead(INI_FILE, sectionName, "Server", "")
                    sMod := IniRead(INI_FILE, sectionName, "Mod", "")
                    g_ListView.Add(, sectionName, sAccount, sServer, sMod)
                }
            }
        }
    } catch as err {
        ; File might not exist yet
    }
}

ShowSettings(*) {
    settingsGui := Gui(, "Global Settings")
    settingsGui.Opt("+Owner" g_MainGui.Hwnd)

    settingsGui.Add("Text", "x10 y20 w100 h20", "Handle64 Path:")
    inputHandle := settingsGui.Add("Edit", "x120 y18 w300 h20", g_Handle64Path)
    btnBrowseHandle := settingsGui.Add("Button", "x430 y18 w30 h20", "...")
    btnBrowseHandle.OnEvent("Click", (*) => BrowseFile(inputHandle, "Select Handle64 Folder", "Executables (*.exe)"))

    settingsGui.Add("Text", "x10 y50 w100 h20", "D2R Folder:")
    inputD2R := settingsGui.Add("Edit", "x120 y48 w300 h20", g_D2RFolder)
    btnBrowseD2R := settingsGui.Add("Button", "x430 y48 w30 h20", "...")
    btnBrowseD2R.OnEvent("Click", (*) => BrowseFolder(inputD2R, "Select D2R Folder"))

    settingsGui.Add("Text", "x10 y80 w100 h20", "Bnet Config:")
    inputConfig := settingsGui.Add("Edit", "x120 y78 w300 h20", g_BnetConfigPath)
    btnBrowseConfig := settingsGui.Add("Button", "x430 y78 w30 h20", "...")
    btnBrowseConfig.OnEvent("Click", (*) => BrowseFile(inputConfig, "Select Battle.net.config", "Config (*.config)"))

    btnSave := settingsGui.Add("Button", "x150 y130 w80 h30", "Save")
    btnSave.OnEvent("Click", SaveGlobalSettings)

    btnCancel := settingsGui.Add("Button", "x250 y130 w80 h30", "Cancel")
    btnCancel.OnEvent("Click", (*) => settingsGui.Destroy())

    settingsGui.Show("w500 h200")

    BrowseFile(editCtrl, title, filter) {
        selected := FileSelect(3, editCtrl.Value, title, filter)
        if selected {
            SplitPath(selected, , &dir)
            editCtrl.Value := dir
        }
    }

    BrowseFolder(editCtrl, title) {
        selected := DirSelect(editCtrl.Value, 3, title)
        if selected
            editCtrl.Value := selected
    }

    SaveGlobalSettings(*) {
        global g_Handle64Path, g_D2RFolder, g_BnetConfigPath
        g_Handle64Path := inputHandle.Value
        g_D2RFolder := inputD2R.Value
        g_BnetConfigPath := inputConfig.Value
        SaveSettings()
        settingsGui.Destroy()
    }
}

AddAccount(*) {
    ShowAccountDialog("", "", "", "", "")
}

EditAccount(*) {
    row := g_ListView.GetNext(0, "Focused")
    if (row == 0) {
        MsgBox("Please select an account to edit.", "Warning", "Icon!")
        return
    }

    sProfile := g_ListView.GetText(row, 1)
    sAccount := IniRead(INI_FILE, sProfile, "Account", "")
    sPassword := IniRead(INI_FILE, sProfile, "Password", "")
    sServer := IniRead(INI_FILE, sProfile, "Server", "")
    sMod := IniRead(INI_FILE, sProfile, "Mod", "")

    ShowAccountDialog(sProfile, sAccount, sPassword, sServer, sMod, true)
}

DeleteAccount(*) {
    row := g_ListView.GetNext(0, "Focused")
    if (row == 0) {
        MsgBox("Please select an account to delete.", "Warning", "Icon!")
        return
    }

    sProfile := g_ListView.GetText(row, 1)

    result := MsgBox("Are you sure you want to delete profile '" sProfile "'?", "Confirm", "YesNo Icon?")
    if (result == "Yes") {
        IniDelete(INI_FILE, sProfile)
        RefreshAccountList()
    }
}

ShowAccountDialog(sProfile, sAccount, sPassword, sServer, sMod, bEdit := false) {
    accGui := Gui(, bEdit ? "Edit Account" : "Add Account")
    accGui.Opt("+Owner" g_MainGui.Hwnd)

    accGui.Add("Text", "x10 y20 w100 h20", "Profile Name:")
    inputProfile := accGui.Add("Edit", "x120 y18 w250 h20", sProfile)
    if bEdit
        inputProfile.Opt("+ReadOnly")

    accGui.Add("Text", "x10 y50 w100 h20", "Account (Email):")
    inputAccount := accGui.Add("Edit", "x120 y48 w250 h20", sAccount)

    accGui.Add("Text", "x10 y80 w100 h20", "Password:")
    inputPassword := accGui.Add("Edit", "x120 y78 w250 h20 Password", sPassword)

    accGui.Add("Text", "x10 y110 w100 h20", "Server (Region):")
    comboServer := accGui.Add("ComboBox", "x120 y108 w250 h20 Choose1", ["asia", "us", "eu"])
    if (sServer != "") {
        try comboServer.Text := sServer
    } else {
        comboServer.Text := "asia"
    }

    accGui.Add("Text", "x10 y140 w150 h20", "Mod / Args (Optional):")
    inputMod := accGui.Add("Edit", "x120 y138 w250 h20", sMod)

    btnSave := accGui.Add("Button", "x100 y190 w80 h30", "Save")
    btnSave.OnEvent("Click", SaveAccount)

    btnCancel := accGui.Add("Button", "x220 y190 w80 h30", "Cancel")
    btnCancel.OnEvent("Click", (*) => accGui.Destroy())

    accGui.Show("w400 h250")

    SaveAccount(*) {
        newProfile := inputProfile.Value
        if (newProfile == "") {
            MsgBox("Profile Name cannot be empty.", "Error", "Iconx")
            return
        }

        IniWrite(inputAccount.Value, INI_FILE, newProfile, "Account")
        IniWrite(inputPassword.Value, INI_FILE, newProfile, "Password")
        IniWrite(comboServer.Text, INI_FILE, newProfile, "Server")
        IniWrite(inputMod.Value, INI_FILE, newProfile, "Mod")

        RefreshAccountList()
        accGui.Destroy()
    }
}

LaunchSelectedAccount(*) {
    row := g_ListView.GetNext(0, "Focused")
    if (row == 0) {
        MsgBox("Please select an account to launch.", "Warning", "Icon!")
        return
    }

    sProfile := g_ListView.GetText(row, 1)
    sAccount := IniRead(INI_FILE, sProfile, "Account", "")
    sPassword := IniRead(INI_FILE, sProfile, "Password", "")
    sServer := IniRead(INI_FILE, sProfile, "Server", "asia")
    sMod := IniRead(INI_FILE, sProfile, "Mod", "")

    ; 1. Map Server Region
    sRegionPrefix := sServer
    if (sServer == "asia")
        sRegionPrefix := "kr"

    ; 2. Update Battle.net Config
    UpdateBnetConfig(sAccount)

    ; 3. Launch D2R
    sD2RExe := g_D2RFolder "\D2R.exe"
    if !FileExist(sD2RExe) {
        MsgBox("D2R.exe not found at: " sD2RExe, "Error", "Iconx")
        return
    }

    ; Build parameter string
    sParams := Format('-username "{1}" -password "{2}" -address "{3}.actual.battle.net"', sAccount, sPassword,
        sRegionPrefix)

    if (sMod == "")
        sParams .= " -ns -w -txt"
    else
        sParams .= " " sMod

    ; Launch
    try {
        Run('"' sD2RExe '" ' sParams, g_D2RFolder, , &iPID)
    } catch as err {
        MsgBox("Failed to launch D2R.exe`n" err.Message, "Error", "Iconx")
        return
    }

    WriteLog('Launched: "' sD2RExe '" ' sParams '  (PID=' iPID ')')

    ; 4. Post-Launch Handle Kill
    WriteLog('Waiting 5 seconds for handle creation...')
    Sleep(5000)
    KillD2RHandleForPID(iPID)

    ; 5. Rename Window
    endTime := A_TickCount + 30000 ; 30 seconds
    hWindow := 0

    while (A_TickCount < endTime) {
        try {
            if WinExist("ahk_pid " iPID " ahk_class Diablo II Class") {
                hWindow := WinExist("Diablo II: Resurrected ahk_pid " iPID)
                if hWindow
                    break
            }
        }
        Sleep(500)
    }

    if hWindow {
        ; Try to rename multiple times as the game might reset it
        loop 5 {
            WinSetTitle(sProfile, hWindow)
            if (WinGetTitle(hWindow) == sProfile)
                break
            Sleep(500)
        }

        ; Skip Cinematics
        Sleep(1000)
        loop 5 {
            try {
                WinActivate(hWindow)
                if WinWaitActive(hWindow, , 1) {
                    Send("{Space}")
                    WriteLog("Sent Space to " sProfile)
                }
            }
            Sleep(500)
        }
    }
}

KillD2RHandleForPID(iPID) {
    sHandleExe := g_Handle64Path "\handle64.exe"
    if !FileExist(sHandleExe) {
        sHandleExe := g_Handle64Path
        if !InStr(sHandleExe, "handle64.exe")
            sHandleExe .= "\handle64.exe"

        if !FileExist(sHandleExe) {
            WriteLog("Error: handle64.exe not found at " sHandleExe)
            return
        }
    }

    WriteLog('Running handle64 for PID ' iPID)

    ; Run handle64 to find handles for specific PID
    ; Use temp file to avoid console window flashing
    tempFile := A_Temp "\d2r_handle_" A_TickCount ".txt"
    try {
        RunWait(A_ComSpec ' /C ""' sHandleExe '" -accepteula -a -p ' iPID ' > "' tempFile '""', , "Hide")
    } catch as err {
        WriteLog("Error running handle64: " err.Message)
        return
    }

    try {
        sOutput := FileRead(tempFile)
        FileDelete(tempFile)
    } catch {
        sOutput := ""
    }

    ; Parse output for "Check For Other Instances"
    loop parse, sOutput, "`n", "`r" {
        if InStr(A_LoopField, "Check For Other Instances") {
            ; Extract Handle ID (Hex)
            if RegExMatch(A_LoopField, "([0-9A-Fa-f]+):\s+Event.*DiabloII Check For Other Instances", &match) {
                sHandleID := match[1]
                WriteLog('Found Handle ' sHandleID ' for PID ' iPID '. Closing...')
                RunWait('"' sHandleExe '" -p ' iPID ' -c ' sHandleID ' -y', , "Hide")
            }
        }
    }
}

UpdateBnetConfig(sEmail) {
    if !FileExist(g_BnetConfigPath) {
        MsgBox("Battle.net.config not found at: " g_BnetConfigPath, "Error", "Iconx")
        return
    }

    try {
        sContent := FileRead(g_BnetConfigPath)
        sNewContent := RegExReplace(sContent, '"SavedAccountNames":\s*".+?",', '"SavedAccountNames": "' sEmail '",')

        if (sContent != sNewContent) {
            FileDelete(g_BnetConfigPath)
            FileAppend(sNewContent, g_BnetConfigPath)
            WriteLog('Updated Battle.net.config SavedAccountNames -> ' sEmail)
        } else {
            WriteLog('Warning: SavedAccountNames pattern not found or already set in Battle.net.config')
        }
    } catch as err {
        WriteLog('Error updating Battle.net.config: ' err.Message)
    }
}
