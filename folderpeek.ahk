; FOLDEDPEEK v1 - extend File Explorer with a tooltip that shows the files inside any hovered folder
; - by DavidBevi https://github.com/DavidBevi/folderpeek
; - with code provided by Plankoe https://www.reddit.com/r/AutoHotkey/comments/1igtojs/comment/masgznv/

;▼ RECOMMENDED SETTINGS
#Requires AutoHotkey v2.0
#SingleInstance Force

FolderPeek() {
    If !WinActive("ahk_class CabinetWClass") {
        ToolTip()
        Return
    }
    Static cache := ""
    path := ""
    Try path := ExplorerGetHoveredItem()
    If (cache!=path and InStr(FileExist(path), "D")) {
        cache := path
        folders := "【 " StrSplit(path,"\")[-1] " 】`n"
        files := ""
        Loop Files, path "\*.*", "DF" {
            (FileExist(A_LoopFileFullPath) ~= "D")?
                folders .= "🖿 " A_LoopFileName "`n":
                files .= "     " A_LoopFileName "`n"
        }
        ToolTip(folders . files)
    } Else If cache!=path {
        ToolTip()
        path=""? cache:=path: {}
    }
}
SetTimer(FolderPeek, 100)

; by PLANKOE https://www.reddit.com/r/AutoHotkey/comments/1igtojs/comment/masgznv/
ExplorerGetHoveredItem() {
    static VT_DISPATCH := 9, F_OWNVALUE := 1, h := DllCall('LoadLibrary', 'str', 'oleacc', 'ptr')

    DllCall('GetCursorPos', 'int64*', &pt:=0)
    hwnd := DllCall('GetAncestor', 'ptr', DllCall('user32.dll\WindowFromPoint', 'int64',  pt), 'uint', 2)
    shellWindow := GetExplorerComObject(hwnd)
    if !IsSet(shellWindow)
        return
    varChild := Buffer(8 + 2*A_PtrSize)
    if DllCall('oleacc\AccessibleObjectFromPoint', 'int64', pt, 'ptr*', &pAcc:=0, 'ptr', varChild) = 0 {
        idChild := NumGet(varChild, 8, 'uint')
        accObj := ComValue(VT_DISPATCH, pAcc, F_OWNVALUE)
    }
    if !IsSet(accObj)
        return
    role := accObj.accRole[idChild]
    if role = 42  ; editable text
        name := accObj.accParent.accName[idChild]
    else if role = 34  ; list item
        name := accObj.accName[idChild]
    if !IsSet(name)
        return
    return RTrim(shellWindow.Document.Folder.Self.Path, '\') '\' name
}

; by PLANKOE https://www.reddit.com/r/AutoHotkey/comments/1igtojs/comment/masgznv/
GetExplorerComObject(hwnd := WinExist('A')) {
    winClass := WinGetClass(hwnd)
    if !RegExMatch(winClass, '^(?:(?<desktop>Progman|WorkerW)|(?:Cabinet|Explore)WClass)$', &M)
       return
    shellWindows := ComObject('Shell.Application').Windows
    if M.Desktop ; https://www.autohotkey.com/boards/viewtopic.php?p=255169#p255169
        return shellWindows.Item(ComValue(0x13, 0x8))
    try activeTab := ControlGetHwnd('ShellTabWindowClass1', hwnd)
    for w in shellWindows { ; https://learn.microsoft.com/en-us/windows/win32/shell/shellfolderview
        if w.hwnd != hwnd
            continue
        if IsSet(activeTab) {
            ; Get explorer active tab for Windows 11
            ; https://www.autohotkey.com/boards/viewtopic.php?f=83&t=109907
            static IID_IShellBrowser := '{000214E2-0000-0000-C000-000000000046}'
            shellBrowser := ComObjQuery(w, IID_IShellBrowser, IID_IShellBrowser)
            ComCall(3, shellBrowser, 'uint*', &thisTab:=0)
            if thisTab != activeTab
                continue
        }
        return w
    }
}
