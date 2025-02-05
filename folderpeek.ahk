    ; FOLDEDPEEK v2 - extend File Explorer with a tooltip that shows the files inside any hovered folder
    ; - Made by DavidBevi https://github.com/DavidBevi/folderpeek
    ; - Help by Plankoe https://www.reddit.com/r/AutoHotkey/comments/1igtojs/comment/masgznv/

    ;▼ RECOMMENDED SETTINGS
    #Requires AutoHotkey v2.0
    #SingleInstance Force

    ;▼ (DOUBLE-CLICK) RELOAD THIS SCRIPT
    ~F2::(A_ThisHotkey=A_PriorHotkey and A_TimeSincePriorHotkey<200)? Reload(): {}

    SetTimer(FolderPeek, 16)

    ; by DavidBevi
    FolderPeek(*) {
        Static mouse:=[0,0]
        MouseGetPos(&x,&y)
        If mouse[1]=x and mouse[2]=y {
            Return
        } Else mouse:=[x,y]
        Static cache:=["",""] ;[path,contents]
        Static dif:= [Ord("𝟎")-Ord("0"), Ord("𝐚")-Ord("a"), Ord("𝐀")-Ord("A")]
        path:=""
        Try path:=ExplorerGetHoveredItem()
        If (cache[1]!=path && FileExist(path)~="D") {
            cache[1]:=path, dirs:="", files:=""
            for letter in StrSplit(StrSplit(path,"\")[-1])        ; boring foldername → 𝐟𝐚𝐧𝐜𝐲 𝐟𝐨𝐥𝐝𝐞𝐫𝐧𝐚𝐦𝐞
                dirs.=  letter~="[0-9]" ? Chr(Ord(letter)+dif[1]) :
                        letter~="[a-z]" ? Chr(Ord(letter)+dif[2]) :
                        letter~="[A-Z]" ? Chr(Ord(letter)+dif[3]) : letter
            Loop Files, path "\*.*", "DF"
                f:=A_LoopFileName, (FileExist(path "\" f)~="D")?  dirs.="`n🖿 " f:  files.="`n     " f
            cache[2]:= dirs . files
        } Else If !(FileExist(path)~="D") {
            cache:=["",""]
        }
        ToolTip(cache[2])
    }

    ; by PLANKOE with edits
    ExplorerGetHoveredItem() {
        static VT_DISPATCH:=9, F_OWNVALUE:=1, h:=DllCall('LoadLibrary','str','oleacc','ptr')
        DllCall('GetCursorPos', 'int64*', &pt:=0)
        hwnd := DllCall('GetAncestor','ptr',DllCall('user32.dll\WindowFromPoint','int64',pt),'uint',2)
        winClass:=WinGetClass(hwnd)
        if RegExMatch(winClass,'^(?:(?<desktop>Progman|WorkerW)|(?:Cabinet|Explore)WClass)$',&M) {
            shellWindows:=ComObject('Shell.Application').Windows
            if M.Desktop ; https://www.autohotkey.com/boards/viewtopic.php?p=255169#p255169
                shellWindow:= shellWindows.Item(ComValue(0x13, 0x8))
            else {
                try activeTab:=ControlGetHwnd('ShellTabWindowClass1',hwnd)
                for w in shellWindows { ; https://learn.microsoft.com/en-us/windows/win32/shell/shellfolderview
                    if w.hwnd!=hwnd
                        continue
                    if IsSet(activeTab) { ; https://www.autohotkey.com/boards/viewtopic.php?f=83&t=109907
                        static IID_IShellBrowser := '{000214E2-0000-0000-C000-000000000046}'
                        shellBrowser := ComObjQuery(w,IID_IShellBrowser,IID_IShellBrowser)
                        ComCall(3,shellBrowser, 'uint*',&thisTab:=0)
                        if thisTab!=activeTab
                            continue
                    }
                    shellWindow:= w
                }
            }
        }
        if !IsSet(shellWindow)
            return
        varChild := Buffer(8 + 2*A_PtrSize)
        if DllCall('oleacc\AccessibleObjectFromPoint', 'int64',pt, 'ptr*',&pAcc:=0, 'ptr',varChild)=0
            idChild:=NumGet(varChild,8,'uint'), accObj:=ComValue(VT_DISPATCH,pAcc,F_OWNVALUE)
        if !IsSet(accObj)
            return
        if accObj.accRole[idChild] = 42  ; editable text
            return RTrim(shellWindow.Document.Folder.Self.Path, '\') '\' accObj.accParent.accName[idChild]
        else return
    }
