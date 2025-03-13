; FOLDERPEEK v3.0.0 - extend File Explorer with a tooltip with the contents of folders and
; 7z archives (if 7z is installed), also previews the first 1500 characters of txt files.
; - Made by DavidBevi https://github.com/DavidBevi/folderpeek
; - Help by Plankoe https://www.reddit.com/r/AutoHotkey/comments/1igtojs/comment/masgznv/

;▼ SETTINGS
#Requires AutoHotkey v2.0
#SingleInstance Force
CoordMode("Mouse")
Try TraySetIcon("Shell32.dll", 309)
FileExist(pathof_7z:=A_ProgramFiles "\7-Zip\7z.exe")? {} :
FileExist(pathof_7z:=A_Programs     "\7-Zip\7z.exe")? {} : 0
A_IconTip:="𝐅𝐎𝐋𝐃𝐄𝐑𝐏𝐄𝐄𝐊 for File Explorer`n  Preview contents of hovered`n  folders" .
            (pathof_7z? ", 7z," : "") . " and txt files."

;▼ (DOUBLE-CLICK) RELOAD THIS SCRIPT
~F2::(A_ThisHotkey=A_PriorHotkey and A_TimeSincePriorHotkey<200)? Reload(): {}

;▼ by DavidBevi
SetTimer(FolderPeek, 16)
FolderPeek(*) {
    Static oldx:=0, oldy:=0, contentof:=Map(), cache:=["",""] ;[path,contents]
    Static dif:= [Ord("𝟎")-Ord("0"), Ord("𝐚")-Ord("a"), Ord("𝐀")-Ord("A"), 0]
    MouseGetPos(&x,&y)
    If x=oldx && y=oldy {
        Return
    } Else oldx:=x, oldy:=y
    path:=ExplorerGetHoveredItem()
    ;▼ folder
    If cache[1]!=path && DirExist(path) {
        dirs:="", contentof[path]:=""
        Loop Parse StrSplit(path,"\")[-1]  ;► 𝐟𝐚𝐧𝐜𝐲 𝐛𝐨𝐥𝐝 𝐟𝐨𝐥𝐝𝐞𝐫-𝐧𝐚𝐦𝐞
            dirs.=Chr(Ord(A_LoopField)+dif[A_LoopField~="[0-9]"?1: 
                A_LoopField~="[a-z]"?2: A_LoopField~="[A-Z]"?3: 4])
        Loop Files, path "\*.*", "DF"
            If A_Index<46
                DirExist(path "\" A_LoopFileName)? dirs.="`n🖿 " A_LoopFileName:
                    contentof[path].="`n     " A_LoopFileName
            Else overflow:="`n`n    (+ " A_Index-45 ")"
        cache:=[path, dirs contentof[path] (IsSet(overflow)? overflow: "")]
    ;▼ 7z
    } Else If path~="\.7z$" && IsSet(pathof_7z) {
        If !(contentof.Has(path)) {
            guipeek("LOADING 7z...","x" x+16 " y" y+20)
            files:="", contentof[path]:=""
            Try contents:=ComObject("WScript.Shell").Exec('"' pathof_7z '" l -ba "' path '"').StdOut.ReadAll()
            Loop Parse StrSplit(path,"\")[-1]  ;► 𝐟𝐚𝐧𝐜𝐲 𝐛𝐨𝐥𝐝 𝐚𝐫𝐜𝐡𝐢𝐯𝐞-𝐧𝐚𝐦𝐞
                contentof[path].=Chr(Ord(A_LoopField)+dif[A_LoopField~="[0-9]"?1: 
                    A_LoopField~="[a-z]"?2: A_LoopField~="[A-Z]"?3: 4])
            Loop Parse contents, "`n"
                StrLen(A_LoopField)<2? {}: files.="`n • " SubStr(A_LoopField,54)
            contentof[path].= Sort(files)
        }
        cache:=[path,contentof[path]]
    ;▼ txt
    } Else If FileExist(path) && path~="\.txt$" {
        (!contentof.Has(path) && FileRead(path)~="^\QFront-end Connection")? ;----- DEV'S CUSTOM FILTER --- ;
            contentof[path]:=SubStr(FileRead(path),start:=RegExMatch(FileRead(path),"F IO C.*UNIQUEN.*hv"), ;
            RegExMatch(FileRead(path),"\| *\R*F IO C.*backpla")-start+1) :{} ;---- can be deleted safely -- ;
        contentof.Has(path)? {}: contentof[path]:=SubStr(FileRead(path),1,1500)
        cache:=[path,contentof[path]]
    } Else If !DirExist(path) {
        cache:=["",""]
    }
    ;▼ router
    DirExist(cache[1]) or path~="\.7z$" ? ToolTip(cache[2]) : ToolTip("")
    cache[1]~="\.txt$" ? guipeek(cache[2],"x" x+16 " y" y+20) : guipeek()
    ;▼ function
    guipeek(text:=0, opts:="") {
        Static g, gt
        guiexists:=0
        Try guiexists:=g.Hwnd
        If !guiexists && text ;► create gui, populate, show
            a:=WinActive("A"), g:=Gui("-Caption +ToolWindow +AlwaysOnTop -DPIScale"),
            g.SetFont(,"Consolas"), g.SetFont(,"Monospace"), g.SetFont(,"Monoid"),
            gt:=g.AddText(,text), g.Show(opts), WinSetTransparent(240,g.hwnd),
            WinActive("A")=a?{}:WinActivate(a)
        Else If guiexists && text && gt.Value!=text ;► destroy gui, recreate
            g.Destroy(),guipeek(text,opts)
        Else If guiexists && !text ;► destroy gui
            g.Destroy()
    }
}
;▼ by PLANKOE with edits
ExplorerGetHoveredItem() {
    static h:=DllCall('LoadLibrary','str','oleacc','ptr')
    DllCall('GetCursorPos', 'int64*', &pt:=0)
    hwnd:=DllCall('GetAncestor','ptr',DllCall('user32.dll\WindowFromPoint','int64',pt),'uint',2)
    If RegExMatch(WinGetClass(hwnd),'^(?:(?<desktop>Progman|WorkerW)|(?:Cabinet|Explore)WClass)$',&M)
        shellWindows:=ComObject('Shell.Application').Windows,
        M.Desktop? shellWindow:=shellWindows.Item(ComValue(0x13,0x8)): GetFromActiveTab()
    GetFromActiveTab() {
        Try activeTab:=ControlGetHwnd('ShellTabWindowClass1',hwnd)
        For w in shellWindows ; https://learn.microsoft.com/en-us/windows/win32/shell/shellfolderview
            (w.hwnd=hwnd && IsSet(activeTab))? ; https://www.autohotkey.com/boards/viewtopic.php?f=83&t=109907
                (ComCall(3,ComObjQuery(w,id:="{000214E2-0000-0000-C000-000000000046}",id), 'uint*',&thisTab:=0),
                thisTab=activeTab? shellWindow:=w :{}) :{}
    }
    If !IsSet(shellWindow)
        Return
    If DllCall('oleacc\AccessibleObjectFromPoint', 'int64',pt, 'ptr*',&pAcc:=0, 'ptr',buf:=Buffer(8+2*A_PtrSize))=0
        idChild:=NumGet(buf,8,'uint'), accObj:=ComValue(9,pAcc,1)
    Switch accObj.accRole[idChild] {
        Case 42: Try name:=accObj.accParent.accName[idChild]    ; editable text
        Case 34: Try name:=accObj.accName[idChild]              ; list item
    }
    Return IsSet(name)? (RTrim(shellWindow.Document.Folder.Self.Path, '\') '\' name) :""
}
