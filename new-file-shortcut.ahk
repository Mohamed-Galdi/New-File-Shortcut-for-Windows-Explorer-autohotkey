#Requires AutoHotkey v2.0
TraySetIcon("shell32.dll", 3)
TrayTip("CreateFileMacro", "Ctrl+N to create a new file in the current Explorer window.")

#HotIf WinActive("ahk_class CabinetWClass") || WinActive("ahk_class Progman") || WinActive("ahk_class WorkerW")
^n::CreateNewFile()
#HotIf

CreateNewFile() {
   ; Get the current active Explorer window's path
   folderPath := GetActiveExplorerPath()
   if !folderPath {
       MsgBox("Failed to retrieve the current folder path.", "Error", "IconX")
       return
   }

   baseName := "New File"
   index := 1
   Loop {
       fileName := baseName (index > 1 ? Format(" ({})", index) : "")
       filePath := folderPath "\" fileName
       if !FileExist(filePath)
           break
       index++
   }

   try {
       FileAppend("", filePath)
   } catch as Err {
       MsgBox("Failed to create the file.`nError: " Err.Message, "Error", "IconX")
       return
   }

   if WinActive("ahk_class CabinetWClass") || WinActive("ahk_class Progman") || WinActive("ahk_class WorkerW") {
       Send("{F5}")       ; Refresh the view
       Sleep(300)         ; Wait for refresh
       Send("^{Home}")    ; Go to top of list
       Sleep(100)         ; Brief pause
       Send(fileName)     ; Type filename to select
       Sleep(100)         ; Brief pause
       Send("{F2}")       ; Start renaming
   }
}

GetActiveExplorerPath() {
   if WinActive("ahk_class Progman") || WinActive("ahk_class WorkerW") {
       return A_Desktop
   }
   
   shell := ComObject("Shell.Application")
   for window in shell.Windows {
       try {
           ; Check if the window is active and get its path
           if (window.HWND = WinActive("A")) {
               return window.Document.Folder.Self.Path
           }
       }
   }
   return ""
}