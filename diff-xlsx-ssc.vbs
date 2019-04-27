
Option Explicit

Dim objArgs
Set objArgs = WScript.Arguments

Dim num
num = objArgs.Count
If num < 2 Then
    MsgBox "Usage: [CScript | WScript] diff-xlsx.vbs base.xlsx new.xlsx", vbExclamation, "Invalid arguments"
    WScript.Quit 1
End If

Dim sBaseFile, sNewFile
sBaseFile = objArgs(0)
sNewFile = objArgs(1)
Set objArgs = Nothing

Dim objFileSystem
Set objFileSystem = CreateObject("Scripting.FileSystemObject")
If objFileSystem.FileExists(sBaseFile) = False Then
    MsgBox "File " + sBaseFile + " does not exist.  Cannot compare the files.", vbExclamation, "File not found"
    WScript.Quit 1
End If
If objFileSystem.FileExists(sNewFile) = False Then
    MsgBox "File " + sNewFile + " does not exist.  Cannot compare the files.", vbExclamation, "File not found"
    WScript.Quit 1
End If

'Compare file size
Dim fBaseFile, fNewFile, sTempFolder, sTempFile
Set fBaseFile = objFileSystem.GetFile(sBaseFile)
Set fNewFile = objFileSystem.GetFile(sNewFile)
sTempFolder = objFileSystem.GetSpecialFolder(2)
sTempFile = sTempFolder + "\temp.txt"
Set objFileSystem = Nothing

'Create temp.txt for save path of 2 xlsx files
Dim fs, f
Set fs = WScript.CreateObject("Scripting.FileSystemObject")
Set f = fs.CreateTextFile(sTempFile, 2, True)
f.WriteLine sBaseFile
f.WriteLine sNewFile
f.Close()
Set fs = Nothing
Set f = Nothing

'Compare files using SPREADSHEETCOMPARE.exe
Dim WshShell, result
Set WshShell = WScript.CreateObject("WScript.Shell")
result = WshShell.Run("""C:\Program Files (x86)\Microsoft Office\Office16\DCF\SPREADSHEETCOMPARE.exe"" " & sTempFile, 0, True)
Set WshShell = Nothing

WScript.Quit