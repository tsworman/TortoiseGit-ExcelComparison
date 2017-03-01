dim objArgs, objFileSystem, sBaseVer, sNewVer, sMessage, sBaseMessage, sNewMessage, bDiffers, sTempFile

Set objArgs = WScript.Arguments
num = objArgs.Count
if num < 2 then
    MsgBox "Usage: [CScript | WScript] diff-xlsx.vbs base.xlsx new.xlsx", vbExclamation, "Invalid arguments"
    WScript.Quit 1
end if

sBaseFile = objArgs(0)
sNewFile = objArgs(1)

Set objFileSystem = CreateObject("Scripting.FileSystemObject")
If objFileSystem.FileExists(sBaseFile) = False Then
    MsgBox "File " + sBaseFile + " does not exist.  Cannot compare the files.", vbExclamation, "File not found"
    Wscript.Quit 1
End If
If objFileSystem.FileExists(sNewFile) = False Then
    MsgBox "File " + sNewFile + " does not exist.  Cannot compare the files.", vbExclamation, "File not found"
    Wscript.Quit 1
End If

Set objScript = Nothing

' Compare file size
dim fBaseFile, fNewFile, fs, f
Set fBaseFile = objFileSystem.GetFile(sBaseFile)
Set fNewFile = objFileSystem.GetFile(sNewFile)
sTempFile = "H:\temp-Git.text"

'Creat temp.txt for save path of 2 xlsx files
Set fs =Wscript.CreateObject("scripting.filesystemobject")
Set f = fs.CreateTextFile(sTempFile, 2, True)
f.WriteLine sBaseFile
f.WriteLine sNewFile
f.Close()

' Compare files using SPREADSHEETCOMPARE.Exe
Set WshShell = WScript.CreateObject("WScript.Shell")
result = WshShell.Run("""C:\Program Files (x86)\Microsoft Office\Office16\DCF\SPREADSHEETCOMPARE.exe"" H:\temp-Git.text", 0, True)

Wscript.Quit