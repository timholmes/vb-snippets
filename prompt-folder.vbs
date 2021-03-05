
' *** Simple function to allow a end user to pick a folder on the file system ***
Function BrowseForFolder()
  Dim oFolder
  Set oFolder = CreateObject("Shell.Application").BrowseForFolder(0,"Select a Folder",0,0)
  If (oFolder Is Nothing) Then
    BrowseForFolder = Empty
  Else 
    BrowseForFolder = oFolder.Self.Path
  End If
End Function

' Call function to prompt for a folder
Dim sFolder
sFolder = BrowseForFolder()

' Loop through all the files and do something
Dim oFSO
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFolder = oFSO.GetFolder(sFolder)
Set oFiles = oFolder.Files
For Each oFile in oFiles
    
    ' Right now this next line will print out the files it finds in the folder
    ' ****  Add your script here...
    Wscript.Echo sFolder & oFile.Name
Next
