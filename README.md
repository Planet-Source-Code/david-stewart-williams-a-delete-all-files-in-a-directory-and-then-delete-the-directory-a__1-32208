<div align="center">

## A Delete All files in a Directory and then Delete the DIrectory A


</div>

### Description

Code Deletes a File and Directory, Note YOu DOnt Have to Reference Ms Scripting Runtime Please Vote for me if this helps you out or if you think it is a good code!!!!! Thank you Sean For letting me know you don't Have to reference Ms Scripting
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[David Stewart Williams](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/david-stewart-williams.md)
**Level**          |Beginner
**User Rating**    |4.4 (101 globes from 23 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/david-stewart-williams-a-delete-all-files-in-a-directory-and-then-delete-the-directory-a__1-32208/archive/master.zip)





### Source Code

```
'you can't just delete an exe with this by itself but it can delete an exe in a directory
'Please vote if you like this code!!
Public Sub DelAll(ByVal DirtoDelete As Variant)
Dim FSO, FS
Set FSO = CreateObject("Scripting.FileSystemObject")
FS = FSO.DeleteFolder(DirtoDelete, True)
End Sub
'so like
Private Sub Command1_Click()
Call delall("c:\")
'that would delete the c:\ drive
End Sub
'if you need more help e-mail me at
'exhibitworks@yahoo.com
```

