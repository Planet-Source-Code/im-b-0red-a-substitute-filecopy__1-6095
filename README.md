<div align="center">

## A substitute 'FileCopy'


</div>

### Description

i'd imagine this has been done before, but if it has i haven't seen it. all it is is a substitute for the FileCopy statement. copies a file byte-for-byte to a new destination. 110% commented just like 'A+ Secure Delete' And 'A "Dummy" File Generator' (both by me). well hope you like this.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Im\_\[B\]0ReD](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/im-b-0red.md)
**Level**          |Intermediate
**User Rating**    |3.7 (26 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/im-b-0red-a-substitute-filecopy__1-6095/archive/master.zip)





### Source Code

```
Function CopyFile(srcFile As String, dstFile As String)
'this copies a file byte-for-byte
'or you could just use good old FileCopy :-)
On Error Resume Next 'If we get an error, keep going
Dim Copy As Long, CopyByteForByte As Byte 'the variables
Open srcFile For Binary Access Write As #1 'open the destination file so we can write to it
Open dstFile For Binary Access Read As #2 'open the source file so we can read from it
For Copy = 1 To LOF(2) 'Copy The SourceFile Byte-For-Byte
  Put #1, , CopyByteForByte 'Put the byte in the destination file
Next Copy 'stop the loop
MsgBox "Done!"
End Function
```

