<div align="center">

## \*\*\*Convert TXT file to Executable EXE\*\*\*


</div>

### Description

This code convert a TXT file to EXE file.When you convert the file start the EXE and the old file will be typed(like TYPE command)

This is really great code

NOTE : RUN THE .EXE FROM MS-DOS MODE
 
### More Info
 
Create a label, a command button and common dialog control

Change the Caption of the button to "Select a file"

And that's all


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Atanas Matev](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/atanas-matev.md)
**Level**          |Unknown
**User Rating**    |4.6 (73 globes from 16 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/atanas-matev-convert-txt-file-to-executable-exe__1-2071/archive/master.zip)





### Source Code

```
Dim a(14) As Byte
Dim i As Integer
Public Function HiByte(ByVal wParam As Integer)
 HiByte = wParam \ &H100 And &HFF&
End Function
Public Function LoByte(ByVal wParam As Integer)
 LoByte = wParam And &HFF&
End Function
Private Sub Command1_Click()
 On Error GoTo 10
 a(0) = 190
 a(1) = 15
 a(2) = 1
 a(3) = 185
 a(4) = 0
 a(5) = 0
 a(6) = 252
 a(7) = 172
 a(8) = 205
 a(9) = 41
 a(10) = 73
 a(11) = 117
 a(12) = 250
 a(13) = 205
 a(14) = 32
 CommonDialog1.Filter = "Text Files|*.txt|"
 CommonDialog1.Action = 1
 Open CommonDialog1.filename For Input As #1
 sourcelen = LOF(1)
 Close #1
 a(4) = LoByte(sourcelen)
 a(5) = HiByte(sourcelen)
 newfilename = Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & ".exe"
 If MsgBox("Are you sure you want to convert `" & CommonDialog1.FileTitle & "` to `" & newfilename & "`", vbYesNo, "Confirm") = vbNo Then Exit Sub
 Open CommonDialog1.filename For Input As #1
 Open newfilename For Output As #2
 t = Input(LOF(1), 1)
 For k = 0 To 14
 st = st & Chr(a(k))
 Next k
 st = st & t
 Print #2, st
 Close #1
 Close #2
 Label1.Caption = "Converted successful"
 Exit Sub
10
 Label1.Caption = "Error"
End Sub
```

