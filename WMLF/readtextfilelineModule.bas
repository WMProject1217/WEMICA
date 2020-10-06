Attribute VB_Name = "readtextfilelineModule"
Sub Main()
Dim TextLine
cmd = Command
Open cmd For Input As #1 ' 打开文件。
Do While Not EOF(1) ' 循环至文件尾。
Line Input #1, TextLine ' 读入一行数据并将其赋予某变量。
'Debug.Print TextLine ' 在立即窗口中显示数据。
Path = App.Path & "\wmca.exe " & TextLine
retval = Shell(Path, vbNormalFocus)
Loop
Close #1 ' 关闭文件。
End Sub
