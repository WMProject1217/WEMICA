Attribute VB_Name = "readtextfilelineModule"
Sub Main()
Dim TextLine
cmd = Command
Open cmd For Input As #1 ' ���ļ���
Do While Not EOF(1) ' ѭ�����ļ�β��
Line Input #1, TextLine ' ����һ�����ݲ����丳��ĳ������
'Debug.Print TextLine ' ��������������ʾ���ݡ�
Path = App.Path & "\wmca.exe " & TextLine
retval = Shell(Path, vbNormalFocus)
Loop
Close #1 ' �ر��ļ���
End Sub
