[General]
SyntaxVersion=2
BeginHotkey=121
BeginHotkeyMod=0
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=64594d19-73b3-4e73-b008-834d326d4aef
Description=��ϵͳjson2.0
Enable=1
AutoRun=0
[Repeat]
Type=0
Number=1
[SetupUI]
Type=2
QUI=
[Relative]
SetupOCXFile=
[Comment]

[Attachment]
UEsDBBQAAgAIADZBilb8rEjUvQEAANsCAAAQABEAanNvbi10by1leGNlbC5weVVUDQAHR8QzZLHEM2RZZuFjbVLdShwxFL4fmHcIepFZWOYBCnNRbZW2UAu9LGUJm+wamU1CJmu9bKtrRRQXRFtaFUX8Qeja3rSri+7LTKbjWzTZ+XGlHYYZ8nG+n3NOaEtwqQCSTYFkRFyHZsBCxFl5EIhhFAHzCuw6rjMJ9No3PbiO++/j/gV4LJvtFmHqlVWQQF9epT+PXEdkx6AU9x8WephEdUmFopwFcJqzRWK8nr+eewkaNCRAcfB0qU5CWMk8k98DvX6Y7H6KB790dzO+2Yv713rYSU8//Pm6kt5u686J3vqY7PwovH2EcQ3lph6kTLRVzWrDKpgnoQjgMwvde1qr/1J5W/3DnRthWcaSbYOmZ8fJQTcddtOjjSKQkYrMKHLx0c/KR17BuRzord28udWO7l25zjuq5gEXhHm20r+PXwVQwiphdY4pawZwduoFrNj1NB65DjCP3V4NI4WMZcOXBGHvQbasYxMs2ezlHeOi3nL9kCMceaVMQe4N7z73wKjwbvlcr62awRdpc/qY2JsJ+514O35lwBMDzUjUMlcNN2w99kvIG/OKh/v6+xez4XzA2WgsyVe8RiyWjWVsNVVAGSZLwQwKI1L5C1BLAQIXCxQAAgAIADZBilb8rEjUvQEAANsCAAAQAAkAAAAAAAAAIACAgQAAAABqc29uLXRvLWV4Y2VsLnB5VVQFAAdHxDNkUEsFBgAAAAABAAEARwAAAPwBAAAAAA==


[Script]
MyDate = Date 'ϵͳ����
MyYear = Year(MyDate)    'ϵͳ����
MyMonth = Month(MyDate)    'ϵͳ����
If MyMonth < 10 Then 
    MyMonth = 0 & MyMonth  '�·�֮ǰ����
End If
MyDay = Day(MyDate)
If MyDay < 10 Then 
    MyDay = "0" & MyDay  '����֮ǰ����
End If
Datetime = MyMonth & Myday  'ƴ���ַ���  �õ�YYYYMMDD
TracePrint Datetime'����̨���


UserVar url="https://gd12345.sz.gov.cn/ticket/GetTicketDealCommon.json?pageAnchor=3&pageSize=100" "URL"
UserVar fathersavepath="D:\Onedrive\MyFile\��֯����\2023����ֵ���\��������\����ͳ��\���ɼ���������ͳ�ƣ�����ȷ�����ڰ�������\ԭʼ����\" "����·��"
//UserVar pages=4 "��ҳ��"
UserVar shortdelay=500 "�����ӳ�"
UserVar longdelay=3000 "��ҳ�ȴ�"

dir_Num  = 0
savepath = fathersavepath
Do
If dir_Num = 0 Then 
	savepath = fathersavepath & Datetime
Else 
	savepath = fathersavepath & Datetime&"-"&dir_Num
End If
If Lib.�ļ�.�ж��ļ����Ƿ����(savepath) Then 
	TracePrint savepath&"�Ѵ���"
	savepath = fathersavepath
	dir_Num=dir_Num+1
Else 
	Exit Do
End If
Loop

Plugin.File.CreateFolder savepath
TracePrint savepath&"�Ѵ���"


Max=0
i = 1


do
	If i = 1 Then 
		runapp "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe --new-window " &url& "&currentPage=1"
		TracePrint "����"		
	Else 
		TracePrint "ҳ����"&i	
		KeyDown 17, 1 //ctrl+t
		KeyPress 84, 1
		KeyUp 17, 1	
		Delay shortdelay
		SayString url & "&currentPage=" & i
		Delay shortdelay
		KeyPress "Enter", 1		
	End If
	Delay longdelay
	KeyPress "Tab", 1
	Delay shortdelay
	KeyPress "Tab", 1
	Delay shortdelay
	KeyDown 17, 1 //Ctrl+A
	KeyPress 65, 1
	KeyUp 17, 1
	Delay shortdelay
	
	KeyDown 17, 1 //Ctrl+C
	KeyPress 67, 1
	KeyUp 17, 1
	Delay shortdelay


	If InStr(Plugin.Sys.GetCLB ,Chr(34) & "totalCount" & Chr(34) & ": 0") > 0 Then 
		Max=i-1
		Exit Do
	Elseif InStr(Plugin.Sys.GetCLB ,Chr(34)&"code"&Chr(34)&": "&Chr(34)&"200"&Chr(34)) > 0 Then 
		TracePrint savepath & "\" & i & ".txt"
		Plugin.File.WriteFileEx savepath & "\" & i & ".txt", Plugin.Sys.GetCLB
	Else 
		MessageBox "��ҳ�쳣�������Ƿ���سɹ����ɳ��Ե��߼��"
		ExitScript
	End If
	i=i+1
Loop



Plugin.File.DeleteFile savepath &"\"& "json-to-excel.bat"
Plugin.File.WriteFileEx savepath &"\"& "json-to-excel.bat", "@echo off"
Plugin.File.WriteFileEx savepath &"\"& "json-to-excel.bat", "cd /d " & savepath
For i = 1 To Max
	Plugin.File.WriteFileEx savepath &"\"& "json-to-excel.bat", "py json-to-excel.py "&i&".txt "&i&".xlsx"
Next
Plugin.File.WriteFileEx savepath & "\" & "json-to-excel.bat", "md excel"
Plugin.File.WriteFileEx savepath & "\" & "json-to-excel.bat", "move *.xlsx .\excel"
Plugin.File.WriteFileEx savepath & "\" & "json-to-excel.bat", "@echo ��������ɣ��س��رմ���"
Plugin.File.WriteFileEx savepath & "\" & "json-to-excel.bat", "pause"
PutAttachment savepath & "\", "json-to-excel.py"
RunApp savepath &"\"& "json-to-excel.bat"

Function GetXY(target, Part)
	//TracePrint "# ����ͼƬ������ʼ��������ʾ�����"
	iX = - 2
	iY = - 2
	iZB = null
	//TracePrint "## ���ó�ʼֵx��y��ZBΪ��"
	//TracePrint iX
	//TracePrint iY
	//TracePrint iZB
	ScreenResolutionX=Plugin.GetSysInfo.GetScreenResolutionX 
	ScreenResolutionY=Plugin.GetSysInfo.GetScreenResolutionY 
	//TracePrint "## ��Ļ�ֱ���Ϊ��"
	//TracePrint ScreenResolutionX
	//TracePrint ScreenResolutionY
	//TracePrint "## ���ҵ�ͼƬĿ���������Ҫ���ǣ�"
	//TracePrint target
	//TracePrint Part
	iXY = Plugin.ColorEx.FindPic(0, 0, ScreenResolutionX, ScreenResolutionY, target, Part, 0)
	Call Plugin.ColorEx.Free()
	//TracePrint "���ҵ���ԭʼֵ���ָ�λ��x��y�ֱ��ǣ�"
	//TracePrint iXY
	Call Plugin.ColorEx.Free()
	iZB = InStr(iXY, "|")
	//TracePrint iZB
	iX = Clng(Left(iXY, iZB - 1)) : iY = Clng(Right(iXY, Len(iXY) - iZB))
	//TracePrint iX
	//TracePrint iY
	GetXY = Array(iX, iY)
	//TracePrint "# ����ͼƬ��������"
End Function

Function MGetXY(target, Part)
	XY = GetXY(Target, Part)
	TracePrint Target
	TracePrint XY(0)&","&XY(1)
	If XY(0)>-1 and XY(1)>-1 Then 
		MGetXY = XY		
		Exit Function
	End If
	For i = 2 To 9999
		NextTarget = Replace(Target, ".bmp", i & ".bmp")
		TracePrint NextTarget
		IsNextTargetFile = Plugin.File.IsFileExist(NextTarget)
		If IsNextTargetFile Then
			XY = GetXY(NextTarget, Part)
			TracePrint XY(0)&","&XY(1)
			If XY(0) > - 1  and XY(1) > - 1  Then 
			MGetXY = XY
			Exit Function
			End If
		Else 
			MGetXY = XY
			Exit Function
		End If
				
	Next
	
End Function

Function TipsSender(Message)	
	Call Plugin.Msg.Tips(Message)
End Function