Set fso = WScript.CreateObject("Scripting.Filesystemobject")
n = 10
Do Until (n-10) = 10
	n = n+1
	num = n-10
	Set f=fso.OpenTextFile("file_sync.conf-02-"+CStr(n)+"-common",1)'以只读方式打开
	str=f.ReadAll'读取所有内容
	str=Replace(str,"PROC["+strNum(num)+"]","PROC["+CStr(n)+"]")'替换文本。
	str=Replace(str,"DATA[01]","DATA[02]")
	str=Replace(str,"sync_data/01/"+strNum(num),"sync_data/02/"+CStr(n))
	str=Replace(str,"600"+strNum(num),"600"+CStr(n))
	str=Replace(str,"192.168.0.72","10.39.4.191")
	str=Replace(str,"640"+strNum(num),"640"+CStr(n))
	str=Replace(str,"TARGET_ID="+strNum(num),"TARGET_ID="+CStr(n))
	str=Replace(str,"400"+strNum(num),"400"+CStr(n))
	f.Close'关闭文件，此句可忽略
	Set f=fso.OpenTextFile("file_sync.conf-02-"+CStr(n)+"-common",2)'以写方式打开
	f.Write(str)'将字符串写入
	f.Close
Loop

Function strNum(numb)
	strNum = CStr(numb)
	if numb < 10 Then strNum = "0"+CStr(numb)
End Function
