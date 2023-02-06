On Error Resume Next
'Set objBrowser = Createobject("InternetExplorer.Application")
'objBrowser.Visible = True

'Source excel declare
StartRow = 0
EndRow = 0
ExcelName = "C:\AD\Active Data.xlsx"
SheetName = "BROWSER"

DataTable.AddSheet SheetName
DataTable.ImportSheet ExcelName, SheetName, SheetName

If DataTable.GetSheet(SheetName).GetRowCount > 0 Then
	For i = StartRow To DataTable.GetSheet(SheetName).GetRowCount -1
		DataTable.SetCurrentRow(i)
	
		BROWSER_TYPE = DataTable.Value("BROWSER_TYPE",SheetName)
		BROWSER_URL = DataTable.Value("URL", SheetName)
		ACTIVE = DataTable.Value("ACTIVE", SheetName)
		
If Active = "Y" Then
	If BROWSER_TYPE = "IE" Then
		SystemUtil.Run "iexplorer.exe", BROWSER_URL
	ElseIf BROWSER_TYPE = "EDGE" Then
		SystemUtil.Run "msedge.exe", BROWSER_URL
	ElseIf BROWSER_TYPE = "Firefox" Then
		SystemUtil.Run "firefox.exe", BROWSER_URL
	ElseIf BROWSER_TYPE = "Chrome" Then
		SystemUtil.Run "chrome.exe", BROWSER_URL
	End If
	Browser("CreationTime:=0").Sync
	Browser("CreationTime:=0").Maximize
End If
Next
End If

'LOGIN
StartRow = 0
EndRow = 0
ExcelName = "C:\AD\Active Data.xlsx"
SheetName = "LOGIN"

DataTable.AddSheet SheetName
DataTable.ImportSheet ExcelName, SheetName, SheetName

If DataTable.GetSheet(SheetName).GetRowCount > 0 Then
	For i = StartRow To DataTable.GetSheet(SheetName).GetRowCount -1
		DataTable.SetCurrentRow(i)
		
		ACTIVE = DataTable.Value("ACTIVE",SheetName)
		USERNAME = DataTable.Value("USERNAME",SheetName)
		PASSWORD = DataTable.Value("PASSWORD",SheetName)
		
		RegisterUserFunc "Page", "CaptureScreenshot", "CaptureScreenshot"
		RegisterUserFunc "Browser", "CaptureScreenshot", "CaptureScreenshot"
		RegisterUserFunc "Frame", "CaptureScreenshot", "CaptureScreenshot"
		RegisterUserFunc "Dialog", "CaptureScreenshot", "CaptureScreenshot"
		RegisterUserFunc "swfWindow", "CaptureScreenshot", "CaptureScreenshot"
		
		RegisterUserFunc "Browser", "Synchronize", "Synchronize"
		RegisterUserFunc "Page", "Synchronize", "Synchronize"
		
		Browser("Login - OSS Berbasis Risiko").Page("Login - OSS Berbasis Risiko").Link("MASUK SEKARANG").Click @@ script infofile_;_ZIP::ssf14.xml_;_
		wait (5)
		Browser("micclass:=Browser").Page("micclass:=Page").CaptureScreenshot micPass, "Page ScreenShot"
		wait (5)
		Browser("Login - OSS Berbasis Risiko").Page("Login - OSS Berbasis Risiko").WebEdit("html id:=input-21").Set USERNAME @@ script infofile_;_ZIP::ssf15.xml_;_
		wait (5)
		Browser("micclass:=Browser").Page("micclass:=Page").CaptureScreenshot micPass, "Page ScreenShot"
		Browser("Login - OSS Berbasis Risiko").Page("Login - OSS Berbasis Risiko").WebEdit("Masukkan kata sandi").Set PASSWORD @@ script infofile_;_ZIP::ssf16.xml_;_
		wait(5)
		Browser("micclass:=Browser").Page("micclass:=Page").CaptureScreenshot micPass, "Page ScreenShot"
		Browser("Login - OSS Berbasis Risiko").Page("Login - OSS Berbasis Risiko").WebButton("Masuk").Click @@ script infofile_;_ZIP::ssf18.xml_;_
		wait(5)
		Browser("micclass:=Browser").Page("micclass:=Page").CaptureScreenshot micPass, "Page ScreenShot"
		wait(5)
		Browser("Login - OSS Berbasis Risiko").Page("OSS Berbasis Risiko").WebElement("F").Click
		wait(5)
		Browser("micclass:=Browser").Page("micclass:=Page").CaptureScreenshot micPass, "Page ScreenShot"
		'Browser("Login - OSS Berbasis Risiko").Page("OSS Berbasis Risiko").WebButton("Keluar").Click
		
		next
		End If
		
		CreateWordFile
		
		'Browser("Login - OSS Berbasis Risiko").Page("Login - OSS Berbasis Risiko_2").WebEdit("Contoh: 081xxxxxxxxx atau nama").Set
		
		
