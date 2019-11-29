Option Strict Off
Imports System
Imports System.Text
Imports System.IO
Imports System.Collections
Imports System.Windows.Forms
Imports System.Windows.Forms.MessageBox
Imports NXOpen
Imports NXOpen.UF

Module NXJournal

	Sub Main(ByVal args() As String)
		Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
		
	
		' Открываем консоль
		Dim lw As ListingWindow = theSession.ListingWindow
		lw.Open()

		' Ищем и читаем лог файлы в массив
		Dim LogSummName(-1) As String
		Dim LogSummText(-1) As String


		Dim fi As New System.IO.FileInfo("C:\testroot\testsub\test.txt")
		lw.Writeline(fi.Directory.Name)
		
		Dim i As integer = 0
		For Each foundFile As String In My.Computer.FileSystem.GetFiles("K:\Industrial\КД\Ящики злектрические",FileIO.SearchOption.SearchAllSubDirectories, "log.txt")
			Redim LogSummName(LogSummName.Length+1)
			Redim LogSummText(LogSummText.Length+1)
			
			'System.IO.FileInfo(foundFile)
			'fi.Directory.Name
			
			LogSummName(i) = My.Computer.FileSystem.ReadAllText(foundFile)
			LogSummText(i) = My.Computer.FileSystem.ReadAllText(foundFile)
			
			i=i+1
		Next
		
		

	End Sub

End Module