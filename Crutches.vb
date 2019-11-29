Option Strict Off
Imports System
Imports System.IO
Imports System.Collections
Imports System.Windows.Forms
Imports System.Windows.Forms.MessageBox
Imports NXOpen
Imports NXOpen.UF
Imports System.Text.RegularExpressions
 
Module NXJournal
 
Dim theSession As Session = Session.GetSession()
Dim workPart As Part = theSession.Parts.Work
Dim displayPart As Part = theSession.Parts.Display


Sub Main

	'Чтение списка атрибутов из файла
    'Dim path As String = "F:\VBNET\NxAttr.txt"		
    'Dim readText() As String = File.ReadAllLines(path)

	'Получаем имя файла без расширения и пытаемся разделить наименование от обозначения
    Dim partNo As String
	Dim rx As New Regex("\b\.\d{3}\s\b", RegexOptions.Compiled Or RegexOptions.IgnoreCase)

	'Имя файла без без расширения
    partNo = IO.Path.GetFileNameWithoutExtension(workPart.FullPath)
	
	'Поиск совпадений
	Dim matches As MatchCollection = rx.Matches(partNo)

	'Обозначение
	Dim PartDesignation As String = Left(partNo, matches.Item(0).Index+4)
	'Наименование
	Dim PartNameRight As String = Right(partNo, len(partNo)-matches.Item(0).Index-5)

	'Переписываем значения атрибутов
	Dim attributeValue As String = PartDesignation
	Dim attributeName As String = "OBOZNACHENIE"
	workPart.SetUserAttribute(attributeName, -1, attributeValue, Update.Option.Now)


End Sub


Function GetFileName()
		Dim strPath as String
		Dim strPart as String
		Dim pos as Integer
		
		'get the full file path
		strPath = displayPart.fullpath
		'get the part file name
		pos = InStrRev(strPath, "\")
		strPart = Mid(strPath, pos + 1)
		
		strPath = Left(strPath, pos)
		'strip off the ".prt" extension
		strPart = Left(strPart, Len(strPart) - 4)
		
		GetFileName = strPart
End Function


Function GetFilePath()
		Dim strPath as String
		Dim strPart as String
		Dim pos as Integer
		
		'get the full file path
		strPath = displayPart.fullpath
		'get the part file name
		pos = InStrRev(strPath, "\")
		strPart = Mid(strPath, pos + 1)
		
		strPath = Left(strPath, pos)
		'strip off the ".prt" extension
		strPart = Left(strPart, Len(strPart) - 4)
		
		GetFilePath = strPath
End Function

End module