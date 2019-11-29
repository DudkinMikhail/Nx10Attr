Option Strict Off
Imports System
Imports NXOpen
Imports NXOpen.BlockStyler


Public Class CorpsContain
	'class members
	Private Shared theSession As Session
	Private Shared theUI As UI
	Private theDlxFileName As String
	Private theDialog As NXOpen.BlockStyler.BlockDialog
	Private toggle0 As NXOpen.BlockStyler.Toggle' Block type: Toggle
	Private toggle01 As NXOpen.BlockStyler.Toggle' Block type: Toggle
	Private toggle02 As NXOpen.BlockStyler.Toggle' Block type: Toggle
	Private toggle04 As NXOpen.BlockStyler.Toggle' Block type: Toggle
	Private toggle05 As NXOpen.BlockStyler.Toggle' Block type: Toggle
	Private toggle03 As NXOpen.BlockStyler.Toggle' Block type: Toggle
	
	'Загрузка DLX файла
	Public Sub New()
		theSession = Session.GetSession()
		theUI = UI.GetUI()
		theDlxFileName = "F:\VB.NET\CorpsContain.dlx"
		theDialog = theUI.CreateDialog(theDlxFileName)
		theDialog.AddApplyHandler(AddressOf apply_cb)
		theDialog.AddOkHandler(AddressOf ok_cb)
		theDialog.AddUpdateHandler(AddressOf update_cb)
		theDialog.AddInitializeHandler(AddressOf initialize_cb)
	End Sub
	
	'Инициализация элементов
	Public Sub initialize_cb()    
		toggle0 = theDialog.TopBlock.FindBlock("toggle0")
		toggle01 = theDialog.TopBlock.FindBlock("toggle01")
		toggle02 = theDialog.TopBlock.FindBlock("toggle02")
		toggle04 = theDialog.TopBlock.FindBlock("toggle04")
		toggle05 = theDialog.TopBlock.FindBlock("toggle05")
		toggle03 = theDialog.TopBlock.FindBlock("toggle03")
	End Sub

	'MAIN
	Public Shared Sub Main()
		Dim theCorpsContain As CorpsContain = Nothing
		theCorpsContain = New CorpsContain()
		theCorpsContain.Show()       
	End Sub
		
	'Открытие диалога
	Public Sub Show()        
		theDialog.Show
	End Sub
		
	'Отменить, закрыть
	Public Sub Dispose()
		theDialog.Dispose()
		theDialog = Nothing
	End Sub


	'Применить
	Public Function apply_cb() As Integer

	End Function

	'Обновление
	Public Function update_cb(ByVal block As NXOpen.BlockStyler.UIBlock) As Integer
	
	End Function
		
	'Ок
	Public Function ok_cb() As Integer

	End Function
		

	'Выгрузка
	Public Shared Function GetUnloadOption(ByVal arg As String) As Integer
	Return CType(Session.LibraryUnloadOption.Immediately, Integer)
	End Function
End Class
