Option Strict Off
Imports System
Imports System.IO
Imports System.Collections
Imports NXOpen
Imports NXOpen.UF
Imports Microsoft.VisualBasic


Module NXJournal

Sub Main(ByVal args() As String)

	Dim NewBoxName, Message, Title, DefaultValue, ServerPath, SourcePath As String
	Dim NewBoxH, NewBoxW, NewBoxT as integer
	dim NewBoxSbody, NewBoxSrear, NewBoxSdoor, NewBoxSpanel003, NewBoxIP as string
	
	'Постоянные переменные
	Message = "Введите наименование корпуса"
	Title = "Самый главный скрипт"
	DefaultValue = ""
	ServerPath = "K:\Industrial\КД\Ящики злектрические"
	SourcePath = "C:\SWPlus\NX10\Шаблоны деталей шкафов"

	'Перечень деталей для замены
	Dim ReplaceParts() As String = {"COMPONENT ЭРА..000.001 Этикетка 1", _
	"COMPONENT ЭРА..100.001 Стенка задняя 1", _
	"COMPONENT ЭРА..200.001 Корпус 1", _
	"COMPONENT ЭРА..300.003 Монтажная панель 1", _
	"COMPONENT ЭРА..300.004 Монтажная панель 1", _
	"COMPONENT ЭРА..300.005 Монтажная панель 1", _
	"COMPONENT ЭРА..400.001 Дверь наружная 1", _
	"COMPONENT ЭРА..400.002 Дверь наружная 1", _
	"COMPONENT ЭРА..300.011 Панель 1", _
	"COMPONENT ЭРА..Материалы.Уплотнитель дверной 1"}


	'Ввод названия
	NewBoxName = InputBox(message, title, DefaultValue)

	'Ввод высоты
	NewBoxH = InputBox("Высота, целое число", title, DefaultValue)
	
	'Ввод ширины
	NewBoxW = InputBox("Ширина, целое число", title, DefaultValue)
	
	'Ввод глубины
	NewBoxT = InputBox("Глубина, целое число", title, DefaultValue)

	'Ввод толщины корпуса
	NewBoxSbody = InputBox("Толщина металла корпуса", title, DefaultValue)

	'Ввод толщины задней стенки
	NewBoxSrear = InputBox("Толщина металла задней стенки", title, NewBoxSbody)

	'Ввод толщины дверки
	NewBoxSdoor = InputBox("Толщина металла дверки", title, NewBoxSbody)

	'Ввод толщины монтажной панели 003
	NewBoxSpanel003 = InputBox("Толщина металла монтажной панели .003", title, "0.8")
	
	'Ввод IP
	NewBoxIP = InputBox("IP? 31 54", title, "31")

	

	'Создаем рабочую папку c название нового шкафа на сервере
	Directory.CreateDirectory(ServerPath & "\" & NewBoxName)

	'Создаем подпапку Out
	Directory.CreateDirectory(ServerPath & "\" & NewBoxName & "\Out")

	'Массив имен файлов шаблонов
	Dim TemplatePartList As String() = Directory.GetFiles(SourcePath, "*.prt")

	'Копируем файлы в новую папку
	For Each i As String In TemplatePartList
		File.Copy(i, ServerPath & "\" & NewBoxName & "\" & Left(i.Substring(SourcePath.Length + 1),4) & NewBoxName & Mid(i.Substring(SourcePath.Length + 1),5))
	Next


	'Открываем сборку
	Dim partLoadStatus As NXOpen.PartLoadStatus
	Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
	theSession.Parts.OpenBaseDisplay(ServerPath & "\" & NewBoxName & "\ЭРА." & NewBoxName & ".000.000 СБ.prt", partLoadStatus)

	'Переключение в режим модели
	theSession.ApplicationSwitchImmediate("UG_APP_MODELING")

	'Объявления
	Dim workPart As NXOpen.Part = theSession.Parts.Work
	Dim displayPart As NXOpen.Part = theSession.Parts.Display
	
	'Объявление и открытие консоли
	'Dim lw As ListingWindow = theSession.ListingWindow
	'lw.Open()

	'Замена компонентов
	Dim kk As integer
	Dim a, b, c, d As string
	kk=0
	For Each k As String In ReplaceParts
		Dim replaceComponentBuilder1 As NXOpen.Assemblies.ReplaceComponentBuilder
		replaceComponentBuilder1 = workPart.AssemblyManager.CreateReplaceComponentBuilder()
		replaceComponentBuilder1.ComponentNameType = NXOpen.Assemblies.ReplaceComponentBuilder.ComponentNameOption.AsSpecified
	'	replaceComponentBuilder1.ComponentName = "ЭРА.ЩУ-МП 295Х190Х120 IP54.000.000 СБ"
		a = "ЭРА." & NewBoxName & ".000.000 СБ"
		replaceComponentBuilder1.ComponentName = a
	'	lw.Writeline(a)
		replaceComponentBuilder1.ComponentName = ""
	'	Dim component1 As NXOpen.Assemblies.Component = CType(workPart.ComponentAssembly.RootComponent.FindObject("COMPONENT ЭРА..200.001 Корпус 1"), NXOpen.Assemblies.Component)
		Dim component1 As NXOpen.Assemblies.Component = CType(workPart.ComponentAssembly.RootComponent.FindObject(ReplaceParts(kk)), NXOpen.Assemblies.Component)
		Dim added1 As Boolean
		added1 = replaceComponentBuilder1.ComponentsToReplace.Add(component1)
	'	replaceComponentBuilder1.ComponentName = "ЭРА..200.001 КОРПУС"
		b = UCase(Mid(ReplaceParts(kk), 12, Len(ReplaceParts(kk))-12))
		replaceComponentBuilder1.ComponentName = b
	'	lw.Writeline(b)
	'	replaceComponentBuilder1.ComponentName = "ЭРА.Test21.200.001 КОРПУС"
		c = "ЭРА." & NewBoxName & UCase(Mid(ReplaceParts(kk), 15, Len(ReplaceParts(kk))-15))
		replaceComponentBuilder1.ComponentName = c
	'	lw.Writeline(c)
	'	replaceComponentBuilder1.ReplacementPart = "K:\Industrial\КД\Ящики злектрические\Test21\ЭРА.Test21.200.001 Корпус.prt"
		d = ServerPath & "\" & NewBoxName & "\ЭРА." & NewBoxName & Mid(ReplaceParts(kk), 15, Len(ReplaceParts(kk))-16) & ".prt"
		replaceComponentBuilder1.ReplacementPart = d
	'	lw.Writeline(d)
		replaceComponentBuilder1.SetComponentReferenceSetType(NXOpen.Assemblies.ReplaceComponentBuilder.ComponentReferenceSet.Maintain, Nothing)
		Dim partLoadStatus1 As NXOpen.PartLoadStatus
		partLoadStatus1 = replaceComponentBuilder1.RegisterReplacePartLoadStatus()
		Dim nXObject1 As NXOpen.NXObject
		nXObject1 = replaceComponentBuilder1.Commit()
		partLoadStatus1.Dispose()
		replaceComponentBuilder1.Destroy()
		kk=kk+1
	Next
		kk=0

	'Скрываем сопряжения
	Dim numberHidden1 As Integer
	numberHidden1 = theSession.DisplayManager.HideByType("SHOW_HIDE_TYPE_ASSEMBLY_CONSTRAINTS", NXOpen.DisplayManager.ShowHideScope.AnyInAssembly)
	workPart.ModelingViews.WorkView.FitAfterShowOrHide(NXOpen.View.ShowOrHideType.HideOnly)

	'Переписываем значения атрибутов сборки
	'Обозначение
	workPart.SetUserAttribute("OBOZNACHENIE", -1, "ЭРА." & NewBoxName & ".000.000", Update.Option.Now)
	'Наименование
	workPart.SetUserAttribute("NAIMENOVANIE", -1, "Корпус " & NewBoxName, Update.Option.Now)
	'ТТ в зависимости от IP
	if NewBoxIP = "31" then
		Dim readText() As String = File.ReadAllLines("C:\SWPlus\NX10\Тектовые шаблоны\ТТ СБ IP31.txt")
		Dim TTFileContent As String
		Dim TTFileContentX As String
		For Each TTFileContent In readText
			TTFileContentX = TTFileContentX & TTFileContent & Chr(13)
		next
		workPart.SetUserAttribute("TT", -1, TTFileContentX, Update.Option.Now)
	end if
	
	if NewBoxIP = "54" then
		Dim readText() As String = File.ReadAllLines("C:\SWPlus\NX10\Тектовые шаблоны\ТТ СБ IP54.txt")
		Dim TTFileContent As String
		Dim TTFileContentX As String
		For Each TTFileContent In readText
			TTFileContentX = TTFileContentX & TTFileContent & Chr(13)
		next
		workPart.SetUserAttribute("TT", -1, TTFileContentX, Update.Option.Now)
	end if

	'Сохранение сборки
	Dim partSaveStatus1 As NXOpen.PartSaveStatus
	partSaveStatus1 = workPart.Save(NXOpen.BasePart.SaveComponents.True, NXOpen.BasePart.CloseAfterSave.False)
	partSaveStatus1.Dispose()

	'Закрытие сборки
	Dim partCloseResponses1 As NXOpen.PartCloseResponses
	partCloseResponses1 = theSession.Parts.NewPartCloseResponses()
	workPart.Close(NXOpen.BasePart.CloseWholeTree.True, NXOpen.BasePart.CloseModified.UseResponses, partCloseResponses1)
	workPart = Nothing
	displayPart = Nothing
	partCloseResponses1.Dispose()

	'Переназначение свойств (обозначение, наименование)
	For Each k As String In ReplaceParts
		'Уплотнитель не требует переназначения свойств
		if ReplaceParts(kk)="COMPONENT ЭРА..Материалы.Уплотнитель дверной 1" Then Continue For
		NewPartEdit(ServerPath, NewBoxName, ReplaceParts(kk))
		kk=kk+1
	Next
	kk=0
	
	'Изменение параметров деталей корпуса
	For Each k As String In ReplaceParts
		if ReplaceParts(kk)="COMPONENT ЭРА..300.004 Монтажная панель 1" then
			kk=kk+1
			Continue For
		end if
		if ReplaceParts(kk)="COMPONENT ЭРА..300.005 Монтажная панель 1" then 
			kk=kk+1
			Continue For
		end if
		NewPartParamEdit(ServerPath, NewBoxName, ReplaceParts(kk), NewBoxH, NewBoxW, NewBoxT, NewBoxSbody, NewBoxSrear, NewBoxSdoor, NewBoxSpanel003)
		kk=kk+1
	Next
	kk=0


	'Переназначение аттрибутов и замена файла сборки в ВП
	VPParamEdit(ServerPath, NewBoxName)
	
	
	'Переназначение аттрибутов и замена файла сборки в исполнении 01
	Door01ParamEdit(ServerPath, NewBoxName)

End sub


'Переназначение аттрибутов и замена файла сборки в ВП
Sub VPParamEdit(ByVal ServerPathX as String, ByVal NewBoxNameX as String)

	'Открываем ведомость покупных
	Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
	Dim basePart1 As NXOpen.BasePart
	Dim partLoadStatus1 As NXOpen.PartLoadStatus
	basePart1 = theSession.Parts.OpenBaseDisplay(ServerPathX & "\" & NewBoxNameX & "\ЭРА." & NewBoxNameX & ".000.000 ВП.prt", partLoadStatus1)
	Dim workPart As NXOpen.Part = theSession.Parts.Work
	Dim displayPart As NXOpen.Part = theSession.Parts.Display

	'Переписываем значения атрибутов ведомости покупных
	'Обозначение
	workPart.SetUserAttribute("OBOZNACHENIE", -1, "ЭРА." & NewBoxNameX & ".000.000 ВП", Update.Option.Now)
	'Наименование
	workPart.SetUserAttribute("NAIMENOVANIE", -1, "Корпус " & NewBoxNameX, Update.Option.Now)

	'Замена сборки
		Dim a, b, c, d As string
		Dim replaceComponentBuilder1 As NXOpen.Assemblies.ReplaceComponentBuilder
		replaceComponentBuilder1 = workPart.AssemblyManager.CreateReplaceComponentBuilder()
		replaceComponentBuilder1.ComponentNameType = NXOpen.Assemblies.ReplaceComponentBuilder.ComponentNameOption.AsSpecified
	'	replaceComponentBuilder1.ComponentName = "ЭРА.ЩУ-МП 295Х190Х120 IP54.000.000 СБ"
		a = "ЭРА." & NewBoxNameX & ".000.000 ВП"
		replaceComponentBuilder1.ComponentName = a
	'	lw.Writeline(a)
		replaceComponentBuilder1.ComponentName = ""
	'	Dim component1 As NXOpen.Assemblies.Component = CType(workPart.ComponentAssembly.RootComponent.FindObject("COMPONENT ЭРА..200.001 Корпус 1"), NXOpen.Assemblies.Component)
		Dim component1 As NXOpen.Assemblies.Component = CType(workPart.ComponentAssembly.RootComponent.FindObject("COMPONENT ЭРА..000.000 СБ 1"), NXOpen.Assemblies.Component)
		Dim added1 As Boolean
		added1 = replaceComponentBuilder1.ComponentsToReplace.Add(component1)
	'	replaceComponentBuilder1.ComponentName = "ЭРА..200.001 КОРПУС"
		b = UCase(Mid("COMPONENT ЭРА..000.000 СБ 1", 12, Len("COMPONENT ЭРА..000.000 СБ 1")-12))
		replaceComponentBuilder1.ComponentName = b
	'	lw.Writeline(b)
	'	replaceComponentBuilder1.ComponentName = "ЭРА.Test21.200.001 КОРПУС"
		c = "ЭРА." & NewBoxNameX & UCase(Mid("COMPONENT ЭРА..000.000 СБ 1", 15, Len("COMPONENT ЭРА..000.000 СБ 1")-15))
		replaceComponentBuilder1.ComponentName = c
	'	lw.Writeline(c)
	'	replaceComponentBuilder1.ReplacementPart = "K:\Industrial\КД\Ящики злектрические\Test21\ЭРА.Test21.200.001 Корпус.prt"
		d = ServerPathX & "\" & NewBoxNameX & "\ЭРА." & NewBoxNameX & Mid("COMPONENT ЭРА..000.000 СБ 1", 15, Len("COMPONENT ЭРА..000.000 СБ 1")-16) & ".prt"
		replaceComponentBuilder1.ReplacementPart = d
	'	lw.Writeline(d)
		replaceComponentBuilder1.SetComponentReferenceSetType(NXOpen.Assemblies.ReplaceComponentBuilder.ComponentReferenceSet.Maintain, Nothing)
		Dim partLoadStatus2 As NXOpen.PartLoadStatus
		partLoadStatus2 = replaceComponentBuilder1.RegisterReplacePartLoadStatus()
		Dim nXObject1 As NXOpen.NXObject
		nXObject1 = replaceComponentBuilder1.Commit()
		partLoadStatus2.Dispose()
		replaceComponentBuilder1.Destroy()

	'Сохранение ВП
	Dim partSaveStatus1 As NXOpen.PartSaveStatus
	partSaveStatus1 = workPart.Save(NXOpen.BasePart.SaveComponents.True, NXOpen.BasePart.CloseAfterSave.False)
	partSaveStatus1.Dispose()

	'Закрытие ВП
	Dim partCloseResponses1 As NXOpen.PartCloseResponses
	partCloseResponses1 = theSession.Parts.NewPartCloseResponses()
	workPart.Close(NXOpen.BasePart.CloseWholeTree.True, NXOpen.BasePart.CloseModified.UseResponses, partCloseResponses1)
	workPart = Nothing
	displayPart = Nothing
	partCloseResponses1.Dispose()

End Sub

'Переназначение аттрибутов и замена файла сборки в исполнении 01
Sub Door01ParamEdit(ByVal ServerPathX as String, ByVal NewBoxNameX as String)

	'Открываем ведомость покупных
	Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
	Dim basePart1 As NXOpen.BasePart
	Dim partLoadStatus1 As NXOpen.PartLoadStatus
	basePart1 = theSession.Parts.OpenBaseDisplay(ServerPathX & "\" & NewBoxNameX & "\ЭРА." & NewBoxNameX & ".000.000-01 СБ.prt", partLoadStatus1)
	Dim workPart As NXOpen.Part = theSession.Parts.Work
	Dim displayPart As NXOpen.Part = theSession.Parts.Display

	'Переписываем значения атрибутов ведомости покупных
	'Наименование
	workPart.SetUserAttribute("NAIMENOVANIE", -1, "Корпус " & NewBoxNameX, Update.Option.Now)

	'Обозначение
	workPart.SetUserAttribute("OBOZNACHENIE", -1, "ЭРА." & NewBoxNameX & ".000.000-01", Update.Option.Now)

	'Замена сборки
		Dim a, b, c, d As string
		Dim replaceComponentBuilder1 As NXOpen.Assemblies.ReplaceComponentBuilder
		replaceComponentBuilder1 = workPart.AssemblyManager.CreateReplaceComponentBuilder()
		replaceComponentBuilder1.ComponentNameType = NXOpen.Assemblies.ReplaceComponentBuilder.ComponentNameOption.AsSpecified
	'	replaceComponentBuilder1.ComponentName = "ЭРА.ЩУ-МП 295Х190Х120 IP54.000.000 СБ"
		a = "ЭРА." & NewBoxNameX & ".000.000-01 СБ"
		replaceComponentBuilder1.ComponentName = a
	'	lw.Writeline(a)
		replaceComponentBuilder1.ComponentName = ""
	'	Dim component1 As NXOpen.Assemblies.Component = CType(workPart.ComponentAssembly.RootComponent.FindObject("COMPONENT ЭРА..200.001 Корпус 1"), NXOpen.Assemblies.Component)
		Dim component1 As NXOpen.Assemblies.Component = CType(workPart.ComponentAssembly.RootComponent.FindObject("COMPONENT ЭРА..000.000 СБ 1"), NXOpen.Assemblies.Component)
		Dim added1 As Boolean
		added1 = replaceComponentBuilder1.ComponentsToReplace.Add(component1)
	'	replaceComponentBuilder1.ComponentName = "ЭРА..200.001 КОРПУС"
		b = UCase(Mid("COMPONENT ЭРА..000.000 СБ 1", 12, Len("COMPONENT ЭРА..000.000 СБ 1")-12))
		replaceComponentBuilder1.ComponentName = b
	'	lw.Writeline(b)
	'	replaceComponentBuilder1.ComponentName = "ЭРА.Test21.200.001 КОРПУС"
		c = "ЭРА." & NewBoxNameX & UCase(Mid("COMPONENT ЭРА..000.000 СБ 1", 15, Len("COMPONENT ЭРА..000.000 СБ 1")-15))
		replaceComponentBuilder1.ComponentName = c
	'	lw.Writeline(c)
	'	replaceComponentBuilder1.ReplacementPart = "K:\Industrial\КД\Ящики злектрические\Test21\ЭРА.Test21.200.001 Корпус.prt"
		d = ServerPathX & "\" & NewBoxNameX & "\ЭРА." & NewBoxNameX & Mid("COMPONENT ЭРА..000.000 СБ 1", 15, Len("COMPONENT ЭРА..000.000 СБ 1")-16) & ".prt"
		replaceComponentBuilder1.ReplacementPart = d
	'	lw.Writeline(d)
		replaceComponentBuilder1.SetComponentReferenceSetType(NXOpen.Assemblies.ReplaceComponentBuilder.ComponentReferenceSet.Maintain, Nothing)
		Dim partLoadStatus2 As NXOpen.PartLoadStatus
		partLoadStatus2 = replaceComponentBuilder1.RegisterReplacePartLoadStatus()
		Dim nXObject1 As NXOpen.NXObject
		nXObject1 = replaceComponentBuilder1.Commit()
		partLoadStatus2.Dispose()
		replaceComponentBuilder1.Destroy()

	'Сохранение ВП
	Dim partSaveStatus1 As NXOpen.PartSaveStatus
	partSaveStatus1 = workPart.Save(NXOpen.BasePart.SaveComponents.True, NXOpen.BasePart.CloseAfterSave.False)
	partSaveStatus1.Dispose()

	'Закрытие ВП
	Dim partCloseResponses1 As NXOpen.PartCloseResponses
	partCloseResponses1 = theSession.Parts.NewPartCloseResponses()
	workPart.Close(NXOpen.BasePart.CloseWholeTree.True, NXOpen.BasePart.CloseModified.UseResponses, partCloseResponses1)
	workPart = Nothing
	displayPart = Nothing
	partCloseResponses1.Dispose()

End Sub


'Изменение параметров деталей корпуса
Sub NewPartParamEdit(ByVal ServerPathX as String, ByVal NewBoxNameX as String, ByVal ReplacePartsX as String, NewBoxH as integer, _
	NewBoxW as integer, NewBoxT as integer, NewBoxSbody as string, NewBoxSrear as string, NewBoxSdoor as string, NewBoxSpanel003 as string)

	'Открываем деталь
	Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
	Dim basePart1 As NXOpen.BasePart
	Dim partLoadStatus1 As NXOpen.PartLoadStatus
	basePart1 = theSession.Parts.OpenBaseDisplay(ServerPathX & "\" & NewBoxNameX & "\ЭРА." & NewBoxNameX & Mid(ReplacePartsX, 15, Len(ReplacePartsX)-16) & ".prt", partLoadStatus1)
	Dim workPart As NXOpen.Part = theSession.Parts.Work
	Dim displayPart As NXOpen.Part = theSession.Parts.Display


	'Корпус
	if ReplacePartsX="COMPONENT ЭРА..200.001 Корпус 1" Then
		'Переключение в режим листового металла
		theSession.ApplicationSwitchImmediate("UG_APP_SBSM")
		Dim expression1 As NXOpen.Expression = CType(workPart.Expressions.FindObject("H"), NXOpen.Expression)
		Dim unit1 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
		workPart.Expressions.EditWithUnits(expression1, unit1, NewBoxH)
		Dim expression2 As NXOpen.Expression = CType(workPart.Expressions.FindObject("W"), NXOpen.Expression)
		Dim unit2 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
		workPart.Expressions.EditWithUnits(expression2, unit2, NewBoxW)
		Dim expression3 As NXOpen.Expression = CType(workPart.Expressions.FindObject("T"), NXOpen.Expression)
		Dim unit3 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
		workPart.Expressions.EditWithUnits(expression3, unit3, NewBoxT-19)
		Dim expression4 As NXOpen.Expression = CType(workPart.Expressions.FindObject("p3267"), NXOpen.Expression)
		Dim unit4 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
		Dim expression5 As NXOpen.Expression = CType(workPart.Expressions.FindObject("p4625"), NXOpen.Expression)
		Dim unit5 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
		if NewBoxSbody = "0.5" then
			workPart.Expressions.EditWithUnits(expression4, unit4, "36.1")
			workPart.Expressions.EditWithUnits(expression5, unit5, "36.1")
			'Замена ТТ (радиус гиба)
			Dim readText() As String = File.ReadAllLines("C:\SWPlus\NX10\Тектовые шаблоны\ТТ Корпус 0.5.txt")
			Dim TTFileContent As String
			Dim TTFileContentX As String
				For Each TTFileContent In readText
					TTFileContentX = TTFileContentX & TTFileContent & Chr(13)
				next
			workPart.SetUserAttribute("TT", -1, TTFileContentX, Update.Option.Now)
		End if
		if NewBoxSbody = "0.6" then
			workPart.Expressions.EditWithUnits(expression4, unit4, "36.3")
			workPart.Expressions.EditWithUnits(expression5, unit5, "36.3")
			'Замена ТТ (радиус гиба)
			Dim readText() As String = File.ReadAllLines("C:\SWPlus\NX10\Тектовые шаблоны\ТТ Корпус 0.6.txt")
			Dim TTFileContent As String
			Dim TTFileContentX As String
				For Each TTFileContent In readText
					TTFileContentX = TTFileContentX & TTFileContent & Chr(13)
				next
			workPart.SetUserAttribute("TT", -1, TTFileContentX, Update.Option.Now)
		End if
		if NewBoxSbody = "0.8" then
			workPart.Expressions.EditWithUnits(expression4, unit4, "36.7")
			workPart.Expressions.EditWithUnits(expression5, unit5, "36.7")
			'Замена ТТ (радиус гиба)
			Dim readText() As String = File.ReadAllLines("C:\SWPlus\NX10\Тектовые шаблоны\ТТ Корпус 0.8.txt")
			Dim TTFileContent As String
			Dim TTFileContentX As String
				For Each TTFileContent In readText
					TTFileContentX = TTFileContentX & TTFileContent & Chr(13)
				next
			workPart.SetUserAttribute("TT", -1, TTFileContentX, Update.Option.Now)
		End if
		workPart.Preferences.PartSheetmetal.SetThickness(False, NewBoxSbody)
		workPart.Preferences.PartSheetmetal.SetBendRadius(False, NewBoxSbody)
		workPart.Preferences.PartSheetmetal.Commit()
		theSession.Preferences.Modeling.UpdatePending = False
	end if

	'Задняя стенка
	if ReplacePartsX="COMPONENT ЭРА..100.001 Стенка задняя 1" Then
		'Переключение в режим листового металла
		theSession.ApplicationSwitchImmediate("UG_APP_SBSM")
		Dim expression1 As NXOpen.Expression = CType(workPart.Expressions.FindObject("H"), NXOpen.Expression)
		Dim unit1 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
		workPart.Expressions.EditWithUnits(expression1, unit1, NewBoxH-5)
		Dim expression2 As NXOpen.Expression = CType(workPart.Expressions.FindObject("W"), NXOpen.Expression)
		Dim unit2 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
		workPart.Expressions.EditWithUnits(expression2, unit2, NewBoxW-5)
		workPart.Preferences.PartSheetmetal.SetThickness(False, NewBoxSrear)
		workPart.Preferences.PartSheetmetal.SetBendRadius(False, NewBoxSrear)
		workPart.Preferences.PartSheetmetal.Commit()
		theSession.Preferences.Modeling.UpdatePending = False
	end if

	'Монтажная панель 003
	if ReplacePartsX="COMPONENT ЭРА..300.003 Монтажная панель 1" Then
		'Переключение в режим листового металла
		theSession.ApplicationSwitchImmediate("UG_APP_SBSM")
		Dim expression1 As NXOpen.Expression = CType(workPart.Expressions.FindObject("H"), NXOpen.Expression)
		Dim unit1 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
		'Суммарный отступ для монтажной панели по высоте 35 мм
		workPart.Expressions.EditWithUnits(expression1, unit1, NewBoxH-35)
		Dim expression2 As NXOpen.Expression = CType(workPart.Expressions.FindObject("W"), NXOpen.Expression)
		Dim unit2 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
		'Суммарный отступ для монтажной панели по ширине 35 мм
		workPart.Expressions.EditWithUnits(expression2, unit2, NewBoxW-35)
		workPart.Preferences.PartSheetmetal.SetThickness(False, NewBoxSpanel003)
		workPart.Preferences.PartSheetmetal.SetBendRadius(False, NewBoxSpanel003)
		workPart.Preferences.PartSheetmetal.Commit()
		theSession.Preferences.Modeling.UpdatePending = False
	end if

	'Коррекция высоты и ширины дверей
	Dim NewBoxHx, NewBoxWx as String
	if NewBoxSbody = "0.5" then
		NewBoxHx = CStr(NewBoxH + 0.9554)
		NewBoxWx = CStr(NewBoxW + 0.9554)
		'Функция Cstr после преобразованя ставит запятую, что недопустимо, меняем на точку
		NewBoxHx = NewBoxHx.Replace(",",".")
		NewBoxWx = NewBoxWx.Replace(",",".")
	end if
	if NewBoxSbody = "0.6" then
		NewBoxHx = CStr(NewBoxH + 1.1465)
		NewBoxWx = CStr(NewBoxW + 1.1465)
		NewBoxHx = NewBoxHx.Replace(",",".")
		NewBoxWx = NewBoxWx.Replace(",",".")
	end if
	if NewBoxSbody = "0.8" then
		NewBoxHx = CStr(NewBoxH + 1.5287)
		NewBoxWx = CStr(NewBoxW + 1.5287)
		NewBoxHx = NewBoxHx.Replace(",",".")
		NewBoxWx = NewBoxWx.Replace(",",".")
	end if
	
	'Дверка 001
	if ReplacePartsX="COMPONENT ЭРА..400.001 Дверь наружная 1" Then
		'Переключение в режим листового металла
		theSession.ApplicationSwitchImmediate("UG_APP_SBSM")
		Dim expression1 As NXOpen.Expression = CType(workPart.Expressions.FindObject("H"), NXOpen.Expression)
		Dim unit1 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
		if NewBoxSbody = "0.5" then
			workPart.Expressions.EditWithUnits(expression1, unit1, NewBoxHx)
			'Замена ТТ (радиус гиба)
			Dim readText() As String = File.ReadAllLines("C:\SWPlus\NX10\Тектовые шаблоны\ТТ Дверка 0.5.txt")
			Dim TTFileContent As String
			Dim TTFileContentX As String
				For Each TTFileContent In readText
					TTFileContentX = TTFileContentX & TTFileContent & Chr(13)
				next
			workPart.SetUserAttribute("TT", -1, TTFileContentX, Update.Option.Now)
		end if
		if NewBoxSbody = "0.6" then
			workPart.Expressions.EditWithUnits(expression1, unit1, NewBoxHx)
			Dim readText() As String = File.ReadAllLines("C:\SWPlus\NX10\Тектовые шаблоны\ТТ Дверка 0.6.txt")
			Dim TTFileContent As String
			Dim TTFileContentX As String
				For Each TTFileContent In readText
					TTFileContentX = TTFileContentX & TTFileContent & Chr(13)
				next
			workPart.SetUserAttribute("TT", -1, TTFileContentX, Update.Option.Now)
		end if
		if NewBoxSbody = "0.8" then
			workPart.Expressions.EditWithUnits(expression1, unit1, NewBoxHx)
			Dim readText() As String = File.ReadAllLines("C:\SWPlus\NX10\Тектовые шаблоны\ТТ Дверка 0.8.txt")
			Dim TTFileContent As String
			Dim TTFileContentX As String
				For Each TTFileContent In readText
					TTFileContentX = TTFileContentX & TTFileContent & Chr(13)
				next
			workPart.SetUserAttribute("TT", -1, TTFileContentX, Update.Option.Now)
		end if
		Dim expression2 As NXOpen.Expression = CType(workPart.Expressions.FindObject("W"), NXOpen.Expression)
		Dim unit2 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
		if NewBoxSbody = "0.5" then workPart.Expressions.EditWithUnits(expression2, unit2, NewBoxWx)
		if NewBoxSbody = "0.6" then workPart.Expressions.EditWithUnits(expression2, unit2, NewBoxWx)
		if NewBoxSbody = "0.8" then workPart.Expressions.EditWithUnits(expression2, unit2, NewBoxWx)
		workPart.Preferences.PartSheetmetal.SetThickness(False, NewBoxSdoor)
		workPart.Preferences.PartSheetmetal.SetBendRadius(False, NewBoxSdoor)
		workPart.Preferences.PartSheetmetal.Commit()
		theSession.Preferences.Modeling.UpdatePending = False
	end if

	'Дверка 001
	if ReplacePartsX="COMPONENT ЭРА..400.002 Дверь наружная 1" Then
		'Переключение в режим листового металла
		theSession.ApplicationSwitchImmediate("UG_APP_SBSM")
		Dim expression1 As NXOpen.Expression = CType(workPart.Expressions.FindObject("H"), NXOpen.Expression)
		Dim unit1 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
		if NewBoxSbody = "0.5" then
			workPart.Expressions.EditWithUnits(expression1, unit1, NewBoxHx)
			'Замена ТТ (разный радиус гиба)
			Dim readText() As String = File.ReadAllLines("C:\SWPlus\NX10\Тектовые шаблоны\ТТ Дверка 0.5.txt")
			Dim TTFileContent As String
			Dim TTFileContentX As String
				For Each TTFileContent In readText
					TTFileContentX = TTFileContentX & TTFileContent & Chr(13)
				next
			workPart.SetUserAttribute("TT", -1, TTFileContentX, Update.Option.Now)
		end if
		if NewBoxSbody = "0.6" then
			workPart.Expressions.EditWithUnits(expression1, unit1, NewBoxHx)
			Dim readText() As String = File.ReadAllLines("C:\SWPlus\NX10\Тектовые шаблоны\ТТ Дверка 0.6.txt")
			Dim TTFileContent As String
			Dim TTFileContentX As String
				For Each TTFileContent In readText
					TTFileContentX = TTFileContentX & TTFileContent & Chr(13)
				next
			workPart.SetUserAttribute("TT", -1, TTFileContentX, Update.Option.Now)
		end if
		if NewBoxSbody = "0.8" then
			workPart.Expressions.EditWithUnits(expression1, unit1, NewBoxHx)
			Dim readText() As String = File.ReadAllLines("C:\SWPlus\NX10\Тектовые шаблоны\ТТ Дверка 0.8.txt")
			Dim TTFileContent As String
			Dim TTFileContentX As String
				For Each TTFileContent In readText
					TTFileContentX = TTFileContentX & TTFileContent & Chr(13)
				next
			workPart.SetUserAttribute("TT", -1, TTFileContentX, Update.Option.Now)
		end if
		Dim expression2 As NXOpen.Expression = CType(workPart.Expressions.FindObject("W"), NXOpen.Expression)
		Dim unit2 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
		if NewBoxSbody = "0.5" then workPart.Expressions.EditWithUnits(expression2, unit2, NewBoxWx)
		if NewBoxSbody = "0.6" then workPart.Expressions.EditWithUnits(expression2, unit2, NewBoxWx)
		if NewBoxSbody = "0.8" then workPart.Expressions.EditWithUnits(expression2, unit2, NewBoxWx)
		workPart.Preferences.PartSheetmetal.SetThickness(False, NewBoxSdoor)
		workPart.Preferences.PartSheetmetal.SetBendRadius(False, NewBoxSdoor)
		workPart.Preferences.PartSheetmetal.Commit()
		theSession.Preferences.Modeling.UpdatePending = False
	end if

	'Уплотнитель
	if ReplacePartsX="COMPONENT ЭРА..Материалы.Уплотнитель дверной 1" Then
		Dim expression1 As NXOpen.Expression = CType(workPart.Expressions.FindObject("H"), NXOpen.Expression)
		Dim unit1 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
		workPart.Expressions.EditWithUnits(expression1, unit1, NewBoxH-30)
		Dim expression2 As NXOpen.Expression = CType(workPart.Expressions.FindObject("W"), NXOpen.Expression)
		Dim unit2 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
		workPart.Expressions.EditWithUnits(expression2, unit2, NewBoxW-30)
		theSession.Preferences.Modeling.UpdatePending = False
	end if


	'Сохраняем
	Dim partSaveStatus1 As NXOpen.PartSaveStatus
	partSaveStatus1 = workPart.Save(NXOpen.BasePart.SaveComponents.True, NXOpen.BasePart.CloseAfterSave.True)
	partSaveStatus1.Dispose()

End sub


'Переназначение атрибутов (скрипт КОСТЫЛЬ как основа)
Sub NewPartEdit(ByVal ServerPathX as String, ByVal NewBoxNameX as String, ByVal ReplacePartsX as String)

	'Открываем деталь
	Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
	Dim basePart1 As NXOpen.BasePart
	Dim partLoadStatus1 As NXOpen.PartLoadStatus
	basePart1 = theSession.Parts.OpenBaseDisplay(ServerPathX & "\" & NewBoxNameX & "\ЭРА." & NewBoxNameX & Mid(ReplacePartsX, 15, Len(ReplacePartsX)-16) & ".prt", partLoadStatus1)
	Dim workPart As NXOpen.Part = theSession.Parts.Work
	Dim displayPart As NXOpen.Part = theSession.Parts.Display

	'Переписываем значения атрибутов детали:
	'Обозначение
	workPart.SetUserAttribute("OBOZNACHENIE", -1, "ЭРА." & NewBoxNameX & Mid(ReplacePartsX, 15, 8), Update.Option.Now)
	'Наименование
	workPart.SetUserAttribute("NAIMENOVANIE", -1, Mid(ReplacePartsX, 24, Len(ReplacePartsX)-25), Update.Option.Now)

	'Переключение в режим листового металла
	theSession.ApplicationSwitchImmediate("UG_APP_SBSM")

	'Сохраняем
	Dim partSaveStatus1 As NXOpen.PartSaveStatus
	partSaveStatus1 = workPart.Save(NXOpen.BasePart.SaveComponents.True, NXOpen.BasePart.CloseAfterSave.True)
	partSaveStatus1.Dispose()

End sub


Public Function GetUnloadOption(ByVal dummy As String) As Integer
GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately
End Function
End module