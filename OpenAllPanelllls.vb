Option Strict Off
Imports System
Imports System.IO
Imports System.Collections
Imports System.Windows.Forms
Imports System.Windows.Forms.MessageBox
Imports NXOpen
Imports NXOpen.UF

Module NXJournal

	Sub Main(ByVal args() As String)
		Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
		Dim lw As ListingWindow = theSession.ListingWindow
		lw.Open()

		For Each foundFile As String In My.Computer.FileSystem.GetFiles("K:\Industrial\КД\Ящики злектрические",FileIO.SearchOption.SearchAllSubDirectories, "* Панель.prt")
			SquaringTheCircle(foundFile)
		Next
	End sub

	Sub SquaringTheCircle(ByVal SBPath as String)
		Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
		Dim basePart1 As NXOpen.BasePart
		Dim partLoadStatus1 As NXOpen.PartLoadStatus
		basePart1 = theSession.Parts.OpenBaseDisplay(SBPath, partLoadStatus1)
		Dim workPart As NXOpen.Part = theSession.Parts.Work
		Dim displayPart As NXOpen.Part = theSession.Parts.Display

	End sub

End Module