Option Strict Off
Imports System
Imports NXOpen
Module NXJournal

	Sub Main ()

	Dim theSession As Session = Session.GetSession()
	Dim workPart As Part = theSession.Parts.Work
	Dim objects(0) As NXObject
	objects(0) = workPart

	Dim attributePropertiesBuilder1 As AttributePropertiesBuilder
		attributePropertiesBuilder1 = workpart.PropertiesManager.CreateAttributePropertiesBuilder(objects)
		attributePropertiesBuilder1.ObjectPicker = AttributePropertiesBaseBuilder.ObjectOptions.ComponentAsPartAttribute
		attributePropertiesBuilder1.DataType = AttributePropertiesBaseBuilder.DataTypeOptions.String
		attributePropertiesBuilder1.Title = "IP"
		attributePropertiesBuilder1.IsArray = False
		attributePropertiesBuilder1.StringValue = "31"

	Dim nXObject1 As NXObject
		nXObject1 = attributePropertiesBuilder1.Commit()
		attributePropertiesBuilder1.Destroy()

	End Sub

End Module 