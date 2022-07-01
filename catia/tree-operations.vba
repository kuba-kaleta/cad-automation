Sub myGetChildren()
    'the code here is structured similar to the c-sharp example.
    'This is the same API, it is just being called from VBA'

    'define variables'
    Dim editor As editor
    Dim root As VPMRootOccurrence
    Dim service As PLMProductService

    'get the active editor and from it get the PLMProductService to find the root.'
    Set editor = CATIA.ActiveEditor
    Set service = editor.GetService("PLMProductService")
    Set root = service.RootOccurrence

    'define a collection data structure'
    Dim children As Collection

    'get the chidlren
    Set children = GetChildren(root)

    Dim child As VPMOccurrence
    'loop over each child and print the name in a message box'
    For Each child In children
        MsgBox "Child Name Is " & child.Name
    Next

End Sub

'function to take a root and return a collection'
Function GetChildren(root As VPMRootOccurrence) As Collection

    Dim col As Collection
    Set col = New Collection

    'loop over nested occurrences in root'
    Dim I As Integer
    For I = 1 To root.Occurrences.Count
        'add occurrence to list'
        col.Add root.Occurrences.Item(I)

        'call recursive function to get all nested occurrences
        'inside current item and add it to the collection'
        GetNestedChildren root.Occurrences.Item(I), col
    Next

    Set GetChildren = col

End Function

'recursive function'
Sub GetNestedChildren(parent As VPMOccurrence, col As Collection)

    Dim I As Integer
    For I = 1 To parent.Occurrences.Count
        'add current occurence to collection'
        col.Add parent.Occurrences.Item(I)
        'call the method on the current occurrence'
        GetNestedChildren parent.Occurrences.Item(I), col
    Next

End Sub


Sub GetRoot()

    Dim oEditor As editor
    Set oEditor = CATIA.ActiveEditor
                 
    Dim oPLMProductService As PLMProductService
    Set oPLMProductService = oEditor.GetService("PLMProductService")
    
    Dim oRootOcc As VPMRootOccurrence
    Set oRootOcc = oPLMProductService.RootOccurrence
    
    Dim oRootRef As VPMReference
    Set oRootRef = oRootOcc.ReferenceRootOccurrenceOf
    
    Dim sVNameRootRef As String
    sVNameRootRef = oRootRef.GetAttributeValue("V_Name")
    
    MsgBox ("The name of the part is: " & sVNameRootRef)

End Sub


Sub CreateParameter()

	Dim editor1 As editor
	Set editor1 = CATIA.ActiveEditor

	Dim part1 As Part
	'MsgBox VarType(editor1.ActiveObject) = 9
	Set part1 = editor1.ActiveObject 

	Dim parameters1 As Parameters
	Set parameters1 = part1.Parameters

	Dim length1 As Length
	Set length1 = parameters1.CreateDimension("", "LENGTH", 0#)

End Sub


Sub CreateGeoSet()
	
	Dim settingControllers1 As SettingControllers
	Set settingControllers1 = CATIA.SettingControllers
	
	Dim settingRepository1 As SettingRepository
	Set settingRepository1 = settingControllers1.Item("Publications")

	boolean1 = settingRepository1.GetAttr("PublishAxisFeature") 'dziala tez bez booleanow
	boolean2 = settingRepository1.GetAttr("PublishAxisFeature")
	boolean3 = settingRepository1.GetAttr("PublishAxisFeature")
	boolean4 = settingRepository1.GetAttr("PublishAxisFeature")
	boolean5 = settingRepository1.GetAttr("PublishAxisFeature")
	boolean6 = settingRepository1.GetAttr("PublishAxisFeature")
	boolean7 = settingRepository1.GetAttr("PublishAxisFeature")
	boolean8 = settingRepository1.GetAttr("PublishAxisFeature")
	boolean9 = settingRepository1.GetAttr("PublishAxisFeature")
	boolean10 = settingRepository1.GetAttr("PublishAxisFeature")
	boolean11 = settingRepository1.GetAttr("PublishAxisFeature")
	boolean12 = settingRepository1.GetAttr("PublishAxisFeature")
	boolean13 = settingRepository1.GetAttr("PublishAxisFeature")
	boolean14 = settingRepository1.GetAttr("PublishAxisFeature")
	boolean15 = settingRepository1.GetAttr("PublishAxisFeature")
	boolean16 = settingRepository1.GetAttr("PublishAxisFeature")
	boolean17 = settingRepository1.GetAttr("PublishAxisFeature")
	boolean18 = settingRepository1.GetAttr("PublishAxisFeature")
	boolean19 = settingRepository1.GetAttr("PublishAxisFeature")
	boolean20 = settingRepository1.GetAttr("PublishAxisFeature")
	boolean21 = settingRepository1.GetAttr("PublishAxisFeature")
	boolean22 = settingRepository1.GetAttr("PublishAxisFeature")
	boolean23 = settingRepository1.GetAttr("PublishAxisFeature")
	boolean24 = settingRepository1.GetAttr("PublishAxisFeature")

	Dim editor1 As editor
	Set editor1 = CATIA.ActiveEditor

	Dim part1 As Part
	Set part1 = editor1.ActiveObject

	part1.Update

	Dim hybridBodies1 As HybridBodies
	Set hybridBodies1 = part1.HybridBodies

	Dim hybridBody1 As HybridBody
	Set hybridBody1 = hybridBodies1.Add()
	hybridBody1.Name = "test1"
    
End Sub


Sub PointCreation()

Dim editor1 As editor
Set editor1 = CATIA.ActiveEditor

Dim part1 As Part
Set part1 = editor1.ActiveObject

Dim hybridBodies1 As HybridBodies
Set hybridBodies1 = part1.HybridBodies
 
Dim hybridBody1 As HybridBody
Set hybridBody1 = hybridBodies1.Item("Geo1")

Dim hybridShapes1 As HybridShapes
Set hybridShapes1 = hybridBody1.HybridShapes
 
Dim hybridShapePointOnCurve1 As HybridShape
Set hybridShapePointOnCurve1 = hybridShapes1.Item("point3")

Dim reference1 As Reference
Set reference1 = part1.CreateReferenceFromObject(hybridShapePointOnCurve1)

Dim hybridShapeFactory1 As Factory
Set hybridShapeFactory1 = part1.HybridShapeFactory

Dim hybridShapePoint1 As HybridShapePointCoord
Set hybridShapePoint1 = hybridShapeFactory1.AddNewPointCoord(0#, 0#, 0#)

Dim hybridBody2 As HybridBody
Set hybridBody2 = hybridBodies1.Item("Geo2")

hybridBody2.AppendHybridShape hybridShapePoint1

End Sub


Sub PointSelection()

Dim oEditor As editor
Set oEditor = CATIA.ActiveEditor
                 
Dim oPLMProductService As PLMProductService
Set oPLMProductService = oEditor.GetService("PLMProductService")
    
Dim oRootOcc As VPMRootOccurrence
Set oRootOcc = oPLMProductService.RootOccurrence
    
Dim products1 As VPMReference
Set products1 = oRootOcc.ReferenceRootOccurrenceOf

Dim selection1 As Selection
Set selection1 = oEditor.Selection
selection1.Clear
selection1.Search ("Name='point3',all")

selection1.Item(1).Value

End Sub












'VB
'proba pomiaru w kontekscie
Set editor1 = CATIA.ActiveEditor
Set part1 = editor1.ActiveObject

Set annotationSets1 = part1.AnnotationSets
Set annotationSet1 = annotationSets1.Add("OBRSTER_3DEXPERIENCE_v2")

Dim coordArray(2)
p.GetCoordinates coordArray

Set reference1 = part1.CreateReferenceFromName(p)

Set userSurfaces1 = part1.UserSurfaces
Set userSurface1 = userSurfaces1.Generate(reference1)

Set annotationFactory1 = annotationSet1.AnnotationFactory
Set annotation1 = annotationFactory1.CreateEvoluateText(userSurface1 , -coordArray(0) -100 , -coordArray(1) , 0.0 , False)

annotation1.Text.Text =  text
Set oText = annotation1.Text.Get2dAnnot
oText.SetFontSize 1, k, 70

End Sub



Language="VBSCRIPT"

Dim coordArray(2)

Set editor1 = CATIA.ActiveEditor

Set part1 = editor1.ActiveObject

Set annotationSets1 = part1.AnnotationSets

Set annotationSet1 = annotationSets1.Add("OBRSTER_3DEXPERIENCE_v2")

Set hybridShapePointCoord1 = p

P.GetCoordinates coordArray

Set reference1 = part1.CreateReferenceFromObject(hybridShapePointCoord1)

Set userSurfaces1 = part1.UserSurfaces

Set userSurface1 = userSurfaces1.Generate(reference1)

Set annotationFactory1 = annotationSet1.AnnotationFactory

Set annotation1 = annotationFactory1.CreateEvoluateText(userSurface1 , -coordArray(0) -100 , -coordArray(1) , 0.0 , False)

annotation1.Text.Text =  text
Set oText = annotation1.Text.Get2dAnnot
oText.SetFontSize 1, k, 70

End Sub
