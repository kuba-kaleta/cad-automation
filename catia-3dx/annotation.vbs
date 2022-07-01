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
