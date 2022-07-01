' ---------------------------------------------------
' ***   Eksport wymiarow rysunku 3DX do excela    ***
' ***             Langage VBA                     ***
' ---------------------------------------------------


Sub CATMain()

Dim settingControllers1 As SettingControllers
Set settingControllers1 = CATIA.SettingControllers

Dim settingRepository1 As SettingRepository
Set settingRepository1 = settingControllers1.Item("Publications")

boolean1 = settingRepository1.GetAttr("PubDisplay")

boolean2 = settingRepository1.GetAttr("PubDisplay")

Dim settingRepository2 As SettingRepository
Set settingRepository2 = settingControllers1.Item("LPCommonEditor")

boolean3 = settingRepository2.GetAttr("DisplayRequirementsMode")

boolean4 = settingRepository2.GetAttr("LowLightMode")

boolean5 = settingRepository1.GetAttr("PubDisplay")

boolean6 = settingRepository2.GetAttr("DisplayRequirementsMode")

Dim editor1 As Editor
Set editor1 = CATIA.ActiveEditor

Dim drawingSheets1 As DrawingSheets
Set drawingSheets1 = editor1.ActiveObject

Dim oTolType As Long
Dim oTolName As String
Dim oUpTol As String
Dim oLowTol As String
Dim odUpTol As Double
Dim odLowTol As Double
Dim oDisplayMode As Long

' *** Wybieranie wszystkich wymiarow ***

Dim selection1 As Selection
Set selection1 = editor1.Selection
selection1.Clear
selection1.Search "CATDrwSearch.DrwDimension,all"

' *** Excel ***

Dim xl As Object 'Excel.Application
On Error Resume Next
Set xl = GetObject(, "Excel.Application")
If Err <> o Then
    Set xl = CreateObject("Excel.Application")
    xl.Visible = True
End If

Set workbooks = xl.Application.workbooks
Set myworkbook = xl.workbooks.Add
Set myworksheet = xl.ActiveWorkbook.Add
Set myworksheet = xl.Sheets.Add

' *** umieszczenie tytulow w arkuszu ***

myworksheet.Range("A1").Value = "Typ"
myworksheet.Range("B1").Value = "Wymiar"
myworksheet.Range("C1").Value = "Tolerancja dolna"
myworksheet.Range("D1").Value = "Tolerancja gorna"
myworksheet.Range("E1").Value = "Widok"

' *** rozmieszczanie wymiarow w arkuszu ***

For i = 1 To selection1.Count
    Set MyDimension = selection1.Item(i).Value
    MyDimensionValue = MyDimension.GetValue.Value
    ' tolerancje
    MyDimension.GetTolerances oTolType, oTolName, oUpTol, oLowTol, odUpTol, odLowTol, oDisplayMode
    myworksheet.cells(i + 1, 2).Value = Round(MyDimensionValue, 2)
    If oTolType = 1 Then 'numeryczne
        myworksheet.cells(i + 1, 3).Value = odLowTol
        myworksheet.cells(i + 1, 4).Value = odUpTol
    End If
    If oTolType = 2 Then 'alfanumeryczne
        myworksheet.cells(i + 1, 3).Value = oLowTol
        myworksheet.cells(i + 1, 4).Value = oUpTol
    End If
        
    ' Rodzaj wymiaru
    MyDimType = MyDimension.DimType
    Select Case MyDimType
        Case 5, 6, 7, 8, 17, 19         'promien
            MyDimTypeTexte = "R"
        Case 9, 10, 11, 12, 13, 18
            MyDimTypeTexte = "&Oslash;"        'srednica
        Case 14
            MyDimTypeTexte = "Ch"       'faza
        Case 4
            MyDimTypeTexte = "Angle"    'kat
        Case Else
            MyDimTypeTexte = "Length"         'dlugosc
    End Select
    myworksheet.cells(i + 1, 1).Value = MyDimTypeTexte
    
    If MyDimTypeTexte = "Angle" Then
        myworksheet.cells(i + 1, 2) = Round(MyDimensionValue * 180 / 3.1415926535, 2)
    End If
    
    myworksheet.cells(i + 1, 5).Value = MyDimension.Parent.Parent.Name
    odLowTol = 0
    odUpTol = 0
    oUpTol = ""
    oLowTol = ""
    
    Set ThisDrawingDim = selection1.Item(i).Value
    'ThisDrawingDim.ValueFrame = catFraRectangle
    
    Dim MyValue As DrawingDimValue
    Set MyValue = ThisDrawingDim.GetValue
    
    'MyValue.FakeDimType = catDimFakeNumValue
    If MyValue.GetFakeDimValue(1) <> "*Fake*" Then
        myworksheet.cells(i + 1, 2) = MyValue.GetFakeDimValue(1)
        myworksheet.cells(i + 1, 6) = "FAKE"
    End If

Next

Dim MyPartName As String
PartName = MyDimension.Parent.Parent.DrawingGenView.GetAssociatedRootProduct.GetAttributeValue("V_Name")

MkDir "C:\DimCapture"

Dim MyPath As String
MyPath = "C:\DimCapture\" & PartName & ".xlsx"
xl.ActiveWorkbook.SaveAs MyPath
xl.ActiveWorkbook.Save

'MsgBox MyDimension.Parent.Parent.GenerativeBehavior.Document.FullName
'MsgBox ProductDrawn.PartNumber

'Dim rootReference As VPMReference
'Set rootReference = GetRootReference(MyDimension.Parent.Parent.Parent)

End Sub


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
