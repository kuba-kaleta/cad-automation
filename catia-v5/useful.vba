'Names
Sub CATMain()
    
    Dim ExcelSheet As Object
    On Error Resume Next

'if Excel is already running, then get the Excel object
            'Set ExcelSheet = GetObject(, "Excel.Application")
  
            'If Err.Number <> 0 Then
'If Excel is not already running, then create a new session of Excel

                Set ExcelSheet = CreateObject("Excel.Application")

                  ExcelSheet.Visible = True

            'End If
    
    'add a new workbook
                  ExcelSheet.Workbooks.Add

'set the workbook as the active document
        Set WBK = Excel.ActiveWorkbook.Sheets(1)
    
    GetNextNode catia.ActiveDocument.Product, ExcelSheet
  
End Sub

 
Sub GetNextNode(oCurrentProduct As Product, ExcelSheet As Object)

    Dim oCurrentTreeNode As Product
    Dim I As Integer
    Dim shift As Integer
    
    ExcelSheet.Application.Cells(1, 1).Value = "Part"
    ExcelSheet.Application.Cells(1, 2).Value = "Name"
    ExcelSheet.Application.Cells(1, 3).Value = "Description"
    ExcelSheet.Application.Cells(1, 4).Value = "Definition"
    ExcelSheet.Application.Cells(1, 5).Value = "DescRef"

    ' Loop through every tree node for the current product
    shift = 0
    For I = 1 To oCurrentProduct.Products.Count
        Set oCurrentTreeNode = oCurrentProduct.Products.Item(I)
        'MsgBox i
        ' Determine if the current node is a part, product or component
        If IsPart(oCurrentTreeNode) = True Then
            'MsgBox oCurrentTreeNode.PartNumber & " is a part"
            ExcelSheet.Application.Cells(I + 1 - shift, 1).Value = oCurrentTreeNode.PartNumber
            ExcelSheet.Application.Cells(I + 1 - shift, 2).Value = oCurrentTreeNode.Name
            ExcelSheet.Application.Cells(I + 1 - shift, 3).Value = oCurrentTreeNode.DescriptionInst
            ExcelSheet.Application.Cells(I + 1 - shift, 4).Value = oCurrentTreeNode.Definition
            ExcelSheet.Application.Cells(I + 1 - shift, 5).Value = oCurrentTreeNode.DescriptionRef
            
        Else
            'MsgBox oCurrentTreeNode.PartNumber & " is a component"
            shift = shift + 1
        End If
       
        ' if sub-nodes exist below the current tree node, call the sub recursively
        If oCurrentTreeNode.Products.Count > 0 Then
            GetNextNode oCurrentTreeNode, ExcelSheet
        End If
     
   Next
   
   'ExcelSheet.Application.Visible = True 'ExcelSheet.Application.Visible = True
   MsgBox "Ok, jesli edycja zakonczona"

   shift = 0
    For I = 1 To oCurrentProduct.Products.Count
        Set oCurrentTreeNode = oCurrentProduct.Products.Item(I)

        ' Determine if the current node is a part, product or component
        If IsPart(oCurrentTreeNode) = True Then
            'MsgBox oCurrentTreeNode.PartNumber & " is a part"
            oCurrentTreeNode.PartNumber = ExcelSheet.Application.Cells(I + 1 - shift, 1).Value
            oCurrentTreeNode.Name = ExcelSheet.Application.Cells(I + 1 - shift, 2).Value
            oCurrentTreeNode.DescriptionInst = ExcelSheet.Application.Cells(I + 1 - shift, 3).Value
            oCurrentTreeNode.Definition = ExcelSheet.Application.Cells(I + 1 - shift, 4).Value
            oCurrentTreeNode.DescriptionRef = ExcelSheet.Application.Cells(I + 1 - shift, 5).Value
           
        Else
            'MsgBox oCurrentTreeNode.PartNumber & " is a component"
            shift = shift + 1
        End If
       
       
        ' if sub-nodes exist below the current tree node, call the sub recursively
        If oCurrentTreeNode.Products.Count > 0 Then
            GetNextNode oCurrentTreeNode, ExcelSheet
        End If
     
   Next
   
    'ExcelSheet.SaveAs "C:\Users\jakub.krajanowski\jakub.kaleta\Catia\Automatyzacja\Names\" + catia.ActiveDocument.Product.Name + ".xlsx" '"\\10.200.11.150\OBR-Konstrukcja\Marcin Brzezinski\" + catia.ActiveDocument.Product.Name + ".xlsx"
    Set ExcelSheet = Nothing ' Release the object variable.

End Sub

Function IsPart(objCurrentProduct As Product) As Boolean

    Dim oTestPart As PartDocument
   
    Set oTestPart = Nothing
   
    On Error Resume Next
     
        Set oTestPart = catia.Documents.Item(objCurrentProduct.PartNumber & ".CATPart")

        If Not oTestPart Is Nothing Then
            IsPart = True
        Else
            IsPart = False
        End If
        
End Function




'pomiar w kontekscie
'http://www.coe.org/p/fo/et/thread=24453


Sub CATMain()
Dim catia As Application
Set catia = GetObject(, "catia.application")

Dim proddoc As ProductDocument
Set proddoc = catia.ActiveDocument

Dim main_prod As Product
Set main_prod = proddoc.Product

Dim main_prods As Products
Set main_prods = main_prod.Products

Dim prod1 'As Product
Set prod1 = main_prods.Item(1)

Dim prods As Products
Set prods = prod1.Products

Dim ClampPart As Part
Set ClampPart = prods.Item(1).ReferenceProduct.Parent.Part

Dim ClampLocationPoint 'As Point
Set ClampLocationPoint = ClampPart.FindObjectByName("Point1")

' create reference to a point on the assembly level
Dim refCLP As Reference
'OLD CODE: Set refCLP = ClampPart.CreateReferenceFromObject(ClampLocationPoint)
Set refCLP = main_prod.CreateReferenceFromName(main_prod.PartNumber & "/" & prod1.Name & "/" & prods.Item(1).Name & "/!Point1")

Dim TheSPAWorkbench As Workbench
Set TheSPAWorkbench = catia.ActiveDocument.GetWorkbench("SPAWorkbench")

Dim TheMeasurable 'As Measurable
Dim Coordinates(8)
Dim min_dist As Double
Dim MainAssyPart As Part
Set MainAssyPart = main_prods.Item(2).ReferenceProduct.Parent.Part

Dim AssyPartOrigin
Set AssyPartOrigin = MainAssyPart.FindObjectByName("OPoint")

' create reference to origin point (on the assembly level)
Dim refAPO As Reference
Set refAPO = main_prod.CreateReferenceFromName(main_prod.PartNumber & "/" & main_prods.Item(2).Name & "/!OPoint")

'OLD CODE: Dim refAxisOrigin As Reference
'OLD CODE: Set refAxisOrigin = MainAssyPart.CreateReferenceFromObject(AssyPartOrigin)
'OLD CODE: Set TheMeasurable = TheSPAWorkbench.GetMeasurable(ClampLocationPoint)

'OLD CODE: TheMeasurable.GetMinimumDistancePoints refAxisOrigin, Coordinates
Set TheMeasurable = TheSPAWorkbench.GetMeasurable(refAPO)

' measure distance between two points (from AssyPartOrigin to ClampLocationPoint)
Dim dDistance ' as Double
dDistance = TheMeasurable.GetMinimumDistance(refCLP)

Debug.Print dDistance

End Sub

'As Measurable


'Names
Sub CATMain()
    
    GetNextNode CATIA.ActiveDocument.Product
  
End Sub

 
Sub GetNextNode(oCurrentProduct As Product)

    Dim oCurrentTreeNode As Product
    Dim I As Integer
    Dim shift As Integer
    
    Dim ExcelSheet As Object
    Set ExcelSheet = CreateObject("Excel.Sheet")
    
    ExcelSheet.Application.Cells(1, 1).Value = "Part"
    ExcelSheet.Application.Cells(1, 2).Value = "Name"
    ExcelSheet.Application.Cells(1, 3).Value = "Description"
    ExcelSheet.Application.Cells(1, 4).Value = "Definition"
    ExcelSheet.Application.Cells(1, 5).Value = "DescRef"

    ' Loop through every tree node for the current product
    shift = 0
    For I = 1 To oCurrentProduct.Products.Count
        Set oCurrentTreeNode = oCurrentProduct.Products.Item(I)
        'MsgBox i
        ' Determine if the current node is a part, product or component
        If IsPart(oCurrentTreeNode) = True Then
            'MsgBox oCurrentTreeNode.PartNumber & " is a part"
            ExcelSheet.Application.Cells(I + 1 - shift, 1).Value = oCurrentTreeNode.PartNumber
            ExcelSheet.Application.Cells(I + 1 - shift, 2).Value = oCurrentTreeNode.Name
            ExcelSheet.Application.Cells(I + 1 - shift, 3).Value = " "
            ExcelSheet.Application.Cells(I + 1 - shift, 4).Value = " "
            ExcelSheet.Application.Cells(I + 1 - shift, 5).Value = " "
            
        Else
            'MsgBox oCurrentTreeNode.PartNumber & " is a component"
            shift = shift + 1
        End If
       
        ' if sub-nodes exist below the current tree node, call the sub recursively
        If oCurrentTreeNode.Products.Count > 0 Then
            GetNextNode oCurrentTreeNode
        End If
     
   Next
   
   ExcelSheet.Application.Visible = True
   MsgBox "Czy wartosci sa wpisane?"

   shift = 0
    For I = 1 To oCurrentProduct.Products.Count
        Set oCurrentTreeNode = oCurrentProduct.Products.Item(I)

        ' Determine if the current node is a part, product or component
        If IsPart(oCurrentTreeNode) = True Then
            'MsgBox oCurrentTreeNode.PartNumber & " is a part"
            oCurrentTreeNode.PartNumber = ExcelSheet.Application.Cells(I + 1 - shift, 1).Value
            oCurrentTreeNode.Name = ExcelSheet.Application.Cells(I + 1 - shift, 2).Value
            oCurrentTreeNode.DescriptionInst = ExcelSheet.Application.Cells(I + 1 - shift, 3).Value
            oCurrentTreeNode.Definition = ExcelSheet.Application.Cells(I + 1 - shift, 4).Value
            oCurrentTreeNode.DescriptionRef = ExcelSheet.Application.Cells(I + 1 - shift, 5).Value
           
        Else
            'MsgBox oCurrentTreeNode.PartNumber & " is a component"
            shift = shift + 1
        End If
       
       
        ' if sub-nodes exist below the current tree node, call the sub recursively
        If oCurrentTreeNode.Products.Count > 0 Then
            GetNextNode oCurrentTreeNode
        End If
     
   Next
   
    ExcelSheet.SaveAs "\\10.200.11.150\OBR-Konstrukcja\Marcin Brzezinski\NamesTable.xlsx" ' Save the sheet to directory.
    Set ExcelSheet = Nothing ' Release the object variable.

End Sub

Function IsPart(objCurrentProduct As Product) As Boolean

    Dim oTestPart As PartDocument
   
    Set oTestPart = Nothing
   
    On Error Resume Next
     
        Set oTestPart = CATIA.Documents.Item(objCurrentProduct.PartNumber & ".CATPart")

        If Not oTestPart Is Nothing Then
            IsPart = True
        Else
            IsPart = False
        End If
        
End Function

'namesTable
Sub CATMain()

'Use InputBox instead of using textbox in Visaul Basic

Dim ElementName as string
ElementName = InputBox("Please eneter element name")


Dim documents1 As Documents
Set documents1 = CATIA.Documents

Dim Selection As Selection
Set Selection = CATIA.Activedocument.selection

Dim ElementsArray(0)
ElementsArray(0) = "AnyObject"
Dim Status As String
Status = Selection.SelectElement3(ElementsArray, "Select Elements for name change", False, CATMultiSelectionMode.CATMultiSelTriggWhenUserValidatesSelection, False)

Dim i As Single
For i = 1 To Selection.Count
Selection.Item(i).Value.name = ElementName & "." & i
Selection.Item(i).Value.DescriptionInst = ElementName & "." & i
Selection.Item(i).Value.PartNumber = ElementName & "." & i
Selection.Item(i).Value.Definition = ElementName & "." & i
Selection.Item(i).Value.DescriptionRef = ElementName & "." & i

Next

End Sub


'titleblock
﻿'COPYRIGHT DASSAULT SYSTEMES 2001

' ****************************************************************************
' Purpose:       To draw a Frame and TitleBlock
'
' Assumptions:   A Drafting document should be active
'
' Author:        Tomasz Godlewski
'
' Languages:     VBScript
' Version:       V5R13
' Reg. Settings: English (United States)
' ****************************************************************************

Public DrwDocument   As DrawingDocument
Public DrwSheets     As DrawingSheets
Public DrwSheet      As DrawingSheet
Public DrwView       As DrawingView
Public DrwTexts      As DrawingTexts
Public Text          As DrawingText
Public Fact          As Factory2D
Public Point         As Point2D
Public Line          As Line2D
Public Cicle         As Circle2D
Public Selection     As Selection
Public GeomElems     As GeometricElements
Public Height        As Double            'Sheet height
Public Width         As Double            'Sheet width
Public Offset        As Double            'Distance between the sheet edges and the frame borders
Public OH            As Double            'Horizontal origin for drawing the titleblock
Public OV            As Double            'Vertical   origin for drawing the titleblock
Public Col(12)        As Double            'Columns coordinates
Public Row(12)        As Double            'Rows    coordinates
Public colRev(5)     As Double            'Columns coordinates of revision block
Public TranslationX  As Double            'Horizontal translation to operate when changing standard
Public TranslationY  As Double            'Vertical   translation to operate when changing standard
Public displayFormat As String            'Sheet format according to standard
Public sheetFormat   As catPaperSize      'Sheet format as integer value

Public TabSpawanieXPos As Double
Public TabSpawanieYPos As Double

Const mm           = 1
Const Inch         = 254
Const RulerLength  = 200
Const MacroID      = "Drawing_Titleblock_Ster"
Const RevRowHeight = 9


Sub CATMain()
  CATInit
  On Error Resume Next
    name = DrwTexts.GetItem("Reference_" + MacroID).Name
  If Err.Number <> 0 Then
    Err.Clear
    name = "none"
  End If
  On Error Goto 0
  If (name = "none") Then
    CATDrw_UtworzTabelke
  Else
    CATDrw_Zmien_rozmiar
   CATDrw_Uaktualnij
  End If
End Sub

Sub CATDrw_UtworzTabelke()
  Dim TextToFill_1 As DrawingText
  '-------------------------------------------------------------------------------
  'How to create the FTB
  '-------------------------------------------------------------------------------
  CATInit       'To init public variables & work in the background view
  If CATCheckRef(1) Then Exit Sub 'To check whether a FTB exists already in the sheet
  CATStandard   'To compute standard sizes
  CATReference  'To place on the drawing a reference point
  CATFrame      'To draw the frame
  CATTitleBlock 'To draw the TitleBlock and fill in it

  Set TextToFill_1 = DrwTexts.GetItem("TitleBlock_Text_Drawn_1")
  TextToFill_1.Text = "J. Kaleta"

End Sub

Sub CATDrw_UtworzKomentarz_Spawanie()
  CATInit
  CATStandard
  Set Text = DrwTexts.Add("EN 15085-CL3/CPD/CT4" & vbcrlf & "Grupa materiałowa XXX wg CR ISO 15608",TabSpawanieXPos, TabSpawanieYPos)
  Text.Name = "Tekst dolny Spawanie - v1"
  Text.SetFontName      0, 0, "Monospac821 BT"
  Text.SetFontSize 0, 0, 2.5
End Sub

Sub CATDrw_SkasujTabelke()
  '-------------------------------------------------------------------------------
  'How to delete the FTB
  '-------------------------------------------------------------------------------
  CATInit
  If CATCheckRef(0) Then Exit Sub

  CATRemoveAll


End Sub

Sub CATDrw_Zmien_rozmiar()
  '-------------------------------------------------------------------------------
  'How to resize the FTB
  '-------------------------------------------------------------------------------
  CATInit
  If CATCheckRef(0) Then Exit Sub
  CATStandard
  CATMoveReference
  If TranslationX <> 0 Or TranslationY <> 0 Then
    CATRemoveFrame
    CATRemovePicture
    CATMoveTitleBlock
	CATTitleBlockStandard
    CATFrame
    CATLinks
  End If
End Sub

Sub CATDrw_Uaktualnij()
  '-------------------------------------------------------------------------------
  'How to update the FTB
  '-------------------------------------------------------------------------------
  CATInit
  If CATCheckRef(0) Then Exit Sub
  CATStandard
  CATTitleBlockStandard
  CATLinks
End Sub

Sub CATDrw_Sprawdzil()
  '-------------------------------------------------------------------------------
  'How to update a bit more the FTB
  '-------------------------------------------------------------------------------
  CATInit
  If CATCheckRef(0) Then Exit Sub
  CATFillField "TitleBlock_Text_Check_1", "TitleBlock_Text_CDate_1", "sprawdzony"
End Sub

Sub CATDrw_Zatwierdzil()
  '-------------------------------------------------------------------------------
  'How to update a bit more the FTB
  '-------------------------------------------------------------------------------
  CATInit
  If CATCheckRef(0) Then Exit Sub
  CATFillField "TitleBlock_Text_Appd_1", "TitleBlock_Text_ADate_1", "zatwierdzony"
End Sub

Sub CATDrw_Konstruowal()
  '-------------------------------------------------------------------------------
  'How to update a bit more the FTB
  '-------------------------------------------------------------------------------
  CATInit
  If CATCheckRef(0) Then Exit Sub
  CATFillField "TitleBlock_Text_Design_1", "TitleBlock_Text_DeDate_1", "zaprojektowany"
End Sub

Sub CATDrw_Kreslil()
  '-------------------------------------------------------------------------------
  'How to update a bit more the FTB
  '-------------------------------------------------------------------------------
  CATInit
  If CATCheckRef(0) Then Exit Sub
  CATFillField "TitleBlock_Text_Drawn_1", "TitleBlock_Text_DrDate_1", "kreslony"
End Sub

Sub CATDrw_DodajTabeleTolerancji()
  '-------------------------------------------------------------------------------
  'How to create or modify a revison block
  '-------------------------------------------------------------------------------
  Dim X As double
  Dim Y As double
  CATInit
  CATRevPos revision, X, Y
  CATRevisionBlock revision, X, Y
End Sub

Sub CATDrw_ZmienRewizje()
  '-------------------------------------------------------------------------------
  'How to create or modify a revison block
  '-------------------------------------------------------------------------------
  Dim X As double
  Dim Y As double
  CATInit
  If CATCheckRef(0) Then Exit Sub
  revision = CATCheckRev
  On Error Resume Next
    DrwTexts.GetItem("TitleBlock_Text_Rev_1").Text = Chr(65 + revision)
  If Err.Number <> 0 Then
    Err.Clear
  End If
  On Error Goto 0
  End Sub

Sub CATInit()
  '-------------------------------------------------------------------------------
  'How to init the dialog and create main objects
  '-------------------------------------------------------------------------------
  Set DrwDocument = CATIA.ActiveDocument
  Set DrwSheets   = DrwDocument.Sheets
  Set Selection   = DrwDocument.Selection
  Set DrwSheet    = DrwSheets.ActiveSheet
  Set DrwView     = DrwSheet.Views.ActiveView
  Set DrwTexts    = DrwView.Texts
  Set Fact        = DrwView.Factory2D
  Set GeomElems   = DrwView.GeometricElements

  Col(1) = -180*mm
  Col(2) = -176*mm
  Col(3) = -157*mm
  Col(4) = -127*mm
  Col(5) = -101*mm
  Col(6) = - 86*mm
  Col(7) = - 83*mm
  Col(8) = - 76*mm
  Col(9) = - 43*mm
  Col(10) = -18*mm
  Col(11) = -10*mm


  Row(1) = + 12.5*mm
  Row(2) = + 15.5*mm
  Row(3) = + 25*mm
  Row(4) = + 30*mm
  Row(5) = + 32*mm
  Row(6) = + 35*mm
  Row(7) = + 39*mm
  Row(8) = + 40*mm
  Row(9) = + 46*mm
  Row(10) = + 54*mm
  Row(11) = - 27*mm
  Row(12) = - 36*mm

End Sub

Sub CATStandard()
  '-------------------------------------------------------------------------------
  'How to compute standard values
  '-------------------------------------------------------------------------------
  Height      = DrwSheet.GetPaperHeight
  Width       = DrwSheet.GetPaperWidth
  sheetFormat = DrwSheet.PaperSize

  TabSpawanieXPos = Width - 164
  TabSpawanieYPos = 75

  Offset = 7.*mm 'Offset default value = 10.
  If (sheetFormat = CatPaperA0 Or sheetFormat = CatPaperA1 Or sheetFormat = CatPaperUser And _
      (DrwSheet.GetPaperWidth > 594.*mm Or DrwSheet.GetPaperHeight > 594.*mm)) Then
    Offset = 15.*mm
  End If

  OH = Width - Offset
  OV = Offset
  HO = Height - Offset - 87

  documentStd = DrwDocument.Standard
  If (documentStd = catISO) Then
    If sheetFormat = 13 Then
      displayFormat = "USER"
    Else
      displayFormat = "A" + CStr(sheetFormat - 2)
    End IF
  Else
    Select Case sheetFormat
      Case 0
        displayFormat = "Letter"
      Case 1
        displayFormat = "Legal"
      Case 7
        displayFormat = "A"
      Case 8
        displayFormat = "B"
      Case 9
        displayFormat = "C"
      Case 10
        displayFormat = "D"
      Case 11
        displayFormat = "E"
      Case 12
        displayFormat = "F"
      Case 13
        displayFormat = "J"
    End Select
  End If

End Sub

Sub CATReference()
  '-------------------------------------------------------------------------------
  'How to create a reference text
  '-------------------------------------------------------------------------------
  Set Text = DrwTexts.Add("", Width - Offset, Offset)
  Text.Name = "Reference_" + MacroID
End Sub

Function CATCheckRef(Mode As Integer) As Integer
  '-------------------------------------------------------------------------------
  'How to check that the called macro is the right one
  '-------------------------------------------------------------------------------
  nbTexts = DrwTexts.Count
  i = 0
  notFound = 0
  While (notFound = 0 And i<nbTexts)
    i = i + 1
    Set Text = DrwTexts.Item(i)
    WholeName = Text.Name
    leftText = Left(WholeName, 10)
    If (leftText = "Reference_") Then
    notFound = 1
    refText = "Reference_" + MacroID
    If (Mode = 1) Then
      MsgBox "Frame and Titleblock already created!"
      CATCheckRef = 1
      Exit Function
    ElseIf (Text.Name <> refText) Then
      MsgBox "Frame and Titleblock created using another style:" + Chr(10) + "        " + MacroID
      CATCheckRef = 1
      Exit Function
    End If
    End If
  Wend
  CATCheckRef = 0

End Function

Function CATCheckRev() As Integer
  '-------------------------------------------------------------------------------
  'How to check that a revision block alredy exists
  '-------------------------------------------------------------------------------
  CATCheckRev = 0
  nbTexts = DrwTexts.Count
  i = 0
  While (i<nbTexts)
    current = 0
    i = i + 1
    Set Text = DrwTexts.Item(i)
    WholeName = Text.Name
    leftText = Left(WholeName, 23)
    If (leftText = "RevisionBlock_Text_Rev_") Then
      CATCheckRev = CATCheckRev + 1
    End If
  Wend

End Function

Sub CATFrame()
  '-------------------------------------------------------------------------------
  'How to create the Frame
  '-------------------------------------------------------------------------------
  Dim Cst_1   As Double  'Length (in cm) between 2 horinzontal marks
  Dim Cst_2   As Double  'Length (in cm) between 2 vertical marks
  Dim Nb_CM_H As Integer 'Number/2 of horizontal centring marks
  Dim Nb_CM_V As Integer 'Number/2 of vertical centring marks
  Dim Ruler   As Integer 'Ruler length (in cm)

  CATFrameStandard     Nb_CM_H, Nb_CM_V, Ruler, Cst_1, Cst_2
  CATFrameBorder
  CATFrameCentringMark Nb_CM_H, Nb_CM_V, Ruler, Cst_1, Cst_2
  CATFrameText         Nb_CM_H, Nb_CM_V, Ruler, Cst_1, Cst_2
'  CATFrameRuler        Ruler, Cst_1
  CATPicture
End Sub

Sub CATFrameStandard(Nb_CM_H As Integer, Nb_CM_V As Integer, Ruler As Integer, Cst_1 As Double, Cst_2 As Double)
  '-------------------------------------------------------------------------------
  'How to compute standard values
  '-------------------------------------------------------------------------------
  Cst_1 = 74.2*mm '297, 594, 1189 are multiples of 74.2
  Cst_2 = 52.5*mm '210, 420, 841  are multiples of 52.2
  If DrwSheet.Orientation = CatPaperPortrait And _
     (sheetFormat = CatPaperA0 Or _
      sheetFormat = CatPaperA2 Or _
      sheetFormat = CatPaperA4) Or _
      DrwSheet.Orientation = CatPaperLandscape And _
     (sheetFormat = CatPaperA1 Or _
      sheetFormat = CatPaperA3) Then
    Cst_1 = 52.5*mm
    Cst_2 = 74.2*mm
  End If

  Nb_CM_H = CInt(.5 * Width / Cst_1)
  Nb_CM_V = CInt(.5 * Height / Cst_2)

  Ruler   = CInt((Nb_CM_H - 1) * Cst_1 / 50) * 100 'here is computed the maximum ruler length
  If RulerLength < Ruler Then
    Ruler = RulerLength
  End If
End Sub

Sub CATFrameBorder()
  '-------------------------------------------------------------------------------
  'How to draw the frame border
  '-------------------------------------------------------------------------------
  On Error Resume Next
    Set Line = Fact.CreateLine(OV, OV             , OH, OV             )
    Line.Name = "Frame_Border_Bottom"
    Set Line = Fact.CreateLine(OH, OV             , OH, Height - Offset)
    Line.Name = "Frame_Border_Left"
    Set Line = Fact.CreateLine(OH, Height - Offset, OV, Height - Offset)
    Line.Name = "Frame_Border_Top"
    Set Line = Fact.CreateLine(OV, Height - Offset, OV, OV             )
    Line.Name = "Frame_Border_Right"
  If Err.Number <> 0 Then
    Err.Clear
  End If
  On Error Goto 0
End Sub

Sub CATFrameCentringMark(Nb_CM_H As Integer, Nb_CM_V As Integer, Ruler As Integer, Cst_1 As Double, Cst_2 As Double)
  '-------------------------------------------------------------------------------
  'How to draw the centring marks
  '-------------------------------------------------------------------------------
  On Error Resume Next
    Set Line = Fact.CreateLine(.5 * Width    , Height - Offset, .5 * Width, Height     )
    Line.Name = "Frame_CentringMark_Top"
    Set Line = Fact.CreateLine(.5 * Width    , OV             , .5 * Width, .0         )
    Line.Name = "Frame_CentringMark_Bottom"
    Set Line = Fact.CreateLine(OV            , .5 * Height    , .0        , .5 * Height)
    Line.Name = "Frame_CentringMark_Left"
    Set Line = Fact.CreateLine(Width - Offset, .5 * Height    , Width     , .5 * Height)
    Line.Name = "Frame_CentringMark_Right"


	For i = 1 To Nb_CM_H
      If (i * Cst_1 < .5 * Width - 1.) Then
        Set Line  = Fact.CreateLine(.5 * Width + i * Cst_1, OV, .5 * Width + i * Cst_1, .25 * Offset)
        Line.Name = "Frame_CentringMark_Bottom"
        Set Line  = Fact.CreateLine(.5 * Width - i * Cst_1, OV, .5 * Width - i * Cst_1, .25 * Offset)
        Line.Name = "Frame_CentringMark_Bottom"
      End If
    Next

    For i = 1 To Nb_CM_H
      If (i * Cst_1 < .5 * Width - 1.) Then
        Set Line  = Fact.CreateLine(.5 * Width + i * Cst_1, Height - Offset, .5 * Width + i * Cst_1, Height - .25 * Offset)
        Line.Name = "Frame_CentringMark_Top"
        Set Line  = Fact.CreateLine(.5 * Width - i * Cst_1, Height - Offset, .5 * Width - i * Cst_1, Height - .25 * Offset)
        Line.Name = "Frame_CentringMark_Top"
      End If
    Next

    For i = 1 To Nb_CM_V
      If (i * Cst_2 < .5 * Height - 1.) Then
        Set Line  = Fact.CreateLine(OV, .5 * Height + i * Cst_2, .25 * Offset        , .5 * Height + i * Cst_2)
        Line.Name = "Frame_CentringMark_Left"
        Set Line  = Fact.CreateLine(OV, .5 * Height - i * Cst_2, .25 * Offset        , .5 * Height - i * Cst_2)
        Line.Name = "Frame_CentringMark_Left"
        Set Line  = Fact.CreateLine(OH, .5 * Height + i * Cst_2, Width - .25 * Offset, .5 * Height + i * Cst_2)
        Line.Name = "Frame_CentringMark_Right"
        Set Line  = Fact.CreateLine(OH, .5 * Height - i * Cst_2, Width - .25 * Offset, .5 * Height - i * Cst_2)
        Line.Name = "Frame_CentringMark_Right"
      End If
    Next

  If Err.Number <> 0 Then
    Err.Clear
  End If
  On Error Goto 0
End Sub

Sub CATFrameText(Nb_CM_H As Integer, Nb_CM_V As Integer, Ruler As Integer, Cst_1 As Double, Cst_2 As Double)
  '-------------------------------------------------------------------------------
  'How to create coordinates
  '-------------------------------------------------------------------------------
  On Error Resume Next

    For i = 1 To Nb_CM_H
      Set Text = DrwTexts.Add( CStr(Nb_CM_H + i), .5 * Width - (i - .5) * Cst_1, .5 * Offset)
      CATFormatFText "Frame_Text_Bottom_1", 0
      Set Text = DrwTexts.Add(CStr(Nb_CM_H - i + 1), .5 * Width + (i - .5) * Cst_1, .5 * Offset)
      CATFormatFText "Frame_Text_Bottom_2", 0
    Next

    For i = 1 To Nb_CM_H
      Set Text = DrwTexts.Add(CStr(Nb_CM_H + i)    , .5 * Width - (i - .5) * Cst_1, Height - .5 * Offset)
      CATFormatFText "Frame_Text_Top_1", -90
      Set Text = DrwTexts.Add(CStr(Nb_CM_H - i + 1), .5 * Width + (i - .5) * Cst_1, Height - .5 * Offset)
      CATFormatFText "Frame_Text_Top_2", -90
    Next

    For i = 1 To Nb_CM_V
      Set Text = DrwTexts.Add(Chr(65 + Nb_CM_V - i)     , .5 * Offset        , .5 * Height - (i - .5) * Cst_2)
      CATFormatFText "Frame_Text_Left_1", -90
      Set Text = DrwTexts.Add(Chr(64 + Nb_CM_V + i) , .5 * Offset        , .5 * Height + (i - .5) * Cst_2)
      CATFormatFText "Frame_Text_Left_2", -90
      Set Text = DrwTexts.Add(Chr(65 + Nb_CM_V - i), Width - .5 * Offset, .5 * Height - (i - .5) * Cst_2)
      CATFormatFText "Frame_Text_Right_1", 0
      Set Text = DrwTexts.Add(Chr(64 + Nb_CM_V + i), Width - .5 * Offset, .5 * Height + (i - .5) * Cst_2)
      CATFormatFText "Frame_Text_Right_2", 0
    Next

  If Err.Number <> 0 Then
    Err.Clear
  End If
  On Error Goto 0
End Sub


Sub CATTitleBlock()
  '-------------------------------------------------------------------------------
  'How to create the TitleBlock
  '-------------------------------------------------------------------------------
  CATTitleBlockFrame    'To draw the geometry
  CATTitleBlockText     'To fill in the title block
  CATTitleBlockStandard
End Sub

Sub CATTitleBlockFrame()
  '-------------------------------------------------------------------------------
  'Tworzenie ramki rysunkowej STER
  '-------------------------------------------------------------------------------
  On Error Resume Next

'Obwodka ramki

    Set Line      = Fact.CreateLine(OH + Col(1), OV         , OH         , OV         )
    Line.Name     = "TitleBlock_Line_Bottom"
    Set Line      = Fact.CreateLine(OH + Col(1), OV         , OH + Col(1), OV + Row(10))
    Line.Name     = "TitleBlock_Line_Left"
    Set Line      = Fact.CreateLine(OH + Col(1), OV + Row(10), OH         , OV + Row(10))
    Line.Name     = "TitleBlock_Line_Top"
    Set Line      = Fact.CreateLine(OH         , OV + Row(10), OH         , OV         )
    Line.Name     = "TitleBlock_Line_Right"

'Linie wewnetrzen

    Set Line      = Fact.CreateLine(OH + Col(2), OV + Row (9), OH + Col(11), OV + Row(9))
    Line.Name     = "TitleBlock_Line_Row_1"
    Set Line      = Fact.CreateLine(OH + Col(3), OV + Row(8), OH + Col(8), OV + Row(8))
    Line.Name     = "TitleBlock_Line_Row_2"
    Set Line      = Fact.CreateLine(OH + Col(8), OV + Row(7), OH         , OV + Row(7))
    Line.Name     = "TitleBlock_Line_Row_3"
    Set Line      = Fact.CreateLine(OH + Col(3), OV + Row(6), OH + Col(8), OV + Row(6))
    Line.Name     = "TitleBlock_Line_Row_4"
    Set Line      = Fact.CreateLine(OH + Col(8), OV + Row(5), OH + Col(11), OV + Row(5))
    Line.Name     = "TitleBlock_Line_Row_5"
   Set Line      = Fact.CreateLine(OH + Col(2), OV + Row(4), OH + Col(8)        , OV + Row(4))
    Line.Name     = "TitleBlock_Line_Row_6"
    Set Line      = Fact.CreateLine(OH + Col(3), OV + Row(3), OH         , OV + Row(3))
    Line.Name     = "TitleBlock_Line_Row_7"
    Set Line      = Fact.CreateLine(OH + Col(2), OV + Row(2), OH + Col(3), OV + Row(2))
    Line.Name     = "TitleBlock_Line_Row_8"
    Set Line      = Fact.CreateLine(OH + Col(3), OV + Row(1), OH         , OV + Row(1))
    Line.Name     = "TitleBlock_Line_Row_9"


    Set Line      = Fact.CreateLine(OH + Col(2), OV         , OH + Col(2), OV + Row(10))
    Line.Name     = "TitleBlock_Line_Column_1"
    Set Line      = Fact.CreateLine(OH + Col(3), OV         , OH + Col(3), OV + Row(10))
    Line.Name     = "TitleBlock_Line_Column_2"
    Set Line      = Fact.CreateLine(OH + Col(4), OV + Row(3), OH + Col(4), OV + Row(10))
    Line.Name     = "TitleBlock_Line_Column_3"
    Set Line      = Fact.CreateLine(OH + Col(5), OV + Row(3), OH + Col(5), OV + Row(10))
    Line.Name     = "TitleBlock_Line_Column_4"
    Set Line      = Fact.CreateLine(OH + Col(6), OV         , OH + Col(6), OV + Row(1))
    Line.Name     = "TitleBlock_Line_Column_5"
    Set Line      = Fact.CreateLine(OH + Col(7), OV + Row(3) , OH + Col(7), OV + Row(10))
    Line.Name     = "TitleBlock_Line_Column_6"
    Set Line      = Fact.CreateLine(OH + Col(8), OV + Row(3) , OH + Col(8), OV + Row(10))
    Line.Name     = "TitleBlock_Line_Column_7"
    Set Line      = Fact.CreateLine(OH + Col(9), OV + Row(3) , OH + Col(9), OV + Row(10))
    Line.Name     = "TitleBlock_Line_Column_8"
    Set Line      = Fact.CreateLine(OH + Col(10), OV , OH + Col(10), OV + Row(1))
    Line.Name     = "TitleBlock_Line_Column_9"
    Set Line      = Fact.CreateLine(OH + Col(11), OV + Row(3) , OH + Col(11), OV + Row(10))
    Line.Name     = "TitleBlock_Line_Column_10"


    'Set Line      = Fact.CreateLine((Height-Offset) + Col(9) - 87 ,Height - Offset         ,   (Height-Offset) + Col(9) -87     , (Height-Offset) + Row(12))
    'Line.Name     = "TitleBlock_Line_Row_10"
    'Set Line      = Fact.CreateLine((Height-Offset) + Col(10) - 87 ,Height - Offset        ,   (Height-Offset) + Col(10) -87       , (Height-Offset) + Row(12))
    'Line.Name     = "TitleBlock_Line_Row_11"

   'Set Line      = Fact.CreateLine((Height-Offset) + Col(9) -87 ,(Height - Offset) + Row(12)         ,   (Height-Offset)  -87    , (Height-Offset) + Row(12))
    'Line.Name     = "TitleBlock_Line_Column_9"
    'Set Line      = Fact.CreateLine((Height-Offset) + Col(9) - 87 ,(Height - Offset) + Row(11)         ,   (Height-Offset)-87     , (Height-Offset) + Row(11))
    'Line.Name     = "TitleBlock_Line_Column_10"
    'Set Line      = Fact.CreateLine((Height-Offset) + Col(9) - 87 ,(Height - Offset) + Row(10)         ,   (Height-Offset)  -87    , (Height-Offset) + Row(10))
    'Line.Name     = "TitleBlock_Line_Column_11"
    'Set Line      = Fact.CreateLine((Height-Offset) + Col(9) - 87 ,(Height - Offset) + Row(9)         ,   (Height-Offset) -87     , (Height-Offset) + Row(9))
    'Line.Name     = "TitleBlock_Line_Column_12"


  If Err.Number <> 0 Then
    Err.Clear
  End If
  On Error Goto 0
End Sub

Sub CATTitleBlockStandard()
  '-------------------------------------------------------------------------------
  'How to create the standard representation
  '-------------------------------------------------------------------------------
  Dim R1   As Double
  Dim R2   As Double
  Dim X(5) As Double
  Dim Y(7) As Double

  R1   = 1*mm
  R2   = 2.5*mm
  X(1) = OH   + Col(2)+1.5*mm
  X(2) = X(1) + 1.5*mm
  X(3) = X(1) + 9.5*mm
  X(4) = X(1) + 13*mm
  X(5) = X(1) + 16*mm
  Y(1) = OV   + (Row(9)+Row(10))/2.
  Y(2) = Y(1) + R1
  Y(3) = Y(1) + R2
  Y(4) = Y(1) + 3*mm
  Y(5) = Y(1) - R1
  Y(6) = Y(1) - R2
  Y(7) = Y(1) - 3*mm

  If sheetProjMethod  <> CatFirstAngle Then
    Xtmp = X(2)
    X(2) = X(1) + X(5) - X(3)
    X(3) = X(1) + X(5) - Xtmp
    X(4) = X(1) + X(5) - X(4)
  End If

  On Error Resume Next
    Set Line   = Fact.CreateLine(X(1), Y(1), X(5), Y(1))
    Line.Name   = "TitleBlock_Standard_Line_Axis_1"
    Set Line   = Fact.CreateLine(X(4), Y(7), X(4), Y(4))
    Line.Name   = "TitleBlock_Standard_Line_Axis_2"
    Set Line   = Fact.CreateLine(X(2), Y(5), X(2), Y(2))
    Line.Name   = "TitleBlock_Standard_Line_1"
    Set Line   = Fact.CreateLine(X(2), Y(2), X(3), Y(3))
    Line.Name   = "TitleBlock_Standard_Line_2"
    Set Line   = Fact.CreateLine(X(3), Y(3), X(3), Y(6))
    Line.Name   = "TitleBlock_Standard_Line_3"
    Set Line   = Fact.CreateLine(X(3), Y(6), X(2), Y(5))
    Line.Name   = "TitleBlock_Standard_Line_4"
    Set Circle = Fact.CreateClosedCircle(X(4), Y(1), R1)
    Circle.Name = "TitleBlock_Standard_Circle_1"
    Set Circle = Fact.CreateClosedCircle(X(4), Y(1), R2)
    Circle.Name = "TitleBlock_Standard_Circle_2"
  If Err.Number <> 0 Then
    Err.Clear
  End If
  On Error Goto 0

End Sub

Sub CATTitleBlockText()
  '-------------------------------------------------------------------------------
  'How to fill in the title block
  '-------------------------------------------------------------------------------
  Text_01 = "Konstruowal/Desig."
  Text_02 = ""
  Text_03 = "Data/Date"
  Text_04 = ""
  Text_05 = "Sprawdzil/Chk'd"
  Text_06 = "Kreslil/Drawn"
  Text_07 = ""
  Text_08 = "Podzialka/ Scale"
  Text_09 = "Masa (kg)/Weight"
  Text_10 = "Arkusz"
  Text_11 = "Format/ Paper size"
  Text_12 = "" ' Paper Format
  Text_13 = "NAZWA RYSUNKU/DRAWING NAME"
  Text_14 = "Rewizja/Rev"
  Text_15 = "A"
  Text_16 = "NR RYSUNKU/DRAWING NUMBER"
  Text_17 = "Wszelkie prawa zastrzezone/All rights reserved"
  Text_18 = " "
  Text_19 = ""

'--------------------------------------------------------------------------
' SÄąÂOWO NAZWISKO
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add("Nazwisko/Name", OH + .5*(Col(4) +Col(5))     , OV + Row(10)-1.5            )
  CATFormatTBText "TitleBlock_Text_Name"      , catTopCenter   ,   2 , 24,0
'--------------------------------------------------------------------------
' SÄąÂOWO DATA
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_03, OH + .5*Col(2)-4       , OV + Row(10) - 1.5           )
  CATFormatTBText "TitleBlock_Text_DeDate"      , catTopCenter  ,    2 ,16,0
'--------------------------------------------------------------------------
' SÄąÂOWO PODPIS
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add("Podpis Sign.", OH + .5*Col(3) -0.5      , OV + Row(10) - 1           )
  CATFormatTBText "TitleBlock_Text_DeDate"      , catTopCenter  ,    1.4, 8,0
'--------------------------------------------------------------------------
' SÄąÂOWO PROJEKTOWAÄąÂ
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_01, OH + Col(3) + 0.5     , OV + Row(9)-1            )
  CATFormatTBText "TitleBlock_Text_Design"      , catTopLeft   ,   1.9 ,30,0
'--------------------------------------------------------------------------
' WSTAW KTO PROJEKTOWAL
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add("J. Kaleta", OH + Col(4) + 0.5, OV + Row(9)-1                  )
  CATFormatTBText "TitleBlock_Text_Design_1"    , catTopLeft , 1.8 ,26,0
'--------------------------------------------------------------------------
'WSTAW DATĂ„Â PROJEKTOWANIA
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Date, OH + Col(6)-5       , OV + Row(9) - 1                     )
  CATFormatTBText "TitleBlock_Text_DeDate_1"    , catTopCenter, 1.7, 19,0
'--------------------------------------------------------------------------
' SÄąÂOWO KREÄąĹˇLIÄąÂ
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_06, OH + Col(3) + 0.5      , OV + Row(8)-1            )
  CATFormatTBText "TitleBlock_Text_Drawn"       , catTopLeft   ,   2 , 30,0
'--------------------------------------------------------------------------
' WSTAW KTO KREÄąĹˇLIÄąÂ
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_02, OH + Col(4) +0.5       , OV + Row(8)-1            )
  CATFormatTBText "TitleBlock_Text_Drawn_1"     , catTopLeft,   1.8 , 27,0
'--------------------------------------------------------------------------
' WSTAW DATĂ„Â KREÄąĹˇLENIA
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Date   , OH + Col(6)-5   , OV + Row(8)-1          )
  CATFormatTBText "TitleBlock_Text_DrDate_1"    ,  catTopCenter , 1.7 , 19,0
'--------------------------------------------------------------------------
' SÄąÂOWO SPRAWDZIÄąÂ
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_05, OH + Col(3) + 0.5       , OV + Row(6)-1            )
  CATFormatTBText "TitleBlock_Text_Check"       , catTopLeft   ,   2, 30,0
'--------------------------------------------------------------------------
' WSTAW KTO SPRAWDZIÄąÂ
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_04, OH + Col(4) + 0.5, OV + Row(6)-1              )
  CATFormatTBText "TitleBlock_Text_Check_1"     , catTopLeft , 1.8 , 26,0
'--------------------------------------------------------------------------
' WSTAW DATĂ„Â SPRAWDZENIA
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_04,OH +Col(6)-5   , OV + Row(6) - 1            )
  CATFormatTBText "TitleBlock_Text_CDate_1"     , catTopCenter, 1.7 , 19,0
'--------------------------------------------------------------------------
' SÄąÂOWO ZATWIERDZIÄąÂ
'--------------------------------------------------------------------------
   Set Text     = DrwTexts.Add("Zatwierdzil/App'd", OH + Col(3) + 0.5      , OV + Row(4)-1            )
   CATFormatTBText "TitleBlock_Text_Appd"  , catTopLeft  ,    2 , 30,0
'--------------------------------------------------------------------------
' WSTAW KTO ZATWIERDZIÄąÂ
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_04, OH + Col(4) + 0.5, OV + Row(4)-1            )
  CATFormatTBText "TitleBlock_Text_Appd_1"     , catTopLeft,   1.8 , 25,0
'--------------------------------------------------------------------------
' WSTAW DATĂ„Â ZATWIERDZENIA
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_04, OH + Col(6)-5   , OV + Row(4) - 1            )
  CATFormatTBText "TitleBlock_Text_ADate_1"     , catTopCenter, 1.7 , 19,0
'--------------------------------------------------------------------------
' SÄąÂOWO PODZIAÄąÂKA
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_08, OH + .5*(Col(2) +Col(3))+2        , OV + Row(4) -1      )
  CATFormatTBText "TitleBlock_Text_Scale"      , catTopCenter,   1.7 , 19,0
'--------------------------------------------------------------------------
' WSTAW PODZIAÄąÂKĂ„Â
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add("",OH + Col(2)+7        , OV + Row(4) -7        )
  CATFormatTBText "TitleBlock_Text_Scale_1"    , catTopCenter,   4 , 19,0
  Text.InsertVariable 1, 0, DrwDocument.Parameters.Item("Drawing\" + DrwSheet.Name + "\ViewMakeUp.1\Scale")
'--------------------------------------------------------------------------
'SLOWO ADRES WWW
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add("www.ster.com.pl", OH + .5*(Col(2) +Col(3))+0.5        , OV + Row(5) +0.2      )
  CATFormatTBText "TitleBlock_Text_Scale"      , catTopCenter,   1.4 , 19,0
'--------------------------------------------------------------------------
' SÄąÂOWO NUMER JEDNOSTKI WYÄąÂ»SZEGO RZĂ„ÂDU
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add("Nr jedn. wyzszego rzedu Upper assm. drawing number", OH + Col(8)+.8  , OV + Row(10)-1 )
  CATFormatTBText "TitleBlock_Text_Jednostka"     , catTopLeft,    1.4 , 34,0
'--------------------------------------------------------------------------
' WSTAW  NUMER JEDNOSTKI WYÄąÂ»SZEGO RZĂ„ÂDU
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add("", OH + Col(9) + 0.6        , OV + Row(10) - 2         )
  CATFormatTBText "TitleBlock_Text_Jednostka"   , catTopLeft,   1.7 , 33,0
'--------------------------------------------------------------------------
' SÄąÂOWO NUMER NORMY
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add("Numer normy/Norm number", OH + Col(8) +0.6   , OV + Row(9) - 2         )
  CATFormatTBText "TitleBlock_Text_Norma"     , catTopLeft,    1.6 , 33,0
'--------------------------------------------------------------------------
' WSTAW NUMER NORMY
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add("PN-EN-ISO", OH + Col(9) + 0.6        , OV + Row(9) - 2         )
  CATFormatTBText "TitleBlock_Text_Norma"   , catTopLeft,   1.7 , 33,0
'--------------------------------------------------------------------------
' SÄąÂOWO WAGA [KG]
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_09, OH + Col(8) +0.6       , OV + Row(7)-2        )
  CATFormatTBText "TitleBlock_Text_Weight"     , catTopLeft,    1.6 , 30,0
'--------------------------------------------------------------------------
' WSTAW WAGĂ„Â
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_04, OH + Col(9) + 1        , OV + Row(7) - 2         )
  CATFormatTBText "TitleBlock_Text_Weight_1"   , catTopLeft,   2 , 30,0
'--------------------------------------------------------------------------
' SÄąÂOWO MATERIAÄąÂ
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add("MATERIAL / MATERIAL", OH + Col(3) + 1        , OV + Row(1)         )
  CATFormatTBText "TitleBlock_Text_Material"     , catTopLeft,    2 , 40,0
'--------------------------------------------------------------------------
' WSTAW MATERIAÄąÂ
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_04, OH + Col(3) + 2      , OV + Row(1) - 4         )
  CATFormatTBText "TitleBlock_Text_Material_1"   , catTopLeft,   3.5 , 68,0
'--------------------------------------------------------------------------
' SÄąÂOWO WYMIAR
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add("Wymiar/Main dimension", OH + Col(8)+0.6        , OV + Row(4) - 1         )
  CATFormatTBText "TitleBlock_Text_Wymiar"     , catTopLeft,   1.6 , 30,0
'--------------------------------------------------------------------------
' WSTAW WYMIAR
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_04, OH + Col(9) + 1        , OV + Row(4) - 1         )
  CATFormatTBText "TitleBlock_Text_Wymiar"   , catTopLeft,   2 , 35,0
'--------------------------------------------------------------------------
' SÄąÂOWO ARKUSZ
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_10, OH+.5*Col(11)              , OV +Row(9) +7      )
  CATFormatTBText "TitleBlock_Text_Sheet"       , catTopCenter,   1.7 , 9,0
'--------------------------------------------------------------------------
' WSTAW ARKUSZ
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_04, OH+.5*Col(11)              , OV + Row(9) -5      )
  CATFormatTBText "TitleBlock_Text_Sheet_1"     , catBottomCenter,  4 , 5,0
'--------------------------------------------------------------------------
' SÄąÂOWO ILOSC
'--------------------------------------------------------------------------
  'Set Text     = DrwTexts.Add("Sztuk", OH+.5*Col(8)              , OV +Row(7)       )
  'CATFormatTBText "TitleBlock_Text_Ilosc"       , catTopCenter,   2 , 100,0
'--------------------------------------------------------------------------
' WSTAW ILOSC
'--------------------------------------------------------------------------
  'Set Text     = DrwTexts.Add("1", OH+.5*Col(8)              , OV + Row(5)        )
  'CATFormatTBText "TitleBlock_Text_Ilosc"     , catBottomCenter,  4 , 100,0
'--------------------------------------------------------------------------
' SÄąÂOWO FORMAT
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_11, OH + .5*(Col(2) +Col(3))        , OV + Row(2)-1            )
  CATFormatTBText "TitleBlock_Text_Size"        , catTopCenter   ,   1.7 , 16,0
'--------------------------------------------------------------------------
' WSTAW FORMAT
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_12, OH + .5*(Col(2) +Col(3))        , OV + Row(2) -7            )
  CATFormatTBText "TitleBlock_Text_Size_1"      , catTopCenter, 4 , 10,0
'--------------------------------------------------------------------------
' SÄąÂOWO NAZWA RYSUNKU
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_13, OH + Col(3)+1           , OV + Row(3)            )
  CATFormatTBText "TitleBlock_Text_Number"      , catTopLeft  ,    2 , 50,0
'--------------------------------------------------------------------------
' WSTAW NUMER RYSUNKU
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_19, OH + 0.4*(Col(8)+Col(9))-2 , OV + Row(1) -4    )
  CATFormatTBText "TitleBlock_Text_Number_1"    , catTopCenter, 3.5 , 62,0
'--------------------------------------------------------------------------
' SÄąÂOWO ILOÄąĹˇĂ„â€  ARKUSZY
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add("Ilosc Ark. Sheet Qty.", OH + Col(11)+0.5   , OV + Row(7)-1 )
  CATFormatTBText "TitleBlock_Text_Arkusze"         , catTopLeft ,    1.1 , 10,0
'--------------------------------------------------------------------------
' WSTAWIENIE ILOSC ARKUSZY
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_7, OH+.4*Col(11)              , OV  + Row(3)+2          )
  CATFormatTBText "TitleBlock_Text_Arkusze_1"       , catBottomCenter, 4 , 7,0
'--------------------------------------------------------------------------
' SÄąÂOWO REWIZJA
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_14, OH + .5*Col(10)+1          , OV + Row(1)            )
  CATFormatTBText "TitleBlock_Text_Rev"         , catTopCenter ,    1.7 , 18,0
'--------------------------------------------------------------------------
' WSTAWIENIE REWIZJI
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_15, OH+ 0.6*Col(11)-3              , OV + Row(1) -4   )
  CATFormatTBText "TitleBlock_Text_Rev_1"       , catTopCenter, 4 , 7,0
'--------------------------------------------------------------------------
' SÄąÂOWO NUMER PROJEKTU
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_16, OH + Col(6) + 1           , OV + Row(2)-3            )
  CATFormatTBText "TitleBlock_Text_Title"       , catTopLeft  ,    2 , 55,0
'--------------------------------------------------------------------------
' WSTAW NAZWĂ„Â PROJEKTU
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_04, OH+1.03*Col(8)     , OV + Row(3)-5            )
  CATFormatTBText "TitleBlock_Text_Title_1"     , catTopCenter, 4 , 155,0
'--------------------------------------------------------------------------
' SÄąÂOWO Prawa
'--------------------------------------------------------------------------
  Set Text     = DrwTexts.Add(Text_17, OH+Col(1) +0.4             , OV+Row(1)-12            )
  CATFormatTBText "TitleBlock_Text_Company"     , catTopLeft, 1.3 , 57 ,90
 '--------------------------------------------------------------------------
' WSTAW KOD
'--------------------------------------------------------------------------
  'Set Text     = DrwTexts.Add(Text_18, OH+.5*(Col(6)+Col(8))  , OV +Row(3)-4           )
  'CATFormatTBText "TitleBlock_Text_Cod_1"     , catTopCenter, 5 , 100,0
'--------------------------------------------------------------------------
' ADRES FIRMY
'--------------------------------------------------------------------------
  'Set Text     = DrwTexts.Add("OSRODEK", OH+.5*(Col(1)+Col(2)) , OV+Row(8)-1  )
  'CATFormatTBText "TitleBlock_Text_Logo_1"      , catTopCenter,1.5 , 100,0
  'Set Text     = DrwTexts.Add("BADAWCZO-ROZWOJOWY", OH+.5*(Col(1)+Col(2)) , OV+Row(8)-3  )
  'CATFormatTBText "TitleBlock_Text_Logo_2"      , catTopCenter,1.5 , 100,0
  'Set Text     = DrwTexts.Add("STER", OH+.5*(Col(1)+Col(2)) , OV+Row(8)-5.5  )
  'CATFormatTBText "TitleBlock_Text_Logo_3"      , catTopCenter,2 , 100,0
  'Set Text     = DrwTexts.Add("ul. CzÄąâ€šuchowska 12", OH+.5*(Col(1)+Col(2)) , OV+Row(8)-10  )
  'CATFormatTBText "TitleBlock_Text_Logo_4"      , catTopCenter,1 , 100,0
  'Set Text     = DrwTexts.Add("60-434 POZNAN", OH+.5*(Col(1)+Col(2)) , OV+Row(8)-12  )
  'CATFormatTBText "TitleBlock_Text_Logo_5"      , catTopCenter,1 , 100,0
  'Set Text     = DrwTexts.Add("POLSKA", OH+.5*(Col(1)+Col(2)) , OV+Row(8)-14  )
  'CATFormatTBText "TitleBlock_Text_Logo_6"      , catTopCenter,1 , 100,0
  'Set Text     = DrwTexts.Add("http://www.ster.com.pl", OH+.5*(Col(1)+Col(2)) , OV+Row(8)-17  )
  'CATFormatTBText "TitleBlock_Text_Logo_7"      , catTopCenter,1 , 100,0
  'Set Text     = DrwTexts.Add("e-mail: obr@ster.com.pl", OH+.5*(Col(1)+Col(2)) , OV+Row(8)-19  )
  'CATFormatTBText "TitleBlock_Text_Logo_8"      , catTopCenter,1 , 100,0

CATLinks
End Sub

Sub CATPicture()
'--------------------------------------------------------------------------------------------------------------------------
' Wstawienie loga firmy i innych obrazow
'--------------------------------------------------------------------------------------------------------------------------

Dim MySheet As DrawingSheet
Set MySheet = CATIA.ActiveDocument.Sheets.ActiveSheet
Dim MyView As DrawingView
Set MyView = MySheet.Views.ActiveView
dim sPicturePath As String
sPicturePath=CATIA.SystemService.Environ("CATInstallPath")
Dim MyDrawingPicture1 As DrawingPicture
Set MyDrawingPicture1 = MyView.Pictures.Add(sPicturePath & "\VBScript\FrameTitleBlock\ster_logom.bmp",OH + Col(2)+1,OV+Row(6)-3)


End Sub
Sub CATRevisionBlock(rev As Integer, X As double, Y As Double)
  '-------------------------------------------------------------------------------
  'How to create the revision block
  '-------------------------------------------------------------------------------
  CATRevisionBlockFrame rev, X, Y 'To draw the geometry
  CATRevisionBlockText  rev, X, Y 'To fill in the title block
End Sub

Sub CATRevisionBlockFrame(rev As Integer, X As double, Y As double)
  '-------------------------------------------------------------------------------
  'How to draw the revision block geometry
  '-------------------------------------------------------------------------------
  colRev(1) = -190*mm
  colRev(2) = -175*mm
  colRev(3) = -57*mm
  colRev(4) = - 25*mm
  rev = rev + 1
  On Error Resume Next
    'Set Line = Fact.CreateLine(X + colRev(1), Y, X + colRev(1), Y - RevRowHeight)
    'Line.Name = "RevisionBlock_Line_Column_" + Chr(rev) + "_1"
    'Set Line = Fact.CreateLine(X + colRev(2), Y, X + colRev(2), Y - RevRowHeight)
    'Line.Name = "RevisionBlock_Line_Column_" + Chr(rev) + "_2"

    Set Line = Fact.CreateLine(X + colRev(3), Y, X + colRev(3), Y - 4*RevRowHeight)
    Line.Name = "RevisionBlock_Line_Column_" + Chr(rev) + "_3"
    Set Line = Fact.CreateLine(X + colRev(4), Y, X + colRev(4), Y - 4*RevRowHeight)
    Line.Name = "RevisionBlock_Line_Column_" + Chr(rev) + "_4"

	Set Line = Fact.CreateLine(X + colRev(3), Y - RevRowHeight, X, Y - RevRowHeight)
	Line.Name = "RevisionBlock_Line_Row_1"
	Set Line = Fact.CreateLine(X + colRev(3), Y - 2*RevRowHeight, X, Y - 2*RevRowHeight)
    	Line.Name = "RevisionBlock_Line_Row_2"
	Set Line = Fact.CreateLine(X + colRev(3), Y - 3*RevRowHeight, X, Y - 3*RevRowHeight)
    	Line.Name = "RevisionBlock_Line_Row_3"
	Set Line = Fact.CreateLine(X + colRev(3), Y - 4*RevRowHeight, X, Y - 4*RevRowHeight)
    	Line.Name = "RevisionBlock_Line_Row_4"

	'If (rev = 1) Then
 	'Set Line = Fact.CreateLine(X + colRev(1), Y - RevRowHeight, X + colRev(1), Y - 2.*RevRowHeight)
  	'Line.Name = "RevisionBlock_Line_Column_" + Chr(rev) + "_1"
   	'Set Line = Fact.CreateLine(X + colRev(2), Y - RevRowHeight, X + colRev(2), Y - 2.*RevRowHeight)
    	'Line.Name = "RevisionBlock_Line_Column_" + Chr(rev) + "_2"
     	'Set Line = Fact.CreateLine(X + colRev(3), Y - RevRowHeight, X + colRev(3), Y - 2.*RevRowHeight)
      'Line.Name = "RevisionBlock_Line_Column_" + Chr(rev) + "_3"
      'Set Line = Fact.CreateLine(X + colRev(4), Y - RevRowHeight, X + colRev(4), Y - 2.*RevRowHeight)
      'Line.Name = "RevisionBlock_Line_Column_" + Chr(rev) + "_4"

      'Set Line = Fact.CreateLine(X + colRev(1), Y - 2.*RevRowHeight, X, Y - 2.*RevRowHeight)
      'Line.Name = "RevisionBlock_Line_Row_" + Chr(rev)
      'End If
  If Err.Number <> 0 Then
    Err.Clear
  End If
  On Error Goto 0
End Sub

Sub CATRevisionBlockText(rev As Integer, X As double, Y As double)
  '-------------------------------------------------------------------------------
  'How to fill in the revision block
  '-------------------------------------------------------------------------------
  Init        = InputBox("Klasa dokladności:", "Tolerance", "None")
  Description_1 = InputBox("Chropowatosc powierzchni:", " Surface roughness", "None")
  Description_2 = InputBox("Obrobka cieplna:", " Heat treated", "None")
  Description_3 = InputBox("Pokrycie powierzchni:", " Surface coat", "None")


  If (rev = 1) Then
    Set Text = DrwTexts.Add("Klasa dokladnosci Tolerance"        , X + colRev(3) + 0.5, Y - 0.5*RevRowHeight)
    CATFormatRBText "RevisionBlock_Text_Rev"          , catMiddleLeft ,1.6,31
    Set Text = DrwTexts.Add("Chropowatosc powierzchni Surface roughness"       , X + colRev(3) + 0.5, Y - 1.5*RevRowHeight)
    CATFormatRBText "RevisionBlock_Text_Date"         , catMiddleLeft ,1.45,31
    Set Text = DrwTexts.Add("Obrobka cieplna Heat treated", X + colRev(3) + 0.5, Y - 2.5*RevRowHeight)
    CATFormatRBText "RevisionBlock_Text_Description"  , catMiddleLeft,1.6,26
    Set Text = DrwTexts.Add("Pokrycie powierzchni Surface coat", X + colRev(3) + 0.5, Y - 3.5*RevRowHeight)
    CATFormatRBText "RevisionBlock_Text_Init"         , catMiddleLeft,1.6,31
    Set Text = DrwTexts.Add(Description_1  , X + colRev(4) +0.5, Y - 1.5*RevRowHeight)
    CATFormatRBText "RevisionBlock_Text_Rev_A"        , catMiddleLeft,1.6,24
    Set Text = DrwTexts.Add(Description_2  , X +colRev(4)+0.5, Y - 2.5*RevRowHeight)
    CATFormatRBText "RevisionBlock_Text_Date_A"       , catMiddleLeft,1.6,24
    Set Text = DrwTexts.Add(Description_3  , X + colRev(4) + 0.5, Y - 3.5*RevRowHeight)
    CATFormatRBText "RevisionBlock_Text_Description_A", catMiddleLeft,1.6,24
    Set Text = DrwTexts.Add(Init         , X +  colRev(4) +0.5 , Y - 0.5*RevRowHeight)
    CATFormatRBText "RevisionBlock_Text_Init_A"       , catMiddleLeft,1.6,24
 ' Else
 '   Set Text = DrwTexts.Add(Chr(64+rev)  , X + .5*(colRev(1)+colRev(2)), Y - .5*RevRowHeight)
 '   CATFormatRBText "RevisionBlock_Text_Rev_" + Chr(64+rev)        , catMiddleCenter
 '   Set Text = DrwTexts.Add(Date         , X + .5*(colRev(2)+colRev(3)), Y - .5*RevRowHeight)
 '   CATFormatRBText "RevisionBlock_Text_Date_" + Chr(64+rev)       , catMiddleCenter
 '   Set Text = DrwTexts.Add(Description  , X + colRev(3) + 1., Y - .5*RevRowHeight)
 '   CATFormatRBText "RevisionBlock_Text_Description_" + Chr(64+rev), catMiddleLeft
 '   Text.SetFontName      0, 0, "Monospac821 BT"
 '   Text.SetFontSize 0, 0, 2.5
 '   Set Text = DrwTexts.Add(Init         , X + .5*colRev(4)  , Y - .5*RevRowHeight)
 '   CATFormatRBText "RevisionBlock_Text_Init_" + Chr(64+rev)       , catMiddleCenter
  End If
End Sub

Sub CATMoveReference()
  '-------------------------------------------------------------------------------
  'How to get the reference text
  '-------------------------------------------------------------------------------
  On Error Resume Next
    Set Text = DrwTexts.GetItem("Reference_" + MacroID)
  If Err.Number <> 0 Then
    Err.Clear
    TranslationX = .0
    TranslationY = .0
    Exit Sub
  End If
  On Error Goto 0
  TranslationX = Width - Offset - Text.x
  TranslationY = Offset - Text.y
  Text.x = Text.x + TranslationX
  Text.y = Text.y + TranslationY
End Sub

Sub CATRemoveAll()
  '-------------------------------------------------------------------------------
  'How to remove all the dress-up elements of the active view
  '-------------------------------------------------------------------------------
  Dim NbTexts As Integer
  NbTexts = DrwTexts.Count
  For j = 1 To NbTexts
    DrwTexts.Remove(1)
  Next
  CATRemoveGeometry()
  CATRemovePicture()
End Sub

Sub CATRemovePicture()
 On Error Resume Next
   selection.Add(DrwView)
   selection.Search "Drafting.Picture,sel"
  If Err.Number <> 0 Then
    Err.Clear
    Selection.Clear
    iNbOfGeomElems = GeomElems.Count
    ii = 1
    While (ii <= iNbOfGeomElems)
      Set GeomElem = GeomElems.Item(ii)
      Selection.Add(GeomElem)
      ii = ii + 1
    Wend
  End If
  Selection.Delete
End Sub

Sub CATRemoveGeometry()
  '-------------------------------------------------------------------------------
  'How to remove all geometric elements of the active view
  '-------------------------------------------------------------------------------
  On Error Resume Next
   selection.Add(DrwView)
   selection.Search "Drafting.Geometry,sel"
  If Err.Number <> 0 Then
    Err.Clear
    Selection.Clear
    iNbOfGeomElems = GeomElems.Count
    ii = 1
    While (ii <= iNbOfGeomElems)
      Set GeomElem = GeomElems.Item(ii)
      Selection.Add(GeomElem)
      ii = ii + 1
    Wend
  End If
  Selection.Delete
  On Error Goto 0
End Sub

Sub CATRemoveFrame()
  '-------------------------------------------------------------------------------
  'How to remove the whole frame
  '-------------------------------------------------------------------------------
  On Error Resume Next
    selection.Add(DrwView)
    Selection.Search("Drafting.Text.Name     ='Frame_Text_'*, Drawing")
    If Err.Number = 0 Then
      Selection.Delete
    Else
      Err.Clear
      iNbOfTexts = DrwTexts.Count
      ii = iNbOfTexts
      While (ii > 0)
        Set Text = DrwTexts.Item(ii)
        if (Left(Text.Name, 11) = "Frame_Text_")  Then
          DrwTexts.Remove(ii)
        End If
        ii = ii - 1
      Wend
    End If

    Selection.Search("Drafting.Geometry.Name ='Frame_'*, Drawing")
    If Err.Number <> 0 Then
      Err.Clear
      Selection.Clear
      iNbOfGeomElems = GeomElems.Count
      ii = 1
      While (ii <= iNbOfGeomElems)
        Set GeomElem = GeomElems.Item(ii)
        if (Left(GeomElem.Name, 6) = "Frame_")  Then
          Selection.Add(GeomElem)
        End If
        ii = ii + 1
      Wend
    End If
    Selection.Delete
    On Error Goto 0
End Sub

Sub CATMoveTitleBlock()
  '-------------------------------------------------------------------------------
  'How to translate the whole title block after changing the page setup
  '-------------------------------------------------------------------------------
  Dim rootName As String

  Dim rootNameLength As Integer
  Dim NbLineToMove   As Integer
  Dim NbCircleToMove As Integer
  Dim NbTextToMove As Integer

  Dim Origin(2)
  Dim Direction(2)
  Dim Radius As Double

  rootName       = "TitleBlock_Line_"
  rootNameLength = Len(rootName)
  NbLineToMove   = GeomElems.Count
  For i = 1 To NbLineToMove
    Set Line = GeomElems.Item(i)
    If  (Left(Line.Name, rootNameLength) = rootName) Then
      Line.GetOrigin(Origin)
      Line.GetDirection(Direction)
      Line.SetData Origin(0)+TranslationX, Origin(1)+TranslationY, Direction(0), Direction(1)
    End If
  Next

  rootName       = "TitleBlock_Standard_Line_"
  rootNameLength = Len(rootName)
  NbLineToMove   = GeomElems.Count
  For i = 1 To NbLineToMove
    Set Line = GeomElems.Item(i)
    If  (Left(Line.Name, rootNameLength) = rootName) Then
      Line.GetOrigin(Origin)
      Line.GetDirection(Direction)
      Line.SetData Origin(0)+TranslationX, Origin(1)+TranslationY, Direction(0), Direction(1)
    End If
  Next

  rootName       = "TitleBlock_Standard_Circle"
  rootNameLength = Len(rootName)
  NbCircleToMove = GeomElems.Count
  For i = 1 To NbCircleToMove
    Set Circle = GeomElems.Item(i)
    If  (Left(Circle.Name, rootNameLength) = rootName) Then
      Circle.GetCenter(Origin)
      Radius = Circle.Radius
      Circle.SetData Origin(0)+TranslationX, Origin(1)+TranslationY, Radius
    End If
  Next

  rootName       = "TitleBlock_Text_"
  rootNameLength = Len(rootName)
  NbTextToMove   = DrwTexts.Count
  For i = 1 To NbTextToMove
    Set Text = DrwTexts.Item(i)
    If  (Left(Text.Name, rootNameLength) = rootName) Then
      Text.x = Text.x + TranslationX
      Text.y = Text.y + TranslationY
    End If
  Next
End Sub

Sub CATFormatFText(textName As String, angle As Double)
  '-------------------------------------------------------------------------------
  'How to format the texts belonging to the frame
  '-------------------------------------------------------------------------------
  Text.Name           = textName
  Text.SetFontName      0, 0, "Monospac821 BT"
  Text.AnchorPosition = CATMiddleCenter
  Text.Angle          = angle

End Sub

Sub CATFormatTBText(textName As String, anchorPosition As String, fontSize, Wrapp As Double,angle As Double)
  '-------------------------------------------------------------------------------
  'How to format the texts belonging to the titleblock
  '-------------------------------------------------------------------------------
  Text.Name           = textName
  Text.SetFontName      0, 0, "Monospac821 BT"
  Text.AnchorPosition = anchorPosition
  Text.SetFontSize      0, 0, fontSize
  Text.WrappingWidth = Wrapp
  Text.Angle          = angle
End Sub

Sub CATFormatRBText(textName As String, anchorPosition As String,fontSize,Wrapp As Double)
  '-------------------------------------------------------------------------------
  'How to format the texts belonging to the titleblock
  '-------------------------------------------------------------------------------
  Text.Name           = textName
  Text.SetFontName      0, 0, "Monospac821 BT"
  Text.AnchorPosition = anchorPosition
  Text.SetFontSize      0, 0, fontSize
  Text.WrappingWidth = Wrapp
End Sub

Sub CATLinks()
  '-------------------------------------------------------------------------------
  'How to fill in texts with data of the part/product linked with current sheet
  '-------------------------------------------------------------------------------
  On Error Resume Next
    Dim ProductDrawn As ProductDocument
    Set ProductDrawn = DrwSheet.Views.Item("Front view").GenerativeBehavior.Document
  If Err.Number = 0 Then
    DrwTexts.GetItem("TitleBlock_Text_Number_1").Text = ProductDrawn.PartNumber
    DrwTexts.GetItem("TitleBlock_Text_Title_1").Text  = ProductDrawn.DescriptionRef
	'DrwTexts.GetItem("TitleBlock_Text_Title_1").Text  = ProductDrawn.Definition
    DrwTexts.GetItem("TitleBlock_Text_Cod_1").Text  = ProductDrawn.Nomenclature
   'dobre!! DrwTexts.GetItem("TitleBlock_Text_Material_1").Text  = ProductDrawn.UserRefProperties.Item("Material").value
   DrwTexts.GetItem("TitleBlock_Text_Material_1").Text  = ProductDrawn.Parameters.Item("Material").value   ' o to chodzilo !!
Material

Dim ProductAnalysis As Analyze
    Set ProductAnalysis = ProductDrawn.Analyze
    DrwTexts.GetItem("TitleBlock_Text_Weight_1").Text = FormatNumber(ProductAnalysis.Mass,3)
  End If

  '-------------------------------------------------------------------------------
  'Display sheet format
  '-------------------------------------------------------------------------------
  Dim textFormat As DrawingText
  Set textFormat = DrwTexts.GetItem("TitleBlock_Text_Size_1")
  textFormat.Text = displayFormat
  If (Len(displayFormat) > 4 ) Then
    textFormat.SetFontSize 0, 0, 2.5
  Else
    textFormat.SetFontSize 0, 0, 4.
  End If

  '-------------------------------------------------------------------------------
  'Display sheet numbering
  '-------------------------------------------------------------------------------
  Dim nbSheet  As Integer
  Dim curSheet As Integer
  nbSheet  = 0
  curSheet = 0

  If (not DrwSheet.IsDetail) Then
    For i = 1 To DrwSheets.Count
      If (not DrwSheets.Item(i).IsDetail) Then
        nbSheet = nbSheet + 1
      End If
    Next
    For i = 1 To DrwSheets.Count
      If (not DrwSheets.Item(i).IsDetail) Then
        On Error Resume Next
        curSheet = curSheet + 1
        DrwSheets.Item(i).Views.Item(2).Texts.GetItem("TitleBlock_Text_Sheet_1").Text = CStr(curSheet)
	DrwSheets.Item(i).Views.Item(2).Texts.GetItem("TitleBlock_Text_Arkusze_1").Text = CStr(nbSheet)
      End If
    Next
  End If
  On Error Goto 0


End Sub

Sub CATFillField(string1 As String, string2 As String, string3 As String)
  '-------------------------------------------------------------------------------
  'How to call a dialog to fill in manually a given text
  '-------------------------------------------------------------------------------
  Dim TextToFill_1 As DrawingText
  Dim TextToFill_2 As DrawingText
  Dim Person As String

  Set TextToFill_1 = DrwTexts.GetItem(string1)
  Set TextToFill_2 = DrwTexts.GetItem(string2)

  Person = TextToFill_1.Text
  If (Person = "") Then
    Person = "Andrzej Perz"
  End If

  Person = InputBox("Ten dokument jest " + string3 + " przez:", "Controller's name", Person)
  If (Person = "") Then
    Person = "Andrzej Perz"
  End If

  TextToFill_1.Text = Person
  TextToFill_2.Text = Date
End Sub

Sub CATRevPos(rev As Integer, oX As Double, oY As Double)
  '-------------------------------------------------------------------------------
  'How to local the the current revision
  '-------------------------------------------------------------------------------
  CATStandard
  oX = OH
  if (rev = 0) Then
    oY = Height - OV
  Else
    oY = DrwTexts.GetItem("RevisionBlock_Text_Rev_" + Chr(64+rev)).y - .5*RevRowHeight
  End If
End Sub

Sub CATMain()

'Use InputBox instead of using textbox in Visaul Basic

Dim ElementName as string
ElementName = InputBox("Please eneter element name")


Dim documents1 As Documents
Set documents1 = CATIA.Documents

Dim Selection As Selection
Set Selection = CATIA.Activedocument.selection

Dim ElementsArray(0)
ElementsArray(0) = "AnyObject"
Dim Status As String
Status = Selection.SelectElement3(ElementsArray, "Select Elements for name change", False, CATMultiSelectionMode.CATMultiSelTriggWhenUserValidatesSelection, False)

Dim i As Single
For i = 1 To Selection.Count
Selection.Item(i).Value.name = ElementName & "." & i
Selection.Item(i).Value.DescriptionInst = ElementName & "." & i
Selection.Item(i).Value.PartNumber = ElementName & "." & i
Selection.Item(i).Value.Definition = ElementName & "." & i
Selection.Item(i).Value.DescriptionRef = ElementName & "." & i

Next

End Sub


