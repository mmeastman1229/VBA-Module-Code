Attribute VB_Name = "MyUtilities"
Option Explicit
Dim GlobalHandler As New GlobalEventHandler


Public Sub OptimizedMode(ByVal enable As Boolean)

Application.EnableEvents = Not enable
Application.Calculation = IIf(enable, xlCalculationManual, xlCalculationAutomatic)
Application.ScreenUpdating = Not enable
Application.EnableAnimations = Not enable
Application.DisplayAlerts = Not enable
Application.DisplayStatusBar = Not enable
Application.PrintCommunication = Not enable

End Sub

Function GetCurrentSelection() As Range
    Set GetCurrentSelection = Application.Selection
End Function

Sub CompletelyResetExcel()
    
Dim Sh                          As Worksheet

OptimizedMode False

For Each Sh In Application.Worksheets
    
    With Sh
        ActiveSheet.Range("A1").Select
        .Cells.RowHeight = 15
        .Cells.ColumnWidth = 8.43
        .Cells.Clear
    End With

Next Sh

End Sub


Sub LeftUpperAlignAllCells(ByVal rng As Range)
OptimizedMode True

With rng
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlTop
End With

OptimizedMode False
End Sub


Sub TransformSelectedCase(ByRef str As String, ByRef rng As Range)
OptimizedMode True

Dim cell As Range

For Each cell In rng
    Select Case str
        Case "upper"
            cell.Value = StrConv(cell, vbUpperCase)
        Case "lower"
            cell.Value = StrConv(cell, vbLowerCase)
        Case "title"
            cell.Value = StrConv(cell, vbProperCase)
    End Select
Next cell

OptimizedMode False
End Sub


Sub FillInteriorColor(ByVal rngRedValue as Range, ByVal rngGreenValue as Range, _
ByVal rngBlueValue as Range, byVal rngDestinationCells as Range)
OptimizedMode True

Dim i                                   As Integer
Dim red                                 As Long
Dim green                               As Long
Dim blue                                As Long

For i = 1 To rngDestinationCell.Cells.Count
     red = rngRedValue(i).Value
     green = rngGreenValue(i).Value
     blue = rngBlueValue(i).Value
     rngDestinationCells(i).Interior.Color = RGB(red, green, blue)

Next

OptimizedMode False
End Sub


Sub AutoFitSelectedCells(ByVal inputRng)

OptimizedMode True

Dim rng                                 As Range
Set rng = inputRng

With rng
    .EntireColumn.AutoFit
    .EntireRow.AutoFit
End With

OptimizedMode False

End Sub

Sub AutoFitSelectedCellsNoArgs()

OptimizedMode True

Dim ws As Workbook


Dim rng As Range

Set rng = Selection

With rng
    .EntireColumn.AutoFit
    .EntireRow.AutoFit
End With

OptimizedMode False

End Sub

Sub ToggleAutoFilter(ByVal rng)

With rng
    .AutoFilter
End With

End Sub

Sub ClearSelection()

Dim rng As Range
Set rng = Selection

With rng
    .Clear
End With

End Sub

' ---------------------------------------------------------------------
' AutoNumber - Numbers target range from 1 to n, based
' on specific cell, then starts over if lookup value changes.
'
' Used for numbering sets of data
'
' Created by Matthew Eastman 4/19/23
' ---------------------------------------------------------------------

Sub AutoNumber()

OptimizedMode True

Dim wb                                  As Workbook
Dim ws                                  As Worksheet

Dim EntityName                          As Range
Dim NextEntityName                      As Range
Dim ViewName                            As Range
Dim NextViewName                        As Range
Dim TargetRange                         As Range
Dim i                                   As Integer
Dim intNum                              As Integer
Dim finalRow                            As Long

finalRow = Cells(Rows.Count, 1).End(xlUp).Row

Set wb = ActiveWorkbook
Set ws = wb.ActiveSheet

With ws
    Set EntityName = .Range("B2:B" & finalRow)
    Set NextEntityName = EntityName.Offset(1, 0)
    Set ViewName = .Range("C2:C" & finalRow)
    Set NextViewName = ViewName.Offset(1, 0)
    Set TargetRange = .Range("D2:D" & finalRow)

    intNum = 0

    For i = 1 To EntityName.Cells.Count
        intNum = intNum + 1
        TargetRange.Cells(i).Value = intNum
        If (NextEntityName.Cells(i) <> EntityName.Cells(i)) And (NextViewName.Cells(i) <> ViewName.Cells(i)) Then
            intNum = 0
        End If
    Next

End With

OptimizedMode False

End Sub

Sub ToggleFreezeUnfreezeAtSelection()

OptimizedMode True

Dim ws As Worksheet
Set ws = ActiveSheet

With ActiveWindow
    If .FreezePanes = True Then
        .FreezePanes = False
    ElseIf .FreezePanes = False Then
        .FreezePanes = True
    End If
End With

OptimizedMode False

End Sub

'Sub CreateHyperlinkWithReturnLocation(ByVal target As Range, _
'    ByVal rtrn As Range)
'
'
'    ' Set up workbook and worksheet information
'    Dim wb As Workbook
'    Dim ws As Worksheet
'
'    Set wb = ActiveWorkbook
'    Set ws = wb.ActiveSheet
'
'    ' Create variables for targetCell and returnCell
'    Dim targetCell As Range
'    Dim returnCell As Range
'
'
'    ' Set targetCell and returnCell to argument variables
'    Set targetCell = target
'    Set returnCell = rtrn
'
'
'    ' Create additional variables necessary for hyperlink
'    Dim TargetScreenTipTxt As String
'    Dim ReturnScreenTipTxt As String
'    Dim NotNeeded As String
'
'    ' Set variables created above
'    screenTipTxt = "*See Note"
'    NotNeeded = ""
'
'
'
'
'
'    ' Create link to targetCell
'    ' Create variables needed for returnCell
'        ' set returnCell initial value
'        ' set returnCell address
'    ' Create returnCell hyperlink
'    ' create target cell information
'
'
'
'    Dim NotNeeded As String
'
'
'    ' Set hyperlink variables
'    Dim returnCellValue As String
'    returnCellValue = returnCell.Value
'
'
'
'
'    ' Add hyperlink to target cell
'    ws.Hyperlinks.Add _
'        Anchor:=targetCell, _
'        Address:=NotNeeded, _
'        SubAddress:=NotNeeded, _
'        screenTip:=screenTipTxt, _
'        TextToDisplay:=screenTipTxt
'
'    ' Store original location in the return cell
'    returnCell = targetCell.Address
'
'    ' create hyperlink back to original cell
'    Dim rtrnTxt As String
'    rtrnTxt = targetCell.Value & "(Return)"
'
'    ws.Hyperlinks.Add _
'        Anchor:=returnCell, _
'        Address:=NotNeeded, _
'        SubAddress:=NotNeeded, _
'        screenTip:=rtrnTxt, _
'        TextToDisplay:=rtrnTxt
'
'End Sub
    



Sub FacilityClaimFilter()
Dim sht                             As Worksheet
Dim rngFilter                       As Range
Dim fieldNumber                     As Integer

Set sht = ActiveSheet
Set rngFilter = sht.Range("G3")
fieldNumber = 7

rngFilter.AutoFilter Field:=fieldNumber, Criteria1:=Array( _
"link_mrn", "post_date_key", "service_date_key", "admit_datetime", "discharge_datetime", "source_provider_id", "main_specialty_code", "provider_type_code", "place_of_service_code", _
"revenue_code", "hcpcs_code", "hcpcs_modifier_code", "procedure_line_number", "diagnosis_line_number", "drg_code", "classification_scheme_code", "outlier_number", "present_on_admission_code", _
"discharge_status_code", "unit_of_service_quantity", "total_cost", "source_system_name", "region_code, encounter_number", _
"sequence_number", "bill_type_code", "performing_provider_key", "bill_facility_type_code", "admission_source_code", _
"admission_type_code", "ndc", "hcpc_modifier_code", "encounter_number", "region_code", "transaction_number", _
"region_key", "cost_run_key", "total_fixed_direct_cost", _
"total_fixed_indirect_cost", "total_variable_direct_cost", "total_variable_indirect_cost", "hcg_detail_code", _
"hcg_pbp_code", "hcg_code", "hcg_case_admit_label", "hcg_unit_day_label", "hcg_procedure_label", "hcg_case_admit_count", "hcg_unit_day_count", _
"hcg_procedure_count", "hcg_pbp_case_admit_count"), _
Operator:=xlFilterValues


End Sub


Sub MedicalClaimFilter()
Dim sht                             As Worksheet
Dim rngFilter                       As Range
Dim fieldNumber                     As Integer

Set sht = ActiveSheet
Set rngFilter = sht.Range("A1")
fieldNumber = 1

rngFilter.AutoFilter Field:=fieldNumber, Criteria1:=Array("Current Ops Dimensional Layer - Medical Encounter"), _
Operator:=xlFilterValues

Set rngFilter = sht.Range("F1")
fieldNumber = 6

rngFilter.AutoFilter Field:=fieldNumber, Criteria1:=Array( _
"person", "medical_encounter", "encounter_provider", "place_of_service", "hcpcs", _
"hcpcs_modifier", "medical_encounter_icd_diagnosis", "drug", "region", "hcg", _
"department"), _
Operator:=xlFilterValues

Set rngFilter = sht.Range("G1")
fieldNumber = 7

rngFilter.AutoFilter Field:=fieldNumber, Criteria1:=Array( _
"link_mrn", "post_date_key", "service_date_key", "source_provider_id", "main_specialty_code", _
"provider_type_code", "place_of_service_code", "hcpcs_code", "hcpcs_modifier_code", _
"diagnosis_line_number", "unit_of_service_quantity", "total_cost", "source_system_name", _
"sequence_number", "performing_provider_key", "service_facility_code", "ndc", "encounter_number", _
"encounter_sequence_number", "region_code", "gl_region_code", "transaction_number", "region_key", _
"cost_run_key", "total_fixed_direct_cost", "total_fixed_indirect_cost", "total_variable_direct_cost", _
"total_variable_indirect_cost", "hcg_detail_code", "hcg_pbp_code", "hcg_code", "hcg_case_admit_label", _
"hcg_unit_day_label", "hcg_procedure_label", "hcg_case_admit_count", "hcg_unit_day_count", "hcg_procedure_count", _
"hcg_pbp_case_admit_count"), _
Operator:=xlFilterValues


End Sub

Sub SupplementalClaimFilterNCAP()
Dim sht                             As Worksheet
Dim rngFilter                       As Range
Dim fieldNumber                     As Integer

Set sht = ActiveSheet
Set rngFilter = sht.Range("G1")
fieldNumber = 7

rngFilter.AutoFilter Field:=fieldNumber, Criteria1:=Array( _
    "link_mrn", "post_date_key", "service_date_key", "source_provider_id", "main_specialty_code", _
    "provider_type_code", "place_of_service_code", "hcpcs_code", "diagnosis_line_number", _
    "procedure_quantity", "total_cost", "source_system_name", "region_code, encounter_number", _
    "encounter_sequence_number", "encounter_number", "region_code", _
    "gl_region_code,encounter_number", "transaction_number", "region_key", "cost_run_key", _
    "total_fixed_direct_cost", "total_fixed_indirect_cost", "total_variable_direct_cost", _
    "total_variable_indirect_cost", "hcg_detail_code", "hcg_pbp_code", _
    "hcg_code", "hcg_case_admit_label", "hcg_unit_day_label", "hcg_procedure_label", _
    "hcg_case_admit_count", "hcg_unit_day_count", "hcg_procedure_count", "hcg_pbp_case_admit_count"), _
Operator:=xlFilterValues

End Sub

Sub PharmacyEncounterFilter()
Dim sht                             As Worksheet
Dim rngFilter                       As Range
Dim fieldNumber                     As Integer

Set sht = ActiveSheet
Set rngFilter = sht.Range("A1")
fieldNumber = 1

rngFilter.AutoFilter Field:=fieldNumber, Criteria1:=Array("Current Ops Dimensional Layer - Pharmacy Encounter"), _
Operator:=xlFilterValues

Set rngFilter = sht.Range("F1")
fieldNumber = 6

rngFilter.AutoFilter Field:=fieldNumber, Criteria1:=Array( _
"person", "pharmacy_encounter", "department", "drug", "region", "hcg"), _
Operator:=xlFilterValues

Set rngFilter = sht.Range("G1")
fieldNumber = 7

rngFilter.AutoFilter Field:=fieldNumber, Criteria1:=Array( _
"link_mrn", "post_date_key", "service_date_key", "pharmacy_id", "department_name", "fill_location_code", _
"ndc", "generic_drug_code", "day_supply_count", "dispensed_quantity", "total_cost", "charge_amount", _
"total_cost", "source_system_name", "region_code", "dispensed_quantity", "encounter_number", "encounter_sequence_number", _
"region_code", "source_system_name", "region_code", "region_key", "cost_run_key", "total_fixed_direct_cost", "total_fixed_indirect_cost", _
"total_variable_direct_cost", "total_variable_indirect_cost", "total_fixed_direct_cost", "total_fixed_indirect_cost", "total_variable_direct_cost", "total_variable_indirect_cost", _
"hcg_detail_code", "hcg_pbp_code", "hcg_code", "hcg_case_admit_label", "hcg_unit_day_label", "hcg_procedure_label", "hcg_unit_day_count", "hcg_procedure_count"), _
Operator:=xlFilterValues


End Sub

Sub ClearAllFiltersRange()
On Error Resume Next
    ActiveSheet.ShowAllData
On Error GoTo 0
End Sub


Sub ResetActiveCell()

OptimizedMode True

Dim wb As Workbook
Dim ws As Worksheet
Dim shtReturn As Worksheet

Set wb = ActiveWorkbook
Set shtReturn = wb.ActiveSheet

For Each ws In wb.Worksheets
    ws.Activate
    ws.Range("A1").Select
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
Next
shtReturn.Select

OptimizedMode False

End Sub

Sub ShowAllColumns()
    Dim ws As Worksheet
    Dim col As Range
    
    Set ws = ActiveSheet
    For Each col In ws.Columns
        If col.Hidden Then
            col.Hidden = False
        End If
    Next col
End Sub

Sub UnFreezeAllSheets()

    Dim ws As Worksheet
    Dim wkbk As Workbook

    OptimizedMode True
    

    Set wkbk = ActiveWorkbook

    For Each ws In wkbk.Sheets
        ws.Activate
        
        With ActiveWindow
            .FreezePanes = False
        End With
        
    Next ws
    
    OptimizedMode False
    
    wkbk.Sheets(1).Select

End Sub

Sub AllSheetQuickFix()

    OptimizedMode True
    
    Dim ws As Worksheet
    Dim wkbk As Workbook

    Set wkbk = ActiveWorkbook

    For Each ws In wkbk.Sheets
        ws.Activate
    
       Range("B:B").ColumnWidth = 50
'       Range("B:B").ColumnWidth = 25
'       Range("C:C").ColumnWidth = 30
'       Range("D:D").ColumnWidth = 45
'       Range("E:E").ColumnWidth = 50
'       Range("F:F").ColumnWidth = 30
'       Range("G:G").ColumnWidth = 35
'       Range("H:H").ColumnWidth = 16
'       Range("I:M").ColumnWidth = 16
'       Range("N:O").ColumnWidth = 15

Cells.RowHeight = 18

       
        
    Next ws
    
    OptimizedMode False

End Sub

Sub Final()

Dim finalRow As Long


finalRow = Cells(Rows.Count, 1).End(xlUp).Row

Range(ActiveCell.Address, finalRow).Select


End Sub

Sub TurnOffGridLinesAllSheets()

OptimizedMode True

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets

    ws.Activate
    
    If ActiveWindow.DisplayGridlines = False Then
        ActiveWindow.DisplayGridlines = True
    ElseIf ActiveWindow.DisplayGridlines = True Then
        ActiveWindow.DisplayGridlines = False
    End If
     
    
    
Next ws

OptimizedMode False

End Sub

Sub RemoveFillColorFromSelection()

OptimizedMode True

Dim rng As Range
Set rng = Selection

With rng
    .Interior.Color = xlNone
End With

OptimizedMode False

End Sub

Sub CenterBoth()

Dim rng As Range
Set rng = Selection

With rng
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

End Sub


Sub CenterTop()

Dim rng As Range
Set rng = Selection

With rng
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlTop
End With

End Sub

Sub LeftBottom()

Dim rng As Range
Set rng = Selection

With rng
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
End With

End Sub

Sub CalibriBoldItalic()

Dim rng As Range
Set rng = Selection

With rng
    .Font.Name = "Calibri Light"
    .Font.Size = 14
    .Font.Bold = True
    .Font.Italic = False
    
    
End With

End Sub

Sub SelectionQuickFix()

Dim rng As Range
Set rng = Selection
'
With rng
    .Font.Name = "Calibri Light"
    .Font.Size = 12
    .Font.Bold = True
    .Font.Italic = False

End With
'Range("B:B, F:F, P:P").ColumnWidth = 30
'Range("D:E, T:W, Y:Y").ColumnWidth = 15
'Range("G:H, O:O, R:R, X:X, Z:AA").ColumnWidth = 35
'Range("I:I, Q:Q").ColumnWidth = 40
'Range("J:J, L:M, S:S").ColumnWidth = 20
'Range("N:N").ColumnWidth = 11


End Sub


Sub CalibriHeader()

Dim rng As Range
Set rng = Selection
'
With rng
    .Font.Name = "Calibri Light"
    .Font.Size = 14
    .Font.Bold = True
    .Font.Italic = False

End With


End Sub

Sub FillColor_Darken()
'PURPOSE: Darken the cell fill by a shade while maintaining Hue (base Color)
'SOURCE: www.TheSpreadsheetGuru.com

Dim HexColor As String
Dim cell As Range
Dim Darken As Integer
Dim r As Integer
Dim g As Integer
Dim b As Integer
Dim r_new As Integer
Dim g_new As Integer
Dim b_new As Integer

'Shade Settings
  Darken = 1 'recommend 3 (1-16)

'Optimize Code
  Application.ScreenUpdating = False

'Loop through each cell in selection
  For Each cell In Selection.Cells
    
    'Determine HEX color code
      HexColor = Right("000000" & Hex(cell.Interior.Color), 6)
    
    'Determine current RGB color code
      r = CInt("&H" & Right(HexColor, 2))
      g = CInt("&H" & Mid(HexColor, 3, 2))
      b = CInt("&H" & Left(HexColor, 2))
    
    'Calculate new RGB color code
      r_new = WorksheetFunction.Round((r * 15 - 255 * Darken) / (15 - Darken), 0)
      g_new = WorksheetFunction.Round((g * 15 - 255 * Darken) / (15 - Darken), 0)
      b_new = WorksheetFunction.Round((b * 15 - 255 * Darken) / (15 - Darken), 0)
    
    'Change enitre selection's fill color
      On Error Resume Next
        cell.Interior.Color = RGB(r_new, g_new, b_new)
      On Error GoTo 0
  
  Next cell

End Sub

Sub FillColor_Lighten()
'PURPOSE: Lighten the cell fill by a shade while maintaining Hue (base Color)
'SOURCE: www.TheSpreadsheetGuru.com

Dim HexColor As String
Dim cell As Range
Dim Lighten As Integer
Dim r As Integer
Dim g As Integer
Dim b As Integer
Dim r_new As Integer
Dim g_new As Integer
Dim b_new As Integer

'Shade Settings
  Lighten = 1 'recommend 3 (1-16)

'Optimize Code
  OptimizedMode True

'Loop through each cell in selection
  For Each cell In Selection.Cells

    'Determine HEX color code
      HexColor = Right("000000" & Hex(cell.Interior.Color), 6)

    'Determine current RGB color code
      r = CInt("&H" & Right(HexColor, 2))
      g = CInt("&H" & Mid(HexColor, 3, 2))
      b = CInt("&H" & Left(HexColor, 2))

    'Calculate new RGB color code
      r_new = WorksheetFunction.Round(r + (Lighten * (255 - r)) / 15, 0)
      g_new = WorksheetFunction.Round(g + (Lighten * (255 - g)) / 15, 0)
      b_new = WorksheetFunction.Round(b + (Lighten * (255 - b)) / 15, 0)

      'Debug.Print r_new, g_new, b_new

    'Change enitre selection's fill color
      cell.Interior.Color = RGB(r_new, g_new, b_new)

  
  Next cell
  
  'Reset optimized mode
  OptimizedMode False

End Sub

Sub SortSheetsTabName()
   OptimizedMode True
    
    Dim iSheets%, i%, j%
    iSheets = Sheets.Count
    For i = 1 To iSheets - 1
        For j = i + 1 To iSheets
            If Sheets(j).Name < Sheets(i).Name Then
                Sheets(j).Move before:=Sheets(i)
            End If
        Next j
    Next i
    
    OptimizedMode False
    
End Sub

Sub ListAllWorksheets()

Dim wkbk As Workbook
Dim sht As Worksheet
Set wkbk = ActiveWorkbook

For Each sht In wkbk.Sheets
    Debug.Print sht.Name
Next sht

End Sub
Function IsMerged(rCell As Range) As Boolean

    IsMerged = rCell.MergeCells

End Function


Sub ToggleMerge()

Dim rng As Range
Set rng = Selection

With rng
    If Not rng.MergeCells Then
        rng.Merge
    Else
        rng.UnMerge
    End If
   
End With

End Sub


Sub GetRGBColor_Fill()
'PURPOSE: Output the RGB color code for the ActiveCell's Fill Color
'SOURCE: www.TheSpreadsheetGuru.com

Dim HexColor As String
Dim RGBcolor As String

HexColor = Right("000000" & Hex(ActiveCell.Interior.Color), 6)

RGBcolor = "RGB (" & CInt("&H" & Right(HexColor, 2)) & _
", " & CInt("&H" & Mid(HexColor, 3, 2)) & _
", " & CInt("&H" & Left(HexColor, 2)) & ")"


MsgBox RGBcolor, vbInformation, "Cell " & ActiveCell.Address(False, False) & ":  Fill Color"

End Sub

Sub GetRGBColor_Font()
'PURPOSE: Output the RGB color code for the ActiveCell's Font Color
'SOURCE: www.TheSpreadsheetGuru.com

Dim HexColor As String
Dim RGBcolor As String

HexColor = Right("000000" & Hex(ActiveCell.Font.Color), 6)

RGBcolor = "RGB (" & CInt("&H" & Right(HexColor, 2)) & _
", " & CInt("&H" & Mid(HexColor, 3, 2)) & _
", " & CInt("&H" & Left(HexColor, 2)) & ")"

MsgBox RGBcolor, vbInformation, "Cell " & ActiveCell.Address(False, False) & ":  Font Color"

End Sub

Sub GetHexColor()
Dim HexColor As String
Dim copyCell As String


HexColor = Right("000000" & Hex(ActiveCell.Interior.Color), 6)

'Reverse the Hex code
HexColor = Right(HexColor, 2) & Mid(HexColor, 3, 2) & Left(HexColor, 2)


MsgBox "#" & HexColor, vbInformation, "Cell " & ActiveCell.Address(False, False) & ":  Fill Color in HEX"
End Sub

Sub CreateLinkToSheet()

Dim targetWorksheet As Worksheet
Dim hyperlinkAddress As String
Dim wsName As String
Dim wkbk As Workbook


' set worksheet name
wsName = "Summary"

' Set the active workbook
Set wkbk = ActiveWorkbook

    
' Set the target worksheet name
Set targetWorksheet = wkbk.Worksheets(wsName)

' Get address of the worksheet
hyperlinkAddress = "'" & targetWorksheet.Name & "'!A1"


' Create a hyperlin in the active cell
ActiveCell.Hyperlinks.Add _
    Anchor:=ActiveCell, _
    Address:=hyperlinkAddress, _
    TextToDisplay:=targetWorksheet.Name
    

End Sub

Sub GetFontInfo()

Dim rng As Range
Dim FontName As String
Dim fontSize As Integer
Dim fontStyle As String

Set rng = Selection

FontName = rng.Font.Name
fontSize = rng.Font.Size
fontStyle = rng.Font.fontStyle


Debug.Print FontName
Debug.Print fontSize
Debug.Print fontStyle



End Sub

Sub QuickFixColumn()

Dim wb As Workbook
Dim ws As Worksheet

Set wb = Application.ActiveWorkbook

For Each ws In wb.Worksheets
    ws.Rows(1).RowHeight = 30
    ws.Range("A1").ColumnWidth = 38
    ws.Range("B1").ColumnWidth = 26
    ws.Range("C1").ColumnWidth = 37
    ws.Range("D1").ColumnWidth = 11
    ws.Range("E1").ColumnWidth = 40
    ws.Range("F1").ColumnWidth = 50
    ws.Range("G1").ColumnWidth = 25
    ws.Range("H1").ColumnWidth = 41
    ws.Range("I1").ColumnWidth = 14
    ws.Range("J1").ColumnWidth = 13.5
    ws.Range("K1:N1").ColumnWidth = 9
    ws.Range("O1").ColumnWidth = 12
    ws.Range("P1").ColumnWidth = 11

Next ws

OptimizedMode True





OptimizedMode False

End Sub

Sub BillType()

Dim rng As Range
Dim cell As Range
Dim searchCritearia As String
Dim searchRange As Range

Set rng = Range("BI2:BI272072")

For Each cell In rng


Next cell


End Sub

Sub SelectFilterCriteria()

Dim sht                             As Worksheet
Dim rngFilter                       As Range
Dim fieldNumber                     As Integer
Dim operatorValue                   As Integer

Set sht = ActiveSheet
Set rngFilter = sht.Range("G1")
fieldNumber = 7
operatorValue = 7

rngFilter.AutoFilter Field:=fieldNumber, Criteria1:=Array( _
"encounter_number", "line", "region_code", "source_claim_type", "medical_claim_header_id", "transaction_number", "region_key", _
"cost_run_key", "ncap_institutional_claim_key", "ncap_institutional_detail_line_number", "ncap_professional_claim_key", _
"ncap_professional_detail_line_number", "total_fixed_direct_cost", "total_fixed_indirect_cost", _
"total_variable_direct_cost", "total_variable_indirect_cost", "total_cost", "hcg_detail_code", _
"hcg_pbp_code", "hcg_code", "hcg_case_admit_label", "hcg_unit_day_label", "hcg_procedure_label", _
"hcg_case_admit_count", "hcg_unit_day_count", "hcg_procedure_count", "hcg_pbp_case_admit_count"), Operator:=operatorValue

End Sub

Sub QuickFixCell()

Dim ws As Worksheet
Dim rng As Range
Dim cell As Range

Set rng = Selection

For Each cell In rng

    If cell.Value = False Then
        cell.Value = "Not NULL"
    Else:
        cell.Value = "NULL"
    End If
    
    
Next cell

End Sub


Sub FillEmptyWithNull()

OptimizedMode True

Dim rng As Range
Dim cell As Range

Set rng = Range("B2:EJ816")

For Each cell In rng

    If cell.Value = " " Then
        cell.Value = "NULL"
    
    End If
Next cell

OptimizedMode False

End Sub

Sub FormatDateYYYYMMDD()

Dim rng As Range
Dim expression As String

Set rng = Selection
expression = "YYYY-MM-DD"

End Sub


Sub S2TConditionalFormatting()

Dim rng As Range
Dim Finalized As FormatCondition
Dim ValidationPending As FormatCondition
Dim LockedForDeployment As FormatCondition
Dim InDevelopment As FormatCondition

Set rng = Range("$B$3")

rng.FormatConditions.Delete 'Clean any existing formatting

Set Finalized = rng.FormatConditions.Add(Type:=xlTextString, String:="Finalized", TextOperator:=xlContains)
Set ValidationPending = rng.FormatConditions.Add(Type:=xlTextString, String:="Validation Pending", TextOperator:=xlContains)
Set LockedForDeployment = rng.FormatConditions.Add(Type:=xlTextString, String:="Locked for Deployment", TextOperator:=xlContains)
Set InDevelopment = rng.FormatConditions.Add(Type:=xlTextString, String:="In Development", TextOperator:=xlContains)

With Finalized
    .Interior.Color = HexToDecimal("C6EFCE")
    .Font.Color = HexToDecimal("006100")
End With


With ValidationPending
    .Interior.Color = HexToDecimal("FFEB9C")
    .Font.Color = HexToDecimal("9C5700")
End With

With LockedForDeployment
    .Interior.Color = HexToDecimal("FFC7CE")
    .Font.Color = HexToDecimal("9C0006")
End With

With InDevelopment
    .Interior.Color = HexToDecimal("FFFFCC")
    .Font.Color = HexToDecimal("000000")
End With

End Sub

Function HexToDecimal(HexColor As String) As Long

    'Remove the # character if it exists
    HexColor = Replace(HexColor, "#", "")
    HexColor = Right(HexColor, 2) & Mid(HexColor, 3, 2) & Left(HexColor, 2)
    
    HexToDecimal = Val("&H" & HexColor)

End Function

Sub AddValidation()

    Dim rng As Range
    Dim formulaString As String


    Set rng = Selection
    ' formulaString = "='Menu Options'!$B$2:$B$5"
    
    formulaString = "='Menu Options'!$e$1:$e$3"
    
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=formulaString
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    
    End With

End Sub

Sub ReplaceSpaceWithUnderscore()

    Dim rng As Range
    Dim cell As Range
    
    Set rng = Selection
    
    For Each cell In rng
        cell.Value = Replace(cell.Value, " ", "_")
    Next cell

End Sub


Sub ReplaceUnderscoreWithSpace()

    Dim rng As Range
    Dim cell As Range
    
    Set rng = Selection
    
    For Each cell In rng
        cell.Value = Replace(cell.Value, "_", " ")
    Next cell

End Sub

Public Function GetHexBGColor()

    Dim HexColor As String
    
    HexColor = Right("000000" & Hex(ActiveCell.Interior.Color), 6)
    'Reverse the Hex code
    HexColor = Right(HexColor, 2) & Mid(HexColor, 3, 2) & Left(HexColor, 2)
    
    GetHexBGColor = "#" & HexColor
    
End Function

Public Function GetHexFontColor()

    Dim HexColor As String
    
    HexColor = Right("000000" & Hex(ActiveCell.Font.Color), 6)
    'Reverse the Hex code
    HexColor = Right(HexColor, 2) & Mid(HexColor, 3, 2) & Left(HexColor, 2)
    
    GetHexFontColor = "#" & HexColor
    
End Function

Sub InitializeGlobalHandler()
    Set GlobalHandler.App = Application
End Sub

Sub ClearEventHandler()
    Set GlobalHandler = Nothing
End Sub

Public Sub ShowCellInspector()
    InitializeGlobalHandler
    frmCellInfo.Show
End Sub

Public Sub FillWithStyle()

Dim ws              As Worksheet
Dim rng             As Range
Dim cell            As Range

Set ws = ActiveSheet
Set rng = Selection

    For Each cell In rng
        If cell.Value = "PASS" Then
            cell.Style = "Good"
        ElseIf cell.Value = "FAIL" Then
            cell.Style = "Bad"
        End If
    Next cell


End Sub

Sub SelectAllRows()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim finalRow As Long
    Dim finalColumn As Long

    Set wb = ActiveWorkbook
    Set ws = ActiveSheet

    finalRow = Cells(Rows.Count, 1).End(xlUp).Row
    finalColumn = Cells(1, Columns.Count).End(xlToLeft).Column
    ws.Range("A1:" & finalColumn & ":" & finalRow).Select
    
End Sub

Sub FormatLongNumbers()

    OptimizedMode True
    
    Dim rng As Range
    Dim format As String
    
    format = "0"
    Set rng = Selection
    
    With rng
        .NumberFormat = format
    End With
    
    OptimizedMode False

End Sub


Public Sub ChangeFontName(ByRef rng As Range, byRef fontName As String, )
    
    With rng
        rng.Font.Name = fontName
    End With

End Sub 