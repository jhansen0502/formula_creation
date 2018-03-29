Attribute VB_Name = "Module1"
Sub convert()
Dim i As Long
Dim j As Long
Dim k As Long
Dim this_sheet As String
Dim new_sheet As String
Dim source_sheet As String
Dim num_ingredients As Long
Dim column_headers(30) As String
Dim ingredients() As Variant
Dim item_number() As Long
Dim counter
Dim n As Long
Dim header_row As Long

this_sheet = "LIQUID MELT ITEMS"
source_sheet = "Source Data"

'create source data sheet and column headers
Sheets.Add(after:=Sheets(this_sheet)).Name = source_sheet
With Sheets(source_sheet)
    .Range("A1").Value = "FORMULA_NO"
    .Range("B1").Value = "FORMULA_VERS"
    .Range("C1").Value = "LINE_TYPE"
    .Range("D1").Value = "LINE_NO"
    .Range("E1").Value = "ITEM_NO"
    .Range("F1").Value = "QUANTITY_PRODUCT"
    .Range("G1").Value = "QUANTITY_INGREDIENT"
    .Range("H1").Value = "ITEM_UOM"
    .Range("I1").Value = "RELEASE_TYPE"
    .Range("J1").Value = "SCRAP_FACTOR"
    .Range("K1").Value = "SCALE_TYPE_DTL"
    .Range("L1").Value = "COST_ALLOC"
    .Range("M1").Value = "PHANTOM_TYPE"
    .Range("N1").Value = "CONTRIBUTE_TO_YIELD"
    .Range("O1").Value = "CONTRIBUTE_TO_STEP_QTY"
    .Range("P1").Value = "SCALE_MULTIPLE"
    .Range("Q1").Value = "SCALE_ROUNDING_VARIANCE"
    .Range("R1").Value = "ROUNDING_DIRECTION"
    .Range("S1").Value = "DTL_ATTRIBUTE_1"
    .Range("T1").Value = "DTL_ATTRIBUTE_2"
    .Range("U1").Value = "DTL_ATTRIBUTE_3"
    .Range("V1").Value = "DTL_ATTRIBUTE_4"
    .Range("W1").Value = "DTL_ATTRIBUTE_5"
    .Range("X1").Value = "ROLL_MILL"
    .Range("Y1").Value = "INPUT_ROSS_FORMULA"
    .Range("Z1").Value = "ORGN_CODE"
    .Range("AA1").Value = "OUTPUT_ROSS_FORMULA"
    .Range("AB1").Value = "FORMULA_ITEM"
    .Range("AC1").Value = "INPUT_ROSS_PART"
    .Range("AD1").Value = "OUTPUT_ROSS_PART"
    .Range("AE1").Value = "EBS_ITEM"
End With

For i = 0 To 30
    column_headers(i) = Cells(1, i + 1)
Next i

header_row = Sheets(this_sheet).UsedRange.Find(what:="Product Amount").Row
MsgBox (header_row)

For n = (header_row + 1) To Sheets(1).UsedRange.Rows.Count

counter = 0

'num_ingredients = Sheets(this_sheet).Range(Sheets(this_sheet).Cells(2, 4), _
'Sheets(this_sheet).Cells(2, Sheets(this_sheet).UsedRange.Columns.Count)).CountA

num_ingredients = Application.WorksheetFunction.Count(Sheets(1).Range(Sheets(1).Cells(n, 4), _
Sheets(1).Cells(n, Sheets(1).UsedRange.Columns.Count)))

Dim this_item As String
ReDim ingredients(num_ingredients)
ReDim item_number(num_ingredients)

this_item = Sheets(1).Cells(n, 1).Value

For j = 4 To Sheets(1).UsedRange.Columns.Count
        
    If Sheets(1).Cells(n, j) <> "" Then
        ingredients(counter) = Sheets(1).Cells(n, j)
        item_number(counter) = Sheets(1).Cells(1, j)
        counter = counter + 1
    End If
Next j

k = 0

Do While k < UBound(ingredients)
    
    nextrow = Sheets(source_sheet).UsedRange.Rows.Count + 1
    With Sheets(source_sheet)
        .Cells(nextrow, 5).Value = item_number(k)
        .Cells(nextrow, 4).Value = k + 1
        .Cells(nextrow, 3).Value = -1
        .Cells(nextrow, 2).Value = 0
        .Cells(nextrow, 1).Value = this_item & "-310"
        .Cells(nextrow, 6).Value = ""
        .Cells(nextrow, 7).Value = ingredients(k)
        .Cells(nextrow, 8).Value = "LBS"
        .Cells(nextrow, 9).Value = 1
        .Cells(nextrow, 10).Value = 0
        .Cells(nextrow, 11).Value = 1
        .Cells(nextrow, 12).Value = 1
        If item_number(k) <= 1048191 And item_number(k) >= 1048182 Then
            .Cells(nextrow, 13).Value = 0
        Else
            .Cells(nextrow, 13).Value = 1
        End If
        .Cells(nextrow, 14).Value = "Y"
        .Cells(nextrow, 15).Value = "Y"
        .Cells(nextrow, 18).Value = 1
    End With
    k = k + 1
Loop

Dim product_row As Long
product_row = Sheets(source_sheet).UsedRange.Rows.Count + 1

With Sheets(source_sheet)
    .Cells(product_row, 5).Value = this_item
    .Cells(product_row, 6).Value = 140000
    .Cells(product_row, 4).Value = 1
    .Cells(product_row, 3).Value = 1
    .Cells(product_row, 2).Value = 0
    .Cells(product_row, 1).Value = this_item & "-310"
    .Cells(product_row, 7).Value = ""
    .Cells(product_row, 8).Value = "LBS"
    .Cells(product_row, 9).Value = 1
    .Cells(product_row, 10).Value = 0
    .Cells(product_row, 11).Value = 1
    .Cells(product_row, 12).Value = 1
    .Cells(product_row, 13).Value = 1
    .Cells(product_row, 14).Value = "Y"
    .Cells(product_row, 15).Value = "Y"
    .Cells(product_row, 18).Value = 1
End With

Next n

End Sub
