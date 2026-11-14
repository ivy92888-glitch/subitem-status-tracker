' =============================================================
' SubItemTally — WPS / Excel VBA
' Groups sub-items (boxes, batches, line items, ...) under a
' master number and tallies how many are processed vs.
' unprocessed, then writes a one-click summary sheet.
'
' MIT License — Copyright (c) 2026 ivy92888-glitch
' =============================================================
Option Explicit

' ---- Configuration: adjust these to fit your own data ----
Private Const SRC_SHEET      As String = ""           ' "" = use the active sheet; otherwise the sheet name
Private Const SUBITEM_COL    As String = "A"          ' column that holds the sub-item code
Private Const REF_COL        As String = "D"          ' column that holds the linked reference (blank = unprocessed)
Private Const SEQ_SEPARATOR  As String = "-"          ' splits <MasterNo><Seq> inside a sub-item code
Private Const RESULT_SHEET   As String = "Summary"    ' name of the output worksheet
Private Const ONLY_LIST_DONE As Boolean = True       ' True = list only master numbers with >=1 processed sub-item

Sub TallySubItems()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim dict As Object, baseDict As Object
    Dim baseNo As String, subItem As String, refNo As String
    Dim resultWs As Worksheet, outputRow As Long

    ' Bind the source worksheet
    If SRC_SHEET = "" Then
        Set ws = ActiveSheet
    Else
        Set ws = ThisWorkbook.Worksheets(SRC_SHEET)
    End If
    lastRow = ws.Cells(ws.Rows.Count, SUBITEM_COL).End(xlUp).Row

    ' Dictionaries: baseDict = master numbers; dict = sub-item processing status
    Set dict = CreateObject("Scripting.Dictionary")
    Set baseDict = CreateObject("Scripting.Dictionary")

    ' Step 1: scan every data row and build the index
    For i = 2 To lastRow
        subItem = Trim(ws.Cells(i, SUBITEM_COL).Value)
        refNo = Trim(ws.Cells(i, REF_COL).Value)

        ' Derive the master number by stripping the sequence part
        Dim pos As Long
        pos = InStr(subItem, SEQ_SEPARATOR)
        If pos > 0 Then
            baseNo = Left(subItem, pos - 1)
        Else
            baseNo = subItem
        End If

        If Not baseDict.exists(baseNo) Then baseDict.Add baseNo, 1
        If Not dict.exists(subItem) Then _
            dict.Add subItem, IIf(refNo <> "", "Processed", "Unprocessed")
    Next i

    ' Step 2: create the summary sheet, or clear it if it already exists
    On Error Resume Next
    Set resultWs = ThisWorkbook.Worksheets(RESULT_SHEET)
    If resultWs Is Nothing Then
        Set resultWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        resultWs.Name = RESULT_SHEET
    Else
        resultWs.Cells.Clear
    End If
    On Error GoTo 0

    With resultWs
        .Range("A1:D1").Value = Array("MasterNo", "Total", "Processed", "Unprocessed")
        .Rows(1).Font.Bold = True
        outputRow = 2
    End With

    ' Step 3: aggregate the counts for each master number
    Dim totalBoxes As Long, processed As Long, unprocessed As Long
    Dim key As Variant, boxKey As Variant
    For Each key In baseDict.keys
        totalBoxes = 0: processed = 0: unprocessed = 0
        For Each boxKey In dict.keys
            If Left(boxKey, Len(key)) = key Then
                totalBoxes = totalBoxes + 1
                If dict(boxKey) = "Processed" Then
                    processed = processed + 1
                Else
                    unprocessed = unprocessed + 1
                End If
            End If
        Next boxKey

        ' Output (skip master numbers that have no processed sub-item when ONLY_LIST_DONE = True)
        If (Not ONLY_LIST_DONE) Or (processed > 0) Then
            With resultWs
                .Cells(outputRow, 1).Value = key
                .Cells(outputRow, 2).Value = totalBoxes
                .Cells(outputRow, 3).Value = processed
                .Cells(outputRow, 4).Value = unprocessed
                ' Highlight the unprocessed count for quick attention
                If unprocessed > 0 Then _
                    .Cells(outputRow, 4).Interior.Color = RGB(255, 200, 200)
            End With
            outputRow = outputRow + 1
        End If
    Next key

    ' Step 4: format the summary sheet
    With resultWs
        .Columns("A:D").AutoFit
        With .Range("A1:D" & outputRow - 1).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        .Range("A1:D1").AutoFilter
        .Activate
        .Rows(2).Select
        ActiveWindow.FreezePanes = True
    End With

    MsgBox "Done. Results saved to [" & RESULT_SHEET & "] worksheet.", vbInformation
End Sub
