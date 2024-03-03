Attribute VB_Name = "EMEG1"
Sub Table_Format_Update()
Dim totalrows As Long
Dim rng As range
Dim matchedcount As Integer
Dim row As Integer
Dim nextStr As String
Dim RefStr As String
Dim tableindex As Integer
Dim tbl As Word.Table
Dim iRow As Integer
' version 6 remove auto TU naming, keep it align with rfnsa
' version 7.1 two row bugfix
If Selection.Information(wdWithInTable) = True Then
    tableindex = ThisTableNumber
    totalrows = FindNumberofRows() ' get the total rows in selected table
Else
    MsgBox "Selection is not in a table."
End If

MsgBox "Table updates will take a while, Please Click OK to start..."

Set tbl = ActiveDocument.Tables(tableindex)
ActiveDocument.Tables(tableindex - 1).Columns(11).Delete 'delete column "port nubmer"
tbl.Columns(11).Delete 'delete column "port nubmer"
ActiveDocument.Tables(tableindex - 1).Columns(8).Width = 50
tbl.Columns(8).Width = 50
ActiveDocument.Tables(tableindex - 1).Columns(2).Width = tbl.Columns(1).Width
tbl.Columns(2).Width = tbl.Columns(1).Width

'replace text
Call Replace_text(tableindex)

'change font size to 2 in the table
Call Change_font_size(tableindex, 2)

RefStr = tbl.cell(1, 1).range.Text  ' reference string
nextStr = tbl.cell(2, 1).range.Text
matchedcount = 0
        row = 1
        '******place cursor at start
        Set rng = tbl.cell(row, 1).range
        rng.Collapse Direction:=wdCollapseStart
        rng.Select
 
Start:
        iRow = row
        Do While nextStr = RefStr 'revert to whole Str compare 2024.01.11
            matchedcount = matchedcount + 1
            If iRow + 1 >= totalrows Then
            Exit Do
            Else
            nextStr = tbl.cell(iRow + 1, 1).range.Text
            End If
            
            iRow = iRow + 1
        Loop
        If (mathchedcount + row < totalrows) Then
        RefStr = tbl.cell(matchedcount + row, 1).range.Text
        nextStr = tbl.cell(matchedcount + row, 1).range.Text
        End If
        
        If matchedcount >= 2 Then
           Call Merge_rows(tableindex, matchedcount)
        End If
        
        row = matchedcount + row
        matchedcount = 0
        
        If row >= totalrows Then
        'change back font size to 11
        Call Change_font_size(tableindex, 11)
        MsgBox "Table updates done"
        
        Exit Sub
        Else
         Set rng = tbl.cell(row, 1).range
            rng.Collapse Direction:=wdCollapseStart
            rng.Select
        GoTo Start
        End If
End Sub

Function FindNumberofRows() As Long
    Dim rows As Long
    FindNumberofRows = Selection.Information(wdMaximumNumberOfRows)
End Function

Function ThisTableNumber() As Integer
Dim CurrentSelection As Long
Dim T_Start As Long
Dim T_End As Long
Dim oTable As Table
Dim j As Long
CurrentSelection = Selection.range.Start
For Each oTable In ActiveDocument.Tables
T_Start = oTable.range.Start
T_End = oTable.range.End
j = j + 1
'ThisTableNumber = "Couldn't determine table number" ' Added error message
If CurrentSelection >= T_Start And _
CurrentSelection <= T_End Then ' added "="
ThisTableNumber = j
Exit For
End If
Next
End Function
Sub Replace_text(tableindex)
    Dim tbl As Table
    Dim cell As cell
    Dim row As Integer
    Dim vendor As String
    Dim height As Double
    
    Set tbl = ActiveDocument.Tables(tableindex)
    For Each cell In tbl.range.Cells
        Select Case cell.ColumnIndex
            Case 1 'column 1 process
              'copy column Diagram Ref to Owner Ref
              tbl.cell(cell.rowIndex, 2).range.Text = cell.range.Text
              
              If InStr(cell.range.Text, "-J") Then
                row = cell.rowIndex
                vendor = tbl.cell(row, 3).range.Text
                'Modify system/sector column to add vendor info
                If InStr(vendor, "Vodafone") Then
                    If InStr(tbl.cell(row, 10), "NR") Then
                    'nothing to do
                    tbl.cell(row, 10).range.Text = Replace(tbl.cell(row, 10).range.Text, "NR", "TPG NR")
                    'remove line breaks
                    tbl.cell(row, 10).range.Text = Replace(tbl.cell(row, 10).range.Text, Chr(13), "")
                    Else
                       tbl.cell(row, 10).range.Text = Replace(tbl.cell(row, 10).range.Text, "LTE", "TPG LTE")
                       tbl.cell(row, 10).range.Text = Replace(tbl.cell(row, 10).range.Text, "WCDMA", "TPG WCDMA")
                       tbl.cell(row, 10).range.Text = Replace(tbl.cell(row, 10).range.Text, "3.5GHz", "TPG NR 3500")
                       tbl.cell(row, 10).range.Text = Replace(tbl.cell(row, 10).range.Text, Chr(13), "")
                    End If
                    'tbl.cell(row, 9).Range.Text = Replace(tbl.cell(row, 9).Range.Text, "LTE", vendor & " NR/LTE")
                ElseIf InStr(vendor, "TPG") Then
                        If InStr(tbl.cell(row, 10), "NR") Then
                        'add operator name
                        tbl.cell(row, 10).range.Text = Replace(tbl.cell(row, 10).range.Text, "NR", "TPG NR")
                        tbl.cell(row, 10).range.Text = Replace(tbl.cell(row, 10).range.Text, Chr(13), "")
                        Else
                            'no need to add NR to TPG TUs, just keep it same as RFNSA
                                                        'tbl.cell(row, 10).Range.Text = Replace(tbl.cell(row, 10).Range.Text, "LTE", "NR/LTE")
                            tbl.cell(row, 10).range.Text = Replace(tbl.cell(row, 10).range.Text, Chr(13), "")
                        End If
                Else
                        'if vendor is Optus, just add vendor in front of the TU
                        tbl.cell(row, 10).range.Text = Replace(tbl.cell(row, 10).range.Text, tbl.cell(row, 10).range.Text, vendor & " " & tbl.cell(row, 10).range.Text)
                End If
                'Change column 2 to joint venture name
                tbl.cell(row, 3).range.Text = "Optus/ Vodafone Joint Venture"
              ElseIf InStr(cell.range.Text, "-V") Then
                row = cell.rowIndex
                vendor = tbl.cell(row, 3).range.Text
                If InStr(tbl.cell(row, 10), "NR") Then
                        'nothing to do
                        tbl.cell(row, 10).range.Text = Replace(tbl.cell(row, 10).range.Text, Chr(13), "")
                ElseIf InStr(tbl.cell(row, 10), "LTE") Then
                    'no need to add NR to TPG TUs, just keep it same as RFNSA
                                        'tbl.cell(row, 10).Range.Text = Replace(tbl.cell(row, 10).Range.Text, "LTE", "NR/LTE")
                    tbl.cell(row, 10).range.Text = Replace(tbl.cell(row, 10).range.Text, Chr(13), "")
                End If
                tbl.cell(row, 10).range.Text = Replace(tbl.cell(row, 10).range.Text, Chr(13), "")
              End If
        
            Case 10
                If InStr(cell.range.Text, "3.64GHz") Then
                    cell.range.Text = Replace(cell.range.Text, "3.64GHz", "NR 3500")
                ElseIf InStr(cell.range.Text, "3.5GHz") Then
                'Debug.Print cell.Range.Text
                    cell.range.Text = Replace(cell.range.Text, "3.5GHz", "NR 3500")
                ElseIf InStr(cell.range.Text, "3.56GHz") Then
                    cell.range.Text = Replace(cell.range.Text, "3.56GHz", "NR 3500")
                ElseIf InStr(cell.range.Text, "Wimax 2300") Then
                    cell.range.Text = Replace(cell.range.Text, "3.64GHz", "NR 3500")
                End If
                'remove line breaks
                cell.range.Text = Replace(cell.range.Text, Chr(13), "")
            
            Case 11
                If InStr(cell.range.Text, "+0") Then
                    cell.range.Text = Replace(cell.range.Text, "+0", "")
                End If
                If InStr(cell.range.Text, "0+0+0+0+") Then
                    cell.range.Text = Replace(cell.range.Text, "0+0+0+0+", "")
                End If
                
                If InStr(cell.range.Text, "0+0+0+") Then
                    cell.range.Text = Replace(cell.range.Text, "0+0+0+", "")
                End If
                
                If InStr(cell.range.Text, "0+0+") Then
                    cell.range.Text = Replace(cell.range.Text, "0+0+", "")
                End If
                
                'remove line breaks
                cell.range.Text = Replace(cell.range.Text, Chr(13), "")
                
        End Select
    Next cell
End Sub
Sub Merge_rows(tableindex, matchedcount)
    Dim row_count As Integer
    Dim n, m, x, y, i As Integer
    Dim selectedRange As range
    Dim tbl As Word.Table
    Dim rowIndex, colIndex As Integer
    Dim newCell As cell
    Set tbl = ActiveDocument.Tables(tableindex)
    
    n = matchedcount
    m = matchedcount - 1
    
    Selection.MoveDown
    
    
    If n = 2 Then
       'Selection.MoveRight Unit:=wdCharacter, Count:=6, Extend:=wdExtend
       Set selectedRange = Selection.range
       rowIndex = Selection.Cells(1).rowIndex
       colIndex = 7 ' from column diagram ref to Column Mech Tilt
       For i = 2 To 7
        Set newCell = tbl.cell(rowIndex, i)
        selectedRange.Expand Unit:=wdCell
        selectedRange.SetRange Start:=selectedRange.Start, End:=newCell.range.End
       Next i
       selectedRange.Select
    Else
       x = n - 2
       Selection.MoveDown Unit:=wdLine, Count:=x, Extend:=wdExtend
       Selection.MoveRight Unit:=wdCharacter, Count:=6, Extend:=wdExtend
    End If
       
    Selection.Delete
    Selection.MoveUp
    Selection.MoveDown Unit:=wdLine, Count:=m, Extend:=wdExtend
    Selection.Cells.Merge
    
    For y = 1 To 6
        Selection.MoveRight
        Selection.MoveDown Count:=m, Extend:=wdExtend
        Selection.Cells.Merge
    Next y
               
End Sub

Sub Change_font_size(tableindex, size)
Dim fontSize As Integer
Dim tbl As Table
Dim cell As cell

fontSize = size
Set tbl = ActiveDocument.Tables(tableindex)
    For Each cell In tbl.range.Cells
        cell.range.Font.size = fontSize
    Next cell

End Sub



