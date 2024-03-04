Attribute VB_Name = "UnzipModule21"
Sub Equipment_Table_Consistency_Check()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' PRD Equipment table and RFNSA STAD consistency check v1.0''
'' email: yong.zhang@skyaus.com.au                          ''
'' 25/01/2024                                               ''
'' Stictly for SkyAus internal use only                     ''
'' Usage: select 1st column/2nd row of the equipment table  ''
'' then run the Macro
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'v1.3 power comparison per port
'v1.4 saveas file name added time stamp
'v1.41 antenna id type check 2024.03.05
Dim full_zip_filename As Variant
Dim FileDialog As FileDialog
Dim zip_filepath As Variant
Dim rfnsa_id As String
Dim text_file_wild_card, csv_file_name As String
Dim stad_csv As String
Dim equip_csv As String
Dim text_file As Variant
Dim dtcode As Double

Dim totalrows As Long
Dim rng As Range
Dim matchedcount As Integer
Dim row As Integer
Dim nextStr As String
Dim RefStr As String
Dim tableindex As Integer
Dim tbl As Word.Table
Dim iRow As Integer
'Dim RFNSAID As String
'remove auto TU naming, keep it align with rfnsa
'put cursor in the first column second row of the equipment table
If Selection.Information(wdWithInTable) = True Then
    tableindex = ThisTableNumber
    totalrows = FindNumberofRows() ' get the total rows in selected table
Else
    MsgBox "Selection is not in a table. Make sure cursor is in the 2nd row/1st column of the equipment table."
    Exit Sub
End If

MsgBox "Equipment table selected, Please select STAD file."

Application.ScreenUpdating = False

Set FileDialog = Application.FileDialog(msoFileDialogOpen)
With FileDialog
    .Title = "Select STAD Zip File"
    .Filters.Clear
    .Filters.Add "Zip Files", "*.zip" 'open STAD *.zip file
        
    ' Show the dialog and store the selected file(s) in the array
    If .Show = -1 Then
        For Each SelectedFile In .SelectedItems
            zip_filepath = FileDialog.InitialFileName
            full_zip_filename = SelectedFile
        Next SelectedFile
    Else
        ' User canceled the dialog
        MsgBox "Operation canceled by user."
    End If
End With

If full_zip_filename = False Then
    MsgBox ("Invalid File")
    Exit Sub
End If

'retrieve rfnsa id
rfnsa_id = Mid(full_zip_filename, Len(full_zip_filename) - 37, 7)
text_file_wild_card = rfnsa_id & "_" & _
                        Mid(full_zip_filename, Len(full_zip_filename) - 17, 8) & "*.txt"

'unzip STAD file
Application.StatusBar = "Unzip STAD started..."
Call UnzipAFile(full_zip_filename, zip_filepath)

text_file = Dir(zip_filepath & text_file_wild_card)
csv_file_name = zip_filepath & _
                rfnsa_id & "_" & _
                Mid(full_zip_filename, Len(full_zip_filename) - 17, 12)

stad_csv = csv_file_name & ".csv"
text_file = zip_filepath + "\" + text_file

'convert RFNSA STAD text file into csv file
Call ConvertToCSV(text_file, stad_csv)
 Kill (text_file)

'export the equipment table into csv file
Dim doc As Word.Document
Set doc = ActiveDocument
Dim docPath As String
docPath = GetCurrentUserDownloadsPath()
'Dim equip_csv As String
equip_csv = docPath + "/" + rfnsa_id + "_equipment_table.csv"

Application.StatusBar = "Saving Data to CSV file..."
Call ExportTableToCSV(tableindex, equip_csv)

Application.StatusBar = "Comparison started..."
Call ImportCSVFilesAndComparison(stad_csv, equip_csv)
Kill (stad_csv)
Kill (equip_csv)

End Sub
Sub ExportTableToCSV(tableindex As Integer, fileName As String)
    Dim doc As Word.Document
    Dim tbl As Table
    Dim fso As Object
    Dim csvFile As Object
    Dim cellText As String
    Dim rowText As String
    Dim colCount As Integer
    Dim rowCount As Integer
    Dim filePath As String
      
    Set doc = ActiveDocument
    
    Set tbl = doc.Tables(tableindex)
      
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set csvFile = fso.CreateTextFile(fileName, True)
    If InStr(tbl.cell(1, 1).Range.Text, "Diagram") Then
        'nothing
    Else
        csvFile.WriteLine "Diagram Ref,Owner Ref,Owner,Type/Make/Model,Height(m),Bearing(°),Mech.Tilt(°),Elec.Tilt(°),Pol,System,Port Number,Power(Watts)"
    End If
    
    For rowCount = 1 To tbl.rows.Count
        rowText = ""
        'Debug.Print tbl.Columns.Count
        For colCount = 1 To tbl.Columns.Count
            cellText = Replace(tbl.cell(rowCount, colCount).Range.Text, ",", ";")
            'Debug.Print rowCount, colCount
            'Debug.Print cellText
            rowText = rowText & cellText & ","
        Next colCount
    
        rowText = Replace(Replace(rowText, "", ""), Chr(13), "") 'replace space and character(13)
        csvFile.WriteLine rowText
    Next rowCount
      
    csvFile.Close
    Set csvFile = Nothing
    Set fso = Nothing

End Sub
    
Sub UnzipAFile(zippedFileFullName As Variant, unzipToPath As Variant)

Dim ShellApp As Object

'Copy the files & folders from the zip into a folder
Set ShellApp = CreateObject("Shell.Application")
ShellApp.Namespace(unzipToPath).CopyHere ShellApp.Namespace(zippedFileFullName).items

End Sub
Sub ConvertToCSV(textfile As Variant, csvFileName As String)
    Dim tabFile As String
    Dim csvFile As String
    Dim inputData As String
    Dim outputData As String
    
    ' Specify the file paths
    tabFile = textfile
    csvFile = csvFileName
    
    ' Open the input file for reading
    Open tabFile For Input As #1
    
    ' Open the output file for writing
    Open csvFile For Output As #2
    
    ' Read each line from the tab-separated file, replace tabs with commas, and write to the CSV file
    Do Until EOF(1)
        Line Input #1, inputData
        outputData = Replace(inputData, vbTab, ",")
        Print #2, outputData
    Loop
    
    ' Close the file handles
    Close #1
    Close #2
    
    'MsgBox "Conversion completed successfully.", vbInformation

End Sub

Function FindNumberofRows() As Long

Dim rows As Long
'MsgBox (Selection.Information(wdMaximumNumberOfRows))

FindNumberofRows = Selection.Information(wdMaximumNumberOfRows)
'MsgBox (rows)
End Function

Function ThisTableNumber() As Integer
Dim CurrentSelection As Long
Dim T_Start As Long
Dim T_End As Long
Dim oTable As Table
Dim j As Long
CurrentSelection = Selection.Range.Start
For Each oTable In ActiveDocument.Tables
T_Start = oTable.Range.Start
T_End = oTable.Range.End
j = j + 1
'ThisTableNumber = "Couldn't determine table number" ' Added error message
If CurrentSelection >= T_Start And _
CurrentSelection <= T_End Then ' added "="
ThisTableNumber = j
Exit For
End If
Next
End Function


Sub ImportCSVFilesAndComparison(stad_csv As String, equip_csv As String)
    Dim excelApp As Object
    Dim wb As Excel.Workbook
    Dim ws1 As Excel.Worksheet
    Dim ws2 As Excel.Worksheet
    Dim docPath As String
    Dim rfnsa_id As String
    Dim cell1 As Range
    Dim cell2 As Range
    Dim lookupValue, a As Variant
    Dim foundCell As Object
    Dim usedRow As Long
    Dim i As Integer
    Dim combined_value As String
    Dim filterColumn As Integer
    Dim criteria As Variant
    Dim rng As Range
    Dim s As String
    Dim carrier As String
    Dim band As String
    Dim columnToReturn As Integer
    Dim result As String 'Variant
    Dim lookupRange As Range
    Dim position As Long
    Dim port_item As Variant
	Dim currentDate As Date
    Dim currentHour As Integer
    Dim currentMinute As Integer
    
    result = "" 'lookup result set to empty
    
    docPath = Left(stad_csv, Len(stad_csv) - 24)
    rfnsa_id = Left(Right(stad_csv, 24), 7)
    
    Set excelApp = CreateObject("Excel.Application")
    ' Make Excel visible (optional, depending on your needs)
    ' Create a new workbook
    Set wb = excelApp.Workbooks.Add
    ' Maximize the workbook window
    wb.Application.WindowState = xlMaximized
    
    ' Add worksheets to the new workbook
    Set ws1 = wb.Worksheets.Add
    ws1.Name = "STAD"
    Set ws2 = wb.Worksheets.Add
    ws2.Name = "Equipment_Table"
    wb.Worksheets("Sheet1").Delete
    
    
    ' Import data from stad_csv to Sheet1
    With ws1.QueryTables.Add(Connection:="TEXT;" & stad_csv, Destination:=ws1.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileCommaDelimiter = True ' Assuming CSV is comma-delimited
        .Refresh
    End With
    ' Close the connection to the external data
    ws1.QueryTables(1).Delete
    
    ws1.Cells(1, 1).ColumnWidth = 20
    ws1.Cells(1, 5).ColumnWidth = 26
    ws1.Cells(1, 11).ColumnWidth = 8
    ws1.Cells(1, 12).ColumnWidth = 8
    ws1.Cells(1, 13).ColumnWidth = 8
    ws1.Cells(1, 14).ColumnWidth = 8
    ws1.Cells(1, 17).ColumnWidth = 8
    ws1.Select
    ws1.Range("E2").Select
    excelApp.ActiveWindow.FreezePanes = True
    excelApp.ActiveWindow.Zoom = 80
    
    ' Import data from equip_csv to Sheet2
    With ws2.QueryTables.Add(Connection:="TEXT;" & equip_csv, Destination:=ws2.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileCommaDelimiter = True ' Assuming CSV is comma-delimited
        .Refresh
    End With
    ' Close the connection to the external data
    ws2.QueryTables(1).Delete
    
    'Sheet Equipment_Table Titls
    'background color change
    'cell width changes,wraptext,alignment
    ws2.Cells(1, 13) = "STAD:System"
    ws2.Cells(1, 14) = "Height(m)"
    ws2.Cells(1, 15) = "Bearing(°)"
    ws2.Cells(1, 16) = "Mech.Tilt(°)"
    ws2.Cells(1, 17) = "Elec.Tilt(°)"
    ws2.Cells(1, 18) = "Power(Watts -> dBm)"
    ws2.Cells(1, 19) = "Antenna Model"
    ws2.Cells(1, 20) = "Port Number"
    
    ws2.Cells(1, 13).Interior.Color = RGB(255, 255, 0)
    ws2.Cells(1, 14).Interior.Color = RGB(255, 255, 0)
    ws2.Cells(1, 15).Interior.Color = RGB(255, 255, 0)
    ws2.Cells(1, 16).Interior.Color = RGB(255, 255, 0)
    ws2.Cells(1, 17).Interior.Color = RGB(255, 255, 0)
    ws2.Cells(1, 18).Interior.Color = RGB(255, 255, 0)
    ws2.Cells(1, 19).Interior.Color = RGB(255, 255, 0)
    ws2.Cells(1, 20).Interior.Color = RGB(255, 255, 0)
    
    ws2.Cells(1, 5).Interior.Color = RGB(0, 255, 0)
    ws2.Cells(1, 6).Interior.Color = RGB(0, 255, 0)
    ws2.Cells(1, 7).Interior.Color = RGB(0, 255, 0)
    ws2.Cells(1, 8).Interior.Color = RGB(0, 255, 0)
    ws2.Cells(1, 9).Interior.Color = RGB(0, 255, 0)
    ws2.Cells(1, 10).Interior.Color = RGB(0, 255, 0)
    ws2.Cells(1, 11).Interior.Color = RGB(0, 255, 0)
    ws2.Cells(1, 12).Interior.Color = RGB(0, 255, 0)
    
    ws2.Cells(1, 1).ColumnWidth = 8
    ws2.Cells(1, 2).ColumnWidth = 8
    ws2.Cells(1, 3).ColumnWidth = 16
    
    ws2.Cells(1, 7).ColumnWidth = 7
    ws2.Cells(1, 9).ColumnWidth = 6
    ws2.Cells(1, 11).ColumnWidth = 25
    ws2.Cells(1, 12).ColumnWidth = 25
    ws2.Cells(1, 13).ColumnWidth = 38
    ws2.Cells(1, 14).ColumnWidth = 8
    ws2.Cells(1, 15).ColumnWidth = 8
    ws2.Cells(1, 16).ColumnWidth = 7
    ws2.Cells(1, 17).ColumnWidth = 9
    ws2.Cells(1, 18).ColumnWidth = ws2.Cells(1, 12).ColumnWidth
    ws2.Cells(1, 19).ColumnWidth = 42
    ws2.Cells(1, 20).ColumnWidth = 18
    
    ws2.rows(1).WrapText = True
    ws2.Select
    ws2.Range("E2").Select
    excelApp.ActiveWindow.FreezePanes = True
    excelApp.ActiveWindow.Zoom = 70
    'ws2.rows(1).HorizontalAlignment = xlCenter
    'ws2.rows(1).VerticalAlignment = xlCenter
    ws2.UsedRange.HorizontalAlignment = xlCenter
    ws2.UsedRange.VerticalAlignment = xlCenter
    Application.StatusBar = "Data imported to excel..."
    usedRow = ws1.UsedRange.rows.Count
    
    'remove Proposed rows of STAD sheet
    filterColumn = 4 'Existing/Proposed column
    criteria = "Proposed"
      
    For i = usedRow To 2 Step -1 ' Assuming data starts from row 2
        ' Check if the cell in the filter column meets the criteria
        If ws1.Cells(i, filterColumn).Value = criteria Then
            ' Delete the entire row
            ws1.rows(i).Delete
        End If
    Next i
   
    'Update sheet STAD column A to anntenna id + system
    'and change format to match vlookup from equipment_table sheet
    usedRow = ws1.UsedRange.rows.Count
    bandarray = Array("700", "850", "900", "1800", "2100", "2300", _
    "LTE 2600", "LTE2600", "3500", "3.5G", "3.64G", "3.56G", "NR26000", "27GHz")
    carrierarray = Array("Telstra", "Vodafone", "Optus", "3GIS")
    'change STAD sheet A1 title to Index
    ws1.Cells(1, 1) = "Index"
    For i = 2 To usedRow:
        band = ""
        For Each a In bandarray:
            If InStr(ws1.Cells(i, 5), a) Then
                band = a
                Exit For
            End If
        Next a
        
        If InStr(ws1.Cells(i, 5), "WCDMA") Then 'put W in font of WCDMA Tus
            band = "W" + band
        End If
        
        'find carrier information from column 5
        carrier = ""
        For Each a In carrierarray:
            If InStr(ws1.Cells(i, 5), a) Then
                carrier = a
                Exit For
            End If
        Next a
		'change the cell format to text if it contains a number
        '2024.03.05
        If IsNumeric(ws1.Cells(i, 8)) Then
            ws1.Cells(i, 8).NumberFormat = "@"
            'convert the cell value to string
            ws1.Cells(i, 8) = CStr(ws1.Cells(i, 8))
        End If
        '2024.03.05
        combined_value = ws1.Cells(i, 8) + "_" + carrier + "_" + band
       'Debug.Print combined_value
        ws1.Cells(i, 1) = combined_value
    Next i
    'Update sheet STAD end
    
    'equipment table sheet update start
    'create lookup value: antenna id + carrier + band
    'get the number of rows
    usedRow = ws2.UsedRange.rows.Count
    For i = 2 To usedRow:
        band = ""
        For Each a In bandarray:
            If InStr(ws2.Cells(i, 10), a) Then
                band = a
                Exit For
            End If
        Next a
        'adjust lookup value to match STAD index
        If InStr(ws2.Cells(i, 10), "WCDMA") Then 'put W in font of WCDMA Tus to be clear with LTE TUs
            band = "W" + band
        End If
        
        If InStr(ws2.Cells(i, 10), "LTE 2600") Then 'change LTE 2600 to LTE2600
            band = Replace(band, " ", "")
        End If
        If InStr(ws2.Cells(i, 10), "27GHz") Then 'change 27GHz to NR26000
            band = Replace(band, "27GHz", "NR26000")
        End If
        
        carrier = ""
        For Each a In carrierarray:
            If InStr(ws2.Cells(i, 3), a) Then
                carrier = a
                Exit For
            End If
        Next a
        
        'some band need to be adjusted
        band = Replace(band, "3.5G", "3500")
        band = Replace(band, "3.56G", "3500")
        band = Replace(band, "3.64G", "3500")
        'lookupvalue:antenna id + carrier + band
        lookupValue = ws2.Cells(i, 1) + "_" + carrier + "_" + band
        
        'lookup start
        'Set the column number to return
        Application.StatusBar = "Data Comparison started..."
        columnToReturn = 18 ' 1st power column in STAD
        On Error Resume Next
        result = Excel.Application.WorksheetFunction.VLookup(lookupValue, ws1.UsedRange, columnToReturn, 0)
        ws2.Cells(i, 13) = Excel.Application.WorksheetFunction.VLookup(lookupValue, ws1.UsedRange, 5, 0) 'column 5 system
        ws2.Cells(i, 19) = Excel.Application.WorksheetFunction.VLookup(lookupValue, ws1.UsedRange, 9, 0) 'column 9 antenna model
        ws2.Cells(i, 20) = result 'Excel.Application.WorksheetFunction.VLookup(lookupValue, ws1.UsedRange, 17, 0) 'column 17 port number
        On Error GoTo 0
        If Len(result) = 0 Then
            'ws2.Cells(i, 13) = "N/A"
            ws2.Cells(i, 18) = "N/A"
            'ws2.Cells(i, 19) = "N/A"
            'ws2.Cells(i, 20) = "N/A"
            'ws2.Cells(i, 13).Interior.Color = RGB(255, 255, 0)
            ws2.Cells(i, 18).Interior.Color = RGB(255, 255, 0)
            'ws2.Cells(i, 19).Interior.Color = RGB(255, 255, 0)
            'ws2.Cells(i, 20).Interior.Color = RGB(255, 255, 0)
            
            ws2.Cells(i, 12).Interior.Color = RGB(255, 255, 0)
            ws2.Cells(i, 11).Interior.Color = RGB(255, 255, 0)
            ws2.Cells(i, 1).Interior.Color = RGB(255, 255, 0)
            'GoTo SkipIteration
        Else
            result = power_adjustment(result)
            ws2.Cells(i, 18) = result
            'split power value divided by +, then merge non-zero values
            ''split power value divided by +, then merge non-zero values Done...
            'change background color to yellow and font color to red if power between prd and stad mismatched.
            result = ws2.Cells(i, 12)
            result = RemoveZeroInPower(result)
            ws2.Cells(i, 12) = result
            'v1.3 power comparison per port
            If Not port_power_match(ws2.Cells(i, 12), ws2.Cells(i, 18)) Then 'ws2.Cells(i, 18) <> ws2.Cells(i, 12) Then 'total power comparison
                ws2.Cells(i, 18).Interior.Color = RGB(255, 255, 0)
                ws2.Cells(i, 12).Interior.Color = RGB(255, 255, 0)
                ws2.Cells(i, 18).Font.Color = RGB(255, 0, 0)
                ws2.Cells(i, 12).Font.Color = RGB(255, 0, 0)
                'ws2.Cells(i, 10).Interior.Color = RGB(255, 255, 0)
                'ws2.Cells(i, 10).Font.Color = RGB(255, 0, 0)
                ws2.Cells(i, 1).Interior.Color = RGB(255, 255, 0)
            End If
            
            result = ""
        End If
        
        result = ""
        columnToReturn = 14 ' 2nd height column in STAD
        On Error Resume Next
        result = Excel.Application.WorksheetFunction.VLookup(lookupValue, ws1.UsedRange, columnToReturn, 0)
        On Error GoTo 0
        If Len(result) = 0 Then
            ws2.Cells(i, 14) = "N/A"
            ws2.Cells(i, 14).Interior.Color = RGB(255, 255, 0)
            ws2.Cells(i, 5).Interior.Color = RGB(255, 255, 0)
            ws2.Cells(i, 1).Interior.Color = RGB(255, 255, 0)
            ws2.Cells(i, 10).Interior.Color = RGB(255, 255, 0)
        Else
            ws2.Cells(i, 14) = result
            If ws2.Cells(i, 14) <> ws2.Cells(i, 5) Then
                ws2.Cells(i, 14).Interior.Color = RGB(255, 255, 0)
                ws2.Cells(i, 5).Interior.Color = RGB(255, 255, 0)
                ws2.Cells(i, 14).Font.Color = RGB(255, 0, 0)
                ws2.Cells(i, 5).Font.Color = RGB(255, 0, 0)
                'ws2.Cells(i, 10).Interior.Color = RGB(255, 255, 0)
                'ws2.Cells(i, 10).Font.Color = RGB(255, 0, 0)
                ws2.Cells(i, 1).Interior.Color = RGB(255, 255, 0)
            End If
            result = ""
        End If
        
        result = ""
        columnToReturn = 11 ' 3rd bearing column in STAD
        On Error Resume Next
        result = Excel.Application.WorksheetFunction.VLookup(lookupValue, ws1.UsedRange, columnToReturn, 0)
        On Error GoTo 0
        If Len(result) = 0 Then
            ws2.Cells(i, 15) = "N/A"
            ws2.Cells(i, 15).Interior.Color = RGB(255, 255, 0)
            ws2.Cells(i, 6).Interior.Color = RGB(255, 255, 0)
            ws2.Cells(i, 1).Interior.Color = RGB(255, 255, 0)
            ws2.Cells(i, 10).Interior.Color = RGB(255, 255, 0)
            
        Else
            ws2.Cells(i, 15) = result
            If ws2.Cells(i, 15) <> ws2.Cells(i, 6) Then
                ws2.Cells(i, 15).Interior.Color = RGB(255, 255, 0)
                ws2.Cells(i, 6).Interior.Color = RGB(255, 255, 0)
                ws2.Cells(i, 15).Font.Color = RGB(255, 0, 0)
                ws2.Cells(i, 6).Font.Color = RGB(255, 0, 0)
                'ws2.Cells(i, 10).Interior.Color = RGB(255, 255, 0)
                'ws2.Cells(i, 10).Font.Color = RGB(255, 0, 0)
                ws2.Cells(i, 1).Interior.Color = RGB(255, 255, 0)
            End If
            result = ""
        End If
        
        result = ""
        columnToReturn = 12 ' 4th m-tilt column in STAD
        On Error Resume Next
        result = Excel.Application.WorksheetFunction.VLookup(lookupValue, ws1.UsedRange, columnToReturn, 0)
        On Error GoTo 0
        If Len(result) = 0 Then
            ws2.Cells(i, 16) = "N/A"
            ws2.Cells(i, 16).Interior.Color = RGB(255, 255, 0)
            ws2.Cells(i, 7).Interior.Color = RGB(255, 255, 0)
            ws2.Cells(i, 1).Interior.Color = RGB(255, 255, 0)
            ws2.Cells(i, 10).Interior.Color = RGB(255, 255, 0)
        Else
            ws2.Cells(i, 16) = result
            If ws2.Cells(i, 16) <> ws2.Cells(i, 7) Then
                ws2.Cells(i, 16).Interior.Color = RGB(255, 255, 0)
                ws2.Cells(i, 7).Interior.Color = RGB(255, 255, 0)
                ws2.Cells(i, 16).Font.Color = RGB(255, 0, 0)
                ws2.Cells(i, 7).Font.Color = RGB(255, 0, 0)
                'ws2.Cells(i, 10).Interior.Color = RGB(255, 255, 0)
                'ws2.Cells(i, 10).Font.Color = RGB(255, 0, 0)
                ws2.Cells(i, 1).Interior.Color = RGB(255, 255, 0)
            End If
            result = ""
        End If
        
        'e-tilt check start
        result = ""
        columnToReturn = 13 ' 5th e-tilt column in STAD
        On Error Resume Next
        result = Excel.Application.WorksheetFunction.VLookup(lookupValue, ws1.UsedRange, columnToReturn, 0)
        On Error GoTo 0
        If Len(result) = 0 Then
            ws2.Cells(i, 17) = "N/A"
            ws2.Cells(i, 17).Interior.Color = RGB(255, 255, 0)
            ws2.Cells(i, 8).Interior.Color = RGB(255, 255, 0)
            ws2.Cells(i, 1).Interior.Color = RGB(255, 255, 0)
            ws2.Cells(i, 10).Interior.Color = RGB(255, 255, 0)
        Else
            position = InStr(result, "(")
            result = Left(result, position - 2)
            ws2.Cells(i, 17) = result
            If ws2.Cells(i, 17) <> ws2.Cells(i, 8) Then
                ws2.Cells(i, 17).Interior.Color = RGB(255, 255, 0)
                ws2.Cells(i, 8).Interior.Color = RGB(255, 255, 0)
                ws2.Cells(i, 17).Font.Color = RGB(255, 0, 0)
                ws2.Cells(i, 8).Font.Color = RGB(255, 0, 0)
                'ws2.Cells(i, 10).Interior.Color = RGB(255, 255, 0)
                'ws2.Cells(i, 10).Font.Color = RGB(255, 0, 0)
                ws2.Cells(i, 1).Interior.Color = RGB(255, 255, 0)
            End If
            result = ""
        End If

        'SYSTEM name check start
        'STAD System name divied by "/", then remove"[Macro]","[IBC]","- LOCKED",REMOVE SPACE
        'EMEG System name remove space
        Dim item_system As Variant
        result = ws2.Cells(i, 13)
        If result <> "" Then
            item_system = Split(result, "/")
            result = Replace(result, item_system(0), "") 'remove 1st part
            result = Right(result, Len(result) - 1) 'remove "/" at the beginning
        End If
    
        Dim stad_system As String
        stad_system = Replace(Replace(Replace(Replace(result, "[Macro]", ""), "[IBC]", ""), "- LOCKED", ""), " ", "")
        'Debug.Print stad_system
        If stad_system <> Replace(ws2.Cells(i, 10), " ", "") Then
            ws2.Cells(i, 1).Interior.Color = RGB(255, 255, 0)
            ws2.Cells(i, 10).Interior.Color = RGB(255, 255, 0)
            ws2.Cells(i, 10).Font.Color = RGB(255, 0, 0)
            ws2.Cells(i, 13).Interior.Color = RGB(255, 255, 0)
            ws2.Cells(i, 13).Font.Color = RGB(255, 0, 0)
        End If
        result = ""
        'SYSTEM name check end
        'SYSTEM name check end
        'SYSTEM name check end
    
        'Antenna model check start
        'diviced by "/", get 3rd part antenna model
        Dim item_antenna As Variant
        result = ws2.Cells(i, 4) ' get antenna model info from equip table
        item_antenna = Split(result, "/")
        result = item_antenna(UBound(item_antenna)) ' last part is the antenna model
    
        If result <> ws2.Cells(i, 19) Then ' column 19 is the STAD antenna model
            'if not matched, hightlight it
            ws2.Cells(i, 4).Interior.Color = RGB(255, 255, 0)
            ws2.Cells(i, 4).Font.Color = RGB(255, 0, 0)
            ws2.Cells(i, 19).Interior.Color = RGB(255, 255, 0)
            ws2.Cells(i, 19).Font.Color = RGB(255, 0, 0)
        End If
        result = ""
        'Antenna model check end
        
        'Port number from STAD power
        Dim item_port As Variant
        Dim port_no As Integer
        Dim ports As String
        ports = ""
        result = ws2.Cells(i, 20)
        item_port = Split(result, ";")
        For port_no = 0 To UBound(item_port)
            If item_port(port_no) <> "" Then
                If ports = "" Then
                    ports = Str(port_no + 1)
                Else
                    ports = ports + ";" + Str(port_no + 1)
                End If
            End If
        Next port_no
        ws2.Cells(i, 20) = Replace(ports, " ", "")
        If ws2.Cells(i, 20) <> ws2.Cells(i, 11) Then
            ws2.Cells(i, 20).Interior.Color = RGB(255, 255, 0)
            ws2.Cells(i, 20).Font.Color = RGB(255, 0, 0)
            ws2.Cells(i, 11).Interior.Color = RGB(255, 255, 0)
            ws2.Cells(i, 11).Font.Color = RGB(255, 0, 0)
        End If
        result = ""
    Next i
    
    ' Save the workbook with a meaningful name
    Application.StatusBar = "Data Comparision done..."
    MsgBox ("Done! Check excel workbook!")
    
    excelApp.Visible = True
    currentDate = Date
    currentHour = Hour(Now)
    currentMinute = Minute(Now)
	wb.SaveAs docPath + rfnsa_id + "_Comparison_PRD vs STAD_" + Format(currentDate, "YYYYMMDD") + Format(currentHour, "00") + Format(currentMinute, "00") + ".xlsx"
End Sub

Public Function dBmToWatts(dBm As Variant) As Double
Dim port_power_dB As Double
    port_power_dB = CDbl(dBm)
    dBmToWatts = Round((10 ^ (port_power_dB / 10)) / 1000, 1)
End Function

Public Function power_adjustment(inputvalue As String) As String
Dim result As Variant
Dim port_item As Variant

    result = Replace(inputvalue, ";;", "") ' remove extra ;; in power string
    port_item = ""
    If InStr(result, ";") Then
        port_item = Split(result, ";")
    Else
        port_item = result 'if only 1 port, no ;
    End If
            
    result = ""
    For Each Item In port_item
        If Item <> "" Then
            If result = "" Then
                result = Format(dBmToWatts(Item), "0.0") 'first port power value assign to result directly
                'Debug.Print Format(dBmToWatts(item), "0.0")
            Else
                result = result + "+" + Format(dBmToWatts(Item), "0.0") '2nd and after connected with "+"
            End If
        End If
    Next Item
    power_adjustment = Replace(result, " ", "")
End Function
Function RemoveZeroInPower(power As String) As String
Dim port_item As Variant
    result = power
    port_item = ""
    'if result with multiple ports,then divided by "+" to remove zeros, then merge again
    If InStr(result, "+") Then
            port_item = Split(result, "+")
            result = ""
            For Each Item In port_item
                If Item <> "0" Then
                    If result = "" Then
                        result = Item 'first port power value assign to result directly
                    Else
                        result = result + "+" + Item '2nd and after connected with "+"
                    End If
                End If
            Next Item
    'RemoveZeroInPower = result
    End If
    RemoveZeroInPower = result
End Function

Public Function GetCurrentUserDownloadsPath()
    ' Downloads Folder Registry Key
    Dim GUID_WIN_DOWNLOADS_FOLDER As String
    GUID_WIN_DOWNLOADS_FOLDER = "{374DE290-123F-4565-9164-39C4925E467B}"
    Dim KEY_PATH As String
    KEY_PATH = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\"
    Dim pathTmp As String
    
    On Error Resume Next
    pathTmp = RegKeyRead(KEY_PATH & GUID_WIN_DOWNLOADS_FOLDER)
    pathTmp = Replace$(pathTmp, "%USERPROFILE%", Environ$("USERPROFILE"))
    On Error GoTo 0
    
    GetCurrentUserDownloadsPath = pathTmp
End Function
Function RegKeyRead(i_RegKey As String) As String
Dim myWS As Object

  On Error Resume Next
  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'read key from registry
  RegKeyRead = myWS.RegRead(i_RegKey)
End Function
Function FindSecondSlashPosition(inputString As String) As Long
    Dim i As Long
    Dim slashCount As Long
    
    slashCount = 0
    
    For i = 1 To Len(inputString)
        If Mid(inputString, i, 1) = "/" Then
            slashCount = slashCount + 1
            If slashCount = 2 Then
                FindSecondSlashPosition = i
                Exit Function
            End If
        End If
    Next i
    
    ' If the second slash is not found, return -1 or any other indicator of not found
    FindSecondSlashPosition = -1
End Function
Function port_power_match(prd_power As String, stad_power As String) As Boolean
Dim port_item_prd As Variant
Dim port_item_stad As Variant
Dim ports_prd As Integer
Dim ports_stad As Integer
Dim power_gap As Double

If InStr(prd_power, "+") Then
    port_item_prd = Split(prd_power, "+")
    ports_prd = UBound(port_item_prd)
Else
    ports_prd = 0
End If

If InStr(stad_power, "+") Then
    port_item_stad = Split(stad_power, "+")
    ports_stad = UBound(port_item_stad)
Else
    ports_stad = 0
End If

'if ports number not match, return false
If ports_prd <> ports_stad Then
    port_power_match = False
    Exit Function
End If

If ports_stad = 0 Then
    power_gap = Round(Abs(stad_power - prd_power), 2)
    If power_gap >= 1 Then
        port_power_match = False
        Exit Function
    End If
    port_power_match = True
    Exit Function
End If
        
'if ports number matched, then start power comparison per port
For i = 0 To ports_prd
    power_gap = Round(Abs(port_item_stad(i) - port_item_prd(i)), 2)
    'Debug.Print power_gap
    If power_gap >= 1 Then
        port_power_match = False
        Exit Function
    End If
Next i
port_power_match = True
End Function
