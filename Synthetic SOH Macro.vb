Sub ApplyCellBorders(rng As Range)
    'Turn off diagonal borders
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    
    'Apply borders to edges and inside
    With rng.Borders
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
End Sub


Sub SOHMacro()
    
    ' Activate the current worksheet
    Set ws = ActiveSheet

    'Unprotect the worksheet
    ActiveSheet.Unprotect

' Get the start time
    startTime = Timer
    Dim endTime As Double
    Dim elapsedTime As Double

'Declaring Variables that will be used throughout code
    Dim lastRow As Long
    Dim i As Long
    Dim cell As Range

    'Declaring Date
    Dim currentDate As Date
    todaysDate = Format(Date, "DD/MM/YYYY")

Debug.Print "Today's date is " & todaysDate

    'Define lastRow
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row

'Cleaning Raw Data
    'Delete column I
    Columns("I").Delete Shift:=xlToLeft
    
'Insert Actual Weeks Cover Column (K)
    'Insert new column after column k
    Columns("K").Insert Shift:=xlToRight
    
    'Rename the header of the new column
    Range("K1").Value = "Actual Weeks Cover"
    
    'Drop g2/k2 formula in col K from row 2 to lastRow
    Range("K2:K" & lastRow).Formula = "=G2/I2"

    'Copy paste special formula from col K
    Range("K2:K" & lastRow).Copy
    Range("K2:K" & lastRow).PasteSpecial xlPasteValues
    'Clear clipboard
    Application.CutCopyMode = False
    
    'Sort column K smallest to largest
    Range("A1:M" & lastRow).Sort Key1:=Range("K1"), Order1:=xlAscending, header:=xlYes


'Replace negative values in column K with 0
    '**This code is hard coded in because cells include both nonnumerical and numerical values. This eliminates the mistype error
    For Each cell In Range("K2:K1000")
        If IsNumeric(cell.Value) And cell.Value < 0 Then
            cell.Value = 0
        End If
    Next cell

 'Replacing errors in col K if G is less than 1
    'Sort Col G Ascending
    Range("A1:M" & lastRow).Sort Key1:=Range("G1"), Order1:=xlAscending, header:=xlYes
    'Sort Col K Descending
    Range("A1:M" & lastRow).Sort Key1:=Range("K1"), Order1:=xlDescending, header:=xlYes
   
    'locating first instance of 1 in col G and assigning it to firstPosRow
    Set FirstPosRow = Range("G2:G" & lastRow).Find(What:="1", LookIn:=xlValues, LookAt:=xlWhole)
Debug.Print "First Positive Row in col G: " & FirstPosRow.Row
    
    'Assigning last negative row and replacing all values in col k from row 2 to lastNegative row
    lastNegative = FirstPosRow.Row - 1
Debug.Print "Last negative row in col G:" & lastNegative
    Range("K2:K" & lastNegative).Value = 0
    
'12+ in Col K different method than manual due to #DIV/0! error code.
    'In order to shorten up the lines of code without having to run into type errors we will sort col K in ascending order. This allows
        'the errors to be at the very bottom of the col. Therefore when using the .Find method we will locate the first instance of 12
        'before encountering a #DIV/0! error. This will bypass our code crashing.

    'Sort Col K in ascending order
    Range("A1:M" & lastRow).Sort Key1:=Range("K1"), Order1:=xlAscending, header:=xlYes
    
    'Locate first instance of 12 and replace all values from firstTwelve to last row
    Set firstTwelve = Range("K2:K" & lastRow).Find(What:="12", LookIn:=xlValues, LookAt:=xlWhole)
Debug.Print "First Twelve in Col K: " & firstTwelve.Row

    'Replace firstTwleve to lastRow in Col k to 12+
    firstTwelve = firstTwelve.Row
    Range("K" & firstTwelve & ":K" & lastRow).Value = "12+"
    
    'Format col K to number with 2 decimal points
    ws.Range("K:K").NumberFormat = "0.00"

'Formatting Col L - Next order ETA
    'Format col to short date
    Columns("L").NumberFormat = "dd/mm/yyyy"

    'Sort date in Descending order to locate largest date first
    Range("A1:M" & lastRow).Sort Key1:=Range("L1"), Order1:=xlDescending, header:=xlYes

    ' The objective is to add comments in col M corresponding to dates that predate today's date.
    'When running reports on a Monday there will not be yesterday's date available. So we will look for
    'dates between today through the previous friday. There will be 3 iterations, each iteration will be -1.
    'So we will be looking Newest (largest) to oldest (smallest) dates. If date == todays date we will + 1 row

    'Define Variables
    todayDate = Date
    Dim latestDateRow As Long

    'Last Row for Date in Col L
    lastRowL = Cells(Rows.Count, "L").End(xlUp).Row
   
    'Search range with lastRow for L
    Set searchRange = Range("L2:L" & lastRowL)
'Loop starting from the bottom up.

    For i = searchRange.Rows.Count To 1 Step -1
        If searchRange.Cells(i, 1).Value = todayDate Then
            latestDate = todayDate
            latestDateRow = searchRange.Cells(i, 1).Row + 1
            Exit For
        End If
    Next i

    If latestDateRow = 0 Then ' Today's date not found, search for yesterday's date
        For i = searchRange.Rows.Count To 1 Step -1
            If searchRange.Cells(i, 1).Value = yesterdayDate Then
                latestDate = yesterdayDate
                latestDateRow = searchRange.Cells(i, 1).Row + 1
                Exit For
            End If
        Next i
    End If

    If latestDateRow = 0 Then ' Yesterday's date not found, search for 2 days ago
        For i = searchRange.Rows.Count To 1 Step -1
            If searchRange.Cells(i, 1).Value = twoDaysAgo Then
                latestDate = twoDaysAgo
                latestDateRow = searchRange.Cells(i, 1).Row + 1
                Exit For
            End If
        Next i
    End If

    If latestDateRow = 0 Then ' Two days ago not found, search for 3 days ago
        For i = searchRange.Rows.Count To 1 Step -1
            If searchRange.Cells(i, 1).Value = threeDaysAgo Then
                latestDate = threeDaysAgo
                latestDateRow = searchRange.Cells(i, 1).Row + 1
                Exit For
            End If
        Next i
    End If
        
Debug.Print "Latest date found: " & Format(latestDate, "DD/MM/YYYY") & vbCrLf & _
    "Latest date row is: " & latestDateRow

'Adding in comments if proper date was found
    'If date was not found code block will be skipped.

    If latestDateRow <= 0 Then
        MsgBox "Cannot locate date within range.", vbInformation, "Comments will not be added to the data. Click OK to continue."
        GoTo ContinueCode
    Else
        'Adding comments into col M for the date rage before today to oldest date.
        'If there is already a comment from E360, it will append a comment.
        Set ws = ActiveSheet
        For currentRow = latestDateRow To lastRowL
            If ws.Cells(currentRow, "M").Value = "" Then
                ws.Cells(currentRow, "M").Value = "Branch follow up"
            Else
                ws.Cells(currentRow, "M").Value = ws.Cells(currentRow, "M").Value & ", Branch follow up"
            End If
            Next currentRow
    End If

ContinueCode:
'Data reporting is complete - Formatting finished report
    'Redefine lastRow
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row

    'Format Data Type
    ws.Range("G:G").NumberFormat = "0"
    ws.Range("H:J").NumberFormat = "General"

    'Sort Description followed by warehouse col in ascending order
    Range("A1:M" & lastRow).Sort Key1:=Range("B1"), Order1:=xlAscending, header:=xlYes
    Range("A1:M" & lastRow).Sort Key1:=Range("F1"), Order1:=xlAscending, header:=xlYes

    'Set font type and size for A1:N lastRow
    With Range("A2:N" & lastRow)
        .Font.Name = "Calibri"
        .Font.Size = 10
    End With

    'Header Format
    With Range("A1:N1")
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(56, 0, 102)
    End With

    'Horizontally center Cols A & C:N
    Range("A:A").HorizontalAlignment = xlCenter
    Range("C:N").HorizontalAlignment = xlCenter


    'Rename headers
     Range("B1").Value = "Item Description"
     Range("G1").Value = "Opening" & Chr(10) & "Balance"
     Range("H1").Value = "On" & Chr(10) & "Order"
     Range("I1").Value = "Avg Wkly" & Chr(10) & "Weeks" & Chr(10) & "Usage"
     Range("J1").Value = "Weeks"
     Range("L1").Value = "Next" & Chr(10) & "Order" & Chr(10) & "ETA Date"
     Range("M1").Value = "Admin Comments"
     Range("N1").Value = "Stakeholder Comments"
    
    ' Wrap text in row 1
    Rows(1).WrapText = True
    ' Autofit rows
    Range("A:V").EntireRow.AutoFit
    Range("A:V").EntireColumn.AutoFit
    
    ' Set the column widths
'    Columns("A:A").ColumnWidth = 5.4
'    Columns("B:B").ColumnWidth = 42.42
'    Columns("C:C").ColumnWidth = 10.8
'    Columns("D:D").ColumnWidth = 8
'    Columns("E:E").ColumnWidth = 4.71
'    Columns("F:F").ColumnWidth = 10.8
'    Columns("G:G").ColumnWidth = 6.8
'    Columns("H:H").ColumnWidth = 5.4
'    Columns("I:I").ColumnWidth = 7
'    Columns("J:K").ColumnWidth = 5.57
'    Columns("L:L").ColumnWidth = 8.29
'    Columns("M:M").ColumnWidth = 36
    Columns("N:N").ColumnWidth = 12.15
    
    'Remove comments from headers
    For Each cell In Range("A1:N1")
        If Not cell.Comment Is Nothing Then
            cell.ClearComments
        End If
    Next cell

    'insert header rows
    Rows("1:3").Insert
    Range("A1").Value = "Company's Name"
    Range("A2").Value = "Daily Stock on Hand Report for Stakeholder"
    Range("A3").Value = Format(todaysDate, "Long Date")

    'Vertically center all cols
    Range("A:V").VerticalAlignment = xlCenter
    Rows(4).VerticalAlignment = xlBottom

'Merge and center headers
    ' Merge and center row 1
    With Range("A1:N1")
        .Merge
        .HorizontalAlignment = xlCenter
    End With
    
    ' Merge and center row 2
    With Range("A2:N2")
        .Merge
        .HorizontalAlignment = xlCenter
    End With
    
    ' Merge and center row 3
    With Range("A3:N3")
        .Merge
        .HorizontalAlignment = xlCenter
    End With

    ' Headers Font
    Dim headerRows As Range
    Set headerRows = Range("A1:N3")

    For Each headerRow In headerRows
        With headerRow
            .Font.Name = "Calibri"
            .Font.Size = 12
            .Font.Bold = True
            .Font.Color = RGB(56, 0, 102)
        End With
    Next headerRow

'Freeze Headers
    Range("A5").Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.SmallScroll Down:=0

'Apply borders
    'redefine lastRow
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    ApplyCellBorders Range("A4:N" & lastRow)
    'Borders for rows 1:3
    With Range("A1:N3").Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
    End With

'Copy and Save Workbook
    'Return to A1
    Range("A1").Select
    Application.CutCopyMode = False

    Set originalWorkbook = ActiveWorkbook
    Set NewWorkbook = Workbooks.Add
    originalWorkbook.ActiveSheet.Copy After:=NewWorkbook.Sheets(NewWorkbook.Sheets.Count)
    
    'suppress confirmation dialogs
    Application.DisplayAlerts = False
    'Delete Sheet1
    Sheets("Sheet1").Delete
    ' rename the current sheet
    ActiveSheet.Name = "Daily SOH"
    ' re-enable confirmation dialogs
    Application.DisplayAlerts = True
    
    
    ' Close the new workbook
    ActiveWorkbook.Close SaveChanges:=False

    'end timer
    endTime = Timer
    
    ' Calculate the elapsed time in seconds
    elapsedTime = endTime - startTime
    
    ' Display the elapsed time in seconds
    MsgBox "Elapsed time: " & Format(elapsedTime, "0.00") & " seconds. This report used to take the team an upwards of 20-25 minutes manually."
End Sub


