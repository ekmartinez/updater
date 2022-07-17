Attribute VB_Name = "Module1"
Sub Form()

FormX.Show

End Sub

Private Sub CommandButton1_Click()

Dim Msg As String
Dim Title As String
Dim Config As Integer
Dim Ans As Integer

    Msg = "If you continue you will delete all data entered into the worksheets. Do you wish to continue?"
    Title = "Cell Cleaner"
    Config = vbYesNo + vbQuestion
    Ans = MsgBox(Msg, Config, Title)
    If Ans = vbYes Then CellCleaner
End Sub

Private Sub CommandButton2_Click()

Dim Msg As String
Dim Title As String
Dim Config As Integer
Dim Ans As Integer

    Msg = "Do you want to summarize all data in the worksheet and create a summary sheet?"
    Title = "Report Generator"
    Config = vbYesNo + vbQuestion
    Ans = MsgBox(Msg, Config, Title)
    If Ans = vbYes Then CopyDataWithoutHeaders
End Sub


Private Sub CommandButton3_Click()
Unload FormX
End Sub



Sub CellCleaner()

'This code removes all data entries from previous processing in all worksheets.

Dim sh As Worksheet

ScreenUpdating = True
    
    For Each sh In ActiveWorkbook.Worksheets
        sh.Range("C8:L33").ClearContents
        sh.Range("C8:L33").Interior.Color = xlNone ' No Fill
        sh.Range("C30") = "Total Hours"
        sh.Range("C31") = "Gross Pay"
        sh.Range("C32") = "Tax Withholding 7%"
        sh.Range("C33") = "Net Pay"
    Next sh
    
Unload FormX
    
   
End Sub


Function LastRow(sh As Worksheet)
    On Error Resume Next
    LastRow = sh.Cells.Find(What:="*", _
                            After:=sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
    On Error GoTo 0
End Function

Function LastCol(sh As Worksheet)
    On Error Resume Next
    LastCol = sh.Cells.Find(What:="*", _
                            After:=sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
    On Error GoTo 0
End Function

Sub CopyDataWithoutHeaders()
    Dim sh As Worksheet
    Dim DestSh As Worksheet
    Dim Last As Long
    Dim shLast As Long
    Dim CopyRng As Range
    Dim StartRow As Long
    Dim i As Integer
    Dim LRow As Long
    Dim Shift1 As Single
    Dim Shift2 As Single
    Dim Shift3 As Single
    Dim Shift1wd As Single
    Dim Shift2wd As Single
    Dim Shift3wd As Single
    Dim Unidades As Single
    Dim Hollydaysandweekends As Single
    
       
    
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With

    ' Delete the summary sheet if it exists.
    Application.DisplayAlerts = True
    On Error Resume Next
    ActiveWorkbook.Worksheets("Summary").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Add a new summary worksheet.
    Set DestSh = ActiveWorkbook.Worksheets.Add
    DestSh.Name = "Summary"
    
    
    
    ' Fill in the start row.
    StartRow = 8
  

    ' Loop through all worksheets and copy the data to the
    ' summary worksheet.
    For Each sh In ActiveWorkbook.Worksheets
        If sh.Name <> DestSh.Name Then

            ' Find the last row with data on the summary
            ' and source worksheets.
            Last = LastRow(DestSh)
            shLast = LastRow(sh)

            ' If source worksheet is not empty and if the last
            ' row >= StartRow, copy the range.
            If shLast > 0 And shLast >= StartRow Then
                'Set the range that you want to copy
                Set CopyRng = sh.Range(sh.Rows(StartRow), sh.Rows(shLast))

                ' This statement copies values and formats.
                CopyRng.Copy
                With DestSh.Cells(Last + 1, "A")
                    .PasteSpecial xlPasteValues
                    Application.CutCopyMode = False
                End With

            End If

        End If
    Next

ExitTheSub:

    Application.Goto DestSh.Cells(1)

    

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    
'Excel MVP Ron de Bruin (https://msdn.microsoft.com/en-us/library/cc793964(v=office.12).aspx)

Columns("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
Columns("C:C").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

Range("C1").EntireColumn.Insert


LRow = Range("B" & Rows.Count).End(xlUp).Row

For i = 1 To LRow

  Cells(i, 3) = Cells(i, 1) & " " & Cells(i, 2)
    
Next i

Columns("A:B").SpecialCells(xlCellTypeBlanks).EntireColumn.Delete
'Columns("L").SpecialCells(xlCellTypeBlanks).EntireColumn.Delete

Rows(1).Insert

    Worksheets("Summary").Range("A1").Value = "Contractor"
    Worksheets("Summary").Range("B1").Value = "Location"
    Worksheets("Summary").Range("C1").Value = "Date"
    Worksheets("Summary").Range("D1").Value = "Shift 1"
    Worksheets("Summary").Range("E1").Value = "Shift2"
    Worksheets("Summary").Range("F1").Value = "Shift3"
    Worksheets("Summary").Range("G1").Value = "Shift1 Weekend"
    Worksheets("Summary").Range("H1").Value = "Shift2 Weekend"
    Worksheets("Summary").Range("I1").Value = "Shift3 Weekend"
    Worksheets("Summary").Range("J1").Value = "Unidades Weekend"
    Worksheets("Summary").Range("K1").Value = "Hollydays & Weekend"
 
    
    ' AutoFit the column width in the summary sheet.
    DestSh.Columns.AutoFit

    Columns("B:K").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
    End With
    
  Range("C1:C500").NumberFormat = "mm/dd/yyyy"
   
    
    Msg = "Do you wish to update master database?"
    Title = "Report Generator"
    Config = vbYesNo + vbQuestion
    Ans = MsgBox(Msg, Config, Title)
    If Ans = vbYes Then DataBaseUpdater
    If Ans = vbNo Then Unload FormX

Unload FormX

End Sub

Private Sub CommandButton4_Click()

Dim Msg As String
Dim Title As String
Dim Config As Integer
Dim Ans As Integer

    Msg = "The following procedure allows you to sort all worksheets in ascending or decending order.  Do you wish to continue?"
    Title = "Report Generator"
    Config = vbYesNo + vbQuestion
    Ans = MsgBox(Msg, Config, Title)
    If Ans = vbYes Then Sort_Active_Book
End Sub


Sub Sort_Active_Book()

'This code arranges all sheets within a workbook in ascending or descending order.


Dim i As Integer
Dim j As Integer
Dim iAnswer As VbMsgBoxResult
'
' Prompt the user as which direction they wish to
' sort the worksheets.
'
   iAnswer = MsgBox("Sort Sheets in Ascending Order?" & Chr(10) _
     & "Clicking No will sort in Descending Order", _
     vbYesNoCancel + vbQuestion + vbDefaultButton1, "Sort Worksheets")
   For i = 1 To Sheets.Count
      For j = 1 To Sheets.Count - 1
'
' If the answer is Yes, then sort in ascending order.
'
         If iAnswer = vbYes Then
            If UCase$(Sheets(j).Name) > UCase$(Sheets(j + 1).Name) Then
               Sheets(j).Move After:=Sheets(j + 1)
            End If
'
' If the answer is No, then sort in descending order.
'
         ElseIf iAnswer = vbNo Then
            If UCase$(Sheets(j).Name) < UCase$(Sheets(j + 1).Name) Then
               Sheets(j).Move After:=Sheets(j + 1)
            End If
         End If
      Next j
   Next i
   
Unload FormX

End Sub

Sub CopyPaster()

Dim sht As Worksheet

ScreenUpdating = True

Sheets("572-0040 MARALY FELICIANO").Range("D8:L27").Copy
    For Each sht In Worksheets
sht.Range("D8:L33").PasteSpecial xlPasteAll
    Next



Application.CutCopyMode = False

Unload FormX

End Sub


'How to sort worksheets alphanumerically in a workbook in Excel
'https://support.microsoft.com/en-us/kb/812386

Private Sub CommandButton5_Click()

Dim Msg As String
Dim Title As String
Dim Config As Integer
Dim Ans As Integer

    Msg = "Are you ready to copy-paste?"
    Title = "Copy-Paster"
    Config = vbYesNo + vbQuestion
    Ans = MsgBox(Msg, Config, Title)
    If Ans = vbYes Then CopyPaster

End Sub


Sub DataBaseUpdater()

Dim Destiny As Workbook
Dim sht As Worksheet
Dim LastRow As Long
Dim LastColumn As Long
Dim StartCell As Range


ScreenUpdating = True

'Copy data from summary report
Set sht = Worksheets("Summary")
Set StartCell = Range("A5")
  LastRow = sht.Cells(sht.Rows.Count, StartCell.Column).End(xlUp).Row  'Find Last Row
  sht.Range("A5:K" & LastRow).Copy     'Select Range
  
'Paste data from summary report to Database
    Set Destiny = Workbooks.Open("N:\Professional Services\Database.xlsx")
        Destiny.Sheets("Database").Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
        
        MsgBox "Database has been actualized. You are welcome!"
        
        'Destiny.Save
        'Destiny.Close

CutCopyMode = False



'Automate the calculation of total hours in database
'LRow = Range("A" & Rows.Count).End(xlUp).Row
'For i = 2 To LRow
'Next i



End Sub


'LRow = Range("B" & Rows.Count).End(xlUp).Row


