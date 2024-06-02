# Reconciling-Two-Raw-Datasets
The macro written here was for an employee that has to extract two raw datasets from two different systems and then reconcile the total hours by date by bill code for the trips in the dataset.  
The macro is split into 3 parts.  The first part organizes the data extract from the first system and organizes it into a usable form for comparison later.  The second part takes the data extract from the second system and organizes it into a usable form for comparison later.  The 3rd step takes the two organized extracts and performs sumifs as well as many other adjustments to the data to give a final output that is easily readable for the employee.  The employee has to compare these reports, on average, about 4-8 times a month and this process saves the employee about 45 minutes of manual data manipulation every time.  This computes to about 4 hours and 30 minutes every month which saves the employee about 45 hours per year or 8.5 work days. 


Part 1:
```vbscript

Public Sub create_worksheets()

    'create trip tracker raw
    Sheets.Add After:=ActiveSheet
    ActiveSheet.name = "tt_raw"
    ActiveSheet.Move After:=Worksheets(Worksheets.Count)
    Sheets.Add After:=ActiveSheet
    ActiveSheet.name = "Acctg_Xref"
    ActiveSheet.Move After:=Worksheets(Worksheets.Count)

End Sub

Public Sub prep_tt_raw()

    Dim ttr As Worksheet: Set ttr = ThisWorkbook.Worksheets("tt_raw")
    
    ttr.Rows(1 & ":" & 4).EntireRow.Delete
    ttr.Rows(1).EntireRow.ClearContents
    ttr.Range("A:A,C:C, G:G, J:J, L:L, N:N, R:R, T:T").Delete
    
    'column headers ttr
    With ttr
        With .Cells(1, 1)
            .Value = "Trip Date"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(1, 2)
            .Value = "Trip ID"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(1, 3)
            .Value = "Bill To"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(1, 4)
            .Value = "Requester"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(1, 5)
            .Value = "Total Hours"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(1, 6)
            .Value = "Total Driver Cost"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(1, 7)
            .Value = "Total Miles"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(1, 8)
            .Value = "Total Mileage Cost"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(1, 9)
            .Value = "Other Cost"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(1, 10)
            .Value = "Budget #"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(1, 11)
            .Value = "Total Due"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(1, 12)
            .Value = "% to be charged"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(1, 13)
            .Value = "Total Due for Account"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
            With .Cells(1, 14)
            .Value = "Calculated - Bill Code"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(1, 15)
            .Value = "Calculated - Date, Bill Code"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
    End With
    
    ttr.Columns.AutoFit
    
End Sub

Public Sub part1b()

    Call bill_codes
    Call date_bill_codes

End Sub

Public Sub bill_codes()

    Dim ttr As Worksheet: Set ttr = ThisWorkbook.Worksheets("tt_raw")
    Dim axr As Worksheet: Set axr = ThisWorkbook.Worksheets("Acctg_Xref")
    Dim axrange As Range
    Dim resumeloop As Boolean
    
    lastrow = ttr.Range("J" & ttr.Rows.Count).End(xlUp).Row
    lastrow_axr = axr.Range("A" & axr.Rows.Count).End(xlUp).Row
    
    'Set axrange = axr.Range(Cells(2, 1), Cells(lastrow_axr, 1))
    
    For i = 2 To lastrow
    
        ttr.Activate
        bill_code = Application.Match(ttr.Cells(i, 10), axr.Range(axr.Cells(1, 1), axr.Cells(lastrow_axr, 1)), 0)
        
        If IsError(bill_code) Then
            
            resumeloop = MsgBox("Value not found. Do you want to continue to the next iteration?", vbYesNo, "Error") = vbYes
            
            If Not resumeloop Then
                Exit Sub
            End If
            
        ElseIf bill_code > 0 Then
        
            ttr.Cells(i, 14) = axr.Cells(bill_code, 3)
                
        End If

    Next i

End Sub

Public Sub date_bill_codes()

    Dim ttr As Worksheet: Set ttr = ThisWorkbook.Worksheets("tt_raw")
    lastrow = ttr.Range("D" & ttr.Rows.Count).End(xlUp).Row
    Dim str1 As String
    Dim str2 As String
    
    ttr.Range("A:A").Insert
    Columns(2).Copy
    Columns(1).PasteSpecial
    Columns(15).NumberFormat = "@"
    Columns(1).NumberFormat = "@"
    
    For i = 2 To lastrow
        
        str1 = Str(ttr.Cells(i, 1))
        str2 = Str(ttr.Cells(i, 15))
        str1 = Trim(str1)
        str2 = Trim(str2)
        ttr.Cells(i, 16) = str1 & str2

    Next i
    
    Columns(1).Delete

End Sub
```

Part 2:
```vbscript

Public Sub prep1_tc_raw()

    'create time clock raw worksheet
    Sheets.Add After:=ActiveSheet
    ActiveSheet.name = "tc_raw"
    ActiveSheet.Move After:=Worksheets(Worksheets.Count)
    'Sheets.Add After:=ActiveSheet
    'ActiveSheet.name = "tc_final"
    'ActiveSheet.Move After:=Worksheets(Worksheets.Count)
    
    Dim tcr As Worksheet: Set tcr = ThisWorkbook.Worksheets("tc_raw")
    
    
'    tcf.Cells(1, 1) = "Calculated - Date, Bill Code, Hours"
'    tcf.Cells(1, 2) = "Name"
'    tcf.Cells(1, 3) = "Date In"
'    tcf.Cells(1, 4) = "Job Code"
'    tcf.Cells(1, 5) = "Calculated - Date, Bill Code"
'    tcf.Cells(1, 6) = "Hours"
'    tcf.Cells(1, 7) = "Time In"
'    tcf.Cells(1, 8) = "Date Out"
'    tcf.Cells(1, 9) = "Time Out"
    
End Sub

Public Sub step5()
    
    Call raw_del_rows
    Call raw_name_fill
    Call raw_convert_hours
    Call raw1_billcode
    Call raw2_hours_convert
    Call raw3_date_out_convert
    Call adjust_raw_cols
    Call raw_as_final

End Sub


Public Sub raw_del_rows()

    Dim tcr As Worksheet: Set tcr = ThisWorkbook.Worksheets("tc_raw")
    
    
    lastrow = tcr.Range("B" & tcr.Rows.Count).End(xlUp).Row
    
    tcr.Cells(6, 2) = "Date In"
    tcr.Cells(6, 3) = "Time In"
    tcr.Cells(6, 4) = "Date Out"
    tcr.Cells(6, 5) = "Time Out"
    
    tcr.Rows(1 & ":" & 5).EntireRow.Delete
    
    For i = lastrow To 2 Step -1
    
        If IsEmpty(tcr.Cells(i, 2)) Then
            tcr.Rows(i).EntireRow.Delete
        Else
        End If
        
    Next i
    
End Sub

Public Sub raw_name_fill()
    
    Dim tcr As Worksheet: Set tcr = ThisWorkbook.Worksheets("tc_raw")
    
    Dim name As String
    
    lastrow = tcr.Range("B" & tcr.Rows.Count).End(xlUp).Row
    name = tcr.Cells(2, 1)
    
    For j = 3 To lastrow
    
        If IsEmpty(tcr.Cells(j, 1)) Then
            tcr.Cells(j, 1) = name
        ElseIf Not IsEmpty(tcr.Cells(j, 1)) Then
            name = tcr.Cells(j, 1)
        End If
    
    Next j

End Sub

Public Sub raw_convert_hours()

    Dim tcr As Worksheet: Set tcr = ThisWorkbook.Worksheets("tc_raw")
    
    Dim str1 As String
    lastrow = tcr.Range("B" & tcr.Rows.Count).End(xlUp).Row
    
    For i = 2 To lastrow
    
        tcr.Cells(i, 7) = WorksheetFunction.Substitute(tcr.Cells(i, 7), ":", ".")
        
        str1 = Right(tcr.Cells(i, 7), 2)
        
        If str1 = "15" Then
            tcr.Cells(i, 7) = WorksheetFunction.Substitute(tcr.Cells(i, 7), "15", "25")
        ElseIf str1 = "30" Then
            tcr.Cells(i, 7) = WorksheetFunction.Substitute(tcr.Cells(i, 7), "30", "50")
        ElseIf str1 = "45" Then
            tcr.Cells(i, 7) = WorksheetFunction.Substitute(tcr.Cells(i, 7), "45", "75")
        Else
        End If
        
    Next i
    

End Sub


Public Sub raw1_billcode()
    
    Dim tcr As Worksheet: Set tcr = ThisWorkbook.Worksheets("tc_raw")
    lastrow = tcr.Range("B" & tcr.Rows.Count).End(xlUp).Row
    
    
    tcr.Range("B:B").Insert
    tcr.Columns(3).NumberFormat = "m/d/yyyy"
    tcr.Columns(2).NumberFormat = "General"

    'NumberFormat = "General" is general
    'NumberFormat = "@" is text
    'NumberFormat = "0" is number
    'NumberFormat = "General" is date
    
    For i = 2 To lastrow
        
        tcr.Cells(i, 2) = DateValue(tcr.Cells(i, 3))
    
    Next i
    
    tcr.Columns(2).NumberFormat = "@"


End Sub


Public Sub raw2_hours_convert()
    
    Dim tcr As Worksheet: Set tcr = ThisWorkbook.Worksheets("tc_raw")
    Dim str3 As String
    Dim str4 As String
    lastrow = tcr.Range("C" & tcr.Rows.Count).End(xlUp).Row
    Dim destrow As Integer
    
    
    tcr.Columns(8).Copy
    tcr.Columns(9).PasteSpecial
    tcr.Columns(9).NumberFormat = "@"
    
    For j = 2 To lastrow
    
        str4 = Right(tcr.Cells(j, 9), 3)
        
        If str4 = ".00" Then
            str3 = tcr.Cells(j, 9)
            str3 = WorksheetFunction.Substitute(str3, ".00", "")
            tcr.Cells(j, 9) = str3
        ElseIf str4 = ".50" Then
            str3 = tcr.Cells(j, 9)
            str3 = WorksheetFunction.Substitute(str3, ".50", ".5")
            tcr.Cells(j, 9) = str3
        End If
        
    Next j


End Sub

Public Sub raw3_date_out_convert()

    Dim tcr As Worksheet: Set tcr = ThisWorkbook.Worksheets("tc_raw")
    Dim str1 As String
    Dim str2 As String
    lastrow = tcr.Range("C" & tcr.Rows.Count).End(xlUp).Row
    
    tcr.Range("F:F").Insert
    str1 = "/" & Year(Date)
    
    For i = 2 To lastrow
    
        str2 = tcr.Cells(i, 5)
        str2 = Trim(str2)
        tcr.Cells(i, 6) = str2 & str1
    
    Next i

End Sub

Public Sub adjust_raw_cols()
    'NOT DONE
    Dim tcr As Worksheet: Set tcr = ThisWorkbook.Worksheets("tc_raw")
    
    'date as text value
    tcr.Columns(2).Copy
    tcr.Columns(20).PasteSpecial
    
    'hours as decimal text
    tcr.Columns(10).Copy
    tcr.Columns(21).PasteSpecial
    
    'name
    tcr.Columns(1).Copy
    tcr.Columns(12).PasteSpecial
    
    'date in
    tcr.Columns(3).Copy
    tcr.Columns(13).PasteSpecial
    
    'job code
    tcr.Columns(8).Copy
    tcr.Columns(14).PasteSpecial
    
    'hours
    tcr.Columns(9).Copy
    tcr.Columns(16).PasteSpecial
    
    'time in
    tcr.Columns(4).Copy
    tcr.Columns(17).PasteSpecial
    
    'date out
    tcr.Columns(6).Copy
    tcr.Columns(18).PasteSpecial
    
    'time out
    tcr.Columns(7).Copy
    tcr.Columns(19).PasteSpecial
    
    
    tcr.Columns("A:J").EntireColumn.Delete

    
End Sub



Public Sub raw_as_final()
    
    Dim tcr As Worksheet: Set tcr = ThisWorkbook.Worksheets("tc_raw")
    Dim str1 As String
    Dim str2 As String
    Dim str3 As String
    lastrow = tcr.Range("C" & tcr.Rows.Count).End(xlUp).Row
    Dim cols As Range: Set cols = tcr.Columns("A:I")
    
    
    
    tcr.Columns(1).NumberFormat = "@"
    
    For i = 2 To lastrow
        
        str1 = tcr.Cells(i, 10)
        str2 = Str(tcr.Cells(i, 4))
        str3 = tcr.Cells(i, 11)
        str1 = Trim(str1)
        str2 = Trim(str2)
        str3 = Trim(str3)
        
        tcr.Cells(i, 1) = str1 & str2 & str3
        tcr.Cells(i, 5) = str1 & str2
        
    Next i
    
    tcr.Columns("J:K").EntireColumn.Delete
    cols.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
    
    tcr.Cells(1, 1) = "Calculated - Date, Bill Code, Hours"
    tcr.Cells(1, 5) = "Calculated - Date, Bill Code"
    tcr.Cells(1, 8) = "Date Out"
    tcr.Columns.AutoFit
    tcr.Activate
    tcr.Cells(1, 1).Select
    
End Sub
```

Part 3:
```vbscript
Public Sub step6()

    Call tt_tc_compares
    Call input_data
    Call total_hours_dups
    Call lineup_dbcs_tc
    Call lineup_dbcs_tt
    Call dif_highlight
    

End Sub


Public Sub tt_tc_compares()
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.name = "TC_comparison"
    ActiveSheet.Move After:=Worksheets(Worksheets.Count)
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.name = "TT_comparison"
    ActiveSheet.Move After:=Worksheets(Worksheets.Count)
    
End Sub


Public Sub input_data()
    
    Dim tcc As Worksheet: Set tcc = ThisWorkbook.Worksheets("TC_comparison")
    Dim ttc As Worksheet: Set ttc = ThisWorkbook.Worksheets("TT_comparison")
    Dim ttr As Worksheet: Set ttr = ThisWorkbook.Worksheets("tt_raw")
    Dim tcr As Worksheet: Set tcr = ThisWorkbook.Worksheets("tc_raw")
    
    tt_lastrow = ttr.Range("O" & ttr.Rows.Count).End(xlUp).Row
    tc_lastrow = tcr.Range("E" & tcr.Rows.Count).End(xlUp).Row
    tcc_lastrow = tcc.Range("A" & tcc.Rows.Count).End(xlUp).Row
    ttc_lastrow = ttc.Range("A" & ttc.Rows.Count).End(xlUp).Row
    
    
    
    tcc.Cells(1, 1) = "TC Date Bill Code"
    tcc.Cells(1, 2) = "TC Trip Date"
    tcc.Cells(1, 3) = "TC Bill Code"
    tcc.Cells(1, 4) = "TC hours"
    tcc.Cells(1, 5) = "TC TOTAL hours"
    tcc.Columns.AutoFit
    
    ttc.Cells(1, 1) = "TT Date Bill Code"
    ttc.Cells(1, 2) = "TT Trip Date"
    ttc.Cells(1, 3) = "TT Bill Code"
    ttc.Cells(1, 4) = "TT hours"
    ttc.Cells(1, 5) = "TT TOTAL hours"
    ttc.Columns.AutoFit
    
    'tt_date_billcode
    Range(ttr.Cells(2, 15), ttr.Cells(tt_lastrow, 15)).Copy
    ttc.Cells(2, 1).PasteSpecial
    
    'tt_tripdate
    Range(ttr.Cells(2, 1), ttr.Cells(tt_lastrow, 1)).Copy
    ttc.Cells(2, 2).PasteSpecial
    
    'tt_billcode
    Range(ttr.Cells(2, 14), ttr.Cells(tt_lastrow, 14)).Copy
    ttc.Cells(2, 3).PasteSpecial
    
    'tt_hours
    Range(ttr.Cells(2, 5), ttr.Cells(tt_lastrow, 5)).Copy
    ttc.Cells(2, 4).PasteSpecial
    
    
    'tc_date_billcode
    Range(tcr.Cells(2, 5), tcr.Cells(tc_lastrow, 5)).Copy
    tcc.Cells(2, 1).PasteSpecial
    
    'tc_tripdate
    Range(tcr.Cells(2, 3), tcr.Cells(tc_lastrow, 3)).Copy
    tcc.Cells(2, 2).PasteSpecial
    
    'tc_billcode
    Range(tcr.Cells(2, 4), tcr.Cells(tc_lastrow, 4)).Copy
    tcc.Cells(2, 3).PasteSpecial
    
    'tc_hours
    For i = 2 To tc_lastrow
    
        tcr.Cells(i, 6) = Val(tcr.Cells(i, 6))
    
    Next i
    Range(tcr.Cells(2, 6), tcr.Cells(tc_lastrow, 6)).Copy
    tcc.Cells(2, 4).PasteSpecial
    
    'number format for hour columns
    ttc.Columns(4).NumberFormat = "0.00"
    tcc.Columns(4).NumberFormat = "0.00"
    
    ttc.Range("A1").CurrentRegion.Sort _
        key1:=ttc.Range("A1"), Order1:=xlAscending, Header:=xlYes
    
    tcc.Range("A1").CurrentRegion.Sort _
        key1:=tcc.Range("A1"), Order1:=xlAscending, Header:=xlYes
    
End Sub


Public Sub total_hours_dups()
    
    Dim tcc As Worksheet: Set tcc = ThisWorkbook.Worksheets("TC_comparison")
    Dim ttc As Worksheet: Set ttc = ThisWorkbook.Worksheets("TT_comparison")
    
    tcc_lastrow = tcc.Range("A" & tcc.Rows.Count).End(xlUp).Row
    ttc_lastrow = ttc.Range("A" & ttc.Rows.Count).End(xlUp).Row
    
    Dim tcc_hours As Range: Set tcc_hours = Range(tcc.Cells(2, 4), tcc.Cells(tcc_lastrow, 4))
    Dim ttc_hours As Range: Set ttc_hours = Range(ttc.Cells(2, 4), ttc.Cells(ttc_lastrow, 4))
    Dim tcc_dbc As Range: Set tcc_dbc = Range(tcc.Cells(2, 1), tcc.Cells(tcc_lastrow, 1))
    Dim ttc_dbc As Range: Set ttc_dbc = Range(ttc.Cells(2, 1), ttc.Cells(ttc_lastrow, 1))
    
    'ttc total hours
    For j = 2 To ttc_lastrow

        ttc.Cells(j, 5) = WorksheetFunction.SumIf(ttc_dbc, ttc.Cells(j, 1), ttc_hours)

    Next j
    
    'tcc total hours
    For k = 2 To tcc_lastrow
    
        tcc.Cells(k, 5) = WorksheetFunction.SumIf(tcc_dbc, tcc.Cells(k, 1), tcc_hours)
    
    Next k
    
    'remove dups
    ttc.Range("A:E").RemoveDuplicates Columns:=1, Header:=xlYes
    tcc.Range("A:E").RemoveDuplicates Columns:=1, Header:=xlYes
    
        
End Sub

Public Sub lineup_dbcs_tc()
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.name = "Final_Output"
    ActiveSheet.Move After:=Worksheets(Worksheets.Count)
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.name = "Missing_Codes"
    ActiveSheet.Move After:=Worksheets(Worksheets.Count)
    
    Dim fo As Worksheet: Set fo = ThisWorkbook.Worksheets("Final_Output")
    Dim mc As Worksheet: Set mc = ThisWorkbook.Worksheets("Missing_Codes")
    Dim tcc As Worksheet: Set tcc = ThisWorkbook.Worksheets("TC_comparison")
    Dim ttc As Worksheet: Set ttc = ThisWorkbook.Worksheets("TT_comparison")
    
    tcc_lastrow = tcc.Range("A" & tcc.Rows.Count).End(xlUp).Row
    ttc_lastrow = ttc.Range("A" & ttc.Rows.Count).End(xlUp).Row
    destrow = 2
    
    
    'headers
    fo.Cells(1, 1) = "Date Bill Code"
    fo.Cells(1, 2) = "Trip Date"
    fo.Cells(1, 3) = "Bill Code"
    fo.Cells(1, 4) = "TC total hours"
    fo.Cells(1, 5) = "TT total hours"
    fo.Cells(1, 6) = "Difference"
    fo.Columns.AutoFit
    mc.Cells(1, 1) = "TC billcodes not found in TT"
    mc.Cells(1, 2) = "TT billcodes not found in TC"
    mc.Columns.AutoFit
    
    
    
    For i = 2 To tcc_lastrow
    
        tcc.Activate
        dbc = Application.Match(tcc.Cells(i, 1), ttc.Range(ttc.Cells(1, 1), ttc.Cells(ttc_lastrow, 1)), 0)
               
           If IsError(dbc) Then
                mc.Cells(destrow, 1) = tcc.Cells(i, 1)
                destrow = destrow + 1
           ElseIf dbc > 0 Then
           
                fo.Cells(i, 1) = tcc.Cells(i, 1)
                fo.Cells(i, 2) = tcc.Cells(i, 2)
                fo.Cells(i, 3) = tcc.Cells(i, 3)
                fo.Cells(i, 4) = tcc.Cells(i, 5)
                fo.Cells(i, 5) = ttc.Cells(dbc, 5)
               
           End If
        
    Next i
    
End Sub

Public Sub lineup_dbcs_tt()
    
    Dim fo As Worksheet: Set fo = ThisWorkbook.Worksheets("Final_Output")
    Dim mc As Worksheet: Set mc = ThisWorkbook.Worksheets("Missing_Codes")
    Dim tcc As Worksheet: Set tcc = ThisWorkbook.Worksheets("TC_comparison")
    Dim ttc As Worksheet: Set ttc = ThisWorkbook.Worksheets("TT_comparison")
    
    tcc_lastrow = tcc.Range("A" & tcc.Rows.Count).End(xlUp).Row
    ttc_lastrow = ttc.Range("A" & ttc.Rows.Count).End(xlUp).Row
    fo_lastrow = fo.Range("A" & fo.Rows.Count).End(xlUp).Row
    destrow = 2
    
    fo.Columns(2).NumberFormat = "m/d/yyy"
    fo.Columns(2).AutoFit
    
    For i = 2 To ttc_lastrow
    
        ttc.Activate
        dbc = Application.Match(ttc.Cells(i, 1), tcc.Range(tcc.Cells(1, 1), tcc.Cells(tcc_lastrow, 1)), 0)
               
           If IsError(dbc) Then
                mc.Cells(destrow, 2) = ttc.Cells(i, 1)
                destrow = destrow + 1
           ElseIf dbc > 0 Then
           
                fo.Cells(fo_lastrow + 1, 1) = ttc.Cells(i, 1)
                fo.Cells(fo_lastrow + 1, 2) = ttc.Cells(i, 2)
                fo.Cells(fo_lastrow + 1, 3) = ttc.Cells(i, 3)
                fo.Cells(fo_lastrow + 1, 4) = ttc.Cells(i, 5)
                fo.Cells(fo_lastrow + 1, 5) = tcc.Cells(dbc, 5)
                
                fo_lastrow = fo_lastrow + 1
               
           End If
        
           
        
    Next i
    
End Sub

Public Sub dif_highlight()

    Dim fo As Worksheet: Set fo = ThisWorkbook.Worksheets("Final_Output")
    fo_lastrow = fo.Range("A" & fo.Rows.Count).End(xlUp).Row
    
    'calc difference and highlight if not zero
    For i = 2 To fo_lastrow
    
        fo.Cells(i, 6) = fo.Cells(i, 4) - fo.Cells(i, 5)
        
        If fo.Cells(i, 6) = 0 Then
        Else
            fo.Cells(i, 6).Interior.Color = vbYellow
        End If
    
    Next i

End Sub


Public Sub reset()

    Dim fo As Worksheet: Set fo = ThisWorkbook.Worksheets("Final_Output")
    Dim mc As Worksheet: Set mc = ThisWorkbook.Worksheets("Missing_Codes")
    Dim tcc As Worksheet: Set tcc = ThisWorkbook.Worksheets("TC_comparison")
    Dim ttc As Worksheet: Set ttc = ThisWorkbook.Worksheets("TT_comparison")
    Dim ttr As Worksheet: Set ttr = ThisWorkbook.Worksheets("tt_raw")
    Dim tcr As Worksheet: Set tcr = ThisWorkbook.Worksheets("tc_raw")
    Dim axr As Worksheet: Set axr = ThisWorkbook.Worksheets("Acctg_Xref")

    fo.Delete
    mc.Delete
    tcc.Activate
    ttc.Delete
    tcr.Delete
    ttr.Delete
    axr.Delete

End Sub

```
