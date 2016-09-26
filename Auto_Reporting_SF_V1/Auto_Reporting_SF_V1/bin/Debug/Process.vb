'This is stored as "Report for SF-All States_vba.vb":

'Run SQL Query 01, from file "SF-All States_SQLs.sql" with username/pwd@nmap.mitchell.com
'There are 3 parts in SQL1, for Comparable, Nada and Redbook reports, run them one by one.
'Export THE "SQL RESULT" To excel and take the backup.
'Mix the result of 2 runs i.e. Nada and Redbook.
'On both parts, Seperately Run Sub StartProcessing() which calls DetermineLastrowtovalidate to count the no. of rows and
'addcols, extractUML which extracts data from UML form.
'Now, Run SeperateOriginals_comp() on comparable part and SeperateOriginals_book() on book part to seperate originals from revised data and gives
'the desired output in tab originalcreators.
'Now create the unique Original and Revised UserId list (all must be in upper case) from the OriginalCreators tab column "D"
'run the SQL Query 02 with taking unique UserId as input , and save SQL Result into tab mapUserId.
'and run the sub mapUserId.
'In both excels Rearrange the Columns position by inserting Office column from "AA" to "D", Loss State from "AB" to "O" and
'Valuation Type from "AC" to "Q".
'Make a new excel file and copy the "Originalcreators" result to 2 different worksheets.
'Rename the 2 tabs to "Comparable" and "Book" containing corresponding Data.


Const FirstRowToValidate = 2
Dim LastRowToValidate As Long
Const ClaimIDCol = "B"


Sub StartProcessing()
 
'delete PL/SQL Numbering column if not already deleted; this assumes sort order on sheet hasn't been changed
If Worksheets("Table").Cells(2, "A").Value = "1" Then
    Worksheets("Table").Columns("A").Delete
End If

Dim wrkSheet As Worksheet

Set wrkSheet = ActiveWorkbook.Worksheets.Add
wrkSheet.Name = "mapUserId"
Set wrkSheet = ActiveWorkbook.Worksheets.Add
wrkSheet.Name = "OriginalCreators"

DetermineLastRowToValidate "Table"
addcols
extractUML
'SeperateOriginals
End Sub


Sub addcols()

Worksheets("OriginalCreators").Range("A1:AA1").HorizontalAlignment = xlCenter
Worksheets("OriginalCreators").Range("A1:AA1").WrapText = True
Worksheets("OriginalCreators").Range("A1:AA1").Font.Bold = True
Worksheets("OriginalCreators").Range("A1:AA1").Font.Size = 8

Worksheets("OriginalCreators").Cells(1, "A").Value = "Claim Number"
Worksheets("OriginalCreators").Cells(1, "B").Value = "Original Exp"
Worksheets("OriginalCreators").Cells(1, "C").Value = "Revised Exp"
Worksheets("OriginalCreators").Cells(1, "D").Value = "Original USER ID"
Worksheets("OriginalCreators").Cells(1, "E").Value = "Revised USER ID"
Worksheets("OriginalCreators").Cells(1, "F").Value = "VIN"
Worksheets("OriginalCreators").Cells(1, "G").Value = "Vehicle Year"
Worksheets("OriginalCreators").Cells(1, "H").Value = "Vehicle Make"
Worksheets("OriginalCreators").Cells(1, "I").Value = "Vehicle Model"
Worksheets("OriginalCreators").Cells(1, "J").Value = "Mileage"
Worksheets("OriginalCreators").Cells(1, "K").Value = "Loss Date"
Worksheets("OriginalCreators").Cells(1, "L").Value = "Loss Zip"
Worksheets("OriginalCreators").Cells(1, "M").Value = "Original Final Value"
Worksheets("OriginalCreators").Cells(1, "N").Value = "Revised Final Value"
Worksheets("OriginalCreators").Cells(1, "O").Value = "Original Condition Rating"
Worksheets("OriginalCreators").Cells(1, "P").Value = "Original Condition Amount"
Worksheets("OriginalCreators").Cells(1, "Q").Value = "Revised Condition Rating"
Worksheets("OriginalCreators").Cells(1, "R").Value = "Revised Condition Amount"
Worksheets("OriginalCreators").Cells(1, "S").Value = "Original AfterMarket Amount"
Worksheets("OriginalCreators").Cells(1, "T").Value = "Revised AfterMarket Amount"
Worksheets("OriginalCreators").Cells(1, "U").Value = "Original Refurbishment Amount"
Worksheets("OriginalCreators").Cells(1, "V").Value = "Revised Refurbishment Amount"
Worksheets("OriginalCreators").Cells(1, "W").Value = "Original Prior Damage Amount"
Worksheets("OriginalCreators").Cells(1, "X").Value = "Revised Prior Damage Amount"
Worksheets("OriginalCreators").Cells(1, "Y").Value = "Office"
Worksheets("OriginalCreators").Cells(1, "Z").Value = "Loss State"
Worksheets("OriginalCreators").Cells(1, "AA").Value = "Valuation Type (Book/Comparable)"
'format entire sheet MS San Serif 8
Worksheets("OriginalCreators").Cells.Select
    With Selection.Font
        .Name = "MS Sans Serif"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
    End With
Selection.RowHeight = 12
Worksheets("OriginalCreators").Range("A1:AA1").Cells.Select
Selection.RowHeight = 50


'freeze headers
Worksheets("OriginalCreators").Rows("2:2").Select
ActiveWindow.FreezePanes = True


End Sub


Sub extractUML()

Dim nFirstPos As Long
Dim nLastPos As Long
Dim strTemp As String
'Clean up Overall Condition Score.  It will look like this:
'<adj:RatingScore>2.0</adj:RatingScore>
'        </adj:ConditionAdjustment>
'        <adj:UIAdjustments>

Worksheets("Table").Range("A2:Y" & LastRowToValidate).Sort Key1:=Worksheets("Table").Range(ClaimIDCol & "2"), Key2:=Worksheets("Table").Range("A" & "2"), order2:=xlDescending

For Each c In Worksheets("Table").Range("E" & "2:" & "E" & LastRowToValidate).Cells
    'find the first > and strip everything to its left and the first < (after the value) and strip it and everything to the right
    nFirstPos = 0
    nLastPos = 0
    strTemp = ""
    nFirstPos = InStr(1, c.Value, ">")
    nLastPos = InStr(3, c.Value, "<")
    If nFirstPos > 0 And nLastPos > 0 Then
        strTemp = Mid(c.Value, nFirstPos + 1, nLastPos - nFirstPos - 1)
        c.Value = strTemp
    End If
Next

For Each c In Worksheets("Table").Range("F" & "2:" & "F" & LastRowToValidate).Cells
    'find the first > and strip everything to its left and the first < (after the value) and strip it and everything to the right
    nFirstPos = 0
    nLastPos = 0
    strTemp = ""
    nFirstPos = InStr(1, c.Value, ">")
    nLastPos = InStr(3, c.Value, "<")
    If nFirstPos > 0 And nLastPos > 0 Then
        strTemp = Mid(c.Value, nFirstPos + 1, nLastPos - nFirstPos - 1)
        c.Value = strTemp
    End If
Next



For Each c In Worksheets("Table").Range("O" & "2:" & "O" & LastRowToValidate).Cells
    'find the first > and strip everything to its left and the first < (after the value) and strip it and everything to the right
    nFirstPos = 0
    nLastPos = 0
    strTemp = ""
    nFirstPos = InStr(1, c.Value, ">")
    nLastPos = InStr(7, c.Value, "<")
    If nFirstPos > 0 And nLastPos > 0 Then
        strTemp = Mid(c.Value, nFirstPos + 1, nLastPos - nFirstPos - 1)
        c.Value = strTemp
    End If
Next


For Each c In Worksheets("Table").Range("P" & "2:" & "P" & LastRowToValidate).Cells
    'find the first > and strip everything to its left and the first < (after the value) and strip it and everything to the right
    nFirstPos = 0
    nLastPos = 0
    strTemp = ""
   nFirstPos = InStr(1, c.Value, "<adj:AdjustmentAmount>")
    nLastPos = InStr(1, c.Value, "</adj:AdjustmentAmoun")
    If nFirstPos > 0 And nLastPos > 0 Then
        strTemp = Mid(c.Value, nFirstPos + 22, nLastPos - nFirstPos - 22)
        c.Value = strTemp
    Else
        c.Value = ""
    End If
Next


For Each c In Worksheets("Table").Range("Q" & "2:" & "Q" & LastRowToValidate).Cells
    'find the first > and strip everything to its left and the first < (after the value) and strip it and everything to the right
    nFirstPos = 0
    nLastPos = 0
    strTemp = ""
   nFirstPos = InStr(1, c.Value, "<adj:AdjustmentAmount>")
    nLastPos = InStr(1, c.Value, "</adj:AdjustmentAmount>")
    If nFirstPos > 0 And nLastPos > 0 Then
        strTemp = Mid(c.Value, nFirstPos + 22, nLastPos - nFirstPos - 22)
        c.Value = strTemp
    End If
Next


For Each c In Worksheets("Table").Range("S" & "2:" & "S" & LastRowToValidate).Cells
    'find the first > and strip everything to its left and the first < (after the value) and strip it and everything to the right
    nFirstPos = 0
    nLastPos = 0
    strTemp = ""
   nFirstPos = InStr(1, c.Value, "<adj:AdjustmentAmount>")
    nLastPos = InStr(1, c.Value, "</adj:AdjustmentAmount>")
    If nFirstPos > 0 And nLastPos > 0 Then
        strTemp = Mid(c.Value, nFirstPos + 22, nLastPos - nFirstPos - 22)
        c.Value = strTemp
    End If
Next

For Each c In Worksheets("Table").Range("T" & "2:" & "T" & LastRowToValidate).Cells
    'find the first > and strip everything to its left and the first < (after the value) and strip it and everything to the right
    nFirstPos = 0
    nLastPos = 0
    strTemp = ""
   nFirstPos = InStr(1, c.Value, "<adj:AdjustmentAmount>")
    nLastPos = InStr(1, c.Value, "</adj:AdjustmentAmount>")
    If nFirstPos > 0 And nLastPos > 0 Then
        strTemp = Mid(c.Value, nFirstPos + 22, nLastPos - nFirstPos - 22)
        c.Value = strTemp
    End If
Next

For Each c In Worksheets("Table").Range("U" & "2:" & "U" & LastRowToValidate).Cells
    If InStr(1, c.Value, "_REGIONAL") > 0 Then
        c.Value = "Comparable"
    Else
        If InStr(1, c.Value, "CVDBREGIONAL") > 0 Then
            c.Value = "CVDBREGIONAL"
        Else
            If InStr(1, c.Value, "DEALERQUOTE") > 0 Then
                c.Value = "DEALERQUOTE"
            Else
                If InStr(1, c.Value, "HISTORICAL") > 0 Then
                   c.Value = "HISTORICAL"
                Else
                    If InStr(1, c.Value, "CVDB") > 0 Then
                        c.Value = "CVDB"
                    Else
                        If InStr(1, c.Value, "NADA") > 0 Then
                            c.Value = "Nada"
                        Else
                            If InStr(1, c.Value, "REDBOOK") > 0 Then
                                c.Value = "Redbook"
                            Else
                                c.Value = "Not Available"
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
Next


For Each c In Worksheets("Table").Range("V" & "2:" & "V" & LastRowToValidate).Cells
    'find Name= and strip everything to its left and the first > (after the value) and strip it and everything to the right
    nFirstPos = 0
    nLastPos = 0
    strTemp = ""
    nFirstPos = InStr(1, c.Value, "Name=")
    nLastPos = InStr(nFirstPos + 6, c.Value, Chr(34)) + 1
    If nFirstPos > 0 And nLastPos > 0 Then
        strTemp = Mid(c.Value, nFirstPos + 6, nLastPos - nFirstPos - 7)
        c.Value = strTemp
        officeFound = True
    End If
Next

End Sub

Sub DetermineLastRowToValidate(sheetName As String)
Dim nRowCounter As Long
nRowCounter = 1
Do Until Worksheets(sheetName).Cells(nRowCounter, "B").Value = ""  'col B at this point is UserID
    nRowCounter = nRowCounter + 1
Loop
LastRowToValidate = nRowCounter - 1
End Sub



Sub SeperateOriginals_comp()
    
    Dim CountRev As Long
    Dim CountOrig As Long
    Dim m As Long
    Dim n As Long
    
    DetermineLastRowToValidate "Table"
    
    CountRev = 0
    CountOrig = 0
    m = 2
    n = 2
    d = Worksheets("Table").Cells(2, "B")
    r = Worksheets("Table").Cells(m, "A")
   
    If (r = "") Then
        CountOrig = 1
    Else: CountRev = 1
    End If
    
    Worksheets("OriginalCreators").Cells(n, "A").Value = d
    Worksheets("OriginalCreators").Cells(n, "D") = Worksheets("Table").Cells(m, "D")
    'Worksheets("OriginalCreators").Cells(n, "E") = Worksheets("Table").Cells(m, "E") & " " & Worksheets("Table").Cells(m, "F")
    Worksheets("OriginalCreators").Cells(n, "F") = Worksheets("Table").Cells(m, "G")
    Worksheets("OriginalCreators").Cells(n, "G") = Worksheets("Table").Cells(m, "H")
    Worksheets("OriginalCreators").Cells(n, "H") = Worksheets("Table").Cells(m, "I")
    Worksheets("OriginalCreators").Cells(n, "I") = Worksheets("Table").Cells(m, "J")
    Worksheets("OriginalCreators").Cells(n, "J") = Worksheets("Table").Cells(m, "K")
    Worksheets("OriginalCreators").Cells(n, "K") = Worksheets("Table").Cells(m, "L")
    Worksheets("OriginalCreators").Cells(n, "L") = Worksheets("Table").Cells(m, "M")
    If (CountOrig = 1) Then
            Worksheets("OriginalCreators").Cells(n, "M") = Worksheets("Table").Cells(m, "N")
            Worksheets("OriginalCreators").Cells(n, "O") = Worksheets("Table").Cells(m, "O")
            Worksheets("OriginalCreators").Cells(n, "P") = Worksheets("Table").Cells(m, "P")
            Worksheets("OriginalCreators").Cells(n, "S") = Worksheets("Table").Cells(m, "Q")
            Worksheets("OriginalCreators").Cells(n, "U") = Worksheets("Table").Cells(m, "S")
            Worksheets("OriginalCreators").Cells(n, "W") = Worksheets("Table").Cells(m, "T")
            Worksheets("OriginalCreators").Cells(n, "Z") = Worksheets("Table").Cells(m, "R")
            Worksheets("OriginalCreators").Cells(n, "Y") = Worksheets("Table").Cells(m, "V")
            Worksheets("OriginalCreators").Cells(n, "AA") = Worksheets("Table").Cells(m, "U")
    End If
    If (CountRev = 1) Then
            Worksheets("OriginalCreators").Cells(n, "N") = Worksheets("Table").Cells(m, "N")
            Worksheets("OriginalCreators").Cells(n, "Q") = Worksheets("Table").Cells(m, "O")
            Worksheets("OriginalCreators").Cells(n, "R") = Worksheets("Table").Cells(m, "P")
            Worksheets("OriginalCreators").Cells(n, "T") = Worksheets("Table").Cells(m, "Q")
            Worksheets("OriginalCreators").Cells(n, "E") = Worksheets("Table").Cells(m, "D")
            'Worksheets("OriginalCreators").Cells(n, "S") = Worksheets("Table").Cells(m, "Q")
            Worksheets("OriginalCreators").Cells(n, "V") = Worksheets("Table").Cells(m, "S")
            Worksheets("OriginalCreators").Cells(n, "X") = Worksheets("Table").Cells(m, "T")
            Worksheets("OriginalCreators").Cells(n, "Z") = Worksheets("Table").Cells(m, "R")
            Worksheets("OriginalCreators").Cells(n, "Y") = Worksheets("Table").Cells(m, "V")
            Worksheets("OriginalCreators").Cells(n, "AA") = Worksheets("Table").Cells(m, "U")
    End If
    'Worksheets("OriginalCreators").Cells(n, "I") = Worksheets("Table").Cells(m, "J")
    m = m + 1
    Worksheets("OriginalCreators").Cells(n, "B").Value = CountOrig
    Worksheets("OriginalCreators").Cells(n, "C").Value = CountRev
    
    For Each c In Worksheets("Table").Range("B" & "3:" & "B" & LastRowToValidate).Cells
        If (c = d) Then
        CountRev = CountRev + 1
        CountOrig = 0
        Else:   d = c
            If (Worksheets("Table").Cells(m, "A").Value = "") Then
                CountOrig = 1
                CountRev = 0
            Else: CountRev = 1
                  CountOrig = 0
            End If
            n = n + 1
            Worksheets("OriginalCreators").Cells(n, "B").Value = CountOrig
            Worksheets("OriginalCreators").Cells(n, "A") = d
            Worksheets("OriginalCreators").Cells(n, "D") = Worksheets("Table").Cells(m, "D")
            'Worksheets("OriginalCreators").Cells(n, "E") = Worksheets("Table").Cells(m, "E") & " " & Worksheets("Table").Cells(m, "F")
            Worksheets("OriginalCreators").Cells(n, "F") = Worksheets("Table").Cells(m, "G")
            Worksheets("OriginalCreators").Cells(n, "G") = Worksheets("Table").Cells(m, "H")
            Worksheets("OriginalCreators").Cells(n, "H") = Worksheets("Table").Cells(m, "I")
            Worksheets("OriginalCreators").Cells(n, "I") = Worksheets("Table").Cells(m, "J")
            Worksheets("OriginalCreators").Cells(n, "J") = Worksheets("Table").Cells(m, "K")
            Worksheets("OriginalCreators").Cells(n, "K") = Worksheets("Table").Cells(m, "L")
            Worksheets("OriginalCreators").Cells(n, "L") = Worksheets("Table").Cells(m, "M")
            
        End If
        If (CountOrig = 1) Then
            Worksheets("OriginalCreators").Cells(n, "M") = Worksheets("Table").Cells(m, "N")
            Worksheets("OriginalCreators").Cells(n, "O") = Worksheets("Table").Cells(m, "O")
            Worksheets("OriginalCreators").Cells(n, "P") = Worksheets("Table").Cells(m, "P")
            Worksheets("OriginalCreators").Cells(n, "S") = Worksheets("Table").Cells(m, "Q")
            Worksheets("OriginalCreators").Cells(n, "U") = Worksheets("Table").Cells(m, "S")
            Worksheets("OriginalCreators").Cells(n, "W") = Worksheets("Table").Cells(m, "T")
            Worksheets("OriginalCreators").Cells(n, "Z") = Worksheets("Table").Cells(m, "R")
            Worksheets("OriginalCreators").Cells(n, "Y") = Worksheets("Table").Cells(m, "V")
            Worksheets("OriginalCreators").Cells(n, "AA") = Worksheets("Table").Cells(m, "U")
    
        End If
        If (CountRev = 1) Then
            Worksheets("OriginalCreators").Cells(n, "N") = Worksheets("Table").Cells(m, "N")
            Worksheets("OriginalCreators").Cells(n, "Q") = Worksheets("Table").Cells(m, "O")
            Worksheets("OriginalCreators").Cells(n, "R") = Worksheets("Table").Cells(m, "P")
            Worksheets("OriginalCreators").Cells(n, "T") = Worksheets("Table").Cells(m, "Q")
            Worksheets("OriginalCreators").Cells(n, "E") = Worksheets("Table").Cells(m, "D")
            Worksheets("OriginalCreators").Cells(n, "V") = Worksheets("Table").Cells(m, "S")
            Worksheets("OriginalCreators").Cells(n, "X") = Worksheets("Table").Cells(m, "T")
            Worksheets("OriginalCreators").Cells(n, "Z") = Worksheets("Table").Cells(m, "R")
            Worksheets("OriginalCreators").Cells(n, "Y") = Worksheets("Table").Cells(m, "V")
            Worksheets("OriginalCreators").Cells(n, "AA") = Worksheets("Table").Cells(m, "U")

        End If
        m = m + 1
        Worksheets("OriginalCreators").Cells(n, "C").Value = CountRev
                
                
    Next
    'Format "Loss Date" to mm/dd/yyyy
    Worksheets("OriginalCreators").Activate
    Worksheets("OriginalCreators").Columns("K").Select
    Selection.NumberFormat = "mm/dd/yyyy"
    Worksheets("OriginalCreators").Range("A2:AA" & LastRowToValidate).Font.Size = 8.5
    

    Worksheets("OriginalCreators").Range("A2:AA" & LastRowToValidate).Sort Key1:=Worksheets("OriginalCreators").Range("D2")
    
End Sub



Sub SeperateOriginals_book()
    
    Dim CountRev As Long
    Dim CountOrig As Long
    Dim bflag As String
    Dim m As Long
    Dim n As Long
    
    DetermineLastRowToValidate "Table"
    
    CountRev = 0
    CountOrig = 0
    m = 2
    n = 2
    d = Worksheets("Table").Cells(2, "B")
    r = Worksheets("Table").Cells(m, "A")
   
    If (r = "") Then
        CountOrig = 1
    Else: CountRev = 1
    End If
    
    Worksheets("OriginalCreators").Cells(n, "A").Value = d
    Worksheets("OriginalCreators").Cells(n, "D") = Worksheets("Table").Cells(m, "D")
    'Worksheets("OriginalCreators").Cells(n, "E") = Worksheets("Table").Cells(m, "E") & " " & Worksheets("Table").Cells(m, "F")
    Worksheets("OriginalCreators").Cells(n, "F") = Worksheets("Table").Cells(m, "G")
    Worksheets("OriginalCreators").Cells(n, "G") = Worksheets("Table").Cells(m, "H")
    Worksheets("OriginalCreators").Cells(n, "H") = Worksheets("Table").Cells(m, "I")
    Worksheets("OriginalCreators").Cells(n, "I") = Worksheets("Table").Cells(m, "J")
    Worksheets("OriginalCreators").Cells(n, "J") = Worksheets("Table").Cells(m, "K")
    Worksheets("OriginalCreators").Cells(n, "K") = Worksheets("Table").Cells(m, "L")
    Worksheets("OriginalCreators").Cells(n, "L") = Worksheets("Table").Cells(m, "M")
    If (CountOrig = 1) Then
            If (Worksheets("Table").Cells(m, "U").Value = "Nada") Then
                bflag = "nada"
            Else: bflag = "redbook"
            End If
            Worksheets("OriginalCreators").Cells(n, "M") = Worksheets("Table").Cells(m, "N")
            Worksheets("OriginalCreators").Cells(n, "O") = Worksheets("Table").Cells(m, "O")
            Worksheets("OriginalCreators").Cells(n, "P") = Worksheets("Table").Cells(m, "P")
            Worksheets("OriginalCreators").Cells(n, "S") = Worksheets("Table").Cells(m, "Q")
            Worksheets("OriginalCreators").Cells(n, "U") = Worksheets("Table").Cells(m, "S")
            Worksheets("OriginalCreators").Cells(n, "W") = Worksheets("Table").Cells(m, "T")
            Worksheets("OriginalCreators").Cells(n, "Z") = Worksheets("Table").Cells(m, "R")
            Worksheets("OriginalCreators").Cells(n, "Y") = Worksheets("Table").Cells(m, "V")
            Worksheets("OriginalCreators").Cells(n, "AA") = Worksheets("Table").Cells(m, "U")
    End If
    If (CountRev = 1) Then
            If (Worksheets("Table").Cells(m, "U").Value = "Nada") Then
                bflag = "nada"
            Else: bflag = "redbook"
            End If
            Worksheets("OriginalCreators").Cells(n, "N") = Worksheets("Table").Cells(m, "N")
            Worksheets("OriginalCreators").Cells(n, "Q") = Worksheets("Table").Cells(m, "O")
            Worksheets("OriginalCreators").Cells(n, "R") = Worksheets("Table").Cells(m, "P")
            Worksheets("OriginalCreators").Cells(n, "T") = Worksheets("Table").Cells(m, "Q")
            Worksheets("OriginalCreators").Cells(n, "E") = Worksheets("Table").Cells(m, "D")
            'Worksheets("OriginalCreators").Cells(n, "S") = Worksheets("Table").Cells(m, "Q")
            Worksheets("OriginalCreators").Cells(n, "V") = Worksheets("Table").Cells(m, "S")
            Worksheets("OriginalCreators").Cells(n, "X") = Worksheets("Table").Cells(m, "T")
            Worksheets("OriginalCreators").Cells(n, "Z") = Worksheets("Table").Cells(m, "R")
            Worksheets("OriginalCreators").Cells(n, "Y") = Worksheets("Table").Cells(m, "V")
            Worksheets("OriginalCreators").Cells(n, "AA") = Worksheets("Table").Cells(m, "U")
    End If
    'Worksheets("OriginalCreators").Cells(n, "I") = Worksheets("Table").Cells(m, "J")
    m = m + 1
    Worksheets("OriginalCreators").Cells(n, "B").Value = CountOrig
    Worksheets("OriginalCreators").Cells(n, "C").Value = CountRev
    
    For Each c In Worksheets("Table").Range("B" & "3:" & "B" & LastRowToValidate).Cells
        If (c = d) Then
            If (Worksheets("Table").Cells(m, "A").Value = "") Then
                If (bflag = "redbook") Then
                    CountOrig = 1
                    CountRev = 0
                Else:   CountRev = 0
                        CountOrig = 0
                End If
            Else
                valreq2 = Worksheets("Table").Cells(m, "A").Value
                If (valreq1 = valreq2) Then
                    If (bflag = "redbook") Then
                        CountRev = 1
                        CountOrig = 0
                    Else:   CountRev = CountRev + 1
                            CountOrig = 0
                    End If
                Else
                    CountRev = CountRev + 1
                    CountOrig = 0
                    bflag = "null"
                End If
            End If
        Else:   d = c
            bflag = "null"
            If (Worksheets("Table").Cells(m, "A").Value = "") Then
                CountOrig = 1
                CountRev = 0
            Else: CountRev = 1
                  CountOrig = 0
                  
            End If
            n = n + 1
            Worksheets("OriginalCreators").Cells(n, "B").Value = CountOrig
            Worksheets("OriginalCreators").Cells(n, "A") = d
            Worksheets("OriginalCreators").Cells(n, "D") = Worksheets("Table").Cells(m, "D")
            'Worksheets("OriginalCreators").Cells(n, "E") = Worksheets("Table").Cells(m, "E") & " " & Worksheets("Table").Cells(m, "F")
            Worksheets("OriginalCreators").Cells(n, "F") = Worksheets("Table").Cells(m, "G")
            Worksheets("OriginalCreators").Cells(n, "G") = Worksheets("Table").Cells(m, "H")
            Worksheets("OriginalCreators").Cells(n, "H") = Worksheets("Table").Cells(m, "I")
            Worksheets("OriginalCreators").Cells(n, "I") = Worksheets("Table").Cells(m, "J")
            Worksheets("OriginalCreators").Cells(n, "J") = Worksheets("Table").Cells(m, "K")
            Worksheets("OriginalCreators").Cells(n, "K") = Worksheets("Table").Cells(m, "L")
            Worksheets("OriginalCreators").Cells(n, "L") = Worksheets("Table").Cells(m, "M")
        End If
        If (CountOrig = 1) Then
            
            Worksheets("OriginalCreators").Cells(n, "M") = Worksheets("Table").Cells(m, "N")
            Worksheets("OriginalCreators").Cells(n, "O") = Worksheets("Table").Cells(m, "O")
            Worksheets("OriginalCreators").Cells(n, "P") = Worksheets("Table").Cells(m, "P")
            Worksheets("OriginalCreators").Cells(n, "S") = Worksheets("Table").Cells(m, "Q")
            Worksheets("OriginalCreators").Cells(n, "U") = Worksheets("Table").Cells(m, "S")
            Worksheets("OriginalCreators").Cells(n, "W") = Worksheets("Table").Cells(m, "T")
            Worksheets("OriginalCreators").Cells(n, "Z") = Worksheets("Table").Cells(m, "R")
            Worksheets("OriginalCreators").Cells(n, "Y") = Worksheets("Table").Cells(m, "V")
            Worksheets("OriginalCreators").Cells(n, "AA") = Worksheets("Table").Cells(m, "U")
            If (Worksheets("Table").Cells(m, "U").Value = "Nada") Then
                bflag = "nada"
            Else: bflag = "redbook"
            End If
            valreq1 = Null
        End If
        If (CountRev = 1) Then
          If (bflag = "null" Or bflag = "redbook") Then
          
            Worksheets("OriginalCreators").Cells(n, "N") = Worksheets("Table").Cells(m, "N")
            Worksheets("OriginalCreators").Cells(n, "Q") = Worksheets("Table").Cells(m, "O")
            Worksheets("OriginalCreators").Cells(n, "R") = Worksheets("Table").Cells(m, "P")
            Worksheets("OriginalCreators").Cells(n, "T") = Worksheets("Table").Cells(m, "Q")
            Worksheets("OriginalCreators").Cells(n, "E") = Worksheets("Table").Cells(m, "D")
            Worksheets("OriginalCreators").Cells(n, "V") = Worksheets("Table").Cells(m, "S")
            Worksheets("OriginalCreators").Cells(n, "X") = Worksheets("Table").Cells(m, "T")
            Worksheets("OriginalCreators").Cells(n, "Z") = Worksheets("Table").Cells(m, "R")
            Worksheets("OriginalCreators").Cells(n, "Y") = Worksheets("Table").Cells(m, "V")
            Worksheets("OriginalCreators").Cells(n, "AA") = Worksheets("Table").Cells(m, "U")
            valreq1 = Worksheets("Table").Cells(m, "A").Value
          End If
            If (Worksheets("Table").Cells(m, "U").Value = "Nada") Then
                bflag = "nada"
            Else:   bflag = "redbook"
            End If
        End If
        m = m + 1
        Worksheets("OriginalCreators").Cells(n, "C").Value = CountRev
                
                
    Next
    'Format "Loss Date" to mm/dd/yyyy
    Worksheets("OriginalCreators").Activate
    Worksheets("OriginalCreators").Columns("K").Select
    Selection.NumberFormat = "mm/dd/yyyy"
    Worksheets("OriginalCreators").Range("A2:AA" & LastRowToValidate).Font.Size = 8.5
    
'    Worksheets("OriginalCreators").Range("A2:AA" & LastRowToValidate).Sort Key1:=Worksheets("OriginalCreators").Range("B2")
 
    Worksheets("OriginalCreators").Range("A2:AA" & LastRowToValidate).Sort Key1:=Worksheets("OriginalCreators").Range("D2")
    
End Sub



Sub mapUserId()

'delete PL/SQL column
If Worksheets("mapUserId").Cells(2, "A").Value = "1" Then
   Worksheets("mapUserId").Columns("A").Delete
End If
'determine LastRow count on OnmapUserId Tab
DetermineLastRowToValidate "mapUserId"

LastRowOnmapUserIdTab = LastRowToValidate

'determine LastRow count on OriginalCreators Tab
DetermineLastRowToValidate "OriginalCreators"

'insert columns for "Original USER Name" and "Revised USER Name"
Worksheets("OriginalCreators").Columns("F").Insert
Worksheets("OriginalCreators").Columns("E").Insert
'add column headers
Worksheets("OriginalCreators").Cells(1, "E").Value = "Original USER Name"
Worksheets("OriginalCreators").Cells(1, "G").Value = "Revised USER Name"

Dim l As Long
Dim m As Long

For l = 2 To LastRowOnmapUserIdTab
    For m = 2 To LastRowToValidate
        
        If Worksheets("mapUserId").Cells(l, "A").Value = Worksheets("OriginalCreators").Cells(m, "D").Value Then
             Worksheets("OriginalCreators").Cells(m, "E").Value = Worksheets("mapUserId").Cells(l, "B").Value
        End If

        If Worksheets("mapUserId").Cells(l, "A").Value = Worksheets("OriginalCreators").Cells(m, "F").Value Then
             Worksheets("OriginalCreators").Cells(m, "G").Value = Worksheets("mapUserId").Cells(l, "B").Value
        End If

    Next m
Next l

Worksheets("OriginalCreators").Columns("O:Z").Select
Selection.NumberFormat = "###,###,##0.00"
Selection.HorizontalAlignment = xlRight

'format headers in MS San Serif 9
Worksheets("OriginalCreators").Columns("A:AC").Select   'Cells.Select
    With Selection.Font
        .Name = "MS Sans Serif"
        .Size = 9
    End With
Selection.RowHeight = 12.5
Worksheets("OriginalCreators").Range("A1:AC1").Cells.Select
Selection.RowHeight = 40
'freeze headers
Worksheets("OriginalCreators").Rows("2:2").Select
ActiveWindow.FreezePanes = True

Worksheets("OriginalCreators").Range("A2:AC" & LastRowToValidate).Sort Key1:=Worksheets("OriginalCreators").Range("A2")
End Sub

Sub Finalize()
'
' Finalize1 Macro
' Macro Created 9/29/2014 by Sambhav
'

'
Dim nRowCounter As Long
nRowCounter = 1
Do Until Worksheets("OriginalCreators").Cells(nRowCounter, "A").Value = ""  'col B at this point is UserID
    nRowCounter = nRowCounter + 1
Loop
LastRowToValidate = nRowCounter - 1
Dim range_Temp As String
range_Temp = "x" & LastRowToValidate
    Columns("AA:AA").Select
    Application.CutCopyMode = False
    Selection.Cut
    ActiveWindow.LargeScroll ToRight:=-1
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
    ActiveWindow.LargeScroll ToRight:=1
    Columns("AB:AB").Select
    Selection.Cut
    ActiveWindow.LargeScroll ToRight:=-1
    Columns("O:O").Select
    Selection.Insert Shift:=xlToRight
    ActiveWindow.LargeScroll ToRight:=1
    Columns("AC:AC").Select
    Selection.Cut
    ActiveWindow.LargeScroll ToRight:=-1
    Columns("Q:Q").Select
    Selection.Insert Shift:=xlToRight
    Range("A1:AE1").Select
    With Selection.Interior
        .ColorIndex = 37
        .Pattern = xlSolid
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Range("Z1").Select
    Selection.Copy
    Range("AD1").Select
    ActiveSheet.Paste
    Range("AA1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AE1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Revised Equipment Amount"
    With ActiveCell.Characters(Start:=1, Length:=24).Font
        .Name = "MS Sans Serif"
        .FontStyle = "Bold"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Range("AD1").Select
    ActiveCell.FormulaR1C1 = "Original Equipmentment Amount"
    With ActiveCell.Characters(Start:=1, Length:=29).Font
        .Name = "MS Sans Serif"
        .FontStyle = "Bold"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Range("AC4").Select
    ActiveWindow.LargeScroll ToRight:=-1
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    range_Temp = "X2:Y" & LastRowToValidate
    Range(range_Temp).Select
    Selection.Copy
    ActiveWindow.LargeScroll ToRight:=0
    ActiveWindow.ScrollRow = 2610
    ActiveWindow.ScrollRow = 2461
    ActiveWindow.ScrollRow = 2243
    ActiveWindow.ScrollRow = 2002
    ActiveWindow.ScrollRow = 1744
    ActiveWindow.ScrollRow = 1523
    ActiveWindow.ScrollRow = 1180
    ActiveWindow.ScrollRow = 834
    ActiveWindow.ScrollRow = 627
    ActiveWindow.ScrollRow = 522
    ActiveWindow.ScrollRow = 423
    ActiveWindow.ScrollRow = 335
    ActiveWindow.ScrollRow = 206
    ActiveWindow.ScrollRow = 128
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 2
    Range("AD2").Select
    ActiveSheet.Paste
    ActiveWindow.LargeScroll ToRight:=-1
    Application.CutCopyMode = False
    ActiveWorkbook.Save
    Columns("A:A").ColumnWidth = 15.43
    Columns("D:D").ColumnWidth = 21.29
    Columns("F:F").ColumnWidth = 22.14
    Columns("F:F").ColumnWidth = 27.57
    Columns("H:H").ColumnWidth = 16.71
    Columns("I:I").ColumnWidth = 15.43
    Columns("I:I").ColumnWidth = 21
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.LargeScroll ToRight:=-1
    Range("A21").Select
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 148
    ActiveWindow.ScrollRow = 308
    ActiveWindow.ScrollRow = 426
    ActiveWindow.ScrollRow = 593
    ActiveWindow.ScrollRow = 966
    ActiveWindow.ScrollRow = 1160
    ActiveWindow.ScrollRow = 1388
    ActiveWindow.ScrollRow = 1547
    ActiveWindow.ScrollRow = 1561
    ActiveWindow.ScrollRow = 1595
    ActiveWindow.ScrollRow = 1666
    ActiveWindow.ScrollRow = 1860
    ActiveWindow.ScrollRow = 2026
    ActiveWindow.ScrollRow = 2111
    ActiveWindow.ScrollRow = 2135
    ActiveWindow.ScrollRow = 2223
    ActiveWindow.ScrollRow = 2243
    ActiveWindow.ScrollRow = 2267
    ActiveWindow.ScrollRow = 2284
    ActiveWindow.ScrollRow = 2325
    ActiveWindow.ScrollRow = 2379
    ActiveWindow.ScrollRow = 2410
    ActiveWindow.ScrollRow = 2420
    ActiveWindow.ScrollRow = 2423
    ActiveWindow.ScrollRow = 2467
    ActiveWindow.ScrollRow = 2562
    ActiveWindow.ScrollRow = 2573
    ActiveWindow.ScrollRow = 2579
    ActiveWindow.ScrollRow = 2637
    ActiveWindow.ScrollRow = 2698
    ActiveWindow.ScrollRow = 2756
    ActiveWindow.ScrollRow = 2773
    ActiveWindow.ScrollRow = 2753
    ActiveWindow.ScrollRow = 2705
    ActiveWindow.ScrollRow = 2678
    ActiveWindow.ScrollRow = 2654
    ActiveWindow.ScrollRow = 2644
    ActiveWindow.ScrollRow = 2641
    range_Temp = "A" & LastRowToValidate & ":AE" & LastRowToValidate
    Range(range_Temp).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    range_Temp = "AE2:AE" & LastRowToValidate
    Range(range_Temp).Select
    range_Temp = "AE" & LastRowToValidate
    Range(range_Temp).Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveWindow.LargeScroll ToRight:=-1
    ActiveWindow.ScrollColumn = 1
    range_Temp = "E2:AE" & LastRowToValidate
    Range(range_Temp).Select
    ActiveWorkbook.Save
    ActiveWindow.LargeScroll ToRight:=-2
End Sub

Sub Beautify()
'
' Beautify Macro
' Created By Sambhav Patni
' Creation Date : 10/28/2014
'
    Rows("1:1").Select
    With Selection
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("F:F").EntireColumn.AutoFit
    Columns("E:E").EntireColumn.AutoFit
    Columns("E:E").Select
    Columns("H:H").EntireColumn.AutoFit
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    Columns("I:I").ColumnWidth = 19.29
    Columns("K:K").EntireColumn.AutoFit
    Columns("L:L").EntireColumn.AutoFit
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.SmallScroll Down:=-15
    ActiveWindow.LargeScroll ToRight:=-2
    Rows("1:1").RowHeight = 43.5
    Columns("Q:Q").ColumnWidth = 15.75
    Columns("Z:Z").ColumnWidth = 11.7
    Columns("AA:AA").ColumnWidth = 12
    Columns("AB:AB").ColumnWidth = 11.45
    Columns("AC:AC").ColumnWidth = 11.45
    Columns("AD:AD").ColumnWidth = 12.75
    Columns("AE:AE").ColumnWidth = 9.0
    ActiveWorkbook.Save
End Sub
