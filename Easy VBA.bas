Attribute VB_Name = "Module1"
Sub HW2()

Dim sh As Worksheet
For Each sh In ActiveWorkbook.Sheets
    sh.Activate
'copy column
Range("A:A").Copy Range("I:I")
'paste column
Range("I:I").Value = Range("A:A").Value
'remove duplicates
Range("I:I").RemoveDuplicates Columns:=1
'Name Headers for answers
    Range("I1").Value = "Stock Name"
            Range("I1").Font.Bold = True
            Cells.Columns.AutoFit

    Range("J1").Value = "Total Volume"
            Range("J1").Font.Bold = True
            Cells.Columns.AutoFit
Next sh
End Sub
Sub ticker()

Dim r As Double
Dim c As Double
Dim Stockname As String
Dim Vol As Double
Dim Lastrow As Double
Dim LRC As Double

Dim sh As Worksheet
For Each sh In ActiveWorkbook.Sheets
    sh.Activate

Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
LRC = Cells(Rows.Count, 9).End(xlUp).Row

Dim total As Double



For c = 2 To LRC
Stockname = Cells(c, 9).Value
total = 0
    For r = 2 To Lastrow

   

        If Cells(r, 1).Value = Stockname Then
            total = total + Cells(r, 7).Value
    
        Else
            total = total + 0
    
        End If

    
    Next r

Cells(c, 10).Value = total

Next c
Next sh
End Sub

