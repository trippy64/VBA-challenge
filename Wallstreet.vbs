Attribute VB_Name = "Module1"


Sub stockticker()
Dim currentticker As String
Dim openprice As Double
Dim closeprice As Double
Dim volume As Double
Dim opendate As Integer
Dim closedate As Integer
Dim Summary_Table As Integer
Dim lrow As Double
Dim changeprice As Double
Dim percentchange As Double
Dim MaxrChange As Range
Dim MinChange As Range
Dim MaxVolRng As Range
Dim maxpercent As Double
Dim minpercent As Double
Dim maxvolume As Double

'set the initial row count from 2 to final row
'set the summary row to 2

lrow = Cells(Rows.Count, 1).End(xlUp).Row
Summary_Table = 2
volume = 0
j = 2



For i = 2 To lrow



    If Cells(i + 1, 1).Value <> Cells(j, 1).Value Then

        currentticker = Cells(i, 1).Value
        volume = volume + Cells(i, 7).Value
        changeprice = Cells(i, 6).Value - Cells(j, 3).Value
        
        If Cells(i, 3).Value <> 0 Then
            percentchange = changeprice / Cells(i, 3).Value
                Else: percentchange = 0
            End If
             

        Range("I" & Summary_Table).Value = currentticker
        Range("L" & Summary_Table).Value = volume
        Range("J" & Summary_Table).Value = changeprice
        Range("K" & Summary_Table).Value = percentchange
        'conditinal formatting for year change
        If Range("J" & Summary_Table).Value >= 0 Then
            Range("J" & Summary_Table).Interior.ColorIndex = 4
        ElseIf Range("J" & Summary_Table).Value < 0 Then
            Range("J" & Summary_Table).Interior.ColorIndex = 3
        End If
       ' add 1 to the summary table
        Summary_Table = Summary_Table + 1
        'Reset volume
                volume = 0
                j = i + 1
        ElseIf Cells(i + 1, 1).Value = Cells(i, 1).Value Then
            volume = volume + Cells(i, 7).Value




  End If

  ' end loop
Next i
  'Get the summary data compiled to a min and max range, and set the range

  k = 2



  maxpercent = Application.WorksheetFunction.Max(ActiveSheet.Range("K2:K999"))
  minpercent = Application.WorksheetFunction.Min(ActiveSheet.Range("K2:K999"))
  maxvolume = Application.WorksheetFunction.Max(ActiveSheet.Range("L2:L999"))


    For k = 2 To 9999
  'find and print the highest percentage change
  If Cells(k, 11).Value = maxpercent Then
    Cells(2, 16).Value = maxpercent
    Cells(2, 15).Value = Cells(k, 9).Value

  'find and print the lowest percentage change
  ElseIf Cells(k, 11).Value = minpercent Then
    Cells(3, 16).Value = minpercent
    Cells(3, 15).Value = Cells(k, 9).Value

  End If

  'find and print the highest volume
  If Cells(k, 12).Value = maxvolume Then
    Cells(4, 16).Value = maxvolume
    Cells(4, 15).Value = Cells(k, 9).Value

  End If


  Next k



End Sub


