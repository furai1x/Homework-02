Attribute VB_Name = "Homework02"
Option Explicit

Sub StockVolume_Easy()

    Dim lgRows As Long
    Dim lgCols As Long
    Dim strTickerVolume(1000, 2) As Variant
    Dim i As Long
    Dim j As Long
    Dim blnExists As Boolean
    Dim lgCount As Long
    Dim ws As Worksheet


    
    
    For Each ws In ActiveWorkbook.Worksheets

        ws.Activate

        Range("A1").Select
    
        lgRows = Cells(Rows.Count, 1).End(xlUp).Row
        lgCols = Cells(1, Columns.Count).End(xlToLeft).Column
        
           
           
        strTickerVolume(0, 0) = Cells(2, 1).Value
        strTickerVolume(0, 1) = Cells(2, 7).Value
        
        lgCount = 0
        
        
        blnExists = False
        
        For i = 3 To lgRows
            For j = 0 To lgCount
                If strTickerVolume(j, 0) = Cells(i, 1).Value Then
                    strTickerVolume(j, 1) = strTickerVolume(j, 1) + Cells(i, 7).Value
                    blnExists = True
                End If
            Next j
                If blnExists = False Then
                    lgCount = lgCount + 1
                    strTickerVolume(lgCount, 0) = Cells(i, 1).Value
                    strTickerVolume(lgCount, 1) = Cells(i, 7).Value
                End If
            blnExists = False
        Next i
        
        Cells(1, lgCols + 2).Value = "Ticker"
        Cells(1, lgCols + 3).Value = "Total Stock Volume"
        
        For i = 2 To lgCount + 2
            Cells(i, lgCols + 2).Value = strTickerVolume(i - 2, 0)
            Cells(i, lgCols + 3).Value = strTickerVolume(i - 2, 1)
        Next i
    Next ws
    

End Sub
Sub StockVolume_Hard()

    Dim lgRows As Long
    Dim lgCols As Long
    Dim strTickerVolume(4000, 4) As Variant
    Dim i As Long
    Dim j As Long
    Dim blnExists As Boolean
    Dim lgCount As Long
    Dim ws As Worksheet
    Dim strDelim As String
    Dim strRange As String
    Dim strIncDecTTL(2, 1) As Variant
    Dim dblPer As Double
    
    strDelim = "$"
    
    
    For Each ws In ActiveWorkbook.Worksheets

        Erase strTickerVolume
        Erase strIncDecTTL

        ws.Activate

        Range("A1").Select
    
        lgRows = Cells(Rows.Count, 1).End(xlUp).Row
        lgCols = Cells(1, Columns.Count).End(xlToLeft).Column
           
        strTickerVolume(0, 0) = Cells(2, 1).Value
        strTickerVolume(0, 1) = Cells(2, 7).Value
        strTickerVolume(0, 2) = Cells(2, 3).Value
        
        lgCount = 0
        
        
        blnExists = False
        
        For i = 3 To lgRows
            For j = 0 To lgCount
                If strTickerVolume(j, 0) = Cells(i, 1).Value Then
                    strTickerVolume(j, 1) = strTickerVolume(j, 1) + Cells(i, 7).Value
                    blnExists = True
                End If
            Next j
                If blnExists = False Then
                    lgCount = lgCount + 1
                    strTickerVolume(lgCount - 1, 3) = Cells(i - 1, 6).Value
                    strTickerVolume(lgCount, 0) = Cells(i, 1).Value
                    strTickerVolume(lgCount, 1) = Cells(i, 7).Value
                    strTickerVolume(lgCount, 2) = Cells(i, 3).Value
                End If
            blnExists = False
            If i = lgRows Then
                strTickerVolume(lgCount, 3) = Cells(i, 6).Value
            End If
        Next i
        
        Cells(1, lgCols + 2).Value = "Ticker"
        Cells(1, lgCols + 3).Value = "Yearly Change"
        Cells(1, lgCols + 4).Value = "Percent Change"
        Cells(1, lgCols + 5).Value = "Total Stock Volume"
        
        
         
        If strTickerVolume(0, 2) = 0 Then
            dblPer = 999999
        Else:  dblPer = (strTickerVolume(0, 3) - strTickerVolume(0, 2)) / strTickerVolume(0, 2)
        End If
         
        strIncDecTTL(0, 0) = strTickerVolume(0, 0)
        strIncDecTTL(0, 1) = dblPer
        strIncDecTTL(1, 0) = strTickerVolume(0, 0)
        strIncDecTTL(1, 1) = dblPer
        strIncDecTTL(2, 0) = strTickerVolume(0, 0)
        strIncDecTTL(2, 1) = strTickerVolume(0, 1)
        
        
        For i = 2 To lgCount + 2
            Cells(i, lgCols + 2).Value = strTickerVolume(i - 2, 0)
            Cells(i, lgCols + 3).Value = strTickerVolume(i - 2, 3) - strTickerVolume(i - 2, 2)
            If strTickerVolume(i - 2, 2) = 0 Then
                If strTickerVolume(i - 2, 3) = 0 Then
                    Cells(i, lgCols + 4).Value = 0
                Else: Cells(i, lgCols + 4).Value = 999999
                End If
            Else: Cells(i, lgCols + 4).Value = (strTickerVolume(i - 2, 3) - strTickerVolume(i - 2, 2)) / strTickerVolume(i - 2, 2)
            End If
            Cells(i, lgCols + 5).Value = strTickerVolume(i - 2, 1)
            
            If strTickerVolume(i - 2, 1) > strIncDecTTL(2, 1) Then
                strIncDecTTL(2, 1) = strTickerVolume(i - 2, 1)
                strIncDecTTL(2, 0) = strTickerVolume(i - 2, 0)
            End If
            
            If strTickerVolume(i - 2, 2) <> 0 Then
                If (strTickerVolume(i - 2, 3) - strTickerVolume(i - 2, 2)) / strTickerVolume(i - 2, 2) > strIncDecTTL(0, 1) Then
                    strIncDecTTL(0, 1) = (strTickerVolume(i - 2, 3) - strTickerVolume(i - 2, 2)) / strTickerVolume(i - 2, 2)
                    strIncDecTTL(0, 0) = strTickerVolume(i - 2, 0)
                ElseIf (strTickerVolume(i - 2, 3) - strTickerVolume(i - 2, 2)) / strTickerVolume(i - 2, 2) < strIncDecTTL(1, 1) Then
                    strIncDecTTL(1, 1) = (strTickerVolume(i - 2, 3) - strTickerVolume(i - 2, 2)) / strTickerVolume(i - 2, 2)
                    strIncDecTTL(1, 0) = strTickerVolume(i - 2, 0)
                End If
            End If
        Next i
        
        
        strRange = Split(Cells(1, lgCols + 3).Address, strDelim)(1) & "2:" & Split(Cells(1, lgCols + 3).Address, strDelim)(1) & Trim(Str(lgCount + 2))
        
        Range(strRange).Select
        
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
            Formula1:="=0"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 5287936
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=0"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False

        strRange = Split(Cells(1, lgCols + 4).Address, strDelim)(1) & "2:" & Split(Cells(1, lgCols + 4).Address, strDelim)(1) & Trim(Str(lgCount + 2))
        
        Range(strRange).Select
    
        Selection.Style = "Percent"
    
        Cells(1, lgCols + 9).Value = "Ticker"
        Cells(1, lgCols + 10).Value = "Value"
        Cells(2, lgCols + 8).Value = "Greatest % Increase"
        Cells(2, lgCols + 9).Value = strIncDecTTL(0, 0)
        Cells(2, lgCols + 10).Value = strIncDecTTL(0, 1)
        Cells(3, lgCols + 8).Value = "Greatest % Decrease"
        Cells(3, lgCols + 9).Value = strIncDecTTL(1, 0)
        Cells(3, lgCols + 10).Value = strIncDecTTL(1, 1)
        Cells(4, lgCols + 8).Value = "Greatest Total Volume"
        Cells(4, lgCols + 9).Value = strIncDecTTL(2, 0)
        Cells(4, lgCols + 10).Value = strIncDecTTL(2, 1)
        
        strRange = Cells(2, lgCols + 10).Address & ":" & Cells(3, lgCols + 10).Address
        
        Range(strRange).Select
    
        Selection.Style = "Percent"
    
    
    Next ws
    

End Sub

