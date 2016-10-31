
Sub matrice()

Dim lastindex As Integer
Dim Nblignes As Integer
Dim i As Integer
Dim l As Integer

Dim a As Variant
Dim aa As Variant
Dim b As Variant
Dim bb As Variant
Dim c As Variant
Dim cc As Variant
Dim d As Variant
Dim mat1 As Variant
Dim mat2 As Variant

Dim k As Integer
Dim h As Integer
Dim j As Integer

Dim tab1 As Variant
Dim range1 As Variant
Dim range2 As Variant
Dim range3 As Variant

Dim mess As String



'ActiveWorkbook.Sheets("index").Activate
'Range(Cells(4, 2), Cells(100, 4)) = ""
'ActiveWorkbook.Sheets("omega").Activate
'Range(Cells(4, 2), Cells(100, 3)) = ""
'ActiveWorkbook.Sheets("omegainv").Activate
'Range(Cells(4, 2), Cells(100, 3)) = ""
'ActiveWorkbook.Sheets("transpose").Activate
'Range(Cells(4, 2), Cells(100, 3)) = ""


ActiveWorkbook.Sheets("index").Activate
l = Cells(2, 17).Value - 1

Range(Cells(3, 18), Cells(3, 18).Offset(l, l)).Select
Range(Cells(25, 18), Cells(25, 18).Offset(l, l)) = Application.WorksheetFunction.MInverse(Selection)

lastindex = Cells(500, 3).End(xlUp).Row

ReDim tab1(lastindex - 3, 3)

k = 1
For h = 2 To 4
    j = 0
    For i = 4 To lastindex
        tab1(j, k) = Cells(i, h)
        j = j + 1
    Next i
    k = k + 1
Next h


dercol = (lastindex - 3) + 29

k = 1
For i = 3 To 6
    j = 0
    For h = 29 To dercol
        Cells(i, h) = tab1(j, k)
        On Error Resume Next
        j = j + 1
    Next h
    k = k + 1
Next i


lastindex = Cells(500, 4).End(xlUp).Row
Range(Cells(4, 3), Cells(lastindex, 3)).Select
range1 = Selection

lastindex = Cells(550, 20).End(xlUp).Row
Range(Cells(25, 18), Cells(lastindex, 20)).Select
range2 = Selection

Range(Cells(5, 29), Cells(5, 29 + l)).Select
range3 = Selection

a = Application.WorksheetFunction.MMult(range3, range2)
aa = Application.WorksheetFunction.MMult(a, range1)
Range("I3") = aa

Range(Cells(4, 29), Cells(4, 29 + l)).Select
range3 = Selection

b = Application.WorksheetFunction.MMult(range3, range2)
bb = Application.WorksheetFunction.MMult(b, range1)
Range("J3") = bb

lastindex = Cells(500, 4).End(xlUp).Row
Range(Cells(4, 4), Cells(lastindex, 4)).Select
range1 = Selection

Range(Cells(5, 29), Cells(5, 29 + l)).Select
range3 = Selection

c = Application.WorksheetFunction.MMult(range3, range2)
cc = Application.WorksheetFunction.MMult(c, range1)
Range("K3") = cc

Range("L3") = Range("J3") * Range("K3") - (Range("I3") ^ 2)
Range("D3").Select


lastindex = Cells(500, 4).End(xlUp).Row
Range(Cells(4, 3), Cells(lastindex, 3)).Select
range1 = Selection
lastindex = Cells(500, 4).End(xlUp).Row
Range(Cells(4, 4), Cells(lastindex, 4)).Select
range3 = Selection


Range("BM1").Select
'Calcul G
Range(Cells(1, 65), Cells(l + 1, 65)) = Application.WorksheetFunction.MMult(range2, range3)
Range(Cells(1, 66), Cells(l + 1, 66)) = Range("J3")
Range(Cells(1, 67), Cells(l + 1, 67)) = Application.WorksheetFunction.MMult(range2, range1)
Range(Cells(1, 68), Cells(l + 1, 68)) = Range("I3")

For i = 1 To l + 1
Cells(i, 69) = Cells(i, 65) * Cells(i, 66)
Next

For i = 1 To l + 1
Cells(i, 70) = Cells(i, 67) * Cells(i, 68)
Next

For i = 1 To l + 1
Cells(i, 71) = Cells(i, 69) - Cells(i, 70)
Next

Range(Cells(1, 72), Cells(l + 1, 72)) = Range("L3")

For i = 1 To l + 1
Cells(i, 73) = Cells(i, 71) / Cells(i, 72)
Next

Range(Cells(1, 73), Cells(l + 1, 73)).Select
Selection.Copy
Range("M3").Select
ActiveSheet.Paste

'Calcul HH
Range(Cells(1, 74), Cells(l + 1, 74)) = Range("K3")

For i = 1 To l + 1
Cells(i, 75) = Cells(i, 74) * Cells(i, 67)
Next

For i = 1 To l + 1
Cells(i, 76) = Cells(i, 68) * Cells(i, 65)
Next

For i = 1 To l + 1
Cells(i, 77) = Cells(i, 75) - Cells(i, 76)
Next

For i = 1 To l + 1
Cells(i, 78) = Cells(i, 77) / Cells(i, 72)
Next

Range(Cells(1, 78), Cells(l + 1, 78)).Select
Selection.Copy
Range("N3").Select
ActiveSheet.Paste



mess = InputBox("Veuillez définir votre taux sans risque", "Taux sans risque")
    If mess <> "" Then
    End If

Range(Cells(1, 79), Cells(l + 1, 79)) = mess

Range(Cells(1, 79), Cells(l + 1, 79)).Select

