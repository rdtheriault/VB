Sub btn_run()

'define variables
Dim closeC, farC, extremeC, closeS, farS, extremeS, count, t1, studentC, t1t, t2t, t3t, t4t As Integer
Dim teacher1, t1c, t1e, teacher2, t2c, t2f, t2e, teacher3, t3c, t3f, t3e, teacher4, t4c, t4f, t4e As String
Dim leng As Integer


'set variables
t1 = 2
t1t = 0
t2t = 0
t3t = 0
t4t = 0
closeS = Worksheets("Data").Range("E3").Value
farS = Worksheets("Data").Range("F3").Value
extremeS = Worksheets("Data").Range("G3").Value
teacher1 = Worksheets("Data").Cells(t1, 3).Value
closeC = 0
farC = 0
extremeC = 0
studentC = 0

'Clear errors
Worksheets("Data").Range("F15").Value = ""


'Teacher one
If (teacher1 Like "") Then
Worksheets("Data").Range("F15").Value = "You did not enter a teacher"
Exit Sub
End If
Do While teacher1 Like Worksheets("Data").Cells(t1, 3).Value
    If Worksheets("Data").Cells(t1, 1).Value <= extremeS And Worksheets("Data").Cells(t1, 1).Value > 0 Then
        t1e = t1e & Worksheets("Data").Cells(t1, 2).Value & ", "
        extremeC = extremeC + 1
        t1t = t1t + 1
    End If
    If Worksheets("Data").Cells(t1, 1).Value <= farS And Worksheets("Data").Cells(t1, 1).Value > extremeS Then
        t1f = t1f & Worksheets("Data").Cells(t1, 2).Value & ", "
        farC = farC + 1
        t1t = t1t + 1
    End If
    If Worksheets("Data").Cells(t1, 1).Value <= closeS And Worksheets("Data").Cells(t1, 1).Value > farS Then
        t1c = t1c & Worksheets("Data").Cells(t1, 2).Value & ", "
        closeC = closeC + 1
        t1t = t1t + 1
    End If
    
    t1 = t1 + 1
Loop
Worksheets("Table 1").Range("B4").Value = teacher1
Worksheets("Table 1").Range("I4").Value = t1e
Worksheets("Table 1").Range("H4").Value = t1f
Worksheets("Table 1").Range("G4").Value = t1c
Worksheets("Table 1").Range("F4").Value = t1t
Worksheets("Table 1").Range("B4").Font.Color = vbBlack
Worksheets("Table 1").Range("I4").Font.Color = vbBlack
Worksheets("Table 1").Range("H4").Font.Color = vbBlack
Worksheets("Table 1").Range("G4").Font.Color = vbBlack
Worksheets("Table 1").Range("I15").Value = extremeC
Worksheets("Table 1").Range("H15").Value = farC
Worksheets("Table 1").Range("G15").Value = closeC

'Teacher two
teacher2 = Worksheets("Data").Cells(t1, 3).Value
If (teacher2 Like "") Then
Worksheets("Data").Range("F15").Value = "You did not enter a second teacher (if you have one)"
Exit Sub
End If
Do While teacher2 Like Worksheets("Data").Cells(t1, 3).Value
    If Worksheets("Data").Cells(t1, 1).Value <= extremeS And Worksheets("Data").Cells(t1, 1).Value > 0 Then
        t2e = t2e & Worksheets("Data").Cells(t1, 2).Value & ", "
        extremeC = extremeC + 1
        t2t = t2t + 1
    End If
    If Worksheets("Data").Cells(t1, 1).Value <= farS And Worksheets("Data").Cells(t1, 1).Value > extremeS Then
        t2f = t2f & Worksheets("Data").Cells(t1, 2).Value & ", "
        farC = farC + 1
        t2t = t2t + 1
    End If
    If Worksheets("Data").Cells(t1, 1).Value <= closeS And Worksheets("Data").Cells(t1, 1).Value > farS Then
        t2c = t2c & Worksheets("Data").Cells(t1, 2).Value & ", "
        closeC = closeC + 1
        t2t = t2t + 1
    End If
     
    t1 = t1 + 1
Loop
Worksheets("Table 1").Range("B6").Value = teacher2
Worksheets("Table 1").Range("I6").Value = t2e
Worksheets("Table 1").Range("H6").Value = t2f
Worksheets("Table 1").Range("G6").Value = t2c
Worksheets("Table 1").Range("F6").Value = t2t
Worksheets("Table 1").Range("B6").Font.Color = vbBlack
Worksheets("Table 1").Range("I6").Font.Color = vbBlack
Worksheets("Table 1").Range("H6").Font.Color = vbBlack
Worksheets("Table 1").Range("G6").Font.Color = vbBlack
Worksheets("Table 1").Range("I15").Value = extremeC
Worksheets("Table 1").Range("H15").Value = farC
Worksheets("Table 1").Range("G15").Value = closeC

'Teacher three
teacher3 = Worksheets("Data").Cells(t1, 3).Value
If (teacher3 Like "") Then
Worksheets("Data").Range("F15").Value = "You did not enter a third teacher (if you have one)"
Exit Sub
End If
Do While teacher3 Like Worksheets("Data").Cells(t1, 3).Value
    If Worksheets("Data").Cells(t1, 1).Value <= extremeS And Worksheets("Data").Cells(t1, 1).Value > 0 Then
        t3e = t3e & Worksheets("Data").Cells(t1, 2).Value & ", "
        extremeC = extremeC + 1
        t3t = t3t + 1
    End If
    If Worksheets("Data").Cells(t1, 1).Value <= farS And Worksheets("Data").Cells(t1, 1).Value > extremeS Then
        t3f = t3f & Worksheets("Data").Cells(t1, 2).Value & ", "
        farC = farC + 1
        t3t = t3t + 1
    End If
    If Worksheets("Data").Cells(t1, 1).Value <= closeS And Worksheets("Data").Cells(t1, 1).Value > farS Then
        t3c = t3c & Worksheets("Data").Cells(t1, 2).Value & ", "
        closeC = closeC + 1
        t3t = t3t + 1
    End If
     
    t1 = t1 + 1
Loop
Worksheets("Table 1").Range("F11").Value = teacher3
Worksheets("Table 1").Range("I11").Value = t3e
Worksheets("Table 1").Range("H11").Value = t3f
Worksheets("Table 1").Range("G11").Value = t3c
Worksheets("Table 1").Range("F11").Value = t3t
Worksheets("Table 1").Range("B11").Font.Color = vbBlack
Worksheets("Table 1").Range("I11").Font.Color = vbBlack
Worksheets("Table 1").Range("H11").Font.Color = vbBlack
Worksheets("Table 1").Range("G11").Font.Color = vbBlack
Worksheets("Table 1").Range("I15").Value = extremeC
Worksheets("Table 1").Range("H15").Value = farC
Worksheets("Table 1").Range("G15").Value = closeC

'Teacher four
teacher4 = Worksheets("Data").Cells(t1, 3).Value
If (teacher4 Like "") Then
Worksheets("Data").Range("F15").Value = "You did not enter a fourth teacher (if you have one)"
Exit Sub
End If
Do While teacher4 Like Worksheets("Data").Cells(t1, 3).Value
     If Worksheets("Data").Cells(t1, 1).Value <= extremeS And Worksheets("Data").Cells(t1, 1).Value > 0 Then
        t4e = t4e & Worksheets("Data").Cells(t1, 2).Value & ", "
        extremeC = extremeC + 1
        t4t = t4t + 1
    End If
    If Worksheets("Data").Cells(t1, 1).Value <= farS And Worksheets("Data").Cells(t1, 1).Value > extremeS Then
        t4f = t4f & Worksheets("Data").Cells(t1, 2).Value & ", "
        farC = farC + 1
        t4t = t4t + 1
    End If
    If Worksheets("Data").Cells(t1, 1).Value <= closeS And Worksheets("Data").Cells(t1, 1).Value > farS Then
        t4c = t4c & Worksheets("Data").Cells(t1, 2).Value & ", "
        closeC = closeC + 1
        t4t = t4t + 1
    End If

    t1 = t1 + 1
Loop
Worksheets("Table 1").Range("B13").Value = teacher4
Worksheets("Table 1").Range("I13").Value = t4e
Worksheets("Table 1").Range("H13").Value = t4f
Worksheets("Table 1").Range("G13").Value = t4c
Worksheets("Table 1").Range("F13").Value = t4t
Worksheets("Table 1").Range("B13").Font.Color = vbBlack
Worksheets("Table 1").Range("I13").Font.Color = vbBlack
Worksheets("Table 1").Range("H13").Font.Color = vbBlack
Worksheets("Table 1").Range("G13").Font.Color = vbBlack
Worksheets("Table 1").Range("I15").Value = extremeC
Worksheets("Table 1").Range("H15").Value = farC
Worksheets("Table 1").Range("G15").Value = closeC


'Just incase colors went wrong
Worksheets("Table 1").Range("I15").Font.Color = vbBlack
Worksheets("Table 1").Range("H15").Font.Color = vbBlack
Worksheets("Table 1").Range("G15").Font.Color = vbBlack


End Sub
Sub btn_clear()


Worksheets("Table 1").Range("B4").Value = ""
Worksheets("Table 1").Range("I4").Value = ""
Worksheets("Table 1").Range("H4").Value = ""
Worksheets("Table 1").Range("G4").Value = ""
Worksheets("Table 1").Range("F4").Value = ""
Worksheets("Table 1").Range("B6").Value = ""
Worksheets("Table 1").Range("I6").Value = ""
Worksheets("Table 1").Range("H6").Value = ""
Worksheets("Table 1").Range("G6").Value = ""
Worksheets("Table 1").Range("F6").Value = ""
Worksheets("Table 1").Range("B11").Value = ""
Worksheets("Table 1").Range("I11").Value = ""
Worksheets("Table 1").Range("H11").Value = ""
Worksheets("Table 1").Range("G11").Value = ""
Worksheets("Table 1").Range("F11").Value = ""
Worksheets("Table 1").Range("B13").Value = ""
Worksheets("Table 1").Range("I13").Value = ""
Worksheets("Table 1").Range("H13").Value = ""
Worksheets("Table 1").Range("G13").Value = ""
Worksheets("Table 1").Range("F13").Value = ""
Worksheets("Table 1").Range("I15").Value = ""
Worksheets("Table 1").Range("H15").Value = ""
Worksheets("Table 1").Range("G15").Value = ""

Worksheets("Table 1").Range("B5").Value = ""
Worksheets("Table 1").Range("H5").Value = ""
Worksheets("Table 1").Range("G5").Value = ""
Worksheets("Table 1").Range("I5").Value = ""
Worksheets("Table 1").Range("F5").Value = ""
Worksheets("Table 1").Range("B7").Value = ""
Worksheets("Table 1").Range("H7").Value = ""
Worksheets("Table 1").Range("G7").Value = ""
Worksheets("Table 1").Range("I7").Value = ""
Worksheets("Table 1").Range("F7").Value = ""
Worksheets("Table 1").Range("B12").Value = ""
Worksheets("Table 1").Range("H12").Value = ""
Worksheets("Table 1").Range("G12").Value = ""
Worksheets("Table 1").Range("I12").Value = ""
Worksheets("Table 1").Range("F12").Value = ""
Worksheets("Table 1").Range("B14").Value = ""
Worksheets("Table 1").Range("H14").Value = ""
Worksheets("Table 1").Range("G14").Value = ""
Worksheets("Table 1").Range("I14").Value = ""
Worksheets("Table 1").Range("F14").Value = ""
Worksheets("Table 1").Range("H16").Value = ""
Worksheets("Table 1").Range("G16").Value = ""
Worksheets("Table 1").Range("I16").Value = ""

'Clear errors
Worksheets("Data").Range("F21").Value = ""
Worksheets("Data").Range("F15").Value = ""


End Sub
Sub btn_test()

'define variables
Dim closeC, farC, extremeC, closeS, farS, extremeS, count, t1, studentC, t1t, t2t, t3t, t4t As Integer
Dim teacher1, t1c, t1f, t1e, teacher2, t2c, t2f, t2e, teacher3, t3c, t3f, t3e, teacher4, t4c, t4f, t4e As String
Dim leng, lenge, lengc, lengH, lengHe, lengHc, arrayL, arrayLe1, arrayLc1, i, beg, le As Integer
Dim fonts(1 To 300, 1 To 3) As String
Dim fontsE(1 To 300, 1 To 3) As String
Dim fontsC(1 To 300, 1 To 3) As String
'Dim co As String

'set variables
t1 = 2
t1t = 0
t2t = 0
t3t = 0
t4t = 0
closeS = Worksheets("Data").Range("E3").Value
farS = Worksheets("Data").Range("F3").Value
extremeS = Worksheets("Data").Range("G3").Value
teacher1 = Worksheets("Data").Cells(t1, 3).Value
closeC = 0
farC = 0
extremeC = 0
studentC = 0
lengH = 0
lengHe = 0
lengHc = 0
arrayL = 1
arrayLe1 = 1
arrayLc1 = 1

'Clear errors
Worksheets("Data").Range("F15").Value = ""


'Teacher one
If (teacher1 Like "") Then
Worksheets("Data").Range("F15").Value = "You did not enter a teacher"
Exit Sub
End If
Do While teacher1 Like Worksheets("Data").Cells(t1, 3).Value
    If Worksheets("Data").Cells(t1, 1).Value <= extremeS And Worksheets("Data").Cells(t1, 1).Value > 0 Then
        holder = Worksheets("Data").Cells(t1, 2).Value
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbBlack Then
            lengHe = lengHe + lenge + 2
            lenge = Len(holder)
            fontsE(arrayLe1, 1) = lengHe - 1
            fontsE(arrayLe1, 2) = lenge
            fontsE(arrayLe1, 3) = vbBlack
            arrayLe1 = arrayLe1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 12611584 Then
            lengHe = lengHe + lenge + 2
            lenge = Len(holder)
            fontsE(arrayLe1, 1) = lengHe - 1
            fontsE(arrayLe1, 2) = lenge
            fontsE(arrayLe1, 3) = vbBlue
            arrayLe1 = arrayLe1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbRed Then
            lengHe = lengHe + lenge + 2
            lenge = Len(holder)
            fontsE(arrayLe1, 1) = lengHe - 1
            fontsE(arrayLe1, 2) = lenge
            fontsE(arrayLe1, 3) = vbRed
            arrayLe1 = arrayLe1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 10498160 Then
            lengHe = lengHe + lenge + 2
            lenge = Len(holder)
            fontsE(arrayLe1, 1) = lengHe - 1
            fontsE(arrayLe1, 2) = lenge
            fontsE(arrayLe1, 3) = 10498160
            arrayLe1 = arrayLe1 + 1
        End If
        t1e = t1e & holder & ", "
        extremeC = extremeC + 1
        t1t = t1t + 1
    End If
    If Worksheets("Data").Cells(t1, 1).Value <= farS And Worksheets("Data").Cells(t1, 1).Value > extremeS Then
        holder = Worksheets("Data").Cells(t1, 2).Value
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbBlack Then
            lengH = lengH + leng + 2
            leng = Len(holder)
            fonts(arrayL, 1) = lengH - 1
            fonts(arrayL, 2) = leng
            fonts(arrayL, 3) = vbBlack
            arrayL = arrayL + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 12611584 Then
            lengH = lengH + leng + 2
            leng = Len(holder)
            fonts(arrayL, 1) = lengH - 1
            fonts(arrayL, 2) = leng
            fonts(arrayL, 3) = vbBlue
            arrayL = arrayL + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbRed Then
            lengH = lengH + leng + 2
            leng = Len(holder)
            fonts(arrayL, 1) = lengH - 1
            fonts(arrayL, 2) = leng
            fonts(arrayL, 3) = vbRed
            arrayL = arrayL + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 10498160 Then
            lengH = lengH + leng + 2
            leng = Len(holder)
            fonts(arrayL, 1) = lengH - 1
            fonts(arrayL, 2) = leng
            fonts(arrayL, 3) = 10498160
            arrayL = arrayL + 1
        End If
        t1f = t1f & holder & ", "
        farC = farC + 1
        t1t = t1t + 1
    End If
    If Worksheets("Data").Cells(t1, 1).Value <= closeS And Worksheets("Data").Cells(t1, 1).Value > farS Then
        holder = Worksheets("Data").Cells(t1, 2).Value
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbBlack Then
            lengHc = lengHc + lengc + 2
            lengc = Len(holder)
            fontsC(arrayLc1, 1) = lengHc - 1
            fontsC(arrayLc1, 2) = lengc
            fontsC(arrayLc1, 3) = vbBlack
            arrayLc1 = arrayLc1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 12611584 Then
            lengHc = lengHc + lengc + 2
            lengc = Len(holder)
            fontsC(arrayLc1, 1) = lengHc - 1
            fontsC(arrayLc1, 2) = lengc
            fontsC(arrayLc1, 3) = vbBlue
            arrayLc1 = arrayLc1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbRed Then
            lengHc = lengHc + lengc + 2
            lengc = Len(holder)
            fontsC(arrayLc1, 1) = lengHc - 1
            fontsC(arrayLc1, 2) = lengc
            fontsC(arrayLc1, 3) = vbRed
            arrayLc1 = arrayLc1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 10498160 Then
            lengHc = lengHc + lengc + 2
            lengc = Len(holder)
            fontsC(arrayLc1, 1) = lengHc - 1
            fontsC(arrayLc1, 2) = lengc
            fontsC(arrayLc1, 3) = 10498160
            arrayLc1 = arrayLc1 + 1
        End If
        t1c = t1c & holder & ", "
        closeC = closeC + 1
        t1t = t1t + 1
    End If
    t1 = t1 + 1
Loop
'Teacher1 add to other sheet
Worksheets("Table 1").Range("I4").Value = t1e
For i = 1 To 300
If fontsE(i, 1) = "" Then
Exit For
Else
    beg = fontsE(i, 1)
    le = fontsE(i, 2)
    col = fontsE(i, 3)
    Worksheets("Table 1").Range("I4").Characters(beg, le).Font.Color = col
End If
Next i
Worksheets("Table 1").Range("H4").Value = t1f
For i = 1 To 300
If fonts(i, 1) = "" Then
Exit For
Else
    beg = fonts(i, 1)
    le = fonts(i, 2)
    col = fonts(i, 3)
    Worksheets("Table 1").Range("H4").Characters(beg, le).Font.Color = col
End If
Next i
Worksheets("Table 1").Range("G4").Value = t1c
For i = 1 To 300
If fontsC(i, 1) = "" Then
Exit For
Else
    beg = fontsC(i, 1)
    le = fontsC(i, 2)
    col = fontsC(i, 3)
    Worksheets("Table 1").Range("G4").Characters(beg, le).Font.Color = col
End If
Next i
Worksheets("Table 1").Range("B4").Value = teacher1
Worksheets("Table 1").Range("F4").Value = t1t
Worksheets("Table 1").Range("I15").Value = extremeC
Worksheets("Table 1").Range("H15").Value = farC
Worksheets("Table 1").Range("G15").Value = closeC


'Teacher two
lengH = 0
lengHe = 0
lengHc = 0
lengc = 0
lenge = 0
leng = 0
arrayL = 1
arrayLe1 = 1
arrayLc1 = 1
Dim fonts2(1 To 300, 1 To 3) As String
Dim fontsE2(1 To 300, 1 To 3) As String
Dim fontsC2(1 To 300, 1 To 3) As String

teacher2 = Worksheets("Data").Cells(t1, 3).Value
If (teacher2 Like "") Then
Worksheets("Data").Range("F15").Value = "You did not enter a second teacher (if you have one)"
Exit Sub
End If
Do While teacher2 Like Worksheets("Data").Cells(t1, 3).Value
    If Worksheets("Data").Cells(t1, 1).Value <= extremeS And Worksheets("Data").Cells(t1, 1).Value > 0 Then
        holder = Worksheets("Data").Cells(t1, 2).Value
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbBlack Then
            lengHe = lengHe + lenge + 2
            lenge = Len(holder)
            fontsE2(arrayLe1, 1) = lengHe - 1
            fontsE2(arrayLe1, 2) = lenge
            fontsE2(arrayLe1, 3) = vbBlack
            arrayLe1 = arrayLe1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 12611584 Then
            lengHe = lengHe + lenge + 2
            lenge = Len(holder)
            fontsE2(arrayLe1, 1) = lengHe - 1
            fontsE2(arrayLe1, 2) = lenge
            fontsE2(arrayLe1, 3) = vbBlue
            arrayLe1 = arrayLe1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbRed Then
            lengHe = lengHe + lenge + 2
            lenge = Len(holder)
            fontsE2(arrayLe1, 1) = lengHe - 1
            fontsE2(arrayLe1, 2) = lenge
            fontsE2(arrayLe1, 3) = vbRed
            arrayLe1 = arrayLe1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 10498160 Then
            lengHe = lengHe + lenge + 2
            lenge = Len(holder)
            fontsE2(arrayLe1, 1) = lengHe - 1
            fontsE2(arrayLe1, 2) = lenge
            fontsE2(arrayLe1, 3) = 10498160
            arrayLe1 = arrayLe1 + 1
        End If
        t2e = t2e & holder & ", "
        extremeC = extremeC + 1
        t2t = t2t + 1
    End If
    If Worksheets("Data").Cells(t1, 1).Value <= farS And Worksheets("Data").Cells(t1, 1).Value > extremeS Then
        holder = Worksheets("Data").Cells(t1, 2).Value
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbBlack Then
            lengH = lengH + leng + 2
            leng = Len(holder)
            fonts2(arrayL, 1) = lengH - 1
            fonts2(arrayL, 2) = leng
            fonts2(arrayL, 3) = vbBlack
            arrayL = arrayL + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 12611584 Then
            lengH = lengH + leng + 2
            leng = Len(holder)
            fonts2(arrayL, 1) = lengH - 1
            fonts2(arrayL, 2) = leng
            fonts2(arrayL, 3) = vbBlue
            arrayL = arrayL + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbRed Then
            lengH = lengH + leng + 2
            leng = Len(holder)
            fonts2(arrayL, 1) = lengH - 1
            fonts2(arrayL, 2) = leng
            fonts2(arrayL, 3) = vbRed
            arrayL = arrayL + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 10498160 Then
            lengH = lengH + leng + 2
            leng = Len(holder)
            fonts2(arrayL, 1) = lengH - 1
            fonts2(arrayL, 2) = leng
            fonts2(arrayL, 3) = 10498160
            arrayL = arrayL + 1
        End If
        t2f = t2f & holder & ", "
        farC = farC + 1
        t2t = t2t + 1
    End If
    If Worksheets("Data").Cells(t1, 1).Value <= closeS And Worksheets("Data").Cells(t1, 1).Value > farS Then
        holder = Worksheets("Data").Cells(t1, 2).Value
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbBlack Then
            lengHc = lengHc + lengc + 2
            lengc = Len(holder)
            fontsC2(arrayLc1, 1) = lengHc - 1
            fontsC2(arrayLc1, 2) = lengc
            fontsC2(arrayLc1, 3) = vbBlack
            arrayLc1 = arrayLc1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 12611584 Then
            lengHc = lengHc + lengc + 2
            lengc = Len(holder)
            fontsC2(arrayLc1, 1) = lengHc - 1
            fontsC2(arrayLc1, 2) = lengc
            fontsC2(arrayLc1, 3) = vbBlue
            arrayLc1 = arrayLc1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbRed Then
            lengHc = lengHc + lengc + 2
            lengc = Len(holder)
            fontsC2(arrayLc1, 1) = lengHc - 1
            fontsC2(arrayLc1, 2) = lengc
            fontsC2(arrayLc1, 3) = vbRed
            arrayLc1 = arrayLc1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 10498160 Then
            lengHc = lengHc + lengc + 2
            lengc = Len(holder)
            fontsC2(arrayLc1, 1) = lengHc - 1
            fontsC2(arrayLc1, 2) = lengc
            fontsC2(arrayLc1, 3) = 10498160
            arrayLc1 = arrayLc1 + 1
        End If
        t2c = t2c & holder & ", "
        closeC = closeC + 1
        t2t = t2t + 1
    End If
    
    t1 = t1 + 1
Loop
'Teacher2 add to other sheet
Worksheets("Table 1").Range("I6").Value = t2e
For i = 1 To 300
If fontsE2(i, 1) = "" Then
Exit For
Else
    beg = fontsE2(i, 1)
    le = fontsE2(i, 2)
    col = fontsE2(i, 3)
    Worksheets("Table 1").Range("I6").Characters(beg, le).Font.Color = col
End If
Next i
Worksheets("Table 1").Range("H6").Value = t2f
For i = 1 To 300
If fonts2(i, 1) = "" Then
Exit For
Else
    beg = fonts2(i, 1)
    le = fonts2(i, 2)
    col = fonts2(i, 3)
    Worksheets("Table 1").Range("H6").Characters(beg, le).Font.Color = col
End If
Next i
Worksheets("Table 1").Range("G6").Value = t2c
For i = 1 To 300
If fontsC2(i, 1) = "" Then
Exit For
Else
    beg = fontsC2(i, 1)
    le = fontsC2(i, 2)
    col = fontsC2(i, 3)
    Worksheets("Table 1").Range("G6").Characters(beg, le).Font.Color = col
End If
Next i
Worksheets("Table 1").Range("B6").Value = teacher2
Worksheets("Table 1").Range("F6").Value = t2t
Worksheets("Table 1").Range("I15").Value = extremeC
Worksheets("Table 1").Range("H15").Value = farC
Worksheets("Table 1").Range("G15").Value = closeC


'Teacher three
lengH = 0
lengHe = 0
lengHc = 0
lengc = 0
lenge = 0
leng = 0
arrayL = 1
arrayLe1 = 1
arrayLc1 = 1
Dim fonts3(1 To 300, 1 To 3) As String
Dim fontsE3(1 To 300, 1 To 3) As String
Dim fontsC3(1 To 300, 1 To 3) As String

teacher3 = Worksheets("Data").Cells(t1, 3).Value
If (teacher3 Like "") Then
Worksheets("Data").Range("F15").Value = "You did not enter a third teacher (if you have one)"
Exit Sub
End If
Do While teacher3 Like Worksheets("Data").Cells(t1, 3).Value
    If Worksheets("Data").Cells(t1, 1).Value <= extremeS And Worksheets("Data").Cells(t1, 1).Value > 0 Then
        holder = Worksheets("Data").Cells(t1, 2).Value
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbBlack Then
            lengHe = lengHe + lenge + 2
            lenge = Len(holder)
            fontsE3(arrayLe1, 1) = lengHe - 1
            fontsE3(arrayLe1, 2) = lenge
            fontsE3(arrayLe1, 3) = vbBlack
            arrayLe1 = arrayLe1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 12611584 Then
            lengHe = lengHe + lenge + 2
            lenge = Len(holder)
            fontsE3(arrayLe1, 1) = lengHe - 1
            fontsE3(arrayLe1, 2) = lenge
            fontsE3(arrayLe1, 3) = vbBlue
            arrayLe1 = arrayLe1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbRed Then
            lengHe = lengHe + lenge + 2
            lenge = Len(holder)
            fontsE3(arrayLe1, 1) = lengHe - 1
            fontsE3(arrayLe1, 2) = lenge
            fontsE3(arrayLe1, 3) = vbRed
            arrayLe1 = arrayLe1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 10498160 Then
            lengHe = lengHe + lenge + 2
            lenge = Len(holder)
            fontsE3(arrayLe1, 1) = lengHe - 1
            fontsE3(arrayLe1, 2) = lenge
            fontsE3(arrayLe1, 3) = 10498160
            arrayLe1 = arrayLe1 + 1
        End If
        t3e = t3e & holder & ", "
        extremeC = extremeC + 1
        t3t = t3t + 1
    End If
    If Worksheets("Data").Cells(t1, 1).Value <= farS And Worksheets("Data").Cells(t1, 1).Value > extremeS Then
        holder = Worksheets("Data").Cells(t1, 2).Value
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbBlack Then
            lengH = lengH + leng + 2
            leng = Len(holder)
            fonts3(arrayL, 1) = lengH - 1
            fonts3(arrayL, 2) = leng
            fonts3(arrayL, 3) = vbBlack
            arrayL = arrayL + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 12611584 Then
            lengH = lengH + leng + 2
            leng = Len(holder)
            fonts3(arrayL, 1) = lengH - 1
            fonts3(arrayL, 2) = leng
            fonts3(arrayL, 3) = vbBlue
            arrayL = arrayL + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbRed Then
            lengH = lengH + leng + 2
            leng = Len(holder)
            fonts3(arrayL, 1) = lengH - 1
            fonts3(arrayL, 2) = leng
            fonts3(arrayL, 3) = vbRed
            arrayL = arrayL + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 10498160 Then
            lengH = lengH + leng + 2
            leng = Len(holder)
            fonts3(arrayL, 1) = lengH - 1
            fonts3(arrayL, 2) = leng
            fonts3(arrayL, 3) = 10498160
            arrayL = arrayL + 1
        End If
        t3f = t3f & holder & ", "
        farC = farC + 1
        t3t = t3t + 1
    End If
    If Worksheets("Data").Cells(t1, 1).Value <= closeS And Worksheets("Data").Cells(t1, 1).Value > farS Then
        holder = Worksheets("Data").Cells(t1, 2).Value
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbBlack Then
            lengHc = lengHc + lengc + 2
            lengc = Len(holder)
            fontsC3(arrayLc1, 1) = lengHc - 1
            fontsC3(arrayLc1, 2) = lengc
            fontsC3(arrayLc1, 3) = vbBlack
            arrayLc1 = arrayLc1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 12611584 Then
            lengHc = lengHc + lengc + 2
            lengc = Len(holder)
            fontsC3(arrayLc1, 1) = lengHc - 1
            fontsC3(arrayLc1, 2) = lengc
            fontsC3(arrayLc1, 3) = vbBlue
            arrayLc1 = arrayLc1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbRed Then
            lengHc = lengHc + lengc + 2
            lengc = Len(holder)
            fontsC3(arrayLc1, 1) = lengHc - 1
            fontsC3(arrayLc1, 2) = lengc
            fontsC3(arrayLc1, 3) = vbRed
            arrayLc1 = arrayLc1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 10498160 Then
            lengHc = lengHc + lengc + 2
            lengc = Len(holder)
            fontsC3(arrayLc1, 1) = lengHc - 1
            fontsC3(arrayLc1, 2) = lengc
            fontsC3(arrayLc1, 3) = 10498160
            arrayLc1 = arrayLc1 + 1
        End If
        t3c = t3c & holder & ", "
        closeC = closeC + 1
        t3t = t3t + 1
    End If
    
    t1 = t1 + 1
Loop
'Teacher3
Worksheets("Table 1").Range("I11").Value = t3e
For i = 1 To 300
If fontsE3(i, 1) = "" Then
Exit For
Else
    beg = fontsE3(i, 1)
    le = fontsE3(i, 2)
    col = fontsE3(i, 3)
    Worksheets("Table 1").Range("I11").Characters(beg, le).Font.Color = col
End If
Next i
Worksheets("Table 1").Range("H11").Value = t3f
For i = 1 To 300
If fonts3(i, 1) = "" Then
Exit For
Else
    beg = fonts3(i, 1)
    le = fonts3(i, 2)
    col = fonts3(i, 3)
    Worksheets("Table 1").Range("H11").Characters(beg, le).Font.Color = col
End If
Next i
Worksheets("Table 1").Range("G11").Value = t3c
For i = 1 To 300
If fontsC3(i, 1) = "" Then
Exit For
Else
    beg = fontsC3(i, 1)
    le = fontsC3(i, 2)
    col = fontsC3(i, 3)
    Worksheets("Table 1").Range("G11").Characters(beg, le).Font.Color = col
End If
Next i
Worksheets("Table 1").Range("B11").Value = teacher3
Worksheets("Table 1").Range("F11").Value = t3t
Worksheets("Table 1").Range("I15").Value = extremeC
Worksheets("Table 1").Range("H15").Value = farC
Worksheets("Table 1").Range("G15").Value = closeC


'Teacher four
lengH = 0
lengHe = 0
lengHc = 0
lengc = 0
lenge = 0
leng = 0
arrayL = 1
arrayLe1 = 1
arrayLc1 = 1
Dim fonts4(1 To 300, 1 To 3) As String
Dim fontsE4(1 To 300, 1 To 3) As String
Dim fontsC4(1 To 300, 1 To 3) As String

teacher4 = Worksheets("Data").Cells(t1, 3).Value
If (teacher4 Like "") Then
Worksheets("Data").Range("F15").Value = "You did not enter a fourth teacher (if you have one)"
Exit Sub
End If
Do While teacher4 Like Worksheets("Data").Cells(t1, 3).Value
    If Worksheets("Data").Cells(t1, 1).Value <= extremeS And Worksheets("Data").Cells(t1, 1).Value > 0 Then
        holder = Worksheets("Data").Cells(t1, 2).Value
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbBlack Then
            lengHe = lengHe + lenge + 2
            lenge = Len(holder)
            fontsE4(arrayLe1, 1) = lengHe - 1
            fontsE4(arrayLe1, 2) = lenge
            fontsE4(arrayLe1, 3) = vbBlack
            arrayLe1 = arrayLe1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 12611584 Then
            lengHe = lengHe + lenge + 2
            lenge = Len(holder)
            fontsE4(arrayLe1, 1) = lengHe - 1
            fontsE4(arrayLe1, 2) = lenge
            fontsE4(arrayLe1, 3) = vbBlue
            arrayLe1 = arrayLe1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbRed Then
            lengHe = lengHe + lenge + 2
            lenge = Len(holder)
            fontsE4(arrayLe1, 1) = lengHe - 1
            fontsE4(arrayLe1, 2) = lenge
            fontsE4(arrayLe1, 3) = vbRed
            arrayLe1 = arrayLe1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 10498160 Then
            lengHe = lengHe + lenge + 2
            lenge = Len(holder)
            fontsE4(arrayLe1, 1) = lengHe - 1
            fontsE4(arrayLe1, 2) = lenge
            fontsE4(arrayLe1, 3) = 10498160
            arrayLe1 = arrayLe1 + 1
        End If
        t4e = t4e & holder & ", "
        extremeC = extremeC + 1
        t4t = t4t + 1
    End If
    If Worksheets("Data").Cells(t1, 1).Value <= farS And Worksheets("Data").Cells(t1, 1).Value > extremeS Then
        holder = Worksheets("Data").Cells(t1, 2).Value
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbBlack Then
            lengH = lengH + leng + 2
            leng = Len(holder)
            fonts4(arrayL, 1) = lengH - 1
            fonts4(arrayL, 2) = leng
            fonts4(arrayL, 3) = vbBlack
            arrayL = arrayL + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 12611584 Then
            lengH = lengH + leng + 2
            leng = Len(holder)
            fonts4(arrayL, 1) = lengH - 1
            fonts4(arrayL, 2) = leng
            fonts4(arrayL, 3) = vbBlue
            arrayL = arrayL + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbRed Then
            lengH = lengH + leng + 2
            leng = Len(holder)
            fonts4(arrayL, 1) = lengH - 1
            fonts4(arrayL, 2) = leng
            fonts4(arrayL, 3) = vbRed
            arrayL = arrayL + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 10498160 Then
            lengH = lengH + leng + 2
            leng = Len(holder)
            fonts4(arrayL, 1) = lengH - 1
            fonts4(arrayL, 2) = leng
            fonts4(arrayL, 3) = 10498160
            arrayL = arrayL + 1
        End If
        t4f = t4f & holder & ", "
        farC = farC + 1
        t4t = t4t + 1
    End If
    If Worksheets("Data").Cells(t1, 1).Value <= closeS And Worksheets("Data").Cells(t1, 1).Value > farS Then
        holder = Worksheets("Data").Cells(t1, 2).Value
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbBlack Then
            lengHc = lengHc + lengc + 2
            lengc = Len(holder)
            fontsC4(arrayLc1, 1) = lengHc - 1
            fontsC4(arrayLc1, 2) = lengc
            fontsC4(arrayLc1, 3) = vbBlack
            arrayLc1 = arrayLc1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 12611584 Then
            lengHc = lengHc + lengc + 2
            lengc = Len(holder)
            fontsC4(arrayLc1, 1) = lengHc - 1
            fontsC4(arrayLc1, 2) = lengc
            fontsC4(arrayLc1, 3) = vbBlue
            arrayLc1 = arrayLc1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = vbRed Then
            lengHc = lengHc + lengc + 2
            lengc = Len(holder)
            fontsC4(arrayLc1, 1) = lengHc - 1
            fontsC4(arrayLc1, 2) = lengc
            fontsC4(arrayLc1, 3) = vbRed
            arrayLc1 = arrayLc1 + 1
        End If
        If Worksheets("Data").Cells(t1, 2).Font.Color = 10498160 Then
            lengHc = lengHc + lengc + 2
            lengc = Len(holder)
            fontsC4(arrayLc1, 1) = lengHc - 1
            fontsC4(arrayLc1, 2) = lengc
            fontsC4(arrayLc1, 3) = 10498160
            arrayLc1 = arrayLc1 + 1
        End If
        t4c = t4c & holder & ", "
        closeC = closeC + 1
        t4t = t4t + 1
    End If
    
    t1 = t1 + 1
Loop
'Teacher4
Worksheets("Table 1").Range("I13").Value = t4e
For i = 1 To 300
If fontsE4(i, 1) = "" Then
Exit For
Else
    beg = fontsE4(i, 1)
    le = fontsE4(i, 2)
    col = fontsE4(i, 3)
    Worksheets("Table 1").Range("I13").Characters(beg, le).Font.Color = col
End If
Next i
Worksheets("Table 1").Range("H13").Value = t4f
For i = 1 To 300
If fonts4(i, 1) = "" Then
Exit For
Else
    beg = fonts4(i, 1)
    le = fonts4(i, 2)
    col = fonts4(i, 3)
    Worksheets("Table 1").Range("H13").Characters(beg, le).Font.Color = col
End If
Next i
Worksheets("Table 1").Range("G13").Value = t4c
For i = 1 To 300
If fontsC4(i, 1) = "" Then
Exit For
Else
    beg = fontsC4(i, 1)
    le = fontsC4(i, 2)
    col = fontsC4(i, 3)
    Worksheets("Table 1").Range("G13").Characters(beg, le).Font.Color = col
End If
Next i
Worksheets("Table 1").Range("B13").Value = teacher4
Worksheets("Table 1").Range("F13").Value = t4t
Worksheets("Table 1").Range("I15").Value = extremeC
Worksheets("Table 1").Range("H15").Value = farC
Worksheets("Table 1").Range("G15").Value = closeC

End Sub
