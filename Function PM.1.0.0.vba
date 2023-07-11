'Base Setting :
Option Explicit

Function PM(fc As Double, fy As Double, D1 As Double, B2 As Double, _
Ccover As Double, Nr2s As Integer, Nr3s As Integer, Nr2sin As Integer, Nr3sin As Integer, _
sizer As Integer, sizestirp As Integer, Pu As Double, Mu2 As Double, Mu3 As Double)
'Purpose :
'   determine stress ratio of rectangle section loaded by biaxial moments and axial force
'Synopsis :
'   Stress ratio
'   PM(fc, fy, D1, B2, Cc, Nr2s, Nr3s, Nr2sin, Nr3sin, sizer, sizestir, Pu, Mu2, Mu3)
' Copyright :
'   Copyright 2023 Yen-Ting,Lin All rights reserved
'Variable Description:
'   fc - concrete compressive strength(kgf/cm2)
'   fy - yield strength of rebar(kgf/cm2)
'   D1 - dimensional length of (2). side of section(cm)
'   B2 - dimensional length of (3). side of section(cm)
'   Ccover - clear cover(cm)
'   Nr2s - number of rebar of (2). side of outer layer
'   Nr3s - number of rebar of (3). side of outer layer
'   Nr2sin - number of rebar of (2). side of inner layer
'   Nr3sin - number of rebar of (3). side of inner layer
'   sizer - size of rebar
'   sizestirp - size of stirrup
'   Pu - design axial force(tf)
'   Mu2(tf-m) - design moment of dir.(2)
'   Mu3(tf-m) - design moment of dir.(3)
'   theta - rotation on M2-M3 diagram, aka. rotation degree of section(rad)
'   e - eccentrial distence(cm)
'   delta - rotation on P-M diagram(rad)
'   seccoord(4x3 array) - coord. of section corner
'   rcoord(Nrx2 array) - coord. of raber location in seciton
'   pointE(cm,Double) - dimension of transformed section in dir.(2)
'   pointF2s(cm,Double) - dimension of transformed section in dir.(2)
'   pointF3s(cm,Double) - dimension of transformed section in dir.(3)
'   c(cm,Double) - depth of neutral axis
'----------------------------------------------------
'Accelerate :
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'----------------------------------------------------
'Declaration :
    Dim rcoord As Variant
    Dim theta As Double
    Dim e As Double, delta As Double, beta As Double
    Dim Ax As Double, Ay As Double, Bx As Double, By As Double
    Dim Cx As Double, Cy As Double, Dx As Double, Dy As Double
    Dim pointE As Variant, pointF2s As Double, pointF3s As Double
    Dim c As Double, x0 As Double, y0 As Double, x1 As Double, y1 As Double
    Dim P As Double, Pmax As Double
    Dim temp As Variant, tempCc As Variant, tempCs As Variant, tempTs As Variant
    Dim i As Integer, j As Integer, Nr As Integer
'-------------------------------------------------------------
'Allocate & Conversion of units :
    Mu2 = (Mu2 + 0.000001) * 100000            'unit: kgf-cm
    Mu3 = (Mu3 + 0.000001) * 100000            'unit: kgf-cm
    Pu = (Pu + 0.000001) * 1000                       'unit: kgf
    If Nr2sin + Nr3sin = 0 Then
        Nr = (Nr2s + Nr3s) * 2 - 4
    Else
        Nr = (Nr2s + Nr3s + Nr2sin + Nr3sin) * 2 - 8
    End If
    ReDim tempCc(1 To 2), tempCs(1 To (Nr - 1), 1 To 2), tempTs(1 To (Nr - 1), 1 To 2)
'-------------------------------------------------------------
    theta = WorksheetFunction.Pi() / 2 - Atn(Abs(Mu3 / Mu2)) + 0.0000000001
    rcoord = coordtransformation(rebarcoord(D1, B2, Ccover, Nr2s, Nr3s, Nr2sin, Nr3sin, sizer, sizestirp), theta)
    e = ((Mu2) ^ 2 + (Mu3) ^ 2) ^ 0.5 / Pu
'section corner coord.
    temp = transformedsection(D1, B2, theta)
    pointE = temp(1)
    pointF2s = temp(2)
    pointF3s = temp(3)
'--------------------------------------------------------------
'calculate Pmax
    c = 10000000000#
    tempCc = Cc(fc, c, pointE, pointF2s, pointF3s) 'staring index 1
    tempCs = Cs(fc, fy, c, pointE, rcoord, sizer) 'starting index 0
    tempTs = Ts(fy, c, pointE, rcoord, sizer) 'starting index 1
    Pmax = tempCc(1)
    For j = 1 To Nr
        Pmax = Pmax + tempCs(j, 2) - tempTs(j, 2)
    Next
    Pmax = 0.8 * Pmax
'--------------------------------------------------------------
' Nonlinear Solver
' Newton-Raphson method
' IF column is under tension case, then initializing c -> 0.
    If Pu < 0 Then
        c = 0.00000001
        Else
' IF column is under pressure case, then initializing c -> 2*pointE.
        c = pointE * 2
    End If
' Maxiter = 500
    For i = 1 To 500
        x0 = c
        tempCc = Cc(fc, x0, pointE, pointF2s, pointF3s)
        tempCs = Cs(fc, fy, x0, pointE, rcoord, sizer)
        tempTs = Ts(fy, x0, pointE, rcoord, sizer)
        P = tempCc(1)
        For j = 1 To Nr
            P = P + tempCs(j, 2) - tempTs(j, 2)
        Next

        If P > Pmax Then
            P = Pmax
        End If
        P = phic(rcoord, tempTs, fy) * P
        y0 = phic(rcoord, tempTs, fy) * Mn(x0, tempCc, tempCs, tempTs, fy, pointE, rcoord) / P - e
               
        x1 = c + 0.0000001
        tempCc = Cc(fc, x1, pointE, pointF2s, pointF3s)
        tempCs = Cs(fc, fy, x1, pointE, rcoord, sizer)
        tempTs = Ts(fy, x1, pointE, rcoord, sizer)
        P = tempCc(1)
        For j = 1 To Nr
            P = P + tempCs(j, 2) - tempTs(j, 2)
        Next

        If P > Pmax Then
            P = Pmax
        End If
        P = phic(rcoord, tempTs, fy) * P
        y1 = (phic(rcoord, tempTs, fy) * Mn(x1, tempCc, tempCs, tempTs, fy, pointE, rcoord)) / P - e
' if converage
        If Abs(y1) < 0.000001 Then
            c = x1
            Exit For
        Else
' didn 't converage , reinitialize c
            c = x2(x0, y0, x1, y1, 0)
' if c goes to unlitimed or negative ,then reinitialize
            If Abs(c) > (50 * pointE) Or c < 0 Then
                c = pointE * Rnd()
            End If
        End If
   Next i
'--------------------------------------------------
'Nonlinear result
    If Abs(y1) > 0.00001 Then
        PM = "NOT CONVERAGE"
    Else
       tempCc = Cc(fc, c, pointE, pointF2s, pointF3s)
       tempCs = Cs(fc, fy, c, pointE, rcoord, sizer)
       tempTs = Ts(fy, c, pointE, rcoord, sizer)
       P = tempCc(1)
       For j = 1 To Nr
           P = P + tempCs(j, 2) - tempTs(j, 2)
       Next

       If P > Pmax Then
           P = Pmax
       End If
       P = phic(rcoord, tempTs, fy) * P
       PM = ((phic(rcoord, tempTs, fy) * Mn(c, tempCc, tempCs, tempTs, fy, pointE, rcoord)) ^ 2 + P ^ 2) ^ 0.5
       PM = ((Pu) ^ 2 + (Mu2) ^ 2 + (Mu3) ^ 2) ^ 0.5 / PM 'D/C
    End If
    Erase rcoord, temp, tempCc, tempCs, tempTs
End Function

Function transformedsection(D1 As Double, B2 As Double, ByVal theta0 As Double)
'Purpose :
'   transform rectangular section into hexagon
'Synopsis :
'   pointE,pointF2s,pointF3s
'Variable Description:
'   D1 - dimensional length of (2). side of section(cm)
'   B2 - dimensional length of (3). side of section(cm)
'   theta0 - rotation on M2-M3 diagram, aka. rotation degree of section(rad)
'   Ax - section corner coord.
'   Ay - section corner coord.
'   Bx - section corner coord.
'   By - section corner coord.
'   Cx - section corner coord.
'   Cy - section corner coord.
'   Dx - section corner coord.
'   Dy - section corner coord.
'   seccoord(4x3 array) - coord. of section corner
'   pointE(cm,Double) - dimension of transformed section in dir.(2)
'   pointF2s(cm,Double) - dimension of transformed section in dir.(2)
'   pointF3s(cm,Double) - dimension of transformed section in dir.(3)
'   tempseccoord  - tempora
'----------------------------------------------------
'Accelerate :
'----------------------------------------------------
'Declaration :
    Dim tempseccoord As Variant, temp As Variant, seccoord As Variant
    Dim Ax As Double, Ay As Double, Bx As Double, By As Double
    Dim Cx As Double, Cy As Double, Dx As Double, Dy  As Double
    Dim pointE As Double, pointF2s As Double, pointF3s As Double
    ReDim tempseccord(1 To 4, 1 To 2), temp(1 To 3), seccoord(1 To 4, 1 To 2)
'----------------------------------------------------
'Allocate & Conversion of units :
'----------------------------------------------------
'section corner coord.
    seccoord(1, 1) = -D1 / 2
    seccoord(1, 2) = -B2 / 2
    seccoord(2, 1) = D1 / 2
    seccoord(2, 2) = -B2 / 2
    seccoord(3, 1) = D1 / 2
    seccoord(3, 2) = B2 / 2
    seccoord(4, 1) = -D1 / 2
    seccoord(4, 2) = B2 / 2
    tempseccoord = coordtransformation(seccoord, theta0)
    Ax = tempseccoord(1, 1)
    Ay = tempseccoord(1, 2)
    Bx = tempseccoord(2, 1)
    By = tempseccoord(2, 2)
    Cx = tempseccoord(3, 1)
    Cy = tempseccoord(3, 2)
    Dx = tempseccoord(4, 1)
    Dy = tempseccoord(4, 2)
'Tranformed section dimension
    theta0 = -theta0 + WorksheetFunction.Pi() / 2
    If D1 <= B2 Then
        If Abs(theta0) <= Atn(D1 / B2) Then
            pointE = Abs(Bx)
            pointF2s = Abs(Ax)
            pointF3s = Abs(Ay) + y2(Cx, Cy, Bx, By, Ax)
        ElseIf Abs(theta0) <= Atn(B2 / D1) Then
            pointE = Abs(Bx)
            pointF2s = Abs(Ax)
            pointF3s = Abs(Ay) + y2(Cx, Cy, Bx, By, Ax)
        ElseIf Abs(theta0) <= WorksheetFunction.Pi() / 2 Then
            pointE = Abs(Bx)
            pointF2s = Abs(Cx)
            pointF3s = Abs(Ay) + y2(Dx, Dy, Cx, Cy, Ax)
        End If
    Else 'D1 > B2
        If Abs(theta0) <= Atn(B2 / D1) Then
            pointE = Abs(Bx)
            pointF2s = Abs(Ax)
            pointF3s = Abs(Ay) + y2(Cx, Cy, Bx, By, Ax)
        ElseIf Abs(theta0) <= Atn(B2 / D1) Then
            pointE = Abs(Bx)
            pointF2s = Abs(Cx)
            pointF3s = Abs(Ay) + y2(Dx, Dy, Cx, Cy, Ax)
        ElseIf Abs(theta0) <= WorksheetFunction.Pi() / 2 Then
            pointE = Abs(Bx)
            pointF2s = Abs(Cx)
            pointF3s = Abs(Ay) + y2(Dx, Dy, Cx, Cy, Ax)
        End If
    End If
    pointF3s = pointF3s / 2
    temp(1) = pointE
    temp(2) = pointF2s
    temp(3) = pointF3s
    transformedsection = temp
End Function

Function rebarcoord(D1 As Double, B2 As Double, Ccover As Double, _
 Nr2s As Integer, Nr3s As Integer, Nr2sin As Integer, Nr3sin As Integer, _
 sizer As Integer, sizestirp As Integer)
' unit : cm
'Purpose :
'   consider section center as original point
'   calculating each rebar coord.
'Synopsis :
'
'Variable
'   Description:
'       D1(Double) - dimensional length of 2. side of section
'       B2(Double) - dimensional length of 3. side of section
'       Ccover(Double) - clear cover
'       Nr2s(integer) - number of rebar of 2. side of outer layer greater than 2
'       Nr2s(integer) - number of rebar of 3. side of outer layer greater than 2
'       Nr2sin(integer) - number of rebar of 2. side of inner layer greater than 2
'       Nr3sin(integer) - number of rebar of 3. side of inner layer greater than 2
'       Nr - total number of rebar
'       sizer(integer) - size of rebar
'       sizestirp(integer) - size of stirrup
'       Nrtotal(integer) - number of total rebar
'       coord(Nrx2 arrray) - rabar coord.
'       rS2s(cm,Double) - rebar spaing on 2. side
'       rS3s(cm,Double) - rebar spaing on 3. side
'-------------------------------------------------------------------
'Basic setting
'-------------------------------------------------------------------
'Declaration :
Dim rS2s As Double, rS3s As Double
Dim Nr As Integer, Nrouter As Integer
If Nr2sin + Nr3sin = 0 Then
    Nr = (Nr2s + Nr3s) * 2 - 4
Else
    Nr = (Nr2s + Nr3s + Nr2sin + Nr3sin) * 2 - 8
End If
ReDim coord(1 To Nr, 1 To 2) As Double
Dim i As Integer
'-------------------------------------------------------------------
'Allocate :
'-------------------------------------------------------------------
' outer layer
rS2s = (D1 - 2 * Ccover - 2 * diarebar(sizestirp) - diarebar(sizer)) / (Nr2s - 1)
rS3s = (B2 - 2 * Ccover - 2 * diarebar(sizestirp) - diarebar(sizer)) / (Nr3s - 1)
'number of total rebar at outer leyer
Nrouter = (Nr2s + Nr3s) * 2 - 4
' outer layer rebar coord.
' along 2.side
For i = 1 To Nr2s Step 1
        coord(i, 1) = -(D1 - 2 * Ccover - 2 * diarebar(sizestirp) - diarebar(sizer)) / 2 + (i - 1) * rS2s
        coord(i, 2) = (B2 - 2 * Ccover - 2 * diarebar(sizestirp) - diarebar(sizer)) / 2
        coord(Nr2s + i, 1) = -(D1 - 2 * Ccover - 2 * diarebar(sizestirp) - diarebar(sizer)) / 2 + (i - 1) * rS2s
        coord(Nr2s + i, 2) = -(B2 - 2 * Ccover - 2 * diarebar(sizestirp) - diarebar(sizer)) / 2
Next i
' along 3.side
For i = (2 * Nr2s + 1) To (2 * Nr2s + Nr3s - 2)
        coord(i, 1) = -(D1 - 2 * Ccover - 2 * diarebar(sizestirp) - diarebar(sizer)) / 2
        coord(i, 2) = -(B2 - 2 * Ccover - 2 * diarebar(sizestirp) - diarebar(sizer)) / 2 + rS3s * (i - 2 * Nr2s)
        coord(i + Nr3s - 2, 1) = (D1 - 2 * Ccover - 2 * diarebar(sizestirp) - diarebar(sizer)) / 2
        coord(i + Nr3s - 2, 2) = -(B2 - 2 * Ccover - 2 * diarebar(sizestirp) - diarebar(sizer)) / 2 + rS3s * (i - 2 * Nr2s)
Next i
'inner layer rebar coord.
rS2s = (D1 - 2 * Ccover - 2 * diarebar(sizestirp) - 3 * diarebar(sizer)) / (Nr2sin - 1)
rS3s = (B2 - 2 * Ccover - 2 * diarebar(sizestirp) - 3 * diarebar(sizer)) / (Nr3sin - 1)
If Nr2sin > 0 Then
' corner bar
        coord(Nrouter + 1, 1) = -(D1 - 2 * Ccover - 2 * diarebar(sizestirp) - 2 * diarebar(sizer) - 2 / Sqr(2) * diarebar(sizer)) / 2
        coord(Nrouter + 1, 2) = (B2 - 2 * Ccover - 2 * diarebar(sizestirp) - 2 * diarebar(sizer) - 2 / Sqr(2) * diarebar(sizer)) / 2
        coord(Nrouter + 2, 1) = (D1 - 2 * Ccover - 2 * diarebar(sizestirp) - 2 * diarebar(sizer) - 2 / Sqr(2) * diarebar(sizer)) / 2
        coord(Nrouter + 2, 2) = -(B2 - 2 * Ccover - 2 * diarebar(sizestirp) - 2 * diarebar(sizer) - 2 / Sqr(2) * diarebar(sizer)) / 2
        coord(Nrouter + 3, 1) = (D1 - 2 * Ccover - 2 * diarebar(sizestirp) - 2 * diarebar(sizer) - 2 / Sqr(2) * diarebar(sizer)) / 2
        coord(Nrouter + 3, 2) = (B2 - 2 * Ccover - 2 * diarebar(sizestirp) - 2 * diarebar(sizer) - 2 / Sqr(2) * diarebar(sizer)) / 2
        coord(Nrouter + 4, 1) = -(D1 - 2 * Ccover - 2 * diarebar(sizestirp) - 2 * diarebar(sizer) - 2 / Sqr(2) * diarebar(sizer)) / 2
        coord(Nrouter + 4, 2) = -(B2 - 2 * Ccover - 2 * diarebar(sizestirp) - 2 * diarebar(sizer) - 2 / Sqr(2) * diarebar(sizer)) / 2
End If
' side bar along 2.side
If Nr2sin > 2 Then
        For i = Nrouter + 5 To Nrouter + Nr2sin + 2
                coord(i, 1) = -(D1 - 2 * Ccover - 2 * diarebar(sizestirp) - 3 * diarebar(sizer)) / 2 + (i - (Nrouter + 4)) * rS2s
                coord(i, 2) = (B2 - 2 * Ccover - 2 * diarebar(sizestirp) - 3 * diarebar(sizer)) / 2
                coord(Nr2sin - 2 + i, 1) = -(D1 - 2 * Ccover - 2 * diarebar(sizestirp) - 3 * diarebar(sizer)) / 2 + (i - (Nrouter + 4)) * rS2s
                coord(Nr2sin - 2 + i, 2) = -(B2 - 2 * Ccover - 2 * diarebar(sizestirp) - 3 * diarebar(sizer)) / 2
        Next i
End If
' side bar along 3.side
If Nr3sin > 2 Then
        For i = Nrouter + Nr2sin * 2 + 1 To Nrouter + Nr2sin * 2 + Nr3sin - 2
                coord(i, 1) = -(D1 - 2 * Ccover - 2 * diarebar(sizestirp) - 3 * diarebar(sizer)) / 2
                coord(i, 2) = -(B2 - 2 * Ccover - 2 * diarebar(sizestirp) - 3 * diarebar(sizer)) / 2 + (i - (Nrouter + Nr2sin * 2)) * rS3s
                coord(Nr3sin - 2 + i, 1) = (D1 - 2 * Ccover - 2 * diarebar(sizestirp) - 3 * diarebar(sizer)) / 2
                coord(Nr3sin - 2 + i, 2) = -(B2 - 2 * Ccover - 2 * diarebar(sizestirp) - 3 * diarebar(sizer)) / 2 + (i - (Nrouter + Nr2sin * 2)) * rS3s
        Next i
End If
' Conclusion
rebarcoord = coord
End Function

Function Mn(c As Variant, comp As Variant, comps As Variant, tendon As Variant, _
 fy As Double, pointE As Variant, rcoord As Variant)
'Purpose :
'   determine : nominal moment strength of the tranformed section
'Synopsis :
'   take original point moment by Cc,Cs,Ts =
'   Mn(comp, fy, pointE, pointFmaj)
'Variable
'   c(cm,Double) - depth of neutral axis
'   comp(2x1 array) - compressive force and center coord. of whitney block
'   fy(kgf/cm2,integer) - yield strength of rebar
'   comps(200x2 array,[N.T.S;kgf]) - compressive rebar strain and force
'   tendon(200x2) array,[N.T.S;kgf]) - tension rebar strain and force of tranformed section
'   M(Double,kgf-cm) - momnet integral , take compresion moment as positive
'   rcoord(200x2 array) - coord. of raber location in seciton
'Declaration :
Dim M As Variant 'Double
Dim Nr As Integer, i As Integer
'Allocate :
M = 0
Nr = WorksheetFunction.Count(rcoord) / 2
'--------------------------------------------------------
    M = comp(1) * comp(2) 'moment by comp. force Cc
    For i = 1 To Nr
        If c <= pointE Then ' if c is shorter than half of section
            If rcoord(i, 1) > (pointE - c) Then
                M = M + comps(i, 2) * Abs(rcoord(i, 1)) 'moment by comp. rebar force Cs
            Else
                If rcoord(i, 1) >= 0 Then
                    M = M - tendon(i, 2) * Abs(rcoord(i, 1)) 'moment by tension rebar force Ts on positive coord.
                Else
                    M = M + tendon(i, 2) * Abs(rcoord(i, 1)) 'moment by tension rebar force Ts  on negative coord.
                End If
            End If
        Else ' if c is longer than half of section
            If rcoord(i, 1) > (pointE - c) Then
                If rcoord(i, 1) >= 0 Then
                    M = M + comps(i, 2) * Abs(rcoord(i, 1)) 'moment by comp. rebar force Cs on posive coord.
                Else
                    M = M - comps(i, 2) * Abs(rcoord(i, 1)) 'moment by comp. rebar force Cs on negative coord.
                End If
            Else
                    M = M + tendon(i, 2) * Abs(rcoord(i, 1))  'moment by tension rebar force Ts only  on negative coord.
            End If
        End If
    Next
    Mn = M
End Function

Function Cc(fc As Double, c As Variant, pointE As Variant, pointF2 As Double, pointF3 As Double)
'Purpose :
'   determine compressive force and center coord. of whitney block in the tranformed section
'Synopsis :
'   compressive force and center coord. of whitney block of tranformed section
'   (double(kgf),coord.(cm))
'Variable
'   Description:
'   Cc(,kgf,cm,2x1 array) - compressive force and center coord. of whitney block
'   fc(kgf/cm2,integer) - concrete compressive strength
'   pointE(cm,Double) - dimension of transformed section in 2. dir.
'   pointF2(cm,Double) - dimension of transformed section in 2. dir.
'   pointF3(cm,Double) - dimension of transformed section in 3. dir.
'   c(cm,Double) - depth of neutral axis
'   beta(Double) - whitney stress block factor
'   area(cm2,Double) - compressive area
'   a(cm,Double) - whitney stress block depth = beta*c
'   temp - temporary place to locate conclusion
'----------------------------------------------------
'Declaration :
    Dim beta As Double, area As Double, a As Double, h As Double
    Dim temp(1 To 2) As Double
'----------------------------------------------------
'Allocate :
c = c + 0.0000000001
'----------------------------------------------------
'Define beta
    If fc <= 280 Then
        beta = 0.85
    Else
        beta = 0.85 - 0.05 * (fc - 280) / 70
    End If

' Compression zone depth
    a = beta * c

' transform calculation
    If a <= (pointE - pointF2) Then
        area = a * (pointF3 / (pointE - pointF2) * a)
        temp(2) = pointE - 2 / 3 * a
        temp(1) = 0.85 * fc * area
    ElseIf a < pointE And a > (pointE - pointF2) Then
        area = (pointE - pointF2) * pointF3 + 2 * pointF3 * (a - (pointE - pointF2))
        temp(2) = ((pointE - pointF2) * pointF3 * ((pointE - pointF2) / 3 + pointF2) + _
        (2 * pointF3 * (a - (pointE - pointF2)) * (pointF2 - (a - (pointE - pointF2)) / 2))) / _
        area
        temp(1) = 0.85 * fc * area
    ElseIf a < (pointE + pointF2) And a > pointE Then
        area = (pointE - pointF2) * pointF3 + 2 * pointF3 * (a - (pointE - pointF2))
        temp(2) = (((pointE - pointF2) * pointF3 * (pointF2 + (pointE - pointF2) / 3)) + _
        (pointF2 * pointF3 * pointF2) - _
        (2 * (a - pointE) * pointF3 * (a - pointE) / 2)) / area
        temp(1) = 0.85 * fc * area
    ElseIf a <= (pointE * 2) And a > (pointE + pointF2) Then
        h = (2 * pointE - a) * 2 * pointF3 / (pointE - pointF2)
        area = pointF3 * (pointE - pointF2) + 4 * pointF2 * pointF3 + (h + 2 * pointF3) * (a - pointE - pointF2) / 2
        temp(2) = ((pointE - pointF2) * pointF3 * (pointF2 + (pointE - pointF2) / 3) - _
        (h + 2 * pointF3) / 2 * (a - pointE - pointF2) * (pointF2 + (a - pointE - pointF2) / 3 * (2 * h + 2 * pointF3) / (h + 2 * pointF3))) / area
        temp(1) = 0.85 * fc * area
    Else
        area = (pointE - pointF2) * pointF3 * 2 + pointF2 * 4 * pointF3
        temp(2) = 0
        temp(1) = 0.85 * fc * area
    End If
    Cc = temp
End Function

Function Cs(fc As Double, fy As Double, c As Variant, pointE As Variant, rcoord As Variant, sizer As Integer)
'Purpose :
'   determine compressive rebar strain and force in the tranformed section
'Synopsis :
'   compressive rebar strain and force of tranformed section =
'   Cs(fc,fy, c, pointE, rcoord, sizer)
'Variable
'   Description:
'   Cs(Nrx2 array,[unitless;kgf]) - compressive rebar strain and force
'   Nr - numbers of rebar
'   epsilon - strain of rebar
'   epsilony - yield strain
'   rcoord(Nrx2 array) - coord. of raber location in seciton
'   pointE(cm,Double) - dimension of transformed section in maj. dir.
'   c(cm,Double) - depth of neutral axis
'   sizer(integer) - size of rebar
'   rebarcs(kgf,Double) - compressive rebar force
'   temp - temporary place to locate conclusion
'----------------------------------------------------
'Declaration :
    Dim epsilon As Variant, rebarcs As Variant, temp As Variant
    Dim epsilony As Double
    Dim beta As Double, a As Double
    Dim Nr As Integer, i As Integer
'----------------------------------------------------
'Allocate
    Nr = WorksheetFunction.Count(rcoord) / 2
    ReDim epsilon(1 To Nr), rebarcs(1 To Nr), temp(1 To Nr, 1 To 2)
    epsilony = fy / 2040000
'----------------------------------------------------
'Calculate strain of rebar in comprssive
' strain of tension rebar = 0
    If c <= pointE Then
        For i = 1 To Nr
           If rcoord(i, 1) >= (pointE - c) Then
                epsilon(i) = 0.003 / c * (c - (pointE - rcoord(i, 1)))
            Else
                epsilon(i) = 0
            End If
        Next
    Else
        For i = 1 To Nr
            If rcoord(i, 1) >= (-c + pointE) Then
                epsilon(i) = 0.003 / c * (c - (pointE - rcoord(i, 1)))
            Else
                epsilon(i) = 0
            End If
        Next
    End If
'compretion zone depth
    If fc <= 280 Then
        beta = 0.85
    Else
        beta = 0.85 - 0.05 * (fc - 280) / 70
    End If
    a = beta * c
    
'Calculate force of rebar in compressive
    For i = 1 To Nr
        If rcoord(i, 1) >= (-a + pointE) Then
            rebarcs(i) = (fs(Abs(epsilon(i)), fy) - 0.85 * fc) * arearebar(sizer)
        ElseIf rcoord(i, 1) >= (-c + pointE) Then
            rebarcs(i) = fs(Abs(epsilon(i)), fy) * arearebar(sizer)
        End If
    Next
'Packing
    For i = 1 To Nr
        temp(i, 1) = epsilon(i)
        temp(i, 2) = rebarcs(i)
    Next
    Cs = temp
End Function

Function Ts(fy As Double, c As Variant, pointE As Variant, rcoord As Variant, sizer As Integer)
'Purpose :
'   determine tension rebar strain and force in the tranformed section
'Synopsis :
'   tension rebar strain and force of tranformed section =
'   Ts(fy, c, pointE, rcoord, sizer)
'Variable
'   Description:
'   fy(kgf/cm2,integer) - yield strength of rebar
'   epsilon - strain of rebar
'   epsilony - yield  strain of rebar
'   rcoord(Nrx2 array) - coord. of raber location in seciton
'   pointE(cm,Double) - dimension of transformed section in maj. dir.
'   c(cm,Double) - depth of neutral axis
'   sizer(integer) - size of rebar
'   rebarts(kgf,Double) - tension rebar force
'   temp - temporary place to locate conclusion
'----------------------------------------------------
'Declaration :
    Dim epsilon() As Double
    Dim epsilony As Double
    Dim rebarts() As Double, temp() As Double
    Dim Nr As Integer, i As Integer
'----------------------------------------------------
'Allocate
    Nr = WorksheetFunction.Count(rcoord) / 2
    'ReDim temp(1 To Nr, 1 To 2) As Double
    ReDim epsilon(1 To Nr), rebarts(1 To Nr), temp(1 To Nr, 1 To 2)
    epsilony = fy / 2040000
    c = c + 0.0000000001
'----------------------------------------------------
'Calculate strain of rebar in tension
' strain of compressive rebar = 0
    If c <= pointE Then
        For i = 1 To Nr
            If rcoord(i, 1) < (pointE - c) Then
                epsilon(i) = 0.003 / c * (pointE - c - rcoord(i, 1))
            Else
                epsilon(i) = 0
            End If
        Next
    Else
        For i = 1 To Nr
            If rcoord(i, 1) < (-c + pointE) Then
                epsilon(i) = 0.003 / c * (-rcoord(i, 1) - (c - pointE))
            Else
                epsilon(i) = 0
            End If
        Next
    End If
'Calculate stress of rebar in tension
    For i = 1 To Nr
        rebarts(i) = fs(epsilon(i), fy) * arearebar(sizer)
        temp(i, 1) = epsilon(i)
        temp(i, 2) = rebarts(i)
    Next
    Ts = temp
End Function

Function phic(rcoord As Variant, rts As Variant, fy As Double)
'Purpose :
'   determine reduction factor
'Synopsis :
'   phic = phic(recoord, rebar tension array,yield strength)
'   find outest rebar location and strain
'Variable
'   Description:
'       rcoord(Nrx2 array) - rebar coord. (cm)
'       rts(Nrx2 array) - strain and force of tension rebar (unitless,kgf)
'       fy(Double) - yield strenth (kgf/cm2)
'       Tstrain(Double) - strain of outest rebar at tension side
'       Label - i-rebar is the outest rebar
'       epsilony - yield strain
'-----------------------------------------------------
'Declaration :
    Dim Tstrain As Double, epsilony As Double
    Dim i As Integer, Lable As Integer, Nr As Integer
    Dim Minrcoord As Double
'-----------------------------------------------------
'Alloate :
    Tstrain = rts(1, 1)
    epsilony = fy / 2040000
    Minrcoord = rcoord(1, 1)
    Nr = WorksheetFunction.Count(rcoord) / 2
'-----------------------------------------------------
'Searching outest rebar
    For i = 1 To Nr Step 1
        If rcoord(i, 1) < Minrcoord Then
            Minrcoord = rcoord(i, 1)
            Tstrain = rts(i, 1)
        End If
    Next

'define phic
    If Tstrain >= 0.005 Then
        phic = 0.9
    ElseIf Tstrain < 0.005 And Tstrain >= epsilony Then
        phic = 0.65 + 0.25 / (0.005 - epsilony) * (Tstrain - epsilony)
    Else
        phic = 0.65
    End If
End Function

Function fs(epsilon As Double, fy As Double)
'Purpose :
'   determine stress by strain
'Synopsis :
'   rebar stress = fs(epsilon,fy)
'Variable
'   Description:
'   fy(kgf/cm2,integer) - yield strength of rebar
'   epsilon - strain of rebar
'   epsilony - yield  strain of rebar
'----------------------------------------------------
Dim epsilony As Double
epsilony = fy / 2040000
If epsilon > epsilony Then fs = fy Else fs = epsilon * 2040000
End Function

Function coordtransformation(rcoord As Variant, theta As Double)
' coord. transformation
' rcoord. = [ x1, y1 ;
'                 x2, y2 ...]
' theta is clockwise ,unit: rad
    Dim coord As Variant
    coord = WorksheetFunction.Transpose(rcoord)
    Dim transformation(1, 1) As Double
    transformation(0, 0) = Cos(theta)
    transformation(1, 0) = Sin(theta)
    transformation(0, 1) = -Sin(theta)
    transformation(1, 1) = Cos(theta)
    coordtransformation = WorksheetFunction.MMult(transformation, coord)
    coordtransformation = WorksheetFunction.Transpose(coordtransformation)
End Function

Function x2(x0 As Double, y0 As Double, x1 As Double, y1 As Double, y2 As Double)
'Purpose :
'   solve quadratic equation
'Synopsis :
'   a*x0+b=y0
'   a*x1+b=y1
'   a*x2+b=y2
'   solve parameter a,b, then x2
Dim a As Double, b As Double
    a = (y0 - y1) / (x0 - x1 + 0.0000000001)
    b = (x0 * y1 - x1 * y0) / (x0 - x1 + 0.0000000001)
    x2 = (y2 - b) / (a + 0.0000000001)
End Function

Function y2(x0 As Double, y0 As Double, x1 As Double, y1 As Double, x2 As Double)
'Purpose :
'   solve quadratic equation
'Synopsis :
'   a*x0+b=y0
'   a*x1+b=y1
'   a*x2+b=y2
'   solve parameter a,b, then y2
Dim a  As Double, b As Double
    a = (y0 - y1) / (x0 - x1 + 0.0000000001)
    b = y0 - (y0 - y1) / (x0 - x1 + 0.0000000001) * x0
    y2 = a * x2 + b
End Function

Function diarebar(sizerebar As Integer)
' unit : cm
' tranfer rebar size to diameter
    Select Case sizerebar
    Case 3
        diarebar = 0.953
    Case 4
        diarebar = 1.27
    Case 5
        diarebar = 1.59
    Case 6
        diarebar = 1.91
    Case 7
        diarebar = 2.22
    Case 8
        diarebar = 2.54
    Case 9
        diarebar = 2.87
    Case 10
        diarebar = 3.22
    Case 11
        diarebar = 3.58
    End Select
End Function

Function arearebar(sizerebar As Integer)
' unit : cm2
' tranfer rebar size into rebar area
    Select Case sizerebar
    Case 3
        arearebar = 0.71
    Case 4
        arearebar = 1.27
    Case 5
        arearebar = 1.986
    Case 6
        arearebar = 2.839
    Case 7
        arearebar = 3.871
    Case 8
        arearebar = 5.097
    Case 9
        arearebar = 6.452
    Case 10
        arearebar = 8.19
    Case 11
        arearebar = 10.06
    End Select
End Function