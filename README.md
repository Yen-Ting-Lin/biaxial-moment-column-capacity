# biaxial-moment-column-capacity
Capacity Calculation for Columns with Biaxial Bending and Axial Force
Function PM(fc,fy,D1,B2,Ccover,Nr2s,Nr3s,Nr2sin,Nr3sin,sizer,sizestirp,Pu,Mu2,Mu3)
'Purpose :
'   determine stress ratio of rectangle section loaded by biaxial moments and axial force
'Synopsis :
'   Stress ratio
'   PM(fc, fy, D1, B2, Ccover, Nr2s, Nr3s, Nr2sin, Nr3sin, sizer, sizestirp, Pu, Mu2, Mu3)
' Copyright :
'   Copyright 2023 Yen-Ting Lin. All rights reserved
'Variable Description:
'   fc(kgf/cm2) - concrete compressive strength
'   fy(kgf/cm2) - yield strength of rebar
'   D1(cm) - dimensional length of (2). side of section
'   B2(cm) - dimensional length of (3). side of section
'   Ccover(cm) - clear cover
'   Nr2s - number of rebar of (2). side of outer layer
'   Nr3s - number of rebar of (3). side of outer layer
'   Nr2sin - number of rebar of (2). side of inner layer
'   Nr3sin - number of rebar of (3). side of inner layer
'   sizer(U.S.Customary) - size of rebar
'   sizestirp(U.S.Customary) - size of stirrup
'   Pu(tf) - design axial force
'   Mu2(tf-m) - design moment of dir.(2)
'   Mu3(tf-m) - design moment of dir.(3)
'   theta(rad) - rotation on M2-M3 diagram, aka. rotation degree of section
'   e(cm) - eccentrial distence
'   seccoord(cm,4x3 array) - coord. of section corner
'   rcoord(cm,Nrx2 array) - coord. of raber location in seciton
'   pointE(cm) - dimension of transformed section in dir.(2)
'   pointF2s(cm) - dimension of transformed section in dir.(2)
'   pointF3s(cm) - dimension of transformed section in dir.(3)
'   c(cm) - depth of neutral axis
