VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3240
      Top             =   240
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   1200
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2520
      Top             =   1080
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   2640
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1920
      Top             =   240
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const VK_SPACE = &H20
Private Const VK_DOWN = &H28
Private Const VK_doubleu = &H57
Private Const VK_SPACE1 = &H31
Private Const VK_ay = &H41
Private Const VK_dee = &H44
Private Const VK_ex = &H58
Private Const VK_UP = &H26
Private Const VK_LEFT = &H25
Private Const VK_RIGHT = &H27
Dim ei As Integer

Dim enema(9)
Dim enemi(9) As Double
Dim fine1 As Integer

Dim enemj(9)
Dim enemt(9)

Dim up1 As Integer
Dim waterno As Integer

Dim tang As Integer

Dim rds As Integer
Dim fm1 As Integer
Dim fm2 As Integer
Dim fm3 As Integer
Dim fm4 As Integer

Dim gm1 As Integer
Dim gm2 As Integer
Dim gm3 As Integer
Dim gm4 As Integer

Dim hm1 As Integer
Dim hm2 As Integer
Dim hm3 As Integer
Dim hm4 As Integer

Dim tm1 As Integer
Dim tm2 As Integer
Dim tm3 As Integer
Dim tm4 As Integer



Dim sp2 As Integer

Dim pacf As Integer
Dim pacf1 As Double


Dim space1 As Double
Dim Dx As DirectX8
Dim D3D As Direct3D8
Dim D3DDevice As Direct3DDevice8
Dim bRunning As Boolean
Dim x As Double
Dim y As Double
Dim t3y As Integer
Dim t3x As Integer

Dim x1 As Integer
Dim y1 As Integer
Dim ls1 As Double
Dim ts1 As Double
Dim i As Integer
Dim s1 As Integer
Dim s2 As Integer
Dim ls As Double
Dim ts As Double
Dim f1 As Integer
Dim rs As Integer
Dim js As Double
Dim ang As Integer
Dim pi As Double
Dim ringsp As Integer

Dim space4 As Integer

Dim p(200, 200)

Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

Private Type TLVERTEX
    x As Single
    y As Single
    Z As Single
    rhw As Single
    color As Long
    specular As Long
    tu As Single
    tv As Single
End Type

Dim dd As Integer

Dim Title(0 To 3) As TLVERTEX
Dim Title1(0 To 3) As TLVERTEX
Dim TriSstrip(0 To 3) As TLVERTEX
Dim TriStrip2(0 To 3) As TLVERTEX

Dim TrisStrip3(0 To 699) As TLVERTEX
Dim TriStrip4(0 To 3) As TLVERTEX

Dim TrisStrip5(0 To 699) As TLVERTEX

Dim TrisStrip6(0 To 699) As TLVERTEX

Dim TrisStrip7(0 To 699) As TLVERTEX
Dim TrisStrip8(0 To 699) As TLVERTEX
Dim TrisStrip9(0 To 699) As TLVERTEX
Dim TrisStrip10(0 To 699) As TLVERTEX
Dim TrisStrip11(0 To 699) As TLVERTEX
Dim TrisStrip12(0 To 699) As TLVERTEX
Dim TrisStrip13(0 To 3) As TLVERTEX

Dim D3DX As D3DX8

Dim start As Direct3DTexture8

Dim Texture As Direct3DTexture8
Dim Texture1 As Direct3DTexture8
Dim TransTexture As Direct3DTexture8
Dim TransTexture1 As Direct3DTexture8
Dim bridge As Direct3DTexture8
Dim bush As Direct3DTexture8
Dim waterfall As Direct3DTexture8
Dim enemy As Direct3DTexture8
Dim fore1 As Direct3DTexture8
Dim finish As Direct3DTexture8
Dim blow As Direct3DTexture8
Dim done As Direct3DTexture8



Public Function Initialise() As Boolean
On Error GoTo ErrHandler:

Dim DispMode As D3DDISPLAYMODE
Dim D3DWindow As D3DPRESENT_PARAMETERS
Dim ColorKeyVal As Long

Set Dx = New DirectX8
Set D3D = Dx.Direct3DCreate()
Set D3DX = New D3DX8




DispMode.Format = D3DFMT_R5G6B5
DispMode.Width = 640
DispMode.Height = 480

D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP
D3DWindow.BackBufferCount = 1
D3DWindow.BackBufferFormat = DispMode.Format
D3DWindow.BackBufferHeight = 480
D3DWindow.BackBufferWidth = 640
D3DWindow.hDeviceWindow = frmmain.hWnd

Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmmain.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                                                        D3DWindow)


D3DDevice.SetVertexShader FVF


D3DDevice.SetRenderState D3DRS_LIGHTING, False

D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True



Set Texture = D3DX.CreateTextureFromFile(D3DDevice, App.Path & "\ExampleTexture.jpg")




ColorKeyVal = &HFF00FF00


Set TransTexture = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\transtexture.jpg", 64, 64, _
                                                                            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
                                                                            D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                                                                            D3DX_FILTER_POINT, ColorKeyVal, _
                                                                            ByVal 0, ByVal 0)

Set TransTexture1 = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\person1.jpg", 400, 400, _
                                                                            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
                                                                            D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                                                                            D3DX_FILTER_POINT, ColorKeyVal, _
                                                                            ByVal 0, ByVal 0)




Set fore1 = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\pellet.jpg", 64, 64, _
                                                                            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
                                                                            D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                                                                            D3DX_FILTER_POINT, ColorKeyVal, _
                                                                            ByVal 0, ByVal 0)
                                                                            
Set blow = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\blow.jpg", 64, 64, _
                                                                            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
                                                                            D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                                                                            D3DX_FILTER_POINT, ColorKeyVal, _
                                                                            ByVal 0, ByVal 0)


Set bridge = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\bridge.jpg", 256, 256, _
                                                                            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
                                                                            D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                                                                            D3DX_FILTER_POINT, ColorKeyVal, _
                                                                            ByVal 0, ByVal 0)

Set bush = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\bush.jpg", 256, 256, _
                                                                            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
                                                                            D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                                                                            D3DX_FILTER_POINT, ColorKeyVal, _
                                                                            ByVal 0, ByVal 0)


Set waterfall = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\waterfall.jpg", 256, 256, _
                                                                            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
                                                                            D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                                                                            D3DX_FILTER_POINT, ColorKeyVal, _
                                                                            ByVal 0, ByVal 0)

Set enemy = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\enemy.jpg", 256, 256, _
                                                                            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
                                                                            D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                                                                            D3DX_FILTER_POINT, ColorKeyVal, _
                                                                            ByVal 0, ByVal 0)
Set finish = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\K.jpg", 256, 256, _
                                                                            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
                                                                            D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                                                                            D3DX_FILTER_POINT, ColorKeyVal, _
                                                                            ByVal 0, ByVal 0)
                                                                            
Set start = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\start.jpg", 256, 256, _
                                                                            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
                                                                            D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                                                                            D3DX_FILTER_POINT, ColorKeyVal, _
                                                                            ByVal 0, ByVal 0)

Set done = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\finish.jpg", 256, 256, _
                                                                            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
                                                                            D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                                                                            D3DX_FILTER_POINT, ColorKeyVal, _
                                                                            ByVal 0, ByVal 0)
If InitialiseGeometry() = True Then
    Initialise = True
    Exit Function
End If


ErrHandler:

Debug.Print "Error Number Returned: " & Err.Number
Initialise = False
End Function

Public Sub Render()

D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0


D3DDevice.BeginScene


    D3DDevice.SetTexture 0, Texture

    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TriStrip2(0), Len(TriStrip2(0))
    
       D3DDevice.SetTexture 0, waterfall
    For i = 0 To 149
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TrisStrip9(i * 4), Len(TrisStrip9(i * 4))
    Next
    D3DDevice.SetTexture 0, finish
    For i = 0 To 149
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TrisStrip12(i * 4), Len(TrisStrip12(i * 4))
    Next
    
    D3DDevice.SetTexture 0, TransTexture
    
  
    
    For i = 0 To 149
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TrisStrip3(i * 4), Len(TrisStrip3(i * 4))
    Next
    
 
    
    
    
    
    D3DDevice.SetTexture 0, bridge
    For i = 0 To 149
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TrisStrip6(i * 4), Len(TrisStrip6(i * 4))
    Next
    
    D3DDevice.SetTexture 0, enemy
    For i = 0 To 149
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TrisStrip11(i * 4), Len(TrisStrip11(i * 4))
    Next
    
    D3DDevice.SetTexture 0, bush
    For i = 0 To 149
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TrisStrip7(i * 4), Len(TrisStrip7(i * 4))
    Next
    
    
    D3DDevice.SetTexture 0, TransTexture1
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TriSstrip(0), Len(TriSstrip(0))
   
    D3DDevice.SetTexture 0, blow
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TrisStrip13(0), Len(TrisStrip13(0))
   
   
   
    D3DDevice.SetTexture 0, fore1
    For i = 0 To 149
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TrisStrip5(i * 4), Len(TrisStrip5(i * 4))
    Next
 
    D3DDevice.SetTexture 0, bush
    For i = 0 To 149
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TrisStrip8(i * 4), Len(TrisStrip8(i * 4))
    Next
    

    
    D3DDevice.SetTexture 0, start
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Title(0), Len(Title(0))
    
    D3DDevice.SetTexture 0, done
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Title1(0), Len(Title1(0))
   
    
    
    
D3DDevice.EndScene


        D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub

Private Sub Form_Click()

If Timer2.Enabled = True Then
Timer1.Enabled = True
Timer2.Enabled = False
End If

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    bRunning = False
End If
End Sub

Private Sub Form_Load()

x1 = 20
y1 = 170
y = 170
For i = 15 To 17
For j = 50 To 179

p(i, j) = 6
Next
Next
For i = 1 To 200
p(i, 180) = 1
p(40, i) = 2


p(i, 10) = 1
p(10, i) = 2
Next

For i = 1 To 200
For j = 1 To 10

p(j, i) = 2
Next
Next

For i = 1 To 200
For j = 181 To 200

p(i, j) = 2
Next
Next

For i = 1 To 200
For j = 40 To 200

p(j, i) = 2
Next
Next


For i = 20 To 22
For j = 175 To 175

p(i, j) = 1
Next
Next


For i = 24 To 27
For j = 173 To 173

p(i, j) = 1
Next
Next


For i = 21 To 24
For j = 169 To 169

p(i, j) = 1
Next
Next


For i = 25 To 27
For j = 169 To 169

p(i, j) = 3
Next
Next

For i = 28 To 32
For j = 169 To 169

p(i, j) = 1
Next
Next

For i = 12 To 24
For j = 169 To 169

p(i, j) = 1
Next
Next

For i = 11 To 11
For j = 165 To 165

p(i, j) = 1
Next
Next

For i = 12 To 20
For j = 161 To 161

p(i, j) = 1
Next
Next

For i = 14 To 18
For j = 161 To 161

p(i, j) = 3
Next
Next

For i = 22 To 25
For j = 158 To 158

p(i, j) = 1
Next
Next

For i = 28 To 33
For j = 162 To 162

p(i, j) = 1
Next
Next

For i = 34 To 38
For j = 157 To 157

p(i, j) = 1
Next
Next

For i = 11 To 33
For j = 151 To 151

p(i, j) = 1
Next
Next

For i = 34 To 34
For j = 153 To 153

p(i, j) = 1
Next
Next

For i = 26 To 29
For j = 151 To 151

p(i, j) = 3
Next
Next

For i = 26 To 29
For j = 146 To 146

p(i, j) = 1
Next
Next

For i = 12 To 24
For j = 143 To 143

p(i, j) = 1
Next
Next

For i = 14 To 18
For j = 143 To 143

p(i, j) = 3
Next
Next

For i = 20 To 33
For j = 137 To 137

p(i, j) = 1
Next
Next

For i = 24 To 30
For j = 137 To 137

p(i, j) = 3
Next
Next

For i = 35 To 38
For j = 134 To 134

p(i, j) = 1
Next
Next

For i = 11 To 15
For j = 134 To 134

p(i, j) = 1
Next
Next

For i = 14 To 19
For j = 128 To 128

p(i, j) = 1
Next
Next

For i = 15 To 17
For j = 50 To 179

p(i, j) = 6
Next
Next

p(18, 127) = 8

enema(0) = 1
enemi(0) = 23
enemj(0) = 168

enema(1) = 1
enemi(1) = 32
enemj(1) = 161

enema(2) = 1
enemi(2) = 23
enemj(2) = 136

enema(3) = 1
enemi(3) = 24
enemj(3) = 150

enema(4) = 1
enemi(4) = 13
enemj(4) = 142

enema(5) = 1
enemi(5) = 13
enemj(5) = 133

p(30, 166) = 4
p(28, 166) = 5

p(27, 177) = 4
p(18, 177) = 5

p(30, 159) = 4
p(19, 158) = 5

p(31, 148) = 4
p(23, 148) = 5

p(22, 134) = 4
p(20, 140) = 5

pi = 3.14159

Me.Show

bRunning = Initialise()
Debug.Print "Device Creation Return Code : ", bRunning

Do While bRunning
    Render
    DoEvents
    
Loop



On Error Resume Next
Set D3DDevice = Nothing
Set D3D = Nothing
Set Dx = Nothing
Debug.Print "All Objects Destroyed"

Unload Me
End
End Sub



Private Function InitialiseGeometry() As Boolean
    
    
On Error GoTo BailOut:




   
            TriStrip2(0) = CreateTLVertex(0, 0, 0, 1, RGB(255, 255, 255), 0, 0, 0)
            
        
            TriStrip2(1) = CreateTLVertex(640, 0, 0, 1, RGB(255, 255, 255), 0, 1, 0)
            
     
            TriStrip2(2) = CreateTLVertex(0, 480, 0, 1, RGB(255, 255, 255), 0, 0, 1)
            
     
            TriStrip2(3) = CreateTLVertex(640, 480, 0, 1, RGB(255, 255, 255), 0, 1, 1)

InitialiseGeometry = True
Exit Function
BailOut:
InitialiseGeometry = False
End Function

Private Function CreateTLVertex(x As Single, y As Single, Z As Single, rhw As Single, color As Long, specular As Long, tu As Single, tv As Single) As TLVERTEX
  
CreateTLVertex.x = x
CreateTLVertex.y = y
CreateTLVertex.Z = Z
CreateTLVertex.rhw = rhw
CreateTLVertex.color = color
CreateTLVertex.specular = specular
CreateTLVertex.tu = tu
CreateTLVertex.tv = tv
End Function

Private Sub Timer1_Timer()




For ei = 0 To 9

If enema(ei) = 1 Then
If enemt(ei) = 1 Then
p(enemi(ei), enemj(ei)) = 0
enemi(ei) = enemi(ei) - 0.25
If p(enemi(ei), enemj(ei)) <> 0 Or p(enemi(ei), enemj(ei) + 1) = 0 Then
enemt(ei) = 0
enemi(ei) = enemi(ei) + 1
End If
p(enemi(ei), enemj(ei)) = 7
End If

If enemt(ei) = 0 Then
p(enemi(ei), enemj(ei)) = 0
enemi(ei) = enemi(ei) + 0.25
If p(enemi(ei), enemj(ei)) <> 0 Or p(enemi(ei), enemj(ei) + 1) = 0 Then
enemt(ei) = 1
enemi(ei) = enemi(ei) - 1
End If

p(enemi(ei), enemj(ei)) = 7

End If
End If
Next



If p(x1 + 6, y1 + 5) = 7 Then
If up1 <> 0 Then

For ei = 0 To 9
If enemj(ei) = y1 + 5 Then
enema(ei) = 0

Timer3.Enabled = True

p(enemi(ei), enemj(ei)) = 0
dd = 6
End If
Next

Else
Timer4.Enabled = True
ang = 0
End If
End If



If p(x1 + 5, y1 + 5) = 7 Then
If up1 <> 0 Then

For ei = 0 To 9
If enemj(ei) = y1 + 5 Then
enema(ei) = 0
Timer3.Enabled = True

p(enemi(ei), enemj(ei)) = 0
dd = 6
End If
Next

End If
End If

If p(x1 + 7, y1 + 5) = 7 Then
If up1 <> 0 Then

For ei = 0 To 9
If enemj(ei) = y1 + 5 Then
enema(ei) = 0
Timer3.Enabled = True

p(enemi(ei), enemj(ei)) = 0
dd = 6
End If
Next

End If
End If




If GetAsyncKeyState(VK_LEFT) <> 0 Then
ang = ang - 50

If p(x1 + 5, y1 + 5) = 0 Or p(x1 + 5, y1 + 5) = 4 Or p(x1 + 5, y1 + 5) = 5 Or p(x1 + 5, y1 + 5) = 6 Or p(x1 + 5, y1 + 5) = 7 Or p(x1 + 5, y1 + 5) = 8 Then x1 = x1 - 1
If p(x1 + 6, y1 + 5) = 7 Then
Timer4.Enabled = True
ang = 0
End If
End If

If GetAsyncKeyState(VK_RIGHT) <> 0 Then
ang = ang + 50

If p(x1 + 7, y1 + 5) = 0 Or p(x1 + 7, y1 + 5) = 4 Or p(x1 + 7, y1 + 5) = 5 Or p(x1 + 7, y1 + 5) = 6 Or p(x1 + 7, y1 + 5) = 7 Or p(x1 + 7, y1 + 5) = 8 Then x1 = x1 + 1
If p(x1 + 6, y1 + 5) = 7 Then
Timer4.Enabled = True
ang = 0
End If
End If



                     
If GetAsyncKeyState(VK_UP) <> 0 And up1 < 10 Then

y1 = y1 - 2
up1 = up1 + 1
End If

If dd <> 0 Then
dd = dd - 1
y1 = y1 - 2
End If

If up1 <> 0 Then up1 = up1 + 1








If p(x1 + 6, y1 + 6) = 0 Or p(x1 + 6, y1 + 6) = 4 Or p(x1 + 6, y1 + 6) = 5 Or p(x1 + 6, y1 + 6) = 6 Or p(x1 + 6, y1 + 6) = 7 Or p(x1 + 6, y1 + 6) = 8 Then y1 = y1 + 1 Else up1 = 0




If p(x1 + 6, y1 + 5) = 8 Then Timer5.Enabled = True






i = 0

For s1 = 0 To 12
For s2 = 0 To 9

If p(x1 + s1, y1 + s2) = 1 Then
TrisStrip3(i) = CreateTLVertex((s1 * 50), (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip3(i + 1) = CreateTLVertex((s1 * 50) + 50, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip3(i + 2) = CreateTLVertex((s1 * 50), (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip3(i + 3) = CreateTLVertex((s1 * 50) + 50, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip5(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip5(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip5(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip5(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip6(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip6(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip6(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip6(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip7(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip7(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip7(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip7(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip8(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip8(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip8(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip8(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip9(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip9(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip9(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip9(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip11(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip11(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip11(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip11(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip12(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip12(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip12(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip12(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)


End If

If p(x1 + s1, y1 + s2) = 2 Then
TrisStrip5(i) = CreateTLVertex((s1 * 50), (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip5(i + 1) = CreateTLVertex((s1 * 50) + 50, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip5(i + 2) = CreateTLVertex((s1 * 50), (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip5(i + 3) = CreateTLVertex((s1 * 50) + 50, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip3(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip3(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip3(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip3(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip6(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip6(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip6(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip6(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip7(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip7(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip7(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip7(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip8(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip8(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip8(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip8(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip9(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip9(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip9(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip9(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip11(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip11(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip11(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip11(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip12(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip12(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip12(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip12(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

End If

If p(x1 + s1, y1 + s2) = 3 Then
TrisStrip6(i) = CreateTLVertex((s1 * 50), ((s2 - 1) * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip6(i + 1) = CreateTLVertex((s1 * 50) + 50, ((s2 - 1) * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip6(i + 2) = CreateTLVertex((s1 * 50), ((s2 - 1) * 50) + 100, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip6(i + 3) = CreateTLVertex((s1 * 50) + 50, ((s2 - 1) * 50) + 100, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip3(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip3(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip3(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip3(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip5(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip5(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip5(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip5(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip7(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip7(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip7(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip7(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip8(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip8(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip8(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip8(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip9(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip9(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip9(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip9(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip11(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip11(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip11(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip11(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip12(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip12(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip12(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip12(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)


End If

If p(x1 + s1, y1 + s2) = 4 Then
TrisStrip7(i) = CreateTLVertex((s1 * 50), (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip7(i + 1) = CreateTLVertex((s1 * 50) + 100, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip7(i + 2) = CreateTLVertex((s1 * 50), (s2 * 50) + 160, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip7(i + 3) = CreateTLVertex((s1 * 50) + 100, (s2 * 50) + 160, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip5(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip5(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip5(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip5(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip6(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip6(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip6(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip6(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip3(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip3(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip3(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip3(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip8(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip8(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip8(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip8(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip9(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip9(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip9(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip9(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip11(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip11(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip11(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip11(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip12(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip12(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip12(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip12(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

End If


If p(x1 + s1, y1 + s2) = 5 Then
TrisStrip8(i) = CreateTLVertex((s1 * 50), (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip8(i + 1) = CreateTLVertex((s1 * 50) + 90, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip8(i + 2) = CreateTLVertex((s1 * 50), (s2 * 50) + 180, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip8(i + 3) = CreateTLVertex((s1 * 50) + 90, (s2 * 50) + 180, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip5(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip5(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip5(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip5(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip6(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip6(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip6(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip6(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip3(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip3(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip3(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip3(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip7(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip7(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip7(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip7(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip9(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip9(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip9(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip9(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip11(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip11(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip11(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip11(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip12(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip12(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip12(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip12(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

End If

If p(x1 + s1, y1 + s2) = 6 Then

If s2 = 0 Then
TrisStrip9(i) = CreateTLVertex((s1 * 50), (s2 * 50) - 50 + waterno, 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip9(i + 1) = CreateTLVertex((s1 * 50) + 50, (s2 * 50) - 50 + waterno, 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip9(i + 2) = CreateTLVertex((s1 * 50), (s2 * 50) + 50 + waterno, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip9(i + 3) = CreateTLVertex((s1 * 50) + 50, (s2 * 50) + 50 + waterno, 0, 1, RGB(255, 255, 255), 0, 1, 1)
Else
TrisStrip9(i) = CreateTLVertex((s1 * 50), (s2 * 50) + waterno, 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip9(i + 1) = CreateTLVertex((s1 * 50) + 50, (s2 * 50) + waterno, 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip9(i + 2) = CreateTLVertex((s1 * 50), (s2 * 50) + 50 + waterno, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip9(i + 3) = CreateTLVertex((s1 * 50) + 50, (s2 * 50) + 50 + waterno, 0, 1, RGB(255, 255, 255), 0, 1, 1)
End If

TrisStrip5(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip5(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip5(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip5(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip6(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip6(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip6(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip6(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip3(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip3(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip3(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip3(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip7(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip7(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip7(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip7(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip8(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip8(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip8(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip8(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip11(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip11(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip11(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip11(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip12(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip12(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip12(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip12(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

End If


If p(x1 + s1, y1 + s2) = 7 Then
TrisStrip11(i) = CreateTLVertex((s1 * 50), (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip11(i + 1) = CreateTLVertex((s1 * 50) + 50, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip11(i + 2) = CreateTLVertex((s1 * 50), (s2 * 50) + 75, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip11(i + 3) = CreateTLVertex((s1 * 50) + 50, (s2 * 50) + 75, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip5(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip5(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip5(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip5(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip6(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip6(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip6(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip6(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip7(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip7(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip7(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip7(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip8(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip8(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip8(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip8(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip9(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip9(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip9(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip9(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip3(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip3(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip3(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip3(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip12(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip12(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip12(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip12(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

End If

If p(x1 + s1, y1 + s2) = 8 Then

TrisStrip12(i) = CreateTLVertex((s1 * 50), (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip12(i + 1) = CreateTLVertex((s1 * 50) + 50, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip12(i + 2) = CreateTLVertex((s1 * 50), (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip12(i + 3) = CreateTLVertex((s1 * 50) + 50, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip5(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip5(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip5(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip5(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip6(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip6(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip6(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip6(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip7(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip7(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip7(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip7(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip8(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip8(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip8(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip8(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip9(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip9(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip9(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip9(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip11(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip11(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip11(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip11(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip3(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip3(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip3(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip3(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

End If


If p(x1 + s1, y1 + s2) = 0 Then
TrisStrip3(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip3(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip3(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip3(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip5(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip5(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip5(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip5(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip6(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip6(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip6(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip6(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip7(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip7(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip7(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip7(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip8(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip8(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip8(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip8(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip9(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip9(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip9(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip9(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip11(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip11(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip11(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip11(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

TrisStrip12(i) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip12(i + 1) = CreateTLVertex(1, (s2 * 50), 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip12(i + 2) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip12(i + 3) = CreateTLVertex(1, (s2 * 50) + 50, 0, 1, RGB(255, 255, 255), 0, 1, 1)

End If








i = i + 4
Next
Next




If ang > 360 And Timer4.Enabled = False Then ang = 0


pi = 3.14129

fm1 = (Sin((ang) * pi / 180) * 50)
fm2 = (Sin((ang + 90) * pi / 180) * 50)
fm3 = (Sin((ang + 180) * pi / 180) * 50)
fm4 = (Sin((ang + 270) * pi / 180) * 50)

gm1 = (Cos((ang) * pi / 180) * 50)
gm2 = (Cos((ang + 90) * pi / 180) * 50)
gm3 = (Cos((ang + 180) * pi / 180) * 50)
gm4 = (Cos((ang + 270) * pi / 180) * 50)

TriSstrip(0) = CreateTLVertex(325 + gm1, 287 + fm1, 0, 1, RGB(255, 255, 255), 0, 0, 0)
TriSstrip(1) = CreateTLVertex(325 + gm2, 287 + fm2, 0, 1, RGB(255, 255, 255), 0, 1, 0)
TriSstrip(2) = CreateTLVertex(325 + gm4, 287 + fm4, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TriSstrip(3) = CreateTLVertex(325 + gm3, 287 + fm3, 0, 1, RGB(255, 255, 255), 0, 1, 1)

sp2 = 0






If rds = 1 Then ringsp = ringsp + 5
If rds = 0 Then ringsp = ringsp - 5

If ringsp > 50 Then rds = 0
If ringsp < 0 Then rds = 1

pacf = pacf + 1
If pacf = 10 Then pacf = 0

Title(0) = CreateTLVertex(1, 150, 0, 1, RGB(255, 255, 255), 0, 0, 0)
Title(1) = CreateTLVertex(1, 150, 0, 1, RGB(255, 255, 255), 0, 1, 0)
Title(2) = CreateTLVertex(1, 325, 0, 1, RGB(255, 255, 255), 0, 0, 1)
Title(3) = CreateTLVertex(1, 325, 0, 1, RGB(255, 255, 255), 0, 1, 1)

waterno = waterno + 10
If waterno = 50 Then waterno = 0


End Sub

Private Sub Timer2_Timer()
Title(0) = CreateTLVertex(100, 150, 0, 1, RGB(255, 255, 255), 0, 0, 0)
Title(1) = CreateTLVertex(500, 150, 0, 1, RGB(255, 255, 255), 0, 1, 0)
Title(2) = CreateTLVertex(100, 325, 0, 1, RGB(255, 255, 255), 0, 0, 1)
Title(3) = CreateTLVertex(500, 325, 0, 1, RGB(255, 255, 255), 0, 1, 1)

End Sub

Private Sub Timer3_Timer()
tang = tang + 30
pi = 3.14129

hm1 = (Sin((tang) * pi / 180) * tang * 2)
hm2 = (Sin((tang + 90) * pi / 180) * tang * 2)
hm3 = (Sin((tang + 180) * pi / 180) * tang * 2)
hm4 = (Sin((tang + 270) * pi / 180) * tang * 2)

tm1 = (Cos((tang) * pi / 180) * tang * 2)
tm2 = (Cos((tang + 90) * pi / 180) * tang * 2)
tm3 = (Cos((tang + 180) * pi / 180) * tang * 2)
tm4 = (Cos((tang + 270) * pi / 180) * tang * 2)

TrisStrip13(0) = CreateTLVertex(325 + tm1, 337 + hm1, 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip13(1) = CreateTLVertex(325 + tm2, 337 + hm2, 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip13(2) = CreateTLVertex(325 + tm4, 337 + hm4, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip13(3) = CreateTLVertex(325 + tm3, 337 + hm3, 0, 1, RGB(255, 255, 255), 0, 1, 1)

If tang > 400 Then
tang = 0
Timer3.Enabled = False
TrisStrip13(0) = CreateTLVertex(1, 287 + hm1, 0, 1, RGB(255, 255, 255), 0, 0, 0)
TrisStrip13(1) = CreateTLVertex(1, 287 + hm2, 0, 1, RGB(255, 255, 255), 0, 1, 0)
TrisStrip13(2) = CreateTLVertex(1, 287 + hm4, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TrisStrip13(3) = CreateTLVertex(1, 287 + hm3, 0, 1, RGB(255, 255, 255), 0, 1, 1)
End If
End Sub

Private Sub Timer4_Timer()



ang = ang + 100
pi = 3.14129

fm1 = (Sin((ang) * pi / 180) * 50)
fm2 = (Sin((ang + 90) * pi / 180) * 50)
fm3 = (Sin((ang + 180) * pi / 180) * 50)
fm4 = (Sin((ang + 270) * pi / 180) * 50)

gm1 = (Cos((ang) * pi / 180) * 50)
gm2 = (Cos((ang + 90) * pi / 180) * 50)
gm3 = (Cos((ang + 180) * pi / 180) * 50)
gm4 = (Cos((ang + 270) * pi / 180) * 50)

TriSstrip(0) = CreateTLVertex(325 + gm1, 287 + fm1, 0, 1, RGB(255, 255, 255), 0, 0, 0)
TriSstrip(1) = CreateTLVertex(325 + gm2, 287 + fm2, 0, 1, RGB(255, 255, 255), 0, 1, 0)
TriSstrip(2) = CreateTLVertex(325 + gm4, 287 + fm4, 0, 1, RGB(255, 255, 255), 0, 0, 1)
TriSstrip(3) = CreateTLVertex(325 + gm3, 287 + fm3, 0, 1, RGB(255, 255, 255), 0, 1, 1)
If ang > 800 Then


x1 = 20
y1 = 170
y = 170
Timer4.Enabled = False

End If


End Sub

Private Sub Timer5_Timer()

fine1 = fine1 + 1
Title1(0) = CreateTLVertex(100, 150, 0, 1, RGB(255, 255, 255), 0, 0, 0)
Title1(1) = CreateTLVertex(500, 150, 0, 1, RGB(255, 255, 255), 0, 1, 0)
Title1(2) = CreateTLVertex(100, 325, 0, 1, RGB(255, 255, 255), 0, 0, 1)
Title1(3) = CreateTLVertex(500, 325, 0, 1, RGB(255, 255, 255), 0, 1, 1)
If fine1 > 100 Then
Title1(0) = CreateTLVertex(1, 150, 0, 1, RGB(255, 255, 255), 0, 0, 0)
Title1(1) = CreateTLVertex(1, 150, 0, 1, RGB(255, 255, 255), 0, 1, 0)
Title1(2) = CreateTLVertex(1, 325, 0, 1, RGB(255, 255, 255), 0, 0, 1)
Title1(3) = CreateTLVertex(1, 325, 0, 1, RGB(255, 255, 255), 0, 1, 1)
Timer5.Enabled = False
End If

End Sub
