VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Wheels Within Wheels"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8715
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   480
      Top             =   1920
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ang As Double
Const frq = 5
Const frq2 = 14
Const frq3 = 40
Const outeroffset = 200
Const outerfrq = 6

Private Sub Timer1_Timer()
    x = 4000 + (3000 + outeroffset * Sin(outerfrq * ang)) * Cos(ang)
    y = 4000 + (3000 + outeroffset * Sin(outerfrq * ang)) * Sin(ang)
    FillColor = RGB(0, 0, 0)
    ForeColor = RGB(0, 0, 0)
    Circle (x, y), 400
    
    X2 = x + 600 * Cos(ang * frq)
    Y2 = y + 600 * Sin(ang * frq)
    FillColor = RGB(255, 0, 0)
    ForeColor = RGB(255, 0, 0)
    Circle (X2, Y2), 200
    
    X3 = X2 + 300 * Cos(ang * frq2)
    Y3 = Y2 + 300 * Sin(ang * frq2)
    FillColor = RGB(0, 0, 255)
    ForeColor = RGB(0, 0, 255)
    Circle (X3, Y3), 100
    
    X4 = X3 + 150 * Cos(ang * frq3)
    Y4 = Y3 + 150 * Sin(ang * frq3)
    FillColor = RGB(255, 255, 255)
    ForeColor = RGB(255, 255, 255)
    Circle (X4, Y4), 50
    
    ang = ang + 0.01
End Sub
