VERSION 5.00
Begin VB.UserControl CtlLiner 
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6570
   ScaleHeight     =   750
   ScaleWidth      =   6570
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   4425
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   540
      X2              =   4890
      Y1              =   210
      Y2              =   210
   End
End
Attribute VB_Name = "CtlLiner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub UserControl_Initialize()
    On Error Resume Next
    UserControl.Height = 30
    UserControl.BackColor = UserControl.Parent.BackColor
End Sub


Private Sub UserControl_Paint()
Line1.X1 = 0
Line1.Y1 = 0
Line1.X2 = UserControl.Width
Line1.Y2 = 0

Line2.X1 = 0
Line2.Y1 = 20
Line2.X2 = UserControl.Width
Line2.Y2 = 20
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 30
    UserControl_Paint
End Sub

