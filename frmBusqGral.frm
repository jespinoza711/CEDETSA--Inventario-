VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBusqGral 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "B�squeda"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7125
   Icon            =   "frmBusqGral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H80000009&
      Caption         =   "Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2130
      Picture         =   "frmBusqGral.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2220
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000009&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3660
      Picture         =   "frmBusqGral.frx":0C0E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2220
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   750
      Width           =   6615
      Begin VB.TextBox txtDescr 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002F2F2F&
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   840
         Width           =   4935
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002F2F2F&
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Descr :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002F2F2F&
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblCodigo 
         BackStyle       =   0  'Transparent
         Caption         =   "C�digo :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002F2F2F&
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   2895
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   4762
            MinWidth        =   4762
            Text            =   "<Aceptar> confirma"
            TextSave        =   "<Aceptar> confirma"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image ImageInv 
      Height          =   480
      Left            =   30
      Picture         =   "frmBusqGral.frx":0F52
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblIntrod 
      BackStyle       =   0  'Transparent
      Caption         =   "Digite una aproximaci�n de lo que Ud busca en el campo correspondiente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002F2F2F&
      Height          =   255
      Left            =   690
      TabIndex        =   8
      Top             =   60
      Width           =   6735
   End
   Begin VB.Image ImageInv2 
      Height          =   405
      Left            =   0
      Picture         =   "frmBusqGral.frx":181C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18810
   End
End
Attribute VB_Name = "frmBusqGral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bCodChar As Boolean

Private Sub cmdAceptar_Click()
Dim lbok As Boolean


If txtCodigo.Text = "" And txtDescr.Text = "" Then
  lbok = Mensaje("Por favor, digite un criterio de b�squeda ...", ICO_INFORMACION, False)
  txtCodigo.SetFocus
  Exit Sub
End If


If txtCodigo.Text <> "" And txtDescr.Text <> "" Then
   lbok = Mensaje("Debe seleccionar un solo criterio por favor ", ICO_INFORMACION, False)
   txtCodigo.Text = ""
   txtDescr.Text = ""
  txtCodigo.SetFocus
  Exit Sub
End If
If txtCodigo.Text <> "" Then
    If OnlythisChar(txtCodigo.Text, "*") Then
       lbok = Mensaje("Ese criterio es incorrecto, por favor digite otro... ", ICO_INFORMACION, False)
       txtCodigo.Text = ""
       txtDescr.Text = ""
      txtCodigo.SetFocus
      Exit Sub
    End If
    
    If OnlythisChar(txtCodigo.Text, "%") Then
       lbok = Mensaje("Ese criterio es incorrecto, por favor digite otro... ", ICO_INFORMACION, False)
       txtCodigo.Text = ""
       txtDescr.Text = ""
      txtCodigo.SetFocus
      Exit Sub
    End If
End If

If txtDescr.Text <> "" Then
    If OnlythisChar(txtDescr.Text, "*") Then
       lbok = Mensaje("Ese criterio es incorrecto, por favor digite otro... ", ICO_INFORMACION, False)
       txtCodigo.Text = ""
       txtDescr.Text = ""
      txtCodigo.SetFocus
      Exit Sub
    End If
    
    If OnlythisChar(txtDescr.Text, "%") Then
       lbok = Mensaje("Ese criterio es incorrecto, por favor digite otro... ", ICO_INFORMACION, False)
       txtCodigo.Text = ""
       txtDescr.Text = ""
      txtCodigo.SetFocus
      Exit Sub
    End If
End If
  Hide
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub



Private Sub txtCodigo_GotFocus()
StatusBar1.Panels(1).Text = "C�digo del Elemento Buscado"
End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  If bCodChar = False Then
    If Val_TextboxNum(txtCodigo) Then
      cmdAceptar.SetFocus
    End If
  Else
      cmdAceptar.SetFocus
  End If
  KeyAscii = 0
'Else
'  ValidaLargo txtCodigo.Text, KeyAscii, 5
End If
End Sub

Private Sub txtDescr_GotFocus()
StatusBar1.Panels(1).Text = "Descripci�n aproximada del Elemento Buscado"
End Sub

Private Sub txtDescr_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
  cmdAceptar.SetFocus
End If
Mayuscula KeyAscii
ValidaLargo txtCodigo.Text, KeyAscii, 40
End Sub
