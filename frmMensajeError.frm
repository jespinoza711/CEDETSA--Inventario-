VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMensajeError 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Atención"
   ClientHeight    =   2580
   ClientLeft      =   6105
   ClientTop       =   2445
   ClientWidth     =   6690
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame 
      Height          =   1095
      Left            =   -30
      TabIndex        =   4
      Top             =   1530
      Width           =   6765
      Begin VB.CommandButton btnNo 
         BackColor       =   &H80000009&
         Caption         =   "&Cancelar"
         Height          =   585
         Left            =   3630
         Picture         =   "frmMensajeError.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   330
         Width           =   1305
      End
      Begin VB.CommandButton btnSi 
         BackColor       =   &H80000009&
         Caption         =   "&Aceptar"
         Height          =   585
         Left            =   1920
         Picture         =   "frmMensajeError.frx":0344
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   330
         Width           =   1305
      End
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   6690
      TabIndex        =   2
      Top             =   0
      Width           =   6690
      Begin VB.Image Image1 
         Height          =   255
         Left            =   150
         Picture         =   "frmMensajeError.frx":0688
         Stretch         =   -1  'True
         Top             =   30
         Width           =   300
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   660
         TabIndex        =   3
         Top             =   30
         Width           =   3675
      End
   End
   Begin RichTextLib.RichTextBox lsMensaje 
      Height          =   1125
      Left            =   870
      TabIndex        =   1
      Top             =   330
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   1984
      _Version        =   393217
      BackColor       =   12648447
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMensajeError.frx":12CC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ListaIconos 
      Left            =   6120
      Top             =   2310
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensajeError.frx":134E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensajeError.frx":18EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensajeError.frx":1F5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensajeError.frx":2661
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensajeError.frx":2B6F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image picImagen 
      Height          =   675
      Left            =   60
      Top             =   540
      Width           =   735
   End
   Begin VB.Label lsMensaje_1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1680
      TabIndex        =   0
      Top             =   870
      Visible         =   0   'False
      Width           =   2895
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMensajeError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Variable que indica si se "Aceptó" o se "Canceló"
Public gbAceptar As Boolean

Private Sub btnNo_Click()
  gbAceptar = False 'Indica que se canceló
  btnSi.Parent.Hide
End Sub

Private Sub btnSi_Click()
  gbAceptar = True  'Indica que se aceptó
  btnSi.Parent.Hide
End Sub

'''''Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'''''  If Shift = 0 Then
'''''    ' F8
'''''    If (KeyCode = vbKeyF8) Then
'''''      btnSi_Click
'''''      KeyCode = 0
'''''    ' F9
'''''    ElseIf (KeyCode = vbKeyF9) Then
'''''      btnNo_Click
'''''      KeyCode = 0
'''''    End If
'''''  End If
'''''End Sub

Private Sub Form_Load()

End Sub
