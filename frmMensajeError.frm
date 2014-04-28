VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMensajeError 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Atención"
   ClientHeight    =   1920
   ClientLeft      =   6105
   ClientTop       =   2445
   ClientWidth     =   4785
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnNo 
      Caption         =   "Ca&ncelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2295
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton btnSi 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox lsMensaje 
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2143
      _Version        =   393217
      BackColor       =   12648447
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMensajeError.frx":0000
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
      Left            =   4080
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensajeError.frx":0082
            Key             =   "Informacion"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensajeError.frx":04D4
            Key             =   "Pregunta"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensajeError.frx":0926
            Key             =   "Advertencia"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensajeError.frx":0D78
            Key             =   "Error"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensajeError.frx":11CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image picImagen 
      Height          =   495
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   120
      Width           =   495
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
      Left            =   1200
      TabIndex        =   2
      Top             =   360
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

