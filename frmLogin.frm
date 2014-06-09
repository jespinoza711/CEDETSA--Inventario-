VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   4335
   ClientLeft      =   6870
   ClientTop       =   4050
   ClientWidth     =   7830
   DrawMode        =   6  'Mask Pen Not
   FillStyle       =   0  'Solid
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":030A
   ScaleHeight     =   2561.261
   ScaleMode       =   0  'User
   ScaleWidth      =   7351.943
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicFooter 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   7800
      TabIndex        =   11
      Top             =   3990
      Width           =   7830
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Solutions 4 you."
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   540
         TabIndex        =   12
         Top             =   30
         Width           =   1245
      End
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   7830
      TabIndex        =   7
      Top             =   0
      Width           =   7830
      Begin VB.Label lbFormCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   930
         TabIndex        =   9
         Top             =   90
         Width           =   855
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese al sistema con su usuario y contraseña asignados."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   165
         Index           =   1
         Left            =   930
         TabIndex        =   8
         Top             =   420
         Width           =   3570
      End
      Begin VB.Image Image 
         Height          =   690
         Index           =   2
         Left            =   60
         Picture         =   "frmLogin.frx":0F67
         Top             =   60
         Width           =   690
      End
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5160
      TabIndex        =   0
      Top             =   1200
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   5160
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1800
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5220
      TabIndex        =   2
      Top             =   2970
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6420
      TabIndex        =   3
      Top             =   2970
      Width           =   1020
   End
   Begin Inventario.CtlLiner CtlLiner 
      Height          =   30
      Left            =   0
      TabIndex        =   10
      Top             =   750
      Width           =   17925
      _ExtentX        =   31618
      _ExtentY        =   53
   End
   Begin VB.Image ImageLogo 
      Appearance      =   0  'Flat
      Height          =   2115
      Left            =   600
      Picture         =   "frmLogin.frx":1BC4
      Top             =   990
      Width           =   2130
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   2250
      Picture         =   "frmLogin.frx":35B8
      Top             =   0
      Width           =   720
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Versión"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   1
      Left            =   3720
      TabIndex        =   5
      Top             =   1860
      Width           =   1155
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIO:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   0
      Left            =   3720
      TabIndex        =   4
      Top             =   1290
      Width           =   975
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
End Sub




Private Sub cmdOK_Click()
Dim lbok As Boolean
    'check for correct password
    txtUserName.Text = UCase(txtUserName.Text)
    If UserCouldIN(txtUserName.Text, txtPassword.Text) Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        LoginSucceeded = True
        gsUSUARIO = txtUserName.Text
        'Marchoso1.Filename = ""
        Me.Hide
        lbok = LoadAccess(txtUserName.Text, txtPassword.Text, C_MODULO)
        If Not lbok Then
            lbok = Mensaje("No se pudieron cargar los accesos del usuario ", ICO_ERROR, False)
            End
        Else
            'lbok = CargaParametros()
            If Not lbok Then
                lbok = Mensaje("No se ha configurado el Sistema... los parametros no se han definido ", ICO_ERROR, False)
                End
            End If
        End If
    Else
        lbok = Mensaje("Login o Password incorrectos...", ICO_ERROR, False)
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub Form_Load()
'Marchoso1.Filename = App.Path & "\SPV4.gif"
'Marchoso1.Filename = App.Path & "\SPVfinal2.gif"
lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
Set grRecordsetAcceso = New ADODB.Recordset
End Sub

