VERSION 5.00
Begin VB.Form frmStylos 
   Caption         =   "Form1"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin Inventario.CtlLiner CtlLiner 
      Height          =   30
      Left            =   0
      TabIndex        =   9
      Top             =   4380
      Width           =   10155
      _extentx        =   17912
      _extenty        =   53
   End
   Begin VB.PictureBox picHeader 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   10200
      TabIndex        =   7
      Top             =   4440
      Width           =   10200
      Begin VB.Image Image 
         Height          =   720
         Index           =   2
         Left            =   0
         Picture         =   "frmStylos.frx":0000
         Top             =   50
         Width           =   720
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Please fill up all fields provided below. Add/Update product/item record."
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   840
         TabIndex        =   8
         Top             =   120
         Width           =   3675
      End
   End
   Begin VB.Image Image 
      Height          =   645
      Index           =   3
      Left            =   0
      Picture         =   "frmStylos.frx":36F0
      Top             =   3570
      Width           =   12960
   End
   Begin VB.Image Image 
      Height          =   480
      Index           =   1
      Left            =   5580
      Picture         =   "frmStylos.frx":40E4
      Top             =   2700
      Width           =   480
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press ENTER to proceed"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   6060
      TabIndex        =   6
      Top             =   2850
      Width           =   1740
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   210
      Picture         =   "frmStylos.frx":49AE
      Top             =   2610
      Width           =   480
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cash change from Cashier transaction"
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
      Left            =   840
      TabIndex        =   5
      Top             =   2910
      Width           =   2340
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CHANGE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   0
      Left            =   900
      TabIndex        =   4
      Top             =   2610
      Width           =   750
   End
   Begin VB.Image imgSearch 
      Height          =   480
      Left            =   4260
      Picture         =   "frmStylos.frx":55F2
      Top             =   810
      Width           =   480
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Maintenance Team And Users Registration Form."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Registration Form"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Image Image 
      Appearance      =   0  'Flat
      Height          =   660
      Index           =   0
      Left            =   8280
      Picture         =   "frmStylos.frx":5EBC
      Top             =   1560
      Width           =   1530
   End
   Begin VB.Shape Shape 
      BackStyle       =   1  'Opaque
      Height          =   1050
      Index           =   1
      Left            =   -150
      Top             =   1350
      Width           =   10695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Test Data"
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
      Height          =   345
      Left            =   780
      TabIndex        =   1
      Top             =   60
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Tabel Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   210
      TabIndex        =   0
      Top             =   930
      Width           =   3015
   End
   Begin VB.Shape Shape 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Index           =   0
      Left            =   120
      Top             =   900
      Width           =   3915
   End
   Begin VB.Image ImageInv 
      Height          =   480
      Left            =   60
      Picture         =   "frmStylos.frx":7152
      Top             =   60
      Width           =   480
   End
   Begin VB.Image ImageInv2 
      Height          =   405
      Left            =   -150
      Picture         =   "frmStylos.frx":7E1C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18810
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      FillColor       =   &H00FFFFFF&
      Height          =   855
      Index           =   2
      Left            =   0
      Top             =   2460
      Width           =   5415
   End
End
Attribute VB_Name = "frmStylos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

