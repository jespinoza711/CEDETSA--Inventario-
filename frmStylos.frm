VERSION 5.00
Object = "{C9680CB9-8919-4ED0-A47D-8DC07382CB7B}#1.0#0"; "StyleButtonx.ocx"
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
      _ExtentX        =   17912
      _ExtentY        =   53
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
   Begin StyleButtonX.StyleButton StyleButton1 
      Height          =   615
      Left            =   7950
      TabIndex        =   10
      Top             =   2730
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   1085
      UpColorTop1     =   -2147483628
      UpColorTop2     =   -2147483633
      UpColorTop3     =   -2147483633
      UpColorTop4     =   -2147483633
      UpColorButtom1  =   -2147483627
      UpColorButtom2  =   -2147483633
      UpColorButtom3  =   -2147483633
      UpColorButtom4  =   -2147483633
      UpColorLeft1    =   -2147483628
      UpColorLeft2    =   -2147483633
      UpColorLeft3    =   -2147483633
      UpColorLeft4    =   -2147483633
      UpColorRight1   =   -2147483627
      UpColorRight2   =   -2147483633
      UpColorRight3   =   -2147483633
      UpColorRight4   =   -2147483633
      DownColorTop1   =   -2147483627
      DownColorTop2   =   -2147483633
      DownColorTop3   =   -2147483633
      DownColorTop4   =   -2147483633
      DownColorButtom1=   -2147483628
      DownColorButtom2=   -2147483633
      DownColorButtom3=   -2147483633
      DownColorButtom4=   -2147483633
      DownColorLeft1  =   -2147483627
      DownColorLeft2  =   -2147483633
      DownColorLeft3  =   -2147483633
      DownColorLeft4  =   -2147483633
      DownColorRight1 =   -2147483628
      DownColorRight2 =   -2147483633
      DownColorRight3 =   -2147483633
      DownColorRight4 =   -2147483633
      HoverColorTop1  =   -2147483628
      HoverColorTop2  =   -2147483633
      HoverColorTop3  =   -2147483633
      HoverColorTop4  =   -2147483633
      HoverColorButtom1=   -2147483627
      HoverColorButtom2=   -2147483633
      HoverColorButtom3=   -2147483633
      HoverColorButtom4=   -2147483633
      HoverColorLeft1 =   -2147483628
      HoverColorLeft2 =   -2147483633
      HoverColorLeft3 =   -2147483633
      HoverColorLeft4 =   -2147483633
      HoverColorRight1=   -2147483627
      HoverColorRight2=   -2147483633
      HoverColorRight3=   -2147483633
      HoverColorRight4=   -2147483633
      FocusColorTop1  =   -2147483628
      FocusColorTop2  =   -2147483633
      FocusColorTop3  =   -2147483633
      FocusColorTop4  =   -2147483633
      FocusColorButtom1=   -2147483627
      FocusColorButtom2=   -2147483632
      FocusColorButtom3=   -2147483633
      FocusColorButtom4=   -2147483633
      FocusColorLeft1 =   -2147483628
      FocusColorLeft2 =   -2147483633
      FocusColorLeft3 =   -2147483633
      FocusColorLeft4 =   -2147483633
      FocusColorRight1=   -2147483627
      FocusColorRight2=   -2147483632
      FocusColorRight3=   -2147483633
      FocusColorRight4=   -2147483633
      DisabledColorTop1=   -2147483628
      DisabledColorTop2=   -2147483633
      DisabledColorTop3=   -2147483633
      DisabledColorTop4=   -2147483633
      DisabledColorButtom1=   -2147483627
      DisabledColorButtom2=   -2147483633
      DisabledColorButtom3=   -2147483633
      DisabledColorButtom4=   -2147483633
      DisabledColorLeft1=   -2147483628
      DisabledColorLeft2=   -2147483633
      DisabledColorLeft3=   -2147483633
      DisabledColorLeft4=   -2147483633
      DisabledColorRight1=   -2147483627
      DisabledColorRight2=   -2147483633
      DisabledColorRight3=   -2147483633
      DisabledColorRight4=   -2147483633
      Caption         =   "Asocia Usuarios"
      BackColorUp     =   16777215
      BackColorDown   =   -2147483637
      BackColorHover  =   -2147483629
      BackColorFocus  =   -2147483629
      BackColorDisabled=   -2147483636
      DotsInCornerColor=   16777215
      ForeColorDisabled=   12632256
      PictureUp       =   "frmStylos.frx":36F0
      PictureDown     =   "frmStylos.frx":4344
      PictureHover    =   "frmStylos.frx":4F98
      PictureFocus    =   "frmStylos.frx":5BEC
      PictureDisabled =   "frmStylos.frx":6840
      BeginProperty FontUp {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFocus {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAlignment=   2
      DistanceBetweenPictureAndCaption=   50
   End
   Begin VB.Image Image 
      Height          =   645
      Index           =   3
      Left            =   0
      Picture         =   "frmStylos.frx":7494
      Top             =   3570
      Width           =   12960
   End
   Begin VB.Image Image 
      Height          =   480
      Index           =   1
      Left            =   5580
      Picture         =   "frmStylos.frx":7E88
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
      Picture         =   "frmStylos.frx":8752
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
      Picture         =   "frmStylos.frx":9396
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
      Picture         =   "frmStylos.frx":9C60
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
      Picture         =   "frmStylos.frx":AEF6
      Top             =   60
      Width           =   480
   End
   Begin VB.Image ImageInv2 
      Height          =   405
      Left            =   -150
      Picture         =   "frmStylos.frx":BBC0
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

