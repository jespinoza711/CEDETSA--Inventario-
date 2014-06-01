VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmListadoTraslados 
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13785
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
   ScaleHeight     =   7440
   ScaleWidth      =   13785
   StartUpPosition =   3  'Windows Default
   Begin TrueOleDBGrid60.TDBGrid TDBGrid 
      Height          =   3795
      Left            =   240
      OleObjectBlob   =   "frmListadoTraslados.frx":0000
      TabIndex        =   25
      Top             =   3510
      Width           =   10395
   End
   Begin VB.Frame frmCab 
      Height          =   2475
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   10395
      Begin VB.Frame Frame 
         Caption         =   "Filtrar por: "
         Height          =   585
         Left            =   5250
         TabIndex        =   21
         Top             =   1110
         Width           =   3765
         Begin VB.OptionButton optAmbas 
            Caption         =   "Ambas"
            Height          =   315
            Left            =   2730
            TabIndex        =   24
            Top             =   210
            Width           =   855
         End
         Begin VB.OptionButton optSalidas 
            Caption         =   "Salidas"
            Height          =   315
            Left            =   240
            TabIndex        =   23
            Top             =   210
            Width           =   1035
         End
         Begin VB.OptionButton optEntradas 
            Caption         =   "Entradas"
            Height          =   315
            Left            =   1500
            TabIndex        =   22
            Top             =   210
            Width           =   1035
         End
      End
      Begin VB.CommandButton cmdRefrescar 
         Caption         =   "Refrescar"
         Height          =   645
         Left            =   9210
         TabIndex        =   20
         Top             =   300
         Width           =   885
      End
      Begin VB.CheckBox chkViewPendienteAplicar 
         Caption         =   "Ver solo pendiente de Aplicar"
         Height          =   375
         Left            =   3840
         TabIndex        =   19
         Top             =   750
         Width           =   2505
      End
      Begin VB.TextBox Text3 
         Height          =   345
         Left            =   6960
         TabIndex        =   18
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   345
         Left            =   1620
         TabIndex        =   16
         Top             =   1740
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   345
         Left            =   1620
         TabIndex        =   14
         Top             =   1260
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker dtpFechaFinal 
         Height          =   345
         Left            =   7560
         TabIndex        =   12
         Top             =   780
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         _Version        =   393216
         Format          =   20774913
         CurrentDate     =   41787
      End
      Begin MSComCtl2.DTPicker dtpFechaInicial 
         Height          =   315
         Left            =   1620
         TabIndex        =   11
         Top             =   780
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20774913
         CurrentDate     =   41787
      End
      Begin VB.CommandButton cmdCajero 
         Height          =   320
         Left            =   2850
         Picture         =   "frmListadoTraslados.frx":2840
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   300
         Width           =   300
      End
      Begin VB.TextBox txtDescrBodega 
         Height          =   345
         Left            =   3270
         TabIndex        =   7
         Top             =   300
         Width           =   5715
      End
      Begin VB.TextBox txtIDBodega 
         Height          =   345
         Left            =   1620
         TabIndex        =   6
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "Num Entrada:"
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   6
         Left            =   5910
         TabIndex        =   17
         Top             =   1860
         Width           =   1005
      End
      Begin VB.Label Label 
         Caption         =   "Num Salida:"
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label Label 
         Caption         =   "Documento Inv:"
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   1350
         Width           =   1155
      End
      Begin VB.Label Label 
         Caption         =   "Fecha Final:"
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   3
         Left            =   6540
         TabIndex        =   10
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label 
         Caption         =   "Fecha Inicial:"
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label 
         Caption         =   "Bodega:"
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   13785
      TabIndex        =   0
      Top             =   0
      Width           =   13785
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
         TabIndex        =   2
         Top             =   90
         Width           =   855
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Listados de Traslados"
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
         TabIndex        =   1
         Top             =   420
         Width           =   1305
      End
      Begin VB.Image Image 
         Height          =   645
         Index           =   2
         Left            =   60
         Picture         =   "frmListadoTraslados.frx":2B82
         Stretch         =   -1  'True
         Top             =   60
         Width           =   720
      End
   End
   Begin Inventario.CtlLiner CtlLiner 
      Height          =   30
      Left            =   0
      TabIndex        =   3
      Top             =   750
      Width           =   17925
      _extentx        =   31618
      _extenty        =   53
   End
End
Attribute VB_Name = "frmListadoTraslados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

