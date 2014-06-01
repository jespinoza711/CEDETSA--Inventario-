VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProductos 
   Caption         =   "Maestro de Productos"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17040
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmProductos.frx":0000
   ScaleHeight     =   9150
   ScaleWidth      =   17040
   WindowState     =   2  'Maximized
   Begin Inventario.CtlLiner CtlLiner 
      Height          =   30
      Left            =   -420
      TabIndex        =   99
      Top             =   750
      Width           =   17925
      _ExtentX        =   31618
      _ExtentY        =   53
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
      ScaleWidth      =   17040
      TabIndex        =   96
      Top             =   0
      Width           =   17040
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
         TabIndex        =   98
         Top             =   90
         Width           =   855
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Actualización del Maestro de Productos"
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
         TabIndex        =   97
         Top             =   420
         Width           =   2400
      End
      Begin VB.Image Image 
         Height          =   720
         Index           =   2
         Left            =   150
         Picture         =   "frmProductos.frx":0CCA
         Top             =   45
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdUndo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3420
      Picture         =   "frmProductos.frx":1614
      Style           =   1  'Graphical
      TabIndex        =   95
      ToolTipText     =   "Deshacer / Cancelar"
      Top             =   4605
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton cmdEliminar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3420
      Picture         =   "frmProductos.frx":22DE
      Style           =   1  'Graphical
      TabIndex        =   94
      ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
      Top             =   3405
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton cmdAdd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3420
      Picture         =   "frmProductos.frx":2FA8
      Style           =   1  'Graphical
      TabIndex        =   93
      ToolTipText     =   "Agrega el item con los datos digitados..."
      Top             =   2205
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3420
      Picture         =   "frmProductos.frx":3C72
      Style           =   1  'Graphical
      TabIndex        =   92
      ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
      Top             =   4005
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton cmdEditItem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3420
      Picture         =   "frmProductos.frx":593C
      Style           =   1  'Graphical
      TabIndex        =   91
      ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
      Top             =   2805
      Visible         =   0   'False
      Width           =   555
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7275
      Left            =   120
      TabIndex        =   64
      Top             =   1230
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   12832
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   4210752
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.TextBox txtCodigo 
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
      Height          =   285
      Left            =   4860
      TabIndex        =   60
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtDescr 
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
      Height          =   285
      Left            =   7680
      TabIndex        =   59
      Top             =   870
      Width           =   6945
   End
   Begin VB.CheckBox chkActivo 
      Caption         =   "Activo ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15120
      TabIndex        =   58
      Top             =   870
      Width           =   1095
   End
   Begin ActiveTabs.SSActiveTabs SSActiveTabs1 
      Height          =   7275
      Left            =   4035
      TabIndex        =   0
      Top             =   1215
      Width           =   12780
      _ExtentX        =   22543
      _ExtentY        =   12832
      _Version        =   131083
      TabCount        =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHotTracking {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tabs            =   "frmProductos.frx":6606
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   6885
         Left            =   30
         TabIndex        =   1
         Top             =   360
         Width           =   12720
         _ExtentX        =   22437
         _ExtentY        =   12144
         _Version        =   131083
         TabGuid         =   "frmProductos.frx":66D7
         Begin VB.CommandButton Command1 
            Caption         =   "Refrescar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10440
            TabIndex        =   10
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton cmdBodMov 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   1440
            Picture         =   "frmProductos.frx":66FF
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   360
            Width           =   300
         End
         Begin VB.TextBox txtDescrBodMov 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H002F2F2F&
            Height          =   285
            Left            =   3720
            TabIndex        =   8
            Top             =   360
            Width           =   3615
         End
         Begin VB.TextBox txtBodMov 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H002F2F2F&
            Height          =   285
            Left            =   2040
            TabIndex        =   7
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdDelBodMov 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   3240
            Picture         =   "frmProductos.frx":6A41
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   360
            Width           =   300
         End
         Begin VB.CommandButton cmdDelTipoMov 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   9960
            Picture         =   "frmProductos.frx":6E83
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   360
            Width           =   315
         End
         Begin VB.CommandButton cmdTipoMov 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   8160
            Picture         =   "frmProductos.frx":72C5
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   360
            Width           =   315
         End
         Begin VB.TextBox txtTipoMov 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H002F2F2F&
            Height          =   285
            Left            =   8640
            TabIndex        =   3
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtDescrTipoMov 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   10320
            TabIndex        =   2
            Top             =   360
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker DTPFecInic 
            Height          =   255
            Left            =   2880
            TabIndex        =   11
            Top             =   855
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   3092271
            Format          =   62128129
            CurrentDate     =   41692
            MinDate         =   41690
         End
         Begin MSComCtl2.DTPicker DTPFechaFin 
            Height          =   255
            Left            =   6240
            TabIndex        =   12
            Top             =   840
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   3092271
            Format          =   62128129
            CurrentDate     =   41698
         End
         Begin TrueOleDBGrid60.TDBGrid TDBGMov 
            Height          =   5145
            Left            =   150
            OleObjectBlob   =   "frmProductos.frx":7607
            TabIndex        =   13
            Top             =   1290
            Width           =   12375
         End
         Begin VB.Label Label3 
            Caption         =   "Bodega:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H002F2F2F&
            Height          =   255
            Left            =   600
            TabIndex        =   17
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Desde :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H002F2F2F&
            Height          =   255
            Left            =   2040
            TabIndex        =   16
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Hasta :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H002F2F2F&
            Height          =   255
            Left            =   5280
            TabIndex        =   15
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H002F2F2F&
            Height          =   255
            Left            =   7440
            TabIndex        =   14
            Top             =   360
            Width           =   495
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   6885
         Left            =   30
         TabIndex        =   18
         Top             =   360
         Width           =   12720
         _ExtentX        =   22437
         _ExtentY        =   12144
         _Version        =   131083
         TabGuid         =   "frmProductos.frx":E0A7
         Begin VB.CommandButton cmdDelBodega 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   3720
            Picture         =   "frmProductos.frx":E0CF
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   240
            Width           =   300
         End
         Begin VB.TextBox txtBodega 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H002F2F2F&
            Height          =   285
            Left            =   2520
            TabIndex        =   23
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtDescrBodega 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H002F2F2F&
            Height          =   285
            Left            =   4200
            TabIndex        =   22
            Top             =   240
            Width           =   4455
         End
         Begin VB.CommandButton cmdBodega 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   1920
            Picture         =   "frmProductos.frx":E511
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   240
            Width           =   300
         End
         Begin VB.CommandButton cmdRefresExistencia 
            Caption         =   "Refrescar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9120
            TabIndex        =   20
            Top             =   240
            Width           =   1575
         End
         Begin TrueOleDBGrid60.TDBGrid TDBGExistencia 
            Height          =   5745
            Left            =   120
            OleObjectBlob   =   "frmProductos.frx":E853
            TabIndex        =   19
            Top             =   660
            Width           =   12465
         End
         Begin VB.Label Label2 
            Caption         =   "Bodega:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H002F2F2F&
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   25
            Top             =   240
            Width           =   1095
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   6885
         Index           =   0
         Left            =   30
         TabIndex        =   26
         Top             =   360
         Width           =   12720
         _ExtentX        =   22437
         _ExtentY        =   12144
         _Version        =   131083
         TabGuid         =   "frmProductos.frx":12A76
         Begin VB.Frame Frame1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   3405
            Left            =   210
            TabIndex        =   65
            Top             =   3450
            Width           =   12345
            Begin VB.Frame Frame4 
               Caption         =   "Bonificación del Proveedor "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   735
               Left            =   6150
               TabIndex        =   112
               Top             =   2580
               Width           =   5415
               Begin VB.TextBox txtCOBonifica 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H002F2F2F&
                  Height          =   285
                  Left            =   3960
                  TabIndex        =   114
                  Top             =   330
                  Width           =   1095
               End
               Begin VB.TextBox txtCOPorCada 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H002F2F2F&
                  Height          =   285
                  Left            =   1440
                  TabIndex        =   113
                  Top             =   300
                  Width           =   1095
               End
               Begin VB.Label Label10 
                  Caption         =   "Bonifica :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H002F2F2F&
                  Height          =   255
                  Left            =   3000
                  TabIndex        =   116
                  Top             =   360
                  Width           =   975
               End
               Begin VB.Label Label11 
                  Caption         =   "Por Cada :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H002F2F2F&
                  Height          =   255
                  Left            =   240
                  TabIndex        =   115
                  Top             =   360
                  Width           =   1095
               End
            End
            Begin VB.Frame Frame3 
               Caption         =   "Facturación "
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
               Height          =   735
               Left            =   120
               TabIndex        =   109
               Top             =   2550
               Width           =   5415
               Begin VB.CommandButton cmdBonifica 
                  Caption         =   "Escala de Bonificación"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   3120
                  TabIndex        =   111
                  Top             =   150
                  Width           =   2175
               End
               Begin VB.CheckBox chkBonificaFA 
                  Caption         =   "Bonifica en Facturación"
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
                  Height          =   250
                  Left            =   360
                  TabIndex        =   110
                  Top             =   330
                  Width           =   2655
               End
            End
            Begin VB.CheckBox chkBajaPrecioProveedor 
               Caption         =   "Se afecta el Precio con un Descuento/Alza del Proveedor?"
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
               Height          =   255
               Left            =   120
               TabIndex        =   80
               Top             =   840
               Width           =   5415
            End
            Begin VB.CheckBox chkBajaPrecioDistribuidor 
               Caption         =   "Se afecta el Precio con un Descuento/Alza del Distribuidor?"
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
               Height          =   255
               Left            =   5640
               TabIndex        =   79
               Top             =   840
               Width           =   5535
            End
            Begin VB.CommandButton cmdDelProveedor 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Left            =   3120
               Picture         =   "frmProductos.frx":12A9E
               Style           =   1  'Graphical
               TabIndex        =   78
               Top             =   360
               Width           =   300
            End
            Begin VB.TextBox txtCodProveedor 
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
               Height          =   285
               Left            =   1920
               TabIndex        =   77
               Top             =   360
               Width           =   1095
            End
            Begin VB.TextBox txtDescrProveedor 
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
               Height          =   285
               Left            =   3600
               TabIndex        =   76
               Top             =   360
               Width           =   8415
            End
            Begin VB.CommandButton cmdProveedor 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Left            =   1320
               Picture         =   "frmProductos.frx":12EE0
               Style           =   1  'Graphical
               TabIndex        =   75
               Top             =   360
               Width           =   300
            End
            Begin VB.TextBox txtPorcDescAlzaProveedor 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   5640
               TabIndex        =   74
               ToolTipText     =   "Si el signo es Positivo  es un alza, si es Negativo es una rebaja"
               Top             =   1200
               Width           =   1095
            End
            Begin VB.TextBox txtCostoUltPromDolar 
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
               Height          =   285
               Left            =   10950
               TabIndex        =   73
               Top             =   2130
               Width           =   1095
            End
            Begin VB.TextBox txtCostoUltPromLocal 
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
               Height          =   285
               Left            =   7890
               TabIndex        =   72
               Top             =   2160
               Width           =   1095
            End
            Begin VB.TextBox txtCostoUltDolar 
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
               Height          =   285
               Left            =   4740
               TabIndex        =   71
               Top             =   2160
               Width           =   1095
            End
            Begin VB.TextBox txtCostoUltLocal 
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
               Height          =   285
               Left            =   1650
               TabIndex        =   70
               Top             =   2160
               Width           =   1095
            End
            Begin VB.TextBox txtPrecioFOBLocal 
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
               Height          =   285
               Left            =   10950
               TabIndex        =   69
               Top             =   1650
               Width           =   1095
            End
            Begin VB.TextBox txtPrecioCIFLocal 
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
               Height          =   285
               Left            =   7890
               TabIndex        =   68
               Top             =   1680
               Width           =   1095
            End
            Begin VB.TextBox txtPrecioFarmaciaLocal 
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
               Height          =   285
               Left            =   4740
               TabIndex        =   67
               Top             =   1680
               Width           =   1095
            End
            Begin VB.TextBox txtPrecioPublico 
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
               Height          =   285
               Left            =   1650
               TabIndex        =   66
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label Label 
               Caption         =   "*"
               Height          =   255
               Index           =   7
               Left            =   6810
               TabIndex        =   108
               Top             =   1200
               Width           =   225
            End
            Begin VB.Label lblProveedor 
               Caption         =   "Proveedor:"
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
               Height          =   255
               Left            =   120
               TabIndex        =   90
               Top             =   390
               Width           =   1095
            End
            Begin VB.Label lblPorcDescAlzaProveedor 
               Caption         =   "Porcentaje Alza/Baja Proveedor en Precios:"
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
               Height          =   255
               Left            =   120
               TabIndex        =   89
               Top             =   1200
               Width           =   4335
            End
            Begin VB.Label lblCostoUltPromDolar 
               Caption         =   "Costo Ult Promedio $:"
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
               Height          =   255
               Left            =   9150
               TabIndex        =   88
               Top             =   2160
               Width           =   1935
            End
            Begin VB.Label lblCostoUltPromLocal 
               Caption         =   "Costo Ult Promedio C$:"
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
               Height          =   255
               Left            =   5970
               TabIndex        =   87
               Top             =   2160
               Width           =   2055
            End
            Begin VB.Label lblCostoUltDolar 
               Caption         =   "Costo Ultimo $ :"
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
               Height          =   255
               Left            =   3030
               TabIndex        =   86
               Top             =   2160
               Width           =   1815
            End
            Begin VB.Label lblCostoUltLocal 
               Caption         =   "Costo Ultimo C$:"
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
               Height          =   255
               Left            =   120
               TabIndex        =   85
               Top             =   2160
               Width           =   1575
            End
            Begin VB.Label lblPrecioFOBLocal 
               Caption         =   "Precio/Costo FOB C$:"
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
               Height          =   255
               Left            =   9150
               TabIndex        =   84
               Top             =   1680
               Width           =   1935
            End
            Begin VB.Label lblCostoCIF 
               Caption         =   "Precio/Costo CIF C$:"
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
               Height          =   255
               Left            =   6090
               TabIndex        =   83
               Top             =   1680
               Width           =   1935
            End
            Begin VB.Label lblPrecioFarmacia 
               Caption         =   "Precio Farmacia C$ :"
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
               Height          =   255
               Left            =   3030
               TabIndex        =   82
               Top             =   1680
               Width           =   1815
            End
            Begin VB.Label lblPrecioPublico 
               Caption         =   "Precio Público C$:"
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
               Height          =   255
               Left            =   120
               TabIndex        =   81
               Top             =   1680
               Width           =   1815
            End
         End
         Begin VB.Frame Frame2 
            Height          =   3195
            Left            =   210
            TabIndex        =   27
            Top             =   150
            Width           =   12315
            Begin VB.CheckBox chkEsMuestra 
               Caption         =   "Muestra Médica ?"
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
               Height          =   255
               Left            =   9210
               TabIndex        =   51
               Top             =   600
               Width           =   1935
            End
            Begin VB.CheckBox chkEtico 
               Caption         =   "Etico ?"
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
               Height          =   255
               Left            =   9210
               TabIndex        =   50
               Top             =   1080
               Width           =   1095
            End
            Begin VB.CheckBox chkControlado 
               Caption         =   "Controlado ?"
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
               Height          =   255
               Left            =   9210
               TabIndex        =   49
               Top             =   1560
               Width           =   1575
            End
            Begin VB.CommandButton cmdDelclasif3 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Left            =   3240
               Picture         =   "frmProductos.frx":13222
               Style           =   1  'Graphical
               TabIndex        =   48
               Top             =   1290
               Width           =   300
            End
            Begin VB.TextBox txtClasif3 
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   2160
               TabIndex        =   47
               Top             =   1290
               Width           =   855
            End
            Begin VB.TextBox txtDecrClasif3 
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   3720
               TabIndex        =   46
               Top             =   1290
               Width           =   5085
            End
            Begin VB.CommandButton cmdClasif3 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Left            =   1560
               Picture         =   "frmProductos.frx":13664
               Style           =   1  'Graphical
               TabIndex        =   45
               Top             =   1290
               Width           =   300
            End
            Begin VB.CommandButton cmdDelclasif2 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Left            =   3240
               Picture         =   "frmProductos.frx":139A6
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   825
               Width           =   300
            End
            Begin VB.TextBox txtClasif2 
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   2160
               TabIndex        =   43
               Top             =   810
               Width           =   855
            End
            Begin VB.TextBox txtDecrClasif2 
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   3720
               TabIndex        =   42
               Top             =   810
               Width           =   5085
            End
            Begin VB.CommandButton cmdClasif2 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Left            =   1560
               Picture         =   "frmProductos.frx":13DE8
               Style           =   1  'Graphical
               TabIndex        =   41
               Top             =   810
               Width           =   300
            End
            Begin VB.TextBox txtCodigoBarra 
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   2160
               TabIndex        =   40
               Top             =   2640
               Width           =   3975
            End
            Begin VB.CommandButton cmdDelclasif1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Left            =   3240
               Picture         =   "frmProductos.frx":1412A
               Style           =   1  'Graphical
               TabIndex        =   39
               Top             =   360
               Width           =   300
            End
            Begin VB.TextBox txtClasif1 
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   2160
               TabIndex        =   38
               Top             =   360
               Width           =   855
            End
            Begin VB.TextBox txtDecrClasif1 
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   3720
               TabIndex        =   37
               Top             =   360
               Width           =   5085
            End
            Begin VB.CommandButton cmdClasif1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Left            =   1560
               Picture         =   "frmProductos.frx":1456C
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   360
               Width           =   300
            End
            Begin VB.CommandButton cmdDelPresentacion 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Left            =   3240
               Picture         =   "frmProductos.frx":148AE
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   2190
               Width           =   300
            End
            Begin VB.TextBox txtIDPresentacion 
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   2160
               TabIndex        =   34
               Top             =   2190
               Width           =   855
            End
            Begin VB.TextBox txtDescrPresentacion 
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   3720
               TabIndex        =   33
               Top             =   2190
               Width           =   5085
            End
            Begin VB.CommandButton cmdPresentacion 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Left            =   1560
               Picture         =   "frmProductos.frx":14CF0
               Style           =   1  'Graphical
               TabIndex        =   32
               Top             =   2190
               Width           =   300
            End
            Begin VB.CommandButton cmdImpuesto 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Left            =   1560
               Picture         =   "frmProductos.frx":15032
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   1740
               Width           =   300
            End
            Begin VB.TextBox txtDescrImpuesto 
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   3720
               TabIndex        =   30
               Top             =   1740
               Width           =   5085
            End
            Begin VB.TextBox txtImpuesto 
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   2160
               TabIndex        =   29
               Top             =   1740
               Width           =   855
            End
            Begin VB.CommandButton cmdDelImpuesto 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Left            =   3240
               Picture         =   "frmProductos.frx":15374
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   1740
               Width           =   300
            End
            Begin VB.Label Label 
               Caption         =   "*"
               Height          =   255
               Index           =   6
               Left            =   8880
               TabIndex        =   107
               Top             =   2190
               Width           =   225
            End
            Begin VB.Label Label 
               Caption         =   "*"
               Height          =   255
               Index           =   5
               Left            =   8880
               TabIndex        =   106
               Top             =   1740
               Width           =   225
            End
            Begin VB.Label Label 
               Caption         =   "*"
               Height          =   255
               Index           =   4
               Left            =   8880
               TabIndex        =   105
               Top             =   1290
               Width           =   225
            End
            Begin VB.Label Label 
               Caption         =   "*"
               Height          =   255
               Index           =   3
               Left            =   8880
               TabIndex        =   104
               Top             =   810
               Width           =   225
            End
            Begin VB.Label Label 
               Caption         =   "*"
               Height          =   255
               Index           =   0
               Left            =   8880
               TabIndex        =   102
               Top             =   360
               Width           =   225
            End
            Begin VB.Label lblClasif3 
               Caption         =   "Clasificación3:"
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
               Height          =   255
               Left            =   240
               TabIndex        =   57
               Top             =   1320
               Width           =   1335
            End
            Begin VB.Label lblClasif2 
               Caption         =   "Clasificación2:"
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
               Height          =   255
               Left            =   240
               TabIndex        =   56
               Top             =   840
               Width           =   1335
            End
            Begin VB.Label lblCodigoBarra 
               Caption         =   "Código Barra:"
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
               Height          =   255
               Left            =   240
               TabIndex        =   55
               Top             =   2640
               Width           =   1335
            End
            Begin VB.Label lblClasif1 
               Caption         =   "Clasificación1:"
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
               Height          =   255
               Left            =   240
               TabIndex        =   54
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label lblPresentación 
               Caption         =   "Presentación:"
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
               Height          =   255
               Left            =   240
               TabIndex        =   53
               Top             =   2250
               Width           =   1335
            End
            Begin VB.Label lblImpuesto 
               Caption         =   "Impuesto:"
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
               Height          =   255
               Left            =   240
               TabIndex        =   52
               Top             =   1770
               Width           =   1215
            End
         End
      End
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   1695
      Left            =   5400
      OleObjectBlob   =   "frmProductos.frx":157B6
      TabIndex        =   61
      Top             =   9270
      Visible         =   0   'False
      Width           =   12585
   End
   Begin VB.Label Label 
      Caption         =   "*"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   103
      Top             =   0
      Width           =   225
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Los elemento con el signo de un (*) son requeridos."
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   4500
      TabIndex        =   101
      Top             =   8520
      Width           =   3705
   End
   Begin VB.Image Image 
      Height          =   480
      Index           =   0
      Left            =   4020
      Picture         =   "frmProductos.frx":1B593
      Top             =   8400
      Width           =   480
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press ENTER to proceed"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   480
      TabIndex        =   100
      Top             =   150
      Width           =   1740
   End
   Begin VB.Image Image 
      Height          =   480
      Index           =   1
      Left            =   0
      Picture         =   "frmProductos.frx":1BE5D
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código :"
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
      Height          =   255
      Left            =   4140
      TabIndex        =   63
      Top             =   870
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción :"
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
      Height          =   255
      Left            =   6570
      TabIndex        =   62
      Top             =   900
      Width           =   1095
   End
End
Attribute VB_Name = "frmProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As ADODB.Recordset
Dim rst2 As ADODB.Recordset
Dim rst3 As ADODB.Recordset
Dim Accion As TypAccion
Public gsFormCaption As String
Public gsTitle As String

Private Sub HabilitarBotones()
    Select Case Accion
        Case TypAccion.Add, TypAccion.Edit
            cmdSave.Enabled = True
            cmdUndo.Enabled = True
            cmdEliminar.Enabled = False
            cmdAdd.Enabled = False
            cmdEditItem.Enabled = False
        Case TypAccion.View
            If rst.State = adStateClosed Then
                cmdEditItem.Enabled = False
                cmdSave.Enabled = False
                cmdUndo.Enabled = False
                cmdEliminar.Enabled = False
                cmdAdd.Enabled = True
                HabilitarBotonesMain
                Exit Sub
            End If
            If rst.RecordCount <> 0 Then
                cmdSave.Enabled = False
                cmdUndo.Enabled = False
                cmdEliminar.Enabled = True
                cmdAdd.Enabled = True
                cmdEditItem.Enabled = True
            Else
                cmdAdd.Enabled = True
                cmdEditItem.Enabled = False
                cmdSave.Enabled = False
                cmdUndo.Enabled = False
                cmdEliminar.Enabled = False
                cmdAdd.Enabled = True
            End If
    End Select
    HabilitarBotonesMain
End Sub

Public Sub HabilitarControles()
    Select Case Accion
        Case TypAccion.Add
            txtCodigo.Text = "1000"
            txtCodigo.Enabled = False
            txtDescr.Enabled = True
            chkActivo.Enabled = True
            chkActivo.value = 1
            
            txtDescr.Text = ""
            fmtTextbox txtCodigo, "R"
            fmtTextbox txtDescr, "O"
            
            txtClasif1.Enabled = True
            txtClasif1.Text = ""
            txtClasif2.Enabled = True
            txtClasif2.Text = ""
            txtClasif3.Enabled = True
            txtClasif3.Text = ""
            txtCodigoBarra.Enabled = True
            txtCodigoBarra.Text = "ND"
            txtCodProveedor.Enabled = True
            txtDescrProveedor.Text = ""
            txtCodProveedor.Text = ""
            txtCostoUltDolar.Enabled = True
            txtCostoUltDolar.Text = "0"
            txtCostoUltLocal.Enabled = True
            txtCostoUltLocal.Text = "0"
            txtCostoUltPromDolar.Enabled = True
            txtCostoUltPromDolar.Text = "0"
            txtCostoUltPromLocal.Enabled = True
            txtCostoUltPromLocal.Text = "0"
            txtDecrClasif1.Text = ""
            txtDecrClasif1.Enabled = False
            txtDecrClasif2.Text = ""
            txtDecrClasif2.Enabled = False
            txtDecrClasif3.Text = ""
            txtDecrClasif3.Enabled = False
            txtDescrImpuesto.Text = ""
            txtDescrImpuesto.Enabled = False
            txtDescrPresentacion.Text = ""
            txtDescrPresentacion.Enabled = False
            txtIDPresentacion.Text = ""
            txtIDPresentacion.Enabled = True
            txtImpuesto.Text = ""
            txtImpuesto.Enabled = True
            txtPorcDescAlzaProveedor.Enabled = True
            txtPorcDescAlzaProveedor.Text = "0"
            txtPrecioCIFLocal.Enabled = True
            txtPrecioCIFLocal.Text = "0"
            txtPrecioFarmaciaLocal.Enabled = True
            txtPrecioFarmaciaLocal.Text = "0"
            txtPrecioFOBLocal.Enabled = True
            txtPrecioFOBLocal.Text = "0"
            txtPrecioPublico.Enabled = True
            txtPrecioPublico.Text = "0"
            
            chkBonificaFA.Enabled = False
            cmdBonifica.Enabled = False
            
            txtCOBonifica.Enabled = True
            txtCOBonifica.Text = "0"
            txtCOPorCada.Enabled = True
            txtCOPorCada.Text = "0"
            
            Me.cmdClasif1.Enabled = True
            Me.cmdClasif2.Enabled = True
            Me.cmdClasif3.Enabled = True
            Me.cmdImpuesto.Enabled = True
            Me.cmdPresentacion.Enabled = True
            Me.cmdProveedor.Enabled = True
            Me.SSActiveTabs1.Tabs(2).Enabled = False
            Me.SSActiveTabs1.Tabs(3).Enabled = False
            Me.TDBG.Enabled = False
        Case TypAccion.Edit
            txtCodigo.Enabled = True
            txtDescr.Enabled = True
            fmtTextbox txtCodigo, "R"
            fmtTextbox txtDescr, "O"
            chkActivo.Enabled = True
            txtClasif1.Enabled = True
            txtClasif2.Enabled = True
            txtClasif3.Enabled = True
            txtCodigoBarra.Enabled = True
            txtCodProveedor.Enabled = True
            txtCostoUltDolar.Enabled = True
            txtCostoUltDolar.Text = "0"
            txtCostoUltLocal.Enabled = True
            txtCostoUltLocal.Text = "0"
            txtCostoUltPromDolar.Enabled = True
            txtCostoUltPromDolar.Text = "0"
            txtCostoUltPromLocal.Enabled = True
            txtCostoUltPromLocal.Text = "0"
            txtDecrClasif1.Enabled = False
            txtDecrClasif2.Enabled = False
            txtDecrClasif3.Enabled = False
            txtDescrImpuesto.Enabled = False
            txtDescrPresentacion.Enabled = False
            txtIDPresentacion.Enabled = True
            txtImpuesto.Enabled = True
            txtPorcDescAlzaProveedor.Enabled = True
            txtPorcDescAlzaProveedor.Text = "0"
            txtPrecioCIFLocal.Enabled = True
            txtPrecioCIFLocal.Text = "0"
            txtPrecioFarmaciaLocal.Enabled = True
            txtPrecioFarmaciaLocal.Text = "0"
            txtPrecioFOBLocal.Enabled = True
            txtPrecioFOBLocal.Text = "0"
            txtPrecioPublico.Enabled = True
            txtPrecioPublico.Text = "0"
            Me.cmdClasif1.Enabled = True
            Me.cmdClasif2.Enabled = True
            Me.cmdClasif3.Enabled = True
            Me.cmdImpuesto.Enabled = True
            Me.cmdPresentacion.Enabled = True
            Me.cmdProveedor.Enabled = True
            Me.SSActiveTabs1.Tabs(2).Enabled = False
            Me.SSActiveTabs1.Tabs(3).Enabled = False
            
            chkBonificaFA.Enabled = True
            If chkBonificaFA.value = 1 Then
                cmdBonifica.Enabled = True
            Else
                cmdBonifica.Enabled = False
            End If
            
            Me.TDBG.Enabled = False
        Case TypAccion.View
            txtCodigo.Enabled = False
            txtDescr.Enabled = False
            chkActivo.Enabled = False
            txtClasif1.Enabled = False
            txtClasif2.Enabled = False
            txtClasif3.Enabled = False
            txtCodigoBarra.Enabled = False
            txtCodProveedor.Enabled = False
            txtDescrProveedor.Enabled = False
            txtCostoUltDolar.Enabled = False
            txtCostoUltLocal.Enabled = False
            txtCostoUltPromDolar.Enabled = False
            txtCostoUltPromLocal.Enabled = False
            txtDecrClasif1.Enabled = False
            txtDecrClasif2.Enabled = False
            txtDecrClasif3.Enabled = False
            txtDescrImpuesto.Enabled = False
            txtDescrPresentacion.Enabled = False
            txtIDPresentacion.Enabled = False
            txtImpuesto.Enabled = False
            txtPorcDescAlzaProveedor.Enabled = False
            txtPrecioCIFLocal.Enabled = False
            txtPrecioFarmaciaLocal.Enabled = False
            txtPrecioFOBLocal.Enabled = False
            txtPrecioPublico.Enabled = False
            Me.cmdClasif1.Enabled = False
            Me.cmdClasif2.Enabled = False
            Me.cmdClasif3.Enabled = False
            Me.cmdImpuesto.Enabled = False
            Me.cmdPresentacion.Enabled = False
            Me.cmdProveedor.Enabled = False
            Me.SSActiveTabs1.Tabs(2).Enabled = True
            Me.SSActiveTabs1.Tabs(3).Enabled = True
            
            chkBonificaFA.Enabled = False
            If chkBonificaFA.value = 1 Then
                cmdBonifica.Enabled = True
            Else
                cmdBonifica.Enabled = False
            End If
            cmdBonifica.Enabled = False
            
            Me.TDBG.Enabled = True
    End Select
End Sub

Public Sub CommandPass(ByVal srcPerformWhat As String)
    On Error GoTo err
    Select Case srcPerformWhat
        Case "Nuevo"
            cmdAdd_Click
        Case "Editar"
            cmdEditItem_Click
        Case "Eliminar"
            cmdEliminar_Click
        Case "Cancelar"
            cmdUndo_Click
        Case "Imprimir"
            MsgBox "Imprimir"
        Case "Cerrar"
            Unload Me
        Case "Guardar"
            cmdSave_Click
    End Select
    Exit Sub
    'Trap the error
err:
    If err.Number = -2147467259 Then
        MsgBox "You cannot delete this record because it was used by other records! If you want to delete this record" & vbCrLf & _
               "you will first have to delete or change the records that currenly used this record as shown bellow." & vbCrLf & vbCrLf & _
               err.Description, , "Delete Operation Failed!"
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub HabilitarBotonesMain()
    Select Case Accion
        Case TypAccion.Add, TypAccion.Edit
            MDIMain.tbMenu.Buttons(12).Enabled = True 'Guardar
            MDIMain.tbMenu.Buttons(13).Enabled = True 'Undo
            MDIMain.tbMenu.Buttons(11).Enabled = False 'Eliminar
            MDIMain.tbMenu.Buttons(8).Enabled = False 'Nuevo
            MDIMain.tbMenu.Buttons(10).Enabled = False 'Editar
        Case TypAccion.View
            MDIMain.tbMenu.Buttons(12).Enabled = False 'Guardar
            MDIMain.tbMenu.Buttons(13).Enabled = False 'Undo
            MDIMain.tbMenu.Buttons(11).Enabled = True 'Eliminar
            MDIMain.tbMenu.Buttons(8).Enabled = True 'Nuevo
            MDIMain.tbMenu.Buttons(10).Enabled = True 'Editar
    End Select
End Sub

Private Sub chkBonificaFA_Click()
    If (Me.chkBonificaFA.value = 1) Then Me.cmdBonifica.Enabled = True
End Sub

Private Sub cmdBonifica_Click()
    Dim frm As New frmEscalaBonificacion
    frm.gsFormCaption = "Escalas de Bonificación"
    frm.gsTitle = "Escalas de Bonificación"
    frm.gsIDProducto = txtCodigo.Text
    frm.gsDescr = txtDescr.Text
    frm.Show vbModal
    Set frm = Nothing
End Sub


Private Sub cmdAdd_Click()
    Accion = Add
    HabilitarBotones
    HabilitarControles
    txtDescr.SetFocus
End Sub

Private Sub cmdBodega_Click()
   Dim frm As New frmBrowseCat
    
    frm.gsCaptionfrm = "Bodega"
    frm.gsTablabrw = "invBODEGA"
    frm.gsCodigobrw = "IDBodega"
    frm.gbTypeCodeStr = True
    frm.gsDescrbrw = "Descr"
    frm.gbFiltra = False
    'frm.gsFiltro = "IdPaquete='" & Me.gsIDTipoTransaccion & "'"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      Me.txtBodega.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      Me.txtBodega.Text = frm.gsDescrbrw
      fmtTextbox txtBodega, "R"
    End If
End Sub

Private Sub cmdClasif1_Click()
   Dim frm As frmBrowseCat
   Dim lbl As Label
    Set lbl = Me.lblClasif1
    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = Mid(lbl.Caption, 1, Len(lbl.Caption) - 1) '& lblund.Caption
    frm.gsTablabrw = "vGlobalCatalogo"
    frm.gsCodigobrw = "IDCATALOGO"
    frm.gbTypeCodeStr = True
    frm.gsDescrbrw = "DESCR"
    frm.gbFiltra = True
    frm.gsFiltro = "CATALOGO='" & lbl.Tag & "'"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      txtClasif1.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtDecrClasif1.Text = frm.gsDescrbrw
      fmtTextbox txtDecrClasif1, "R"
    End If

End Sub

Private Sub cmdClasif2_Click()
   Dim frm As frmBrowseCat
   Dim lbl As Label
    Set lbl = Me.lblClasif2
    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = Mid(lbl.Caption, 1, Len(lbl.Caption) - 1) '& lblund.Caption
    frm.gsTablabrw = "vGlobalCatalogo"
    frm.gsCodigobrw = "IDCATALOGO"
    frm.gbTypeCodeStr = True
    frm.gsDescrbrw = "DESCR"
    frm.gbFiltra = True
    frm.gsFiltro = "CATALOGO='" & lbl.Tag & "'"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      txtClasif2.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtDecrClasif2.Text = frm.gsDescrbrw
      fmtTextbox txtDecrClasif2, "R"
    End If
End Sub

Private Sub cmdClasif3_Click()
   Dim frm As frmBrowseCat
   Dim lbl As Label
    Set lbl = Me.lblClasif3
    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = Mid(lbl.Caption, 1, Len(lbl.Caption) - 1) '& lblund.Caption
    frm.gsTablabrw = "vGlobalCatalogo"
    frm.gsCodigobrw = "IDCATALOGO"
    frm.gbTypeCodeStr = True
    frm.gsDescrbrw = "DESCR"
    frm.gbFiltra = True
    frm.gsFiltro = "CATALOGO='" & lbl.Tag & "'"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      txtClasif3.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtDecrClasif3.Text = frm.gsDescrbrw
      fmtTextbox txtDecrClasif3, "R"
    End If
End Sub


Private Sub cmdEditItem_Click()
    Accion = Edit
    GetDataFromGridToControl
    HabilitarBotones
    HabilitarControles
End Sub
Private Sub GetDataFromGridToControl()
    If Not (rst.EOF And rst.BOF) Then
        txtCodigo.Text = rst("IDProducto").value
        txtDescr.Text = rst("Descr").value
        If rst("Activo").value = "SI" Then
            chkActivo.value = 1
        Else
            chkActivo.value = 0
        End If
        txtClasif1.Text = rst("Linea").value
        txtDecrClasif1.Text = rst("DESCRCLASIF1").value
        txtClasif2.Text = rst("Familia").value
        txtDecrClasif2.Text = rst("DESCRCLASIF2").value
        txtClasif3.Text = rst("SubFamilia").value
        txtDecrClasif3.Text = rst("DESCRCLASIF3").value
        txtCodProveedor.Text = rst("IDPROVEEDOR").value
        txtDescrProveedor.Text = rst("NOMBRE").value
        txtIDPresentacion.Text = rst("IDPRESENTACION").value
        txtDescrPresentacion.Text = rst("DESCRPRESENTACION").value
        txtImpuesto.Text = rst("IMPUESTO").value
        txtDescrImpuesto.Text = rst("DESCRIMPUESTO").value
        txtCodigoBarra.Text = rst("CODIGOBARRA").value
        If rst("ESMUESTRA").value = "SI" Then
            chkEsMuestra.value = 1
        Else
            chkEsMuestra.value = 0
        End If
        If rst("ESCONTROLADO").value = "SI" Then
            chkControlado.value = 1
        Else
            chkControlado.value = 0
        End If
        If rst("ESETICO").value = "SI" Then
            chkEtico.value = 1
        Else
            chkEtico.value = 0
        End If
        
        If rst("BonificaFA").value = True Then
            chkBonificaFA.value = 1
        Else
            chkBonificaFA.value = 0
        End If
        
        If rst("BAJAPRECIODISTRIBUIDOR").value = "SI" Then
            chkBajaPrecioDistribuidor.value = 1
        Else
            chkBajaPrecioDistribuidor.value = 0
        End If
        
        If rst("BAJAPRECIOPROVEEDOR").value = "SI" Then
            chkBajaPrecioProveedor.value = 1
        Else
            chkBajaPrecioProveedor.value = 0
        End If
        txtCostoUltDolar.Text = rst("CostoUltDolar").value
        txtCostoUltLocal = rst("CostoUltLocal").value
        txtCostoUltPromDolar = rst("CostoUltPromDolar").value
        txtCostoUltPromLocal = rst("CostoUltPromLocal").value
        txtPorcDescAlzaProveedor.Text = rst("PorcDescAlzaProveedor").value
        txtPrecioCIFLocal.Text = rst("PrecioCIFLocal").value
        txtPrecioFarmaciaLocal = rst("PrecioFarmaciaLocal").value
        txtPrecioFOBLocal.Text = rst("PrecioFOBLocal").value
        txtPrecioPublico.Text = rst("PrecioPublicoLocal").value
        
        txtCOBonifica.Text = rst("BonifCOCantidad").value
        txtCOPorCada.Text = rst("BonifCOPorCada").value

    Else
        txtCodigo.Text = ""
        txtDescr.Text = ""
    End If

End Sub

Private Sub cmdEliminar_Click()
    Dim lbok As Boolean
    Dim sMsg As String
    Dim sActivo As String
    Dim sEsMuestra As String
    Dim sEsControlado As String
    Dim sEsEtico As String
    Dim sBajaPrecioDistribuidor As String
    Dim sBajaPrecioProveedor As String
    Dim sBonificaFA As String
    
        If txtCodigo.Text = "" Then
            lbok = Mensaje("El Código del Producto no puede estar en Blanco", ICO_ERROR, False)
            Exit Sub
        End If
        
        If chkActivo.value = 1 Then
            sActivo = "1"
        Else
            sActivo = "0"
        End If
     
         
        If chkEsMuestra.value = 1 Then
            sEsMuestra = "1"
        Else
            sEsMuestra = "0"
        End If
        
        If chkEsMuestra.value = 1 Then
            sEsEtico = "1"
        Else
            sEsEtico = "0"
        End If
        
        If chkBajaPrecioDistribuidor.value = 1 Then
            sBajaPrecioDistribuidor = "1"
        Else
            sBajaPrecioDistribuidor = "0"
        End If
        
         If chkBajaPrecioProveedor.value = 1 Then
            sBajaPrecioProveedor = "1"
        Else
            sBajaPrecioProveedor = "0"
        End If
        
        If chkBonificaFA.value = 1 Then
            sBonificaFA = "1"
        Else
            sBonificaFA = "0"
        End If
       
        ' hay que validar la integridad referencial
        ' if exists dependencias then No se puede eliminar
        lbok = Mensaje("Está seguro de eliminar el Producto " & rst("Descr").value, ICO_PREGUNTA, True)
        If lbok Then
                lbok = invUpdateProducto("D", txtCodigo.Text, txtDescr.Text, txtImpuesto.Text, sEsMuestra, sEsControlado, txtClasif1.Text, _
                    txtClasif2.Text, txtClasif3.Text, sEsEtico, sBajaPrecioDistribuidor, txtCodProveedor.Text, txtCostoUltLocal.Text, txtCostoUltDolar.Text, _
                    txtCostoUltPromLocal.Text, txtCostoUltPromDolar.Text, txtPrecioPublico.Text, Me.txtPrecioFarmaciaLocal.Text, Me.txtPrecioCIFLocal.Text, _
                    txtPrecioFOBLocal.Text, txtIDPresentacion.Text, sBajaPrecioProveedor, Me.txtPorcDescAlzaProveedor.Text, gsUSUARIO, gsUSUARIO, sActivo, txtCodigoBarra.Text, _
                    sBonificaFA, txtCOPorCada.Text, txtCOBonifica.Text)
    
            If lbok Then
                sMsg = "Borrado Exitosamente ... "
                lbok = Mensaje(sMsg, ICO_OK, False)
                ' actualiza datos
                cargaGrid
            End If
        End If
End Sub

Private Sub cmdImpuesto_Click()
   Dim frm As frmBrowseCat
   Dim lbl As Label
    Set lbl = lblImpuesto
    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = Mid(lbl.Caption, 1, Len(lbl.Caption) - 1) '& lblund.Caption
    frm.gsTablabrw = "vGlobalCatalogo"
    frm.gsCodigobrw = "IDCATALOGO"
    frm.gbTypeCodeStr = True
    frm.gsDescrbrw = "DESCR"
    frm.gbFiltra = True
    frm.gsFiltro = "CATALOGO='" & lbl.Tag & "'"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      txtImpuesto.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtDescrImpuesto.Text = frm.gsDescrbrw
      fmtTextbox txtDescrImpuesto, "R"
    End If
End Sub

Private Sub cmdPresentacion_Click()
   Dim frm As frmBrowseCat
   Dim lbl As Label
    Set lbl = lblPresentación
    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = Mid(lbl.Caption, 1, Len(lbl.Caption) - 1) '& lblund.Caption
    frm.gsTablabrw = "vGlobalCatalogo"
    frm.gsCodigobrw = "IDCATALOGO"
    frm.gbTypeCodeStr = True
    frm.gsDescrbrw = "DESCR"
    frm.gbFiltra = True
    frm.gsFiltro = "CATALOGO='" & lbl.Tag & "'"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      txtIDPresentacion.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtDescrPresentacion.Text = frm.gsDescrbrw
      fmtTextbox txtDescrPresentacion, "R"
    End If
End Sub

Private Sub cmdProveedor_Click()
   Dim frm As frmBrowseCat

    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = "Proveedor" '& lblund.Caption
    frm.gsTablabrw = "cpProveedor"
    frm.gsCodigobrw = "IDProveedor"
    frm.gbTypeCodeStr = False
    frm.gsDescrbrw = "Nombre"
    frm.gbFiltra = False
    
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      Me.txtCodProveedor.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtDescrProveedor.Text = frm.gsDescrbrw
      fmtTextbox txtDescrProveedor, "R"
    End If
End Sub

Private Sub cmdRefresExistencia_Click()
    Dim lbok As Boolean
    Dim sIDArticulo As String
    Dim sIDBodega As String
    If txtBodMov.Text = "" Then
        sIDBodega = "-1"
    End If
    
    sIDArticulo = txtCodigo.Text
    
    lbok = CargaExistenciaBodega(sIDArticulo, sIDBodega)

End Sub

Private Sub cmdSave_Click()
    Dim lbok As Boolean
    Dim sMsg As String
    Dim sActivo As String
    Dim sFiltro As String
    
    Dim sEsMuestra As String
    Dim sEsControlado As String
    Dim sEsEtico As String
    Dim sBajaPrecioDistribuidor As String
    Dim sBajaPrecioProveedor As String
    Dim sBonificaFA As String
    
        If Not ControlsOk() Then
            Exit Sub
        End If
        
        If chkActivo.value = 1 Then
            sActivo = "1"
        Else
            sActivo = "0"
        End If
        
         
        If chkEsMuestra.value = 1 Then
            sEsMuestra = "1"
        Else
            sEsMuestra = "0"
        End If
        
        If chkEsMuestra.value = 1 Then
            sEsEtico = "1"
        Else
            sEsEtico = "0"
        End If
        
        If chkBajaPrecioDistribuidor.value = 1 Then
            sBajaPrecioDistribuidor = "1"
        Else
            sBajaPrecioDistribuidor = "0"
        End If
        
         If chkBajaPrecioProveedor.value = 1 Then
            sBajaPrecioProveedor = "1"
        Else
            sBajaPrecioProveedor = "0"
        End If
        
        
         If chkControlado.value = 1 Then
            sEsControlado = "1"
        Else
            sEsControlado = "0"
        End If
        
         If chkBonificaFA.value = 1 Then
            sBonificaFA = "1"
        Else
            sBonificaFA = "0"
        End If
            
    If Accion = Add Then
    
        If Not (rst.EOF And rst.BOF) Then
            sFiltro = "IDProducto = " & txtCodigo.Text
            If ExiteRstKey(rst, sFiltro) Then
               lbok = Mensaje("Ya existe el producto ", ICO_ERROR, False)
                txtCodigo.SetFocus
            Exit Sub
            End If
        End If
    
                lbok = invUpdateProducto("I", txtCodigo.Text, txtDescr.Text, txtImpuesto.Text, sEsMuestra, sEsControlado, txtClasif1.Text, _
                    txtClasif2.Text, txtClasif3.Text, sEsEtico, sBajaPrecioDistribuidor, txtCodProveedor.Text, txtCostoUltLocal.Text, txtCostoUltDolar.Text, _
                    txtCostoUltPromLocal.Text, txtCostoUltPromDolar.Text, txtPrecioPublico.Text, Me.txtPrecioFarmaciaLocal.Text, Me.txtPrecioCIFLocal.Text, _
                    txtPrecioFOBLocal.Text, txtIDPresentacion.Text, sBajaPrecioProveedor, Me.txtPorcDescAlzaProveedor.Text, gsUSUARIO, gsUSUARIO, sActivo, txtCodigoBarra.Text, _
                    sBonificaFA, txtCOPorCada.Text, txtCOBonifica.Text)
            
            If lbok Then
                sMsg = "El Producto ha sido registrado exitosamente ... "
                lbok = Mensaje(sMsg, ICO_OK, False)
                ' actualiza datos
                cargaGrid
                Accion = View
                HabilitarControles
                HabilitarBotones
            Else
                 sMsg = "Ha ocurrido un error tratando de Agregar el Produto ... "
                lbok = Mensaje(sMsg, ICO_ERROR, False)
            End If
    End If ' si estoy adicionando
    If Accion = Edit Then
        If Not (rst.EOF And rst.BOF) Then
                lbok = invUpdateProducto("U", txtCodigo.Text, txtDescr.Text, txtImpuesto.Text, sEsMuestra, sEsControlado, txtClasif1.Text, _
                    txtClasif2.Text, txtClasif3.Text, sEsEtico, sBajaPrecioDistribuidor, txtCodProveedor.Text, txtCostoUltLocal.Text, txtCostoUltDolar.Text, _
                    txtCostoUltPromLocal.Text, txtCostoUltPromDolar.Text, txtPrecioPublico.Text, Me.txtPrecioFarmaciaLocal.Text, Me.txtPrecioCIFLocal.Text, _
                    txtPrecioFOBLocal.Text, txtIDPresentacion.Text, sBajaPrecioProveedor, Me.txtPorcDescAlzaProveedor.Text, gsUSUARIO, gsUSUARIO, sActivo, txtCodigoBarra.Text, _
                    sBonificaFA, txtCOPorCada.Text, txtCOBonifica.Text)
    
            If lbok Then
                sMsg = "Los datos fueron grabados Exitosamente ... "
                lbok = Mensaje(sMsg, ICO_OK, False)
                ' actualiza datos
                cargaGrid
                Accion = View
                HabilitarControles
                HabilitarBotones
            Else
                sMsg = "Ha ocurrido un error tratando de actualizar los datos del producto... "
                lbok = Mensaje(sMsg, ICO_ERROR, False)
            End If
        End If
       End If ' si estoy adicionando

End Sub

Private Sub cmdUndo_Click()
    GetDataFromGridToControl
    Accion = View
    HabilitarControles
    HabilitarBotones
End Sub

Private Sub Form_Activate()
    HighlightInWin Me.Name
    SetupFormToolbar (Me.Name)
End Sub

Private Sub Form_Load()
    Dim lbok As Boolean
    MDIMain.AddForm Me.Name
    Set rst = New ADODB.Recordset
    If rst.State = adStateOpen Then rst.Close
    rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rst.CursorLocation = adUseClient ' Cursor local al cliente
    rst.LockType = adLockOptimistic
    
    Set rst2 = New ADODB.Recordset
    If rst2.State = adStateOpen Then rst2.Close
    rst2.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst2.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rst2.CursorLocation = adUseClient ' Cursor local al cliente
    rst2.LockType = adLockOptimistic
    
    Set rst3 = New ADODB.Recordset
    If rst3.State = adStateOpen Then rst3.Close
    rst3.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst3.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rst3.CursorLocation = adUseClient ' Cursor local al cliente
    rst3.LockType = adLockOptimistic

    Me.Caption = gsFormCaption
    Me.lbFormCaption = gsTitle

    Accion = View
    HabilitarControles
    HabilitarBotones
    lbok = CargaTablas()
    cargaGrid

End Sub

Private Sub InicializaListView()
    Dim sItem As String
        With ListView1
            ' Las pruebas serán en modo "detalle"
            .View = lvwReport
            ' al seleccionar un elemento, seleccionar la línea completa
            .FullRowSelect = True
            ' Mostrar las líneas de la cuadrícula
            .GridLines = True
            ' No permitir la edición automática del texto
            .LabelEdit = lvwManual
            ' Permitir múltiple selección
            .MultiSelect = False
            ' Para que al perder el foco,
            ' se siga viendo el que está seleccionado
            .HideSelection = False
            .LabelWrap = False
            .ForeColor = vbBlue
    
        End With
    
        With ListView1.ColumnHeaders.Add(, , "Descripción", 4500)
            
                .Tag = cTexto
        End With
    
     
        '
        ' Eliminar las cabeceras
        ListView1.ColumnHeaders.Clear
        '
        ' Asignar las cabeceras
        With ListView1.ColumnHeaders.Add(, , "***************   PRODUCTOS   **************", 4500)
            .Tag = cTexto
        End With
        
     ListView1.ListItems.Clear
        ' Asignar algunos datos aleatorios
        If Not (rst.EOF And rst.BOF) Then
            rst.MoveFirst
            While Not rst.EOF
            
            sItem = Trim(right("00000" + Trim(Str(rst("IDProducto").value)), 5)) + "-" + rst("Descr").value
            If Len(sItem) > 50 Then
                sItem = Mid(sItem, 1, 50) & vbLf & Mid(sItem, 51, Len(sItem))
                
            End If
            
                With ListView1.ListItems.Add(, , sItem)
                    ' Cada subitem debe corresponder con cada una de las cabeceras
                    ' la segunda cabecera es el Subitems(1) y así sucesivamente
                End With
                rst.MoveNext
            Wend
        End If

End Sub




Private Sub cargaGrid()
    If rst.State = adStateOpen Then rst.Close
    rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst.CursorType = adOpenKeyset 'adOpenKeyset  'Asigna un cursor dinamico
    rst.CursorLocation = adUseClient ' Cursor local al cliente
    rst.LockType = adLockOptimistic
    GSSQL = gsCompania & ".invGetProductos -1"
    If rst.State = adStateOpen Then rst.Close
    Set rst = GetRecordset(GSSQL)
    If Not (rst.EOF And rst.BOF) Then
      Set TDBG.DataSource = rst
      'CargarDatos rst, TDBG, "Codigo", "Descr"
      TDBG.Refresh
      'IniciaIconos
      rst.MoveFirst
                InicializaListView
                ListView1.ListItems(1).Selected = True
                ShowSelectedItem
           ' IniciaIconos
            If chkBonificaFA.value = 1 Then
                cmdBonifica.Enabled = True
            Else
                cmdBonifica.Enabled = False
            End If

    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (rst Is Nothing) Then Set rst = Nothing
    If Not (rst2 Is Nothing) Then Set rst2 = Nothing
    SetupFormToolbar ("no name")
    MDIMain.SubtractForm Me.Name
    Set frmProductos = Nothing
End Sub

Private Sub ListView1_Click()
    ShowSelectedItem
End Sub


Private Sub ShowSelectedItem()
    Dim sValor As String
    Dim sFiltro As String
    Dim i As Integer
    If Not (rst.EOF And rst.BOF) Then
        With ListView1.SelectedItem
            sValor = .Text
            .ToolTipText = sValor
            i = InStr(sValor, "-")
            i = i - 1
            sValor = Mid(sValor, 1, i)
            sFiltro = "IDProducto=" & Str(Val(sValor))
            rst.MoveFirst
            rst.Find sFiltro
            'rst.Bookmark = getPositionRecord(rst, sFiltro)
            GetDataFromGridToControl
        End With
    End If
End Sub


Private Sub SSActiveTabs1_BeforeTabClick(ByVal NewTab As ActiveTabs.SSTab, ByVal Cancel As ActiveTabs.SSReturnBoolean)
    If NewTab.Key = "GENERAL" Then
    '    SSActiveTabs1.Height = 3255
        'TDBGMov.Height = 3255
    End If
    If NewTab.Key = "PRECIOS" Then
    '    SSActiveTabs1.Height = 3255
        'TDBGMov.Height = 3255
    End If
    
    If NewTab.Key = "EXISTENCIA" Then
        txtBodega.Text = ""
        txtDescrBodega.Text = ""
    '    SSActiveTabs1.Height = 3255
        'TDBGMov.Height = 3255
    End If
    If NewTab.Key = "MOVIMIENTOS" Then
    '    SSActiveTabs1.Height = Me.Height - 200
        'TDBGMov.Height = Me.Height - 100
    End If

End Sub

Private Sub TDBG_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    GetDataFromGridToControl
End Sub

Private Function CargaExistenciaBodega(sIDArticulo As String, sIDBodega As String)
    Dim lbok As Boolean
    On Error GoTo error
    lbok = True
      GSSQL = gsCompania & ".invGetExistenciaBodega " & sIDArticulo & " , " & sIDBodega
    
      'Set rst2 = gConet.Execute(GSSQL, adCmdText)  'Ejecuta la sentencia
      rst3.Open GSSQL, gConet, adOpenDynamic, adLockOptimistic
    
      If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
        gsOperacionError = "No existe ese cliente." 'Asigna msg de error
        lbok = False  'Indica que no es válido
        
      ElseIf Not (rst3.BOF And rst3.EOF) Then  'Si no es válido
        Set TDBGExistencia.DataSource = rst3
        TDBGExistencia.Refresh
      End If
      CargaExistenciaBodega = lbok
      'rst3.Close
      Exit Function
error:
      lbok = False
      gsOperacionError = "Ocurrió un error en la operación de los datos " & err.Description
      Resume Next
End Function

Private Function CargaTablas() As Boolean
    Dim lbok As Boolean
    On Error GoTo error
    lbok = True
      GSSQL = gsCompania & ".globalGetTablas -1 "
              
      'Set rst2 = gConet.Execute(GSSQL, adCmdText)  'Ejecuta la sentencia
      rst2.Open GSSQL, gConet, adOpenDynamic, adLockOptimistic
    
      If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
        gsOperacionError = "No existe ese cliente." 'Asigna msg de error
        lbok = False  'Indica que no es válido
        
      ElseIf Not (rst2.BOF And rst2.EOF) Then  'Si no es válido
        rst2.MoveNext
        lbok = SetLable(rst2, "NOMBRE='LINEA'", lblClasif1)
        lbok = SetLable(rst2, "NOMBRE='FAMILIA'", lblClasif2)
        lbok = SetLable(rst2, "NOMBRE='SUBFAMILIA'", lblClasif3)
        lbok = SetLable(rst2, "NOMBRE='PRESENTACION'", lblPresentación)
        lbok = SetLable(rst2, "NOMBRE='IMPUESTO'", lblImpuesto)
        lbok = True
      End If
      CargaTablas = lbok
      rst2.Close
      Exit Function
error:
      lbok = False
      gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
      Resume Next
End Function

Private Function SetLable(rstFuente As ADODB.Recordset, sFiltro As String, lbl As Label) As Boolean
    Dim lbok As Boolean
    Dim rstClone As ADODB.Recordset
    Dim bmPos As Variant
    lbok = False
    If Not (rstFuente.EOF And rstFuente.BOF) Then
        Set rstClone = New ADODB.Recordset
            bmPos = rstFuente.Bookmark
            rstClone.Filter = adFilterNone
            Set rstClone = rstFuente.Clone
            rstClone.Filter = sFiltro
            If Not rstClone.EOF Then ' Si existe
              lbl.Caption = rstClone("DescrUsuario").value & " :"
              lbl.Tag = rstClone("Nombre").value
              lbok = True
            End If
            rstFuente.Filter = adFilterNone
            rstFuente.Bookmark = bmPos
        rstClone.Filter = adFilterNone
    End If
    SetLable = lbok
End Function

Private Function ControlsOk() As Boolean
    Dim lbok As Boolean
    If txtCodigo.Text = "" Then
        lbok = Mensaje("El Código del Producto no puede quedar en Blanco...", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If Not Val_TextboxNum(txtCodigo) Then
        lbok = Mensaje("El Código del Producto debe ser Numérico", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    
    
    If txtDescr.Text = "" Then
        lbok = Mensaje("La Descripción del Producto no puede quedar en Blanco...", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If txtClasif1.Text = "" Then
        lbok = Mensaje("La Clasificación1 del Producto no puede quedar en Blanco...", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If txtClasif2.Text = "" Then
        lbok = Mensaje("La Clasificación2 del Producto no puede quedar en Blanco...", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If txtClasif3.Text = "" Then
        lbok = Mensaje("La Clasificación3 del Producto no puede quedar en Blanco...", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If txtImpuesto.Text = "" Then
        lbok = Mensaje("EL Impuesto del Producto no puede quedar en Blanco...", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If txtIDPresentacion.Text = "" Then
        lbok = Mensaje("La Presentación del Producto no puede quedar en Blanco...", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    If txtCodProveedor.Text = "" Then
        lbok = Mensaje("EL Proveedor del Producto no puede quedar en Blanco...", ICO_ERROR, False)
        SSActiveTabs1.SelectedTab = 2
        ControlsOk = False
        Exit Function
    End If
    
    If txtCodigoBarra.Text = "" Then
        txtCodigoBarra.Text = "ND"
    End If
    
    If Not Val_TextboxNum(txtCostoUltDolar) Then
        lbok = Mensaje("El Costo Ultimo Dolar del Producto debe ser Numérico", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If Not Val_TextboxNum(txtCostoUltLocal) Then
        lbok = Mensaje("El Costo Ultimo Córdoba del Producto debe ser Numérico", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If Not Val_TextboxNum(txtCostoUltPromDolar) Then
        lbok = Mensaje("El Costo Ultimo Promedio Dolar del Producto debe ser Numérico", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If Not Val_TextboxNum(txtCostoUltPromLocal) Then
        lbok = Mensaje("El Costo Ultimo Promedio Córdoba del Producto debe ser Numérico", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If Not Val_TextboxNum(txtPorcDescAlzaProveedor) Then
        lbok = Mensaje("El % de Alza o Descuento del Proveedor debe ser Numérico", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If Not Val_TextboxNum(txtPrecioCIFLocal) Then
        lbok = Mensaje("El Precio CIF del Proveedor debe ser Numérico", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If Not Val_TextboxNum(txtPrecioFOBLocal) Then
        lbok = Mensaje("El Precio FOB del Proveedor debe ser Numérico", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If Not Val_TextboxNum(txtPrecioFarmaciaLocal) Then
        lbok = Mensaje("El Precio Farmacia del Producto debe ser Numérico", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If Not Val_TextboxNum(txtPrecioPublico) Then
        lbok = Mensaje("El Precio Público del Producto debe ser Numérico", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If Not Val_TextboxNum(txtCOPorCada) Then
        lbok = Mensaje("El Valor de Bonificación <Por Cada> debe ser Numérico", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If Not Val_TextboxNum(txtCOBonifica) Then
        lbok = Mensaje("El Valor de Bonificación debe ser Numérico", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
   

    
    ControlsOk = True
End Function


