VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProductosINICIAL 
   Caption         =   "Catálogo de Productos "
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   8730
   StartUpPosition =   1  'CenterOwner
   Begin ActiveTabs.SSActiveTabs SSActiveTabs1 
      Height          =   3255
      Left            =   240
      TabIndex        =   12
      Top             =   600
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   5741
      _Version        =   131083
      TabCount        =   4
      Tabs            =   "frmProductosINICIAL.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   2865
         Left            =   -99969
         TabIndex        =   79
         Top             =   360
         Width           =   14715
         _ExtentX        =   25956
         _ExtentY        =   5054
         _Version        =   131083
         TabGuid         =   "frmProductosINICIAL.frx":0114
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
            Left            =   10920
            TabIndex        =   94
            Top             =   240
            Width           =   3375
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   9240
            TabIndex        =   93
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdTipoMov 
            Height          =   320
            Left            =   8760
            Picture         =   "frmProductosINICIAL.frx":013C
            Style           =   1  'Graphical
            TabIndex        =   91
            Top             =   240
            Width           =   315
         End
         Begin VB.CommandButton cmdDelTipoMov 
            Height          =   320
            Left            =   10560
            Picture         =   "frmProductosINICIAL.frx":047E
            Style           =   1  'Graphical
            TabIndex        =   90
            Top             =   240
            Width           =   315
         End
         Begin VB.CommandButton cmdDelBodMov 
            Height          =   320
            Left            =   3120
            Picture         =   "frmProductosINICIAL.frx":08C0
            Style           =   1  'Graphical
            TabIndex        =   84
            Top             =   240
            Width           =   300
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1920
            TabIndex        =   83
            Top             =   240
            Width           =   1095
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   3600
            TabIndex        =   82
            Top             =   240
            Width           =   4335
         End
         Begin VB.CommandButton cmdBodMov 
            Height          =   320
            Left            =   1320
            Picture         =   "frmProductosINICIAL.frx":0D02
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   240
            Width           =   300
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Refrescar"
            Height          =   375
            Left            =   11160
            TabIndex        =   80
            Top             =   720
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker DTPFecInic 
            Height          =   255
            Left            =   1920
            TabIndex        =   86
            Top             =   720
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
            Format          =   61997057
            CurrentDate     =   41692
            MinDate         =   41690
         End
         Begin MSComCtl2.DTPicker DTPFechaFin 
            Height          =   255
            Left            =   5280
            TabIndex        =   87
            Top             =   720
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
            Format          =   61997057
            CurrentDate     =   41698
         End
         Begin TrueOleDBGrid60.TDBGrid TDBGMov 
            Height          =   8295
            Left            =   120
            OleObjectBlob   =   "frmProductosINICIAL.frx":1044
            TabIndex        =   95
            Top             =   1200
            Width           =   14505
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
            Height          =   255
            Left            =   8040
            TabIndex        =   92
            Top             =   240
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
            Height          =   255
            Left            =   4320
            TabIndex        =   89
            Top             =   720
            Width           =   735
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
            Height          =   255
            Left            =   1080
            TabIndex        =   88
            Top             =   720
            Width           =   735
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
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   240
            Width           =   1095
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   2865
         Left            =   -99969
         TabIndex        =   71
         Top             =   360
         Width           =   14715
         _ExtentX        =   25956
         _ExtentY        =   5054
         _Version        =   131083
         TabGuid         =   "frmProductosINICIAL.frx":7AE4
         Begin TrueOleDBGrid60.TDBGrid TDBGExistencia 
            Height          =   1695
            Left            =   360
            OleObjectBlob   =   "frmProductosINICIAL.frx":7B0C
            TabIndex        =   78
            Top             =   720
            Width           =   13095
         End
         Begin VB.CommandButton cmdRefresExistencia 
            Caption         =   "Refrescar"
            Height          =   375
            Left            =   8760
            TabIndex        =   77
            Top             =   120
            Width           =   1575
         End
         Begin VB.CommandButton cmdBodega 
            Height          =   320
            Left            =   1560
            Picture         =   "frmProductosINICIAL.frx":BD2F
            Style           =   1  'Graphical
            TabIndex        =   75
            Top             =   120
            Width           =   300
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   3840
            TabIndex        =   74
            Top             =   120
            Width           =   4455
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2160
            TabIndex        =   73
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton cmdDelBodega 
            Height          =   320
            Left            =   3360
            Picture         =   "frmProductosINICIAL.frx":C071
            Style           =   1  'Graphical
            TabIndex        =   72
            Top             =   120
            Width           =   300
         End
         Begin VB.Label Label2 
            Caption         =   "Bodegar:"
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
            Left            =   360
            TabIndex        =   76
            Top             =   120
            Width           =   1095
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   2865
         Index           =   0
         Left            =   30
         TabIndex        =   13
         Top             =   360
         Width           =   14715
         _ExtentX        =   25956
         _ExtentY        =   5054
         _Version        =   131083
         TabGuid         =   "frmProductosINICIAL.frx":C4B3
         Begin VB.Frame Frame2 
            Height          =   2775
            Left            =   120
            TabIndex        =   15
            Top             =   0
            Width           =   14415
            Begin VB.CommandButton cmdDelImpuesto 
               Height          =   320
               Left            =   3240
               Picture         =   "frmProductosINICIAL.frx":C4DB
               Style           =   1  'Graphical
               TabIndex        =   62
               Top             =   1800
               Width           =   300
            End
            Begin VB.TextBox txtImpuesto 
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
               Left            =   2160
               TabIndex        =   61
               Top             =   1800
               Width           =   855
            End
            Begin VB.TextBox txtDescrImpuesto 
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
               Left            =   3720
               TabIndex        =   60
               Top             =   1800
               Width           =   6135
            End
            Begin VB.CommandButton cmdImpuesto 
               Height          =   320
               Left            =   1560
               Picture         =   "frmProductosINICIAL.frx":C91D
               Style           =   1  'Graphical
               TabIndex        =   59
               Top             =   1800
               Width           =   300
            End
            Begin VB.CommandButton cmdPresentacion 
               Height          =   320
               Left            =   1560
               Picture         =   "frmProductosINICIAL.frx":CC5F
               Style           =   1  'Graphical
               TabIndex        =   56
               Top             =   2280
               Width           =   300
            End
            Begin VB.TextBox txtDescrPresentacion 
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
               Left            =   3720
               TabIndex        =   55
               Top             =   2280
               Width           =   6135
            End
            Begin VB.TextBox txtIDPresentacion 
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
               Left            =   2160
               TabIndex        =   54
               Top             =   2280
               Width           =   855
            End
            Begin VB.CommandButton cmdDelPresentacion 
               Height          =   320
               Left            =   3240
               Picture         =   "frmProductosINICIAL.frx":CFA1
               Style           =   1  'Graphical
               TabIndex        =   53
               Top             =   2280
               Width           =   300
            End
            Begin VB.CommandButton cmdClasif1 
               Height          =   320
               Left            =   1560
               Picture         =   "frmProductosINICIAL.frx":D3E3
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   360
               Width           =   300
            End
            Begin VB.TextBox txtDecrClasif1 
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
               Left            =   3690
               TabIndex        =   30
               Top             =   345
               Width           =   6135
            End
            Begin VB.TextBox txtClasif1 
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
               Left            =   2160
               TabIndex        =   29
               Top             =   360
               Width           =   855
            End
            Begin VB.CommandButton cmdDelclasif1 
               Height          =   320
               Left            =   3240
               Picture         =   "frmProductosINICIAL.frx":D725
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   360
               Width           =   300
            End
            Begin VB.TextBox txtCodigoBarra 
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
               TabIndex        =   27
               Top             =   2280
               Width           =   3975
            End
            Begin VB.CommandButton cmdClasif2 
               Height          =   320
               Left            =   1560
               Picture         =   "frmProductosINICIAL.frx":DB67
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   840
               Width           =   300
            End
            Begin VB.TextBox txtDecrClasif2 
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
               Left            =   3720
               TabIndex        =   25
               Top             =   840
               Width           =   6135
            End
            Begin VB.TextBox txtClasif2 
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
               Left            =   2160
               TabIndex        =   24
               Top             =   840
               Width           =   855
            End
            Begin VB.CommandButton cmdDelclasif2 
               Height          =   320
               Left            =   3240
               Picture         =   "frmProductosINICIAL.frx":DEA9
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   840
               Width           =   300
            End
            Begin VB.CommandButton cmdClasif3 
               Height          =   320
               Left            =   1560
               Picture         =   "frmProductosINICIAL.frx":E2EB
               Style           =   1  'Graphical
               TabIndex        =   22
               Top             =   1320
               Width           =   300
            End
            Begin VB.TextBox txtDecrClasif3 
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
               Left            =   3720
               TabIndex        =   21
               Top             =   1320
               Width           =   6135
            End
            Begin VB.TextBox txtClasif3 
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
               Left            =   2160
               TabIndex        =   20
               Top             =   1320
               Width           =   855
            End
            Begin VB.CommandButton cmdDelclasif3 
               Height          =   320
               Left            =   3240
               Picture         =   "frmProductosINICIAL.frx":E62D
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   1320
               Width           =   300
            End
            Begin VB.CheckBox chkControlado 
               Caption         =   "Controlado ?"
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
               Left            =   11400
               TabIndex        =   18
               Top             =   1440
               Width           =   1575
            End
            Begin VB.CheckBox chkEtico 
               Caption         =   "Etico ?"
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
               Left            =   11400
               TabIndex        =   17
               Top             =   960
               Width           =   1095
            End
            Begin VB.CheckBox chkEsMuestra 
               Caption         =   "Muestra Médica ?"
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
               Left            =   11400
               TabIndex        =   16
               Top             =   480
               Width           =   1935
            End
            Begin VB.Label lblImpuesto 
               Caption         =   "Impuesto:"
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
               Left            =   240
               TabIndex        =   63
               Top             =   1800
               Width           =   1215
            End
            Begin VB.Label lblPresentación 
               Caption         =   "Presentación:"
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
               Left            =   240
               TabIndex        =   57
               Top             =   2280
               Width           =   1335
            End
            Begin VB.Label lblClasif1 
               Caption         =   "Clasificación1:"
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
               Left            =   240
               TabIndex        =   35
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label lblCodigoBarra 
               Caption         =   "Código Barra:"
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
               Left            =   10320
               TabIndex        =   34
               Top             =   1920
               Width           =   1335
            End
            Begin VB.Label lblClasif2 
               Caption         =   "Clasificación2:"
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
               Left            =   240
               TabIndex        =   33
               Top             =   840
               Width           =   1335
            End
            Begin VB.Label lblClasif3 
               Caption         =   "Clasificación3:"
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
               Left            =   240
               TabIndex        =   32
               Top             =   1320
               Width           =   1335
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   2865
         Index           =   1
         Left            =   -99969
         TabIndex        =   14
         Top             =   360
         Width           =   14715
         _ExtentX        =   25956
         _ExtentY        =   5054
         _Version        =   131083
         TabGuid         =   "frmProductosINICIAL.frx":EA6F
         Begin VB.Frame Frame1 
            Height          =   2655
            Left            =   120
            TabIndex        =   36
            Top             =   0
            Width           =   14295
            Begin VB.CheckBox chkBajaPrecioProveedor 
               Caption         =   "Se afecta el Precio con un Descuento/Alza del Proveedor?"
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
               Left            =   2640
               TabIndex        =   70
               Top             =   840
               Width           =   5415
            End
            Begin VB.CheckBox chkBajaPrecioDistribuidor 
               Caption         =   "Se afecta el Precio con un Descuento/Alza del Distribuidor?"
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
               Left            =   8280
               TabIndex        =   69
               Top             =   840
               Width           =   5535
            End
            Begin VB.CommandButton cmdDelProveedor 
               Height          =   320
               Left            =   3120
               Picture         =   "frmProductosINICIAL.frx":EA97
               Style           =   1  'Graphical
               TabIndex        =   67
               Top             =   360
               Width           =   300
            End
            Begin VB.TextBox txtCodProveedor 
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
               Left            =   1920
               TabIndex        =   66
               Top             =   360
               Width           =   1095
            End
            Begin VB.TextBox txtDescrProveedor 
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
               Left            =   3600
               TabIndex        =   65
               Top             =   360
               Width           =   8415
            End
            Begin VB.CommandButton cmdProveedor 
               Height          =   320
               Left            =   1320
               Picture         =   "frmProductosINICIAL.frx":EED9
               Style           =   1  'Graphical
               TabIndex        =   64
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
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   5640
               TabIndex        =   58
               ToolTipText     =   "Si el signo es Positivo  es un alza, si es Negativo es una rebaja"
               Top             =   1200
               Width           =   1095
            End
            Begin VB.TextBox txtCostoUltPromDolar 
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
               Left            =   12720
               TabIndex        =   51
               Top             =   2160
               Width           =   1095
            End
            Begin VB.TextBox txtCostoUltPromLocal 
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
               Left            =   9240
               TabIndex        =   49
               Top             =   2160
               Width           =   1095
            End
            Begin VB.TextBox txtCostoUltDolar 
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
               Left            =   5640
               TabIndex        =   47
               Top             =   2160
               Width           =   1095
            End
            Begin VB.TextBox txtCostoUltLocal 
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
               Left            =   1920
               TabIndex        =   45
               Top             =   2160
               Width           =   1095
            End
            Begin VB.TextBox txtPrecioFOBLocal 
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
               Left            =   12720
               TabIndex        =   43
               Top             =   1680
               Width           =   1095
            End
            Begin VB.TextBox txtPrecioCIFLocal 
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
               Left            =   9240
               TabIndex        =   41
               Top             =   1680
               Width           =   1095
            End
            Begin VB.TextBox txtPrecioFarmaciaLocal 
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
               Left            =   5640
               TabIndex        =   39
               Top             =   1680
               Width           =   1095
            End
            Begin VB.TextBox txtPrecioPublico 
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
               Left            =   1920
               TabIndex        =   37
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label lblProveedor 
               Caption         =   "Proveedor:"
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
               Left            =   120
               TabIndex        =   68
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label lblPorcDescAlzaProveedor 
               Caption         =   "Porcentaje Alza/Baja Proveedor en Precios:"
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
               Left            =   120
               TabIndex        =   0
               Top             =   1200
               Width           =   4335
            End
            Begin VB.Label lblCostoUltPromDolar 
               Caption         =   "Costo Ult Promedio $:"
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
               Left            =   10680
               TabIndex        =   52
               Top             =   2160
               Width           =   1935
            End
            Begin VB.Label lblCostoUltPromLocal 
               Caption         =   "Costo Ult Promedio C$:"
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
               Left            =   7080
               TabIndex        =   50
               Top             =   2160
               Width           =   2055
            End
            Begin VB.Label lblCostoUltDolar 
               Caption         =   "Costo Ultimo $ :"
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
               Left            =   3720
               TabIndex        =   48
               Top             =   2160
               Width           =   1815
            End
            Begin VB.Label lblCostoUltLocal 
               Caption         =   "Costo Ultimo C$:"
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
               Left            =   120
               TabIndex        =   46
               Top             =   2160
               Width           =   1575
            End
            Begin VB.Label lblPrecioFOBLocal 
               Caption         =   "Precio/Costo FOB C$:"
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
               Left            =   10680
               TabIndex        =   44
               Top             =   1680
               Width           =   1935
            End
            Begin VB.Label lblCostoCIF 
               Caption         =   "Precio/Costo CIF C$:"
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
               Left            =   7200
               TabIndex        =   42
               Top             =   1680
               Width           =   1935
            End
            Begin VB.Label lblPrecioFarmacia 
               Caption         =   "Precio Farmacia C$ :"
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
               Left            =   3720
               TabIndex        =   40
               Top             =   1680
               Width           =   1815
            End
            Begin VB.Label lblPrecioPublico 
               Caption         =   "Precio Público C$:"
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
               Left            =   120
               TabIndex        =   38
               Top             =   1680
               Width           =   1575
            End
         End
      End
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
      Left            =   12720
      TabIndex        =   9
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtDescr 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "xdfsdfdsf"
      Top             =   240
      Width           =   8055
   End
   Begin VB.TextBox txtCodigo 
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
      Left            =   1800
      TabIndex        =   7
      Top             =   240
      Width           =   1095
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
      Height          =   495
      Left            =   14520
      Picture         =   "frmProductosINICIAL.frx":F21B
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Deshacer / Cancelar"
      Top             =   6720
      Width           =   495
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
      Height          =   495
      Left            =   14520
      Picture         =   "frmProductosINICIAL.frx":F65D
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
      Top             =   4920
      Width           =   495
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
      Height          =   495
      Left            =   14520
      Picture         =   "frmProductosINICIAL.frx":FA9F
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Agrega el item con los datos digitados..."
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   495
      Left            =   14520
      Picture         =   "frmProductosINICIAL.frx":FDA9
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
      Top             =   6120
      Width           =   495
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
      Height          =   495
      Left            =   14520
      Picture         =   "frmProductosINICIAL.frx":100B3
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
      Top             =   4320
      Width           =   495
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   4815
      Left            =   240
      OleObjectBlob   =   "frmProductosINICIAL.frx":1097D
      TabIndex        =   6
      Top             =   4200
      Width           =   14025
   End
   Begin VB.Label Label4 
      Caption         =   "Descripción :"
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
      Left            =   3000
      TabIndex        =   11
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Código :"
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
      Left            =   720
      TabIndex        =   10
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmProductosINICIAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As ADODB.Recordset
Dim rst2 As ADODB.Recordset
Dim rst3 As ADODB.Recordset
Dim bOrdenCodigo As Boolean
Dim bOrdenDescr As Boolean
Dim sCodSucursal As String
Dim bEdit As Boolean
Dim bAdd As Boolean
Dim sSoloActivo As String

Private Sub cmdAdd_Click()
bAdd = True
bEdit = False
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



cmdSave.Enabled = True
cmdEliminar.Enabled = False
cmdAdd.Enabled = False
txtDescr.SetFocus
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
bEdit = True
bAdd = False
GetDataFromGridToControl
txtCodigo.Enabled = True
txtDescr.Enabled = True
fmtTextbox txtCodigo, "R"
fmtTextbox txtDescr, "O"
chkActivo.Enabled = True
cmdSave.Enabled = True
cmdEliminar.Enabled = False
cmdAdd.Enabled = False

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
    txtClasif1.Text = rst("CLASIFICACION1").value
    txtDecrClasif1.Text = rst("DESCRCLASIF1").value
    txtClasif2.Text = rst("CLASIFICACION2").value
    txtDecrClasif2.Text = rst("DESCRCLASIF2").value
    txtClasif3.Text = rst("CLASIFICACION3").value
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
    
    
Else
    txtCodigo.Text = ""
    txtDescr.Text = ""
End If

End Sub

Private Sub cmdEliminar_Click()
Dim lbOk As Boolean
Dim sMsg As String
Dim sTipo As String
Dim sFiltro As String
Dim sActivo As String
Dim sEsMuestra As String
Dim sEsControlado As String
Dim sEsEtico As String
Dim sBajaPrecioDistribuidor As String
Dim sBajaPrecioProveedor As String

    If txtCodigo.Text = "" Then
        lbOk = Mensaje("El Código del Producto no puede estar en Blanco", ICO_ERROR, False)
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
   
    ' hay que validar la integridad referencial
    ' if exists dependencias then No se puede eliminar
    lbOk = Mensaje("Está seguro de eliminar el Producto " & rst("Descr").value, ICO_PREGUNTA, True)
    If lbOk Then
'                lbOk = invUpdateProducto("D", txtCodigo.Text, txtDescr.Text, txtImpuesto.Text, sEsMuestra, sEsControlado, txtClasif1.Text, _
'                txtClasif2.Text, txtClasif3.Text, sEsEtico, sBajaPrecioDistribuidor, txtCodProveedor.Text, txtCostoUltLocal.Text, txtCostoUltDolar.Text, _
'                txtCostoUltPromLocal.Text, txtCostoUltPromDolar.Text, txtPrecioPublico.Text, Me.txtPrecioFarmaciaLocal.Text, Me.txtPrecioCIFLocal.Text, _
'                txtPrecioFOBLocal.Text, txtIDPresentacion.Text, sBajaPrecioProveedor, Me.txtPorcDescAlzaProveedor.Text, gsUSUARIO, gsUSUARIO, sActivo)
        
        If lbOk Then
            sMsg = "Borrado Exitosamente ... "
            lbOk = Mensaje(sMsg, ICO_OK, False)
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
Dim lbOk As Boolean
Dim sIDArticulo As String
Dim sIDBodega As String
If txtBodMov.Text = "" Then
    sIDBodega = "-1"
End If

sIDArticulo = txtCodigo.Text

lbOk = CargaExistenciaBodega(sIDArticulo, sIDBodega)

End Sub

Private Sub cmdSave_Click()
Dim lbOk As Boolean
Dim sMsg As String
Dim sActivo As String
Dim sFiltro As String

Dim sEsMuestra As String
Dim sEsControlado As String
Dim sEsEtico As String
Dim sBajaPrecioDistribuidor As String
Dim sBajaPrecioProveedor As String

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
        
If bAdd Then

    If Not (rst.EOF And rst.BOF) Then
        sFiltro = "IDProducto = " & txtCodigo.Text
        If ExiteRstKey(rst, sFiltro) Then
           lbOk = Mensaje("Ya existe ese Departamento ", ICO_ERROR, False)
            txtCodigo.SetFocus
        Exit Sub
        End If
    End If

            lbOk = invUpdateProducto("I", txtCodigo.Text, txtDescr.Text, txtImpuesto.Text, sEsMuestra, sEsControlado, txtClasif1.Text, _
                txtClasif2.Text, txtClasif3.Text, sEsEtico, sBajaPrecioDistribuidor, txtCodProveedor.Text, txtCostoUltLocal.Text, txtCostoUltDolar.Text, _
                txtCostoUltPromLocal.Text, txtCostoUltPromDolar.Text, txtPrecioPublico.Text, Me.txtPrecioFarmaciaLocal.Text, Me.txtPrecioCIFLocal.Text, _
                txtPrecioFOBLocal.Text, txtIDPresentacion.Text, sBajaPrecioProveedor, Me.txtPorcDescAlzaProveedor.Text, gsUSUARIO, gsUSUARIO, sActivo, txtCodigoBarra.Text)
        
        If lbOk Then
            sMsg = "El Producto ha sido registrado exitosamente ... "
            lbOk = Mensaje(sMsg, ICO_OK, False)
            ' actualiza datos
            cargaGrid
            bEdit = False
            bAdd = False
            initControles
            IniciaIconos
        End If
bAdd = False
End If ' si estoy adicionando
If bEdit Then
    If Not (rst.EOF And rst.BOF) Then
            lbOk = invUpdateProducto("U", txtCodigo.Text, txtDescr.Text, txtImpuesto.Text, sEsMuestra, sEsControlado, txtClasif1.Text, _
                txtClasif2.Text, txtClasif3.Text, sEsEtico, sBajaPrecioDistribuidor, txtCodProveedor.Text, txtCostoUltLocal.Text, txtCostoUltDolar.Text, _
                txtCostoUltPromLocal.Text, txtCostoUltPromDolar.Text, txtPrecioPublico.Text, Me.txtPrecioFarmaciaLocal.Text, Me.txtPrecioCIFLocal.Text, _
                txtPrecioFOBLocal.Text, txtIDPresentacion.Text, sBajaPrecioProveedor, Me.txtPorcDescAlzaProveedor.Text, gsUSUARIO, gsUSUARIO, sActivo, txtCodigoBarra.Text)

        If lbOk Then
            sMsg = "Los datos fueron grabados Exitosamente ... "
            lbOk = Mensaje(sMsg, ICO_OK, False)
            ' actualiza datos
            cargaGrid
            bEdit = False
            bAdd = False
            initControles
            IniciaIconos
        End If
    End If
bEdit = False
End If ' si estoy adicionando

End Sub

Private Sub cmdUndo_Click()
GetDataFromGridToControl
IniciaIconos
End Sub

Private Sub Form_Load()
Dim lbOk As Boolean
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

bEdit = False
bAdd = False
initControles
lbOk = CargaTablas()
cargaGrid
End Sub

Private Sub IniciaIconos()
cmdSave.Enabled = False
cmdEditItem.Enabled = True
cmdEliminar.Enabled = True
cmdAdd.Enabled = True
bEdit = False
bAdd = False

End Sub
Private Sub initControles()
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
End Sub

Private Sub cargaGrid()
Dim sIndependiente As String
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
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not (rst Is Nothing) Then Set rst = Nothing
If Not (rst2 Is Nothing) Then Set rst2 = Nothing
End Sub

Private Sub SSActiveTabs1_BeforeTabClick(ByVal NewTab As ActiveTabs.SSTab, ByVal Cancel As ActiveTabs.SSReturnBoolean)
If NewTab.Key = "GENERAL" Then
    SSActiveTabs1.Height = 3255
    'TDBGMov.Height = 3255
End If
If NewTab.Key = "PRECIOS" Then
    SSActiveTabs1.Height = 3255
    'TDBGMov.Height = 3255
End If

If NewTab.Key = "EXISTENCIA" Then
    txtBodega.Text = ""
    txtDescrBodega.Text = ""
    SSActiveTabs1.Height = 3255
    'TDBGMov.Height = 3255
End If
If NewTab.Key = "MOVIMIENTOS" Then
    SSActiveTabs1.Height = Me.Height - 200
    'TDBGMov.Height = Me.Height - 100
End If

End Sub

Private Sub TDBG_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
GetDataFromGridToControl
'
'IniciaIconos
End Sub

Private Function CargaExistenciaBodega(sIDArticulo As String, sIDBodega As String)
Dim lbOk As Boolean
Dim iResultado As Integer
On Error GoTo error
lbOk = True
  GSSQL = gsCompania & ".invGetExistenciaBodega " & sIDArticulo & " , " & sIDBodega

  'Set rst2 = gConet.Execute(GSSQL, adCmdText)  'Ejecuta la sentencia
  rst3.Open GSSQL, gConet, adOpenDynamic, adLockOptimistic

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbOk = False  'Indica que no es válido
    
  ElseIf Not (rst3.BOF And rst3.EOF) Then  'Si no es válido
    Set TDBGExistencia.DataSource = rst3
    TDBGExistencia.Refresh
  End If
  CargaExistenciaBodega = lbOk
  'rst3.Close
  Exit Function
error:
  lbOk = False
  gsOperacionError = "Ocurrió un error en la operación de los datos " & err.Description
  Resume Next


End Function

Private Function CargaTablas() As Boolean
Dim lbOk As Boolean
Dim iResultado As Integer
On Error GoTo error
lbOk = True
  GSSQL = gsCompania & ".globalGetTablas -1 "
          
  'Set rst2 = gConet.Execute(GSSQL, adCmdText)  'Ejecuta la sentencia
  rst2.Open GSSQL, gConet, adOpenDynamic, adLockOptimistic

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbOk = False  'Indica que no es válido
    
  ElseIf Not (rst2.BOF And rst2.EOF) Then  'Si no es válido
    rst2.MoveNext
    lbOk = SetLable(rst2, "NOMBRE='LINEA'", lblClasif1)
    lbOk = SetLable(rst2, "NOMBRE='FAMILIA'", lblClasif2)
    lbOk = SetLable(rst2, "NOMBRE='SUBFAMILIA'", lblClasif3)
    lbOk = SetLable(rst2, "NOMBRE='PRESENTACION'", lblPresentación)
    lbOk = SetLable(rst2, "NOMBRE='IMPUESTO'", lblImpuesto)
    lbOk = True
  End If
  CargaTablas = lbOk
  rst2.Close
  Exit Function
error:
  lbOk = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function

Private Function SetLable(rstFuente As ADODB.Recordset, sFiltro As String, lbl As Label) As Boolean
Dim lbOk As Boolean
Dim rstClone As ADODB.Recordset
Dim bmPos As Variant
lbOk = False
If Not (rstFuente.EOF And rstFuente.BOF) Then
    Set rstClone = New ADODB.Recordset
        bmPos = rstFuente.Bookmark
        rstClone.Filter = adFilterNone
        Set rstClone = rstFuente.Clone
        rstClone.Filter = sFiltro
        If Not rstClone.EOF Then ' Si existe
          lbl.Caption = rstClone("DescrUsuario").value & " :"
          lbl.Tag = rstClone("Nombre").value
          lbOk = True
        End If
        rstFuente.Filter = adFilterNone
        rstFuente.Bookmark = bmPos
    rstClone.Filter = adFilterNone
End If
SetLable = lbOk
End Function

Private Function ControlsOk() As Boolean
Dim lbOk As Boolean
If txtCodigo.Text = "" Then
    lbOk = Mensaje("El Código del Producto no puede quedar en Blanco...", ICO_ERROR, False)
    ControlsOk = False
    Exit Function
End If

If Not Val_TextboxNum(txtCodigo) Then
    lbOk = Mensaje("El Código del Producto debe ser Numérico", ICO_ERROR, False)
    ControlsOk = False
    Exit Function
End If



If txtDescr.Text = "" Then
    lbOk = Mensaje("La Descripción del Producto no puede quedar en Blanco...", ICO_ERROR, False)
    ControlsOk = False
    Exit Function
End If

If txtClasif1.Text = "" Then
    lbOk = Mensaje("La Clasificación1 del Producto no puede quedar en Blanco...", ICO_ERROR, False)
    ControlsOk = False
    Exit Function
End If

If txtClasif2.Text = "" Then
    lbOk = Mensaje("La Clasificación2 del Producto no puede quedar en Blanco...", ICO_ERROR, False)
    ControlsOk = False
    Exit Function
End If

If txtClasif3.Text = "" Then
    lbOk = Mensaje("La Clasificación3 del Producto no puede quedar en Blanco...", ICO_ERROR, False)
    ControlsOk = False
    Exit Function
End If

If txtImpuesto.Text = "" Then
    lbOk = Mensaje("EL Impuesto del Producto no puede quedar en Blanco...", ICO_ERROR, False)
    ControlsOk = False
    Exit Function
End If

If txtIDPresentacion.Text = "" Then
    lbOk = Mensaje("La Presentación del Producto no puede quedar en Blanco...", ICO_ERROR, False)
    ControlsOk = False
    Exit Function
End If
If txtCodProveedor.Text = "" Then
    lbOk = Mensaje("EL Proveedor del Producto no puede quedar en Blanco...", ICO_ERROR, False)
    SSActiveTabs1.SelectedTab = 2
    ControlsOk = False
    Exit Function
End If

If txtCodigoBarra.Text = "" Then
    txtCodigoBarra.Text = "ND"
End If

If Not Val_TextboxNum(txtCostoUltDolar) Then
    lbOk = Mensaje("El Costo Ultimo Dolar del Producto debe ser Numérico", ICO_ERROR, False)
    ControlsOk = False
    Exit Function
End If

If Not Val_TextboxNum(txtCostoUltLocal) Then
    lbOk = Mensaje("El Costo Ultimo Córdoba del Producto debe ser Numérico", ICO_ERROR, False)
    ControlsOk = False
    Exit Function
End If

If Not Val_TextboxNum(txtCostoUltPromDolar) Then
    lbOk = Mensaje("El Costo Ultimo Promedio Dolar del Producto debe ser Numérico", ICO_ERROR, False)
    ControlsOk = False
    Exit Function
End If

If Not Val_TextboxNum(txtCostoUltPromLocal) Then
    lbOk = Mensaje("El Costo Ultimo Promedio Córdoba del Producto debe ser Numérico", ICO_ERROR, False)
    ControlsOk = False
    Exit Function
End If

If Not Val_TextboxNum(txtPorcDescAlzaProveedor) Then
    lbOk = Mensaje("El % de Alza o Descuento del Proveedor debe ser Numérico", ICO_ERROR, False)
    ControlsOk = False
    Exit Function
End If

If Not Val_TextboxNum(txtPrecioCIFLocal) Then
    lbOk = Mensaje("El Precio CIF del Proveedor debe ser Numérico", ICO_ERROR, False)
    ControlsOk = False
    Exit Function
End If

If Not Val_TextboxNum(txtPrecioFOBLocal) Then
    lbOk = Mensaje("El Precio FOB del Proveedor debe ser Numérico", ICO_ERROR, False)
    ControlsOk = False
    Exit Function
End If

If Not Val_TextboxNum(txtPrecioFarmaciaLocal) Then
    lbOk = Mensaje("El Precio Farmacia del Producto debe ser Numérico", ICO_ERROR, False)
    ControlsOk = False
    Exit Function
End If

If Not Val_TextboxNum(txtPrecioPublico) Then
    lbOk = Mensaje("El Precio Público del Producto debe ser Numérico", ICO_ERROR, False)
    ControlsOk = False
    Exit Function
End If
ControlsOk = True
End Function
