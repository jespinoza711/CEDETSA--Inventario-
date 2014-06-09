VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRegistrarTransaccion 
   Caption         =   "Form1"
   ClientHeight    =   8850
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegistrarTransaccion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8850
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
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
      ScaleWidth      =   11400
      TabIndex        =   56
      Top             =   0
      Width           =   11400
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
         TabIndex        =   58
         Top             =   90
         Width           =   855
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registrar Transacciones el Producto"
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
         TabIndex        =   57
         Top             =   420
         Width           =   2175
      End
      Begin VB.Image Image 
         Height          =   480
         Index           =   2
         Left            =   210
         Picture         =   "frmRegistrarTransaccion.frx":08CA
         Top             =   120
         Width           =   480
      End
   End
   Begin ActiveTabs.SSActiveTabs sTabTransaccion 
      Height          =   7005
      Left            =   180
      TabIndex        =   5
      Top             =   1620
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   12356
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tabs            =   "frmRegistrarTransaccion.frx":150E
      Begin ActiveTabs.SSActiveTabPanel sPabelLinea 
         Height          =   6615
         Left            =   -99969
         TabIndex        =   6
         Top             =   360
         Width           =   14235
         _ExtentX        =   25109
         _ExtentY        =   11668
         _Version        =   131083
         TabGuid         =   "frmRegistrarTransaccion.frx":15C1
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
            Left            =   12810
            Picture         =   "frmRegistrarTransaccion.frx":15E9
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
            Top             =   210
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
            Left            =   13410
            Picture         =   "frmRegistrarTransaccion.frx":22B3
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
            Top             =   210
            Width           =   555
         End
         Begin TrueOleDBGrid60.TDBGrid TDBG 
            Height          =   5670
            Left            =   300
            OleObjectBlob   =   "frmRegistrarTransaccion.frx":2F7D
            TabIndex        =   7
            Top             =   840
            Width           =   13725
         End
      End
      Begin ActiveTabs.SSActiveTabPanel sPanelTransaccion 
         Height          =   6615
         Left            =   -99969
         TabIndex        =   15
         Top             =   360
         Width           =   14235
         _ExtentX        =   25109
         _ExtentY        =   11668
         _Version        =   131083
         TabGuid         =   "frmRegistrarTransaccion.frx":BB42
         Begin VB.Frame Frame3 
            Height          =   2475
            Left            =   1230
            TabIndex        =   38
            Top             =   3330
            Width           =   11445
            Begin VB.TextBox txtPrecioDolar 
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
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   2310
               TabIndex        =   52
               Top             =   1200
               Width           =   1905
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
               Left            =   10590
               Picture         =   "frmRegistrarTransaccion.frx":BB6A
               Style           =   1  'Graphical
               TabIndex        =   49
               ToolTipText     =   "Agrega el item con los datos digitados..."
               Top             =   1410
               Width           =   555
            End
            Begin VB.CheckBox chkAutoSugiereLotes 
               Caption         =   "Auto Sugiere Lotes"
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
               Left            =   8220
               TabIndex        =   43
               Top             =   1410
               Width           =   2145
            End
            Begin VB.TextBox txtCostoDolar 
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
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   2310
               TabIndex        =   41
               Top             =   780
               Width           =   1905
            End
            Begin VB.CommandButton cmdDelLote 
               Height          =   320
               Left            =   3840
               Picture         =   "frmRegistrarTransaccion.frx":C834
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   1680
               Width           =   300
            End
            Begin VB.TextBox txtDescrLote 
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
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   4230
               TabIndex        =   48
               Top             =   1710
               Width           =   6135
            End
            Begin VB.CommandButton cmdLote 
               Height          =   320
               Left            =   3480
               Picture         =   "frmRegistrarTransaccion.frx":E4FE
               Style           =   1  'Graphical
               TabIndex        =   46
               Top             =   1680
               Width           =   300
            End
            Begin VB.TextBox txtLote 
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
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   2310
               TabIndex        =   45
               Top             =   1710
               Width           =   1095
            End
            Begin VB.TextBox txtCantidad 
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
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   2310
               TabIndex        =   39
               Top             =   270
               Width           =   1905
            End
            Begin VB.Label Label12 
               Caption         =   "Precio Dolar:"
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
               Height          =   300
               Left            =   540
               TabIndex        =   53
               Top             =   1260
               Width           =   1215
            End
            Begin VB.Label Label9 
               Caption         =   "Lote:"
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
               Height          =   300
               Left            =   540
               TabIndex        =   44
               Top             =   1680
               Width           =   735
            End
            Begin VB.Label Label10 
               Caption         =   "Cantidad:"
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
               Height          =   300
               Left            =   540
               TabIndex        =   40
               Top             =   330
               Width           =   1095
            End
            Begin VB.Label Label11 
               Caption         =   "Costo Dolar:"
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
               Height          =   300
               Left            =   540
               TabIndex        =   42
               Top             =   840
               Width           =   1095
            End
         End
         Begin VB.Frame Frame2 
            Height          =   915
            Left            =   1230
            TabIndex        =   32
            Top             =   2370
            Width           =   11445
            Begin VB.CommandButton cmdDelArticulo 
               Height          =   320
               Left            =   3900
               Picture         =   "frmRegistrarTransaccion.frx":E840
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   360
               Width           =   300
            End
            Begin VB.TextBox txtArticulo 
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
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   2310
               TabIndex        =   33
               Top             =   390
               Width           =   1095
            End
            Begin VB.TextBox txtDescrArticulo 
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
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   4335
               TabIndex        =   36
               Top             =   360
               Width           =   6675
            End
            Begin VB.CommandButton cmdArticulo 
               Height          =   320
               Left            =   3510
               Picture         =   "frmRegistrarTransaccion.frx":1050A
               Style           =   1  'Graphical
               TabIndex        =   34
               Top             =   360
               Width           =   300
            End
            Begin VB.Label Label8 
               Caption         =   "Articulo:"
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
               Height          =   300
               Left            =   570
               TabIndex        =   37
               Top             =   390
               Width           =   1005
            End
         End
         Begin VB.Frame Frame1 
            Height          =   1980
            Left            =   1230
            TabIndex        =   16
            Top             =   300
            Width           =   11460
            Begin VB.CommandButton cmdDelBodegaOrigen 
               Height          =   320
               Left            =   3900
               Picture         =   "frmRegistrarTransaccion.frx":1084C
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   855
               Width           =   300
            End
            Begin VB.CommandButton cmdBodegaDestino 
               Height          =   320
               Left            =   3510
               Picture         =   "frmRegistrarTransaccion.frx":12516
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   1305
               Width           =   300
            End
            Begin VB.TextBox txtDescrBodegaDestino 
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
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   4335
               TabIndex        =   31
               Top             =   1320
               Width           =   6675
            End
            Begin VB.TextBox txtBodegaDestino 
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
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   2310
               TabIndex        =   27
               Top             =   1320
               Width           =   1095
            End
            Begin VB.CommandButton cmdDelBodegaDestino 
               Height          =   320
               Left            =   3900
               Picture         =   "frmRegistrarTransaccion.frx":12858
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   1305
               Width           =   300
            End
            Begin VB.CommandButton cmdBodegaOrigen 
               Height          =   320
               Left            =   3510
               Picture         =   "frmRegistrarTransaccion.frx":14522
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   855
               Width           =   300
            End
            Begin VB.TextBox txtDescrBodegaOrigen 
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
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   4335
               TabIndex        =   26
               Top             =   870
               Width           =   6675
            End
            Begin VB.TextBox txtBodegaOrigen 
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
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   2310
               TabIndex        =   22
               Top             =   855
               Width           =   1095
            End
            Begin VB.CommandButton cmdTipoTransaccion 
               Height          =   320
               Left            =   3510
               Picture         =   "frmRegistrarTransaccion.frx":14864
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   420
               Width           =   300
            End
            Begin VB.TextBox txtDescrTipoTransaccion 
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
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   4335
               TabIndex        =   21
               Top             =   420
               Width           =   6675
            End
            Begin VB.TextBox txtTipoTransaccion 
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
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   2310
               TabIndex        =   18
               Top             =   420
               Width           =   1095
            End
            Begin VB.CommandButton cmdDelTipoTransaccion 
               Height          =   320
               Left            =   3900
               Picture         =   "frmRegistrarTransaccion.frx":14BA6
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   405
               Width           =   300
            End
            Begin VB.Label lblBodegaDestino 
               Caption         =   "Bodega Destino:"
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
               Height          =   300
               Left            =   555
               TabIndex        =   29
               Top             =   1320
               Width           =   1635
            End
            Begin VB.Label Label6 
               Caption         =   "Bodega Origen:"
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
               Height          =   300
               Left            =   555
               TabIndex        =   25
               Top             =   870
               Width           =   1635
            End
            Begin VB.Label Label5 
               Caption         =   "Tipo Transacción:"
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
               Height          =   300
               Left            =   540
               TabIndex        =   17
               Top             =   420
               Width           =   1635
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel sTabPanelDocumento 
         Height          =   6615
         Left            =   30
         TabIndex        =   8
         Top             =   360
         Width           =   14235
         _ExtentX        =   25109
         _ExtentY        =   11668
         _Version        =   131083
         TabGuid         =   "frmRegistrarTransaccion.frx":16870
         Begin VB.TextBox txtUsuario 
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
            Height          =   315
            Left            =   1350
            TabIndex        =   13
            Top             =   3330
            Width           =   1815
         End
         Begin RichTextLib.RichTextBox txtConcepto 
            Height          =   1785
            Left            =   540
            TabIndex        =   12
            Top             =   1110
            Width           =   13035
            _ExtentX        =   22992
            _ExtentY        =   3149
            _Version        =   393217
            ScrollBars      =   2
            TextRTF         =   $"frmRegistrarTransaccion.frx":16898
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
         Begin MSComCtl2.DTPicker dtpFecha 
            Height          =   315
            Left            =   1140
            TabIndex        =   9
            Top             =   240
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   556
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
            Format          =   97320961
            CurrentDate     =   41095
         End
         Begin VB.Label Label7 
            Caption         =   "Usuario:"
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
            Height          =   300
            Left            =   600
            TabIndex        =   14
            Top             =   3360
            Width           =   795
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha:"
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
            Height          =   300
            Left            =   510
            TabIndex        =   10
            Top             =   270
            Width           =   915
         End
         Begin VB.Label Label4 
            Caption         =   "Concepto:"
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
            Height          =   300
            Left            =   540
            TabIndex        =   11
            Top             =   870
            Width           =   915
         End
      End
   End
   Begin VB.TextBox txtIDDocumento 
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
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   2340
      TabIndex        =   1
      Top             =   1050
      Width           =   2445
   End
   Begin VB.TextBox txtPaquete 
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
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   1290
      TabIndex        =   0
      Top             =   1050
      Width           =   975
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
      Left            =   14490
      Picture         =   "frmRegistrarTransaccion.frx":16913
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Deshacer / Cancelar"
      Top             =   2580
      Width           =   555
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   555
      Left            =   14490
      Picture         =   "frmRegistrarTransaccion.frx":175DD
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
      Top             =   1950
      Width           =   555
   End
   Begin Inventario.CtlLiner CtlLiner 
      Height          =   30
      Left            =   0
      TabIndex        =   59
      Top             =   750
      Width           =   17925
      _ExtentX        =   31618
      _ExtentY        =   53
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Documento:"
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
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1020
   End
   Begin VB.Label lblTransaccion 
      BackStyle       =   0  'Transparent
      Caption         =   "Transacción:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   12540
      TabIndex        =   4
      Top             =   930
      Width           =   2640
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transacción:"
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
      Height          =   300
      Left            =   11355
      TabIndex        =   3
      Top             =   930
      Width           =   1170
   End
End
Attribute VB_Name = "frmRegistrarTransaccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Type typDatosProductos
     CostoUltLocal As Double
     CostoUltDolar As Double
     CostoPromLocal As Double
     CostoPromDolar As Double
 End Type
 

Dim rst As ADODB.Recordset
Dim rstLS As ADODB.Recordset
Dim Accion As TypAccion
Public gsFormCaption As String
Public gsTitle As String
Public gsIDTipoTransaccion As Integer
Dim sMensajeError As String
Dim bIsAutoSugiereLotes As Boolean
Dim dTotalSugeridoLotes As Double
Dim iFactor As Integer
Dim sPaquete As String
Dim cNaturaleza As String * 1
Dim iTipoTransaccion As Integer
Private rstTmpMovimiento As ADODB.Recordset
Dim rstLote As ADODB.Recordset

Dim gTrans As Boolean ' se dispara si hubo error en medio de la transacción
Dim gBeginTransNoEnd As Boolean ' Indica si hubo un begin sin rollback o commit
Dim gTasaCambio As Double
Dim tDatosDelProducto As typDatosProductos

Public scontrol As String


Private Sub HabilitarBotones()
    Select Case Accion
        Case TypAccion.Add
            cmdSave.Enabled = True
            cmdUndo.Enabled = False
            cmdEliminar.Enabled = True
            cmdAdd.Enabled = True
            cmdEditItem.Enabled = True
        Case TypAccion.Edit
            cmdSave.Enabled = False
            cmdUndo.Enabled = True
            cmdEliminar.Enabled = False
            cmdAdd.Enabled = True
            cmdEditItem.Enabled = False
        Case TypAccion.View
            cmdSave.Enabled = False
            cmdUndo.Enabled = False
            cmdEliminar.Enabled = False
            cmdAdd.Enabled = False
            cmdEditItem.Enabled = False
    End Select
    ActivarAccionesByTransacciones
End Sub

Public Sub HabilitarControles()
    Select Case Accion
        Case TypAccion.Add
            dtpFecha.value = Format(Now, "YYYY/MM/DD")
            txtTipoTransaccion.Text = ""
            fmtTextbox txtTipoTransaccion, "O"
            Me.txtDescrTipoTransaccion.Text = ""
            fmtTextbox Me.txtDescrTipoTransaccion, "R"
            
            txtBodegaDestino.Text = ""
            fmtTextbox txtBodegaDestino, "R"
            txtDescrBodegaDestino.Text = ""
            fmtTextbox txtDescrBodegaDestino, "R"
            
            txtBodegaOrigen.Text = ""
            fmtTextbox txtBodegaOrigen, "O"
            txtDescrBodegaOrigen.Text = ""
            fmtTextbox txtDescrBodegaOrigen, "R"
            
            fmtTextbox txtArticulo, "O"
            txtArticulo.Text = ""
            fmtTextbox Me.txtDescrArticulo, "R"
            txtDescrArticulo.Text = ""
            
            txtLote.Text = ""
            fmtTextbox txtLote, "R"
            txtDescrLote.Text = ""
            fmtTextbox txtDescrLote, "R"
            
            txtCantidad.Text = ""
            fmtTextbox txtCantidad, "O"
            
            txtCostoDolar.Text = ""
            fmtTextbox txtCostoDolar, "R"
            
            txtPrecioDolar.Text = ""
            fmtTextbox txtPrecioDolar, "R"
            
            Me.cmdTipoTransaccion.Enabled = True
            Me.cmdDelTipoTransaccion.Enabled = True
            Me.cmdBodegaOrigen.Enabled = True
            Me.cmdDelBodegaOrigen.Enabled = True
            Me.cmdBodegaDestino.Enabled = False
            Me.cmdDelBodegaDestino.Enabled = False
            Me.cmdArticulo.Enabled = True
            Me.cmdDelArticulo.Enabled = True
            Me.cmdLote.Enabled = False
            Me.cmdDelLote.Enabled = True
            Me.TDBG.Enabled = True
            
        Case TypAccion.Edit
            dtpFecha.value = Format(Now, "YYYY/MM/DD")
          
            fmtTextbox txtTipoTransaccion, "R"
            fmtTextbox Me.txtDescrTipoTransaccion, "R"
            
            fmtTextbox txtBodegaDestino, "R"
            fmtTextbox txtDescrBodegaDestino, "R"
            
            fmtTextbox txtBodegaOrigen, "R"
            fmtTextbox txtDescrBodegaOrigen, "R"
            
            fmtTextbox txtArticulo, "R"
            fmtTextbox Me.txtDescrArticulo, "R"
                        
            
            fmtTextbox txtLote, "R"
            fmtTextbox txtDescrLote, "R"
            
            fmtTextbox txtCantidad, "O"
            
            fmtTextbox txtCostoDolar, "O"
                  
            fmtTextbox txtPrecioDolar, "O"
            
            Me.cmdTipoTransaccion.Enabled = False
            Me.cmdDelTipoTransaccion.Enabled = False
            Me.cmdBodegaOrigen.Enabled = False
            Me.cmdDelBodegaOrigen.Enabled = False
            Me.cmdBodegaDestino.Enabled = False
            Me.cmdDelBodegaDestino.Enabled = False
            Me.cmdArticulo.Enabled = False
            Me.cmdDelArticulo.Enabled = False
            Me.cmdLote.Enabled = False
            Me.cmdDelLote.Enabled = False
            Me.TDBG.Enabled = False
            
        Case TypAccion.View
            dtpFecha.value = Format(Now, "YYYY/MM/DD")
            txtTipoTransaccion.Text = ""
            fmtTextbox txtTipoTransaccion, "R"
            txtBodegaDestino.Text = ""
            fmtTextbox txtBodegaDestino, "R"
            txtDescrBodegaDestino.Text = ""
            fmtTextbox txtDescrBodegaDestino, "R"
            txtBodegaOrigen.Text = ""
            fmtTextbox txtBodegaDestino, "R"
            txtDescrBodegaOrigen.Text = ""
            fmtTextbox txtArticulo, "R"
            txtArticulo.Text = ""
            fmtTextbox txtDescrBodegaOrigen, "R"
            txtDescrBodegaOrigen.Text = True
            txtArticulo.Text = ""
            fmtTextbox txtArticulo, "R"
            txtLote.Text = ""
            fmtTextbox txtLote, "R"
            txtCantidad.Text = ""
            fmtTextbox txtCantidad, "R"
            fmtTextbox txtCostoDolar, "R"
            fmtTextbox txtPrecioDolar, "R"
            txtCostoDolar.Text = ""
            
            Me.cmdTipoTransaccion.Enabled = False
            Me.cmdDelTipoTransaccion.Enabled = False
            Me.cmdBodegaOrigen.Enabled = False
            Me.cmdDelBodegaOrigen.Enabled = False
            Me.cmdBodegaDestino.Enabled = False
            Me.cmdDelBodegaDestino.Enabled = False
            Me.cmdArticulo.Enabled = False
            Me.cmdDelArticulo.Enabled = False
            Me.cmdLote.Enabled = False
            Me.cmdDelLote.Enabled = False
           
            Me.TDBG.Enabled = True
    End Select
End Sub

Private Function ValCtrls() As Boolean
    Dim Valida As Boolean
    Valida = True
    If (Me.txtConcepto.Text = "") Then
        sMensajeError = "Por favor ingrese en el concepto del documento..."
        Valida = False
    ElseIf (Me.txtDescrTipoTransaccion.Text = "") Then
        sMensajeError = "Por favor seleccione el tipo de Transacción..."
        Valida = False
    ElseIf (Me.txtDescrBodegaOrigen.Text = "") Then
        sMensajeError = "Por favor seleccione la bodega Origen"
        Valida = False
    ElseIf (Me.txtDescrBodegaDestino.Text = "") And (Me.gsIDTipoTransaccion = 4) Then 'Cuando la transaccion implica un traslado
        sMensajeError = "Por favor seleccione la bodega Destino"
        Valida = False
    ElseIf (Me.txtDescrArticulo.Text = "") Then
        sMensajeError = "Por favor seleccione producto"
        Valida = False
    ElseIf (Me.txtDescrLote.Text = "" And bIsAutoSugiereLotes = False) Then
        sMensajeError = "Por favor seleccione el lote del producto"
        Valida = False
    ElseIf (Me.txtCantidad.Text = "") Then
        sMensajeError = "Por favor digite la cantidad del producto"
        Valida = False
    ElseIf (Me.txtCostoDolar.Text = "") And (iTipoTransaccion = 1 Or iTipoTransaccion = 3 Or iTipoTransaccion = 7 Or iTipoTransaccion = 11) Then
        sMensajeError = "Por favor seleccione el Costo Dolar"
        Valida = False
    ElseIf (Me.txtPrecioDolar.Text = "") And (iTipoTransaccion = 2) Then
        sMensajeError = "Por favor seleccione el Precio del producto"
        Valida = False
    End If
    ValCtrls = Valida
End Function

Private Sub chkAutoSugiereLotes_Click()
      HabilitarAutoSugerirLotes Me.chkAutoSugiereLotes.value
End Sub

Private Sub HabilitarAutoSugerirLotes(IsAutoSugiereLotes As Boolean)
    If IsAutoSugiereLotes = True Then
        Me.txtLote.Enabled = False
        Me.txtDescrLote.Enabled = False
        Me.cmdLote.Enabled = False
        Me.cmdDelLote.Enabled = False
        bIsAutoSugiereLotes = True
    Else
        If (Accion = Add) Then
            Me.txtLote.Enabled = True
            Me.txtDescrLote.Enabled = True
            Me.cmdLote.Enabled = True
            Me.cmdDelLote.Enabled = True
        End If
        bIsAutoSugiereLotes = False
    End If
End Sub

Private Sub cmdAdd_Click()
 
   
  
    
    Dim lbok As Boolean
    If Not ValCtrls Then
        lbok = Mensaje("Revise sus datos por favor !!! " & sMensajeError, ICO_ERROR, False)
        Exit Sub
    End If
    
    
 
    
    If (Accion = Add) Then
        If (bIsAutoSugiereLotes = True) Then
            Set rstLS = New ADODB.Recordset
            If rstLS.State = adStateOpen Then rstLS.Close
            rstLS.ActiveConnection = gConet 'Asocia la conexión de trabajo
            rstLS.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
            rstLS.CursorLocation = adUseClient ' Cursor local al cliente
            rstLS.LockType = adLockOptimistic
        
            Dim frmAutosugiere As New frmAutoSugiereLotes
            frmAutosugiere.gsTitle = "Lotes Autosugeridos"
            frmAutosugiere.gsFormCaption = "Lotes"
            frmAutosugiere.gdCantidad = CDbl(txtCantidad.Text)
            frmAutosugiere.gsIDProducto = txtArticulo.Text
            frmAutosugiere.gsDescrProducto = Me.txtDescrArticulo.Text
            frmAutosugiere.gsIDBodega = Me.txtBodegaOrigen.Text
            frmAutosugiere.gsDescrBodega = Me.txtDescrBodegaOrigen.Text
            dTotalSugeridoLotes = frmAutosugiere.getTotalSugeridoporLote()
            If (dTotalSugeridoLotes < CDbl(Me.txtCantidad.Text)) Then
                lbok = Mensaje("No hay suficiente existencia del producto, la existencia actual es " & dTotalSugeridoLotes, ICO_ERROR, True)
                Set rstLS = Nothing
                Set frmAutosugiere = Nothing
            Else
                frmAutosugiere.Show vbModal
                Set rstLS = frmAutosugiere.grst
            End If
                    
            If rstLS Is Nothing Then Exit Sub
            
            If Not (rstLS.EOF And rstLS.BOF) Then
                rstLS.MoveFirst
                While Not rstLS.EOF
                    If ExiteRstKey(rstTmpMovimiento, "BODEGAOrigen=" & Me.txtBodegaOrigen.Text & " AND BodegaDestino =" & IIf(Me.txtBodegaDestino.Text = "", -1, Me.txtBodegaDestino.Text) & " AND IDPRODUCTO=" & Me.txtArticulo.Text & _
                                                " AND IDLOTE=" & rstLS!IdLote & " AND IDTIPO=" & Me.txtTipoTransaccion.Text) Then
                        lbok = Mensaje("Ya existe ese el registro en la transacción", ICO_ERROR, False)
                        Exit Sub
                    End If
                    Set rstLote = New ADODB.Recordset
                      rstLote.ActiveConnection = gConet
                    CargaDatosLotes rstLote, CInt(rstLS!IdLote)
                    ' Carga los datos del detalle de transacciones para ser grabados a la bd
                    rstTmpMovimiento.AddNew
                    rstTmpMovimiento!BodegaOrigen = Me.txtBodegaOrigen.Text
                    rstTmpMovimiento!DescrBodegaOrigen = Me.txtDescrBodegaOrigen.Text
                    'Pendiente: Aplicar los dos campos siguientes solo para traslados
                    rstTmpMovimiento!BodegaDestino = IIf(Me.txtBodegaDestino.Text = "", -1, Me.txtBodegaDestino.Text)
                    rstTmpMovimiento!DescrBodegaDestino = Me.txtDescrBodegaDestino.Text
                    rstTmpMovimiento!IdProducto = Me.txtArticulo.Text
                    rstTmpMovimiento!DescrProducto = Me.txtDescrArticulo.Text
                    rstTmpMovimiento!IdLote = rstLS!IdLote
                    rstTmpMovimiento!FechaVencimiento = rstLote!FechaVencimiento
                    rstTmpMovimiento!FechaFabricacion = rstLote!FechaFabricacion
                    rstTmpMovimiento!LoteInterno = rstLote!LoteInterno
                    rstTmpMovimiento!IdTipo = Me.txtTipoTransaccion.Text
                    rstTmpMovimiento!DESCRTipo = sPaquete
                    rstTmpMovimiento!Cantidad = rstLS!Cantidad
                    rstTmpMovimiento!Fecha = Format(Me.dtpFecha.value, "YYYY/MM/DD")
                    rstTmpMovimiento!CostoLocal = tDatosDelProducto.CostoPromLocal  'GetLastCostoProm(me.txtcodProducto.Text, "C")
                    rstTmpMovimiento!CostoDolar = tDatosDelProducto.CostoPromDolar 'CDbl(Me.txtCostoDolar.Text) 'GetLastCostoProm(txtCodProdAI.Text, "D")
                    rstTmpMovimiento!PrecioLocal = 0 '(rstTransDETAI!cant * rstTransDETAI!Costo)
                    rstTmpMovimiento!PrecioDolar = 0 '(rstTransDETAI!cant * rstTransDETAI!Costod)
                    rstTmpMovimiento!UserInsert = gsUser
                    rstTmpMovimiento.Update
                    rstTmpMovimiento.MoveFirst
                      
                    rstLS.MoveNext
                Wend
            End If
        Else
                
            If ExiteRstKey(rstTmpMovimiento, "BODEGAORigen=" & Me.txtBodegaOrigen.Text & " and  BodegaDestino=" & IIf(Me.txtBodegaDestino.Text = "", -1, Me.txtBodegaDestino.Text) & "  AND IDPRODUCTO=" & Me.txtArticulo.Text & _
                                          " AND IDLOTE=" & Me.txtLote.Text & " AND IDTIPO=" & Me.txtTipoTransaccion.Text) Then
              lbok = Mensaje("Ya existe ese el registro en la transacción", ICO_ERROR, False)
        
              Exit Sub
            End If
            Set rstLote = New ADODB.Recordset
              rstLote.ActiveConnection = gConet
            CargaDatosLotes rstLote, CInt(Trim(Me.txtLote.Text))
            ' Carga los datos del detalle de transacciones para ser grabados a la bd
            rstTmpMovimiento.AddNew
            rstTmpMovimiento!BodegaOrigen = Me.txtBodegaOrigen.Text
            rstTmpMovimiento!DescrBodegaOrigen = Me.txtDescrBodegaOrigen.Text
            rstTmpMovimiento!BodegaDestino = IIf(Me.txtBodegaDestino.Text = "", -1, Me.txtBodegaDestino.Text)
            rstTmpMovimiento!DescrBodegaDestino = Me.txtBodegaDestino.Text
            rstTmpMovimiento!IdProducto = Me.txtArticulo.Text
            rstTmpMovimiento!DescrProducto = Me.txtDescrArticulo.Text
            rstTmpMovimiento!IdLote = Me.txtLote.Text
            rstTmpMovimiento!FechaVencimiento = rstLote!FechaVencimiento
            rstTmpMovimiento!FechaFabricacion = rstLote!FechaFabricacion
            rstTmpMovimiento!LoteInterno = Me.txtDescrLote.Text
            rstTmpMovimiento!IdTipo = Me.txtTipoTransaccion.Text
            rstTmpMovimiento!DESCRTipo = sPaquete
            rstTmpMovimiento!Cantidad = Me.txtCantidad.Text
            rstTmpMovimiento!Fecha = Format(Me.dtpFecha.value, "YYYY/MM/DD")
            'Obtener el costo para las transacciones de ingreso de productos
            If (Me.txtTipoTransaccion.Text = 1 Or Me.txtTipoTransaccion.Text = 3 Or Me.txtTipoTransaccion.Text = 5 Or Me.txtTipoTransaccion.Text = 7) Then
                rstTmpMovimiento!CostoDolar = Me.txtCostoDolar.Text 'GetLastCostoProm(me.txtcodProducto.Text, "C")
                rstTmpMovimiento!CostoLocal = CDbl(Me.txtCostoDolar.Text) * gTasaCambio 'GetLastCostoProm(txtCodProdAI.Text, "D")
            Else
                rstTmpMovimiento!CostoLocal = tDatosDelProducto.CostoPromLocal 'GetLastCostoProm(me.txtcodProducto.Text, "C")
                rstTmpMovimiento!CostoDolar = tDatosDelProducto.CostoPromDolar
            End If
            'Obtener el precio para la transaccion tipo factura
            If (Me.txtTipoTransaccion.Text = 2) Then
                rstTmpMovimiento!PrecioLocal = CDbl(Me.txtPrecioDolar * gTasaCambio)  '(rstTransDETAI!cant * rstTransDETAI!Costo)
                rstTmpMovimiento!PrecioDolar = CDbl(Me.txtPrecioDolar) '(rstTransDETAI!cant * rstTransDETAI!Costod)
            Else
                rstTmpMovimiento!PrecioLocal = 0 '(rstTransDETAI!cant * rstTransDETAI!Costo)
                rstTmpMovimiento!PrecioDolar = 0 '(rstTransDETAI!cant * rstTransDETAI!Costod)
            End If
            rstTmpMovimiento!UserInsert = gsUser
            rstTmpMovimiento.Update
            rstTmpMovimiento.MoveFirst
        End If
    ElseIf (Accion = Edit) Then
        ' Actualiza el rst temporal
            rstTmpMovimiento!BodegaOrigen = Me.txtBodegaOrigen.Text
            rstTmpMovimiento!DescrBodegaOrigen = Me.txtDescrBodegaOrigen.Text
            rstTmpMovimiento!BodegaDestino = Me.txtBodegaDestino.Text
            rstTmpMovimiento!DescrBodegaDestino = IIf(Me.txtBodegaDestino.Text = "", -1, Me.txtBodegaDestino.Text)
            rstTmpMovimiento!IdProducto = Me.txtArticulo.Text
            rstTmpMovimiento!DescrProducto = Me.txtDescrArticulo.Text
            rstTmpMovimiento!IdLote = Me.txtLote.Text
            rstTmpMovimiento!LoteInterno = Me.txtDescrLote.Text
            rstTmpMovimiento!IdTipo = Me.txtTipoTransaccion.Text
            rstTmpMovimiento!DESCRTipo = sPaquete
'            rstTmpMovimiento!Cantidad = Me.txtCantidad.Text
            rstTmpMovimiento!Fecha = Format(Me.dtpFecha.value, "YYYY/MM/DD")
            'Obtener el costo para la todas las transacciones de ingreso de productos
            If (Me.txtTipoTransaccion.Text = 1 Or Me.txtTipoTransaccion.Text = 3 Or Me.txtTipoTransaccion.Text = 5 Or Me.txtTipoTransaccion.Text = 7) Then
                rstTmpMovimiento!CostoDolar = Me.txtCostoDolar.Text 'GetLastCostoProm(me.txtcodProducto.Text, "C")
                rstTmpMovimiento!CostoLocal = CDbl(Me.txtCostoDolar.Text) * gTasaCambio 'GetLastCostoProm(txtCodProdAI.Text, "D")
            Else
                rstTmpMovimiento!CostoLocal = tDatosDelProducto.CostoPromLocal 'GetLastCostoProm(me.txtcodProducto.Text, "C")
                rstTmpMovimiento!CostoDolar = tDatosDelProducto.CostoPromDolar
            End If
            'Obtener el precio para la transaccion tipo factura
            If (Me.txtTipoTransaccion.Text = 2) Then
                rstTmpMovimiento!PrecioLocal = CDbl(Me.txtPrecioDolar * gTasaCambio)  '(rstTransDETAI!cant * rstTransDETAI!Costo)
                rstTmpMovimiento!PrecioDolar = CDbl(Me.txtPrecioDolar) '(rstTransDETAI!cant * rstTransDETAI!Costod)
            Else
                rstTmpMovimiento!PrecioLocal = 0 '(rstTransDETAI!cant * rstTransDETAI!Costo)
                rstTmpMovimiento!PrecioDolar = 0 '(rstTransDETAI!cant * rstTransDETAI!Costod)
            End If
            rstTmpMovimiento!UserInsert = gsUser
            rstTmpMovimiento.Update
            
        End If
   
       Me.cmdSave.Enabled = True
    
      Set TDBG.DataSource = rstTmpMovimiento
      TDBG.ReBind
      
      If Not (rstTmpMovimiento.EOF And rstTmpMovimiento.BOF) Then Me.sTabTransaccion.Tabs(3).Enabled = True Else Me.sTabTransaccion.Tabs(3).Enabled = False
      
    Accion = Add
      'Me.dtgAjuste.Columns("Descr").FooterText = "Items de la transacción :     " & rstTransAI.RecordCount
HabilitarControles
HabilitarBotones
End Sub

Private Sub cmdArticulo_Click()
Dim frm As New frmBrowseCat
 Dim dicDatosProducto As Dictionary
 
    frm.gsCaptionfrm = "Artículos"
    frm.gsTablabrw = "vinvProducto"
    frm.gsCodigobrw = "IdProducto"
    frm.gbTypeCodeStr = True
    frm.gsDescrbrw = "Descr"
    frm.gbFiltra = False
    'frm.gsFiltro = "IdPaquete='" & Me.gsIDTipoTransaccion & "'"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      Me.txtArticulo.Text = frm.gsCodigobrw
      'Traer el costo promedio del producto
        If (getValueFieldsFromTable("invPRODUCTO", "CostoUltLocal,CostoUltDolar,CostoUltPromLocal,CostoUltPromDolar", "IdProducto=" & Me.txtArticulo.Text, dicDatosProducto) = True) Then
            tDatosDelProducto.CostoPromDolar = CDbl(dicDatosProducto("CostUlPromDolar"))
            tDatosDelProducto.CostoPromLocal = CDbl(dicDatosProducto("CostoUltPromLocal"))
            tDatosDelProducto.CostoUltDolar = CDbl((dicDatosProducto("CostoUltDolar")))
            tDatosDelProducto.CostoUltLocal = CDbl((dicDatosProducto("CostoUltLocal")))
            Select Case Me.txtTipoTransaccion.Text
                Case 1, 3, 7, 11:
                    Me.txtCostoDolar.Text = tDatosDelProducto.CostoPromDolar
            End Select
        End If
    End If
    
    If frm.gsDescrbrw <> "" Then
      Me.txtDescrArticulo.Text = frm.gsDescrbrw
      fmtTextbox Me.txtDescrArticulo, "R"
      Me.txtLote.Enabled = True
      Me.cmdLote.Enabled = True
    End If
End Sub

Private Sub cmdBodegaDestino_Click()
Dim frm As New frmBrowseCat
    
    frm.gsCaptionfrm = "Bodega Origen"
    frm.gsTablabrw = "invBODEGA"
    frm.gsCodigobrw = "IDBodega"
    frm.gbTypeCodeStr = True
    frm.gsDescrbrw = "Descr"
    frm.gbFiltra = False
    'frm.gsFiltro = "IdPaquete='" & Me.gsIDTipoTransaccion & "'"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      Me.txtBodegaDestino.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      Me.txtDescrBodegaDestino.Text = frm.gsDescrbrw
      fmtTextbox txtDescrBodegaDestino, "R"
    End If
End Sub

Private Sub cmdBodegaOrigen_Click()
  Dim frm As New frmBrowseCat
    
    frm.gsCaptionfrm = "Bodega Origen"
    frm.gsTablabrw = "invBODEGA"
    frm.gsCodigobrw = "IDBodega"
    frm.gbTypeCodeStr = True
    frm.gsDescrbrw = "Descr"
    frm.gbFiltra = False
    'frm.gsFiltro = "IdPaquete='" & Me.gsIDTipoTransaccion & "'"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      Me.txtBodegaOrigen.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      Me.txtDescrBodegaOrigen.Text = frm.gsDescrbrw
      fmtTextbox txtDescrBodegaOrigen, "R"
    End If
End Sub

Private Sub cmdDelBodegaOrigen_Click()
    Me.txtBodegaOrigen.Text = ""
    Me.txtDescrBodegaOrigen.Text = ""
End Sub

Private Sub cmdDelTipoTransaccion_Click()
    Me.txtTipoTransaccion.Text = ""
    Me.txtDescrTipoTransaccion.Text = ""
End Sub

Private Sub cmdEditItem_Click()
    Accion = Edit
    GetDataFromGridToControl
    HabilitarBotones
    HabilitarControles
    ActivarAccionesByTransacciones
End Sub

Private Sub GetDataFromGridToControl() 'EDITAR
'
    If Not (rstTmpMovimiento.EOF And rstTmpMovimiento.BOF) Then
        Me.txtTipoTransaccion.Text = rstTmpMovimiento("IDTipo").value
        Me.txtDescrTipoTransaccion.Text = rstTmpMovimiento("DescrTipo").value
        'Contemplar para traslados
        Me.txtBodegaOrigen.Text = rstTmpMovimiento("BodegaOrigen").value
        Me.txtDescrBodegaOrigen.Text = rstTmpMovimiento("DescrBodegaOrigen").value
        Me.txtBodegaDestino.Text = rstTmpMovimiento("BodegaDestino").value
        Me.txtDescrBodegaDestino.Text = rstTmpMovimiento("DescrBodegaDestino").value
        Me.txtCostoDolar.Text = rstTmpMovimiento("CostoDolar").value
        Me.txtPrecioDolar.Text = rstTmpMovimiento("PrecioDolar").value
        Me.txtArticulo.Text = rstTmpMovimiento("IDProducto").value
        Me.txtDescrArticulo.Text = rstTmpMovimiento("DescrProducto").value
        Me.txtLote.Text = rstTmpMovimiento("IDLote").value
        Me.txtDescrLote.Text = rstTmpMovimiento("LoteInterno").value
        Me.txtCantidad.Text = rstTmpMovimiento("Cantidad").value
        Me.sTabTransaccion.Tabs(2).Selected = True
        
    Else
      
        HabilitarControles
    End If

End Sub

Private Sub cmdEliminar_Click()
    Dim lbok As Boolean
    
    lbok = Mensaje("Esta seguro que desea eliminar el registro seleccionado?", ICO_INFORMACION, True)
    If (lbok) Then
        rstTmpMovimiento.Delete
        Accion = Add
        HabilitarBotones
        HabilitarControles
        TDBG.ReBind
    End If
End Sub

Private Sub cmdLote_Click()
    Dim frm As New frmBrowseCat
    
    frm.gsCaptionfrm = "Lote de Productos"
    If (Me.txtTipoTransaccion.Text = "1" Or Me.txtTipoTransaccion.Text = "3" Or Me.txtTipoTransaccion.Text = "7" Or Me.txtTipoTransaccion.Text = "11") Then
        frm.gsTablabrw = "invLOTE"
        frm.gsCodigobrw = "IdLote"
        frm.gbTypeCodeStr = True
        frm.gsDescrbrw = "LoteInterno"
        frm.gsDescrbrw = "LoteProveedor"
        frm.gsMuestraExtra = "SI"
        frm.gsFieldExtrabrw = "FechaVencimiento"
        frm.gbFiltra = False
        frm.gsNombrePantallaExtra = "frmMasterLotes"
        'frm.gsFiltro = "IdPaquete='" & Me.gsIDTipoTransaccion & "'"
    Else
        frm.gsCaptionfrm = "Lotes"
        frm.gsTablabrw = "vinvExistenciaLote"
        frm.gsCodigobrw = "IDLote"
        frm.gbTypeCodeStr = False
        frm.gsDescrbrw = "LoteProveedor"
        frm.gsMuestraExtra = "SI"
        frm.gsFieldExtrabrw = "FechaVencimiento"
        frm.gsMuestraExtra2 = "SI"
        frm.gsFieldExtrabrw2 = "Existencia"
        frm.gbFiltra = True
        frm.gsFiltro = "IDBodega=" & Me.txtBodegaOrigen.Text & " and IDProducto=" & Me.txtArticulo.Text & " and Existencia>0"
    End If
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      Me.txtLote.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      Me.txtDescrLote.Text = frm.gsDescrbrw
      fmtTextbox Me.txtDescrLote, "R"
    End If
End Sub

Private Sub cmdSave_Click()
  Dim lbok As Boolean
    'On Error GoTo errores
    
    If rstTmpMovimiento.RecordCount > 0 Then
      
'      If gTasaCambio = 0 Then
'        lbOk = Mensaje("La tasa de cambio es Cero llame a informática por favor ", ICO_ERROR, False)
'        Exit Sub
'      End If
      
      
      gConet.BeginTrans ' inicio aqui la transacción
      
      gBeginTransNoEnd = True
      Dim sDocumento As String
      sDocumento = invSaveCabeceraTransaccion()
        If sDocumento <> "" Then ' salva la cabecera
        SaveRstBatch rstTmpMovimiento, sDocumento ' salva el detalle que esta en batch
        If (gTrans = True) Then
            invMasterAcutalizaSaldosInventarioPaquete sDocumento, gsIDTipoTransaccion, Me.gsIDTipoTransaccion, gsUser
        End If
''        lbOk = Costo_Promedio_Batch(gRegistrosCODET, Format(CDate(Me.dtpFecha.value), "yyyy-mm-dd"), ParametrosGenerales.CodTranCompra)        ' Costo Promedio
''        If lbOk = False Then
''          If gBeginTransNoEnd Then
''            conn.RollbackTrans
''            gBeginTransNoEnd = False
''          End If
''          lbOk = Mensaje("Ha ocurrido un error en el cálculo del costo promedio, llame a informática", error, False)
''
''          Exit Sub
''        End If
        
  
'        lbOk = Update_Inventory(gRegistrosCODET, "COMP")
'        If lbOk = False Then
'          lbOk = Mensaje("Ha ocurrido un error en la actualización del inventario, llame a informática", error, False)
'          conn.RollbackTrans
'          Exit Sub
'        End If
        
'        If lbOk And SetFlgOk("TRANSACCION", "FLGOK", ParametrosGenerales.CodTranCompra, Str(lCorrelativo)) Then
'          '----------- Progress bar
'            ProgressBar1.Value = 100
'            lblProgress.Caption = "Fin"
'            lblProgress.Refresh
'          '----------- Progress bar
          If (gTrans = True) Then
            lbok = Mensaje("La transacción ha sido guardada exitosamente", ICO_OK, False)
       
          
         ' lblNoTra.Caption = ""
            Me.cmdAdd.Enabled = False
          
'
            Accion = View
            HabilitarBotones
            HabilitarControles
                   
            gConet.CommitTrans
            gBeginTransNoEnd = False
            'InicializaFormulario
            Exit Sub
          Else
            gConet.RollbackTrans
            gTrans = False
            gBeginTransNoEnd = False
          End If
        gBeginTransNoEnd = False
            
       
        Exit Sub
      End If
      
    
    End If
    gTrans = False
    
    If gBeginTransNoEnd Then
      gConet.RollbackTrans
      gBeginTransNoEnd = False
    End If
    lbok = Mensaje("Hubo un error en el proceso de salvado " & Chr(13) & err.Description, ICO_ERROR, False)
    'InicializaFormulario

End Sub

Private Function invSaveCabeceraTransaccion() As String
  
    Dim lbok As Boolean
    On Error GoTo errores
    lbok = False
    Dim sDocumento As String
    Dim rst As ADODB.Recordset

    GSSQL = "invInsertCabMovimientos " & Me.gsIDTipoTransaccion & ",'" & sDocumento & "','" & Format(Str(dtpFecha.value), "yyyymmdd") & _
                "','" & Me.txtConcepto.Text & "','" & sDocumento & "','" & gsUser & "','" & gsUser & "',1"
 Set rst = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    sDocumento = "" 'Indica que ocurrió un error
    sMensajeError = "Error en la búsqueda del descuento !!!" & err.Description
  Else  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    sDocumento = rst("Documento").value
  End If
  invSaveCabeceraTransaccion = sDocumento

    Exit Function
errores:
    gTrans = False
    invSaveCabeceraTransaccion = ""
    'gConet.RollbackTrans
    Exit Function
End Function

Private Sub cmdTipoTransaccion_Click()
    Dim frm As New frmBrowseCat
    
    frm.gsCaptionfrm = "Tipo Transacción"
    frm.gsTablabrw = "vinvPaqueteTipoMovimiento"
    frm.gsCodigobrw = "IDTipo"
    frm.gbTypeCodeStr = True
    frm.gsDescrbrw = "DESCRTipo"
    frm.gbFiltra = True
    frm.gsFiltro = "IdPaquete='" & Me.gsIDTipoTransaccion & "'"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      Me.txtTipoTransaccion.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      Me.txtDescrTipoTransaccion.Text = frm.gsDescrbrw
      fmtTextbox txtDescrTipoTransaccion, "R"
    End If
    ActivarAccionesByTransacciones
End Sub

Private Sub ActivarAccionesByTransacciones()
    Dim dicValue  As New Dictionary
    If (Me.txtDescrTipoTransaccion.Text = "") Then
        Exit Sub
    Else
        iTipoTransaccion = CInt(Me.txtTipoTransaccion.Text)
    End If
    

    
    If (getValueFieldsFromTable("invTIPOMOVIMIENTO", "Transaccion,Naturaleza,Factor", "IdTipo=" & Me.txtTipoTransaccion.Text, dicValue) = True) Then
        cNaturaleza = dicValue("Naturaleza")
        sPaquete = dicValue("Transaccion")
        iFactor = CInt(dicValue("Factor"))
    Else
        lbok = Mensaje("Ha ocurrido un error trantando de obtener información de la transacción", ICO_ERROR, False)
    End If
    
    '    1   COM Compra  E
    '2   FAC Facturación S
    '3   AJE Ajuste por Entrada  E
    '4   AJS Ajuste por Salida   S
    '5   BON Bonificación    E
    '6   PRS Préstamo Salida S
    '7   PRE Préstamo Entrada    E
    '8   CON Consumo S
    '9   TRS Traslado Salida S
    '10  TRE Traslado Entrada    E
    '11  FIE Ajuste Físico Entrada   E
    '12  FIS Ajuste Físico Salida    S
    Dim iTipoTran As Integer
    iTipoTran = Val(Me.txtTipoTransaccion.Text)
    
    Select Case iTipoTran
        Case 1, 3, 5, 7, 10, 11, 12: 'Todas las transacciones de ingreso
            'If (idtipotran = 1 Or idtipotran = 3 Or idtipotran = 7 Or iTipoTran = 11) Then 'Las transacciones que necesitan costo
            Select Case iTipoTran
                Case 1, 3, 7, 11: 'Todas las trasacciones de ingreso
                        If (Accion = Add) Then Me.txtCostoDolar.Text = ""
                        Me.txtCostoDolar.Enabled = True
                        fmtTextbox Me.txtCostoDolar, "O"
                        Me.txtBodegaDestino.Enabled = False
                        Me.txtBodegaDestino.Enabled = False
                        Me.cmdBodegaDestino.Enabled = False
                        Me.cmdDelBodegaDestino.Enabled = False
                        fmtTextbox Me.txtPrecioDolar, "R"
                        Me.TDBG.Columns(1).Visible = False
                Case 5, 10, 12: 'todas las trasacciones de salida
                    fmtTextbox Me.txtPrecioDolar, "R"
                    Me.txtCostoDolar.Text = ""
                    Me.txtCostoDolar.Enabled = False
                    fmtTextbox Me.txtCostoDolar, "R"
                    Me.txtBodegaDestino.Enabled = False
                    Me.txtBodegaDestino.Enabled = False
                    Me.cmdBodegaDestino.Enabled = False
                    Me.cmdDelBodegaDestino.Enabled = False
                    Me.TDBG.Columns(1).Visible = False
                    If (iTipoTran = 10) Then 'Los traslados
                        If (Accion = Add) Then
                            txtBodegaDestino.Text = ""
                            txtDescrBodegaDestino.Text = ""
                        End If
                        fmtTextbox txtBodegaDestino, "O"
                        
                        fmtTextbox txtDescrBodegaDestino, "R"
                        Me.txtBodegaDestino.Enabled = True
                        Me.cmdBodegaDestino.Enabled = True
                        Me.cmdDelBodegaDestino.Enabled = True
                        Me.cmdDelArticulo.Enabled = True
                        Me.TDBG.Columns(1).Visible = True
                    End If
            End Select
            Me.chkAutoSugiereLotes.Enabled = False
            Me.chkAutoSugiereLotes.value = vbUnchecked
            HabilitarAutoSugerirLotes Me.chkAutoSugiereLotes.value
        Case 2, 4, 6, 8, 9: 'Todas las transacciones de salida
            If (iTipoTran = 9) Then 'Si es traslado activar la bodega destino
                If (Accion = Add) Then
                    Me.txtBodegaDestino.Text = ""
                    Me.txtDescrBodegaDestino.Text = ""
                End If
                fmtTextbox Me.txtBodegaDestino, "O"
                Me.cmdBodegaDestino.Enabled = True
                Me.cmdDelBodegaDestino.Enabled = True
                Me.TDBG.Columns(1).Visible = True
            Else
                Me.txtBodegaDestino.Enabled = False
                Me.cmdDelBodegaDestino.Enabled = False
                Me.TDBG.Columns(1).Visible = False
            End If
            If (iTipoTran = 2) Then
                Me.txtPrecioDolar.Enabled = True
                fmtTextbox Me.txtPrecioDolar, "O"
                Me.txtPrecioDolar.Text = ""
                fmtTextbox Me.txtPrecioDolar, "O"
            End If
            Me.txtCostoDolar.Text = ""
            Me.txtCostoDolar.Enabled = False
            fmtTextbox Me.txtCostoDolar, "R"
            Me.chkAutoSugiereLotes.Enabled = True
            Me.chkAutoSugiereLotes.value = vbChecked
            HabilitarAutoSugerirLotes Me.chkAutoSugiereLotes.value
    End Select
End Sub

Private Sub cmdUndo_Click()
    GetDataFromGridToControl
    Accion = Add
    HabilitarBotones
    HabilitarControles
End Sub
' Public Function getDescrCatalogo(txtCodigo As TextBox, sFieldNameCode As String, sTableName As String, sFieldNameDescr As String, Optional bCodeChar As Boolean) As String
'Dim lbok As Boolean
'Dim sDescr As String
'Dim sValor As String
'lbok = False
'If txtCodigo.Text <> "" Then
'    If bCodeChar = True Then
'        sValor = "'" & txtCodigo.Text & "'"
'    Else
'        sValor = txtCodigo.Text
'    End If
'
'    sDescr = GetDescrCat(sFieldNameCode, sValor, sTableName, sFieldNameDescr)
'Else
'    sDescr = ""
'End If
'getDescrCatalogo = sDescr
'End Function

Private Sub SetTextBoxReadOnly()
    fmtTextbox txtUsuario, "R"
    fmtTextbox txtIDDocumento, "R"
    fmtTextbox txtPaquete, "R"
End Sub

Private Sub Form_Load()
    Set rstTmpMovimiento = New ADODB.Recordset
    If rstTmpMovimiento.State = adStateOpen Then rst.Close
    rstTmpMovimiento.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rstTmpMovimiento.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rstTmpMovimiento.CursorLocation = adUseClient ' Cursor local al cliente
    rstTmpMovimiento.LockType = adLockOptimistic
    
    
    gTasaCambio = 25.6
    
    Me.Caption = gsFormCaption
    Me.lbFormCaption = gsTitle
    'gTasaCambio = GetTasadeCambio(Format(Now, "YYYY/MM/DD"))
    PreparaRst ' Prepara los Recordsets
    Set Me.TDBG.DataSource = rstTmpMovimiento
    Me.TDBG.Refresh
    
    Dim DatosPaquete As New Dictionary
    getValueFieldsFromTable "invPAQUETE", "Paquete,Descr,Documento", "IDPaquete=" & Me.gsIDTipoTransaccion, DatosPaquete
    Me.txtPaquete.Text = DatosPaquete("Paquete")
    Me.txtIDDocumento.Text = DatosPaquete("Documento")
    Me.lblTransaccion.Caption = DatosPaquete("Descr")
    Me.txtUsuario.Text = gsUser
    Me.sTabTransaccion.Tabs(2).Enabled = False
    Me.sTabTransaccion.Tabs(3).Enabled = False
    SetTextBoxReadOnly
    Accion = Add
    HabilitarBotones
    HabilitarControles
    Me.chkAutoSugiereLotes.value = vbChecked
    HabilitarAutoSugerirLotes Me.chkAutoSugiereLotes.value
End Sub

Private Sub PreparaRst()
      ' preparacion del recordset fuente del grid de movimientos
      
      Set rstTmpMovimiento = New ADODB.Recordset
      If rstTmpMovimiento.State = adStateOpen Then rstTmpMovimiento.Close
      rstTmpMovimiento.ActiveConnection = gConet 'Asocia la conexión de trabajo
      rstTmpMovimiento.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
      rstTmpMovimiento.CursorLocation = adUseClient ' Cursor local al cliente
      rstTmpMovimiento.LockType = adLockOptimistic
                     
      If rstTmpMovimiento.State = adStateOpen Then rstTmpMovimiento.Close
      GSSQL = "invGetDetalleMovimiento " & gsIDTipoTransaccion & ",'SOFTFORYOU'"
      
      gTrans = True ' asume que NO va a haber un error en la transacción
      Set rstTmpMovimiento = GetRecordset(GSSQL) ' para el detalle
      

End Sub

Private Sub TDBG_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'    GetDataFromGridToControl
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (rstTmpMovimiento Is Nothing) Then Set rstTmpMovimiento = Nothing
End Sub

Public Sub CargaDatosLotes(rst As ADODB.Recordset, iIDLote As Integer)
    Dim lbok As Boolean
    'On Error GoTo error
    lbok = True
      GSSQL = "SELECT IDLote, LoteInterno, LoteProveedor, FechaVencimiento, FechaFabricacion"
    
      GSSQL = GSSQL & " FROM " & " dbo.invLOTE " 'Constuye la sentencia SQL
      GSSQL = GSSQL & " WHERE IDLote=" & iIDLote
      If rst.State = adStateOpen Then rst.Close
      rst.Open GSSQL, , adOpenKeyset, adLockOptimistic
    
    If (rst.BOF And rst.EOF) Then  'Si no es válido
        lbok = False  'Indica que no es válido
    End If
End Sub

Private Sub txtArticulo_KeyPress(KeyAscii As Integer)
    LoadDescrProducto txtArticulo, KeyAscii
    If (KeyAscii = 13) Then
        Me.txtCantidad.SetFocus
    End If
End Sub

Private Sub txtArticulo_LostFocus()
    LoadDescrProducto txtArticulo, 13
End Sub

Private Sub txtBodegaDestino_KeyPress(KeyAscii As Integer)
    LoadDescrBodegaDestino txtBodegaDestino, KeyAscii
    If (KeyAscii = 13) Then
        Me.txtArticulo.SetFocus
    End If
End Sub

Private Sub txtBodegaDestino_LostFocus()
    LoadDescrBodegaDestino txtBodegaDestino, 13
End Sub



Private Sub txtBodegaOrigen_KeyPress(KeyAscii As Integer)
    LoadDescrBodega txtBodegaOrigen, KeyAscii
    If (KeyAscii = 13) Then
        If Me.txtBodegaDestino.Enabled = True Then Me.txtBodegaDestino.SetFocus
    End If
End Sub

Private Sub txtBodegaOrigen_LostFocus()
    LoadDescrBodega txtBodegaOrigen, 13
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
 Dim lbok As Boolean
    If KeyAscii = vbKeyReturn Then
    
        If txtCantidad.Text <> "" Then
        
            If Not Val_TextboxNum(txtCantidad) Then
              lbok = Mensaje("Digite un valor correcto por favor ", ICO_ERROR, False)
              txtCantidad.SetFocus
              Exit Sub
            End If
            
                     
             ' txtTotal.Text = Format(txtCantidad.Text * txtCosto.Text, "###,###,##0.#0")
         
        Else
            lbok = Mensaje("Debe digitar la Cantidad", ICO_ERROR, False)
           ' txtCosto.Text = ""
            Exit Sub
        End If
        Me.txtCostoDolar.SetFocus
    End If
End Sub

Private Sub txtConcepto_Change()
    If (txtConcepto.Text <> "") Then
        Me.sTabTransaccion.Tabs(2).Enabled = True
    Else
        Me.sTabTransaccion.Tabs(2).Enabled = False
    End If
End Sub

Private Sub txtCostoDolar_KeyPress(KeyAscii As Integer)
 Dim lbok As Boolean
    If KeyAscii = vbKeyReturn Then
    
        If Me.txtCostoDolar.Text <> "" Then
        
            If Not Val_TextboxNum(txtCostoDolar) Then
              lbok = Mensaje("Digite un valor correcto por favor ", ICO_ERROR, False)
              txtCostoDolar.SetFocus
              Exit Sub
            End If
            
                     
             ' txtTotal.Text = Format(txtCantidad.Text * txtCosto.Text, "###,###,##0.#0")
         
        Else
            lbok = Mensaje("Debe digitar el Costo Dolar", ICO_ERROR, False)
           ' txtCosto.Text = ""
            Exit Sub
        End If
        Me.cmdAdd.SetFocus
    End If

End Sub

Public Function invSaveDetalleTransaccion(sIDPaquete As String, sIDBodega As String, sIDProducto As String, _
    sIDLote As String, sDocumento As String, sFecha As String, sIDTipo As String, STransaccion As String, _
    sNaturaleza As String, sCantidad As String, sCostoDolar As String, sCostoLocal As String, _
    sPrecioDolar As String, sPrecioLocal As String, sUserInsert As String, sUserUpdate As String) As Boolean
    Dim lbok As Boolean
   
    
    lbok = True
    
      GSSQL = ""
      GSSQL = gsCompania & ".invInsertMovimientos " & sIDPaquete & "," & sIDBodega & "," & sIDProducto & "," & sIDLote & ",'" & sDocumento & "','" & sFecha & "'," & sIDTipo & ",'"
      GSSQL = GSSQL & STransaccion & "','" & sNaturaleza & "'," & sCantidad & "," & sCostoDolar & "," & sCostoLocal & ","
      GSSQL = GSSQL & sPrecioDolar & "," & sPrecioLocal & ",'" & sUserInsert & "','" & sUserUpdate & "'"
        
     gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords   'Ejecuta la sentencia
    
        If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
          'gsOperacionError = "Eliminando el Beneficiado. " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & err.Description
          SetMsgError "Ocurrió un error insertando la transacción . ", err
          lbok = False
        End If
    
    invSaveDetalleTransaccion = lbok
    Exit Function
    

End Function



Public Function invMasterAcutalizaSaldosInventarioPaquete(sDocumento As String, sPaquete As Integer, sIDTipoTransaccion As String, _
    sUsuario As String) As Boolean
    Dim lbok As Boolean
   
    
    lbok = True
      GSSQL = ""
      GSSQL = gsCompania & ".invUpdateMasterExistenciaBodega '" & sDocumento & "'," & sPaquete & "," & sIDTipoTransaccion & ",'" & sUsuario & "'"
        
     gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords   'Ejecuta la sentencia
    
        If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
          'gsOperacionError = "Eliminando el Beneficiado. " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & err.Description
          SetMsgError "Ocurrió un error insertando la transacción . ", err
          lbok = False
        End If
    
    invMasterAcutalizaSaldosInventarioPaquete = lbok
    Exit Function
    

End Function

Private Sub SaveRstBatch(rst As ADODB.Recordset, sCodTra As String)
    On Error GoTo errores
    'Set lRegistros = New ADODB.Recordset  'Inicializa la variable de los registros
    'gConet.BeginTrans
    Dim bOk  As Boolean
    bOk = True
    
    If rst.RecordCount > 0 Then
      rst.MoveFirst
      While Not rst.EOF And bOk
     
            
            bOk = invSaveDetalleTransaccion(Me.gsIDTipoTransaccion, _
                                    rst.Fields("BodegaOrigen").value, _
                                    rst.Fields("IDProducto").value, _
                                    rst.Fields("IDLote").value, _
                                    sCodTra, _
                                    rst.Fields("Fecha").value, _
                                    rst.Fields("IdTipo").value, _
                                    rst.Fields("DESCRTipo").value, _
                                    cNaturaleza, _
                                    Abs(rst.Fields("Cantidad").value), _
                                    rst.Fields("CostoDolar").value, _
                                    rst.Fields("CostoLocal").value, _
                                    rst.Fields("PrecioDolar").value, _
                                    rst.Fields("PrecioLocal").value, _
                                    rst.Fields("UserInsert").value, _
                                    rst.Fields("UserInsert").value)
                                    
            If (rst.Fields("IdTipo").value = 10 Or rst.Fields("IdTipo").value = 9) And bOk Then
                If (rst.Fields("IdTipo").value = 10) Then
                    rst.Fields("IdTipo").value = 9
                    rst.Fields("DESCRTipo").value = "TRS"
                Else
                    rst.Fields("IdTipo").value = 10
                    rst.Fields("DESCRTipo").value = "TRE"
                End If
                
                bOk = invSaveDetalleTransaccion(Me.gsIDTipoTransaccion, _
                                    rst.Fields("BodegaDestino").value, _
                                    rst.Fields("IDProducto").value, _
                                    rst.Fields("IDLote").value, _
                                    sCodTra, _
                                    rst.Fields("Fecha").value, _
                                    rst.Fields("IdTipo").value, _
                                    rst.Fields("DESCRTipo").value, _
                                    cNaturaleza, _
                                    Abs(rst.Fields("Cantidad").value), _
                                    rst.Fields("CostoDolar").value, _
                                    rst.Fields("CostoLocal").value, _
                                    rst.Fields("PrecioDolar").value, _
                                    rst.Fields("PrecioLocal").value, _
                                    rst.Fields("UserInsert").value, _
                                    rst.Fields("UserInsert").value)
                
            End If
                                

            rst.MoveNext
      Wend
      rst.MoveFirst

    End If
    Exit Sub
errores:
    gTrans = False
    'gConet.RollbackTrans 'Descomentarie esto
End Sub

Public Function GetDescrCat(sfldCodCat As String, sValorCodigo As String, sTabla As String, sfldNameDescr As String, Optional bFiltroAdicional As Boolean = False, Optional sFiltroAdicional As String = "") As String
Dim sDescr As String
On Error GoTo error
  sDescr = ""
  GSSQL = "SELECT  " & sfldNameDescr & _
          " FROM " & gsCompania & "." & sTabla & _
          " WHERE " & sfldCodCat & " = " & sValorCodigo  'Constuye la sentencia SQL
  If bFiltroAdicional = True Then
    GSSQL = GSSQL & " AND " & sFiltroAdicional
  End If
    
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    sDescr = ""  'Indica que ocurrió un error
    gsOperacionError = "Error en la búsqueda del artículo !!!" & err.Description
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    sDescr = gRegistrosCmd(sfldNameDescr).value
  End If
  GetDescrCat = sDescr
  gRegistrosCmd.Close
  Exit Function
error:
  sDescr = ""
  gsOperacionError = "Ocurrió un error en la operación de búsqueda de la descripción " & err.Description
  Resume Next
End Function

Private Sub LoadDescrTipoTransaccion(ByRef txtCaja As TextBox, KeyAscii As Integer)
   Dim sDescr As String
    Dim lbok As Boolean
    If KeyAscii = vbKeyReturn Then
        sDescr = getDescrCatalogo(txtCaja, "IDTipo", "vinvPaqueteTipoMovimiento", "DescrTipo")
        If sDescr <> "" Then
            txtDescrTipoTransaccion.Text = sDescr
        Else
            lbok = Mensaje("La transacción no existe por favor revise", ICO_ERROR, False)
        End If
    End If
End Sub

Private Sub LoadDescrBodega(ByRef txtCaja As TextBox, KeyAscii As Integer)
   Dim sDescr As String
    Dim lbok As Boolean
    If KeyAscii = vbKeyReturn Then
        sDescr = getDescrCatalogo(txtCaja, "IDBodega", "invBodega", "Descr")
        If sDescr <> "" Then
            Me.txtDescrBodegaOrigen.Text = sDescr
        Else
            lbok = Mensaje("La Bodega no existe por favor revise", ICO_ERROR, False)
        End If
    End If
End Sub

Private Sub LoadDescrBodegaDestino(ByRef txtCaja As TextBox, KeyAscii As Integer)
   Dim sDescr As String
    Dim lbok As Boolean
    If KeyAscii = vbKeyReturn Then
        sDescr = getDescrCatalogo(txtCaja, "IDBodega", "invBodega", "Descr")
        If sDescr <> "" Then
            txtDescrBodegaDestino.Text = sDescr
        Else
            lbok = Mensaje("La Bodega no existe por favor revise", ICO_ERROR, False)
        End If
    End If
End Sub

Private Sub LoadDescrProducto(ByRef txtCaja As TextBox, KeyAscii As Integer)
   Dim sDescr As String
    Dim lbok As Boolean
    If KeyAscii = vbKeyReturn Then
        sDescr = getDescrCatalogo(txtCaja, "IDProducto", "invProducto", "Descr")
        If sDescr <> "" Then
            txtDescrArticulo.Text = sDescr
        Else
            lbok = Mensaje("El Producto no existe por favor revise", ICO_ERROR, False)
        End If
    End If
End Sub
'#revisar
Private Sub LoadDescrLote(ByRef txtCaja As TextBox, KeyAscii As Integer)
   Dim sDescr As String
    Dim lbok As Boolean
    If KeyAscii = vbKeyReturn Then
        sDescr = getDescrCatalogo(txtCaja, "IDProducto", "invProdcuto", "Descr")
        If sDescr <> "" Then
            txtDescrLote.Text = sDescr
        Else
            lbok = Mensaje("El Producto no existe por favor revise", ICO_ERROR, False)
        End If
    End If
End Sub



Private Sub txtLote_KeyPress(KeyAscii As Integer)
  LoadDescrLote txtLote, KeyAscii
    If (KeyAscii = 13) Then
        Me.cmdAdd.SetFocus
    End If
End Sub

Private Sub txtLote_LostFocus()
 LoadDescrLote txtLote, 13
End Sub

Private Sub txtPrecioDolar_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If txtPrecioDolar.Text <> "" Then

            If Not Val_TextboxNum(txtPrecioDolar) Then
              lbok = Mensaje("Digite un valor correcto por favor ", ICO_ADVERTENCIA, False)
              txtPrecioDolar.SetFocus
              Exit Sub
            End If
                
                     
             ' txtTotal.Text = Format(txtCantidad.Text * txtCosto.Text, "###,###,##0.#0")
         
        Else
            lbok = Mensaje("Debe digitar el Precio Dolar", ICO_ERROR, False)
           ' txtCosto.Text = ""
            Exit Sub
        End If
        Me.cmdAdd.SetFocus
        
    End If

End Sub

Private Sub txtTipoTransaccion_KeyPress(KeyAscii As Integer)
    LoadDescrTipoTransaccion txtTipoTransaccion, KeyAscii
    If (KeyAscii = 13) Then
        Me.txtBodegaOrigen.SetFocus
    End If
End Sub

Private Sub txtTipoTransaccion_LostFocus()
    LoadDescrTipoTransaccion txtTipoTransaccion, 13
End Sub

