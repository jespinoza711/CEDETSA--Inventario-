VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRegistrarTransaccion 
   BackColor       =   &H00F4D5BB&
   Caption         =   "Form1"
   ClientHeight    =   8445
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13470
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8445
   ScaleWidth      =   13470
   StartUpPosition =   2  'CenterScreen
   Begin ActiveTabs.SSActiveTabs sTabTransaccion 
      Height          =   7005
      Left            =   330
      TabIndex        =   6
      Top             =   1320
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   12356
      _Version        =   131083
      TabCount        =   3
      Tabs            =   "frmRegistrarTransaccion.frx":0000
      Begin ActiveTabs.SSActiveTabPanel sPabelLinea 
         Height          =   6615
         Left            =   -99969
         TabIndex        =   7
         Top             =   360
         Width           =   12060
         _ExtentX        =   21273
         _ExtentY        =   11668
         _Version        =   131083
         TabGuid         =   "frmRegistrarTransaccion.frx":00B3
         Begin TrueOleDBGrid60.TDBGrid TDBG 
            Height          =   5970
            Left            =   240
            OleObjectBlob   =   "frmRegistrarTransaccion.frx":00DB
            TabIndex        =   8
            Top             =   420
            Width           =   11595
         End
      End
      Begin ActiveTabs.SSActiveTabPanel sPanelTransaccion 
         Height          =   6615
         Left            =   -99969
         TabIndex        =   16
         Top             =   360
         Width           =   12060
         _ExtentX        =   21273
         _ExtentY        =   11668
         _Version        =   131083
         TabGuid         =   "frmRegistrarTransaccion.frx":8C90
         Begin VB.Frame Frame3 
            Height          =   2205
            Left            =   300
            TabIndex        =   39
            Top             =   3180
            Width           =   11445
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
               Picture         =   "frmRegistrarTransaccion.frx":8CB8
               Style           =   1  'Graphical
               TabIndex        =   50
               ToolTipText     =   "Agrega el item con los datos digitados..."
               Top             =   1110
               Width           =   555
            End
            Begin VB.CheckBox chkAutoSugiereLotes 
               Caption         =   "Auto Sugiere Lotes"
               Height          =   255
               Left            =   8520
               TabIndex        =   44
               Top             =   1050
               Width           =   1845
            End
            Begin VB.TextBox txtCostoDolar 
               Height          =   315
               Left            =   2310
               TabIndex        =   42
               Top             =   780
               Width           =   1905
            End
            Begin VB.CommandButton cmdDelLote 
               Height          =   320
               Left            =   3840
               Picture         =   "frmRegistrarTransaccion.frx":9982
               Style           =   1  'Graphical
               TabIndex        =   48
               Top             =   1320
               Width           =   300
            End
            Begin VB.TextBox txtDescrLote 
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
               Left            =   4230
               TabIndex        =   49
               Top             =   1350
               Width           =   6135
            End
            Begin VB.CommandButton cmdLote 
               Height          =   320
               Left            =   3480
               Picture         =   "frmRegistrarTransaccion.frx":B64C
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   1320
               Width           =   300
            End
            Begin VB.TextBox txtLote 
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
               Left            =   2310
               TabIndex        =   46
               Top             =   1350
               Width           =   1095
            End
            Begin VB.TextBox txtCantidad 
               Height          =   285
               Left            =   2310
               TabIndex        =   40
               Top             =   270
               Width           =   1905
            End
            Begin VB.Label Label9 
               Caption         =   "Lote:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H002F2F2F&
               Height          =   300
               Left            =   540
               TabIndex        =   45
               Top             =   1320
               Width           =   735
            End
            Begin VB.Label Label10 
               Caption         =   "Cantidad:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H002F2F2F&
               Height          =   300
               Left            =   540
               TabIndex        =   41
               Top             =   330
               Width           =   1095
            End
            Begin VB.Label Label11 
               Caption         =   "Costo Dolar:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H002F2F2F&
               Height          =   300
               Left            =   540
               TabIndex        =   43
               Top             =   840
               Width           =   1095
            End
         End
         Begin VB.Frame Frame2 
            Height          =   915
            Left            =   300
            TabIndex        =   33
            Top             =   2220
            Width           =   11445
            Begin VB.CommandButton cmdDelArticulo 
               Height          =   320
               Left            =   3900
               Picture         =   "frmRegistrarTransaccion.frx":B98E
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   360
               Width           =   300
            End
            Begin VB.TextBox txtArticulo 
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
               Left            =   2310
               TabIndex        =   34
               Top             =   360
               Width           =   1095
            End
            Begin VB.TextBox txtDescrArticulo 
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
               Left            =   4335
               TabIndex        =   37
               Top             =   360
               Width           =   6675
            End
            Begin VB.CommandButton cmdArticulo 
               Height          =   320
               Left            =   3510
               Picture         =   "frmRegistrarTransaccion.frx":D658
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   360
               Width           =   300
            End
            Begin VB.Label Label8 
               Caption         =   "Articulo:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H002F2F2F&
               Height          =   300
               Left            =   570
               TabIndex        =   38
               Top             =   390
               Width           =   1005
            End
         End
         Begin VB.Frame Frame1 
            Height          =   1980
            Left            =   300
            TabIndex        =   17
            Top             =   150
            Width           =   11460
            Begin VB.CommandButton cmdDelBodegaOrigen 
               Height          =   320
               Left            =   3900
               Picture         =   "frmRegistrarTransaccion.frx":D99A
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   855
               Width           =   300
            End
            Begin VB.CommandButton cmdBodegaDestino 
               Height          =   320
               Left            =   3510
               Picture         =   "frmRegistrarTransaccion.frx":F664
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   1305
               Width           =   300
            End
            Begin VB.TextBox txtDescrBodegaDestino 
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
               Left            =   4335
               TabIndex        =   32
               Top             =   1320
               Width           =   6675
            End
            Begin VB.TextBox txtBodegaDestino 
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
               Left            =   2310
               TabIndex        =   28
               Top             =   1305
               Width           =   1095
            End
            Begin VB.CommandButton cmdDelBodegaDestino 
               Height          =   320
               Left            =   3900
               Picture         =   "frmRegistrarTransaccion.frx":F9A6
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   1305
               Width           =   300
            End
            Begin VB.CommandButton cmdBodegaOrigen 
               Height          =   320
               Left            =   3510
               Picture         =   "frmRegistrarTransaccion.frx":11670
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   855
               Width           =   300
            End
            Begin VB.TextBox txtDescrBodegaOrigen 
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
               Left            =   4335
               TabIndex        =   27
               Top             =   870
               Width           =   6675
            End
            Begin VB.TextBox txtBodegaOrigen 
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
               Left            =   2310
               TabIndex        =   23
               Top             =   855
               Width           =   1095
            End
            Begin VB.CommandButton cmdTipoTransaccion 
               Height          =   320
               Left            =   3510
               Picture         =   "frmRegistrarTransaccion.frx":119B2
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   420
               Width           =   300
            End
            Begin VB.TextBox txtDescrTipoTransaccion 
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
               Left            =   4335
               TabIndex        =   22
               Top             =   420
               Width           =   6675
            End
            Begin VB.TextBox txtTipoTransaccion 
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
               Left            =   2310
               TabIndex        =   19
               Top             =   420
               Width           =   1095
            End
            Begin VB.CommandButton cmdDelTipoTransaccion 
               Height          =   320
               Left            =   3900
               Picture         =   "frmRegistrarTransaccion.frx":11CF4
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   405
               Width           =   300
            End
            Begin VB.Label lblBodegaDestino 
               Caption         =   "Bodega Destino:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H002F2F2F&
               Height          =   300
               Left            =   555
               TabIndex        =   30
               Top             =   1320
               Width           =   1635
            End
            Begin VB.Label Label6 
               Caption         =   "Bodega Origen:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H002F2F2F&
               Height          =   300
               Left            =   555
               TabIndex        =   26
               Top             =   870
               Width           =   1635
            End
            Begin VB.Label Label5 
               Caption         =   "Tipo Transacci�n:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H002F2F2F&
               Height          =   300
               Left            =   540
               TabIndex        =   18
               Top             =   420
               Width           =   1635
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel sTabPanelDocumento 
         Height          =   6615
         Left            =   30
         TabIndex        =   9
         Top             =   360
         Width           =   12060
         _ExtentX        =   21273
         _ExtentY        =   11668
         _Version        =   131083
         TabGuid         =   "frmRegistrarTransaccion.frx":139BE
         Begin VB.TextBox txtUsuario 
            Height          =   315
            Left            =   1350
            TabIndex        =   14
            Top             =   3330
            Width           =   1815
         End
         Begin RichTextLib.RichTextBox txtConcepto 
            Height          =   1785
            Left            =   540
            TabIndex        =   13
            Top             =   1110
            Width           =   11085
            _ExtentX        =   19553
            _ExtentY        =   3149
            _Version        =   393217
            ScrollBars      =   2
            TextRTF         =   $"frmRegistrarTransaccion.frx":139E6
         End
         Begin MSComCtl2.DTPicker dtpFecha 
            Height          =   315
            Left            =   1140
            TabIndex        =   10
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
            Format          =   57278465
            CurrentDate     =   41095
         End
         Begin VB.Label Label7 
            Caption         =   "Usuario:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H002F2F2F&
            Height          =   300
            Left            =   600
            TabIndex        =   15
            Top             =   3360
            Width           =   795
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H002F2F2F&
            Height          =   300
            Left            =   510
            TabIndex        =   11
            Top             =   270
            Width           =   915
         End
         Begin VB.Label Label4 
            Caption         =   "Concepto:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H002F2F2F&
            Height          =   300
            Left            =   540
            TabIndex        =   12
            Top             =   870
            Width           =   915
         End
      End
   End
   Begin VB.TextBox txtIDDocumento 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   2400
      TabIndex        =   2
      Top             =   780
      Width           =   2445
   End
   Begin VB.TextBox txtPaquete 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   1350
      TabIndex        =   1
      Top             =   780
      Width           =   975
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
      Left            =   12600
      Picture         =   "frmRegistrarTransaccion.frx":13A5D
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
      Top             =   3360
      Width           =   555
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
      Left            =   12600
      Picture         =   "frmRegistrarTransaccion.frx":14727
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Deshacer / Cancelar"
      Top             =   3990
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
      Left            =   12585
      Picture         =   "frmRegistrarTransaccion.frx":153F1
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
      Top             =   2730
      Width           =   555
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   555
      Left            =   12600
      Picture         =   "frmRegistrarTransaccion.frx":160BB
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Aplica y Guarda los datos de la transacci�n en Firme ..."
      Top             =   1710
      Width           =   555
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Documento:"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   390
      TabIndex        =   3
      Top             =   810
      Width           =   1020
   End
   Begin VB.Label lblTransaccion 
      BackStyle       =   0  'Transparent
      Caption         =   "Transacci�n:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   10590
      TabIndex        =   5
      Top             =   810
      Width           =   2640
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transacci�n:"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9405
      TabIndex        =   4
      Top             =   810
      Width           =   1170
   End
   Begin VB.Label lbFormCaption 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Titulo Catalogo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002F2F2F&
      Height          =   375
      Left            =   -210
      TabIndex        =   0
      Top             =   0
      Width           =   13650
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   -405
      Picture         =   "frmRegistrarTransaccion.frx":17D85
      Stretch         =   -1  'True
      Top             =   -300
      Width           =   13815
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
Dim bOrdenCodigo As Boolean
Dim bOrdenDescr As Boolean
Dim sCodSucursal As String
Dim sSoloActivo As String
Dim Accion As TypAccion
Public gsFormCaption As String
Public gsTitle As String
Public gsIDTipoTransaccion As Integer
Dim sMensajeError As String
Dim bIsAutoSugiereLotes As Boolean
Dim dTotalSugeridoLotes As Double
Dim iFactor As Integer
Dim sPaquete As String
Dim iPaquete As Integer
Dim cNaturaleza As String * 1
Dim iTipoTransaccion As Integer
Private rstTmpMovimiento As ADODB.Recordset
Dim rstLote As ADODB.Recordset
Dim lbAIenProceso As Boolean ' Indica si un ajuste est� en proceso

Dim gTrans As Boolean ' se dispara si hubo error en medio de la transacci�n
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
            
            Me.cmdTipoTransaccion.Enabled = True
            Me.cmdDelTipoTransaccion.Enabled = True
            Me.cmdBodegaOrigen.Enabled = True
            Me.cmdDelBodegaOrigen.Enabled = True
            Me.cmdBodegaDestino.Enabled = False
            Me.cmdDelBodegaDestino.Enabled = False
            Me.cmdArticulo.Enabled = True
            Me.cmdDelArticulo.Enabled = True
            Me.cmdLote.Enabled = True
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
            
            fmtTextbox txtCostoDolar, "R"
            
            Me.cmdTipoTransaccion.Enabled = True
            Me.cmdDelTipoTransaccion.Enabled = True
            Me.cmdBodegaOrigen.Enabled = True
            Me.cmdDelBodegaOrigen.Enabled = True
            Me.cmdBodegaDestino.Enabled = True
            Me.cmdDelBodegaDestino.Enabled = True
            Me.cmdArticulo.Enabled = True
            Me.cmdDelArticulo.Enabled = True
            Me.cmdLote.Enabled = True
            Me.cmdDelLote.Enabled = True
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
        sMensajeError = "Por favor seleccione el tipo de Transacci�n..."
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
        Me.txtLote.Enabled = True
        Me.txtDescrLote.Enabled = True
        Me.cmdLote.Enabled = True
        Me.cmdDelLote.Enabled = True
        bIsAutoSugiereLotes = False
    End If
End Sub

Private Sub cmdAdd_Click()
 
    Dim dicDatosProducto As Dictionary
  
    
    Dim lbOk As Boolean
    If Not ValCtrls Then
        lbOk = Mensaje("Revise sus datos por favor !!! " & sMensajeError, ICO_ERROR, False)
        Exit Sub
    End If
    
    
    If (getValueFieldsFromTable("invPRODUCTO", "CostoUltLocal,CostoUltDolar,CostoUltPromLocal,CostoUltPromDolar", "IdProducto=" & Me.txtArticulo.Text, dicDatosProducto) = True) Then
        tDatosDelProducto.CostoPromDolar = CDbl(dicDatosProducto("CostUlPromDolar"))
        tDatosDelProducto.CostoPromLocal = CDbl(dicDatosProducto("CostoUltPromLocal"))
        tDatosDelProducto.CostoUltDolar = CDbl((dicDatosProducto("CostoUltDolar")))
        tDatosDelProducto.CostoUltLocal = CDbl((dicDatosProducto("CostoUltLocal")))
    End If
    
    If (Accion = Add) Then
        If (bIsAutoSugiereLotes = True) Then
            Set rstLS = New ADODB.Recordset
            If rstLS.State = adStateOpen Then rstLS.Close
            rstLS.ActiveConnection = gConet 'Asocia la conexi�n de trabajo
            rstLS.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
            rstLS.CursorLocation = adUseClient ' Cursor local al cliente
            rstLS.LockType = adLockOptimistic
        
            Dim frmAutosugiere As New frmAutoSugiereLotes
            frmAutosugiere.gsTitle = "Lotes Autosugeridos"
            frmAutosugiere.gsFormCaption = "Lotes"
            frmAutosugiere.gdCantidad = CDbl(txtCantidad.Text)
            frmAutosugiere.gsIdProducto = txtArticulo.Text
            frmAutosugiere.gsDescrProducto = Me.txtDescrArticulo.Text
            frmAutosugiere.gsIDBodega = Me.txtBodegaOrigen.Text
            frmAutosugiere.gsDescrBodega = Me.txtDescrBodegaOrigen.Text
            dTotalSugeridoLotes = frmAutosugiere.GetTotalSugeridoLotes()
            If (dTotalSugeridoLotes < CDbl(Me.txtCantidad.Text)) Then
                lbOk = Mensaje("No hay suficiente existencia del producto, la existencia actual es " & dTotalSugeridoLotes, ICO_ERROR, True)
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
                        lbOk = Mensaje("Ya existe ese el registro en la transacci�n", ICO_ERROR, False)
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
                    rstTmpMovimiento!DESCRTipo = Me.txtDescrTipoTransaccion.Text
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
                
            If ExiteRstKey(rstTmpMovimiento, "BODEGAORigen=" & Me.txtBodegaOrigen.Text & " BodegaDestino=" & IIf(Me.txtBodegaDestino.Text = "", -1, Me.txtBodegaDestino.Text) & "  AND IDPRODUCTO=" & Me.txtArticulo.Text & _
                                          " AND IDLOTE=" & Me.txtLote.Text & " AND IDTIPO=" & Me.txtTipoTransaccion.Text) Then
              lbOk = Mensaje("Ya existe ese el registro en la transacci�n", ICO_ERROR, False)
        
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
            rstTmpMovimiento!DESCRTipo = Me.txtDescrTipoTransaccion.Text
            rstTmpMovimiento!Cantidad = Me.txtCantidad.Text
            rstTmpMovimiento!Fecha = Format(Me.dtpFecha.value, "YYYY/MM/DD")
            If (Me.txtTipoTransaccion.Text = 3 Or Me.txtTipoTransaccion.Text = 5 Or Me.txtTipoTransaccion.Text = 7) Then
                rstTmpMovimiento!CostoLocal = Me.txtCostoDolar.Text 'GetLastCostoProm(me.txtcodProducto.Text, "C")
                rstTmpMovimiento!CostoDolar = Me.txtCostoDolar.Text * gTasaCambio 'GetLastCostoProm(txtCodProdAI.Text, "D")
            Else
                rstTmpMovimiento!CostoLocal = tDatosDelProducto.CostoPromLocal 'GetLastCostoProm(me.txtcodProducto.Text, "C")
                rstTmpMovimiento!CostoDolar = tDatosDelProducto.CostoPromDolar
            End If
            rstTmpMovimiento!PrecioLocal = 0 '(rstTransDETAI!cant * rstTransDETAI!Costo)
            rstTmpMovimiento!PrecioDolar = 0 '(rstTransDETAI!cant * rstTransDETAI!Costod)
            rstTmpMovimiento!UserInsert = gsUser
            rstTmpMovimiento.Update
            rstTmpMovimiento.MoveFirst
        End If
    ElseIf (Accion = Edit) Then
        ' Actualiza el rst temporal
            rstTmpMovimiento!IdBodega = Me.txtBodegaOrigen.Text
            rstTmpMovimiento!DescrBodega = Me.txtDescrBodegaOrigen.Text
            rstTmpMovimiento!IdProducto = Me.txtArticulo.Text
            rstTmpMovimiento!DescrProducto = Me.txtDescrArticulo.Text
            rstTmpMovimiento!IdLote = Me.txtLote.Text
            rstTmpMovimiento!LoteInterno = Me.txtDescrLote.Text
            rstTmpMovimiento!IdTipo = Me.txtTipoTransaccion.Text
            rstTmpMovimiento!DESCRTipo = Me.txtDescrTipoTransaccion.Text
            rstTmpMovimiento!Cantidad = Me.txtCantidad.Text
            rstTmpMovimiento!Fecha = Format(Me.dtpFecha.value, "YYYY/MM/DD")
            rstTmpMovimiento!CostoLocal = tDatosDelProducto.CostoPromLocal 'GetLastCostoProm(me.txtcodProducto.Text, "C")
            rstTmpMovimiento!CostoDolar = tDatosDelProducto.CostoPromDolar 'GetLastCostoProm(txtCodProdAI.Text, "D")
            rstTmpMovimiento!PrecioLocal = 0 '(rstTransDETAI!cant * rstTransDETAI!Costo)
            rstTmpMovimiento!PrecioDolar = 0 '(rstTransDETAI!cant * rstTransDETAI!Costod)
            rstTmpMovimiento!UserInsert = gsUser
            rstTmpMovimiento.Update
            
        End If
      
        
    
    
    
       Me.cmdSave.Enabled = True
    
      Set TDBG.DataSource = rstTmpMovimiento
      TDBG.ReBind
      
      If Not (rstTmpMovimiento.EOF And rstTmpMovimiento.BOF) Then Me.sTabTransaccion.Tabs(3).Enabled = True Else Me.sTabTransaccion.Tabs(3).Enabled = False
      

      'Me.dtgAjuste.Columns("Descr").FooterText = "Items de la transacci�n :     " & rstTransAI.RecordCount
HabilitarControles
HabilitarBotones
End Sub

Private Sub cmdArticulo_Click()
Dim frm As New frmBrowseCat
    
    frm.gsCaptionfrm = "Art�culos"
    frm.gsTablabrw = "vinvProducto"
    frm.gsCodigobrw = "IdProducto"
    frm.gbTypeCodeStr = True
    frm.gsDescrbrw = "Descr"
    frm.gbFiltra = False
    'frm.gsFiltro = "IdPaquete='" & Me.gsIDTipoTransaccion & "'"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      Me.txtArticulo.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      Me.txtDescrArticulo.Text = frm.gsDescrbrw
      fmtTextbox Me.txtDescrArticulo, "R"
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
        Me.txtDescrBodegaDestino.Text = rstTmpMovimiento("DescrBodega").value
        Me.txtCostoDolar.Text = rstTmpMovimiento("CostoDolar").value
        Me.txtBodegaOrigen.Text = rstTmpMovimiento("IDBodega").value
        Me.txtDescrBodegaOrigen.Text = rstTmpMovimiento("DescrBodega").value
        Me.txtArticulo.Text = rstTmpMovimiento("IDProducto").value
        Me.txtDescrArticulo.Text = rstTmpMovimiento("DescrProducto").value
        Me.txtLote.Text = rstTmpMovimiento("IDLote").value
        Me.txtDescrLote.Text = rstTmpMovimiento("LoteInterno").value
        Me.txtCantidad.Text = rstTmpMovimiento("Cantidad").value
        Me.txtCostoDolar.Text = rstTmpMovimiento("CostoDolar").value
        
        
    Else
      
        HabilitarControles
    End If

End Sub

Private Sub cmdEliminar_Click()
    Dim lbOk As Boolean
    
    lbOk = Mensaje("Esta seguro que desea eliminar el registro seleccionado?", ICO_INFORMACION, True)
    If (lbOk) Then
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
    frm.gsTablabrw = "invLOTE"
    frm.gsCodigobrw = "IdLote"
    frm.gbTypeCodeStr = True
    frm.gsDescrbrw = "LoteInterno"
    frm.gbFiltra = False
    'frm.gsFiltro = "IdPaquete='" & Me.gsIDTipoTransaccion & "'"
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
  Dim lbOk As Boolean
    'On Error GoTo errores
    
    If rstTmpMovimiento.RecordCount > 0 Then
      
'      If gTasaCambio = 0 Then
'        lbOk = Mensaje("La tasa de cambio es Cero llame a inform�tica por favor ", ICO_ERROR, False)
'        Exit Sub
'      End If
      
      
      gConet.BeginTrans ' inicio aqui la transacci�n
      
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
''          lbOk = Mensaje("Ha ocurrido un error en el c�lculo del costo promedio, llame a inform�tica", error, False)
''
''          Exit Sub
''        End If
        
  
'        lbOk = Update_Inventory(gRegistrosCODET, "COMP")
'        If lbOk = False Then
'          lbOk = Mensaje("Ha ocurrido un error en la actualizaci�n del inventario, llame a inform�tica", error, False)
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
            lbOk = Mensaje("La transacci�n ha sido guardada exitosamente", ICO_OK, False)
       
          
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
    lbOk = Mensaje("Hubo un error en el proceso de salvado " & Chr(13) & err.Description, ICO_ERROR, False)
    'InicializaFormulario

End Sub

Private Function invSaveCabeceraTransaccion() As String
  
    Dim lbOk As Boolean
    On Error GoTo errores
    lbOk = False
    Dim sDocumento As String
    Dim rst As ADODB.Recordset

    GSSQL = "invInsertCabMovimientos " & Me.gsIDTipoTransaccion & ",'" & sDocumento & "','" & Format(Str(dtpFecha.value), "yyyymmdd") & _
                "','" & Me.txtConcepto.Text & "','" & sDocumento & "','" & gsUser & "','" & gsUser & "'"
 Set rst = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecuci�n
    sDocumento = "" 'Indica que ocurri� un error
    sMensajeError = "Error en la b�squeda del descuento !!!" & err.Description
  Else  'Si no es v�lido
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
    
    frm.gsCaptionfrm = "Tipo Transacci�n"
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
        lbOk = Mensaje("Ha ocurrido un error trantando de obtener informaci�n de la transacci�n", ICO_ERROR, False)
    End If
    
    '    1   COM Compra  E
    '2   FAC Facturaci�n S
    '3   AJE Ajuste por Entrada  E
    '4   AJS Ajuste por Salida   S
    '5   BON Bonificaci�n    E
    '6   PRS Pr�stamo Salida S
    '7   PRE Pr�stamo Entrada    E
    '8   CON Consumo S
    '9   TRS Traslado Salida S
    '10  TRE Traslado Entrada    E
    '11  FIE Ajuste F�sico Entrada   E
    '12  FIS Ajuste F�sico Salida    S
    Dim iTipoTran As Integer
    iTipoTran = Val(Me.txtTipoTransaccion.Text)
    
    Select Case iTipoTran
        Case 1, 3, 5, 7, 10, 11, 12: 'Todas las transacciones de ingreso
            'If (idtipotran = 1 Or idtipotran = 3 Or idtipotran = 7 Or iTipoTran = 11) Then 'Las transacciones que necesitan costo
            Select Case iTipoTran
                Case 1, 3, 7, 11: 'Todas las trasacciones de ingreso
                    Me.txtCostoDolar.Text = ""
                    Me.txtCostoDolar.Enabled = True
                    fmtTextbox Me.txtCostoDolar, "O"
                    Me.txtBodegaDestino.Enabled = False
                    Me.txtBodegaDestino.Enabled = False
                    Me.cmdBodegaDestino.Enabled = False
                    Me.cmdDelBodegaDestino.Enabled = False
                    Me.cmdDelArticulo.Enabled = False
                    Me.TDBG.Columns(1).Visible = False
                Case 5, 10, 12: 'todas las trasacciones de salida
                    Me.txtCostoDolar.Text = ""
                    Me.txtCostoDolar.Enabled = False
                    fmtTextbox Me.txtCostoDolar, "R"
                    Me.txtBodegaDestino.Enabled = False
                    Me.txtBodegaDestino.Enabled = False
                    Me.cmdBodegaDestino.Enabled = False
                    Me.cmdDelBodegaDestino.Enabled = False
                    Me.cmdDelArticulo.Enabled = False
                    Me.TDBG.Columns(1).Visible = False
                    If (iTipoTran = 10) Then 'Los traslados
                        txtBodegaDestino.Text = ""
                        fmtTextbox txtBodegaDestino, "O"
                        txtDescrBodegaDestino.Text = ""
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
                Me.txtBodegaDestino.Enabled = True
                Me.cmdDelBodegaDestino.Enabled = True
                Me.cmdDelArticulo.Enabled = True
                Me.TDBG.Columns(1).Visible = True
            Else '
                Me.txtBodegaDestino.Enabled = False
                Me.cmdDelBodegaDestino.Enabled = False
                Me.cmdDelArticulo.Enabled = False
                Me.TDBG.Columns(1).Visible = False
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
    rstTmpMovimiento.ActiveConnection = gConet 'Asocia la conexi�n de trabajo
    rstTmpMovimiento.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rstTmpMovimiento.CursorLocation = adUseClient ' Cursor local al cliente
    rstTmpMovimiento.LockType = adLockOptimistic
    
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
      rstTmpMovimiento.ActiveConnection = gConet 'Asocia la conexi�n de trabajo
      rstTmpMovimiento.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
      rstTmpMovimiento.CursorLocation = adUseClient ' Cursor local al cliente
      rstTmpMovimiento.LockType = adLockOptimistic
                     
      If rstTmpMovimiento.State = adStateOpen Then rstTmpMovimiento.Close
      GSSQL = "invGetDetalleMovimiento " & gsIDTipoTransaccion & ",'SOFTFORYOU'"
      
      gTrans = True ' asume que NO va a haber un error en la transacci�n
      Set rstTmpMovimiento = GetRecordset(GSSQL) ' para el detalle
      

End Sub

Private Sub TDBG_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'    GetDataFromGridToControl
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (rstTmpMovimiento Is Nothing) Then Set rstTmpMovimiento = Nothing
End Sub

Public Sub CargaDatosLotes(rst As ADODB.Recordset, iIDLote As Integer)
    Dim lbOk As Boolean
    'On Error GoTo error
    lbOk = True
      GSSQL = "SELECT IDLote, LoteInterno, LoteProveedor, FechaVencimiento, FechaFabricacion"
    
      GSSQL = GSSQL & " FROM " & " dbo.invLOTE " 'Constuye la sentencia SQL
      GSSQL = GSSQL & " WHERE IDLote=" & iIDLote
      If rst.State = adStateOpen Then rst.Close
      rst.Open GSSQL, , adOpenKeyset, adLockOptimistic
    
    If (rst.BOF And rst.EOF) Then  'Si no es v�lido
        lbOk = False  'Indica que no es v�lido
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
 Dim lbOk As Boolean
    If KeyAscii = vbKeyReturn Then
    
        If txtCantidad.Text <> "" Then
        
            If Not Val_TextboxNum(txtCantidad) Then
              lbOk = Mensaje("Digite un valor correcto por favor ", ICO_ERROR, False)
              txtCantidad.SetFocus
              Exit Sub
            End If
            
                     
             ' txtTotal.Text = Format(txtCantidad.Text * txtCosto.Text, "###,###,##0.#0")
         
        Else
            lbOk = Mensaje("Debe digitar la Cantidad", ICO_ERROR, False)
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
 Dim lbOk As Boolean
    If KeyAscii = vbKeyReturn Then
    
        If Me.txtCostoDolar.Text <> "" Then
        
            If Not Val_TextboxNum(txtCostoDolar) Then
              lbOk = Mensaje("Digite un valor correcto por favor ", ICO_ERROR, False)
              txtCostoDolar.SetFocus
              Exit Sub
            End If
            
                     
             ' txtTotal.Text = Format(txtCantidad.Text * txtCosto.Text, "###,###,##0.#0")
         
        Else
            lbOk = Mensaje("Debe digitar el Costo Dolar", ICO_ERROR, False)
           ' txtCosto.Text = ""
            Exit Sub
        End If
        Me.cmdAdd.SetFocus
    End If

End Sub

Public Function invSaveDetalleTransaccion(sIDPaquete As String, sIDBodega As String, sIDProducto As String, _
    sIDlote As String, sDocumento As String, sFecha As String, sIDTipo As String, STransaccion As String, _
    sNaturaleza As String, sCantidad As String, sCostoDolar As String, sCostoLocal As String, _
    sPrecioDolar As String, sPrecioLocal As String, sUserInsert As String, sUserUpdate As String) As Boolean
    Dim lbOk As Boolean
   
    
    lbOk = True
      GSSQL = ""
      GSSQL = gsCompania & ".invInsertMovimientos " & sIDPaquete & "," & sIDBodega & "," & sIDProducto & "," & sIDlote & ",'" & sDocumento & "','" & sFecha & "'," & sIDTipo & ",'"
      GSSQL = GSSQL & STransaccion & "','" & sNaturaleza & "'," & sCantidad & "," & sCostoDolar & "," & sCostoLocal & ","
      GSSQL = GSSQL & sPrecioDolar & "," & sPrecioLocal & ",'" & sUserInsert & "','" & sUserUpdate & "'"
        
     gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords   'Ejecuta la sentencia
    
        If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecuci�n
          'gsOperacionError = "Eliminando el Beneficiado. " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & err.Description
          SetMsgError "Ocurri� un error insertando la transacci�n . ", err
          lbOk = False
        End If
    
    invSaveDetalleTransaccion = lbOk
    Exit Function
    

End Function



Public Function invMasterAcutalizaSaldosInventarioPaquete(sDocumento As String, sPaquete As Integer, sIDTipoTransaccion As String, _
    sUsuario As String) As Boolean
    Dim lbOk As Boolean
   
    
    lbOk = True
      GSSQL = ""
      GSSQL = gsCompania & ".invUpdateMasterExistenciaBodega '" & sDocumento & "'," & sPaquete & "," & sIDTipoTransaccion & ",'" & sUsuario & "'"
        
     gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords   'Ejecuta la sentencia
    
        If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecuci�n
          'gsOperacionError = "Eliminando el Beneficiado. " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & err.Description
          SetMsgError "Ocurri� un error insertando la transacci�n . ", err
          lbOk = False
        End If
    
    invMasterAcutalizaSaldosInventarioPaquete = lbOk
    Exit Function
    

End Function

Private Sub SaveRstBatch(rst As ADODB.Recordset, sCodTra As String)
    Dim i As Integer
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

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecuci�n
    sDescr = ""  'Indica que ocurri� un error
    gsOperacionError = "Error en la b�squeda del art�culo !!!" & err.Description
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es v�lido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    sDescr = gRegistrosCmd(sfldNameDescr).value
  End If
  GetDescrCat = sDescr
  gRegistrosCmd.Close
  Exit Function
error:
  sDescr = ""
  gsOperacionError = "Ocurri� un error en la operaci�n de b�squeda de la descripci�n " & err.Description
  Resume Next
End Function

Private Sub LoadDescrTipoTransaccion(ByRef txtCaja As TextBox, KeyAscii As Integer)
   Dim sDescr As String
    Dim lbOk As Boolean
    If KeyAscii = vbKeyReturn Then
        sDescr = getDescrCatalogo(txtCaja, "IDTipo", "vinvPaqueteTipoMovimiento", "DescrTipo")
        If sDescr <> "" Then
            txtDescrTipoTransaccion.Text = sDescr
        Else
            lbOk = Mensaje("La transacci�n no existe por favor revise", ICO_ERROR, False)
        End If
    End If
End Sub

Private Sub LoadDescrBodega(ByRef txtCaja As TextBox, KeyAscii As Integer)
   Dim sDescr As String
    Dim lbOk As Boolean
    If KeyAscii = vbKeyReturn Then
        sDescr = getDescrCatalogo(txtCaja, "IDBodega", "invBodega", "Descr")
        If sDescr <> "" Then
            Me.txtDescrBodegaOrigen.Text = sDescr
        Else
            lbOk = Mensaje("La Bodega no existe por favor revise", ICO_ERROR, False)
        End If
    End If
End Sub

Private Sub LoadDescrBodegaDestino(ByRef txtCaja As TextBox, KeyAscii As Integer)
   Dim sDescr As String
    Dim lbOk As Boolean
    If KeyAscii = vbKeyReturn Then
        sDescr = getDescrCatalogo(txtCaja, "IDBodega", "invBodega", "Descr")
        If sDescr <> "" Then
            txtDescrBodegaDestino.Text = sDescr
        Else
            lbOk = Mensaje("La Bodega no existe por favor revise", ICO_ERROR, False)
        End If
    End If
End Sub

Private Sub LoadDescrProducto(ByRef txtCaja As TextBox, KeyAscii As Integer)
   Dim sDescr As String
    Dim lbOk As Boolean
    If KeyAscii = vbKeyReturn Then
        sDescr = getDescrCatalogo(txtCaja, "IDProducto", "invProducto", "Descr")
        If sDescr <> "" Then
            txtDescrArticulo.Text = sDescr
        Else
            lbOk = Mensaje("El Producto no existe por favor revise", ICO_ERROR, False)
        End If
    End If
End Sub

Private Sub LoadDescrLote(ByRef txtCaja As TextBox, KeyAscii As Integer)
   Dim sDescr As String
    Dim lbOk As Boolean
    If KeyAscii = vbKeyReturn Then
        sDescr = getDescrCatalogo(txtCaja, "IDProducto", "invProdcuto", "Descr")
        If sDescr <> "" Then
            txtDescrLote.Text = sDescr
        Else
            lbOk = Mensaje("El Producto no existe por favor revise", ICO_ERROR, False)
        End If
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

