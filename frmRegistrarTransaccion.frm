VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRegistrarTransaccion 
   Caption         =   "Form1"
   ClientHeight    =   7815
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10485
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
   ScaleHeight     =   7815
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
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
      Left            =   9600
      Picture         =   "frmRegistrarTransaccion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
      Top             =   3165
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
      Left            =   9600
      Picture         =   "frmRegistrarTransaccion.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Deshacer / Cancelar"
      Top             =   3795
      Width           =   555
   End
   Begin VB.Frame Frame1 
      Height          =   2670
      Left            =   1125
      TabIndex        =   11
      Top             =   4680
      Width           =   8415
      Begin VB.TextBox txtCostoDolar 
         Height          =   315
         Left            =   6540
         TabIndex        =   40
         Top             =   2190
         Width           =   1515
      End
      Begin VB.TextBox txtCantidad 
         Height          =   315
         Left            =   2385
         TabIndex        =   39
         Top             =   2190
         Width           =   1515
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
         Left            =   2385
         TabIndex        =   38
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdLote 
         Height          =   320
         Left            =   1785
         Picture         =   "frmRegistrarTransaccion.frx":1994
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1800
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
         Left            =   4020
         TabIndex        =   33
         Top             =   1815
         Width           =   4035
      End
      Begin VB.CommandButton cmdDelLote 
         Height          =   320
         Left            =   3585
         Picture         =   "frmRegistrarTransaccion.frx":1CD6
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1800
         Width           =   300
      End
      Begin VB.CommandButton cmdArticulo 
         Height          =   320
         Left            =   1770
         Picture         =   "frmRegistrarTransaccion.frx":2118
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1410
         Width           =   300
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
         Left            =   4005
         TabIndex        =   29
         Top             =   1425
         Width           =   4035
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
         Left            =   2370
         TabIndex        =   28
         Top             =   1410
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelArticulo 
         Height          =   320
         Left            =   3570
         Picture         =   "frmRegistrarTransaccion.frx":245A
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1410
         Width           =   300
      End
      Begin VB.CommandButton cmdBodegaDestino 
         Height          =   320
         Left            =   1770
         Picture         =   "frmRegistrarTransaccion.frx":289C
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1035
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
         Left            =   4005
         TabIndex        =   24
         Top             =   1050
         Width           =   4035
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
         Left            =   2370
         TabIndex        =   23
         Top             =   1035
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelBodegaDestino 
         Height          =   320
         Left            =   3570
         Picture         =   "frmRegistrarTransaccion.frx":2BDE
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1035
         Width           =   300
      End
      Begin VB.CommandButton cmdBodegaOrigen 
         Height          =   320
         Left            =   1770
         Picture         =   "frmRegistrarTransaccion.frx":3020
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   630
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
         Left            =   4005
         TabIndex        =   19
         Top             =   660
         Width           =   4035
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
         Left            =   2370
         TabIndex        =   18
         Top             =   645
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelBodegaOrigen 
         Height          =   320
         Left            =   3570
         Picture         =   "frmRegistrarTransaccion.frx":3362
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   645
         Width           =   300
      End
      Begin VB.CommandButton cmdTipoTransaccion 
         Height          =   320
         Left            =   1770
         Picture         =   "frmRegistrarTransaccion.frx":37A4
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   225
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
         Left            =   4005
         TabIndex        =   15
         Top             =   240
         Width           =   4035
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
         Left            =   2370
         TabIndex        =   14
         Top             =   225
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelTipoTransaccion 
         Height          =   320
         Left            =   3570
         Picture         =   "frmRegistrarTransaccion.frx":3AE6
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   225
         Width           =   300
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
         Left            =   5400
         TabIndex        =   37
         Top             =   2235
         Width           =   1095
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
         Left            =   225
         TabIndex        =   36
         Top             =   2235
         Width           =   1095
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
         Left            =   240
         TabIndex        =   35
         Top             =   1815
         Width           =   1635
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
         Left            =   225
         TabIndex        =   31
         Top             =   1425
         Width           =   1635
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
         Left            =   225
         TabIndex        =   26
         Top             =   1050
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
         Left            =   225
         TabIndex        =   21
         Top             =   660
         Width           =   1635
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo Transacción:"
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
         Left            =   210
         TabIndex        =   12
         Top             =   255
         Width           =   1635
      End
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
      Left            =   9600
      Picture         =   "frmRegistrarTransaccion.frx":3F28
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
      Top             =   2535
      Width           =   555
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   555
      Left            =   9600
      Picture         =   "frmRegistrarTransaccion.frx":4BF2
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
      Top             =   1485
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
      Left            =   9600
      Picture         =   "frmRegistrarTransaccion.frx":68BC
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Agrega el item con los datos digitados..."
      Top             =   4815
      Width           =   555
   End
   Begin RichTextLib.RichTextBox txtConcepto 
      Height          =   915
      Left            =   1140
      TabIndex        =   7
      Top             =   1305
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   1614
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmRegistrarTransaccion.frx":7586
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   2220
      Left            =   1095
      OleObjectBlob   =   "frmRegistrarTransaccion.frx":75FD
      TabIndex        =   1
      Top             =   2370
      Width           =   8415
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   1155
      TabIndex        =   2
      Top             =   750
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
      Format          =   61603841
      CurrentDate     =   41095
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
      Left            =   255
      TabIndex        =   6
      Top             =   1335
      Width           =   915
   End
   Begin VB.Label lblTransaccion 
      Caption         =   "Transacción:"
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
      Left            =   7710
      TabIndex        =   5
      Top             =   735
      Width           =   2640
   End
   Begin VB.Label Label1 
      Caption         =   "Transacción:"
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
      Left            =   6510
      TabIndex        =   4
      Top             =   750
      Width           =   1170
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
      Left            =   270
      TabIndex        =   3
      Top             =   780
      Width           =   915
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
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   10710
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   -90
      Picture         =   "frmRegistrarTransaccion.frx":FB42
      Stretch         =   -1  'True
      Top             =   -300
      Width           =   11490
   End
End
Attribute VB_Name = "frmRegistrarTransaccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rst As ADODB.Recordset
Dim bOrdenCodigo As Boolean
Dim bOrdenDescr As Boolean
Dim sCodSucursal As String
Dim sSoloActivo As String
Dim Accion As TypAccion
Public gsFormCaption As String
Public gsTitle As String
Public gsIDTipoTransaccion As Integer
Dim sMensajeError As String

Private rstTmpMovimiento As ADODB.Recordset
Dim rstLote As ADODB.Recordset
'Private rstMovimiento As ADODB.Recordset
'Private rstTransCABAI As ADODB.Recordset
Dim lbAIenProceso As Boolean ' Indica si un ajuste está en proceso

Dim gTrans As Boolean ' se dispara si hubo error en medio de la transacción
Dim gBeginTransNoEnd As Boolean ' Indica si hubo un begin sin rollback o commit
Dim gTasaCambio As Double

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
End Sub

Public Sub HabilitarControles()
    Select Case Accion
        Case TypAccion.Add
            dtpFecha.value = Format(Now, "YYYY/MM/DD")
            txtTipoTransaccion.Text = ""
            fmtTextbox txtTipoTransaccion, "R"
            Me.txtDescrTipoTransaccion.Text = ""
            fmtTextbox Me.txtDescrTipoTransaccion, "R"
            
            txtBodegaDestino.Text = ""
            fmtTextbox txtBodegaDestino, "R"
            txtDescrBodegaDestino.Text = ""
            fmtTextbox txtDescrBodegaDestino, "R"
            
            txtBodegaOrigen.Text = ""
            fmtTextbox txtBodegaOrigen, "R"
            txtDescrBodegaOrigen.Text = ""
            fmtTextbox txtDescrBodegaOrigen, "R"
            
            fmtTextbox txtArticulo, "R"
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
            Me.cmdBodegaDestino.Enabled = True
            Me.cmdDelBodegaDestino.Enabled = True
            Me.cmdArticulo.Enabled = True
            Me.cmdDelArticulo.Enabled = True
            Me.cmdLote.Enabled = True
            Me.cmdDelLote.Enabled = True
            
'            txtBodega.Text = "100"
'            txtDescrBodega.Text = ""
'            fmtTextbox txtBodega, "R"
'            fmtTextbox txtDescrBodega, "O"
            
        Case TypAccion.Edit
            dtpFecha.value = Format(Now, "YYYY/MM/DD")
            txtTipoTransaccion.Text = ""
            fmtTextbox txtTipoTransaccion, "R"
            Me.txtDescrTipoTransaccion.Text = ""
            fmtTextbox Me.txtDescrTipoTransaccion, "R"
            
            txtBodegaDestino.Text = ""
            fmtTextbox txtBodegaDestino, "R"
            txtDescrBodegaDestino.Text = ""
            fmtTextbox txtDescrBodegaDestino, "R"
            
            txtBodegaOrigen.Text = ""
            fmtTextbox txtBodegaOrigen, "R"
            txtDescrBodegaOrigen.Text = ""
            fmtTextbox txtDescrBodegaOrigen, "R"
            
            fmtTextbox txtArticulo, "R"
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
            Me.cmdBodegaDestino.Enabled = True
            Me.cmdDelBodegaDestino.Enabled = True
            Me.cmdArticulo.Enabled = True
            Me.cmdDelArticulo.Enabled = True
            Me.cmdLote.Enabled = True
            Me.cmdDelLote.Enabled = True
           
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
    ElseIf (Me.txtDescrLote.Text = "") Then
        sMensajeError = "Por favor seleccione el lote del producto"
        Valida = False
    ElseIf (Me.txtCantidad.Text = "") Then
        sMensajeError = "Por favor digite la cantidad del producto"
        Valida = False
    ElseIf (Me.txtCostoDolar.Text = "") And (Me.gsIDTipoTransaccion = 5) Then
        sMensajeError = "Por favor seleccione producto"
        Valida = False
    End If
    ValCtrls = Valida
End Function

Private Sub cmdAdd_Click()
   
    Dim lbOk As Boolean
    If Not ValCtrls Then
    lbOk = Mensaje("Revise sus datos por favor !!! " & sMensajeError, ICO_ERROR, False)
    Exit Sub
    End If
    
      ' Actualiza el rst temporal
      Me.cmdSave.Enabled = True
      If (Accion = Add) Then
          If ExiteRstKey(rstTmpMovimiento, "IDBODEGA=" & Me.txtBodegaOrigen.Text & " AND IDPRODUCTO=" & Me.txtArticulo.Text & _
                                        " AND IDLOTE=" & Me.txtLote.Text & " AND IDTIPO=" & Me.txtTipoTransaccion.Text) Then
            lbOk = Mensaje("Ya existe ese el registro en la transacción", ICO_ERROR, False)
'            txtCodProducto.Text = ""
'            txtDescrProducto.Text = ""
'            Me.txtCantidad.Text = ""
'            Me.txtCodBodega.Text = ""
'            Me.txtDescrBodega.Text = ""
'            Me.txtCodProducto.SetFocus
            Exit Sub
          End If
          Set rstLote = New ADODB.Recordset
            rstLote.ActiveConnection = conn
          CargaDatosLotes rstLote, CInt(Trim(Me.txtLote.Text))
          ' Carga los datos del detalle de transacciones para ser grabados a la bd
          rstTmpMovimiento.AddNew
          rstTmpMovimiento!IDBodega = Me.txtBodegaOrigen.Text
          rstTmpMovimiento!DescrBodega = Me.txtDescrBodegaOrigen.Text
          rstTmpMovimiento!IDPRODUCTO = Me.txtArticulo.Text
          rstTmpMovimiento!DescrProducto = Me.txtDescrArticulo.Text
          rstTmpMovimiento!IDLote = Me.txtLote.Text
          rstTmpMovimiento!FechaVencimiento = rstLote!FechaVencimiento
          rstTmpMovimiento!FechaFabricacion = rstLote!FechaFabricacion
          rstTmpMovimiento!LoteInterno = Me.txtDescrLote.Text
          rstTmpMovimiento!IDTipo = Me.txtTipoTransaccion.Text
          rstTmpMovimiento!DescrTipo = Me.txtDescrTipoTransaccion.Text
          rstTmpMovimiento!Cantidad = Me.txtCantidad.Text
          rstTmpMovimiento!Fecha = Format(Me.dtpFecha.value, "YYYY/MM/DD")
          rstTmpMovimiento!CostoLocal = 0 'GetLastCostoProm(me.txtcodProducto.Text, "C")
          rstTmpMovimiento!CostoDolar = 0 'GetLastCostoProm(txtCodProdAI.Text, "D")
          rstTmpMovimiento!PrecioLocal = 0 '(rstTransDETAI!cant * rstTransDETAI!Costo)
          rstTmpMovimiento!PrecioDolar = 0 '(rstTransDETAI!cant * rstTransDETAI!Costod)
          rstTmpMovimiento!UserInsert = gsUser
          rstTmpMovimiento.Update
          rstTmpMovimiento.MoveFirst
          
          
'          ' carga los datos para ser mostrados en el grid
'          rstTmpMovimiento.AddNew
'          rstTmpMovimiento!CodTipoTran = ParametrosGenerales.CodTranAjuste
'          rstTmpMovimiento!CorTran = lCorrelativo
'        '  rstTransAI!CODPROVEEDOR = 0
'          rstTmpMovimiento!codproducto = Me.txtCodProducto.Text
'          rstTmpMovimiento!CodBodega = Me.txtCodBodega.Text
'          rstTmpMovimiento!Cantidad = Me.txtCantidad.Text
'        '  rstTransAI!Costo = txtCosto.Text
'        '  rstTransAI!Monto = rstTransCO!Cant * rstTransCO!Costo
'          rstTmpMovimiento!Descr = Me.txtDescrProducto.Text
'          rstTransAI!DescBod = Me.txtDescrBodega.Text
'        '  rstTransAI!numfac = txtFactura.Text
'          rstTmpMovimiento!Fecha = Format(Me.dtpFecha.value, "YYYY/MM/DD")
      ElseIf (Accion = Edit) Then
          rstTmpMovimiento!IDBodega = Me.txtBodegaOrigen.Text
          rstTmpMovimiento!DecrBodega = Me.txtDescrBodegaOrigen.Text
          rstTmpMovimiento!IDPRODUCTO = Me.txtArticulo.Text
          rstTmpMovimiento!DescrProducto = Me.txtDescrArticulo.Text
          rstTmpMovimiento!IDLote = Me.txtLote.Text
          rstTmpMovimiento!LoteInterno = Me.txtDescrLote.Text
          rstTmpMovimiento!IDTipo = Me.txtTipoTransaccion.Text
          rstTmpMovimiento!DescrTipo = Me.txtDescrTipoTransaccion.Text
          rstTmpMovimiento!Cantidad = Me.txtCantidad.Text
          rstTmpMovimiento!Fecha = Format(Me.dtpFecha.value, "YYYY/MM/DD")
          rstTmpMovimiento!CostoLocal = 0 'GetLastCostoProm(me.txtcodProducto.Text, "C")
          rstTmpMovimiento!CostoDolar = 0 'GetLastCostoProm(txtCodProdAI.Text, "D")
          rstTmpMovimiento!PrecioLocal = 0 '(rstTransDETAI!cant * rstTransDETAI!Costo)
          rstTmpMovimiento!PrecioDolar = 0 '(rstTransDETAI!cant * rstTransDETAI!Costod)
          rstTmpMovimiento!UserInsert = gsUser
        
          rstTmpMovimiento.Update
          
       
          
         ' Me.cmdFindProducto.Enabled = True
          
      End If
      
      Set TDBG.DataSource = rstTmpMovimiento
      TDBG.ReBind
    '  rstTransAI!Descr = txtDescrProdAI.Text
      
      
      'Me.dtgAjuste.Columns("Descr").FooterText = "Items de la transacción :     " & rstTransAI.RecordCount
     
      
'      Me.txtCodProducto.Text = ""
'      Me.txtDescrProducto.Text = ""
'      Me.txtCantidad.Text = ""
'      Me.cmdEliminar.Enabled = True
'      Me.cmdModificar.Enabled = True
'      Me.cmdGuardar.Enabled = True
'      Me.cmdAgregar.Enabled = False
      Accion = Add
      HabilitarControles
End Sub

Private Sub cmdArticulo_Click()
Dim frm As New frmBrowseCat
    
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
End Sub
Private Sub GetDataFromGridToControl() 'EDITAR
'If Not (rst.EOF And rst.BOF) Then
'    txtBodega.Text = rst("IDBodega").value
'    txtDescrBodega.Text = rst("DescrBodega").value
'    If rst("Activo").value = True Then
'        chkActivo.value = 1
'    Else
'        chkActivo.value = 0
'    End If
'    If rst("Factura").value = True Then
'        chkFactura.value = 1
'    Else
'        chkFactura.value = 0
'    End If
'Else
'    txtBodega.Text = ""
'    txtDescrBodega.Text = ""
'End If

End Sub

Private Sub cmdEliminar_Click()
    Dim lbOk As Boolean
    Dim sMsg As String
    Dim sTipo As String
    Dim sFiltro As String
    Dim sActivo As String
    Dim sFactura As String

    If txtBodega.Text = "" Then
        lbOk = Mensaje("La Bodega no puede estar en Blanco", ICO_ERROR, False)
        Exit Sub
    End If
    If chkActivo.value = 1 Then
        sActivo = "1"
    Else
        sActivo = "0"

    End If
    
    If chkFactura.value = 1 Then
        sFactura = "1"
    Else
        sFactura = "0"
    End If
    
    ' hay que validar la integridad referencial
    lbOk = Mensaje("Está seguro de eliminar la Bodega " & rst("IDBodega").value, ICO_PREGUNTA, True)
    If lbOk Then
                lbOk = invUpdateBodega("D", txtBodega.Text, txtDescrBodega.Text, sActivo, sFactura)
        
        If lbOk Then
            sMsg = "Borrado Exitosamente ... "
            lbOk = Mensaje(sMsg, ICO_OK, False)
            ' actualiza datos
            cargaGrid
        End If
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
    Dim sMsg As String
    Dim sActivo As String
    Dim sFactura As String
    Dim sFiltro As String
    If txtBodega.Text = "" Then
        lbOk = Mensaje("La Bodega no puede estar en Blanco", ICO_ERROR, False)
        Exit Sub
    End If
    If chkActivo.value = 1 Then
        sActivo = "1"
    Else
        sActivo = "0"
    End If
    If chkFactura.value = 1 Then
        sFactura = "1"
    Else
        sFactura = "0"
    End If
    If txtDescrBodega.Text = "" Then
        lbOk = Mensaje("La Descripción del Centro no puede estar en blanco", ICO_ERROR, False)
        Exit Sub
    End If
    

        
    If Accion = Add Then
    
        If Not (rst.EOF And rst.BOF) Then
            sFiltro = "IDBodega = '" & txtBodega.Text & "'"
            If ExiteRstKey(rst, sFiltro) Then
               lbOk = Mensaje("Ya exista Bodega ", ICO_ERROR, False)
                txtBodega.SetFocus
            Exit Sub
            End If
        End If
    
            lbOk = invUpdateBodega("I", txtBodega.Text, txtDescrBodega.Text, sActivo, sFactura)
            
            If lbOk Then
                sMsg = "La Bodega ha sido registrada exitosamente ... "
                lbOk = Mensaje(sMsg, ICO_OK, False)
                ' actualiza datos
                cargaGrid
                Accion = View
                HabilitarControles
                HabilitarBotones
            Else
                 sMsg = "Ha ocurrido un error tratando de Agregar la Bodega... "
                lbOk = Mensaje(sMsg, ICO_ERROR, False)
            End If
    End If ' si estoy adicionando
        If Accion = Edit Then
            If Not (rst.EOF And rst.BOF) Then
                lbOk = invUpdateBodega("U", txtBodega.Text, txtDescrBodega.Text, sActivo, sFactura)
                If lbOk Then
                    sMsg = "Ha ocurrido un error tratando de Actualizar la Bodega... "
                    lbOk = Mensaje(sMsg, ICO_ERROR, False)
                    ' actualiza datos
                    cargaGrid
                    Accion = View
                    HabilitarControles
                    HabilitarBotones
                End If
            End If
        
    End If ' si estoy adicionando

End Sub

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
End Sub

Private Sub cmdUndo_Click()
    GetDataFromGridToControl
    Accion = Add
    HabilitarBotones
    HabilitarControles
End Sub

Private Sub Form_Load()
    Set rstTmpMovimiento = New ADODB.Recordset
    If rstTmpMovimiento.State = adStateOpen Then rst.Close
    rstTmpMovimiento.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rstTmpMovimiento.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rstTmpMovimiento.CursorLocation = adUseClient ' Cursor local al cliente
    rstTmpMovimiento.LockType = adLockOptimistic
    Me.Caption = gsFormCaption
    Me.lbFormCaption = gsTitle
    'gTasaCambio = GetTasadeCambio(Format(Now, "YYYY/MM/DD"))
    PreparaRst ' Prepara los Recordsets
    Set Me.TDBG.DataSource = rstTmpMovimiento
    Me.TDBG.Refresh
    
    Accion = Add
    HabilitarBotones
    HabilitarControles
    
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


Private Sub cargaGrid()
    Dim sIndependiente As String
    If rst.State = adStateOpen Then rst.Close
    rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rst.CursorLocation = adUseClient ' Cursor local al cliente
    rst.LockType = adLockOptimistic
    GSSQL = gsCompania & ".invGetBodegas -1"
    If rst.State = adStateOpen Then rst.Close
    Set rst = GetRecordset(GSSQL)
    If Not (rst.EOF And rst.BOF) Then
      Set TDBG.DataSource = rst
      'CargarDatos rst, TDBG, "Codigo", "Descr"
      TDBG.Refresh
      'IniciaIconos
    End If
End Sub


Private Sub TDBG_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    GetDataFromGridToControl
    HabilitarControles
    HabilitarBotones
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (rst Is Nothing) Then Set rst = Nothing
End Sub

Public Sub CargaDatosLotes(rst As ADODB.Recordset, iIDLote As Integer)
    Dim lbOk As Boolean
    'On Error GoTo error
    lbOk = True
      gstrSQL = "SELECT IDLote, LoteInterno, LoteProveedor, FechaVencimiento, FechaFabricacion"
    
      gstrSQL = gstrSQL & " FROM " & " dbo.invLOTE " 'Constuye la sentencia SQL
      gstrSQL = gstrSQL & " WHERE IDLote=" & iIDLote
      If rst.State = adStateOpen Then rst.Close
      rst.Open gstrSQL, , adOpenKeyset, adLockOptimistic
    
    If (rst.BOF And rst.EOF) Then  'Si no es válido
        lbOk = False  'Indica que no es válido
    End If
End Sub

