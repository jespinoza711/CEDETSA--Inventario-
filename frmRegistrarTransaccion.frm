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
      Format          =   21364737
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
            Me.TDBG.Enabled = True
'            txtBodega.Text = "100"
'            txtDescrBodega.Text = ""
'            fmtTextbox txtBodega, "R"
'            fmtTextbox txtDescrBodega, "O"
            
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
            rstLote.ActiveConnection = gConet
          CargaDatosLotes rstLote, CInt(Trim(Me.txtLote.Text))
          ' Carga los datos del detalle de transacciones para ser grabados a la bd
          rstTmpMovimiento.AddNew
          rstTmpMovimiento!IdBodega = Me.txtBodegaOrigen.Text
          rstTmpMovimiento!DescrBodega = Me.txtDescrBodegaOrigen.Text
          rstTmpMovimiento!IDProducto = Me.txtArticulo.Text
          rstTmpMovimiento!DescrProducto = Me.txtDescrArticulo.Text
          rstTmpMovimiento!IDLote = Me.txtLote.Text
          rstTmpMovimiento!FechaVencimiento = rstLote!FechaVencimiento
          rstTmpMovimiento!FechaFabricacion = rstLote!FechaFabricacion
          rstTmpMovimiento!LoteInterno = Me.txtDescrLote.Text
          rstTmpMovimiento!IdTipo = Me.txtTipoTransaccion.Text
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
          rstTmpMovimiento!IdBodega = Me.txtBodegaOrigen.Text
          rstTmpMovimiento!DescrBodega = Me.txtDescrBodegaOrigen.Text
          rstTmpMovimiento!IDProducto = Me.txtArticulo.Text
          rstTmpMovimiento!DescrProducto = Me.txtDescrArticulo.Text
          rstTmpMovimiento!IDLote = Me.txtLote.Text
          rstTmpMovimiento!LoteInterno = Me.txtDescrLote.Text
          rstTmpMovimiento!IdTipo = Me.txtTipoTransaccion.Text
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
      HabilitarBotones
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
'
    If Not (rstTmpMovimiento.EOF And rstTmpMovimiento.BOF) Then
        Me.txtTipoTransaccion.Text = rstTmpMovimiento("IDTipo").value
        Me.txtDescrTipoTransaccion.Text = rstTmpMovimiento("DescrTipo").value
        'Contemplar para traslados
        'Me.txtDescrBodegaDestino.Text = rstTmpMovimiento("DescrBodega").value
        
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

'Private Function CreaCabeceraTran() As Boolean
'    Dim lTotal As Double
'    Dim lTotalD As Double
'    Dim lbOk As Boolean
'    On Error GoTo errores
'    lbOk = False
'
'    gConet.Execute "invInsertCabMovimientos "
'    ' preparacion del recordset CABECERA
'    Dim rstTransCABCO As ADODB.Recordset
'    Set rstTransCABCO = New ADODB.Recordset
'    If rstTransCABCO.State = adStateOpen Then rstTransCABCO.Close
'    rstTransCABCO.ActiveConnection = gConet 'Asocia la conexión de trabajo
'    rstTransCABCO.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
'    rstTransCABCO.CursorLocation = adUseClient ' Cursor local al cliente
'    rstTransCABCO.LockType = adLockPessimistic 'adLockOptimistic
'    If rstTransCABCO.State = adStateOpen Then rstTransCABCO.Close
'    gstrSQL = "SELECT * FROM TRANSACCION WHERE 1=2"
'    rstTransCABCO.Open gstrSQL
'
'    'gConet.BeginTrans
'
'    If Not rstTransDETCO.EOF Then
'      lTotal = 0
'        rstTransCABCO.AddNew
'        rstTransCABCO!Fecha = Format(Me.dtpFecha.value, "YYYY/MM/DD") 'Format(Now, )
'        rstTransCABCO!CorTran = lCorrelativo
'        rstTransCABCO!CodTipoTran = ParametrosGenerales.CodTranCompra
'        rstTransCABCO!Documento = Me.txtNumFactura.Text
'        rstTransCABCO!Descr = Me.txtReferencia.Text
'        rstTransCABCO!Fecha = Format(Me.dtpFactura.value, "YYYY/MM/DD")
'        rstTransCABCO!Usuario = gNombreUsuario
'        rstTransCABCO.Update
'        lbOk = True
'        CreaCabeceraCO = lbOk
'        rstTransDETCO.MoveFirst ' se ubica en el inicio
'        Exit Function
'    End If
'errores:
'    gTrans = False
'    CreaCabeceraCO = lbOk
'    'gConet.RollbackTrans
'    Exit Function
'End Function


Private Sub cmdSave_Click()
  Dim lbOk As Boolean
    'On Error GoTo errores
    
    If rstTmpMovimiento.RecordCount > 0 Then
'        rstTmpMovimiento!CorTran = lCorrelativo
      
'      If gTasaCambio = 0 Then
'        lbOk = Mensaje("La tasa de cambio es Cero llame a informática por favor ", ICO_ERROR, False)
'        Exit Sub
'      End If
      
      
      gConet.BeginTrans ' inicio aqui la transacción
      gBeginTransNoEnd = True
      Dim sDocumento As String
      sDocumento = CreaCabecera()
      If sDocumento <> "" Then ' salva la cabecera
'      '----------- Progress bar
'        ProgressBar1.Visible = True
'        ProgressBar1.Min = 1
'        ProgressBar1.Max = 100
'        ProgressBar1.Value = 20
'        lblProgress.Caption = "Preparando datos"
'        lblProgress.Refresh
'      '----------- Progress bar
        SaveRstBatch rstTmpMovimiento, sDocumento ' salva el detalle que esta en batch
        'rstTransCO.Update
'      '----------- Progress bar
'        ProgressBar1.Value = 70
'        lblProgress.Caption = "Costo Promedio"
'        lblProgress.Refresh
'      '----------- Progress bar
        
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
        
'      '----------- Progress bar
'        ProgressBar1.Value = 90
'        lblProgress.Caption = "Act. inventario"
'        lblProgress.Refresh
'      '----------- Progress bar
        
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
            lbOk = Mensaje("La transacción ha sido guardada exitosamente", ICO_OK, False)
         ' ProgressBar1.Visible = False
          'lblProgress.Caption = ""
          
         ' lblNoTra.Caption = ""
            Me.cmdAdd.Enabled = False
          
'          Me.tdgCompra.Columns("Descr").FooterText = "Items de la transacción :     "
'          tdgCompra.Columns("Costo").FooterText = "Total : "
'          tdgCompra.Columns("Monto").NumberFormat = "###,###,##0.#0"
'          tdgCompra.Columns("Monto").FooterText = "0"
  
        
            Accion = View
            HabilitarBotones
            HabilitarControles
        
'          If rstTransCO.State = adStateOpen Then rstTransCO.Close
'          If rstTransDETCO.State = adStateOpen Then rstTransDETCO.Close
'          If gRegistrosCODET.State = adStateOpen Then gRegistrosCODET.Close
'          If rstTransCABCO.State = adStateOpen Then rstTransCABCO.Close
'
           
            gConet.CommitTrans
            gBeginTransNoEnd = False
            'InicializaFormulario
            Exit Sub
          Else
            gConet.RollbackTrans
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

Private Function CreaCabecera() As String
    Dim lTotal As Double
    Dim lTotalD As Double
    Dim lbOk As Boolean
    On Error GoTo errores
    lbOk = False
    
    ' preparacion del recordset CABECERA
    Dim sqlCmd As New ADODB.Command
    
    Dim sDocumento As String
    sqlCmd.ActiveConnection = gConet
    sqlCmd.CommandText = "invInsertCabMovimientos "
    sqlCmd.CommandType = adCmdStoredProc
        
    sqlCmd.Parameters.Append sqlCmd.CreateParameter(, adInteger, adParamInput, 100, Me.gsIDTipoTransaccion)
    sqlCmd.Parameters.Append sqlCmd.CreateParameter(, adVarChar, adParamOutput, 400, sDocumento)
    sqlCmd.Parameters.Append sqlCmd.CreateParameter(, adDate, adParamInput, 100, Me.dtpFecha.value)
    sqlCmd.Parameters.Append sqlCmd.CreateParameter(, adVarChar, adParamInput, 255, Me.txtConcepto.Text)
    sqlCmd.Parameters.Append sqlCmd.CreateParameter(, adVarChar, adParamInput, 255, gsUser)
    sqlCmd.Parameters.Append sqlCmd.CreateParameter(, adVarChar, adParamInput, 255, gsUser)
    sqlCmd.Execute
    CreaCabecera = sqlCmd.Parameters(1).value
    Exit Function
errores:
    gTrans = False
    CreaCabecera = lbOk
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
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (rst Is Nothing) Then Set rst = Nothing
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
    
    If (rst.BOF And rst.EOF) Then  'Si no es válido
        lbOk = False  'Indica que no es válido
    End If
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


Private Sub SaveRstBatch(rst As ADODB.Recordset, sCodTra As String)
    Dim i As Integer
    On Error GoTo errores
    'Set lRegistros = New ADODB.Recordset  'Inicializa la variable de los registros
    'gConet.BeginTrans
    
    
    If rst.RecordCount > 0 Then
      rst.MoveFirst
      While Not rst.EOF
            Dim sqlCmd As New ADODB.Command
            With sqlCmd
                .ActiveConnection = gConet
            
                .CommandText = "invInsertMovimientos"
                .CommandType = adCmdStoredProc
                .CommandTimeout = 30
                .Parameters.Append .CreateParameter("@IDPaquete", adInteger, adParamInput, , Me.gsIDTipoTransaccion)
                .Parameters.Append .CreateParameter("@IDBodega", adInteger, adParamInput, , rst.Fields("IdBodega").value)
                .Parameters.Append .CreateParameter("@IDProducto", adInteger, adParamInput, , rst.Fields("IDProducto").value)
                .Parameters.Append .CreateParameter("@IDLote", adInteger, adParamInput, , rst.Fields("IDLote").value)
                .Parameters.Append .CreateParameter("@Documento", adVarChar, adParamInput, 20, sCodTra)
                .Parameters.Append .CreateParameter("@Fecha", adDate, adParamInput, , rst.Fields("Fecha").value)
                .Parameters.Append .CreateParameter("@IdTipo", adInteger, adParamInput, , rst.Fields("IdTipo").value)
                .Parameters.Append .CreateParameter("@Transaccion", adVarChar, adParamInput, 10, rst.Fields("Transaccion").value)
                .Parameters.Append .CreateParameter("@Naturaleza", adVarChar, adParamInput, 1, rst.Fields("Naturaleza").value)
                .Parameters.Append .CreateParameter("@Cantidad", adDecimal, adParamInput, , rst.Fields("Cantidad").value)
                .Parameters.Append .CreateParameter("@CostoDolar", adDecimal, adParamInput, , rst.Fields("CostoDolar").value)
                .Parameters.Append .CreateParameter("@CostoLocal", adDecimal, adParamInput, , rst.Fields("CostoLocal").value)
                .Parameters.Append .CreateParameter("@PrecioLocal", adDecimal, adParamInput, , rst.Fields("PrecioLocal").value)
                .Parameters.Append .CreateParameter("@PrecioDolar", adDecimal, adParamInput, , rst.Fields("PrecioDolar").value)
                .Parameters.Append .CreateParameter("@UserInsert", adVarChar, adParamInput, 20, rst.Fields("UserInsert").value)
                .Parameters.Append .CreateParameter("@UserUpdate", adVarChar, adParamInput, 20, rst.Fields("UserInsert").value)
                .Parameters("@Cantidad").Precision = 28
                .Parameters("@Cantidad").NumericScale = 8
                .Parameters("@CostoDolar").Precision = 28
                .Parameters("@CostoDolar").NumericScale = 8
                .Parameters("@CostoLocal").Precision = 28
                .Parameters("@CostoLocal").NumericScale = 8
                .Parameters("@PrecioLocal").Precision = 28
                .Parameters("@PrecioLocal").NumericScale = 8
                .Parameters("@PrecioDolar").Precision = 28
                .Parameters("@PrecioDolar").NumericScale = 8
                .Execute
               
                
            End With
            rst.MoveNext
      Wend
      

    End If
    Exit Sub
errores:
    gTrans = False
    'gConet.RollbackTrans
End Sub

