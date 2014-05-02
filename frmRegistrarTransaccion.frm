VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRegistrarTransaccion 
   Caption         =   "Form1"
   ClientHeight    =   7815
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13410
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
   ScaleHeight     =   7815
   ScaleWidth      =   13410
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
      Left            =   12570
      Picture         =   "frmRegistrarTransaccion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
      Top             =   3030
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
      Left            =   12570
      Picture         =   "frmRegistrarTransaccion.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Deshacer / Cancelar"
      Top             =   3660
      Width           =   555
   End
   Begin VB.Frame Frame1 
      Height          =   2670
      Left            =   1110
      TabIndex        =   11
      Top             =   4755
      Width           =   11190
      Begin VB.TextBox txtCostoDolar 
         Height          =   315
         Left            =   6690
         TabIndex        =   40
         Top             =   2190
         Width           =   1515
      End
      Begin VB.TextBox txtCantidad 
         Height          =   315
         Left            =   2535
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
         Left            =   2535
         TabIndex        =   38
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdLote 
         Height          =   320
         Left            =   1935
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
         Left            =   4170
         TabIndex        =   33
         Top             =   1815
         Width           =   6675
      End
      Begin VB.CommandButton cmdDelLote 
         Height          =   320
         Left            =   3735
         Picture         =   "frmRegistrarTransaccion.frx":1CD6
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1800
         Width           =   300
      End
      Begin VB.CommandButton cmdArticulo 
         Height          =   320
         Left            =   1920
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
         Left            =   4155
         TabIndex        =   29
         Top             =   1425
         Width           =   6675
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
         Left            =   2520
         TabIndex        =   28
         Top             =   1410
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelArticulo 
         Height          =   320
         Left            =   3720
         Picture         =   "frmRegistrarTransaccion.frx":245A
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1410
         Width           =   300
      End
      Begin VB.CommandButton cmdBodegaDestino 
         Height          =   320
         Left            =   1920
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
         Left            =   4155
         TabIndex        =   24
         Top             =   1050
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
         Left            =   2520
         TabIndex        =   23
         Top             =   1035
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelBodegaDestino 
         Height          =   320
         Left            =   3720
         Picture         =   "frmRegistrarTransaccion.frx":2BDE
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1035
         Width           =   300
      End
      Begin VB.CommandButton cmdBodegaOrigen 
         Height          =   320
         Left            =   1920
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
         Left            =   4155
         TabIndex        =   19
         Top             =   660
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
         Left            =   2520
         TabIndex        =   18
         Top             =   645
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelBodegaOrigen 
         Height          =   320
         Left            =   3720
         Picture         =   "frmRegistrarTransaccion.frx":3362
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   645
         Width           =   300
      End
      Begin VB.CommandButton cmdTipoTransaccion 
         Height          =   320
         Left            =   1920
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
         Left            =   4155
         TabIndex        =   15
         Top             =   240
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
         Left            =   2520
         TabIndex        =   14
         Top             =   225
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelTipoTransaccion 
         Height          =   320
         Left            =   3720
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
         Left            =   5550
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
         Left            =   375
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
         Left            =   390
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
         Left            =   375
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
         Left            =   375
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
         Left            =   375
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
         Left            =   360
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
      Left            =   12555
      Picture         =   "frmRegistrarTransaccion.frx":3F28
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
      Top             =   2400
      Width           =   555
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   555
      Left            =   12570
      Picture         =   "frmRegistrarTransaccion.frx":4BF2
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
      Top             =   1380
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
      Left            =   12570
      Picture         =   "frmRegistrarTransaccion.frx":68BC
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Agrega el item con los datos digitados..."
      Top             =   4845
      Width           =   555
   End
   Begin RichTextLib.RichTextBox txtConcepto 
      Height          =   915
      Left            =   1140
      TabIndex        =   7
      Top             =   1305
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   1614
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmRegistrarTransaccion.frx":7586
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   2220
      Left            =   1110
      OleObjectBlob   =   "frmRegistrarTransaccion.frx":75FD
      TabIndex        =   1
      Top             =   2370
      Width           =   11205
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
      Left            =   9615
      TabIndex        =   5
      Top             =   810
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
      Left            =   8415
      TabIndex        =   4
      Top             =   825
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
      Left            =   -210
      TabIndex        =   0
      Top             =   0
      Width           =   13650
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   -405
      Picture         =   "frmRegistrarTransaccion.frx":FB42
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
          

          
      End If
      
      Set TDBG.DataSource = rstTmpMovimiento
      TDBG.ReBind
      
      
      'Me.dtgAjuste.Columns("Descr").FooterText = "Items de la transacción :     " & rstTransAI.RecordCount

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


Private Sub cmdSave_Click()
  Dim lbOk As Boolean
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
            gTrans = False
          If (gTrans = True) Then
            lbOk = Mensaje("La transacción ha sido guardada exitosamente", ICO_OK, False)
       
          
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
    invSaveCabeceraTransaccion = sqlCmd.Parameters(1).value
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
    
        If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
          'gsOperacionError = "Eliminando el Beneficiado. " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & err.Description
          SetMsgError "Ocurrió un error insertando la transacción . ", err
          lbOk = False
        End If
    
    invSaveDetalleTransaccion = lbOk
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
                                    rst.Fields("IdBodega").value, _
                                    rst.Fields("IDProducto").value, _
                                    rst.Fields("IDLote").value, _
                                    sCodTra, _
                                    rst.Fields("Fecha").value, _
                                    rst.Fields("IdTipo").value, _
                                    "", _
                                    "", _
                                    rst.Fields("Cantidad").value, _
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
    'gConet.RollbackTrans
End Sub

