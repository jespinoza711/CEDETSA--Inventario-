VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRegistrarTraslado 
   Caption         =   "Form1"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   8490
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgTitle 
      Left            =   4170
      Top             =   990
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistrarTraslado.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistrarTraslado.frx":1CDA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmDetalle 
      Caption         =   " Agregar Detalle:"
      Height          =   1965
      Left            =   210
      TabIndex        =   30
      Top             =   6510
      Width           =   12045
      Begin VB.TextBox txtIDProducto 
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
         Height          =   345
         Left            =   1170
         TabIndex        =   40
         Top             =   540
         Width           =   1095
      End
      Begin VB.TextBox txtDescrProducto 
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
         Height          =   345
         Left            =   2670
         TabIndex        =   39
         Top             =   540
         Width           =   5715
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
         Left            =   8790
         TabIndex        =   38
         Top             =   540
         Width           =   2145
      End
      Begin VB.TextBox txtIDLote 
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
         Height          =   345
         Left            =   1170
         TabIndex        =   37
         Top             =   1020
         Width           =   1095
      End
      Begin VB.TextBox txtDescrLote 
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
         Height          =   345
         Left            =   2670
         TabIndex        =   36
         Top             =   1020
         Width           =   5715
      End
      Begin VB.TextBox txtCantidad 
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
         Height          =   345
         Left            =   9720
         TabIndex        =   35
         Top             =   1020
         Width           =   1095
      End
      Begin VB.CommandButton cmdProducto 
         Height          =   315
         Left            =   2310
         Picture         =   "frmRegistrarTraslado.frx":39B4
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   540
         Width           =   300
      End
      Begin VB.CommandButton cmdLote 
         Height          =   315
         Left            =   2310
         Picture         =   "frmRegistrarTraslado.frx":3CF6
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1020
         Width           =   300
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
         Left            =   11160
         Picture         =   "frmRegistrarTraslado.frx":4038
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Agrega el item con los datos digitados..."
         Top             =   750
         Width           =   555
      End
      Begin VB.TextBox txtCantidadRemitida 
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
         Height          =   345
         Left            =   9720
         TabIndex        =   31
         Top             =   1470
         Width           =   1095
      End
      Begin VB.Label lblProducto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Producto:"
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
         Height          =   225
         Left            =   420
         TabIndex        =   44
         Top             =   570
         Width           =   825
      End
      Begin VB.Label lblLote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
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
         Height          =   225
         Left            =   450
         TabIndex        =   43
         Top             =   1050
         Width           =   825
      End
      Begin VB.Label lblCantidad 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
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
         Height          =   225
         Left            =   8850
         TabIndex        =   42
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label lblCantidadRemitida 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   " Cantidad Remitida:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   7860
         TabIndex        =   41
         Top             =   1560
         Width           =   1815
      End
   End
   Begin VB.TextBox txtNumSalida 
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
      Height          =   345
      Left            =   7140
      TabIndex        =   28
      Top             =   2730
      Width           =   1725
   End
   Begin VB.Frame Frame1 
      Height          =   2685
      Left            =   12300
      TabIndex        =   21
      Top             =   3210
      Width           =   765
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
         Left            =   90
         Picture         =   "frmRegistrarTraslado.frx":4D02
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Deshacer / Cancelar"
         Top             =   2040
         Width           =   555
      End
      Begin VB.CommandButton cmdSave 
         Enabled         =   0   'False
         Height          =   555
         Left            =   90
         Picture         =   "frmRegistrarTraslado.frx":59CC
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
         Top             =   1410
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
         Height          =   585
         Left            =   90
         Picture         =   "frmRegistrarTraslado.frx":7696
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
         Top             =   810
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
         Left            =   90
         Picture         =   "frmRegistrarTraslado.frx":8360
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
         Top             =   180
         Width           =   555
      End
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   3165
      Left            =   240
      OleObjectBlob   =   "frmRegistrarTraslado.frx":902A
      TabIndex        =   20
      Top             =   3300
      Width           =   11955
   End
   Begin VB.TextBox txtNumReferencia 
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
      Height          =   345
      Left            =   1500
      TabIndex        =   19
      Top             =   2760
      Width           =   1725
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   345
      Left            =   1500
      TabIndex        =   17
      Top             =   1380
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   4210752
      Format          =   97320961
      CurrentDate     =   41787
   End
   Begin VB.TextBox txtEstado 
      BackColor       =   &H8000000F&
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
      Left            =   11160
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   870
      Width           =   1635
   End
   Begin VB.TextBox txtBodegaDestino 
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
      Height          =   345
      Left            =   1500
      TabIndex        =   13
      Top             =   2310
      Width           =   1095
   End
   Begin VB.TextBox txtDescrBodegaDestino 
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
      Height          =   345
      Left            =   3150
      TabIndex        =   12
      Top             =   2310
      Width           =   5715
   End
   Begin VB.CommandButton cmdBodegaDestino 
      Height          =   320
      Left            =   2730
      Picture         =   "frmRegistrarTraslado.frx":D8B7
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2310
      Width           =   300
   End
   Begin VB.TextBox txtBodegaOrigen 
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
      Height          =   345
      Left            =   1500
      TabIndex        =   10
      Top             =   1860
      Width           =   1095
   End
   Begin VB.TextBox txtDescrBodegaOrigen 
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
      Height          =   345
      Left            =   3150
      TabIndex        =   9
      Top             =   1860
      Width           =   5715
   End
   Begin VB.CommandButton cmdBodegaOrigen 
      Height          =   320
      Left            =   2730
      Picture         =   "frmRegistrarTraslado.frx":DBF9
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1860
      Width           =   300
   End
   Begin VB.TextBox txtIDTraslado 
      BackColor       =   &H8000000F&
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
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   930
      Width           =   1635
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
      ScaleWidth      =   8490
      TabIndex        =   0
      Top             =   0
      Width           =   8490
      Begin VB.Image imgCaption 
         Height          =   645
         Left            =   150
         Top             =   60
         Width           =   735
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maestro de Clientes"
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
         TabIndex        =   2
         Top             =   420
         Width           =   1230
      End
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
         TabIndex        =   1
         Top             =   90
         Width           =   855
      End
   End
   Begin Inventario.CtlLiner CtlLiner 
      Height          =   30
      Left            =   0
      TabIndex        =   3
      Top             =   750
      Width           =   19275
      _ExtentX        =   33999
      _ExtentY        =   53
   End
   Begin MSComCtl2.DTPicker dtpFechaSalida 
      Height          =   345
      Left            =   7200
      TabIndex        =   26
      Top             =   1380
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   4210752
      Format          =   97320961
      CurrentDate     =   41787
   End
   Begin VB.Label lblNumSalida 
      Caption         =   "Num Salida:"
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
      Height          =   225
      Left            =   5880
      TabIndex        =   29
      Top             =   2790
      Width           =   1245
   End
   Begin VB.Label lblFechaSalida 
      Caption         =   "Fecha Salida:"
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   5910
      TabIndex        =   27
      Top             =   1440
      Width           =   1065
   End
   Begin VB.Label lblNumReferencia 
      Caption         =   "Num Salida:"
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
      Height          =   225
      Left            =   240
      TabIndex        =   18
      Top             =   2820
      Width           =   1245
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha:"
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
      Height          =   225
      Left            =   210
      TabIndex        =   16
      Top             =   1440
      Width           =   645
   End
   Begin VB.Label Label3 
      Caption         =   "Estado:"
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   10500
      TabIndex        =   15
      Top             =   900
      Width           =   675
   End
   Begin VB.Label Label2 
      Caption         =   "Bodega Destino:"
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
      Height          =   225
      Left            =   210
      TabIndex        =   7
      Top             =   2370
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "Bodega Origen:"
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
      Height          =   225
      Left            =   210
      TabIndex        =   6
      Top             =   1920
      Width           =   1245
   End
   Begin VB.Label lbl 
      Caption         =   "IDTrasaldo:"
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
      Height          =   225
      Left            =   210
      TabIndex        =   5
      Top             =   930
      Width           =   885
   End
End
Attribute VB_Name = "frmRegistrarTraslado"
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

Dim sMensajeError As String
Dim bIsAutoSugiereLotes As Boolean
Dim dTotalSugeridoLotes As Double

Private rstCabecera As ADODB.Recordset
Private rstDetalle As ADODB.Recordset
Dim rstLote As ADODB.Recordset


Dim gTrans As Boolean ' se dispara si hubo error en medio de la transacción
Dim gBeginTransNoEnd As Boolean ' Indica si hubo un begin sin rollback o commit
Dim gTasaCambio As Double
Dim tDatosDelProducto As typDatosProductos
Dim sIdStatus As String
Public scontrol As String
Dim sTipoMovSalida As String
Dim sTipoMovEntrada As String
Dim sPaqueteTraslado As String
Public sAccion As String
Public sDocumentoTraslado As String

Private Sub InicializaControles()
    dtpFecha.value = Format(Now, "YYYY/MM/DD")
    txtIDTraslado.Text = "TP000000000000?"
    fmtTextbox txtIDTraslado, "O"
    
    txtBodegaOrigen.Text = ""
    fmtTextbox txtBodegaOrigen, "O"
    txtDescrBodegaOrigen.Text = ""
    fmtTextbox txtDescrBodegaOrigen, "R"
    
    txtBodegaDestino.Text = ""
    fmtTextbox txtBodegaDestino, "O"
    txtDescrBodegaDestino.Text = ""
    fmtTextbox txtDescrBodegaDestino, "R"
    
    txtNumReferencia.Text = ""
    fmtTextbox txtNumReferencia, "O"
    
    sIdStatus = "16-3"
    sTipoMovSalida = "9"
    sTipoMovEntrada = "10"
    sPaqueteTraslado = "7"
    Me.txtEstado.Text = "Pendiente"
    
    Me.TDBG.Columns(6).Visible = False
    Me.TDBG.Columns(1).Width = 5600
    
    
    Me.lblFechaSalida.Visible = False
    Me.dtpFechaSalida.Visible = False
    
    Me.lblCantidadRemitida.Visible = False
    Me.txtCantidadRemitida.Visible = False
    Me.lblNumReferencia.Caption = "Num Salida"
    Me.lblNumSalida.Visible = False
    Me.txtNumSalida.Visible = False
    Me.imgCaption.Picture = Me.imgTitle.ListImages(2).Picture
End Sub

Private Sub HabilitarBotones()
    Select Case Accion
        Case TypAccion.Add
            cmdSave.Enabled = True
            cmdUndo.Enabled = False
            If (sAccion = "Salida") Then
                cmdEliminar.Enabled = True
            Else
                cmdEliminar.Enabled = False
            End If
            cmdAdd.Enabled = True
            cmdEditItem.Enabled = True
        Case TypAccion.Edit
            cmdSave.Enabled = False
            cmdUndo.Enabled = True
            cmdEliminar.Enabled = False
            cmdAdd.Enabled = True
            cmdEditItem.Enabled = False
        Case TypAccion.View
           If rstDetalle.State = adStateClosed Then
                cmdEditItem.Enabled = False
                cmdSave.Enabled = False
                cmdUndo.Enabled = False
                cmdEliminar.Enabled = False
                cmdAdd.Enabled = True
                Exit Sub
            End If
            If rstDetalle.RecordCount <> 0 Then
                cmdSave.Enabled = False
                cmdUndo.Enabled = False
                If (sAccion = "Salida") Then
                    cmdEliminar.Enabled = True
                Else
                    cmdEliminar.Enabled = False
                End If
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
    'ActivarAccionesByTransacciones
End Sub

Public Sub HabilitarControles()

    Select Case Accion

        Case TypAccion.Add
                        
            If sAccion = "Salida" Then
                fmtTextbox Me.txtBodegaDestino, "O"
                fmtTextbox Me.txtDescrBodegaDestino, "R"
                
                fmtTextbox Me.txtBodegaOrigen, "O"
                fmtTextbox Me.txtDescrBodegaOrigen, "R"
                
                fmtTextbox Me.txtBodegaOrigen, "O"
                fmtTextbox Me.txtDescrBodegaOrigen, "R"
            ElseIf sAccion = "Entrada" Then
                fmtTextbox Me.txtBodegaDestino, "R"
                fmtTextbox Me.txtDescrBodegaDestino, "R"
                
                fmtTextbox Me.txtBodegaOrigen, "R"
                fmtTextbox Me.txtDescrBodegaOrigen, "R"
                
                fmtTextbox Me.txtBodegaOrigen, "R"
                fmtTextbox Me.txtDescrBodegaOrigen, "R"
                
                fmtTextbox Me.txtNumSalida, "R"
                
                Me.dtpFechaSalida.Enabled = False
            End If

            txtIDProducto.Text = ""
            fmtTextbox txtIDProducto, "O"
            txtDescrProducto.Text = ""
            fmtTextbox txtDescrProducto, "R"
            
            txtIDLote.Text = ""
            fmtTextbox txtIDLote, "O"
            txtDescrLote.Text = ""
            fmtTextbox txtDescrLote, "R"
            
            txtCantidad.Text = ""
            fmtTextbox txtCantidad, "O"
            
            If (sAccion = "Entrada") Then
                txtCantidadRemitida.Text = ""
                fmtTextbox txtCantidadRemitida, "R"
            End If
                       
            Me.cmdBodegaDestino.Enabled = True
            Me.cmdBodegaOrigen.Enabled = True
            Me.cmdProducto.Enabled = True
            Me.cmdLote.Enabled = True
            Me.TDBG.Enabled = True
            
            Me.chkAutoSugiereLotes.value = vbChecked
            HabilitarAutoSugerirLotes Me.chkAutoSugiereLotes.value
            
        Case TypAccion.Edit
            
            If sAccion = "Salida" Then
                fmtTextbox Me.txtBodegaDestino, "O"
                fmtTextbox Me.txtDescrBodegaDestino, "R"
                
                fmtTextbox Me.txtBodegaOrigen, "O"
                fmtTextbox Me.txtDescrBodegaOrigen, "R"
                
                fmtTextbox Me.txtBodegaOrigen, "O"
                fmtTextbox Me.txtDescrBodegaOrigen, "R"
            ElseIf sAccion = "Entrada" Then
                fmtTextbox Me.txtBodegaDestino, "R"
                fmtTextbox Me.txtDescrBodegaDestino, "R"
                
                fmtTextbox Me.txtBodegaOrigen, "R"
                fmtTextbox Me.txtDescrBodegaOrigen, "R"
                
                fmtTextbox Me.txtBodegaOrigen, "R"
                fmtTextbox Me.txtDescrBodegaOrigen, "R"
                
                fmtTextbox Me.txtNumSalida, "R"
                
                Me.dtpFechaSalida.Enabled = False
            End If
            
            fmtTextbox txtIDProducto, "O"
            fmtTextbox Me.txtDescrProducto, "R"
            
            fmtTextbox txtIDLote, "O"
            fmtTextbox txtDescrLote, "R"
            
            fmtTextbox txtCantidad, "O"
            
            Me.cmdProducto.Enabled = False
            Me.cmdLote.Enabled = False
            Me.TDBG.Enabled = False
            
        Case TypAccion.View
           
            fmtTextbox Me.txtBodegaDestino, "R"
            fmtTextbox Me.txtDescrBodegaDestino, "R"
                
            fmtTextbox Me.txtBodegaOrigen, "R"
            fmtTextbox Me.txtDescrBodegaOrigen, "R"
                
            fmtTextbox Me.txtBodegaOrigen, "R"
            fmtTextbox Me.txtDescrBodegaOrigen, "R"
                
            fmtTextbox Me.txtNumSalida, "R"
            fmtTextbox Me.txtNumReferencia, "R"
            
            Me.dtpFecha.Enabled = False
            Me.dtpFechaSalida.Enabled = False
          
            fmtTextbox txtIDProducto, "O"
            txtIDLote.Text = ""
            fmtTextbox txtIDLote, "O"
            txtCantidad.Text = ""
            fmtTextbox txtCantidad, "O"
            
            Me.cmdProducto.Enabled = False
            Me.cmdLote.Enabled = False
           
            If (sAccion = "Entrada") Then
                txtCantidadRemitida.Text = ""
                fmtTextbox txtCantidadRemitida, "R"
            End If
           
            Me.TDBG.Enabled = True
    End Select

End Sub

Private Function ValCtrlsCabecera() As Boolean
    Dim Valida As Boolean
    Valida = True
    If (Me.txtBodegaOrigen.Text = "") Then
        sMensajeError = "Por favor selecciona la Bodega Origen..."
        Valida = False
    ElseIf (Me.txtBodegaDestino.Text = "") Then
        sMensajeError = "Por favor seleccione la Bodega Destino..."
        Valida = False
    ElseIf (Me.txtNumReferencia.Text = "") Then
        sMensajeError = "Por favor digite el numero de salida..."
        Valida = False
'    ElseIf (rstDetalle.RecordCount = 0) Then
'        sMensajeError = "La transación debe de tener al menos un registro en su detalle..."
'        Valida = False
    End If
    ValCtrlsCabecera = Valida
End Function

Private Function ValCtrlsDetalle() As Boolean
    Dim Valida As Boolean
    Valida = True
    If (Me.txtIDProducto.Text = "") Then
        sMensajeError = "Por favor seleccione el producto..."
        Valida = False
    ElseIf (Me.txtIDLote.Text = "" And Me.chkAutoSugiereLotes.value = False) Then
        sMensajeError = "Por favor seleccione el lote del producto..."
        Valida = False
    ElseIf (Me.txtCantidad.Text = "") Then
        sMensajeError = "Por favor digite la cantidad del traslado..."
        Valida = False
    End If
    ValCtrlsDetalle = Valida
End Function

Private Sub chkAutoSugiereLotes_Click()
      HabilitarAutoSugerirLotes Me.chkAutoSugiereLotes.value
End Sub

Private Sub HabilitarAutoSugerirLotes(IsAutoSugiereLotes As Boolean)
    If IsAutoSugiereLotes = True Then
        Me.txtIDLote.Enabled = False
        Me.txtDescrLote.Enabled = False
        Me.cmdLote.Enabled = False
        'Me.cmdDelLote.Enabled = False
        bIsAutoSugiereLotes = True
    Else
        If (Accion = Add) Then
            Me.txtIDLote.Enabled = True
            Me.txtDescrLote.Enabled = True
            Me.cmdLote.Enabled = True
            'Me.cmdDelLote.Enabled = True
        End If
        bIsAutoSugiereLotes = False
    End If
End Sub


Private Sub cmdAdd_Click()
   Dim lbok As Boolean
    Dim dicDatosExistencia As Dictionary
    If Not ValCtrlsCabecera Then
        lbok = Mensaje("Revise sus datos por favor !!! " & sMensajeError, ICO_ERROR, False)
        Exit Sub
    End If
   
    If Not ValCtrlsDetalle Then
        lbok = Mensaje("Revise sus datos por favor !!! " & sMensajeError, ICO_ERROR, False)
        Exit Sub
    End If
    
    
    If (Accion = Add And sAccion = "Salida") Then
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
            frmAutosugiere.gsIDProducto = Me.txtIDProducto.Text
            frmAutosugiere.gsDescrProducto = Me.txtDescrProducto.Text
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
                    If ExiteRstKey(rstDetalle, "IDPRODUCTO=" & Me.txtIDProducto.Text & _
                                                " AND IDLOTE=" & rstLS!IdLote) Then
                        lbok = Mensaje("Ya existe ese el registro en la transacción", ICO_ERROR, False)
                        Exit Sub
                    End If
                    Set rstLote = New ADODB.Recordset
                      rstLote.ActiveConnection = gConet
                    CargaDatosLotes rstLote, CInt(rstLS!IdLote)
                    ' Carga los datos del detalle de transacciones para ser grabados a la bd
                    rstDetalle.AddNew
                    rstDetalle!IdProducto = Me.txtIDProducto.Text
                    rstDetalle!DescrProducto = Me.txtDescrProducto.Text
                    'Pendiente: Aplicar los dos campos siguientes solo para traslados
                    rstDetalle!IdLote = rstLS!IdLote
                    rstDetalle!LoteInterno = rstLote!LoteInterno
                    rstDetalle!FechaVencimiento = rstLote!FechaVencimiento
                    rstDetalle!Cantidad = rstLS!Cantidad
                    rstDetalle.Update
                    rstDetalle.MoveFirst
                      
                    rstLS.MoveNext
                Wend
            End If
        Else
                
             If ExiteRstKey(rstDetalle, "IDPRODUCTO=" & Me.txtIDProducto.Text & _
                                                " AND IDLOTE=" & Me.txtIDLote.Text) Then
              lbok = Mensaje("Ya existe ese el registro en la transacción", ICO_ERROR, False)
        
              Exit Sub
            End If
            Set rstLote = New ADODB.Recordset
              rstLote.ActiveConnection = gConet
            CargaDatosLotes rstLote, CInt(Trim(Me.txtIDLote.Text))
            'Verificar existenciaas suficientes
            If (getValueFieldsFromTable("invEXISTENCIALOTE", "Existencia", "IdProducto=" & Me.txtIDProducto.Text & " and  IDBodega='" & Me.txtBodegaOrigen.Text & "'", dicDatosExistencia) = True) Then
                 If (CDbl(dicDatosExistencia("Existencia")) < CDbl(Me.txtCantidad.Text)) Then
                    lbok = Mensaje("No hay suficiente existencia para satisfacer la transacción.", ICO_ERROR, False)
                    Me.txtCantidad.SetFocus
                    Exit Sub
                 End If
               
            End If
            ' Carga los datos del detalle de transacciones para ser grabados a la bd
            rstDetalle.AddNew
            rstDetalle!IdProducto = Me.txtIDProducto.Text
            rstDetalle!DescrProducto = Me.txtDescrProducto.Text
            rstDetalle!IdLote = Me.txtIDLote.Text
            rstDetalle!LoteInterno = Me.txtDescrLote.Text
            rstDetalle!FechaVencimiento = rstLote!FechaVencimiento
            rstDetalle!Cantidad = Val(Me.txtCantidad.Text)
            
            rstDetalle.Update
            rstDetalle.MoveFirst
        End If
    ElseIf (Accion = Edit) Then
        ' Actualiza el rst temporal
            rstDetalle!IdProducto = Me.txtIDProducto.Text
            rstDetalle!DescrProducto = Me.txtDescrProducto.Text
            rstDetalle!IdLote = Me.txtIDLote.Text
            rstDetalle!LoteInterno = Me.txtDescrLote.Text
            'rstDetalle!FechaVencimiento = rstLote!FechaVencimiento
            If (sAccion = "Salida") Then
                rstDetalle!Cantidad = Val(Me.txtCantidad.Text)
                rstDetalle!CantidadRecibida = 0
                rstDetalle!Ajuste = 0
                rstDetalle!RecibidoParcial = 0
                rstDetalle!RecibidoTotal = 0
            Else
                rstDetalle!CantidadRecibida = Val(Me.txtCantidad.Text)
                rstDetalle!Cantidad = Val(Me.txtCantidadRemitida.Text)
                rstDetalle!Ajuste = rstDetalle!Cantidad - rstDetalle!CantidadRecibida
                rstDetalle!RecibidoParcial = IIf(rstDetalle!CantidadRecibida <> rstDetalle!Cantidad, "1", "0")
                rstDetalle!RecibidoTotal = IIf(rstDetalle!CantidadRecibida <> rstDetalle!Cantidad, "0", "1")
            End If
            
    End If
   
    Me.cmdSave.Enabled = True
    
    Set TDBG.DataSource = rstDetalle
    TDBG.ReBind
             
    Accion = Add
      'Me.dtgAjuste.Columns("Descr").FooterText = "Items de la transacción :     " & rstTransAI.RecordCount
HabilitarControles
HabilitarBotones
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
      'Validar si la bodega seleccionada es diferente a la bodega origen
      If (Me.txtBodegaDestino.Text = Me.txtBodegaOrigen.Text) Then
         lbok = Mensaje("No puede seleccionar la misma bodega para el traslado, por favor verifique?", ICO_INFORMACION, True)
         Me.txtBodegaDestino.Text = ""
         Exit Sub
      End If
      
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
      If (Me.txtBodegaDestino.Text = Me.txtBodegaOrigen.Text) Then
         lbok = Mensaje("No puede seleccionar la misma bodega para el traslado, por favor verifique?", ICO_INFORMACION, True)
         Me.txtBodegaOrigen.Text = ""
         Exit Sub
      End If
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      Me.txtDescrBodegaOrigen.Text = frm.gsDescrbrw
      fmtTextbox txtDescrBodegaOrigen, "R"
    End If
End Sub

Private Sub cmdEditItem_Click()
    Accion = Edit
    GetDataFromGridToControl
    HabilitarBotones
    HabilitarControles
   
End Sub

Private Sub GetDataFromGridToControl() 'EDITAR
'
    If Not (rstDetalle.EOF And rstDetalle.BOF) Then
        Me.txtIDProducto.Text = rstDetalle("IDProducto").value
        Me.txtDescrProducto.Text = rstDetalle("DescrProducto").value
        'Contemplar para traslados
        Me.txtIDLote.Text = rstDetalle("IDLote").value
        Me.txtDescrLote.Text = rstDetalle("LoteInterno").value
        If (Me.sAccion = "Salida") Then
            Me.txtCantidad.Text = rstDetalle("Cantidad").value
        Else
            Me.txtCantidad.Text = rstDetalle("CantidadRecibida").value
            Me.txtCantidadRemitida.Text = rstDetalle("Cantidad").value
        End If
    Else
      
        HabilitarControles
    End If

End Sub


Private Sub cmdEliminar_Click()
Dim lbok As Boolean
    
    lbok = Mensaje("Esta seguro que desea eliminar el registro seleccionado?", ICO_INFORMACION, True)
    If (lbok) Then
        rstDetalle.Delete
        Accion = Add
        HabilitarBotones
        HabilitarControles
        TDBG.ReBind
    End If
End Sub

Private Sub cmdLote_Click()
    Dim frm As frmBrowseCat

    Set frm = New frmBrowseCat
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
    frm.gsFiltro = "IDBodega=" & Me.txtBodegaOrigen.Text & " and IDProducto=" & Me.txtIDProducto.Text & " and Existencia>0"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
        txtIDLote.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      Me.txtDescrLote.Text = frm.gsDescrbrw
      fmtTextbox Me.txtIDLote, "R"
      'txtExistenciaLote.Text = frm.gsExtraValor2
    End If


'    Dim frm As New frmBrowseCat
'
'    frm.gsCaptionfrm = "Lote de Productos"
'    frm.gsTablabrw = "invLOTE"
'    frm.gsCodigobrw = "IdLote"
'    frm.gbTypeCodeStr = True
'    frm.gsDescrbrw = "LoteInterno"
'    frm.gbFiltra = False
'    frm.gsNombrePantallaExtra = "frmMasterLotes"
'    'frm.gsFiltro = "IdPaquete='" & Me.gsIDTipoTransaccion & "'"
'    frm.Show vbModal
'    If frm.gsCodigobrw <> "" Then
'      Me.txtIDLote.Text = frm.gsCodigobrw
'
'    End If
'
'    If frm.gsDescrbrw <> "" Then
'      Me.txtDescrLote.Text = frm.gsDescrbrw
'      fmtTextbox Me.txtDescrLote, "R"
'    End If
End Sub

Private Sub cmdProducto_Click()
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
      Me.txtIDProducto.Text = frm.gsCodigobrw
      'Traer el costo promedio del producto
        If (getValueFieldsFromTable("invPRODUCTO", "CostoUltLocal,CostoUltDolar,CostoUltPromLocal,CostoUltPromDolar", "IdProducto=" & Me.txtIDProducto.Text, dicDatosProducto) = True) Then
            tDatosDelProducto.CostoPromDolar = CDbl(dicDatosProducto("CostUlPromDolar"))
            tDatosDelProducto.CostoPromLocal = CDbl(dicDatosProducto("CostoUltPromLocal"))
            tDatosDelProducto.CostoUltDolar = CDbl((dicDatosProducto("CostoUltDolar")))
            tDatosDelProducto.CostoUltLocal = CDbl((dicDatosProducto("CostoUltLocal")))
           
        End If
    End If
    
    If frm.gsDescrbrw <> "" Then
      Me.txtDescrProducto.Text = frm.gsDescrbrw
      fmtTextbox Me.txtDescrProducto, "R"
    
      If (Me.chkAutoSugiereLotes.value = vbChecked) Then
        Me.cmdLote.Enabled = False
         Me.txtIDLote.Enabled = False
      Else
        Me.cmdLote.Enabled = True
        Me.txtIDLote.Enabled = True
      End If
    End If
End Sub

Private Function invSaveCabeceraTraslado(sOperacion As String) As String
  
    Dim lbok As Boolean
    On Error GoTo errores
    lbok = False
    Dim sDocumento As String
    Dim rst As ADODB.Recordset
    
    Dim sIDStatusRecibido As String
    Dim sFechaRemision As String
    Dim sFechaEntrada As String
    Dim sNumEntrada As String
    Dim sNumSalida As String
    Dim sDocumentoAjuste As String
    Dim sAplicado As String
    
    If (sOperacion = "I") Then
        sIDStatusRecibido = "16-3"
    Else
        sIDStatusRecibido = "16-1"
    End If
    If (sAccion = "Salida") Then
        sFechaRemision = Format(Str(Me.dtpFecha.value), "yyyymmdd")
        sFechaEntrada = Format("1980-01-01", "yyyymmdd")
        sNumEntrada = ""
        sNumSalida = Me.txtNumReferencia.Text
        sDocumentoAjuste = ""
        sAplicado = "0"
        sDocumentoTraslado = ""
    Else
        sFechaRemision = Format(Str(Me.dtpFechaSalida.value), "yyyymmdd")
        sFechaEntrada = Format(Str(Me.dtpFecha.value), "yyyymmdd")
        sNumEntrada = Me.txtNumReferencia.Text
        sNumSalida = Me.txtNumSalida.Text
        sDocumentoAjuste = ""
        sAplicado = "1"
    End If
    GSSQL = "invUpdateCabTraslados '" & sOperacion & "','" & sDocumentoTraslado & "','" & Me.txtBodegaOrigen.Text & "','" & Me.txtBodegaDestino.Text & "','" & sIDStatusRecibido & "','" & sFechaRemision & _
                "','" & sFechaEntrada & "','" & sNumEntrada & "','" & sNumSalida & "','" & sDocumentoAjuste & "'," & sAplicado
 Set rst = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  
  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    sDocumento = "" 'Indica que ocurrió un error
    sMensajeError = "Ha ocurrido un error tratando de ingresar la cabecera del traslado!!!" & err.Description
  Else  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    If (sAccion = "Salida") Then
        sDocumento = rst("IDTraslado").value
    Else
        sDocumento = sDocumentoTraslado
    End If
  End If
  invSaveCabeceraTraslado = sDocumento

    Exit Function
errores:
    gTrans = False
    invSaveCabeceraTraslado = ""
    'gConet.RollbackTrans
    Exit Function
End Function


Public Function invUpdateDetalleTraslados(sOperacion As String, sIDTraslado As String, sBodegaOrigen As String, _
    sBodegaDestino As String, sIDProducto As String, sIDLote As String, sCantidad As String, sCantidadRecibida As String, _
    sAjuste As String, sRecibidoParcial As String, sRecibidoTotal As String) As Boolean
    Dim lbok As Boolean
   
    
    lbok = True
    
      GSSQL = ""
      GSSQL = gsCompania & ".invUpdateDetalleTraslados '" & sOperacion & "','" & sIDTraslado & "'," & sBodegaOrigen & "," & sBodegaDestino & "," & sIDProducto & "," & sIDLote & "," & sCantidad & ","
      GSSQL = GSSQL & sCantidadRecibida & "," & sAjuste & "," & sRecibidoParcial & "," & sRecibidoTotal
        
     gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords   'Ejecuta la sentencia
    
        If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
          'gsOperacionError = "Eliminando el Beneficiado. " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & err.Description
          SetMsgError "Ocurrió un error insertando la transacción . ", err
          lbok = False
        End If
    
    invUpdateDetalleTraslados = lbok
    Exit Function
    

End Function

Private Sub SaveRstDetalle(rst As ADODB.Recordset, sDocumento As String, sOperacion As String)
    On Error GoTo errores
    'Set lRegistros = New ADODB.Recordset  'Inicializa la variable de los registros
    'gConet.BeginTrans
    Dim bOk  As Boolean
    bOk = True
    
    If rst.RecordCount > 0 Then
      rst.MoveFirst
      While Not rst.EOF And bOk
     
            
            bOk = invUpdateDetalleTraslados(sOperacion, _
                                    sDocumento, _
                                    Me.txtBodegaOrigen.Text, _
                                    Me.txtBodegaDestino.Text, _
                                    rst.Fields("IDProducto").value, _
                                    rst.Fields("IDLote").value, _
                                    rst.Fields("Cantidad").value, _
                                    rst.Fields("CantidadRecibida").value, _
                                    rst.Fields("Ajuste").value, _
                                    CInt(rst.Fields("RecibidoParcial").value), _
                                    CInt(rst.Fields("RecibidoTotal").value))
                                

            rst.MoveNext
      Wend
      
      rst.MoveFirst

    
    End If
    
    Exit Sub
errores:
    gTrans = False
    'gConet.RollbackTrans 'Descomentarie esto
End Sub

Private Function invGeneraCabeceraTraslado(sDocumento As String) As String
    Dim lbok As Boolean
    On Error GoTo errores
    lbok = False
   
    Dim rst As ADODB.Recordset

    GSSQL = "invInsertCabMovimientos " & sPaqueteTraslado & ",'" & sDocumento & "','" & Format(Str(dtpFecha.value), "yyyymmdd") & _
                "','Traslado de Bodega " & Me.txtDescrBodegaOrigen.Text & " a bodega " & Me.txtDescrBodegaDestino.Text & "','" & sDocumento & "','" & gsUser & "','" & gsUser & "',0"
 Set rst = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    sDocumento = "" 'Indica que ocurrió un error
    sMensajeError = "Ha ocurrido un error tratando de grabar la cabecera de la transacción !!!" & err.Description
  Else  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    sDocumento = rst("Documento").value
  End If
  invGeneraCabeceraTraslado = sDocumento

    Exit Function
errores:
    gTrans = False
    invGeneraCabeceraTraslado = ""
    'gConet.RollbackTrans
    Exit Function
End Function


Private Function invGeneraAjusteByTraslado(sIDTraslado As String, sUsuario As String) As Boolean
    Dim lbok As Boolean
    On Error GoTo errores
    lbok = False

    Dim rst As ADODB.Recordset

    GSSQL = "invGeneraAjusteByTraslado '" & sIDTraslado & "','" & sUsuario & "'"
    Set rst = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    'sDocumento = "" 'Indica que ocurrió un error
        sMensajeError = "Ha ocurrido un error tratando de guardar el detalle del movimiento!!!" & err.Description
 
    End If
    invGeneraAjusteByTraslado = lbok

    Exit Function
errores:
    gTrans = False
    invGeneraAjusteByTraslado = False
    'gConet.RollbackTrans
    Exit Function
End Function


Private Function invGeneraDetalleMovimientoTraslado(sIDTraslado As String, sIsSalida As String, sUserInsert As String, sUserUpdate As String) As Boolean
    Dim lbok As Boolean
    On Error GoTo errores
    lbok = False

    Dim rst As ADODB.Recordset

    GSSQL = "invGeneraDetalleMovimientoTraslado '" & sIDTraslado & "'," & sIsSalida & ",'" & sUserInsert & "','" & sUserUpdate & "'"
 Set rst = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    'sDocumento = "" 'Indica que ocurrió un error
    sMensajeError = "Ha ocurrido un error tratando de guardar el detalle del movimiento!!!" & err.Description
 
  End If
  invGeneraDetalleMovimientoTraslado = lbok

    Exit Function
errores:
    gTrans = False
    invGeneraDetalleMovimientoTraslado = False
    'gConet.RollbackTrans
    Exit Function
End Function


Private Function ValidaEntradaParcial(rst As ADODB.Recordset) As Boolean
     On Error GoTo errores
    'Set lRegistros = New ADODB.Recordset  'Inicializa la variable de los registros
    'gConet.BeginTrans
    Dim bOk  As Boolean
    bOk = True
    If rst.RecordCount > 0 Then
      rst.MoveFirst
      While Not rst.EOF And bOk
            If (rst!CantidadRecibida < rst!Cantidad) Then
                ValidaEntradaParcial = True
                Exit Sub
            rst.MoveNext
      Wend
      rst.MoveFirst
    End If
    Exit Function
errores:
    gTrans = False
End Function

Private Sub cmdSave_Click()
    Dim lbok As Boolean
    'On Error GoTo errores
    Dim sOperacion As String
    If Not ValCtrlsCabecera Then
        lbok = Mensaje("Revise sus datos por favor !!! " & sMensajeError, ICO_ERROR, False)
        Exit Sub
    End If
    
    If rstDetalle.RecordCount > 0 Then
        gConet.BeginTrans ' inicio aqui la transacción
        gBeginTransNoEnd = True
        
        Dim sDocumento As String
        sOperacion = IIf(sAccion = "Salida", "I", "U")
        sDocumento = invSaveCabeceraTraslado(sOperacion)
        If sDocumento <> "" Then ' salva la cabecera
            SaveRstDetalle rstDetalle, sDocumento, sOperacion ' salva el detalle que esta en batch
            If (gTrans = True) And sAccion = "Salida" Then
              invGeneraCabeceraTraslado sDocumento
            End If
            If (gTrans = True) Then
                'invMasterAcutalizaSaldosInventarioPaquete sDocumento, gsIDTipoTransaccion, Me.gsIDTipoTransaccion, gsUser
                invGeneraDetalleMovimientoTraslado sDocumento, IIf(sAccion = "Salida", "1", "0"), gsUser, gsUser
            End If
            If (gTrans = True And sAccion = "Entrada" And ValidaEntradaParcial(rstDetalle) = True) Then
                'Generar los ajustes correspondientes
                invGeneraAjusteByTraslado sDocumento, gsUser
            End If
            If (gTrans = True) Then
                lbok = Mensaje("La transacción ha sido guardada exitosamente", ICO_OK, False)
           
                Me.cmdAdd.Enabled = False
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
    Else
        gTrans = False
    
        If gBeginTransNoEnd Then gConet.RollbackTrans
        gBeginTransNoEnd = False
    End If
    lbok = Mensaje("Hubo un error en el proceso de salvado " & Chr(13) & err.Description, ICO_ERROR, False)
End Sub


    '      If gTasaCambio = 0 Then
    '        lbOk = Mensaje("La tasa de cambio es Cero llame a informática por favor ", ICO_ERROR, False)
    '        Exit Sub
    '      End If
       
       
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

Private Sub cmdUndo_Click()
    GetDataFromGridToControl
    Accion = Add
    HabilitarBotones
    HabilitarControles
End Sub

Private Sub InicializaControlesEntrada()
    GetCabeceraTraslado sDocumentoTraslado
    GetDetalleTrasladoByDocumento sDocumentoTraslado
    Me.TDBG.Columns(6).Visible = True
    Me.TDBG.Columns(1).Width = 4050.142
    Me.TDBG.Columns(5).Caption = "Cant Remitida"
    Me.lblFechaSalida.Visible = True
    Me.dtpFechaSalida.Visible = True
    Me.lblCantidadRemitida.Visible = True
    Me.txtCantidadRemitida.Visible = True
    fmtTextbox Me.txtCantidadRemitida, "R"
    Me.lblNumReferencia.Caption = "Num Entrada"
    Me.lblNumSalida.Visible = True
    Me.txtNumSalida.Visible = True
    Me.imgCaption.Picture = Me.imgTitle.ListImages(1).Picture
End Sub

Private Sub Form_Activate()
    HighlightInWin Me.Name
    SetupFormToolbar (Me.Name)
End Sub

Private Sub Form_Load()
    Set rstDetalle = New ADODB.Recordset
    If rstDetalle.State = adStateOpen Then rst.Close
    rstDetalle.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rstDetalle.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rstDetalle.CursorLocation = adUseClient ' Cursor local al cliente
    rstDetalle.LockType = adLockOptimistic
    
    
    gTasaCambio = 25.6
    
    Me.Caption = gsFormCaption
    Me.lbFormCaption = gsTitle
    Accion = Add
    
    'gTasaCambio = GetTasadeCambio(Format(Now, "YYYY/MM/DD"))
    If sAccion = "Salida" Then
        PreparaRstDetalle ' Prepara los Recordsets
        Set Me.TDBG.DataSource = rstDetalle
        Me.TDBG.Refresh
        InicializaControles
    ElseIf sAccion = "Entrada" Then
        InicializaControlesEntrada
        Set Me.TDBG.DataSource = rstDetalle
        Me.TDBG.Refresh
        EnlazarControles
    ElseIf sAccion = "View" Then
        InicializaControlesEntrada
        Set Me.TDBG.DataSource = rstDetalle
        Me.TDBG.Refresh
        EnlazarControles
        Frame1.Visible = False
        Me.frmDetalle.Visible = False
        Accion = View
    End If
   
    'SetTextBoxReadOnly
    
    HabilitarBotones
    HabilitarControles
    Me.chkAutoSugiereLotes.value = vbChecked
    HabilitarAutoSugerirLotes Me.chkAutoSugiereLotes.value
End Sub

Private Sub EnlazarControles()
    'Enlazar Datos de Cabecera
     If Not (rstCabecera.EOF And rstCabecera.BOF) Then
        Me.txtIDTraslado.Text = rstCabecera!IDTraslado
        Me.txtBodegaOrigen.Text = rstCabecera!BodegaOrigen
        Me.txtDescrBodegaOrigen.Text = rstCabecera!DescrBodegaOrigen
        Me.txtBodegaDestino.Text = rstCabecera!BodegaDestino
        Me.txtDescrBodegaDestino.Text = rstCabecera!DescrBodegaDestino
        Me.dtpFecha.value = rstCabecera!FechaRemision
        Me.txtEstado.Text = rstCabecera!DescrStatusRecibido
        Me.txtNumSalida.Text = rstCabecera!NumSalida
     End If
End Sub

Private Sub GetCabeceraTraslado(sDocumento As String)
      ' preparacion del recordset fuente del grid de movimientos
      
      Set rstCabecera = New ADODB.Recordset
      If rstCabecera.State = adStateOpen Then rstCabecera.Close
      rstCabecera.ActiveConnection = gConet 'Asocia la conexión de trabajo
      rstCabecera.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
      rstCabecera.CursorLocation = adUseClient ' Cursor local al cliente
      rstCabecera.LockType = adLockOptimistic
                     
      If rstCabecera.State = adStateOpen Then rstCabecera.Close
      GSSQL = "invGetCabTraslados '" & sDocumento & "'"
      
      gTrans = True ' asume que NO va a haber un error en la transacción
      Set rstCabecera = GetRecordset(GSSQL) ' para el detalle
End Sub


Private Sub GetDetalleTrasladoByDocumento(sDocumento As String)
      ' preparacion del recordset fuente del grid de movimientos
      
      Set rstDetalle = New ADODB.Recordset
      If rstDetalle.State = adStateOpen Then rstDetalle.Close
      rstDetalle.ActiveConnection = gConet 'Asocia la conexión de trabajo
      rstDetalle.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
      rstDetalle.CursorLocation = adUseClient ' Cursor local al cliente
      rstDetalle.LockType = adLockOptimistic
                     
      If rstDetalle.State = adStateOpen Then rstDetalle.Close
      GSSQL = "invGetDetalleTraslados '" & sDocumento & "'"
      
      gTrans = True ' asume que NO va a haber un error en la transacción
      Set rstDetalle = GetRecordset(GSSQL) ' para el detalle
End Sub


Private Sub PreparaRstDetalle()
      ' preparacion del recordset fuente del grid de movimientos
      
      Set rstDetalle = New ADODB.Recordset
      If rstDetalle.State = adStateOpen Then rstDetalle.Close
      rstDetalle.ActiveConnection = gConet 'Asocia la conexión de trabajo
      rstDetalle.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
      rstDetalle.CursorLocation = adUseClient ' Cursor local al cliente
      rstDetalle.LockType = adLockOptimistic
                     
      If rstDetalle.State = adStateOpen Then rstDetalle.Close
      GSSQL = "invPreparaDetalleTraslados"
      
      gTrans = True ' asume que NO va a haber un error en la transacción
      Set rstDetalle = GetRecordset(GSSQL) ' para el detalle
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

