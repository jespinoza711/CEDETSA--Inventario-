VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmDetPedido 
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   11400
      TabIndex        =   4
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
         TabIndex        =   6
         Top             =   90
         Width           =   855
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle del Pedido.."
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
         TabIndex        =   5
         Top             =   420
         Width           =   1215
      End
      Begin VB.Image Image 
         Height          =   480
         Index           =   2
         Left            =   210
         Picture         =   "frmDetPedido.frx":0000
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdAprobar 
      Caption         =   "Aprobar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8220
      TabIndex        =   2
      Top             =   930
      Width           =   1695
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "Anular"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   1
      Top             =   930
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9990
      TabIndex        =   0
      Top             =   7020
      Visible         =   0   'False
      Width           =   1695
   End
   Begin TrueOleDBGrid60.TDBGrid TDBGFAC 
      Height          =   5565
      Left            =   300
      OleObjectBlob   =   "frmDetPedido.frx":0C44
      TabIndex        =   3
      Top             =   1320
      Width           =   11355
   End
   Begin Inventario.CtlLiner CtlLiner 
      Height          =   30
      Left            =   0
      TabIndex        =   7
      Top             =   750
      Width           =   17925
      _ExtentX        =   31618
      _ExtentY        =   53
   End
End
Attribute VB_Name = "frmDetPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsttmpProdFac As ADODB.Recordset
Public gsIDPedido As String
Public gsIDCliente As String
Public gsIDBodega As String
Public gsNombre As String
Public gsFecha As String
Public gbHuboAprobacion As Boolean
Public gbHuboAnulacion As Boolean


Private Sub cmdAnular_Click()
Dim sBackOrder As String
Dim sAprobado As String
Dim lbok As Boolean
lbok = Mensaje("Está seguro que Ud quiere ANULAR el Pedido No. " & gsIDPedido & " del Cliente " & gsNombre, ICO_PREGUNTA, True)
If lbok Then
    If Not (rsttmpProdFac.EOF And rsttmpProdFac.BOF) Then
        rsttmpProdFac.MoveFirst
        If rsttmpProdFac("BackOrder").value = True Then
            sBackOrder = "1"
        Else
            sBackOrder = "0"
        End If
        If rsttmpProdFac("Aprobado").value = True Then
            lbok = Mensaje("Ud. No Puede Anular un Pedido Aprobado", ICO_ERROR, False)
            Exit Sub
            sAprobado = "1"
        Else
            sAprobado = "0"
        End If
    
        lbok = fafUpdatePedido("U", gsIDPedido, gsIDBodega, gsIDCliente, "0", gsFecha, sAprobado, sBackOrder, "1")
        If lbok Then
            lbok = Mensaje("El Pedido ha sido Anulado exitosamente...", ICO_OK, False)
            gbHuboAnulacion = True
            Unload Me
        End If
        
    End If
End If
End Sub

Private Sub cmdAprobar_Click()
Dim sBackOrder As String
Dim sAnulado As String
Dim lbok As Boolean
lbok = Mensaje("Está seguro que Ud quiere APROBAR el Pedido No. " & gsIDPedido & " del Cliente " & gsNombre, ICO_PREGUNTA, True)
If lbok Then
    If Not (rsttmpProdFac.EOF And rsttmpProdFac.BOF) Then
        rsttmpProdFac.MoveFirst
        If rsttmpProdFac("BackOrder").value = True Then
            sBackOrder = "1"
        Else
            sBackOrder = "0"
        End If
        If rsttmpProdFac("Anulado").value = True Then
            lbok = Mensaje("Ud. No Puede Aprobar un Pedido Anulado", ICO_ERROR, False)
            Exit Sub
            sAnulado = "1"
        Else
            sAnulado = "0"
        End If
    
        lbok = fafUpdatePedido("U", gsIDPedido, gsIDBodega, gsIDCliente, "0", gsFecha, "1", sBackOrder, sAnulado)
        If lbok Then
            lbok = Mensaje("El Pedido ha sido Aprobado exitosamente...", ICO_OK, False)
            gbHuboAprobacion = True
            Unload Me
        End If
    End If
End If

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Activate()
HighlightInWin Me.Name
SetupFormToolbar (Me.Name)
End Sub



Private Sub Form_Load()
PreparaRst
SetColumnSizeGrid
gbHuboAnulacion = False
gbHuboAprobacion = False
End Sub
Private Sub PreparaRst()
      ' preparacion del recordset fuente del grid de compra
      ' recordar que este recordset va a ser temporal, no se hara addnew a la bd
      ' lleva además de los campos de la tabla detalle de compra, la descripcion del producto
      Set rsttmpProdFac = New ADODB.Recordset
      If rsttmpProdFac.State = adStateOpen Then rsttmpProdFac.Close
      rsttmpProdFac.ActiveConnection = gConet 'Asocia la conexión de trabajo
      rsttmpProdFac.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
      rsttmpProdFac.CursorLocation = adUseClient ' Cursor local al cliente
      rsttmpProdFac.LockType = adLockOptimistic
        
      GSSQL = gsCompania & ".fafGetDetallePedidoToGrid " & gsIDBodega & "," & gsIDCliente & "," & gsIDPedido
           
      If rsttmpProdFac.State = adStateOpen Then rsttmpProdFac.Close
      Set rsttmpProdFac = GetRecordset(GSSQL)
      Set TDBGFAC.DataSource = rsttmpProdFac
        TDBGFAC.Refresh
      lblTitulo.Caption = " Cliente : " & gscliente & " " & gsNombre & " Pedido # " & gsIDPedido & " Fecha : " & gsFecha
      lblTitulo.Refresh
End Sub
Private Sub SetColumnSizeGrid()
TDBGFAC.Columns("IDProducto").Width = 1110.047
TDBGFAC.Columns("Descr").Width = 5040
TDBGFAC.Columns("PRecio").Width = 1454.74
TDBGFAC.Columns("Cantidad").Width = 1574.929
TDBGFAC.Columns("SubTotal").Width = 1785.26
TDBGFAC.Columns("IMpuesto").Width = 1574.929
TDBGFAC.Columns("Total").Width = 1635.024

End Sub



Private Sub Form_Unload(Cancel As Integer)
    SetupFormToolbar ("no form")
    
    Set frmDetPedido = Nothing
End Sub












