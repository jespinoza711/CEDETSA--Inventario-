VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPedidos 
   Caption         =   "Form1"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14970
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   14970
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Filtro por Número de Pedidos"
      Height          =   735
      Left            =   180
      TabIndex        =   21
      Top             =   120
      Width           =   5175
      Begin VB.TextBox txtPedidoInicial 
         Height          =   285
         Left            =   1320
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtPedidoFinal 
         Height          =   285
         Left            =   3810
         TabIndex        =   22
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label1 
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
         Left            =   360
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
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
         Left            =   2880
         TabIndex        =   24
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtro por Rango de Fechas"
      Height          =   735
      Left            =   5700
      TabIndex        =   16
      Top             =   120
      Width           =   6255
      Begin MSComCtl2.DTPicker DTPDesde 
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   62128129
         CurrentDate     =   41640
      End
      Begin MSComCtl2.DTPicker DTPHasta 
         Height          =   255
         Left            =   4320
         TabIndex        =   18
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   62128129
         CurrentDate     =   47484
      End
      Begin VB.Label Label9 
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
         Left            =   3360
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
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
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Filtro por Cliente y Vendedor"
      Height          =   1215
      Left            =   180
      TabIndex        =   5
      Top             =   960
      Width           =   11775
      Begin VB.CommandButton cmdDelCliente 
         Height          =   320
         Left            =   2580
         Picture         =   "frmPedidos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   300
      End
      Begin VB.CommandButton cmdCliente 
         Height          =   320
         Left            =   2160
         Picture         =   "frmPedidos.frx":1CCA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   300
      End
      Begin VB.TextBox txtCliente 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   1080
         TabIndex        =   11
         Top             =   240
         Width           =   945
      End
      Begin VB.TextBox txtNombre 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2955
         TabIndex        =   10
         Top             =   240
         Width           =   8400
      End
      Begin VB.CommandButton cmdDelVendedor 
         Height          =   320
         Left            =   2580
         Picture         =   "frmPedidos.frx":200C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   720
         Width           =   300
      End
      Begin VB.CommandButton cmdVendedor 
         Height          =   320
         Left            =   2160
         Picture         =   "frmPedidos.frx":3CD6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   720
         Width           =   300
      End
      Begin VB.TextBox txtVendedor 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   720
         Width           =   945
      End
      Begin VB.TextBox txtDescrVendedor 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2955
         TabIndex        =   6
         Top             =   720
         Width           =   8400
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente :"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002F2F2F&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor :"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002F2F2F&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   780
         Width           =   855
      End
   End
   Begin VB.CheckBox chkDesaprobados 
      Caption         =   "Solo Desaprobados ?"
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
      Left            =   12180
      TabIndex        =   4
      Top             =   240
      Width           =   2295
   End
   Begin VB.CheckBox ChkAnuladas 
      Caption         =   "Solo Anulados ?"
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
      Left            =   12180
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
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
      Left            =   12300
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   300
      TabIndex        =   0
      Top             =   6960
      Width           =   14535
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   $"frmPedidos.frx":4018
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   14055
      End
   End
   Begin TrueOleDBGrid60.TDBGrid TDBGFAC 
      Height          =   4095
      Left            =   300
      OleObjectBlob   =   "frmPedidos.frx":40A9
      TabIndex        =   26
      Top             =   2520
      Width           =   14415
   End
End
Attribute VB_Name = "frmPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sPedidoInicial As String
Dim sPedidoFinal As String
Dim sFechaInicial As String
Dim sFechaFinal As String
Dim sCliente As String
Dim sVendedor As String
Dim sDesaprobados As String
Dim sAnulados As String
Dim rsttmpProdFac As ADODB.Recordset ' para la fuente del grid

Private Sub cmdCliente_Click()
   Dim frm As frmBrowseCat

    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = "Clientes"
    frm.gsTablabrw = "ccCLIENTE"
    frm.gsCodigobrw = "IDCLIENTE"
    frm.gbTypeCodeStr = False
    frm.gsDescrbrw = "Nombre"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
        txtCliente.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtNombre.Text = frm.gsDescrbrw
      fmtTextbox txtNombre, "R"

    End If
End Sub

Private Sub cmdDelCliente_Click()
txtCliente.Text = ""
txtNombre.Text = ""
End Sub

Private Sub cmdDelVendedor_Click()
txtVendedor.Text = ""
txtDescrVendedor.Text = ""
End Sub

Private Sub cmdRefresh_Click()
Dim lbok As Boolean
If ValCtrls() Then
      GSSQL = gsCompania & ".fafgetPedidos 'C', " & sPedidoInicial & "," & sPedidoFinal & ",'" & sFechaInicial & "','" & sFechaFinal & "'," & sCliente & "," & sVendedor & "," & sDesaprobados & "," & sAnulados
      If rsttmpProdFac.State = adStateOpen Then rsttmpProdFac.Close
      Set rsttmpProdFac = GetRecordset(GSSQL)
      Set TDBGFAC.DataSource = rsttmpProdFac
      TDBGFAC.Refresh
End If
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
        

End Sub


Private Sub cmdVendedor_Click()
  Dim frm As frmBrowseCat

    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = "Vendedor"
    frm.gsTablabrw = "fafVendedor"
    frm.gsCodigobrw = "IDVendedor"
    frm.gbTypeCodeStr = False
    frm.gsDescrbrw = "Nombre"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
        txtVendedor.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtDescrVendedor.Text = frm.gsDescrbrw
      fmtTextbox txtDescrVendedor, "R"
    End If
End Sub

Private Function ValCtrls() As Boolean
Dim lbok As Boolean
Dim sDescr As String
On Error GoTo salir
gsOperacionError = ""
lbok = True
If txtPedidoInicial.Text <> "" Then
    If Not Val_TextboxNum(txtPedidoInicial) Then
        gsOperacionError = "El Número del Pedido inicial debe ser numérico."
        lbok = Mensaje(gsOperacionError, ICO_ADVERTENCIA, False)
        lbok = False
        'txtCodVendedor.SetFocus
        txtPedidoInicial.SetFocus
        GoTo salir
        
    End If
End If

If txtPedidoFinal.Text <> "" Then
    If Not Val_TextboxNum(txtPedidoFinal) Then
     gsOperacionError = "El Número del Pedido final debe ser numérico."
     lbok = Mensaje(gsOperacionError, ICO_ADVERTENCIA, False)
     lbok = False
     'txtCodVendedor.SetFocus
     txtPedidoFinal.SetFocus
     GoTo salir
    End If
End If

If Val(txtPedidoInicial.Text) > Val(txtPedidoFinal.Text) Then
     gsOperacionError = "El Número del Pedido inicial debe ser mejor que el final."
     lbok = Mensaje(gsOperacionError, ICO_ADVERTENCIA, False)
     lbok = False
     'txtCodVendedor.SetFocus
     txtPedidoFinal.SetFocus
     GoTo salir
    
End If
If txtPedidoInicial.Text = "" Then
    sPedidoInicial = "0"
Else
    sPedidoInicial = txtPedidoInicial.Text
End If
If txtPedidoFinal.Text = "" Then
    sPedidoFinal = "0"
Else
    sPedidoFinal = txtPedidoFinal.Text
End If
If Format(DTPDesde.value, "yyyy-mm-dd") > Format(DTPHasta.value, "yyyy-mm-dd") Then
lbok = Mensaje("La fecha inicial no puede ser mayor que la fecha final ", ICO_PREGUNTA, False)
    If lbok = True Then
        lbok = False
        DTPHasta.SetFocus
        GoTo salir
    Else
        lbok = True
    End If
End If

    sFechaInicial = Format(Str(DTPDesde.value), "yyyymmdd")
    sFechaFinal = Format(Str(DTPHasta.value), "yyyymmdd")
If txtCliente.Text <> "" Then
    If Not Val_TextboxNum(txtCliente) Then
     gsOperacionError = "El código del Cleinte debe ser numérico."
     lbok = Mensaje(gsOperacionError, ICO_ADVERTENCIA, False)
     lbok = False
     cmdCliente.SetFocus
     'txtcodBodega.SetFocus
     GoTo salir
    End If
End If

If txtCliente.Text <> "" Then
    sDescr = GetDescrCat("IDCLIENTE", txtCliente.Text, "ccCLIENTE", "Nombre")
  If sDescr = "" Then
    gsOperacionError = "El CLIENTE no existe."
    lbok = Mensaje(gsOperacionError, ICO_ADVERTENCIA, False)
    txtCodCliente.Text = ""
    txtNombre.Text = ""
    lbok = False
    'txtCodCliente.SetFocus
    cmdCliente.SetFocus
    
    GoTo salir
  Else
    txtNombre.Text = sDescr
  End If
End If


If txtCliente.Text = "" Then
    sCliente = "-1"
Else
    sCliente = txtCliente.Text
End If

If txtVendedor.Text <> "" Then
    If Not Val_TextboxNum(txtVendedor) Then
     gsOperacionError = "El código del Vendedor debe ser numérico."
     lbok = Mensaje(gsOperacionError, ICO_ADVERTENCIA, False)
     lbok = False
     'txtCodVendedor.SetFocus
     cmdVendedor.SetFocus
     GoTo salir
    End If
End If

If txtVendedor.Text <> "" Then
    sDescr = GetDescrCat("IDVENDEDOR", txtVendedor.Text, "FAFVENDEDOR", "Nombre")
  If sDescr = "" Then
    gsOperacionError = "El Vendedor no existe."
    lbok = Mensaje(gsOperacionError, ICO_ADVERTENCIA, False)
    txtVendedor.Text = ""
    txtDescrVendedor.Text = ""
    lbok = False
    'txtCodVendedor.SetFocus
    cmdVendedor.SetFocus
 
    GoTo salir
  Else
    txtDescrVendedor.Text = sDescr
  End If
End If
If txtVendedor.Text = "" Then
    sVendedor = "-1"
Else
    sVendedor = txtVendedor.Text
End If

If txtCliente.Text = "" Then
    sCliente = "-1"
Else
    sCliente = txtCliente.Text
End If

If chkDesaprobados.value = 1 Then
    sDesaprobados = 1
Else
    sDesaprobados = 0
End If

If ChkAnuladas.value = 1 Then
    sAnulados = 1
Else
    sAnulados = 0
End If
lbok = True
ValCtrls = lbok
Exit Function
salir:

ValCtrls = lbok
End Function


Private Sub Form_Load()
PreparaRst
SetColumnSizeGrid
End Sub

Private Sub SetColumnSizeGrid()
TDBGFAC.Columns("IDCliente").Width = 824.882
TDBGFAC.Columns("Nombre").Width = 3479.811
TDBGFAC.Columns("Pedido").Width = 1019.906
TDBGFAC.Columns("Fecha").Width = 1649.764
TDBGFAC.Columns("SubTotal").Width = 1484.787
TDBGFAC.Columns("TotalImpuesto").Width = 1379.906
TDBGFAC.Columns("Total").Width = 1470.047
TDBGFAC.Columns("Aprobado").Width = 1244.976
TDBGFAC.Columns("Anulado").Width = 1230.236

End Sub




Private Sub TDBGFAC_DblClick()
Dim lbok As Boolean
Dim vPosition As Variant
If rsttmpProdFac.State = adStateClosed Then Exit Sub
If Not (rsttmpProdFac.EOF And rsttmpProdFac.BOF) Then
    vPosition = rsttmpProdFac.Bookmark
    Dim frm As frmDetPedido
    Set frm = New frmDetPedido
    frm.gsIDBodega = rsttmpProdFac("IDBodega").value
    frm.gsIDCliente = rsttmpProdFac("IDcliente").value
    frm.gsIDPedido = rsttmpProdFac("IDPedido").value
    frm.gsFecha = rsttmpProdFac("Fecha").value
    frm.gsNombre = rsttmpProdFac("Nombre").value
    frm.Show vbModal
    If frm.gbHuboAnulacion Then
        ChkAnuladas.value = 0
    End If
    If frm.gbHuboAprobacion Then
        chkDesaprobados.value = 0
    End If
    If frm.gbHuboAnulacion Or frm.gbHuboAprobacion Then
        cmdRefresh_Click
    End If

    Set frm = Nothing
    rsttmpProdFac.Bookmark = vPosition
End If
End Sub

Sub txtCliente_KeyPress(KeyAscii As Integer)
Dim sDescr As String
Dim lbok As Boolean
If KeyAscii = vbKeyReturn Then
    sDescr = getDescrCatalogo(txtCliente, "CODCliente", "ccCliente", "Nombre")
    If sDescr <> "" Then
        txtNombre.Text = sDescr
    Else
        lbok = Mensaje("Ese Cliente No Existe", ICO_ERROR, False)
    End If
End If
End Sub

Private Sub txtCliente_LostFocus()
Dim sDescr As String
Dim lbok As Boolean

    sDescr = getDescrCatalogo(txtCliente, "CodCliente", "ccCliente", "Nombre")
    If sDescr <> "" Then
        txtNombre.Text = sDescr
    Else
        lbok = Mensaje("Ese Cliente No Existe", ICO_ERROR, False)
    End If
End Sub

Private Sub txtVendedor_KeyPress(KeyAscii As Integer)
Dim sDescr As String
Dim lbok As Boolean
If KeyAscii = vbKeyReturn Then
    sDescr = getDescrCatalogo(txtVendedor, "IDVendedor", "fafVendedor", "Nombre")
    If sDescr <> "" Then
        txtDescrVendedor.Text = sDescr
    Else
        lbok = Mensaje("Ese Vendedor No Existe", ICO_ERROR, False)
    End If
End If
End Sub

Private Sub txtVendedor_LostFocus()
Dim sDescr As String
Dim lbok As Boolean

    sDescr = getDescrCatalogo(txtVendedor, "IDVendedor", "fafVendedor", "Nombre")
    If sDescr <> "" Then
        txtDescrVendedor.Text = sDescr
    Else
        lbok = Mensaje("Ese Vendedor No Existe", ICO_ERROR, False)
    End If

End Sub

