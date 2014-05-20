VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmAutoSugiereLotes 
   Caption         =   "v"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Height          =   320
      Left            =   2670
      Picture         =   "frmAutoSugiereLotes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2520
      Width           =   300
   End
   Begin VB.CommandButton cmdLote 
      Height          =   320
      Left            =   2280
      Picture         =   "frmAutoSugiereLotes.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2535
      Width           =   300
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4560
      TabIndex        =   21
      Top             =   6450
      Width           =   1155
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2640
      TabIndex        =   20
      Top             =   6450
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   1665
      Left            =   240
      TabIndex        =   11
      Top             =   540
      Width           =   8175
      Begin VB.TextBox txtCantidadTotal 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1650
         TabIndex        =   19
         Top             =   1140
         Width           =   975
      End
      Begin VB.TextBox txtIdBodega 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1650
         TabIndex        =   16
         Top             =   330
         Width           =   975
      End
      Begin VB.TextBox txtDescrBodega 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2730
         TabIndex        =   15
         Top             =   330
         Width           =   4935
      End
      Begin VB.TextBox txtProducto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1650
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtDescrProducto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2730
         TabIndex        =   12
         Top             =   720
         Width           =   4935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   270
         TabIndex        =   18
         Top             =   1170
         Width           =   1365
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
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
         Height          =   345
         Left            =   300
         TabIndex        =   17
         Top             =   390
         Width           =   765
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Producto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   300
         TabIndex        =   14
         Top             =   780
         Width           =   765
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
      Left            =   7650
      Picture         =   "frmAutoSugiereLotes.frx":200C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
      Top             =   3930
      Width           =   555
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   555
      Left            =   7650
      Picture         =   "frmAutoSugiereLotes.frx":2CD6
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
      Top             =   5130
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
      Left            =   7650
      Picture         =   "frmAutoSugiereLotes.frx":49A0
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Agrega el item con los datos digitados..."
      Top             =   3330
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
      Left            =   7650
      Picture         =   "frmAutoSugiereLotes.frx":566A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
      Top             =   4530
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
      Left            =   7650
      Picture         =   "frmAutoSugiereLotes.frx":6334
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Deshacer / Cancelar"
      Top             =   5730
      Width           =   555
   End
   Begin VB.TextBox txtCantidad 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1170
      TabIndex        =   3
      Top             =   2910
      Width           =   975
   End
   Begin VB.TextBox txtLoteInterno 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3060
      TabIndex        =   2
      Top             =   2520
      Width           =   4785
   End
   Begin VB.TextBox txtIdLote 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1170
      TabIndex        =   0
      Top             =   2520
      Width           =   975
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   2835
      Left            =   240
      OleObjectBlob   =   "frmAutoSugiereLotes.frx":6FFE
      TabIndex        =   10
      Top             =   3330
      Width           =   7185
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
      Left            =   -750
      TabIndex        =   24
      Top             =   0
      Width           =   10140
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   0
      Picture         =   "frmAutoSugiereLotes.frx":C9EF
      Stretch         =   -1  'True
      Top             =   -330
      Width           =   11490
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
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
      Left            =   300
      TabIndex        =   4
      Top             =   2910
      Width           =   765
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lote:"
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
      Left            =   300
      TabIndex        =   1
      Top             =   2520
      Width           =   765
   End
End
Attribute VB_Name = "frmAutoSugiereLotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public grst As ADODB.Recordset
Dim bOrdenCodigo As Boolean
Dim bOrdenDescr As Boolean
Dim sCodSucursal As String
Dim Accion As TypAccion
Dim sSoloActivo As String
Public gsFormCaption As String
Public gsTitle As String

Public gsIDBodega As Integer
Public gsIDProducto As Integer
Public gsDescrProducto As String
Public gsDescrBodega As String
Public gdCantidad As Double


Private Sub HabilitarBotones()
    Select Case Accion
        Case TypAccion.Add, TypAccion.Edit
            cmdSave.Enabled = True
            cmdUndo.Enabled = True
            cmdEliminar.Enabled = False
            cmdAdd.Enabled = False
            cmdEditItem.Enabled = False
            cmdCancelar.Enabled = False
            cmdAceptar.Enabled = False
        Case TypAccion.View
            cmdSave.Enabled = False
            cmdUndo.Enabled = False
            cmdEliminar.Enabled = True
            cmdAdd.Enabled = True
            cmdEditItem.Enabled = True
            cmdCancelar.Enabled = True
            cmdAceptar.Enabled = True
            
    End Select
End Sub

Public Sub HabilitarControles()
    Select Case Accion
        Case TypAccion.Add
            txtIdLote.Enabled = True
            txtLoteInterno.Enabled = True
            txtCantidad.Enabled = True
            cmdLote.Enabled = True
            cmdClear.Enabled = True
            txtCantidad.Text = ""
            txtIdLote.Text = ""
            txtLoteInterno.Text = ""
            fmtTextbox txtIdLote, "R"
            fmtTextbox txtLoteInterno, "O"
            Me.TDBG.Enabled = False
        Case TypAccion.Edit
            txtIdLote.Enabled = True
            txtLoteInterno.Enabled = True
            cmdLote.Enabled = False
            cmdClear.Enabled = False
            fmtTextbox txtIdLote, "R"
            fmtTextbox txtLoteInterno, "R"
            txtCantidad.Enabled = True
            Me.TDBG.Enabled = False
        Case TypAccion.View
            cmdLote.Enabled = False
            cmdClear.Enabled = False
            fmtTextbox txtIdLote, "R"
            fmtTextbox txtCantidad, "O"
            fmtTextbox txtLoteInterno, "R"
            Me.TDBG.Enabled = True
    End Select
End Sub


Private Sub cmdAceptar_Click()
    ValidarCantidadLotes
    Hide
End Sub

Private Sub cmdAdd_Click()
    Accion = Add
    HabilitarBotones
    HabilitarControles
    txtLoteInterno.SetFocus
End Sub

Private Sub cmdCancelar_Click()
    Set grst = Nothing
    Hide
End Sub

Private Sub cmdClear_Click()
    Me.txtIdLote.Text = ""
    Me.txtLoteInterno.Text = ""
End Sub

Private Sub cmdEditItem_Click()
    Accion = Edit
    GetDataFromGridToControl
    HabilitarBotones
    HabilitarControles
    txtCantidad.SetFocus
End Sub
Private Sub GetDataFromGridToControl()
    If Not (grst.EOF And grst.BOF) Then
        txtIdLote.Text = grst("IDLote").value
        txtLoteInterno.Text = grst("LoteInterno").value
        txtCantidad.Text = grst("Cantidad").value
    Else
        txtIdLote.Text = ""
        txtLoteInterno.Text = ""
        txtCantidad.Text = ""
    End If
End Sub

Private Sub cmdEliminar_Click()
    Dim lbok As Boolean
    
    lbok = Mensaje("Esta seguro que desea eliminar el registro seleccionado?", ICO_INFORMACION, True)
    If (lbok) Then
        grst.Delete
        Accion = View
        HabilitarBotones
        HabilitarControles
        Me.TDBG.ReBind
    End If
End Sub

Private Sub ValidarCantidadLotes()
   Dim CantLotes As Double, CantidadEdicion As Double
    CantidadEdicion = IIf(Accion = View, 0, Val(txtCantidad.Text))
    CantLotes = GetTotalLotes(grst)
   If (CantLotes + CantidadEdicion) > Me.gdCantidad Then
        lbok = Mensaje("La Cantidad del detalle de Lote no puede mayor a la cantidad total del producto", ICO_ERROR, False)
        Exit Sub
    ElseIf (CantLotes + CantidadEdicion) < Me.gdCantidad Then
        lbok = Mensaje("La Cantidad del detalle de Lote no puede menor a la cantidad total del producto", ICO_ERROR, False)
        Exit Sub
    End If
End Sub

Private Sub cmdSave_Click()
    Dim lbok As Boolean
    Dim sMsg As String
    Dim sActivo As String
    Dim sFactura As String
    Dim sFiltro As String
        If txtIdLote.Text = "" Then
            lbok = Mensaje("El Lote no puede estar en Blanco", ICO_ERROR, False)
            Exit Sub
        End If

        If txtCantidad.Text = "" Then
            lbok = Mensaje("La Cantidad no puede estar en Blanco", ICO_ERROR, False)
            Exit Sub
        End If
        If txtLoteInterno.Text = "" Then
            lbok = Mensaje("La Descripción del Lote no puede estar en blanco", ICO_ERROR, False)
            Exit Sub
        End If

    'Validar que la cantidad total cuadre con la cantidad de lotes
 
    
    ValidarCantidadLotes
    
    
    
    If (Accion = Add) Then
          If ExiteRstKey(grst, "IDLote=" & Me.txtIdLote.Text & " AND IDPRODUCTO=" & Me.gsIDProducto & _
                                        " AND IdBodega=" & gsIDBodega) Then
            lbok = Mensaje("Ya existe ese el registro en la transacción", ICO_ERROR, False)

            Exit Sub
          End If
        
          ' Carga los datos del detalle de transacciones para ser grabados a la bd
        
        Dim datosLote As New Dictionary
        getValueFieldsFromTable "invlote", "LoteInterno,LoteProveedor,FechaVencimiento,FechaFabricacion", " IDLote=" & Me.txtIdLote.Text, datosLote
        grst.AddNew
        grst!IdBodega = Me.gsIDBodega
        grst!IdProducto = Me.gsIDProducto
        grst!IDLote = Me.txtIdLote.Text
        grst!Cantidad = Me.txtCantidad.Text
        grst!FechaVencimiento = datosLote("FechaVencimiento")
        grst!FechaFabricacion = datosLote("FechaFabricacion")
        grst!LoteProveedor = datosLote("LoteProveedor")
        grst!LoteInterno = datosLote("LoteInterno")
          
        grst.Update
        
        grst.MoveFirst
    ElseIf (Accion = Edit) Then
      grst!IdBodega = gsIDBodega
      grst!IdProducto = gsIDProducto
      grst!IDLote = Me.txtIdLote.Text
      grst!Cantidad = Me.txtCantidad.Text
      grst.Update
    End If



      Set TDBG.DataSource = grst
      TDBG.ReBind
      
      
      'Me.dtgAjuste.Columns("Descr").FooterText = "Items de la transacción :     " & grstTransAI.RecordCount

      Accion = View
      HabilitarControles
      HabilitarBotones

End Sub

Private Sub cmdLote_Click()
   Dim frm As frmBrowseCat

    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = "Lote Producto" '& lblund.Caption
    frm.gsTablabrw = "vinvExistenciaLote"
    frm.gsCodigobrw = "IDLote"
    frm.gbTypeCodeStr = False
    frm.gsDescrbrw = "LoteInterno"
    frm.gbFiltra = True
    frm.gsFiltro = " IdProducto=" & gsIDProducto & " and Existencia>0"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      Me.txtIdLote.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      Me.txtLoteInterno.Text = frm.gsDescrbrw
      fmtTextbox txtLoteInterno, "R"
    End If
End Sub

Private Sub cmdUndo_Click()
    GetDataFromGridToControl
    Accion = View
    HabilitarControles
    HabilitarBotones
End Sub

Private Sub Form_Load()
Dim dTotalExistenciaLote As Double
    Set grst = New ADODB.Recordset
    If grst.State = adStateOpen Then grst.Close
    grst.ActiveConnection = gConet 'Asocia la conexión de trabajo
    grst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    grst.CursorLocation = adUseClient ' Cursor local al cliente
    grst.LockType = adLockOptimistic

    'Set grst = invGetSugeridoLote(gsIDBodega, gsIdProducto, gdCantidad)
    
    Caption = gsFormCaption
    lbFormCaption = gsTitle
    txtProducto.Text = gsIDProducto
    txtDescrProducto.Text = gsDescrProducto
    Me.txtIdBodega.Text = gsIDBodega
    Me.txtDescrBodega.Text = Me.gsDescrBodega
    Me.txtCantidadTotal.Text = Me.gdCantidad
    Accion = View
    HabilitarBotones
    HabilitarControles
    cargaGrid
    dTotalExistenciaLote = GetTotalLotes(grst)
    If dTotalExistenciaLote < gdCantidad Then
        lbok = Mensaje("No hay suficiente existencias para satisfacer el producto, solamente dispone de " & Str(dTotalExistenciaLote), ICO_ERROR, False)
        Set grst = Nothing
        Unload Me
    End If
End Sub
    
Private Sub cargaGrid()

   If grst.State = adStateOpen Then grst.Close
    grst.ActiveConnection = gConet 'Asocia la conexión de trabajo
    grst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    grst.CursorLocation = adUseClient ' Cursor local al cliente
    grst.LockType = adLockOptimistic
    If grst.State = adStateOpen Then grst.Close
    Set grst = invGetSugeridoLote(gsIDBodega, gsIDProducto, gdCantidad)
    If Not (grst.EOF And grst.BOF) Then
      Set TDBG.DataSource = grst
      TDBG.Refresh
    End If
End Sub




Private Sub TDBG_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    GetDataFromGridToControl
    HabilitarBotones
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Accion = Edit Or Accion = Add Then
        Cancel = True
        Exit Sub
    End If
    If Not (grst Is Nothing) Then Set grst = Nothing
End Sub

Private Function GetTotalLotes(grst As Recordset) As Double

    Dim dResult As Double
    dResult = 0
    
    If Not (grst.EOF And grst.BOF) Then
      grst.MoveFirst
      While Not grst.EOF
            dResult = dResult + grst!Cantidad
   
        grst.MoveNext
      Wend
      grst.MoveFirst
    End If
    GetTotalLotes = dResult
End Function

Public Function getTotalSugeridoporLote() As Double
Dim rst As New ADODB.Recordset
Dim dResultado As Double
rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
rst.CursorLocation = adUseClient ' Cursor local al cliente
rst.LockType = adLockOptimistic
If rst.State = adStateOpen Then rst.Close
Set rst = invGetSugeridoLote(gsIDBodega, gsIDProducto, gdCantidad)
dResultado = 0
If Not (rst.EOF And rst.BOF) Then
    dResultado = GetTotalLotes(rst)
End If

getTotalSugeridoporLote = dResultado

End Function



