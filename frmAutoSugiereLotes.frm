VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmAutoSugiereLotes 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8550
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAutoSugiereLotes.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Inventario.CtlLiner CtlLiner1 
      Height          =   30
      Left            =   -450
      TabIndex        =   25
      Top             =   780
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   53
   End
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
      Height          =   780
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   8550
      TabIndex        =   24
      Top             =   0
      Width           =   8550
      Begin VB.Label lbFormCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "TITLE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   780
         TabIndex        =   31
         Top             =   150
         Width           =   1185
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auto sugiere lotes, para la bodega y productos seleccionados"
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
         Left            =   780
         TabIndex        =   26
         Top             =   420
         Width           =   3750
      End
      Begin VB.Image Image 
         Height          =   540
         Index           =   2
         Left            =   150
         Picture         =   "frmAutoSugiereLotes.frx":058A
         Top             =   90
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdClear 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   2760
      Picture         =   "frmAutoSugiereLotes.frx":0F9D
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3270
      Width           =   300
   End
   Begin VB.CommandButton cmdLote 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   2370
      Picture         =   "frmAutoSugiereLotes.frx":2C67
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3285
      Width           =   300
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H80000009&
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
      Height          =   555
      Left            =   4470
      Picture         =   "frmAutoSugiereLotes.frx":2FA9
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7830
      Width           =   1155
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H80000009&
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
      Height          =   555
      Left            =   2760
      Picture         =   "frmAutoSugiereLotes.frx":32ED
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7800
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   " Detalle de Productos a Distribuir "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1665
      Left            =   210
      TabIndex        =   11
      Top             =   960
      Width           =   8175
      Begin VB.TextBox txtCantidadTotal 
         Appearance      =   0  'Flat
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1650
         TabIndex        =   19
         Top             =   1140
         Width           =   975
      End
      Begin VB.TextBox txtIdBodega 
         Appearance      =   0  'Flat
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1650
         TabIndex        =   16
         Top             =   330
         Width           =   975
      End
      Begin VB.TextBox txtDescrBodega 
         Appearance      =   0  'Flat
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   2730
         TabIndex        =   15
         Top             =   330
         Width           =   4935
      End
      Begin VB.TextBox txtProducto 
         Appearance      =   0  'Flat
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1650
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtDescrProducto 
         Appearance      =   0  'Flat
         ForeColor       =   &H00404040&
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
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
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
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
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
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Index           =   0
         Left            =   300
         TabIndex        =   14
         Top             =   780
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdEditItem 
      BackColor       =   &H80000009&
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
      Left            =   7620
      Picture         =   "frmAutoSugiereLotes.frx":3631
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
      Top             =   4980
      Width           =   555
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H80000009&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   7620
      Picture         =   "frmAutoSugiereLotes.frx":42FB
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
      Top             =   6180
      Width           =   555
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H80000009&
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
      Left            =   7620
      Picture         =   "frmAutoSugiereLotes.frx":5FC5
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Agrega el item con los datos digitados..."
      Top             =   4380
      Width           =   555
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H80000009&
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
      Left            =   7620
      Picture         =   "frmAutoSugiereLotes.frx":6C8F
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
      Top             =   5580
      Width           =   555
   End
   Begin VB.CommandButton cmdUndo 
      BackColor       =   &H80000009&
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
      Left            =   7620
      Picture         =   "frmAutoSugiereLotes.frx":7959
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Deshacer / Cancelar"
      Top             =   6780
      Width           =   555
   End
   Begin VB.TextBox txtCantidad 
      Appearance      =   0  'Flat
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   1260
      TabIndex        =   3
      Top             =   3660
      Width           =   975
   End
   Begin VB.TextBox txtLoteInterno 
      Appearance      =   0  'Flat
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   3150
      TabIndex        =   2
      Top             =   3270
      Width           =   4785
   End
   Begin VB.TextBox txtIdLote 
      Appearance      =   0  'Flat
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   1260
      TabIndex        =   0
      Top             =   3270
      Width           =   975
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   3285
      Left            =   360
      OleObjectBlob   =   "frmAutoSugiereLotes.frx":8623
      TabIndex        =   10
      Top             =   4170
      Width           =   7185
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   240
      Picture         =   "frmAutoSugiereLotes.frx":E014
      Top             =   7350
      Width           =   480
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Puede modificar los lotes sugeridos por el sistema."
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   720
      TabIndex        =   30
      Top             =   7470
      Width           =   3600
   End
   Begin VB.Label Label7 
      Caption         =   "____________________________ _ _ _"
      Height          =   405
      Left            =   5280
      TabIndex        =   29
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label Label6 
      Caption         =   " _ _ _ ____________________________"
      Height          =   405
      Left            =   240
      TabIndex        =   28
      Top             =   2760
      Width           =   3045
   End
   Begin VB.Label Label5 
      Caption         =   "Desglose de Lotes"
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
      Height          =   375
      Left            =   3450
      TabIndex        =   27
      Top             =   2850
      Width           =   1605
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   390
      TabIndex        =   4
      Top             =   3690
      Width           =   765
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   390
      TabIndex        =   1
      Top             =   3270
      Width           =   765
   End
End
Attribute VB_Name = "frmAutoSugiereLotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public grst As ADODB.Recordset
Dim Accion As TypAccion
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
            txtIDLote.Enabled = True
            txtLoteInterno.Enabled = True
            txtCantidad.Enabled = True
            cmdLote.Enabled = True
            cmdClear.Enabled = True
            txtCantidad.Text = ""
            txtIDLote.Text = ""
            txtLoteInterno.Text = ""
            fmtTextbox txtIDLote, "R"
            fmtTextbox txtLoteInterno, "O"
            Me.TDBG.Enabled = False
        Case TypAccion.Edit
            txtIDLote.Enabled = True
            txtLoteInterno.Enabled = True
            cmdLote.Enabled = False
            cmdClear.Enabled = False
            fmtTextbox txtIDLote, "R"
            fmtTextbox txtLoteInterno, "R"
            txtCantidad.Enabled = True
            Me.TDBG.Enabled = False
        Case TypAccion.View
            cmdLote.Enabled = False
            cmdClear.Enabled = False
            fmtTextbox txtIDLote, "R"
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
    Me.txtIDLote.Text = ""
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
        txtIDLote.Text = grst("IDLote").value
        txtLoteInterno.Text = grst("LoteInterno").value
        txtCantidad.Text = grst("Cantidad").value
    Else
        txtIDLote.Text = ""
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
        lbok = Mensaje("La Cantidad del detalle de Lote no puede ser mayor a la cantidad total del producto", ICO_ERROR, False)
        Exit Sub
    ElseIf (CantLotes + CantidadEdicion) < Me.gdCantidad Then
        lbok = Mensaje("La Cantidad del detalle de Lote no puede ser menor a la cantidad total del producto", ICO_ERROR, False)
        Exit Sub
    End If
End Sub

Private Sub cmdSave_Click()
    Dim lbok As Boolean

        If txtIDLote.Text = "" Then
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
          If ExiteRstKey(grst, "IDLote=" & Me.txtIDLote.Text & " AND IDPRODUCTO=" & Me.gsIDProducto & _
                                        " AND IdBodega=" & gsIDBodega) Then
            lbok = Mensaje("Ya existe ese el registro en la transacción", ICO_ERROR, False)

            Exit Sub
          End If
        
          ' Carga los datos del detalle de transacciones para ser grabados a la bd
        
        Dim datosLote As New Dictionary
        getValueFieldsFromTable "invlote", "LoteInterno,LoteProveedor,FechaVencimiento,FechaFabricacion", " IDLote=" & Me.txtIDLote.Text, datosLote
        grst.AddNew
        grst!IdBodega = Me.gsIDBodega
        grst!IdProducto = Me.gsIDProducto
        grst!IdLote = Me.txtIDLote.Text
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
      grst!IdLote = Me.txtIDLote.Text
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
      Me.txtIDLote.Text = frm.gsCodigobrw
      
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
    Me.txtIDBodega.Text = gsIDBodega
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



