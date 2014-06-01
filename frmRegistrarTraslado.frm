VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRegistrarTraslado 
   Caption         =   "Form1"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13230
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   13230
   WindowState     =   2  'Maximized
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
      Left            =   10980
      Picture         =   "frmRegistrarTraslado.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Agrega el item con los datos digitados..."
      Top             =   6870
      Width           =   555
   End
   Begin VB.Frame Frame1 
      Height          =   2685
      Left            =   12300
      TabIndex        =   32
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
         Picture         =   "frmRegistrarTraslado.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Deshacer / Cancelar"
         Top             =   2040
         Width           =   555
      End
      Begin VB.CommandButton cmdSave 
         Enabled         =   0   'False
         Height          =   555
         Left            =   90
         Picture         =   "frmRegistrarTraslado.frx":1994
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
         Top             =   1440
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
         Picture         =   "frmRegistrarTraslado.frx":365E
         Style           =   1  'Graphical
         TabIndex        =   34
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
         Picture         =   "frmRegistrarTraslado.frx":4328
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
         Top             =   180
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdLote 
      Height          =   320
      Left            =   2160
      Picture         =   "frmRegistrarTraslado.frx":4FF2
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7140
      Width           =   300
   End
   Begin VB.CommandButton cmdProducto 
      Height          =   320
      Left            =   2160
      Picture         =   "frmRegistrarTraslado.frx":5334
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6660
      Width           =   300
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
      Left            =   9570
      TabIndex        =   29
      Top             =   7140
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
      Left            =   2520
      TabIndex        =   27
      Top             =   7140
      Width           =   5715
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
      Left            =   1020
      TabIndex        =   26
      Top             =   7140
      Width           =   1095
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
      Left            =   8640
      TabIndex        =   24
      Top             =   6660
      Width           =   2145
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
      Left            =   2520
      TabIndex        =   23
      Top             =   6660
      Width           =   5715
   End
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
      Left            =   1020
      TabIndex        =   22
      Top             =   6660
      Width           =   1095
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   3165
      Left            =   240
      OleObjectBlob   =   "frmRegistrarTraslado.frx":5676
      TabIndex        =   20
      Top             =   3300
      Width           =   11955
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
      Format          =   61800449
      CurrentDate     =   41787
   End
   Begin VB.TextBox Text4 
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
      Picture         =   "frmRegistrarTraslado.frx":988B
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
      Picture         =   "frmRegistrarTraslado.frx":9BCD
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
      Top             =   900
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
      ScaleWidth      =   13230
      TabIndex        =   0
      Top             =   0
      Width           =   13230
      Begin VB.Image Image 
         Height          =   480
         Index           =   2
         Left            =   270
         Picture         =   "frmRegistrarTraslado.frx":9F0F
         Top             =   120
         Width           =   480
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
   Begin VB.Label Label8 
      Caption         =   "Cantidad:"
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   8700
      TabIndex        =   28
      Top             =   7200
      Width           =   825
   End
   Begin VB.Label Label7 
      Caption         =   "Lote:"
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   300
      TabIndex        =   25
      Top             =   7170
      Width           =   825
   End
   Begin VB.Label Label6 
      Caption         =   "Producto:"
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   270
      TabIndex        =   21
      Top             =   6690
      Width           =   825
   End
   Begin VB.Label Label5 
      Caption         =   "Num Salida:"
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   210
      TabIndex        =   18
      Top             =   2820
      Width           =   1245
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha:"
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
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   210
      TabIndex        =   7
      Top             =   2370
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "Bodega Origen:"
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   210
      TabIndex        =   6
      Top             =   1920
      Width           =   1245
   End
   Begin VB.Label lbl 
      Caption         =   "IDTrasaldo:"
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

Private rstDetalle As ADODB.Recordset
Dim rstLote As ADODB.Recordset


Dim gTrans As Boolean ' se dispara si hubo error en medio de la transacción
Dim gBeginTransNoEnd As Boolean ' Indica si hubo un begin sin rollback o commit
Dim gTasaCambio As Double
Dim tDatosDelProducto As typDatosProductos

Public scontrol As String

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
    
    txtNumSalida.Text = ""
    fmtTextbox txtNumSalida, "O"
End Sub

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
                cmdEliminar.Enabled = True
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
           
            
            txtIDProducto.Text = ""
            fmtTextbox txtIDProducto, "O"
            txtDescrProducto.Text = ""
            fmtTextbox txtDescrProducto, "R"
            
            txtIdLote.Text = ""
            fmtTextbox txtIdLote, "O"
            txtDescrLote.Text = ""
            fmtTextbox txtDescrLote, "R"
            
            txtCantidad.Text = ""
            fmtTextbox txtCantidad, "O"
            
                       
            Me.cmdBodegaDestino.Enabled = True
            Me.cmdBodegaOrigen.Enabled = True
            Me.cmdProducto.Enabled = True
            Me.cmdLote.Enabled = True
            Me.TDBG.Enabled = True
            
            Me.chkAutoSugiereLotes.value = vbChecked
            HabilitarAutoSugerirLotes Me.chkAutoSugiereLotes.value
            
        Case TypAccion.Edit
            
            
            fmtTextbox txtIDProducto, "O"
            fmtTextbox Me.txtDescrProducto, "R"
                        
            
            fmtTextbox txtIdLote, "O"
            fmtTextbox txtDescrLote, "R"
            
            fmtTextbox txtCantidad, "O"
        
            
            
            Me.cmdProducto.Enabled = False
            Me.cmdLote.Enabled = False
            Me.TDBG.Enabled = False
            
        Case TypAccion.View
           
          
            fmtTextbox txtIDProducto, "O"
            txtIdLote.Text = ""
            fmtTextbox txtIdLote, "O"
            txtCantidad.Text = ""
            fmtTextbox txtCantidad, "O"
            
            Me.cmdProducto.Enabled = False
            Me.cmdLote.Enabled = False
           
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
    ElseIf (Me.txtNumSalida.Text = "") Then
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
    ElseIf (Me.txtIdLote.Text = "" And Me.chkAutoSugiereLotes.value = False) Then
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
        Me.txtIdLote.Enabled = False
        Me.txtDescrLote.Enabled = False
        Me.cmdLote.Enabled = False
        'Me.cmdDelLote.Enabled = False
        bIsAutoSugiereLotes = True
    Else
        If (Accion = Add) Then
            Me.txtIdLote.Enabled = True
            Me.txtDescrLote.Enabled = True
            Me.cmdLote.Enabled = True
            'Me.cmdDelLote.Enabled = True
        End If
        bIsAutoSugiereLotes = False
    End If
End Sub


Private Sub cmdAdd_Click()
   Dim lbok As Boolean
    
    If Not ValCtrlsCabecera Then
        lbok = Mensaje("Revise sus datos por favor !!! " & sMensajeError, ICO_ERROR, False)
        Exit Sub
    End If
   
    If Not ValCtrlsDetalle Then
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
                                                " AND IDLOTE=" & Me.txtIdLote.Text) Then
              lbok = Mensaje("Ya existe ese el registro en la transacción", ICO_ERROR, False)
        
              Exit Sub
            End If
            Set rstLote = New ADODB.Recordset
              rstLote.ActiveConnection = gConet
            CargaDatosLotes rstLote, CInt(Trim(Me.txtIdLote.Text))
            ' Carga los datos del detalle de transacciones para ser grabados a la bd
            rstDetalle.AddNew
            rstDetalle!IdProducto = Me.txtIDProducto.Text
            rstDetalle!DescrProducto = Me.txtDescrProducto.Text
            rstDetalle!IdLote = Me.txtIdLote.Text
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
            rstDetalle!IdLote = Me.txtIdLote.Text
            rstDetalle!LoteInterno = Me.txtDescrLote.Text
            rstDetalle!FechaVencimiento = rstLote!FechaVencimiento
            rstDetalle!Cantidad = Val(Me.txtCantidad.Text)
            
            
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
        Me.txtIdLote.Text = rstDetalle("IDLote").value
        Me.txtDescrLote.Text = rstDetalle("DescrLote").value
        Me.txtCantidad.Text = rstDetalle("Cantidad").value
        
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
        txtIdLote.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      Me.txtDescrLote.Text = frm.gsDescrbrw
      fmtTextbox Me.txtIdLote, "R"
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
      Me.txtIdLote.Enabled = True
      If (Me.chkAutoSugiereLotes.value = True) Then
        Me.cmdLote.Enabled = False
      Else
        Me.cmdLote.Enabled = True
      End If
    End If
End Sub

Private Function invSaveCabeceraTraslado() As String
  
    Dim lbok As Boolean
    On Error GoTo errores
    lbok = False
    Dim sDocumento As String
    Dim rst As ADODB.Recordset

    GSSQL = "invUpdateCabTraslados 'I','','" & Me.txtBodegaOrigen.Text & "','" & Me.txtBodegaDestino.Text & "','1-16','" & Format(Str(Me.dtpFecha.value), "yyyymmdd") & _
                "','" & Format("1980-01-01", "yyyymmdd") & "','','" & Me.txtNumSalida.Text & "','',0"
 Set rst = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    sDocumento = "" 'Indica que ocurrió un error
    sMensajeError = "Ha ocurrido un error tratando de ingresar la cabecera del traslado!!!" & err.Description
  Else  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    sDocumento = rst("IDTraslado").value
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
                                    "0", _
                                    "0", _
                                    "0", _
                                    "0")
                                

            rst.MoveNext
      Wend
      rst.MoveFirst

    End If
    Exit Sub
errores:
    gTrans = False
    'gConet.RollbackTrans 'Descomentarie esto
End Sub


Private Sub cmdSave_Click()
 Dim lbok As Boolean
    'On Error GoTo errores
    
    If Not ValCtrlsCabecera Then
        lbok = Mensaje("Revise sus datos por favor !!! " & sMensajeError, ICO_ERROR, False)
        Exit Sub
    End If
    
    If rstDetalle.RecordCount > 0 Then
      
'      If gTasaCambio = 0 Then
'        lbOk = Mensaje("La tasa de cambio es Cero llame a informática por favor ", ICO_ERROR, False)
'        Exit Sub
'      End If
      
      
      gConet.BeginTrans ' inicio aqui la transacción
      
      gBeginTransNoEnd = True
      Dim sDocumento As String
      sDocumento = invSaveCabeceraTraslado()
        If sDocumento <> "" Then ' salva la cabecera
        SaveRstDetalle rstDetalle, sDocumento, "I" ' salva el detalle que esta en batch
'        If (gTrans = True) Then
'            invMasterAcutalizaSaldosInventarioPaquete sDocumento, gsIDTipoTransaccion, Me.gsIDTipoTransaccion, gsUser
'        End If



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




Private Sub cmdUndo_Click()
    GetDataFromGridToControl
    Accion = Add
    HabilitarBotones
    HabilitarControles
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
    'gTasaCambio = GetTasadeCambio(Format(Now, "YYYY/MM/DD"))
    PreparaRstDetalle ' Prepara los Recordsets
    Set Me.TDBG.DataSource = rstDetalle
    Me.TDBG.Refresh
    
    InicializaControles
   
    'SetTextBoxReadOnly
    Accion = Add
    HabilitarBotones
    HabilitarControles
    Me.chkAutoSugiereLotes.value = vbChecked
    HabilitarAutoSugerirLotes Me.chkAutoSugiereLotes.value
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


