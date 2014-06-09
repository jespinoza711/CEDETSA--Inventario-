VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmListadoTraslados 
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   8490
   WindowState     =   2  'Maximized
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   3795
      Left            =   240
      OleObjectBlob   =   "frmListadoTraslados.frx":0000
      TabIndex        =   25
      Top             =   3510
      Width           =   10395
   End
   Begin VB.Frame frmCab 
      Height          =   2475
      Left            =   240
      TabIndex        =   4
      Top             =   990
      Width           =   10425
      Begin VB.Frame Frame 
         Caption         =   "Filtrar por: "
         Height          =   585
         Left            =   5250
         TabIndex        =   21
         Top             =   1110
         Width           =   3765
         Begin VB.OptionButton optAmbas 
            Caption         =   "Ambas"
            Height          =   315
            Left            =   2730
            TabIndex        =   24
            Top             =   210
            Width           =   855
         End
         Begin VB.OptionButton optSalidas 
            Caption         =   "Salidas"
            Height          =   315
            Left            =   240
            TabIndex        =   23
            Top             =   210
            Width           =   1035
         End
         Begin VB.OptionButton optEntradas 
            Caption         =   "Entradas"
            Height          =   315
            Left            =   1500
            TabIndex        =   22
            Top             =   210
            Width           =   1035
         End
      End
      Begin VB.CommandButton cmdRefrescar 
         Caption         =   "Refrescar"
         Height          =   885
         Left            =   9210
         TabIndex        =   20
         Top             =   300
         Width           =   945
      End
      Begin VB.CheckBox chkViewPendienteAplicar 
         Caption         =   "Pendiente de Ingresar"
         Height          =   375
         Left            =   3840
         TabIndex        =   19
         Top             =   750
         Width           =   2505
      End
      Begin VB.TextBox txtNumEntrada 
         Height          =   345
         Left            =   6960
         TabIndex        =   18
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtNumSalida 
         Height          =   345
         Left            =   1620
         TabIndex        =   16
         Top             =   1740
         Width           =   2055
      End
      Begin VB.TextBox txtDocumentoInv 
         Height          =   345
         Left            =   1620
         TabIndex        =   14
         Top             =   1260
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker dtpFechaFinal 
         Height          =   345
         Left            =   7560
         TabIndex        =   12
         Top             =   780
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         _Version        =   393216
         Format          =   97320961
         CurrentDate     =   41787
      End
      Begin MSComCtl2.DTPicker dtpFechaInicial 
         Height          =   315
         Left            =   1620
         TabIndex        =   11
         Top             =   780
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   393216
         Format          =   97320961
         CurrentDate     =   41787
      End
      Begin VB.CommandButton cmdBodega 
         Height          =   320
         Left            =   2850
         Picture         =   "frmListadoTraslados.frx":5BED
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   300
         Width           =   300
      End
      Begin VB.TextBox txtDescrBodega 
         Height          =   345
         Left            =   3270
         TabIndex        =   7
         Top             =   300
         Width           =   5715
      End
      Begin VB.TextBox txtIDBodega 
         Height          =   345
         Left            =   1620
         TabIndex        =   6
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "Num Entrada:"
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   6
         Left            =   5910
         TabIndex        =   17
         Top             =   1860
         Width           =   1005
      End
      Begin VB.Label Label 
         Caption         =   "Num Salida:"
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label Label 
         Caption         =   "Documento Inv:"
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   1350
         Width           =   1155
      End
      Begin VB.Label Label 
         Caption         =   "Fecha Final:"
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   3
         Left            =   6540
         TabIndex        =   10
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label 
         Caption         =   "Fecha Inicial:"
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label 
         Caption         =   "Bodega:"
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   645
      End
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
      Height          =   750
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   8490
      TabIndex        =   0
      Top             =   0
      Width           =   8490
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
         TabIndex        =   2
         Top             =   90
         Width           =   855
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Listados de Traslados"
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
         TabIndex        =   1
         Top             =   420
         Width           =   1305
      End
      Begin VB.Image Image 
         Height          =   645
         Index           =   2
         Left            =   60
         Picture         =   "frmListadoTraslados.frx":5F2F
         Stretch         =   -1  'True
         Top             =   60
         Width           =   720
      End
   End
   Begin Inventario.CtlLiner CtlLiner 
      Height          =   30
      Left            =   0
      TabIndex        =   3
      Top             =   750
      Width           =   17925
      _ExtentX        =   31618
      _ExtentY        =   53
   End
End
Attribute VB_Name = "frmListadoTraslados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As ADODB.Recordset
Public gsFormCaption As String
Public gsTitle As String

Private Sub cargaGrid()
    If rst.State = adStateOpen Then rst.Close
    rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rst.CursorLocation = adUseClient ' Cursor local al cliente
    rst.LockType = adLockOptimistic
    
    Dim sEntradas  As String
    Dim sViewPendiente As String
    
    
    If (Me.optSalidas.value = True) Then
        sEntradas = "0"
    ElseIf Me.optEntradas.value = True Then
        sEntradas = "1"
    Else
        sEntradas = "-1"
    End If
    
    sViewPendiente = IIf(Me.chkViewPendienteAplicar.value = vbChecked, 1, 0)
    
    
    GSSQL = gsCompania & ".invGetCabTrasladosConFiltros '" & IIf(Trim(Me.txtDocumentoInv.Text) = "", "*", Trim(Me.txtDocumentoInv.Text)) & _
                                                    "'," & IIf(Trim(Me.txtIdBodega.Text) = "", -1, Trim(Me.txtIdBodega)) & _
                                                    ",'" & IIf(Trim(Me.txtNumEntrada.Text) = "", "*", Trim(Me.txtNumEntrada.Text)) & _
                                                    "','" & IIf(Trim(Me.txtNumSalida.Text) = "", "*", Trim(Me.txtNumSalida.Text)) & _
                                                    "','" & Me.dtpFechaInicial.value & "','" & Me.dtpFechaFinal.value & _
                                                    "'," & sEntradas & "," & sViewPendiente
    If rst.State = adStateOpen Then rst.Close
    Set rst = GetRecordset(GSSQL)
    If Not (rst.EOF And rst.BOF) Then
      Set TDBG.DataSource = rst
      TDBG.Refresh
    End If
End Sub

Private Sub InicializarControles()
    fmtTextbox Me.txtIdBodega, "R"
    fmtTextbox Me.txtDescrBodega, "R"
    
    Me.dtpFechaFinal.value = DateTime.Now
    Me.dtpFechaInicial.value = DateTime.DateAdd("W", -1, DateTime.Now)
    
    Me.optSalidas.value = True
End Sub


Private Sub chkViewPendienteAplicar_Click()
     If (Me.chkViewPendienteAplicar.value = vbChecked) Then
        Me.optEntradas.value = True
     End If
End Sub

Private Sub cmdBodega_Click()
    Dim frm As New frmBrowseCat
    
    frm.gsCaptionfrm = "Bodega"
    frm.gsTablabrw = "invBODEGA"
    frm.gsCodigobrw = "IDBodega"
    frm.gbTypeCodeStr = True
    frm.gsDescrbrw = "Descr"
    frm.gbFiltra = False
    'frm.gsFiltro = "IdPaquete='" & Me.gsIDTipoTransaccion & "'"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      Me.txtIdBodega.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      Me.txtDescrBodega.Text = frm.gsDescrbrw
      fmtTextbox txtDescrBodega, "R"
    End If
End Sub

Private Sub cmdRefrescar_Click()
    cargaGrid
End Sub

Private Sub HabilitarBotonesMain()
    Select Case Accion
        Case TypAccion.Add, TypAccion.Edit
            MDIMain.tbMenu.Buttons(8).Enabled = False 'Nuevo
            MDIMain.tbMenu.Buttons(10).Enabled = False 'Editar
        Case TypAccion.View
            MDIMain.tbMenu.Buttons(8).Enabled = True 'Nuevo
            MDIMain.tbMenu.Buttons(10).Enabled = True 'Editar
    End Select
End Sub

Private Sub Form_Activate()
    HighlightInWin Me.Name
    SetupFormToolbar (Me.Name)
End Sub

Public Sub CommandPass(ByVal srcPerformWhat As String)

    On Error GoTo err

    Select Case srcPerformWhat

        Case "SalidaProductos"
            CrearTraslado

        Case "IngresoProductos"
            CrearEntradaTraslado

        Case "Cerrar"
            Unload Me
    End Select

    Exit Sub

    'Trap the error
err:

    If err.Number = -2147467259 Then
        MsgBox "You cannot delete this record because it was used by other records! If you want to delete this record" & vbCrLf & "you will first have to delete or change the records that currenly used this record as shown bellow." & vbCrLf & vbCrLf & err.Description, , "Delete Operation Failed!"
        Me.MousePointer = vbDefault
    End If

End Sub

Private Sub CrearEntradaTraslado()
    Dim vPosition As Variant
    If rst.State = adStateClosed Then Exit Sub
    vPosition = rst.Bookmark
    Dim ofrmTraslado As New frmRegistrarTraslado
    ofrmTraslado.gsFormCaption = "Traslados"
    ofrmTraslado.gsTitle = "REGISTRO ENTRADA TRASLADO"
    ofrmTraslado.sAccion = "Entrada"
    ofrmTraslado.sDocumentoTraslado = rst("IDTraslado").value
    LoadForm ofrmTraslado
    rst.Bookmark = vPosition
End Sub

Private Sub CrearTraslado()

    Dim ofrmTraslado As New frmRegistrarTraslado

    ofrmTraslado.gsFormCaption = "Traslados"
    ofrmTraslado.gsTitle = "REGISTRO SALIDA TRASLADO"
    ofrmTraslado.sAccion = "Salida"
    LoadForm ofrmTraslado
End Sub


Private Sub Form_Load()
    MDIMain.AddForm Me.Name
    Set rst = New ADODB.Recordset
    If rst.State = adStateOpen Then rst.Close
    rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rst.CursorLocation = adUseClient ' Cursor local al cliente
    rst.LockType = adLockOptimistic

    Me.Caption = gsFormCaption
    Me.lbFormCaption = gsTitle
   
    InicializarControles
    
    cargaGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetupFormToolbar ("no form")
    Set frmProductos = Nothing
End Sub


Private Sub Form_Resize()
 On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500
        
        
        'Frame2.Width = ScaleWidth - CONTROL_MARGIN
        
        TDBG.Width = Me.ScaleWidth - CONTROL_MARGIN
        TDBG.Height = (Me.ScaleHeight - Me.picHeader.Height) - TDBG.top
        
    End If
    'TrueDBGridResize 0
End Sub

Public Sub TrueDBGridResize(iIndex As Integer)
    'If WindowState <> vbMaximized Then Exit Sub
    Dim i As Integer
    Dim dAnchoTotal As Double
    Dim dAnchocol As Double
    dAnchoTotal = 0
    dAnchocol = 0
    For i = 0 To Me.TDBG.Columns.Count - 1
        If (i = iIndex) Then
            dAnchocol = TDBG.Columns(i).Width
        Else
            dAnchoTotal = dAnchoTotal + TDBG.Columns(i).Width
        End If
    Next i

    Me.TDBG.Columns(iIndex).Width = (Me.ScaleWidth - dAnchoTotal) - CONTROL_MARGIN
End Sub



Private Sub TDBG_DblClick()
Dim vPosition As Variant
If rst.State = adStateClosed Then Exit Sub
If Not (rst.EOF And rst.BOF) Then
    vPosition = rst.Bookmark
    Dim frm As frmRegistrarTraslado
    Set frm = New frmRegistrarTraslado
    frm.sAccion = "View"
    frm.sDocumentoTraslado = rst("IDTraslado").value
    frm.gsFormCaption = "Traslado"
    frm.gsTitle = "Traslado"
    frm.Show
    

    Set frm = Nothing
    rst.Bookmark = vPosition
End If
End Sub

Private Sub TDBG_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not (rst.BOF And rst.EOF) Then
        If (rst!IDStatusRecibido = "16-3") Then 'El traslado no se ha recibido aun
            MDIMain.tbMenu.Buttons(21).Enabled = True
        Else
            MDIMain.tbMenu.Buttons(21).Enabled = False
        End If
    End If
End Sub

