VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransacciones 
   Caption         =   "Form1"
   ClientHeight    =   8070
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11370
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H002F2F2F&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   11370
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   11370
      TabIndex        =   15
      Top             =   0
      Width           =   11370
      Begin VB.Image Image 
         Height          =   540
         Index           =   2
         Left            =   150
         Picture         =   "frmTransacciones.frx":0000
         Top             =   45
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Actualización del Maestro de Productos"
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
         TabIndex        =   17
         Top             =   420
         Width           =   2400
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
         TabIndex        =   16
         Top             =   90
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1485
      Left            =   10590
      TabIndex        =   12
      Top             =   3180
      Visible         =   0   'False
      Width           =   735
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
         Left            =   90
         Picture         =   "frmTransacciones.frx":0B72
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Agrega el item con los datos digitados..."
         Top             =   180
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
         Left            =   90
         Picture         =   "frmTransacciones.frx":183C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Deshacer / Cancelar"
         Top             =   810
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   120
      TabIndex        =   0
      Top             =   930
      Width           =   11130
      Begin VB.TextBox txtPaquete 
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   2100
         TabIndex        =   9
         Top             =   780
         Width           =   1275
      End
      Begin VB.TextBox txtDescrPaquete 
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   4320
         TabIndex        =   8
         Top             =   780
         Width           =   5460
      End
      Begin VB.CommandButton cmdDelclasif1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3870
         Picture         =   "frmTransacciones.frx":2506
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   780
         Width           =   330
      End
      Begin VB.CommandButton cmdPaquete 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3450
         Picture         =   "frmTransacciones.frx":41D0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   780
         Width           =   330
      End
      Begin VB.CommandButton cmbBuscar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   10110
         Picture         =   "frmTransacciones.frx":4512
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   450
         Width           =   600
      End
      Begin MSComCtl2.DTPicker dtpFechaInicial 
         Height          =   315
         Left            =   2055
         TabIndex        =   2
         Top             =   315
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
         CalendarForeColor=   4210752
         CalendarTitleForeColor=   4210752
         Format          =   61800449
         CurrentDate     =   41095
      End
      Begin MSComCtl2.DTPicker dtpFechaFinal 
         Height          =   315
         Left            =   7710
         TabIndex        =   3
         Top             =   330
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
         CalendarForeColor=   4210752
         CalendarTitleForeColor=   4210752
         Format          =   61800449
         CurrentDate     =   41095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Paquete:"
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
         Height          =   315
         Left            =   870
         TabIndex        =   10
         Top             =   810
         Width           =   825
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final:"
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
         Height          =   300
         Left            =   6450
         TabIndex        =   5
         Top             =   390
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicial:"
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
         Height          =   300
         Left            =   840
         TabIndex        =   4
         Top             =   390
         Width           =   1335
      End
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   5400
      Left            =   120
      OleObjectBlob   =   "frmTransacciones.frx":61DC
      TabIndex        =   11
      Top             =   2490
      Width           =   10365
   End
   Begin Inventario.CtlLiner CtlLiner 
      Height          =   30
      Left            =   0
      TabIndex        =   18
      Top             =   750
      Width           =   17925
      _ExtentX        =   31618
      _ExtentY        =   53
   End
End
Attribute VB_Name = "frmTransacciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As ADODB.Recordset
Public gsFormCaption As String
Public gsTitle As String

Private Sub cmbBuscar_Click()
    cargaGrid
End Sub

Private Sub cmdBuscarTranasaccion_Click()
'    frmCatalogoTransacciones.Show vbModal
'    Me.txtCodTran.Text = frmCatalogoTransacciones.sCodigo
'    Me.txtDescrTran.Text = frmCatalogoTransacciones.sDescripcion
'
'    GetTransaccion rst, Me.txtCodTran.Text, Me.dtpFechaInicial.value, Me.dtpFechaFinal.value
'    Me.tdgTransac.DataSource = rst
'
End Sub

Private Sub cargaGrid()
    If rst.State = adStateOpen Then rst.Close
    rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rst.CursorLocation = adUseClient ' Cursor local al cliente
    rst.LockType = adLockOptimistic
    GSSQL = gsCompania & ".invGetCabeceraDocumento " & IIf(Trim(Me.txtPaquete.Text) = "", -1, Val(Me.txtPaquete.Text)) & ",'*','" & Me.dtpFechaInicial.value & "','" & Me.dtpFechaFinal.value & "'"
    If rst.State = adStateOpen Then rst.Close
    Set rst = GetRecordset(GSSQL)
    If Not (rst.EOF And rst.BOF) Then
      Set TDBG.DataSource = rst
      TDBG.Refresh
    End If
End Sub


Private Sub cmdAdd_Click()
    If (Me.txtDescrPaquete.Text = "") Then
       Mensaje "Debe seleccionar un paquete", ICO_ERROR, False
        Exit Sub
    End If
    
    Dim ofrmRegistrar As New frmRegistrarTransaccion
    ofrmRegistrar.gsFormCaption = "Registrar Transacción"
    ofrmRegistrar.gsTitle = "Registrar Transacción"
    ofrmRegistrar.gsIDTipoTransaccion = CInt(Trim(Me.txtPaquete.Text))
    ofrmRegistrar.Show vbModal
End Sub

Private Sub cmdDelclasif1_Click()
    txtPaquete.Text = ""
    txtDescrPaquete.Text = ""
    cargaGrid
End Sub

Private Sub cmdPaquete_Click()
    Dim frm As frmBrowseCat
    
    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = "Lista Paquetes" '& lblund.Caption
    frm.gsTablabrw = "invPaquete"
    frm.gsCodigobrw = "IDPaquete"
    frm.gbTypeCodeStr = True
    frm.gsDescrbrw = "DESCR"
    frm.gbFiltra = False
    'frm.gsFiltro = "CATALOGO='" & lbl.Tag & "'"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      txtPaquete.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtDescrPaquete.Text = frm.gsDescrbrw
      fmtTextbox txtDescrPaquete, "R"
    End If
End Sub

Private Sub Form_Activate()
    HighlightInWin Me.Name
    SetupFormToolbar (Me.Name)
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
    fmtTextbox txtDescrPaquete, "R"
    fmtTextbox txtPaquete, "R"
    Me.dtpFechaFinal.value = DateTime.Now
    Me.dtpFechaInicial.value = DateTime.DateAdd("M", -1, DateTime.Now)
    cargaGrid
End Sub

Private Sub tdgTransac_DblClick()
'    Select Case Me.txtCodTran.Text
'        Case ParametrosGenerales.CodTranCompra ' en caso de que sea COMPRAS
'            If UserMayAccess(gNombreUsuario, SECREALIZACOMPRAPRODUCTO, GIDMODULO) Then
'                frmAgregarCompra.Accion = View
'                frmAgregarCompra.iDocumento = rst!Documento
'                frmAgregarCompra.iNumTransaccion = rst!CorTran
'                frmAgregarCompra.Show vbModal
'            End If
'        Case ParametrosGenerales.CodTranAjuste
'            If UserMayAccess(gNombreUsuario, SECCREAAJUSTEPRODUCTO, GIDMODULO) Then
'                frmAjustes.Accion = View
'                frmAjustes.iDocumento = rst!Documento
'                frmAjustes.iCorTran = rst!Documento
'                frmAjustes.Show vbModal
'            End If
'        Case ParametrosGenerales.CodTranAnulaFactura
'            If UserMayAccess(gNombreUsuario, SECREALIZAANULACIONFACTURA, GIDMODULO) Then
''                 frmAnulacionFactura.Accion = View
''                frmAnulacionFactura.iDocumento = rst!Documento
''                frmAnulacionFactura.iCorTran = rst!Documento
''
'                frmAnulacionFactura.Show vbModal
'            End If
'    End Select
End Sub


Private Sub Form_Unload(Cancel As Integer)
    SetupFormToolbar ("no name")
    'Main.SubtractForm Me.Name
    Set frmTransacciones = Nothing
End Sub

Public Sub CommandPass(ByVal srcPerformWhat As String)
    On Error GoTo err
    Select Case srcPerformWhat
        Case "Nuevo"
            cmdAdd_Click

        Case "Cerrar"
            Unload Me
      
    End Select
    Exit Sub
    'Trap the error
err:
    If err.Number = -2147467259 Then
        MsgBox "You cannot delete this record because it was used by other records! If you want to delete this record" & vbCrLf & _
               "you will first have to delete or change the records that currenly used this record as shown bellow." & vbCrLf & vbCrLf & _
               err.Description, , "Delete Operation Failed!"
        Me.MousePointer = vbDefault
    End If
End Sub




Private Sub Form_Resize()
 On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500
        
        'center_obj_horizontal Me, Frame2
        'Frame2.Width = ScaleWidth - CONTROL_MARGIN
        
        TDBG.Width = Me.ScaleWidth - CONTROL_MARGIN
        TDBG.Height = (Me.ScaleHeight - Me.picHeader.Height) - TDBG.top
        
    End If
    TrueDBGridResize 1
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


