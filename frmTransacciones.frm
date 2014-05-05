VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransacciones 
   BackColor       =   &H00F4D5BB&
   Caption         =   "Form1"
   ClientHeight    =   7035
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
   ScaleHeight     =   7035
   ScaleWidth      =   11370
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00F4D5BB&
      Height          =   1485
      Left            =   10530
      TabIndex        =   13
      Top             =   2040
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
         Picture         =   "frmTransacciones.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Picture         =   "frmTransacciones.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Deshacer / Cancelar"
         Top             =   810
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F4D5BB&
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
      TabIndex        =   1
      Top             =   585
      Width           =   11130
      Begin VB.TextBox txtPaquete 
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
         Left            =   2100
         TabIndex        =   10
         Top             =   780
         Width           =   1275
      End
      Begin VB.TextBox txtDescrPaquete 
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
         Left            =   4320
         TabIndex        =   9
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
         Picture         =   "frmTransacciones.frx":1994
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "frmTransacciones.frx":365E
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   10140
         Picture         =   "frmTransacciones.frx":39A0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   465
         Width           =   600
      End
      Begin MSComCtl2.DTPicker dtpFechaInicial 
         Height          =   315
         Left            =   2055
         TabIndex        =   3
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
         Format          =   20971521
         CurrentDate     =   41095
      End
      Begin MSComCtl2.DTPicker dtpFechaFinal 
         Height          =   315
         Left            =   7710
         TabIndex        =   4
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
         Format          =   20971521
         CurrentDate     =   41095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Paquete:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   840
         TabIndex        =   11
         Top             =   930
         Width           =   825
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   6450
         TabIndex        =   6
         Top             =   390
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicial:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   840
         TabIndex        =   5
         Top             =   390
         Width           =   1335
      End
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   4710
      Left            =   120
      OleObjectBlob   =   "frmTransacciones.frx":566A
      TabIndex        =   12
      Top             =   2130
      Width           =   10305
   End
   Begin VB.Label lbFormCaption 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Titulo Form"
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
      Left            =   -600
      TabIndex        =   0
      Top             =   0
      Width           =   12225
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   -210
      Picture         =   "frmTransacciones.frx":B727
      Stretch         =   -1  'True
      Top             =   -330
      Width           =   11850
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
    frmCatalogoTransacciones.Show vbModal
    Me.txtCodTran.Text = frmCatalogoTransacciones.sCodigo
    Me.txtDescrTran.Text = frmCatalogoTransacciones.sDescripcion
       
    GetTransaccion rst, Me.txtCodTran.Text, Me.dtpFechaInicial.value, Me.dtpFechaFinal.value
    Me.tdgTransac.DataSource = rst
    
End Sub

Private Sub cargaGrid()
    If rst.State = adStateOpen Then rst.Close
    rst.ActiveConnection = gConet 'Asocia la conexi�n de trabajo
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

Private Sub Form_Load()
    Set rst = New ADODB.Recordset
    If rst.State = adStateOpen Then rst.Close
    rst.ActiveConnection = gConet 'Asocia la conexi�n de trabajo
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


