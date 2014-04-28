VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmVendedor 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5490
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   9030
   ForeColor       =   &H00414141&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmVendedor.frx":0000
   ScaleHeight     =   5490
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   8220
      Picture         =   "frmVendedor.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
      Top             =   2940
      Width           =   555
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   555
      Left            =   8220
      Picture         =   "frmVendedor.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
      Top             =   4140
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
      Left            =   8220
      Picture         =   "frmVendedor.frx":365E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Agrega el item con los datos digitados..."
      Top             =   2340
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
      Left            =   8220
      Picture         =   "frmVendedor.frx":4328
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
      Top             =   3540
      Width           =   555
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
      Height          =   1575
      Left            =   180
      TabIndex        =   2
      Top             =   675
      Width           =   8775
      Begin VB.TextBox txtDescrTipo 
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
         Left            =   3315
         TabIndex        =   16
         Top             =   1020
         Width           =   5280
      End
      Begin VB.TextBox txtTipo 
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
         Left            =   1440
         TabIndex        =   15
         Top             =   1020
         Width           =   1080
      End
      Begin VB.TextBox txtVendedor 
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
         Left            =   1440
         TabIndex        =   14
         Top             =   585
         Width           =   1050
      End
      Begin VB.CommandButton cmdTipo 
         Height          =   320
         Left            =   2580
         Picture         =   "frmVendedor.frx":4FF2
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1020
         Width           =   300
      End
      Begin VB.CommandButton cmdDelModulo 
         Height          =   320
         Left            =   2940
         Picture         =   "frmVendedor.frx":5334
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1020
         Width           =   300
      End
      Begin VB.CheckBox chkActivo 
         Caption         =   "Activo ?"
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
         Height          =   255
         Left            =   7680
         TabIndex        =   4
         Top             =   240
         Width           =   1005
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
         Left            =   2565
         TabIndex        =   3
         Top             =   585
         Width           =   6030
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo :"
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
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor :"
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
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
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
      Left            =   8220
      Picture         =   "frmVendedor.frx":6FFE
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Deshacer / Cancelar"
      Top             =   4740
      Width           =   555
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   2955
      Left            =   150
      OleObjectBlob   =   "frmVendedor.frx":7CC8
      TabIndex        =   10
      Top             =   2340
      Width           =   7905
   End
   Begin VB.Label lbFormCaption 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Catálogo de Vendedor"
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
      Left            =   -510
      TabIndex        =   12
      Top             =   0
      Width           =   10140
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   -60
      Picture         =   "frmVendedor.frx":D709
      Stretch         =   -1  'True
      Top             =   -300
      Width           =   11490
   End
End
Attribute VB_Name = "frmVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As ADODB.Recordset
Dim bOrdenCodigo As Boolean
Dim bOrdenDescr As Boolean
Dim sCodSucursal As String
Dim Accion As TypAccion
Dim sSoloActivo As String
Public gsFormCaption As String
Public gsTitle As String

Private Sub HabilitarBotones()
    Select Case Accion
        Case TypAccion.Add, TypAccion.Edit
            cmdSave.Enabled = True
            cmdUndo.Enabled = True
            cmdEliminar.Enabled = False
            cmdAdd.Enabled = False
            cmdEditItem.Enabled = False
        Case TypAccion.View
            cmdSave.Enabled = False
            cmdUndo.Enabled = False
            cmdEliminar.Enabled = True
            cmdAdd.Enabled = True
            cmdEditItem.Enabled = True
    End Select
End Sub

Public Sub HabilitarControles()
    Select Case Accion
        Case TypAccion.Add
            txtVendedor.Enabled = True
            txtNombre.Enabled = True
            chkActivo.Enabled = True
            txtTipo.Enabled = True
            cmdTipo.Enabled = True
            txtTipo.Enabled = True
            txtTipo.Text = ""
            txtDescrTipo.Text = ""
            chkActivo.value = 1
            txtVendedor.Text = "100"
            txtNombre.Text = ""
            fmtTextbox txtVendedor, "R"
            fmtTextbox txtNombre, "O"
            Me.TDBG.Enabled = False
        Case TypAccion.Edit
            txtVendedor.Enabled = True
            txtNombre.Enabled = True
            fmtTextbox txtVendedor, "R"
            fmtTextbox txtNombre, "O"
            chkActivo.Enabled = True
            txtTipo.Enabled = True
            Me.TDBG.Enabled = False
        Case TypAccion.View
            chkActivo.Enabled = False
            cmdTipo.Enabled = False
            fmtTextbox txtVendedor, "R"
            fmtTextbox txtTipo, "R"
            fmtTextbox txtDescrTipo, "R"
            fmtTextbox txtNombre, "R"
            Me.TDBG.Enabled = True
    End Select
End Sub


Private Sub cmdAdd_Click()
    Accion = Add
    HabilitarBotones
    HabilitarControles
    txtNombre.SetFocus
End Sub

Private Sub cmdEditItem_Click()
    Accion = Edit
    GetDataFromGridToControl
    HabilitarBotones
    HabilitarControles
End Sub
Private Sub GetDataFromGridToControl()
    If Not (rst.EOF And rst.BOF) Then
        txtVendedor.Text = rst("IDVendedor").value
        txtNombre.Text = rst("Nombre").value
        If rst("Activo").value = True Then
            chkActivo.value = 1
        Else
            chkActivo.value = 0
        End If
        txtTipo.Text = rst("Tipo").value
        txtDescrTipo.Text = rst("DescrTipo").value
    Else
        txtVendedor.Text = ""
        txtNombre.Text = ""
        txtDescrTipo.Text = ""
        
    End If
End Sub

Private Sub cmdEliminar_Click()
    Dim lbOk As Boolean
    Dim sMsg As String
    Dim sTipo As String
    Dim sFiltro As String
    Dim sActivo As String
    Dim sFactura As String
    
        If txtVendedor.Text = "" Then
            lbOk = Mensaje("El Vendedor no puede estar en Blanco", ICO_ERROR, False)
            Exit Sub
        End If
        If chkActivo.value = 1 Then
           sActivo = "1"
        Else
            sActivo = "0"
        End If
        
        If txtTipo.Text = "" Then
            lbOk = Mensaje("El Vendedor no puede estar en Blanco", ICO_ERROR, False)
            Exit Sub
        End If
        ' hay que validar la integridad referencial
        lbOk = Mensaje("Está seguro de eliminar el Vendedor " & rst("Nombre").value, ICO_PREGUNTA, True)
        If lbOk Then
                    lbOk = fafUpdateVendedor("D", txtVendedor.Text, txtNombre.Text, txtTipo.Text, sActivo)
            
            If lbOk Then
                sMsg = "Borrado Exitosamente ... "
                lbOk = Mensaje(sMsg, ICO_OK, False)
                ' actualiza datos
                cargaGrid
            End If
        End If
End Sub

Private Sub cmdSave_Click()
    Dim lbOk As Boolean
    Dim sMsg As String
    Dim sActivo As String
    Dim sFactura As String
    Dim sFiltro As String
        If txtVendedor.Text = "" Then
            lbOk = Mensaje("El Vendedor no puede estar en Blanco", ICO_ERROR, False)
            Exit Sub
        End If
        If chkActivo.value = 1 Then
            sActivo = "1"
        Else
            sActivo = "0"
        End If
        If txtTipo.Text = "" Then
            lbOk = Mensaje("El Tipo del Vendedor no puede estar en Blanco", ICO_ERROR, False)
            Exit Sub
        End If
        If txtNombre.Text = "" Then
            lbOk = Mensaje("La Descripción del Centro no puede estar en blanco", ICO_ERROR, False)
            Exit Sub
        End If
        
    
            
    If Accion = Add Then
    
        If Not (rst.EOF And rst.BOF) Then
            sFiltro = "IDVendedor = '" & txtVendedor.Text & "'"
            If ExiteRstKey(rst, sFiltro) Then
               lbOk = Mensaje("Ya existe el Vendedor ", ICO_ERROR, False)
                txtVendedor.SetFocus
            Exit Sub
            End If
        End If
    
        lbOk = fafUpdateVendedor("I", txtVendedor.Text, txtNombre.Text, txtTipo.Text, sActivo)
        
        If lbOk Then
            sMsg = "El Vendedor ha sido registrada exitosamente ... "
            lbOk = Mensaje(sMsg, ICO_OK, False)
            ' actualiza datos
            Accion = View
            cargaGrid
            HabilitarControles
            HabilitarBotones
        Else
            sMsg = "Ha ocurrido un error tratando de Agregar el Vendedor... "
            lbOk = Mensaje(sMsg, ICO_ERROR, False)
        End If
      
    End If ' si estoy adicionando
    If Accion = Edit Then
        If Not (rst.EOF And rst.BOF) Then
            lbOk = fafUpdateVendedor("U", txtVendedor.Text, txtNombre.Text, txtTipo.Text, sActivo)
            If lbOk Then
                sMsg = "Los datos fueron grabados Exitosamente ... "
                lbOk = Mensaje(sMsg, ICO_OK, False)
                ' actualiza datos
                Accion = View
                cargaGrid
                HabilitarControles
                HabilitarBotones
            Else
                sMsg = "Ha ocurrido un error tratando de Actualizar el Vendedor... "
                lbOk = Mensaje(sMsg, ICO_ERROR, False)
            End If
        End If
        
    End If ' si estoy adicionando

End Sub

Private Sub cmdTipo_Click()
   Dim frm As frmBrowseCat

    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = "VENDEDOR" '& lblund.Caption
    frm.gsTablabrw = "vfafTipoVendedor"
    frm.gsCodigobrw = "Codigo"
    frm.gbTypeCodeStr = True
    frm.gsDescrbrw = "DESCR"
    frm.gbFiltra = False
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      txtTipo.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtDescrTipo.Text = frm.gsDescrbrw
      fmtTextbox txtDescrTipo, "R"
    End If
End Sub

Private Sub cmdUndo_Click()
    GetDataFromGridToControl
    Accion = View
    HabilitarControles
    HabilitarBotones
End Sub

Private Sub Form_Load()
    Set rst = New ADODB.Recordset
    If rst.State = adStateOpen Then rst.Close
    rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rst.CursorLocation = adUseClient ' Cursor local al cliente
    rst.LockType = adLockOptimistic
    Me.Caption = gsFormCaption
    Me.lbFormCaption = gsTitle
    Accion = View
    HabilitarBotones
    HabilitarControles
    cargaGrid
End Sub


Private Sub cargaGrid()
    Dim sIndependiente As String
    If rst.State = adStateOpen Then rst.Close
    rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rst.CursorLocation = adUseClient ' Cursor local al cliente
    rst.LockType = adLockOptimistic
    GSSQL = gsCompania & ".fafGetVendedores -1"
    If rst.State = adStateOpen Then rst.Close
    Set rst = GetRecordset(GSSQL)
    If Not (rst.EOF And rst.BOF) Then
      Set TDBG.DataSource = rst
      TDBG.Refresh
    End If
End Sub


Private Sub TDBG_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    GetDataFromGridToControl
    HabilitarBotones
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (rst Is Nothing) Then Set rst = Nothing
End Sub
