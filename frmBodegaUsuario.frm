VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmBodegaUsuario 
   Caption         =   "Form1"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows Default
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
      Left            =   9330
      Picture         =   "frmBodegaUsuario.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
      Top             =   3360
      Width           =   555
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   555
      Left            =   9330
      Picture         =   "frmBodegaUsuario.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
      Top             =   4560
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
      Left            =   9330
      Picture         =   "frmBodegaUsuario.frx":2994
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Agrega el item con los datos digitados..."
      Top             =   2760
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
      Left            =   9330
      Picture         =   "frmBodegaUsuario.frx":365E
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
      Top             =   3960
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
      Height          =   1920
      Left            =   450
      TabIndex        =   1
      Top             =   570
      Width           =   8775
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
         Left            =   3315
         TabIndex        =   9
         Top             =   1020
         Width           =   5280
      End
      Begin VB.TextBox txtUsuario 
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
         TabIndex        =   8
         Top             =   1020
         Width           =   1080
      End
      Begin VB.TextBox txtBodega 
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
         TabIndex        =   7
         Top             =   585
         Width           =   1050
      End
      Begin VB.CommandButton cmdUsuario 
         Height          =   320
         Left            =   2580
         Picture         =   "frmBodegaUsuario.frx":4328
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1020
         Width           =   300
      End
      Begin VB.CommandButton cmdDelModulo 
         Height          =   320
         Left            =   2940
         Picture         =   "frmBodegaUsuario.frx":466A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1020
         Width           =   300
      End
      Begin VB.TextBox txtDescrBodega 
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
         TabIndex        =   4
         Top             =   585
         Width           =   6030
      End
      Begin VB.CheckBox chkFactura 
         Caption         =   "El Usuario puede Facturar ?"
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
         Left            =   1440
         TabIndex        =   3
         Top             =   1440
         Width           =   2685
      End
      Begin VB.CheckBox chkConsultaInv 
         Caption         =   "El Usuario puede Consultar Inventario ?"
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
         Left            =   4320
         TabIndex        =   2
         Top             =   1440
         Width           =   3765
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario :"
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
         TabIndex        =   11
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bodega :"
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
         TabIndex        =   10
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
      Left            =   9330
      Picture         =   "frmBodegaUsuario.frx":6334
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Deshacer / Cancelar"
      Top             =   5160
      Width           =   555
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   2955
      Left            =   330
      OleObjectBlob   =   "frmBodegaUsuario.frx":6FFE
      TabIndex        =   16
      Top             =   2730
      Width           =   7905
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
      Height          =   495
      Left            =   30
      TabIndex        =   17
      Top             =   -30
      Width           =   10140
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   -30
      Picture         =   "frmBodegaUsuario.frx":E0BB
      Stretch         =   -1  'True
      Top             =   -30
      Width           =   11490
   End
End
Attribute VB_Name = "frmBodegaUsuario"
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
Public gsIDBodega As String
Public gsDescrBodega As String

Private Sub HabilitarBotones()
    Select Case Accion
        Case TypAccion.Add, TypAccion.Edit
            cmdSave.Enabled = True
            cmdUndo.Enabled = True
            cmdEliminar.Enabled = False
            cmdAdd.Enabled = False
            cmdEditItem.Enabled = False
        Case TypAccion.View
            If rst.State = adStateClosed Then
                cmdEditItem.Enabled = False
                cmdSave.Enabled = False
                cmdUndo.Enabled = False
                cmdEliminar.Enabled = False
                cmdAdd.Enabled = True
                Exit Sub
            End If
            If rst.RecordCount <> 0 Then
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
End Sub

Public Sub HabilitarControles()
    Select Case Accion
        Case TypAccion.Add
            txtBodega.Enabled = False
            txtDescrBodega.Enabled = False
            chkFactura.Enabled = True
            chkFactura.value = 0
            chkConsultaInv.Enabled = True
            chkConsultaInv.value = 0
            txtUsuario.Enabled = True
            cmdUsuario.Enabled = True
            txtUsuario.Enabled = True
            txtUsuario.Text = ""
            txtNombre.Text = ""

            fmtTextbox txtBodega, "R"
            fmtTextbox txtDescrBodega, "R"
            TDBG.Enabled = False
        Case TypAccion.Edit
            txtBodega.Enabled = True
            txtDescrBodega.Enabled = True
            fmtTextbox txtBodega, "R"
            fmtTextbox txtDescrBodega, "R"
            chkFactura.Enabled = True
            chkConsultaInv.Enabled = True
            txtUsuario.Enabled = True
            Me.TDBG.Enabled = False
        Case TypAccion.View
            chkFactura.Enabled = False
            chkConsultaInv.Enabled = False
            cmdUsuario.Enabled = False
            fmtTextbox txtBodega, "R"
            fmtTextbox txtUsuario, "R"
            fmtTextbox txtNombre, "R"
            fmtTextbox txtDescrBodega, "R"
            TDBG.Enabled = True
    End Select
End Sub


Private Sub cmdAdd_Click()
    Accion = Add
    HabilitarBotones
    HabilitarControles
    cmdUsuario.SetFocus
End Sub

Private Sub cmdEditItem_Click()
    Accion = Edit
    GetDataFromGridToControl
    HabilitarBotones
    HabilitarControles
End Sub
Private Sub GetDataFromGridToControl()
    If Not (rst.EOF And rst.BOF) Then
        txtNombre.Text = rst("Nombre").value
        txtUsuario.Text = rst("Usuario").value
        If rst("Factura").value = True Then
            chkFactura.value = 1
        Else
            chkFactura.value = 0
        End If
        If rst("ConsultaInv").value = True Then
            chkConsultaInv.value = 1
        Else
            chkConsultaInv.value = 0
        End If
        
    Else
        txtUsuario.Text = ""
        txtNombre.Text = ""
        
    End If
End Sub

Private Sub cmdEliminar_Click()
    Dim lbok As Boolean
    Dim sMsg As String
    Dim sTipo As String
    Dim sFiltro As String
    Dim sFactura As String
    Dim sConsultaInv As String
    
        If txtBodega.Text = "" Then
            lbok = Mensaje("El Vendedor no puede estar en Blanco", ICO_ERROR, False)
            Exit Sub
        End If
        If chkFactura.value = 1 Then
           sFactura = "1"
        Else
            sFactura = "0"
        End If
        If chkConsultaInv.value = 1 Then
           sConsultaInv = "1"
        Else
            sConsultaInv = "0"
        End If
        
        If txtUsuario.Text = "" Then
            lbok = Mensaje("El Usuario no puede estar en Blanco", ICO_ERROR, False)
            Exit Sub
        End If
        ' hay que validar la integridad referencial
        lbok = Mensaje("Está seguro de eliminar el Vendedor " & rst("Nombre").value, ICO_PREGUNTA, True)
        If lbok Then
                    lbok = invUpdateBodegaUsuario("D", txtBodega.Text, txtUsuario.Text, sFactura, sConsultaInv)
            
            If lbok Then
                sMsg = "Borrado Exitosamente ... "
                lbok = Mensaje(sMsg, ICO_OK, False)
                ' actualiza datos
                cargaGrid
            End If
        End If
End Sub

Private Sub cmdSave_Click()
    Dim lbok As Boolean
    Dim sMsg As String
    Dim sActivo As String
    Dim sFactura As String
    Dim sConsInv As String
    Dim sFiltro As String
        If txtBodega.Text = "" Then
            lbok = Mensaje("La Bodega no puede estar en Blanco", ICO_ERROR, False)
            Exit Sub
        End If
        If chkFactura.value = 1 Then
            sFactura = "1"
        Else
            sFactura = "0"
        End If
        If chkConsultaInv.value = 1 Then
            sConsInv = "1"
        Else
            sConsInv = "0"
        End If
        If txtUsuario.Text = "" Then
            lbok = Mensaje("Usuario no puede estar en Blanco", ICO_ERROR, False)
            Exit Sub
        End If
            
    If Accion = Add Then
    
        If Not (rst.EOF And rst.BOF) Then
            sFiltro = "IDBodega = '" & txtBodega.Text & "'" & " and Usuario='" & txtUsuario.Text & "'"
            If ExiteRstKey(rst, sFiltro) Then
               lbok = Mensaje("Ya existe el usuario en esa bodega ", ICO_ERROR, False)
                txtUsuario.SetFocus
            Exit Sub
            End If
        End If
    
        lbok = invUpdateBodegaUsuario("I", txtBodega.Text, txtUsuario.Text, sFactura, sConsInv)
        
        If lbok Then
            sMsg = "El Usuario ha sido registrado exitosamente ... "
            lbok = Mensaje(sMsg, ICO_OK, False)
            ' actualiza datos
            Accion = View
            cargaGrid
            HabilitarControles
            HabilitarBotones
        Else
            sMsg = "Ha ocurrido un error tratando de Agregar el Vendedor... "
            lbok = Mensaje(sMsg, ICO_ERROR, False)
        End If
      
    End If ' si estoy adicionando
    If Accion = Edit Then
        If Not (rst.EOF And rst.BOF) Then
            lbok = invUpdateBodegaUsuario("U", txtBodega.Text, txtUsuario.Text, sFactura, sConsInv)
            If lbok Then
                sMsg = "Los datos fueron grabados Exitosamente ... "
                lbok = Mensaje(sMsg, ICO_OK, False)
                ' actualiza datos
                Accion = View
                cargaGrid
                HabilitarControles
                HabilitarBotones
            Else
                sMsg = "Ha ocurrido un error tratando de Actualizar el Vendedor... "
                lbok = Mensaje(sMsg, ICO_ERROR, False)
            End If
        End If
        
    End If ' si estoy adicionando

End Sub

Private Sub cmdUsuario_Click()
   Dim frm As frmBrowseCat

    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = "Usuarios" '& lblund.Caption
    frm.gsTablabrw = "secUsuario"
    frm.gsCodigobrw = "Usuario"
    frm.gbTypeCodeStr = True
    frm.gsDescrbrw = "DESCR"
    frm.gbFiltra = True
    frm.gsFiltro = "Activo=1"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      txtUsuario.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtNombre.Text = frm.gsDescrbrw
      fmtTextbox txtNombre, "R"
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
    'sFormCaption = "Catalogo de Vendedores"
    Caption = gsFormCaption
    lbFormCaption = gsTitle
    Accion = View
    HabilitarBotones
    HabilitarControles
    txtDescrBodega.Text = gsDescrBodega
    txtBodega.Text = gsIDBodega
    cargaGrid
End Sub


Private Sub cargaGrid()
    Dim sIndependiente As String
    If rst.State = adStateOpen Then rst.Close
    rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rst.CursorLocation = adUseClient ' Cursor local al cliente
    rst.LockType = adLockOptimistic
    GSSQL = gsCompania & ".getBodegaUsuario -1"
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

