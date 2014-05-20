VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmCatalogos 
   BackColor       =   &H00FEE3DA&
   Caption         =   "Catalogos"
   ClientHeight    =   8205
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13800
   Icon            =   "frmCatalogos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   13800
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
      Height          =   495
      Left            =   12720
      Picture         =   "frmCatalogos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   495
      Left            =   12720
      Picture         =   "frmCatalogos.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
      Top             =   4200
      Width           =   495
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
      Height          =   495
      Left            =   12720
      Picture         =   "frmCatalogos.frx":1016
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Agrega el item con los datos digitados..."
      Top             =   3600
      Width           =   495
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
      Height          =   495
      Left            =   12720
      Picture         =   "frmCatalogos.frx":1320
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
      Top             =   3000
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   480
      TabIndex        =   2
      Top             =   0
      Width           =   10695
      Begin VB.CheckBox chkActivo 
         Caption         =   "Activo ?"
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
         Left            =   8160
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtDescr 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3240
         TabIndex        =   14
         Top             =   1080
         Width           =   7095
      End
      Begin VB.TextBox txtCodigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox chkDejarTabla 
         Caption         =   "Dejar seteada la Tabla ?"
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
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CommandButton cmdTabla 
         Height          =   320
         Left            =   720
         Picture         =   "frmCatalogos.frx":1762
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   600
         Width           =   300
      End
      Begin VB.TextBox txtTablaNombre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3240
         TabIndex        =   10
         Top             =   600
         Width           =   7095
      End
      Begin VB.TextBox txtIDTabla 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdDelStatus 
         Height          =   320
         Left            =   2040
         Picture         =   "frmCatalogos.frx":1AA4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   600
         Width           =   300
      End
      Begin VB.TextBox txtIDCatalogo 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkUsaValor 
         Caption         =   "Usa Valor ?"
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
         TabIndex        =   6
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox txtNombreValor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3720
         TabIndex        =   5
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7320
         TabIndex        =   4
         Top             =   1560
         Width           =   975
      End
      Begin VB.CheckBox chkProtected 
         Caption         =   "Protegido ?"
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
         Left            =   8760
         TabIndex        =   3
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblDescr 
         Caption         =   "Descripción :"
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
         Left            =   2040
         TabIndex        =   22
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Código :"
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
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Descr :"
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
         Left            =   2400
         TabIndex        =   20
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Tabla :"
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
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "IDCatalogo :"
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
         Left            =   4920
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre Valor :"
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
         Left            =   1920
         TabIndex        =   17
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Valor :"
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
         Left            =   6480
         TabIndex        =   16
         Top             =   1560
         Width           =   615
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
      Height          =   495
      Left            =   12720
      Picture         =   "frmCatalogos.frx":1EE6
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Deshacer / Cancelar"
      Top             =   4800
      Width           =   495
   End
   Begin VB.CheckBox chkVerTabla 
      Caption         =   "Filtrar Tabla Seleccionada?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11400
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   5655
      Left            =   0
      OleObjectBlob   =   "frmCatalogos.frx":2328
      TabIndex        =   27
      Top             =   2160
      Width           =   12465
   End
End
Attribute VB_Name = "frmCatalogos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As ADODB.Recordset
Dim bOrdenCodigo As Boolean
Dim bOrdenDescr As Boolean
Dim sCodSucursal As String
Dim bEdit As Boolean
Dim bAdd As Boolean
Dim sSoloActivo As String
Public gsCatalogoName As String


Private Sub chkUsaValor_Click()
If chkUsaValor.value = 1 Then
    txtNombreValor.Enabled = True
    txtValor.Enabled = True
Else
    txtNombreValor.Enabled = False
    txtValor.Enabled = False
End If
End Sub

Private Sub chkVerTabla_Click()
cargaGrid
End Sub

Private Sub cmdAdd_Click()
bAdd = True
bEdit = False
If chkDejarTabla.value = False Then
    txtIDTabla.Text = ""
    txtTablaNombre.Text = ""
End If
txtCodigo.Enabled = False
txtDescr.Enabled = True
chkActivo.Enabled = True
chkUsaValor.Enabled = True
chkUsaValor.value = 0
chkActivo.value = 1
chkProtected.value = 0
txtCodigo.Text = "1000"
txtIDCatalogo.Text = "00"
txtDescr.Text = ""
'txtNombreValor.Enabled = True
txtNombreValor.Text = "ND"
'txtValor.Enabled = True
txtValor.Text = 0
fmtTextbox txtCodigo, "R"
fmtTextbox txtDescr, "O"
cmdSave.Enabled = True
cmdEliminar.Enabled = False
cmdAdd.Enabled = False
txtDescr.SetFocus
End Sub

Private Sub cmdDelStatus_Click()
txtIDTabla.Text = ""
txtTablaNombre.Text = ""
cargaGrid
End Sub

Private Sub cmdEditItem_Click()
Dim lbok  As Boolean

If Me.txtCodigo.Text = "0" Then
    lbok = Mensaje("El valor No Definido ND es un valor protegido por el sistema... Ud no puede modificarlo.", ICO_ADVERTENCIA, False)
    Exit Sub
End If

If Not (rst.EOF And rst.BOF) Then
    If rst("Protected").value = True Then
        lbok = Mensaje("El valor del Registro es valor Protegido por el sistema... Ud no puede modificarlo.", ICO_ADVERTENCIA, False)
    
        Exit Sub
    End If
        
End If

bEdit = True
bAdd = False
GetDataFromGridToControl
txtCodigo.Enabled = True
txtDescr.Enabled = True
fmtTextbox txtCodigo, "R"
fmtTextbox txtDescr, "O"
chkActivo.Enabled = True
chkUsaValor.Enabled = True
txtNombreValor.Enabled = True
txtValor.Enabled = True
cmdSave.Enabled = True
cmdEliminar.Enabled = False
cmdAdd.Enabled = False
End Sub
Private Sub GetDataFromGridToControl()
If Not (rst.EOF And rst.BOF) Then
    txtIDCatalogo.Text = rst("IDCatalogo").value
    txtIDTabla.Text = rst("IDTabla").value
    txtTablaNombre.Text = rst("Nombre").value
    txtCodigo.Text = rst("Codigo").value
    txtDescr.Text = rst("Descr").value
    If rst("Activo").value = True Then
        chkActivo.value = 1
    Else
        chkActivo.value = 0
    End If
    
    If rst("Protected").value = True Then
        chkProtected.value = 1
    Else
        chkProtected.value = 0
    End If
    
    If rst("UsaValor").value = True Then
        chkUsaValor.value = 1
    Else
        chkUsaValor.value = 0
    End If
    txtNombreValor.Text = rst("NombreValor").value
    txtValor.Text = rst("Valor").value
Else
    txtCodigo.Text = ""
    txtDescr.Text = ""
End If

End Sub

Private Sub cmdEliminar_Click()
Dim lbok As Boolean
Dim sMsg As String
Dim sTipo As String
Dim sFiltro As String
Dim sActivo As String
Dim sUsaValor As String

If txtCodigo.Text = "0" Then
    lbok = Mensaje("El valor No Definido ND es un valor protegido por el sistema... Ud no puede modificarlo.", ICO_ADVERTENCIA, False)
    Exit Sub
End If

If Not (rst.EOF And rst.BOF) Then
    If rst("Protected").value = True Then
        lbok = Mensaje("El valor del Registro es valor Protegido por el sistema... Ud no puede eliminarlo.", ICO_ADVERTENCIA, False)
    
        Exit Sub
    End If
        
End If

    If txtIDCatalogo.Text = "" Then
        lbok = Mensaje("El ID del Catálogo no puede estar en Blanco", ICO_ERROR, False)
        Exit Sub
    End If


    If txtIDTabla.Text = "" Then
        lbok = Mensaje("El Código de la tabla no puede estar en Blanco", ICO_ERROR, False)
        Exit Sub
    End If
    
    If txtCodigo.Text = "" Then
        lbok = Mensaje("El Código del elemento de la tabla no puede estar en Blanco", ICO_ERROR, False)
        Exit Sub
    End If
    
    If chkActivo.value = 1 Then
        sActivo = "1"
    Else
        sActivo = "0"
    End If
    If chkUsaValor.value = 1 Then
        sUsaValor = "1"
    Else
        sUsaValor = "0"
    End If
    ' hay que validar la integridad referencial
    lbok = Mensaje("Está seguro de eliminar el Registro " & rst("Descr").value, ICO_PREGUNTA, True)
    If lbok Then
                lbok = spGlobalUpdateCatalogo("D", txtIDTabla.Text, txtDescr.Text, sActivo, sUsaValor, txtNombreValor.Text, txtValor.Text, txtIDCatalogo.Text)
        
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
Dim sbkDocumentos As String
Dim sUsaValor As String
Dim sFiltro As String
    If txtIDCatalogo.Text = "" Then
        lbok = Mensaje("El ID del Catálogo no puede estar en Blanco", ICO_ERROR, False)
        Exit Sub
    End If


    If txtIDTabla.Text = "" Then
        lbok = Mensaje("El Código de la tabla no puede estar en Blanco", ICO_ERROR, False)
        Exit Sub
    End If
    
'    If txtCodigo.Text = "" Then
'        lbok = Mensaje("El Código del elemento de la tabla no puede estar en Blanco", ICO_ERROR, False)
'        Exit Sub
'    End If

    If Not Val_TextboxNum(txtValor) Then
        txtValor.Text = "0"
        lbok = Mensaje("El Valor debe ser numérico", ICO_ERROR, False)
        Exit Sub
    End If
    
    If Not Val_TextboxNum(txtIDTabla) Then
        txtIDTabla.Text = ""
        lbok = Mensaje("El Código de la tabla debe ser numérico", ICO_ERROR, False)
        Exit Sub
    End If

'    If Not Val_TextboxNum(txtCodigo) Then
'        lbok = Mensaje("El Código del elemento de la tabla debe ser numérico ", ICO_ERROR, False)
'        Exit Sub
'    End If
    

    If txtCodigo.Text = "0" Then
        lbok = Mensaje("El valor No Definido ND es un valor protegido por el sistema... Ud no puede agregarlo.", ICO_ADVERTENCIA, False)
        Exit Sub
    End If
    
    If Me.chkActivo.value = 1 Then
        sActivo = "1"
    Else
        sActivo = "0"
    End If
    
    
    If Me.chkUsaValor.value = 1 Then
        sUsaValor = "1"
    Else
        sUsaValor = "0"
    End If
    
    If txtDescr.Text = "" Then
        lbok = Mensaje("La Descripción del elemento de la tabla no puede estar en blanco", ICO_ERROR, False)
        Exit Sub
    End If
    

        
If bAdd Then


        lbok = spGlobalUpdateCatalogo("I", txtIDTabla.Text, txtDescr.Text, sActivo, sUsaValor, txtNombreValor.Text, txtValor.Text, txtIDCatalogo.Text)
        
        If lbok Then
            sMsg = "El Codigo ha sido registrado exitosamente ... "
            lbok = Mensaje(sMsg, ICO_OK, False)
            ' actualiza datos
            cargaGrid
            bEdit = False
            bAdd = False
            initControles
            IniciaIconos
        End If
bAdd = False
End If ' si estoy adicionando
If bEdit Then
    If Not (rst.EOF And rst.BOF) Then
        lbok = spGlobalUpdateCatalogo("U", txtIDTabla.Text, txtDescr.Text, sActivo, sUsaValor, txtNombreValor.Text, txtValor.Text, txtIDCatalogo.Text)
        If lbok Then
            sMsg = "Los datos fueron grabados Exitosamente ... "
            lbok = Mensaje(sMsg, ICO_OK, False)
            ' actualiza datos
            cargaGrid
            bEdit = False
            bAdd = False
            initControles
            IniciaIconos
        End If
    End If
bEdit = False
End If ' si estoy adicionando

End Sub

Private Sub cmdStatus_Click()

End Sub

Private Sub cmdTabla_Click()
   Dim frm As frmBrowseCat
    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = "Tabla" '& lblund.Caption
    frm.gsTablabrw = "globalTABLAS"
    frm.gsCodigobrw = "IDTabla"
    frm.gbTypeCodeStr = False
    frm.gsDescrbrw = "Nombre"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      txtIDTabla = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtTablaNombre = frm.gsDescrbrw
       fmtTextbox txtTablaNombre, "R"
    End If
    If txtIDTabla.Text <> "" Then
        If bAdd = False And bEdit = False Then
            cargaGrid
            'txtDescr.SetFocus
        End If
        
    End If
    
End Sub

Private Sub cmdUndo_Click()
GetDataFromGridToControl
IniciaIconos
End Sub

Private Sub Form_Load()
Set rst = New ADODB.Recordset
If rst.State = adStateOpen Then rst.Close
rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
rst.CursorLocation = adUseClient ' Cursor local al cliente
rst.LockType = adLockOptimistic
bEdit = False
bAdd = False
initControles
cargaGrid
End Sub

Private Sub IniciaIconos()
cmdSave.Enabled = False
cmdEditItem.Enabled = True
cmdEliminar.Enabled = True
cmdAdd.Enabled = True
bEdit = False
bAdd = False

End Sub
Private Sub initControles()
txtCodigo.Enabled = False
txtDescr.Enabled = False
chkActivo.Enabled = False
txtIDTabla.Enabled = False
txtTablaNombre.Enabled = False
chkUsaValor.Enabled = False
txtNombreValor.Enabled = False
txtValor.Enabled = False
chkProtected.Enabled = False
End Sub

Private Sub cargaGrid()
Dim sTabla As String
If rst.State = adStateOpen Then rst.Close
rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
rst.CursorLocation = adUseClient ' Cursor local al cliente
rst.LockType = adLockOptimistic
Caption = gsCatalogoName
If chkVerTabla.value = 0 Then
    sTabla = -1 ' Mostrará Todas las Tablas
Else
    sTabla = txtIDTabla.Text
End If
GSSQL = gsCompania & ".globalGetCatalogos " & sTabla
If rst.State = adStateOpen Then rst.Close
Set rst = GetRecordset(GSSQL)
If Not (rst.EOF And rst.BOF) Then
  Set TDBG.DataSource = rst
  'CargarDatos rst, TDBG, "Codigo", "Descr"
    TDBG.Refresh
    TDBG.Columns(4).FooterAlignment = dbgRight
    TDBG.Columns(4).FooterText = " Total de elementos :    " & Str(rst.RecordCount)
  
  'IniciaIconos
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not (rst Is Nothing) Then Set rst = Nothing
End Sub

Private Sub TDBG_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
GetDataFromGridToControl
'sCodSucursal = txtCodSucursal.Text
IniciaIconos
End Sub



