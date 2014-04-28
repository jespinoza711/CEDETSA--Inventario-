VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmCatalogos 
   Caption         =   "Catalogos"
   ClientHeight    =   8205
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11100
   Icon            =   "frmCatalogos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   11100
   StartUpPosition =   1  'CenterOwner
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
      Left            =   11640
      TabIndex        =   26
      Top             =   1320
      Width           =   1815
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
      Left            =   12960
      Picture         =   "frmCatalogos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Deshacer / Cancelar"
      Top             =   4920
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   720
      TabIndex        =   6
      Top             =   120
      Width           =   10695
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
         Left            =   9360
         TabIndex        =   24
         Top             =   1560
         Width           =   975
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
         Left            =   5760
         TabIndex        =   22
         Top             =   1560
         Width           =   2535
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
         Left            =   2280
         TabIndex        =   21
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox txtIDCatalogo 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6120
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdDelStatus 
         Height          =   320
         Left            =   2040
         Picture         =   "frmCatalogos.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   600
         Width           =   300
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
         TabIndex        =   15
         Top             =   600
         Width           =   855
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
         TabIndex        =   14
         Top             =   600
         Width           =   7095
      End
      Begin VB.CommandButton cmdTabla 
         Height          =   320
         Left            =   720
         Picture         =   "frmCatalogos.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   600
         Width           =   300
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
         TabIndex        =   0
         Top             =   1080
         Width           =   855
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
         TabIndex        =   1
         Top             =   1080
         Width           =   7095
      End
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
         TabIndex        =   7
         Top             =   240
         Width           =   1095
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
         Left            =   8520
         TabIndex        =   25
         Top             =   1560
         Width           =   615
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
         Left            =   3960
         TabIndex        =   23
         Top             =   1560
         Width           =   1455
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
         TabIndex        =   20
         Top             =   240
         Width           =   1095
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
         TabIndex        =   18
         Top             =   600
         Width           =   615
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
         TabIndex        =   17
         Top             =   600
         Width           =   735
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
         TabIndex        =   9
         Top             =   1080
         Width           =   735
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
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
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
      Left            =   12960
      Picture         =   "frmCatalogos.frx":1008
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
      Top             =   3120
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
      Left            =   12960
      Picture         =   "frmCatalogos.frx":144A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Agrega el item con los datos digitados..."
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   495
      Left            =   12960
      Picture         =   "frmCatalogos.frx":1754
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
      Top             =   4320
      Width           =   495
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
      Height          =   495
      Left            =   12960
      Picture         =   "frmCatalogos.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
      Top             =   2520
      Width           =   495
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   5655
      Left            =   240
      OleObjectBlob   =   "frmCatalogos.frx":2328
      TabIndex        =   11
      Top             =   2280
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
chkActivo.value = 1
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


