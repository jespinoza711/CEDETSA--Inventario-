VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmCatalogos 
   Caption         =   "Catalogos"
   ClientHeight    =   10155
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12555
   Icon            =   "frmCatalogos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10155
   ScaleWidth      =   12555
   StartUpPosition =   1  'CenterOwner
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
      ScaleWidth      =   12555
      TabIndex        =   28
      Top             =   0
      Width           =   12555
      Begin VB.Image Image 
         Height          =   480
         Index           =   2
         Left            =   180
         Picture         =   "frmCatalogos.frx":0442
         Top             =   60
         Width           =   480
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Catálogo Genérico"
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
         TabIndex        =   30
         Top             =   420
         Width           =   1140
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
         TabIndex        =   29
         Top             =   90
         Width           =   855
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
      Height          =   495
      Left            =   11880
      Picture         =   "frmCatalogos.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
      Top             =   4020
      Width           =   495
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   495
      Left            =   11880
      Picture         =   "frmCatalogos.frx":15D6
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
      Top             =   5820
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
      Left            =   11880
      Picture         =   "frmCatalogos.frx":18E0
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Agrega el item con los datos digitados..."
      Top             =   5220
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
      Left            =   11880
      Picture         =   "frmCatalogos.frx":1BEA
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
      Top             =   4620
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Height          =   2835
      Left            =   180
      TabIndex        =   2
      Top             =   1020
      Width           =   11415
      Begin VB.CheckBox chkActivo 
         Caption         =   "Activo ?"
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
         Left            =   9930
         TabIndex        =   15
         Top             =   390
         Width           =   1095
      End
      Begin VB.TextBox txtDescr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3780
         TabIndex        =   14
         Top             =   1710
         Width           =   7335
      End
      Begin VB.TextBox txtCodigo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1230
         TabIndex        =   13
         Top             =   1710
         Width           =   1305
      End
      Begin VB.CheckBox chkDejarTabla 
         Caption         =   "Dejar seteada la Tabla ?"
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
         Left            =   210
         TabIndex        =   12
         Top             =   330
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CommandButton cmdTabla 
         Height          =   320
         Left            =   870
         Picture         =   "frmCatalogos.frx":202C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1170
         Width           =   300
      End
      Begin VB.TextBox txtTablaNombre 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3780
         TabIndex        =   10
         Top             =   1200
         Width           =   7335
      End
      Begin VB.TextBox txtIDTabla 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1230
         TabIndex        =   9
         Top             =   1200
         Width           =   1275
      End
      Begin VB.CommandButton cmdDelStatus 
         Height          =   320
         Left            =   2610
         Picture         =   "frmCatalogos.frx":236E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1200
         Width           =   300
      End
      Begin VB.TextBox txtIDCatalogo 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1380
         TabIndex        =   7
         Top             =   750
         Width           =   975
      End
      Begin VB.CheckBox chkUsaValor 
         Caption         =   "Usa Valor ?"
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
         Left            =   210
         TabIndex        =   6
         Top             =   2220
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox txtNombreValor 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3780
         TabIndex        =   5
         Top             =   2190
         Width           =   2535
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7470
         TabIndex        =   4
         Top             =   2190
         Width           =   975
      End
      Begin VB.CheckBox chkProtected 
         Caption         =   "Protegido ?"
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
         Left            =   9840
         TabIndex        =   3
         Top             =   2220
         Width           =   1305
      End
      Begin VB.Label lblDescr 
         Caption         =   "Descripción :"
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
         Left            =   2580
         TabIndex        =   22
         Top             =   1710
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Código :"
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
         Left            =   210
         TabIndex        =   21
         Top             =   1770
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Descr :"
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
         Left            =   3060
         TabIndex        =   20
         Top             =   1230
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Tabla :"
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
         Left            =   210
         TabIndex        =   19
         Top             =   1230
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "IDCatalogo :"
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
         Left            =   210
         TabIndex        =   18
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre Valor :"
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
         Left            =   2070
         TabIndex        =   17
         Top             =   2220
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Valor :"
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
         Left            =   6780
         TabIndex        =   16
         Top             =   2220
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
      Left            =   11880
      Picture         =   "frmCatalogos.frx":27B0
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Deshacer / Cancelar"
      Top             =   6420
      Width           =   495
   End
   Begin VB.CheckBox chkVerTabla 
      Caption         =   "Filtrar Tabla Seleccionada?"
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
      Height          =   405
      Left            =   9000
      TabIndex        =   0
      Top             =   4020
      Width           =   2625
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   6105
      Left            =   300
      OleObjectBlob   =   "frmCatalogos.frx":2BF2
      TabIndex        =   27
      Top             =   4470
      Width           =   11295
   End
   Begin Inventario.CtlLiner CtlLiner 
      Height          =   30
      Left            =   0
      TabIndex        =   31
      Top             =   750
      Width           =   17925
      _ExtentX        =   31618
      _ExtentY        =   53
   End
End
Attribute VB_Name = "frmCatalogos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As ADODB.Recordset

Dim bEdit As Boolean
Dim bAdd As Boolean
Public gsCatalogoName As String
Public gsFormCaption As String
Public gsTitle As String


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
HabilitarBotonesMain
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
HabilitarBotonesMain
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
Dim sUsaValor As String
    
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
HabilitarBotonesMain
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
bEdit = False
bAdd = False
initControles
cargaGrid
HabilitarBotonesMain
  Me.Caption = gsFormCaption
    Me.lbFormCaption = gsTitle
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
    SetupFormToolbar ("no name")
    Set frmCatalogos = Nothing
End Sub

Private Sub TDBG_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
GetDataFromGridToControl
'sCodSucursal = txtCodSucursal.Text
IniciaIconos
End Sub


Public Sub CommandPass(ByVal srcPerformWhat As String)
    On Error GoTo err
    Select Case srcPerformWhat
        Case "Nuevo"
            cmdAdd_Click
        Case "Editar"
            cmdEditItem_Click
        Case "Eliminar"
            cmdEliminar_Click
        Case "Cancelar"
            cmdUndo_Click
        Case "Imprimir"
            MsgBox "Imprimir"
        Case "Cerrar"
            Unload Me
        Case "Guardar"
            cmdSave_Click
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


Private Sub HabilitarBotonesMain()
    Select Case Accion
        Case TypAccion.Add, TypAccion.Edit
            MDIMain.tbMenu.Buttons(12).Enabled = True 'Guardar
            MDIMain.tbMenu.Buttons(13).Enabled = True 'Undo
            MDIMain.tbMenu.Buttons(11).Enabled = False 'Eliminar
            MDIMain.tbMenu.Buttons(8).Enabled = False 'Nuevo
            MDIMain.tbMenu.Buttons(10).Enabled = False 'Editar
        Case TypAccion.View
            MDIMain.tbMenu.Buttons(12).Enabled = False 'Guardar
            MDIMain.tbMenu.Buttons(13).Enabled = False 'Undo
            MDIMain.tbMenu.Buttons(11).Enabled = True 'Eliminar
            MDIMain.tbMenu.Buttons(8).Enabled = True 'Nuevo
            MDIMain.tbMenu.Buttons(10).Enabled = True 'Editar
    End Select
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


