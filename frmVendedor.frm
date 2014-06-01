VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmVendedor 
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9120
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00414141&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmVendedor.frx":0000
   ScaleHeight     =   6585
   ScaleWidth      =   9120
   WindowState     =   2  'Maximized
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
      ScaleWidth      =   9120
      TabIndex        =   17
      Top             =   0
      Width           =   9120
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
         TabIndex        =   19
         Top             =   90
         Width           =   855
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Catálogo de Bodegas"
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
         TabIndex        =   18
         Top             =   420
         Width           =   1320
      End
      Begin VB.Image Image 
         Height          =   480
         Index           =   2
         Left            =   240
         Picture         =   "frmVendedor.frx":0CCA
         Top             =   90
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   8190
      TabIndex        =   11
      Top             =   2820
      Visible         =   0   'False
      Width           =   795
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
         Left            =   120
         Picture         =   "frmVendedor.frx":190E
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Deshacer / Cancelar"
         Top             =   2640
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
         Left            =   120
         Picture         =   "frmVendedor.frx":25D8
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
         Top             =   1440
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
         Left            =   120
         Picture         =   "frmVendedor.frx":32A2
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Agrega el item con los datos digitados..."
         Top             =   240
         Width           =   555
      End
      Begin VB.CommandButton cmdSave 
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
         Left            =   120
         Picture         =   "frmVendedor.frx":3F6C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
         Top             =   2040
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
         Left            =   120
         Picture         =   "frmVendedor.frx":5C36
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
         Top             =   840
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   90
      TabIndex        =   1
      Top             =   1020
      Width           =   8775
      Begin VB.TextBox txtDescrTipo 
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
         Left            =   3315
         TabIndex        =   10
         Top             =   1020
         Width           =   5280
      End
      Begin VB.TextBox txtTipo 
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
         Left            =   1440
         TabIndex        =   9
         Top             =   1020
         Width           =   1080
      End
      Begin VB.TextBox txtVendedor 
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
         Left            =   1440
         TabIndex        =   8
         Top             =   585
         Width           =   1050
      End
      Begin VB.CommandButton cmdTipo 
         Appearance      =   0  'Flat
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
         Left            =   2580
         Picture         =   "frmVendedor.frx":6900
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1020
         Width           =   300
      End
      Begin VB.CommandButton cmdDelModulo 
         Appearance      =   0  'Flat
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
         Left            =   2940
         Picture         =   "frmVendedor.frx":6C42
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1020
         Width           =   300
      End
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
         Left            =   7680
         TabIndex        =   3
         Top             =   240
         Width           =   1005
      End
      Begin VB.TextBox txtNombre 
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
         Left            =   2565
         TabIndex        =   2
         Top             =   585
         Width           =   6030
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo :"
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
         Left            =   360
         TabIndex        =   0
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor :"
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
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   3435
      Left            =   300
      OleObjectBlob   =   "frmVendedor.frx":890C
      TabIndex        =   5
      Top             =   2730
      Width           =   7665
   End
   Begin Inventario.CtlLiner CtlLiner 
      Height          =   30
      Left            =   0
      TabIndex        =   20
      Top             =   750
      Width           =   17925
      _ExtentX        =   31618
      _ExtentY        =   53
   End
End
Attribute VB_Name = "frmVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As ADODB.Recordset
Dim Accion As TypAccion
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
    HabilitarBotonesMain
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
    Dim lbok As Boolean
    Dim sMsg As String
    Dim sActivo As String
    
        If txtVendedor.Text = "" Then
            lbok = Mensaje("El Vendedor no puede estar en Blanco", ICO_ERROR, False)
            Exit Sub
        End If
        If chkActivo.value = 1 Then
           sActivo = "1"
        Else
            sActivo = "0"
        End If
        
        If txtTipo.Text = "" Then
            lbok = Mensaje("El Vendedor no puede estar en Blanco", ICO_ERROR, False)
            Exit Sub
        End If
        ' hay que validar la integridad referencial
        lbok = Mensaje("Está seguro de eliminar el Vendedor " & rst("Nombre").value, ICO_PREGUNTA, True)
        If lbok Then
                    lbok = fafUpdateVendedor("D", txtVendedor.Text, txtNombre.Text, txtTipo.Text, sActivo)
            
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
    Dim sFiltro As String
        If txtVendedor.Text = "" Then
            lbok = Mensaje("El Vendedor no puede estar en Blanco", ICO_ERROR, False)
            Exit Sub
        End If
        If chkActivo.value = 1 Then
            sActivo = "1"
        Else
            sActivo = "0"
        End If
        If txtTipo.Text = "" Then
            lbok = Mensaje("El Tipo del Vendedor no puede estar en Blanco", ICO_ERROR, False)
            Exit Sub
        End If
        If txtNombre.Text = "" Then
            lbok = Mensaje("La Descripción del Centro no puede estar en blanco", ICO_ERROR, False)
            Exit Sub
        End If
        
    
            
    If Accion = Add Then
    
        If Not (rst.EOF And rst.BOF) Then
            sFiltro = "IDVendedor = '" & txtVendedor.Text & "'"
            If ExiteRstKey(rst, sFiltro) Then
               lbok = Mensaje("Ya existe el Vendedor ", ICO_ERROR, False)
                txtVendedor.SetFocus
            Exit Sub
            End If
        End If
    
        lbok = fafUpdateVendedor("I", txtVendedor.Text, txtNombre.Text, txtTipo.Text, sActivo)
        
        If lbok Then
            sMsg = "El Vendedor ha sido registrada exitosamente ... "
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
            lbok = fafUpdateVendedor("U", txtVendedor.Text, txtNombre.Text, txtTipo.Text, sActivo)
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
    Accion = View
    HabilitarBotones
    HabilitarControles
    cargaGrid
End Sub


Private Sub cargaGrid()
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
     SetupFormToolbar ("no name")
    MDIMain.SubtractForm Me.Name
    Set frmVendedor = Nothing
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


