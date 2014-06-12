VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form frmFiltroExistenciaProducto 
   Appearance      =   0  'Flat
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8415
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
   ScaleHeight     =   5535
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   3060
      TabIndex        =   20
      Top             =   5010
      Width           =   990
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   4620
      TabIndex        =   19
      Top             =   5010
      Width           =   990
   End
   Begin ActiveTabs.SSActiveTabs tabFiltroExistencia 
      Height          =   4425
      Left            =   180
      TabIndex        =   0
      Top             =   510
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   7805
      _Version        =   131083
      TabCount        =   2
      Tabs            =   "frmFiltroExistenciaProducto.frx":0000
      Begin ActiveTabs.SSActiveTabPanel stbLineas 
         Height          =   4035
         Left            =   -99969
         TabIndex        =   11
         Top             =   360
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   7117
         _Version        =   131083
         TabGuid         =   "frmFiltroExistenciaProducto.frx":0084
         Begin VB.CheckBox chkSeleccionarAllSubFamilias 
            Caption         =   "Seleccionar todas las SubFamilias"
            Height          =   375
            Left            =   5490
            TabIndex        =   24
            Top             =   3540
            Width           =   2625
         End
         Begin VB.CheckBox chkSelecionarAllFamilias 
            Caption         =   "Seleccionar todas las Familias"
            Height          =   285
            Left            =   2790
            TabIndex        =   23
            Top             =   3570
            Width           =   2445
         End
         Begin VB.CheckBox chkSeleccionarAllLineas 
            Caption         =   "Seleccionar todas las Lineas"
            Height          =   285
            Left            =   180
            TabIndex        =   22
            Top             =   3600
            Width           =   2445
         End
         Begin VB.ListBox lstSubFamilias 
            ForeColor       =   &H00404040&
            Height          =   2985
            Left            =   5430
            Style           =   1  'Checkbox
            TabIndex        =   14
            Top             =   510
            Width           =   2385
         End
         Begin VB.ListBox lstFamilias 
            ForeColor       =   &H00404040&
            Height          =   2985
            Left            =   2790
            Style           =   1  'Checkbox
            TabIndex        =   13
            Top             =   510
            Width           =   2385
         End
         Begin VB.ListBox lstLineas 
            ForeColor       =   &H00404040&
            Height          =   2985
            Left            =   180
            Style           =   1  'Checkbox
            TabIndex        =   12
            Top             =   510
            Width           =   2385
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Listado de Familias:"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   2820
            TabIndex        =   17
            Top             =   270
            Width           =   1500
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Listado de SubFamilias:"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   5430
            TabIndex        =   16
            Top             =   270
            Width           =   1665
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Listado de Lineas:"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   210
            TabIndex        =   15
            Top             =   270
            Width           =   1290
         End
      End
      Begin ActiveTabs.SSActiveTabPanel sPanelBodega 
         Height          =   4035
         Left            =   30
         TabIndex        =   1
         Top             =   360
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   7117
         _Version        =   131083
         TabGuid         =   "frmFiltroExistenciaProducto.frx":00AC
         Begin VB.CheckBox chkSeleccionarAllBodegas 
            Caption         =   "Seleccionar todas las bodegas"
            Height          =   285
            Left            =   240
            TabIndex        =   21
            Top             =   3660
            Width           =   2685
         End
         Begin VB.Frame frmDetalla 
            Caption         =   " Métodos de Visualización "
            ForeColor       =   &H00404040&
            Height          =   2145
            Left            =   3720
            TabIndex        =   6
            Top             =   570
            Width           =   3135
            Begin VB.OptionButton optDetallaXArticulo 
               Caption         =   "Detalla por Articulo"
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   210
               TabIndex        =   10
               Top             =   1500
               Width           =   2355
            End
            Begin VB.OptionButton optDetallaXBodega 
               Caption         =   "Detalla por Bodega"
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   210
               TabIndex        =   9
               Top             =   1140
               Width           =   2355
            End
            Begin VB.OptionButton optConsolidaXArticulo 
               Caption         =   "Consolidar por Articulo"
               ForeColor       =   &H00404040&
               Height          =   405
               Left            =   210
               TabIndex        =   8
               Top             =   750
               Width           =   2355
            End
            Begin VB.OptionButton optConsolidaXBodega 
               Caption         =   "Consolidar por Bodega"
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   210
               TabIndex        =   7
               Top             =   420
               Width           =   2355
            End
         End
         Begin VB.CheckBox chkIncluirExistenciaTransito 
            Caption         =   "Incluir Existencias en Transito"
            ForeColor       =   &H00404040&
            Height          =   435
            Left            =   3840
            TabIndex        =   5
            Top             =   2850
            Width           =   2625
         End
         Begin VB.CheckBox chkSoloProductosConExistencia 
            Caption         =   "&Solo productos con Existencia"
            ForeColor       =   &H00404040&
            Height          =   435
            Left            =   3840
            TabIndex        =   4
            Top             =   3300
            Width           =   2625
         End
         Begin VB.ListBox lstBodegas 
            ForeColor       =   &H00404040&
            Height          =   2760
            Left            =   240
            Style           =   1  'Checkbox
            TabIndex        =   2
            Top             =   810
            Width           =   3105
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Listado de Bodegas:"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   210
            TabIndex        =   3
            Top             =   510
            Width           =   1455
         End
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filtro de Existencia de Productos"
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
      Left            =   2400
      TabIndex        =   18
      Top             =   120
      Width           =   3675
   End
End
Attribute VB_Name = "frmFiltroExistenciaProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstBodegas As New ADODB.Recordset
Dim rstLineas As New ADODB.Recordset
Dim rstFamilias As New ADODB.Recordset
Dim rstSubFamilias As New ADODB.Recordset

Dim sBodegasSeleccionadas  As String
Dim sLineasSeleccionadas As String
Dim sFamiliasSeleccionadas As String
Dim sSubFamiliasSeleccionadas As String

Enum TypModoVisualizacion
    ConsolidadobyBodega
    ConsolidadobyProducto
    DetalladabyBodega
    DetalladaByProducto
End Enum

Dim Modovisualizacion As TypModoVisualizacion
Dim bSoloProductosConExistencia As Boolean
Dim bIncluirExistenciaTransito As Boolean


Private Sub Command1_Click()
    GetDatosPestanaBodega
End Sub

Private Sub chkSeleccionarAllBodegas_Click()
    Dim i As Integer
    If (Me.chkSeleccionarAllBodegas.Value = vbChecked) Then
        For i = 0 To Me.lstBodegas.ListCount - 1
            Me.lstBodegas.Selected(i) = True
        Next i
    Else
        For i = 0 To Me.lstBodegas.ListCount - 1
            Me.lstBodegas.Selected(i) = False
        Next i
    End If
End Sub

Private Sub chkSeleccionarAllLineas_Click()
    Dim i As Integer
    If (Me.chkSeleccionarAllLineas.Value = vbChecked) Then
        For i = 0 To Me.lstLineas.ListCount - 1
            Me.lstLineas.Selected(i) = True
        Next i
    Else
        For i = 0 To Me.lstLineas.ListCount - 1
            Me.lstLineas.Selected(i) = False
        Next i
    End If
End Sub

Private Sub chkSeleccionarAllSubFamilias_Click()
    Dim i As Integer
    If (Me.chkSeleccionarAllSubFamilias.Value = vbChecked) Then
        For i = 0 To Me.lstSubFamilias.ListCount - 1
            Me.lstSubFamilias.Selected(i) = True
        Next i
    Else
        For i = 0 To Me.lstSubFamilias.ListCount - 1
            Me.lstSubFamilias.Selected(i) = False
        Next i
    End If
End Sub

Private Sub chkSelecionarAllFamilias_Click()
    Dim i As Integer
    If (Me.chkSelecionarAllFamilias.Value = vbChecked) Then
        For i = 0 To Me.lstFamilias.ListCount - 1
            Me.lstFamilias.Selected(i) = True
        Next i
    Else
        For i = 0 To Me.lstFamilias.ListCount - 1
            Me.lstFamilias.Selected(i) = False
        Next i
    End If
End Sub

Private Sub Form_Load()
    PreparaRstBodegas
    PreparaRstLineas
    PreparaRstFamilias
    PreparaRstSubFamilias
    LlenarListaBodegas
    LlenarListaClasificaciones Me.lstLineas, rstLineas
    LlenarListaClasificaciones Me.lstFamilias, rstFamilias
    LlenarListaClasificaciones Me.lstSubFamilias, rstSubFamilias
End Sub

Private Sub PreparaRstBodegas()
      Set rstBodegas = New ADODB.Recordset
      If rstBodegas.State = adStateOpen Then rstBodegas.Close
      rstBodegas.ActiveConnection = gConet 'Asocia la conexión de trabajo
      rstBodegas.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
      rstBodegas.CursorLocation = adUseClient ' Cursor local al cliente
      rstBodegas.LockType = adLockOptimistic
        
      GSSQL = "dbo.invGetBodegaByUsuario " & gsUSUARIO
           
      If rstBodegas.State = adStateOpen Then rstBodegas.Close
      Set rstBodegas = GetRecordset(GSSQL)
End Sub

Private Sub PreparaRstLineas()
      Set rstLineas = New ADODB.Recordset
      If rstLineas.State = adStateOpen Then rstLineas.Close
      rstLineas.ActiveConnection = gConet 'Asocia la conexión de trabajo
      rstLineas.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
      rstLineas.CursorLocation = adUseClient ' Cursor local al cliente
      rstLineas.LockType = adLockOptimistic
        
      GSSQL = "dbo.invGetLineas '*' "
           
      If rstLineas.State = adStateOpen Then rstLineas.Close
      Set rstLineas = GetRecordset(GSSQL)
End Sub


Private Sub PreparaRstFamilias()
      Set rstFamilias = New ADODB.Recordset
      If rstFamilias.State = adStateOpen Then rstFamilias.Close
      rstFamilias.ActiveConnection = gConet 'Asocia la conexión de trabajo
      rstFamilias.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
      rstFamilias.CursorLocation = adUseClient ' Cursor local al cliente
      rstFamilias.LockType = adLockOptimistic
        
      GSSQL = "dbo.invGetFamilias '*' "
           
      If rstFamilias.State = adStateOpen Then rstFamilias.Close
      Set rstFamilias = GetRecordset(GSSQL)
End Sub



Private Sub PreparaRstSubFamilias()
      Set rstSubFamilias = New ADODB.Recordset
      If rstSubFamilias.State = adStateOpen Then rstSubFamilias.Close
      rstSubFamilias.ActiveConnection = gConet 'Asocia la conexión de trabajo
      rstSubFamilias.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
      rstSubFamilias.CursorLocation = adUseClient ' Cursor local al cliente
      rstSubFamilias.LockType = adLockOptimistic
        
      GSSQL = "dbo.invGetSubFamilias '*' "
           
      If rstSubFamilias.State = adStateOpen Then rstSubFamilias.Close
      Set rstSubFamilias = GetRecordset(GSSQL)
End Sub

Private Sub LlenarListaBodegas()
    If Not (rstBodegas.EOF And rstBodegas.BOF) Then
        Do While Not rstBodegas.EOF
            Me.lstBodegas.AddItem CStr(rstBodegas!IdBodega) + "-" + rstBodegas!DescrBodega
            rstBodegas.MoveNext
        Loop
    End If
End Sub


Private Sub LlenarListaClasificaciones(ByRef lstv As ListBox, rst As ADODB.Recordset)
    If Not (rst.EOF And rst.BOF) Then
        rst.MoveFirst
        Do While Not rst.EOF
            lstv.AddItem CStr(rst!Codigo) + "-" + rst!descr
            rst.MoveNext
        Loop
    End If
End Sub

Public Sub GetDatosPestanaClasificaciones()
    Dim i As Integer
    'Lineas
    For i = 0 To Me.lstLineas.ListCount - 1
        If Me.lstLineas.Selected(i) Then
            Dim sElemento() As String
            sElemento = Split(Me.lstLineas.List(i), "-")
            sLineasSeleccionadas = sLineasSeleccionadas + sElemento(0) + ","
        End If
    Next i
    'Familias
    For i = 0 To Me.lstFamilias.ListCount - 1
        If Me.lstFamilias.Selected(i) Then
            Dim sElemento() As String
            sElemento = Split(Me.lstFamilias.List(i), "-")
            sFamiliasSeleccionadas = sFamiliasSeleccionadas + sElemento(0) + ","
        End If
    Next i
    'Familias
    For i = 0 To Me.lstSubFamilias.ListCount - 1
        If Me.lstSubFamilias.Selected(i) Then
            Dim sElemento() As String
            sElemento = Split(Me.lstSubFamilias.List(i), "-")
            sSubFamiliasSeleccionadas = sSubFamiliasSeleccionadas + sElemento(0) + ","
        End If
    Next i
    
    sLineasSeleccionadas = Mid$(sLineasSeleccionadas, 1, Len(sLineasSeleccionadas) - 1)
    sFamiliasSeleccionadas = Mid$(sFamiliasSeleccionadas, 1, Len(sFamiliasSeleccionadas) - 1)
    sSubFamiliasSeleccionadas = Mid$(sSubFamiliasSeleccionadas, 1, Len(sSubFamiliasSeleccionadas) - 1)
    
End Sub

Public Sub GetDatosPestanaBodega()
    Dim i As Integer
        
    For i = 0 To Me.lstBodegas.ListCount - 1
        If Me.lstBodegas.Selected(i) Then
            Dim sElemento() As String
            sElemento = Split(Me.lstBodegas.List(i), "-")
            sBodegasSeleccionadas = sBodegasSeleccionadas + sElemento(0) + ","
        End If
    Next i
    
    sBodegasSeleccionadas = Mid$(sBodegasSeleccionadas, 1, Len(sBodegasSeleccionadas) - 1)
    
    bSoloProductosConExistencia = IIf(Me.chkSoloProductosConExistencia.Value = vbChecked, True, False)
    bIncluirExistenciaTransito = IIf(Me.chkIncluirExistenciaTransito.Value = vbChecked, True, False)
    
    If (Me.optConsolidaXArticulo.Value = True) Then
        Modovisualizacion = ConsolidadobyProducto
    ElseIf (Me.optConsolidaXBodega.Value = True) Then
        Modovisualizacion = ConsolidadobyBodega
    ElseIf (Me.optDetallaXArticulo.Value = True) Then
        Modovisualizacion = DetalladaByProducto
    ElseIf (Me.optDetallaXBodega.Value = True) Then
        Modovisualizacion = DetalladabyBodega
    End If
    
    
End Sub

