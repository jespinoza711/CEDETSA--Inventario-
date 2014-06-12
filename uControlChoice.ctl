VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl uControlChoice 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7800
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   4665
   ScaleWidth      =   7800
   Begin VB.CheckBox chkSeleccionarTodo 
      Caption         =   "Seleccionar Todo"
      Height          =   225
      Left            =   330
      TabIndex        =   4
      Top             =   4320
      Width           =   1875
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "<<"
      Height          =   360
      Left            =   3600
      TabIndex        =   1
      Top             =   2310
      Width           =   660
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   ">>"
      Height          =   360
      Left            =   3600
      TabIndex        =   0
      Top             =   1800
      Width           =   660
   End
   Begin MSComctlLib.ListView lstvSource 
      Height          =   3675
      Left            =   270
      TabIndex        =   2
      Top             =   570
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   6482
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lstSelected 
      Height          =   3675
      Left            =   4350
      TabIndex        =   3
      Top             =   570
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   6482
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblTituloSeleccionado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccionado"
      Height          =   195
      Left            =   4380
      TabIndex        =   6
      Top             =   360
      Width           =   3195
   End
   Begin VB.Label lblTituloOrigen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   330
      TabIndex        =   5
      Top             =   360
      Width           =   3165
   End
End
Attribute VB_Name = "uControlChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim gsTituloSource As String
Dim gsTituloSeleccionado As String

Public m_rstsource As New ADODB.Recordset
Public m_rstSeleccionados As New ADODB.Recordset
Public bSelectAll As Boolean


Public Property Get TituloSource() As String
    TituloSource = gsTituloSource
End Property
Public Property Let TituloSource(Value As String)
    gsTituloSource = Value
    UserControl.lblTituloOrigen.Caption = gsTituloSource
    PropertyChanged ("TituloSource")
End Property


Public Property Get TituloSeleccionado() As String
    TituloSeleccionado = gsTituloSeleccionado
End Property
Public Property Let TituloSeleccionado(Value As String)
    gsTituloSeleccionado = Value
    UserControl.lblTituloSeleccionado.Caption = gsTituloSeleccionado
    PropertyChanged ("TituloSeleccionado")
End Property

Public Property Get rstsource() As ADODB.Recordset
    rstsource = m_rstsource
End Property
Public Property Let rstsource(Value As ADODB.Recordset)
    m_rstsource = Value
    PropertyChanged ("rstsource")
End Property


Private Sub EnlazarControles()
   EnlazaListView lstvSource, rstsource
   EnlazaListView lstSelected, rstSeleccionados
End Sub

Private Sub cmdAdd_Click()
    If Not (rstsource.EOF And rstsource.BOF) Then
        If (lstvSource.SelectedItem <> Null) Then
            rstSeleccionados.AddNew
            rstSeleccionados!Codigo = lstvSource.SelectedItem.Text
            rstSeleccionados!Descripcion = lstvSource.SelectedItem.SubItems(1)
            rstSeleccionados.Save
            rstsource.Find ("Codigo=" & lstvSource.SelectedItem.Text)
            rstsource.Delete
            EnlazarControles
        End If
    End If
End Sub

Private Sub cmdDel_Click()
    If Not (rstSeleccionados.EOF And rstSeleccionados.BOF) Then
        If (lstSelected.SelectedItem <> Null) Then
            rstsource.AddNew
            rstsource!Codigo = lstSelected.SelectedItem.Text
            rstsource!Descripcion = lstSelected.SelectedItem.SubItems(1)
            rstsource.Save
            rstSeleccionados.Find ("Codigo=" & lstSelected.SelectedItem.Text)
            rstSeleccionados.Delete
            EnlazarControles
        End If
    End If
End Sub

Private Sub UserControl_GetDataMember(DataMember As String, Data As Object)

End Sub

Private Sub UserControl_Initialize()
   ' Inicializar
End Sub

Public Sub Inicializar()
    lblTituloOrigen.Caption = gsTituloSource
    lblTituloSeleccionado.Caption = gsTituloSeleccionado
    If (rstsource.State = adStateOpen) Then
        If Not (rstsource.EOF And rstsource.BOF) Then
            EnlazaListView lstvSource, rstsource
            EnlazaListView lstSelected, rstSeleccionados
        End If
    End If
End Sub


Private Sub EnlazaListView(lstv As ListView, rst As ADODB.Recordset)
    Dim sItem As String
        With lstv
            ' Las pruebas serán en modo "detalle"
            .View = lvwReport
            ' al seleccionar un elemento, seleccionar la línea completa
            .FullRowSelect = True
            ' Mostrar las líneas de la cuadrícula
            .GridLines = True
            ' No permitir la edición automática del texto
            .LabelEdit = lvwManual
            ' Permitir múltiple selección
            .MultiSelect = False
            ' Para que al perder el foco,
            ' se siga viendo el que está seleccionado
            .HideSelection = False
            .LabelWrap = False
            .ForeColor = vbBlue
    
        End With
    
        With lstv.ColumnHeaders.Add(, , "Codigo", 2000)
                '.Tag = cCodigo
        End With
    
        With lstv.ColumnHeaders.Add(, , "Descripción", 4500)
                '.Tag = cTexto
        End With
    
        lstv.ListItems.Clear
        ' Asignar algunos datos aleatorios
        If Not (rst.EOF And rst.BOF) Then
            rst.MoveFirst
            While Not rst.EOF
            
                Dim item As ListItem
                item = lstv.ListItems.Add(, , rst("IDCodigo").Value)
                item.ListSubItems.Add , , rst("Descripcion").Value
                rst.MoveNext
            Wend
        End If

End Sub

Private Sub UserControl_InitProperties()
    Inicializar
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    gsTituloSource = PropBag.ReadProperty("TituloSource", "Establezca el titulo")
    gsTituloSeleccionado = PropBag.ReadProperty("TituloSeleccionado", "Estableza el titulo")
End Sub

Private Sub UserControl_Show()
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "TituloSource", gsTituloSource, "Estableza el titulo"
    PropBag.WriteProperty "TituloSeleccionado", gsTituloSeleccionado, "Estableza el titulo"
End Sub
