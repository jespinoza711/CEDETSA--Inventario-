VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmBrowseCat 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Catálogo"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   8145
   DrawMode        =   4  'Mask Not Pen
   Icon            =   "frmBrowseCat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H80000009&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2580
      Picture         =   "frmBrowseCat.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6600
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H80000009&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4110
      Picture         =   "frmBrowseCat.frx":0C0E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6600
      Width           =   1155
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   8145
      TabIndex        =   6
      Top             =   0
      Width           =   8145
      Begin VB.Image Image 
         Height          =   480
         Index           =   2
         Left            =   210
         Picture         =   "frmBrowseCat.frx":0F52
         Top             =   90
         Width           =   480
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmBrowseCat.frx":181C
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
         Height          =   315
         Index           =   1
         Left            =   930
         TabIndex        =   8
         Top             =   390
         Width           =   7050
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
         TabIndex        =   7
         Top             =   90
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   585
      Left            =   7230
      Picture         =   "frmBrowseCat.frx":18B6
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   645
   End
   Begin VB.CommandButton cmdAdicionar 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   585
      Left            =   6240
      Picture         =   "frmBrowseCat.frx":1BC0
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Agrega un nuevo elemento de Catalogo"
      Top             =   960
      Width           =   645
   End
   Begin VB.CommandButton cmdQuitafiltro 
      BackColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   4350
      Picture         =   "frmBrowseCat.frx":288A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Quita el filtro existente y muestra todo el catálogo."
      Top             =   990
      Width           =   645
   End
   Begin VB.CommandButton cmdFiltra 
      BackColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   3660
      Picture         =   "frmBrowseCat.frx":2BA9
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Filtra el catálogo con el criterio digitado por Ud."
      Top             =   960
      Width           =   645
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   7575
      Begin TrueOleDBGrid60.TDBGrid TDBG 
         Height          =   4215
         Left            =   90
         OleObjectBlob   =   "frmBrowseCat.frx":2EBC
         TabIndex        =   13
         Top             =   210
         Width           =   7395
      End
   End
   Begin VB.CommandButton cmdBuscar 
      BackColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   5550
      Picture         =   "frmBrowseCat.frx":63E1
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Realiza una búsqueda, con F3 sigue buscando el criterio"
      Top             =   960
      Width           =   645
   End
   Begin Inventario.CtlLiner CtlLiner 
      Height          =   30
      Left            =   0
      TabIndex        =   9
      Top             =   750
      Width           =   17925
      _ExtentX        =   31618
      _ExtentY        =   53
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Doble Click o Enter sobre el elemento deseado en el detalle..."
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   630
      TabIndex        =   12
      Top             =   6180
      Width           =   4350
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   150
      Picture         =   "frmBrowseCat.frx":70AB
      Top             =   6060
      Width           =   480
   End
End
Attribute VB_Name = "frmBrowseCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public gsCodigobrw As String
Public gsDescrbrw As String
Public gsTablabrw As String
Public gsFieldsbrw As String
Public gsMuestraExtra As String
Public gsFieldExtrabrw As String

Public gsMuestraExtra2 As String
Public gsFieldExtrabrw2 As String

Public gsOrderfld As String
Public gsCaptionfrm As String
Public gbTypeCodeStr As Boolean ' indica si el codigo a buscar es tipo numerico o char
Public gbFiltra As Boolean ' indica si se le puede pasar un filtro al catalogo
Public gsFiltro As String ' para digitar el filtro del catálogo
Public gsExtraValor1 As String
Public gsExtraValor2 As String
Public gdWidthExtra1 As Double

Public gdWidthExtra2 As Double
' Estas variables son utilizadas para el llamado de formularios extras por ejemplo en Add del Catalogo
Public gsNombrePantallaExtra As String
Public gsCodigo1FormularioExtra As String
Public gsCodigo2FormularioExtra As String
Public gsCodigo3FormularioExtra As String
Public gsDescr1FormularioExtra As String
Public gsDescr2FormularioExtra As String
Public gsDescr3FormularioExtra As String

Dim rst As ADODB.Recordset
Dim sFindF3 As String ' string con el string seach del F3
Dim BookFirstF3 As Variant ' bookmark de la primera busqueda F3
Dim sfldFind As String
Dim gsCampoCodigoTabla As String
Dim gsCampoDescrTabla As String
Dim bFiltroAntes As Boolean ' indica que ya filtro la informacion antes de adicionar
Dim bOrdenCodigo As Boolean
Dim bOrdenDescr As Boolean

Private Sub cmdAceptar_Click()
  Hide
End Sub

Private Sub cmdAdicionar_Click()
'Select Case UCase(gsNombrePantallaExtra)
'    Case "FRMUSUARIO":
'            Dim frm2 As frmUsuario
'            Set frm2 = New frmUsuario
'            frm2.Show vbModal
'            Set frm2 = Nothing
'End Select



End Sub

Private Sub cmdBuscar_Click()
Busqueda sfldFind
TDBG.SetFocus
End Sub

Private Sub cmdCancelar_Click()
  gsCodigobrw = ""
  gsDescrbrw = ""
  Hide
End Sub

Private Sub cmdFiltra_Click()
Dim sTipoCodigo As String
If gbTypeCodeStr = True Then
    sTipoCodigo = "C"
Else
    sTipoCodigo = "N"
End If
Filtra gsCampoCodigoTabla, sTipoCodigo, gsCampoDescrTabla
bFiltroAntes = True
End Sub

Private Sub cmdQuitafiltro_Click()
QuitaFiltro
bFiltroAntes = False
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdSalir_Click()
cmdCancelar_Click
End Sub

Private Sub Form_Load()
'Set gRegistrosBrw = New ADODB.Recordset
bFiltroAntes = False

Set gRegistrosBrw = New ADODB.Recordset  'Inicializa la variable de los registros
    Caption = gsCaptionfrm 'Caption & " " & gsTablabrw & "S" 'Mid(gsCaptionfrm, 1, Len(gsCaptionfrm) - 1)
    gRegistrosBrw.ActiveConnection = gConet 'Asocia la conexión de trabajo
    gRegistrosBrw.CursorType = adOpenKeyset  'Asigna un cursor estático
    gRegistrosBrw.CursorLocation = adUseClient ' Cursor local al cliente
    gRegistrosBrw.LockType = adLockOptimistic
    sfldFind = gsFieldsbrw
    If BrowseCatalogo(gsTablabrw, gsCodigobrw, gsDescrbrw, gsMuestraExtra, gsFieldExtrabrw, gsMuestraExtra2, gsFieldExtrabrw2, gbFiltra, gsFiltro, gsOrderfld) Then
      TDBG.Caption = "Selección"
      TDBG.Columns(0).DataField = gsCodigobrw
      TDBG.Columns(0).Caption = gsCodigobrw
      TDBG.Columns(1).Caption = gsDescrbrw
      TDBG.Columns(1).DataField = gsDescrbrw
      
      If UCase(gsMuestraExtra) = "SI" Then
        If gdWidthExtra1 = 0 Then
          gdWidthExtra1 = 1500
        End If
         TDBG.Columns(1).Width = 2500
        TDBG.Columns(2).Width = gdWidthExtra1
        TDBG.Columns(2).Caption = gsFieldExtrabrw
        TDBG.Columns(2).DataField = gsFieldExtrabrw
      Else
        TDBG.Columns(2).Visible = False
      End If
      
      If UCase(gsMuestraExtra2) = "SI" Then
        If gdWidthExtra1 = 0 Then
          gdWidthExtra1 = 1500
        End If
        TDBG.Columns(3).Width = gdWidthExtra1
        TDBG.Columns(3).Caption = gsFieldExtrabrw2
        TDBG.Columns(3).DataField = gsFieldExtrabrw2
      Else
        TDBG.Columns(3).Visible = False
      End If
      If UCase(gsMuestraExtra) <> "SI" And UCase(gsMuestraExtra2) <> "SI" Then
         TDBG.Columns(1).Width = 5000
        'TDBG.Columns(1).Width = TDBG.Columns(1).Width + 1000
        TDBG.Columns(2).Visible = False
        TDBG.Columns(3).Visible = False
      End If
      
      If UCase(gsMuestraExtra) = "SI" And UCase(gsMuestraExtra2) <> "SI" Then
      TDBG.Columns(1).Width = 3500
      End If
      gsCampoCodigoTabla = gsCodigobrw
      gsCampoDescrTabla = gsDescrbrw
      
    
      
      Set TDBG.DataSource = gRegistrosBrw
      TDBG.ReBind
      gsCodigobrw = ""
      gsDescrbrw = ""
    End If
End Sub

Private Sub Filtra(sFldnameCode As String, sTypeCodigo As String, sfldNameDescr As String)
    Dim frmFind As New frmBusqGral
    Dim strCodigo As String, strDescr As String
    Dim sFiltro As String
      frmFind.Caption = "Filtro de Información"
      frmFind.Show vbModal
      If frmFind.txtCodigo <> "" Then
        strCodigo = frmFind.txtCodigo.Text
      End If
    
      If frmFind.txtDescr <> "" Then
        strDescr = frmFind.txtDescr.Text
      End If
    If gRegistrosBrw.State = adStateClosed Then
        Exit Sub
    End If
      If strCodigo <> "" Or strDescr <> "" Then
         sFiltro = ""
         If strCodigo <> "" Then
            If UCase(sTypeCodigo) = "N" Then
                sFiltro = gsCampoCodigoTabla & " =" & strCodigo
            Else
                strCodigo = strCodigo '& "%" ' SustituyeChar(strCodigo, "*", "%")
                sFiltro = gsCampoCodigoTabla & " like '%" & strCodigo & "%'"
            End If
         End If
    
         If strDescr <> "" Then
                If sFiltro <> "" Then sFiltro = sFiltro & " and "
                strDescr = SustituyeChar(strDescr, "*", "%")
                sFiltro = sfldNameDescr & " like '" & strDescr & "'"
         End If
           gRegistrosBrw.Filter = adFilterNone
           gRegistrosBrw.Filter = sFiltro
           If gRegistrosBrw.EOF Then
             MsgBox "No hay registros que cumplan con ese criterio", vbOKOnly, "Error en filtro"
             QuitaFiltro
           End If
    
      End If
    
    
      Unload frmFind

End Sub

Private Sub Quitafiltro_Click()
    QuitaFiltro
End Sub

Private Sub QuitaFiltro()
    If gRegistrosBrw.State = adStateClosed Then
        Exit Sub
    End If
    gRegistrosBrw.Filter = adFilterNone
    gRegistrosBrw.Filter = adFilterNone
    If Not gRegistrosBrw.EOF And Not gRegistrosBrw.BOF Then
        gRegistrosBrw.MoveFirst
    End If
End Sub





Private Sub Form_Unload(Cancel As Integer)
If (gRegistrosBrw.State = adStateClosed) Then Exit Sub
If (gRegistrosBrw.BOF And gRegistrosBrw.EOF) Then
  gsCodigobrw = ""
  gsDescrbrw = ""
End If
End Sub

Private Sub TDBG_DblClick()
Hide
End Sub

Private Sub TDBG_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF3 Then
  If Not gRegistrosBrw.EOF And sFindF3 <> "" Then
    gRegistrosBrw.MoveNext
    If gRegistrosBrw.EOF Then
        gRegistrosBrw.MovePrevious
    End If
    If Not gRegistrosBrw.EOF Then
      gRegistrosBrw.Find sFindF3
      If gRegistrosBrw.EOF And BookFirstF3 <> vbEmpty Then
          gRegistrosBrw.Bookmark = BookFirstF3
'      Else
'          gRecordset.MoveFirst
      End If
    End If
  End If
End If
End Sub

Private Sub TDBG_HeadClick(ByVal ColIndex As Integer)
Dim iTmp As Integer
If gRegistrosBrw.EOF And gRegistrosBrw.BOF Then
  Exit Sub
End If

  iTmp = TDBG.Splits(0).LeftCol
  If ColIndex = 0 And TDBG.Columns(ColIndex).Name <> "" Then
    If bOrdenCodigo = False Then
      gRegistrosBrw.Sort = gRegistrosBrw(ColIndex).Name & " ASC" 'TDBG.Columns(ColIndex).Name & " ASC"
      bOrdenCodigo = True
    Else
      gRegistrosBrw.Sort = gRegistrosBrw(ColIndex).Name & " DESC" 'TDBG.Columns(ColIndex).Name & " DESC"
      bOrdenCodigo = False
    End If
  Else
    If TDBG.Columns(ColIndex).Name <> "" Then
        If bOrdenDescr = False Then
          gRegistrosBrw.Sort = gRegistrosBrw(ColIndex).Name & " ASC"  ' TDBG.Columns(ColIndex).Name & " ASC"
          bOrdenDescr = True
        Else
          gRegistrosBrw.Sort = gRegistrosBrw(ColIndex).Name & " DESC" 'TDBG.Columns(ColIndex).Name & " DESC"
          bOrdenDescr = False
        End If
    End If
  End If
  
  gRegistrosBrw.MoveFirst
'  gRegistrosBrw.Sort = TDBG.Columns(ColIndex).Name
  TDBG.Splits(0).LeftCol = iTmp
End Sub

Private Sub TDBG_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
  cmdAceptar_Click
End If
End Sub



Private Sub TDBG_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If gRegistrosBrw.State = adStateClosed Then
gsCodigobrw = ""
gsDescrbrw = ""
  Exit Sub
End If
If Not gRegistrosBrw.EOF Then
  gsCodigobrw = gRegistrosBrw.Fields(0).value
  gsDescrbrw = gRegistrosBrw.Fields(1).value
  If gsMuestraExtra = "SI" Then
    gsExtraValor1 = gRegistrosBrw.Fields(2).value
  Else
    gsExtraValor1 = ""
  End If
  If gsMuestraExtra2 = "SI" Then
    gsExtraValor2 = gRegistrosBrw.Fields(3).value
'  Else
'   gsExtraValor2 = ""
  End If
Else
  gsCodigobrw = ""
  gsDescrbrw = ""
End If
End Sub

' Levanta la pantalla de búsqueda
Private Sub Busqueda(sFldname As String)
Dim frmFind As New frmBusqGral
Dim strCodigo As String, strDescr As String
Dim rstClone As ADODB.Recordset
Set rstClone = New ADODB.Recordset
sFindF3 = ""
  frmFind.Show vbModal
  If frmFind.txtCodigo <> "" Then
    strCodigo = frmFind.txtCodigo.Text
  End If
If gRegistrosBrw.State = adStateClosed Then
    Exit Sub
End If
  If frmFind.txtDescr <> "" Then
    strDescr = frmFind.txtDescr.Text
  End If

  If strCodigo <> "" Or strDescr <> "" Then
       rstClone.Filter = adFilterNone
      Set rstClone = gRegistrosBrw.Clone
      
      If strCodigo <> "" Then
            If gbTypeCodeStr = False Then
              sFindF3 = gsCampoCodigoTabla & "=" & strCodigo
              rstClone.Filter = gsCampoCodigoTabla & "=" & strCodigo
            End If
      End If
     If strCodigo <> "" And gbTypeCodeStr = True Then
              sFindF3 = gsCampoCodigoTabla & " like '" & strCodigo & "%'"
              rstClone.Filter = sFindF3 ' gsCampoCodigoTabla & "=" & strCodigo
            
      End If
      
      If strDescr <> "" Then
        'strDescr = SustituyeChar(strDescr, "*", "%")
        'strDescr = strDescr & "*"
        sFindF3 = gsCampoDescrTabla & " like '%" & strDescr & "%'"
        rstClone.Filter = sFindF3 ' gsCampoDescrTabla & " like '" & strDescr & "'"
      End If
      
      If Not rstClone.EOF Then
        gRegistrosBrw.Bookmark = rstClone.Bookmark
        BookFirstF3 = rstClone.Bookmark
        TDBG.Bookmark = gRegistrosBrw.Bookmark
      Else
        gRegistrosBrw.MoveFirst
        sFindF3 = ""
        BookFirstF3 = gRegistrosBrw.Bookmark
      End If
  End If
  
  If rstClone.State = adStateOpen Then
      rstClone.Close
  End If
  Unload frmFind

End Sub

Private Sub Busquedaant()
Dim frmFind As New frmBusqGral
Dim strCodigo As String
Dim strDescr As String
Dim sField As String
Dim rstClone As ADODB.Recordset
Set rstClone = New ADODB.Recordset

  frmFind.Show vbModal
  If frmFind.txtCodigo <> "" Then
    strCodigo = frmFind.txtCodigo.Text
  End If

  If frmFind.txtDescr <> "" Then
    strDescr = frmFind.txtDescr.Text
  End If
  
  If strCodigo <> "" Or strDescr <> "" Then
       rstClone.Filter = adFilterNone
      Set rstClone = gRegistrosBrw.Clone
      sField = rstClone.Fields(0).Name
      If strCodigo <> "" Then
        If Not gbTypeCodeStr Then
            rstClone.Filter = sField & "=" & strCodigo
        Else
            rstClone.Filter = sField & " ='" & strCodigo & "'"
        End If
      End If
      
      If strDescr <> "" Then
        rstClone.Filter = "descr like '%" & strDescr & "%'"
      End If
      
      If Not rstClone.EOF Then
        gRegistrosBrw.Bookmark = rstClone.Bookmark
        TDBG.Bookmark = gRegistrosBrw.Bookmark
      End If
  End If
  
  If rstClone.State = adStateOpen Then
      rstClone.Close
  End If
  Unload frmFind

End Sub

Public Function BrowseCatalogo(sTabla As String, sFldCod As String, sDescr As String, gsMuestraExtra As String, gsFieldExtrabrw As String, gsMuestraExtra2 As String, gsFieldExtrabrw2 As String, bFiltra As Boolean, sFiltro As String, Optional sOrderFld As String) As Boolean
Dim lbok As Boolean
Dim sOrden As String
On Error GoTo error
  lbok = True
  If UCase(gsMuestraExtra) = "SI" Then
    GSSQL = "SELECT " & sFldCod & "," & sDescr & "," & gsFieldExtrabrw
  End If
  
  If UCase(gsMuestraExtra2) = "SI" And UCase(gsMuestraExtra) = "SI" Then
    GSSQL = GSSQL & "," & gsFieldExtrabrw2
  End If
  
  
  If UCase(gsMuestraExtra) <> "SI" And UCase(gsMuestraExtra2) <> "SI" Then
    GSSQL = "SELECT " & sFldCod & "," & sDescr
  End If
  
  GSSQL = GSSQL & " FROM " & gsCompania & "." & sTabla          'Constuye la sentencia SQL
  If sOrderFld = "" Then
  sOrden = sFldCod
  Else
  sOrden = sOrderFld
  End If
  
  If bFiltra = True And sFiltro <> "" Then
    GSSQL = GSSQL & " WHERE " & sFiltro
  End If
  GSSQL = GSSQL & " ORDER BY " & sOrden
  If gRegistrosBrw.State = adStateOpen Then gRegistrosBrw.Close
  gRegistrosBrw.Open GSSQL, gConet, adOpenDynamic, adLockOptimistic, adCmdText    'Ejecuta la sentencia

If (gRegistrosBrw.BOF And gRegistrosBrw.EOF) Then  'Si no es válido
    gsOperacionError = "No existe ese item." 'Asigna msg de error
    lbok = False  'Indica que no es válido
End If

BrowseCatalogo = lbok
Exit Function
error:
  lbok = False
  BrowseCatalogo = lbok

End Function






