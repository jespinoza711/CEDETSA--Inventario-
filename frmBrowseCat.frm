VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmBrowseCat 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Catálogo"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7920
   DrawMode        =   4  'Mask Not Pen
   Icon            =   "frmBrowseCat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Height          =   585
      Left            =   7020
      Picture         =   "frmBrowseCat.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   60
      Width           =   645
   End
   Begin VB.CommandButton cmdAdicionar 
      Enabled         =   0   'False
      Height          =   585
      Left            =   7020
      Picture         =   "frmBrowseCat.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   840
      Width           =   645
   End
   Begin VB.CommandButton cmdQuitafiltro 
      Height          =   585
      Left            =   4695
      Picture         =   "frmBrowseCat.frx":189E
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Quita el filtro existente y muestra todo el catálogo."
      Top             =   840
      Width           =   645
   End
   Begin VB.CommandButton cmdFiltra 
      Height          =   585
      Left            =   3990
      Picture         =   "frmBrowseCat.frx":2568
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Filtra el catálogo con el criterio digitado por Ud."
      Top             =   840
      Width           =   645
   End
   Begin VB.Frame frmIntrod 
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6255
      Begin VB.Label lblIntrod 
         Alignment       =   2  'Center
         Caption         =   $"frmBrowseCat.frx":3232
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002F2F2F&
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   6015
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   7575
      Begin TrueOleDBGrid60.TDBGrid TDBG 
         Height          =   4215
         Left            =   120
         OleObjectBlob   =   "frmBrowseCat.frx":32D6
         TabIndex        =   0
         Top             =   240
         Width           =   7455
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdBuscar 
      Height          =   585
      Left            =   6330
      Picture         =   "frmBrowseCat.frx":619B
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Realiza una búsqueda, con F3 sigue buscando el criterio"
      Top             =   840
      Width           =   645
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6765
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9701
            MinWidth        =   9701
            Text            =   "Doble Click o Enter sobre el elemento deseado en el detalle..."
            TextSave        =   "Doble Click o Enter sobre el elemento deseado en el detalle..."
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
Public gsOrderfld As String
Public gsCaptionfrm As String
Public gbTypeCodeStr As Boolean ' indica si el codigo a buscar es tipo numerico o char
Public gbFiltra As Boolean ' indica si se le puede pasar un filtro al catalogo
Public gsFiltro As String ' para digitar el filtro del catálogo
Public gsExtraValor1 As String
Public gdWidthExtra1 As Double
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
Select Case UCase(gsNombrePantallaExtra)
    Case "FRMUSUARIO":
            Dim frm2 As frmUsuario
            Set frm2 = New frmUsuario
            frm2.Show vbModal
            Set frm2 = Nothing
End Select



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
    If BrowseCatalogo(gsTablabrw, gsCodigobrw, gsDescrbrw, gsMuestraExtra, gsFieldExtrabrw, gbFiltra, gsFiltro, gsOrderfld) Then
      TDBG.Caption = "Selección"
      TDBG.Columns(0).DataField = gsCodigobrw
      TDBG.Columns(0).Caption = gsCodigobrw
      TDBG.Columns(1).Caption = gsDescrbrw
      TDBG.Columns(1).DataField = gsDescrbrw
      
      If UCase(gsMuestraExtra) = "SI" Then
      If gdWidthExtra1 = 0 Then
        gdWidthExtra1 = 1500
      End If
      
        TDBG.Columns(2).Width = gdWidthExtra1
        TDBG.Columns(2).Caption = gsFieldExtrabrw
        TDBG.Columns(2).DataField = gsFieldExtrabrw
      Else
        TDBG.Columns(1).Width = TDBG.Columns(1).Width + 1000
        TDBG.Columns(2).Visible = False
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
  End If
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

Public Function BrowseCatalogo(sTabla As String, sFldCod As String, sDescr As String, gsMuestraExtra As String, gsFieldExtrabrw As String, bFiltra As Boolean, sFiltro As String, Optional sOrderFld As String) As Boolean
Dim lbok As Boolean
Dim sOrden As String
On Error GoTo error
  lbok = True
  If UCase(gsMuestraExtra) = "SI" Then
    GSSQL = "SELECT " & sFldCod & "," & sDescr & "," & gsFieldExtrabrw
  Else
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





