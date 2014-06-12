VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2535
      ItemData        =   "frmTest.frx":0000
      Left            =   1950
      List            =   "frmTest.frx":000D
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstsource As New ADODB.Recordset
Dim rstSeleccionaos As New ADODB.Recordset




Private Sub Command_Click()
 
End Sub

Private Sub PreparaRstSource()
      ' preparacion del recordset fuente del grid de compra
      ' recordar que este recordset va a ser temporal, no se hara addnew a la bd
      ' lleva además de los campos de la tabla detalle de compra, la descripcion del producto
      Set rstsource = New ADODB.Recordset
      If rstsource.State = adStateOpen Then rstsource.Close
      rstsource.ActiveConnection = gConet 'Asocia la conexión de trabajo
      rstsource.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
      rstsource.CursorLocation = adUseClient ' Cursor local al cliente
      rstsource.LockType = adLockOptimistic
        
      GSSQL = "dbo.invGetBodegas -1"
           
      If rstsource.State = adStateOpen Then rstsource.Close
      Set rstsource = GetRecordset(GSSQL)
End Sub

Private Sub PreparaRstSeleccionados()
      ' preparacion del recordset fuente del grid de compra
      ' recordar que este recordset va a ser temporal, no se hara addnew a la bd
      ' lleva además de los campos de la tabla detalle de compra, la descripcion del producto
      Set rstSeleccionados = New ADODB.Recordset
      If rstSeleccionados.State = adStateOpen Then rstSeleccionados.Close
      rstSeleccionados.ActiveConnection = gConet 'Asocia la conexión de trabajo
      rstSeleccionados.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
      rstSeleccionados.CursorLocation = adUseClient ' Cursor local al cliente
      rstSeleccionados.LockType = adLockOptimistic
        
      GSSQL = "SELECT * FROM dbo.invBODEGA WHERE 1=2"
           
      If rstSeleccionados.State = adStateOpen Then rstSeleccionados.Close
      Set rstSeleccionados = GetRecordset(GSSQL)
End Sub


Private Sub Form_Load()
    PreparaRstSource
    PreparaRstSeleccionados
    
    'Me.uControl.rstSeleccionados = rstSeleccionados
End Sub

