VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      Height          =   705
      Left            =   690
      Picture         =   "frmTest.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As New ADODB.Recordset

Private Sub Command_Click()
 Set rst = New ADODB.Recordset
    If rst.State = adStateOpen Then grst.Close
    rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rst.CursorLocation = adUseClient ' Cursor local al cliente
    rst.LockType = adLockOptimistic
    
    Dim frmAuto As New frmAutoSugiereLotes
    frmAuto.gsIDBodega = 1
    frmAuto.gsIDProducto = 1
    frmAuto.gdCantidad = 100
    frmAuto.gsDescrProducto = "Producto2"
    frmAuto.gsDescrBodega = "Descr Bodega"
    frmAuto.gsFormCaption = "Titulo Formulario"
    frmAuto.gsTitle = "titulo"
    
    Set frmAuto.grst = rst
    frmAuto.Show vbModal
    'MsgBox frmAuto.grst.RecordCount
    Set rst = frmAuto.grst
   MsgBox rst.BOF And rst.EOF
End Sub
