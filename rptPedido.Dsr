VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} rptPedido 
   Caption         =   "Pedido"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   18960
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   33443
   _ExtentY        =   19315
   SectionData     =   "rptPedido.dsx":0000
End
Attribute VB_Name = "rptPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PageHeader_Format()
fldFecha.Text = Now
lblCliente.Caption = Str(DataControl1.Recordset!IDCliente) + " " + (DataControl1.Recordset!NOMBRE)
lblNoPedido.Caption = Str(DataControl1.Recordset!IDPedido)
lblVendedor.Caption = Str(DataControl1.Recordset!IDVendedor) + " " + (DataControl1.Recordset!DescrVendedor)
lblFechaPedido.Caption = Format(Str(DataControl1.Recordset!Fecha), "dd-mm-YYYY")
End Sub
