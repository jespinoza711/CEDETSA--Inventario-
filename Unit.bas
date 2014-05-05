Attribute VB_Name = "Unit"
Public Const ICO_INFORMACION = 1
Public Const ICO_PREGUNTA = 2
Public Const ICO_ADVERTENCIA = 3
Public Const ICO_ERROR = 4
Public Const ICO_OK = 5

'Constantes para Modulo
Public Const C_MODULO = 1000
Public Const A_CATALOGOS = 1
Public Const A_PRODUCTOS = 2


Public gdFechaCortePreparacion As Date
Public gsCompaniaActiva As String
Public gsNombreCompania As String
Public gsBaseDatosActiva As String
Public gsPropiedadesPos As String
Public gsRutaCargaPosOffLine As String
Public gsCompania As String

'Public gskey As New EncryptedData
Public gbEquipo64Bits As Boolean


Dim TimerId As Long, Var As Long

Public gConet As ADODB.Connection
    'Objeto de conexión a la base de datos
Public gRegistros As ADODB.Recordset
    'Objeto que va a mantener el resultado de las consultas hechas
Public gRegistrosCmd As ADODB.Recordset
    'Objeto que va a mantener el resultado de la ejecucion de un comando de la conexion
Public gRegistrosBrw As ADODB.Recordset
    'Objeto que va a mantener el resultado para el browse de los catalogos
Public GSSQL As String
    'Contiene la sentencia SQL a ejecutar o que generó el error
Public gsOperacionError As String
Public gsConetstr As String


Public gsNombreServidor As String 'Contiene el nombre del servidor
Public gsNombreBaseDatos As String 'Contiene el nombre de la base de datos
Public gsUser As String 'Contiene el usuario
Public gsPassword As String 'Contiene el password del usuario
Public gsTypeDB As String 'Contiene el password del usuario


Public gsUSUARIO As String ' este es el usuario login

Public grRecordsetAcceso As ADODB.Recordset

Type TParametros
  NombreEmpresa As String
End Type

Public Enum TypAccion
    Add
    Edit
    View
End Enum

Type TparametrosPrepDifCamb
  FechaCorte As Date
End Type
    ' Tipo de parametros del Sistema de Inventario
Public gparametros As TParametros
Public gParametrosPrepDifCC As TparametrosPrepDifCamb

Public gClientePreparacion As String
Public gCategoriaPreparacion As String
Public gFechaCortePreparacion As Date
Private Declare Function GetProcAddress Lib "kernel32" _
    (ByVal hModule As Long, _
    ByVal lpProcName As String) As Long
    
Private Declare Function GetModuleHandle Lib "kernel32" _
    Alias "GetModuleHandleA" _
    (ByVal lpModuleName As String) As Long
    
Private Declare Function GetCurrentProcess Lib "kernel32" _
    () As Long

Private Declare Function IsWow64Process Lib "kernel32" _
    (ByVal hProc As Long, _
    bWow64Process As Boolean) As Long



' Declaraciones para uso del Win32 API
Private Declare Function GetPrivateProfileString Lib "kernel32" _
  Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" _
  Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
  

Public Sub WriteIniKey(FileName As String, Section As String, Key As String, value As String)
WritePrivateProfileString Section, Key, value, FileName
End Sub




'******************************************************
'******************************************************
'*************** Programa Principal *******************

Sub Main()


    
    Dim lbOk As Boolean
    
    gbEquipo64Bits = Is64bit()
    If Not LeerConfiguracion = True Then
      lbOk = Mensaje("Error al leer el archivo de configuración ", ICO_ERROR, False)
      End
    End If
    
    If ConexionBD Then
        If (Command$ <> "") Then
            LoginForCommandLine
        Else
        
            Dim frm As frmLogin
            Set frm = New frmLogin
            frm.Show vbModal
            lbOk = frm.LoginSucceeded
            If lbOk Then
              Unload frm
              frmMain.Show
              'frmPrepare.Show
              'frmAnalisisVencimiento.Show
            Else
                Unload frm
                End
            End If
        End If
    Else
      lbOk = Mensaje("No pudo conectarse a la base de datos :" & gsOperacionError, ICO_ERROR, False)
      End
    End If
    
'Dim lbOk As Boolean
'If Not LeerConfiguracion = True Then
'  lbOk = Mensaje("Error al leer el archivo de configuración ", ICO_ERROR, False)
'  End
'End If
'
'If ConexionBD Then
'frmFacturas.Show
'
''''lbOk = CargaParametros
'''''  frmabout.Show 'vbModal, MDIMenu
'''''Dim Inicio As Long
'''''Dim Final As Long
'''''Dim tiempopausa As Integer
'''''
'''''   tiempopausa = 3   ' Asigna hora de inicio.
'''''   Inicio = Timer   ' Establece la hora de inicio.
'''''   Do While Timer < Inicio + tiempopausa
'''''      DoEvents   ' Cambia a otros procesos.
'''''   Loop
'''''   Final = Time   ' Asigna hora de finalización.
'''''   Unload frmabout
''''MDIMenu.Show
'
'Else
'  lbOk = Mensaje("No pudo conectarse a la base de datos :" & gsOperacionError, ICO_ERROR, False)
'  End
'End If




'''Dim lbok As Boolean
'''
'''gbEquipo64Bits = Is64bit()
'''If Not LeerConfiguracion = True Then
'''  lbok = Mensaje("Error al leer el archivo de configuración ", ICO_ERROR, False)
'''  End
'''End If
'''
'''If ConexionBD Then
'''Dim frm As frmLogin
'''Set frm = New frmLogin
'''frm.Show vbModal
'''lbok = frm.LoginSucceeded
'''If lbok Then
'''  Unload frm
'''  frmMain.Show
'''  'frmPrepare.Show
'''  'frmAnalisisVencimiento.Show
'''Else
'''    Unload frm
'''    End
'''End If
'''Else
'''  lbok = Mensaje("No pudo conectarse a la base de datos :" & gsOperacionError, ICO_ERROR, False)
'''  End
'''End If
''''Dim lbOk As Boolean
''''If Not LeerConfiguracion = True Then
''''  lbOk = Mensaje("Error al leer el archivo de configuración ", ICO_ERROR, False)
''''  End
''''End If
''''
''''If ConexionBD Then
''''frmFacturas.Show
''''
'''''''lbOk = CargaParametros
''''''''  frmabout.Show 'vbModal, MDIMenu
''''''''Dim Inicio As Long
''''''''Dim Final As Long
''''''''Dim tiempopausa As Integer
''''''''
''''''''   tiempopausa = 3   ' Asigna hora de inicio.
''''''''   Inicio = Timer   ' Establece la hora de inicio.
''''''''   Do While Timer < Inicio + tiempopausa
''''''''      DoEvents   ' Cambia a otros procesos.
''''''''   Loop
''''''''   Final = Time   ' Asigna hora de finalización.
''''''''   Unload frmabout
'''''''MDIMenu.Show
''''
''''Else
''''  lbOk = Mensaje("No pudo conectarse a la base de datos :" & gsOperacionError, ICO_ERROR, False)
''''  End
''''End If

End Sub
'**************** Fin de Principal *********************
'*******************************************************
'*******************************************************

Private Sub LoginForCommandLine()

    Dim sArgs() As String
    Dim sUsuario As String
    Dim sPass As String
    Dim i As Integer
        
    sArgs = Split(Command$, " ")
    
    sUsuario = sArgs(0)
    sPass = sArgs(1)
    
   
    If UserCouldIN(sUsuario, sPass) Then
        gsUSUARIO = sUsuario
    
        lbOk = LoadAccess(sUsuario, sPass, C_MODULO)
        If Not lbOk Then
            lbOk = Mensaje("No se pudieron cargar los accesos del usuario ", ICO_ERROR, False)
            End
        Else
        'lbok = CargaParametros()
            If Not lbOk Then
                lbOk = Mensaje("No se ha configurado el Sistema... los parametros no se han definido ", ICO_ERROR, False)
                End
            End If
        End If
    Else
        lbOk = Mensaje("Login o Password incorrectos...", ICO_ERROR, False)
        End
    End If
    
    
    If lbOk Then
        frmMain.Show
    Else
        End
    End If

End Sub


Public Function Is64bit() As Boolean
    Dim handle As Long, bolFunc As Boolean

    ' Assume initially that this is not a Wow64 process
    bolFunc = False

    ' Now check to see if IsWow64Process function exists
    handle = GetProcAddress(GetModuleHandle("kernel32"), _
                   "IsWow64Process")

    If handle > 0 Then ' IsWow64Process function exists
        ' Now use the function to determine if
        ' we are running under Wow64
        IsWow64Process GetCurrentProcess(), bolFunc
    End If

    Is64bit = bolFunc

End Function


Public Function Update_Parametros(sNombreEmpresa As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
  GSSQL = gsCompania & "." & "uspparmaUpdateParametros '" & sNombreEmpresa & "'"
  
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = " ."
      lbOk = False
        
    End If

Update_Parametros = lbOk
Exit Function

error:
  lbOk = False
  Resume Next

End Function

Public Function Update_UltimoDiaCerrado(sUltAnioCerrado As String, sUltMesCerrado As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
  GSSQL = gsCompania & "." & "sgvUpdateUltimoMesCerrado " & sUltAnioCerrado & "," & sUltMesCerrado
  
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = " ."
      lbOk = False
        
    End If

Update_UltimoDiaCerrado = lbOk
Exit Function

error:
  lbOk = False
  Resume Next

End Function

Public Function Update_AnulaEsquela(sEsquela As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
  GSSQL = "Update " & gsCompania & "." & "sgvEsquela set Anulada=1 where Consecutivo='" & sEsquela & "'"
  
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = " ."
      lbOk = False
        
    End If

Update_AnulaEsquela = lbOk
Exit Function

error:
  lbOk = False
  Resume Next

End Function



Public Function FechaEnMesAbierto(dFecha As Date) As Boolean
Dim lbOk As Boolean
Dim dUltimoMesCerrado As Date
Dim dFirstDayAllow As Date
lbOk = False

dUltimoMesCerrado = UltimoMesCerradoFirstDay
dFirstDayAllow = DateAdd("m", 1, dUltimoMesCerrado)
If dFecha >= dFirstDayAllow Then
    lbOk = True
End If
FechaEnMesAbierto = lbOk
End Function



Private Function LimpiaNulos(ByVal psParam As String) As String   'Toma una hilera y elimina sus nulos
  Dim lnCont As Integer 'Es el contador del ciclo
  Dim lnTam As Integer  'Mantiene el tamaño de la hilera
  Dim lsHilera As String  'String de trabajo
  
  lnTam = CInt(Len(psParam)) 'Obtiene el tamaño de la hilera
  lnCont = InStr(1, psParam, Chr(0), 1) - 1
  If lnCont <> -1 Then
  lsHilera = Left(psParam, lnCont) & Space(lnTam - lnCont)  'Asigna espacios al resto
  Else
  lsHilera = psParam
  End If
  
  LimpiaNulos = lsHilera
End Function

Public Sub ValidaLargo( _
  ByVal psTexto As String, _
  ByRef KeyAscii As Integer, _
  ByVal pnLargo As Integer)
  Dim lbOk As Boolean
On Error GoTo err
  
  If (pnLargo <= Len(psTexto)) Then 'Si se alcanzó el límite
    If (KeyAscii <> 8) Then
      KeyAscii = 0
    End If
  End If
  Exit Sub

err:
    lbOk = Mensaje("Función: ValidaLargo ..." & err.Description, ICO_ERROR, False)
    Resume Next
End Sub

Public Function ReadIniKey(sFileName As String, sKey As String, sVariable As String) As String
Dim sResultado As String
Dim lsTemp As String * 30
Dim sFileNameLocal As String
On Error GoTo error
sResultado = ""
sFileNameLocal = ""
'sFileNameLocal = App.Path & "\" & sFileName
sFileNameLocal = sFileName
GetPrivateProfileString sKey, sVariable, "NO", lsTemp, 32, sFileNameLocal
sResultado = LimpiaNulos(lsTemp)
sResultado = RTrim(sResultado)
ReadIniKey = sResultado
Exit Function
error:
    ReadIniKey = ""
    Exit Function
End Function

Public Sub Mayuscula(KeyAscii As Integer)
  'Hace el upper case del caracter
  Dim lsChar As String 'Caracter
  
  lsChar = Chr(KeyAscii) 'Asigna el caracter
  lsChar = UCase(lsChar) 'Hace el upper case
  KeyAscii = Asc(lsChar) 'Asigna el código
End Sub

' Para emitir mensajes del sistema
Public Function Mensaje(sMsg As String, icono As Integer, cancelar As Boolean) As Boolean
Dim sTitle As String
    Select Case icono
      Case ICO_ADVERTENCIA:
        sTitle = "Advertencia"
      Case ICO_ERROR
        sTitle = "Error"
      Case ICO_PREGUNTA
        sTitle = "Pregunta"
      Case ICO_INFORMACION
        sTitle = "Información"
      Case ICO_OK
        sTitle = "Exito"
    End Select
    
    Set frmMensajeError = Nothing
    frmMensajeError.Caption = sTitle
    frmMensajeError.lsMensaje = sMsg
    frmMensajeError.picImagen.Picture = frmMensajeError.ListaIconos.ListImages.Item(icono).Picture
    
    If Not cancelar Then

      frmMensajeError.btnNo.Visible = False
      frmMensajeError.btnSi.Left = 1718
    Else
      frmMensajeError.btnNo.TabIndex = 0
    End If
    
    frmMensajeError.Show vbModal

    Mensaje = frmMensajeError.gbAceptar
    Unload frmMensajeError

End Function

' Setea un valor en una llave del registro
'Public Function SetKeyValueReg(sPathKey As String, sKeyAffected As String, sNewValue As String) As Boolean
'SetKeyValueReg = UpdateKey(HKEY_LOCAL_MACHINE, sPathKey, sKeyAffected, sNewValue)
'End Function

'Public Function FindAkey(sPathKey As String, sKeyAffected As String, sValue As String) As Boolean
'Dim sVal As String
'On Error GoTo ErrH
'    If Not GetKeyValue(HKEY_LOCAL_MACHINE, sPathKey, sKeyAffected, sVal) Then
'        err.Raise vbObjectError + 1, "ConfServer/ConnectionStrings/GetConnString", "No pudo recuperarse el nombre de la base de datos."
'        Exit Function
'    End If
'    sVal = Trim(sVal)
'    If sVal = "" Then
'        err.Raise vbObjectError + 1, "ConfServer/ConnectionStrings/GetConnString", "No pudo recuperarse el nombre de la base de datos."
'        Exit Function
'    End If
'    If sValue = sVal Then
'        FindAkey = True
'    Else
'        FindAkey = False
'    End If
'    Exit Function
'ErrH:
'    err.Raise err.Number, "ConfServer/ConnectionStrings/GetConnString", err.Description
'
'End Function
'
'
'Public Function GetAKey(sKey As String) As String
''On Error GoTo ErrH
'    Dim sResult As String
'    Dim sRutaKey As String
'    Dim bOk As Boolean
'    sResult = ""
'    sRutaKey = "SOFTWARE\EXACTUS\POSOFFLINE\TIENDA\"
'
'    If Not GetKeyValue(HKEY_LOCAL_MACHINE, sRutaKey, "RutaSincronizacion", sResult) Then
'       bOk = Mensaje("Hubo un error al obtener la ruta del Carga Pos", ICO_ERROR, False)
'        Exit Function
'    End If
'    sResult = Trim(sResult)
'    GetAKey = sResult
'
'End Function

Public Function ConexionBD() As Boolean    'Función que inicia la conexión a la base de datos
  Dim sPassword As String
  Dim lbOk As Boolean 'Indica que el proceso está bien
  Dim lsConexion As String  'Va a almacenar la hilera de conexión a la base de datos
  On Error GoTo error:  'Para capturar el evento de error
  
  Set gConet = New ADODB.Connection 'Inicializa la variable de conexión
  lbOk = True
    lsConexion = GetConectString(gsTypeDB, gsNombreServidor, gsNombreBaseDatos, gsUser, gsPassword)
 '   gConet.Open "Provider=SQLNCLI.1;Password=admin;Persist Security Info=True;User ID=SysUser;Initial Catalog=exactus;Data Source=serverg4s"
    gConet.Open lsConexion 'Realiza la conexión a la BD
    gConet.CommandTimeout = 300
    If (gConet.Errors.Count > 0) Then  'Si no hubo conexión
      lbOk = False  'No se cuenta con una conexión
    End If
    gsConetstr = lsConexion
    ConexionBD = lbOk 'Asigna el valor de retorno
  Exit Function
  
error:
    lbOk = False
    gsOperacionError = "constr :" & lsConexion & err.Description
    ConexionBD = lbOk 'Asigna el valor de retorno
End Function

Public Function DesconectaBD() As Boolean
Dim lbOk As Boolean
On Error GoTo error
If gConet.State = adStateOpen Then
  lbOk = True
  gConet.Close
End If
DesconectaBD = lbOk
Exit Function
error:
  lbOk = False
  Resume Next
End Function


Public Function GetConectString(sTypeDB As String, gsNombreServidor As String, gsNombreBaseDatos As String, gsUser As String, gsClave As String) As String
Dim strConn As String
Select Case UCase(sTypeDB)
    Case "SQL2000":
        strConn = "Provider=SQLOLEDB.1;" & _
                    "Password=" & gsPassword & ";" & _
                    "Persist Security Info=False;" & _
                    "User ID=" & gsUser & ";" & _
                    "Initial Catalog=" & gsNombreBaseDatos & ";" & _
                    "Data Source=" & gsNombreServidor & ";" & _
                    "Connect Timeout=5000"
    Case "SQL2005":
        strConn = "Provider=SQLNCLI.1;" & _
                    "Password=" & gsPassword & ";" & _
                    "Persist Security Info=True;" & _
                    "User ID=" & gsUser & ";" & _
                    "Initial Catalog=" & gsNombreBaseDatos & ";" & _
                    "Data Source=" & gsNombreServidor & ";" & _
                    "Connect Timeout=5000" & ";"
    Case "SQL2008":
     If gbEquipo64Bits = False Then
        strConn = "Provider=SQLNCLI.10;" & _
                    "Pwd=" & gsPassword & ";" & _
                    "Persist Security Info=True;" & _
                    "UID=" & gsUser & ";" & _
                    "DataBase=" & gsNombreBaseDatos & ";" & _
                    "Server=" & gsNombreServidor & ";" & _
                    "Connect Timeout=5000" & ";"
     Else
        strConn = "Provider=SQLNCLI10;" & _
                    "Server=" & gsNombreServidor & ";" & _
                    "DataBase=" & gsNombreBaseDatos & ";" & _
                    "UID=" & gsUser & ";" & _
                    "Pwd=" & gsPassword & ";" & _
                    "Connect Timeout=5000" & ";"

     End If

    

End Select
GetConectString = strConn
End Function


Public Function LeerConfiguracion() As Boolean
  'Lee del archivo INI
  Dim lsTemp As String * 30
  Dim lbOk As Boolean

  On Error GoTo error
  lbOk = True
  '***************SECCION DE SERVIDOR
  gsTypeDB = ""
  GetPrivateProfileString "SERVIDOR", "TYPEDB", _
      "NO", lsTemp, 32, "SDF.ini" '*** STRING ***
  gsTypeDB = LimpiaNulos(lsTemp)
  gsTypeDB = RTrim(gsTypeDB) 'Almacena el nombre del servidor de base de datos
  If gsTypeDB = "" Or gsTypeDB = "NO" Then GoTo error
  
  gsNombreServidor = ""
  GetPrivateProfileString "SERVIDOR", "SERVER", _
      "NO", lsTemp, 32, "SDF.ini" '*** STRING ***
  gsNombreServidor = LimpiaNulos(lsTemp)
  gsNombreServidor = RTrim(gsNombreServidor) 'Almacena el nombre del servidor de base de datos
  If gsNombreServidor = "" Or gsNombreServidor = "NO" Then GoTo error
  
  gsNombreBaseDatos = ""
  GetPrivateProfileString "SERVIDOR", "DATABASE", _
      "NO", lsTemp, 32, "SDF.ini" '*** STRING ***
  gsNombreBaseDatos = LimpiaNulos(lsTemp)
  gsNombreBaseDatos = RTrim(gsNombreBaseDatos) 'Almacena el nombre de la de base de datos
  If gsNombreBaseDatos = "" Or gsNombreBaseDatos = "NO" Then GoTo error
 
  
  gsCompania = ""
  GetPrivateProfileString "SERVIDOR", "COMPANIA", _
      "NO", lsTemp, 32, "SDF.ini" '*** STRING ***
  gsCompania = LimpiaNulos(lsTemp)
  gsCompania = RTrim(gsCompania) 'Almacena el nombre de la de base de datos
  If gsCompania = "" Or gsCompania = "NO" Then GoTo error
  
  gsUser = ""
  GetPrivateProfileString "SERVIDOR", "USER", _
      "NO", lsTemp, 16, "SDF.ini" '*** STRING ***
  gsUser = LimpiaNulos(lsTemp)
  gsUser = RTrim(gsUser) 'Almacenará el nombre del servidor
  If gsUser = "" Or gsUser = "NO" Then GoTo error
  
    gsPassword = ""
  GetPrivateProfileString "SERVIDOR", "PASSWORD", _
      "NO", lsTemp, 30, "SDF.ini" '*** STRING ***
  
  gsPassword = LimpiaNulos(lsTemp)
  gsPassword = RTrim(gsPassword) 'Almacenará el nombre del servidor
'  gsPassword = Decrypt(gsPassword, 13)
  
  If gsPassword = "" Or gsPassword = "NO" Then GoTo error
  

  LeerConfiguracion = lbOk
  Exit Function

error:  'En caso de haber un error
  lbOk = False
  LeerConfiguracion = lbOk
End Function

Public Sub CargaNominas(rst As ADODB.Recordset, sTipo As String, cbo As ComboBox)
Dim lbOk As Boolean
'On Error GoTo error
lbOk = True
  GSSQL = "Select Nomina "
  GSSQL = GSSQL & " FROM " & "." & "zNominaTransfGFS "  'Constuye la sentencia SQL
  GSSQL = GSSQL & " WHERE TIPO ='" & sTipo & "'"
  GSSQL = GSSQL & " Order by orden "
  If rst.State = adStateOpen Then rst.Close
  rst.ActiveConnection = gConet
  rst.Open GSSQL ' , gConet, adOpenDynamic, adLockOptimistic, adCmdText    'Ejecuta la sentencia
cbo.Clear
If Not (rst.BOF And rst.EOF) Then  'Si no es válido
    rst.MoveFirst
    While Not rst.EOF
        cbo.AddItem rst("NOMINA").value
        rst.MoveNext
    Wend
End If

End Sub

'Public Function GetConectString(gsNombreServidor As String, gsNombreBaseDatos As String, gsUser As String, gsClave As String) As String
'Dim strConn As String
'strConn = "Provider=SQLOLEDB.1;" & _
'            "Password=" & gsPassword & ";" & _
'            "Persist Security Info=False;" & _
'            "User ID=" & gsUser & ";" & _
'            "Initial Catalog=" & gsNombreBaseDatos & ";" & _
'            "Data Source=" & gsNombreServidor & ";" & _
'            "Connect Timeout=30"
'GetConectString = strConn
'End Function

Public Function GetRecordset(strSource As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Cerrar As Boolean
    Set rs = New ADODB.Recordset

    Set rs.ActiveConnection = gConet
    rs.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rs.CursorLocation = adUseClient ' Cursor local al cliente
    rs.LockType = adLockOptimistic
'    rs.CursorLocation = adUseClient
'    rs.CursorType = adOpenStatic
'    rs.LockType = adLockBatchOptimistic
    rs.Source = strSource
    rs.Open
    Set rs.ActiveConnection = Nothing
    Set GetRecordset = rs
End Function

Public Function invGetSugeridoLote(IdBodega As Integer, IdProducto As Integer, Cantidad As Double) As ADODB.Recordset
    
'Dim rst As ADODB.Recordset
'Set rst = New ADODB.Recordset
'rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
'rst.CursorType = adOpenStatic  'Asigna un cursor dinamico
'rst.CursorLocation = adUseClient ' Cursor local al cliente
'rst.LockType = adLockOptimistic
'
'
'On Error GoTo error
'
'  GSSQL = "invGetSugeridoLote " & IdBodega & "," & IdProducto & "," & Cantidad
'
'  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia
'
'  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
'      Set rst = Nothing ' "Error en la búsqueda del artículo !!!" & err.Description
'  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
'    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
'   Set rst = gRegistrosCmd
'
'  End If
'
'  Set invGetSugeridoLote = rst
'   'gRegistrosCmd.Close
'  Exit Function
'error:
'  Set rst = gRegistrosCmd
'  gsOperacionError = "Ocurrió un error en la operación de búsqueda de la descripción " & err.Description
'  Resume Next
 
    

    Dim rs As ADODB.Recordset
    Dim Cerrar As Boolean
    Set rs = New ADODB.Recordset

    Set rs.ActiveConnection = gConet


    rs.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rs.CursorLocation = adUseClient ' Cursor local al cliente
    rs.LockType = adLockOptimistic
'    rs.CursorLocation = adUseClient
'    rs.CursorType = adOpenStatic
'    rs.LockType = adLockBatchOptimistic
    rs.Source = gsCompania & ".invGetSugeridoLote " & IdBodega & "," & IdProducto & "," & Cantidad
    rs.Open
    Set rs.ActiveConnection = Nothing
    Set invGetSugeridoLote = rs
End Function

Public Function GetDisctinctRecordCount(rs As ADODB.Recordset) As Integer
Dim sFacturaAnterior As String
Dim posActual As Variant
Dim iCount As Integer
iCount = 0
If Not rs.EOF Then
    posActual = rs.Bookmark
    rs.MoveFirst
    sFacturaAnterior = ""
        While Not rs.EOF
            If sFacturaAnterior <> rs("Factura").value Then
                iCount = iCount + 1
            End If
        sFacturaAnterior = rs("Factura").value
        rs.MoveNext
        Wend
    rs.Bookmark = posActual
End If
GetDisctinctRecordCount = iCount
End Function



Public Function SustituyeChar(strFuente As String, sChar As String, sNewChar As String)
Dim i As Integer
Dim lbOk As Boolean
Dim strTemporal As String
Dim strResult As String
lbOk = False
i = 1
If strFuente <> "" Then

If OnlythisChar(strFuente, ".") Then
    SustituyeChar = ""
    Exit Function
End If

strTemporal = strFuente
    If InStr(strTemporal, sChar) = 0 Then
      SustituyeChar = "%" & strTemporal & "%"
      Exit Function
    End If
    strResult = ""
      While Not lbOk And i <= Len(strFuente)
        If InStr(strTemporal, sChar) <> 0 Then
           i = InStr(strTemporal, sChar) ' posición del caracter
           strResult = strResult & Mid(strTemporal, 1, i - 1) & sNewChar '& Mid(strTemporal, i + 1)
           strTemporal = Mid(strTemporal, i + 1, Len(strFuente))
           i = i + 1
        Else
           strResult = strResult & strTemporal
           lbOk = True
        End If
      Wend
    SustituyeChar = strResult
Else
  SustituyeChar = ""
End If
End Function

Public Function OnlythisChar(strFuente As String, sChar As String) As Boolean
Dim i As Integer
Dim lbOk As Boolean
Dim stmpChar As String
stmpChar = ""
lbOk = True
i = 1
While i <= Len(strFuente) And lbOk
stmpChar = Mid(strFuente, i, 1)
If stmpChar <> sChar Then
    lbOk = False
End If
i = i + 1
Wend
OnlythisChar = lbOk
End Function

' Formato de un text box
' R : read only O : Obrigatorio P : Optional
Public Sub fmtTextbox(txtbox As TextBox, iModo As String)
Select Case UCase(iModo)
  Case "R" ' formatode lectura solamente
    txtbox.BackColor = &HE0E0E0   '&H8000000F
    txtbox.ForeColor = vbBlack
    txtbox.Locked = True
  Case "O" ' que sea obligatorio
    txtbox.BackColor = &HC0FFFF
    txtbox.Locked = False
  Case "P"
    txtbox.BackColor = &H80000005  ' que sea opcional
    txtbox.Locked = False
End Select
End Sub
' Valida un control text box para un dato que debe ser NUMERO
Public Function Val_TextboxNum(txtbox As TextBox)
Dim lbOk As Boolean
lbOk = True
If Trim(txtbox.Text) = "" Then

  lbOk = False
  Val_TextboxNum = lbOk
  Exit Function
End If


If Not IsNumeric(txtbox.Text) Then
  
  lbOk = False
  Val_TextboxNum = lbOk

End If

Val_TextboxNum = lbOk

End Function

' ----------------- Para el common dialog
Function GetFileName(FileName As Variant)
    ' Muestra un cuadro de diálogo Guardar como y devuelve un nombre de archivo.
    ' Si el usuario elige Cancelar, devuelve una cadena vacía.
    On Error Resume Next
    
    frmMain.CommonDialog1.InitDir = App.Path
    frmMain.CommonDialog1.DialogTitle = "Exportar a formato Excel"
    frmMain.CommonDialog1.FileName = FileName
    frmMain.CommonDialog1.Filter = " EXL (*.xls)|*.xls"
    frmMain.CommonDialog1.ShowSave
    If err <> 32755 Then    ' El usuario eligió Cancelar.
        GetFileName = frmMain.CommonDialog1.FileName
    Else
        GetFileName = ""
    End If
End Function

Function GetFileNamePDF(FileName As Variant)
    ' Muestra un cuadro de diálogo Guardar como y devuelve un nombre de archivo.
    ' Si el usuario elige Cancelar, devuelve una cadena vacía.
    On Error Resume Next
    
    frmMain.CommonDialog1.InitDir = App.Path
    frmMain.CommonDialog1.DialogTitle = "Exportar a formato PDF"
    frmMain.CommonDialog1.FileName = FileName
    frmMain.CommonDialog1.Filter = " PDF (*.pdf)|*.pdf"
    frmMain.CommonDialog1.ShowSave
    If err <> 32755 Then    ' El usuario eligió Cancelar.
        GetFileNamePDF = frmMain.CommonDialog1.FileName
    Else
        GetFileNamePDF = ""
    End If
End Function

' Abre Excel y el archivo que se le pasa como parametro
' iAbrir = 0 --> no abrir; iabrir = 1 --> abrir
Public Function OpenExcel(iAbrir As Integer, sNombre As String) As Boolean
Dim lbOk As Boolean
lbOk = False
On Error GoTo errores
Dim objExcel
If iAbrir > 0 Then
Set objExcel = CreateObject("Excel.Application")
  objExcel.Visible = True
  objExcel.Workbooks.Open sNombre
  lbOk = True
End If
  OpenExcel = lbOk
Exit Function
errores:
  MsgBox "Error al abrir la aplicación Excel", vbOKOnly, "Error "
  lbOk = False
  Resume Next
End Function

Public Function GetNameMonth(d As Date) As String
Dim sMes As String
sMes = ""
Select Case Month(d)
 Case 1: sMes = "Enero"
 Case 2: sMes = "Febrero"
 Case 3: sMes = "Marzo"
 Case 4: sMes = "Abril"
 Case 5: sMes = "Mayo"
 Case 6: sMes = "Junio"
 Case 7: sMes = "Julio"
 Case 8: sMes = "Agosto"
 Case 9: sMes = "Septiembre"
 Case 10: sMes = "Octubre"
 Case 11: sMes = "Noviembre"
 Case 12: sMes = "Diciembre"
 
End Select

GetNameMonth = sMes
End Function

Public Function ExiteRstKey(rstFuente As ADODB.Recordset, sFiltro As String) As Boolean
Dim lbOk As Boolean
Dim rstClone As ADODB.Recordset
Dim bmPos As Variant
lbOk = False
If Not (rstFuente.EOF And rstFuente.BOF) Then
    Set rstClone = New ADODB.Recordset
        bmPos = rstFuente.Bookmark
        rstClone.Filter = adFilterNone
        Set rstClone = rstFuente.Clone
        rstClone.Filter = sFiltro
        If Not rstClone.EOF Then ' Si existe
          lbOk = True
        End If
        rstFuente.Filter = adFilterNone
        rstFuente.Bookmark = bmPos
    rstClone.Filter = adFilterNone
End If
ExiteRstKey = lbOk
End Function



Public Function UserMayAccess(sUsuario As String, iAccion As Integer, iModulo As Integer) As Boolean
Dim lbOk As Boolean
On Error GoTo errores
Dim sCriterioAcceso As String
sCriterioAcceso = "USUARIO ='" & UCase(sUsuario) & "' AND IDACCION=" & Str(iAccion) & " AND IDMODULO=" & Str(iModulo)
If ExiteRstKey(grRecordsetAcceso, sCriterioAcceso) Then
    lbOk = True
Else
    lbOk = False
End If
UserMayAccess = lbOk
Exit Function
errores:
lbOk = False
UserMayAccess = lbOk
End Function

' carga parametros de la preparacion diferencial cambiario CC
Public Function CargaParametrosPreparacionDifCC() As Boolean
Dim lbOk As Boolean
On Error GoTo error
lbOk = True
  GSSQL = "SELECT top 1 FechaCorte " & _
          " FROM " & gsCompania & "." & "parmaPreparacionProcesoDiferencial "

  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbOk = False  'Indica que no es válido

  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    gParametrosPrepDifCC.FechaCorte = gRegistrosCmd("FechaCorte").value
    
    lbOk = True
  End If
  CargaParametrosPreparacionDifCC = lbOk
  gRegistrosCmd.Close
  Exit Function
error:
  lbOk = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next

End Function

Public Function CargaParametrosPreparacionDifCP() As Boolean
Dim lbOk As Boolean
On Error GoTo error
lbOk = True
  GSSQL = "SELECT top 1 FechaCorte " & _
          " FROM " & gsCompania & "." & "parmaPreparacionProcesoDiferencialCP "

  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbOk = False  'Indica que no es válido

  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    gParametrosPrepDifCC.FechaCorte = gRegistrosCmd("FechaCorte").value
    
    lbOk = True
  End If
  CargaParametrosPreparacionDifCP = lbOk
  gRegistrosCmd.Close
  Exit Function
error:
  lbOk = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next

End Function

Public Sub Destructor()
                If Not IsNull(grRecordsetAcceso) Then
                    Set grRecordsetAcceso = Nothing
                End If
                If Not IsNull(gRegistrosCmd) Then
                    Set gRegistrosCmd = Nothing
                End If
                If Not IsNull(gConet) Then
                    If gConet.State = 1 Then
                        gConet.Close
                        Set gConet = Nothing
                    End If
                End If
                If Not IsNull(gskey) Then
                    Set gskey = Nothing
                End If

                
End Sub

Public Function Update_TipoCambio(sTipoCambio As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
  GSSQL = gsCompania & "." & "repActualizaTipoCambioContrato " & sTipoCambio
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbOk = False
    End If

Update_TipoCambio = lbOk
Exit Function

error:
  lbOk = False
  Resume Next

End Function


Public Function Cierra_Periodo(sEmpleado As String, sPeriodo As String, sFecha As String, sPeriodoNuevo As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
GSSQL = gsCompania & ".sgvGetDetalleMovimientosEmpleado" & " '" & sEmpleado & "', '" & sPeriodo & "' , '" & sFecha & "', 1, 1, " & sPeriodoNuevo
gConet.BeginTrans
gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbOk = False
      gConet.RollbackTrans
      Cierra_Periodo = lbOk
      Exit Function
    End If

Cierra_Periodo = lbOk
gConet.CommitTrans
Exit Function

error:
  lbOk = False
  Resume Next

End Function

Public Function Baja_Empleado(sEmpleado As String, sPeriodo As String, sFecha As String, sComentario As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
GSSQL = gsCompania & ".sgvGetDetalleMovimientosEmpleado" & " '" & sEmpleado & "', " & sPeriodo & " , '" & sFecha & "', 1, null, null, 1,0, '" & gsUSUARIO & "','S', '" & sComentario & "'"
gConet.BeginTrans
gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbOk = False
      gConet.RollbackTrans
      Baja_Empleado = lbOk
      Exit Function
    End If

Baja_Empleado = lbOk
gConet.CommitTrans
Exit Function

error:
  lbOk = False
  Resume Next

End Function

Public Function Prepara_AnalisisAntiguedad(sUsuario As String, sFecha As String, sCodCliente As String, sCodCategoria As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
GSSQL = gsCompania & ".uspparmaPrepareAnalisisVencimiento" & " '" & sUsuario & "' , '" & sFecha & "' , '" & sCodCliente & "','" & sCodCategoria & "'"
'gConet.BeginTrans

gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbOk = False
      'gConet.RollbackTrans
      Prepara_AnalisisAntiguedad = lbOk
      Exit Function
    End If

Prepara_AnalisisAntiguedad = lbOk
'gConet.CommitTrans
Exit Function

error:
  lbOk = False
  Resume Next

End Function

Public Function uspparmaGetClasificacionporPago(sCodCliente As String, sFechaInicio As String, sFechaFin As String, sCodCategoria As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
GSSQL = gsCompania & ".uspparmaGetClasificacionporPago" & " '" & sCodCliente & "' , '" & sFechaInicio & "' , '" & sFechaFin & "','" & sCodCategoria & "'"
'gConet.BeginTrans

gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbOk = False
      'gConet.RollbackTrans
      uspparmaGetClasificacionporPago = lbOk
      Exit Function
    End If

uspparmaGetClasificacionporPago = lbOk
'gConet.CommitTrans
Exit Function

error:
  lbOk = False
  Resume Next

End Function



Public Function Prepara_Reportes(sPeriodo As String, sFecha As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
GSSQL = gsCompania & ".sgvGetDetalleMovimientosEmpleado" & " '*', '" & sPeriodo & "' , '" & sFecha & "', null, null,null,null, 1,'" & gsUSUARIO & "'"
gConet.BeginTrans
gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbOk = False
      gConet.RollbackTrans
      Prepara_Reportes = lbOk
      Exit Function
    End If

Prepara_Reportes = lbOk
gConet.CommitTrans
Exit Function

error:
  lbOk = False
  Resume Next

End Function



' Carga Parámetros del sistema
Public Function CargaParametros() As Boolean
Dim lbOk As Boolean
Dim iResultado As Integer
On Error GoTo error
lbOk = True
  GSSQL = "SELECT * " & _
          " FROM " & gsCompania & "." & "INVPARAMETROS "
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbOk = False  'Indica que no es válido
    
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    
    gparametros.NombreEmpresa = gRegistrosCmd("NombreEmpresa").value

    lbOk = True
  End If
  CargaParametros = lbOk
  gRegistrosCmd.Close
  Exit Function
error:
  lbOk = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function



' devuelve Proximo Consecutivo de la esquela
Public Function getNextConsecEsquela() As String
Dim lbOk As Boolean
Dim sConsecutivo As String
On Error GoTo error
sConsecutivo = ""
  GSSQL = "SELECT " & gsCompania & "." & "sgvGetNextConsecEsquela() NextConsecutivo"
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    sConsecutivo = ""  'Indica que no es válido
    
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    sConsecutivo = gRegistrosCmd("NextConsecutivo").value
  End If
  getNextConsecEsquela = sConsecutivo
  gRegistrosCmd.Close
  Exit Function
error:
  sConsecutivo = ""
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function

Public Function sgvRangoExistente(sInicio As String, sFin As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error
lbOk = False
  GSSQL = "SELECT * from " & gsCompania & "." & "sgvAlertas where " & sInicio & " between Inicio and Fin or " & sFin & " between inicio and fin "
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbOk = False
    
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbOk = True
  End If
  sgvRangoExistente = lbOk
  gRegistrosCmd.Close
  Exit Function
error:
  lbOk = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function


Public Function sgvEMPLEADOBAJA(sEmpleado As String) As String
Dim lbOk As Boolean
Dim iBAJA As Integer
On Error GoTo error
sConsecutivo = ""
  GSSQL = "SELECT " & gsCompania & "." & "sgvEMPLEADOBAJA('" & sEmpleado & "') BAJA"
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    iBAJA = 1  'Indica que no es válido
    
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    iBAJA = gRegistrosCmd("BAJA").value
  End If
  sgvEMPLEADOBAJA = iBAJA
  gRegistrosCmd.Close
  Exit Function
error:
  sConsecutivo = ""
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function



Public Function esSaldoInicialPeriodo(sTipoAccion As String) As Boolean
Dim lbOk As Boolean
Dim iesPeriodo As Boolean
On Error GoTo error
iesPeriodo = False
  GSSQL = "SELECT SaldoInicialPeriodo from " & gsCompania & "." & "SGVTIPOACCION where Activo = 1 and Tipo_Accion = '" & sTipoAccion & "'"
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    iesPeriodo = False  'Indica que no es válido
    
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    
    iesPeriodo = gRegistrosCmd("SaldoInicialPeriodo").value
  End If
  esSaldoInicialPeriodo = iesPeriodo
  gRegistrosCmd.Close
  Exit Function
error:
  iesPeriodo = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function

Public Function ExisteSaldoInicialPeriodo() As Boolean
Dim lbOk As Boolean
Dim iesPeriodo As Boolean
On Error GoTo error
iesPeriodo = False
  GSSQL = "SELECT SaldoInicialPeriodo from " & gsCompania & "." & "SGVTIPOACCION where Activo = 1 and SaldoInicialPeriodo=1"
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    iesPeriodo = False  'Indica que no es válido
    
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    
    iesPeriodo = gRegistrosCmd("SaldoInicialPeriodo").value
  End If
  ExisteSaldoInicialPeriodo = iesPeriodo
  gRegistrosCmd.Close
  Exit Function
error:
  iesPeriodo = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function


Public Function Num2Text(ByVal value As Double) As String


    Select Case value
        Case 0: Num2Text = "CERO"
        Case 1: Num2Text = "UN"
        Case 2: Num2Text = "DOS"
        Case 3: Num2Text = "TRES"
        Case 4: Num2Text = "CUATRO"
        Case 5: Num2Text = "CINCO"
        Case 6: Num2Text = "SEIS"
        Case 7: Num2Text = "SIETE"
        Case 8: Num2Text = "OCHO"
        Case 9: Num2Text = "NUEVE"
        Case 10: Num2Text = "DIEZ"
        Case 11: Num2Text = "ONCE"
        Case 12: Num2Text = "DOCE"
        Case 13: Num2Text = "TRECE"
        Case 14: Num2Text = "CATORCE"
        Case 15: Num2Text = "QUINCE"
        Case Is < 20: Num2Text = "DIECI" & Num2Text(value - 10)
        Case 20: Num2Text = "VEINTE"
        Case Is < 30: Num2Text = "VEINTI" & Num2Text(value - 20)
        Case 30: Num2Text = "TREINTA"
        Case 40: Num2Text = "CUARENTA"
        Case 50: Num2Text = "CINCUENTA"
        Case 60: Num2Text = "SESENTA"
        Case 70: Num2Text = "SETENTA"
        Case 80: Num2Text = "OCHENTA"
        Case 90: Num2Text = "NOVENTA"
        Case Is < 100: Num2Text = Num2Text(Int(value \ 10) * 10) & " Y " & Num2Text(value Mod 10)
        Case 100: Num2Text = "CIEN"
        Case Is < 200: Num2Text = "CIENTO " & Num2Text(value - 100)
        Case 200, 300, 400, 600, 800: Num2Text = Num2Text(Int(value \ 100)) & "CIENTOS"
        Case 500: Num2Text = "QUINIENTOS"
        Case 700: Num2Text = "SETECIENTOS"
        Case 900: Num2Text = "NOVECIENTOS"
        Case Is < 1000: Num2Text = Num2Text(Int(value \ 100) * 100) & " " & Num2Text(value Mod 100)
        Case 1000: Num2Text = "MIL"
        Case Is < 2000: Num2Text = "MIL " & Num2Text(value Mod 1000)
        Case Is < 1000000: Num2Text = Num2Text(Int(value \ 1000)) & " MIL"
            If value Mod 1000 Then Num2Text = Num2Text & " " & Num2Text(value Mod 1000)
        Case 1000000: Num2Text = "UN MILLON"
        Case Is < 2000000: Num2Text = "UN MILLON " & Num2Text(value Mod 1000000)
        Case Is < 1000000000000#: Num2Text = Num2Text(Int(value / 1000000)) & " MILLONES "
            If (value - Int(value / 1000000) * 1000000) Then Num2Text = Num2Text & " " & Num2Text(value - Int(value / 1000000) * 1000000)
        Case 1000000000000#: Num2Text = "UN BILLON"
        Case Is < 2000000000000#: Num2Text = "UN BILLON " & Num2Text(value - Int(value / 1000000000000#) * 1000000000000#)
        Case Else: Num2Text = Num2Text(Int(value / 1000000000000#)) & " BILLONES"
            If (value - Int(value / 1000000000000#) * 1000000000000#) Then Num2Text = Num2Text & " " & Num2Text(value - Int(value / 1000000000000#) * 1000000000000#)
    End Select


End Function

Public Function zNumberToText(valor As Double, Optional sMoneda As String) As String
Dim dDecimal As Double
Dim iEntero As Long 'Integer
Dim sResultado As String
iEntero = Fix(valor) ' Fix(CLng(valor)) 'Int(valor)
dDecimal = valor - iEntero

sResultado = Num2Text(iEntero)
If dDecimal > 0 Then
    sResultado = sResultado & " " & sMoneda & " con " & Str(Round(dDecimal, 4)) & "/100"
End If
zNumberToText = sResultado & " " & sMoneda
End Function

Public Function UpdateUsuarioModuloRole(sUsuario As String, sModulo As String, lst As ListBox) As Boolean
On Error GoTo errores
Dim lbOk As Boolean
Dim rst As ADODB.Recordset
Dim sFiltro As String
Dim iRole As Integer
Dim i As Integer
Dim lsRole As String
Dim lbHayRegistros As Boolean
lbOk = False
lbHayRegistros = False
Set rst = New ADODB.Recordset
If rst.State = adStateOpen Then rst.Close
rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
rst.CursorLocation = adUseClient ' Cursor local al cliente
rst.LockType = adLockOptimistic
GSSQL = "SELECT DISTINCT IDMODULO, DESCRMODULO, USUARIO, IDROLE, DESCRROLE " & _
          " FROM " & gsCompania & "." & "vsecPrivilegios "
GSSQL = GSSQL & " WHERE USUARIO = '" & sUsuario & "'"
GSSQL = GSSQL & " AND IDMODULO = " & sModulo
GSSQL = GSSQL & " ORDER BY IDMODULO, IDROLE "
If rst.State = adStateOpen Then rst.Close
Set rst = GetRecordset(GSSQL)
    If (rst.EOF And rst.BOF) Then
        lbHayRegistros = False
    Else
        lbHayRegistros = True
    End If

'If Not (rst.EOF And rst.BOF) Then

For i = 0 To lst.ListCount - 1
    If lbHayRegistros Then
        rst.MoveFirst
    End If
    lsRole = GetCodeBeforeMinus(lst.List(i))
    sFiltro = "IDRole =" & lsRole
    If lst.Selected(i) And Not lbHayRegistros Then
        ' Insertar Registro lst.Selected(i) = True
            lbOk = Insert_UsuarioRole(sModulo, lsRole, sUsuario)
    
    End If
    If Not lst.Selected(i) And lbHayRegistros Then
        If ExiteRstKey(rst, sFiltro) Then
        ' Eliminar el Registro
            lbOk = Delete_UsuarioRole(sModulo, lsRole, sUsuario)
        End If
    End If
Next i

'End If


Set rst = Nothing
UpdateUsuarioModuloRole = lbOk
Exit Function
errores:
lbOk = False
UpdateUsuarioModuloRole = lbOk

End Function

Public Function Insert_UsuarioRole(sModulo As String, sRole As String, sUsuario As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
  GSSQL = "INSERT  " & gsCompania & "." & "secUSUARIOROLE (IDMODULO, IDROLE, USUARIO)"
  GSSQL = GSSQL & " VALUES (" & sModulo & " , " & sRole & " , '" & sUsuario & "')"

    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbOk = False
    End If

Insert_UsuarioRole = lbOk
Exit Function

error:
  lbOk = False
  Resume Next

End Function


Public Function LoadModuloRole(sModulo As String, lst As ListBox) As Boolean
On Error GoTo errores
Dim lbOk As Boolean
Dim sItem As String
Dim rst As ADODB.Recordset

lbOk = False
Set rst = New ADODB.Recordset
If rst.State = adStateOpen Then rst.Close
rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
rst.CursorLocation = adUseClient ' Cursor local al cliente
rst.LockType = adLockOptimistic
GSSQL = "SELECT distinct IDROLE, DESCR " & _
          " FROM " & gsCompania & "." & "vsecModuloRoleAccion "
GSSQL = GSSQL & " WHERE IDMODULO = " & sModulo
GSSQL = GSSQL & " ORDER BY  IDROLE "

If rst.State = adStateOpen Then rst.Close
Set rst = GetRecordset(GSSQL)

If Not (rst.EOF And rst.BOF) Then
rst.MoveFirst
lst.Clear
While Not rst.EOF
    sItem = Str(rst("IDROLE").value) & " - " & rst("DESCR").value
    lst.AddItem sItem
    lst.ItemData(lst.ListIndex) = rst("IDROLE").value
    lst.Tag = Str(rst("IDROLE").value)
    rst.MoveNext
Wend
End If
Set rst = Nothing
LoadModuloRole = lbOk
Exit Function
errores:
lbOk = False
LoadModuloRole = lbOk
End Function

Public Sub UncheckListBox(lst As ListBox)
Dim i As Integer
    For i = 0 To lst.ListCount - 1
            lst.Selected(i) = False
    Next i
End Sub




Public Function LoadRoleModulo(sUsuario As String, sModulo As String, lst As ListBox) As Boolean
On Error GoTo errores
Dim lbOk As Boolean
Dim rst As ADODB.Recordset
Dim sFiltro As String
Dim iRole As Integer
Dim i As Integer
lbOk = True
Set rst = New ADODB.Recordset
If rst.State = adStateOpen Then rst.Close
rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
rst.CursorLocation = adUseClient ' Cursor local al cliente
rst.LockType = adLockOptimistic
GSSQL = "SELECT DISTINCT IDMODULO, DESCRMODULO, USUARIO, IDROLE, DESCRROLE " & _
          " FROM " & gsCompania & "." & "vsecPrivilegios "
GSSQL = GSSQL & " WHERE USUARIO = '" & sUsuario & "'"
GSSQL = GSSQL & " AND IDMODULO = " & sModulo
GSSQL = GSSQL & " ORDER BY IDMODULO, IDROLE "
If rst.State = adStateOpen Then rst.Close
Set rst = GetRecordset(GSSQL)

If Not (rst.EOF And rst.BOF) Then
'rst.MoveFirst
'While rst.EOF
    For i = 0 To lst.ListCount - 1
        rst.MoveFirst
        sFiltro = "IDRole =" & GetCodeBeforeMinus(lst.List(i))
        If ExiteRstKey(rst, sFiltro) Then
            lst.Selected(i) = True
        Else
            lst.Selected(i) = False
        End If
    Next i
'rst.MoveNext
'Wend
lst.ListIndex = 0
End If


Set rst = Nothing
LoadRoleModulo = lbOk
Exit Function
errores:
lbOk = False
LoadRoleModulo = lbOk
End Function

Public Function GetCodeBeforeMinus(gsDato As String) As String
Dim i As Integer
Dim sResult As String
sResult = ""
i = InStr(1, gsDato, "-", vbTextCompare)
If i > 0 Then
    sResult = Trim(Left(gsDato, i - 1))
End If
GetCodeBeforeMinus = sResult
End Function

Public Function GetCodeBeforevbTab(gsDato As String) As String
Dim i As Integer
Dim sResult As String
sResult = ""
i = InStr(1, gsDato, vbTab, vbTextCompare)
If i > 0 Then
    sResult = Trim(Left(gsDato, i - 1))
End If
GetCodeBeforevbTab = sResult
End Function


Public Function LoadAccionRole(sModulo As String, sRole As String, lst As ListBox) As Boolean
On Error GoTo errores
Dim lbOk As Boolean
Dim sItem As String
Dim rst As ADODB.Recordset

lbOk = False
Set rst = New ADODB.Recordset
If rst.State = adStateOpen Then rst.Close
rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
rst.CursorLocation = adUseClient ' Cursor local al cliente
rst.LockType = adLockOptimistic
GSSQL = "SELECT IDROLE, IDACCION, DESCRACCION " & _
          " FROM " & gsCompania & "." & "vsecModuloRoleAccion "
GSSQL = GSSQL & " WHERE IDMODULO = " & sModulo & " AND IDROLE =" & sRole
GSSQL = GSSQL & " ORDER BY IDROLE, IDACCION "

If rst.State = adStateOpen Then rst.Close
Set rst = GetRecordset(GSSQL)

If Not (rst.EOF And rst.BOF) Then
rst.MoveFirst
lst.Clear
While Not rst.EOF
    sItem = Str(rst("IDACCION").value) & " - " & rst("DESCRACCION").value
    lst.AddItem sItem
    rst.MoveNext
Wend
Else
lst.Clear
End If
Set rst = Nothing
LoadAccionRole = lbOk
Exit Function
errores:
lbOk = False
LoadAccionRole = lbOk
End Function

Public Function Delete_UsuarioRole(sModulo As String, sRole As String, sUsuario As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
  GSSQL = "DELETE FROM " & gsCompania & "." & "secUSUARIOROLE "
  GSSQL = GSSQL & " WHERE IDMODULO =" & sModulo & " AND IDROLE = " & sRole & " AND USUARIO = " & "'" & sUsuario & "'"

    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto el UsuarioRole."
      lbOk = False
    End If

Delete_UsuarioRole = lbOk
Exit Function

error:
  lbOk = False
  Resume Next

End Function

Public Function sgvActualizaSaldoEmpleado(sOperacion As String, sPeriodo As String, sTipo_Accion As String, sEmpleado As String, _
    sFecha As String, sCantDias As String, sComentario As String, sUsuario As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
  GSSQL = gsCompania & ".sgvActualizaSaldoEmpleado '" & sOperacion & "'," & sPeriodo & ", '" & sTipo_Accion & "','" & sEmpleado & "', '" & _
  sFecha & "', " & sCantDias & ", '" & sComentario & "','" & sUsuario & "'"

    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el Beneficiado."
      lbOk = False
    End If

sgvActualizaSaldoEmpleado = lbOk
Exit Function

error:
  lbOk = False
  Resume Next

End Function

Public Function uspsgvSetBajaEmpleado(sEmpleado As String, sPeriodo As String, _
    sFecha As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
  GSSQL = gsCompania & ".uspsgvSetBajaEmpleado '" & sEmpleado & "'," & sPeriodo & ", '" & _
  sFecha & "'"

    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el Beneficiado."
      lbOk = False
    End If

uspsgvSetBajaEmpleado = lbOk
Exit Function

error:
  lbOk = False
  Resume Next

End Function


Public Function sgvActualizaTipoAccion(sTipo_Accion As String, sDescr As String, sPrioridad As String, sFactor As String, _
    sflgRRHH As String, sAfectaVacaciones As String, sCalculoVacacional As String, sEsAjuste As String, sEsSaldoInicialPeriodo As String, sEsSaldoSalidaEmpleado As String, _
    sActivo As String, sSeRepiteMes As String, sUnaVezPeriodo As String, sUnaVezMes As String, sMaximoUno As String, sOperacion As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
  GSSQL = gsCompania & ".sgvActualizaTipoAccion '" & sTipo_Accion & "','" & sDescr & "' , " & sPrioridad & ", " & sFactor & ", " & sflgRRHH & "," & sAfectaVacaciones & "," _
  & sCalculoVacacional & "," & sEsAjuste & "," & sEsSaldoInicialPeriodo & "," & sEsSaldoSalidaEmpleado & "," & sActivo & "," & sSeRepiteMes & "," & sUnaVezPeriodo & "," & sUnaVezMes & "," & sMaximoUno & ",'" & sOperacion & "'"

    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el Beneficiado."
      lbOk = False
    End If

sgvActualizaTipoAccion = lbOk
Exit Function

error:
  lbOk = False
  Resume Next

End Function

Public Function sgvUpdatesgvHeaderEsquelas(sOperacion As String, sEmpleado As String, sAnio As String, sMes As String, sConsecutivo As String, sFechaCreacion As String, sFechaSaldo As String, sSaldo As String, sDiasSolicitados As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
sFechaCreacion = Format(sFechaCreacion, "YYYYmmdd")
sFechaSaldo = Format(sFechaSaldo, "YYYYmmdd")

  GSSQL = gsCompania & ".sgvUpdatesgvHeaderEsquelas '" & sOperacion & "','" & sEmpleado & "'," & sAnio & "," & sMes & ",'" & sConsecutivo & "' , '" & sFechaCreacion & "', '" & sFechaSaldo & "', " & sSaldo & "," & sDiasSolicitados

    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el Beneficiado."
      lbOk = False
    End If

sgvUpdatesgvHeaderEsquelas = lbOk
Exit Function

error:
  lbOk = False
  Resume Next

End Function


Public Function sgvUpdatesgvEsqueladetalle(sOperacion As String, sEmpleado As String, sAnio As String, sMes As String, sConsecutivo As String, sDiaSolicitado As String, sCantidad As String, sComentario As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
sDiaSolicitado = Format(sDiaSolicitado, "YYYYmmdd")


  GSSQL = gsCompania & ".sgvUpdatesgvEsqueladetalle '" & sOperacion & "','" & sEmpleado & "'," & sAnio & "," & sMes & ",'" & sConsecutivo & "' , '" & sDiaSolicitado & "'," & sCantidad & ",'" & sComentario & "'"

    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el Beneficiado."
      lbOk = False
    End If

sgvUpdatesgvEsqueladetalle = lbOk
Exit Function

error:
  lbOk = False
  Resume Next

End Function


Public Function invUpdateBodega(sOperacion As String, sIDBodega As String, sDescrBodega As String, sActivo As String, sFactura As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True

  GSSQL = gsCompania & ".invUpdateBodega '" & sOperacion & "'," & sIDBodega & ",'" & sDescrBodega & "'," & sActivo & "," & sFactura

    
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
        'gsOperacionError = "Eliminando el Beneficiado."
        SetMsgError "Ocurrió un error actualizando La bodega. ", err
      lbOk = False
    End If

invUpdateBodega = lbOk
Exit Function

error:
  lbOk = False
  Resume Next

End Function

Public Function invUpdateProducto(sOperacion As String, sIDProducto As String, sDescr As String, sImpuesto As String, _
sEsMuestra As String, sEsControlado As String, sClasif1 As String, sClasif2 As String, sClasif3 As String, _
sEsEtico As String, sBajaPrecioDistribuidor As String, sIDProveedor As String, sCostoUltLocal As String, _
sCostoUltDolar As String, sCostoUltPromLocal As String, sCostoUltPromDolar As String, sPrecioPublicoLocal As String, _
sPrecioFarmaciaLocal As String, sPrecioCIFLocal As String, sPrecioFOBLocal As String, sIDPresentacion As String, _
sBajaPrecioProveedor As String, sPorcDescAlzaProveedor As String, sUserInsert As String, sUserUpdate As String, _
sActivo As String, sCodigoBarra As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
  GSSQL = ""
  GSSQL = gsCompania & ".invUpdateProducto '" & sOperacion & "'," & sIDProducto & ",'" & sDescr & "','" & sImpuesto & "',"
  GSSQL = GSSQL & sEsMuestra & "," & sEsControlado & ",'" & sClasif1 & "','" & sClasif2 & "','" & sClasif3 & "'," & sEsEtico & ","
  GSSQL = GSSQL & sBajaPrecioDistribuidor & "," & sIDProveedor & "," & sCostoUltLocal & "," & sCostoUltDolar & "," & sCostoUltPromLocal & ","
  GSSQL = GSSQL & sCostoUltPromDolar & "," & sPrecioPublicoLocal & "," & sPrecioFarmaciaLocal & "," & sPrecioCIFLocal & ","
  GSSQL = GSSQL & sPrecioFOBLocal & ",'" & sIDPresentacion & "'," & sBajaPrecioProveedor & "," & sPorcDescAlzaProveedor & ",'"
  GSSQL = GSSQL & sUserInsert & "','" & sUserUpdate & "'," & sActivo & ",'" & sCodigoBarra & "'"
    
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      'gsOperacionError = "Eliminando el Beneficiado. " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & err.Description
      SetMsgError "Ocurrió un error actualizando el producto. ", err
      lbOk = False
    End If

invUpdateProducto = lbOk
Exit Function

error:
  
  lbOk = False
  Resume Next

End Function

Public Function sgvActualizaAlerta(sCodigo As String, sDescr As String, sInicio As String, sFin As String, _
        sOperacion As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True

  GSSQL = gsCompania & ".sgvActualizaAlertas " & sCodigo & ",'" & sDescr & "' , " & sInicio & ", " & sFin & ",'" & sOperacion & "'"
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el Beneficiado."
      lbOk = False
    End If

sgvActualizaAlerta = lbOk
Exit Function

error:
  lbOk = False
  Resume Next

End Function



Public Function sgvActualizaPeriodo(sPeriodo As String, sFechaInicio As String, sFechaFinal As String, sEstado As String, _
        sOperacion As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True

sFechaInicio = Format(sFechaInicio, "yyyymmdd") & " 00:00:00.000"
sFechaFinal = Format(sFechaFinal, "yyyymmdd") & " 23:59:59.000"

  GSSQL = gsCompania & ".sgvActualizaPeriodo " & sPeriodo & ",'" & sFechaInicio & "' , '" & sFechaFinal & "', '" & sEstado & "','" & sOperacion & "'"

    
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el Beneficiado."
      lbOk = False
    End If

sgvActualizaPeriodo = lbOk
Exit Function

error:
  lbOk = False
  Resume Next

End Function

Public Function secUpdateUsuario(sUsuario As String, sNombre As String, sPassword As String, sActivo As String, sIDCDIS As String, _
        sOperacion As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True


  GSSQL = gsCompania & ".secUpdateUsuario '" & sOperacion & "','" & sUsuario & "','" & sNombre & "' , '" & sPassword & "', '" & sActivo & "'," & sIDCDIS

    
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el Beneficiado."
      lbOk = False
    End If

secUpdateUsuario = lbOk
Exit Function

error:
  lbOk = False
  Resume Next

End Function

Public Function secUpdateEmpleado(sEmpleado As String, sNombre As String, sFecha_Ingreso As String, sFecha_Salida As String, sSalario_Referencia As String, sActivo As String, _
       sDepartamento As String, sCentro As String, sOperacion As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True

'sFecha_Ingreso = Format(sFecha_Ingreso, "yyyymmdd")
'sFecha_Salida = Format(sFecha_Salida, "yyyymmdd")

  GSSQL = gsCompania & ".sgvUpdateglobalEmpleado '" & sOperacion & "','" & sEmpleado & "','" & sNombre & "' , '" & sFecha_Ingreso & "', '" & sFecha_Salida & "'," & sSalario_Referencia & ",'" & sActivo & "','" & sDepartamento & "','" & sCentro & "'"

    
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el Beneficiado."
      lbOk = False
    End If

secUpdateEmpleado = lbOk
Exit Function

error:
  lbOk = False
  Resume Next

End Function


Public Function secActualizaRoleUsuario(sIDModulo As String, sUsuario As String, sIDRole As String, sOperacion As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True

  GSSQL = gsCompania & ".secActualizaRoleUsuario " & sIDModulo & ",'" & sUsuario & "' , " & sIDRole & ", '" & sOperacion & "'"

    
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el Beneficiado."
      lbOk = False
    End If

secActualizaRoleUsuario = lbOk
Exit Function

error:
  lbOk = False
  Resume Next

End Function


Public Function setFechaPeriodo(sDia As String, sMes As String, dFechaSelected As Date) As Date
Dim dFechaResultado As Date

    dFechaResultado = CDate(Str(Year(dFechaSelected)) + "-" + sMes + "-" + sDia)
setFechaPeriodo = dFechaResultado
End Function


Public Function sgvGetCicloAbierto() As String
Dim lbOk As Boolean
Dim sResultado As String
On Error GoTo error
lbOk = True
  GSSQL = "SELECT " & gsCompania & "." & "sgvGetCicloAbierto() Periodo "
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbOk = False  'Indica que no es válido
    sResultado = "-1"
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    sResultado = Str(gRegistrosCmd("Periodo").value)
    lbOk = True
  End If
  sgvGetCicloAbierto = sResultado
  gRegistrosCmd.Close
  Exit Function
error:
  lbOk = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function

Public Function GetDatoFromTable(sFieldName As String, sTabla As String, sFiltro As String) As String
Dim sDato As String
On Error GoTo error
  sDescr = ""
  GSSQL = "SELECT top 1  " & sFieldName & _
          " FROM " & gsCompania & "." & sTabla & _
          " WHERE " & sFiltro 'Constuye la sentencia SQL
  
  
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    sDato = ""  'Indica que ocurrió un error
    gsOperacionError = "Error en la búsqueda del artículo !!!" & err.Description
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    sDato = gRegistrosCmd(sFieldName).value
  End If
  GetDatoFromTable = sDato
  gRegistrosCmd.Close
  Exit Function
error:
  sDato = ""
  SetMsgError "Ocurrió un error en la operación de búsqueda de la descripción ", err
  'gsOperacionError = "Ocurrió un error en la operación de búsqueda de la descripción " & err.Description
  Resume Next
End Function

Public Function GetrstAccionPropiedades(sAccion As String) As ADODB.Recordset
Dim sDato As String
Dim rst As ADODB.Recordset
Set rst = New ADODB.Recordset
rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
rst.CursorLocation = adUseClient ' Cursor local al cliente
rst.LockType = adLockOptimistic
On Error GoTo error

  GSSQL = " SELECT  Tipo_Accion, Descr, Prioridad, Factor, AfectaVacaciones, CalculoVacacional, EsAjuste, SaldoInicialPeriodo, SaldoSalidaEmpleado, SeRepiteMes, UnaVezPeriodo, UnaVezMes, ValorMaximoUno, Activo " & _
          " FROM " & gsCompania & "." & "sgvTipoAccion where Tipo_Accion='" & sAccion & "' AND Activo = 1 "

  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      Set rst = Nothing ' "Error en la búsqueda del artículo !!!" & err.Description
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
   Set rst = gRegistrosCmd
    
  End If
   
  Set GetrstAccionPropiedades = rst
   'gRegistrosCmd.Close
  Exit Function
error:
  Set rst = gRegistrosCmd
  gsOperacionError = "Ocurrió un error en la operación de búsqueda de la descripción " & err.Description
  Resume Next
End Function



Public Function Update_Password(sUsuario As String, sNewPassword As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
'sNewPassword = zEncrypt(sNewPassword)
  GSSQL = "UPDATE " & gsCompania & ".secUSUARIO SET Password= '" & sNewPassword & "'"
  GSSQL = GSSQL & " WHERE USUARIO ='" & sUsuario & "'"

    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbOk = False
    End If

Update_Password = lbOk
Exit Function

error:
  lbOk = False
  Resume Next

End Function

Public Function GetDescrCat(sfldCodCat As String, sValorCodigo As String, sTabla As String, sfldNameDescr As String, Optional bFiltroAdicional As Boolean = False, Optional sFiltroAdicional As String = "") As String
Dim sDescr As String
On Error GoTo error
  sDescr = ""
  GSSQL = "SELECT  " & sfldNameDescr & _
          " FROM " & gsCompania & "." & sTabla & _
          " WHERE " & sfldCodCat & " = " & sValorCodigo  'Constuye la sentencia SQL
  If bFiltroAdicional = True Then
    GSSQL = GSSQL & " AND " & sFiltroAdicional
  End If
    
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    sDescr = ""  'Indica que ocurrió un error
    gsOperacionError = "Error en la búsqueda del artículo !!!" & err.Description
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    sDescr = gRegistrosCmd(sfldNameDescr).value
  End If
  GetDescrCat = sDescr
  gRegistrosCmd.Close
  Exit Function
error:
  sDescr = ""
  gsOperacionError = "Ocurrió un error en la operación de búsqueda de la descripción " & err.Description
  Resume Next
End Function


Public Function GetDescrFromTable(sfldnameCodigo As String, bCodigoStr As Boolean, sfldNameDescr As String, sValueCodigo As String, sTabla As String, Optional bFiltroAdicional As Boolean = False, Optional sFiltroAdicional As String = "") As String
Dim sDescr As String
Dim sValor As String
On Error GoTo error
  sDescr = ""
  If bCodigoStr Then
    sValor = "'" & sValueCodigo & "'"
  Else
    sValor = sValueCodigo
  End If
  GSSQL = " SELECT " & sfldNameDescr & _
          " FROM " & gsCompania & "." & sTabla & _
          " WHERE " & sfldnameCodigo & " = " & sValor  'Constuye la sentencia SQL
  
  If bFiltroAdicional = True Then
    GSSQL = GSSQL & " AND " & sFiltroAdicional
  End If
  
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    sDescr = ""  'Indica que ocurrió un error
    gsOperacionError = "Error en la búsqueda del artículo !!!" & err.Description
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    sDescr = gRegistrosCmd(0).value
  End If
  GetDescrFromTable = sDescr
  gRegistrosCmd.Close
  Exit Function
error:
  sDescr = ""
  gsOperacionError = "Ocurrió un error en la operación de búsqueda de la descripción " & err.Description
  Resume Next
End Function


Public Function ExistCodeInTable(sfldnameCodigo As String, bCodigoStr As Boolean, sValueCodigo As String, sTabla As String, Optional bFiltroAdicional As Boolean = False, Optional sFiltroAdicional As String = "") As Boolean
Dim sDescr As String
Dim sValor As String
Dim lbOk As Boolean
On Error GoTo error
lbOk = False

  If bCodigoStr Then
    sValor = "'" & sValueCodigo & "'"
  Else
    sValor = sValueCodigo
  End If
  GSSQL = " SELECT " & sfldnameCodigo & _
          " FROM " & gsCompania & "." & sTabla & _
          " WHERE " & sfldnameCodigo & " = " & sValor  'Constuye la sentencia SQL
  
  If bFiltroAdicional = True Then
    GSSQL = GSSQL & " AND " & sFiltroAdicional
  End If
  
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    lbOk = False
    gsOperacionError = "Error en la búsqueda del artículo !!!" & err.Description
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbOk = True ' Si Existe el codigo
  End If
  ExistCodeInTable = lbOk
  gRegistrosCmd.Close
  Exit Function
error:
  gsOperacionError = "Ocurrió un error en la operación de búsqueda de la descripción " & err.Description
  Resume Next
End Function

Public Sub CargarDatos(xSet As ADODB.Recordset, tdbgData As TDBGrid, sNameFieldCode As String, sNameFieldDescr As String)
On Error GoTo EH
Dim i As Integer, j As Integer, bExiste As Boolean

    Set tdbgData.DataSource = Nothing
    With xSet
        .Find sNameFieldCode & " = null "
        For i = 1 To .Fields.Count - 1
            If Not LCase(.Fields(i).Name) = LCase(sNameFieldCode) Then
                If LCase(.Fields(i).Name) = LCase(sNameFieldDescr) Then
                    tdbgData.Columns(.Fields(i).Name).FooterAlignment = dbgRight
                    tdbgData.Columns(.Fields(i).Name).FooterText = " Total de elementos :    " & .Fields(i)
                Else

                    tdbgData.Columns(.Fields(i).Name).FooterText = Format(.Fields(i), "#,###,###.#0")

                End If
            End If
        Next
        .Delete
        .MoveFirst
    End With
    Set tdbgData.DataSource = xSet

    Exit Sub
EH:
    MsgBox err.Number & "-" & err.Source & "-" & err.Description
End Sub

Public Function getDayName(d As Date) As String
Dim sDia As String
sDia = ""
Select Case Weekday(d)
Case 1: sDia = "DOMINGO"
Case 2: sDia = "LUNES"
Case 3: sDia = "MARTES"
Case 4: sDia = "MIERCOLES"
Case 5: sDia = "JUEVES"
Case 6: sDia = "VIERNES"
Case 7: sDia = "SABADO"
End Select
getDayName = sDia
End Function


Public Function ExistePreparacion(sUsuario As String) As Boolean
Dim lbOk As Boolean
Dim bTienePreparacion As Boolean
On Error GoTo error
iesPeriodo = False
  GSSQL = "SELECT Usuario, Fecha from " & gsCompania & "." & "parmaPreparacionUsuario where Usuario = '" & sUsuario & "'"
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    bTienePreparacion = False  'Indica que no es válido
    
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    gdFechaCortePreparacion = gRegistrosCmd("Fecha").value
    bTienePreparacion = True
  End If
  ExistePreparacion = bTienePreparacion
  gRegistrosCmd.Close
  Exit Function
error:
  bTienePreparacion = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function


Public Function uspparmaUpdateCliente(sCliente As String, sSupervisor As String, sCanal As String, sStatus As String, sResponsable As String, sEsPreventa As String, sCDIS As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True


  GSSQL = gsCompania & ".uspparmaUpdateCliente '" & sCliente & "'," & sSupervisor & " , " & sCanal & ", " & sStatus & " , " & sResponsable & "," & sEsPreventa & "," & sCDIS

    
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Error al Actualizar el Cliente."
      lbOk = False
    End If

uspparmaUpdateCliente = lbOk
Exit Function

error:
  lbOk = False
  Resume Next

End Function


Public Function zEncrypt(sTexto As String) As String
Dim scadena As String
Dim sResultado As String
scadena = Trim(sTexto)
On Error GoTo error
gskey.Content = scadena
gskey.Algorithm.Name = CAPICOM_ENCRYPTION_ALGORITHM_3DES
gskey.SetSecret "1*&%$1290+-)(@!#}{]["
sResultado = gskey.Encrypt
zEncrypt = sResultado
Exit Function
error:
sResultado = ""
Resume Next
End Function

Public Function zDecrypt(sTexto As String) As String
Dim scadena As String
Dim sResultado As String
On Error GoTo error
scadena = Trim(sTexto)
gskey.Algorithm.Name = CAPICOM_ENCRYPTION_ALGORITHM_3DES
gskey.SetSecret "1*&%$1290+-)(@!#}{]["
gskey.Decrypt scadena
sResultado = gskey.Content
zDecrypt = sResultado
Exit Function
error:
sResultado = ""
Resume Next
End Function

Public Function uspparmaUpdateAplicacionesFromAsiento(sAsiento As String, sFecha As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
  GSSQL = gsCompania & "." & "uspparmaUpdateAplicacionesFromAsiento '" & sAsiento & "','" & sFecha & "'"
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbOk = False
    End If

uspparmaUpdateAplicacionesFromAsiento = lbOk
Exit Function

error:
  lbOk = False
  Resume Next

End Function

Public Function Prepara_DiferencialCC(sUsuario As String, sFecha As String, sTipoCambio As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
GSSQL = gsCompania & ".uspparmaPrepareDiferencialCC" & " '" & sUsuario & "' , '" & sFecha & "' , " & sTipoCambio
'gConet.BeginTrans

gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbOk = False
      'gConet.RollbackTrans
      Prepara_DiferencialCC = lbOk
      Exit Function
    End If

Prepara_DiferencialCC = lbOk
'gConet.CommitTrans
Exit Function

error:
  lbOk = False
  Resume Next

End Function

Public Function uspparmaInsertPrepDiferencialCC(sFechaUltProc As String, sFechaCorte As String, sTipoCambio As String, sPaquete As String, _
sDescrPaquete As String, sTipo As String, sDescrTipo As String, sUsuario As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
GSSQL = gsCompania & ".uspparmaInsertPrepDiferencialCC" & " '" & sFechaUltProc & "' , '" & sFechaCorte & "' , " & _
sTipoCambio & ",'" & sPaquete & "','" & sDescrPaquete & "','" & sTipo & "','" & sDescrTipo & "','" & sUsuario & "'"
'gConet.BeginTrans

gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbOk = False
      'gConet.RollbackTrans
      uspparmaInsertPrepDiferencialCC = lbOk
      Exit Function
    End If

uspparmaInsertPrepDiferencialCC = lbOk
'gConet.CommitTrans
Exit Function

error:
  lbOk = False
  Resume Next

End Function

Public Function Prepara_DiferencialCambiario(sModo As String, sFecha As String, sFechaCorte As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
GSSQL = gsCompania & ".uspparmaSetProcesoDifCamb" & "'" & sModo & "' , '" & sFecha & "','" & sFechaCorte & "'"
'gConet.BeginTrans

gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbOk = False
      'gConet.RollbackTrans
      Prepara_DiferencialCambiario = lbOk
      Exit Function
    End If

Prepara_DiferencialCambiario = lbOk
'gConet.CommitTrans
Exit Function

error:
  lbOk = False
  Resume Next

End Function

Public Function Prepara_DiferencialCambiarioCP(sModo As String, sFecha As String, sFechaCorte As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
GSSQL = gsCompania & ".uspparmaSetProcesoDifCambCP" & "'" & sModo & "' , '" & sFecha & "','" & sFechaCorte & "'"
'gConet.BeginTrans

gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbOk = False
      'gConet.RollbackTrans
      Prepara_DiferencialCambiarioCP = lbOk
      Exit Function
    End If

Prepara_DiferencialCambiarioCP = lbOk
'gConet.CommitTrans
Exit Function

error:
  lbOk = False
  Resume Next

End Function




Public Function parmaExistePreparacionCC() As Boolean
Dim lbOk As Boolean
Dim iResultado As Integer
On Error GoTo error
lbOk = False
  GSSQL = "SELECT " & gsCompania & "." & "parmagetExistePreparacionCC() Status"
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    iResultado = gRegistrosCmd("Status").value
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    iResultado = gRegistrosCmd("Status").value
  End If
  If iResultado = 1 Then lbOk = True Else lbOk = False
  parmaExistePreparacionCC = lbOk
  gRegistrosCmd.Close
  Exit Function
error:
  iResultado = 0
  lbOk = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function


Public Function parmaExistePreparacionCP() As Boolean
Dim lbOk As Boolean
Dim iResultado As Integer
On Error GoTo error
lbOk = False
  GSSQL = "SELECT " & gsCompania & "." & "parmagetExistePreparacionCP() Status"
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    iResultado = gRegistrosCmd("Status").value
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    iResultado = gRegistrosCmd("Status").value
  End If
  If iResultado = 1 Then lbOk = True Else lbOk = False
  parmaExistePreparacionCP = lbOk
  gRegistrosCmd.Close
  Exit Function
error:
  iResultado = 0
  lbOk = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function

Public Function uspparmaAplicaRecibosLiquidacion(sLiquidacion As String, sUsuario As String, sTotalPago As String, sFechaRecibo As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True
GSSQL = gsCompania & ".uspparmaAplicaRecibosLiquidacion " & "'" & sLiquidacion & "' , '" & sUsuario & "', " & sTotalPago & ", '" & sFechaRecibo & "'"
'gConet.BeginTrans

gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "aplicando"
      lbOk = False
      'gConet.RollbackTrans
      uspparmaAplicaRecibosLiquidacion = lbOk
      Exit Function
    End If

uspparmaAplicaRecibosLiquidacion = lbOk
'gConet.CommitTrans
Exit Function

error:
  lbOk = False
  Resume Next

End Function

Public Function ExisteUsuarioExactus(sUsuario As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error
lbOk = False
  GSSQL = "SELECT * from " & DBO & "." & "Usuario where Usuario ='" & sUsuario & "'"
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbOk = False
    
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbOk = True
  End If
  ExisteUsuarioExactus = lbOk
  gRegistrosCmd.Close
  Exit Function
error:
  lbOk = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function

Public Function LastDate(pFecha As Date) As Date
    LastDay = DateSerial(Year(pFecha), Month(pFecha) + 1, 1 - 1)
End Function

' Description: This function is called to export the specified report to an excel format under the specified file name.
'Public Sub ExportRptToExcel(ByVal rpt As Object, ByVal sFileName As String)
'Dim oExport As ActiveReportsExcelExport.ARExportExcel
'
'On Error Resume Next
'    Set oExport = CreateObject("ActiveReportsExcelExport.ARExportExcel")
'On Error GoTo 0
'
'    If oExport Is Nothing Then
'        'MsgBox "There was an error creating the Excel Export object.  Please reinstall and register the component.", vbCritical, "Oops..."
'        Exit Sub
'
'    End If
'
'    oExport.FileName = sFileName
'
'    oExport.Export rpt.Pages
'
'End Sub

Public Function Encrypt(Name As String, Key As Long) As String
Dim v As Long, c1 As String, z As String
For v = 1 To Len(Name)
c1 = Asc(Mid(Name, v, 1))
c1 = Chr(c1 + Key) ' private key here !
z = z & c1
Next v
Encrypt = z
End Function
Public Function Decrypt(Name As String, Key As Long) As String
Dim v As Long, c1 As String, z As String
For v = 1 To Len(Name)
c1 = Asc(Mid(Name, v, 1))
c1 = Chr(c1 - Key) ' private key here !
z = z & c1
Next v
Decrypt = z
End Function



Public Function parmaUpdateparmaDocumentsToApply(sOperation As String, sCliente As String, sDocumento As String, sFecha As String, _
sTipo As String, sValor As String, sIDFile As String, sflgAplicado As String) As Boolean
Dim lbOk As Boolean
Dim iResultado As Integer
Dim gRegistrosCmd As ADODB.Recordset
Dim sResultado As String

On Error GoTo error


lbOk = True
  GSSQL = gsCompania & ".[parmaUpdateparmaDocumentsToApply]  '" & sOperation & "','" & sCliente & "' ,'" & sDocumento & "','" & sFecha & "','" & _
    sTipo & "'," & sValor & "," & sIDFile & "," & sflgAplicado

  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbOk = False  'Indica que no es válido

  Else 'If Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido

    'sResultado = gRegistrosCmd("Resultado").value

    lbOk = True
  End If
  parmaUpdateparmaDocumentsToApply = lbOk
  gRegistrosCmd.Close
  Exit Function
error:
  lbOk = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next

End Function


Public Function ExisteCliente(sCliente As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error
lbOk = False
  GSSQL = "SELECT * from " & gsCompania & "." & "CLIENTE where CLIENTE = '" & sCliente & "'"
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbOk = False
    
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbOk = True
  End If
  ExisteCliente = lbOk
  gRegistrosCmd.Close
  Exit Function
error:
  lbOk = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function

Public Function ExisteLiquidacion(sLiquidacion As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error
lbOk = False
  GSSQL = "SELECT * from " & gsCompania & "." & "parmaAS400Liquidacion where idliquidacion = '" & sLiquidacion & "'"
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbOk = False
    
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbOk = True
  End If
  ExisteLiquidacion = lbOk
  gRegistrosCmd.Close
  Exit Function
error:
  lbOk = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function

Public Function parmaActualizaLiquidacionDesdeAS400(sIDLiquidacion As String, sFecha As String, sVendedor As String, sAux1 As String, sAux2 As String, sAux3 As String, _
sTotalLiquidacion As String, sTotalFaltante) As Boolean
Dim lbOk As Boolean
Dim iResultado As Integer
Dim gRegistrosCmd As ADODB.Recordset
Dim sResultado As String

On Error GoTo error


lbOk = True
  GSSQL = gsCompania & ".[parmaUpdateLiquidacionAS400]  'I','" & sIDLiquidacion & "','" & sFecha & "' ," & sTotalLiquidacion & "," & sTotalFaltante & ""
gConet.BeginTrans
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbOk = False  'Indica que no es válido
    GoTo error
  End If
    
    If sVendedor <> "" Then
        GSSQL = gsCompania & ".[parmaLiquidacionVendedorAS400]  'I','" & sIDLiquidacion & "','" & sVendedor & "','V'"
        
        Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia
            If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
              gsOperacionError = "No existe ese cliente." 'Asigna msg de error
              lbOk = False  'Indica que no es válido
              GoTo error
            End If
     End If

    If sAux1 <> "" Then
        GSSQL = gsCompania & ".[parmaLiquidacionVendedorAS400]  'I','" & sIDLiquidacion & "','" & sAux1 & "','A'"
        Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia
        If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
          gsOperacionError = "No existe ese cliente." 'Asigna msg de error
          lbOk = False  'Indica que no es válido
          GoTo error
        End If
    End If
    If sAux2 <> "" Then
     GSSQL = gsCompania & ".[parmaLiquidacionVendedorAS400]  'I','" & sIDLiquidacion & "','" & sAux2 & "','A'"
     Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia
        If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
           gsOperacionError = "No existe ese cliente." 'Asigna msg de error
           lbOk = False  'Indica que no es válido
           GoTo error
        End If
    End If
    If sAux3 <> "" Then
       GSSQL = gsCompania & ".[parmaLiquidacionVendedorAS400]  'I','" & sIDLiquidacion & "','" & sAux3 & "','A'"
       Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia
        If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
           gsOperacionError = "No existe ese cliente." 'Asigna msg de error
           lbOk = False  'Indica que no es válido
           GoTo error
        End If
    End If
                       
        
        If lbOk Then
            
            GSSQL = gsCompania & ".[parmaUpdateLiquidacionAS400]  'F','" & sIDLiquidacion & "'"
            Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia
            If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
               gsOperacionError = "No existe ese cliente." 'Asigna msg de error
               lbOk = False  'Indica que no es válido
               GoTo error
            End If
            
        End If

    If lbOk = False Then
        gConet.RollbackTrans
    Else
        gConet.CommitTrans
    End If
    
    
  
  parmaActualizaLiquidacionDesdeAS400 = lbOk
  
  'gRegistrosCmd.Close
  Exit Function
error:
  lbOk = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  gConet.RollbackTrans
   parmaActualizaLiquidacionDesdeAS400 = lbOk

End Function


Public Function ExisteNotasAnuladas(sLiquidacion As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error
lbOk = False
  GSSQL = "SELECT * from " & gsCompania & "." & "Documentos_CC where Tipo = 'N/D' AND Documento like '%" & Mid(sLiquidacion, 1, 8) & "%' and Anulado='S' AND CLIENTE IN " & _
  " (SELECT CLIENTE FROM " & gsCompania & ".parmaAS400LiquidacionDetFaltante WHERE IDLIQUIDACION = '" & sLiquidacion & "')"
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbOk = False
    
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbOk = True
  End If
  ExisteNotasAnuladas = lbOk
  gRegistrosCmd.Close
  Exit Function
error:
  lbOk = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function


Public Function SaldosNotasMenorQueFaltante(sLiquidacion As String, dFaltante As Double) As Boolean
Dim lbOk As Boolean
On Error GoTo error
lbOk = False
  GSSQL = "SELECT Sum(Saldo) Saldo from " & gsCompania & "." & "Documentos_CC where Tipo = 'N/D' AND Documento like '%" & Mid(sLiquidacion, 1, 8) & "%' and Anulado='N'" & " AND CLIENTE IN " & _
  " (SELECT CLIENTE FROM " & gsCompania & ".parmaAS400LiquidacionDetFaltante WHERE IDLIQUIDACION = '" & sLiquidacion & "')"
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbOk = False
    
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    If Round(gRegistrosCmd(0).value, 2) < dFaltante Then
        lbOk = True
    End If
  End If
  SaldosNotasMenorQueFaltante = lbOk
  gRegistrosCmd.Close
  Exit Function
error:
  lbOk = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function


Public Function CambiaStatusLiquidacion(sLiquidacion As String, sStatus As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error
lbOk = False
  GSSQL = "Update " & gsCompania & "." & "parmaAS400Liquidacion set statusCobro = '" & sStatus & "' where idliquidacion = '" & sLiquidacion & "'"
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbOk = False
    
  Else
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbOk = True
  End If
  CambiaStatusLiquidacion = lbOk
  'gRegistrosCmd.Close
  Exit Function
error:
  lbOk = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function


Public Function spGlobalUpdateCatalogo(sOperacion As String, sIDTable As String, sDescr As String, sActivo As String, sUsaValor As String, sNombreValor As String, sValor As String, Optional sIDCatalogo As String) As Boolean
Dim lbOk As Boolean
Dim sIDCat As String
On Error GoTo error

lbOk = True
    If sIDCatalogo = "" Then
        sIDCat = "ND"
    Else
        sIDCat = sIDCatalogo
    End If
    

  GSSQL = gsCompania & ".spGlobalUpdateCatalogo '" & sOperacion & "','" & sIDCat & "'," & sIDTable & ",0,'" & sDescr & "'," & sActivo & "," & sUsaValor & ",'" & sNombreValor & "'," & sValor

    
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      'gsOperacionError = "Eliminando el Registro."
      SetMsgError "Eliminando el Registro.", err
      lbOk = False
    End If

spGlobalUpdateCatalogo = lbOk
Exit Function

error:
  lbOk = False
  Resume Next

End Function


Public Function getPositionRecord(rstFuente As ADODB.Recordset, sFiltro As String) As Variant
Dim lbmk As Variant
Dim rstClone As ADODB.Recordset
Dim bmPos As Variant
lbmk = -1
If Not (rstFuente.EOF And rstFuente.BOF) Then
    Set rstClone = New ADODB.Recordset
        bmPos = rstFuente.Bookmark
        rstClone.Filter = adFilterNone
        Set rstClone = rstFuente.Clone
        rstClone.Filter = sFiltro
        If Not rstClone.EOF Then ' Si existe
          lbmk = rstClone.Bookmark
        End If
        rstFuente.Filter = adFilterNone
        rstFuente.Bookmark = bmPos
    rstClone.Filter = adFilterNone
End If
getPositionRecord = lbmk
End Function

Public Function LoadAccess(sUsuario As String, sPassword As String, sModulo As String)
On Error GoTo errores
Dim lbOk As Boolean
Dim lsFilename As String
Dim rst As ADODB.Recordset
Dim rstCSV As ADODB.Recordset
lbOk = False
Set rst = New ADODB.Recordset
If rst.State = adStateOpen Then rst.Close
rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
rst.CursorLocation = adUseClient ' Cursor local al cliente
rst.LockType = adLockOptimistic
GSSQL = "SELECT IDMODULO, USUARIO, IDROLE, IDACCION " & _
          " FROM " & gsCompania & "." & " vsecPrivilegios "
GSSQL = GSSQL & " WHERE USUARIO = '" & sUsuario & "'"
GSSQL = GSSQL & " AND IDMODULO = " & sModulo

If rst.State = adStateOpen Then rst.Close
Set rst = GetRecordset(GSSQL)

If Not (rst.EOF And rst.BOF) Then
    Set grRecordsetAcceso = rst
    lbOk = True
End If
LoadAccess = lbOk
Exit Function
errores:
lbOk = False
LoadAccess = lbOk
End Function

Public Function UserCouldIN(sUsuario As String, sPassword As String)
On Error GoTo errores
Dim lbOk As Boolean
Dim rst As ADODB.Recordset
lbOk = False
Set rst = New ADODB.Recordset
If rst.State = adStateOpen Then rst.Close
rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
rst.CursorLocation = adUseClient ' Cursor local al cliente
rst.LockType = adLockOptimistic
'sPassword = zEncrypt(sPassword)

GSSQL = "SELECT USUARIO, PASSWORD " & _
          " FROM " & gsCompania & "." & "secUSUARIO WHERE USUARIO = '" & sUsuario & "'" & "  AND PASSWORD ='" & sPassword & "'"
          
If rst.State = adStateOpen Then rst.Close
rst.Open GSSQL
If Not (rst.EOF And rst.BOF) Then
    Set grRecordsetAcceso = rst
    lbOk = True
'    If zDecrypt(rst("Password").value) = sPassword Then
'        Set grRecordsetAcceso = rst
'        lbok = True
'    Else
'        lbok = False
'    End If
End If
UserCouldIN = lbOk
Exit Function
errores:
lbOk = False
UserCouldIN = lbOk
End Function

Public Function fafUpdateVendedor(sOperacion As String, sIDVendedor As String, sNombre As String, sTipo As String, sActivo As String) As Boolean
Dim lbOk As Boolean
On Error GoTo error

lbOk = True

  GSSQL = gsCompania & ".fafUpdateVendedor '" & sOperacion & "'," & sIDVendedor & ",'" & sNombre & "'," & sActivo & ",'" & sTipo & "'"

    
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
       'gsOperacionError = "Eliminando el Beneficiado."
       SetMsgError "Eliminando el Vendedor.", err
      lbOk = False
    End If

fafUpdateVendedor = lbOk
Exit Sub

error:
  lbOk = False
  Resume Next

End Sub



Public Function ExiteItem(rst As ADODB.Recordset, sCondicion As String) As Boolean
Dim lbOk As Boolean
lbOk = False
If Not rst.EOF Then
  rst.MoveFirst
  rst.Find sCondicion, 0, adSearchForward, 0
  If Not rst.EOF Then
    lbOk = True
  End If
End If
ExiteItem = lbOk
End Function


Public Sub SetMsgError(sError As String, oError As error)
    gsOperacionError = sError & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & oError.Description
End Sub


Public Function getValueFieldsFromTable(sTabla As String, sListFieldName As String, sFiltro As String, ByRef dicResult As Dictionary) As String
    Dim lbOk As Boolean
    Dim sResultado As String
    Dim sOrden As String
    Dim rst As ADODB.Recordset
    Dim i As Integer
    
    Set dicResult = New Dictionary
    On Error GoTo error
    lbOk = False
    sResultado = "ND"
    
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst.CursorType = adOpenKeyset  'Asigna un cursor estático
    rst.CursorLocation = adUseClient ' Cursor local al cliente
    rst.LockType = adLockOptimistic
    
    lbOk = True
      
    'Armar el listado de Campos
      
    Dim arrFields() As String
    arrFields = Split(sListFieldName, ",")
    
    GSSQL = "SELECT TOP 1 "
    
    For i = 0 To UBound(arrFields)
         GSSQL = GSSQL & arrFields(i) & ","
    Next i
    
    GSSQL = Mid(GSSQL, 1, Len(GSSQL) - 1)
    GSSQL = GSSQL & " FROM " & gsCompania & "." & sTabla          'Constuye la sentencia SQL
    GSSQL = GSSQL & " WHERE " & sFiltro
    
    If rst.State = adStateOpen Then rst.Close
    rst.Open GSSQL, gConet, adOpenDynamic, adLockOptimistic, adCmdText    'Ejecuta la sentencia
    
    If Not (rst.BOF And rst.EOF) Then
        For i = 0 To UBound(arrFields)
            dicResult.Add arrFields(i), rst(arrFields(i)).value
        Next i
        lbOk = True
    End If
    
    getValueFieldFromTable = lbOk
    Set rst = Nothing
    Exit Function
error:
      lbOk = False
      getValueFieldFromTable = lbOk
    
End Function




