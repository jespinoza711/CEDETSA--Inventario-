Attribute VB_Name = "Unit"
Public Const ICO_INFORMACION = 4
Public Const ICO_PREGUNTA = 3
Public Const ICO_ADVERTENCIA = 1
Public Const ICO_ERROR = 2
Public Const ICO_OK = 5
Public Const CONTROL_MARGIN = 300
'Constantes para Modulo
Public Const C_MODULO = 1000
Public Const A_CATALOGOS = 1
Public Const A_PRODUCTOS = 2

'Seccion del MDI

Public Type MENU_BUTTONS_INFO
  form_name As String
  form_buttons As String
End Type

Public ButtonsAvailable() As MENU_BUTTONS_INFO


Public Declare Function GetCursorPos Lib "user32" _
            (lpPoint As POINTAPI) As Long
            
Public Type POINTAPI
        x As Long
        Y As Long
End Type



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
Public lbok As Boolean

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


    
    Dim lbok As Boolean
    
    gbEquipo64Bits = Is64bit()
    If Not LeerConfiguracion = True Then
      lbok = Mensaje("Error al leer el archivo de configuración ", ICO_ERROR, False)
      End
    End If
    
    If ConexionBD Then
        If (Command$ <> "") Then
            LoginForCommandLine
        Else
        
            Dim frm As frmLogin
            Set frm = New frmLogin
            frm.Show vbModal
            lbok = frm.LoginSucceeded
            If lbok Then
              Unload frm
              MDIMain.Show
              'frmPrepare.Show
              'frmAnalisisVencimiento.Show
            Else
                Unload frm
                End
            End If
        End If
    Else
      lbok = Mensaje("No pudo conectarse a la base de datos :" & gsOperacionError, ICO_ERROR, False)
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
            
    sArgs = Split(Command$, " ")
    
    sUsuario = sArgs(0)
    sPass = sArgs(1)
    
   
    If UserCouldIN(sUsuario, sPass) Then
        gsUSUARIO = sUsuario
    
        lbok = LoadAccess(sUsuario, sPass, C_MODULO)
        If Not lbok Then
            lbok = Mensaje("No se pudieron cargar los accesos del usuario ", ICO_ERROR, False)
            End
        Else
        'lbok = CargaParametros()
            If Not lbok Then
                lbok = Mensaje("No se ha configurado el Sistema... los parametros no se han definido ", ICO_ERROR, False)
                End
            End If
        End If
    Else
        lbok = Mensaje("Login o Password incorrectos...", ICO_ERROR, False)
        End
    End If
    
    
    If lbok Then
        MDIMain.Show
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
Dim lbok As Boolean
On Error GoTo error

lbok = True
  GSSQL = gsCompania & "." & "uspparmaUpdateParametros '" & sNombreEmpresa & "'"
  
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = " ."
      lbok = False
        
    End If

Update_Parametros = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function

Public Function Update_UltimoDiaCerrado(sUltAnioCerrado As String, sUltMesCerrado As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True
  GSSQL = gsCompania & "." & "sgvUpdateUltimoMesCerrado " & sUltAnioCerrado & "," & sUltMesCerrado
  
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = " ."
      lbok = False
        
    End If

Update_UltimoDiaCerrado = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function

Public Function Update_AnulaEsquela(sEsquela As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True
  GSSQL = "Update " & gsCompania & "." & "sgvEsquela set Anulada=1 where Consecutivo='" & sEsquela & "'"
  
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = " ."
      lbok = False
        
    End If

Update_AnulaEsquela = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function



Public Function FechaEnMesAbierto(dFecha As Date) As Boolean
Dim lbok As Boolean
Dim dUltimoMesCerrado As Date
Dim dFirstDayAllow As Date
lbok = False

dUltimoMesCerrado = UltimoMesCerradoFirstDay
dFirstDayAllow = DateAdd("m", 1, dUltimoMesCerrado)
If dFecha >= dFirstDayAllow Then
    lbok = True
End If
FechaEnMesAbierto = lbok
End Function



Private Function LimpiaNulos(ByVal psParam As String) As String   'Toma una hilera y elimina sus nulos
  Dim lnCont As Integer 'Es el contador del ciclo
  Dim lnTam As Integer  'Mantiene el tamaño de la hilera
  Dim lsHilera As String  'String de trabajo
  
  lnTam = CInt(Len(psParam)) 'Obtiene el tamaño de la hilera
  lnCont = InStr(1, psParam, Chr(0), 1) - 1
  If lnCont <> -1 Then
  lsHilera = left(psParam, lnCont) & Space(lnTam - lnCont)  'Asigna espacios al resto
  Else
  lsHilera = psParam
  End If
  
  LimpiaNulos = lsHilera
End Function

Public Sub ValidaLargo( _
  ByVal psTexto As String, _
  ByRef KeyAscii As Integer, _
  ByVal pnLargo As Integer)
  Dim lbok As Boolean
On Error GoTo err
  
  If (pnLargo <= Len(psTexto)) Then 'Si se alcanzó el límite
    If (KeyAscii <> 8) Then
      KeyAscii = 0
    End If
  End If
  Exit Sub

err:
    lbok = Mensaje("Función: ValidaLargo ..." & err.Description, ICO_ERROR, False)
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
      frmMensajeError.btnSi.left = 2850
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
  Dim lbok As Boolean 'Indica que el proceso está bien
  Dim lsConexion As String  'Va a almacenar la hilera de conexión a la base de datos
  On Error GoTo error:  'Para capturar el evento de error
  
  Set gConet = New ADODB.Connection 'Inicializa la variable de conexión
  lbok = True
    lsConexion = GetConectString(gsTypeDB, gsNombreServidor, gsNombreBaseDatos, gsUser, gsPassword)
 '   gConet.Open "Provider=SQLNCLI.1;Password=admin;Persist Security Info=True;User ID=SysUser;Initial Catalog=exactus;Data Source=serverg4s"
    gConet.Open lsConexion 'Realiza la conexión a la BD
    gConet.CommandTimeout = 300
    'gConet.Properties("Multiple Connections").value = False
    If (gConet.Errors.Count > 0) Then  'Si no hubo conexión
      lbok = False  'No se cuenta con una conexión
    End If
    gsConetstr = lsConexion
    ConexionBD = lbok 'Asigna el valor de retorno
  Exit Function
  
error:
    lbok = False
    gsOperacionError = "constr :" & lsConexion & err.Description
    ConexionBD = lbok 'Asigna el valor de retorno
End Function



Public Function ConnectToDBOpen() As Boolean
    If (gConet.State = adStateOpen) Then
        gConet.Close
    End If
    gConet.CursorLocation = adUseClient
    gConet.Open lsConexion
     gConet.CommandTimeout = 300
    gConet.Properties("Multiple Connections").value = False
    ConnectToDBOpen = True
End Function

Public Function DesconectaBD() As Boolean
Dim lbok As Boolean
On Error GoTo error
If gConet.State = adStateOpen Then
  lbok = True
  gConet.Close
End If
DesconectaBD = lbok
Exit Function
error:
  lbok = False
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
  Dim lbok As Boolean

  On Error GoTo error
  lbok = True
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
  

  LeerConfiguracion = lbok
  Exit Function

error:  'En caso de haber un error
  lbok = False
  LeerConfiguracion = lbok
End Function

Public Sub CargaNominas(rst As ADODB.Recordset, sTipo As String, cbo As ComboBox)
Dim lbok As Boolean
'On Error GoTo error
lbok = True
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
Dim lbok As Boolean
Dim strTemporal As String
Dim strResult As String
lbok = False
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
      While Not lbok And i <= Len(strFuente)
        If InStr(strTemporal, sChar) <> 0 Then
           i = InStr(strTemporal, sChar) ' posición del caracter
           strResult = strResult & Mid(strTemporal, 1, i - 1) & sNewChar '& Mid(strTemporal, i + 1)
           strTemporal = Mid(strTemporal, i + 1, Len(strFuente))
           i = i + 1
        Else
           strResult = strResult & strTemporal
           lbok = True
        End If
      Wend
    SustituyeChar = strResult
Else
  SustituyeChar = ""
End If
End Function

Public Function OnlythisChar(strFuente As String, sChar As String) As Boolean
Dim i As Integer
Dim lbok As Boolean
Dim stmpChar As String
stmpChar = ""
lbok = True
i = 1
While i <= Len(strFuente) And lbok
stmpChar = Mid(strFuente, i, 1)
If stmpChar <> sChar Then
    lbok = False
End If
i = i + 1
Wend
OnlythisChar = lbok
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
Dim lbok As Boolean
lbok = True
If Trim(txtbox.Text) = "" Then

  lbok = False
  Val_TextboxNum = lbok
  Exit Function
End If


If Not IsNumeric(txtbox.Text) Then
  
  lbok = False
  Val_TextboxNum = lbok

End If

Val_TextboxNum = lbok

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
Dim lbok As Boolean
lbok = False
On Error GoTo errores
Dim objExcel
If iAbrir > 0 Then
Set objExcel = CreateObject("Excel.Application")
  objExcel.Visible = True
  objExcel.Workbooks.Open sNombre
  lbok = True
End If
  OpenExcel = lbok
Exit Function
errores:
  MsgBox "Error al abrir la aplicación Excel", vbOKOnly, "Error "
  lbok = False
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
Dim lbok As Boolean
Dim rstClone As ADODB.Recordset
Dim bmPos As Variant
lbok = False
If Not (rstFuente.EOF And rstFuente.BOF) Then
    Set rstClone = New ADODB.Recordset
        bmPos = rstFuente.Bookmark
        rstClone.Filter = adFilterNone
        Set rstClone = rstFuente.Clone
        rstClone.Filter = sFiltro
        If Not rstClone.EOF Then ' Si existe
          lbok = True
        End If
        rstFuente.Filter = adFilterNone
        rstFuente.Bookmark = bmPos
    rstClone.Filter = adFilterNone
End If
ExiteRstKey = lbok
End Function



Public Function UserMayAccess(sUsuario As String, iAccion As Integer, iModulo As Integer) As Boolean
Dim lbok As Boolean
On Error GoTo errores
Dim sCriterioAcceso As String
sCriterioAcceso = "USUARIO ='" & UCase(sUsuario) & "' AND IDACCION=" & Str(iAccion) & " AND IDMODULO=" & Str(iModulo)
If ExiteRstKey(grRecordsetAcceso, sCriterioAcceso) Then
    lbok = True
Else
    lbok = False
End If
UserMayAccess = lbok
Exit Function
errores:
lbok = False
UserMayAccess = lbok
End Function

' carga parametros de la preparacion diferencial cambiario CC
Public Function CargaParametrosPreparacionDifCC() As Boolean
Dim lbok As Boolean
On Error GoTo error
lbok = True
  GSSQL = "SELECT top 1 FechaCorte " & _
          " FROM " & gsCompania & "." & "parmaPreparacionProcesoDiferencial "

  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbok = False  'Indica que no es válido

  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    gParametrosPrepDifCC.FechaCorte = gRegistrosCmd("FechaCorte").value
    
    lbok = True
  End If
  CargaParametrosPreparacionDifCC = lbok
  gRegistrosCmd.Close
  Exit Function
error:
  lbok = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next

End Function

Public Function CargaParametrosPreparacionDifCP() As Boolean
Dim lbok As Boolean
On Error GoTo error
lbok = True
  GSSQL = "SELECT top 1 FechaCorte " & _
          " FROM " & gsCompania & "." & "parmaPreparacionProcesoDiferencialCP "

  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbok = False  'Indica que no es válido

  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    gParametrosPrepDifCC.FechaCorte = gRegistrosCmd("FechaCorte").value
    
    lbok = True
  End If
  CargaParametrosPreparacionDifCP = lbok
  gRegistrosCmd.Close
  Exit Function
error:
  lbok = False
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
Dim lbok As Boolean
On Error GoTo error

lbok = True
  GSSQL = gsCompania & "." & "repActualizaTipoCambioContrato " & sTipoCambio
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbok = False
    End If

Update_TipoCambio = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function


Public Function Cierra_Periodo(sEmpleado As String, sPeriodo As String, sFecha As String, sPeriodoNuevo As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True
GSSQL = gsCompania & ".sgvGetDetalleMovimientosEmpleado" & " '" & sEmpleado & "', '" & sPeriodo & "' , '" & sFecha & "', 1, 1, " & sPeriodoNuevo
gConet.BeginTrans
gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbok = False
      gConet.RollbackTrans
      Cierra_Periodo = lbok
      Exit Function
    End If

Cierra_Periodo = lbok
gConet.CommitTrans
Exit Function

error:
  lbok = False
  Resume Next

End Function

Public Function Baja_Empleado(sEmpleado As String, sPeriodo As String, sFecha As String, sComentario As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True
GSSQL = gsCompania & ".sgvGetDetalleMovimientosEmpleado" & " '" & sEmpleado & "', " & sPeriodo & " , '" & sFecha & "', 1, null, null, 1,0, '" & gsUSUARIO & "','S', '" & sComentario & "'"
gConet.BeginTrans
gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbok = False
      gConet.RollbackTrans
      Baja_Empleado = lbok
      Exit Function
    End If

Baja_Empleado = lbok
gConet.CommitTrans
Exit Function

error:
  lbok = False
  Resume Next

End Function

Public Function Prepara_AnalisisAntiguedad(sUsuario As String, sFecha As String, sCodCliente As String, sCodCategoria As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True
GSSQL = gsCompania & ".uspparmaPrepareAnalisisVencimiento" & " '" & sUsuario & "' , '" & sFecha & "' , '" & sCodCliente & "','" & sCodCategoria & "'"
'gConet.BeginTrans

gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbok = False
      'gConet.RollbackTrans
      Prepara_AnalisisAntiguedad = lbok
      Exit Function
    End If

Prepara_AnalisisAntiguedad = lbok
'gConet.CommitTrans
Exit Function

error:
  lbok = False
  Resume Next

End Function

Public Function uspparmaGetClasificacionporPago(sCodCliente As String, sFechaInicio As String, sFechaFin As String, sCodCategoria As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True
GSSQL = gsCompania & ".uspparmaGetClasificacionporPago" & " '" & sCodCliente & "' , '" & sFechaInicio & "' , '" & sFechaFin & "','" & sCodCategoria & "'"
'gConet.BeginTrans

gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbok = False
      'gConet.RollbackTrans
      uspparmaGetClasificacionporPago = lbok
      Exit Function
    End If

uspparmaGetClasificacionporPago = lbok
'gConet.CommitTrans
Exit Function

error:
  lbok = False
  Resume Next

End Function



Public Function Prepara_Reportes(sPeriodo As String, sFecha As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True
GSSQL = gsCompania & ".sgvGetDetalleMovimientosEmpleado" & " '*', '" & sPeriodo & "' , '" & sFecha & "', null, null,null,null, 1,'" & gsUSUARIO & "'"
gConet.BeginTrans
gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbok = False
      gConet.RollbackTrans
      Prepara_Reportes = lbok
      Exit Function
    End If

Prepara_Reportes = lbok
gConet.CommitTrans
Exit Function

error:
  lbok = False
  Resume Next

End Function



' Carga Parámetros del sistema
Public Function CargaParametros() As Boolean
Dim lbok As Boolean
On Error GoTo error
lbok = True
  GSSQL = "SELECT * " & _
          " FROM " & gsCompania & "." & "INVPARAMETROS "
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbok = False  'Indica que no es válido
    
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    
    gparametros.NombreEmpresa = gRegistrosCmd("NombreEmpresa").value

    lbok = True
  End If
  CargaParametros = lbok
  gRegistrosCmd.Close
  Exit Function
error:
  lbok = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function



' devuelve Proximo Consecutivo de la esquela
Public Function getNextConsecEsquela() As String
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
Dim lbok As Boolean
On Error GoTo error
lbok = False
  GSSQL = "SELECT * from " & gsCompania & "." & "sgvAlertas where " & sInicio & " between Inicio and Fin or " & sFin & " between inicio and fin "
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbok = False
    
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbok = True
  End If
  sgvRangoExistente = lbok
  gRegistrosCmd.Close
  Exit Function
error:
  lbok = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function


Public Function sgvEMPLEADOBAJA(sEmpleado As String) As String

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
Dim lbok As Boolean
Dim rst As ADODB.Recordset
Dim sFiltro As String

Dim i As Integer
Dim lsRole As String
Dim lbHayRegistros As Boolean
lbok = False
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
            lbok = Insert_UsuarioRole(sModulo, lsRole, sUsuario)
    
    End If
    If Not lst.Selected(i) And lbHayRegistros Then
        If ExiteRstKey(rst, sFiltro) Then
        ' Eliminar el Registro
            lbok = Delete_UsuarioRole(sModulo, lsRole, sUsuario)
        End If
    End If
Next i

'End If


Set rst = Nothing
UpdateUsuarioModuloRole = lbok
Exit Function
errores:
lbok = False
UpdateUsuarioModuloRole = lbok

End Function

Public Function Insert_UsuarioRole(sModulo As String, sRole As String, sUsuario As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True
  GSSQL = "INSERT  " & gsCompania & "." & "secUSUARIOROLE (IDMODULO, IDROLE, USUARIO)"
  GSSQL = GSSQL & " VALUES (" & sModulo & " , " & sRole & " , '" & sUsuario & "')"

    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbok = False
    End If

Insert_UsuarioRole = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function


Public Function LoadModuloRole(sModulo As String, lst As ListBox) As Boolean
On Error GoTo errores
Dim lbok As Boolean
Dim sItem As String
Dim rst As ADODB.Recordset

lbok = False
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
LoadModuloRole = lbok
Exit Function
errores:
lbok = False
LoadModuloRole = lbok
End Function

Public Sub UncheckListBox(lst As ListBox)
Dim i As Integer
    For i = 0 To lst.ListCount - 1
            lst.Selected(i) = False
    Next i
End Sub




Public Function LoadRoleModulo(sUsuario As String, sModulo As String, lst As ListBox) As Boolean
On Error GoTo errores
Dim lbok As Boolean
Dim rst As ADODB.Recordset
Dim sFiltro As String
Dim i As Integer
lbok = True
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
LoadRoleModulo = lbok
Exit Function
errores:
lbok = False
LoadRoleModulo = lbok
End Function

Public Function GetCodeBeforeMinus(gsDato As String) As String
Dim i As Integer
Dim sResult As String
sResult = ""
i = InStr(1, gsDato, "-", vbTextCompare)
If i > 0 Then
    sResult = Trim(left(gsDato, i - 1))
End If
GetCodeBeforeMinus = sResult
End Function

Public Function GetCodeBeforevbTab(gsDato As String) As String
Dim i As Integer
Dim sResult As String
sResult = ""
i = InStr(1, gsDato, vbTab, vbTextCompare)
If i > 0 Then
    sResult = Trim(left(gsDato, i - 1))
End If
GetCodeBeforevbTab = sResult
End Function


Public Function LoadAccionRole(sModulo As String, sRole As String, lst As ListBox) As Boolean
On Error GoTo errores
Dim lbok As Boolean
Dim sItem As String
Dim rst As ADODB.Recordset

lbok = False
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
LoadAccionRole = lbok
Exit Function
errores:
lbok = False
LoadAccionRole = lbok
End Function

Public Function Delete_UsuarioRole(sModulo As String, sRole As String, sUsuario As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True
  GSSQL = "DELETE FROM " & gsCompania & "." & "secUSUARIOROLE "
  GSSQL = GSSQL & " WHERE IDMODULO =" & sModulo & " AND IDROLE = " & sRole & " AND USUARIO = " & "'" & sUsuario & "'"

    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto el UsuarioRole."
      lbok = False
    End If

Delete_UsuarioRole = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function

Public Function sgvActualizaSaldoEmpleado(sOperacion As String, sPeriodo As String, sTipo_Accion As String, sEmpleado As String, _
    sFecha As String, sCantDias As String, sComentario As String, sUsuario As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True
  GSSQL = gsCompania & ".sgvActualizaSaldoEmpleado '" & sOperacion & "'," & sPeriodo & ", '" & sTipo_Accion & "','" & sEmpleado & "', '" & _
  sFecha & "', " & sCantDias & ", '" & sComentario & "','" & sUsuario & "'"

    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el Beneficiado."
      lbok = False
    End If

sgvActualizaSaldoEmpleado = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function

Public Function uspsgvSetBajaEmpleado(sEmpleado As String, sPeriodo As String, _
    sFecha As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True
  GSSQL = gsCompania & ".uspsgvSetBajaEmpleado '" & sEmpleado & "'," & sPeriodo & ", '" & _
  sFecha & "'"

    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el Beneficiado."
      lbok = False
    End If

uspsgvSetBajaEmpleado = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function


Public Function sgvActualizaTipoAccion(sTipo_Accion As String, sDescr As String, sPrioridad As String, sFactor As String, _
    sflgRRHH As String, sAfectaVacaciones As String, sCalculoVacacional As String, sEsAjuste As String, sEsSaldoInicialPeriodo As String, sEsSaldoSalidaEmpleado As String, _
    sActivo As String, sSeRepiteMes As String, sUnaVezPeriodo As String, sUnaVezMes As String, sMaximoUno As String, sOperacion As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True
  GSSQL = gsCompania & ".sgvActualizaTipoAccion '" & sTipo_Accion & "','" & sDescr & "' , " & sPrioridad & ", " & sFactor & ", " & sflgRRHH & "," & sAfectaVacaciones & "," _
  & sCalculoVacacional & "," & sEsAjuste & "," & sEsSaldoInicialPeriodo & "," & sEsSaldoSalidaEmpleado & "," & sActivo & "," & sSeRepiteMes & "," & sUnaVezPeriodo & "," & sUnaVezMes & "," & sMaximoUno & ",'" & sOperacion & "'"

    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el Beneficiado."
      lbok = False
    End If

sgvActualizaTipoAccion = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function

Public Function sgvUpdatesgvHeaderEsquelas(sOperacion As String, sEmpleado As String, sAnio As String, sMes As String, sConsecutivo As String, sFechaCreacion As String, sFechaSaldo As String, sSaldo As String, sDiasSolicitados As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True
sFechaCreacion = Format(sFechaCreacion, "YYYYmmdd")
sFechaSaldo = Format(sFechaSaldo, "YYYYmmdd")

  GSSQL = gsCompania & ".sgvUpdatesgvHeaderEsquelas '" & sOperacion & "','" & sEmpleado & "'," & sAnio & "," & sMes & ",'" & sConsecutivo & "' , '" & sFechaCreacion & "', '" & sFechaSaldo & "', " & sSaldo & "," & sDiasSolicitados

    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el Beneficiado."
      lbok = False
    End If

sgvUpdatesgvHeaderEsquelas = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function


Public Function sgvUpdatesgvEsqueladetalle(sOperacion As String, sEmpleado As String, sAnio As String, sMes As String, sConsecutivo As String, sDiaSolicitado As String, sCantidad As String, sComentario As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True
sDiaSolicitado = Format(sDiaSolicitado, "YYYYmmdd")


  GSSQL = gsCompania & ".sgvUpdatesgvEsqueladetalle '" & sOperacion & "','" & sEmpleado & "'," & sAnio & "," & sMes & ",'" & sConsecutivo & "' , '" & sDiaSolicitado & "'," & sCantidad & ",'" & sComentario & "'"

    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el Beneficiado."
      lbok = False
    End If

sgvUpdatesgvEsqueladetalle = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function


Public Function invUpdateBodega(sOperacion As String, sIDBodega As String, sDescrBodega As String, sActivo As String, sFactura As String, sPreFactura As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True

  GSSQL = gsCompania & ".invUpdateBodega '" & sOperacion & "'," & sIDBodega & ",'" & sDescrBodega & "'," & sActivo & "," & sFactura & ",'" & sPreFactura & "'"

    
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
        'gsOperacionError = "Eliminando el Beneficiado."
        SetMsgError "Ocurrió un error actualizando La bodega. ", err
      lbok = False
    End If

invUpdateBodega = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function



Public Function invUpdateLote(sOperacion As String, sIDLote As String, sLoteInterno As String, sLoteProveedor As String, sFechaVencimiento As String, sFechaProduccion) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True

  GSSQL = gsCompania & ".invUpdateLote '" & sOperacion & "'," & sIDLote & ",'" & sLoteInterno & "','" & sLoteProveedor & "','" & sFechaVencimiento & "','" & sFechaProduccion & "'"

    
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
        'gsOperacionError = "Eliminando el Beneficiado."
        SetMsgError "Ocurrió un error actualizando el Lote. ", err
      lbok = False
    End If

invUpdateLote = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function

Public Function ccUpdateCliente(sOperacion As String, sIDCliente As String, sNombre As String, sRazonSocial As String, sDireccion As String, sTelefono1 As String, sTelefono2 As String, sTelefono3 As String, _
                sCelular1 As String, sCelular2 As String, semail As String, sEsFarmacia As String, sNombreFarmacia As String, sRUC As String, _
                sPropietario As String, sIDBodega As String, sIDPlazo As String, sIDMoneda As String, sIDCategoria As String, sIDDepartamento As String, _
                sIDMunicipio As String, sIDZona As String, sIDVendedor As String, _
                sTechoCredito As String, sActivo As String, sUsuario As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True
  GSSQL = ""
  GSSQL = gsCompania & ".ccUpdateCliente '" & sOperacion & "'," & sIDCliente & ",'" & sNombre & "','" & sRazonSocial & "',"
  GSSQL = GSSQL & "'" & sDireccion & "','" & sTelefono1 & "','" & sTelefono2 & "','" & sTelefono3 & "','" & sCelular1 & "','" & sCelular2 & "',"
  GSSQL = GSSQL & "'" & semail & "'," & sEsFarmacia & ",'" & sNombreFarmacia & "','" & sRUC & "',"
  GSSQL = GSSQL & "'" & sPropietario & "'," & sIDBodega & ",'" & sIDPlazo & "','" & sIDMoneda & "','" & sIDCategoria & "','"
  GSSQL = GSSQL & sIDDepartamento & "','" & sIDMunicipio & "','" & sIDZona & "'," & sIDVendedor & ","
  GSSQL = GSSQL & sTechoCredito & " ," & sActivo & " , '"
  GSSQL = GSSQL & sUsuario & "'"
  
    
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      'gsOperacionError = "Eliminando el Beneficiado. " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & err.Description
      SetMsgError "Ocurrió un error actualizando el cliente. ", err
      lbok = False
    End If

ccUpdateCliente = lbok
Exit Function

error:
  
  lbok = False
  Resume Next

End Function


Public Function invUpdateProducto(sOperacion As String, sIDProducto As String, sDescr As String, sImpuesto As String, _
sEsMuestra As String, sEsControlado As String, sClasif1 As String, sClasif2 As String, sClasif3 As String, _
sEsEtico As String, sBajaPrecioDistribuidor As String, sIDProveedor As String, sCostoUltLocal As String, _
sCostoUltDolar As String, sCostoUltPromLocal As String, sCostoUltPromDolar As String, sPrecioPublicoLocal As String, _
sPrecioFarmaciaLocal As String, sPrecioCIFLocal As String, sPrecioFOBLocal As String, sIDPresentacion As String, _
sBajaPrecioProveedor As String, sPorcDescAlzaProveedor As String, sUserInsert As String, sUserUpdate As String, _
sActivo As String, sCodigoBarra As String, sBonificaFA As String, _
sBonifCOPorCada As String, sBonifCOCantidad As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True
  GSSQL = ""
  GSSQL = gsCompania & ".invUpdateProducto '" & sOperacion & "'," & sIDProducto & ",'" & sDescr & "','" & sImpuesto & "',"
  GSSQL = GSSQL & sEsMuestra & "," & sEsControlado & ",'" & sClasif1 & "','" & sClasif2 & "','" & sClasif3 & "'," & sEsEtico & ","
  GSSQL = GSSQL & sBajaPrecioDistribuidor & "," & sIDProveedor & "," & sCostoUltLocal & "," & sCostoUltDolar & "," & sCostoUltPromLocal & ","
  GSSQL = GSSQL & sCostoUltPromDolar & "," & sPrecioPublicoLocal & "," & sPrecioFarmaciaLocal & "," & sPrecioCIFLocal & ","
  GSSQL = GSSQL & sPrecioFOBLocal & ",'" & sIDPresentacion & "'," & sBajaPrecioProveedor & "," & sPorcDescAlzaProveedor & ",'"
  GSSQL = GSSQL & sUserInsert & "','" & sUserUpdate & "'," & sActivo & ",'" & sCodigoBarra & "', "
  GSSQL = GSSQL & sBonificaFA & " , " & sBonifCOPorCada & "," & sBonifCOCantidad
    
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      'gsOperacionError = "Eliminando el Beneficiado. " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & err.Description
      SetMsgError "Ocurrió un error actualizando el producto. ", err
      lbok = False
    End If

invUpdateProducto = lbok
Exit Function

error:
  
  lbok = False
  Resume Next

End Function

Public Function sgvActualizaAlerta(sCodigo As String, sDescr As String, sInicio As String, sFin As String, _
        sOperacion As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True

  GSSQL = gsCompania & ".sgvActualizaAlertas " & sCodigo & ",'" & sDescr & "' , " & sInicio & ", " & sFin & ",'" & sOperacion & "'"
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el Beneficiado."
      lbok = False
    End If

sgvActualizaAlerta = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function



Public Function sgvActualizaPeriodo(sPeriodo As String, sFechaInicio As String, sFechaFinal As String, sEstado As String, _
        sOperacion As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True

sFechaInicio = Format(sFechaInicio, "yyyymmdd") & " 00:00:00.000"
sFechaFinal = Format(sFechaFinal, "yyyymmdd") & " 23:59:59.000"

  GSSQL = gsCompania & ".sgvActualizaPeriodo " & sPeriodo & ",'" & sFechaInicio & "' , '" & sFechaFinal & "', '" & sEstado & "','" & sOperacion & "'"

    
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el Beneficiado."
      lbok = False
    End If

sgvActualizaPeriodo = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function

Public Function secUpdateUsuario(sUsuario As String, sNombre As String, sPassword As String, sActivo As String, sIDCDIS As String, _
        sOperacion As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True


  GSSQL = gsCompania & ".secUpdateUsuario '" & sOperacion & "','" & sUsuario & "','" & sNombre & "' , '" & sPassword & "', '" & sActivo & "'," & sIDCDIS

    
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el Beneficiado."
      lbok = False
    End If

secUpdateUsuario = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function

Public Function secUpdateEmpleado(sEmpleado As String, sNombre As String, sFecha_Ingreso As String, sFecha_Salida As String, sSalario_Referencia As String, sActivo As String, _
       sDepartamento As String, sCentro As String, sOperacion As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True

'sFecha_Ingreso = Format(sFecha_Ingreso, "yyyymmdd")
'sFecha_Salida = Format(sFecha_Salida, "yyyymmdd")

  GSSQL = gsCompania & ".sgvUpdateglobalEmpleado '" & sOperacion & "','" & sEmpleado & "','" & sNombre & "' , '" & sFecha_Ingreso & "', '" & sFecha_Salida & "'," & sSalario_Referencia & ",'" & sActivo & "','" & sDepartamento & "','" & sCentro & "'"

    
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el Beneficiado."
      lbok = False
    End If

secUpdateEmpleado = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function


Public Function secActualizaRoleUsuario(sIDModulo As String, sUsuario As String, sIDRole As String, sOperacion As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True

  GSSQL = gsCompania & ".secActualizaRoleUsuario " & sIDModulo & ",'" & sUsuario & "' , " & sIDRole & ", '" & sOperacion & "'"

    
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el Beneficiado."
      lbok = False
    End If

secActualizaRoleUsuario = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function


Public Function setFechaPeriodo(sDia As String, sMes As String, dFechaSelected As Date) As Date
Dim dFechaResultado As Date

    dFechaResultado = CDate(Str(Year(dFechaSelected)) + "-" + sMes + "-" + sDia)
setFechaPeriodo = dFechaResultado
End Function


Public Function sgvGetCicloAbierto() As String
Dim lbok As Boolean
Dim sResultado As String
On Error GoTo error
lbok = True
  GSSQL = "SELECT " & gsCompania & "." & "sgvGetCicloAbierto() Periodo "
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbok = False  'Indica que no es válido
    sResultado = "-1"
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    sResultado = Str(gRegistrosCmd("Periodo").value)
    lbok = True
  End If
  sgvGetCicloAbierto = sResultado
  gRegistrosCmd.Close
  Exit Function
error:
  lbok = False
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
Dim lbok As Boolean
On Error GoTo error

lbok = True
'sNewPassword = zEncrypt(sNewPassword)
  GSSQL = "UPDATE " & gsCompania & ".secUSUARIO SET Password= '" & sNewPassword & "'"
  GSSQL = GSSQL & " WHERE USUARIO ='" & sUsuario & "'"

    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbok = False
    End If

Update_Password = lbok
Exit Function

error:
  lbok = False
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

Dim sValor As String
Dim lbok As Boolean
On Error GoTo error
lbok = False

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
    lbok = False
    gsOperacionError = "Error en la búsqueda del artículo !!!" & err.Description
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbok = True ' Si Existe el codigo
  End If
  ExistCodeInTable = lbok
  gRegistrosCmd.Close
  Exit Function
error:
  gsOperacionError = "Ocurrió un error en la operación de búsqueda de la descripción " & err.Description
  Resume Next
End Function

Public Sub CargarDatos(xSet As ADODB.Recordset, tdbgData As TDBGrid, sNameFieldCode As String, sNameFieldDescr As String)
On Error GoTo EH
Dim i As Integer

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
Dim lbok As Boolean
On Error GoTo error

lbok = True


  GSSQL = gsCompania & ".uspparmaUpdateCliente '" & sCliente & "'," & sSupervisor & " , " & sCanal & ", " & sStatus & " , " & sResponsable & "," & sEsPreventa & "," & sCDIS

    
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Error al Actualizar el Cliente."
      lbok = False
    End If

uspparmaUpdateCliente = lbok
Exit Function

error:
  lbok = False
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
Dim lbok As Boolean
On Error GoTo error

lbok = True
  GSSQL = gsCompania & "." & "uspparmaUpdateAplicacionesFromAsiento '" & sAsiento & "','" & sFecha & "'"
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbok = False
    End If

uspparmaUpdateAplicacionesFromAsiento = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function

Public Function Prepara_DiferencialCC(sUsuario As String, sFecha As String, sTipoCambio As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True
GSSQL = gsCompania & ".uspparmaPrepareDiferencialCC" & " '" & sUsuario & "' , '" & sFecha & "' , " & sTipoCambio
'gConet.BeginTrans

gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbok = False
      'gConet.RollbackTrans
      Prepara_DiferencialCC = lbok
      Exit Function
    End If

Prepara_DiferencialCC = lbok
'gConet.CommitTrans
Exit Function

error:
  lbok = False
  Resume Next

End Function

Public Function uspparmaInsertPrepDiferencialCC(sFechaUltProc As String, sFechaCorte As String, sTipoCambio As String, sPaquete As String, _
sDescrPaquete As String, sTipo As String, sDescrTipo As String, sUsuario As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True
GSSQL = gsCompania & ".uspparmaInsertPrepDiferencialCC" & " '" & sFechaUltProc & "' , '" & sFechaCorte & "' , " & _
sTipoCambio & ",'" & sPaquete & "','" & sDescrPaquete & "','" & sTipo & "','" & sDescrTipo & "','" & sUsuario & "'"
'gConet.BeginTrans

gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbok = False
      'gConet.RollbackTrans
      uspparmaInsertPrepDiferencialCC = lbok
      Exit Function
    End If

uspparmaInsertPrepDiferencialCC = lbok
'gConet.CommitTrans
Exit Function

error:
  lbok = False
  Resume Next

End Function

Public Function Prepara_DiferencialCambiario(sModo As String, sFecha As String, sFechaCorte As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True
GSSQL = gsCompania & ".uspparmaSetProcesoDifCamb" & "'" & sModo & "' , '" & sFecha & "','" & sFechaCorte & "'"
'gConet.BeginTrans

gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbok = False
      'gConet.RollbackTrans
      Prepara_DiferencialCambiario = lbok
      Exit Function
    End If

Prepara_DiferencialCambiario = lbok
'gConet.CommitTrans
Exit Function

error:
  lbok = False
  Resume Next

End Function

Public Function Prepara_DiferencialCambiarioCP(sModo As String, sFecha As String, sFechaCorte As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True
GSSQL = gsCompania & ".uspparmaSetProcesoDifCambCP" & "'" & sModo & "' , '" & sFecha & "','" & sFechaCorte & "'"
'gConet.BeginTrans

gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "Eliminando el producto del Proveedor."
      lbok = False
      'gConet.RollbackTrans
      Prepara_DiferencialCambiarioCP = lbok
      Exit Function
    End If

Prepara_DiferencialCambiarioCP = lbok
'gConet.CommitTrans
Exit Function

error:
  lbok = False
  Resume Next

End Function




Public Function parmaExistePreparacionCC() As Boolean
Dim lbok As Boolean
Dim iResultado As Integer
On Error GoTo error
lbok = False
  GSSQL = "SELECT " & gsCompania & "." & "parmagetExistePreparacionCC() Status"
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    iResultado = gRegistrosCmd("Status").value
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    iResultado = gRegistrosCmd("Status").value
  End If
  If iResultado = 1 Then lbok = True Else lbok = False
  parmaExistePreparacionCC = lbok
  gRegistrosCmd.Close
  Exit Function
error:
  iResultado = 0
  lbok = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function


Public Function parmaExistePreparacionCP() As Boolean
Dim lbok As Boolean
Dim iResultado As Integer
On Error GoTo error
lbok = False
  GSSQL = "SELECT " & gsCompania & "." & "parmagetExistePreparacionCP() Status"
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    iResultado = gRegistrosCmd("Status").value
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    iResultado = gRegistrosCmd("Status").value
  End If
  If iResultado = 1 Then lbok = True Else lbok = False
  parmaExistePreparacionCP = lbok
  gRegistrosCmd.Close
  Exit Function
error:
  iResultado = 0
  lbok = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function

Public Function uspparmaAplicaRecibosLiquidacion(sLiquidacion As String, sUsuario As String, sTotalPago As String, sFechaRecibo As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True
GSSQL = gsCompania & ".uspparmaAplicaRecibosLiquidacion " & "'" & sLiquidacion & "' , '" & sUsuario & "', " & sTotalPago & ", '" & sFechaRecibo & "'"
'gConet.BeginTrans

gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
      gsOperacionError = "aplicando"
      lbok = False
      'gConet.RollbackTrans
      uspparmaAplicaRecibosLiquidacion = lbok
      Exit Function
    End If

uspparmaAplicaRecibosLiquidacion = lbok
'gConet.CommitTrans
Exit Function

error:
  lbok = False
  Resume Next

End Function

Public Function ExisteUsuarioExactus(sUsuario As String) As Boolean
Dim lbok As Boolean
On Error GoTo error
lbok = False
  GSSQL = "SELECT * from " & dbo & "." & "Usuario where Usuario ='" & sUsuario & "'"
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbok = False
    
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbok = True
  End If
  ExisteUsuarioExactus = lbok
  gRegistrosCmd.Close
  Exit Function
error:
  lbok = False
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
Dim lbok As Boolean

Dim gRegistrosCmd As ADODB.Recordset


On Error GoTo error


lbok = True
  GSSQL = gsCompania & ".[parmaUpdateparmaDocumentsToApply]  '" & sOperation & "','" & sCliente & "' ,'" & sDocumento & "','" & sFecha & "','" & _
    sTipo & "'," & sValor & "," & sIDFile & "," & sflgAplicado

  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbok = False  'Indica que no es válido

  Else 'If Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido

    'sResultado = gRegistrosCmd("Resultado").value

    lbok = True
  End If
  parmaUpdateparmaDocumentsToApply = lbok
  gRegistrosCmd.Close
  Exit Function
error:
  lbok = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next

End Function


Public Function ExisteCliente(sCliente As String) As Boolean
Dim lbok As Boolean
On Error GoTo error
lbok = False
  GSSQL = "SELECT * from " & gsCompania & "." & "CLIENTE where CLIENTE = '" & sCliente & "'"
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbok = False
    
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbok = True
  End If
  ExisteCliente = lbok
  gRegistrosCmd.Close
  Exit Function
error:
  lbok = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function

Public Function ExisteLiquidacion(sLiquidacion As String) As Boolean
Dim lbok As Boolean
On Error GoTo error
lbok = False
  GSSQL = "SELECT * from " & gsCompania & "." & "parmaAS400Liquidacion where idliquidacion = '" & sLiquidacion & "'"
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbok = False
    
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbok = True
  End If
  ExisteLiquidacion = lbok
  gRegistrosCmd.Close
  Exit Function
error:
  lbok = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function

Public Function parmaActualizaLiquidacionDesdeAS400(sIDLiquidacion As String, sFecha As String, sVendedor As String, sAux1 As String, sAux2 As String, sAux3 As String, _
sTotalLiquidacion As String, sTotalFaltante) As Boolean
Dim lbok As Boolean

Dim gRegistrosCmd As ADODB.Recordset


On Error GoTo error


lbok = True
  GSSQL = gsCompania & ".[parmaUpdateLiquidacionAS400]  'I','" & sIDLiquidacion & "','" & sFecha & "' ," & sTotalLiquidacion & "," & sTotalFaltante & ""
gConet.BeginTrans
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbok = False  'Indica que no es válido
    GoTo error
  End If
    
    If sVendedor <> "" Then
        GSSQL = gsCompania & ".[parmaLiquidacionVendedorAS400]  'I','" & sIDLiquidacion & "','" & sVendedor & "','V'"
        
        Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia
            If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
              gsOperacionError = "No existe ese cliente." 'Asigna msg de error
              lbok = False  'Indica que no es válido
              GoTo error
            End If
     End If

    If sAux1 <> "" Then
        GSSQL = gsCompania & ".[parmaLiquidacionVendedorAS400]  'I','" & sIDLiquidacion & "','" & sAux1 & "','A'"
        Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia
        If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
          gsOperacionError = "No existe ese cliente." 'Asigna msg de error
          lbok = False  'Indica que no es válido
          GoTo error
        End If
    End If
    If sAux2 <> "" Then
     GSSQL = gsCompania & ".[parmaLiquidacionVendedorAS400]  'I','" & sIDLiquidacion & "','" & sAux2 & "','A'"
     Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia
        If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
           gsOperacionError = "No existe ese cliente." 'Asigna msg de error
           lbok = False  'Indica que no es válido
           GoTo error
        End If
    End If
    If sAux3 <> "" Then
       GSSQL = gsCompania & ".[parmaLiquidacionVendedorAS400]  'I','" & sIDLiquidacion & "','" & sAux3 & "','A'"
       Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia
        If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
           gsOperacionError = "No existe ese cliente." 'Asigna msg de error
           lbok = False  'Indica que no es válido
           GoTo error
        End If
    End If
                       
        
        If lbok Then
            
            GSSQL = gsCompania & ".[parmaUpdateLiquidacionAS400]  'F','" & sIDLiquidacion & "'"
            Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia
            If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
               gsOperacionError = "No existe ese cliente." 'Asigna msg de error
               lbok = False  'Indica que no es válido
               GoTo error
            End If
            
        End If

    If lbok = False Then
        gConet.RollbackTrans
    Else
        gConet.CommitTrans
    End If
    
    
  
  parmaActualizaLiquidacionDesdeAS400 = lbok
  
  'gRegistrosCmd.Close
  Exit Function
error:
  lbok = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  gConet.RollbackTrans
   parmaActualizaLiquidacionDesdeAS400 = lbok

End Function


Public Function ExisteNotasAnuladas(sLiquidacion As String) As Boolean
Dim lbok As Boolean
On Error GoTo error
lbok = False
  GSSQL = "SELECT * from " & gsCompania & "." & "Documentos_CC where Tipo = 'N/D' AND Documento like '%" & Mid(sLiquidacion, 1, 8) & "%' and Anulado='S' AND CLIENTE IN " & _
  " (SELECT CLIENTE FROM " & gsCompania & ".parmaAS400LiquidacionDetFaltante WHERE IDLIQUIDACION = '" & sLiquidacion & "')"
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbok = False
    
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbok = True
  End If
  ExisteNotasAnuladas = lbok
  gRegistrosCmd.Close
  Exit Function
error:
  lbok = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function


Public Function SaldosNotasMenorQueFaltante(sLiquidacion As String, dFaltante As Double) As Boolean
Dim lbok As Boolean
On Error GoTo error
lbok = False
  GSSQL = "SELECT Sum(Saldo) Saldo from " & gsCompania & "." & "Documentos_CC where Tipo = 'N/D' AND Documento like '%" & Mid(sLiquidacion, 1, 8) & "%' and Anulado='N'" & " AND CLIENTE IN " & _
  " (SELECT CLIENTE FROM " & gsCompania & ".parmaAS400LiquidacionDetFaltante WHERE IDLIQUIDACION = '" & sLiquidacion & "')"
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbok = False
    
  ElseIf Not (gRegistrosCmd.BOF And gRegistrosCmd.EOF) Then  'Si no es válido
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    If Round(gRegistrosCmd(0).value, 2) < dFaltante Then
        lbok = True
    End If
  End If
  SaldosNotasMenorQueFaltante = lbok
  gRegistrosCmd.Close
  Exit Function
error:
  lbok = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function


Public Function CambiaStatusLiquidacion(sLiquidacion As String, sStatus As String) As Boolean
Dim lbok As Boolean
On Error GoTo error
lbok = False
  GSSQL = "Update " & gsCompania & "." & "parmaAS400Liquidacion set statusCobro = '" & sStatus & "' where idliquidacion = '" & sLiquidacion & "'"
          
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbok = False
    
  Else
    'gsOperacionError = "No existe ese cliente." 'Asigna msg de error
    lbok = True
  End If
  CambiaStatusLiquidacion = lbok
  'gRegistrosCmd.Close
  Exit Function
error:
  lbok = False
  gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
  Resume Next
                    
End Function


Public Function spGlobalUpdateCatalogo(sOperacion As String, sIDTable As String, sDescr As String, sActivo As String, sUsaValor As String, sNombreValor As String, sValor As String, Optional sIDCatalogo As String) As Boolean
Dim lbok As Boolean
Dim sIDCat As String
On Error GoTo error

lbok = True
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
      lbok = False
    End If

spGlobalUpdateCatalogo = lbok
Exit Function

error:
  lbok = False
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
Dim lbok As Boolean
Dim lsFilename As String
Dim rst As ADODB.Recordset
Dim rstCSV As ADODB.Recordset
lbok = False
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
    lbok = True
End If
LoadAccess = lbok
Exit Function
errores:
lbok = False
LoadAccess = lbok
End Function

Public Function UserCouldIN(sUsuario As String, sPassword As String)
On Error GoTo errores
Dim lbok As Boolean
Dim rst As ADODB.Recordset
lbok = False
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
    lbok = True
'    If zDecrypt(rst("Password").value) = sPassword Then
'        Set grRecordsetAcceso = rst
'        lbok = True
'    Else
'        lbok = False
'    End If
End If
UserCouldIN = lbok
Exit Function
errores:
lbok = False
UserCouldIN = lbok
End Function

Public Function fafUpdateVendedor(sOperacion As String, sIDVendedor As String, sNombre As String, sTipo As String, sActivo As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True

  GSSQL = gsCompania & ".fafUpdateVendedor '" & sOperacion & "'," & sIDVendedor & ",'" & sNombre & "'," & sActivo & ",'" & sTipo & "'"

    
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
       'gsOperacionError = "Eliminando el Beneficiado."
       SetMsgError "Eliminando el Vendedor.", err
      lbok = False
    End If

fafUpdateVendedor = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function

Public Function fafUpdateEscalaBonificacion(sOperacion As String, sIDProducto As String, sEscala As String, sPorCada As String, sBonifica As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True

  GSSQL = gsCompania & ".fafUpdateEscalaBonificacion '" & sOperacion & "'," & sIDProducto & "," & sEscala & "," & sPorCada & "," & sBonifica
    
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
       'gsOperacionError = "Eliminando el Beneficiado."
       SetMsgError "Eliminando Escala Bonificacion.", err
      lbok = False
    End If

fafUpdateEscalaBonificacion = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function



Public Function ExiteItem(rst As ADODB.Recordset, sCondicion As String) As Boolean
Dim lbok As Boolean
lbok = False
If Not rst.EOF Then
  rst.MoveFirst
  rst.Find sCondicion, 0, adSearchForward, 0
  If Not rst.EOF Then
    lbok = True
  End If
End If
ExiteItem = lbok
End Function


Public Sub SetMsgError(sError As String, oError As error)
    gsOperacionError = sError & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & oError.Description
End Sub

Public Function getValueFieldFromTable(sTabla As String, sFieldName As String, sFiltro As String, sFieldName2 As String, _
ByRef sResult1 As String, ByRef sResult2 As String, _
Optional bISNumericFirstField As Boolean, Optional bISNumericSecondField As Boolean) As String
Dim lbok As Boolean
Dim sResultado As String

Dim rst As ADODB.Recordset
On Error GoTo error
lbok = False
sResultado = "ND"
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst.CursorType = adOpenKeyset  'Asigna un cursor estático
    rst.CursorLocation = adUseClient ' Cursor local al cliente
    rst.LockType = adLockOptimistic

  lbok = True
  If sFieldName2 = "" Then
    GSSQL = "SELECT top 1 " & sFieldName
  Else
    GSSQL = "SELECT top 1 " & sFieldName & "," & sFieldName2
  End If
  GSSQL = GSSQL & " FROM " & gsCompania & "." & sTabla          'Constuye la sentencia SQL
  
    GSSQL = GSSQL & " WHERE " & sFiltro

  If rst.State = adStateOpen Then rst.Close
  rst.Open GSSQL, gConet, adOpenDynamic, adLockOptimistic, adCmdText    'Ejecuta la sentencia

If Not (rst.BOF And rst.EOF) Then
    If bISNumericFirstField Then
        sResult1 = Str(rst(sFieldName).value)  ' lo retorna siempre caracter
    Else
        sResult1 = rst(sFieldName).value
    End If
    
    If bISNumericSecondField Then
        sResult2 = Str(rst(sFieldName2).value)  ' lo retorna siempre caracter
    Else
        sResult2 = rst(sFieldName2).value
    End If
    lbok = True
End If

getValueFieldFromTable = lbok
Set rst = Nothing
Exit Function
error:
  lbok = False
  getValueFieldFromTable = lbok

End Function


Public Function getValueFieldsFromTable(sTabla As String, sListFieldName As String, sFiltro As String, ByRef dicResult As Dictionary) As Boolean
    Dim lbok As Boolean
    Dim sResultado As String
    Dim rst As ADODB.Recordset
    Dim i As Integer
    
    Set dicResult = New Dictionary
    On Error GoTo error
    lbok = False
    sResultado = "ND"
    
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst.CursorType = adOpenKeyset  'Asigna un cursor estático
    rst.CursorLocation = adUseClient ' Cursor local al cliente
    rst.LockType = adLockOptimistic
    
    lbok = True
      
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
        lbok = True
    End If
    
    getValueFieldsFromTable = lbok
    Set rst = Nothing
    Exit Function
error:
      lbok = False
      getValueFieldsFromTable = lbok
    
End Function


Public Function invGetSugeridoLote(IdBodega As Integer, IdProducto As Integer, Cantidad As Double) As ADODB.Recordset
  
    Dim rs As ADODB.Recordset
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

' obtiene la descr del catálogo de un codigo numerico
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

Public Function getDescrCatalogo(txtCodigo As TextBox, sFieldNameCode As String, sTableName As String, sFieldNameDescr As String, Optional bCodeChar As Boolean) As String
Dim lbok As Boolean
Dim sDescr As String
Dim sValor As String
lbok = False
If txtCodigo.Text <> "" Then
    If bCodeChar = True Then
        sValor = "'" & txtCodigo.Text & "'"
    Else
        sValor = txtCodigo.Text
    End If
    
    sDescr = GetDescrCat(sFieldNameCode, sValor, sTableName, sFieldNameDescr)
Else
    sDescr = ""
End If
getDescrCatalogo = sDescr
End Function



Public Function ExisteDependencia(sTablaRelacion As String, sFldname As String, sValFld As String, sType As String) As Boolean
Dim lbok As Boolean
On Error GoTo error
lbok = False
Dim rs As ADODB.Recordset
  gstrSQL = "SELECT TOP 1 " & sFldname
  gstrSQL = gstrSQL & " FROM " & sTablaRelacion
  If UCase(sType) = "N" Then
    gstrSQL = gstrSQL & " WHERE " & sFldname & " = " & sValFld 'Constuye la sentencia SQL
  Else
    gstrSQL = gstrSQL & " WHERE " & sFldname & " = '" & sValFld & "'" 'Constuye la sentencia SQL
  End If

  Set rs = gConet.Execute(gstrSQL, , adCmdText)  'Ejecuta la sentencia

  If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
    lbok = False  'Indica que ocurrió un error
    sMensajeError = "Error en la búsqueda  !!!" & err.Description
  ElseIf Not (rs.BOF And rs.EOF) Then  'Si no es válido
    
    lbok = True  'Indica que ya existe
  End If
  rs.Close

ExisteDependencia = lbok
Exit Function
error:
    lbok = False
    Resume Next
End Function



Public Function fafUpdatePedido(sOperacion As String, ByRef sIDPedido As String, sIDBodega As String, sIDCliente As String, sIDVendedor As String, sFecha As String, _
sAprobado As String, sBackOrder As String, sAnulado As String) As Boolean
Dim lbok As Boolean
Dim sResultado As String
Dim gRegistrosCmd As ADODB.Recordset
On Error GoTo error

lbok = True

  GSSQL = gsCompania & ".fafUpdatePedido '" & sOperacion & "'," & sIDPedido & "," & sIDBodega & "," & sIDCliente & "," & sIDVendedor & ",'" & sFecha & "'," & _
  sAprobado & "," & sBackOrder & "," & sAnulado
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)
    
    'gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
        'gsOperacionError = "Eliminando el Beneficiado."
        
        SetMsgError "Ocurrió un error actualizando La bodega. ", err
      lbok = False
    Else
       sResultado = gRegistrosCmd(0).value
       sIDPedido = sResultado
    End If

fafUpdatePedido = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function

Public Function fafUpdatePedidoLinea(sOperacion As String, ByRef sIDPedido As String, sIDBodega As String, sIDCliente As String, sIDVendedor As String, sFecha As String, _
sIDProducto As String, sCantidadPedida As String, sPrecio As String, sSubTotal As String, sTotalImpuesto As String, sTotal As String) As Boolean
Dim lbok As Boolean

Dim gRegistrosCmd As ADODB.Recordset
On Error GoTo error

lbok = True

  GSSQL = gsCompania & ".fafUpdatePedidoLinea '" & sOperacion & "'," & sIDPedido & "," & sIDBodega & "," & sIDCliente & "," & sIDVendedor & ",'" & sFecha & "'," & _
  sIDProducto & "," & sCantidadPedida & "," & sPrecio & "," & sSubTotal & "," & sTotalImpuesto & "," & sTotal
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)
    
    'gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
        'gsOperacionError = "Eliminando el Beneficiado."
        
        SetMsgError "Ocurrió un error actualizando La bodega. ", err
      lbok = False
    Else
        lbok = True
    End If

fafUpdatePedidoLinea = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function


Public Sub ImprimePedido(rs As ADODB.Recordset, Optional bDirectPrint As Boolean = False)
Dim rep As rptPedido
Dim sconstr As String
Set rep = New rptPedido
sconstr = gsConetstr
rep.DataControl1.ConnectionString = sconstr
If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
End If
If Not rs.EOF Then
  Set rep.DataControl1.Recordset = rs

  rep.Printer.Orientation = ddOPortrait
  rep.Toolbar.Tools.Add "Export PDF"
  rep.lblNombreEmpresa.Caption = gparametros.NombreEmpresa '  "Parmalat Centroamérica, S.A." 'gParametros.NombreEmpresa
  rep.lblTitulo.Caption = "PEDIDO "
  'rep.lblFechacorte = "Fecha de Corte : " & sFecha
  'rep.Detail.Visible = False
  rep.Printer.PaperSize = 5
  rep.Printer.Orientation = ddOPortrait
  If bDirectPrint = True Then
     rep.PrintReport False ' directo a la impresora
  Else
     rep.Show vbModal
  End If
  
    rs.MoveFirst
Else
  lbok = Mensaje("No hay registros para imprimir...", ICO_ERROR, False)
End If
End Sub



Public Function fafgetCantBodegaFacturableForUser(sUsuario As String) As Integer
Dim lbok As Boolean
Dim iResultado As Integer
Dim gRegistrosCmd As ADODB.Recordset
On Error GoTo error

lbok = True

  GSSQL = "Select " & gsCompania & ".fafgetCantBodegaFacturableForUser( '" & sUsuario & "') as Resultado"
  
  Set gRegistrosCmd = gConet.Execute(GSSQL, , adCmdText)
    
    'gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
        'gsOperacionError = "Eliminando el Beneficiado."
        
        SetMsgError "Ocurrió un error actualizando La bodega. ", err
      iResultado = 0
    Else
       iResultado = gRegistrosCmd(0).value
       fafgetCantBodegaFacturableForUser = iResultado
    End If

fafgetCantBodegaFacturableForUser = iResultado
Exit Function

error:
  lbok = False
  Resume Next

End Function


Public Function invUpdateBodegaUsuario(sOperacion As String, sBodega As String, sUsuario As String, sFactura As String, sConsInv As String) As Boolean
Dim lbok As Boolean
On Error GoTo error

lbok = True

  GSSQL = gsCompania & ".invUpdateBodegaUsuario '" & sOperacion & "'," & sBodega & ",'" & sUsuario & "'," & sFactura & "," & sConsInv

    
    gConet.Execute GSSQL, , adCmdText + adExecuteNoRecords 'Ejecuta la sentencia

    If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
       'gsOperacionError = "Eliminando el Beneficiado."
       SetMsgError "Eliminando el Vendedor.", err
      lbok = False
    End If

invUpdateBodegaUsuario = lbok
Exit Function

error:
  lbok = False
  Resume Next

End Function

' Obtener Unidades bonificadas en Facturación
Public Function getUnidadesBonificadas(dCantidad As Double, dPorCada As Double, dBonif As Double) As Double
Dim dResultado As Double
Dim dBonifNoRound As Double
dResultado = 0
If dPorCada <> 0 Then
    dBonifNoRound = dCantidad / dPorCada
End If
    dResultado = Floor(dBonifNoRound) * dBonif
getUnidadesBonificadas = dResultado
End Function




'Devuelve el entero más pequeño no menor que X.
'Ejemplo: Ceiling(1.23) = 2, Ceiling(-1.23) = -1
Public Function Ceiling(ByVal x As Double) As Long
   Ceiling = -Int(x * (-1))
End Function

'Devuelve el entero más grande no mayor que X.
'Ejemplo: Floor(1.23) = 1, Floor(-1.23) = -2
Public Function Floor(ByVal x As Double) As Long
   Floor = (-Int(x) * (-1))
End Function

Public Function EsEntero(sNumero As String) As Boolean
Dim lbok As Boolean
lbok = False
If IsNumeric(sNumero) Then
   If CInt(sNumero) = sNumero Then
       lbok = True
   Else
       lbok = False
   End If
End If
EsEntero = lbok
End Function



Public Sub SetupMenuButtons()
'
' Each form window has its own set of Toolbar buttons.
' Including the split there are 20 buttons.
' Buttons that are to be made visible are indicated by the presence of a letter in the string
' The value of that letter indicates the position in the toolbar list of the button that we want.
' For Example the string "CDEST" means make visible buttons in the toolbar array at index positions:
' 1, 2, 3, 17, 18
' 1 = Left, 2 = Right, 3 = separator, 17 = User's Guide, 18 = About


    ReDim ButtonsAvailable(0 To 19)
    
    ButtonsAvailable(0).form_buttons = "1,2,3,25,23,24"
    ButtonsAvailable(0).form_name = "no form"
    
    ButtonsAvailable(1).form_buttons = "1,2,3,8,10,11,12,13,17,18,19,22,25"
    ButtonsAvailable(1).form_name = "frmProductos"
    
    ButtonsAvailable(2).form_buttons = "1,2,3,8,10,11,12,13,17,18,19,22,25"
    ButtonsAvailable(2).form_name = "frmBodega"
    
    ButtonsAvailable(3).form_buttons = "1,2,3,8,10,11,12,13,17,18,19,22,25"
    ButtonsAvailable(3).form_name = "frmMasterLotes"
    
    ButtonsAvailable(4).form_buttons = "1,2,3,8,10,11,12,13,17,18,19,22,25"
    ButtonsAvailable(4).form_name = "frmCatalogos"
    
    ButtonsAvailable(5).form_buttons = "1,2,3,8,10,11,12,13,17,18,19,22,25"
    ButtonsAvailable(5).form_name = "frmCliente"
    
    ButtonsAvailable(6).form_buttons = "1,2,3,24,8,22"
    ButtonsAvailable(6).form_name = "frmPedidos"
    
    ButtonsAvailable(7).form_buttons = "1,2,3,8,10,11,12,13,17,18,19,22,25"
    ButtonsAvailable(7).form_name = "frmTablas"
    
    ButtonsAvailable(8).form_buttons = "1,2,3,24,8,22"
    ButtonsAvailable(8).form_name = "frmTransacciones"
    
    ButtonsAvailable(9).form_buttons = "1,2,3,8,10,11,12,13,17,18,19,22,25"
    ButtonsAvailable(9).form_name = "frmVendedor"
    
    ButtonsAvailable(10).form_buttons = "1,2,3,24,8,20,21,22"
    ButtonsAvailable(10).form_name = "frmListadoTraslados"
    
    ButtonsAvailable(11).form_buttons = "1,2,3,24,22"
    ButtonsAvailable(11).form_name = "frmRegistrarTraslado"
    '#################  hasta aca #######################3
    
    ButtonsAvailable(12).form_buttons = "CDEFLPQRU"
    ButtonsAvailable(12).form_name = "frmStockMonitoring"
    
    ButtonsAvailable(13).form_buttons = "CDEIKLOPQRU"
    ButtonsAvailable(13).form_name = "frmStockReceive"
    
    ButtonsAvailable(14).form_buttons = "CDEIJLNPQRU"
    ButtonsAvailable(14).form_name = "frmLoading"
    
    ButtonsAvailable(15).form_buttons = "CDEIJLNPRU"
    ButtonsAvailable(15).form_name = "frmInvoice"
    
    ButtonsAvailable(16).form_buttons = "CDEIJLNPRU"
    ButtonsAvailable(16).form_name = "frmVanCollection"
    
    ButtonsAvailable(17).form_buttons = "CDEIJLNPRU"
    ButtonsAvailable(17).form_name = "frmVanInventory"
    
    ButtonsAvailable(18).form_buttons = "CDEIJLNPRU"
    ButtonsAvailable(18).form_name = "frmVanRemmitance"

       
    
End Sub

Public Sub LoadForm(ByRef srcForm As Form)
    On Error Resume Next
    srcForm.Show
    srcForm.WindowState = vbMaximized
    srcForm.SetFocus
End Sub


Public Sub SetupFormToolbar(frm As String)
' Given a form name set up the toolbar buttons appropriate for that form
    'Dim pattern As String
    Dim pattern() As String
    Dim p As Integer
    Dim j As Integer
    Dim i As Integer
    Dim visibility(1 To 32) As Boolean
    Dim foundfrm As Boolean
    
    foundfrm = False
    For i = 1 To 32
        visibility(i) = False                   'Initially assume all toolbar buttons are invisible
    Next i
    
    'pattern = ""
    
    For i = 0 To 18                            'There are 24 types of forms from type 0 to type 23
        If frm = ButtonsAvailable(i).form_name Then
            'pattern = ButtonsAvailable(i).form_buttons
            pattern = Split(ButtonsAvailable(i).form_buttons, ",")
            'For j = 1 To Len(pattern)
            For j = 0 To UBound(pattern)
                'p = Asc(Mid(pattern, j, 1)) - 64    'if it was "C" than p = 1
                p = pattern(j)
                visibility(p) = True
              
            Next j
            For j = 1 To 24
                MDIMain.tbMenu.Buttons.Item(j).Visible = visibility(j)
            Next j
            foundfrm = True
            Exit For
        End If
    Next i              'only continue of form name didn't match
    
    If foundfrm = False Then        'If the form name was not found then default to buttons pattern for "no form"
        'pattern = ButtonsAvailable(0).form_buttons
        pattern = Split(ButtonsAvailable(i).form_buttons, ",")
        For j = 0 To UBound(pattern)
                'p = Asc(Mid(pattern, j, 1)) - 64
                p = pattern(j)
                visibility(p) = True
              
            Next j
            For j = 1 To 23
                MDIMain.tbMenu.Buttons.Item(j).Visible = visibility(j)
            Next j
        
    End If
    
    

End Sub



'Procedure used to center form
Public Sub centerForm(ByRef sForm As Form, ByVal sHeight As Integer, ByVal sWidth As Integer)
    sForm.Move (sWidth - sForm.Width) / 2, (sHeight - sForm.Height) / 2
End Sub
'Procedure used to center object horizontal
Public Sub center_obj_horizontal(ByVal sParentObj As Variant, ByRef sMoveObj As Variant)
    sMoveObj.left = (sParentObj.Width - sMoveObj.Width) / 2
End Sub
'Procedure used to center vertical
Public Sub center_obj_vertical(ByVal sParentObj As Variant, ByRef sMoveObj As Variant)
    sMoveObj.top = (sParentObj.Height - sMoveObj.Height) / 2
End Sub

Public Sub HighlightInWin(ByVal srcKey As String)
    With MDIMain.lvWin
        If .ListItems.Count > 0 Then
            If .SelectedItem.Key <> srcKey Then
                Dim c As Integer
                For c = 1 To .ListItems.Count
                    If .ListItems(c).Key = srcKey Then
                        .ListItems(c).Selected = True
                        .ListItems(c).EnsureVisible
                        Exit For
                    End If
                Next c
            End If
        End If
    End With
End Sub

Public Function IsWindowInListbox(s As String, LBox As ListBox) As Integer
's is the name of a window's caption
'Scan the strings in the listbox  comparing each string to s
'If found return the index number of that string (zero based)
'if not found return -1
'
Dim ls As String
Dim i As Integer
    If (LBox.ListCount = 0) Then
        IsWindowInListbox = -1
        Exit Function
    End If
    For i = 0 To LBox.ListCount - 1
        ls = LBox.List(i)
        If ls = s Then
            IsWindowInListbox = i
            Exit Function
        End If
    Next i
    
    IsWindowInListbox = -1
End Function





