Attribute VB_Name = "Mdl_localizar_servidor_SQL"
'--- retvals
Private Const SQL_ERROR                     As Integer = -1
Private Const SQL_INVALID_HANDLE            As Integer = -2
Private Const SQL_NEED_DATA                 As Integer = 99
Private Const SQL_NO_DATA_FOUND             As Integer = 100
Private Const SQL_SUCCESS                   As Integer = 0
Private Const SQL_SUCCESS_WITH_INFO         As Integer = 1
'--- for SQLSetConnectOption
Private Const SQL_ATTR_LOGIN_TIMEOUT        As Long = 103
Private Const SQL_ATTR_CONNECTION_TIMEOUT   As Long = 113
Private Const SQL_ATTR_QUERY_TIMEOUT        As Long = 0
Private Const SQL_COPT_SS_BASE              As Long = 1200
Private Const SQL_COPT_SS_INTEGRATED_SECURITY As Long = (SQL_COPT_SS_BASE + 3) ' Force integrated security on login
Private Const SQL_COPT_SS_BASE_EX           As Long = 1240
Private Const SQL_COPT_SS_BROWSE_CACHE_DATA As Long = (SQL_COPT_SS_BASE_EX + 5) ' Determines if we should cache browse info. Used when returned buffer is greater then ODBC limit (32K)
'--- param type
Private Const SQL_IS_UINTEGER               As Integer = (-5)
Private Const SQL_IS_INTEGER                As Integer = (-6)
'--- for SQL_COPT_SS_INTEGRATED_SECURITY
Private Const SQL_IS_OFF                    As Long = 0
Private Const SQL_IS_ON                     As Long = 1
'--- for SQL_COPT_SS_BROWSE_CACHE_DATA
Private Const SQL_CACHE_DATA_NO             As Long = 0
Private Const SQL_CACHE_DATA_YES            As Long = 1
'--- for SQLSetEnvAttr
Private Const SQL_ATTR_ODBC_VERSION         As Long = 200
Private Const SQL_OV_ODBC3                  As Long = 3

Private Declare Function SQLAllocEnv Lib "odbc32.dll" (phEnv As Long) As Integer
Private Declare Function SQLAllocConnect Lib "odbc32.dll" (ByVal hEnv As Long, phDbc As Long) As Integer
Private Declare Function SQLSetEnvAttr Lib "odbc32" (ByVal EnvironmentHandle As Long, ByVal Attrib As Long, Value As Any, ByVal StringLength As Long) As Integer
Private Declare Function SQLBrowseConnect Lib "odbc32.dll" (ByVal hDbc As Long, ByVal szConnStrIn As String, ByVal cbConnStrIn As Integer, ByVal szConnStrOut As String, ByVal cbConnStrOutMax As Integer, pcbConnStrOut As Integer) As Integer
Private Declare Function SQLDisconnect Lib "odbc32.dll" (ByVal hDbc As Long) As Integer
Private Declare Function SQLFreeConnect Lib "odbc32.dll" (ByVal hDbc As Long) As Integer
Private Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal hEnv As Long) As Integer
Private Declare Function SQLSetConnectOption Lib "odbc32.dll" (ByVal ConnectionHandle As Long, ByVal Option_ As Integer, ByVal Value As Long) As Integer
Private Declare Function SQLGetConnectOption Lib "odbc32.dll" (ByVal ConnectionHandle As Long, ByVal Option_ As Integer, Value As Long) As Integer
Private Declare Function SQLError Lib "odbc32.dll" (ByVal EnvironmentHandle As Long, ByVal ConnectionHandle As Long, ByVal StatementHandle As Long, ByVal Sqlstate As String, NativeError As Long, ByVal MessageText As String, ByVal BufferLength As Integer, TextLength As Integer) As Integer
'--- ODBC 3.0
Private Declare Function SQLSetConnectAttr Lib "odbc32" Alias "SQLSetConnectAttrA" (ByVal ConnectionHandle As Long, ByVal Attrib As Long, Value As Any, ByVal StringLength As Long) As Integer
Private Declare Function SQLGetConnectAttr Lib "odbc32" Alias "SQLGetConnectAttrA" (ByVal ConnectionHandle As Long, ByVal Attrib As Long, Value As Any, ByVal BufferLength As Long, StringLength As Long) As Integer

Private Const STR_NO_USER_DBS           As String = "<No user databases>"

Public Function EnumSqlServers() As Variant
On Error GoTo tratar_erro
Const CONN_STR      As String = "DRIVER={SQL Server}"
Const PREFIX        As String = "SERVER:Servidor={"
Const SUFFIX        As String = "}"

EnumSqlServers = pvBrowseConnect(CONN_STR, PREFIX, SUFFIX)

Exit Function
tratar_erro:
    MsgBox ("Descri��o do erro : " + Error()), vbCritical
    Exit Function
End Function

Public Function EnumSqlDbs(sServer As String, Optional sUser As String, Optional sPass As String) As Variant
On Error GoTo tratar_erro
Const CONN_STR      As String = "DRIVER={SQL Server};SERVER=%1;UID=%2;PWD=%3;"
Const PREFIX        As String = "Database={"
Const SUFFIX        As String = "}"
Dim sConnStr        As String
    
EnumSqlDbs = pvBrowseConnect( _
Replace(Replace(Replace(CONN_STR, _
        "%1", sServer), _
        "%2", sUser), _
        "%3", sPass), _
PREFIX, SUFFIX, Len(sUser) = 0)
    
Exit Function
tratar_erro:
    MsgBox ("Descri��o do erro : " + Error()), vbCritical
    Exit Function
End Function

Private Function pvBrowseConnect(sConnStr As String, sPrefix As String, sSuffix As String, Optional ByVal bItegrated As Boolean)
On Error GoTo tratar_erro
Dim rc              As Integer
Dim hEnv            As Long
Dim hDbc            As Long
Dim sBuffer         As String
Dim nReqBufSize     As Integer
Dim lStart          As Long
Dim lEnd            As Long
Dim dwSec           As Long
Dim lStrLen         As Long

'--- init environment
rc = SQLAllocEnv(hEnv)
rc = SQLSetEnvAttr(hEnv, SQL_ATTR_ODBC_VERSION, ByVal SQL_OV_ODBC3, SQL_IS_INTEGER)
'--- init conn
rc = SQLAllocConnect(hEnv, hDbc)
'--- timeouts to ~5 secs
rc = SQLSetConnectOption(hDbc, SQL_ATTR_CONNECTION_TIMEOUT, 3)
rc = SQLSetConnectOption(hDbc, SQL_ATTR_LOGIN_TIMEOUT, 3)
'--- integrated security
If bItegrated Then
    rc = SQLSetConnectOption(hDbc, SQL_COPT_SS_INTEGRATED_SECURITY, SQL_IS_ON)
End If
'--- improve performance
rc = SQLSetConnectOption(hDbc, SQL_COPT_SS_BROWSE_CACHE_DATA, SQL_CACHE_DATA_YES)
'--- initial buffer size
nReqBufSize = 1000
'--- repeat getting info until buffer gets large enough
Do
    sBuffer = String(nReqBufSize + 1, 0)
    '.lblStatus.Caption = "Status: Localizando servidor(es)..."
    rc = SQLBrowseConnect(hDbc, sConnStr, Len(sConnStr), sBuffer, Len(sBuffer), nReqBufSize)
    
Loop While rc = SQL_NEED_DATA And nReqBufSize >= Len(sBuffer)
'--- if ok -> parse buffer
If rc = SQL_SUCCESS Or rc = SQL_NEED_DATA Then
    '--- find prefix
    sPrefix = "{"
   lStart = InStr(1, sBuffer, sPrefix)
    lStart = InStr(1, sBuffer, "{")
    'frmDbSettings.lblStatus.Caption = "Status: Servidor(es)localizados com sucesso!..."

    'MsgBox (sPrefix)

    If lStart > 0 Then
        lStart = lStart + Len(sPrefix)
        '--- find suffix
        lEnd = InStr(lStart, sBuffer, sSuffix)
        If lEnd > 0 Then
            lEnd = lEnd - Len(sSuffix) + 1
            '--- success
            pvBrowseConnect = Split(Mid(sBuffer, lStart, lEnd - lStart), ",")
        End If
    Else
        Err.Raise vbObjectError, "ODBC", pvGetError(rc, hEnv, hDbc, 0)
    End If
End If
'--- disconnect
rc = SQLDisconnect(hDbc)
'--- free handles
rc = SQLFreeConnect(hDbc)
rc = SQLFreeEnv(hEnv)
'--- on failure -> return Array(0 To -1)
If Not IsArray(pvBrowseConnect) Then
    pvBrowseConnect = Split("")
End If

Exit Function
tratar_erro:
    MsgBox ("Descri��o do erro : " + Error()), vbCritical
    Exit Function
End Function

Private Function pvGetError(ByVal rc As Long, ByVal hEnv As Long, ByVal hDbc As Long, ByVal hStm As Long) As String
On Error GoTo tratar_erro
Dim sSqlState       As String * 5
Dim lNativeError    As Long
Dim sMsg            As String * 512
Dim nTextLength     As Integer
    
SQLError hEnv, hDbc, hStm, sSqlState, lNativeError, sMsg, Len(sMsg), nTextLength
pvGetError = "ODBC Result: 0x" & Hex(rc) & vbCrLf & vbCrLf & Left(sMsg, nTextLength)

Exit Function
tratar_erro:
    MsgBox ("Descri��o do erro : " + Error()), vbCritical
    Exit Function
End Function

Public Function EnumSqlDbAdo(sServer As String, Optional sUser As String, Optional sPass As String) As Variant
On Error GoTo tratar_erro
Dim cn              As New adodb.Connection
Dim vRet            As Variant
Dim lIdx            As Long
Dim sSQL            As String
    
cn.ConnectionTimeout = 3
cn.CursorLocation = adUseClient
For lIdx = 1 To 3
    If sUser <> "" Then
        cn.Open "Provider=SQLOLEDB;Persist Security Info=false;Data Source=" & sServer, sUser, sPass
    Else
        cn.Open "Provider=SQLOLEDB;Persist Security Info=false;Integrated Security=SSPI;Data Source=" & sServer
    End If
    If Err.Number = 0 Then
        sSQL = "select  name" & vbCrLf & _
               "from    sysdatabases " & vbCrLf & _
               "where   (status & 512) = 0" & vbCrLf & _
               "        and charindex('|' + name + '|', " & vbCrLf & _
               "            '|master|model|tempdb|msdb|distribution|') = 0"
        With cn.Execute(sSQL)
            ReDim vRet(.RecordCount - 1)
            lIdx = 0
            Do While Not .EOF
                If Left(!Name, 5) <> "DBNFE" Then
                    vRet(lIdx) = !Name
                    lIdx = lIdx + 1
                End If
                .MoveNext
            Loop
            If lIdx = 0 Then
                ReDim vRet(0)
                vRet(0) = STR_NO_USER_DBS
            Else
                ReDim Preserve vRet(lIdx - 1)
            End If
        End With
        EnumSqlDbAdo = vRet
        Exit Function
    End If
Next
'--- on failure -> return Array(0 To -1)
If Not IsArray(EnumSqlDbAdo) Then
    EnumSqlDbAdo = Split("")
End If

Exit Function
tratar_erro:
    If Err.Number = "-2147217843" Then Exit Function
    'MsgBox ("Descri��o do erro : " + Error()), vbCritical
    Exit Function
End Function
