Attribute VB_Name = "ORA"
Option Explicit
''''''''''''''''''''''''''''
' Oracle Objects for OLE global constant file.
' This file can be loaded into a code module.
''''''''''''''''''''''''''''

'Editmode property values
' These are intended to match similar constants in the
' Visual Basic file CONSTANT.TXT
Global Const ORADATA_EDITNONE = 0
Global Const ORADATA_EDITMODE = 1
Global Const ORADATA_EDITADD = 2

' Field Data Types
' These are intended to match similar constants in the
' Visual Basic file DATACONS.TXT
Global Const ORADB_BOOLEAN = 1
Global Const ORADB_BYTE = 2
Global Const ORADB_INTEGER = 3
Global Const ORADB_LONG = 4
Global Const ORADB_CURRENCY = 5
Global Const ORADB_SINGLE = 6
Global Const ORADB_DOUBLE = 7
Global Const ORADB_DATE = 8
Global Const ORADB_OBJECT = 9
Global Const ORADB_TEXT = 10
Global Const ORADB_LONGBINARY = 11
Global Const ORADB_MEMO = 12

'Parameter Types
Global Const ORAPARM_INPUT = 1
Global Const ORAPARM_OUTPUT = 2
Global Const ORAPARM_BOTH = 3

'Parameter Status
Global Const ORAPSTAT_INPUT = &H1&
Global Const ORAPSTAT_OUTPUT = &H2&
Global Const ORAPSTAT_AUTOENABLE = &H4&
Global Const ORAPSTAT_ENABLE = &H8&

'CreateDynaset Method Options
Global Const ORADYN_DEFAULT = &H0&
Global Const ORADYN_NO_AUTOBIND = &H1&
Global Const ORADYN_NO_BLANKSTRIP = &H2&
Global Const ORADYN_READONLY = &H4&
Global Const ORADYN_NOCACHE = &H8&
Global Const ORADYN_ORAMODE = &H10&
Global Const ORADYN_NO_REFETCH = &H20&
Global Const ORADYN_NO_MOVEFIRST = &H40&
Global Const ORADYN_DIRTY_WRITE = &H80&

'OpenDatabase Method Options
Global Const ORADB_DEFAULT = &H0&
Global Const ORADB_ORAMODE = &H1&
Global Const ORADB_NOWAIT = &H2&
Global Const ORADB_DBDEFAULT = &H4&
Global Const ORADB_DEFERRED = &H8&
Global Const ORADB_ENLIST_IN_MTS = &H10&

'Oracle type codes
Global Const ORATYPE_VARCHAR2 = 1
Global Const ORATYPE_NUMBER = 2
Global Const ORATYPE_SINT = 3
Global Const ORATYPE_FLOAT = 4
Global Const ORATYPE_STRING = 5
Global Const ORATYPE_DECIMAL = 7
Global Const ORATYPE_VARCHAR = 9
Global Const ORATYPE_DATE = 12
Global Const ORATYPE_REAL = 21
Global Const ORATYPE_DOUBLE = 22
Global Const ORATYPE_UNSIGNED8 = 23
Global Const ORATYPE_UNSIGNED16 = 25
Global Const ORATYPE_UNSIGNED32 = 26
Global Const ORATYPE_SIGNED8 = 27
Global Const ORATYPE_SIGNED16 = 28
Global Const ORATYPE_SIGNED32 = 29
Global Const ORATYPE_PTR = 32
Global Const ORATYPE_OPAQUE = 58
Global Const ORATYPE_UINT = 68
Global Const ORATYPE_RAW = 95
Global Const ORATYPE_CHAR = 96
Global Const ORATYPE_CHARZ = 97
Global Const ORATYPE_CURSOR = 102
Global Const ORATYPE_ROWID = 104
Global Const ORATYPE_MLSLABEL = 105
Global Const ORATYPE_OBJECT = 108
Global Const ORATYPE_REF = 110
Global Const ORATYPE_CLOB = 112
Global Const ORATYPE_BLOB = 113
Global Const ORATYPE_BFILE = 114
Global Const ORATYPE_CFILE = 115
Global Const ORATYPE_RSLT = 116
Global Const ORATYPE_NAMEDCOLLECTION = 122
Global Const ORATYPE_COLL = 122
Global Const ORATYPE_SYSFIRST = 228
Global Const ORATYPE_SYSLAST = 235
Global Const ORATYPE_OCTET = 245
Global Const ORATYPE_SMALLINT = 246
Global Const ORATYPE_VARRAY = 247
Global Const ORATYPE_TABLE = 248
Global Const ORATYPE_OTMLAST = 320
Global Const ORATYPE_RAW_BIN = 2000


'CreateSql Method options
Global Const ORASQL_DEFAULT = &H0&
Global Const ORASQL_NO_AUTOBIND = &H1&
Global Const ORASQL_FAILEXEC = &H2&
Global Const ORASQL_NONBLK = &H4&

'OraLob operation return codes
Global Const ORALOB_SUCCESS = 0
Global Const ORALOB_NEED_DATA = 99
Global Const ORALOB_NODATA = 100

'OraLob Write operation chunck  modes
Global Const ORALOB_ONE_PIECE = 0
Global Const ORALOB_FIRST_PIECE = 1
Global Const ORALOB_NEXT_PIECE = 2
Global Const ORALOB_LAST_PIECE = 3

'OraRef Lock operation
Global Const ORAREF_NO_LOCK = 1
Global Const ORAREF_EXCLUSIVE_LOCK = 2
Global Const ORAREF_NOWAIT_LOCK = 3

'OraRef Pin operaion
Global Const ORAREF_READ_ANY = 3
Global Const ORAREF_READ_RECENT = 4
Global Const ORAREF_READ_LATEST = 5

'OIP errors returned as part of the OLE Automation error.
Global Const OERROR_ADVISEULINK = 4096  ' Invalid advisory connection
Global Const OERROR_POSITION = 4098 ' Invalid database position
Global Const OERROR_NOFIELDNAME = 4099  ' Field 'field-name' not found
Global Const OERROR_TRANSIP = 4101  ' Transaction already in process
Global Const OERROR_TRANSNIPC = 4104    ' Commit detected with no active transaction
Global Const OERROR_TRANSNIPR = 4105    ' Rollback detected with no active transaction
Global Const OERROR_NODSET = 4106   ' No such set attached to connection
Global Const OERROR_INVROWNUM = 4108    ' Invalid row reference
Global Const OERROR_TEMPFILE = 4109 ' Error creating temporary file
Global Const OERROR_DUPSESSION = 4110   ' Duplicate session name
Global Const OERROR_NOSESSION = 4111    ' Session not found during detach
Global Const OERROR_NOOBJECTN = 4112    ' No such object named 'object-name'
Global Const OERROR_DUPCONN = 4113  ' Duplicate connection name
Global Const OERROR_NOCONN = 4114   ' No such connection during detach
Global Const OERROR_BFINDEX = 4115  ' Invalid field index
Global Const OERROR_CURNREADY = 4116    ' Cursor not ready for I/O
Global Const OERROR_NOUPDATES = 4117    ' Not an updatable set
Global Const OERROR_NOTEDITING = 4118   ' Attempt to update without edit or add operation
Global Const OERROR_DATACHANGE = 4119   ' Data has been modified
Global Const OERROR_NOBUFMEM = 4120 ' No memory for data transfer buffers
Global Const OERROR_INVBKMRK = 4121 ' Invalid bookmark
Global Const OERROR_BNDVNOEN = 4122 ' Bind variable not fully enabled
Global Const OERROR_DUPPARAM = 4123 ' Duplicate parameter name
Global Const OERROR_INVARGVAL = 4124    ' Invalid argument value
Global Const OERROR_INVFLDTYPE = 4125   ' Invalid field type
Global Const OERROR_TRANSFORUP = 4127   ' For Update detected with no active transaction
Global Const OERROR_NOTUPFORUP = 4128   ' For Update detected but not updatable set
Global Const OERROR_TRANSLOCK = 4129    ' Commit/Rollback with SELECT FOR UPDATE in progress
Global Const OERROR_CACHEPARM = 4130    ' Invalid cache parameter
Global Const OERROR_FLDRQROWID = 4131   ' Field processing requires ROWID
Global Const OERROR_OUTOFMEMORY = 4132  ' Internal Error
Global Const OERROR_MAXSIZE = 4135      ' Element size specified in AddTable exceeds the maximum allowed size for that variable type. See AddTable Method for more details.
Global Const OERROR_INVDIMENSION = 4136 ' Dimension specified in AddTable is invalid (i.e. negative). See AddTable Method for more details.
Global Const OERROR_MAXBUFFER = 4137    ' Buffer size for parameter array variable exceeds 32512 bytes (OCI limit).
Global Const OERROR_ARRAYSIZ = 4138 ' Dimensions of array parameters used in insert/update/delete statements are not equal.
Global Const OERROR_ARRAYFAILP = 4139   ' Error processing arrays. For details refer to OO4OERR.LOG in the windows directory.
Global Const OERROR_CREATEPOOL = 4147   ' Database Pool Already exists for this session.
Global Const OERROR_GETDB = 4148    ' Unable to obtain a free database object from the pool.

Global Const OERROR_NOOBJECT = 4796     'Creating Oracle object instance in client side object cache is failed
Global Const OERROR_BINDERR = 4797      'Binding  Oracle object instance to the SQL statement  is failed
Global Const OERROR_NOATTRNAME = 4798   'Getting attribute name of Oracle object instance is failed
Global Const OERROR_NOATTRINDEX = 4799  'Getting attribute index of Oracle object instance is failed
Global Const OERROR_INVINPOBJECT = 4801 'Invalid input object type for binding operation
Global Const OERROR_BAD_INDICATOR = 4802 'Fetched Oracle Object instance comes with invalid indicator structure
Global Const OERROR_OBJINSTNULL = 4803  'Operation on NULL Oracle object instance is failed. See IsNull property on OraObject
Global Const OERROR_REFNULL = 4804      'Pin Operation on NULL  Ref value is failed. See IsRefNull property on OraRef

Global Const OERROR_INVPOLLPARAMS = 4896 'Invalid  polling amount and chunksize specified for LOB read/write operation.
Global Const OERROR_INVSEEKPARAMS = 4897 'Invalid seek value is specified for LOB read/write operation.
Global Const OERROR_LOBREAD = 4898      'Read operation failed
Global Const OERROR_LOBWRITE = 4899     'Write operation failure
Global Const OERROR_INVCLOBBUF = 4900   'Input buffer type is not string for CLOB write operation
Global Const OERROR_INVBLOBBUF = 4901   'Input buffer type is not bytes for BLOB write operation
Global Const OERROR_INVLOBLEN = 4902    'Invalid buffer length for LOB write operation
Global Const OERROR_NOEDIT = 4903       'Write,Trim ,Append,Copy operation is allowed outside the dynaset edit
Global Const OERROR_INVINPUTLOB = 4904  'Invalid input LOB for bind operation
Global Const OERROR_NOEDITONCLONE = 4905 'Write,Trim,Append,Copy is not allowed for clone LOB object
Global Const OERROR_LOBFILEOPEN = 4906  'Specified file could not be opened in LOB operation
Global Const OERROR_LOBFILEIOERR = 4907 'File Read or Write failed in LOB Operation.
Global Const OERROR_LOBNULL = 4908    'Operation on NULL LOB has failed.

Global Const OERROR_AQCREATEERR = 4996    'Error creating AQ object
Global Const OERROR_MSGCREATEERR = 4997   'Error creating AQMsg object
Global Const OERROR_PAYLOADCREATEERR = 4998 ' Error creating Payload object
Global Const OERROR_MAXAGENTS = 4998       ' Maximum number of subscribers exceeded.
Global Const OERROR_AGENTCREATEERR = 5000  ' Error creating AQ Agent

Global Const OERROR_COLLINSTNULL = 5196 'Operation on NULL Oracle collection is  failed. See IsNull property on OraCollection
Global Const OERROR_NOELEMENT = 5197    'Element does not exist for given index
Global Const OERROR_INVINDEX = 5198     'Invalid collection index is specified
Global Const OERROR_NODELETE = 5199     'Delete operation is not supported for VARRAY collection type
Global Const OERROR_SAFEARRINVELEM = 5200  'Variant SafeArray cannot be created from the collection having non scalar element types

Global Const OERROR_NULLNUMBER = 5296   'Operation on NULL Oracle Number  is  failed.

' meta data type, OraMetaData.type returns one of the following
Global Const ORAMD_TABLE = 1
Global Const ORAMD_VIEW = 2
Global Const ORAMD_COLUMN = 3
Global Const ORAMD_COLUMN_LIST = 4
Global Const ORAMD_TYPE = 5
Global Const ORAMD_TYPE_ATTR = 6
Global Const ORAMD_TYPE_ATTR_LIST = 7
Global Const ORAMD_TYPE_METHOD = 8
Global Const ORAMD_TYPE_METHOD_LIST = 9
Global Const ORAMD_TYPE_ARG = 10
Global Const ORAMD_TYPE_RESULT = 11
Global Const ORAMD_PROC = 12
Global Const ORAMD_FUNC = 13
Global Const ORAMD_ARG = 14
Global Const ORAMD_ARG_LIST = 15
Global Const ORAMD_PACKAGE = 16
Global Const ORAMD_SUBPROG_LIST = 17
Global Const ORAMD_COLLECTION = 18
Global Const ORAMD_SYNONYM = 19
Global Const ORAMD_SEQENCE = 20
Global Const ORAMD_SCHEMA = 21
Global Const ORAMD_OBJECT_LIST = 22
Global Const ORAMD_SCHEMA_LIST = 23
Global Const ORAMD_DATABASE = 24

' AQ Options
' AQ Visible options
Global Const ORAAQ_ENQ_IMMEDIATE = 1
Global Const ORAAQ_ENQ_ON_COMMIT = 2

' AQ MessageID options
Global Const ORAAQ_NULL_MSGID = Null

' Selection Criteria for filtering messages
Global Const ORAAQ_ANY = 0
Global Const ORAAQ_CONSUMER = 1
Global Const ORAAQ_MSGID = 2

' Locking behaviour while dequeueing messages
Global Const ORAAQ_DQ_BROWSE = 1
Global Const ORAAQ_DQ_LOCKED = 2
Global Const ORAAQ_DQ_REMOVE = 3

' Message Position criteria for dequeuing
Global Const ORAAQ_DQ_FIRST_MSG = 1
Global Const ORAAQ_DQ_NEXT_TRANS = 2
Global Const ORAAQ_DQ_NEXT_MSG = 3

' Wait options for a dequeue operation
Global Const ORAAQ_DQ_WAIT_FOREVER = -1
Global Const ORAAQ_DQ_NOWAIT = 0


' Values of various OraAQMsg properties

' Number of Seconds to delay a newly enqueued message
' before it is available for dequeueing
Global Const ORAAQ_MSG_NO_DELAY = 0
' Prioirity values for messages
Global Const ORAAQ_MSG_PRIORITY_NORMAL = 0
Global Const ORAAQ_MSG_PRIORITY_HIGH = -10
Global Const ORAAQ_MSG_PRIORITY_LOW = 10

' Message Expiration in seconds
Global Const ORAAQ_MSG_NO_EXPIRE = 0
Global Const ORAAQ_MAX_AGENTS = 10

'Non Blocking return values
Global Const ORASQL_STILL_EXECUTING = -3123
Global Const ORASQL_SUCCESS = 0

' --------------------------------------------------
'  ここまでオラクル付属のファイルをコピー
'---------------------------------------------------

Public OraDB As OraDatabase 'oracle db object
Public OraSess As OraSession 'oracle session object
'Private Const cOracleSVName = "CM1"
'Private Const cOracleUidPwd = "cm1/cm1"

'複数工場接続対応 2009/05/28 SETsw kubota
Public OraDB_Other As OraDatabase 'oracle db object
Public OraSess_Other As OraSession 'oracle session object


' VAX用
Public VaxOraDB As OraDatabase 'oracle db object
Public VaxOraSess As OraSession 'oracle session object
Private Const cVaxOracleSVName = "vax"
Private Const cVaxOracleUidPwd = "vax/vax"


'概要      :Oracleのセッションを開く
'説明      :アプリケーションの起動時に呼ぶ
'履歴      :2001/06/04 作成  人見
Public Function OraDBOpen() As Boolean
    'Oracle Session Object
        Dim sDbName As String
    Dim sUID As String
    Dim sPWD As String
    
    Select Case gsFactryCd
    Case "10"               ''野田工場
        sDbName = "NODA"
        sUID = "oracle"
        sPWD = "oracle"
    Case "30"               ''生野工場
        sDbName = "IKNO"
        sUID = "oracle"
        sPWD = "oracle"
    Case "40"               ''米沢工場
        sDbName = "YONE"
        sUID = "oracle"
        sPWD = "oracle"
    Case "42"               '’３００ｍｍ
        sDbName = "cm1"
        sUID = "cm1"
        sPWD = "cm1"
    Case "43"               '’３００ｍｍ
        sDbName = "cmt"
        sUID = "cm1"
        sPWD = "cm1"
    Case "90"               ''テスト環境
        sDbName = "TEST0"
        sUID = "oracle"
        sPWD = "oracle"
    Case "91"               ''テスト環境(新) 2007/04/05追加 SETsw kubota
                            ''テスト環境(米沢) 2009/11/16追加 SSS.Marushita
        sDbName = "CLA0X"
        sUID = "oracle"
        sPWD = "oracle"
    Case "92"               ''テスト環境(生野) 2009/11/16追加 SSS.Marushita
        sDbName = "CLA0X"
        sUID = "oracle"
        sPWD = "oracle"
    Case "93"               ''テスト環境(生野A1) 2010/04/14追加 SETsw kubota
        sDbName = "CLA1"
        sUID = "oracle"
        sPWD = "oracle"
    Case "94"               ''テスト環境(尼崎A1) 2009/11/16追加 SSS.Marushita
        sDbName = "CLA1"
        sUID = "oracle"
        sPWD = "oracle"
    Case "99"               ''仮
        sDbName = "BOIS"
        sUID = "BOIS"
        sPWD = "BOIS"
    Case "AM"               ''尼崎工場 2009/06/02追加 SSS.Marushita
        sDbName = "CLK0"
        sUID = "oracle"
        sPWD = "oracle"
    Case Else               ''外販
        sDbName = "oracle"
        sUID = "oracle"
        sPWD = "oracle"
    End Select

    On Error GoTo ErrHandler
    Set OraSess = CreateObject("OracleInProcServer.XOraSession")
'   Set OraDB = OraSess.OpenDatabase(cOracleSVName, cOracleUidPwd, 0&)
    Set OraDB = OraSess.OpenDatabase(sDbName, sUID & "/" & sPWD, 0&)
    OraDBOpen = True
    Exit Function
ErrHandler:
    If Not OraSess Is Nothing Then
        Set OraSess = Nothing
    End If
    OraDBOpen = False
End Function

'概要      :VAX用 Oracleのセッションを開く
'説明      :アプリケーションの起動時に呼ぶ
'履歴      :2001/06/04 作成  人見
Public Function VaxOraDBOpen() As Boolean
    'Oracle Session Object
    On Error GoTo ErrHandler
    Set VaxOraSess = CreateObject("OracleInProcServer.XOraSession")
    Set VaxOraDB = VaxOraSess.OpenDatabase(cVaxOracleSVName, cVaxOracleUidPwd, 0&)
    VaxOraDBOpen = True
    Exit Function
ErrHandler:
    If Not VaxOraSess Is Nothing Then
        Set VaxOraSess = Nothing
    End If
    VaxOraDBOpen = False
End Function

'概要      :Oracleのセッションを閉じる
'説明      :アプリケーションの終了時に呼ぶ
'履歴      :2001/06/04 作成  人見
Public Sub OraDBClose()
    On Error Resume Next
    If Not OraDB Is Nothing Then
        OraDB.Close
        Set OraDB = Nothing
    End If
    If Not OraSess Is Nothing Then
        Set OraSess = Nothing
    End If
End Sub

'概要      :VAX用 Oracleのセッションを閉じる
'説明      :アプリケーションの終了時に呼ぶ
'履歴      :2001/06/04 作成  人見
Public Sub VaxOraDBClose()
    On Error Resume Next
    If Not VaxOraDB Is Nothing Then
        VaxOraDB.Close
        Set VaxOraDB = Nothing
    End If
    If Not VaxOraSess Is Nothing Then
        Set VaxOraSess = Nothing
    End If
End Sub

'概要      :Oracleのエラーを表示する
'説明      :
'履歴      :2001/06/04 作成  人見
Public Sub GetOraErr()
    On Error Resume Next
    If OraSess.LastServerErrText <> vbNullString Then
        MsgBox OraSess.LastServerErrText, vbOKOnly + vbCritical, "ERROR"
    End If
    If OraDB.LastServerErrText <> vbNullString Then
        MsgBox OraDB.LastServerErrText, vbOKOnly + vbCritical, "ERROR"
    End If
End Sub

'概要      :VAX用 Oracleのエラーを表示する
'説明      :
'履歴      :2001/06/04 作成  人見
Public Sub VaxGetOraErr()
    On Error Resume Next
    If Not IsNull(VaxOraSess.LastServerErrText) Then
        MsgBox VaxOraSess.LastServerErrText, vbOKOnly + vbCritical, "ERROR"
    End If
    If Not IsNull(VaxOraDB.LastServerErrText) Then
        MsgBox VaxOraDB.LastServerErrText, vbOKOnly + vbCritical, "ERROR"
    End If
End Sub

'概要      :Oracleのフィールド型名を得る
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                         ,説明
'          :typeNo        ,I  ,OracleInProcServer.vbTypes ,フィールド型
'          :戻り値        ,O  ,String                     ,フィールド型名
'説明      :
'履歴      :2001/06/08 作成  野村
'Public Function oraTypeName(typeNo As OracleInProcServer.vbTypes) As String
'    Select Case typeNo
'      Case ORADB_BOOLEAN
'        oraTypeName = "ORADB_BOOLEAN"
'      Case ORADB_BYTE
'        oraTypeName = "ORADB_BYTE"
'      Case ORADB_CURRENCY
'        oraTypeName = "ORADB_CURRENCY"
'      Case ORADB_DATE
'        oraTypeName = "ORADB_DATE"
'      Case ORADB_DOUBLE
'        oraTypeName = "ORADB_DOUBLE"
'      Case ORADB_INTEGER
'        oraTypeName = "ORADB_INTEGER"
'      Case ORADB_LONG
'        oraTypeName = "ORADB_LONG"
'      Case ORADB_LONGBINARY
'        oraTypeName = "ORADB_LONGBINARY"
'      Case ORADB_MEMO
'        oraTypeName = "ORADB_MEMO"
'      Case ORADB_OBJECT
'        oraTypeName = "ORADB_OBJECT"
'      Case ORADB_SINGLE
'        oraTypeName = "ORADB_SINGLE"
'      Case ORADB_TEXT
'        oraTypeName = "ORADB_TEXT"
'      Case Else
'        oraTypeName = "UnKnown"
'    End Select
'End Function


Public Function oraGetSysdate() As Date
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    sql = "select SYSDATE from TBCMB003"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    oraGetSysdate = rs("SYSDATE")
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If
End Function


'>>>>> 複数工場接続対応 2009/05/28追加 SETsw kubota ------------------------

'///////////////////////////////////////////////////
' @(f)
' 機能    : 起動工場ではない他の工場のＤＢにコネクトする
'
' 返り値  : 正常 - true
'           異常 - false
'
' 引き数  : sFactryCd - 工場コード
'
' 機能説明: ＤＢにコネクトする
'           ｺﾈｸﾄ先は、関数に渡された引数の工場ｺｰﾄﾞにより換える
'
'///////////////////////////////////////////////////
Public Function OraDBOpen_Other(ByVal sFactryCd As String) As Boolean
    Dim sDbName As String
    Dim sUID As String
    Dim sPWD As String
    
    Call GetConnectStr(sFactryCd, sDbName, sUID, sPWD)
    
    On Error GoTo ErrHandler
    
    ''オラクル接続
    Set OraSess_Other = CreateObject("OracleInProcServer.XOraSession")
    Set OraDB_Other = OraSess_Other.OpenDatabase(sDbName, sUID & "/" & sPWD, 0&)
    
    OraDBOpen_Other = True
    Exit Function

ErrHandler:
    If Not OraSess_Other Is Nothing Then
        Set OraSess_Other = Nothing
    End If
    OraDBOpen_Other = False

End Function

'///////////////////////////////////////////////////
' @(f)
' 機能    :ＤＢ開放
'
' 返り値  : 正常 - true
'           異常 - false
'
' 機能説明: ＤＢ開放
'
'///////////////////////////////////////////////////
Public Function OraDBClose_Other() As Boolean
    
    On Error GoTo ErrProc
    
    ''オラクル切断
    OraDB_Other.Close
    
    ''解放
    Set OraDB_Other = Nothing
    Set OraSess_Other = Nothing
    
    OraDBClose_Other = True
    Exit Function
    
ErrProc:
    OraDBClose_Other = False
End Function

'///////////////////////////////////////////////////
' @(f)
' 機能    : 引数工場コードのDB接続文字列を取得
' 返り値  : なし
' 引き数  : sFactryCd - 工場コード
' 　　　　  sDbName - DB名
' 　　　　  sUID - ユーザ
' 　　　　  sPWD - パスワード
' 機能説明:
'///////////////////////////////////////////////////
Public Sub GetConnectStr(ByVal sFactryCd As String _
                       , ByRef sDbName As String _
                       , ByRef sUID As String _
                       , ByRef sPWD As String _
                       )

    Select Case sFactryCd
    Case "10"               ''野田工場
        sDbName = "NODA"
        sUID = "oracle"
        sPWD = "oracle"
    Case "30"               ''生野工場
        sDbName = "IKNO"
        sUID = "oracle"
        sPWD = "oracle"
    Case "AM"               ''尼崎工場
        sDbName = "CLK0"
        sUID = "oracle"
        sPWD = "oracle"
    Case "40"               ''米沢工場
        sDbName = "YONE"
        sUID = "oracle"
        sPWD = "oracle"
    Case "42"               '’３００ｍｍ
        sDbName = "cm1"
        sUID = "cm1"
        sPWD = "cm1"
    Case "43"               '’３００ｍｍ
        sDbName = "cmt"
        sUID = "cm1"
        sPWD = "cm1"
    Case "44"               ''米沢工場(CLE5) Add 2011/01/18
        sDbName = "CLE5"
        sUID = "oracle"
        sPWD = "oracle"
    Case "50"               ''千歳工場
        sDbName = "CLF0"    ''2011/01/18　SUMCO Akizuki 追加
        sUID = "oracle"
        sPWD = "oracle"
    Case "80"               ''SPTI(インドネシア)
        sDbName = "CLW0"    ''2011/01/19　SUMCO Akizuki 追加
        sUID = "oracle"
        sPWD = "oracle"
    Case "90"               ''テスト環境
        sDbName = "TEST0"
        sUID = "oracle"
        sPWD = "oracle"
    Case "91"               ''テスト環境(新)
                            ''テスト環境(米沢) 2009/11/16追加 SSS.Marushita
        sDbName = "CLA0X"
        sUID = "oracle"
        sPWD = "oracle"
    Case "92"               ''テスト環境(生野) 2009/11/16追加 SSS.Marushita
        sDbName = "CLA0X"
        sUID = "oracle"
        sPWD = "oracle"
    Case "93"               ''テスト環境(生野A1) 2010/04/14追加 SETsw kubota
        sDbName = "CLA1"
        sUID = "oracle"
        sPWD = "oracle"
    Case "94"               ''テスト環境(尼崎A1) 2009/11/16追加 SSS.Marushita
        sDbName = "CLA1"
        sUID = "oracle"
        sPWD = "oracle"
    Case "95"               ''DWH            Add 2009/05/28
        sDbName = "DWH"
        sUID = "DWHMGR"
        sPWD = "DWHMGR"
    Case "99"               ''仮
        sDbName = "BOIS"
        sUID = "BOIS"
        sPWD = "BOIS"
    Case Else               ''外販
        sDbName = "oracle"
        sUID = "oracle"
        sPWD = "oracle"
    End Select

End Sub

'///////////////////////////////////////////////////
' @(f)
' 機能    :オラクルダイナセットの作成
'
' 返り値  : 正常 - true
'           異常 - false
'
' 引き数  : ARG1 - ダイナセットセットオブジェクト
'           ARG2 - SQL文
'           ARG3 - ダイナセットオプション
'
' 機能説明: オラクルダイナセット作成
'
'///////////////////////////////////////////////////
Public Function DynSet_Other(ByRef objOraDynaset As Object, sSqlStmt As String, Optional vOpt = &H4&) As Boolean
    On Error GoTo DynErr
    
    ''オラクルダイナセット作成
    Set objOraDynaset = OraDB_Other.CreateDynaset(sSqlStmt, vOpt)
    DynSet_Other = True
    Exit Function
    
DynErr:
    DynSet_Other = False
End Function

'///////////////////////////////////////////////////
' @(f)
' 機能    : ＳＱＬ文実行
'
' 返り値  : 0以上:処理件数
'           　-1：異常
'
' 引き数  : ARG1 - SQL文
'
' 機能説明: ＳＱＬ文実行し、処理件数を返す
'
'///////////////////////////////////////////////////
Public Function SqlExec_Other(sSqlStmt As String) As Long
    On Error GoTo ErrProc
    
    ''オラクルＳＱＬ実行
    SqlExec_Other = OraDB_Other.DbExecuteSQL(sSqlStmt)
    
    Exit Function
    
ErrProc:
    SqlExec_Other = -1
End Function

'<<<<< 複数工場接続対応 2009/05/28追加 SETsw kubota ------------------------

