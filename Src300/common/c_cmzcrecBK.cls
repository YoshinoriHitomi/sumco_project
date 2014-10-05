VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_cmzcrec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "RecordSetの必要な情報を保持するクラス"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"c_cmzcflds"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'ﾌﾟﾛﾊﾟﾃｨ値を保持するためのﾛｰｶﾙ変数。
Private mvarTableName As String 'ﾛｰｶﾙ ｺﾋﾟｰ
Private mvarFields As c_cmzcflds


Public Function Clone() As c_cmzcrec
Dim CloneTo As c_cmzcrec
Dim fld As c_cmzcfld

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcrec.cls -- Function Clone"

    Set CloneTo = New c_cmzcrec
    CloneTo.TABLENAME = mvarTableName
    For Each fld In mvarFields
        With fld
            CloneTo.Fields.Add .FieldName, .Value, .FieldType, .FldSize, .FldScale
        End With
    Next
    
    Set Clone = CloneTo
    Set CloneTo = Nothing

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Set Clone = Nothing
    Resume proc_exit
End Function

Public Sub SetRecDefault()
Attribute SetRecDefault.VB_Description = "該当テーブルの全フィールドについて、未設定なら既定値を設定する"
Dim rs As OraDynaset
Dim FldName As String
Dim fld As OraField
Dim dt As Variant


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcrec.cls -- Sub SetRecDefault"

    If mvarFields Is Nothing Then
        Set mvarFields = New c_cmzcflds
    End If
    Set rs = OraDB.CreateDynaset("select * from " & mvarTableName, ORADYN_NO_BLANKSTRIP)
    For Each fld In rs.Fields
        ''そのフィールドが未登録なら、既定値で登録する
        If mvarFields.Exist(fld.Name) = False Then
            Select Case fld.TYPE
              Case ORADB_BOOLEAN
                dt = 0
              Case ORADB_DATE
                If fld.Name = "OTHERTIME" Then
                    dt = ""
                Else
                    dt = Format$(Date, "yy-mm-dd")
                End If
              Case ORADB_SINGLE, ORADB_DOUBLE
                dt = 0#
              Case ORADB_BYTE, ORADB_INTEGER, ORADB_LONG, ORADB_CURRENCY
                dt = 0
              Case ORADB_TEXT, ORADB_MEMO
                If (fld.Name = "SENDFLAG") Or (fld.Name = "SUMMITSENDFLAG") Then
                    dt = "0"
                Else
                    dt = " "
                End If
              Case Else
                dt = "UnKnown"
            End Select
            mvarFields.Add fld.Name, dt, fld.TYPE, fld.Size, fld.OraScale
        End If
    Next
    rs.Close

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub



Public Function SqlInsert() As String
Attribute SqlInsert.VB_Description = "保存されているフィールド内容を元に insert用のsqlを生成する"
Dim sqlTop As String
Dim sqlVal As String
Dim fld As c_cmzcfld
Dim s As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcrec.cls -- Function SqlInsert"

    If mvarFields.COUNT = 0 Then
        SqlInsert = vbNullString
        GoTo proc_exit
    End If
    
    sqlTop = "insert into " & mvarTableName & " ("
    sqlVal = "values ("
    For Each fld In mvarFields
        With fld
            sqlTop = sqlTop & .FieldName & ", "
            If IsNull(.Value) = True Then
                sqlVal = sqlVal & "NULL, "
            Else
                Select Case .FieldType
                  Case ORADB_BOOLEAN
                        sqlVal = sqlVal & .Value & ", "
                  Case ORADB_BYTE, ORADB_CURRENCY, ORADB_INTEGER, ORADB_LONG
                        sqlVal = sqlVal & .Value & ", "
                  Case ORADB_SINGLE, ORADB_DOUBLE
                        sqlVal = sqlVal & .Value & ", "
                  Case ORADB_TEXT, ORADB_MEMO
                        s = Trim$(.Value)
                        If (.FieldName = "SENDFLAG") Or (.FieldName = "SUMMITSENDFLAG") Then
                            If s = vbNullString Then
                                sqlVal = sqlVal & "'0', "
                            Else
                                sqlVal = sqlVal & "'" & s & "', "
                            End If
                        Else
                            If s = vbNullString Then
                                sqlVal = sqlVal & "' ', "
                            Else
                                sqlVal = sqlVal & "'" & s & "', "
                            End If
                        End If
                  Case ORADB_DATE
                        Select Case .FieldName
                          Case "REGDATE", "UPDDATE", "SENDDATE"
                                sqlVal = sqlVal & "SYSDATE, "
                          Case Else
                                sqlVal = sqlVal & "'" & Format$(.Value, "yy-mm-dd") & "', "
                        End Select
                
                  Case Else
                        Debug.Print "c_cmzcrec: SqlInsert(" & mvarTableName & "." & .FieldName & ") -- 未対応のOracle型"
                End Select
            End If
            If .FieldName = "EPDSETCH" Then Debug.Print s

        End With
    Next
    Mid$(sqlTop, Len(sqlTop) - 1, 1) = ")"
    Mid$(sqlVal, Len(sqlVal) - 1, 1) = ")"
    SqlInsert = sqlTop & sqlVal

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

Public Function SqlUpdate(orgRec As c_cmzcrec, ParamArray keys() As Variant) As String
Attribute SqlUpdate.VB_Description = "変更部分を探索し、update用のsqlを生成する"
Dim sqlTop As String
Dim sqlVal As String
Dim sqlWhere As String
Dim fld As c_cmzcfld
Dim i As Integer
Dim keyMax As Integer
Dim isKey As Boolean
Dim fldValue As String
Dim s As String
Dim TEST As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcrec.cls -- Function SqlUpdate"

    If mvarFields.COUNT = 0 Then
        SqlUpdate = vbNullString
        GoTo proc_exit
    End If
    
    keyMax = UBound(keys)
    If keyMax = 0 Then
            SqlUpdate = vbNullString
        GoTo proc_exit
    End If
    
    ''UPDDATEフィールドの「入力値」を更新する
    mvarFields.Add "UPDDATE", Date, ORADB_DATE '値が変わってさえいれば、後でSYSDATEに置き換わる
    
    sqlTop = "update " & mvarTableName
    sqlVal = " set "
    sqlWhere = " where "
    For Each fld In mvarFields
        With fld
            ''フィールドの値をSQL文字列の形式に整形する
            If IsNull(.Value) = True Then
                fldValue = "NULL "
            Else
                Select Case .FieldType
                    Case ORADB_BOOLEAN
                        fldValue = .Value
                    Case ORADB_BYTE, ORADB_CURRENCY, ORADB_INTEGER, ORADB_LONG
                        fldValue = .Value
                    Case ORADB_SINGLE, ORADB_DOUBLE
                        fldValue = .Value
                    Case ORADB_TEXT, ORADB_MEMO
                        s = Trim$(.Value)
                        If (.FieldName = "SENDFLAG") Or (.FieldName = "SUMMITSENDFLAG") Then
                            If s = vbNullString Then
                                fldValue = "'0'"
                            Else
                                fldValue = "'" & s & "'"
                            End If
                        Else
                            If s = vbNullString Then
                                fldValue = "' '"
                            Else
                                fldValue = "'" & s & "'"
                            End If
                        End If
                    Case ORADB_DATE
                        Select Case .FieldName
                            Case "REGDATE", "UPDDATE", "SENDDATE"
                                fldValue = "SYSDATE"
                            Case Else
                                fldValue = "'" & Format$(.Value, "yy-mm-dd") & "'"
                        End Select
            
                    Case Else
                        Debug.Print "c_cmzcrec: SqlUpdate(" & mvarTableName & "." & .FieldName & ") -- 未対応のOracle型"
                End Select
           End If
         ''keyフィールドかを調べる
            isKey = False
            For i = 0 To keyMax
                If .FieldName = keys(i) Then
                    isKey = True
                    Exit For
                End If
            Next
        
            If isKey Then
                ''keyフィールドならwhere句に
                sqlWhere = sqlWhere & "(" & .FieldName & " = " & fldValue & ") and "
            Else
                ''keyフィールドでなければset句に(ただし、未変更のフィールドは無視する)
                If orgRec.Fields.Exist(.FieldName) Then
                    If CStr(.Value) <> CStr(orgRec(.FieldName)) Then
                        sqlVal = sqlVal & .FieldName & "=" & fldValue & ", "
                    End If
                Else
                    sqlVal = sqlVal & .FieldName & "=" & fldValue & ", "
                End If
            End If
        End With
    Next
    sqlWhere = Left$(sqlWhere, Len(sqlWhere) - Len(" and "))    '最後の「 and 」を消す
    sqlVal = Left$(sqlVal, Len(sqlVal) - Len(", "))             '最後の「, 」を消す
    
    If Len(sqlVal) < Len(" set ") Then
        SqlUpdate = vbNullString                '変更されていなかった
    Else
        SqlUpdate = sqlTop & sqlVal & sqlWhere  '通常時
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function



'概要      :フィールドコレクション(Get)
'説明      :
'履歴      :2001/06/08 作成  野村
Public Property Get Fields() As c_cmzcflds
Attribute Fields.VB_UserMemId = 0

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcrec.cls -- Property Get Fields"

    If mvarFields Is Nothing Then
        Set mvarFields = New c_cmzcflds
    End If


    Set Fields = mvarFields

proc_exit:
    '終了
    gErr.Pop
    Exit Property

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Property


'概要      :フィールドコレクション(Set)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型         ,説明
'          :vData         ,   ,c_cmzcflds ,
'説明      :
'履歴      :2001/06/08 作成  野村
Public Property Set Fields(vData As c_cmzcflds)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcrec.cls -- Property Set Fields"

    Set mvarFields = vData

proc_exit:
    '終了
    gErr.Pop
    Exit Property

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Property



'概要      :Oracleのレコードセットからデータを一括コピーする
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型         ,説明
'          :TbName        ,I  ,String     ,テーブル名
'          :rs            ,I  ,OraDynaset ,レコードセット
'説明      :
'履歴      :2001/06/08 作成  野村
Public Sub CopyFromRs(TbName$, rs As OraDynaset)
Attribute CopyFromRs.VB_Description = "RecordSet の内容を一括コピーする"
Dim i As Integer
Dim fldCnt As Integer


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcrec.cls -- Sub CopyFromRs"

    TABLENAME = TbName
    Set mvarFields = New c_cmzcflds
    fldCnt = rs.Fields.COUNT
    For i = 0 To fldCnt - 1
        mvarFields.Add rs.FieldName(i), rs.Fields(i), rs.Fields(i).TYPE, rs.Fields(i).Size, rs.Fields(i).OraScale
    Next

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub



'概要      :Oracleのレコードセットにデータを一括コピーする
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型         ,説明
'          :rs            ,O  ,OraDynaset ,レコードセット
'説明      :
'履歴      :2001/06/12 作成  野村
Public Sub CopyToRs(rs As OraDynaset)
Dim fld As c_cmzcfld


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcrec.cls -- Sub CopyToRs"

    For Each fld In mvarFields
        rs(fld.FieldName) = fld.Value
    Next

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub



Private Sub Class_Initialize()
    Set mvarFields = New c_cmzcflds
End Sub

'概要      :Teriminate時処理
'説明      :
'履歴      :2001/06/08 作成  野村
Private Sub Class_Terminate()

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcrec.cls -- Sub Class_Terminate"

  Set mvarFields = Nothing

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub



'概要      :TableNameプロパティ(Let)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :vData         ,I  ,String    ,テーブル名
'説明      :
'履歴      :2001/06/08 作成  野村
Public Property Let TABLENAME(ByVal vData As String)
Attribute TABLENAME.VB_Description = "テーブル名"
'ﾌﾟﾛﾊﾟﾃｨに値を代入するときに、代入式の左辺で使用します。
'Syntax: X.TableName = 5

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcrec.cls -- Property Let TableName"

    mvarTableName = Trim$(vData)

proc_exit:
    '終了
    gErr.Pop
    Exit Property

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Property


'概要      :TableNameプロパティ(Get)
'説明      :
'履歴      :2001/06/08 作成  野村
Public Property Get TABLENAME() As String
'ﾌﾟﾛﾊﾟﾃｨの値を取得するときに、代入式の右辺で使用します。
'Syntax: Debug.Print X.TableName

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcrec.cls -- Property Get TableName"

    TABLENAME = mvarTableName

proc_exit:
    '終了
    gErr.Pop
    Exit Property

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Property



