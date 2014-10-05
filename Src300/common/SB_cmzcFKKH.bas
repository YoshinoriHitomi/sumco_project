Attribute VB_Name = "SB_cmzcFKKH"
Option Explicit
'                                     2003/09/01
'======================================================
' 振替可能候補品番
' 概要    : 振替元品番より振替可能候補品番を一覧表示し、
'           振替先品番として決定する。
' 参照    :
'======================================================

Public FKKH_MotoHinban As String            ' 振替元品番
Public FKKH_Proccd As String                ' 工程コード
Public FKKH_Crynum As String                ' 結晶番号

Public FKKH_SakiHinban As String            ' 振替先品番


'======================================================
' 特採番号入力
' 概要    : 特採番号と特採理由の入力を行う。
' 参照    :
'======================================================
Public TBN_MotoHinban As String             ' 振替元品番
Public TBN_SakiHinban As String             ' 振替先品番

Public TBN_Bangou As String                 ' 特採番号
Public TBN_Riyuu As String                  ' 特採理由
Public TBN_Msg As String                    ' 特採時メッセージ
