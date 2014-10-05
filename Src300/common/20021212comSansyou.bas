Attribute VB_Name = "kpkcommon"
' 生産管理用
Public SeisanOraDB As OraDatabase 'oracle db object
Public SeisanOraSess As OraSession 'oracle session object

'TextComboButtonChengeのTYPE定義
Public Const TXT_CHENGE As Long = 1     '対象はテキストのみ
Public Const COM_CHENGE As Long = 2     '対象はコンボのみ
Public Const BTN_CHENGE As Long = 4     '対象はボタンのみ
Public Const T_C_CHENGE As Long = TXT_CHENGE + COM_CHENGE   '対象はテキスト、コンボ
Public Const T_B_CHENGE As Long = TXT_CHENGE + BTN_CHENGE   '対象はテキスト、ボタン
Public Const C_B_CHENGE As Long = COM_CHENGE + BTN_CHENGE   '対象はコンボ、ボタン
Public Const T_C_B_CHENGE As Long = TXT_CHENGE + COM_CHENGE + BTN_CHENGE  '対象はテキスト、コンボ、ボタン

Public l_OverDay As Long '10000以上のデータがある場合に加算する
' @(f)
' 機能    : <開始日と時間><日締めする日付と時間>を返す
'
' 返り値  : 正常=True　異常=False
'
' 引き数:
'           sStDate  : 開始日
'           sEndDate : 終了日
'
Public Function HidimeKeisan(sStDate As String, SENDDATE As String) As Boolean
    Dim bErrFlg As Boolean
    Dim sTime As String
    Dim sWk_Date As String
    Dim dateWk   As Date
    Dim lNumber As Long
    
    lNumber = 86400 - 1  '２４時間（秒）－１秒
    bErrFlg = True
        
    sTime = F_DbConectEndTime("X", "80", "1")
    sStDate = Mid(sStDate, 1, 4) & "/" & Mid(sStDate, 5, 2) & "/" & Mid(sStDate, 7, 2) & " " & sTime
    SENDDATE = Mid(SENDDATE, 1, 4) & "/" & Mid(SENDDATE, 5, 2) & "/" & Mid(SENDDATE, 7, 2) & " " & sTime
    dateWk = SENDDATE
    SENDDATE = DateAdd("S", lNumber, dateWk)

    HidimeKeisan = bErrFlg
End Function

' @(f)
' 機能    : 日締めする時間をTBから抜き出す
'
' 返り値  : "7:00:00"or 各TBデータ
'
' 引き数:   s_SYSCA9   : SYSCA9の呼出し条件
'           s_SHUCA9   : SHUCA9の呼出し条件
'           s_CODEA9   : CODEA9の呼出し条件
'
' 機能説明: Table:KODA9から条件呼出
'
Public Function F_DbConectEndTime(s_SYSCA9 As String, _
                                  s_SHUCA9 As String, _
                                  s_CODEA9 As String) As String
  Dim dynOraDyn As OraDynaset
  Dim wk_koteiCdName As String
  Dim s_SQL As String
  Dim i_Lp As Integer
  
  '初期値
  F_DbConectEndTime = "07:00:00"
  
            s_SQL = "SELECT"
    s_SQL = s_SQL + " KCODE01A9 "
    s_SQL = s_SQL + "FROM"
    s_SQL = s_SQL + " KODA9 "
    s_SQL = s_SQL + "WHERE"
    s_SQL = s_SQL + " SYSCA9 = '" + s_SYSCA9 + "' AND"
    s_SQL = s_SQL + " SHUCA9 = '" + s_SHUCA9 + "' AND"
    s_SQL = s_SQL + " CODEA9 = '" + s_CODEA9 + "'"
    
    'オラクル接続
    If DynSet2(dynOraDyn, s_SQL) = False Then
    
        ''ダイナセット作成失敗
        Call MsgOut(100, "", ERR_DISP_LOG, "KODA9")
    Else
        If (dynOraDyn(0).Value = "") Or _
           (IsNull(dynOraDyn(0).Value) = True) Or _
           (IsEmpty(dynOraDyn(0).Value) = True) Then
           Exit Function
        Else
            F_DbConectEndTime = dynOraDyn(0).Value
        End If
    End If
            
End Function
' @(f)
' 機能    : ComBoxにDBから取得した値を入れる。
'
' 返り値  : True ＞ 正常   False ＞　異常
'
' 引き数:   ComBoxName : 書き込むコンボBox   Null値不可
'           s_SYSCA9   : SYSCA9の呼出し条件  NULL値不可
'           s_SHUCA9   : SHUCA9の呼出し条件  NULL値不可
'           s_CODEA9   : CODEA9の呼出し条件  省略可
'           s_CTR01A9  : CTR01A9の呼出し条件 省略可
'
' 機能説明: Table:KODA9からNAMESJA9を条件で呼出
'
Public Function F_DbConectAddComItems(ComBoxName As ComboBox, _
                                        s_SYSCA9 As String, _
                                        s_SHUCA9 As String, _
                                        Optional s_CODEA9 As String, _
                                        Optional s_CTR01A9 As String) As Boolean
  Dim dynOraDyn As OraDynaset
  Dim wk_koteiCdName As String
  Dim s_SQL As String
  Dim i_Sec As Integer
  Dim i_Lp  As Integer
    
    F_DbConectAddComItems = False

    ComBoxName.Clear

            s_SQL = "SELECT"
    s_SQL = s_SQL + " CODEA9,"
    s_SQL = s_SQL + " NAMESJA9,"
    s_SQL = s_SQL + " KCODE01A9 "
    s_SQL = s_SQL + "FROM"
    s_SQL = s_SQL + " KODA9 "
    s_SQL = s_SQL + "WHERE"
    s_SQL = s_SQL + " SYSCA9 = '" + s_SYSCA9 + "' AND"
    s_SQL = s_SQL + " SHUCA9 = '" + s_SHUCA9 + "'"
    
    If s_CODEA9 <> "" Then _
        s_SQL = s_SQL + " AND CODEA9 = '" + s_CODEA9 + "'"
    
    If s_CTR01A9 <> "" Then _
        s_SQL = s_SQL + " AND CTR01A9 = '" + s_CTR01A9 + "'"
    
    'オラクル接続
    If DynSet2(dynOraDyn, s_SQL) = False Then
    ''ダイナセット作成失敗
        Call MsgOut(100, "", ERR_DISP_LOG, "TBCMB002")
        Exit Function
    End If
            
    'コンボボックスに項目を表示する
    i_Sec = 0
    While dynOraDyn.EOF = False
        If IsNull(dynOraDyn(0)) = False Then
            wk_koteiCdName = ""
            wk_koteiCdName = dynOraDyn(0).Value & " " & dynOraDyn(1).Value
            ComBoxName.AddItem wk_koteiCdName
        End If
        If (dynOraDyn(2) = "1") And (i_Sec = 0) Then
            i_Sec = i_Lp
        End If
        i_Lp = i_Lp + 1
        dynOraDyn.DbMoveNext
    Wend
    
    ComBoxName.Tag = i_Sec
    
    F_DbConectAddComItems = True
    
End Function
' @(f)
' 機能    : メインフォームのコンボBoxがNull値だった場合、取得したIndexを表示する
'
' 返り値  : なし
'
' 引き数 : コンボBox
'
Public Sub F_ComboIndex(ComBoxName As ComboBox)
    With ComBoxName
        If (.Enabled = True) Then
            If .Text = "" Then
                .ListIndex = .Tag
            End If
        End If
    End With
End Sub


' @(f)
'
' 機能      : 開始日から終了日までの時間をストリングで返す
'
' 返り値    :　"0000:00.0"
'
'       StatDay ： 開始日（Date型に変換できる形式）
'       EndDay  ： 終了日（Date型に変換できる形式）
'
Function F_TimeStatEnd(StatDay, EndDay) As String
  Dim l_Day  As Long    '日付
  Dim d_Time As Double  '時間
  Dim d_Min  As Double
  Dim s_Day  As String
  Dim i_Canma As Integer
    
    F_TimeStatEnd = "0000$00.0"
    
    '日が無い場合は処理を抜ける
    If StatDay = "" Or EndDay = "" Then _
        Exit Function
    
    '開始日より終了日が後ろになった場合は処理を抜ける
    If CDate(StatDay) > CDate(EndDay) Then Exit Function
    
    d_Min = DateDiff("n", StatDay, EndDay)
    '1日以上の場合
    If d_Min >= 1440 Then
        s_Day = CStr((d_Min / 60) / 24)
        i_Canma = InStr(1, s_Day, ".")
        If i_Canma <> 0 Then
            s_Day = Left(s_Day, i_Canma)
        End If
        l_Day = CInt(s_Day)
    Else
        l_Day = 0
        d_Time = d_Min / 60
    End If
    
    '小数点以下第一位まで表示する
    d_Time = CInt(d_Time * 10)
    d_Time = d_Time * 0.1
            
    F_TimeStatEnd = F_DayTimeAr(l_Day, d_Time)

End Function
' @(f)
'
' 機能      : 合計時間から平均を割り出す
'
' 返り値    :　"0000:00.0"   データが無い "0000$00.0"
'
'       SumTime  ： 合計時間(0000:00.0)
'       i_AvgCnt ： 割数 (整数)
'
Function F_DayTimeAvg(SumTime As String, i_AvgCnt As Integer, Optional OverDay As Long) As String
  Dim l_Day   As Long
  Dim l_Day_1 As Double
  Dim d_Time  As Double
  Dim d_Time2 As Double
  Dim d_Time3 As Double
  Dim i_Time  As Integer
  
    'データが無いので処理を行わない
    If SumTime = "0000$00.0" Then Exit Function
    If SumTime = "" Then Exit Function
    
    '日付の平均
    l_Day_1 = Left(SumTime, 4) / i_AvgCnt
    If l_Day_1 >= 1 Then
        l_Day = l_Day_1
    Else
        l_Day = 0
    End If
    '日付の小数点部を時間換算
    d_Time3 = (l_Day_1 - l_Day) * 24
    If d_Time3 < 0 Then
        l_Day = l_Day - 1
        d_Time3 = (l_Day_1 - (l_Day)) * 24
    End If
    '時間の計算
    d_Time = (CDbl(Right(SumTime, 4)) / i_AvgCnt) + d_Time3
    
    If d_Time >= 24 Then
        l_Day = (l_Day + 1) / i_AvgCnt
        d_Time = d_Time - 24
    End If
    
    '小数点以下第一位まで表示する
    i_Time = CInt(d_Time * 10)
    d_Time = i_Time * 0.1
    
    If OverDay <> 0 Then _
        l_Day = l_Day + (OverDay * 10000)
    
    F_DayTimeAvg = F_DayTimeAr(l_Day, d_Time)

End Function
' @(f)
'
' 機能      : 合計を算出
'
' 返り値    :　"0000:00.0"
'       Sumtime_1 ： 合計する二つの数の一つ'
'       Sumtime_2 ： 合計する二つの数の一つ
'
Function F_TimeSum(SumTime_1 As String, SumTime_2 As String) As String
  Dim l_Day      As Long
  Dim i_DaySub   As Long
  Dim d_Time     As Double
  Dim d_TimeSub  As Double
  
    If SumTime_1 <> "" Then
        l_Day = CInt(Left(SumTime_1, 4))
        d_Time = CDbl(Right(SumTime_1, 4))
    Else
        l_Day = 0
        d_Time = 0
    End If
    i_DaySub = CInt(Left(SumTime_2, 4))
    d_TimeSub = CDbl(Right(SumTime_2, 4))
    
    l_Day = l_Day + i_DaySub
    d_Time = d_Time + d_TimeSub
    
    If d_Time > 24 Then
        d_Time = d_Time - 24
        l_Day = l_Day + 1
    End If
    
    '9999日を上回ったデータはエラーを表示する(最大で２５万件以上）
    If l_Day > 10000 Then
        l_OverDay = l_OverDay + 1
        l_Day = l_Day - 10000
        Exit Function
    End If

    d_Time = CInt(d_Time * 10) * 0.1
        
    F_TimeSum = F_DayTimeAr(l_Day, d_Time)

End Function
' @(f)
'
' 機能      : 表示形式変換
'
' 返り値    :　"####:##.#"
'
Function F_DayTimeAr(l_Day As Long, d_Time As Double) As String
  Dim s_Sp1 As String
  Dim s_Sp2 As String
  Dim s_Sp3 As String

    F_DayTimeAr = "0000:00.0"
    
    If InStr(1, CStr(d_Time), ".") = 0 Then
        s_Sp3 = ".0"
    End If

    '日スペース
    Select Case l_Day
      Case 0 To 9: s_Sp1 = "000"
      Case 10 To 99: s_Sp1 = "00"
      Case 100 To 999: s_Sp1 = "0"
      Case 1000 To 9999: s_Sp1 = ""
    End Select
    '時間スペース
    Select Case d_Time
      Case 0 To 9.9: s_Sp2 = "0"
      Case 10 To 24: s_Sp2 = ""
    End Select
    
    F_DayTimeAr = s_Sp1 & CStr(l_Day) & ":" & s_Sp2 & CStr(d_Time) & s_Sp3

End Function
' @(f)
'
' 機能      : 表示形式変換
'
' 返り値    :　"###日:##.#"
'
Function F_DispDayTime(s_DayTime As String) As String
  Dim s_Sp1 As String
  Dim s_Sp2 As String
  Dim s_Sp3 As String
  Dim l_Day As Long
  Dim d_Time As Double

    F_DispDayTime = "  0日 0.0"
    If s_DayTime = "" Then Exit Function
    
    l_Day = CInt(Left(s_DayTime, 4))
    d_Time = CDbl(Right(s_DayTime, 4))
    
    
    If InStr(1, CStr(d_Time), ".") = 0 Then
        s_Sp3 = ".0"
    End If

    '日スペース
    Select Case l_Day
      Case 0 To 9: s_Sp1 = "  "
      Case 10 To 99: s_Sp1 = " "
      Case 100 To 999: s_Sp1 = ""
    End Select
    '時間スペース
    Select Case d_Time
      Case 0 To 9.9: s_Sp2 = " "
      Case 10 To 24: s_Sp2 = ""
    End Select
    
    F_DispDayTime = s_Sp1 & CStr(l_Day) & "日" & s_Sp2 & CStr(d_Time) & s_Sp3

End Function
' @(f)
'
' 機能      : 表示形式変換
'
' 返り値    :　"###日##.#" → 分 "######.#"
'
Function F_ReTime(s_DayTime As String) As String
  Dim l_Day  As Long
  Dim d_Time As Double
    
    l_Day = Left(s_DayTime, 3)
    d_Time = Right(s_DayTime, 4)

    If l_Day > 1 Then
        l_Day = l_Day * 24
    End If
    
    F_ReTime = l_Day + d_Time

End Function


' @(f)
'
' 機能      : 長さから重量を求める計算処理
'
' 返り値    :　重量
'
Function fncNagaWeightChg(lNagasa As Long) As Long
    fncNagaWeightChg = (301 / 2) ^ 2 * 3.1416 * lNagasa * 2.33 / 1000
End Function

' @(f)
'
' 機能      : 生産管理ＤＢ ＯＰＥＮ
'
' 返り値    :　重量
'

Public Function OraDBSeisanOpen() As Boolean
    'Oracle Session Object
        Dim sDbName As String
    Dim sUID As String
    Dim sPWD As String
    
'    Select Case gsFactryCd
'    Case "42"               '’３００ｍｍ
        sDbName = "cp1"
        sUID = "cp1"
        sPWD = "cp1"
'    End Select

    On Error GoTo ErrHandler
    Set SeisanOraSess = CreateObject("OracleInProcServer.XOraSession")
    Set SeisanOraDB = SeisanOraSess.OpenDatabase(sDbName, sUID & "/" & sPWD, 0&)
    OraDBSeisanOpen = True
    Exit Function
ErrHandler:
    If Not SeisanOraSess Is Nothing Then
        Set SeisanOraSess = Nothing
    End If
    OraDBSeisanOpen = False
End Function

'概要      :Oracleのセッションを閉じる(生産管理ＤＢ)
'説明      :アプリケーションの終了時に呼ぶ
'履歴      :
Public Sub OraSeisanDBClose()
    On Error Resume Next
    If Not SeisanOraDB Is Nothing Then
        SeisanOraDB.Close
        Set SeisanOraDB = Nothing
    End If
    If Not SeisanOraSess Is Nothing Then
        Set SeisanOraSess = Nothing
    End If
End Sub

'///////////////////////////////////////////////////
' @(f)
' 機能    :オラクルダイナセットの作成(生産管理ＤＢ)
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
Public Function DynSetSeisan(ByRef objOraDynaset As Object, sSqlStmt As String, Optional vOpt = &H4&) As Boolean
    On Error GoTo DynErr
    
    ''オラクルダイナセット作成
    Set objOraDynaset = SeisanOraDB.CreateDynaset(sSqlStmt, vOpt)
    DynSetSeisan = True
    Exit Function
    
DynErr:
    DynSetSeisan = False
End Function

'///////////////////////////////////////////////////
' @(f)
' 機能    : フォーム上のコントロールを使用不可にする
'
' 返り値  :
'
' 引き数  : フォーム
'        ： 対象(lType)  1:テキスト 2:コンボボックス 4:ボタン
'        ： 格納値(bSet) true/false
'
' 機能説明: 指定したフォームの「[Ｆ１]ﾒｲﾝﾒﾆｭｰ」ボタン以外
'           のコントロールを使用不可にする
'
'///////////////////////////////////////////////////
Public Sub TextComboButtonChenge(frmForm As Form, lType As Long, bSet As Boolean)
    Dim iIdx As Integer
    Dim ctlControl As Control
    
    
    ''フォーム上のコントロールを全て使用不可にする
    For Each ctlControl In frmForm.Controls
        If TypeOf ctlControl Is TextBox Then
            If (lType And TXT_CHENGE) = TXT_CHENGE Then
                ctlControl.Enabled = bSet
            End If
        ElseIf TypeOf ctlControl Is ComboBox Then
            If (lType And COM_CHENGE) = COM_CHENGE Then
                ctlControl.Enabled = bSet
            End If
        ElseIf TypeOf ctlControl Is CommandButton Then
            If (lType And BTN_CHENGE) = BTN_CHENGE Then
                ctlControl.Enabled = bSet
            End If
        End If
    Next ctlControl
    
    ''「[Ｆ１]ﾒｲﾝﾒﾆｭｰ」ボタンを使用可能にする
    If ((lType And BTN_CHENGE) = BTN_CHENGE) Then
        frmForm.cmdF(1).Enabled = True
    End If
End Sub


