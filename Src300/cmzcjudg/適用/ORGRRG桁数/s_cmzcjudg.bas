Attribute VB_Name = "s_cmzcjudg"
Option Explicit

'Public Enum FUNCTION_RETURN                 ''関数の戻り値
'    FUNCTION_RETURN_SUCCESS = 0             '' 正常
'    FUNCTION_RETURN_FAILURE = -1            '' 異常
'End Enum

Public Const EZJ00 = "EZJ00" ''測定対象コード %s には対応してません。
Public Const ZJ001 = "ZJ001" ''処理方法コード %s は、無効です。
Public Const ZJ002 = "ZJ002" ''測定点数異常でBMD最大値が求められません。

Public Const JUDG_OK = True                 ''判定結果がOKの場合帰す値
Public Const JUDG_NG = False                ''判定結果がNGの場合帰す値

Public Const OI_JUDG = "酸素濃度判定"       ''判定項目文字列(Oi)
Public Const CS_JUDG = "炭素濃度判定"       ''判定項目文字列(Cs)
Public Const RES_JUDG = "比抵抗判定"        ''判定項目文字列(比抵抗)
Public Const GFA_JUDG = "GFA判定"           ''判定項目文字列(GFA)
Public Const BMD_JUDG = "BMD判定"           ''判定項目文字列(BMD)
Public Const OSF_JUDG = "OSF判定"           ''判定項目文字列(OSF)
Public Const DEN_JUDG = "DEN判定"           ''判定項目文字列(Den)
Public Const LDL_JUDG = "L/DL判定"          ''判定項目文字列(L/DL)
Public Const DVD2_JUDG = "DVD2判定"         ''判定項目文字列(DVD2)
Public Const LT_JUDG = "ライフタイム判定"   ''判定項目文字列(ライフタイム)
Public Const EPD_JUDG = "EPD判定"           ''判定項目文字列(EPD)
Public Const DOI_JUDG = "ΔOi判定"          ''判定項目文字列(ΔOi)
Public Const DZ_JUDG = "DZ判定"             ''判定項目文字列(DZ)
Public Const DSOD_JUDG = "DSOD判定"         ''判定項目文字列(DSOD)
Public Const SPV_JUDG = "SPV判定"           ''判定項目文字列(SPV)
Public Const AOI_JUDG = "AOi判定"           ''判定項目文字列(AOi)　03/12/09 ooba

Public Const WFRES_JUDG = 1                 ''判定識別フラグ(RES)
Public Const WFOI_JUDG = 2                  ''判定識別フラグ(Oi)
Public Const WFDOI_JUDG = 3                 ''判定識別フラグ(ΔOi)
Public Const WFOSF_JUDG = 4                 ''判定識別フラグ(OSF)
Public Const WFBMD_JUDG = 5                 ''判定識別フラグ(BMD)
Public Const WFDZ_JUDG = 6                  ''判定識別フラグ(DZ)
Public Const WFDSOD_JUDG = 7                ''判定識別フラグ(DSOD)
Public Const WFSPV_JUDG = 8                 ''判定識別フラグ(SPV)
Public Const WFAOI_JUDG = 9                 ''判定識別フラグ(AOi)　03/12/09 ooba

Public Const ObjCode01 = "1"                ''中心測定値
Public Const ObjCode02 = "2"                ''測定値の中央値
Public Const ObjCode03 = "3"                ''全測定点
Public Const ObjCode04 = "6"                ''R/2
Public Const ObjCode05 = "A"                ''全点の平均値
Public Const ObjCode06 = "B"                ''全点の最大値
Public Const ObjCode07 = "C"                ''全点の平均値と最大値
Public Const ObjCode08 = "D"                ''全点の最小値
Public Const ObjCode09 = "E"                ''内周部2点、外周部2点(5点測定で1,2,4,5)
Public Const ObjCode10 = "F"                ''MAX(2,4点目)
Public Const ObjCode11 = "G"                ''MAX(2,3,4点目)
Public Const ObjCode12 = "H"                ''個数保証
Public Const ObjCode13 = "N"                ''狙い
Public Const ObjCode14 = "Z"                ''形状測定(平坦度、反返り、WARP)
Public Const ObjCode15 = " "                ''規格なし
Public Const ObjCode16 = "K"                ''2001/09/19 S.Sano 全点の最小値と最大値
Public Const ObjCode17 = "L"                ''AVE+MIN　08/03/13 ooba
Public Const ObjCode18 = "7"                ''AVE+外周1点   '' 2008/10/20 BMD評価,外周1点保証機能追加 ADD By Systech
Public Const ObjCodeGrp01 = "1236"          ''FTIR、GFA、WF比抵抗、WF酸素濃度、ΔOi
'-----TEST2004/10
Public Const ObjCodeGrp05 = "1236N"          ''比抵抗
'2001/09/19 S.SanoPublic Const ObjCodeGrp02 = "ABCDFG"        ''BMD、DZ

'' 2008/10/20 BMD評価,外周1点保証機能追加 UPD By Systech Start
''Public Const ObjCodeGrp02 = "ABCDFGK"        ''2001/09/19 S.Sano BMD、DZ
Public Const ObjCodeGrp02 = "ABCDFGK7"      ''BMD、DZ
'' 2008/10/20 BMD評価,外周1点保証機能追加 UPD By Systech End

Public Const ObjCodeGrp03 = "ABC"           ''OSF
Public Const ObjCodeGrp04 = "3"             ''DSOD、SPV
Public Const ObjCodeGrp06 = "ABCDFGK3"        ''2004/12/15 S.Sano BMD、DZ

Public Const PosCode01 = "E"                ''
Public Const PosCode02 = "G"                ''
Public Const PosCode03 = "H"                ''
Public Const PosCode04 = "J"                ''
Public Const PosCode05 = "M"                ''
Public Const PosCode06 = "Q"                ''
Public Const PosCode07 = "N"                ''
Public Const PosCode08 = "P"                ''
Public Const PosCode09 = "R"                ''
Public Const PosCodeGrp01 = "EGHJMQNPRR"    ''

Public Const JudgCodeC01 = "H"              ''結晶判定有り
Public Const JudgCodeC02 = "BS X"           ''結晶判定無し、全てOK
Public Const JudgCodeW01 = "H"              ''WF判定有り
Public Const JudgCodeW02 = "BS H"           ''WF判定無し、全てOK
Public Const JudgCodeW03 = "XS"             ''測定有り
'----- TEST2004/10
Public Const KSTAFF_J002 = "SHIJI"                    '

'Add Start 2011/01/26 SMPK Miyata
''Cu-deco 製品仕様パターン区分
Public Const CudecoSpcPtnNR = "1"           '' リング無し指定
Public Const CudecoSpcPtnND = "2"           '' ディスク無し指定
Public Const CudecoSpcPtnNP = "3"           '' パターン無し指定
Public Const CudecoSpcPtnN = "4"            '' 不問 (選択なし)
Public Const CudecoSpcPtnNB = "5"           '' バンド無し指定
Public Const CudecoSpcPtnNPB = "6"          '' Pバンド無し指定
Public Const CudecoSpcPtnNBB = "7"          '' Bバンド指定無し
''Cu-deco 実績パターン区分
Public Const CudecoJskPtnN = "0"            '' None
Public Const CudecoJskPtnR = "1"            '' Ring
Public Const CudecoJskPtnD = "2"            '' Disk
Public Const CudecoJskPtnDR = "3"           '' Disk & Ring
Public Const CudecoJskPtnPB_B = "5"         '' PB-band
Public Const CudecoJskPtnP_B = "6"          '' P-band
Public Const CudecoJskPtnB_B = "7"          '' B-band
'Add End   2011/01/26 SMPK Miyata


'品質保証情報構造体
Public Type Guarantee
    cMeth  As String                        ''測定位置_方
    cCount As String                        ''測定位置_点
    cPos   As String                        ''測定位置_位(OSFの場合 領)
    cObj   As String                        ''保証方法_対
    cJudg  As String                        ''保証方法_処
'C－OSF3判定機能追加 2007/04/23 M.Kaga START   ---
'    cJudg2  As String                       ''保証方法_処
'C－OSF3判定機能追加 2007/04/23 M.Kaga END     ---
    cBunp  As String                        ''面内分布計算式    ' Res，Oi 面内分布計算式追加依頼  No.030205  yakimura  2003.06.06
End Type
'エラー情報構造体
Public Type ERROR_INFOMATION
    ErrCode     As Variant                  ''エラーコード
    ErrStr(4)   As Variant                  ''オプション文字列
End Type

'概要      :配列内の最小値を求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :d()           ,I  ,double    ,測定値
'          :戻り値        ,O  ,double    ,最大値
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function JudgMin(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim temp As Double
    Dim High As Integer
    
    If iMax > 0 Then
        High = iMax - 1
    Else '' 配列の上限を取得します。
        High = UBound(d)
    End If
    
    temp = d(0)
    For c0 = 1 To High
        If d(c0) <> -1 Then
            If d(c0) < temp Then
                temp = d(c0)
            End If
        End If
    Next
    
    JudgMin = temp
End Function

'概要      :配列内の最大値を求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :d()           ,I  ,double    ,測定値
'          :戻り値        ,O  ,double    ,最大値
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function JudgMax(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim temp As Double
    Dim High As Integer
    
    If iMax > 0 Then
        High = iMax - 1
    Else '' 配列の上限を取得します。
        High = UBound(d)
    End If
    
    temp = d(0)
    For c0 = 1 To High
        If d(c0) <> -1 Then
            If d(c0) > temp Then
                temp = d(c0)
            End If
        End If
    Next
    
    JudgMax = temp
End Function

'概要      :配列内データの平均値を求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :d()           ,I  ,double    ,測定値
'          :戻り値        ,O  ,double    ,平均値
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function JudgAve(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim c1 As Integer
    Dim temp As Double
    Dim High As Integer
    
    If iMax > 0 Then
        High = iMax - 1
    Else '' 配列の上限を取得します。
        High = UBound(d)
    End If
    
    temp = 0
    c1 = 0
    For c0 = 0 To High
        If d(c0) <> -1 Then
            c1 = c1 + 1
            temp = temp + d(c0)
        End If
    Next
    If c1 = 0 Then
        JudgAve = 0
    Else
        JudgAve = temp / c1
    End If
End Function

'概要      :配列内データの平均値を求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :d()           ,I  ,double    ,測定値
'          :戻り値        ,O  ,double    ,平均値
'説明      :配列内データがすべてNULL(-1)の場合-1を返す。
'履歴      :新規作成 2005/06/22 ffc)tanabe
Public Function JudgAve3(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim c1 As Integer
    Dim temp As Double
    Dim High As Integer
    
    If iMax > 0 Then
        High = iMax - 1
    Else '' 配列の上限を取得します。
        High = UBound(d)
    End If
    
    temp = 0
    c1 = 0
    For c0 = 0 To High
        If d(c0) <> -1 Then
            c1 = c1 + 1
            temp = temp + d(c0)
        End If
    Next
    If c1 = 0 Then
        JudgAve3 = -1
    Else
        JudgAve3 = temp / c1
    End If
End Function

'概要      :配列内データの中央値を求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :d()           ,I  ,double    ,測定値
'          :戻り値        ,O  ,double    ,中央値
'説明      :中央値とは、最大値と最小値の中心値に最も近い値。
'履歴      :2001/06/06 佐野 信哉 作成
Public Function JudgCenter(d() As Double) As Double
'住友流
    Dim High As Integer
    Dim temp() As Double
    Dim c0 As Integer
    Dim c1 As Integer
    
    c1 = 0
    For c0 = 0 To UBound(d)
        If d(c0) <> -1 Then
            ReDim Preserve temp(c1) As Double
            temp(c1) = d(c0)
            c1 = c1 + 1
        End If
    Next
    
    If c1 <> 0 Then
        BubbleSort temp()
        
        High = UBound(temp)
        JudgCenter = temp(Int((High + 1) / 2))
    Else
        JudgCenter = -9999
    End If

'三菱流
'    Dim c0 As Integer
'    Dim temp As double
'    Dim temp1 As double
'    Dim temp2 As double
'    Dim Center As double
'    Dim High As Integer
'
'    '' 配列の上限を取得します。
'    High = UBound(d)
'
'    ''最大値を求める。
'    temp1 = JudgMax(d())
'    ''現在最少の絶対値を求めた測定値とする。
'    temp2 = temp1
'
'    ''中心値を求める。
'    Center = (JudgMin(d()) + temp1) / 2
'
'    ''最大値と中心値の絶対値を求める。
'    ''現在最少の絶対値とする。
'    temp1 = Abs(temp1)
'
'    ''最大値と最小値に最も近い測定値を求める。
'    For c0 = 0 To High
'        If d(c0) <> -1 Then
'            ''測定値と中心値の絶対値を求める。
'            temp = Abs(Center - d(c0))
'            ''前回求めた絶対値と今回求めた絶対値を比較し、
'            ''今回求めた絶対値が小さかった場合。
'            If temp < temp1 Then
'                ''今回求めた絶対値を現在最少の絶対値とする。
'                temp1 = temp
'                ''現在最少の絶対値を求めた測定値とする。
'                temp2 = d(c0)
'            End If
'        End If
'    Next
'
'    ''中心値との絶対値が最も小さい測定値を返す。
'    JudgCenter = temp2
End Function

'概要      :配列のコピーを作成する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :d()           ,I  ,double    ,測定値
'          :d1()          ,O  ,double    ,測定値のコピー
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Sub DataCopy(d() As Double, d1() As Double)
    Dim High As Integer
    Dim c0 As Integer
    
    '' 配列の上限を取得します。
    High = UBound(d)
    
    ''第2引数に第1引数をコピーします。
    For c0 = 0 To High
        d1(c0) = d(c0)
    Next
End Sub

'概要      :バブルソートを行います。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :d()           ,IO ,double    ,測定値のコピー
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Sub BubbleSort(d() As Double)
    Dim High As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim temp As Double
    '' 配列の上限を取得します。
    High = UBound(d)
    
    '' 配列の個々の要素に対して繰り返し処理します。
    For c0 = 0 To High - 1
        '' 配列の個々の要素に対して繰り返し処理します。
        For c1 = c0 + 1 To High
            '' 配列の前方にある値が、配列の後方にある値より
            '' 大きい場合には、それらを交換します。
            If d(c0) > d(c1) Then
                temp = d(c0)
                d(c0) = d(c1)
                d(c1) = temp
            End If
        Next
    Next
End Sub


'概要      :Sideデータの平均値を求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :d()           ,I  ,double    ,測定値
'          :戻り値        ,O  ,double    ,平均値
'説明      :
'履歴      :2003/06/06 yakimura 作成
Public Function JudgSideAve(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim cnt As Integer
    Dim temp As Double
    Dim High As Integer
    
    If iMax > 0 Then
        High = iMax - 1
    Else '' 配列の上限を取得します。
        High = UBound(d)
    End If
    
    temp = 0
    cnt = 0
    For c0 = 1 To High
        If d(c0) <> -1 Then
            cnt = cnt + 1
            temp = temp + d(c0)
        End If
    Next
    If cnt = 0 Then
        JudgSideAve = 0
    Else
        JudgSideAve = temp / cnt
    End If
End Function

'概要      :配列内(0～2)の最小値を求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :d()           ,I  ,double    ,測定値
'          :戻り値        ,O  ,double    ,最大値
'説明      :
'履歴      :2003/06/06  yakimura  作成
Public Function JudgMin2(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim temp As Double
    Dim High As Integer
    
    temp = d(0)
    For c0 = 1 To 2
        If d(c0) <> -1 Then
            If d(c0) < temp Then
                temp = d(c0)
            End If
        End If
    Next
    
    JudgMin2 = temp
End Function

'概要      :配列内(0～2)の最大値を求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :d()           ,I  ,double    ,測定値
'          :戻り値        ,O  ,double    ,最大値
'説明      :
'履歴      :2003/06/06  yakimura  作成
Public Function JudgMax2(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim temp As Double
    Dim High As Integer
    
    temp = d(0)
    For c0 = 1 To 2
        If d(c0) <> -1 Then
            If d(c0) > temp Then
                temp = d(c0)
            End If
        End If
    Next
    
    JudgMax2 = temp
End Function

'概要      :配列内(0～2)データの平均値を求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :d()           ,I  ,double    ,測定値
'          :戻り値        ,O  ,double    ,平均値
'説明      :
'履歴      :2003/06/06  yakimura  作成
Public Function JudgAve2(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim c1 As Integer
    Dim temp As Double
    Dim High As Integer
    
    temp = 0
    c1 = 0
    For c0 = 0 To 2
        If d(c0) <> -1 Then
            c1 = c1 + 1
            temp = temp + d(c0)
        End If
    Next
    If c1 = 0 Then
        JudgAve2 = 0
    Else
        JudgAve2 = temp / c1
    End If
End Function

'概要      :Side(1～2)データの平均値を求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :d()           ,I  ,double    ,測定値
'          :戻り値        ,O  ,double    ,平均値
'説明      :
'履歴      :2003/06/06 yakimura 作成
Public Function JudgSideAve2(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim cnt As Integer
    Dim temp As Double
    Dim High As Integer
    
    temp = 0
    cnt = 0
    For c0 = 1 To 2
        If d(c0) <> -1 Then
            cnt = cnt + 1
            temp = temp + d(c0)
        End If
    Next
    If cnt = 0 Then
        JudgSideAve2 = 0
    Else
        JudgSideAve2 = temp / cnt
    End If
End Function

'概要      :Side(1～4)の最小値を求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :d()           ,I  ,double    ,測定値
'          :戻り値        ,O  ,double    ,最大値
'説明      :
'履歴      :2003/06/06  yakimura  作成
Public Function JudgSideMin(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim temp As Double
    Dim High As Integer
    
    temp = d(1)
    For c0 = 2 To 4
        If d(c0) <> -1 Then
            If d(c0) < temp Then
                temp = d(c0)
            End If
        End If
    Next
    
    JudgSideMin = temp
End Function

'概要      :Side(1～4)の最大値を求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :d()           ,I  ,double    ,測定値
'          :戻り値        ,O  ,double    ,最大値
'説明      :
'履歴      :2003/06/06  yakimura  作成
Public Function JudgSideMax(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim temp As Double
    Dim High As Integer
    
    temp = d(1)
    For c0 = 2 To 4
        If d(c0) <> -1 Then
            If d(c0) > temp Then
                temp = d(c0)
            End If
        End If
    Next
    
    JudgSideMax = temp
End Function


'概要      :エラー情報構造体に値を代入する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :[Code]        ,I  ,String           ,オプションエラーコード
'          :[Str1]        ,I  ,String           ,オプション文字列１
'          :[Str2]        ,I  ,String           ,オプション文字列２
'          :[Str3]        ,I  ,String           ,オプション文字列３
'          :[Str4]        ,I  ,String           ,オプション文字列４
'          :[Str5]        ,I  ,String           ,オプション文字列５
'          :戻り値        ,O  ,FUNCTION_RETURN  ,FUNCTION_RETURN_FAILURE
'説明      :必ず関数異常終了コードを返す。
'　　      :オプション引数は、省略した場合、""が代入される。
'　　      :オプション引数を、全て省略した場合、エラー情報構造体の初期化が行われる。
'　　      :オプション引数を、全て省略した場合、戻り値は、無視する。
'履歴      :2001/06/06 佐野 信哉 作成
Public Function SetErrInfo(ErrInfo As ERROR_INFOMATION, Optional CODE, Optional Str1, Optional Str2, Optional Str3, Optional Str4, Optional Str5) As FUNCTION_RETURN
    ErrInfo.ErrCode = CODE
    ErrInfo.ErrStr(0) = Str1
    ErrInfo.ErrStr(1) = Str2
    ErrInfo.ErrStr(2) = Str3
    ErrInfo.ErrStr(3) = Str4
    ErrInfo.ErrStr(4) = Str5
    SetErrInfo = FUNCTION_RETURN_FAILURE
End Function

'概要      :測定値の範囲判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :JudgData      ,I  ,double    ,測定値
'          :SpecMin       ,I  ,double    ,下限値
'          :SpecMax       ,I  ,double    ,上限値
'          :戻り値        ,O  ,Boolean   ,JUDG_OK or JUDG_NG
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function RangeDecision(JudgData As Double, SpecMin As Double, SpecMax As Double) As Boolean
    RangeDecision = ((JudgData >= SpecMin) And (JudgData <= SpecMax))
End Function

'概要      :OSFのパターンの判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :Kubun         ,I  ,string    ,パターン区分
'          :JudgData()    ,I  ,string    ,パターン実績
'          :戻り値        ,O  ,Boolean   ,JUDG_OK or JUDG_NG
'説明      :
'履歴      :2003/05/17　ooba
Public Function JudgPattern(KUBUN As String, JudgData() As String * 1) As Boolean
    Dim ct As Integer
    Dim RD As String
    
    '【パターン区分】　1：リング無し　2：ディスク無し　3：パターン無し　4：不問
    Select Case KUBUN
        Case "1"
            RD = "D "
        Case "2"
            RD = "R "
        Case "3"
            RD = " "
        Case "4", " "
            RD = "RD "
    End Select
    
    'パターン区分に該当する文字列とパターン実績を比べ、判定を行う。
    For ct = 0 To 2
        JudgPattern = (InStr(RD, JudgData(ct)) > 0)
        If JudgPattern = False Then
            Exit For
        End If
    Next
End Function

'概要      :DSODのﾊﾟﾀｰﾝの判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :Kubun         ,I  ,string    ,ﾊﾟﾀｰﾝ区分
'          :JudgData()    ,I  ,string    ,ﾊﾟﾀｰﾝ実績
'          :戻り値        ,O  ,Boolean   ,JUDG_OK or JUDG_NG
'説明      :
'履歴      :2004/07/28　ooba
Public Function JudgDsodPattern(KUBUN As String, JudgData() As String * 3) As Boolean

    Dim iCnt As Integer
    Dim sPtn As String
    
    'DSODﾊﾟﾀｰﾝ実績結果に対して、仕様上のﾊﾟﾀｰﾝ区分ｺｰﾄﾞと判定を行う。
    For iCnt = 0 To 1
    
        JudgDsodPattern = False
        sPtn = Trim(JudgData(iCnt))
        
        Select Case KUBUN
            Case "1"        'ﾘﾝｸﾞ無し
                If sPtn = "" Or sPtn = "D" Then
                    JudgDsodPattern = True
                End If
            Case "2"        'ﾃﾞｨｽｸ無し
                If sPtn = "" Or sPtn = "R" Then
                    JudgDsodPattern = True
                End If
            Case "3"        'ﾊﾟﾀｰﾝ無し
                If sPtn = "" Then
                    JudgDsodPattern = True
                End If
            Case "4", " "   '不問
                If sPtn = "" Or sPtn = "R" Or sPtn = "D" Or sPtn = "R,D" Then
                    JudgDsodPattern = True
                End If
        End Select
        If JudgDsodPattern = False Then
            Exit For
        End If
    Next
    
End Function

'概要      :BMDの面内分布計算(面内分布"P")を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :d()           ,I  ,double    ,測定値
'          :戻り値        ,O  ,double    ,平均値
'説明      :
'履歴      :2003/05/21　ooba
Public Function JudgBmdMBP(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim side As Double
    Dim center As Double
    Dim High As Integer
    Dim deverrflag As Boolean
    
    deverrflag = False
    JudgBmdMBP = -1
    
    If iMax > 0 Then
        High = iMax - 1
    Else '' 配列の上限を取得します。
        High = UBound(d)
    End If
    
    '配列の最初の値をsideにセット
    side = d(0)
    '配列の最後の値をcenterにセット
    For c0 = 1 To High
        If d(c0) <> -1 Then
            center = d(c0)
        End If
    Next
    
    'sideとsenterを比べ、大きい方を分子へ
    If side < center Then
        If side > 0 Then
            JudgBmdMBP = center / side * 100
        ElseIf side = 0 Then
            deverrflag = True
        End If
    Else
        If center > 0 Then
            JudgBmdMBP = side / center * 100
        ElseIf center = 0 Then
            deverrflag = True
        End If
    End If
    JudgBmdMBP = Round(JudgBmdMBP, 1)
    
    'デバッグ用処理をコメント　2003/05/29 ooba
'    If deverrflag Then
'        WFCJudgDialog.WFCErrorMessage "分布計算 0 除算エラー"
'    ElseIf JudgBmdMBP = -1 Then
'        WFCJudgDialog.WFCErrorMessage "測定位置、対象データ、分布計算矛盾"
'    End If
End Function
'概要      :RRG、ORGを求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :d()           ,I  ,double    ,測定値
'          :iMax          ,I  ,Integer   ,測定点数
'          :戻り値        ,O  ,double    ,RRG,ORG
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function RGCal(d() As Double, iMax As Integer) As Double
    Dim temp As Double
    Dim High As Integer
    
    '' 配列の上限を取得します。
    High = UBound(d)
    If High < iMax - 1 Then
        RGCal = -1
        Exit Function
    End If
    
    temp = JudgMin(d(), iMax)
    If temp <> 0 Then
        temp = (JudgMax(d(), iMax) - temp) * 100 / temp
    End If
    
    RGCal = temp
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'概要      :比抵抗，酸素濃度の面内計算値を求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :戻り値        ,O  ,double    ,Mennai
'説明      :
'履歴      :2003/06/06  yakimura  作成
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'端数丸めを行う関数と行わない関数の二つに分離 2011/11/25 SETsw kubota
'Public Function MENNAI_Cal(JudgFlag As String, R() As Double, G As Guarantee, calcode As String) As Double
Public Function MENNAI_Cal_NotRound(JudgFlag As String, R() As Double, G As Guarantee, calcode As String) As Double

Dim Min        As Double
Dim max        As Double
Dim AVE        As Double
Dim center     As Double
Dim side       As Double
Dim Side_Ave   As Double
Dim Mennai     As Double
Dim w_Max1     As Double
Dim w_Max2     As Double
Dim w_N1       As Double
Dim w_N2       As Double
Dim w_N3       As Double
Dim w_N4       As Double
Dim w_Nx       As Double
    
    Mennai = -1
    
    Select Case calcode

    Case "A"                '---> (max-min)/min×100
        Select Case JudgFlag
           Case RES_JUDG ''判定識別フラグ(RES)
              
              Min = Min_Cal(RES_JUDG, R(), G)
              max = Max_Cal(RES_JUDG, R(), G)
           
           Case OI_JUDG ''判定識別フラグ(Oi)
              
              Min = Min_Cal(OI_JUDG, R(), G)
              max = Max_Cal(OI_JUDG, R(), G)
        
        End Select
        
        If (Min <> -9999) And (max <> -9999) Then
            If (Min <> 0) Then
                Mennai = (max - Min) * 100 / Min
            Else
                Mennai = -1
            End If
        End If

    Case "B"                '---> (max-min)/max×100
        Select Case JudgFlag
           Case RES_JUDG ''判定識別フラグ(RES)
              
              Min = Min_Cal(RES_JUDG, R(), G)
              max = Max_Cal(RES_JUDG, R(), G)
           
           Case OI_JUDG ''判定識別フラグ(Oi)
              
              Min = Min_Cal(OI_JUDG, R(), G)
              max = Max_Cal(OI_JUDG, R(), G)
        
        End Select
        
        If (Min <> -9999) And (max <> -9999) Then
            If (max <> 0) Then
                Mennai = (max - Min) * 100 / max
            Else
                Mennai = -1
            End If
        Else
            Mennai = -1
        End If

    Case "C"                '---> (max-min)/center×100
        Select Case JudgFlag
           Case RES_JUDG ''判定識別フラグ(RES)
              
              Min = Min_Cal(RES_JUDG, R(), G)
              max = Max_Cal(RES_JUDG, R(), G)
           
           Case OI_JUDG ''判定識別フラグ(Oi)
              
              Min = Min_Cal(OI_JUDG, R(), G)
              max = Max_Cal(OI_JUDG, R(), G)
        
        End Select
        
        center = R(0)
        If (Min <> -9999) And (max <> -9999) And (center <> -9999) Then
            If (center <> 0) Then
                Mennai = (max - Min) * 100 / center
            Else
                Mennai = -1
            End If
        Else
            Mennai = -1
        End If

    Case "D"                '---> |center-side|max/center×100
        
        Select Case JudgFlag
        Case RES_JUDG ''判定識別フラグ(RES)
           
           center = R(0)
           Select Case G.cPos
           Case "1", "G", "H", "J", "M"

'特殊対応　2003.08.21 yakimura start
              If G.cCount = 2 Then
                 side = R(1)
                 If (side <> -1) And (center <> -1) Then
                     If (center <> 0) Then
                         Mennai = Abs(center - side) * 100 / center
                     Else
                         Mennai = -1
                     End If
                 Else
                     Mennai = -1
                 End If
              End If
'特殊対応　2003.08.21 yakimura end
              
              If G.cCount = 3 Then
                 side = R(2)
                 If (side <> -1) And (center <> -1) Then
                     If (center <> 0) Then
                         Mennai = Abs(center - side) * 100 / center
                     Else
                         Mennai = -1
                     End If
                 Else
                     Mennai = -1
                 End If
              End If
                  
              If G.cCount = 5 Then
                 side = R(3)
                 If (side <> -1) And (center <> -1) Then
                     If (center <> 0) Then
                         w_Max1 = Abs(center - side) * 100 / center
                     Else
                         w_Max1 = -1
                     End If
                 Else
                     w_Max1 = -1
                 End If
              
                 side = R(4)
                 If (side <> -1) And (center <> -1) Then
                     If (center <> 0) Then
                         w_Max2 = Abs(center - side) * 100 / center
                     Else
                         w_Max2 = -1
                     End If
                 Else
                     w_Max2 = -1
                 End If
              
                 Mennai = IIf(w_Max1 >= w_Max2, w_Max1, w_Max2)
                
              End If
                  
           Case Else
              
              w_Max1 = SideMin_Cal(RES_JUDG, R(), G)
              w_Max1 = Abs(center - w_Max1)
           
              w_Max2 = SideMax_Cal(RES_JUDG, R(), G)
              w_Max2 = Abs(center - w_Max2)
           
              side = IIf(w_Max1 >= w_Max2, w_Max1, w_Max2)
        
              If (side <> -9999) And (center <> -9999) Then
                  If (center <> 0) Then
                      Mennai = side * 100 / center
                  Else
                      Mennai = -1
                  End If
              Else
                  Mennai = -1
              End If
           
           End Select
           
        Case OI_JUDG ''判定識別フラグ(Oi)
           
           center = R(0)
           Select Case G.cPos
           Case "1", "G", "H", "J", "M", "N", "Q"
           
'特殊対応　2003.08.21 yakimura start
              
              If G.cCount = 2 Then
                 side = R(1)
                 If (side <> -1) And (center <> -1) Then
                     If (center <> 0) Then
                         Mennai = Abs(center - side) * 100 / center
                     Else
                         Mennai = -1
                     End If
                 Else
                     Mennai = -1
                 End If
              End If

'特殊対応　2003.08.21 yakimura start
              
              If G.cCount = 3 Then
                 side = R(2)
                 If (side <> -1) And (center <> -1) Then
                     If (center <> 0) Then
                         Mennai = Abs(center - side) * 100 / center
                     Else
                         Mennai = -1
                     End If
                 Else
                     Mennai = -1
                 End If
              End If
                  
              If G.cCount = 5 Then
                 side = R(3)
                 If (side <> -1) And (center <> -1) Then
                     If (center <> 0) Then
                         w_Max1 = Abs(center - side) * 100 / center
                     Else
                         w_Max1 = -1
                     End If
                 Else
                     w_Max1 = -1
                 End If
              
                 side = R(4)
                 If (side <> -1) And (center <> -1) Then
                     If (center <> 0) Then
                         w_Max2 = Abs(center - side) * 100 / center
                     Else
                         w_Max2 = -1
                     End If
                 Else
                     w_Max2 = -1
                 End If
              
                 Mennai = IIf(w_Max1 >= w_Max2, w_Max1, w_Max2)
                
              End If
           
           Case Else
              
              w_Max1 = SideMin_Cal(OI_JUDG, R(), G)
              w_Max1 = Abs(center - w_Max1)
           
              w_Max2 = SideMax_Cal(OI_JUDG, R(), G)
              w_Max2 = Abs(center - w_Max2)
        
              side = IIf(w_Max1 >= w_Max2, w_Max1, w_Max2)
        
              If (side <> -9999) And (center <> -9999) Then
                  If (center <> 0) Then
                      Mennai = side * 100 / center
                  Else
                      Mennai = -1
                  End If
              Else
                  Mennai = -1
              End If
           
           End Select
        
        End Select
    
    Case "E"                '---> |(centerave-sideave)|/centerave×100
        
        Select Case G.cPos
           Case "1", "G", "H", "J", "M"

'特殊対応　2003.08.21 yakimura start
              
              If G.cCount = 2 Then
                 If R(0) <> 0 Then
                    Mennai = Abs(R(0) - R(1)) * 100 / R(0)
                 Else
                    Mennai = -1
                 End If
              End If

'特殊対応　2003.08.21 yakimura end
              
              If G.cCount = 3 Then
                 If R(0) <> 0 Then
                    Mennai = Abs(R(0) - R(2)) * 100 / R(0)
                 Else
                    Mennai = -1
                 End If
              End If

' 2003.07.30 米沢事業所 工藤氏の確認により、測定位置値の平均を求めてから面内値を算出する  yakimura
              If G.cCount = 5 Then
'                 If R(0) <> 0 Then
'                    w_N3 = Abs(R(0) - R(3)) * 100 / R(0)
'                 Else
'                    w_N3 = -1
'                 End If
'                 If R(0) <> 0 Then
'                    w_N4 = Abs(R(0) - R(4)) * 100 / R(0)
'                 Else
'                    w_N4 = -1
'                 End If

'                 If w_N3 <> -1 and w_N4 <> -1 Then
'                    Mennai = (w_N3 + w_N4) / 2
'                 ElseIf w_N3 = -1 Then
'                    Mennai = w_N4
'                 ElseIf w_N4 = -1 Then
'                    Mennai = w_N3
'                 Else
'                    Mennai = -1
'                 End If

                 If R(3) <> -1 And R(4) <> -1 Then
                    w_Nx = (R(3) + R(4)) / 2
                 ElseIf R(3) = -1 Then
                    w_Nx = R(4)
                 ElseIf R(4) = -1 Then
                    w_Nx = R(3)
                 End If

                 If R(0) <> 0 Then
                    Mennai = Abs(R(0) - w_Nx) * 100 / R(0)
                 Else
                    Mennai = -1
                 End If

              End If

           Case Else

              center = R(0)
              Side_Ave = SideAve_Cal(RES_JUDG, R(), G)

              If (Side_Ave <> -9999) And (center <> -9999) Then
                  If (center <> 0) Then
                      Mennai = Abs(center - Side_Ave) * 100 / center
                  Else
                      Mennai = -1
                  End If
              Else
                  Mennai = -1
              End If

        End Select

    Case "M"                '---> (max-min)/ave×100
        Select Case JudgFlag
           Case RES_JUDG ''判定識別フラグ(RES)
              
              AVE = Ave_Cal(RES_JUDG, R(), G)
              Min = Min_Cal(RES_JUDG, R(), G)
              max = Max_Cal(RES_JUDG, R(), G)
        
              If (Min <> -9999) And (max <> -9999) And (AVE <> -9999) Then
                 If (AVE <> 0) Then
                    Mennai = (max - Min) * 100 / AVE
                 Else
                    Mennai = -1
                 End If
              Else
                 Mennai = -1
              End If
           
           Case OI_JUDG ''判定識別フラグ(Oi)　　　　　　特殊処理　「Oi」 は、"A" で計算

              Min = Min_Cal(OI_JUDG, R(), G)
              max = Max_Cal(OI_JUDG, R(), G)
        
              If (Min <> -9999) And (max <> -9999) Then
                  If (Min <> 0) Then
                      Mennai = (max - Min) * 100 / Min
                  Else
                      Mennai = -1
                  End If
              Else
                  Mennai = -1
              End If
        
        End Select

    Case "N"                '---> |(center-side)/(center+side)|×200
        
        Select Case JudgFlag
        Case RES_JUDG ''判定識別フラグ(RES)
           
           Select Case G.cPos
           Case "1", "G", "H", "J", "M"
                  
'特殊対応　2003.08.21 yakimura start
              
              If G.cCount = 2 Then
                 If R(0) <> 0 And R(1) <> 0 Then
                    Mennai = Abs(R(0) - R(1)) / (R(0) + R(1)) * 200
                 Else
                    Mennai = -1
                 End If
              End If

'特殊対応　2003.08.21 yakimura end
              
              If G.cCount = 3 Then
                 If R(0) <> 0 And R(2) <> 0 Then
                    Mennai = Abs(R(0) - R(2)) / (R(0) + R(2)) * 200
                 Else
                    Mennai = -1
                 End If
              End If
              
              If G.cCount = 5 Then
                 If R(0) <> 0 And R(3) <> 0 Then
                    w_N3 = Abs(R(0) - R(3)) / (R(0) + R(3)) * 200
                 Else
                    w_N3 = -1
                 End If
                 If R(0) <> 0 And R(4) <> 0 Then
                    w_N4 = Abs(R(0) - R(4)) / (R(0) + R(4)) * 200
                 Else
                    w_N4 = -1
                 End If
              
                 Mennai = IIf(w_N3 >= w_N4, w_N3, w_N4)
              End If
              
           Case Else
              
'特殊対応　2003.08.21 yakimura start
              
              If G.cCount = 2 Then
                 If R(0) <> 0 And R(1) <> 0 Then
                    Mennai = Abs(R(0) - R(1)) / (R(0) + R(1)) * 200
                 Else
                    Mennai = -1
                 End If
              
              ElseIf G.cCount = 3 Then
'特殊対応　2003.08.21 yakimura end
                 If R(0) <> 0 And R(1) <> 0 Then
                    w_N1 = Abs(R(0) - R(1)) / (R(0) + R(1)) * 200
                 Else
                    w_N1 = -1
                 End If
                 If R(0) <> 0 And R(2) <> 0 Then
                    w_N2 = Abs(R(0) - R(2)) / (R(0) + R(2)) * 200
                 Else
                    w_N2 = -1
                 End If
              
                 Mennai = IIf(w_N1 >= w_N2, w_N1, w_N2)
              
              ElseIf G.cCount = 5 Then
              
                 If R(0) <> 0 And R(1) <> 0 Then
                    w_N1 = Abs(R(0) - R(1)) / (R(0) + R(1)) * 200
                 Else
                    w_N1 = -1
                 End If
                 If R(0) <> 0 And R(2) <> 0 Then
                    w_N2 = Abs(R(0) - R(2)) / (R(0) + R(2)) * 200
                 Else
                    w_N2 = -1
                 End If
                 If R(0) <> 0 And R(3) <> 0 Then
                    w_N3 = Abs(R(0) - R(3)) / (R(0) + R(3)) * 200
                 Else
                    w_N3 = -1
                 End If
                 If R(0) <> 0 And R(4) <> 0 Then
                    w_N4 = Abs(R(0) - R(4)) / (R(0) + R(4)) * 200
                 Else
                    w_N4 = -1
                 End If
           
                 Mennai = IIf(w_N1 >= w_N2, w_N1, w_N2)
                 Mennai = IIf(Mennai >= w_N3, Mennai, w_N3)
                 Mennai = IIf(Mennai >= w_N4, Mennai, w_N4)
              
              End If
           
           End Select
           
        Case OI_JUDG ''判定識別フラグ(Oi)
           
           Select Case G.cPos
           Case "1", "G", "H", "J", "M", "N", "Q"
                  
'特殊対応　2003.08.21 yakimura start
              
              If G.cCount = 2 Then
                 If R(0) <> 0 And R(1) <> 0 Then
                    Mennai = Abs(R(0) - R(1)) / (R(0) + R(1)) * 200
                 Else
                    Mennai = -1
                 End If
              End If
              
'特殊対応　2003.08.21 yakimura end
              
              If G.cCount = 3 Then
                 If R(0) <> 0 And R(2) <> 0 Then
                    Mennai = Abs(R(0) - R(2)) / (R(0) + R(2)) * 200
                 Else
                    Mennai = -1
                 End If
              End If
              
              If G.cCount = 5 Then
                 If R(0) <> 0 And R(3) <> 0 Then
                    w_N3 = Abs(R(0) - R(3)) / (R(0) + R(3)) * 200
                 Else
                    w_N3 = -1
                 End If
                 If R(0) <> 0 And R(4) <> 0 Then
                    w_N4 = Abs(R(0) - R(4)) / (R(0) + R(4)) * 200
                 Else
                    w_N4 = -1
                 End If
              
                 Mennai = IIf(w_N3 >= w_N4, w_N3, w_N4)
              End If
              
           Case Else
              
'特殊対応　2003.08.21 yakimura start
              
              If G.cCount = 2 Then
                 If R(0) <> 0 And R(1) <> 0 Then
                    Mennai = Abs(R(0) - R(1)) / (R(0) + R(1)) * 200
                 Else
                    Mennai = -1
                 End If

              ElseIf G.cCount = 3 Then
'特殊対応　2003.08.21 yakimura end
                 If R(0) <> 0 And R(1) <> 0 Then
                    w_N1 = Abs(R(0) - R(1)) / (R(0) + R(1)) * 200
                 Else
                    w_N1 = -1
                 End If
                 If R(0) <> 0 And R(2) <> 0 Then
                    w_N2 = Abs(R(0) - R(2)) / (R(0) + R(2)) * 200
                 Else
                    w_N2 = -1
                 End If
              
                 Mennai = IIf(w_N1 >= w_N2, w_N1, w_N2)
              
              ElseIf G.cCount = 5 Then
              
                 If R(0) <> 0 And R(1) <> 0 Then
                    w_N1 = Abs(R(0) - R(1)) / (R(0) + R(1)) * 200
                 Else
                    w_N1 = -1
                 End If
                 If R(0) <> 0 And R(2) <> 0 Then
                    w_N2 = Abs(R(0) - R(2)) / (R(0) + R(2)) * 200
                 Else
                    w_N2 = -1
                 End If
                 If R(0) <> 0 And R(3) <> 0 Then
                    w_N3 = Abs(R(0) - R(3)) / (R(0) + R(3)) * 200
                 Else
                    w_N3 = -1
                 End If
                 If R(0) <> 0 And R(4) <> 0 Then
                    w_N4 = Abs(R(0) - R(4)) / (R(0) + R(4)) * 200
                 Else
                    w_N4 = -1
                 End If
           
                 Mennai = IIf(w_N1 >= w_N2, w_N1, w_N2)
                 Mennai = IIf(Mennai >= w_N3, Mennai, w_N3)
                 Mennai = IIf(Mennai >= w_N4, Mennai, w_N4)
              
              End If
           
           End Select
           
        End Select
           

    Case " ", ""
        
        ' 計算も判定も行わない
        
    Case Else               '---> (max-min)/min×100     特殊ケース　"A" として計算する
        Select Case JudgFlag
           Case RES_JUDG ''判定識別フラグ(RES)
              
              Min = Min_Cal(RES_JUDG, R(), G)
              max = Max_Cal(RES_JUDG, R(), G)
           
           Case OI_JUDG ''判定識別フラグ(Oi)
              
              Min = Min_Cal(OI_JUDG, R(), G)
              max = Max_Cal(OI_JUDG, R(), G)
        
        End Select
        
        If (Min <> -9999) And (max <> -9999) Then
            If (Min <> 0) Then
                Mennai = (max - Min) * 100 / Min
            Else
                Mennai = -1
            End If
        End If
    
    End Select

    'MENNAI_Cal = RoundUp(Mennai, 4)              '''既存の処理ではこうなるが…
    MENNAI_Cal_NotRound = Mennai        '丸めない値を返すように変更 2011/11/25 SETsw kubota

End Function
'端数丸めを行う関数と行わない関数の二つに分離 2011/11/25 SETsw kubota
Public Function MENNAI_Cal(JudgFlag As String, R() As Double, G As Guarantee, calcode As String) As Double
    '端数丸めを行わない関数を呼び出し、小数4桁(5桁目切り上げ)にして返す
    MENNAI_Cal = RoundUp(MENNAI_Cal_NotRound(JudgFlag, R(), G, calcode), 4)
End Function



'概要      :判定対象MINデータを求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :戻り値        ,O  ,double    ,RRG
'説明      :
'履歴      :2003/06/06  yakimura  作成
Public Function Min_Cal(JudgFlag As String, d() As Double, G As Guarantee) As Double
Dim Min    As Double
    
    Min = -9999
    
    Select Case G.cCount
    Case "1"
            Min = d(0)
    Case "2"
            Min = IIf(d(0) <= d(1), d(0), d(1))
    Case "3"
            Min = JudgMin2(d())
    Case "5"
            Min = JudgMin(d())
    End Select
    
    Min_Cal = Min

End Function

'概要      :判定対象MAXデータを求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :戻り値        ,O  ,double    ,RRG
'説明      :
'履歴      :2003/06/06  yakimura  作成
Public Function Max_Cal(JudgFlag As String, d() As Double, G As Guarantee) As Double
Dim max    As Double
    
    max = -9999
    
    Select Case G.cCount
    Case "1"
            max = d(0)
    Case "2"
            max = IIf(d(0) <= d(1), d(1), d(0))
    Case "3"
            max = JudgMax2(d())
    Case "5"
            max = JudgMax(d())
    End Select

    Max_Cal = max

End Function

'概要      :判定対象aveデータを求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :戻り値        ,O  ,double    ,RRG
'説明      :
'履歴      :2003/06/06  yakimura  作成
Public Function Ave_Cal(JudgFlag As String, d() As Double, G As Guarantee) As Double
Dim AVE As Double
    
    AVE = -9999
    
    Select Case G.cCount
    Case "1"
            AVE = d(0)
    Case "2"
            If d(0) <> -1 And d(1) <> -1 Then
               AVE = (d(0) + d(1)) / 2
            ElseIf d(0) = -1 Then
               AVE = d(1)
            ElseIf d(1) = -1 Then
               AVE = d(0)
            End If
    Case "3"
            AVE = JudgAve2(d())
    Case "5"
            AVE = JudgAve(d())
    End Select

    Ave_Cal = AVE

End Function

'概要      :判定対象side_aveデータを求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :戻り値        ,O  ,double    ,RRG
'説明      :
'履歴      :2003/06/06  yakimura  作成
Public Function SideAve_Cal(JudgFlag As String, d() As Double, G As Guarantee) As Double
Dim sideave As Double
    
    sideave = -9999
    
    Select Case G.cCount
    Case "2"
            sideave = d(1)
    Case "3"
            sideave = JudgSideAve2(d())
    Case "5"
            sideave = JudgSideAve(d())
    End Select

    SideAve_Cal = sideave

End Function


'概要      :判定対象side_minデータを求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :戻り値        ,O  ,double    ,RRG
'説明      :
'履歴      :2003/06/06  yakimura  作成
Public Function SideMin_Cal(JudgFlag As String, d() As Double, G As Guarantee) As Double
Dim SideMin As Double
    
    SideMin = -9999
    
    Select Case G.cCount
    Case "2"
            SideMin = d(1)
    Case "3"
            If d(1) <> -1 And d(2) <> -1 Then
               SideMin = IIf(d(1) <= d(2), d(1), d(2))
            ElseIf d(1) = -1 Then
               SideMin = d(2)
            ElseIf d(2) = -1 Then
               SideMin = d(1)
            End If
    Case "5"
            SideMin = JudgSideMin(d())
    End Select

    SideMin_Cal = SideMin

End Function


'概要      :判定対象side_maxデータを求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :戻り値        ,O  ,double    ,RRG
'説明      :
'履歴      :2003/06/06  yakimura  作成
Public Function SideMax_Cal(JudgFlag As String, d() As Double, G As Guarantee) As Double
Dim SideMax As Double
    
    SideMax = -9999
    
    Select Case G.cCount
    Case "2"
            SideMax = d(1)
    Case "3"
            If d(1) <> -1 And d(2) <> -1 Then
               SideMax = IIf(d(1) <= d(2), d(2), d(1))
            ElseIf d(1) = -1 Then
               SideMax = d(2)
            ElseIf d(2) = -1 Then
               SideMax = d(1)
            End If
    Case "5"
            SideMax = JudgSideMax(d())
    End Select

    SideMax_Cal = SideMax

End Function

'概要      :測定値がNULLだった場合の範囲判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :JudgData      ,I  ,double    ,測定値
'          :SpecMin       ,I  ,double    ,下限値
'          :SpecMax       ,I  ,double    ,上限値
'          :戻り値        ,O  ,Boolean   ,JUDG_OK or JUDG_NG
'説明      :
'履歴      :2003/12/11 新規作成 システムブレイン
Public Function RangeDecision_nl(JudgData As Double, SpecMin As Double, SpecMax As Double) As Boolean
    RangeDecision_nl = False
    If (JudgData >= SpecMin) Or (SpecMin = -1) Then
        If (JudgData <= SpecMax) Or (SpecMax = -1) Then
            RangeDecision_nl = True
        End If
    End If
'    RangeDecision = ((JudgData >= SpecMin) And (JudgData <= SpecMax))
End Function

' 指定された実測抵抗データに仕様測定位置より抽出した抵抗データをセットしなおす
'sokuteiTensu  = 測定点数(HSXRSPOT)
'sokuteiIchi   = 測定位置(HSXRSPOI)
'typ_RS        = 実測データ(TBCMJ002)
'TEST2004/10
Public Function Set_Rs_Ichi(sokuteiTensu As String, sokuteiIchi As String, MEAS1 As Double, _
                            MEAS2 As Double, MEAS3 As Double, MEAS4 As Double, MEAS5 As Double) As FUNCTION_RETURN
Dim sTensu As String
Dim sName As String
Dim sMeas(1 To 5) As Double
Dim sMeas1(1 To 5) As Double
Dim i As Integer
Set_Rs_Ichi = FUNCTION_RETURN_FAILURE

''現在1,2,3,5点のみ対応
If InStr("1235", sokuteiTensu) = 0 Then
    Exit Function
End If

''TBCMB005(1点=info1,2点=info2,3点=info3,5点=info5）
sName = "info" & sokuteiTensu

sTensu = GetCodeField("SC", "30", sokuteiIchi, sName)
If sTensu = "" Then
    Exit Function
End If
''

sMeas(1) = MEAS1
sMeas(2) = MEAS2
sMeas(3) = MEAS3
sMeas(4) = MEAS4
sMeas(5) = MEAS5

For i = 1 To 5
    sMeas1(i) = -1
Next
For i = 1 To sokuteiTensu
    'コードＤＢ = ex...1点=1,3点=133,5点=13333
    sMeas1(i) = sMeas(Mid(sTensu, i, 1))
Next

MEAS1 = sMeas1(1)
MEAS2 = sMeas1(2)
MEAS3 = sMeas1(3)
MEAS4 = sMeas1(4)
MEAS5 = sMeas1(5)

Set_Rs_Ichi = FUNCTION_RETURN_SUCCESS

End Function

'Add Start 2011/01/26 SMPK Miyata
'概要      :Cudecoのパターンの判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :SpcKbn        ,I  ,string    ,パターン区分
'          :JskKbn        ,I  ,string    ,パターン実績
'          :HanteiKbn     ,I  ,string    ,OKパターン区分(１文字目はバターン区分、２文字以降がOKバターン区分)
'          :戻り値        ,O  ,Boolean   ,JUDG_OK or JUDG_NG
'説明      :
'履歴      :
Public Function CudecoJudgPattern(SpcKbn As String, JskKbn As String, HanteiKbn() As String) As Boolean
    Dim ii As Integer
    
    CudecoJudgPattern = JUDG_NG
    For ii = 0 To UBound(HanteiKbn)
        If Mid(HanteiKbn(ii), 1, 1) = SpcKbn Then
            If InStr(2, HanteiKbn(ii), JskKbn) > 0 Then
                CudecoJudgPattern = JUDG_OK
                Exit For
            End If
        End If
    Next ii
    
End Function
'Add End   2011/01/26 SMPK Miyata

