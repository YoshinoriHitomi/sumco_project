Attribute VB_Name = "SB_WfJudg_SQL"
Option Explicit

'' WFセンター総合判定待ち一覧

' SXL管理
Public Type DBDRV_scmzc_fcmlc001b_SXL
    CRYNUM      As String * 12      ' 結晶番号
    INGOTPOS    As Integer          ' 結晶内開始位置
    Length      As Integer          ' 長さ
    SXLID       As String * 13      ' SXLID
    KRPROCCD    As String * 5       ' 管理工程
    NOWPROC     As String * 5       ' 現在工程
    LPKRPROCCD  As String * 5       ' 最終通過管理工程
    LASTPASS    As String * 5       ' 最終通過工程
    DELCLS      As String * 1       ' 削除区分
    LSTATCLS    As String * 1       ' 最終状態区分
    HOLDCLS     As String * 1       ' ホールド区分
    hinban      As String * 8       ' 品番
    REVNUM      As Integer          ' 製品番号改訂番号
    factory     As String * 1       ' 工場
    opecond     As String * 1       ' 操業条件
    COUNT       As Integer          ' 枚数
    REGDATE     As Date             ' 登録日付
    UPDDATE     As Date             ' 更新日付
    KETURAKU    As Boolean          ' 欠落情報有無フラグ
    WFSMP()     As typ_XSDCW        ' サンプル管理（TOP、TAIL順 ２レコード）
End Type


' WFセンター総合判定
' 入力用
Public Type type_DBDRV_scmzc_fcmlc001c_In
    HIN         As tFullHinban      ' 品番(full)
    SAMPLEID    As String * 16      ' サンプルID
    SXLID       As String * 13      ' SXLID
    WFSMP       As typ_XSDCW        ' ｻﾝﾌﾟﾙ管理
End Type

' WF製品仕様取得用
Public Type type_DBDRV_scmzc_fcmlc001c_Siyou
    HWFTYPE As String * 1           ' 品ＷＦタイプ
    HWFCDIR As String * 1           ' 品ＷＦ結晶面方
    HWFCDOP As String * 1           ' 品ＷＦ結晶ドープ
    HWFRMIN As Double               ' 品ＷＦ比抵抗下限
    HWFRMAX As Double               ' 品ＷＦ比抵抗上限
    HWFRSPOH As String * 1          ' 品ＷＦ比抵抗測定位置＿方
    HWFRSPOT As String * 1          ' 品ＷＦ比抵抗測定位置＿点
    HWFRSPOI As String * 1          ' 品ＷＦ比抵抗測定位置＿位
    HWFRHWYT As String * 1          ' 品ＷＦ比抵抗保証方法＿対
    HWFRHWYS As String * 1          ' 品ＷＦ比抵抗保証方法＿処
    HWFRMCAL As String * 1          ' 品ＷＦ比抵抗面内計算
    HWFRAMIN As Double              ' 品ＷＦ比抵抗平均下限
    HWFRAMAX As Double              ' 品ＷＦ比抵抗平均上限
    HWFRMBNP As Double              ' 品ＷＦ比抵抗面内分布
    
    HWFMKMIN As Double              ' 品ＷＦ無欠陥層下限
    HWFMKMAX As Double              ' 品ＷＦ無欠陥層上限
    HWFMKSPH As String * 1          ' 品ＷＦ無欠陥層測定位置＿方
    HWFMKSPT As String * 1          ' 品ＷＦ無欠陥層測定位置＿点
    HWFMKSPR As String * 1          ' 品ＷＦ無欠陥層測定位置＿領
    HWFMKHWT As String * 1          ' 品ＷＦ無欠陥層保証方法＿対
    HWFMKHWS As String * 1          ' 品ＷＦ無欠陥層保証方法＿処

    HWFONMIN As Double              ' 品ＷＦ酸素濃度下限
    HWFONMAX As Double              ' 品ＷＦ酸素濃度上限
    HWFONSPH As String * 1          ' 品ＷＦ酸素濃度測定位置＿方
    HWFONSPT As String * 1          ' 品ＷＦ酸素濃度測定位置＿点
    HWFONSPI As String * 1          ' 品ＷＦ酸素濃度測定位置＿位
    HWFONHWT As String * 1          ' 品ＷＦ酸素濃度保証方法＿対
    HWFONHWS As String * 1          ' 品ＷＦ酸素濃度保証方法＿処
    HWFONMCL As String * 1          ' 品ＷＦ酸素濃度面内計算
    HWFONMBP As Double              ' 品ＷＦ酸素濃度面内分布
    HWFONAMN As Double              ' 品ＷＦ酸素濃度平均下限
    HWFONAMX As Double              ' 品ＷＦ酸素濃度平均上限

    HWFOS1MN As Double              ' 品ＷＦ酸素析出１下限
    HWFOS1MX As Double              ' 品ＷＦ酸素析出１上限
    HWFOS1SH As String * 1          ' 品ＷＦ酸素析出１測定位置＿方
    HWFOS1ST As String * 1          ' 品ＷＦ酸素析出１測定位置＿点
    HWFOS1SI As String * 1          ' 品ＷＦ酸素析出１測定位置＿位
    HWFOS1HT As String * 1          ' 品ＷＦ酸素析出１保証方法＿対
    HWFOS1HS As String * 1          ' 品ＷＦ酸素析出１保証方法＿処
    HWFOS2SH As String * 1          ' 品ＷＦ酸素析出２測定位置＿方
    HWFOS2ST As String * 1          ' 品ＷＦ酸素析出２測定位置＿点
    HWFOS2SI As String * 1          ' 品ＷＦ酸素析出２測定位置＿位
    HWFOS2MN As Double              ' 品ＷＦ酸素析出２下限
    HWFOS2MX As Double              ' 品ＷＦ酸素析出２上限
    HWFOS2HT As String * 1          ' 品ＷＦ酸素析出２保証方法＿対
    HWFOS2HS As String * 1          ' 品ＷＦ酸素析出２保証方法＿処
    HWFOS3MN As Double              ' 品ＷＦ酸素析出３下限
    HWFOS3MX As Double              ' 品ＷＦ酸素析出３上限
    HWFOS3SH As String * 1          ' 品ＷＦ酸素析出３測定位置＿方
    HWFOS3ST As String * 1          ' 品ＷＦ酸素析出３測定位置＿点
    HWFOS3SI As String * 1          ' 品ＷＦ酸素析出３測定位置＿位
    HWFOS3HT As String * 1          ' 品ＷＦ酸素析出３保証方法＿対
    HWFOS3HS As String * 1          ' 品ＷＦ酸素析出３保証方法＿処

    HWFZOMIN As Double              ' 品ＷＦ残存酸素下限
    HWFZOMAX As Double              ' 品ＷＦ残存酸素上限
    HWFZOSPH As String * 1          ' 品ＷＦ残存酸素測定位置＿方
    HWFZOSPT As String * 1          ' 品ＷＦ残存酸素測定位置＿点
    HWFZOSPI As String * 1          ' 品ＷＦ残存酸素測定位置＿位
    HWFZOHWT As String * 1          ' 品ＷＦ残存酸素保証方法＿対
    HWFZOHWS As String * 1          ' 品ＷＦ残存酸素保証方法＿処
    
    HWFDSOMX As Double              ' 品ＷＦＤＳＯＤ上限
    HWFDSOMN As Double              ' 品ＷＦＤＳＯＤ下限
    HWFDSOAX As Integer             ' 品ＷＦＤＳＯＤ領域上限
    HWFDSOAN As Integer             ' 品ＷＦＤＳＯＤ領域下限
    HWFDSOHT As String * 1          ' 品ＷＦＤＳＯＤ保証方法＿対
    HWFDSOHS As String * 1          ' 品ＷＦＤＳＯＤ保証方法＿処
    HWFDSOPTK As String * 1         ' 品ＷＦＤＳＯＤパタン区分
    
    HWFSPVMX As Double              ' 品ＷＦＳＰＶＦＥ上限
    HWFSPVAM As Double              ' 品ＷＦＳＰＶＦＥ平均上限
    HWFSPVSH As String * 1          ' 品ＷＦＳＰＶＦＥ測定位置＿方
    HWFSPVST As String * 1          ' 品ＷＦＳＰＶＦＥ測定位置＿点
    HWFSPVSI As String * 1          ' 品ＷＦＳＰＶＦＥ測定位置＿位
    HWFSPVHT As String * 1          ' 品ＷＦＳＰＶＦＥ保証方法＿対
    HWFSPVHS As String * 1          ' 品ＷＦＳＰＶＦＥ保証方法＿処
    HWFDLSPH As String * 1          ' 品ＷＦ拡散長測定位置＿方
    HWFDLSPT As String * 1          ' 品ＷＦ拡散長測定位置＿点
    HWFDLSPI As String * 1          ' 品ＷＦ拡散長測定位置＿位
    HWFDLHWT As String * 1          ' 品ＷＦ拡散長保証方法＿対
    HWFDLHWS As String * 1          ' 品ＷＦ拡散長保証方法＿処
    HWFDLMIN As Integer             ' 品ＷＦ拡散長下限
    HWFDLMAX As Integer             ' 品ＷＦ拡散長上限
    HWFNRHS As String * 1           ' 品ＷＦＳＰＶＮＲ保証方法＿処
    HWFNRKN As String * 1           ' 品ＷＦＳＰＶＮＲ保証方法＿抜
    
    HWFOF1AX As Double              ' 品ＷＦＯＳＦ１平均上限
    HWFOF1MX As Double              ' 品ＷＦＯＳＦ１上限
    HWFOF1SH As String * 1          ' 品ＷＦＯＳＦ１測定位置＿方
    HWFOF1ST As String * 1          ' 品ＷＦＯＳＦ１測定位置＿点
    HWFOF1SR As String * 1          ' 品ＷＦＯＳＦ１測定位置＿領
    HWFOF1HT As String * 1          ' 品ＷＦＯＳＦ１保証方法＿対
    HWFOF1HS As String * 1          ' 品ＷＦＯＳＦ１保証方法＿処
    HWFOF2AX As Double              ' 品ＷＦＯＳＦ２平均上限
    HWFOF2MX As Double              ' 品ＷＦＯＳＦ２上限
    HWFOF2SH As String * 1          ' 品ＷＦＯＳＦ２測定位置＿方
    HWFOF2ST As String * 1          ' 品ＷＦＯＳＦ２測定位置＿点
    HWFOF2SR As String * 1          ' 品ＷＦＯＳＦ２測定位置＿領
    HWFOF2HT As String * 1          ' 品ＷＦＯＳＦ２保証方法＿対
    HWFOF2HS As String * 1          ' 品ＷＦＯＳＦ２保証方法＿処
    HWFOF3AX As Double              ' 品ＷＦＯＳＦ３平均上限
    HWFOF3MX As Double              ' 品ＷＦＯＳＦ３上限
    HWFOF3SH As String * 1          ' 品ＷＦＯＳＦ３測定位置＿方
    HWFOF3ST As String * 1          ' 品ＷＦＯＳＦ３測定位置＿点
    HWFOF3SR As String * 1          ' 品ＷＦＯＳＦ３測定位置＿領
    HWFOF3HT As String * 1          ' 品ＷＦＯＳＦ３保証方法＿対
    HWFOF3HS As String * 1          ' 品ＷＦＯＳＦ３保証方法＿処
    HWFOF4AX As Double              ' 品ＷＦＯＳＦ４平均上限
    HWFOF4MX As Double              ' 品ＷＦＯＳＦ４上限
    HWFOF4SH As String * 1          ' 品ＷＦＯＳＦ４測定位置＿方
    HWFOF4ST As String * 1          ' 品ＷＦＯＳＦ４測定位置＿点
    HWFOF4SR As String * 1          ' 品ＷＦＯＳＦ４測定位置＿領
    HWFOF4HT As String * 1          ' 品ＷＦＯＳＦ４保証方法＿対
    HWFOF4HS As String * 1          ' 品ＷＦＯＳＦ４保証方法＿処
    HWFOSF1PTK As String * 1        ' 品ＷＦＯＳＦ１パタン区分
    HWFOSF2PTK As String * 1        ' 品ＷＦＯＳＦ２パタン区分
    HWFOSF3PTK As String * 1        ' 品ＷＦＯＳＦ３パタン区分
    HWFOSF4PTK As String * 1        ' 品ＷＦＯＳＦ４パタン区分
    
    HWFBM1AN As Double              ' 品ＷＦＢＭＤ１平均下限
    HWFBM1AX As Double              ' 品ＷＦＢＭＤ１平均上限
    HWFBM1SH As String * 1          ' 品ＷＦＢＭＤ１測定位置＿方
    HWFBM1ST As String * 1          ' 品ＷＦＢＭＤ１測定位置＿点
    HWFBM1SR As String * 1          ' 品ＷＦＢＭＤ１測定位置＿領
    HWFBM1HT As String * 1          ' 品ＷＦＢＭＤ１保証方法＿対
    HWFBM1HS As String * 1          ' 品ＷＦＢＭＤ１保証方法＿処
    HWFBM2AN As Double              ' 品ＷＦＢＭＤ２平均下限
    HWFBM2AX As Double              ' 品ＷＦＢＭＤ２平均上限
    HWFBM2SH As String * 1          ' 品ＷＦＢＭＤ２測定位置＿方
    HWFBM2ST As String * 1          ' 品ＷＦＢＭＤ２測定位置＿点
    HWFBM2SR As String * 1          ' 品ＷＦＢＭＤ２測定位置＿領
    HWFBM2HT As String * 1          ' 品ＷＦＢＭＤ２保証方法＿対
    HWFBM2HS As String * 1          ' 品ＷＦＢＭＤ２保証方法＿処
    HWFBM3AN As Double              ' 品ＷＦＢＭＤ３平均下限
    HWFBM3AX As Double              ' 品ＷＦＢＭＤ３平均上限
    HWFBM3SH As String * 1          ' 品ＷＦＢＭＤ３測定位置＿方
    HWFBM3ST As String * 1          ' 品ＷＦＢＭＤ３測定位置＿点
    HWFBM3SR As String * 1          ' 品ＷＦＢＭＤ３測定位置＿領
    HWFBM3HT As String * 1          ' 品ＷＦＢＭＤ３保証方法＿対
    HWFBM3HS As String * 1          ' 品ＷＦＢＭＤ３保証方法＿処
    HWFBM1MBP As Single             ' 品ＷＦＢＭＤ１面内分布
    HWFBM2MBP As Single             ' 品ＷＦＢＭＤ２面内分布
    HWFBM3MBP As Single             ' 品ＷＦＢＭＤ３面内分布
    HWFBM1MCL As String * 2         ' 品ＷＦＢＭＤ１面内計算
    HWFBM2MCL As String * 2         ' 品ＷＦＢＭＤ２面内計算
    HWFBM3MCL As String * 2         ' 品ＷＦＢＭＤ３面内計算
    
    HWFOS1NS As String * 2          ' 品ＷＦ酸素析出１熱処理法
    HWFOS2NS As String * 2          ' 品ＷＦ酸素析出２熱処理法
    HWFOS3NS As String * 2          ' 品ＷＦ酸素析出３熱処理法
    HWFZONSW As String * 2          ' 品ＷＦ残存酸素熱処理法
    HWFOF1NS As String * 2          ' 品ＷＦＯＳＦ１熱処理法
    HWFOF2NS As String * 2          ' 品ＷＦＯＳＦ２熱処理法
    HWFOF3NS As String * 2          ' 品ＷＦＯＳＦ３熱処理法
    HWFOF4NS As String * 2          ' 品ＷＦＯＳＦ４熱処理法
    HWFBM1NS As String * 2          ' 品ＷＦＢＭＤ１熱処理法
    HWFBM2NS As String * 2          ' 品ＷＦＢＭＤ２熱処理法
    HWFBM3NS As String * 2          ' 品ＷＦＢＭＤ３熱処理法
    
    HWFANTIM As Integer             ' 品ＷＦＡＮ時間
    HWFANTNP As Integer             ' 品ＷＦＡＮ温度

    HWFOF1ET As Integer             ' 品ＷＦＯＳＦ１選択ＥＴ代
    HWFOF2ET As Integer             ' 品ＷＦＯＳＦ２選択ＥＴ代
    HWFOF3ET As Integer             ' 品ＷＦＯＳＦ３選択ＥＴ代
    HWFOF4ET As Integer             ' 品ＷＦＯＳＦ４選択ＥＴ代
    HWFBM1ET As Integer             ' 品ＷＦＢＭＤ１選択ＥＴ代
    HWFBM2ET As Integer             ' 品ＷＦＢＭＤ２選択ＥＴ代
    HWFBM3ET As Integer             ' 品ＷＦＢＭＤ３選択ＥＴ代

    HWFOF1SZ As String * 1          ' 品ＷＦＯＳＦ１測定条件
    HWFOF2SZ As String * 1          ' 品ＷＦＯＳＦ２測定条件
    HWFOF3SZ As String * 1          ' 品ＷＦＯＳＦ３測定条件
    HWFOF4SZ As String * 1          ' 品ＷＦＯＳＦ４測定条件
    HWFBM1SZ As String * 1          ' 品ＷＦＢＭＤ１測定条件
    HWFBM2SZ As String * 1          ' 品ＷＦＢＭＤ２測定条件
    HWFBM3SZ As String * 1          ' 品ＷＦＢＭＤ３測定条件
    
    HWFDENKU As String * 1          ' 品ＷＦＤｅｎ検査有無
    HWFDENMX As Integer             ' 品ＷＦＤｅｎ上限
    HWFDENMN As Integer             ' 品ＷＦＤｅｎ下限
    HWFDENHT As String * 1          ' 品ＷＦＤｅｎ保証方法＿対
    HWFDENHS As String * 1          ' 品ＷＦＤｅｎ保証方法＿処
    HWFDVDKU As String * 1          ' 品ＷＦＤＶＤ２検査有無
    HWFDVDMXN As Integer            ' 品ＷＦＤＶＤ２上限
    HWFDVDMNN As Integer            ' 品ＷＦＤＶＤ２下限
    HWFDVDHT As String * 1          ' 品ＷＦＤＶＤ２保証方法＿対
    HWFDVDHS As String * 1          ' 品ＷＦＤＶＤ２保証方法＿処
    HWFLDLKU As String * 1          ' 品ＷＦＬ／ＤＬ検査有無
    HWFLDLMX As Integer             ' 品ＷＦＬ／ＤＬ上限
    HWFLDLMN As Integer             ' 品ＷＦＬ／ＤＬ下限
    HWFLDLHT As String * 1          ' 品ＷＦＬ／ＤＬ保証方法＿対
    HWFLDLHS As String * 1          ' 品ＷＦＬ／ＤＬ保証方法＿処
    HWFGDSPH As String * 1          ' 品ＷＦＧＤ測定位置＿方
    HWFGDSPT As String * 1          ' 品ＷＦＧＤ測定位置＿点
    HWFGDSPR As String * 1          ' 品ＷＦＧＤ測定位置＿領
    HWFGDLINE As Single             '品ＷＦＧＤライン数
    HWFRKHNN As String * 1          ' 品ＷＦ比抵抗検査頻度＿抜
    HWFONKHN As String * 1          ' 品ＷＦ酸素濃度検査頻度＿抜
    HWFOF1KN As String * 1          ' 品ＷＦＯＳＦ１検査頻度＿抜
    HWFOF2KN As String * 1          ' 品ＷＦＯＳＦ２検査頻度＿抜
    HWFOF3KN As String * 1          ' 品ＷＦＯＳＦ３検査頻度＿抜
    HWFOF4KN As String * 1          ' 品ＷＦＯＳＦ４検査頻度＿抜
    HWFBM1KN As String * 1          ' 品ＷＦＢＭＤ１検査頻度＿抜
    HWFBM2KN As String * 1          ' 品ＷＦＢＭＤ２検査頻度＿抜
    HWFBM3KN As String * 1          ' 品ＷＦＢＭＤ３検査頻度＿抜
    HWFOS1KN As String * 1          ' 品ＷＦ酸素析出１検査頻度＿抜
    HWFOS2KN As String * 1          ' 品ＷＦ酸素析出２検査頻度＿抜
    HWFOS3KN As String * 1          ' 品ＷＦ酸素析出３検査頻度＿抜
    HWFDSOKN As String * 1          ' 品ＷＦＤＳＯＤ検査頻度＿抜
    HWFMKKHN As String * 1          ' 品ＷＦ無欠陥層検査頻度＿抜
    HWFSPVKN As String * 1          ' 品ＷＦＳＰＶＦＥ検査頻度＿抜
    HWFDLKHN As String * 1          ' 品ＷＦ拡散長検査頻度＿抜
    HWFZOKHN As String * 1          ' 品ＷＦ残存酸素検査頻度＿抜
    HWFGDKHN As String * 1          ' 品ＷＦＧＤ検査頻度＿抜
    BLOCKID() As String * 12        ' ブロックID

''SPV判定処理追加
''既存の構造体に項目追加するとVBの制限に引っかかるので、別で管理する。
''WF製品仕様取得用(CMBC039用)
    HWFSPVPUG As Double             ' 品ＷＦＳＰＶＦＥＰＵＡ限
    HWFSPVPUR As Double             ' 品ＷＦＳＰＶＦＥＰＵＡ率
    HWFSPVSTD As Double             ' 品ＷＦＳＰＶＦＥ標準偏差
    HWFNRMX   As Double             ' 品ＷＦＳＰＶＮＲ上限
    HWFNRPUG  As Double             ' 品ＷＦＳＰＶＮＲＰＵＡ限
    HWFNRPUR  As Double             ' 品ＷＦＳＰＶＮＲＰＵＡ率
    HWFNRSTD  As Double             ' 品ＷＦＳＰＶＮＲ標準偏差
    HWFDLPUG  As Double             ' 品ＷＦ拡散長ＰＵＡ限
    HWFDLPUR  As Double             ' 品ＷＦ拡散長ＰＵＡ率
    HWFNRAM   As Double             ' 品ＷＦＳＰＶＮＲ平均
    HWFNRSH   As String * 1         ' 品ＷＦＳＰＶＮＲ測定位置_方
    HWFNRST   As String * 1         ' 品ＷＦＳＰＶＮＲ測定位置_点
    HWFNRHT   As String * 1         ' 品ＷＦＳＰＶＮＲ保証方法_対
    HWFNRSI   As String * 1         ' 品ＷＦＳＰＶＮＲ測定位置_位

' エピ先行評価追加対応
    HEPHS      As Boolean           ' エピ仕様フラグ(有:1,無:0)
    HEPANTNP   As Integer           ' 品EPAN温度
    HEPOF1AX   As Double            ' 品EPOSF1平均上限
    HEPOF1MX   As Double            ' 品EPOSF1上限
    HEPOF1ET   As Double            ' 品EPOSF1選択ET代
    HEPOF1NS   As String * 2        ' 品EPOSF1熱処理法
    HEPOF1SZ   As String * 1        ' 品EPOSF1測定条件
    HEPOF1SH   As String * 1        ' 品EPOSF1測定位置_方
    HEPOF1ST   As String * 1        ' 品EPOSF1測定位置_点
    HEPOF1SR   As String * 1        ' 品EPOSF1測定位置_領
    HEPOF1HT   As String * 1        ' 品EPOSF1保証方法_対
    HEPOF1HS   As String * 1        ' 品EPOSF1保証方法_処
    HEPOF1KM   As String * 1        ' 品EPOSF1検査頻度_枚
    HEPOF1KN   As String * 1        ' 品EPOSF1検査頻度_抜
    HEPOF1KH   As String * 1        ' 品EPOSF1検査頻度_保
    HEPOF1KU   As String * 1        ' 品EPOSF1検査頻度_ｳ
    HEPOSF1PTK As String * 1        ' 品EPOSF1ﾊﾟﾀﾝ区分
    HEPOF2AX   As Double            ' 品EPOSF2平均上限
    HEPOF2MX   As Double            ' 品EPOSF2上限
    HEPOF2ET   As Double            ' 品EPOSF2選択ET代
    HEPOF2NS   As String * 2        ' 品EPOSF2熱処理法
    HEPOF2SZ   As String * 1        ' 品EPOSF2測定条件
    HEPOF2SH   As String * 1        ' 品EPOSF2測定位置_方
    HEPOF2ST   As String * 1        ' 品EPOSF2測定位置_点
    HEPOF2SR   As String * 1        ' 品EPOSF2測定位置_領
    HEPOF2HT   As String * 1        ' 品EPOSF2保証方法_対
    HEPOF2HS   As String * 1        ' 品EPOSF2保証方法_処
    HEPOF2KM   As String * 1        ' 品EPOSF2検査頻度_枚
    HEPOF2KN   As String * 1        ' 品EPOSF2検査頻度_抜
    HEPOF2KH   As String * 1        ' 品EPOSF2検査頻度_保
    HEPOF2KU   As String * 1        ' 品EPOSF2検査頻度_ｳ
    HEPOSF2PTK As String * 1        ' 品EPOSF2ﾊﾟﾀﾝ区分
    HEPOF3AX   As Double            ' 品EPOSF3平均上限
    HEPOF3MX   As Double            ' 品EPOSF3上限
    HEPOF3ET   As Double            ' 品EPOSF3選択ET代
    HEPOF3NS   As String * 2        ' 品EPOSF3熱処理法
    HEPOF3SZ   As String * 1        ' 品EPOSF3測定条件
    HEPOF3SH   As String * 1        ' 品EPOSF3測定位置_方
    HEPOF3ST   As String * 1        ' 品EPOSF3測定位置_点
    HEPOF3SR   As String * 1        ' 品EPOSF3測定位置_領
    HEPOF3HT   As String * 1        ' 品EPOSF3保証方法_対
    HEPOF3HS   As String * 1        ' 品EPOSF3保証方法_処
    HEPOF3KM   As String * 1        ' 品EPOSF3検査頻度_枚
    HEPOF3KN   As String * 1        ' 品EPOSF3検査頻度_抜
    HEPOF3KH   As String * 1        ' 品EPOSF3検査頻度_保
    HEPOF3KU   As String * 1        ' 品EPOSF3検査頻度_ｳ
    HEPOSF3PTK As String * 1        ' 品EPOSF3ﾊﾟﾀﾝ区分
    HEPBM1AN   As Double            ' 品EPBMD1平均下限
    HEPBM1AX   As Double            ' 品EPBMD1平均上限
    HEPBM1ET   As Double            ' 品EPBMD1選択ET代
    HEPBM1NS   As String * 2        ' 品EPBMD1熱処理法
    HEPBM1SZ   As String * 1        ' 品EPBMD1測定条件
    HEPBM1SH   As String * 1        ' 品EPBMD1測定位置_方
    HEPBM1ST   As String * 1        ' 品EPBMD1測定位置_点
    HEPBM1SR   As String * 1        ' 品EPBMD1測定位置_領
    HEPBM1HT   As String * 1        ' 品EPBMD1保証方法_対
    HEPBM1HS   As String * 1        ' 品EPBMD1保証方法_処
    HEPBM1KM   As String * 1        ' 品EPBMD1検査頻度_枚
    HEPBM1KN   As String * 1        ' 品EPBMD1検査頻度_抜
    HEPBM1KH   As String * 1        ' 品EPBMD1検査頻度_保
    HEPBM1KU   As String * 1        ' 品EPBMD1検査頻度_ｳ
    HEPBM1MBP  As Double            ' 品EPBMD1面内分布
    HEPBM1MCL  As String * 2        ' 品EPBMD1面内計算
    HEPBM2AN   As Double            ' 品EPBMD2平均下限
    HEPBM2AX   As Double            ' 品EPBMD2平均上限
    HEPBM2ET   As Double            ' 品EPBMD2選択ET代
    HEPBM2NS   As String * 2        ' 品EPBMD2熱処理法
    HEPBM2SZ   As String * 1        ' 品EPBMD2測定条件
    HEPBM2SH   As String * 1        ' 品EPBMD2測定位置_方
    HEPBM2ST   As String * 1        ' 品EPBMD2測定位置_点
    HEPBM2SR   As String * 1        ' 品EPBMD2測定位置_領
    HEPBM2HT   As String * 1        ' 品EPBMD2保証方法_対
    HEPBM2HS   As String * 1        ' 品EPBMD2保証方法_処
    HEPBM2KM   As String * 1        ' 品EPBMD2検査頻度_枚
    HEPBM2KN   As String * 1        ' 品EPBMD2検査頻度_抜
    HEPBM2KH   As String * 1        ' 品EPBMD2検査頻度_保
    HEPBM2KU   As String * 1        ' 品EPBMD2検査頻度_ｳ
    HEPBM2MBP  As Double            ' 品EPBMD2面内分布
    HEPBM2MCL  As String * 2        ' 品EPBMD2面内計算
    HEPBM3AN   As Double            ' 品EPBMD3平均下限
    HEPBM3AX   As Double            ' 品EPBMD3平均上限
    HEPBM3GSAN As Double            ' 品EPBMD3平均下限(外周)　09/05/07 ooba
    HEPBM3GSAX As Double            ' 品EPBMD3平均上限(外周)　09/05/07 ooba
    HEPBM3ET   As Double            ' 品EPBMD3選択ET代
    HEPBM3NS   As String * 2        ' 品EPBMD3熱処理法
    HEPBM3SZ   As String * 1        ' 品EPBMD3測定条件
    HEPBM3SH   As String * 1        ' 品EPBMD3測定位置_方
    HEPBM3ST   As String * 1        ' 品EPBMD3測定位置_点
    HEPBM3SR   As String * 1        ' 品EPBMD3測定位置_領
    HEPBM3HT   As String * 1        ' 品EPBMD3保証方法_対
    HEPBM3HS   As String * 1        ' 品EPBMD3保証方法_処
    HEPBM3KM   As String * 1        ' 品EPBMD3検査頻度_枚
    HEPBM3KN   As String * 1        ' 品EPBMD3検査頻度_抜
    HEPBM3KH   As String * 1        ' 品EPBMD3検査頻度_保
    HEPBM3KU   As String * 1        ' 品EPBMD3検査頻度_ｳ
    HEPBM3MBP  As Double            ' 品EPBMD3面内分布
    HEPBM3MCL  As String * 2        ' 品EPBMD3面内計算

'--------------- 2008/08/25 INSERT START  By Systech ---------------
    HSXDKTMP    As String * 1       ' DK温度（仕様）
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    HSXLDLRMN   As Integer          ' 品SXL/DL連続0下限
    HSXLDLRMX   As Integer          ' 品SXL/DL連続0上限
    HWFLDLRMN   As Integer          ' 品WFL/DL連続0下限
    HWFLDLRMX   As Integer          ' 品WFL/DL連続0上限
    HWFGDPTK    As String * 1       ' 品ＷＦＧＤパタン区分
    HSXGDPTK    As String * 1       ' 品ＳＸＧＤパタン区分
    WFHSGDCW    As String * 1       ' 保証FLG（GD)
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
    
'' ↓Add 2008/10/01 SIRD対応 Y.Hitomi
    HWFSIRDMX   As Integer          ' 軸状転位上限
    HWFSIRDSZ   As String * 1       ' 軸状転位測定条件
    HWFSIRDHT   As String * 1       ' 軸状転位保証方法＿対
    HWFSIRDHS   As String * 1       ' 軸状転位保証方法＿処
    HWFSIRDKM   As String * 1       ' 軸状転位検査頻度＿枚
    HWFSIRDKN   As String * 1       ' 軸状転位検査頻度＿抜
    HWFSIRDKH   As String * 1       ' 軸状転位検査頻度＿保
    HWFSIRDKU   As String * 1       ' 軸状転位検査頻度＿ウ
'' ↑Add 2008/10/01 SIRD対応 Y.Hitomi
End Type

'*******************************************************************************************
'*    関数名        : SetInitData
'*
'*    処理概要      : 1.初期設定処理
'*
'*    パラメータ    : 変数名      ,IO  ,型                           ,説明
'*                   intSXLID     ,I   ,String                       ,SXL-ID
'*                   udtNew_Hinban,I   ,tFullHinban                  ,該当品番(構造体)
'*                   udtSXL       ,O   ,DBDRV_scmzc_fcmlc001b_SXL    ,SXL管理用
'*                   intSmpGetFlg ,I   ,Integer                      ,ｻﾝﾌﾟﾙ管理取得ﾌﾗｸﾞ(0:ｻﾝﾌﾟﾙ指定なし, 1:ｻﾝﾌﾟﾙ指定あり)
'*                   sSamplID1    ,I   ,String                       ,TOPｻﾝﾌﾟﾙID(省略可)
'*                   sSamplID2    ,I   ,String                       ,BOTｻﾝﾌﾟﾙID(省略可)
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************
Public Function SetInitData(intSXLID As String, udtNew_Hinban As tFullHinban, udtSXL As DBDRV_scmzc_fcmlc001b_SXL, _
                            intSmpGetFlg As Integer, sSamplID1 As String, sSamplID2 As String) As FUNCTION_RETURN

    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim intRecCnt   As Integer
    Dim i           As Integer
    Dim intIngotpos As Integer              ' 結晶内位置
    Dim intLength   As Integer              ' 長さ
    Dim sCryNum     As String               ' 結晶番号
    
    ' エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function SetInitData"

    SetInitData = FUNCTION_RETURN_SUCCESS

Debug.Print "1-1 " & Now & " SXL管理を取得 SQL実行"
    ' SXL管理を取得
    sSQL = "select "
    sSQL = sSQL & "xtalcb as CRYNUM, "      ' 結晶番号
    sSQL = sSQL & "inposcb as INGOTPOS, "   ' 結晶内開始位置
    sSQL = sSQL & "rlencb as LENGTH, "      ' 長さ
    sSQL = sSQL & "sxlidcb as SXLID, "      ' SXLID
    sSQL = sSQL & "' ' as KRPROCCD, "       ' 管理工程
    sSQL = sSQL & "gnwkntcb as NOWPROC, "   ' 現在工程
    sSQL = sSQL & "' ' as LPKRPROCCD, "     ' 最終通過管理工程
    sSQL = sSQL & "newkntcb as LASTPASS, "  ' 最終通過工程
    sSQL = sSQL & "livkcb as DELCLS, "      ' 削除区分
    sSQL = sSQL & "lstccb as LSTATCLS, "    ' 最終状態区分
    sSQL = sSQL & "sholdclscb HOLDCLS, "    ' ホールド区分
    sSQL = sSQL & "hinbcb as HINBAN, "      ' 品番
    sSQL = sSQL & "revnumcb as REVNUM, "    ' 製品番号改訂番号
    sSQL = sSQL & "factorycb as FACTORY, "  ' 工場
    sSQL = sSQL & "opecb as OPECOND, "      ' 操業条件
    sSQL = sSQL & "tdaycb as REGDATE, "     ' 登録日付
    sSQL = sSQL & "kdaycb as UPDDATE "      ' 更新日付
    sSQL = sSQL & " from XSDCB "
    sSQL = sSQL & " where sxlidcb = '" & intSXLID & "'"
    
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_DEFAULT)

    ' レコード0件時正常終了
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        SetInitData = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
Debug.Print "1-2 " & Now & " 配列にセット"
    With udtSXL
        .CRYNUM = rs("CRYNUM")                                          ' 結晶番号
        .INGOTPOS = rs("INGOTPOS")                                      ' 結晶内開始位置
        If IsNull(rs("LENGTH")) = False Then .Length = rs("LENGTH")     ' 長さ
        .SXLID = rs("SXLID")                                            ' SXLID
        .KRPROCCD = rs("KRPROCCD")                                      ' 管理工程
        .NOWPROC = rs("NOWPROC")                                        ' 現在工程
        .LPKRPROCCD = rs("LPKRPROCCD")                                  ' 最終通過管理工程
        .LASTPASS = rs("LASTPASS")                                      ' 最終通過工程
        .DELCLS = rs("DELCLS")                                          ' 削除区分
        .LSTATCLS = rs("LSTATCLS")                                      ' 最終状態区分
        If IsNull(rs("HOLDCLS")) = False Then .HOLDCLS = rs("HOLDCLS")  ' ホールド区分
        .hinban = rs("HINBAN")                                          ' 品番
        .REVNUM = rs("REVNUM")                                          ' 製品番号改訂番号
        .factory = rs("FACTORY")                                        ' 工場
        .opecond = rs("OPECOND")                                        ' 操業条件
        .REGDATE = rs("REGDATE")                                        ' 登録日付
        .UPDDATE = rs("UPDDATE")                                        ' 更新日付
    End With
    
    Set rs = Nothing
    
    ' 工程実績ﾃﾞｰﾀ取得関数からﾃﾞｰﾀを取得し設定する
    If intSmpGetFlg <> 0 Then
        If GET_hurikaeC3(intSXLID, wiKcnt, intIngotpos, intLength, sCryNum) = FUNCTION_RETURN_FAILURE Then
            SetInitData = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
            
        With udtSXL
            .CRYNUM = sCryNum                                           ' 結晶番号
            .INGOTPOS = intIngotpos                                     ' 結晶内開始位置
            .Length = intLength                                         ' 長さ
        End With
    End If
        
Debug.Print "2-1 " & Now & " 新サンプル管理(SXL)を取得 SQL実行"
    ' 新サンプル管理(SXL)を取得
    ' エピ先行評価追加対応
    sSQL = "select SXLIDCW, SMPKBNCW, TBKBNCW, REVNUMCW, XTALCW, INPOSCW, REPSMPLIDCW, HINBCW, FACTORYCW, OPECW, KTKBNCW, SMCRYNUMCW, "
    sSQL = sSQL & "WFSMPLIDRSCW, WFSMPLIDRS1CW, WFSMPLIDRS2CW, WFINDRSCW, WFRESRS1CW, WFRESRS2CW, WFSMPLIDOICW, WFINDOICW, WFRESOICW, "
    sSQL = sSQL & "WFSMPLIDB1CW, WFINDB1CW, WFRESB1CW, WFSMPLIDB2CW, WFINDB2CW, WFRESB2CW, WFSMPLIDB3CW, WFINDB3CW, WFRESB3CW, "
    sSQL = sSQL & "WFSMPLIDL1CW, WFINDL1CW, WFRESL1CW, WFSMPLIDL2CW, WFINDL2CW, WFRESL2CW, WFSMPLIDL3CW, WFINDL3CW, WFRESL3CW, "
    sSQL = sSQL & "WFSMPLIDL4CW, WFINDL4CW, WFRESL4CW, WFSMPLIDDSCW, WFINDDSCW, WFRESDSCW, WFSMPLIDDZCW, WFINDDZCW, WFRESDZCW, "
    sSQL = sSQL & "WFSMPLIDSPCW, WFINDSPCW, WFRESSPCW, WFSMPLIDDO1CW, WFINDDO1CW, WFRESDO1CW, WFSMPLIDDO2CW, WFINDDO2CW, WFRESDO2CW, "
    sSQL = sSQL & "WFSMPLIDDO3CW, WFINDDO3CW, WFRESDO3CW, WFSMPLIDOT1CW, WFINDOT1CW, WFRESOT1CW, WFSMPLIDOT2CW, WFINDOT2CW, WFRESOT2CW, "
    sSQL = sSQL & "WFSMPLIDAOICW , WFINDAOICW, WFRESAOICW, SMPLNUMCW, SMPLPATCW, TSTAFFCW, TDAYCW, KSTAFFCW, KDAYCW, SNDKCW, SNDDAYCW, "
    sSQL = sSQL & "WFSMPLIDGDCW, WFINDGDCW, WFRESGDCW, WFHSGDCW "
    sSQL = sSQL & ",EPSMPLIDB1CW, EPINDB1CW, EPRESB1CW, EPSMPLIDB2CW, EPINDB2CW, EPRESB2CW, EPSMPLIDB3CW, EPINDB3CW, EPRESB3CW, "
    sSQL = sSQL & "EPSMPLIDL1CW, EPINDL1CW, EPRESL1CW, EPSMPLIDL2CW, EPINDL2CW, EPRESL2CW, EPSMPLIDL3CW, EPINDL3CW, EPRESL3CW "
    sSQL = sSQL & "from XSDCW "
    
    If intSmpGetFlg = 0 Then        ' SXL-IDで検索(生死区分=生ﾛｯﾄ)
        sSQL = sSQL & "where SXLIDCW = '" & intSXLID & "' and "
        sSQL = sSQL & "      LIVKCW = '0' "
    Else                            ' 結晶番号とｻﾝﾌﾟﾙIDで検索
        sSQL = sSQL & "where XTALCW = substr('" & intSXLID & "', 1, 9) || '000' and "
        sSQL = sSQL & "      REPSMPLIDCW in ('" & sSamplID1 & "', '" & sSamplID2 & "') "
    End If
    sSQL = sSQL & "order by INPOSCW"
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_DEFAULT)

    ' レコード0件時正常終了
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        SetInitData = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

Debug.Print "2-2 " & Now & " 配列にセット"
    intRecCnt = rs.RecordCount
    ReDim udtSXL.WFSMP(intRecCnt)
    For i = 1 To intRecCnt
        With udtSXL.WFSMP(i)
            .SXLIDCW = rs("SXLIDCW")                                                            ' SXL-ID
            .SMPKBNCW = rs("SMPKBNCW")                                                          ' サンプル区分
            .TBKBNCW = rs("TBKBNCW")                                                            ' T/B区分
            
            If IsNull(rs("REPSMPLIDCW")) = False Then .REPSMPLIDCW = rs("REPSMPLIDCW")          ' 代表サンプルID
            If IsNull(rs("XTALCW")) = False Then .XTALCW = rs("XTALCW")                         ' 結晶番号
            If IsNull(rs("INPOSCW")) = False Then .INPOSCW = rs("INPOSCW")                      ' 結晶内位置
            If IsNull(rs("HINBCW")) = False Then .HINBCW = rs("HINBCW")                         ' 品番
            If IsNull(rs("REVNUMCW")) = False Then .REVNUMCW = rs("REVNUMCW")                   ' 製品番号改訂番号
            If IsNull(rs("FACTORYCW")) = False Then .FACTORYCW = rs("FACTORYCW")                ' 工場
            If IsNull(rs("OPECW")) = False Then .OPECW = rs("OPECW")                            ' 操業条件
            If IsNull(rs("KTKBNCW")) = False Then .KTKBNCW = rs("KTKBNCW")                      ' 確定区分
            If IsNull(rs("SMCRYNUMCW")) = False Then .SMCRYNUMCW = rs("SMCRYNUMCW")             ' ｻﾝﾌﾟﾙﾌﾞﾛｯｸID
            If IsNull(rs("WFSMPLIDRSCW")) = False Then .WFSMPLIDRSCW = rs("WFSMPLIDRSCW")       ' サンプルID(Rs)
            If IsNull(rs("WFSMPLIDRS1CW")) = False Then .WFSMPLIDRS1CW = rs("WFSMPLIDRS1CW")    ' 推定サンプルID1(Rs)
            If IsNull(rs("WFSMPLIDRS2CW")) = False Then .WFSMPLIDRS2CW = rs("WFSMPLIDRS2CW")    ' 推定サンプルID2(Rs)
            If IsNull(rs("WFINDRSCW")) = False Then .WFINDRSCW = rs("WFINDRSCW")                ' 状態FLG(Rs)
            If IsNull(rs("WFRESRS1CW")) = False Then .WFRESRS1CW = rs("WFRESRS1CW")             ' 実績FLG1(Rs)
            If IsNull(rs("WFRESRS2CW")) = False Then .WFRESRS2CW = rs("WFRESRS2CW")             ' 実績FLG2(Rs)
            If IsNull(rs("WFSMPLIDOICW")) = False Then .WFSMPLIDOICW = rs("WFSMPLIDOICW")       ' サンプルID(Oi)
            If IsNull(rs("WFINDOICW")) = False Then .WFINDOICW = rs("WFINDOICW")                ' 状態FLG(Oi)
            If IsNull(rs("WFRESOICW")) = False Then .WFRESOICW = rs("WFRESOICW")                ' 実績FLG(Oi)
            If IsNull(rs("WFSMPLIDB1CW")) = False Then .WFSMPLIDB1CW = rs("WFSMPLIDB1CW")       ' サンプルID(B1)
            If IsNull(rs("WFINDB1CW")) = False Then .WFINDB1CW = rs("WFINDB1CW")                ' 状態FLG(B1)
            If IsNull(rs("WFRESB1CW")) = False Then .WFRESB1CW = rs("WFRESB1CW")                ' 実績FLG(B1)
            If IsNull(rs("WFSMPLIDB2CW")) = False Then .WFSMPLIDB2CW = rs("WFSMPLIDB2CW")       ' サンプルID(B2)
            If IsNull(rs("WFINDB2CW")) = False Then .WFINDB2CW = rs("WFINDB2CW")                ' 状態FLG(B2)
            If IsNull(rs("WFRESB2CW")) = False Then .WFRESB2CW = rs("WFRESB2CW")                ' 実績FLG(B2)
            If IsNull(rs("WFSMPLIDB3CW")) = False Then .WFSMPLIDB3CW = rs("WFSMPLIDB3CW")       ' サンプルID(B3)
            If IsNull(rs("WFINDB3CW")) = False Then .WFINDB3CW = rs("WFINDB3CW")                ' 状態FLG(B3)
            If IsNull(rs("WFRESB3CW")) = False Then .WFRESB3CW = rs("WFRESB3CW")                ' 実績FLG(B3)
            If IsNull(rs("WFSMPLIDL1CW")) = False Then .WFSMPLIDL1CW = rs("WFSMPLIDL1CW")       ' サンプルID(L1)
            If IsNull(rs("WFINDL1CW")) = False Then .WFINDL1CW = rs("WFINDL1CW")                ' 状態FLG(L1)
            If IsNull(rs("WFRESL1CW")) = False Then .WFRESL1CW = rs("WFRESL1CW")                ' 実績FLG(L1)
            If IsNull(rs("WFSMPLIDL2CW")) = False Then .WFSMPLIDL2CW = rs("WFSMPLIDL2CW")       ' サンプルID(L2)
            If IsNull(rs("WFINDL2CW")) = False Then .WFINDL2CW = rs("WFINDL2CW")                ' 状態FLG(L2)
            If IsNull(rs("WFRESL2CW")) = False Then .WFRESL2CW = rs("WFRESL2CW")                ' 実績FLG(L2)
            If IsNull(rs("WFSMPLIDL3CW")) = False Then .WFSMPLIDL3CW = rs("WFSMPLIDL3CW")       ' サンプルID(L3)
            If IsNull(rs("WFINDL3CW")) = False Then .WFINDL3CW = rs("WFINDL3CW")                ' 状態FLG(L3)
            If IsNull(rs("WFRESL3CW")) = False Then .WFRESL3CW = rs("WFRESL3CW")                ' 実績FLG(L3)
            If IsNull(rs("WFSMPLIDL4CW")) = False Then .WFSMPLIDL4CW = rs("WFSMPLIDL4CW")       ' サンプルID(L4)
            If IsNull(rs("WFINDL4CW")) = False Then .WFINDL4CW = rs("WFINDL4CW")                ' 状態FLG(L4)
            If IsNull(rs("WFRESL4CW")) = False Then .WFRESL4CW = rs("WFRESL4CW")                ' 実績FLG(L4)
            If IsNull(rs("WFSMPLIDDSCW")) = False Then .WFSMPLIDDSCW = rs("WFSMPLIDDSCW")       ' サンプルID(DS)
            If IsNull(rs("WFINDDSCW")) = False Then .WFINDDSCW = rs("WFINDDSCW")                ' 状態FLG(DS)
            If IsNull(rs("WFRESDSCW")) = False Then .WFRESDSCW = rs("WFRESDSCW")                ' 実績FLG(DS)
            If IsNull(rs("WFSMPLIDDZCW")) = False Then .WFSMPLIDDZCW = rs("WFSMPLIDDZCW")       ' サンプルID(DZ)
            If IsNull(rs("WFINDDZCW")) = False Then .WFINDDZCW = rs("WFINDDZCW")                ' 状態FLG(DZ)
            If IsNull(rs("WFRESDZCW")) = False Then .WFRESDZCW = rs("WFRESDZCW")                ' 実績FLG(DZ)
            If IsNull(rs("WFSMPLIDSPCW")) = False Then .WFSMPLIDSPCW = rs("WFSMPLIDSPCW")       ' サンプルID(SP)
            If IsNull(rs("WFINDSPCW")) = False Then .WFINDSPCW = rs("WFINDSPCW")                ' 状態FLG(SP)
            If IsNull(rs("WFRESSPCW")) = False Then .WFRESSPCW = rs("WFRESSPCW")                ' 実績FLG(SP)
            If IsNull(rs("WFSMPLIDDO1CW")) = False Then .WFSMPLIDDO1CW = rs("WFSMPLIDDO1CW")    ' サンプルID(DO1)
            If IsNull(rs("WFINDDO1CW")) = False Then .WFINDDO1CW = rs("WFINDDO1CW")             ' 状態FLG(DO1)
            If IsNull(rs("WFRESDO1CW")) = False Then .WFRESDO1CW = rs("WFRESDO1CW")             ' 実績FLG(DO1)
            If IsNull(rs("WFSMPLIDDO2CW")) = False Then .WFSMPLIDDO2CW = rs("WFSMPLIDDO2CW")    ' サンプルID(DO2)
            If IsNull(rs("WFINDDO2CW")) = False Then .WFINDDO2CW = rs("WFINDDO2CW")             ' 状態FLG(DO2)
            If IsNull(rs("WFRESDO2CW")) = False Then .WFRESDO2CW = rs("WFRESDO2CW")             ' 実績FLG(DO2)
            If IsNull(rs("WFSMPLIDDO3CW")) = False Then .WFSMPLIDDO3CW = rs("WFSMPLIDDO3CW")    ' サンプルID(DO3)
            If IsNull(rs("WFINDDO3CW")) = False Then .WFINDDO3CW = rs("WFINDDO3CW")             ' 状態FLG(DO3)
            If IsNull(rs("WFRESDO3CW")) = False Then .WFRESDO3CW = rs("WFRESDO3CW")             ' 実績FLG(DO3)
            If IsNull(rs("WFSMPLIDOT1CW")) = False Then .WFSMPLIDOT1CW = rs("WFSMPLIDOT1CW")    ' サンプルID(OT1)
            If IsNull(rs("WFINDOT1CW")) = False Then .WFINDOT1CW = rs("WFINDOT1CW")             ' 状態FLG(OT1)
            If IsNull(rs("WFRESOT1CW")) = False Then .WFRESOT1CW = rs("WFRESOT1CW")             ' 実績FLG(OT1)
            If IsNull(rs("WFSMPLIDOT2CW")) = False Then .WFSMPLIDOT2CW = rs("WFSMPLIDOT2CW")    ' サンプルID(OT2)
            If IsNull(rs("WFINDOT2CW")) = False Then .WFINDOT2CW = rs("WFINDOT2CW")             ' 状態FLG(OT2)
            If IsNull(rs("WFRESOT2CW")) = False Then .WFRESOT2CW = rs("WFRESOT2CW")             ' 実績FLG(OT2)
            If IsNull(rs("WFSMPLIDAOICW")) = False Then .WFSMPLIDAOICW = rs("WFSMPLIDAOICW")    ' サンプルID(AOI)
            If IsNull(rs("WFINDAOICW")) = False Then .WFINDAOICW = rs("WFINDAOICW")             ' 状態FLG(AOI)
            If IsNull(rs("WFRESAOICW")) = False Then .WFRESAOICW = rs("WFRESAOICW")             ' 実績FLG(AOI)
            If IsNull(rs("SMPLNUMCW")) = False Then .SMPLNUMCW = rs("SMPLNUMCW")                ' ｻﾝﾌﾟﾙ枚数
            If IsNull(rs("SMPLPATCW")) = False Then .SMPLPATCW = rs("SMPLPATCW")                ' ｻﾝﾌﾟﾙﾊﾟﾀｰﾝ
            If IsNull(rs("TSTAFFCW")) = False Then .TSTAFFCW = rs("TSTAFFCW")                   ' 登録社員ID
            If IsNull(rs("TDAYCW")) = False Then .TDAYCW = rs("TDAYCW")                         ' 登録日付
            If IsNull(rs("KSTAFFCW")) = False Then .KSTAFFCW = rs("KSTAFFCW")                   ' 更新社員ID
            If IsNull(rs("KDAYCW")) = False Then .KDAYCW = rs("KDAYCW")                         ' 更新日付
            If IsNull(rs("SNDKCW")) = False Then .SNDKCW = rs("SNDKCW")                         ' 送信ﾌﾗｸﾞ
            If IsNull(rs("SNDDAYCW")) = False Then .SNDDAYCW = rs("SNDDAYCW")                   ' 送信日付
            If IsNull(rs("WFSMPLIDGDCW")) = False Then .WFSMPLIDGDCW = rs("WFSMPLIDGDCW")       ' サンプルID(GD)
            If IsNull(rs("WFINDGDCW")) = False Then .WFINDGDCW = rs("WFINDGDCW")                ' 状態FLG(GD)
            If IsNull(rs("WFRESGDCW")) = False Then .WFRESGDCW = rs("WFRESGDCW")                ' 実績FLG(GD)
            If IsNull(rs("WFHSGDCW")) = False Then .WFHSGDCW = rs("WFHSGDCW")                   ' 保証FLG(GD)
            
            ' エピ先行評価追加
            If IsNull(rs("EPSMPLIDB1CW")) = False Then .EPSMPLIDB1CW = rs("EPSMPLIDB1CW")       ' サンプルID(B1E)
            If IsNull(rs("EPINDB1CW")) = False Then .EPINDB1CW = rs("EPINDB1CW")                ' 状態FLG(B1E)
            If IsNull(rs("EPRESB1CW")) = False Then .EPRESB1CW = rs("EPRESB1CW")                ' 実績FLG(B1E)
            If IsNull(rs("EPSMPLIDB2CW")) = False Then .EPSMPLIDB2CW = rs("EPSMPLIDB2CW")       ' サンプルID(B2E)
            If IsNull(rs("EPINDB2CW")) = False Then .EPINDB2CW = rs("EPINDB2CW")                ' 状態FLG(B2E)
            If IsNull(rs("EPRESB2CW")) = False Then .EPRESB2CW = rs("EPRESB2CW")                ' 実績FLG(B2E)
            If IsNull(rs("EPSMPLIDB3CW")) = False Then .EPSMPLIDB3CW = rs("EPSMPLIDB3CW")       ' サンプルID(B3E)
            If IsNull(rs("EPINDB3CW")) = False Then .EPINDB3CW = rs("EPINDB3CW")                ' 状態FLG(B3E)
            If IsNull(rs("EPRESB3CW")) = False Then .EPRESB3CW = rs("EPRESB3CW")                ' 実績FLG(B3E)
            If IsNull(rs("EPSMPLIDL1CW")) = False Then .EPSMPLIDL1CW = rs("EPSMPLIDL1CW")       ' サンプルID(L1E)
            If IsNull(rs("EPINDL1CW")) = False Then .EPINDL1CW = rs("EPINDL1CW")                ' 状態FLG(L1E)
            If IsNull(rs("EPRESL1CW")) = False Then .EPRESL1CW = rs("EPRESL1CW")                ' 実績FLG(L1E)
            If IsNull(rs("EPSMPLIDL2CW")) = False Then .EPSMPLIDL2CW = rs("EPSMPLIDL2CW")       ' サンプルID(L2E)
            If IsNull(rs("EPINDL2CW")) = False Then .EPINDL2CW = rs("EPINDL2CW")                ' 状態FLG(L2E)
            If IsNull(rs("EPRESL2CW")) = False Then .EPRESL2CW = rs("EPRESL2CW")                ' 実績FLG(L2E)
            If IsNull(rs("EPSMPLIDL3CW")) = False Then .EPSMPLIDL3CW = rs("EPSMPLIDL3CW")       ' サンプルID(L3E)
            If IsNull(rs("EPINDL3CW")) = False Then .EPINDL3CW = rs("EPINDL3CW")                ' 状態FLG(L3E)
            If IsNull(rs("EPRESL3CW")) = False Then .EPRESL3CW = rs("EPRESL3CW")                ' 実績FLG(L3E)
        End With
        rs.MoveNext
    Next i
    
    Set rs = Nothing

Debug.Print "3 " & Now & " 欠落有無を取得"
    ' 欠落情報取得
    If KeturakuInfo(udtSXL) = FUNCTION_RETURN_FAILURE Then
        SetInitData = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

Debug.Print "4 " & Now & " 枚数を取得"
    ' 枚数取得
    If GetMaisu(udtSXL) = FUNCTION_RETURN_FAILURE Then
        SetInitData = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
Debug.Print "8 " & Now

    SetInitData = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' 終了
'    gErr.Pop
    Exit Function

proc_err:
    ' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    SetInitData = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    関数名        :
'*
'*    処理概要      : 1.欠落有無取得
'*
'*    パラメータ    : 変数名        ,IO ,型                        ,説明
'*                    udtSXL        ,IO ,DBDRV_scmzc_fcmlc001b_SXL ,SXL管理
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Function KeturakuInfo(udtSXL As DBDRV_scmzc_fcmlc001b_SXL) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim i           As Long
    Dim lngRecCnt   As Long
    Dim sSXLID      As String

    ' エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function KeturakuInfo"

    KeturakuInfo = FUNCTION_RETURN_SUCCESS

    sSQL = "select distinct SXL.SXLIDCB "
    sSQL = sSQL & "from XSDCB SXL, TBCME040 BLK, TBCMY012 REJ "
    sSQL = sSQL & "where"
    sSQL = sSQL & "  REJ.LOTID=BLK.BLOCKID"
    sSQL = sSQL & "  and SXL.SXLIDCB = '" & udtSXL.SXLID & "'"
    sSQL = sSQL & "  and SXL.XTALCB=BLK.CRYNUM"
    sSQL = sSQL & "  and SXL.LIVKCB<>'1'"
    sSQL = sSQL & "  and ("
    sSQL = sSQL & "    ("
    sSQL = sSQL & "      REJ.ALLSCRAP='Y'"
    sSQL = sSQL & "      and SXL.INPOSCB<BLK.INGOTPOS+BLK.LENGTH"
    sSQL = sSQL & "      and SXL.INPOSCB+SXL.RLENCB>BLK.INGOTPOS"
    sSQL = sSQL & "    ) or ("
    sSQL = sSQL & "      REJ.ALLSCRAP='N'"
    sSQL = sSQL & "      and REJ.REJCAT='A'"
    sSQL = sSQL & "      and (SXL.INPOSCB < BLK.INGOTPOS + REJ.LENTO)"
    sSQL = sSQL & "      and (SXL.INPOSCB + SXL.RLENCB > BLK.INGOTPOS + REJ.LENFROM)"
    sSQL = sSQL & "    ) or ("
    sSQL = sSQL & "      REJ.REJCAT='B'"
    sSQL = sSQL & "      and BLK.INGOTPOS + REJ.TOP_POS/10.0 between SXL.INPOSCB and SXL.INPOSCB + SXL.RLENCB"
    sSQL = sSQL & "    )"
    sSQL = sSQL & "  )"
    
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    ' SQL結果のSXLIDが欠落ありのSXLID
    If rs.RecordCount = 0 Then
        udtSXL.KETURAKU = False
    Else
        udtSXL.KETURAKU = True
    End If
    Set rs = Nothing

proc_exit:
    ' 終了
'    gErr.Pop
    Exit Function

proc_err:
    ' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    KeturakuInfo = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    関数名        : GetMaisu
'*
'*    処理概要      : 1.WF枚数取得
'*
'*    パラメータ    : 変数名        ,IO ,型                        ,説明
'*                    udtSXL        ,IO ,DBDRV_scmzc_fcmlc001b_SXL ,SXL管理
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Function GetMaisu(udtSXL As DBDRV_scmzc_fcmlc001b_SXL) As FUNCTION_RETURN
    Dim sSQL    As String
    Dim rs      As OraDynaset

    ' エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function getMaisu"

    GetMaisu = FUNCTION_RETURN_SUCCESS

    sSQL = sSQL & "select MAICB from XSDCB where SXLIDCB = '" & udtSXL.SXLID & "'"
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount = 0 Then
        udtSXL.COUNT = 0
    Else
        udtSXL.COUNT = rs("MAICB")
    End If
    Set rs = Nothing

proc_exit:
    ' 終了
    Exit Function

proc_err:
    ' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    GetMaisu = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    関数名        : funWfcGetDataEtc
'*
'*    処理概要      : 1.WF総合判定 各種データ取得
'*
'*    パラメータ    : 変数名      ,IO  ,型                                 ,説明
'*               　　udtTypIn     ,I   ,type_DBDRV_scmzc_fcmlc001c_In      ,入力用
'*               　　Siyou        ,O   ,type_DBDRV_scmzc_fcmlc001c_Siyou   ,WF仕様用
'*               　　udtSokutei   ,O   ,typ_TBCMY013                       ,測定評価結果
'*               　　sErrMsg 　　 ,O   ,String    　　　　　　　　　　　   ,エラーメッセージ
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function funWfcGetDataEtc(udtTypIn As type_DBDRV_scmzc_fcmlc001c_In, udtNew_Hinban As tFullHinban, intSmpGetFlg As Integer, _
                                 udtSiyou As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                                 udtSokutei() As typ_TBCMY013, _
                                 sErrMsg As String) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim intRecCnt   As Integer
    Dim i           As Long
    Dim sDBName     As String
    Dim intPos      As Integer      ' ｻﾝﾌﾟﾙ位置
    
    ' エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function funWfcGetDataEtc"

    funWfcGetDataEtc = FUNCTION_RETURN_SUCCESS
    
    ' WF仕様取得
    sDBName = "V001"
    sSQL = "select "
    sSQL = sSQL & "E021HWFTYPE, "            ' 品ＷＦタイプ
    sSQL = sSQL & "E022HWFCDIR, "            ' 品ＷＦ結晶面方
    sSQL = sSQL & "E023HWFCDOP, "            ' 品ＷＦ結晶ドープ
    
    sSQL = sSQL & "E021HWFRMIN, "            ' 品ＷＦ比抵抗下限
    sSQL = sSQL & "E021HWFRMAX, "            ' 品ＷＦ比抵抗上限
    sSQL = sSQL & "E021HWFRSPOH, "           ' 品ＷＦ比抵抗測定位置＿方
    sSQL = sSQL & "E021HWFRSPOT, "           ' 品ＷＦ比抵抗測定位置＿点
    sSQL = sSQL & "E021HWFRSPOI, "           ' 品ＷＦ比抵抗測定位置＿位
    sSQL = sSQL & "E021HWFRHWYT, "           ' 品ＷＦ比抵抗保証方法＿対
    sSQL = sSQL & "E021HWFRHWYS, "           ' 品ＷＦ比抵抗保証方法＿処
    sSQL = sSQL & "E021HWFRMCAL, "           ' 品ＷＦ比抵抗面内計算
    sSQL = sSQL & "E021HWFRAMIN, "           ' 品ＷＦ比抵抗平均下限
    sSQL = sSQL & "E021HWFRAMAX, "           ' 品ＷＦ比抵抗平均上限
    sSQL = sSQL & "E021HWFRMBNP, "           ' 品ＷＦ比抵抗面内分布

    sSQL = sSQL & "E024HWFMKMIN, "           ' 品ＷＦ無欠陥層下限
    sSQL = sSQL & "E024HWFMKMAX, "           ' 品ＷＦ無欠陥層上限
    sSQL = sSQL & "E024HWFMKSPH, "           ' 品ＷＦ無欠陥層測定位置＿方
    sSQL = sSQL & "E024HWFMKSPT, "           ' 品ＷＦ無欠陥層測定位置＿点
    sSQL = sSQL & "E024HWFMKSPR, "           ' 品ＷＦ無欠陥層測定位置＿領
    sSQL = sSQL & "E024HWFMKHWT, "           ' 品ＷＦ無欠陥層保証方法＿対
    sSQL = sSQL & "E024HWFMKHWS, "           ' 品ＷＦ無欠陥層保証方法＿処

    sSQL = sSQL & "E025HWFONMIN, "           ' 品ＷＦ酸素濃度下限
    sSQL = sSQL & "E025HWFONMAX, "           ' 品ＷＦ酸素濃度上限
    sSQL = sSQL & "E025HWFONSPH, "           ' 品ＷＦ酸素濃度測定位置＿方
    sSQL = sSQL & "E025HWFONSPT, "           ' 品ＷＦ酸素濃度測定位置＿点
    sSQL = sSQL & "E025HWFONSPI, "           ' 品ＷＦ酸素濃度測定位置＿位
    sSQL = sSQL & "E025HWFONHWT, "           ' 品ＷＦ酸素濃度保証方法＿対
    sSQL = sSQL & "E025HWFONHWS, "           ' 品ＷＦ酸素濃度保証方法＿処
    sSQL = sSQL & "E025HWFONMCL, "           ' 品ＷＦ酸素濃度面内計算
    sSQL = sSQL & "E025HWFONMBP, "           ' 品ＷＦ酸素濃度面内分布
    sSQL = sSQL & "E025HWFONAMN, "           ' 品ＷＦ酸素濃度平均下限
    sSQL = sSQL & "E025HWFONAMX, "           ' 品ＷＦ酸素濃度平均上限

    sSQL = sSQL & "E025HWFOS1MN, "           ' 品ＷＦ酸素析出１下限
    sSQL = sSQL & "E025HWFOS1MX, "           ' 品ＷＦ酸素析出１上限
    sSQL = sSQL & "E025HWFOS1SH, "           ' 品ＷＦ酸素析出１測定位置＿方
    sSQL = sSQL & "E025HWFOS1ST, "           ' 品ＷＦ酸素析出１測定位置＿点
    sSQL = sSQL & "E025HWFOS1SI, "           ' 品ＷＦ酸素析出１測定位置＿位
    sSQL = sSQL & "E025HWFOS1HT, "           ' 品ＷＦ酸素析出１保証方法＿対
    sSQL = sSQL & "E025HWFOS1HS, "           ' 品ＷＦ酸素析出１保証方法＿処
    sSQL = sSQL & "E025HWFOS2SH, "           ' 品ＷＦ酸素析出２測定位置＿方
    sSQL = sSQL & "E025HWFOS2ST, "           ' 品ＷＦ酸素析出２測定位置＿点
    sSQL = sSQL & "E025HWFOS2SI, "           ' 品ＷＦ酸素析出２測定位置＿位
    sSQL = sSQL & "E025HWFOS2MN, "           ' 品ＷＦ酸素析出２下限
    sSQL = sSQL & "E025HWFOS2MX, "           ' 品ＷＦ酸素析出２上限
    sSQL = sSQL & "E025HWFOS2HT, "           ' 品ＷＦ酸素析出２保証方法＿対
    sSQL = sSQL & "E025HWFOS2HS, "           ' 品ＷＦ酸素析出２保証方法＿処
    sSQL = sSQL & "E025HWFOS3MN, "           ' 品ＷＦ酸素析出３下限
    sSQL = sSQL & "E025HWFOS3MX, "           ' 品ＷＦ酸素析出３上限
    sSQL = sSQL & "E025HWFOS3SH, "           ' 品ＷＦ酸素析出３測定位置＿方
    sSQL = sSQL & "E025HWFOS3ST, "           ' 品ＷＦ酸素析出３測定位置＿点
    sSQL = sSQL & "E025HWFOS3SI, "           ' 品ＷＦ酸素析出３測定位置＿位
    sSQL = sSQL & "E025HWFOS3HT, "           ' 品ＷＦ酸素析出３保証方法＿対
    sSQL = sSQL & "E025HWFOS3HS, "           ' 品ＷＦ酸素析出３保証方法＿処

    sSQL = sSQL & "E026HWFDSOMX, "           ' 品ＷＦＤＳＯＤ上限
    sSQL = sSQL & "E026HWFDSOMN, "           ' 品ＷＦＤＳＯＤ下限
    sSQL = sSQL & "E026HWFDSOAX, "           ' 品ＷＦＤＳＯＤ領域上限
    sSQL = sSQL & "E026HWFDSOAN, "           ' 品ＷＦＤＳＯＤ領域下限
    sSQL = sSQL & "E026HWFDSOHT, "           ' 品ＷＦＤＳＯＤ保証方法＿対
    sSQL = sSQL & "E026HWFDSOHS, "           ' 品ＷＦＤＳＯＤ保証方法＿処
    
    sSQL = sSQL & "E028HWFSPVMX, "           ' 品ＷＦＳＰＶＦＥ上限
    sSQL = sSQL & "E028HWFSPVSH, "           ' 品ＷＦＳＰＶＦＥ測定位置＿方
    sSQL = sSQL & "E028HWFSPVST, "           ' 品ＷＦＳＰＶＦＥ測定位置＿点
    sSQL = sSQL & "E028HWFSPVSI, "           ' 品ＷＦＳＰＶＦＥ測定位置＿位
    sSQL = sSQL & "E028HWFSPVHT, "           ' 品ＷＦＳＰＶＦＥ保証方法＿対
    sSQL = sSQL & "E028HWFSPVHS, "           ' 品ＷＦＳＰＶＦＥ保証方法＿処
    sSQL = sSQL & "E028HWFDLSPH, "           ' 品ＷＦ拡散長測定位置＿方
    sSQL = sSQL & "E028HWFDLSPT, "           ' 品ＷＦ拡散長測定位置＿点
    sSQL = sSQL & "E028HWFDLSPI, "           ' 品ＷＦ拡散長測定位置＿位
    sSQL = sSQL & "E028HWFDLHWT, "           ' 品ＷＦ拡散長保証方法＿対
    sSQL = sSQL & "E028HWFDLHWS, "           ' 品ＷＦ拡散長保証方法＿処
    sSQL = sSQL & "E028HWFDLMIN, "           ' 品ＷＦ拡散長下限
    sSQL = sSQL & "E028HWFDLMAX, "           ' 品ＷＦ拡散長上限
    
    sSQL = sSQL & "E029HWFOF1AX, "           ' 品ＷＦＯＳＦ１平均上限
    sSQL = sSQL & "E029HWFOF1MX, "           ' 品ＷＦＯＳＦ１上限
    sSQL = sSQL & "E029HWFOF1SH, "           ' 品ＷＦＯＳＦ１測定位置＿方
    sSQL = sSQL & "E029HWFOF1ST, "           ' 品ＷＦＯＳＦ１測定位置＿点
    sSQL = sSQL & "E029HWFOF1SR, "           ' 品ＷＦＯＳＦ１測定位置＿領
    sSQL = sSQL & "E029HWFOF1HT, "           ' 品ＷＦＯＳＦ１保証方法＿対
    sSQL = sSQL & "E029HWFOF1HS, "           ' 品ＷＦＯＳＦ１保証方法＿処
    sSQL = sSQL & "E029HWFOF2AX, "           ' 品ＷＦＯＳＦ２平均上限
    sSQL = sSQL & "E029HWFOF2MX, "           ' 品ＷＦＯＳＦ２上限
    sSQL = sSQL & "E029HWFOF2SH, "           ' 品ＷＦＯＳＦ２測定位置＿方
    sSQL = sSQL & "E029HWFOF2ST, "           ' 品ＷＦＯＳＦ２測定位置＿点
    sSQL = sSQL & "E029HWFOF2SR, "           ' 品ＷＦＯＳＦ２測定位置＿領
    sSQL = sSQL & "E029HWFOF2HT, "           ' 品ＷＦＯＳＦ２保証方法＿対
    sSQL = sSQL & "E029HWFOF2HS, "           ' 品ＷＦＯＳＦ２保証方法＿処
    sSQL = sSQL & "E029HWFOF3AX, "           ' 品ＷＦＯＳＦ３平均上限
    sSQL = sSQL & "E029HWFOF3MX, "           ' 品ＷＦＯＳＦ３上限
    sSQL = sSQL & "E029HWFOF3SH, "           ' 品ＷＦＯＳＦ３測定位置＿方
    sSQL = sSQL & "E029HWFOF3ST, "           ' 品ＷＦＯＳＦ３測定位置＿点
    sSQL = sSQL & "E029HWFOF3SR, "           ' 品ＷＦＯＳＦ３測定位置＿領
    sSQL = sSQL & "E029HWFOF3HT, "           ' 品ＷＦＯＳＦ３保証方法＿対
    sSQL = sSQL & "E029HWFOF3HS, "           ' 品ＷＦＯＳＦ３保証方法＿処
    sSQL = sSQL & "E029HWFOF4AX, "           ' 品ＷＦＯＳＦ４平均上限
    sSQL = sSQL & "E029HWFOF4MX, "           ' 品ＷＦＯＳＦ４上限
    sSQL = sSQL & "E029HWFOF4SH, "           ' 品ＷＦＯＳＦ４測定位置＿方
    sSQL = sSQL & "E029HWFOF4ST, "           ' 品ＷＦＯＳＦ４測定位置＿点
    sSQL = sSQL & "E029HWFOF4SR, "           ' 品ＷＦＯＳＦ４測定位置＿領
    sSQL = sSQL & "E029HWFOF4HT, "           ' 品ＷＦＯＳＦ４保証方法＿対
    sSQL = sSQL & "E029HWFOF4HS, "           ' 品ＷＦＯＳＦ４保証方法＿処
    sSQL = sSQL & "E029HWFBM1AN, "           ' 品ＷＦＢＭＤ１平均下限
    sSQL = sSQL & "E029HWFBM1AX, "           ' 品ＷＦＢＭＤ１平均上限
    sSQL = sSQL & "E029HWFBM1SH, "           ' 品ＷＦＢＭＤ１測定位置＿方
    sSQL = sSQL & "E029HWFBM1ST, "           ' 品ＷＦＢＭＤ１測定位置＿点
    sSQL = sSQL & "E029HWFBM1SR, "           ' 品ＷＦＢＭＤ１測定位置＿領
    sSQL = sSQL & "E029HWFBM1HT, "           ' 品ＷＦＢＭＤ１保証方法＿対
    sSQL = sSQL & "E029HWFBM1HS, "           ' 品ＷＦＢＭＤ１保証方法＿処
    sSQL = sSQL & "E029HWFBM2AN, "           ' 品ＷＦＢＭＤ２平均下限
    sSQL = sSQL & "E029HWFBM2AX, "           ' 品ＷＦＢＭＤ２平均上限
    sSQL = sSQL & "E029HWFBM2SH, "           ' 品ＷＦＢＭＤ２測定位置＿方
    sSQL = sSQL & "E029HWFBM2ST, "           ' 品ＷＦＢＭＤ２測定位置＿点
    sSQL = sSQL & "E029HWFBM2SR, "           ' 品ＷＦＢＭＤ２測定位置＿領
    sSQL = sSQL & "E029HWFBM2HT, "           ' 品ＷＦＢＭＤ２保証方法＿対
    sSQL = sSQL & "E029HWFBM2HS, "           ' 品ＷＦＢＭＤ２保証方法＿処
    sSQL = sSQL & "E029HWFBM3AN, "           ' 品ＷＦＢＭＤ３平均下限
    sSQL = sSQL & "E029HWFBM3AX, "           ' 品ＷＦＢＭＤ３平均上限
    sSQL = sSQL & "E029HWFBM3SH, "           ' 品ＷＦＢＭＤ３測定位置＿方
    sSQL = sSQL & "E029HWFBM3ST, "           ' 品ＷＦＢＭＤ３測定位置＿点
    sSQL = sSQL & "E029HWFBM3SR, "           ' 品ＷＦＢＭＤ３測定位置＿領
    sSQL = sSQL & "E029HWFBM3HT, "           ' 品ＷＦＢＭＤ３保証方法＿対
    sSQL = sSQL & "E029HWFBM3HS, "           ' 品ＷＦＢＭＤ３保証方法＿処
    sSQL = sSQL & "E029HWFOSF1PTK, "         ' 品ＷＦＯＳＦ１パタン区分
    sSQL = sSQL & "E029HWFOSF2PTK, "         ' 品ＷＦＯＳＦ２パタン区分
    sSQL = sSQL & "E029HWFOSF3PTK, "         ' 品ＷＦＯＳＦ３パタン区分
    sSQL = sSQL & "E029HWFOSF4PTK, "         ' 品ＷＦＯＳＦ４パタン区分
    sSQL = sSQL & "E029HWFBM1MBP, "          ' 品ＷＦＢＭＤ１面内分布
    sSQL = sSQL & "E029HWFBM2MBP, "          ' 品ＷＦＢＭＤ２面内分布
    sSQL = sSQL & "E029HWFBM3MBP, "          ' 品ＷＦＢＭＤ３面内分布
    sSQL = sSQL & "E029HWFBM1MCL, "          ' 品ＷＦＢＭＤ１面内計算
    sSQL = sSQL & "E029HWFBM2MCL, "          ' 品ＷＦＢＭＤ２面内計算
    sSQL = sSQL & "E029HWFBM3MCL, "          ' 品ＷＦＢＭＤ３面内計算
    sSQL = sSQL & "E025HWFOS1NS, "           ' 品ＷＦ酸素析出１熱処理法
    sSQL = sSQL & "E025HWFOS2NS, "           ' 品ＷＦ酸素析出２熱処理法
    sSQL = sSQL & "E025HWFOS3NS, "           ' 品ＷＦ酸素析出３熱処理法
    
    sSQL = sSQL & "E029HWFOF1NS, "           ' 品ＷＦＯＳＦ１熱処理法
    sSQL = sSQL & "E029HWFOF2NS, "           ' 品ＷＦＯＳＦ２熱処理法
    sSQL = sSQL & "E029HWFOF3NS, "           ' 品ＷＦＯＳＦ３熱処理法
    sSQL = sSQL & "E029HWFOF4NS, "           ' 品ＷＦＯＳＦ４熱処理法
    
    sSQL = sSQL & "E029HWFBM1NS, "           ' 品ＷＦＢＭＤ１熱処理法
    sSQL = sSQL & "E029HWFBM2NS, "           ' 品ＷＦＢＭＤ２熱処理法
    sSQL = sSQL & "E029HWFBM3NS, "           ' 品ＷＦＢＭＤ３熱処理法

    sSQL = sSQL & "E025HWFANTIM, "           ' 品ＷＦＡＮ時間
    sSQL = sSQL & "E025HWFANTNP, "           ' 品ＷＦＡＮ温度

    sSQL = sSQL & "E029HWFOF1ET, "           ' 品ＷＦＯＳＦ１選択ＥＴ代
    sSQL = sSQL & "E029HWFOF2ET, "           ' 品ＷＦＯＳＦ２選択ＥＴ代
    sSQL = sSQL & "E029HWFOF3ET, "           ' 品ＷＦＯＳＦ３選択ＥＴ代
    sSQL = sSQL & "E029HWFOF4ET, "           ' 品ＷＦＯＳＦ４選択ＥＴ代
    sSQL = sSQL & "E029HWFBM1ET, "           ' 品ＷＦＢＭＤ１選択ＥＴ代
    sSQL = sSQL & "E029HWFBM2ET, "           ' 品ＷＦＢＭＤ２選択ＥＴ代
    sSQL = sSQL & "E029HWFBM3ET, "           ' 品ＷＦＢＭＤ３選択ＥＴ代

    sSQL = sSQL & "E029HWFOF1SZ, "           ' 品ＷＦＯＳＦ１測定条件
    sSQL = sSQL & "E029HWFOF2SZ, "           ' 品ＷＦＯＳＦ２測定条件
    sSQL = sSQL & "E029HWFOF3SZ, "           ' 品ＷＦＯＳＦ３測定条件
    sSQL = sSQL & "E029HWFOF4SZ, "           ' 品ＷＦＯＳＦ４測定条件
    sSQL = sSQL & "E029HWFBM1SZ, "           ' 品ＷＦＢＭＤ１測定条件
    sSQL = sSQL & "E029HWFBM2SZ, "           ' 品ＷＦＢＭＤ２測定条件
    sSQL = sSQL & "E029HWFBM3SZ, "           ' 品ＷＦＢＭＤ３測定条件
    sSQL = sSQL & "E028HWFSPVAM "            ' 品ＷＦＳＰＶ平均上限
    
    '' SPV9点対応  整数部2桁→3桁変更対応
    sSQL = sSQL & ",E028HWFSPVMXN"
    sSQL = sSQL & ",E028HWFSPVAMN"
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    sSQL = sSQL & ",NVL(E36.HSXDKTMP, ' ') HSXDKTMP"
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    sSQL = sSQL & ",HSXLDLRMN"              ' 品SXL/DL連続0下限
    sSQL = sSQL & ",HSXLDLRMX"              ' 品SXL/DL連続0上限
    sSQL = sSQL & ",HWFLDLRMN"              ' 品WFL/DL連続0下限
    sSQL = sSQL & ",HWFLDLRMX"              ' 品WFL/DL連続0上限
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
    

'--------------- 2008/08/25 UPDATE START  By Systech ---------------
'    sSql = sSql & " from VECME001"
    sSQL = sSQL & " from VECME001, TBCME036 E36"
'--------------- 2008/08/25 UPDATE  END   By Systech ---------------
    
    If intSmpGetFlg = 0 Then
        sSQL = sSQL & " where E018HINBAN = '" & udtTypIn.HIN.hinban & "' and "
        sSQL = sSQL & " E018MNOREVNO = " & udtTypIn.HIN.mnorevno & " and "
        sSQL = sSQL & " E018FACTORY = '" & udtTypIn.HIN.factory & "' and "
        sSQL = sSQL & " E018OPECOND = '" & udtTypIn.HIN.opecond & "' "
    Else
        sSQL = sSQL & " where E018HINBAN = '" & udtNew_Hinban.hinban & "' and "
        sSQL = sSQL & " E018MNOREVNO = " & udtNew_Hinban.mnorevno & " and "
        sSQL = sSQL & " E018FACTORY = '" & udtNew_Hinban.factory & "' and "
        sSQL = sSQL & " E018OPECOND = '" & udtNew_Hinban.opecond & "' "
    End If

'--------------- 2008/08/25 INSERT START  By Systech ---------------
    sSQL = sSQL & " AND E36.HINBAN = E018HINBAN"
    sSQL = sSQL & " AND E36.MNOREVNO = E018MNOREVNO"
    sSQL = sSQL & " AND E36.FACTORY = E018FACTORY"
    sSQL = sSQL & " AND E36.OPECOND = E018OPECOND"
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    
    Debug.Print sSQL
    
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' レコード0件はエラー終了
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sErrMsg = GetMsgStr("EGET2", sDBName)
        funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    With udtSiyou
        .HWFTYPE = rs("E021HWFTYPE")                    ' 品ＷＦタイプ
        .HWFCDIR = rs("E022HWFCDIR")                    ' 品ＷＦ結晶面方
        .HWFCDOP = rs("E023HWFCDOP")                    ' 品ＷＦ結晶ドープ

        .HWFRMIN = fncNullCheck(rs("E021HWFRMIN"))      ' 品ＷＦ比抵抗下限
        .HWFRMAX = fncNullCheck(rs("E021HWFRMAX"))      ' 品ＷＦ比抵抗上限
        .HWFRSPOH = rs("E021HWFRSPOH")                  ' 品ＷＦ比抵抗測定位置＿方
        .HWFRSPOT = rs("E021HWFRSPOT")                  ' 品ＷＦ比抵抗測定位置＿点
        .HWFRSPOI = rs("E021HWFRSPOI")                  ' 品ＷＦ比抵抗測定位置＿位
        .HWFRHWYT = rs("E021HWFRHWYT")                  ' 品ＷＦ比抵抗保証方法＿対
        .HWFRHWYS = rs("E021HWFRHWYS")                  ' 品ＷＦ比抵抗保証方法＿処
        .HWFRMCAL = rs("E021HWFRMCAL")                  ' 品ＷＦ比抵抗面内計算
        .HWFRAMIN = fncNullCheck(rs("E021HWFRAMIN"))    ' 品ＷＦ比抵抗平均下限
        .HWFRAMAX = fncNullCheck(rs("E021HWFRAMAX"))    ' 品ＷＦ比抵抗平均上限
        .HWFRMBNP = fncNullCheck(rs("E021HWFRMBNP"))    ' 品ＷＦ比抵抗面内分布

        .HWFMKMIN = fncNullCheck(rs("E024HWFMKMIN"))    ' 品ＷＦ無欠陥層下限
        .HWFMKMAX = fncNullCheck(rs("E024HWFMKMAX"))    ' 品ＷＦ無欠陥層上限
        .HWFMKSPH = rs("E024HWFMKSPH")                  ' 品ＷＦ無欠陥層測定位置＿方
        .HWFMKSPT = rs("E024HWFMKSPT")                  ' 品ＷＦ無欠陥層測定位置＿点
        .HWFMKSPR = rs("E024HWFMKSPR")                  ' 品ＷＦ無欠陥層測定位置＿領
        .HWFMKHWT = rs("E024HWFMKHWT")                  ' 品ＷＦ無欠陥層保証方法＿対
        .HWFMKHWS = rs("E024HWFMKHWS")                  ' 品ＷＦ無欠陥層保証方法＿処

        .HWFONMIN = fncNullCheck(rs("E025HWFONMIN"))    ' 品ＷＦ酸素濃度下限
        .HWFONMAX = fncNullCheck(rs("E025HWFONMAX"))    ' 品ＷＦ酸素濃度上限
        .HWFONSPH = rs("E025HWFONSPH")                  ' 品ＷＦ酸素濃度測定位置＿方
        .HWFONSPT = rs("E025HWFONSPT")                  ' 品ＷＦ酸素濃度測定位置＿点
        .HWFONSPI = rs("E025HWFONSPI")                  ' 品ＷＦ酸素濃度測定位置＿位
        .HWFONHWT = rs("E025HWFONHWT")                  ' 品ＷＦ酸素濃度保証方法＿対
        .HWFONHWS = rs("E025HWFONHWS")                  ' 品ＷＦ酸素濃度保証方法＿処
        .HWFONMCL = rs("E025HWFONMCL")                  ' 品ＷＦ酸素濃度面内計算
        .HWFONMBP = fncNullCheck(rs("E025HWFONMBP"))    ' 品ＷＦ酸素濃度面内分布
        .HWFONAMN = fncNullCheck(rs("E025HWFONAMN"))    ' 品ＷＦ酸素濃度平均下限
        .HWFONAMX = fncNullCheck(rs("E025HWFONAMX"))    ' 品ＷＦ酸素濃度平均上限

        .HWFOS1MN = fncNullCheck(rs("E025HWFOS1MN"))    ' 品ＷＦ酸素析出１下限
        .HWFOS1MX = fncNullCheck(rs("E025HWFOS1MX"))    ' 品ＷＦ酸素析出１上限
        .HWFOS1SH = rs("E025HWFOS1SH")                  ' 品ＷＦ酸素析出１測定位置＿方
        .HWFOS1ST = rs("E025HWFOS1ST")                  ' 品ＷＦ酸素析出１測定位置＿点
        .HWFOS1SI = rs("E025HWFOS1SI")                  ' 品ＷＦ酸素析出１測定位置＿位
        .HWFOS1HT = rs("E025HWFOS1HT")                  ' 品ＷＦ酸素析出１保証方法＿対
        .HWFOS1HS = rs("E025HWFOS1HS")                  ' 品ＷＦ酸素析出１保証方法＿処
        .HWFOS2SH = rs("E025HWFOS2SH")                  ' 品ＷＦ酸素析出２測定位置＿方
        .HWFOS2ST = rs("E025HWFOS2ST")                  ' 品ＷＦ酸素析出２測定位置＿点
        .HWFOS2SI = rs("E025HWFOS2SI")                  ' 品ＷＦ酸素析出２測定位置＿位
        .HWFOS2MN = fncNullCheck(rs("E025HWFOS2MN"))    ' 品ＷＦ酸素析出２下限
        .HWFOS2MX = fncNullCheck(rs("E025HWFOS2MX"))    ' 品ＷＦ酸素析出２上限
        .HWFOS2HT = rs("E025HWFOS2HT")                  ' 品ＷＦ酸素析出２保証方法＿対
        .HWFOS2HS = rs("E025HWFOS2HS")                  ' 品ＷＦ酸素析出２保証方法＿処
        .HWFOS3MN = fncNullCheck(rs("E025HWFOS3MN"))    ' 品ＷＦ酸素析出３下限
        .HWFOS3MX = fncNullCheck(rs("E025HWFOS3MX"))    ' 品ＷＦ酸素析出３上限
        .HWFOS3SH = rs("E025HWFOS3SH")                  ' 品ＷＦ酸素析出３測定位置＿方
        .HWFOS3ST = rs("E025HWFOS3ST")                  ' 品ＷＦ酸素析出３測定位置＿点
        .HWFOS3SI = rs("E025HWFOS3SI")                  ' 品ＷＦ酸素析出３測定位置＿位
        .HWFOS3HT = rs("E025HWFOS3HT")                  ' 品ＷＦ酸素析出３保証方法＿対
        .HWFOS3HS = rs("E025HWFOS3HS")                  ' 品ＷＦ酸素析出３保証方法＿処

        .HWFDSOMX = fncNullCheck(rs("E026HWFDSOMX"))    ' 品ＷＦＤＳＯＤ上限
        .HWFDSOMN = fncNullCheck(rs("E026HWFDSOMN"))    ' 品ＷＦＤＳＯＤ下限
        .HWFDSOAX = fncNullCheck(rs("E026HWFDSOAX"))    ' 品ＷＦＤＳＯＤ領域上限
        .HWFDSOAN = fncNullCheck(rs("E026HWFDSOAN"))    ' 品ＷＦＤＳＯＤ領域下限
        .HWFDSOHT = rs("E026HWFDSOHT")                  ' 品ＷＦＤＳＯＤ保証方法＿対
        .HWFDSOHS = rs("E026HWFDSOHS")                  ' 品ＷＦＤＳＯＤ保証方法＿処
        
        ' SPV9点対応  整数部2桁→3桁変更対応
        .HWFSPVMX = fncNullCheck(rs("E028HWFSPVMXN"))   ' 品ＷＦＳＰＶＦＥ上限
        
        .HWFSPVSH = rs("E028HWFSPVSH")                  ' 品ＷＦＳＰＶＦＥ測定位置＿方
        .HWFSPVST = rs("E028HWFSPVST")                  ' 品ＷＦＳＰＶＦＥ測定位置＿点
        .HWFSPVSI = rs("E028HWFSPVSI")                  ' 品ＷＦＳＰＶＦＥ測定位置＿位
        .HWFSPVHT = rs("E028HWFSPVHT")                  ' 品ＷＦＳＰＶＦＥ保証方法＿対
        .HWFSPVHS = rs("E028HWFSPVHS")                  ' 品ＷＦＳＰＶＦＥ保証方法＿処
        .HWFDLSPH = rs("E028HWFDLSPH")                  ' 品ＷＦ拡散長測定位置＿方
        .HWFDLSPT = rs("E028HWFDLSPT")                  ' 品ＷＦ拡散長測定位置＿点
        .HWFDLSPI = rs("E028HWFDLSPI")                  ' 品ＷＦ拡散長測定位置＿位
        .HWFDLHWT = rs("E028HWFDLHWT")                  ' 品ＷＦ拡散長保証方法＿対
        .HWFDLHWS = rs("E028HWFDLHWS")                  ' 品ＷＦ拡散長保証方法＿処
        .HWFDLMIN = fncNullCheck(rs("E028HWFDLMIN"))    ' 品ＷＦ拡散長下限
        .HWFDLMAX = fncNullCheck(rs("E028HWFDLMAX"))    ' 品ＷＦ拡散長上限
                    
        .HWFOF1AX = fncNullCheck(rs("E029HWFOF1AX"))    ' 品ＷＦＯＳＦ１平均上限
        .HWFOF1MX = fncNullCheck(rs("E029HWFOF1MX"))    ' 品ＷＦＯＳＦ１上限
        .HWFOF1SH = rs("E029HWFOF1SH")                  ' 品ＷＦＯＳＦ１測定位置＿方
        .HWFOF1ST = rs("E029HWFOF1ST")                  ' 品ＷＦＯＳＦ１測定位置＿点
        .HWFOF1SR = rs("E029HWFOF1SR")                  ' 品ＷＦＯＳＦ１測定位置＿領
        .HWFOF1HT = rs("E029HWFOF1HT")                  ' 品ＷＦＯＳＦ１保証方法＿対
        .HWFOF1HS = rs("E029HWFOF1HS")                  ' 品ＷＦＯＳＦ１保証方法＿処
        .HWFOF2AX = fncNullCheck(rs("E029HWFOF2AX"))    ' 品ＷＦＯＳＦ２平均上限
        .HWFOF2MX = fncNullCheck(rs("E029HWFOF2MX"))    ' 品ＷＦＯＳＦ２上限
        .HWFOF2SH = rs("E029HWFOF2SH")                  ' 品ＷＦＯＳＦ２測定位置＿方
        .HWFOF2ST = rs("E029HWFOF2ST")                  ' 品ＷＦＯＳＦ２測定位置＿点
        .HWFOF2SR = rs("E029HWFOF2SR")                  ' 品ＷＦＯＳＦ２測定位置＿領
        .HWFOF2HT = rs("E029HWFOF2HT")                  ' 品ＷＦＯＳＦ２保証方法＿対
        .HWFOF2HS = rs("E029HWFOF2HS")                  ' 品ＷＦＯＳＦ２保証方法＿処
        .HWFOF3AX = fncNullCheck(rs("E029HWFOF3AX"))    ' 品ＷＦＯＳＦ３平均上限
        .HWFOF3MX = fncNullCheck(rs("E029HWFOF3MX"))    ' 品ＷＦＯＳＦ３上限
        .HWFOF3SH = rs("E029HWFOF3SH")                  ' 品ＷＦＯＳＦ３測定位置＿方
        .HWFOF3ST = rs("E029HWFOF3ST")                  ' 品ＷＦＯＳＦ３測定位置＿点
        .HWFOF3SR = rs("E029HWFOF3SR")                  ' 品ＷＦＯＳＦ３測定位置＿領
        .HWFOF3HT = rs("E029HWFOF3HT")                  ' 品ＷＦＯＳＦ３保証方法＿対
        .HWFOF3HS = rs("E029HWFOF3HS")                  ' 品ＷＦＯＳＦ３保証方法＿処
        .HWFOF4AX = fncNullCheck(rs("E029HWFOF4AX"))    ' 品ＷＦＯＳＦ４平均上限
        .HWFOF4MX = fncNullCheck(rs("E029HWFOF4MX"))    ' 品ＷＦＯＳＦ４上限
        .HWFOF4SH = rs("E029HWFOF4SH")                  ' 品ＷＦＯＳＦ４測定位置＿方
        .HWFOF4ST = rs("E029HWFOF4ST")                  ' 品ＷＦＯＳＦ４測定位置＿点
        .HWFOF4SR = rs("E029HWFOF4SR")                  ' 品ＷＦＯＳＦ４測定位置＿領
        .HWFOF4HT = rs("E029HWFOF4HT")                  ' 品ＷＦＯＳＦ４保証方法＿対
        .HWFOF4HS = rs("E029HWFOF4HS")                  ' 品ＷＦＯＳＦ４保証方法＿処
        If IsNull(rs("E029HWFOSF1PTK")) = False Then .HWFOSF1PTK = rs("E029HWFOSF1PTK")       ' 品ＷＦＯＳＦ１パタン区分　▼2003/05/14 ooba
        If IsNull(rs("E029HWFOSF2PTK")) = False Then .HWFOSF2PTK = rs("E029HWFOSF2PTK")       ' 品ＷＦＯＳＦ２パタン区分
        If IsNull(rs("E029HWFOSF3PTK")) = False Then .HWFOSF3PTK = rs("E029HWFOSF3PTK")       ' 品ＷＦＯＳＦ３パタン区分
        If IsNull(rs("E029HWFOSF4PTK")) = False Then .HWFOSF4PTK = rs("E029HWFOSF4PTK")       ' 品ＷＦＯＳＦ４パタン区分　▲2003/05/14 ooba

        ' BMDべき乗数変更対応
        .HWFBM1AN = fncNullCheck(rs("E029HWFBM1AN"))    ' 品ＷＦＢＭＤ１平均下限
        .HWFBM1AX = fncNullCheck(rs("E029HWFBM1AX"))    ' 品ＷＦＢＭＤ１平均上限
        .HWFBM1SH = rs("E029HWFBM1SH")                  ' 品ＷＦＢＭＤ１測定位置＿方
        .HWFBM1ST = rs("E029HWFBM1ST")                  ' 品ＷＦＢＭＤ１測定位置＿点
        .HWFBM1SR = rs("E029HWFBM1SR")                  ' 品ＷＦＢＭＤ１測定位置＿領
        .HWFBM1HT = rs("E029HWFBM1HT")                  ' 品ＷＦＢＭＤ１保証方法＿対
        .HWFBM1HS = rs("E029HWFBM1HS")                  ' 品ＷＦＢＭＤ１保証方法＿処
        
        'BMDべき乗数変更対応
        .HWFBM2AN = fncNullCheck(rs("E029HWFBM2AN"))    ' 品ＷＦＢＭＤ２平均下限
        .HWFBM2AX = fncNullCheck(rs("E029HWFBM2AX"))    ' 品ＷＦＢＭＤ２平均上限
        .HWFBM2SH = rs("E029HWFBM2SH")                  ' 品ＷＦＢＭＤ２測定位置＿方
        .HWFBM2ST = rs("E029HWFBM2ST")                  ' 品ＷＦＢＭＤ２測定位置＿点
        .HWFBM2SR = rs("E029HWFBM2SR")                  ' 品ＷＦＢＭＤ２測定位置＿領
        .HWFBM2HT = rs("E029HWFBM2HT")                  ' 品ＷＦＢＭＤ２保証方法＿対
        .HWFBM2HS = rs("E029HWFBM2HS")                  ' 品ＷＦＢＭＤ２保証方法＿処

        ' BMDべき乗数変更対応
        .HWFBM3AN = fncNullCheck(rs("E029HWFBM3AN"))    ' 品ＷＦＢＭＤ３平均下限
        .HWFBM3AX = fncNullCheck(rs("E029HWFBM3AX"))    ' 品ＷＦＢＭＤ３平均上限
        .HWFBM3SH = rs("E029HWFBM3SH")                  ' 品ＷＦＢＭＤ３測定位置＿方
        .HWFBM3ST = rs("E029HWFBM3ST")                  ' 品ＷＦＢＭＤ３測定位置＿点
        .HWFBM3SR = rs("E029HWFBM3SR")                  ' 品ＷＦＢＭＤ３測定位置＿領
        .HWFBM3HT = rs("E029HWFBM3HT")                  ' 品ＷＦＢＭＤ３保証方法＿対
        .HWFBM3HS = rs("E029HWFBM3HS")                  ' 品ＷＦＢＭＤ３保証方法＿処
        
        .HWFBM1MBP = fncNullCheck(rs("E029HWFBM1MBP"))  ' 品ＷＦＢＭＤ１面内分布
        .HWFBM2MBP = fncNullCheck(rs("E029HWFBM2MBP"))  ' 品ＷＦＢＭＤ２面内分布
        .HWFBM3MBP = fncNullCheck(rs("E029HWFBM3MBP"))  ' 品ＷＦＢＭＤ３面内分布
        If IsNull(rs("E029HWFBM1MCL")) = False Then .HWFBM1MCL = rs("E029HWFBM1MCL")         ' 品ＷＦＢＭＤ１面内計算
        If IsNull(rs("E029HWFBM2MCL")) = False Then .HWFBM2MCL = rs("E029HWFBM2MCL")         ' 品ＷＦＢＭＤ２面内計算
        If IsNull(rs("E029HWFBM3MCL")) = False Then .HWFBM3MCL = rs("E029HWFBM3MCL")         ' 品ＷＦＢＭＤ３面内計算　▲2003/05/14 ooba

        .HWFOS1NS = rs("E025HWFOS1NS")                  ' 品ＷＦ酸素析出１熱処理法
        .HWFOS2NS = rs("E025HWFOS2NS")                  ' 品ＷＦ酸素析出２熱処理法
        .HWFOS3NS = rs("E025HWFOS3NS")                  ' 品ＷＦ酸素析出３熱処理法
        .HWFOF1NS = rs("E029HWFOF1NS")                  ' 品ＷＦＯＳＦ１熱処理法
        .HWFOF2NS = rs("E029HWFOF2NS")                  ' 品ＷＦＯＳＦ２熱処理法
        .HWFOF3NS = rs("E029HWFOF3NS")                  ' 品ＷＦＯＳＦ３熱処理法
        .HWFOF4NS = rs("E029HWFOF4NS")                  ' 品ＷＦＯＳＦ４熱処理法
        .HWFBM1NS = rs("E029HWFBM1NS")                  ' 品ＷＦＢＭＤ１熱処理法
        .HWFBM2NS = rs("E029HWFBM2NS")                  ' 品ＷＦＢＭＤ２熱処理法
        .HWFBM3NS = rs("E029HWFBM3NS")                  ' 品ＷＦＢＭＤ３熱処理法

        .HWFANTIM = fncNullCheck(rs("E025HWFANTIM"))    ' 品ＷＦＡＮ時間
        .HWFANTNP = fncNullCheck(rs("E025HWFANTNP"))    ' 品ＷＦＡＮ温度

        .HWFOF1ET = fncNullCheck(rs("E029HWFOF1ET"))    ' 品ＷＦＯＳＦ１選択ＥＴ代
        .HWFOF2ET = fncNullCheck(rs("E029HWFOF2ET"))    ' 品ＷＦＯＳＦ２選択ＥＴ代
        .HWFOF3ET = fncNullCheck(rs("E029HWFOF3ET"))    ' 品ＷＦＯＳＦ３選択ＥＴ代
        .HWFOF4ET = fncNullCheck(rs("E029HWFOF4ET"))    ' 品ＷＦＯＳＦ４選択ＥＴ代
        .HWFBM1ET = fncNullCheck(rs("E029HWFBM1ET"))    ' 品ＷＦＢＭＤ１選択ＥＴ代
        .HWFBM2ET = fncNullCheck(rs("E029HWFBM2ET"))    ' 品ＷＦＢＭＤ２選択ＥＴ代
        .HWFBM3ET = fncNullCheck(rs("E029HWFBM3ET"))    ' 品ＷＦＢＭＤ３選択ＥＴ代

        .HWFOF1SZ = rs("E029HWFOF1SZ")                  ' 品ＷＦＯＳＦ１測定条件
        .HWFOF2SZ = rs("E029HWFOF2SZ")                  ' 品ＷＦＯＳＦ２測定条件
        .HWFOF3SZ = rs("E029HWFOF3SZ")                  ' 品ＷＦＯＳＦ３測定条件
        .HWFOF4SZ = rs("E029HWFOF4SZ")                  ' 品ＷＦＯＳＦ４測定条件
        .HWFBM1SZ = rs("E029HWFBM1SZ")                  ' 品ＷＦＢＭＤ１測定条件
        .HWFBM2SZ = rs("E029HWFBM2SZ")                  ' 品ＷＦＢＭＤ２測定条件
        .HWFBM3SZ = rs("E029HWFBM3SZ")                  ' 品ＷＦＢＭＤ３測定条件
    
        ' SPV9点対応  整数部2桁→3桁変更対応
        .HWFSPVAM = fncNullCheck(rs("E028HWFSPVAMN"))   ' 品ＷＦＳＰＶＦＥ平均上限

'--------------- 2008/08/25 INSERT START  By Systech ---------------
        .HSXDKTMP = rs("HSXDKTMP")                      ' DK温度（仕様）
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
        .HSXLDLRMN = fncNullCheck(rs("HSXLDLRMN"))      ' 品SXL/DL連続0下限
        .HSXLDLRMX = fncNullCheck(rs("HSXLDLRMX"))      ' 品SXL/DL連続0上限
        .HWFLDLRMN = fncNullCheck(rs("HWFLDLRMN"))      ' 品WFL/DL連続0下限
        .HWFLDLRMX = fncNullCheck(rs("HWFLDLRMX"))      ' 品WFL/DL連続0上限
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
        
        End With
    Set rs = Nothing
    
    ' 検査頻度_抜ﾃﾞｰﾀ取得
    sSQL = "select "
    sSQL = sSQL & "TBCME026.HWFDSOPTK, "                ' 品WFDSODパタン区分
    sSQL = sSQL & "TBCME021.HWFRKHNN, "                 ' 品ＷＦ比抵抗検査頻度＿抜
    sSQL = sSQL & "TBCME025.HWFONKHN, "                 ' 品ＷＦ酸素濃度検査頻度＿抜
    sSQL = sSQL & "TBCME029.HWFOF1KN, "                 ' 品ＷＦＯＳＦ１検査頻度＿抜
    sSQL = sSQL & "TBCME029.HWFOF2KN, "                 ' 品ＷＦＯＳＦ２検査頻度＿抜
    sSQL = sSQL & "TBCME029.HWFOF3KN, "                 ' 品ＷＦＯＳＦ３検査頻度＿抜
    sSQL = sSQL & "TBCME029.HWFOF4KN, "                 ' 品ＷＦＯＳＦ４検査頻度＿抜
    sSQL = sSQL & "TBCME029.HWFBM1KN, "                 ' 品ＷＦＢＭＤ１検査頻度＿抜
    sSQL = sSQL & "TBCME029.HWFBM2KN, "                 ' 品ＷＦＢＭＤ２検査頻度＿抜
    sSQL = sSQL & "TBCME029.HWFBM3KN, "                 ' 品ＷＦＢＭＤ３検査頻度＿抜
    sSQL = sSQL & "TBCME025.HWFOS1KN, "                 ' 品ＷＦ酸素析出１検査頻度＿抜
    sSQL = sSQL & "TBCME025.HWFOS2KN, "                 ' 品ＷＦ酸素析出２検査頻度＿抜
    sSQL = sSQL & "TBCME025.HWFOS3KN, "                 ' 品ＷＦ酸素析出３検査頻度＿抜
    sSQL = sSQL & "TBCME026.HWFDSOKN, "                 ' 品ＷＦＤＳＯＤ検査頻度＿抜
    sSQL = sSQL & "TBCME024.HWFMKKHN, "                 ' 品ＷＦ無欠陥層検査頻度＿抜
    sSQL = sSQL & "TBCME028.HWFSPVKN, "                 ' 品ＷＦＳＰＶＦＥ検査頻度＿抜
    sSQL = sSQL & "TBCME028.HWFDLKHN, "                 ' 品ＷＦ拡散長検査頻度＿抜
    sSQL = sSQL & "TBCME025.HWFZOKHN, "                 ' 品ＷＦ残存酸素検査頻度＿抜
    sSQL = sSQL & "TBCME026.HWFGDKHN "                  ' 品ＷＦＧＤ検査頻度＿抜
    sSQL = sSQL & "from TBCME021, TBCME024, TBCME025, TBCME026, TBCME028, TBCME029 "
    sSQL = sSQL & "where TBCME021.HINBAN = TBCME024.HINBAN "
    sSQL = sSQL & "and TBCME021.MNOREVNO = TBCME024.MNOREVNO "
    sSQL = sSQL & "and TBCME021.FACTORY = TBCME024.FACTORY "
    sSQL = sSQL & "and TBCME021.OPECOND = TBCME024.OPECOND "
    sSQL = sSQL & "and TBCME021.HINBAN = TBCME025.HINBAN "
    sSQL = sSQL & "and TBCME021.MNOREVNO = TBCME025.MNOREVNO "
    sSQL = sSQL & "and TBCME021.FACTORY = TBCME025.FACTORY "
    sSQL = sSQL & "and TBCME021.OPECOND = TBCME025.OPECOND "
    sSQL = sSQL & "and TBCME021.HINBAN = TBCME026.HINBAN "
    sSQL = sSQL & "and TBCME021.MNOREVNO = TBCME026.MNOREVNO "
    sSQL = sSQL & "and TBCME021.FACTORY = TBCME026.FACTORY "
    sSQL = sSQL & "and TBCME021.OPECOND = TBCME026.OPECOND "
    sSQL = sSQL & "and TBCME021.HINBAN = TBCME028.HINBAN "
    sSQL = sSQL & "and TBCME021.MNOREVNO = TBCME028.MNOREVNO "
    sSQL = sSQL & "and TBCME021.FACTORY = TBCME028.FACTORY "
    sSQL = sSQL & "and TBCME021.OPECOND = TBCME028.OPECOND "
    sSQL = sSQL & "and TBCME021.HINBAN = TBCME029.HINBAN "
    sSQL = sSQL & "and TBCME021.MNOREVNO = TBCME029.MNOREVNO "
    sSQL = sSQL & "and TBCME021.FACTORY = TBCME029.FACTORY "
    sSQL = sSQL & "and TBCME021.OPECOND = TBCME029.OPECOND "
    
    If intSmpGetFlg = 0 Then
        sSQL = sSQL & "and TBCME021.HINBAN = '" & udtTypIn.HIN.hinban & "' "
        sSQL = sSQL & "and TBCME021.MNOREVNO = " & udtTypIn.HIN.mnorevno & " "
        sSQL = sSQL & "and TBCME021.FACTORY = '" & udtTypIn.HIN.factory & "' "
        sSQL = sSQL & "and TBCME021.OPECOND = '" & udtTypIn.HIN.opecond & "' "
    Else
        sSQL = sSQL & "and TBCME021.HINBAN = '" & udtNew_Hinban.hinban & "' "
        sSQL = sSQL & "and TBCME021.MNOREVNO = " & udtNew_Hinban.mnorevno & " "
        sSQL = sSQL & "and TBCME021.FACTORY = '" & udtNew_Hinban.factory & "' "
        sSQL = sSQL & "and TBCME021.OPECOND = '" & udtNew_Hinban.opecond & "' "
    End If

    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' レコード0件はエラー終了
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sErrMsg = GetMsgStr("EGET2", sDBName)
        funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    With udtSiyou
        If IsNull(rs("HWFDSOPTK")) = False Then .HWFDSOPTK = rs("HWFDSOPTK") Else .HWFDSOPTK = " "  ' 品WFDSODパタン区分
        If IsNull(rs("HWFRKHNN")) = False Then .HWFRKHNN = rs("HWFRKHNN") Else .HWFRKHNN = " "      ' 品ＷＦ比抵抗検査頻度＿抜
        If IsNull(rs("HWFONKHN")) = False Then .HWFONKHN = rs("HWFONKHN") Else .HWFONKHN = " "      ' 品ＷＦ酸素濃度検査頻度＿抜
        If IsNull(rs("HWFOF1KN")) = False Then .HWFOF1KN = rs("HWFOF1KN") Else .HWFOF1KN = " "      ' 品ＷＦＯＳＦ１検査頻度＿抜
        If IsNull(rs("HWFOF2KN")) = False Then .HWFOF2KN = rs("HWFOF2KN") Else .HWFOF2KN = " "      ' 品ＷＦＯＳＦ２検査頻度＿抜
        If IsNull(rs("HWFOF3KN")) = False Then .HWFOF3KN = rs("HWFOF3KN") Else .HWFOF3KN = " "      ' 品ＷＦＯＳＦ３検査頻度＿抜
        If IsNull(rs("HWFOF4KN")) = False Then .HWFOF4KN = rs("HWFOF4KN") Else .HWFOF4KN = " "      ' 品ＷＦＯＳＦ４検査頻度＿抜
        If IsNull(rs("HWFBM1KN")) = False Then .HWFBM1KN = rs("HWFBM1KN") Else .HWFBM1KN = " "      ' 品ＷＦＢＭＤ１検査頻度＿抜
        If IsNull(rs("HWFBM2KN")) = False Then .HWFBM2KN = rs("HWFBM2KN") Else .HWFBM2KN = " "      ' 品ＷＦＢＭＤ２検査頻度＿抜
        If IsNull(rs("HWFBM3KN")) = False Then .HWFBM3KN = rs("HWFBM3KN") Else .HWFBM3KN = " "      ' 品ＷＦＢＭＤ３検査頻度＿抜
        If IsNull(rs("HWFOS1KN")) = False Then .HWFOS1KN = rs("HWFOS1KN") Else .HWFOS1KN = " "      ' 品ＷＦ酸素析出１検査頻度＿抜
        If IsNull(rs("HWFOS2KN")) = False Then .HWFOS2KN = rs("HWFOS2KN") Else .HWFOS2KN = " "      ' 品ＷＦ酸素析出２検査頻度＿抜
        If IsNull(rs("HWFOS3KN")) = False Then .HWFOS3KN = rs("HWFOS3KN") Else .HWFOS3KN = " "      ' 品ＷＦ酸素析出３検査頻度＿抜
        If IsNull(rs("HWFDSOKN")) = False Then .HWFDSOKN = rs("HWFDSOKN") Else .HWFDSOKN = " "      ' 品ＷＦＤＳＯＤ検査頻度＿抜
        If IsNull(rs("HWFMKKHN")) = False Then .HWFMKKHN = rs("HWFMKKHN") Else .HWFMKKHN = " "      ' 品ＷＦ無欠陥層検査頻度＿抜
        If IsNull(rs("HWFSPVKN")) = False Then .HWFSPVKN = rs("HWFSPVKN") Else .HWFSPVKN = " "      ' 品ＷＦＳＰＶＦＥ検査頻度＿抜
        If IsNull(rs("HWFDLKHN")) = False Then .HWFDLKHN = rs("HWFDLKHN") Else .HWFDLKHN = " "      ' 品ＷＦ拡散長検査頻度＿抜
        If IsNull(rs("HWFZOKHN")) = False Then .HWFZOKHN = rs("HWFZOKHN") Else .HWFZOKHN = " "      ' 品ＷＦ残存酸素検査頻度＿抜
        If IsNull(rs("HWFGDKHN")) = False Then .HWFGDKHN = rs("HWFGDKHN") Else .HWFGDKHN = " "      ' 品ＷＦＧＤ検査頻度＿抜
    End With
    Set rs = Nothing
    
    ''残存酸素仕様取得
    sDBName = "E025"
    sSQL = "select HWFZOMIN, HWFZOMAX, HWFZOSPH, HWFZOSPT, HWFZOSPI, HWFZOHWT, "
    sSQL = sSQL & "HWFZOHWS, HWFZONSW from TBCME025 "
    If intSmpGetFlg = 0 Then
        sSQL = sSQL & " where HINBAN = '" & udtTypIn.HIN.hinban & "' and "
        sSQL = sSQL & " MNOREVNO = " & udtTypIn.HIN.mnorevno & " and "
        sSQL = sSQL & " FACTORY = '" & udtTypIn.HIN.factory & "' and "
        sSQL = sSQL & " OPECOND = '" & udtTypIn.HIN.opecond & "' "
    Else
        sSQL = sSQL & " where HINBAN = '" & udtNew_Hinban.hinban & "' and "
        sSQL = sSQL & " MNOREVNO = " & udtNew_Hinban.mnorevno & " and "
        sSQL = sSQL & " FACTORY = '" & udtNew_Hinban.factory & "' and "
        sSQL = sSQL & " OPECOND = '" & udtNew_Hinban.opecond & "' "
    End If
    
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' レコード0件はエラー終了
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sErrMsg = GetMsgStr("EGET2", sDBName)
        funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    udtSiyou.HWFZOMIN = fncNullCheck(rs("HWFZOMIN"))                                                        ' 品ＷＦ残存酸素下限
    udtSiyou.HWFZOMAX = fncNullCheck(rs("HWFZOMAX"))                                                        ' 品ＷＦ残存酸素上限
    If IsNull(rs("HWFZOSPH")) = False Then udtSiyou.HWFZOSPH = rs("HWFZOSPH") Else udtSiyou.HWFZOSPH = " "  ' 品ＷＦ残存酸素測定位置＿方
    If IsNull(rs("HWFZOSPT")) = False Then udtSiyou.HWFZOSPT = rs("HWFZOSPT") Else udtSiyou.HWFZOSPT = " "  ' 品ＷＦ残存酸素測定位置＿点
    If IsNull(rs("HWFZOSPI")) = False Then udtSiyou.HWFZOSPI = rs("HWFZOSPI") Else udtSiyou.HWFZOSPI = " "  ' 品ＷＦ残存酸素測定位置＿位
    If IsNull(rs("HWFZOHWT")) = False Then udtSiyou.HWFZOHWT = rs("HWFZOHWT") Else udtSiyou.HWFZOHWT = " "  ' 品ＷＦ残存酸素保証方法＿対
    If IsNull(rs("HWFZOHWS")) = False Then udtSiyou.HWFZOHWS = rs("HWFZOHWS") Else udtSiyou.HWFZOHWS = " "  ' 品ＷＦ残存酸素保証方法＿処
    If IsNull(rs("HWFZONSW")) = False Then udtSiyou.HWFZONSW = rs("HWFZONSW") Else udtSiyou.HWFZONSW = " "  ' 品ＷＦ残存酸素熱処理法
        
    Set rs = Nothing
    
    ' GD仕様取得
    sDBName = "E026"
    sSQL = "select "
    sSQL = sSQL & "HWFDENKU, "        ' 品ＷＦＤｅｎ検査有無
    sSQL = sSQL & "HWFDENMX, "        ' 品ＷＦＤｅｎ上限
    sSQL = sSQL & "HWFDENMN, "        ' 品ＷＦＤｅｎ下限
    sSQL = sSQL & "HWFDENHT, "        ' 品ＷＦＤｅｎ保証方法＿対
    sSQL = sSQL & "HWFDENHS, "        ' 品ＷＦＤｅｎ保証方法＿処
    sSQL = sSQL & "HWFDVDKU, "        ' 品ＷＦＤＶＤ２検査有無
    sSQL = sSQL & "HWFDVDMXN, "       ' 品ＷＦＤＶＤ２上限
    sSQL = sSQL & "HWFDVDMNN, "       ' 品ＷＦＤＶＤ２下限
    sSQL = sSQL & "HWFDVDHT, "        ' 品ＷＦＤＶＤ２保証方法＿対
    sSQL = sSQL & "HWFDVDHS, "        ' 品ＷＦＤＶＤ２保証方法＿処
    sSQL = sSQL & "HWFLDLKU, "        ' 品ＷＦＬ／ＤＬ検査有無
    sSQL = sSQL & "HWFLDLMX, "        ' 品ＷＦＬ／ＤＬ上限
    sSQL = sSQL & "HWFLDLMN, "        ' 品ＷＦＬ／ＤＬ下限
    sSQL = sSQL & "HWFLDLHT, "        ' 品ＷＦＬ／ＤＬ保証方法＿対
    sSQL = sSQL & "HWFLDLHS, "        ' 品ＷＦＬ／ＤＬ保証方法＿処
    sSQL = sSQL & "HWFGDSPH, "        ' 品ＷＦＧＤ測定位置＿方
    sSQL = sSQL & "HWFGDSPT, "        ' 品ＷＦＧＤ測定位置＿点
    sSQL = sSQL & "HWFGDSPR "         ' 品ＷＦＧＤ測定位置＿領
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    sSQL = sSQL & ",HWFGDPTK "        ' 品ＷＦＧＤパタン区分
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
    sSQL = sSQL & "from TBCME026 "
    
    If intSmpGetFlg = 0 Then
        sSQL = sSQL & " where HINBAN = '" & udtTypIn.HIN.hinban & "' and "
        sSQL = sSQL & " MNOREVNO = " & udtTypIn.HIN.mnorevno & " and "
        sSQL = sSQL & " FACTORY = '" & udtTypIn.HIN.factory & "' and "
        sSQL = sSQL & " OPECOND = '" & udtTypIn.HIN.opecond & "' "
    Else
        sSQL = sSQL & " where HINBAN = '" & udtNew_Hinban.hinban & "' and "
        sSQL = sSQL & " MNOREVNO = " & udtNew_Hinban.mnorevno & " and "
        sSQL = sSQL & " FACTORY = '" & udtNew_Hinban.factory & "' and "
        sSQL = sSQL & " OPECOND = '" & udtNew_Hinban.opecond & "' "
    End If
    
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' レコード0件はエラー終了
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sErrMsg = GetMsgStr("EGET2", sDBName)
        funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
        
    With udtSiyou
        .HWFDENKU = rs("HWFDENKU")                      ' 品ＷＦＤｅｎ検査有無
        .HWFDENMX = fncNullCheck(rs("HWFDENMX"))        ' 品ＷＦＤｅｎ上限
        .HWFDENMN = fncNullCheck(rs("HWFDENMN"))        ' 品ＷＦＤｅｎ下限
        .HWFDENHT = rs("HWFDENHT")                      ' 品ＷＦＤｅｎ保証方法＿対
        .HWFDENHS = rs("HWFDENHS")                      ' 品ＷＦＤｅｎ保証方法＿処
        .HWFDVDKU = rs("HWFDVDKU")                      ' 品ＷＦＤＶＤ２検査有無
        .HWFDVDMXN = fncNullCheck(rs("HWFDVDMXN"))      ' 品ＷＦＤＶＤ２上限
        .HWFDVDMNN = fncNullCheck(rs("HWFDVDMNN"))      ' 品ＷＦＤＶＤ２下限
        .HWFDVDHT = rs("HWFDVDHT")                      ' 品ＷＦＤＶＤ２保証方法＿対
        .HWFDVDHS = rs("HWFDVDHS")                      ' 品ＷＦＤＶＤ２保証方法＿処
        .HWFLDLKU = rs("HWFLDLKU")                      ' 品ＷＦＬ／ＤＬ検査有無
        .HWFLDLMX = fncNullCheck(rs("HWFLDLMX"))        ' 品ＷＦＬ／ＤＬ上限
        .HWFLDLMN = fncNullCheck(rs("HWFLDLMN"))        ' 品ＷＦＬ／ＤＬ下限
        .HWFLDLHT = rs("HWFLDLHT")                      ' 品ＷＦＬ／ＤＬ保証方法＿対
        .HWFLDLHS = rs("HWFLDLHS")                      ' 品ＷＦＬ／ＤＬ保証方法＿処
        .HWFGDSPH = rs("HWFGDSPH")                      ' 品ＷＦＧＤ測定位置＿方
        .HWFGDSPT = rs("HWFGDSPT")                      ' 品ＷＦＧＤ測定位置＿点
        .HWFGDSPR = rs("HWFGDSPR")                      ' 品ＷＦＧＤ測定位置＿領
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
        If IsNull(rs("HWFGDPTK")) = False Then .HWFGDPTK = rs("HWFGDPTK") Else .HWFGDPTK = " "      ' 品ＷＦＧＤパタン区分
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
    End With
    Set rs = Nothing

    ' 品ＷＦＧＤライン数の取得
    sDBName = "E036"
    
    sSQL = "select "
    sSQL = sSQL & "HWFGDLINE "        ' 品ＷＦＧＤライン数の取得
    sSQL = sSQL & "from TBCME036 "
    If intSmpGetFlg = 0 Then
        sSQL = sSQL & " where HINBAN = '" & udtTypIn.HIN.hinban & "' and "
        sSQL = sSQL & " MNOREVNO = " & udtTypIn.HIN.mnorevno & " and "
        sSQL = sSQL & " FACTORY = '" & udtTypIn.HIN.factory & "' and "
        sSQL = sSQL & " OPECOND = '" & udtTypIn.HIN.opecond & "' "
    Else
        sSQL = sSQL & " where HINBAN = '" & udtNew_Hinban.hinban & "' and "
        sSQL = sSQL & " MNOREVNO = " & udtNew_Hinban.mnorevno & " and "
        sSQL = sSQL & " FACTORY = '" & udtNew_Hinban.factory & "' and "
        sSQL = sSQL & " OPECOND = '" & udtNew_Hinban.opecond & "' "
    End If
    
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' レコード0件はエラー終了
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sErrMsg = GetMsgStr("EGET2", sDBName)
        funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
        
    With udtSiyou
        .HWFGDLINE = fncNullCheck(rs("HWFGDLINE"))  ' 品ＷＦＧＤライン数の取得
    End With
    
    Set rs = Nothing
    
    
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    ' 品ＳＸＧＤパタン区分の取得
    sDBName = "E020"
    
    sSQL = "select "
    sSQL = sSQL & "HSXGDPTK "        ' 品ＳＸＧＤパタン区分
    sSQL = sSQL & "from TBCME020 "
    If intSmpGetFlg = 0 Then
        sSQL = sSQL & " where HINBAN = '" & udtTypIn.HIN.hinban & "' and "
        sSQL = sSQL & " MNOREVNO = " & udtTypIn.HIN.mnorevno & " and "
        sSQL = sSQL & " FACTORY = '" & udtTypIn.HIN.factory & "' and "
        sSQL = sSQL & " OPECOND = '" & udtTypIn.HIN.opecond & "' "
    Else
        sSQL = sSQL & " where HINBAN = '" & udtNew_Hinban.hinban & "' and "
        sSQL = sSQL & " MNOREVNO = " & udtNew_Hinban.mnorevno & " and "
        sSQL = sSQL & " FACTORY = '" & udtNew_Hinban.factory & "' and "
        sSQL = sSQL & " OPECOND = '" & udtNew_Hinban.opecond & "' "
    End If
    
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' レコード0件はエラー終了
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sErrMsg = GetMsgStr("EGET2", sDBName)
        funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
        
    With udtSiyou
        If IsNull(rs("HSXGDPTK")) = False Then .HSXGDPTK = rs("HSXGDPTK") Else .HSXGDPTK = " "
    End With
    
    Set rs = Nothing
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
    
    ' SPVNr濃度仕様取得
    ' Add 2010/01/06 SIRD対応 Y.Hitomi
    sDBName = "E048"
    sSQL = "select "
    sSQL = sSQL & "HWFNRHS, "                       ' 品WFSPVNR保証方法_処
    sSQL = sSQL & "HWFNRKN, "                       ' 品WFSPVNR保証方法_抜
    ' ↓Add 2010/01/06 SIRD対応 Y.Hitomi
    sSQL = sSQL & "HWFSIRDMX, "                     ' 軸状転位上限
    sSQL = sSQL & "HWFSIRDSZ, "                     ' 軸状転位測定条件
    sSQL = sSQL & "HWFSIRDHT, "                     ' 軸状転位保証方法＿対
    sSQL = sSQL & "HWFSIRDHS, "                     ' 軸状転位保証方法＿処
    sSQL = sSQL & "HWFSIRDKM, "                     ' 軸状転位検査頻度＿枚
    sSQL = sSQL & "HWFSIRDKN, "                     ' 軸状転位検査頻度＿抜
    sSQL = sSQL & "HWFSIRDKH, "                     ' 軸状転位検査頻度＿保
    sSQL = sSQL & "HWFSIRDKU  "                     ' 軸状転位検査頻度＿ウ
    ' ↑Add 2010/01/06 SIRD対応 Y.Hitomi
    sSQL = sSQL & "from TBCME048 "
    sSQL = sSQL & "where HINBAN = '" & udtNew_Hinban.hinban & "' "
    sSQL = sSQL & "and MNOREVNO = " & udtNew_Hinban.mnorevno & " "
    sSQL = sSQL & "and FACTORY = '" & udtNew_Hinban.factory & "' "
    sSQL = sSQL & "and OPECOND = '" & udtNew_Hinban.opecond & "' "
    
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sErrMsg = GetMsgStr("EGET2", sDBName)
        funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    If IsNull(rs("HWFNRHS")) = False Then udtSiyou.HWFNRHS = rs("HWFNRHS") Else udtSiyou.HWFNRHS = " "
    If IsNull(rs("HWFNRKN")) = False Then udtSiyou.HWFNRKN = rs("HWFNRKN") Else udtSiyou.HWFNRKN = " "
    ' ↓Add 2010/01/06 SIRD対応 Y.Hitomi
    If IsNull(rs("HWFSIRDMX")) = False Then udtSiyou.HWFSIRDMX = rs("HWFSIRDMX") Else udtSiyou.HWFSIRDMX = fncNullCheck(rs("HWFSIRDMX"))
    If IsNull(rs("HWFSIRDSZ")) = False Then udtSiyou.HWFSIRDSZ = rs("HWFSIRDSZ") Else udtSiyou.HWFSIRDSZ = " "
    If IsNull(rs("HWFSIRDHT")) = False Then udtSiyou.HWFSIRDHT = rs("HWFSIRDHT") Else udtSiyou.HWFSIRDHT = " "
    If IsNull(rs("HWFSIRDHS")) = False Then udtSiyou.HWFSIRDHS = rs("HWFSIRDHS") Else udtSiyou.HWFSIRDHS = " "
    If IsNull(rs("HWFSIRDKM")) = False Then udtSiyou.HWFSIRDKM = rs("HWFSIRDKM") Else udtSiyou.HWFSIRDKM = " "
    If IsNull(rs("HWFSIRDKN")) = False Then udtSiyou.HWFSIRDKN = rs("HWFSIRDKN") Else udtSiyou.HWFSIRDKN = " "
    If IsNull(rs("HWFSIRDKH")) = False Then udtSiyou.HWFSIRDKH = rs("HWFSIRDKH") Else udtSiyou.HWFSIRDKH = " "
    If IsNull(rs("HWFSIRDKU")) = False Then udtSiyou.HWFSIRDKU = rs("HWFSIRDKU") Else udtSiyou.HWFSIRDKU = " "
    ' ↑Add 2010/01/06 SIRD対応 Y.Hitomi
    
    rs.Close
    
    ' 測定評価結果取得
    sDBName = "Y013"
    If funGetTBCMY013(udtTypIn, udtSokutei()) = FUNCTION_RETURN_FAILURE Then
        funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
        
    ' GD実績取得
    If udtTypIn.WFSMP.WFINDGDCW <> "0" Then
        ' ｻﾝﾌﾟﾙ位置ｾｯﾄ
        If udtTypIn.WFSMP.TBKBNCW = "T" Then intPos = SxlTop Else intPos = SxlTail
        
        ' 結晶GD実績取得
        If udtTypIn.WFSMP.WFHSGDCW = "1" Then
            sDBName = "J006"
            If funGetGDJisseki_J006(udtTypIn.WFSMP.XTALCW, udtTypIn.WFSMP.WFSMPLIDGDCW, _
                                        typ_J015_WFGDJudg(intPos)) = FUNCTION_RETURN_FAILURE Then
                funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        ' WF_GD実績取得
        Else
            sDBName = "J015"
            If funGetGDJisseki_J015(udtTypIn.WFSMP.XTALCW, udtTypIn.WFSMP.WFSMPLIDGDCW, _
                                        udtTypIn.WFSMP.WFHSGDCW, typ_J015_WFGDJudg(intPos)) _
                                                                = FUNCTION_RETURN_FAILURE Then
                funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        End If
    End If

    '↓SIRD実績取得 Add 2010/01/07 SIRD対応 Y.Hitomi
        ' ｻﾝﾌﾟﾙ位置ｾｯﾄ
    If udtTypIn.WFSMP.TBKBNCW = "T" Then
        intPos = SxlTop
    Else
        intPos = SxlTail
    End If

    ' WF_SIRD実績取得
    If udtTypIn.WFSMP.WFINDL4CW <> "0" Then
        sDBName = "J022"
        If funGetSDJisseki_J022(udtTypIn.WFSMP.XTALCW, udtTypIn.WFSMP.WFSMPLIDL4CW, _
                                   typ_J022_WFSDJudg(intPos)) = FUNCTION_RETURN_FAILURE Then
            funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    '↑SIRD実績取得 Add 2010/01/07 SIRD対応 Y.Hitomi

    ' SPV9点対応
    ' ｻﾝﾌﾟﾙ位置ｾｯﾄ
    If udtTypIn.WFSMP.TBKBNCW = "T" Then intPos = SxlTop Else intPos = SxlTail
    
    If udtTypIn.WFSMP.WFINDSPCW <> "0" Then
        sDBName = "J016"
        If funGetSPVJisseki_J016(udtTypIn.WFSMP.XTALCW, udtTypIn.WFSMP.WFSMPLIDSPCW, _
                                    typ_J016_WFSPVJudg(intPos), udtSiyou) = FUNCTION_RETURN_FAILURE Then
            funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

'    HWFRSPOT As String * 1          ' 品ＷＦ比抵抗測定位置＿点
'    HWFRSPOI As String * 1          ' 品ＷＦ比抵抗測定位置＿位
'    HWFONSPT As String * 1          ' 品ＷＦ酸素濃度測定位置＿点
'    HWFONSPI As String * 1          ' 品ＷＦ酸素濃度測定位置＿位
'の代わりに
' 製品仕様SXLﾃﾞｰﾀ１
'Public Type typ_TBCME018
'    HSXRSPOT As String * 1          ' 品ＳＸ比抵抗測定位置＿点
'    HSXRSPOI As String * 1          ' 品ＳＸ比抵抗測定位置＿位
' 製品仕様SXLﾃﾞｰﾀ２
'Public Type typ_TBCME019
'    HSXONSPT As String * 1          ' 品ＳＸ酸素濃度測定位置＿点
'    HSXONSPI As String * 1          ' 品ＳＸ酸素濃度測定位置＿位
'を使用する。
    sSQL = "select "
    sSQL = sSQL & " HSXRSPOT, "
    sSQL = sSQL & " HSXRSPOI, "
    sSQL = sSQL & " HSXONSPT, "
    sSQL = sSQL & " HSXONSPI "
    sSQL = sSQL & " from "
    sSQL = sSQL & " TBCME018 K01, TBCME019 K12 "
    
    If intSmpGetFlg = 0 Then
        sSQL = sSQL & " where K01.HINBAN='" & udtTypIn.HIN.hinban & "' and "
        sSQL = sSQL & " K01.MNOREVNO=" & udtTypIn.HIN.mnorevno & " and "
        sSQL = sSQL & " K01.FACTORY='" & udtTypIn.HIN.factory & "' and "
        sSQL = sSQL & " K01.OPECOND='" & udtTypIn.HIN.opecond & "' and "
        sSQL = sSQL & " K12.HINBAN='" & udtTypIn.HIN.hinban & "' and "
        sSQL = sSQL & " K12.MNOREVNO=" & udtTypIn.HIN.mnorevno & " and "
        sSQL = sSQL & " K12.FACTORY='" & udtTypIn.HIN.factory & "' and "
        sSQL = sSQL & " K12.OPECOND='" & udtTypIn.HIN.opecond & "'"
    Else
        sSQL = sSQL & " where K01.HINBAN='" & udtNew_Hinban.hinban & "' and "
        sSQL = sSQL & " K01.MNOREVNO=" & udtNew_Hinban.mnorevno & " and "
        sSQL = sSQL & " K01.FACTORY='" & udtNew_Hinban.factory & "' and "
        sSQL = sSQL & " K01.OPECOND='" & udtNew_Hinban.opecond & "' and "
        sSQL = sSQL & " K12.HINBAN='" & udtNew_Hinban.hinban & "' and "
        sSQL = sSQL & " K12.MNOREVNO=" & udtNew_Hinban.mnorevno & " and "
        sSQL = sSQL & " K12.FACTORY='" & udtNew_Hinban.factory & "' and "
        sSQL = sSQL & " K12.OPECOND='" & udtNew_Hinban.opecond & "'"
    End If
    
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    ' レコード0件はエラー終了
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sErrMsg = GetMsgStr("EGET2", sDBName)
        funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    With udtSiyou
        .HWFRSPOT = rs("HSXRSPOT")      ' 品ＳＸ比抵抗測定位置＿点
        .HWFRSPOI = rs("HSXRSPOI")      ' 品ＳＸ比抵抗測定位置＿位
        .HWFONSPT = rs("HSXONSPT")      ' 品ＳＸ酸素濃度測定位置＿点
        .HWFONSPI = rs("HSXONSPI")      ' 品ＳＸ酸素濃度測定位置＿位
    End With
    
    Set rs = Nothing
    
    '' エピ仕様取得(BMD,OSF)
    sSQL = "select "
    sSQL = sSQL & "HEPANTNP, "          ' 品EPAN温度
    sSQL = sSQL & "HEPBM1HS, "          ' 品EPBMD1保証方法＿処
    sSQL = sSQL & "HEPBM1AN, "          ' 品EPBMD1平均下限
    sSQL = sSQL & "HEPBM1AX, "          ' 品EPBMD1平均上限
    sSQL = sSQL & "HEPBM2HS, "          ' 品EPBMD1保証方法＿処
    sSQL = sSQL & "HEPBM2AN, "          ' 品EPBMD2平均下限
    sSQL = sSQL & "HEPBM2AX, "          ' 品EPBMD2平均上限
    sSQL = sSQL & "HEPBM3HS, "          ' 品EPBMD1保証方法＿処
    sSQL = sSQL & "HEPBM3AN, "          ' 品EPBMD3平均下限
    sSQL = sSQL & "HEPBM3AX, "          ' 品EPBMD3平均上限
    sSQL = sSQL & "HEPBM3GSAN, "        ' 品EPBMD3平均下限(外周)　09/05/07 ooba
    sSQL = sSQL & "HEPBM3GSAX, "        ' 品EPBMD3平均上限(外周)　09/05/07 ooba
    sSQL = sSQL & "HEPOF1HS, "          ' 品EPOSF1保証方法＿処
    sSQL = sSQL & "HEPOF1AX, "          ' 品EPOSF1平均上限
    sSQL = sSQL & "HEPOF1MX, "          ' 品EPOSF1上限
    sSQL = sSQL & "HEPOF2HS, "          ' 品EPOSF2保証方法＿処
    sSQL = sSQL & "HEPOF2AX, "          ' 品EPOSF2平均上限
    sSQL = sSQL & "HEPOF2MX, "          ' 品EPOSF2上限
    sSQL = sSQL & "HEPOF3HS, "          ' 品EPOSF3保証方法＿処
    sSQL = sSQL & "HEPOF3AX, "          ' 品EPOSF3平均上限
    sSQL = sSQL & "HEPOF3MX  "          ' 品EPOSF3上限
    sSQL = sSQL & "from TBCME050 "      ' 製品仕様エピデータ１
    
    If intSmpGetFlg = 0 Then
        sSQL = sSQL & " where HINBAN = '" & udtTypIn.HIN.hinban & "' and "
        sSQL = sSQL & " MNOREVNO = " & udtTypIn.HIN.mnorevno & " and "
        sSQL = sSQL & " FACTORY = '" & udtTypIn.HIN.factory & "' and "
        sSQL = sSQL & " OPECOND = '" & udtTypIn.HIN.opecond & "' "
    Else
        sSQL = sSQL & " where HINBAN = '" & udtNew_Hinban.hinban & "' and "
        sSQL = sSQL & " MNOREVNO = " & udtNew_Hinban.mnorevno & " and "
        sSQL = sSQL & " FACTORY = '" & udtNew_Hinban.factory & "' and "
        sSQL = sSQL & " OPECOND = '" & udtNew_Hinban.opecond & "' "
    End If
    
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' データ無しの場合は終了
    If rs.RecordCount = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDBName)
        funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    ' 品EPBMD1保証方法＿処が"H","S"のものが一つでもあれば、エピ仕様あり
    If rs("HEPBM1HS") = "H" Or rs("HEPBM2HS") = "H" Or rs("HEPBM3HS") = "H" Or _
       rs("HEPOF1HS") = "H" Or rs("HEPOF2HS") = "H" Or rs("HEPOF3HS") = "H" Or _
       rs("HEPBM1HS") = "S" Or rs("HEPBM2HS") = "S" Or rs("HEPBM3HS") = "S" Or _
       rs("HEPOF1HS") = "S" Or rs("HEPOF2HS") = "S" Or rs("HEPOF3HS") = "S" Then
        udtSiyou.HEPHS = True
    Else
        udtSiyou.HEPHS = False
    End If
    
    If udtSiyou.HEPHS = True Then
        With udtSiyou
            .HEPBM1AN = fncNullCheck(rs("HEPBM1AN"))   ' 品EPBMD1平均下限
            .HEPBM1AX = fncNullCheck(rs("HEPBM1AX"))   ' 品EPBMD1平均上限
            .HEPBM2AN = fncNullCheck(rs("HEPBM2AN"))   ' 品EPBMD2平均下限
            .HEPBM2AX = fncNullCheck(rs("HEPBM2AX"))   ' 品EPBMD2平均上限
            .HEPBM3AN = fncNullCheck(rs("HEPBM3AN"))   ' 品EPBMD3平均下限
            .HEPBM3AX = fncNullCheck(rs("HEPBM3AX"))   ' 品EPBMD3平均上限
            .HEPBM3GSAN = fncNullCheck(rs("HEPBM3GSAN"))    ' 品EPBMD3平均下限(外周)　09/05/07 ooba
            .HEPBM3GSAX = fncNullCheck(rs("HEPBM3GSAX"))    ' 品EPBMD3平均上限(外周)　09/05/07 ooba
            .HEPOF1AX = fncNullCheck(rs("HEPOF1AX"))   ' 品EPOSF1平均下限
            .HEPOF1MX = fncNullCheck(rs("HEPOF1MX"))   ' 品EPOSF1上限
            .HEPOF2AX = fncNullCheck(rs("HEPOF2AX"))   ' 品EPOSF2平均下限
            .HEPOF2MX = fncNullCheck(rs("HEPOF2MX"))   ' 品EPOSF2上限
            .HEPOF3AX = fncNullCheck(rs("HEPOF3AX"))   ' 品EPOSF3平均下限
            .HEPOF3MX = fncNullCheck(rs("HEPOF3MX"))   ' 品EPOSF3上限
            .HEPANTNP = fncNullCheck(rs("HEPANTNP"))   ' 品EPAN温度
        End With
    End If
    
    rs.Close

proc_exit:
    ' 終了
'    gErr.Pop
    Exit Function

proc_err:
    ' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'************************************************************************************
'*    関数名        : funGetTBCMY013
'*
'*    処理概要      : 1.テーブル「TBCMY013」から条件にあったレコードを抽出する
'*                      (測定評価結果取得)
'*
'*    パラメータ    : 変数名       ,IO ,型                             ,説明
'*                   udtTypIn      ,I  ,type_DBDRV_scmzc_fcmlc001c_In  ,入力用
'*                   udtRecords()  ,O  ,typ_TBCMY013                   ,抽出レコード
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'************************************************************************************
Private Function funGetTBCMY013(udtTypIn As type_DBDRV_scmzc_fcmlc001c_In, udtRecords() As typ_TBCMY013) As FUNCTION_RETURN
    Dim sSQL        As String       ' SQL全体
    Dim rs          As OraDynaset   ' RecordSet
    Dim lngRecCnt   As Long         ' レコード数
    Dim i           As Long

    ' エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function funGetTBCMY013"

    ' SQLを組み立てる
    sSQL = "select SAMPLEID, OSITEM, MAISU, SPEC, NETSU, ET, MES, DKAN, MESDATA1, MESDATA2, MESDATA3, MESDATA4, MESDATA5, "
    sSQL = sSQL & "MESDATA6, MESDATA7, MESDATA8, MESDATA9, MESDATA10, MESDATA11, MESDATA12, MESDATA13, MESDATA14, MESDATA15, "
    sSQL = sSQL & "TXID, REGDATE, SENDFLAG, SENDDATE "
    sSQL = sSQL & "from TBCMY013 "
    sSQL = sSQL & "where ('" & udtTypIn.WFSMP.WFINDRSCW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDRSCW & "' and SPEC = '" & OSWFRES & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDOICW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDOICW & "' and SPEC = '" & OSWFOI & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDB1CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDB1CW & "' and SPEC = '" & OSWFBMD1 & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDB2CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDB2CW & "' and SPEC = '" & OSWFBMD2 & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDB3CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDB3CW & "' and SPEC = '" & OSWFBMD3 & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDL1CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDL1CW & "' and SPEC = '" & OSWFOSF1 & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDL2CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDL2CW & "' and SPEC = '" & OSWFOSF2 & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDL3CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDL3CW & "' and SPEC = '" & OSWFOSF3 & "') or "
'    sSql = sSql & "      ('" & udtTypIn.WFSMP.WFINDL4CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDL4CW & "' and SPEC = '" & OSWFOSF4 & "') or "
'Upd 2010/01/07 SIRD対応 Y.Hitomi
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDL4CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDL4CW & "' and SPEC = '" & OSWFSIRD & "') or "
    
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDDSCW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDDSCW & "' and SPEC = '" & OSWFDS & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDDZCW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDDZCW & "' and SPEC = '" & OSWFDZ & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDDO1CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDDO1CW & "' and SPEC = '" & OSWFDOI1 & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDDO2CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDDO2CW & "' and SPEC = '" & OSWFDOI2 & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDDO3CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDDO3CW & "' and SPEC = '" & OSWFDOI3 & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDOT1CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDOT1CW & "' and SPEC = '" & OSWFOT1 & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDOT2CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDOT2CW & "' and SPEC = '" & OSWFOT2 & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDAOICW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDAOICW & "' and SPEC = '" & OSWFAOI & "')"
    
    Debug.Print sSQL
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)
    If rs Is Nothing Then
        Set rs = Nothing
        ReDim udtRecords(0)
        funGetTBCMY013 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ' 抽出結果を格納する
    lngRecCnt = rs.RecordCount
    ReDim udtRecords(lngRecCnt)
    For i = 1 To lngRecCnt
        With udtRecords(i)
            .SAMPLEID = rs("SAMPLEID")       ' サンプルID
            .OSITEM = rs("OSITEM")           ' 評価項目
            .MAISU = rs("MAISU")             ' 評価枚数
            .Spec = rs("SPEC")               ' 規格値
            .NETSU = rs("NETSU")             ' 熱処理条件
            .ET = rs("ET")                   ' エッチング条件
            .MES = rs("MES")                 ' 計測方法
            .DKAN = rs("DKAN")               ' ＤＫアニール条件
            .MESDATA1 = rs("MESDATA1")       ' 測定データその１
            .MESDATA2 = rs("MESDATA2")       ' 測定データその２
            .MESDATA3 = rs("MESDATA3")       ' 測定データその３
            .MESDATA4 = rs("MESDATA4")       ' 測定データその４
            .MESDATA5 = rs("MESDATA5")       ' 測定データその５
            .MESDATA6 = rs("MESDATA6")       ' 測定データその６
            .MESDATA7 = rs("MESDATA7")       ' 測定データその７
            .MESDATA8 = rs("MESDATA8")       ' 測定データその８
            .MESDATA9 = rs("MESDATA9")       ' 測定データその９
            .MESDATA10 = rs("MESDATA10")     ' 測定データその１０
            .MESDATA11 = rs("MESDATA11")     ' 測定データその1１
            .MESDATA12 = rs("MESDATA12")     ' 測定データその1２
            .MESDATA13 = rs("MESDATA13")     ' 測定データその1３
            .MESDATA14 = rs("MESDATA14")     ' 測定データその1４
            .MESDATA15 = rs("MESDATA15")     ' 測定データその1５
            .TXID = rs("TXID")               ' トランザクションID
            .REGDATE = rs("REGDATE")         ' 登録日付
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
        End With
        rs.MoveNext
    Next
    Set rs = Nothing

    funGetTBCMY013 = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' 終了
    Exit Function

proc_err:
    ' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    funGetTBCMY013 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'************************************************************************************
'*    関数名        : funGetGDJisseki_J006
'*
'*    処理概要      : 1.結晶GD実績(TBCMJ006)の取得処理
'*
'*    パラメータ    : 変数名       ,IO ,型            ,説明
'*                   sCryNum       ,I  ,String        ,入力用
'*                   sSmplID       ,I  ,String        ,抽出レコード
'*                   udtGDjisseki  ,O  ,typ_TBCMJ015  ,結晶GD実績(構造体)
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'************************************************************************************
Public Function funGetGDJisseki_J006(sCryNum As String, sSmplID As String, _
                                                    udtGDjisseki As typ_TBCMJ015) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim lngSmplID   As Long         ' ﾃﾞｰﾀ型変更
    
    ' エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function funGetGDJisseki_J006"
    
    ' ｻﾝﾌﾟﾙIDが数値でない場合
    If IsNumeric(sSmplID) = False Then
        funGetGDJisseki_J006 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    lngSmplID = CLng(sSmplID)         'ﾃﾞｰﾀ型変更
    
    ' 結晶番号、ｻﾝﾌﾟﾙIDからTBCMJ006の結晶GD実績値を検索する。
    sSQL = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, MSRSDEN, MSRSLDL, MSRSDVD2, "
    sSQL = sSQL & "MS01LDL1, MS01LDL2, MS01LDL3, MS01LDL4, MS01LDL5, MS01DEN1, MS01DEN2, MS01DEN3, MS01DEN4, MS01DEN5, "
    sSQL = sSQL & "MS02LDL1, MS02LDL2, MS02LDL3, MS02LDL4, MS02LDL5, MS02DEN1, MS02DEN2, MS02DEN3, MS02DEN4, MS02DEN5, "
    sSQL = sSQL & "MS03LDL1, MS03LDL2, MS03LDL3, MS03LDL4, MS03LDL5, MS03DEN1, MS03DEN2, MS03DEN3, MS03DEN4, MS03DEN5, "
    sSQL = sSQL & "MS04LDL1, MS04LDL2, MS04LDL3, MS04LDL4, MS04LDL5, MS04DEN1, MS04DEN2, MS04DEN3, MS04DEN4, MS04DEN5, "
    sSQL = sSQL & "MS05LDL1, MS05LDL2, MS05LDL3, MS05LDL4, MS05LDL5, MS05DEN1, MS05DEN2, MS05DEN3, MS05DEN4, MS05DEN5, "
    sSQL = sSQL & "MS06LDL1, MS06LDL2, MS06LDL3, MS06LDL4, MS06LDL5, MS06DEN1, MS06DEN2, MS06DEN3, MS06DEN4, MS06DEN5, "
    sSQL = sSQL & "MS07LDL1, MS07LDL2, MS07LDL3, MS07LDL4, MS07LDL5, MS07DEN1, MS07DEN2, MS07DEN3, MS07DEN4, MS07DEN5, "
    sSQL = sSQL & "MS08LDL1, MS08LDL2, MS08LDL3, MS08LDL4, MS08LDL5, MS08DEN1, MS08DEN2, MS08DEN3, MS08DEN4, MS08DEN5, "
    sSQL = sSQL & "MS09LDL1, MS09LDL2, MS09LDL3, MS09LDL4, MS09LDL5, MS09DEN1, MS09DEN2, MS09DEN3, MS09DEN4, MS09DEN5, "
    sSQL = sSQL & "MS10LDL1, MS10LDL2, MS10LDL3, MS10LDL4, MS10LDL5, MS10DEN1, MS10DEN2, MS10DEN3, MS10DEN4, MS10DEN5, "
    sSQL = sSQL & "MS11LDL1, MS11LDL2, MS11LDL3, MS11LDL4, MS11LDL5, MS11DEN1, MS11DEN2, MS11DEN3, MS11DEN4, MS11DEN5, "
    sSQL = sSQL & "MS12LDL1, MS12LDL2, MS12LDL3, MS12LDL4, MS12LDL5, MS12DEN1, MS12DEN2, MS12DEN3, MS12DEN4, MS12DEN5, "
    sSQL = sSQL & "MS13LDL1, MS13LDL2, MS13LDL3, MS13LDL4, MS13LDL5, MS13DEN1, MS13DEN2, MS13DEN3, MS13DEN4, MS13DEN5, "
    sSQL = sSQL & "MS14LDL1, MS14LDL2, MS14LDL3, MS14LDL4, MS14LDL5, MS14DEN1, MS14DEN2, MS14DEN3, MS14DEN4, MS14DEN5, "
    sSQL = sSQL & "MS15LDL1, MS15LDL2, MS15LDL3, MS15LDL4, MS15LDL5, MS15DEN1, MS15DEN2, MS15DEN3, MS15DEN4, MS15DEN5, "
    sSQL = sSQL & "MS01DVD2, MS02DVD2, MS03DVD2, MS04DVD2, MS05DVD2, REGDATE "
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    sSQL = sSQL & ", MSZEROMN, MSZEROMX "
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
    sSQL = sSQL & "from TBCMJ006 "
    sSQL = sSQL & "where CRYNUM = '" & sCryNum & "' and "
    sSQL = sSQL & "      SMPLNO = " & lngSmplID & " and "
    sSQL = sSQL & "      TRANCNT = (select max(TRANCNT) from TBCMJ006 "
    sSQL = sSQL & "                 where CRYNUM = '" & sCryNum & "' and "
    sSQL = sSQL & "                       SMPLNO = " & lngSmplID & ")"
    
    ' SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' 該当ﾃﾞｰﾀなし
    If rs.EOF Then
        udtGDjisseki.SMPLNO = ""
        funGetGDJisseki_J006 = FUNCTION_RETURN_SUCCESS
        Set rs = Nothing
        GoTo proc_exit
    End If

    With udtGDjisseki
        .CRYNUM = rs("CRYNUM")          ' 結晶番号
        .POSITION = rs("POSITION")      ' 位置
        .SMPKBN = rs("SMPKBN")          ' サンプル区分
        .TRANCOND = rs("TRANCOND")      ' 処理条件
        .TRANCNT = rs("TRANCNT")        ' 処理回数
        .HSFLG = "0"                    ' 保証フラグ
        .SMPLNO = CStr(rs("SMPLNO"))    ' サンプルＮｏ
        .SMPLUMU = rs("SMPLUMU")        ' サンプル有無
        .MSRSDEN = rs("MSRSDEN")        ' 測定結果 Den
        .MSRSLDL = rs("MSRSLDL")        ' 測定結果 L/DL
        .MSRSDVD2 = rs("MSRSDVD2")      ' 測定結果 DVD2
        .MS01LDL1 = rs("MS01LDL1")      ' 測定値01 L/DL1
        .MS01LDL2 = rs("MS01LDL2")      ' 測定値01 L/DL2
        .MS01LDL3 = rs("MS01LDL3")      ' 測定値01 L/DL3
        .MS01LDL4 = rs("MS01LDL4")      ' 測定値01 L/DL4
        .MS01LDL5 = rs("MS01LDL5")      ' 測定値01 L/DL5
        .MS01DEN1 = rs("MS01DEN1")      ' 測定値01 Den1
        .MS01DEN2 = rs("MS01DEN2")      ' 測定値01 Den2
        .MS01DEN3 = rs("MS01DEN3")      ' 測定値01 Den3
        .MS01DEN4 = rs("MS01DEN4")      ' 測定値01 Den4
        .MS01DEN5 = rs("MS01DEN5")      ' 測定値01 Den5
        .MS02LDL1 = rs("MS02LDL1")      ' 測定値02 L/DL1
        .MS02LDL2 = rs("MS02LDL2")      ' 測定値02 L/DL2
        .MS02LDL3 = rs("MS02LDL3")      ' 測定値02 L/DL3
        .MS02LDL4 = rs("MS02LDL4")      ' 測定値02 L/DL4
        .MS02LDL5 = rs("MS02LDL5")      ' 測定値02 L/DL5
        .MS02DEN1 = rs("MS02DEN1")      ' 測定値02 Den1
        .MS02DEN2 = rs("MS02DEN2")      ' 測定値02 Den2
        .MS02DEN3 = rs("MS02DEN3")      ' 測定値02 Den3
        .MS02DEN4 = rs("MS02DEN4")      ' 測定値02 Den4
        .MS02DEN5 = rs("MS02DEN5")      ' 測定値02 Den5
        .MS03LDL1 = rs("MS03LDL1")      ' 測定値03 L/DL1
        .MS03LDL2 = rs("MS03LDL2")      ' 測定値03 L/DL2
        .MS03LDL3 = rs("MS03LDL3")      ' 測定値03 L/DL3
        .MS03LDL4 = rs("MS03LDL4")      ' 測定値03 L/DL4
        .MS03LDL5 = rs("MS03LDL5")      ' 測定値03 L/DL5
        .MS03DEN1 = rs("MS03DEN1")      ' 測定値03 Den1
        .MS03DEN2 = rs("MS03DEN2")      ' 測定値03 Den2
        .MS03DEN3 = rs("MS03DEN3")      ' 測定値03 Den3
        .MS03DEN4 = rs("MS03DEN4")      ' 測定値03 Den4
        .MS03DEN5 = rs("MS03DEN5")      ' 測定値03 Den5
        .MS04LDL1 = rs("MS04LDL1")      ' 測定値04 L/DL1
        .MS04LDL2 = rs("MS04LDL2")      ' 測定値04 L/DL2
        .MS04LDL3 = rs("MS04LDL3")      ' 測定値04 L/DL3
        .MS04LDL4 = rs("MS04LDL4")      ' 測定値04 L/DL4
        .MS04LDL5 = rs("MS04LDL5")      ' 測定値04 L/DL5
        .MS04DEN1 = rs("MS04DEN1")      ' 測定値04 Den1
        .MS04DEN2 = rs("MS04DEN2")      ' 測定値04 Den2
        .MS04DEN3 = rs("MS04DEN3")      ' 測定値04 Den3
        .MS04DEN4 = rs("MS04DEN4")      ' 測定値04 Den4
        .MS04DEN5 = rs("MS04DEN5")      ' 測定値04 Den5
        .MS05LDL1 = rs("MS05LDL1")      ' 測定値05 L/DL1
        .MS05LDL2 = rs("MS05LDL2")      ' 測定値05 L/DL2
        .MS05LDL3 = rs("MS05LDL3")      ' 測定値05 L/DL3
        .MS05LDL4 = rs("MS05LDL4")      ' 測定値05 L/DL4
        .MS05LDL5 = rs("MS05LDL5")      ' 測定値05 L/DL5
        .MS05DEN1 = rs("MS05DEN1")      ' 測定値05 Den1
        .MS05DEN2 = rs("MS05DEN2")      ' 測定値05 Den2
        .MS05DEN3 = rs("MS05DEN3")      ' 測定値05 Den3
        .MS05DEN4 = rs("MS05DEN4")      ' 測定値05 Den4
        .MS05DEN5 = rs("MS05DEN5")      ' 測定値05 Den5
        .MS06LDL1 = rs("MS06LDL1")      ' 測定値06 L/DL1
        .MS06LDL2 = rs("MS06LDL2")      ' 測定値06 L/DL2
        .MS06LDL3 = rs("MS06LDL3")      ' 測定値06 L/DL3
        .MS06LDL4 = rs("MS06LDL4")      ' 測定値06 L/DL4
        .MS06LDL5 = rs("MS06LDL5")      ' 測定値06 L/DL5
        .MS06DEN1 = rs("MS06DEN1")      ' 測定値06 Den1
        .MS06DEN2 = rs("MS06DEN2")      ' 測定値06 Den2
        .MS06DEN3 = rs("MS06DEN3")      ' 測定値06 Den3
        .MS06DEN4 = rs("MS06DEN4")      ' 測定値06 Den4
        .MS06DEN5 = rs("MS06DEN5")      ' 測定値06 Den5
        .MS07LDL1 = rs("MS07LDL1")      ' 測定値07 L/DL1
        .MS07LDL2 = rs("MS07LDL2")      ' 測定値07 L/DL2
        .MS07LDL3 = rs("MS07LDL3")      ' 測定値07 L/DL3
        .MS07LDL4 = rs("MS07LDL4")      ' 測定値07 L/DL4
        .MS07LDL5 = rs("MS07LDL5")      ' 測定値07 L/DL5
        .MS07DEN1 = rs("MS07DEN1")      ' 測定値07 Den1
        .MS07DEN2 = rs("MS07DEN2")      ' 測定値07 Den2
        .MS07DEN3 = rs("MS07DEN3")      ' 測定値07 Den3
        .MS07DEN4 = rs("MS07DEN4")      ' 測定値07 Den4
        .MS07DEN5 = rs("MS07DEN5")      ' 測定値07 Den5
        .MS08LDL1 = rs("MS08LDL1")      ' 測定値08 L/DL1
        .MS08LDL2 = rs("MS08LDL2")      ' 測定値08 L/DL2
        .MS08LDL3 = rs("MS08LDL3")      ' 測定値08 L/DL3
        .MS08LDL4 = rs("MS08LDL4")      ' 測定値08 L/DL4
        .MS08LDL5 = rs("MS08LDL5")      ' 測定値08 L/DL5
        .MS08DEN1 = rs("MS08DEN1")      ' 測定値08 Den1
        .MS08DEN2 = rs("MS08DEN2")      ' 測定値08 Den2
        .MS08DEN3 = rs("MS08DEN3")      ' 測定値08 Den3
        .MS08DEN4 = rs("MS08DEN4")      ' 測定値08 Den4
        .MS08DEN5 = rs("MS08DEN5")      ' 測定値08 Den5
        .MS09LDL1 = rs("MS09LDL1")      ' 測定値09 L/DL1
        .MS09LDL2 = rs("MS09LDL2")      ' 測定値09 L/DL2
        .MS09LDL3 = rs("MS09LDL3")      ' 測定値09 L/DL3
        .MS09LDL4 = rs("MS09LDL4")      ' 測定値09 L/DL4
        .MS09LDL5 = rs("MS09LDL5")      ' 測定値09 L/DL5
        .MS09DEN1 = rs("MS09DEN1")      ' 測定値09 Den1
        .MS09DEN2 = rs("MS09DEN2")      ' 測定値09 Den2
        .MS09DEN3 = rs("MS09DEN3")      ' 測定値09 Den3
        .MS09DEN4 = rs("MS09DEN4")      ' 測定値09 Den4
        .MS09DEN5 = rs("MS09DEN5")      ' 測定値09 Den5
        .MS10LDL1 = rs("MS10LDL1")      ' 測定値10 L/DL1
        .MS10LDL2 = rs("MS10LDL2")      ' 測定値10 L/DL2
        .MS10LDL3 = rs("MS10LDL3")      ' 測定値10 L/DL3
        .MS10LDL4 = rs("MS10LDL4")      ' 測定値10 L/DL4
        .MS10LDL5 = rs("MS10LDL5")      ' 測定値10 L/DL5
        .MS10DEN1 = rs("MS10DEN1")      ' 測定値10 Den1
        .MS10DEN2 = rs("MS10DEN2")      ' 測定値10 Den2
        .MS10DEN3 = rs("MS10DEN3")      ' 測定値10 Den3
        .MS10DEN4 = rs("MS10DEN4")      ' 測定値10 Den4
        .MS10DEN5 = rs("MS10DEN5")      ' 測定値10 Den5
        .MS11LDL1 = rs("MS11LDL1")      ' 測定値11 L/DL1
        .MS11LDL2 = rs("MS11LDL2")      ' 測定値11 L/DL2
        .MS11LDL3 = rs("MS11LDL3")      ' 測定値11 L/DL3
        .MS11LDL4 = rs("MS11LDL4")      ' 測定値11 L/DL4
        .MS11LDL5 = rs("MS11LDL5")      ' 測定値11 L/DL5
        .MS11DEN1 = rs("MS11DEN1")      ' 測定値11 Den1
        .MS11DEN2 = rs("MS11DEN2")      ' 測定値11 Den2
        .MS11DEN3 = rs("MS11DEN3")      ' 測定値11 Den3
        .MS11DEN4 = rs("MS11DEN4")      ' 測定値11 Den4
        .MS11DEN5 = rs("MS11DEN5")      ' 測定値11 Den5
        .MS12LDL1 = rs("MS12LDL1")      ' 測定値12 L/DL1
        .MS12LDL2 = rs("MS12LDL2")      ' 測定値12 L/DL2
        .MS12LDL3 = rs("MS12LDL3")      ' 測定値12 L/DL3
        .MS12LDL4 = rs("MS12LDL4")      ' 測定値12 L/DL4
        .MS12LDL5 = rs("MS12LDL5")      ' 測定値12 L/DL5
        .MS12DEN1 = rs("MS12DEN1")      ' 測定値12 Den1
        .MS12DEN2 = rs("MS12DEN2")      ' 測定値12 Den2
        .MS12DEN3 = rs("MS12DEN3")      ' 測定値12 Den3
        .MS12DEN4 = rs("MS12DEN4")      ' 測定値12 Den4
        .MS12DEN5 = rs("MS12DEN5")      ' 測定値12 Den5
        .MS13LDL1 = rs("MS13LDL1")      ' 測定値13 L/DL1
        .MS13LDL2 = rs("MS13LDL2")      ' 測定値13 L/DL2
        .MS13LDL3 = rs("MS13LDL3")      ' 測定値13 L/DL3
        .MS13LDL4 = rs("MS13LDL4")      ' 測定値13 L/DL4
        .MS13LDL5 = rs("MS13LDL5")      ' 測定値13 L/DL5
        .MS13DEN1 = rs("MS13DEN1")      ' 測定値13 Den1
        .MS13DEN2 = rs("MS13DEN2")      ' 測定値13 Den2
        .MS13DEN3 = rs("MS13DEN3")      ' 測定値13 Den3
        .MS13DEN4 = rs("MS13DEN4")      ' 測定値13 Den4
        .MS13DEN5 = rs("MS13DEN5")      ' 測定値13 Den5
        .MS14LDL1 = rs("MS14LDL1")      ' 測定値14 L/DL1
        .MS14LDL2 = rs("MS14LDL2")      ' 測定値14 L/DL2
        .MS14LDL3 = rs("MS14LDL3")      ' 測定値14 L/DL3
        .MS14LDL4 = rs("MS14LDL4")      ' 測定値14 L/DL4
        .MS14LDL5 = rs("MS14LDL5")      ' 測定値14 L/DL5
        .MS14DEN1 = rs("MS14DEN1")      ' 測定値14 Den1
        .MS14DEN2 = rs("MS14DEN2")      ' 測定値14 Den2
        .MS14DEN3 = rs("MS14DEN3")      ' 測定値14 Den3
        .MS14DEN4 = rs("MS14DEN4")      ' 測定値14 Den4
        .MS14DEN5 = rs("MS14DEN5")      ' 測定値14 Den5
        .MS15LDL1 = rs("MS15LDL1")      ' 測定値15 L/DL1
        .MS15LDL2 = rs("MS15LDL2")      ' 測定値15 L/DL2
        .MS15LDL3 = rs("MS15LDL3")      ' 測定値15 L/DL3
        .MS15LDL4 = rs("MS15LDL4")      ' 測定値15 L/DL4
        .MS15LDL5 = rs("MS15LDL5")      ' 測定値15 L/DL5
        .MS15DEN1 = rs("MS15DEN1")      ' 測定値15 Den1
        .MS15DEN2 = rs("MS15DEN2")      ' 測定値15 Den2
        .MS15DEN3 = rs("MS15DEN3")      ' 測定値15 Den3
        .MS15DEN4 = rs("MS15DEN4")      ' 測定値15 Den4
        .MS15DEN5 = rs("MS15DEN5")      ' 測定値15 Den5
        If IsNull(rs("MS01DVD2")) = False Then .MS01DVD2 = rs("MS01DVD2") Else .MS01DVD2 = -1   ' 測定値01 DVD2
        If IsNull(rs("MS02DVD2")) = False Then .MS02DVD2 = rs("MS02DVD2") Else .MS02DVD2 = -1   ' 測定値02 DVD2
        If IsNull(rs("MS03DVD2")) = False Then .MS03DVD2 = rs("MS03DVD2") Else .MS03DVD2 = -1   ' 測定値03 DVD2
        If IsNull(rs("MS04DVD2")) = False Then .MS04DVD2 = rs("MS04DVD2") Else .MS04DVD2 = -1   ' 測定値04 DVD2
        If IsNull(rs("MS05DVD2")) = False Then .MS05DVD2 = rs("MS05DVD2") Else .MS05DVD2 = -1   ' 測定値05 DVD2
        .REGDATE = rs("REGDATE")        ' 登録日付
        
        '↓追加 熱処理判断処理追加
        '2.1.3 AN温度 実績反映チェック追加
        '結晶のデータはAN温度を持っていないので、表示しないようにする
        'DBData2DispDateでデータ整形しているので、それにあわせて-1をいれる
        .DKAN = "  -1"
        
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
        If IsNull(rs("MSZEROMN")) = False Then .MSZEROMN = rs("MSZEROMN") Else .MSZEROMN = -1   ' L/DL0連続数最小値
        If IsNull(rs("MSZEROMX")) = False Then .MSZEROMX = rs("MSZEROMX") Else .MSZEROMX = -1   ' L/DL0連続数最大値
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
        
    End With
    
    Set rs = Nothing

    funGetGDJisseki_J006 = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' 終了
    Exit Function

proc_err:
    ' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    Set rs = Nothing
    funGetGDJisseki_J006 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'************************************************************************************
'*    関数名        : funGetGDJisseki_J015
'*
'*    処理概要      : 1.GD実績(TBCMJ015)の取得処理
'*
'*    パラメータ    : 変数名       ,IO ,型            ,説明
'*                   sCryNum       ,I  ,String        ,入力用
'*                   sSmplID       ,I  ,String        ,抽出レコード
'*                   sHsFlg_XSDCW  ,I  ,String        ,保証FLG(0:WF実績、1:結晶実績)
'*                   udtGDjisseki  ,O  ,typ_TBCMJ015  ,結晶GD実績(構造体)
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'************************************************************************************
Public Function funGetGDJisseki_J015(sCryNum As String, sSmplID As String, _
                                        sHsFlg_XSDCW As String, udtGDjisseki As typ_TBCMJ015) As FUNCTION_RETURN
    Dim sSQL            As String
    Dim rs              As OraDynaset
    Dim intSmplID       As Integer
    Dim sHsFlg_J015     As String       ' 保証FLG(1:WF実績)
    
    ' エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function funGetGDJisseki_J015"
    
    If sHsFlg_XSDCW = "0" Then
        ' WF実績
        sHsFlg_J015 = "1"
    Else
        ' 結晶実績
        sHsFlg_J015 = "0"
    End If
    
    '結晶番号、ｻﾝﾌﾟﾙID、保証FLGからTBCMJ015のGD実績値を検索する。
    sSQL = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, HSFLG, SMPLNO, SMPLUMU, "
    sSQL = sSQL & "HINBAN, REVNUM, FACTORY, OPECOND, SXLID, KRPROCCD, PROCCODE, GOUKI, "
    sSQL = sSQL & "OSITEM, MAISU, SPEC, NETSU, ET, MES, DKAN, ETMAE_RYO01, ETATO_RYO01, MSRSDEN, MSRSLDL, MSRSDVD2, "
    sSQL = sSQL & "MS01LDL1, MS01LDL2, MS01LDL3, MS01LDL4, MS01LDL5, MS01DEN1, MS01DEN2, MS01DEN3, MS01DEN4, MS01DEN5, "
    sSQL = sSQL & "MS02LDL1, MS02LDL2, MS02LDL3, MS02LDL4, MS02LDL5, MS02DEN1, MS02DEN2, MS02DEN3, MS02DEN4, MS02DEN5, "
    sSQL = sSQL & "MS03LDL1, MS03LDL2, MS03LDL3, MS03LDL4, MS03LDL5, MS03DEN1, MS03DEN2, MS03DEN3, MS03DEN4, MS03DEN5, "
    sSQL = sSQL & "MS04LDL1, MS04LDL2, MS04LDL3, MS04LDL4, MS04LDL5, MS04DEN1, MS04DEN2, MS04DEN3, MS04DEN4, MS04DEN5, "
    sSQL = sSQL & "MS05LDL1, MS05LDL2, MS05LDL3, MS05LDL4, MS05LDL5, MS05DEN1, MS05DEN2, MS05DEN3, MS05DEN4, MS05DEN5, "
    sSQL = sSQL & "MS06LDL1, MS06LDL2, MS06LDL3, MS06LDL4, MS06LDL5, MS06DEN1, MS06DEN2, MS06DEN3, MS06DEN4, MS06DEN5, "
    sSQL = sSQL & "MS07LDL1, MS07LDL2, MS07LDL3, MS07LDL4, MS07LDL5, MS07DEN1, MS07DEN2, MS07DEN3, MS07DEN4, MS07DEN5, "
    sSQL = sSQL & "MS08LDL1, MS08LDL2, MS08LDL3, MS08LDL4, MS08LDL5, MS08DEN1, MS08DEN2, MS08DEN3, MS08DEN4, MS08DEN5, "
    sSQL = sSQL & "MS09LDL1, MS09LDL2, MS09LDL3, MS09LDL4, MS09LDL5, MS09DEN1, MS09DEN2, MS09DEN3, MS09DEN4, MS09DEN5, "
    sSQL = sSQL & "MS10LDL1, MS10LDL2, MS10LDL3, MS10LDL4, MS10LDL5, MS10DEN1, MS10DEN2, MS10DEN3, MS10DEN4, MS10DEN5, "
    sSQL = sSQL & "MS11LDL1, MS11LDL2, MS11LDL3, MS11LDL4, MS11LDL5, MS11DEN1, MS11DEN2, MS11DEN3, MS11DEN4, MS11DEN5, "
    sSQL = sSQL & "MS12LDL1, MS12LDL2, MS12LDL3, MS12LDL4, MS12LDL5, MS12DEN1, MS12DEN2, MS12DEN3, MS12DEN4, MS12DEN5, "
    sSQL = sSQL & "MS13LDL1, MS13LDL2, MS13LDL3, MS13LDL4, MS13LDL5, MS13DEN1, MS13DEN2, MS13DEN3, MS13DEN4, MS13DEN5, "
    sSQL = sSQL & "MS14LDL1, MS14LDL2, MS14LDL3, MS14LDL4, MS14LDL5, MS14DEN1, MS14DEN2, MS14DEN3, MS14DEN4, MS14DEN5, "
    sSQL = sSQL & "MS15LDL1, MS15LDL2, MS15LDL3, MS15LDL4, MS15LDL5, MS15DEN1, MS15DEN2, MS15DEN3, MS15DEN4, MS15DEN5, "
    sSQL = sSQL & "MS01DVD2, MS02DVD2, MS03DVD2, MS04DVD2, MS05DVD2, "
    sSQL = sSQL & "TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    sSQL = sSQL & ", MSZEROMN , MSZEROMX "
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
    sSQL = sSQL & "from TBCMJ015 "
    sSQL = sSQL & "where CRYNUM = '" & sCryNum & "' and "
    sSQL = sSQL & "      SMPLNO = '" & sSmplID & "' and "
    sSQL = sSQL & "      HSFLG = '" & sHsFlg_J015 & "' and "
    sSQL = sSQL & "      TRANCNT = (select max(TRANCNT) from TBCMJ015 "
    sSQL = sSQL & "                 where CRYNUM = '" & sCryNum & "' and "
    sSQL = sSQL & "                       SMPLNO = '" & sSmplID & "' and "
    sSQL = sSQL & "                       HSFLG = '" & sHsFlg_J015 & "')"
    
    ' SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' 該当ﾃﾞｰﾀなし
    If rs.EOF Then
        udtGDjisseki.SMPLNO = ""
        funGetGDJisseki_J015 = FUNCTION_RETURN_SUCCESS
        Set rs = Nothing
        GoTo proc_exit
    End If

    With udtGDjisseki
        If IsNull(rs("CRYNUM")) = False Then .CRYNUM = rs("CRYNUM")                             ' 結晶番号
        If IsNull(rs("POSITION")) = False Then .POSITION = rs("POSITION")                       ' 位置
        If IsNull(rs("SMPKBN")) = False Then .SMPKBN = rs("SMPKBN")                             ' サンプル区分
        If IsNull(rs("TRANCOND")) = False Then .TRANCOND = rs("TRANCOND")                       ' 処理条件
        If IsNull(rs("TRANCNT")) = False Then .TRANCNT = rs("TRANCNT")                          ' 処理回数
        If IsNull(rs("HSFLG")) = False Then .HSFLG = rs("HSFLG")                                ' 保証フラグ
        If IsNull(rs("SMPLNO")) = False Then .SMPLNO = rs("SMPLNO")                             ' サンプルＮｏ
        If IsNull(rs("SMPLUMU")) = False Then .SMPLUMU = rs("SMPLUMU")                          ' サンプル有無
        If IsNull(rs("HINBAN")) = False Then .hinban = rs("HINBAN")                             ' 品番
        If IsNull(rs("REVNUM")) = False Then .REVNUM = rs("REVNUM")                             ' 製品番号改訂番号
        If IsNull(rs("FACTORY")) = False Then .factory = rs("FACTORY")                         ' 工場
        If IsNull(rs("OPECOND")) = False Then .opecond = rs("OPECOND")                          ' 操業条件
        If IsNull(rs("SXLID")) = False Then .SXLID = rs("SXLID")                                ' SXLID
        If IsNull(rs("KRPROCCD")) = False Then .KRPROCCD = rs("KRPROCCD")                       ' 管理工程コード
        If IsNull(rs("PROCCODE")) = False Then .PROCCODE = rs("PROCCODE")                       ' 工程コード
        If IsNull(rs("GOUKI")) = False Then .GOUKI = rs("GOUKI")                                ' 号機
        If IsNull(rs("OSITEM")) = False Then .OSITEM = rs("OSITEM")                             ' 評価項目
        If IsNull(rs("MAISU")) = False Then .MAISU = rs("MAISU")                                ' 評価枚数
        If IsNull(rs("SPEC")) = False Then .Spec = rs("SPEC")                                   ' 規格値
        If IsNull(rs("NETSU")) = False Then .NETSU = rs("NETSU")                                ' 熱処理条件
        If IsNull(rs("ET")) = False Then .ET = rs("ET")                                         ' エッチング条件
        If IsNull(rs("MES")) = False Then .MES = rs("MES")                                      ' 計測方法
        If IsNull(rs("DKAN")) = False Then .DKAN = rs("DKAN")                                   ' ＤＫアニール条件
        If IsNull(rs("ETMAE_RYO01")) = False Then .ETMAE_RYO01 = rs("ETMAE_RYO01")              ' ET前重量01
        If IsNull(rs("ETATO_RYO01")) = False Then .ETATO_RYO01 = rs("ETATO_RYO01")              ' ET後重量01
        
        If IsNull(rs("MSRSDEN")) = False Then .MSRSDEN = rs("MSRSDEN") Else .MSRSDEN = -1       ' 測定結果 Den
        If IsNull(rs("MSRSLDL")) = False Then .MSRSLDL = rs("MSRSLDL") Else .MSRSLDL = -1       ' 測定結果 L/DL
        If IsNull(rs("MSRSDVD2")) = False Then .MSRSDVD2 = rs("MSRSDVD2") Else .MSRSDVD2 = -1   ' 測定結果 DVD2
        If IsNull(rs("MS01LDL1")) = False Then .MS01LDL1 = rs("MS01LDL1") Else .MS01LDL1 = -1   ' 測定値01 L/DL1
        If IsNull(rs("MS01LDL2")) = False Then .MS01LDL2 = rs("MS01LDL2") Else .MS01LDL2 = -1   ' 測定値01 L/DL2
        If IsNull(rs("MS01LDL3")) = False Then .MS01LDL3 = rs("MS01LDL3") Else .MS01LDL3 = -1   ' 測定値01 L/DL3
        If IsNull(rs("MS01LDL4")) = False Then .MS01LDL4 = rs("MS01LDL4") Else .MS01LDL4 = -1   ' 測定値01 L/DL4
        If IsNull(rs("MS01LDL5")) = False Then .MS01LDL5 = rs("MS01LDL5") Else .MS01LDL5 = -1   ' 測定値01 L/DL5
        If IsNull(rs("MS01DEN1")) = False Then .MS01DEN1 = rs("MS01DEN1") Else .MS01DEN1 = -1   ' 測定値01 Den1
        If IsNull(rs("MS01DEN2")) = False Then .MS01DEN2 = rs("MS01DEN2") Else .MS01DEN2 = -1   ' 測定値01 Den2
        If IsNull(rs("MS01DEN3")) = False Then .MS01DEN3 = rs("MS01DEN3") Else .MS01DEN3 = -1   ' 測定値01 Den3
        If IsNull(rs("MS01DEN4")) = False Then .MS01DEN4 = rs("MS01DEN4") Else .MS01DEN4 = -1   ' 測定値01 Den4
        If IsNull(rs("MS01DEN5")) = False Then .MS01DEN5 = rs("MS01DEN5") Else .MS01DEN5 = -1   ' 測定値01 Den5
        If IsNull(rs("MS02LDL1")) = False Then .MS02LDL1 = rs("MS02LDL1") Else .MS02LDL1 = -1   ' 測定値02 L/DL1
        If IsNull(rs("MS02LDL2")) = False Then .MS02LDL2 = rs("MS02LDL2") Else .MS02LDL2 = -1   ' 測定値02 L/DL2
        If IsNull(rs("MS02LDL3")) = False Then .MS02LDL3 = rs("MS02LDL3") Else .MS02LDL3 = -1   ' 測定値02 L/DL3
        If IsNull(rs("MS02LDL4")) = False Then .MS02LDL4 = rs("MS02LDL4") Else .MS02LDL4 = -1   ' 測定値02 L/DL4
        If IsNull(rs("MS02LDL5")) = False Then .MS02LDL5 = rs("MS02LDL5") Else .MS02LDL5 = -1   ' 測定値02 L/DL5
        If IsNull(rs("MS02DEN1")) = False Then .MS02DEN1 = rs("MS02DEN1") Else .MS02DEN1 = -1   ' 測定値02 Den1
        If IsNull(rs("MS02DEN2")) = False Then .MS02DEN2 = rs("MS02DEN2") Else .MS02DEN2 = -1   ' 測定値02 Den2
        If IsNull(rs("MS02DEN3")) = False Then .MS02DEN3 = rs("MS02DEN3") Else .MS02DEN3 = -1   ' 測定値02 Den3
        If IsNull(rs("MS02DEN4")) = False Then .MS02DEN4 = rs("MS02DEN4") Else .MS02DEN4 = -1   ' 測定値02 Den4
        If IsNull(rs("MS02DEN5")) = False Then .MS02DEN5 = rs("MS02DEN5") Else .MS02DEN5 = -1   ' 測定値02 Den5
        If IsNull(rs("MS03LDL1")) = False Then .MS03LDL1 = rs("MS03LDL1") Else .MS03LDL1 = -1   ' 測定値03 L/DL1
        If IsNull(rs("MS03LDL2")) = False Then .MS03LDL2 = rs("MS03LDL2") Else .MS03LDL2 = -1   ' 測定値03 L/DL2
        If IsNull(rs("MS03LDL3")) = False Then .MS03LDL3 = rs("MS03LDL3") Else .MS03LDL3 = -1   ' 測定値03 L/DL3
        If IsNull(rs("MS03LDL4")) = False Then .MS03LDL4 = rs("MS03LDL4") Else .MS03LDL4 = -1   ' 測定値03 L/DL4
        If IsNull(rs("MS03LDL5")) = False Then .MS03LDL5 = rs("MS03LDL5") Else .MS03LDL5 = -1   ' 測定値03 L/DL5
        If IsNull(rs("MS03DEN1")) = False Then .MS03DEN1 = rs("MS03DEN1") Else .MS03DEN1 = -1   ' 測定値03 Den1
        If IsNull(rs("MS03DEN2")) = False Then .MS03DEN2 = rs("MS03DEN2") Else .MS03DEN2 = -1   ' 測定値03 Den2
        If IsNull(rs("MS03DEN3")) = False Then .MS03DEN3 = rs("MS03DEN3") Else .MS03DEN3 = -1   ' 測定値03 Den3
        If IsNull(rs("MS03DEN4")) = False Then .MS03DEN4 = rs("MS03DEN4") Else .MS03DEN4 = -1   ' 測定値03 Den4
        If IsNull(rs("MS03DEN5")) = False Then .MS03DEN5 = rs("MS03DEN5") Else .MS03DEN5 = -1   ' 測定値03 Den5
        If IsNull(rs("MS04LDL1")) = False Then .MS04LDL1 = rs("MS04LDL1") Else .MS04LDL1 = -1   ' 測定値04 L/DL1
        If IsNull(rs("MS04LDL2")) = False Then .MS04LDL2 = rs("MS04LDL2") Else .MS04LDL2 = -1   ' 測定値04 L/DL2
        If IsNull(rs("MS04LDL3")) = False Then .MS04LDL3 = rs("MS04LDL3") Else .MS04LDL3 = -1   ' 測定値04 L/DL3
        If IsNull(rs("MS04LDL4")) = False Then .MS04LDL4 = rs("MS04LDL4") Else .MS04LDL4 = -1   ' 測定値04 L/DL4
        If IsNull(rs("MS04LDL5")) = False Then .MS04LDL5 = rs("MS04LDL5") Else .MS04LDL5 = -1   ' 測定値04 L/DL5
        If IsNull(rs("MS04DEN1")) = False Then .MS04DEN1 = rs("MS04DEN1") Else .MS04DEN1 = -1   ' 測定値04 Den1
        If IsNull(rs("MS04DEN2")) = False Then .MS04DEN2 = rs("MS04DEN2") Else .MS04DEN2 = -1   ' 測定値04 Den2
        If IsNull(rs("MS04DEN3")) = False Then .MS04DEN3 = rs("MS04DEN3") Else .MS04DEN3 = -1   ' 測定値04 Den3
        If IsNull(rs("MS04DEN4")) = False Then .MS04DEN4 = rs("MS04DEN4") Else .MS04DEN4 = -1   ' 測定値04 Den4
        If IsNull(rs("MS04DEN5")) = False Then .MS04DEN5 = rs("MS04DEN5") Else .MS04DEN5 = -1   ' 測定値04 Den5
        If IsNull(rs("MS05LDL1")) = False Then .MS05LDL1 = rs("MS05LDL1") Else .MS05LDL1 = -1   ' 測定値05 L/DL1
        If IsNull(rs("MS05LDL2")) = False Then .MS05LDL2 = rs("MS05LDL2") Else .MS05LDL2 = -1   ' 測定値05 L/DL2
        If IsNull(rs("MS05LDL3")) = False Then .MS05LDL3 = rs("MS05LDL3") Else .MS05LDL3 = -1   ' 測定値05 L/DL3
        If IsNull(rs("MS05LDL4")) = False Then .MS05LDL4 = rs("MS05LDL4") Else .MS05LDL4 = -1   ' 測定値05 L/DL4
        If IsNull(rs("MS05LDL5")) = False Then .MS05LDL5 = rs("MS05LDL5") Else .MS05LDL5 = -1   ' 測定値05 L/DL5
        If IsNull(rs("MS05DEN1")) = False Then .MS05DEN1 = rs("MS05DEN1") Else .MS05DEN1 = -1   ' 測定値05 Den1
        If IsNull(rs("MS05DEN2")) = False Then .MS05DEN2 = rs("MS05DEN2") Else .MS05DEN2 = -1   ' 測定値05 Den2
        If IsNull(rs("MS05DEN3")) = False Then .MS05DEN3 = rs("MS05DEN3") Else .MS05DEN3 = -1   ' 測定値05 Den3
        If IsNull(rs("MS05DEN4")) = False Then .MS05DEN4 = rs("MS05DEN4") Else .MS05DEN4 = -1   ' 測定値05 Den4
        If IsNull(rs("MS05DEN5")) = False Then .MS05DEN5 = rs("MS05DEN5") Else .MS05DEN5 = -1   ' 測定値05 Den5
        If IsNull(rs("MS06LDL1")) = False Then .MS06LDL1 = rs("MS06LDL1") Else .MS06LDL1 = -1   ' 測定値06 L/DL1
        If IsNull(rs("MS06LDL2")) = False Then .MS06LDL2 = rs("MS06LDL2") Else .MS06LDL2 = -1   ' 測定値06 L/DL2
        If IsNull(rs("MS06LDL3")) = False Then .MS06LDL3 = rs("MS06LDL3") Else .MS06LDL3 = -1   ' 測定値06 L/DL3
        If IsNull(rs("MS06LDL4")) = False Then .MS06LDL4 = rs("MS06LDL4") Else .MS06LDL4 = -1   ' 測定値06 L/DL4
        If IsNull(rs("MS06LDL5")) = False Then .MS06LDL5 = rs("MS06LDL5") Else .MS06LDL5 = -1   ' 測定値06 L/DL5
        If IsNull(rs("MS06DEN1")) = False Then .MS06DEN1 = rs("MS06DEN1") Else .MS06DEN1 = -1   ' 測定値06 Den1
        If IsNull(rs("MS06DEN2")) = False Then .MS06DEN2 = rs("MS06DEN2") Else .MS06DEN2 = -1   ' 測定値06 Den2
        If IsNull(rs("MS06DEN3")) = False Then .MS06DEN3 = rs("MS06DEN3") Else .MS06DEN3 = -1   ' 測定値06 Den3
        If IsNull(rs("MS06DEN4")) = False Then .MS06DEN4 = rs("MS06DEN4") Else .MS06DEN4 = -1   ' 測定値06 Den4
        If IsNull(rs("MS06DEN5")) = False Then .MS06DEN5 = rs("MS06DEN5") Else .MS06DEN5 = -1   ' 測定値06 Den5
        If IsNull(rs("MS07LDL1")) = False Then .MS07LDL1 = rs("MS07LDL1") Else .MS07LDL1 = -1   ' 測定値07 L/DL1
        If IsNull(rs("MS07LDL2")) = False Then .MS07LDL2 = rs("MS07LDL2") Else .MS07LDL2 = -1   ' 測定値07 L/DL2
        If IsNull(rs("MS07LDL3")) = False Then .MS07LDL3 = rs("MS07LDL3") Else .MS07LDL3 = -1   ' 測定値07 L/DL3
        If IsNull(rs("MS07LDL4")) = False Then .MS07LDL4 = rs("MS07LDL4") Else .MS07LDL4 = -1   ' 測定値07 L/DL4
        If IsNull(rs("MS07LDL5")) = False Then .MS07LDL5 = rs("MS07LDL5") Else .MS07LDL5 = -1   ' 測定値07 L/DL5
        If IsNull(rs("MS07DEN1")) = False Then .MS07DEN1 = rs("MS07DEN1") Else .MS07DEN1 = -1   ' 測定値07 Den1
        If IsNull(rs("MS07DEN2")) = False Then .MS07DEN2 = rs("MS07DEN2") Else .MS07DEN2 = -1   ' 測定値07 Den2
        If IsNull(rs("MS07DEN3")) = False Then .MS07DEN3 = rs("MS07DEN3") Else .MS07DEN3 = -1   ' 測定値07 Den3
        If IsNull(rs("MS07DEN4")) = False Then .MS07DEN4 = rs("MS07DEN4") Else .MS07DEN4 = -1   ' 測定値07 Den4
        If IsNull(rs("MS07DEN5")) = False Then .MS07DEN5 = rs("MS07DEN5") Else .MS07DEN5 = -1   ' 測定値07 Den5
        If IsNull(rs("MS08LDL1")) = False Then .MS08LDL1 = rs("MS08LDL1") Else .MS08LDL1 = -1   ' 測定値08 L/DL1
        If IsNull(rs("MS08LDL2")) = False Then .MS08LDL2 = rs("MS08LDL2") Else .MS08LDL2 = -1   ' 測定値08 L/DL2
        If IsNull(rs("MS08LDL3")) = False Then .MS08LDL3 = rs("MS08LDL3") Else .MS08LDL3 = -1   ' 測定値08 L/DL3
        If IsNull(rs("MS08LDL4")) = False Then .MS08LDL4 = rs("MS08LDL4") Else .MS08LDL4 = -1   ' 測定値08 L/DL4
        If IsNull(rs("MS08LDL5")) = False Then .MS08LDL5 = rs("MS08LDL5") Else .MS08LDL5 = -1   ' 測定値08 L/DL5
        If IsNull(rs("MS08DEN1")) = False Then .MS08DEN1 = rs("MS08DEN1") Else .MS08DEN1 = -1   ' 測定値08 Den1
        If IsNull(rs("MS08DEN2")) = False Then .MS08DEN2 = rs("MS08DEN2") Else .MS08DEN2 = -1   ' 測定値08 Den2
        If IsNull(rs("MS08DEN3")) = False Then .MS08DEN3 = rs("MS08DEN3") Else .MS08DEN3 = -1   ' 測定値08 Den3
        If IsNull(rs("MS08DEN4")) = False Then .MS08DEN4 = rs("MS08DEN4") Else .MS08DEN4 = -1   ' 測定値08 Den4
        If IsNull(rs("MS08DEN5")) = False Then .MS08DEN5 = rs("MS08DEN5") Else .MS08DEN5 = -1   ' 測定値08 Den5
        If IsNull(rs("MS09LDL1")) = False Then .MS09LDL1 = rs("MS09LDL1") Else .MS09LDL1 = -1   ' 測定値09 L/DL1
        If IsNull(rs("MS09LDL2")) = False Then .MS09LDL2 = rs("MS09LDL2") Else .MS09LDL2 = -1   ' 測定値09 L/DL2
        If IsNull(rs("MS09LDL3")) = False Then .MS09LDL3 = rs("MS09LDL3") Else .MS09LDL3 = -1   ' 測定値09 L/DL3
        If IsNull(rs("MS09LDL4")) = False Then .MS09LDL4 = rs("MS09LDL4") Else .MS09LDL4 = -1   ' 測定値09 L/DL4
        If IsNull(rs("MS09LDL5")) = False Then .MS09LDL5 = rs("MS09LDL5") Else .MS09LDL5 = -1   ' 測定値09 L/DL5
        If IsNull(rs("MS09DEN1")) = False Then .MS09DEN1 = rs("MS09DEN1") Else .MS09DEN1 = -1   ' 測定値09 Den1
        If IsNull(rs("MS09DEN2")) = False Then .MS09DEN2 = rs("MS09DEN2") Else .MS09DEN2 = -1   ' 測定値09 Den2
        If IsNull(rs("MS09DEN3")) = False Then .MS09DEN3 = rs("MS09DEN3") Else .MS09DEN3 = -1   ' 測定値09 Den3
        If IsNull(rs("MS09DEN4")) = False Then .MS09DEN4 = rs("MS09DEN4") Else .MS09DEN4 = -1   ' 測定値09 Den4
        If IsNull(rs("MS09DEN5")) = False Then .MS09DEN5 = rs("MS09DEN5") Else .MS09DEN5 = -1   ' 測定値09 Den5
        If IsNull(rs("MS10LDL1")) = False Then .MS10LDL1 = rs("MS10LDL1") Else .MS10LDL1 = -1   ' 測定値10 L/DL1
        If IsNull(rs("MS10LDL2")) = False Then .MS10LDL2 = rs("MS10LDL2") Else .MS10LDL2 = -1   ' 測定値10 L/DL2
        If IsNull(rs("MS10LDL3")) = False Then .MS10LDL3 = rs("MS10LDL3") Else .MS10LDL3 = -1   ' 測定値10 L/DL3
        If IsNull(rs("MS10LDL4")) = False Then .MS10LDL4 = rs("MS10LDL4") Else .MS10LDL4 = -1   ' 測定値10 L/DL4
        If IsNull(rs("MS10LDL5")) = False Then .MS10LDL5 = rs("MS10LDL5") Else .MS10LDL5 = -1   ' 測定値10 L/DL5
        If IsNull(rs("MS10DEN1")) = False Then .MS10DEN1 = rs("MS10DEN1") Else .MS10DEN1 = -1   ' 測定値10 Den1
        If IsNull(rs("MS10DEN2")) = False Then .MS10DEN2 = rs("MS10DEN2") Else .MS10DEN2 = -1   ' 測定値10 Den2
        If IsNull(rs("MS10DEN3")) = False Then .MS10DEN3 = rs("MS10DEN3") Else .MS10DEN3 = -1   ' 測定値10 Den3
        If IsNull(rs("MS10DEN4")) = False Then .MS10DEN4 = rs("MS10DEN4") Else .MS10DEN4 = -1   ' 測定値10 Den4
        If IsNull(rs("MS10DEN5")) = False Then .MS10DEN5 = rs("MS10DEN5") Else .MS10DEN5 = -1   ' 測定値10 Den5
        If IsNull(rs("MS11LDL1")) = False Then .MS11LDL1 = rs("MS11LDL1") Else .MS11LDL1 = -1   ' 測定値11 L/DL1
        If IsNull(rs("MS11LDL2")) = False Then .MS11LDL2 = rs("MS11LDL2") Else .MS11LDL2 = -1   ' 測定値11 L/DL2
        If IsNull(rs("MS11LDL3")) = False Then .MS11LDL3 = rs("MS11LDL3") Else .MS11LDL3 = -1   ' 測定値11 L/DL3
        If IsNull(rs("MS11LDL4")) = False Then .MS11LDL4 = rs("MS11LDL4") Else .MS11LDL4 = -1   ' 測定値11 L/DL4
        If IsNull(rs("MS11LDL5")) = False Then .MS11LDL5 = rs("MS11LDL5") Else .MS11LDL5 = -1   ' 測定値11 L/DL5
        If IsNull(rs("MS11DEN1")) = False Then .MS11DEN1 = rs("MS11DEN1") Else .MS11DEN1 = -1   ' 測定値11 Den1
        If IsNull(rs("MS11DEN2")) = False Then .MS11DEN2 = rs("MS11DEN2") Else .MS11DEN2 = -1   ' 測定値11 Den2
        If IsNull(rs("MS11DEN3")) = False Then .MS11DEN3 = rs("MS11DEN3") Else .MS11DEN3 = -1   ' 測定値11 Den3
        If IsNull(rs("MS11DEN4")) = False Then .MS11DEN4 = rs("MS11DEN4") Else .MS11DEN4 = -1   ' 測定値11 Den4
        If IsNull(rs("MS11DEN5")) = False Then .MS11DEN5 = rs("MS11DEN5") Else .MS11DEN5 = -1   ' 測定値11 Den5
        If IsNull(rs("MS12LDL1")) = False Then .MS12LDL1 = rs("MS12LDL1") Else .MS12LDL1 = -1   ' 測定値12 L/DL1
        If IsNull(rs("MS12LDL2")) = False Then .MS12LDL2 = rs("MS12LDL2") Else .MS12LDL2 = -1   ' 測定値12 L/DL2
        If IsNull(rs("MS12LDL3")) = False Then .MS12LDL3 = rs("MS12LDL3") Else .MS12LDL3 = -1   ' 測定値12 L/DL3
        If IsNull(rs("MS12LDL4")) = False Then .MS12LDL4 = rs("MS12LDL4") Else .MS12LDL4 = -1   ' 測定値12 L/DL4
        If IsNull(rs("MS12LDL5")) = False Then .MS12LDL5 = rs("MS12LDL5") Else .MS12LDL5 = -1   ' 測定値12 L/DL5
        If IsNull(rs("MS12DEN1")) = False Then .MS12DEN1 = rs("MS12DEN1") Else .MS12DEN1 = -1   ' 測定値12 Den1
        If IsNull(rs("MS12DEN2")) = False Then .MS12DEN2 = rs("MS12DEN2") Else .MS12DEN2 = -1   ' 測定値12 Den2
        If IsNull(rs("MS12DEN3")) = False Then .MS12DEN3 = rs("MS12DEN3") Else .MS12DEN3 = -1   ' 測定値12 Den3
        If IsNull(rs("MS12DEN4")) = False Then .MS12DEN4 = rs("MS12DEN4") Else .MS12DEN4 = -1   ' 測定値12 Den4
        If IsNull(rs("MS12DEN5")) = False Then .MS12DEN5 = rs("MS12DEN5") Else .MS12DEN5 = -1   ' 測定値12 Den5
        If IsNull(rs("MS13LDL1")) = False Then .MS13LDL1 = rs("MS13LDL1") Else .MS13LDL1 = -1   ' 測定値13 L/DL1
        If IsNull(rs("MS13LDL2")) = False Then .MS13LDL2 = rs("MS13LDL2") Else .MS13LDL2 = -1   ' 測定値13 L/DL2
        If IsNull(rs("MS13LDL3")) = False Then .MS13LDL3 = rs("MS13LDL3") Else .MS13LDL3 = -1   ' 測定値13 L/DL3
        If IsNull(rs("MS13LDL4")) = False Then .MS13LDL4 = rs("MS13LDL4") Else .MS13LDL4 = -1   ' 測定値13 L/DL4
        If IsNull(rs("MS13LDL5")) = False Then .MS13LDL5 = rs("MS13LDL5") Else .MS13LDL5 = -1   ' 測定値13 L/DL5
        If IsNull(rs("MS13DEN1")) = False Then .MS13DEN1 = rs("MS13DEN1") Else .MS13DEN1 = -1   ' 測定値13 Den1
        If IsNull(rs("MS13DEN2")) = False Then .MS13DEN2 = rs("MS13DEN2") Else .MS13DEN2 = -1   ' 測定値13 Den2
        If IsNull(rs("MS13DEN3")) = False Then .MS13DEN3 = rs("MS13DEN3") Else .MS13DEN3 = -1   ' 測定値13 Den3
        If IsNull(rs("MS13DEN4")) = False Then .MS13DEN4 = rs("MS13DEN4") Else .MS13DEN4 = -1   ' 測定値13 Den4
        If IsNull(rs("MS13DEN5")) = False Then .MS13DEN5 = rs("MS13DEN5") Else .MS13DEN5 = -1   ' 測定値13 Den5
        If IsNull(rs("MS14LDL1")) = False Then .MS14LDL1 = rs("MS14LDL1") Else .MS14LDL1 = -1   ' 測定値14 L/DL1
        If IsNull(rs("MS14LDL2")) = False Then .MS14LDL2 = rs("MS14LDL2") Else .MS14LDL2 = -1   ' 測定値14 L/DL2
        If IsNull(rs("MS14LDL3")) = False Then .MS14LDL3 = rs("MS14LDL3") Else .MS14LDL3 = -1   ' 測定値14 L/DL3
        If IsNull(rs("MS14LDL4")) = False Then .MS14LDL4 = rs("MS14LDL4") Else .MS14LDL4 = -1   ' 測定値14 L/DL4
        If IsNull(rs("MS14LDL5")) = False Then .MS14LDL5 = rs("MS14LDL5") Else .MS14LDL5 = -1   ' 測定値14 L/DL5
        If IsNull(rs("MS14DEN1")) = False Then .MS14DEN1 = rs("MS14DEN1") Else .MS14DEN1 = -1   ' 測定値14 Den1
        If IsNull(rs("MS14DEN2")) = False Then .MS14DEN2 = rs("MS14DEN2") Else .MS14DEN2 = -1   ' 測定値14 Den2
        If IsNull(rs("MS14DEN3")) = False Then .MS14DEN3 = rs("MS14DEN3") Else .MS14DEN3 = -1   ' 測定値14 Den3
        If IsNull(rs("MS14DEN4")) = False Then .MS14DEN4 = rs("MS14DEN4") Else .MS14DEN4 = -1   ' 測定値14 Den4
        If IsNull(rs("MS14DEN5")) = False Then .MS14DEN5 = rs("MS14DEN5") Else .MS14DEN5 = -1   ' 測定値14 Den5
        If IsNull(rs("MS15LDL1")) = False Then .MS15LDL1 = rs("MS15LDL1") Else .MS15LDL1 = -1   ' 測定値15 L/DL1
        If IsNull(rs("MS15LDL2")) = False Then .MS15LDL2 = rs("MS15LDL2") Else .MS15LDL2 = -1   ' 測定値15 L/DL2
        If IsNull(rs("MS15LDL3")) = False Then .MS15LDL3 = rs("MS15LDL3") Else .MS15LDL3 = -1   ' 測定値15 L/DL3
        If IsNull(rs("MS15LDL4")) = False Then .MS15LDL4 = rs("MS15LDL4") Else .MS15LDL4 = -1   ' 測定値15 L/DL4
        If IsNull(rs("MS15LDL5")) = False Then .MS15LDL5 = rs("MS15LDL5") Else .MS15LDL5 = -1   ' 測定値15 L/DL5
        If IsNull(rs("MS15DEN1")) = False Then .MS15DEN1 = rs("MS15DEN1") Else .MS15DEN1 = -1   ' 測定値15 Den1
        If IsNull(rs("MS15DEN2")) = False Then .MS15DEN2 = rs("MS15DEN2") Else .MS15DEN2 = -1   ' 測定値15 Den2
        If IsNull(rs("MS15DEN3")) = False Then .MS15DEN3 = rs("MS15DEN3") Else .MS15DEN3 = -1   ' 測定値15 Den3
        If IsNull(rs("MS15DEN4")) = False Then .MS15DEN4 = rs("MS15DEN4") Else .MS15DEN4 = -1   ' 測定値15 Den4
        If IsNull(rs("MS15DEN5")) = False Then .MS15DEN5 = rs("MS15DEN5") Else .MS15DEN5 = -1   ' 測定値15 Den5
        If IsNull(rs("MS01DVD2")) = False Then .MS01DVD2 = rs("MS01DVD2") Else .MS01DVD2 = -1   ' 測定値01 DVD2
        If IsNull(rs("MS02DVD2")) = False Then .MS02DVD2 = rs("MS02DVD2") Else .MS02DVD2 = -1   ' 測定値02 DVD2
        If IsNull(rs("MS03DVD2")) = False Then .MS03DVD2 = rs("MS03DVD2") Else .MS03DVD2 = -1   ' 測定値03 DVD2
        If IsNull(rs("MS04DVD2")) = False Then .MS04DVD2 = rs("MS04DVD2") Else .MS04DVD2 = -1   ' 測定値04 DVD2
        If IsNull(rs("MS05DVD2")) = False Then .MS05DVD2 = rs("MS05DVD2") Else .MS05DVD2 = -1   ' 測定値05 DVD2
        
        If IsNull(rs("TSTAFFID")) = False Then .TSTAFFID = rs("TSTAFFID")                       ' 登録社員ID
        If IsNull(rs("REGDATE")) = False Then .REGDATE = rs("REGDATE")                          ' 登録日付
        If IsNull(rs("KSTAFFID")) = False Then .KSTAFFID = rs("KSTAFFID")                       ' 更新社員ID
        If IsNull(rs("UPDDATE")) = False Then .UPDDATE = rs("UPDDATE")                          ' 更新日付
        If IsNull(rs("SENDFLAG")) = False Then .SENDFLAG = rs("SENDFLAG")                       ' 送信フラグ
        If IsNull(rs("SENDDATE")) = False Then .SENDDATE = rs("SENDDATE")                       ' 送信日付
    
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
        If IsNull(rs("MSZEROMN")) = False Then .MSZEROMN = rs("MSZEROMN") Else .MSZEROMN = -1   ' L/DL0連続数最小値
        If IsNull(rs("MSZEROMX")) = False Then .MSZEROMX = rs("MSZEROMX") Else .MSZEROMX = -1   ' L/DL0連続数最大値
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
    
    End With
    
    Set rs = Nothing

    funGetGDJisseki_J015 = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' 終了
    Exit Function

proc_err:
    ' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    Set rs = Nothing
    funGetGDJisseki_J015 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function
'************************************************************************************
'*    関数名        : funGetSDJisseki_J022
'*
'*    処理概要      : 1.SIRD実績(TBCMJ022)の取得処理
'*
'*    パラメータ    : 変数名       ,IO ,型            ,説明
'*                   sCryNum       ,I  ,String        ,入力用
'*                   sSmplID       ,I  ,String        ,抽出レコード
'*                   sHsFlg_XSDCW  ,I  ,String        ,保証FLG(0:WF実績、1:結晶実績)
'*                   udtGDjisseki  ,O  ,typ_TBCMJ022  ,SIRD実績(構造体)
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'************************************************************************************
Public Function funGetSDJisseki_J022(sCryNum As String, sSmplID As String, _
                                         udtSDjisseki As typ_TBCMJ022) As FUNCTION_RETURN
    Dim sSQL            As String
    Dim rs              As OraDynaset
    Dim intSmplID       As Integer
    Dim sHsFlg_J022     As String       ' 保証FLG(1:WF実績)

    ' エラーハンドラの設定
    On Error GoTo proc_err

'    If sHsFlg_XSDCW = "0" Then
'        ' WF実績
'        sHsFlg_J022 = "1"
'    Else
'        ' 結晶実績
'        sHsFlg_J015 = "0"
'    End If

    '結晶番号、ｻﾝﾌﾟﾙIDからTBCMJ022のSIRD実績値を検索する。
    sSQL = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, HSFLG, SMPLNO, SMPLUMU, "
    sSQL = sSQL & "HINBAN, REVNUM, FACTORY, OPECOND, SXLID, KRPROCCD, PROCCODE, GOUKI, "
    sSQL = sSQL & "OSITEM, MAISU, SPEC, NETSU, ET, MES, DKAN, "
    sSQL = sSQL & "SIRDCNT,"
    sSQL = sSQL & "TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sSQL = sSQL & "from TBCMJ022 "
    sSQL = sSQL & "where CRYNUM = '" & sCryNum & "' and "
    sSQL = sSQL & "      SMPLNO = '" & sSmplID & "' and "
    sSQL = sSQL & "      TRANCNT = 0 "

    ' SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    ' 該当ﾃﾞｰﾀなし
    If rs.EOF Then
        udtSDjisseki.SMPLNO = ""
        funGetSDJisseki_J022 = FUNCTION_RETURN_SUCCESS
        Set rs = Nothing
        GoTo proc_exit
    End If

    With udtSDjisseki
        If IsNull(rs("CRYNUM")) = False Then .CRYNUM = rs("CRYNUM")                             ' 結晶番号
        If IsNull(rs("POSITION")) = False Then .POSITION = rs("POSITION")                       ' 位置
        If IsNull(rs("SMPKBN")) = False Then .SMPKBN = rs("SMPKBN")                             ' サンプル区分
        If IsNull(rs("TRANCOND")) = False Then .TRANCOND = rs("TRANCOND")                       ' 処理条件
        If IsNull(rs("TRANCNT")) = False Then .TRANCNT = rs("TRANCNT")                          ' 処理回数
        If IsNull(rs("HSFLG")) = False Then .HSFLG = rs("HSFLG")                                ' 保証フラグ
        If IsNull(rs("SMPLNO")) = False Then .SMPLNO = rs("SMPLNO")                             ' サンプルＮｏ
        If IsNull(rs("SMPLUMU")) = False Then .SMPLUMU = rs("SMPLUMU")                          ' サンプル有無
        If IsNull(rs("HINBAN")) = False Then .hinban = rs("HINBAN")                             ' 品番
        If IsNull(rs("REVNUM")) = False Then .REVNUM = rs("REVNUM")                             ' 製品番号改訂番号
        If IsNull(rs("FACTORY")) = False Then .factory = rs("FACTORY")                         ' 工場
        If IsNull(rs("OPECOND")) = False Then .opecond = rs("OPECOND")                          ' 操業条件
        If IsNull(rs("SXLID")) = False Then .SXLID = rs("SXLID")                                ' SXLID
        If IsNull(rs("KRPROCCD")) = False Then .KRPROCCD = rs("KRPROCCD")                       ' 管理工程コード
        If IsNull(rs("PROCCODE")) = False Then .PROCCODE = rs("PROCCODE")                       ' 工程コード
        If IsNull(rs("GOUKI")) = False Then .GOUKI = rs("GOUKI")                                ' 号機
        If IsNull(rs("OSITEM")) = False Then .OSITEM = rs("OSITEM")                             ' 評価項目
        If IsNull(rs("MAISU")) = False Then .MAISU = rs("MAISU")                                ' 評価枚数
        If IsNull(rs("SPEC")) = False Then .Spec = rs("SPEC")                                   ' 規格値
        If IsNull(rs("NETSU")) = False Then .NETSU = rs("NETSU")                                ' 熱処理条件
        If IsNull(rs("ET")) = False Then .ET = rs("ET")                                         ' エッチング条件
        If IsNull(rs("MES")) = False Then .MES = rs("MES")                                      ' 計測方法
        If IsNull(rs("DKAN")) = False Then .DKAN = rs("DKAN")                                   ' ＤＫアニール条件
        If IsNull(rs("SIRDCNT")) = False Then .SIRDCNT = rs("SIRDCNT")                          ' 面内個数（SIRD個数)
    End With

    Set rs = Nothing

    funGetSDJisseki_J022 = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' 終了
    Exit Function

proc_err:
    ' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    Set rs = Nothing
    funGetSDJisseki_J022 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'************************************************************************************
'*    関数名        : s_cmmc001db_sSql
'*
'*    処理概要      : 1.引上げ終了実績取得関数
'*
'*    パラメータ    : 変数名       ,IO ,型            ,説明
'*                   sCryNum       ,I  ,String        ,入力用
'*                   udtTbcmh004   ,O  ,typ_TBCMH004  ,引上げ終了実績取得用
'*
'*    戻り値        : (Double)正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'************************************************************************************
Public Function s_cmmc001db_Sql(ByVal sCryNum As String, _
                udtTbcmh004() As typ_TBCMH004) As Double
    Dim sSQL    As String
    Dim intRET  As Integer
    
    ' エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function s_cmmc001db_sSql"

    sSQL = " where CRYNUM = '" & sCryNum & "' "

    If DBDRV_GetTBCMH004(udtTbcmh004, sSQL, "order by CRYNUM") = FUNCTION_RETURN_FAILURE Then
        s_cmmc001db_Sql = FUNCTION_RETURN_FAILURE
    Else
        s_cmmc001db_Sql = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    ' 終了
    Exit Function

proc_err:
    ' エラーハンドラ
    s_cmmc001db_Sql = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'****************************************************************************************
'*    関数名        : DBDRV_GetTBCMH001
'*
'*    処理概要      : 1.テーブル「TBCMH001」から条件にあったレコードを抽出する
'*
'*    パラメータ    : 変数名       ,IO ,型           ,説明
'*                   udtRecords()  ,O  ,typ_TBCMH001 ,抽出レコード
'*                   sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'*                   sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'****************************************************************************************
Public Function DBDRV_GetTBCMH001(udtRecords() As typ_TBCMH001, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
    Dim sSQL        As String       ' SQL全体
    Dim sSqlBase    As String       ' SQL基本部(WHERE節の前まで)
    Dim rs          As OraDynaset   ' RecordSet
    Dim lngRecCnt   As Long         ' レコード数
    Dim i           As Long

    ' SQLを組み立てる
    sSqlBase = "Select UPINDNO, KRPROCCD, PROCCODE, MODEL, GOUKI, PGID, CPORGIND, HINBAN, NMNOREVNO, NFACTORY, NOPECOND, NUMNOTE1," & _
              " NUMNOTE2, SEED, SEKIERTB, DPNTCLS, DOPANT, AMRESIST, CRYDOPCL, CRYDOPVL, UPBTCHNM, ADDDOPCL, ADDDOPVL, ADDDOPPT," & _
              " BCNT1COD, BCNT1CMT, BCNT2COD, BCNT2CMT, MTCLS1, MTWGHT1, ESWGHT1, MTCLS2, MTWGHT2, ESWGHT2, MTCLS3, MTWGHT3," & _
              " ESWGHT3, MTCLS4, MTWGHT4, ESWGHT4, MTCLS5, MTWGHT5, ESWGHT5, MTCLS6, MTWGHT6, ESWGHT6, MTCLS7, MTWGHT7, ESWGHT7," & _
              " MTCLS8, MTWGHT8, ESWGHT8, MTCLS9, MTWGHT9, ESWGHT9, MTCLS10, MTWGHT10, ESWGHT10, MTCLS11, MTWGHT11, ESWGHT11," & _
              " MTCLS12, MTWGHT12, ESWGHT12, MTCLS13, MTWGHT13, ESWGHT13, MTCLS14, MTWGHT14, ESWGHT14, MTCLS15, MTWGHT15," & _
              " ESWGHT15, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sSqlBase = sSqlBase & "From TBCMH001"
    sSQL = sSqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sSQL = sSQL & " " & sqlWhere & " " & sqlOrder
    End If

    ' データを抽出する
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim udtRecords(0)
        DBDRV_GetTBCMH001 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ' 抽出結果を格納する
    lngRecCnt = rs.RecordCount
    ReDim udtRecords(lngRecCnt)
    For i = 1 To lngRecCnt
        With udtRecords(i)
            .UPINDNO = rs("UPINDNO")         ' 引上げ指示No.
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCODE = rs("PROCCODE")       ' 工程コード
            .MODEL = rs("MODEL")             ' 機種
            .GOUKI = rs("GOUKI")             ' 号機
            .PGID = rs("PGID")               ' PG-ID
            .CPORGIND = rs("CPORGIND")       ' 複写元指示No
            .hinban = rs("HINBAN")           ' 品番
            .NMNOREVNO = rs("NMNOREVNO")     ' 製品番号改訂番号
            .NFACTORY = rs("NFACTORY")       ' 工場
            .NOPECOND = rs("NOPECOND")       ' 操業条件
            .NUMNOTE1 = rs("NUMNOTE1")       ' 品番備考１
            .NUMNOTE2 = rs("NUMNOTE2")       ' 品番備考２
            .SEED = rs("SEED")               ' シード
            .SEKIERTB = rs("SEKIERTB")       ' 石英ルツボ
            .DPNTCLS = rs("DPNTCLS")         ' ドーパント種類
            .DOPANT = rs("DOPANT")           ' ドーパント量
            .AMRESIST = rs("AMRESIST")       ' ねらい抵抗
            .CRYDOPCL = rs("CRYDOPCL")       ' 結晶ドープ種類
            .CRYDOPVL = rs("CRYDOPVL")       ' 結晶ドープ量
            .UPBTCHNM = rs("UPBTCHNM")       ' 引上げバッチ数
            .ADDDOPCL = rs("ADDDOPCL")       ' 追加ドーパント種類
            .ADDDOPVL = rs("ADDDOPVL")       ' 追加ドーパント量
            .ADDDOPPT = rs("ADDDOPPT")       ' 追加ドーパント位置
            .BCNT1COD = rs("BCNT1COD")       ' バッチ備考1（コード）
            .BCNT1CMT = rs("BCNT1CMT")       ' バッチ備考1（ｺﾒﾝﾄ）
            .BCNT2COD = rs("BCNT2COD")       ' バッチ備考2（コード）
            .BCNT2CMT = rs("BCNT2CMT")       ' バッチ備考2（ｺﾒﾝﾄ）
            .MTCLS1 = rs("MTCLS1")           ' 原料種類1
            .MTWGHT1 = rs("MTWGHT1")         ' 原料重量1
            .ESWGHT1 = rs("ESWGHT1")         ' 推定残重量1
            .MTCLS2 = rs("MTCLS2")           ' 原料種類2
            .MTWGHT2 = rs("MTWGHT2")         ' 原料重量2
            .ESWGHT2 = rs("ESWGHT2")         ' 推定残重量2
            .MTCLS3 = rs("MTCLS3")           ' 原料種類3
            .MTWGHT3 = rs("MTWGHT3")         ' 原料重量3
            .ESWGHT3 = rs("ESWGHT3")         ' 推定残重量3
            .MTCLS4 = rs("MTCLS4")           ' 原料種類4
            .MTWGHT4 = rs("MTWGHT4")         ' 原料重量4
            .ESWGHT4 = rs("ESWGHT4")         ' 推定残重量4
            .MTCLS5 = rs("MTCLS5")           ' 原料種類5
            .MTWGHT5 = rs("MTWGHT5")         ' 原料重量5
            .ESWGHT5 = rs("ESWGHT5")         ' 推定残重量5
            .MTCLS6 = rs("MTCLS6")           ' 原料種類6
            .MTWGHT6 = rs("MTWGHT6")         ' 原料重量6
            .ESWGHT6 = rs("ESWGHT6")         ' 推定残重量6
            .MTCLS7 = rs("MTCLS7")           ' 原料種類7
            .MTWGHT7 = rs("MTWGHT7")         ' 原料重量7
            .ESWGHT7 = rs("ESWGHT7")         ' 推定残重量7
            .MTCLS8 = rs("MTCLS8")           ' 原料種類8
            .MTWGHT8 = rs("MTWGHT8")         ' 原料重量8
            .ESWGHT8 = rs("ESWGHT8")         ' 推定残重量8
            .MTCLS9 = rs("MTCLS9")           ' 原料種類9
            .MTWGHT9 = rs("MTWGHT9")         ' 原料重量9
            .ESWGHT9 = rs("ESWGHT9")         ' 推定残重量9
            .MTCLS10 = rs("MTCLS10")         ' 原料種類10
            .MTWGHT10 = rs("MTWGHT10")       ' 原料重量10
            .ESWGHT10 = rs("ESWGHT10")       ' 推定残重量10
            .MTCLS11 = rs("MTCLS11")         ' 原料種類11
            .MTWGHT11 = rs("MTWGHT11")       ' 原料重量11
            .ESWGHT11 = rs("ESWGHT11")       ' 推定残重量11
            .MTCLS12 = rs("MTCLS12")         ' 原料種類12
            .MTWGHT12 = rs("MTWGHT12")       ' 原料重量12
            .ESWGHT12 = rs("ESWGHT12")       ' 推定残重量12
            .MTCLS13 = rs("MTCLS13")         ' 原料種類13
            .MTWGHT13 = rs("MTWGHT13")       ' 原料重量13
            .ESWGHT13 = rs("ESWGHT13")       ' 推定残重量13
            .MTCLS14 = rs("MTCLS14")         ' 原料種類14
            .MTWGHT14 = rs("MTWGHT14")       ' 原料重量14
            .ESWGHT14 = rs("ESWGHT14")       ' 推定残重量14
            .MTCLS15 = rs("MTCLS15")         ' 原料種類15
            .MTWGHT15 = rs("MTWGHT15")       ' 原料重量15
            .ESWGHT15 = rs("ESWGHT15")       ' 推定残重量15
            .TSTAFFID = rs("TSTAFFID")       ' 登録社員ID
            .REGDATE = rs("REGDATE")         ' 登録日付
            .KSTAFFID = rs("KSTAFFID")       ' 更新社員ID
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMH001 = FUNCTION_RETURN_SUCCESS
End Function

'****************************************************************************************
'*    関数名        : DBDRV_GetTBCMH004
'*
'*    処理概要      : 1.テーブル「TBCMH004」から条件にあったレコードを抽出する
'*
'*    パラメータ    : 変数名       ,IO ,型           ,説明
'*                   udtRecords()  ,O  ,typ_TBCMH004 ,抽出レコード
'*                   sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'*                   sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'****************************************************************************************
Public Function DBDRV_GetTBCMH004(udtRecords() As typ_TBCMH004, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
    Dim sSQL        As String       ' SQL全体
    Dim sSqlBase    As String       ' SQL基本部(WHERE節の前まで)
    Dim rs          As OraDynaset   ' RecordSet
    Dim lngRecCnt   As Long         ' レコード数
    Dim i           As Long

    ' SQLを組み立てる
    sSqlBase = "Select CRYNUM, KRPROCCD, PROCCODE, LENGTOP, LENGTKDO, LENGTAIL, LENGFREE, DM1, DM2, DM3, WGHTTOP, WGHTTKDO," & _
              " WGHTTAIL, WGHTFREE, WGTOPCUT, UPWEIGHT, CHARGE, SEED, STATCLS, JDGECODE, PWTIME, ADDDPPOS, ADDDPCLS, ADDDPVAL," & _
              " ADDDPNAM, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE "
    sSqlBase = sSqlBase & "From TBCMH004"
    sSQL = sSqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sSQL = sSQL & " " & sqlWhere & " " & sqlOrder
    End If

    ' データを抽出する
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim udtRecords(0)
        DBDRV_GetTBCMH004 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ' 抽出結果を格納する
    lngRecCnt = rs.RecordCount
    ReDim udtRecords(lngRecCnt)
    For i = 1 To lngRecCnt
        With udtRecords(i)
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCODE = rs("PROCCODE")       ' 工程コード
            .LENGTOP = rs("LENGTOP")         ' 長さ（TOP）
            .LENGTKDO = rs("LENGTKDO")       ' 長さ（直胴）
            .LENGTAIL = rs("LENGTAIL")       ' 長さ（TAIL）
            .LENGFREE = rs("LENGFREE")       ' フリー長さ
            .DM1 = rs("DM1")                 ' 直胴直径１
            .DM2 = rs("DM2")                 ' 直胴直径２
            .DM3 = rs("DM3")                 ' 直胴直径３
            .WGHTTOP = rs("WGHTTOP")         ' 重量（TOP）
            .WGHTTKDO = rs("WGHTTKDO")       ' 重量（直胴）
            .WGHTTAIL = rs("WGHTTAIL")       ' 重量（TAIL)
            .WGHTFREE = rs("WGHTFREE")       ' 重量（フリー長さ）
            .WGTOPCUT = rs("WGTOPCUT")       ' トップカット重量
            .UPWEIGHT = rs("UPWEIGHT")       ' 引上げ重量
            .CHARGE = rs("CHARGE")           ' チャージ量
            .SEED = rs("SEED")               ' シード
            .STATCLS = rs("STATCLS")         ' BOT状況区分
            .JDGECODE = rs("JDGECODE")       ' 判定コード
            .PWTIME = rs("PWTIME")           ' パワー時間
            .ADDDPPOS = rs("ADDDPPOS")       ' 追加ドープ位置
            .ADDDPCLS = rs("ADDDPCLS")       ' 追加ドーパント種類
            .ADDDPVAL = rs("ADDDPVAL")       ' 追加ドープ量
            .ADDDPNAM = rs("ADDDPNAM")       ' 追加ドープ名
            .TSTAFFID = rs("TSTAFFID")       ' 登録社員ID
            .REGDATE = rs("REGDATE")         ' 登録日付
            .KSTAFFID = rs("KSTAFFID")       ' 更新社員ID
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMH004 = FUNCTION_RETURN_SUCCESS
End Function

'*********************************************************************************************
'*    関数名        : funGetSPVJisseki_J016
'*
'*    処理概要      : 1.SPV実績(TBCMJ016)の取得処理
'*
'*    パラメータ    : 変数名       ,IO ,型                               ,説明
'*                   sCryNum       ,I  ,String                           ,結晶番号
'*                   sSmplID       ,I  ,String                           ,ｻﾝﾌﾟﾙID
'*                   tSPVjisseki   ,O  ,typ_TBCMJ016                     ,結晶SPV実績(構造体)
'*                   Siyou         ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou ,WF仕様用
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************
Public Function funGetSPVJisseki_J016(sCryNum As String, sSmplID As String, udtSPVJisseki As typ_TBCMJ016 _
                                    , udtSiyou As type_DBDRV_scmzc_fcmlc001c_Siyou) As FUNCTION_RETURN
    Dim sSQL            As String
    Dim rs              As OraDynaset
    Dim sSokutei        As String       ''測定方法(測定位置＿方 + 測定位置＿点 + 測定位置＿位)
    
    ' エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function funGetSPVJisseki_J016"
    
    With udtSPVJisseki
        .MAX_FE = -2
        .MIN_FE = -2
        .AVE_FE = -2
        .CENTER_FE = -2
        .MAX_DIFF = -2
        .MIN_DIFF = -2
        .AVE_DIFF = -2
        .CENTER_DIFF = -2
    End With
            
    ' 結晶番号、ｻﾝﾌﾟﾙIDからTBCMJ016の結晶SPV実績値を検索する。
    sSQL = ""
    sSQL = sSQL & " select CRYNUM,POSITION,SMPKBN,TRANCOND,TRANCNT,HSFLG,SMPLNO,SMPLUMU" & vbLf
    sSQL = sSQL & "       ,HINBAN,REVNUM,FACTORY,OPECOND,SXLID,KRPROCCD,PROCCODE,GOUKI" & vbLf
    sSQL = sSQL & "       ,OSITEM,MAISU,SPEC,NETSU,ET,MES,DKAN" & vbLf
    sSQL = sSQL & "       ,SPV_Fe_MAX,SPV_Fe_AVE,SPV_Fe_MIN" & vbLf
    sSQL = sSQL & "       ,ms01_SPV_Fe,ms02_SPV_Fe,ms03_SPV_Fe,ms04_SPV_Fe,ms05_SPV_Fe" & vbLf
    sSQL = sSQL & "       ,ms06_SPV_Fe,ms07_SPV_Fe,ms08_SPV_Fe,ms09_SPV_Fe" & vbLf
    sSQL = sSQL & "       ,SPV_Diff_MAX,SPV_Diff_AVE,SPV_Diff_MIN" & vbLf
    sSQL = sSQL & "       ,ms01_SPV_Diff,ms02_SPV_Diff,ms03_SPV_Diff,ms04_SPV_Diff,ms05_SPV_Diff" & vbLf
    sSQL = sSQL & "       ,ms06_SPV_Diff,ms07_SPV_Diff,ms08_SPV_Diff,ms09_SPV_Diff" & vbLf
    sSQL = sSQL & "       ,TSTAFFID,REGDATE,KSTAFFID,UPDDATE,SENDFLAG,SENDDATE" & vbLf

    ' SPV判定処理追加
    '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
    sSQL = sSQL & "       ,SPV_Fe_PUA,SPV_Fe_PUAP,SPV_Fe_STD,SPV_Diff_PUA,SPV_Diff_PUAP" & vbLf
    sSQL = sSQL & "       ,SPV_Nr_MAX,SPV_Nr_AVE,SPV_Nr_STD,SPV_Nr_PUA,SPV_Nr_PUAP" & vbLf
    sSQL = sSQL & " from   TBCMJ016 " & vbLf
    sSQL = sSQL & " where  CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & " and    SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & " and    HSFLG = '1'" & vbLf
    sSQL = sSQL & " and    TRANCNT = ( select   max(TRANCNT) from TBCMJ016 " & vbLf
    sSQL = sSQL & "                    where    CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "                    and      SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "                    and      HSFLG = '1')" & vbLf
    
    ' SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' 該当ﾃﾞｰﾀなし
    If rs.EOF Then
        
        udtSPVJisseki.SMPLNO = "0"
        funGetSPVJisseki_J016 = FUNCTION_RETURN_SUCCESS
        Set rs = Nothing
        GoTo proc_exit
    End If

    With udtSPVJisseki
        If IsNull(rs("CRYNUM")) = False Then .CRYNUM = rs("CRYNUM")                                                 ' 結晶番号
        If IsNull(rs("POSITION")) = False Then .POSITION = rs("POSITION")                                           ' 位置
        If IsNull(rs("SMPKBN")) = False Then .SMPKBN = rs("SMPKBN")                                                 ' サンプル区分
        If IsNull(rs("TRANCOND")) = False Then .TRANCOND = rs("TRANCOND")                                           ' 処理条件
        If IsNull(rs("TRANCNT")) = False Then .TRANCNT = rs("TRANCNT")                                              ' 処理回数
        If IsNull(rs("HSFLG")) = False Then .HSFLG = rs("HSFLG")                                                    ' 保証フラグ
        If IsNull(rs("SMPLNO")) = False Then .SMPLNO = rs("SMPLNO")                                                 ' サンプルＮｏ
        If IsNull(rs("SMPLUMU")) = False Then .SMPLUMU = rs("SMPLUMU")                                              ' サンプル有無
        If IsNull(rs("HINBAN")) = False Then .hinban = rs("HINBAN")                                                 ' 品番
        If IsNull(rs("REVNUM")) = False Then .REVNUM = rs("REVNUM")                                                 ' 製品番号改訂番号
        If IsNull(rs("FACTORY")) = False Then .factory = rs("FACTORY")                                              ' 工場
        If IsNull(rs("OPECOND")) = False Then .opecond = rs("OPECOND")                                              ' 操業条件
        If IsNull(rs("SXLID")) = False Then .SXLID = rs("SXLID")                                                    ' SXLID
        If IsNull(rs("KRPROCCD")) = False Then .KRPROCCD = rs("KRPROCCD")                                           ' 管理工程コード
        If IsNull(rs("PROCCODE")) = False Then .PROCCODE = rs("PROCCODE")                                           ' 工程コード
        If IsNull(rs("GOUKI")) = False Then .GOUKI = rs("GOUKI")                                                    ' 号機
        If IsNull(rs("OSITEM")) = False Then .OSITEM = rs("OSITEM")                                                 ' 評価項目
        If IsNull(rs("MAISU")) = False Then .MAISU = rs("MAISU")                                                    ' 評価枚数
        If IsNull(rs("SPEC")) = False Then .Spec = rs("SPEC")                                                       ' 規格値
        If IsNull(rs("NETSU")) = False Then .NETSU = rs("NETSU")                                                    ' 熱処理条件
        If IsNull(rs("ET")) = False Then .ET = rs("ET")                                                             ' エッチング条件
        If IsNull(rs("MES")) = False Then .MES = rs("MES")                                                          ' 計測方法
        If IsNull(rs("DKAN")) = False Then .DKAN = rs("DKAN")                                                       ' ＤＫアニール条件
    
        If IsNull(rs("SPV_Fe_MAX")) = False Then .SPV_Fe_MAX = rs("SPV_Fe_MAX") Else .SPV_Fe_MAX = -1               ' SPV_Fe_MAX
        If IsNull(rs("SPV_Fe_AVE")) = False Then .SPV_Fe_AVE = rs("SPV_Fe_AVE") Else .SPV_Fe_AVE = -1               ' SPV_Fe_AVE
        If IsNull(rs("SPV_Fe_MIN")) = False Then .SPV_Fe_MIN = rs("SPV_Fe_MIN") Else .SPV_Fe_MIN = -1               ' SPV_Fe_MIN
        If IsNull(rs("ms01_SPV_Fe")) = False Then .ms01_SPV_Fe = rs("ms01_SPV_Fe") Else .ms01_SPV_Fe = -1           ' 測定値01 SPV_Fe
        If IsNull(rs("ms02_SPV_Fe")) = False Then .ms02_SPV_Fe = rs("ms02_SPV_Fe") Else .ms02_SPV_Fe = -1           ' 測定値02 SPV_Fe
        If IsNull(rs("ms03_SPV_Fe")) = False Then .ms03_SPV_Fe = rs("ms03_SPV_Fe") Else .ms03_SPV_Fe = -1           ' 測定値03 SPV_Fe
        If IsNull(rs("ms04_SPV_Fe")) = False Then .ms04_SPV_Fe = rs("ms04_SPV_Fe") Else .ms04_SPV_Fe = -1           ' 測定値04 SPV_Fe
        If IsNull(rs("ms05_SPV_Fe")) = False Then .ms05_SPV_Fe = rs("ms05_SPV_Fe") Else .ms05_SPV_Fe = -1           ' 測定値05 SPV_Fe
        If IsNull(rs("ms06_SPV_Fe")) = False Then .ms06_SPV_Fe = rs("ms06_SPV_Fe") Else .ms06_SPV_Fe = -1           ' 測定値06 SPV_Fe
        If IsNull(rs("ms07_SPV_Fe")) = False Then .ms07_SPV_Fe = rs("ms07_SPV_Fe") Else .ms07_SPV_Fe = -1           ' 測定値07 SPV_Fe
        If IsNull(rs("ms08_SPV_Fe")) = False Then .ms08_SPV_Fe = rs("ms08_SPV_Fe") Else .ms08_SPV_Fe = -1           ' 測定値08 SPV_Fe
        If IsNull(rs("ms09_SPV_Fe")) = False Then .ms09_SPV_Fe = rs("ms09_SPV_Fe") Else .ms09_SPV_Fe = -1           ' 測定値09 SPV_Fe
        If IsNull(rs("SPV_Diff_MAX")) = False Then .SPV_Diff_MAX = rs("SPV_Diff_MAX") Else .SPV_Diff_MAX = -1       ' SPV_拡散長_MAX
        If IsNull(rs("SPV_Diff_AVE")) = False Then .SPV_Diff_AVE = rs("SPV_Diff_AVE") Else .SPV_Diff_AVE = -1       ' SPV_拡散長_AVE
        If IsNull(rs("SPV_Diff_MIN")) = False Then .SPV_Diff_MIN = rs("SPV_Diff_MIN") Else .SPV_Diff_MIN = -1       ' SPV_拡散長_MIN
        If IsNull(rs("ms01_SPV_Diff")) = False Then .ms01_SPV_Diff = rs("ms01_SPV_Diff") Else .ms01_SPV_Diff = -1   ' 測定値01 SPV_拡散長
        If IsNull(rs("ms02_SPV_Diff")) = False Then .ms02_SPV_Diff = rs("ms02_SPV_Diff") Else .ms02_SPV_Diff = -1   ' 測定値02 SPV_拡散長
        If IsNull(rs("ms03_SPV_Diff")) = False Then .ms03_SPV_Diff = rs("ms03_SPV_Diff") Else .ms03_SPV_Diff = -1   ' 測定値03 SPV_拡散長
        If IsNull(rs("ms04_SPV_Diff")) = False Then .ms04_SPV_Diff = rs("ms04_SPV_Diff") Else .ms04_SPV_Diff = -1   ' 測定値04 SPV_拡散長
        If IsNull(rs("ms05_SPV_Diff")) = False Then .ms05_SPV_Diff = rs("ms05_SPV_Diff") Else .ms05_SPV_Diff = -1   ' 測定値05 SPV_拡散長
        If IsNull(rs("ms06_SPV_Diff")) = False Then .ms06_SPV_Diff = rs("ms06_SPV_Diff") Else .ms06_SPV_Diff = -1   ' 測定値06 SPV_拡散長
        If IsNull(rs("ms07_SPV_Diff")) = False Then .ms07_SPV_Diff = rs("ms07_SPV_Diff") Else .ms07_SPV_Diff = -1   ' 測定値07 SPV_拡散長
        If IsNull(rs("ms08_SPV_Diff")) = False Then .ms08_SPV_Diff = rs("ms08_SPV_Diff") Else .ms08_SPV_Diff = -1   ' 測定値08 SPV_拡散長
        If IsNull(rs("ms09_SPV_Diff")) = False Then .ms09_SPV_Diff = rs("ms09_SPV_Diff") Else .ms09_SPV_Diff = -1   ' 測定値09 SPV_拡散長
        
        If IsNull(rs("TSTAFFID")) = False Then .TSTAFFID = rs("TSTAFFID")                                           ' 登録社員ID
        If IsNull(rs("REGDATE")) = False Then .REGDATE = rs("REGDATE")                                              ' 登録日付
        If IsNull(rs("KSTAFFID")) = False Then .KSTAFFID = rs("KSTAFFID")                                           ' 更新社員ID
        If IsNull(rs("UPDDATE")) = False Then .UPDDATE = rs("UPDDATE")                                              ' 更新日付
        If IsNull(rs("SENDFLAG")) = False Then .SENDFLAG = rs("SENDFLAG")                                           ' 送信フラグ
        If IsNull(rs("SENDDATE")) = False Then .SENDDATE = rs("SENDDATE")                                           ' 送信日付

        ' SPV判定処理追加
        ' 項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
        If IsNull(rs("SPV_Fe_PUA")) = False Then .SPV_Fe_PUA = rs("SPV_Fe_PUA") Else .SPV_Fe_PUA = -1               ' SPV_Fe PUA値
        If IsNull(rs("SPV_Fe_PUAP")) = False Then .SPV_Fe_PUAP = rs("SPV_Fe_PUAP") Else .SPV_Fe_PUAP = -1           ' SPV_Fe PUA％値
        If IsNull(rs("SPV_Fe_STD")) = False Then .SPV_Fe_STD = rs("SPV_Fe_STD") Else .SPV_Fe_STD = -1               ' SPV_Fe STD
        If IsNull(rs("SPV_Diff_PUA")) = False Then .SPV_Diff_PUA = rs("SPV_Diff_PUA") Else .SPV_Diff_PUA = -1       ' SPV_拡散長 PUA値
        If IsNull(rs("SPV_Diff_PUAP")) = False Then .SPV_Diff_PUAP = rs("SPV_Diff_PUAP") Else .SPV_Diff_PUAP = -1   ' SPV_拡散長 PUA％値
        If IsNull(rs("SPV_Nr_MAX")) = False Then .SPV_Nr_MAX = rs("SPV_Nr_MAX") Else .SPV_Nr_MAX = -1               ' SPV_OtherRecords_MAX
        If IsNull(rs("SPV_Nr_AVE")) = False Then .SPV_Nr_AVE = rs("SPV_Nr_AVE") Else .SPV_Nr_AVE = -1               ' SPV_OtherRecords_AVE
        If IsNull(rs("SPV_Nr_STD")) = False Then .SPV_Nr_STD = rs("SPV_Nr_STD") Else .SPV_Nr_STD = -1               ' SPV_OtherRecords_STD
        If IsNull(rs("SPV_Nr_PUA")) = False Then .SPV_Nr_PUA = rs("SPV_Nr_PUA") Else .SPV_Nr_PUA = -1               ' SPV_OtherRecords_PUA値
        If IsNull(rs("SPV_Nr_PUAP")) = False Then .SPV_Nr_PUAP = rs("SPV_Nr_PUAP") Else .SPV_Nr_PUAP = -1           ' SPV_OtherRecords_PUA％値

        ' Fe濃度測定方法
        sSokutei = Trim(udtSiyou.HWFSPVSH) & Trim(udtSiyou.HWFSPVST) & Trim(udtSiyou.HWFSPVSI)
    
        ' MAP測定の場合
        If sSokutei = "AMX" Then
            .MAX_FE = .SPV_Fe_MAX
            ' SPV判定処理追加
            ' Map測定(AMX)の場合は、表示データ2(MIN)を表示しないように修正
            .MIN_FE = -1
            .AVE_FE = .SPV_Fe_AVE
            .CENTER_FE = -1

            ' SPV判定処理追加
            .PUA_FE = .SPV_Fe_PUA
            .PUAP_FE = .SPV_Fe_PUAP
            .STD_FE = .SPV_Fe_STD
        ' 9点測定の場合
        ElseIf sSokutei = "V9T" Then
            ' Fe濃度のMAX,MIN,AVEを取得
            If funGetSPVJisseki_J016_Fe(.CRYNUM, .SMPLNO, .TRANCNT, _
                                    udtSPVJisseki) = FUNCTION_RETURN_FAILURE Then
                funGetSPVJisseki_J016 = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
            
            .CENTER_FE = .ms01_SPV_Fe
            
            ' SPV判定処理追加
            .PUA_FE = -1
            .PUAP_FE = -1
            .STD_FE = -1
        Else
            .MAX_FE = -1
            .MIN_FE = -1
            .AVE_FE = -1
            .CENTER_FE = -1
            
            'SPV判定処理追加
            .PUA_FE = -1
            .PUAP_FE = -1
            .STD_FE = -1
        End If
    
        ' 拡散長測定方法
        sSokutei = Trim(udtSiyou.HWFDLSPH) & Trim(udtSiyou.HWFDLSPT) & Trim(udtSiyou.HWFDLSPI)
    
        ' MAP測定の場合
        If sSokutei = "AMX" Then
            .MAX_DIFF = .SPV_Diff_MAX
            .MIN_DIFF = .SPV_Diff_MIN
            .AVE_DIFF = .SPV_Diff_AVE
            .CENTER_DIFF = -1
            
            ' SPV判定処理追加
            .PUA_DIFF = .SPV_Diff_PUA
            .PUAP_DIFF = .SPV_Diff_PUAP
        
        ' 9点測定の場合
        ElseIf sSokutei = "V9T" Then
            ' 拡散長のMAX,MIN,AVEを取得
            If funGetSPVJisseki_J016_Diff(.CRYNUM, .SMPLNO, .TRANCNT, _
                                    udtSPVJisseki) = FUNCTION_RETURN_FAILURE Then
                funGetSPVJisseki_J016 = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
            
            .CENTER_DIFF = .ms01_SPV_Diff
            
            ' SPV判定処理追加
            .PUA_DIFF = -1
            .PUAP_DIFF = -1
        Else
            .MAX_DIFF = -1
            .MIN_DIFF = -1
            .AVE_DIFF = -1
            .CENTER_DIFF = -1
            
            ' SPV判定処理追加
            .PUA_DIFF = -1
            .PUAP_DIFF = -1
        End If
        
        ' SPV判定処理追加
        ' Nr濃度測定方法
        sSokutei = Trim(udtSiyou.HWFNRSH) & Trim(udtSiyou.HWFNRST) & Trim(udtSiyou.HWFNRSI)
        
        ' MAP測定の場合
        If sSokutei = "AMX" Then
            .MAX_NR = .SPV_Nr_MAX
            .MIN_NR = -1
            .AVE_NR = .SPV_Nr_AVE
            .CENTER_NR = -1
            .PUA_NR = .SPV_Nr_PUA
            .PUAP_NR = .SPV_Nr_PUAP
            .STD_NR = .SPV_Nr_STD
        
        ' 9点測定の場合
        ElseIf sSokutei = "V9T" Then
            .MAX_NR = -1
            .MIN_NR = -1
            .AVE_NR = -1
            .CENTER_NR = -1
            .PUA_NR = -1
            .PUAP_NR = -1
            .STD_NR = -1
        Else
            .MAX_NR = -1
            .MIN_NR = -1
            .AVE_NR = -1
            .CENTER_NR = -1
            .PUA_NR = -1
            .PUAP_NR = -1
            .STD_NR = -1
        End If
    End With
    
    Set rs = Nothing

    funGetSPVJisseki_J016 = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' 終了
    Exit Function

proc_err:
    ' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    Set rs = Nothing
    funGetSPVJisseki_J016 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*********************************************************************************************
'*    関数名        : funGetSPVJisseki_J016_Fe
'*
'*    処理概要      : 1.SPV実績(TBCMJ016)のFe濃度9点測定値のMAX・MIN・AVEを取得する
'*
'*    パラメータ    : 変数名       ,IO ,型                               ,説明
'*                   sCryNum       ,I  ,String                           ,結晶番号
'*                   sSmplID       ,I  ,String                           ,ｻﾝﾌﾟﾙID
'*                   intTrancnt    ,I  ,Integer                          ,処理回数
'*                   udtSPVJisseki ,O  ,typ_TBCMJ016                     ,結晶SPV実績(構造体)
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************
Public Function funGetSPVJisseki_J016_Fe(sCryNum As String, sSmplID As String, intTrancnt As Integer, _
                                        udtSPVJisseki As typ_TBCMJ016) As FUNCTION_RETURN
    Dim sSQL            As String
    Dim rs              As OraDynaset
    
    ' エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function funGetSPVJisseki_J016_Fe"
            
    sSQL = ""
    sSQL = sSQL & " SELECT  MAX(SPV_FE) AS MAX_FE,MIN(SPV_FE) AS MIN_FE,AVG(SPV_FE) AS AVE_FE" & vbLf
    sSQL = sSQL & " FROM   (SELECT  CRYNUM,SMPLNO,TRANCNT,ms01_SPV_FE AS SPV_FE" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms02_SPV_FE AS SPV_FE" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms03_SPV_FE AS SPV_FE" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms04_SPV_FE AS SPV_FE" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms05_SPV_FE AS SPV_FE" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms06_SPV_FE AS SPV_FE" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms07_SPV_FE AS SPV_FE" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms08_SPV_FE AS SPV_FE" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms09_SPV_FE AS SPV_FE" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "        )" & vbLf
    
    ' SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' 該当ﾃﾞｰﾀなし
    If rs.EOF Then
        funGetSPVJisseki_J016_Fe = FUNCTION_RETURN_SUCCESS
    
        With udtSPVJisseki
            .MAX_FE = -1
            .MIN_FE = -1
            .AVE_FE = -1
        End With
        
        Set rs = Nothing
        GoTo proc_exit
    End If

    With udtSPVJisseki
        If IsNull(rs("MAX_FE")) = False Then .MAX_FE = rs("MAX_FE") Else .MAX_FE = -1
        If IsNull(rs("MIN_FE")) = False Then .MIN_FE = rs("MIN_FE") Else .MIN_FE = -1
        If IsNull(rs("AVE_FE")) = False Then .AVE_FE = rs("AVE_FE") Else .AVE_FE = -1
    End With
    
    Set rs = Nothing
    
    funGetSPVJisseki_J016_Fe = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' 終了
    Exit Function

proc_err:
    ' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    Set rs = Nothing
    funGetSPVJisseki_J016_Fe = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*********************************************************************************************
'*    関数名        : funGetSPVJisseki_J016_Diff
'*
'*    処理概要      : 1.SPV実績(TBCMJ016)の拡散長9点測定値のMAX・MIN・AVEを取得する
'*
'*    パラメータ    : 変数名       ,IO ,型                               ,説明
'*                   sCryNum       ,I  ,String                           ,結晶番号
'*                   sSmplID       ,I  ,String                           ,ｻﾝﾌﾟﾙID
'*                   intTrancnt    ,I  ,Integer                          ,処理回数
'*                   udtSPVJisseki ,O  ,typ_TBCMJ016                     ,結晶SPV実績(構造体)
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************
Public Function funGetSPVJisseki_J016_Diff(sCryNum As String, sSmplID As String, intTrancnt As Integer, _
                                        udtSPVJisseki As typ_TBCMJ016) As FUNCTION_RETURN
    Dim sSQL            As String
    Dim rs              As OraDynaset
    
    ' エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function funGetSPVJisseki_J016_Diff"
            
    sSQL = ""
    sSQL = sSQL & " SELECT  MAX(SPV_DIFF) AS MAX_DIFF,MIN(SPV_DIFF) AS MIN_DIFF,AVG(SPV_DIFF) AS AVE_DIFF" & vbLf
    sSQL = sSQL & " FROM   (SELECT  CRYNUM,SMPLNO,TRANCNT,ms01_SPV_DIFF AS SPV_DIFF" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms02_SPV_DIFF AS SPV_DIFF" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms03_SPV_DIFF AS SPV_DIFF" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms04_SPV_DIFF AS SPV_DIFF" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms05_SPV_DIFF AS SPV_DIFF" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms06_SPV_DIFF AS SPV_DIFF" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms07_SPV_DIFF AS SPV_DIFF" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms08_SPV_DIFF AS SPV_DIFF" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms09_SPV_DIFF AS SPV_DIFF" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "        )" & vbLf
    
    ' SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' 該当ﾃﾞｰﾀなし
    If rs.EOF Then
        funGetSPVJisseki_J016_Diff = FUNCTION_RETURN_SUCCESS
    
        With udtSPVJisseki
            .MAX_DIFF = -1
            .MIN_DIFF = -1
            .AVE_DIFF = -1
        End With
        
        Set rs = Nothing
        GoTo proc_exit
    End If

    With udtSPVJisseki
        If IsNull(rs("MAX_DIFF")) = False Then .MAX_DIFF = rs("MAX_DIFF") Else .MAX_DIFF = -1
        If IsNull(rs("MIN_DIFF")) = False Then .MIN_DIFF = rs("MIN_DIFF") Else .MIN_DIFF = -1
        If IsNull(rs("AVE_DIFF")) = False Then .AVE_DIFF = rs("AVE_DIFF") Else .AVE_DIFF = -1
    End With
    
    Set rs = Nothing
    
    funGetSPVJisseki_J016_Diff = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' 終了
    Exit Function

proc_err:
    ' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    Set rs = Nothing
    funGetSPVJisseki_J016_Diff = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*********************************************************************************************
'*    関数名        : funGetSiyou_Warp
'*
'*    処理概要      : 1.Warp仕様値の取得処理
'*
'*    パラメータ    : 変数名    ,IO ,型           ,説明
'*                   udtHIN     ,I  ,tFullHinban  ,品番
'*                   dblWarpMax ,I  ,Double       ,Warp上限
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************
Public Function funGetSiyou_Warp(udtHin As tFullHinban, dblWarpMax As Double) As FUNCTION_RETURN

    Dim sSQL    As String           ' SQL全体
    Dim rs      As OraDynaset       ' RecordSet

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function funGetSiyou_Warp"

    sSQL = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
    sSQL = sSQL & "HWFWARMX "
    sSQL = sSQL & "from TBCME027 "
    sSQL = sSQL & "Where HINBAN = '" & udtHin.hinban & "' and "
    sSQL = sSQL & "      MNOREVNO = " & udtHin.mnorevno & " and "
    sSQL = sSQL & "      FACTORY = '" & udtHin.factory & "' and "
    sSQL = sSQL & "      OPECOND = '" & udtHin.opecond & "'"
    
    ' データを抽出する
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGetSiyou_Warp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ' 抽出結果を格納する
    dblWarpMax = fncNullCheck(rs("HWFWARMX"))         ' 品WFWARP上限
        
    Set rs = Nothing

    funGetSiyou_Warp = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' 終了
    Exit Function

proc_err:
    ' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    funGetSiyou_Warp = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*********************************************************************************************
'*    関数名        : funGetSiyou_Kaku
'*
'*    処理概要      : 1.合成角度仕様値の取得処理
'*
'*    パラメータ    : 変数名    ,IO ,型           ,説明
'*                   udtHIN     ,I  ,tFullHinban  ,品番
'*                   dblKakuMin ,I  ,Double       ,結晶面傾下限
'*                   dblWarpMax ,I  ,Double       ,結晶面傾上限
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************
Public Function funGetSiyou_Kaku(udtHin As tFullHinban, dblKakuMin As Double, dblKakuMax As Double) As FUNCTION_RETURN
    Dim sSQL    As String           ' SQL全体
    Dim rs      As OraDynaset       ' RecordSet

    ' エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function funGetSiyou_Kaku"

    sSQL = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
    sSQL = sSQL & "HWFCSMIN, HWFCSMAX "
    sSQL = sSQL & "from TBCME022 "
    sSQL = sSQL & "Where HINBAN = '" & udtHin.hinban & "' and "
    sSQL = sSQL & "      MNOREVNO = " & udtHin.mnorevno & " and "
    sSQL = sSQL & "      FACTORY = '" & udtHin.factory & "' and "
    sSQL = sSQL & "      OPECOND = '" & udtHin.opecond & "'"
    
    ' データを抽出する
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGetSiyou_Kaku = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ' 抽出結果を格納する
    dblKakuMin = fncNullCheck(rs("HWFCSMIN"))         ' 品WF結晶面傾下限
    dblKakuMax = fncNullCheck(rs("HWFCSMAX"))         ' 品WF結晶面傾上限
    
    Set rs = Nothing

    funGetSiyou_Kaku = FUNCTION_RETURN_SUCCESS
  
proc_exit:
    ' 終了
'    gErr.Pop
    Exit Function

proc_err:
    ' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    funGetSiyou_Kaku = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*********************************************************************************************
'*    関数名        : funGet_TBCMY018
'*
'*    処理概要      : 1.標準測定ﾃﾞｰﾀ(TBCMY018)の取得処理
'*
'*    パラメータ    : 変数名    ,IO ,型                  ,説明
'*                   sBlockID   ,I  ,String              ,ﾌﾞﾛｯｸID
'*                   sMeasItem  ,I  ,String              ,測定項目名
'*                   udtMEAS()  ,O  ,typ_WarpKakuData    ,測定ﾃﾞｰﾀ
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************
Public Function funGet_TBCMY018(sBlockId As String, sMeasItem As String, udtMEAS() As typ_WarpKakuData) As FUNCTION_RETURN
    Dim sSQL        As String           ' SQL全体
    Dim rs          As OraDynaset       ' RecordSet
    Dim lngRecCnt   As Long             ' ﾚｺｰﾄﾞ数
    Dim i           As Integer

    ' エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function funGet_TBCMY018"

    sSQL = "select SUBLOTID, MEASITEM, WAFID, MEASDATA "
    sSQL = sSQL & "from TBCMY018 Y018 "
    sSQL = sSQL & "Where SUBLOTID = '" & sBlockId & "' and "
    sSQL = sSQL & "      MEASITEM like '%" & sMeasItem & "%' and "
    sSQL = sSQL & "      TRANCNT = (select MAX(TRANCNT) from TBCMY018 "
    sSQL = sSQL & "                 where SUBLOTID = Y018.SUBLOTID "
    sSQL = sSQL & "                 and WAFID = Y018.WAFID) "
    sSQL = sSQL & "order by WAFID"
    
    ' データを抽出する
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    If rs Is Nothing Then
        ReDim udtMEAS(0)
        rs.Close
        funGet_TBCMY018 = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If
    
    lngRecCnt = rs.RecordCount
    ReDim udtMEAS(lngRecCnt)
    
    ' 抽出結果を格納する
    For i = 1 To lngRecCnt
        ' ﾌﾞﾛｯｸID
        udtMEAS(i).BLOCKID = sBlockId
        
        ' ｳｪﾊｰID
        If IsNull(rs("WAFID")) Then
            udtMEAS(i).WAFID = -1
        ElseIf Not IsNumeric(rs("WAFID")) Then
            udtMEAS(i).WAFID = -1
        Else
            udtMEAS(i).WAFID = CDbl(rs("WAFID"))
        End If
        
        ' 測定値
        If IsNull(rs("MEASDATA")) Then
            udtMEAS(i).MEASDATA = -1
        ElseIf Not IsNumeric(rs("MEASDATA")) Then
            udtMEAS(i).MEASDATA = -1
        Else
            udtMEAS(i).MEASDATA = CDbl(rs("MEASDATA"))
        End If
        
        rs.MoveNext
    Next i
    rs.Close

    funGet_TBCMY018 = FUNCTION_RETURN_SUCCESS
  
proc_exit:
    ' 終了
    Exit Function

proc_err:
    ' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    funGet_TBCMY018 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*********************************************************************************************
'*    関数名        : funWfcGetDataEtc_SPV
'*
'*    処理概要      : 1.WF総合判定 各種データ取得(SPV用)
'*
'*    パラメータ    : 変数名      ,IO  ,型                                    ,説明
'*                   udtNew_Hinban,I   ,tFullHinban                           ,品番情報
'*                   Siyou        ,O   ,type_DBDRV_scmzc_fcmlc001c_Siyou_SPV  ,WF仕様用
'*                   sErrMsg 　　 ,O   ,String    　　　　　　　　　　　    　,エラーメッセージ
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************
Public Function funWfcGetDataEtc_SPV(udtNew_Hinban As tFullHinban, _
                                 udtSiyou As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                                 Optional sErrMsg As String = vbNullString) As FUNCTION_RETURN
    Dim sSQL    As String
    Dim rs      As OraDynaset
    Dim sDBName As String

    ' エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function funWfcGetDataEtc_SPV"

    funWfcGetDataEtc_SPV = FUNCTION_RETURN_SUCCESS

    ' WF仕様取得

    sDBName = "E048"
    sSQL = "select "
    sSQL = sSQL & "HWFSPVPUG,"        ' 品ＷＦＳＰＶＦＥＰＵＡ限
    sSQL = sSQL & "HWFSPVPUR,"        ' 品ＷＦＳＰＶＦＥＰＵＡ率
    sSQL = sSQL & "HWFSPVSTD,"        ' 品ＷＦＳＰＶＦＥ標準偏差
    sSQL = sSQL & "HWFDLPUG,"         ' 品ＷＦ拡散長ＰＵＡ限
    sSQL = sSQL & "HWFDLPUR,"         ' 品ＷＦ拡散長ＰＵＡ率
    sSQL = sSQL & "HWFNRMX,"          ' 品ＷＦＳＰＶＮＲ上限
    sSQL = sSQL & "HWFNRAM,"          ' 品ＷＦＳＰＶＮＲ平均
    sSQL = sSQL & "HWFNRPUG,"         ' 品ＷＦＳＰＶＮＲＰＵＡ限
    sSQL = sSQL & "HWFNRPUR,"         ' 品ＷＦＳＰＶＮＲＰＵＡ率
    sSQL = sSQL & "HWFNRSTD,"         ' 品ＷＦＳＰＶＮＲ標準偏差
    sSQL = sSQL & "HWFNRKN,"          ' 品ＷＦＳＰＶＮＲ検査頻度＿抜
    sSQL = sSQL & "HWFNRHS,"          ' 品ＷＦＳＰＶＮＲ保証方法＿処
    sSQL = sSQL & "HWFNRSH,"          ' 品ＷＦＳＰＶＮＲ測定位置＿方
    sSQL = sSQL & "HWFNRST,"          ' 品ＷＦＳＰＶＮＲ測定位置＿点
    sSQL = sSQL & "HWFNRHT,"          ' 品ＷＦＳＰＶＮＲ保証方法＿対
    sSQL = sSQL & "HWFNRSI "          ' 品ＷＦＳＰＶＮＲ測定位置＿位
    sSQL = sSQL & "from TBCME048 "
    sSQL = sSQL & "where HINBAN = '" & udtNew_Hinban.hinban & "' "
    sSQL = sSQL & "and MNOREVNO = " & udtNew_Hinban.mnorevno & " "
    sSQL = sSQL & "and FACTORY = '" & udtNew_Hinban.factory & "' "
    sSQL = sSQL & "and OPECOND = '" & udtNew_Hinban.opecond & "' "

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sErrMsg = GetMsgStr("EGET2", sDBName)
        funWfcGetDataEtc_SPV = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    With udtSiyou
        ' 品ＷＦＳＰＶＦＥＰＵＡ限
        If IsNull(rs("HWFSPVPUG")) = False Then .HWFSPVPUG = rs("HWFSPVPUG") Else .HWFSPVPUG = -1
        
        ' 品ＷＦＳＰＶＦＥＰＵＡ率
        If IsNull(rs("HWFSPVPUR")) = False Then .HWFSPVPUR = rs("HWFSPVPUR") Else .HWFSPVPUR = -1
        
        ' 品ＷＦＳＰＶＦＥ標準偏差
        If IsNull(rs("HWFSPVSTD")) = False Then .HWFSPVSTD = rs("HWFSPVSTD") Else .HWFSPVSTD = -1
        
        ' 品ＷＦＳＰＶＮＲ上限
        .HWFNRMX = fncNullCheck(rs("HWFNRMX"))
        
        ' 品ＷＦＳＰＶＮＲＰＵＡ限
        If IsNull(rs("HWFNRPUG")) = False Then .HWFNRPUG = rs("HWFNRPUG") Else .HWFNRPUG = -1
        
        ' 品ＷＦＳＰＶＮＲＰＵＡ率
        If IsNull(rs("HWFNRPUR")) = False Then .HWFNRPUR = rs("HWFNRPUR") Else .HWFNRPUR = -1
        
        ' 品ＷＦＳＰＶＮＲ標準偏差
        If IsNull(rs("HWFNRSTD")) = False Then .HWFNRSTD = rs("HWFNRSTD") Else .HWFNRSTD = -1
        
        ' 品ＷＦ拡散長ＰＵＡ限
        If IsNull(rs("HWFDLPUG")) = False Then .HWFDLPUG = rs("HWFDLPUG") Else .HWFDLPUG = -1
        
        ' 品ＷＦ拡散長ＰＵＡ率
        If IsNull(rs("HWFDLPUR")) = False Then .HWFDLPUR = rs("HWFDLPUR") Else .HWFDLPUR = -1
        
        ' 品ＷＦＳＰＶＮＲ平均
        .HWFNRAM = fncNullCheck(rs("HWFNRAM"))
        
        ' 品ＷＦＳＰＶＮＲ検査頻度＿抜
        If IsNull(rs("HWFNRKN")) = False Then .HWFNRKN = rs("HWFNRKN") Else .HWFNRKN = vbNullString
        
        ' 品ＷＦＳＰＶＮＲ保証方法＿処
        If IsNull(rs("HWFNRHS")) = False Then .HWFNRHS = rs("HWFNRHS") Else .HWFNRHS = vbNullString
        
        ' 品ＷＦＳＰＶＮＲ測定位置_方
        If IsNull(rs("HWFNRSH")) = False Then .HWFNRSH = rs("HWFNRSH") Else .HWFNRSH = vbNullString
        
        ' 品ＷＦＳＰＶＮＲ測定位置_点
        If IsNull(rs("HWFNRST")) = False Then .HWFNRST = rs("HWFNRST") Else .HWFNRST = vbNullString
        
        ' 品ＷＦＳＰＶＮＲ測定位置_処
        If IsNull(rs("HWFNRHT")) = False Then .HWFNRHT = rs("HWFNRHT") Else .HWFNRHT = vbNullString
        
        ' 品ＷＦＳＰＶＮＲ測定位置_位
        If IsNull(rs("HWFNRSI")) = False Then .HWFNRSI = rs("HWFNRSI") Else .HWFNRSI = vbNullString
    End With
    rs.Close

    Set rs = Nothing

proc_exit:
    ' 終了
    Exit Function

proc_err:
    ' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    funWfcGetDataEtc_SPV = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function
