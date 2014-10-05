Attribute VB_Name = "s_cmbc039_SQL"
Option Explicit

'' WFセンター総合判定待ち一覧

' SXL管理
Public Type DBDRV_scmzc_fcmlc001b_SXL039
    CRYNUMCA As String * 12        ' 結晶番号
    INPOSCA As Integer             ' 結晶内開始位置
    GNLCA As Integer               ' 現在長さ
    SXLIDCA As String * 13         ' SXLID
    GNKKNTCA As String * 5         ' 現在管理工程(未使用)
    GNWKNTCA As String * 5         ' 現在工程
    NOWPROC As String * 5          ' 現在工程
    NEKKNTCA As String * 5         ' 最終通過管理工程(未使用)
    NEWKNTCA As String * 5         ' 最終通過工程
    SAKJCA As String * 1           ' 削除区分
    LSTATBCA As String * 1         ' 最終状態区分
    HOLDBCA As String * 1          ' ホールド区分
    HINBCA As String * 8           ' 品番
    REVNUMCA As Integer            ' 製品番号改訂番号
    FACTORYCA As String * 1        ' 工場
    OPECA As String * 1            ' 操業条件
    MAICB As Integer               ' 枚数
    TDAYCB As Date                 ' 登録日付
    KDAYCA As Date                 ' 更新日付
    HOLDBCB As String * 1          ' ﾎｰﾙﾄﾞ区分　06/02/08 ooba
    WFHOLDFLGCB As String * 1      ' WFﾎｰﾙﾄﾞ区分　06/02/08 ooba
    KETURAKU As Boolean            ' 欠落情報有無フラグ
    WFSMP() As typ_XSDCW           ' サンプル管理（TOP、TAIL順 ２レコード）
    PLANTCAT As String             ' 向先 07/09/04 SPK Tsutsumi Add
    KANREN As String * 1            ' 関連ﾌﾞﾛｯｸ有無　08/01/31 ooba
    AGRSTATUS  As String            ' 承認確認区分 add SETkimizuka
    STOP    As String               ' 停止 add SETkimizuka
    CAUSE   As String               ' 停止理由 add SETkimizuka
    PRINTNO As String               ' 先行評価 add SETkimizuka
End Type

'WFセンター総合判定

'入力用
Public Type type_DBDRV_scmzc_fcmlc001c_In039
    HIN As tFullHinban             ' 品番(full)
    SAMPLEID As String * 16        ' サンプルID
    SXLID As String * 13           ' SXLID
End Type

'WF製品仕様取得用
Public Type type_DBDRV_scmzc_fcmlc001c_Siyou039
    HWFTYPE As String * 1          ' 品ＷＦタイプ
    HWFCDIR As String * 1          ' 品ＷＦ結晶面方
    HWFCDOP As String * 1          ' 品ＷＦ結晶ドープ

    HWFRMIN As Double              ' 品ＷＦ比抵抗下限
    HWFRMAX As Double              ' 品ＷＦ比抵抗上限
    HWFRSPOH As String * 1         ' 品ＷＦ比抵抗測定位置＿方
    HWFRSPOT As String * 1         ' 品ＷＦ比抵抗測定位置＿点
    HWFRSPOI As String * 1         ' 品ＷＦ比抵抗測定位置＿位
    HWFRHWYT As String * 1         ' 品ＷＦ比抵抗保証方法＿対
    HWFRHWYS As String * 1         ' 品ＷＦ比抵抗保証方法＿処
    HWFRMCAL As String * 1         ' 品ＷＦ比抵抗面内計算 2001/11/08 S.Sano
    HWFRAMIN As Double             ' 品ＷＦ比抵抗平均下限
    HWFRAMAX As Double             ' 品ＷＦ比抵抗平均上限
    HWFRMBNP As Double             ' 品ＷＦ比抵抗面内分布

    HWFMKMIN As Double             ' 品ＷＦ無欠陥層下限
    HWFMKMAX As Double             ' 品ＷＦ無欠陥層上限
    HWFMKSPH As String * 1         ' 品ＷＦ無欠陥層測定位置＿方
    HWFMKSPT As String * 1         ' 品ＷＦ無欠陥層測定位置＿点
    HWFMKSPR As String * 1         ' 品ＷＦ無欠陥層測定位置＿領
    HWFMKHWT As String * 1         ' 品ＷＦ無欠陥層保証方法＿対
    HWFMKHWS As String * 1         ' 品ＷＦ無欠陥層保証方法＿処

    HWFONMIN As Double             ' 品ＷＦ酸素濃度下限
    HWFONMAX As Double             ' 品ＷＦ酸素濃度上限
    HWFONSPH As String * 1         ' 品ＷＦ酸素濃度測定位置＿方
    HWFONSPT As String * 1         ' 品ＷＦ酸素濃度測定位置＿点
    HWFONSPI As String * 1         ' 品ＷＦ酸素濃度測定位置＿位
    HWFONHWT As String * 1         ' 品ＷＦ酸素濃度保証方法＿対
    HWFONHWS As String * 1         ' 品ＷＦ酸素濃度保証方法＿処
    HWFONMCL As String * 1         ' 品ＷＦ酸素濃度面内計算 2001/11/08 S.Sano
    HWFONMBP As Double             ' 品ＷＦ酸素濃度面内分布
    HWFONAMN As Double             ' 品ＷＦ酸素濃度平均下限
    HWFONAMX As Double             ' 品ＷＦ酸素濃度平均上限

    HWFOS1MN As Double             ' 品ＷＦ酸素析出１下限
    HWFOS1MX As Double             ' 品ＷＦ酸素析出１上限
    HWFOS1SH As String * 1         ' 品ＷＦ酸素析出１測定位置＿方
    HWFOS1ST As String * 1         ' 品ＷＦ酸素析出１測定位置＿点
    HWFOS1SI As String * 1         ' 品ＷＦ酸素析出１測定位置＿位
    HWFOS1HT As String * 1         ' 品ＷＦ酸素析出１保証方法＿対
    HWFOS1HS As String * 1         ' 品ＷＦ酸素析出１保証方法＿処
    HWFOS2SH As String * 1         ' 品ＷＦ酸素析出２測定位置＿方
    HWFOS2ST As String * 1         ' 品ＷＦ酸素析出２測定位置＿点
    HWFOS2SI As String * 1         ' 品ＷＦ酸素析出２測定位置＿位
    HWFOS2MN As Double             ' 品ＷＦ酸素析出２下限
    HWFOS2MX As Double             ' 品ＷＦ酸素析出２上限
    HWFOS2HT As String * 1         ' 品ＷＦ酸素析出２保証方法＿対
    HWFOS2HS As String * 1         ' 品ＷＦ酸素析出２保証方法＿処
    HWFOS3MN As Double             ' 品ＷＦ酸素析出３下限
    HWFOS3MX As Double             ' 品ＷＦ酸素析出３上限
    HWFOS3SH As String * 1         ' 品ＷＦ酸素析出３測定位置＿方
    HWFOS3ST As String * 1         ' 品ＷＦ酸素析出３測定位置＿点
    HWFOS3SI As String * 1         ' 品ＷＦ酸素析出３測定位置＿位
    HWFOS3HT As String * 1         ' 品ＷＦ酸素析出３保証方法＿対
    HWFOS3HS As String * 1         ' 品ＷＦ酸素析出３保証方法＿処

    HWFDSOMX As Double             ' 品ＷＦＤＳＯＤ上限              '2003/11/17 SystemBrain Integer ⇒ Double
    HWFDSOMN As Double             ' 品ＷＦＤＳＯＤ下限              '2003/11/17 SystemBrain Integer ⇒ Double
    HWFDSOAX As Integer            ' 品ＷＦＤＳＯＤ領域上限
    HWFDSOAN As Integer            ' 品ＷＦＤＳＯＤ領域下限
    HWFDSOHT As String * 1         ' 品ＷＦＤＳＯＤ保証方法＿対
    HWFDSOHS As String * 1         ' 品ＷＦＤＳＯＤ保証方法＿処

    HWFSPVMX As Double             ' 品ＷＦＳＰＶＦＥ上限
    HWFSPVSH As String * 1         ' 品ＷＦＳＰＶＦＥ測定位置＿方
    HWFSPVST As String * 1         ' 品ＷＦＳＰＶＦＥ測定位置＿点
    HWFSPVSI As String * 1         ' 品ＷＦＳＰＶＦＥ測定位置＿位
    HWFSPVHT As String * 1         ' 品ＷＦＳＰＶＦＥ保証方法＿対
    HWFSPVHS As String * 1         ' 品ＷＦＳＰＶＦＥ保証方法＿処
    HWFDLSPH As String * 1         ' 品ＷＦ拡散長測定位置＿方
    HWFDLSPT As String * 1         ' 品ＷＦ拡散長測定位置＿点
    HWFDLSPI As String * 1         ' 品ＷＦ拡散長測定位置＿位
    HWFDLHWT As String * 1         ' 品ＷＦ拡散長保証方法＿対
    HWFDLHWS As String * 1         ' 品ＷＦ拡散長保証方法＿処
    HWFDLMIN As Integer            ' 品ＷＦ拡散長下限
    HWFDLMAX As Integer            ' 品ＷＦ拡散長上限

    HWFOF1AX As Double             ' 品ＷＦＯＳＦ１平均上限
    HWFOF1MX As Double             ' 品ＷＦＯＳＦ１上限
    HWFOF1SH As String * 1         ' 品ＷＦＯＳＦ１測定位置＿方
    HWFOF1ST As String * 1         ' 品ＷＦＯＳＦ１測定位置＿点
    HWFOF1SR As String * 1         ' 品ＷＦＯＳＦ１測定位置＿領
    HWFOF1HT As String * 1         ' 品ＷＦＯＳＦ１保証方法＿対
    HWFOF1HS As String * 1         ' 品ＷＦＯＳＦ１保証方法＿処
    HWFOF2AX As Double             ' 品ＷＦＯＳＦ２平均上限
    HWFOF2MX As Double             ' 品ＷＦＯＳＦ２上限
    HWFOF2SH As String * 1         ' 品ＷＦＯＳＦ２測定位置＿方
    HWFOF2ST As String * 1         ' 品ＷＦＯＳＦ２測定位置＿点
    HWFOF2SR As String * 1         ' 品ＷＦＯＳＦ２測定位置＿領
    HWFOF2HT As String * 1         ' 品ＷＦＯＳＦ２保証方法＿対
    HWFOF2HS As String * 1         ' 品ＷＦＯＳＦ２保証方法＿処
    HWFOF3AX As Double             ' 品ＷＦＯＳＦ３平均上限
    HWFOF3MX As Double             ' 品ＷＦＯＳＦ３上限
    HWFOF3SH As String * 1         ' 品ＷＦＯＳＦ３測定位置＿方
    HWFOF3ST As String * 1         ' 品ＷＦＯＳＦ３測定位置＿点
    HWFOF3SR As String * 1         ' 品ＷＦＯＳＦ３測定位置＿領
    HWFOF3HT As String * 1         ' 品ＷＦＯＳＦ３保証方法＿対
    HWFOF3HS As String * 1         ' 品ＷＦＯＳＦ３保証方法＿処
    HWFOF4AX As Double             ' 品ＷＦＯＳＦ４平均上限
    HWFOF4MX As Double             ' 品ＷＦＯＳＦ４上限
    HWFOF4SH As String * 1         ' 品ＷＦＯＳＦ４測定位置＿方
    HWFOF4ST As String * 1         ' 品ＷＦＯＳＦ４測定位置＿点
    HWFOF4SR As String * 1         ' 品ＷＦＯＳＦ４測定位置＿領
    HWFOF4HT As String * 1         ' 品ＷＦＯＳＦ４保証方法＿対
    HWFOF4HS As String * 1         ' 品ＷＦＯＳＦ４保証方法＿処
    HWFOSF1PTK As String * 1       ' 品ＷＦＯＳＦ１パタン区分　▼2003/05/14 ooba
    HWFOSF2PTK As String * 1       ' 品ＷＦＯＳＦ２パタン区分
    HWFOSF3PTK As String * 1       ' 品ＷＦＯＳＦ３パタン区分
    HWFOSF4PTK As String * 1       ' 品ＷＦＯＳＦ４パタン区分　▲2003/05/14 ooba

    HWFBM1AN As Double             ' 品ＷＦＢＭＤ１平均下限
    HWFBM1AX As Double             ' 品ＷＦＢＭＤ１平均上限
    HWFBM1SH As String * 1         ' 品ＷＦＢＭＤ１測定位置＿方
    HWFBM1ST As String * 1         ' 品ＷＦＢＭＤ１測定位置＿点
    HWFBM1SR As String * 1         ' 品ＷＦＢＭＤ１測定位置＿領
    HWFBM1HT As String * 1         ' 品ＷＦＢＭＤ１保証方法＿対
    HWFBM1HS As String * 1         ' 品ＷＦＢＭＤ１保証方法＿処
    HWFBM2AN As Double             ' 品ＷＦＢＭＤ２平均下限
    HWFBM2AX As Double             ' 品ＷＦＢＭＤ２平均上限
    HWFBM2SH As String * 1         ' 品ＷＦＢＭＤ２測定位置＿方
    HWFBM2ST As String * 1         ' 品ＷＦＢＭＤ２測定位置＿点
    HWFBM2SR As String * 1         ' 品ＷＦＢＭＤ２測定位置＿領
    HWFBM2HT As String * 1         ' 品ＷＦＢＭＤ２保証方法＿対
    HWFBM2HS As String * 1         ' 品ＷＦＢＭＤ２保証方法＿処
    HWFBM3AN As Double             ' 品ＷＦＢＭＤ３平均下限
    HWFBM3AX As Double             ' 品ＷＦＢＭＤ３平均上限
    HWFBM3SH As String * 1         ' 品ＷＦＢＭＤ３測定位置＿方
    HWFBM3ST As String * 1         ' 品ＷＦＢＭＤ３測定位置＿点
    HWFBM3SR As String * 1         ' 品ＷＦＢＭＤ３測定位置＿領
    HWFBM3HT As String * 1         ' 品ＷＦＢＭＤ３保証方法＿対
    HWFBM3HS As String * 1         ' 品ＷＦＢＭＤ３保証方法＿処
    HWFBM1MBP As Double            ' 品ＷＦＢＭＤ１面内分布　▼2003/05/14 ooba
    HWFBM2MBP As Double            ' 品ＷＦＢＭＤ２面内分布
    HWFBM3MBP As Double            ' 品ＷＦＢＭＤ３面内分布
    HWFBM1MCL As String * 2        ' 品ＷＦＢＭＤ１面内計算
    HWFBM2MCL As String * 2        ' 品ＷＦＢＭＤ２面内計算
    HWFBM3MCL As String * 2        ' 品ＷＦＢＭＤ３面内計算　▲2003/05/14 ooba

    HWFOS1NS As String * 2         ' 品ＷＦ酸素析出１熱処理法
    HWFOS2NS As String * 2         ' 品ＷＦ酸素析出２熱処理法
    HWFOS3NS As String * 2         ' 品ＷＦ酸素析出３熱処理法
    HWFOF1NS As String * 2         ' 品ＷＦＯＳＦ１熱処理法
    HWFOF2NS As String * 2         ' 品ＷＦＯＳＦ２熱処理法
    HWFOF3NS As String * 2         ' 品ＷＦＯＳＦ３熱処理法
    HWFOF4NS As String * 2         ' 品ＷＦＯＳＦ４熱処理法
    HWFBM1NS As String * 2         ' 品ＷＦＢＭＤ１熱処理法
    HWFBM2NS As String * 2         ' 品ＷＦＢＭＤ２熱処理法
    HWFBM3NS As String * 2         ' 品ＷＦＢＭＤ３熱処理法

    HWFANTIM As Integer            ' 品ＷＦＡＮ時間
    HWFANTNP As Integer            ' 品ＷＦＡＮ温度

    HWFOF1ET As Integer            ' 品ＷＦＯＳＦ１選択ＥＴ代
    HWFOF2ET As Integer            ' 品ＷＦＯＳＦ２選択ＥＴ代
    HWFOF3ET As Integer            ' 品ＷＦＯＳＦ３選択ＥＴ代
    HWFOF4ET As Integer            ' 品ＷＦＯＳＦ４選択ＥＴ代
    HWFBM1ET As Integer            ' 品ＷＦＢＭＤ１選択ＥＴ代
    HWFBM2ET As Integer            ' 品ＷＦＢＭＤ２選択ＥＴ代
    HWFBM3ET As Integer            ' 品ＷＦＢＭＤ３選択ＥＴ代

    HWFOF1SZ As String * 1         ' 品ＷＦＯＳＦ１測定条件
    HWFOF2SZ As String * 1         ' 品ＷＦＯＳＦ２測定条件
    HWFOF3SZ As String * 1         ' 品ＷＦＯＳＦ３測定条件
    HWFOF4SZ As String * 1         ' 品ＷＦＯＳＦ４測定条件
    HWFBM1SZ As String * 1         ' 品ＷＦＢＭＤ１測定条件
    HWFBM2SZ As String * 1         ' 品ＷＦＢＭＤ２測定条件
    HWFBM3SZ As String * 1         ' 品ＷＦＢＭＤ３測定条件

    BLOCKID() As String * 12       ' ブロックID
End Type

'SXL管理更新用（現在工程、最終通過工程）
Public Type type_DBDRV_scmzc_fcmlc001c_UpdSXL1
    CRYNUM As String * 12          ' 結晶番号
    INGOTPOS As Integer            ' 結晶内開始位置
    NOWPROC As String * 5          ' 現在工程
    LASTPASS As String * 5         ' 最終通過工程
End Type

'SXL管理更新用（削除区分、最終状態区分）
Public Type type_DBDRV_scmzc_fcmlc001c_UpdSXL2
    CRYNUM As String * 12          ' 結晶番号
    INGOTPOS As Integer            ' 結晶内開始位置
    DELCLS As String * 1           ' 削除区分
    LSTATCLS As String * 1         ' 最終状態区分
End Type

'WFサンプル管理更新用
Public Type type_DBDRV_scmzc_fcmlc001c_UpdWfCrySmp
    CRYNUM As String * 12          ' 結晶番号
    INGOTPOS As Integer            ' 結晶内位置
    SMPKBN As String * 1           ' サンプル区分
End Type


' 再抜試指示
'入力用
Type type_DBDRV_scmzc_fcmlc001d_In
    CRYNUM As String * 12          '結晶番号
    HIN As tFullHinban             '品番
    LENGHT As Integer
End Type

'WF仕様取得用
Public Type type_DBDRV_scmzc_fcmlc001d_WfSiyou
    HWFRMIN As Double              ' 品ＷＦ比抵抗下限
    HWFRMAX As Double              ' 品ＷＦ比抵抗上限
    HWFRHWYS As String * 1         ' 品ＷＦ比抵抗保証方法＿処(Rs)
    HWFONHWS As String * 1         ' 品ＷＦ酸素濃度保証方法＿処(Oi)
    HWFBM1HS As String * 1         ' 品ＷＦＢＭＤ１保証方法＿処(B1)
    HWFBM2HS As String * 1         ' 品ＷＦＢＭＤ２保証方法＿処(B2)
    HWFBM3HS As String * 1         ' 品ＷＦＢＭＤ３保証方法＿処(B3)
    HWFOF1HS As String * 1         ' 品ＷＦＯＳＦ１保証方法＿処(L1)
    HWFOF2HS As String * 1         ' 品ＷＦＯＳＦ２保証方法＿処(L2)
    HWFOF3HS As String * 1         ' 品ＷＦＯＳＦ３保証方法＿処(L3)
    HWFOF4HS As String * 1         ' 品ＷＦＯＳＦ４保証方法＿処(L4)
    HWFDSOHS As String * 1         ' 品ＷＦＤＳＯＤ保証方法＿処(DS)
    HWFMKHWS As String * 1         ' 品ＷＦ無欠陥層保証方法＿処(DZ)
    HWFSPVHS As String * 1         ' 品ＷＦＳＰＶＦＥ保証方法＿処(SP)
    HWFDLHWS As String * 1         ' 品ＷＦ拡散長保証方法＿処(KL)　06/06/08 ooba
    HWFNRHS  As String * 1         ' 品ＷＦＳＰＶＮＲ保証方法＿処(NR)　06/06/08 ooba
    HWFOS1HS As String * 1         ' 品ＷＦ酸素析出１保証方法＿処(D1)
    HWFOS2HS As String * 1         ' 品ＷＦ酸素析出２保証方法＿処(D2)
    HWFOS3HS As String * 1         ' 品ＷＦ酸素析出３保証方法＿処(D3)
    HWFZOHWS As String * 1         ' 品ＷＦ残存酸素保証方法＿処(AO)    ''追加　03/12/15 ooba
    HWFDENHS As String * 1         ' 品ＷＦＤｅｎ保証方法＿処(GD)      '追加　05/02/17 ooba START ====>
    HWFDVDHS As String * 1         ' 品ＷＦＤＶＤ２保証方法＿処(GD)
    HWFLDLHS As String * 1         ' 品ＷＦＬ／ＤＬ保証方法＿処(GD)    '追加　05/02/17 ooba END ======>
    HWFOT1   As String * 1         ' 03/05/26
    HWFOT2   As String * 1         ' 03/05/26
    KEIKAKUL As Integer            ' 計画長
    HWFMAI1   As String * 1        ' 04/07/16
    HWFMAI2   As String * 1        ' 04/07/16
' 06/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    HEPOF1HS As String * 1         ' 検査有無(OSF1E)
    HEPOF2HS As String * 1         ' 検査有無(OSF2E)
    HEPOF3HS As String * 1         ' 検査有無(OSF3E)
    HEPBM1HS As String * 1         ' 検査有無(BMD1E)
    HEPBM2HS As String * 1         ' 検査有無(BMD2E)
    HEPBM3HS As String * 1         ' 検査有無(BMD3E)
' 06/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
' ↓10/01/06 Add SIRD対応 Y.Hitomi
    HWFSIRDHS As String * 1         ' 検査有無(SIRD)
' ↑10/01/06 Add SIRD対応 Y.Hitomi
    CHUTAN   As Integer            ' 中間抜試単位(枚)
    CHUKYO   As Integer            ' 中間抜試許容値
    CHUFLG   As String             ' 中間抜試フラグ
End Type

'WFサンプル管理（TOP、TAIL ２レコード）
Public Type type_DBDRV_scmzc_fcmlc001d_WfSmp
    INGOTPOS As Integer            ' 結晶内位置
    SMPLID As String * 16          ' サンプルID
    hinban As String * 8           ' 品番
    REVNUM As Integer              ' 製品番号改訂番号
    factory As String * 1          ' 工場
    opecond As String * 1          ' 操業条件
    WFINDRS As String * 1          ' 状態FLG（Rs)
    WFINDOI As String * 1          ' 状態FLG（Oi)
    WFINDB1 As String * 1          ' 状態FLG（B1)
    WFINDB2 As String * 1          ' 状態FLG（B2）
    WFINDB3 As String * 1          ' 状態FLG（B3)
    WFINDL1 As String * 1          ' 状態FLG（L1)
    WFINDL2 As String * 1          ' 状態FLG（L2)
    WFINDL3 As String * 1          ' 状態FLG（L3)
    WFINDL4 As String * 1          ' 状態FLG（L4)
    WFINDDS As String * 1          ' 状態FLG（DS)
    WFINDDZ As String * 1          ' 状態FLG（DZ)
    WFINDSP As String * 1          ' 状態FLG（SP)
    WFINDDO1 As String * 1         ' 状態FLG（DO1)
    WFINDDO2 As String * 1         ' 状態FLG（DO2)
    WFINDDO3 As String * 1         ' 状態FLG（DO3)
    WFINDOTHER1 As String * 1      ' 検査有無(OT2) ''Add.03/05/20 後藤
    WFINDOTHER2 As String * 1      ' 検査有無(OT1) ''Add.03/05/20
    WFINDAOI As String * 1         ' 状態FLG (AOi)     '残存酸素追加　03/12/15 ooba
    WFINDGD As String * 1          ' 状態FLG (GD)      'GD追加　05/02/17 ooba
    WFHSGD As String * 1           ' 保証FLG (GD)      'GD追加　05/02/17 ooba
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    EPINDB1CW As String * 1        '状態FLG(BMD1)
    EPINDB2CW As String * 1        '状態FLG(BMD2)
    EPINDB3CW As String * 1        '状態FLG(BMD3)
    EPINDL1CW As String * 1        '状態FLG(OSF1)
    EPINDL2CW As String * 1        '状態FLG(OSF2)
    EPINDL3CW As String * 1        '状態FLG(OSF3)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
End Type

Public Type typ_WfSampleGr
    BLOCKID As String
    blockp As Integer
    WFSMP As typ_XSDCW
    HINUP As tFullHinban           '上品番
    HINDN As tFullHinban           '下品番
    ERRDNFLG As Boolean            '下品番エラーフラグ
End Type

' 欠落ウェハー情報
Public Type typ_LackWaf
    BLOCKID As String * 12         'ブロックID
    WAFERNO As Integer             'ウェハー連番
    TOP_POS As Integer             'ウェハー開始位置
    TAIL_POS As Integer            'ウェハー終了位置
End Type
'比抵抗
Public Type NoTest_RES
    HWFRHWYS As String * 1         '品WF比抵抗保証方法＿処
End Type
'酸素濃度
Public Type NoTest_OI
    HWFONHWS As String * 1         '品WF酸素濃度保証方法＿処
End Type
'BMDx
Public Type NoTest_BMD
    HWFBMxHS As String * 1         '品WFBMDx保証方法＿処
    HWFBMxET As Integer            '品WFBMD1選択ET代
    HWFBMxNS As String * 2         '品WFBMD1熱処理法
    HWFBMxSZ As String * 1         '品WFBMD1測定条件
    HWFBMxSH As String * 1         '品WFBMD1測定位置_方
    HWFBMxST As String * 1         '品WFBMD1測定位置_点
    HWFBMxSR As String * 1         '品WFBMD1測定位置_領
End Type
'OSFx
Public Type NoTest_OSF
    HWFOFxHS As String * 1         '品WFOSFx保証方法＿処
    HWFOFxET As Integer            '品WFOSF1選択ET代
    HWFOFxNS As String * 2         '品WFOSF1熱処理法
    HWFOFxSZ As String * 1         '品WFOSF1測定条件
    HWFOFxSH As String * 1         '品WFOSF1測定位置_方
    HWFOFxST As String * 1         '品WFOSF1測定位置_点
    HWFOFxSR As String * 1         '品WFOSF1測定位置_領
End Type
'DSOD
Public Type NoTest_DSOD
    HWFDSOHS As String * 1         '品WFDSOD保証方法＿処
    HWFDSOKE As String * 1         '品WFDSOD検査
End Type
'DZ
Public Type NoTest_DZ
    HWFMKHWS As String * 1         '品WF無欠陥層保証方法＿処
    HWFMKSZY As String * 1         '品WF無欠陥層測定条件
    HWFMKSPH As String * 1         '品WF無欠陥層測定位置＿方
    HWFMKSPT As String * 1         '品WF無欠陥層測定位置＿点
    HWFMKSPR As String * 1         '品WF無欠陥層測定位置＿領
End Type
'SPVFE
Public Type NoTest_SPVFE
    HWFSPVHS As String * 1         '品WFSPVFE保証方法＿処
    HWFSPVSH As String * 1         '品WFSPVFE測定位置＿方
    HWFSPVST As String * 1         '品WFSPVFE測定位置＿点
    HWFSPVSI As String * 1         '品WFSPVFE測定位置＿位
End Type
'拡散長
Public Type NoTest_SPV
    HWFDLHWS As String * 1         '品WF拡散長保証方法＿処
    HWFDLSPH As String * 1         '品WF拡散長測定位置＿方
    HWFDLSPT As String * 1         '品WF拡散長測定位置＿点
    HWFDLSPI As String * 1         '品WF拡散長測定位置＿位
End Type
'⊿Oix
Public Type NoTest_DOI
    HWFOSxHS As String * 1         '品WF酸素析出x保証方法＿処
    HWFOSxNS As String * 2         '品WF酸素析出1熱処理法
    HWFOSxSH As String * 1         '品WF酸素析出1測定位置＿方
    HWFOSxST As String * 1         '品WF酸素析出1測定位置＿点
    HWFOSxSI As String * 1         '品WF酸素析出1測定位置＿位
End Type

Public Type NoTest_Info
    Res As NoTest_RES
    Oi As NoTest_OI
    BMD(2) As NoTest_BMD
    OSF(3) As NoTest_OSF
    Dsod As NoTest_DSOD
    DZ As NoTest_DZ
    SpvFe As NoTest_SPVFE
    Spv As NoTest_SPV
    Doi(2) As NoTest_DOI
End Type

' WFサンプル仕様(*は未チェックのパラメータ)
Public Type typ_SpWFSamp
    HIN As tFullHinban             ' 品番

    HWFRHWYS As String * 1         ' 処理方法(Rs)
    HWFRSPOH As String * 1         ' 測定方法(Rs)*
    HWFRSPOT As String * 1         ' 測定点数(Rs) -> Heavy
    HWFRSPOI As String * 1         ' 測定位置(Rs)*

    HWFONHWS As String * 1         ' 処理方法(Oi)
    HWFONKWY As String * 2         ' 検査方法(Oi)
    HWFONSPH As String * 1         ' 測定方法(Oi)
    HWFONSPT As String * 1         ' 測定点数(Oi) -> Heavy
    HWFONSPI As String * 1         ' 測定位置(Oi)

    HWFBM1HS As String * 1         ' 処理方法(B1)
    HWFBM1SH As String * 1         ' 測定方法(B1)
    HWFBM1ST As String * 1         ' 測定点数(B1)
    HWFBM1SR As String * 1         ' 除外領域(B1)
    HWFBM1NS As String * 2         ' 熱処理法(B1)
    HWFBM1SZ As String * 1         ' 測定条件(B1)
    HWFBM1ET As Integer            ' 選択エッチ(B1)

    HWFBM2HS As String * 1         ' 処理方法(B2)
    HWFBM2SH As String * 1         ' 測定方法(B2)
    HWFBM2ST As String * 1         ' 測定点数(B2)
    HWFBM2SR As String * 1         ' 除外領域(B2)
    HWFBM2NS As String * 2         ' 熱処理法(B2)
    HWFBM2SZ As String * 1         ' 測定条件(B2)
    HWFBM2ET As Integer            ' 選択エッチ(B2)

    HWFBM3HS As String * 1         ' 処理方法(B3)
    HWFBM3SH As String * 1         ' 測定方法(B3)
    HWFBM3ST As String * 1         ' 測定点数(B3)
    HWFBM3SR As String * 1         ' 除外領域(B3)
    HWFBM3NS As String * 2         ' 熱処理法(B3)
    HWFBM3SZ As String * 1         ' 測定条件(B3)
    HWFBM3ET As Integer            ' 選択エッチ(B3)

    HWFOF1HS As String * 1         ' 処理方法(L1)
    HWFOF1SH As String * 1         ' 測定方法(L1)
    HWFOF1ST As String * 1         ' 測定点数(L1)
    HWFOF1SR As String * 1         ' 除外領域(L1)
    HWFOF1NS As String * 2         ' 熱処理法(L1)
    HWFOF1SZ As String * 1         ' 測定条件(L1)
    HWFOF1ET As Integer            ' 選択エッチ(L1)

    HWFOF2HS As String * 1         ' 処理方法(L2)
    HWFOF2SH As String * 1         ' 測定方法(L2)
    HWFOF2ST As String * 1         ' 測定点数(L2)
    HWFOF2SR As String * 1         ' 除外領域(L2)
    HWFOF2NS As String * 2         ' 熱処理法(L2)
    HWFOF2SZ As String * 1         ' 測定条件(L2)
    HWFOF2ET As Integer            ' 選択エッチ(L2)

    HWFOF3HS As String * 1         ' 処理方法(L3)
    HWFOF3SH As String * 1         ' 測定方法(L3)
    HWFOF3ST As String * 1         ' 測定点数(L3)
    HWFOF3SR As String * 1         ' 除外領域(L3)
    HWFOF3NS As String * 2         ' 熱処理法(L3)
    HWFOF3SZ As String * 1         ' 測定条件(L3)
    HWFOF3ET As Integer            ' 選択エッチ(L3)

'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''    HWFOF4HS As String * 1         ' 処理方法(L4)
'''    HWFOF4SH As String * 1         ' 測定方法(L4)
'''    HWFOF4ST As String * 1         ' 測定点数(L4)
'''    HWFOF4SR As String * 1         ' 除外領域(L4)
'''    HWFOF4NS As String * 2         ' 熱処理法(L4)
'''    HWFOF4SZ As String * 1         ' 測定条件(L4)
'''    HWFOF4ET As Integer            ' 選択エッチ(L4)
    
    HWFSIRDMX As Integer       '軸状転位上限(SIRD)
    HWFSIRDSZ As String * 1    '軸状転位測定条件(SIRD)
    HWFSIRDHT As String * 1    '軸状転位保証方法＿対(SIRD)
    HWFSIRDHS As String * 1    '軸状転位保証方法＿処(SIRD)
    HWFSIRDKM As String * 1    '軸状転位検査頻度＿枚(SIRD)
    HWFSIRDKH As String * 1    '軸状転位検査頻度＿保(SIRD)
    HWFSIRDKU As String * 1    '軸状転位検査頻度＿ウ(SIRD)
    HWFSIRDPS As String * 2    '軸状転位TB保証位置(SIRD)
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)

    HWFDSOHS As String * 1         ' 処理方法(DS)

    HWFMKHWS As String * 1         ' 処理方法(DZ)
    HWFMKSPH As String * 1         ' 測定方法(DZ)
    HWFMKSPT As String * 1         ' 測定点数(DZ)
    HWFMKSPR As String * 1         ' 除外領域(DZ)
    HWFMKNSW As String * 2         ' 熱処理法(DZ)
    HWFMKSZY As String * 1         ' 測定条件(DZ)
    HWFMKCET As Integer            ' 選択エッチ(DZ)

    HWFSPVHS As String * 1         ' 処理方法(SP/Fe濃度)
    HWFSPVSH As String * 1         ' 測定方法(SP/Fe濃度)*
    HWFSPVST As String * 1         ' 測定点数(SP/Fe濃度)*
    HWFSPVSI As String * 1         ' 測定位置(SP/Fe濃度)*
    HWFDLHWS As String * 1         ' 処理方法(SP/拡散長)
    HWFDLSPH As String * 1         ' 測定方法(SP/拡散長)*
    HWFDLSPT As String * 1         ' 測定点数(SP/拡散長)*
    HWFDLSPI As String * 1         ' 測定位置(SP/拡散長)*
    HWFNRHS  As String * 1         ' 処理方法(SP/Nr濃度)               06/06/08 ooba START ======>
    HWFNRSH  As String * 1         ' 測定方法(SP/Nr濃度)*
    HWFNRST  As String * 1         ' 測定点数(SP/Nr濃度)*
    HWFNRSI  As String * 1         ' 測定位置(SP/Nr濃度)*
    HWFSPVPUG   As String * 10     ' PUA限(SP/Fe濃度)*
    HWFSPVPUR   As String * 10     ' PUA率(SP/Fe濃度)*
    HWFSPVSTD   As String * 10     ' 標準偏差(SP/Fe濃度)*
    HWFDLPUG    As String * 10     ' PUA限(SP/拡散長)*
    HWFDLPUR    As String * 10     ' PUA率(SP/拡散長)*
    HWFNRPUG    As String * 10     ' PUA限(SP/Nr濃度)*
    HWFNRPUR    As String * 10     ' PUA率(SP/Nr濃度)*
    HWFNRSTD    As String * 10     ' 標準偏差(SP/Nr濃度)*      06/06/08 ooba END ========>

    HWFOS1HS As String * 1         ' 処理方法(D1)
    HWFOS1SH As String * 1         ' 測定方法(D1)*
    HWFOS1ST As String * 1         ' 測定点数(D1)*
    HWFOS1SI As String * 1         ' 測定位置(D1)*
    HWFOS1NS As String * 2         ' 熱処理法(D1)

    HWFOS2HS As String * 1         ' 処理方法(D2)
    HWFOS2SH As String * 1         ' 測定方法(D2)*
    HWFOS2ST As String * 1         ' 測定点数(D2)*
    HWFOS2SI As String * 1         ' 測定位置(D2)*
    HWFOS2NS As String * 2         ' 熱処理法(D2)

    HWFOS3HS As String * 1         ' 処理方法(D3)
    HWFOS3SH As String * 1         ' 測定方法(D3)*
    HWFOS3ST As String * 1         ' 測定点数(D3)*
    HWFOS3SI As String * 1         ' 測定位置(D3)*
    HWFOS3NS As String * 2         ' 熱処理法(D3)

    HWFZOHWS As String * 1         ' 処理方法(AO)  ''追加 03/12/15 ooba START ======>
    HWFZOSPH As String * 1         ' 測定方法(AO)*
    HWFZOSPT As String * 1         ' 測定点数(AO)*
    HWFZOSPI As String * 1         ' 測定位置(AO)*
    HWFZONSW As String * 2         ' 熱処理法(AO)  ''追加 03/12/15 ooba END ========>

    HWFDENHS As String * 1         ' 処理方法(GD/DEN)  '追加　05/02/18 ooba START ====>
    HWFLDLHS As String * 1         ' 処理方法(GD/LDL)
    HWFDVDHS As String * 1         ' 処理方法(GD/DVD2) '追加　05/02/18 ooba END ======>
    HWFGDSPH As String * 1         ' 測定方法(GD)　    '05/10/25 ooba
    HWFGDSPT As String * 1         ' 測定点数(GD)　    '05/10/25 ooba
    HWFGDZAR As String * 1         ' 除外領域(GD)　    '05/10/25 ooba

    HWFRKHNN As String * 1         ' 検査頻度_抜(Rs)   '追加　04/04/12 ooba START ====>
    HWFONKHN As String * 1         ' 検査頻度_抜(Oi)
    HWFOF1KN As String * 1         ' 検査頻度_抜(L1)
    HWFOF2KN As String * 1         ' 検査頻度_抜(L2)
    HWFOF3KN As String * 1         ' 検査頻度_抜(L3)
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
''    HWFOF4KN As String * 1         ' 検査頻度_抜(L4)
    HWFSIRDKN As String * 1  ' 検査頻度_抜(SIRD)
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
    HWFBM1KN As String * 1         ' 検査頻度_抜(B1)
    HWFBM2KN As String * 1         ' 検査頻度_抜(B2)
    HWFBM3KN As String * 1         ' 検査頻度_抜(B3)
    HWFOS1KN As String * 1         ' 検査頻度_抜(D1)
    HWFOS2KN As String * 1         ' 検査頻度_抜(D2)
    HWFOS3KN As String * 1         ' 検査頻度_抜(D3)
    HWFDSOKN As String * 1         ' 検査頻度_抜(DS)
    HWFMKKHN As String * 1         ' 検査頻度_抜(DZ)
    HWFSPVKN As String * 1         ' 検査頻度_抜(SP/Fe濃度)
    HWFDLKHN As String * 1         ' 検査頻度_抜(SP/拡散長)
    HWFZOKHN As String * 1         ' 検査頻度_抜(AO)   '追加　04/04/12 ooba END ======>
    HWFGDKHN As String * 1         ' 検査頻度_抜(GD)　05/02/18 ooba
    HWFNRKN  As String * 1         ' 検査頻度_抜(SP/Nr濃度)  06/06/08 ooba

    HWFIGKBN As String * 1         ' IG区分
    HWFANTNP As Integer            ' DKアニール条件(温度)
    HWFANTIM As Integer            ' DKアニール条件(時間)
    HWFANGZY As String * 1         ' DKアニール条件(ガス)　04/07/29 ooba
    HWOTHER1 As String * 1         ' 検査有無(OT2) ''Add.03/05/20 後藤
    HWOTHER2 As String * 1         ' 検査有無(OT1) ''Add.03/05/20
    HWOTHER1MAI As String * 1      ' 04/07/16
    HWOTHER2MAI As String * 1      ' 04/07/16

''Upd Start (TCS)T.Terauchi 2005/10/05  抜試指示4.5ﾗｲﾝ対応
    HWFGDLINE   As String * 3      '品WFGDﾗｲﾝ数(TBCME036)

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    HEPOF1NS As String * 2         ' 品熱処理法(OSF1E)
    HEPOF1SZ As String * 1         ' 品測定条件(OSF1E)
    HEPOF1ET As Integer            ' 品選択ET代(OSF1E)
    HEPOF1HS As String * 1         ' 品保証方法_処(OSF1E)
    HEPOF1SH As String * 1         ' 品測定位置_方(OSF1E)
    HEPOF1ST As String * 1         ' 品測定位置_点(OSF1E)
    HEPOF1SR As String * 1         ' 品測定位置_領(OSF1E)
    HEPOF1KN As String * 1         ' 品検査頻度_抜(OSF1E)
    HEPOF2NS As String * 2         ' 品熱処理法(OSF2E)
    HEPOF2SZ As String * 1         ' 品測定条件(OSF2E)
    HEPOF2ET As Integer            ' 品選択ET代(OSF2E)
    HEPOF2HS As String * 1         ' 品保証方法_処(OSF2E)
    HEPOF2SH As String * 1         ' 品測定位置_方(OSF2E)
    HEPOF2ST As String * 1         ' 品測定位置_点(OSF2E)
    HEPOF2SR As String * 1         ' 品測定位置_領(OSF2E)
    HEPOF2KN As String * 1         ' 品検査頻度_抜(OSF2E)
    HEPOF3NS As String * 2         ' 品熱処理法(OSF3E)
    HEPOF3SZ As String * 1         ' 品測定条件(OSF3E)
    HEPOF3ET As Integer            ' 品選択ET代(OSF3E)
    HEPOF3HS As String * 1         ' 品保証方法_処(OSF3E)
    HEPOF3SH As String * 1         ' 品測定位置_方(OSF3E)
    HEPOF3ST As String * 1         ' 品測定位置_点(OSF3E)
    HEPOF3SR As String * 1         ' 品測定位置_領(OSF3E)
    HEPOF3KN As String * 1         ' 品検査頻度_抜(OSF3E)
    HEPBM1NS As String * 2         ' 品熱処理法(BMD1E)
    HEPBM1SZ As String * 1         ' 品測定条件(BMD1E)
    HEPBM1ET As Integer            ' 品選択ET代(BMD1E)
    HEPBM1HS As String * 1         ' 品保証方法_処(BMD1E)
    HEPBM1SH As String * 1         ' 品測定位置_方(BMD1E)
    HEPBM1ST As String * 1         ' 品測定位置_点(BMD1E)
    HEPBM1SR As String * 1         ' 品測定位置_領(BMD1E)
    HEPBM1KN As String * 1         ' 品検査頻度_抜(BMD1E)
    HEPBM2NS As String * 2         ' 品熱処理法(BMD2E)
    HEPBM2SZ As String * 1         ' 品測定条件(BMD2E)
    HEPBM2ET As Integer            ' 品選択ET代(BMD2E)
    HEPBM2HS As String * 1         ' 品保証方法_処(BMD2E)
    HEPBM2SH As String * 1         ' 品測定位置_方(BMD2E)
    HEPBM2ST As String * 1         ' 品測定位置_点(BMD2E)
    HEPBM2SR As String * 1         ' 品測定位置_領(BMD2E)
    HEPBM2KN As String * 1         ' 品検査頻度_抜(BMD2E)
    HEPBM3NS As String * 2         ' 品熱処理法(BMD3E)
    HEPBM3SZ As String * 1         ' 品測定条件(BMD3E)
    HEPBM3ET As Integer            ' 品選択ET代(BMD3E)
    HEPBM3HS As String * 1         ' 品保証方法_処(BMD3E)
    HEPBM3SH As String * 1         ' 品測定位置_方(BMD3E)
    HEPBM3ST As String * 1         ' 品測定位置_点(BMD3E)
    HEPBM3SR As String * 1         ' 品測定位置_領(BMD3E)
    HEPBM3KN As String * 1         ' 品検査頻度_抜(BMD3E)
    HEPACEN  As Double             ' 品E1厚中心
    HEPANTNP As Integer            ' 品EPAN温度
    HEPANTIM As Integer            ' 品EPAN時間
    HEPIGKBN As String * 1         ' 品EPIG区分
    HEPANGZY As String * 1         ' 品EP高温ANガス条件
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    HWFGDSZY As String * 1         ' 品ＷＦＧＤ測定条件

'--------------- 2008/08/25 INSERT START  By Systech ---------------
    HSXDKTMP As String * 1         ' DK温度
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
End Type

' WFサンプルテーブル
Public Type typ_WFSample
    CRYINDRS As String * 1         ' 検査項目(Rs)
    CRYINDOI As String * 1         ' 検査項目(Oi)
    CRYINDB1 As String * 1         ' 検査項目(B1)
    CRYINDB2 As String * 1         ' 検査項目(B2）
    CRYINDB3 As String * 1         ' 検査項目(B3)
    CRYINDL1 As String * 1         ' 検査項目(L1)
    CRYINDL2 As String * 1         ' 検査項目(L2)
    CRYINDL3 As String * 1         ' 検査項目(L3)
    CRYINDL4 As String * 1         ' 検査項目(L4)
    CRYINDDS As String * 1         ' 検査項目(DS)
    CRYINDDZ As String * 1         ' 検査項目(DZ)
    CRYINDSP As String * 1         ' 検査項目(SP)
    CRYINDD1 As String * 1         ' 検査項目(D1)
    CRYINDD2 As String * 1         ' 検査項目(D2)
    CRYINDD3 As String * 1         ' 検査項目(D3)
    CRYOTHER1 As String * 1        ' 検査有無(OT2) ''Add.03/05/20 後藤
    CRYOTHER2 As String * 1        ' 検査有無(OT1) ''Add.03/05/20
    CRYINDAO As String * 1         ' 検査項目(AO)      ''追加　03/12/15 ooba
    CRYINDGD As String * 1         ' 検査有無(GD)      '追加 05/01/18 ooba
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    CRYINDGD2 As String * 1        ' 検査有無(GD測定条件用)
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
    WFHSGD As String * 1           ' 保証FLG(GD)       '追加 05/01/18 ooba
' 06/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    EPIINDL1 As String * 1         ' 検査有無(OSF1E)
    EPIINDL2 As String * 1         ' 検査有無(OSF2E)
    EPIINDL3 As String * 1         ' 検査有無(OSF3E)
    EPIINDB1 As String * 1         ' 検査有無(BMD1E)
    EPIINDB2 As String * 1         ' 検査有無(BMD2E)
    EPIINDB3 As String * 1         ' 検査有無(BMD3E)
End Type

'2002/09/11 ADD hitec)N.MATSUMOTO Start
Public strBlockID()    As String
Public Const PROCD_WFC_SAINUKISI = "CW760"  'WFセンター再抜試
Public Const PROCD_SXL_MAP = "TX860"        'シングルマップ
Public Const WF_HANTEI_FORM As Integer = 1  '画面の判定（WFセンター総合判定）
Public Const SAINUKISI_FORM As Integer = 2  '画面の判定（再抜試指示）


'2002/09/11 ADD hitec)N.MATSUMOTO  End


'=================================
'2003/02/28 ADD HITEC)okazaki start

Public Type type_DBDRV_Nukisi
    LOTID       As String * 12     ' ブロックID
    SXLID       As String * 13     ' SXLID
    MinMax      As Integer         ' 0:MIN 1:MAX
    BLOCKSEQ    As String * 3      ' ブロック内連番
    WFSTA       As String * 1      ' WF状態
    hinban      As String * 8      ' 品番
    RTOP_POS    As Double          ' 論理ブロック内位置
    RITOP_POS   As Double          ' 論理結晶内位置
    SMPLEID     As String * 16     ' 抜試位置
    SHAFLAG     As String * 1      ' サンプルフラグ
    INDTM       As Date
    BASKETID    As String * 6
    SLOTNO      As Integer
    CURRWPCS    As Integer
    EXISTFLG    As String * 1
    TOP_POS     As Integer
    REJCAT      As String * 1
    TXID        As String * 6
    REGDATE     As Date
    SUMMITSENDFLAG As String * 1
    SENDFLAG    As String * 1
    SENDDATE    As Date
    HREJCODE    As String * 4
    UPDPROC     As String * 5
    UPDDATE     As Date
    REVNUM      As Integer
    factory     As String * 1
    opecond     As String * 1
    KANKBN      As String * 1
    NREJCODE    As String * 6
    SMPLEFLG    As String
End Type
Public Type type_DBDRV_LOTSXL
    LOTID       As String * 12     ' ブロックID
    SXLID       As String * 13     ' SXLID
End Type
'2003/02/28 Add HITEC)okazaki end

'2003/02/28 Hitec)okazaki add start
Public tExamine() As type_DBDRV_Nukisi  '画面表示時
                                        'ウェハーセンター入庫情報テーブル
Public tKeturaku() As typ_TBCMY012

Public tSXLID() As type_DBDRV_LOTSXL
'2003/02/28 Hitec)okazaki add end

'add  2003/03/15 hitec)matsumoto ---------------
Public bWfmapView As Boolean

Public CngSmpID_UD()    As String       ' UD→TB変更用　2004/01/29 ooba
Public bMotoGDcpyFlg(2) As Boolean      ' 初期行の結晶GD引継ぎ有無　05/08/04 ooba

'add 2003/03/25 hitec)matsumoto ｸﾞﾛｰﾊﾞﾙ関数として使いたいので、f_cmbc039_3.frmより移動----------------
Public SIngotP As Integer              ' インゴット上側位置
Public EIngotP As Integer              ' インゴット下側位置
'add 2003/03/25 hitec)matsumoto ------------------------------

Public tblSXL As DBDRV_scmzc_fcmlc001b_SXL ' SXL管理（待ち一覧から）    'upd 2003/04/27 hitec)matsumoto f_cmbc039_3より移動

'WFサンプル実績FLG更新対象ﾁｪｯｸ結果構造体　04/02/06 tuku
Public Type type_chkUP
    rs As String * 1               ' 受信FLG（Rs)
    Oi As String * 1               ' 受信FLG（Oi)
    B1 As String * 1               ' 受信FLG（B1)
    B2 As String * 1               ' 受信FLG（B2）
    B3 As String * 1               ' 受信FLG（B3)
    L1 As String * 1               ' 受信FLG（L1)
    L2 As String * 1               ' 受信FLG（L2)
    L3 As String * 1               ' 受信FLG（L3)
    L4 As String * 1               ' 受信FLG（L4)
    DS As String * 1               ' 受信FLG（DS)
    DZ As String * 1               ' 受信FLG（DZ)
    sp As String * 1               ' 受信FLG（SP)
    DO1 As String * 1              ' 受信FLG（DO1)
    DO2 As String * 1              ' 受信FLG（DO2)
    DO3 As String * 1              ' 受信FLG（DO3)
    OT1 As String * 1              ' 受信FLG (OT1)
    OT2 As String * 1              ' 受信FLG (OT2)
    AOI As String * 1              ' 受信FLG (AOi)
    GD As String * 1               ' 受信FLG (GD)   '05/02/04 ooba
    B1E As String * 1              ' 受信FLG（B1E)
    B2E As String * 1              ' 受信FLG（B2E）
    B3E As String * 1              ' 受信FLG（B3E)
    L1E As String * 1              ' 受信FLG（L1E)
    L2E As String * 1              ' 受信FLG（L2E)
    L3E As String * 1              ' 受信FLG（L3E)
End Type

'**************************************************************************************
'*    関数名        : KeturakuInfo
'*
'*    処理概要      : 1.欠落有無取得
'*
'*    パラメータ    : 変数名        ,IO ,型                            ,説明
'*                    SXL           ,O  ,DBDRV_scmzc_fcmlc001b_SXL039  ,SXL管理
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'**************************************************************************************
Private Function KeturakuInfo(sxl() As DBDRV_scmzc_fcmlc001b_SXL039) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim i           As Long
    Dim j           As Long
    Dim rs          As OraDynaset
    Dim lngRecCnt   As Long
    Dim intSXLCnt   As Integer
    Dim sSXLID      As String

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_sSql.bas -- Function KeturakuInfo"

    KeturakuInfo = FUNCTION_RETURN_SUCCESS

    intSXLCnt = UBound(sxl)

#If True Then   'New Version  2002.1.24
    sSQL = "select distinct SXL.sSXLID "
    sSQL = sSQL & "from TBCME042 SXL, TBCME040 BLK, TBCMY012 REJ "
    sSQL = sSQL & "where"
    sSQL = sSQL & "  REJ.LOTID=BLK.BLOCKID"
    sSQL = sSQL & "  and SXL.CRYNUM=BLK.CRYNUM"
    sSQL = sSQL & "  and SXL.DELCLS<>'1'"
    sSQL = sSQL & "  and ("
    sSQL = sSQL & "    ("
    sSQL = sSQL & "      REJ.ALLSCRAP='Y'"
    sSQL = sSQL & "      and SXL.INGOTPOS<BLK.INGOTPOS+BLK.LENGTH"
    sSQL = sSQL & "      and SXL.INGOTPOS+SXL.LENGTH>BLK.INGOTPOS"
    sSQL = sSQL & "    ) or ("
    sSQL = sSQL & "      REJ.ALLSCRAP='N'"
    sSQL = sSQL & "      and REJ.REJCAT='A'"
    sSQL = sSQL & "      and (SXL.INGOTPOS < BLK.INGOTPOS + REJ.LENTO)"
    sSQL = sSQL & "      and (SXL.INGOTPOS + SXL.LENGTH > BLK.INGOTPOS + REJ.LENFROM)"
    sSQL = sSQL & "    ) or ("
    sSQL = sSQL & "      REJ.REJCAT='B'"
    sSQL = sSQL & "      and BLK.INGOTPOS + REJ.TOP_POS/10.0 between SXL.INGOTPOS and SXL.INGOTPOS + SXL.LENGTH"
    sSQL = sSQL & "    )"
    sSQL = sSQL & "  )"
#Else
    sSQL = "select "
    sSQL = sSQL & " distinct sSXLID "
    sSQL = sSQL & " from "
    sSQL = sSQL & " VECMW002 K, XSDCA A, TBCME040 B "
    sSQL = sSQL & " where "
    sSQL = sSQL & " A.CRYNUMCA = B.CRYNUM "
    sSQL = sSQL & " and B.BLOCKID = K.BLOCKID "
    sSQL = sSQL & " and ( "
    sSQL = sSQL & " ((B.INGOTPOS + K.TOP_POS) >= A.INPOSCA and (B.INGOTPOS + K.TOP_POS) < (A.INPOSCA + A.GNLCA)) "
    sSQL = sSQL & " or ((B.INGOTPOS + K.TAIL_POS) > A.INPOSCA and (B.INGOTPOS + K.TAIL_POS) < (A.INPOSCA + A.GNLCA)) "
    sSQL = sSQL & " or (A.INPOSCA >= (B.INGOTPOS + K.TOP_POS)  and A.INPOSCA < (B.INGOTPOS + K.TAIL_POS)) "
    sSQL = sSQL & " or ((A.INPOSCA + A.GNLCA) > (B.INGOTPOS + K.TOP_POS) and (A.INPOSCA + A.GNLCA) < (B.INGOTPOS + K.TAIL_POS)) "
    sSQL = sSQL & " and S.sSXLID in ("
    For i = 1 To intSXLCnt
        If i = intSXLCnt Then
            sSQL = sSQL & "'" & sxl(i).sSXLID & "' "
        Else
            sSQL = sSQL & "'" & sxl(i).sSXLID & "', "
        End If
    Next
    sSQL = sSQL & ") "
#End If
    Debug.Print sSQL
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    lngRecCnt = rs.RecordCount

    '初期化
    For i = 1 To intSXLCnt
        sxl(i).KETURAKU = False
    Next

    'sSql結果のsSXLIDが欠落ありのsSXLID
    For i = 1 To lngRecCnt
        sSXLID = rs("sSXLID")
        For j = 1 To intSXLCnt
            If sSXLID = sxl(j).CRYNUMCA Then
                sxl(j).KETURAKU = True
            End If
        Next
        rs.MoveNext
    Next
    rs.Close

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error sSql ======"
    Debug.Print sSQL
    KeturakuInfo = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'***********************************************************************************
'*    関数名        : GetMaisu
'*
'*    処理概要      : 1.WF枚数取得
'*
'*    パラメータ    : 変数名        ,IO ,型                            ,説明
'*                    SXL           ,O  ,DBDRV_scmzc_fcmlc001b_SXL039  ,SXL管理
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'***********************************************************************************
Private Function GetMaisu(sxl() As DBDRV_scmzc_fcmlc001b_SXL039) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim i           As Long
    Dim j           As Long
    Dim rs          As OraDynaset
    Dim lngRecCnt   As Long
    Dim intSXLCnt   As Integer
    Dim sSXLID      As String

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_sSql.bas -- Function getMaisu"

    GetMaisu = FUNCTION_RETURN_SUCCESS

    intSXLCnt = UBound(sxl)

    sSQL = sSQL & "SELECT sSXLIDCB,MAICB "
    sSQL = sSQL & "FROM XSDCB "
    sSQL = sSQL & "GROUP BY sSXLIDCB,MAICB"

    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    lngRecCnt = rs.RecordCount

    '初期化
    For i = 1 To intSXLCnt
        sxl(i).MAICB = 0
    Next

    '枚数格納
    For i = 1 To lngRecCnt
        sSXLID = rs("sSXLIDCB")
        For j = 1 To intSXLCnt
            If sSXLID = sxl(j).SXLIDCA Then
                sxl(j).MAICB = rs("MAICB")
                Exit For
            End If
        Next
        rs.MoveNext
    Next
    rs.Close

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error sSql ======"
    Debug.Print sSQL
    GetMaisu = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'*******************************************************************************************
'*    関数名        : GetsSXLIDINBlkid
'*
'*    処理概要      : 1.SXLの全ブロック入庫チェック
'*
'*    パラメータ    : 変数名        ,IO ,型                            ,説明
'*                    SXL           ,O  ,DBDRV_scmzc_fcmlc001b_SXL039  ,SXL管理
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************
Private Function GetsSXLIDINBlkid(sxl() As DBDRV_scmzc_fcmlc001b_SXL039) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim i           As Long
    Dim j           As Long
    Dim rs          As OraDynaset
    Dim lngRecCnt   As Long
    Dim intSXLCnt   As Integer
    Dim sSXLID      As String

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_sSql.bas -- Function GetsSXLIDINBlkid"

    GetsSXLIDINBlkid = FUNCTION_RETURN_SUCCESS

    intSXLCnt = UBound(sxl)
    ReDim WFJudgExecOkFlag(intSXLCnt) As Boolean    'WF総合判定実行可能フラグ

    '入庫待ちブロックを含むSXLはグレー表示
    sSQL = "select distinct SXL.sSXLID "
    sSQL = sSQL & "from TBCME042 SXL, TBCME040 BLK "
    sSQL = sSQL & "where SXL.DELCLS='0' and SXL.NOWPROC='CW750'"
    sSQL = sSQL & "  and BLK.CRYNUM=SXL.CRYNUM and BLK.INGOTPOS>=0"
    sSQL = sSQL & "  and SXL.INGOTPOS<BLK.INGOTPOS+BLK.LENGTH"
    sSQL = sSQL & "  and SXL.INGOTPOS+SXL.LENGTH>BLK.INGOTPOS"
    sSQL = sSQL & "  and not exists (select LOTID from TBCMY011 where LOTID=BLK.BLOCKID)"
    sSQL = sSQL & "  and not exists (select LOTID from TBCMY012 where LOTID=BLK.BLOCKID and ALLSCRAP='Y')"

    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_DEFAULT)
    lngRecCnt = rs.RecordCount

    For i = 1 To intSXLCnt
        WFJudgExecOkFlag(i) = True
    Next

    For i = 1 To lngRecCnt
        sSXLID = rs("sSXLID")
        For j = 1 To intSXLCnt
            If sSXLID = sxl(j).SXLIDCA Then
                WFJudgExecOkFlag(j) = False
                Exit For
            End If
        Next
        rs.MoveNext
    Next
    rs.Close
    Set rs = Nothing
    Debug.Print sSQL

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error sSql ======"
    Debug.Print sSQL
    GetsSXLIDINBlkid = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'***************************************************************************************
'*    関数名        : MeasRsltCheck
'*
'*    処理概要      : 1.測定評価結果受信確認
'*
'*    パラメータ    : 変数名        ,IO ,型                            ,説明
'*                    SXL           ,O  ,DBDRV_scmzc_fcmlc001b_SXL039  ,SXL管理
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'***************************************************************************************
Private Function MeasRsltCheck(sxl() As DBDRV_scmzc_fcmlc001b_SXL039) As FUNCTION_RETURN
    Dim sSQL            As String
    Dim rs              As OraDynaset
    Dim intSXLCnt       As Integer
    Dim udtSokutei()    As typ_TBCMY013
    Dim intGDcnt        As Integer           'GD実績ﾚｺｰﾄﾞ数
''SPV9点対応
    Dim intSPVCnt       As Integer          'SPV実績レコード数
    Dim c0              As Integer
    Dim c1              As Integer
    Dim c2              As Integer
    Dim blChangeFlag    As Boolean
    Dim blPassFlg       As Boolean
    Dim sSqlWhere       As String
#If SPEEDUP Then
    Dim sChkWF()        As String
    Dim i               As Integer
#End If
    Dim udtChkUp        As type_chkUP

    'エラーハンドラの設定
    'On Error GoTo proc_err
'    gErr.Push "s_cmbc039_sSql.bas -- Function MeasRsltCheck"

    MeasRsltCheck = FUNCTION_RETURN_SUCCESS

    '測定評価結果取得sSql変更
    intSXLCnt = UBound(sxl)
    If intSXLCnt = 0 Then GoTo proc_exit

    sSQL = sSQL & "select "
    sSQL = sSQL & "SXLIDCW, "
    sSQL = sSQL & "REPSMPLIDCW, "
    sSQL = sSQL & "decode(trim(RES_SPEC),'RES','1','0') RES, "
    sSQL = sSQL & "decode(trim(OI_SPEC),'OI','1','0') OI, "
    sSQL = sSQL & "decode(trim(BMD1_SPEC),'BMD1','1','0') BMD1, "
    sSQL = sSQL & "decode(trim(BMD2_SPEC),'BMD2','1','0') BMD2, "
    sSQL = sSQL & "decode(trim(BMD3_SPEC),'BMD3','1','0') BMD3, "
    sSQL = sSQL & "decode(trim(OSF1_SPEC),'OSF1','1','0') OSF1, "
    sSQL = sSQL & "decode(trim(OSF2_SPEC),'OSF2','1','0') OSF2, "
    sSQL = sSQL & "decode(trim(OSF3_SPEC),'OSF3','1','0') OSF3, "
'    sSql = sSql & "decode(trim(OSF4_SPEC),'OSF4','1','0') OSF4, " SIRD対応
'    sSQL = sSQL & "decode(trim(SIRD_SPEC),'SIRD','1','0') SIRD, " 2010/05/19 REP Y.Hitomi
    sSQL = sSQL & "decode(trim(SIRD_SPEC),'TENI','1','0') SIRD, "
    sSQL = sSQL & "decode(trim(DSOD_SPEC),'DSOD','1','0') DSOD, "
    sSQL = sSQL & "decode(trim(DZ_SPEC),'DZ','1','0') DZ, "
    sSQL = sSQL & "decode(trim(DOI1_SPEC),'DOI1','1','0') DOI1, "
    sSQL = sSQL & "decode(trim(DOI2_SPEC),'DOI2','1','0') DOI2, "
    sSQL = sSQL & "decode(trim(DOI3_SPEC),'DOI3','1','0') DOI3, "
    sSQL = sSQL & "decode(trim(AOI_SPEC),'AOI','1','0') AOI, "
    sSQL = sSQL & "decode(trim(GD_SPEC),'GD','1','0') GD, "
    sSQL = sSQL & "decode(trim(SPV_SPEC),'SPV','1','0') SPV, "
    sSQL = sSQL & "decode(trim(IJO),'1','1','0') IJO "

    'エピ先行評価追加対応
    sSQL = sSQL & ",decode(trim(BMD1E_SPEC),'BMD1','1','0') BMD1E, "
    sSQL = sSQL & "decode(trim(BMD2E_SPEC),'BMD2','1','0') BMD2E, "
    sSQL = sSQL & "decode(trim(BMD3E_SPEC),'BMD3','1','0') BMD3E, "
    sSQL = sSQL & "decode(trim(OSF1E_SPEC),'OSF1','1','0') OSF1E, "
    sSQL = sSQL & "decode(trim(OSF2E_SPEC),'OSF2','1','0') OSF2E, "
    sSQL = sSQL & "decode(trim(OSF3E_SPEC),'OSF3','1','0') OSF3E "
    sSQL = sSQL & "from "
    sSQL = sSQL & "(select SXLIDCW,REPSMPLIDCW,INPOSCW,WFSMPLIDL4CW from XSDCW where SXLIDCW in ("

    For c0 = 1 To intSXLCnt
        sSQL = sSQL & "'" & sxl(c0).SXLIDCA & "'"
        If c0 <> intSXLCnt Then sSQL = sSQL & ", " Else sSQL = sSQL & ") "
    Next c0

    sSQL = sSQL & "and LIVKCW = '0'), "
    sSQL = sSQL & "(select SAMPLEID,SPEC RES_SPEC from TBCMY013 where SPEC = 'RES') RES, "
    sSQL = sSQL & "(select SAMPLEID,SPEC OI_SPEC from TBCMY013 where SPEC = 'OI') OI, "
    sSQL = sSQL & "(select SAMPLEID,SPEC BMD1_SPEC from TBCMY013 where SPEC = 'BMD1') BMD1, "
    sSQL = sSQL & "(select SAMPLEID,SPEC BMD2_SPEC from TBCMY013 where SPEC = 'BMD2') BMD2, "
    sSQL = sSQL & "(select SAMPLEID,SPEC BMD3_SPEC from TBCMY013 where SPEC = 'BMD3') BMD3, "
    sSQL = sSQL & "(select SAMPLEID,SPEC OSF1_SPEC from TBCMY013 where SPEC = 'OSF1') OSF1, "
    sSQL = sSQL & "(select SAMPLEID,SPEC OSF2_SPEC from TBCMY013 where SPEC = 'OSF2') OSF2, "
    sSQL = sSQL & "(select SAMPLEID,SPEC OSF3_SPEC from TBCMY013 where SPEC = 'OSF3') OSF3, "
'    sSql = sSql & "(select SAMPLEID,SPEC OSF4_SPEC from TBCMY013 where SPEC = 'OSF4') OSF4, "  'SIRD対応
'    sSQL = sSQL & "(select SMPLNO,SPEC SIRD_SPEC from TBCMJ022 where SPEC = 'SIRD') SIRD, "    '2010/05/19 REP Y.Hitomi
    sSQL = sSQL & "(select SMPLNO,SPEC SIRD_SPEC from TBCMJ022 where SPEC = 'TENI') SIRD, "
    sSQL = sSQL & "(select SAMPLEID,SPEC DSOD_SPEC from TBCMY013 where SPEC = 'DSOD') DSOD, "
    sSQL = sSQL & "(select SAMPLEID,SPEC DZ_SPEC from TBCMY013 where SPEC = 'DZ') DZ, "
    sSQL = sSQL & "(select SAMPLEID,SPEC DOI1_SPEC from TBCMY013 where SPEC = 'DOI1') DOI1, "
    sSQL = sSQL & "(select SAMPLEID,SPEC DOI2_SPEC from TBCMY013 where SPEC = 'DOI2') DOI2, "
    sSQL = sSQL & "(select SAMPLEID,SPEC DOI3_SPEC from TBCMY013 where SPEC = 'DOI3') DOI3, "
    sSQL = sSQL & "(select SAMPLEID,SPEC AOI_SPEC from TBCMY013 where SPEC = 'AOI') AOI, "
    sSQL = sSQL & "(select SMPLNO,'GD' GD_SPEC from TBCMJ015 where HSFLG = '1') GD, "
    sSQL = sSQL & "(select SMPLNO,'SPV' SPV_SPEC from TBCMJ016 where HSFLG = '1') SPV, "
    sSQL = sSQL & "(select SAMPLEID,'1' IJO from TBCMY016) Y16 "

    'エピ先行評価追加対応
    sSQL = sSQL & ",(select SAMPLEID,SPEC BMD1E_SPEC from TBCMY022 where SPEC = 'BMD1') BMD1E, "
    sSQL = sSQL & "(select SAMPLEID,SPEC BMD2E_SPEC from TBCMY022 where SPEC = 'BMD2') BMD2E, "
    sSQL = sSQL & "(select SAMPLEID,SPEC BMD3E_SPEC from TBCMY022 where SPEC = 'BMD3') BMD3E, "
    sSQL = sSQL & "(select SAMPLEID,SPEC OSF1E_SPEC from TBCMY022 where SPEC = 'OSF1') OSF1E, "
    sSQL = sSQL & "(select SAMPLEID,SPEC OSF2E_SPEC from TBCMY022 where SPEC = 'OSF2') OSF2E, "
    sSQL = sSQL & "(select SAMPLEID,SPEC OSF3E_SPEC from TBCMY022 where SPEC = 'OSF3') OSF3E "
    sSQL = sSQL & "where REPSMPLIDCW = RES.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = OI.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = BMD1.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = BMD2.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = BMD3.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = OSF1.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = OSF2.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = OSF3.SAMPLEID(+) "
'    sSql = sSql & "and REPSMPLIDCW = OSF4.SAMPLEID(+) " SIRD_Y.Hitomi
    sSQL = sSQL & "and WFSMPLIDL4CW = SIRD.SMPLNO(+) "
    sSQL = sSQL & "and REPSMPLIDCW = DSOD.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = DZ.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = DOI1.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = DOI2.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = DOI3.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = AOI.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = GD.SMPLNO(+) "
    sSQL = sSQL & "and REPSMPLIDCW = SPV.SMPLNO(+) "
    sSQL = sSQL & "and REPSMPLIDCW = Y16.SAMPLEID(+) "

    'エピ先行評価追加対応
    sSQL = sSQL & "and REPSMPLIDCW = BMD1E.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = BMD2E.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = BMD3E.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = OSF1E.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = OSF2E.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = OSF3E.SAMPLEID(+) "
    sSQL = sSQL & "order by SXLIDCW,INPOSCW "

    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_DEFAULT)
    For c0 = 1 To rs.RecordCount
        For c1 = 1 To intSXLCnt
            For c2 = 1 To UBound(sxl(c1).WFSMP())
                If sxl(c1).WFSMP(c2).REPSMPLIDCW = rs("REPSMPLIDCW") Then
                    '受信ﾁｪｯｸFLGの初期化
                    With udtChkUp
                        .rs = "0"           ' 受信FLG（Rs)
                        .Oi = "0"           ' 受信FLG（Oi)
                        .B1 = "0"           ' 受信FLG（B1)
                        .B2 = "0"           ' 受信FLG（B2）
                        .B3 = "0"           ' 受信FLG（B3)
                        .L1 = "0"           ' 受信FLG（L1)
                        .L2 = "0"           ' 受信FLG（L2)
                        .L3 = "0"           ' 受信FLG（L3)
                        .L4 = "0"           ' 受信FLG（L4)
                        .DS = "0"           ' 受信FLG（DS)
                        .DZ = "0"           ' 受信FLG（DZ)
                        .sp = "0"           ' 受信FLG（SP)
                        .DO1 = "0"          ' 受信FLG（DO1)
                        .DO2 = "0"          ' 受信FLG（DO2)
                        .DO3 = "0"          ' 受信FLG（DO3)
                        .OT1 = "0"          ' 受信FLG (OT2)
                        .OT2 = "0"          ' 受信FLG (OT1)
                        .AOI = "0"          ' 受信FLG (AOi)
                        .GD = "0"           ' 受信FLG (GD)

                        'エピ先行評価追加対応
                        .B1E = "0"           ' 受信FLG（B1E)
                        .B2E = "0"           ' 受信FLG（B2E）
                        .B3E = "0"           ' 受信FLG（B3E)
                        .L1E = "0"           ' 受信FLG（L1E)
                        .L2E = "0"           ' 受信FLG（L2E)
                        .L3E = "0"           ' 受信FLG（L3E)
                    End With

                    blChangeFlag = False
                    With sxl(c1).WFSMP(c2)
                        If rs("RES") = "1" Then         'RES
                            If (.REPSMPLIDCW = .WFSMPLIDRSCW) And (.WFRESRS1CW = "0") Then
                                .WFRESRS1CW = "1"
                                blChangeFlag = True
                                udtChkUp.rs = "1"
                            End If
                        End If
                        If rs("OI") = "1" Then          'OI
                            If (.REPSMPLIDCW = .WFSMPLIDOICW) And (.WFRESOICW = "0") Then
                                .WFRESOICW = "1"
                                blChangeFlag = True
                                udtChkUp.Oi = "1"
                            End If
                        End If
                        If rs("BMD1") = "1" Then        'BMD1
                            If (.REPSMPLIDCW = .WFSMPLIDB1CW) And (.WFRESB1CW = "0") Then
                                .WFRESB1CW = "1"
                                blChangeFlag = True
                                udtChkUp.B1 = "1"
                            End If
                        End If
                        If rs("BMD2") = "1" Then        'BMD2
                            If (.REPSMPLIDCW = .WFSMPLIDB2CW) And (.WFRESB2CW = "0") Then
                                .WFRESB2CW = "1"
                                blChangeFlag = True
                                udtChkUp.B2 = "1"
                            End If
                        End If
                        If rs("BMD3") = "1" Then        'BMD3
                            If (.REPSMPLIDCW = .WFSMPLIDB3CW) And (.WFRESB3CW = "0") Then
                                .WFRESB3CW = "1"
                                blChangeFlag = True
                                udtChkUp.B3 = "1"
                            End If
                        End If
                        If rs("OSF1") = "1" Then        'OSF1
                            If (.REPSMPLIDCW = .WFSMPLIDL1CW) And (.WFRESL1CW = "0") Then
                                .WFRESL1CW = "1"
                                blChangeFlag = True
                                udtChkUp.L1 = "1"
                            End If
                        End If
                        If rs("OSF2") = "1" Then        'OSF2
                            If (.REPSMPLIDCW = .WFSMPLIDL2CW) And (.WFRESL2CW = "0") Then
                                .WFRESL2CW = "1"
                                blChangeFlag = True
                                udtChkUp.L2 = "1"
                            End If
                        End If
                        If rs("OSF3") = "1" Then        'OSF3
                            If (.REPSMPLIDCW = .WFSMPLIDL3CW) And (.WFRESL3CW = "0") Then
                                .WFRESL3CW = "1"
                                blChangeFlag = True
                                udtChkUp.L3 = "1"
                            End If
                        End If
'                        If rs("OSF4") = "1" Then        'OSF4 SIRD_Y.Hitomi
'                            If (.REPSMPLIDCW = .WFSMPLIDL4CW) And (.WFRESL4CW = "0") Then
'                                .WFRESL4CW = "1"
'                                blChangeFlag = True
'                                udtChkUp.L4 = "1"
'                            End If
'                        End If
                        If rs("SIRD") = "1" Then        'SIRD
                            If (.WFRESL4CW = "0") Then
                                .WFRESL4CW = "1"
                                blChangeFlag = True
                                udtChkUp.L4 = "1"
                            End If
                        End If
                        If rs("DSOD") = "1" Then        'DSOD
                            If (.REPSMPLIDCW = .WFSMPLIDDSCW) And (.WFRESDSCW = "0") Then
                                .WFRESDSCW = "1"
                                blChangeFlag = True
                                udtChkUp.DS = "1"
                            End If
                        End If
                        If rs("DZ") = "1" Then          'DZ
                            If (.REPSMPLIDCW = .WFSMPLIDDZCW) And (.WFRESDZCW = "0") Then
                                .WFRESDZCW = "1"
                                blChangeFlag = True
                                udtChkUp.DZ = "1"
                            End If
                        End If
                        If rs("DOI1") = "1" Then        'DOI1
                            If (.REPSMPLIDCW = .WFSMPLIDDO1CW) And (.WFRESDO1CW = "0") Then
                                .WFRESDO1CW = "1"
                                blChangeFlag = True
                                udtChkUp.DO1 = "1"
                            End If
                        End If
                        If rs("DOI2") = "1" Then        'DOI2
                            If (.REPSMPLIDCW = .WFSMPLIDDO2CW) And (.WFRESDO2CW = "0") Then
                                .WFRESDO2CW = "1"
                                blChangeFlag = True
                                udtChkUp.DO2 = "1"
                            End If
                        End If
                        If rs("DOI3") = "1" Then        'DOI3
                            If (.REPSMPLIDCW = .WFSMPLIDDO3CW) And (.WFRESDO3CW = "0") Then
                                .WFRESDO3CW = "1"
                                blChangeFlag = True
                                udtChkUp.DO3 = "1"
                            End If
                        End If
                        If rs("AOI") = "1" Then         'AOI
                            If (.REPSMPLIDCW = .WFSMPLIDAOICW) And (.WFRESAOICW = "0") Then
                                .WFRESAOICW = "1"
                                blChangeFlag = True
                                udtChkUp.AOI = "1"
                            End If
                        End If
                        If rs("GD") = "1" Then          'GD
                            If (.REPSMPLIDCW = .WFSMPLIDGDCW) And (.WFRESGDCW = "0") Then
                                .WFRESGDCW = "1"
                                blChangeFlag = True
                                udtChkUp.GD = "1"
                            End If
                        End If
                        If rs("SPV") = "1" Then         'SPV
                            If (.REPSMPLIDCW = .WFSMPLIDSPCW) And (.WFRESSPCW = "0") Then
                                .WFRESSPCW = "1"
                                blChangeFlag = True
                                udtChkUp.sp = "1"
                            End If
                        End If

                        'エピ先行評価追加対応
                        If rs("BMD1E") = "1" Then        'BMD1E
                            If (.REPSMPLIDCW = .EPSMPLIDB1CW) And (.EPRESB1CW = "0") Then
                                .EPRESB1CW = "1"
                                blChangeFlag = True
                                udtChkUp.B1E = "1"
                            End If
                        End If
                        If rs("BMD2E") = "1" Then        'BMD2E
                            If (.REPSMPLIDCW = .EPSMPLIDB2CW) And (.EPRESB2CW = "0") Then
                                .EPRESB2CW = "1"
                                blChangeFlag = True
                                udtChkUp.B2E = "1"
                            End If
                        End If
                        If rs("BMD3E") = "1" Then        'BMD3E
                            If (.REPSMPLIDCW = .EPSMPLIDB3CW) And (.EPRESB3CW = "0") Then
                                .EPRESB3CW = "1"
                                blChangeFlag = True
                                udtChkUp.B3E = "1"
                            End If
                        End If
                        If rs("OSF1E") = "1" Then        'OSF1E
                            If (.REPSMPLIDCW = .EPSMPLIDL1CW) And (.EPRESL1CW = "0") Then
                                .EPRESL1CW = "1"
                                blChangeFlag = True
                                udtChkUp.L1E = "1"
                            End If
                        End If
                        If rs("OSF2E") = "1" Then        'OSF2E
                            If (.REPSMPLIDCW = .EPSMPLIDL2CW) And (.EPRESL2CW = "0") Then
                                .EPRESL2CW = "1"
                                blChangeFlag = True
                                udtChkUp.L2E = "1"
                            End If
                        End If
                        If rs("OSF3E") = "1" Then        'OSF3E
                            If (.REPSMPLIDCW = .EPSMPLIDL3CW) And (.EPRESL3CW = "0") Then
                                .EPRESL3CW = "1"
                                blChangeFlag = True
                                udtChkUp.L3E = "1"
                            End If
                        End If
                        If blChangeFlag Then
                            '同サンプルIDの実績付け
                            If WfSmp_Upd_SmplID(.XTALCW, .REPSMPLIDCW, udtChkUp) = FUNCTION_RETURN_FAILURE Then
                                MeasRsltCheck = FUNCTION_RETURN_FAILURE
                            End If
                            'Add 2010/01/21 SIRD対応 Y.Hitomi
                            If WfSmp_Upd_SmplID_SD(.XTALCW, .WFSMPLIDL4CW, udtChkUp) = FUNCTION_RETURN_FAILURE Then
                                MeasRsltCheck = FUNCTION_RETURN_FAILURE
                            End If
                        '抜試異常が登録されている
                        ElseIf rs("IJO") = "1" Then
                            'エピ先行評価追加対応
                            If (.WFINDRSCW <> "0" And .WFRESRS1CW <> "2") Or _
                               (.WFINDOICW <> "0" And .WFRESOICW <> "2") Or _
                               (.WFINDB1CW <> "0" And .WFRESB1CW <> "2") Or _
                               (.WFINDB2CW <> "0" And .WFRESB2CW <> "2") Or _
                               (.WFINDB3CW <> "0" And .WFRESB3CW <> "2") Or _
                               (.WFINDL1CW <> "0" And .WFRESL1CW <> "2") Or _
                               (.WFINDL2CW <> "0" And .WFRESL2CW <> "2") Or _
                               (.WFINDL3CW <> "0" And .WFRESL3CW <> "2") Or _
                               (.WFINDL4CW <> "0" And .WFRESL4CW <> "2") Or _
                               (.WFINDDSCW <> "0" And .WFRESDSCW <> "2") Or _
                               (.WFINDDZCW <> "0" And .WFRESDZCW <> "2") Or _
                               (.WFINDSPCW <> "0" And .WFRESSPCW <> "2") Or _
                               (.WFINDDO1CW <> "0" And .WFRESDO1CW <> "2") Or _
                               (.WFINDDO2CW <> "0" And .WFRESDO2CW <> "2") Or _
                               (.WFINDDO3CW <> "0" And .WFRESDO3CW <> "2") Or _
                               (.WFINDAOICW <> "0" And .WFRESAOICW <> "2") Or _
                               (.WFINDGDCW <> "0" And .WFHSGDCW <> "1" And .WFRESGDCW <> "2") Or _
                               (.EPINDB1CW <> "0" And .EPRESB1CW <> "2") Or _
                               (.EPINDB2CW <> "0" And .EPRESB2CW <> "2") Or _
                               (.EPINDB3CW <> "0" And .EPRESB3CW <> "2") Or _
                               (.EPINDL1CW <> "0" And .EPRESL1CW <> "2") Or _
                               (.EPINDL2CW <> "0" And .EPRESL2CW <> "2") Or _
                               (.EPINDL3CW <> "0" And .EPRESL3CW <> "2") Then

                                If .WFINDRSCW <> "0" Then .WFRESRS1CW = "2"
                                If .WFINDOICW <> "0" Then .WFRESOICW = "2"
                                If .WFINDB1CW <> "0" Then .WFRESB1CW = "2"
                                If .WFINDB2CW <> "0" Then .WFRESB2CW = "2"
                                If .WFINDB3CW <> "0" Then .WFRESB3CW = "2"
                                If .WFINDL1CW <> "0" Then .WFRESL1CW = "2"
                                If .WFINDL2CW <> "0" Then .WFRESL2CW = "2"
                                If .WFINDL3CW <> "0" Then .WFRESL3CW = "2"
                                If .WFINDL4CW <> "0" Then .WFRESL4CW = "2"
                                If .WFINDDSCW <> "0" Then .WFRESDSCW = "2"
                                If .WFINDDZCW <> "0" Then .WFRESDZCW = "2"
                                If .WFINDSPCW <> "0" Then .WFRESSPCW = "2"
                                If .WFINDDO1CW <> "0" Then .WFRESDO1CW = "2"
                                If .WFINDDO2CW <> "0" Then .WFRESDO2CW = "2"
                                If .WFINDDO3CW <> "0" Then .WFRESDO3CW = "2"
                                If .WFINDAOICW <> "0" Then .WFRESAOICW = "2"
                                If .WFINDGDCW <> "0" And .WFHSGDCW <> "1" Then .WFRESGDCW = "2"
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -s-
                                If .EPINDB1CW <> "0" Then .EPRESB1CW = "2"
                                If .EPINDB2CW <> "0" Then .EPRESB2CW = "2"
                                If .EPINDB3CW <> "0" Then .EPRESB3CW = "2"
                                If .EPINDL1CW <> "0" Then .EPRESL1CW = "2"
                                If .EPINDL2CW <> "0" Then .EPRESL2CW = "2"
                                If .EPINDL3CW <> "0" Then .EPRESL3CW = "2"
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -e-
                                If WfSmp_Upd(sxl(c1).WFSMP(c2)) = FUNCTION_RETURN_FAILURE Then
                                    MeasRsltCheck = FUNCTION_RETURN_FAILURE
                                End If
                            End If
                        End If
                    End With
                    GoTo LoopNext
                End If
            Next c2
        Next c1
LoopNext:
        rs.MoveNext
    Next c0
    rs.Close
    '測定評価結果取得sSql変更　06/02/07 ooba END =============================================>

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error sSql ======"
    Debug.Print sSQL
    MeasRsltCheck = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'Add Start 2011/06/16 Y.Hitomi
'***************************************************************************************
'*    関数名        : MeasRsltCheck1
'*
'*    処理概要      : 1.測定評価結果受信確認
'*
'*    パラメータ    : 変数名        ,IO ,型                            ,説明
'*                    SXL           ,O  ,DBDRV_scmzc_fcmlc001b_SXL039  ,SXL管理
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'***************************************************************************************
Private Function MeasRsltCheck1(sxl() As DBDRV_scmzc_fcmlc001b_SXL039) As FUNCTION_RETURN
    Dim sSQL            As String
    Dim rs              As OraDynaset
    Dim intSXLCnt       As Integer
    Dim udtSokutei()    As typ_TBCMY013
    Dim intGDcnt        As Integer           'GD実績ﾚｺｰﾄﾞ数
''SPV9点対応
    Dim intSPVCnt       As Integer          'SPV実績レコード数
    Dim c0              As Integer
    Dim c1              As Integer
    Dim c2              As Integer
    Dim blChangeFlag    As Boolean
    Dim blPassFlg       As Boolean
    Dim sSqlWhere       As String
#If SPEEDUP Then
    Dim sChkWF()        As String
    Dim i               As Integer
#End If
    Dim udtChkUp        As type_chkUP

    'エラーハンドラの設定
    'On Error GoTo proc_err
'    gErr.Push "s_cmbc039_sSql.bas -- Function MeasRsltCheck"

    MeasRsltCheck1 = FUNCTION_RETURN_SUCCESS

    '測定評価結果取得sSql変更
    intSXLCnt = UBound(sxl)
    If intSXLCnt = 0 Then GoTo proc_exit

    sSQL = sSQL & "select "
    sSQL = sSQL & "SXLIDCW, "
    sSQL = sSQL & "REPSMPLIDCW, "
    sSQL = sSQL & "decode(trim(SIRD_SPEC),'TENI','1','0') SIRD "
    sSQL = sSQL & "from "
    sSQL = sSQL & "(select SXLIDCW,REPSMPLIDCW,INPOSCW,WFSMPLIDL4CW from XSDCW where SXLIDCW in ("

    For c0 = 1 To intSXLCnt
        sSQL = sSQL & "'" & sxl(c0).SXLIDCA & "'"
        If c0 <> intSXLCnt Then sSQL = sSQL & ", " Else sSQL = sSQL & ") "
    Next c0

    sSQL = sSQL & "and LIVKCW = '0'), "
    sSQL = sSQL & "(select SMPLNO,SPEC SIRD_SPEC from TBCMJ022 where SPEC = 'TENI') SIRD "
    sSQL = sSQL & "where "

    sSQL = sSQL & "WFSMPLIDL4CW = SIRD.SMPLNO(+) "
    sSQL = sSQL & "order by SXLIDCW,INPOSCW "

    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_DEFAULT)
    For c0 = 1 To rs.RecordCount
        For c1 = 1 To intSXLCnt
            For c2 = 1 To UBound(sxl(c1).WFSMP())
                If sxl(c1).WFSMP(c2).REPSMPLIDCW = rs("REPSMPLIDCW") Then
                    '受信ﾁｪｯｸFLGの初期化
                    With udtChkUp
                        .rs = "0"           ' 受信FLG（Rs)
                        .Oi = "0"           ' 受信FLG（Oi)
                        .B1 = "0"           ' 受信FLG（B1)
                        .B2 = "0"           ' 受信FLG（B2）
                        .B3 = "0"           ' 受信FLG（B3)
                        .L1 = "0"           ' 受信FLG（L1)
                        .L2 = "0"           ' 受信FLG（L2)
                        .L3 = "0"           ' 受信FLG（L3)
                        .L4 = "0"           ' 受信FLG（L4)
                        .DS = "0"           ' 受信FLG（DS)
                        .DZ = "0"           ' 受信FLG（DZ)
                        .sp = "0"           ' 受信FLG（SP)
                        .DO1 = "0"          ' 受信FLG（DO1)
                        .DO2 = "0"          ' 受信FLG（DO2)
                        .DO3 = "0"          ' 受信FLG（DO3)
                        .OT1 = "0"          ' 受信FLG (OT2)
                        .OT2 = "0"          ' 受信FLG (OT1)
                        .AOI = "0"          ' 受信FLG (AOi)
                        .GD = "0"           ' 受信FLG (GD)

                        'エピ先行評価追加対応
                        .B1E = "0"           ' 受信FLG（B1E)
                        .B2E = "0"           ' 受信FLG（B2E）
                        .B3E = "0"           ' 受信FLG（B3E)
                        .L1E = "0"           ' 受信FLG（L1E)
                        .L2E = "0"           ' 受信FLG（L2E)
                        .L3E = "0"           ' 受信FLG（L3E)
                    End With

                    blChangeFlag = False
                    With sxl(c1).WFSMP(c2)
                        If rs("SIRD") = "1" Then        'SIRD
                            If (.WFRESL4CW = "0") Then
                                .WFRESL4CW = "1"
                                blChangeFlag = True
                                udtChkUp.L4 = "1"
                            End If
                        End If
                        If blChangeFlag Then
                            If WfSmp_Upd_SmplID_SD(.XTALCW, .WFSMPLIDL4CW, udtChkUp) = FUNCTION_RETURN_FAILURE Then
                                MeasRsltCheck1 = FUNCTION_RETURN_FAILURE
                            End If
                        End If
                    End With
                    GoTo LoopNext
                End If
            Next c2
        Next c1
LoopNext:
        rs.MoveNext
    Next c0
    rs.Close
    '測定評価結果取得sSql変更　06/02/07 ooba END =============================================>

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error sSql ======"
    Debug.Print sSQL
    MeasRsltCheck1 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'Add End 2011/06/16 Y.Hitomi

'*******************************************************************************
'*    関数名        : WfSmp_Upd
'*
'*    処理概要      : 1.WFサンプル管理アップデート
'*
'*    パラメータ    : 変数名        ,IO ,型              ,説明
'*                    WFSMP         ,I  ,typ_XSDCW       ,新サンプル管理（SXL）
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Function WfSmp_Upd(WFSMP As typ_XSDCW) As FUNCTION_RETURN
    Dim sSQL As String

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function WfSmp_Upd"

    WfSmp_Upd = FUNCTION_RETURN_SUCCESS

    With WFSMP
        sSQL = sSQL & "update XSDCW set "
        sSQL = sSQL & "REPSMPLIDCW='" & .REPSMPLIDCW & "',"           'サンプルID
        sSQL = sSQL & "HINBCW='" & .HINBCW & "',"                     '品番
        sSQL = sSQL & "REVNUMCW=" & .REVNUMCW & ","                   '製品番号改訂番号
        sSQL = sSQL & "FACTORYCW='" & .FACTORYCW & "',"               '工場
        sSQL = sSQL & "OPECW='" & .OPECW & "',"                       '操業条件
        sSQL = sSQL & "KTKBNCW='" & .KTKBNCW & "',"                   '確定区分
        sSQL = sSQL & "WFINDRSCW='" & .WFINDRSCW & "',"               '状態FLG(Rs)
        sSQL = sSQL & "WFRESRS1CW='" & .WFRESRS1CW & "',"                '実績FLG(Rs)
        sSQL = sSQL & "WFINDOICW='" & .WFINDOICW & "',"               '状態FLG(Oi)
        sSQL = sSQL & "WFRESOICW='" & .WFRESOICW & "',"                 '実績FLG(Oi)
        sSQL = sSQL & "WFINDB1CW='" & .WFINDB1CW & "',"               '状態FLG(B1)
        sSQL = sSQL & "WFRESB1CW='" & .WFRESB1CW & "',"                 '実績FLG(B1)
        sSQL = sSQL & "WFINDB2CW='" & .WFINDB2CW & "',"               '状態FLG(B2)
        sSQL = sSQL & "WFRESB2CW='" & .WFRESB2CW & "',"                 '実績FLG(B2)
        sSQL = sSQL & "WFINDB3CW='" & .WFINDB3CW & "',"               '状態FLG(B3)
        sSQL = sSQL & "WFRESB3CW='" & .WFRESB3CW & "',"                 '実績FLG(B3)
        sSQL = sSQL & "WFINDL1CW='" & .WFINDL1CW & "',"               '状態FLG(L1)
        sSQL = sSQL & "WFRESL1CW='" & .WFRESL1CW & "',"                 '実績FLG(L1)
        sSQL = sSQL & "WFINDL2CW='" & .WFINDL2CW & "',"               '状態FLG(L2)
        sSQL = sSQL & "WFRESL2CW='" & .WFRESL2CW & "',"                 '実績FLG(L2)
        sSQL = sSQL & "WFINDL3CW='" & .WFINDL3CW & "',"               '状態FLG(L3)
        sSQL = sSQL & "WFRESL3CW='" & .WFRESL3CW & "',"                 '実績FLG(L3)
        sSQL = sSQL & "WFINDL4CW='" & .WFINDL4CW & "',"               '状態FLG(L4)
        sSQL = sSQL & "WFRESL4CW='" & .WFRESL4CW & "',"                 '実績FLG(L4)
        sSQL = sSQL & "WFINDDSCW='" & .WFINDDSCW & "',"               '状態FLG(DS)
        sSQL = sSQL & "WFRESDSCW='" & .WFRESDSCW & "',"                 '実績FLG(DS)
        sSQL = sSQL & "WFINDDZCW='" & .WFINDDZCW & "',"               '状態FLG(DZ)
        sSQL = sSQL & "WFRESDZCW='" & .WFRESDZCW & "',"                 '実績FLG(DZ)
        sSQL = sSQL & "WFINDSPCW='" & .WFINDSPCW & "',"               '状態FLG(SP)
        sSQL = sSQL & "WFRESSPCW='" & .WFRESSPCW & "',"                 '実績FLG(SP)
        sSQL = sSQL & "WFINDDO1CW='" & .WFINDDO1CW & "',"             '状態FLG(DO1)
        sSQL = sSQL & "WFRESDO1CW='" & .WFRESDO1CW & "', "              '実績FLG(DO1)
        sSQL = sSQL & "WFINDDO2CW='" & .WFINDDO2CW & "',"             '状態FLG(DO2)
        sSQL = sSQL & "WFRESDO2CW='" & .WFRESDO2CW & "',"               '実績FLG(DO2)
        sSQL = sSQL & "WFINDDO3CW='" & .WFINDDO3CW & "',"             '状態FLG(DO3)
        sSQL = sSQL & "WFRESDO3CW='" & .WFRESDO3CW & "',"               '実績FLG(DO3)
        ''残存酸素追加　03/12/15 ooba
        sSQL = sSQL & "WFINDAOICW='" & .WFINDAOICW & "',"             '状態FLG(AOi)
        sSQL = sSQL & "WFRESAOICW='" & .WFRESAOICW & "',"               '実績FLG(AOi)
        'GD追加　05/02/04 ooba
        sSQL = sSQL & "WFINDGDCW='" & .WFINDGDCW & "',"               '状態FLG(GD)
        sSQL = sSQL & "WFRESGDCW='" & .WFRESGDCW & "',"               '実績FLG(GD)
        '--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -s-
        sSQL = sSQL & "EPINDB1CW='" & .EPINDB1CW & "',"               '状態FLG(B1E)
        sSQL = sSQL & "EPRESB1CW='" & .EPRESB1CW & "',"                 '実績FLG(B1E)
        sSQL = sSQL & "EPINDB2CW='" & .EPINDB2CW & "',"               '状態FLG(B2E)
        sSQL = sSQL & "EPRESB2CW='" & .EPRESB2CW & "',"                 '実績FLG(B2E)
        sSQL = sSQL & "EPINDB3CW='" & .EPINDB3CW & "',"               '状態FLG(B3E)
        sSQL = sSQL & "EPRESB3CW='" & .EPRESB3CW & "',"                 '実績FLG(B3E)
        sSQL = sSQL & "EPINDL1CW='" & .EPINDL1CW & "',"               '状態FLG(L1E)
        sSQL = sSQL & "EPRESL1CW='" & .EPRESL1CW & "',"                 '実績FLG(L1E)
        sSQL = sSQL & "EPINDL2CW='" & .EPINDL2CW & "',"               '状態FLG(L2E)
        sSQL = sSQL & "EPRESL2CW='" & .EPRESL2CW & "',"                 '実績FLG(L2E)
        sSQL = sSQL & "EPINDL3CW='" & .EPINDL3CW & "',"               '状態FLG(L3E)
        sSQL = sSQL & "EPRESL3CW='" & .EPRESL3CW & "',"                 '実績FLG(L3E)
        '--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -e-
        sSQL = sSQL & "KDAYCW=sysdate,"
        sSQL = sSQL & "SNDKCW='0'"
        sSQL = sSQL & "where XTALCW='" & .XTALCW & "'"
        sSQL = sSQL & "and INPOSCW=" & .INPOSCW & ""
        sSQL = sSQL & "and SMPKBNCW='" & .SMPKBNCW & "'"
    End With

    If 0 >= OraDB.ExecuteSQL(sSQL) Then
        WfSmp_Upd = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    WfSmp_Upd = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'****************************************************************************************************
'*    関数名        : WfSmp_Upd_SmplID
'*
'*    処理概要      : 1.テーブル「XSDCW」の条件にあったレコードを更新する(ｻﾝﾌﾟﾙIDの実績ﾌﾗｸﾞ)
'*
'*    パラメータ    : 変数名       ,IO  ,型                      ,説明
'*                  :XTAL          ,I   ,String                  ,結晶番号
'*                  :WFSMPID       ,I   ,String                  ,サンプルID
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'****************************************************************************************************
Private Function WfSmp_Upd_SmplID(xtal As String, WFSMPID As String, chkUp As type_chkUP) As Integer
    Dim sSQL    As String
    Dim rs      As OraDynaset    'RecordSet

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function WfSmp_Upd_SmplID"

    WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE

    ''サンプル管理更新方法見直し　2004/02/06 TUKU START ===============================================>
    ''検査ごとの更新FLGの結果で検査ごとに実績FLGの更新を行うように変更

    ' ｻﾝﾌﾟﾙID(Rs)の更新
    If chkUp.rs = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESRS1CW='1', "                              ' 実績FLG1(Rs)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' 更新日付
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDRSCW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' ｻﾝﾌﾟﾙID(Oi)の更新
    If chkUp.Oi = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESOICW='1', "                               ' 実績FLG(Oi)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' 更新日付
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDOICW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' ｻﾝﾌﾟﾙID(B1)の更新
    If chkUp.B1 = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESB1CW='1', "                               ' 実績FLG(B1)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' 更新日付
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDB1CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' ｻﾝﾌﾟﾙID(B2)の更新
    If chkUp.B2 = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESB2CW='1', "                               ' 実績FLG(B2)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' 更新日付
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDB2CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' ｻﾝﾌﾟﾙID(B3)の更新
    If chkUp.B3 = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESB3CW='1', "                               ' 実績FLG(B3)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' 更新日付
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDB3CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' ｻﾝﾌﾟﾙID(L1)の更新
    If chkUp.L1 = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESL1CW='1', "                               ' 実績FLG(L1)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' 更新日付
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDL1CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' ｻﾝﾌﾟﾙID(L2)の更新
    If chkUp.L2 = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESL2CW='1', "                               ' 実績FLG(L2)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' 更新日付
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDL2CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' ｻﾝﾌﾟﾙID(L3)の更新
    If chkUp.L3 = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESL3CW='1', "                               ' 実績FLG(L3)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' 更新日付
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDL3CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
'Del 2010/01/21 SIRD対応 Y.Hitomi
'    ' ｻﾝﾌﾟﾙID(L4)の更新
'    If chkUp.L4 = "1" Then
'        sSql = "update XSDCW set "
'        sSql = sSql & "WFRESL4CW='1', "                               ' 実績FLG(L4)
'        sSql = sSql & "KDAYCW=sysdate "                               ' 更新日付
'        sSql = sSql & "WHERE XTALCW = '" & xtal & "'"
'        sSql = sSql & "      WFSMPLIDL4CW = '" & WFSMPID & "'"
'
'        If OraDB.ExecuteSQL(sSql) <= 0 Then
'            rs.Close
'            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
'            GoTo proc_exit
'        End If
'    End If

    ' ｻﾝﾌﾟﾙID(DS)の更新
    If chkUp.DS = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESDSCW='1', "                               ' 実績FLG(DS)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' 更新日付
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDDSCW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' ｻﾝﾌﾟﾙID(DZ)の更新
    If chkUp.DZ = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESDZCW='1', "                               ' 実績FLG(DZ)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' 更新日付
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDDZCW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' ｻﾝﾌﾟﾙID(SP)の更新
    If chkUp.sp = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESSPCW='1', "                               ' 実績FLG(SP)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' 更新日付
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDSPCW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' ｻﾝﾌﾟﾙID(DO1)の更新
    If chkUp.DO1 = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESDO1CW='1', "                              ' 実績FLG(DO1)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' 更新日付
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDDO1CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' ｻﾝﾌﾟﾙID(DO2)の更新
    If chkUp.DO2 = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESDO2CW='1', "                              ' 実績FLG(DO2)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' 更新日付
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDDO2CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' ｻﾝﾌﾟﾙID(DO3)の更新
    If chkUp.DO3 = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESDO3CW='1', "                              ' 実績FLG(DO3)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' 更新日付
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDDO3CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ''残存酸素追加　03/12/15 ooba START ===============================================>
    ' ｻﾝﾌﾟﾙID(AOi)の更新
    If chkUp.AOI = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESAOICW='1', "                              ' 実績FLG(AOi)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' 更新日付
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDAOICW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    ''残存酸素追加　03/12/15 ooba END =================================================>

    ''GD追加　05/02/04 ooba START =====================================================>
    If chkUp.GD = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESGDCW='1', "                               ' 実績FLG(GD)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' 更新日付
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDGDCW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    ''GD追加　05/02/04 ooba END =======================================================>

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    ' ｻﾝﾌﾟﾙID(B1E)の更新
    If chkUp.B1E = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "EPRESB1CW='1', "                               ' 実績FLG(B1)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' 更新日付
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      EPSMPLIDB1CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' ｻﾝﾌﾟﾙID(B2E)の更新
    If chkUp.B2E = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "EPRESB2CW='1', "                               ' 実績FLG(B2)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' 更新日付
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      EPSMPLIDB2CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' ｻﾝﾌﾟﾙID(B3E)の更新
    If chkUp.B3E = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "EPRESB3CW='1', "                               ' 実績FLG(B3)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' 更新日付
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      EPSMPLIDB3CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' ｻﾝﾌﾟﾙID(L1E)の更新
    If chkUp.L1E = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "EPRESL1CW='1', "                               ' 実績FLG(L1)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' 更新日付
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      EPSMPLIDL1CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' ｻﾝﾌﾟﾙID(L2E)の更新
    If chkUp.L2E = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "EPRESL2CW='1', "                               ' 実績FLG(L2)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' 更新日付
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      EPSMPLIDL2CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' ｻﾝﾌﾟﾙID(L3E)の更新
    If chkUp.L3E = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "EPRESL3CW='1', "                               ' 実績FLG(L3)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' 更新日付
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      EPSMPLIDL3CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

    ''サンプル管理更新方法見直し　2004/02/06 TUKU END ===============================================>

    WfSmp_Upd_SmplID = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
''Add 2010/01/21 SIRD評価対応 Y.Hitomi
'****************************************************************************************************
'*    関数名        : WfSmp_Upd_SmplID_SD
'*
'*    処理概要      : 1.テーブル「XSDCW」の条件にあったレコードを更新する(ｻﾝﾌﾟﾙIDの実績ﾌﾗｸﾞ) SIRD評価用
'*
'*    パラメータ    : 変数名       ,IO  ,型                      ,説明
'*                  :XTAL          ,I   ,String                  ,結晶番号
'*                  :WFSMPID       ,I   ,String                  ,サンプルID
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'****************************************************************************************************
Private Function WfSmp_Upd_SmplID_SD(xtal As String, WFSMPID As String, chkUp As type_chkUP) As Integer
    Dim sSQL    As String
    Dim rs      As OraDynaset    'RecordSet

    'エラーハンドラの設定
    On Error GoTo proc_err

    WfSmp_Upd_SmplID_SD = FUNCTION_RETURN_FAILURE


    If chkUp.L4 = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESL4CW='1', "                               ' 実績FLG(SIRD)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' 更新日付
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "'"
        sSQL = sSQL & " and WFSMPLIDL4CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID_SD = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If


    WfSmp_Upd_SmplID_SD = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    WfSmp_Upd_SmplID_SD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'*******************************************************************************
'*    関数名        : TBCMY016Check
'*
'*    処理概要      : 1.抜試異常チェック
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    WFSMPID       ,O  ,String   ,SXL管理
'*
'*    戻り値        : 使用していない
'*
'*******************************************************************************
Private Function TBCMY016Check(WFSMPID As String) As Integer
    Dim sSQL    As String
    Dim rs      As OraDynaset

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function TBCMY016Check"

    sSQL = "select "
    sSQL = sSQL & " SAMPLEID "
    sSQL = sSQL & " from "
    sSQL = sSQL & " TBCMY016 "
    sSQL = sSQL & " where "
    sSQL = sSQL & " SAMPLEID = '" & WFSMPID & "' "
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    TBCMY016Check = rs.RecordCount
    rs.Close

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    TBCMY016Check = -1
    gErr.HandleError
    Resume proc_exit
End Function

'************************************************************************************
'*    関数名        : DBDRV_scmzc_fcmlc001b_Disp
'*
'*    処理概要      : 1.WF総合判定 待ち一覧 表示用ＤＢドライバ
'*
'*    パラメータ    : 変数名       ,IO  ,型                            ,説明
'*              　　:inNowPorc     ,I   ,String                        ,入力用(工程)
'*              　　:SXL           ,O   ,DBDRV_scmzc_fcmlc001b_SXL039  ,SXL管理用
'*              　　:sErrMsg 　　　,O   ,String                        ,エラーメッセージ
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'************************************************************************************
Public Function DBDRV_scmzc_fcmlc001b_Disp(inNowPorc As String, _
                                            sxl() As DBDRV_scmzc_fcmlc001b_SXL039, _
                                            sErrMsg As String _
                                            ) As FUNCTION_RETURN
    Dim udtWKSXL()      As DBDRV_scmzc_fcmlc001b_SXL039
    Dim sSQL            As String
    Dim sDBName         As String
    Dim rs              As OraDynaset
    Dim lngRecCnt       As Long
    Dim sCryNumBuf      As String
    Dim iIngotPosBuf    As Integer
    Dim i               As Long
    Dim j               As Long
    Dim k               As Long
    Dim l               As Long
    Dim lngNullCnt      As Long
    Dim sNullSXLID      As String
    Dim intSxlCount     As Integer
    Dim rs2             As OraDynaset
    Dim lngCmpCnt       As Long             'サンプル数     'Add 2011/03/07 SMPK Miyata
    
    'エラーハンドラの設定
    'On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001b_Disp"
Debug.Print "1 " & Now & " SXL管理、XSDCBを取得 SQL実行"
    DBDRV_scmzc_fcmlc001b_Disp = FUNCTION_RETURN_SUCCESS

    ' SXL管理、XSDCBを取得（ビューを使用しないように変更)　TUKU  2003/10/9
''注意）待ち一覧用の為、仮修正です。
''↓追加START SXL管理（E042）→XSDCB機能移行 '06/1/5 SMP岡本
    sDBName = "(XSDCB)"

    '仕掛SXL取得SQL変更　06/02/07 ooba START ===============================================>
Debug.Print "2 " & Now & " SXLに対する新ｻﾝﾌﾟﾙ情報をセット"
    sSQL = sSQL & "select "
    sSQL = sSQL & "CRYNUM, "
    sSQL = sSQL & "INGOTPOS, "
    sSQL = sSQL & "LENGTH, "
    sSQL = sSQL & "SXLID, "
    sSQL = sSQL & "KRPROCCD, "
    sSQL = sSQL & "NOWPROC, "
    sSQL = sSQL & "LPKRPROCCD, "
    sSQL = sSQL & "LASTPASS, "
    sSQL = sSQL & "DELCLS, "
    sSQL = sSQL & "LSTATCLS, "
    sSQL = sSQL & "HOLDCLS, "
    sSQL = sSQL & "HINBAN, "
    sSQL = sSQL & "REVNUM, "
    sSQL = sSQL & "FACTORY, "
    sSQL = sSQL & "OPECOND, "
    sSQL = sSQL & "MAICB, "
    sSQL = sSQL & "REGDATE, "
    sSQL = sSQL & "UPDDATE, "
    sSQL = sSQL & "HOLDBCB, "             'ﾎｰﾙﾄﾞ区分　06/02/08 ooba
    sSQL = sSQL & "WFHOLDFLGCB, "         'WFﾎｰﾙﾄﾞ区分　06/02/08 ooba
    sSQL = sSQL & "KBLKFLGCB, "           '関連ﾌﾞﾛｯｸﾌﾗｸﾞ　08/01/31 ooba
    sSQL = sSQL & "PLANTCAT, "            '向先 07/09/05 SPK Tsutsumi Add
    sSQL = sSQL & "XTALCW, "
    sSQL = sSQL & "INPOSCW, "
    sSQL = sSQL & "nvl(TBKBNCW,'T') as TBKBNCW, "
    sSQL = sSQL & "SMPKBNCW, "
    sSQL = sSQL & "REPSMPLIDCW, "
    sSQL = sSQL & "HINBCW, "
    sSQL = sSQL & "REVNUMCW, "
    sSQL = sSQL & "FACTORYCW, "
    sSQL = sSQL & "OPECW, "
    sSQL = sSQL & "KTKBNCW, "
    sSQL = sSQL & "WFINDRSCW, "
    sSQL = sSQL & "WFINDOICW, "
    sSQL = sSQL & "WFINDB1CW, "
    sSQL = sSQL & "WFINDB2CW, "
    sSQL = sSQL & "WFINDB3CW, "
    sSQL = sSQL & "WFINDL1CW, "
    sSQL = sSQL & "WFINDL2CW, "
    sSQL = sSQL & "WFINDL3CW, "
    sSQL = sSQL & "WFINDL4CW, "
    sSQL = sSQL & "WFINDDSCW, "
    sSQL = sSQL & "WFINDDZCW, "
    sSQL = sSQL & "WFINDSPCW, "
    sSQL = sSQL & "WFINDDO1CW, "
    sSQL = sSQL & "WFINDDO2CW, "
    sSQL = sSQL & "WFINDDO3CW, "
    sSQL = sSQL & "WFINDAOICW, "
    sSQL = sSQL & "WFINDGDCW, "
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    sSQL = sSQL & "EPINDB1CW, "
    sSQL = sSQL & "EPINDB2CW, "
    sSQL = sSQL & "EPINDB3CW, "
    sSQL = sSQL & "EPINDL1CW, "
    sSQL = sSQL & "EPINDL2CW, "
    sSQL = sSQL & "EPINDL3CW, "
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
    sSQL = sSQL & "WFRESRS1CW, "
    sSQL = sSQL & "WFRESOICW, "
    sSQL = sSQL & "WFRESB1CW, "
    sSQL = sSQL & "WFRESB2CW, "
    sSQL = sSQL & "WFRESB3CW, "
    sSQL = sSQL & "WFRESL1CW, "
    sSQL = sSQL & "WFRESL2CW, "
    sSQL = sSQL & "WFRESL3CW, "
    sSQL = sSQL & "WFRESL4CW, "
    sSQL = sSQL & "WFRESDSCW, "
    sSQL = sSQL & "WFRESDZCW, "
    sSQL = sSQL & "WFRESSPCW, "
    sSQL = sSQL & "WFRESDO1CW, "
    sSQL = sSQL & "WFRESDO2CW, "
    sSQL = sSQL & "WFRESDO3CW, "
    sSQL = sSQL & "WFRESAOICW, "
    sSQL = sSQL & "WFRESGDCW, "
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    sSQL = sSQL & "EPRESB1CW, "
    sSQL = sSQL & "EPRESB2CW, "
    sSQL = sSQL & "EPRESB3CW, "
    sSQL = sSQL & "EPRESL1CW, "
    sSQL = sSQL & "EPRESL2CW, "
    sSQL = sSQL & "EPRESL3CW, "
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
    sSQL = sSQL & "WFSMPLIDRSCW, "
    sSQL = sSQL & "WFSMPLIDOICW, "
    sSQL = sSQL & "WFSMPLIDB1CW, "
    sSQL = sSQL & "WFSMPLIDB2CW, "
    sSQL = sSQL & "WFSMPLIDB3CW, "
    sSQL = sSQL & "WFSMPLIDL1CW, "
    sSQL = sSQL & "WFSMPLIDL2CW, "
    sSQL = sSQL & "WFSMPLIDL3CW, "
    sSQL = sSQL & "WFSMPLIDL4CW, "
    sSQL = sSQL & "WFSMPLIDDSCW, "
    sSQL = sSQL & "WFSMPLIDDZCW, "
    sSQL = sSQL & "WFSMPLIDSPCW, "
    sSQL = sSQL & "WFSMPLIDDO1CW, "
    sSQL = sSQL & "WFSMPLIDDO2CW, "
    sSQL = sSQL & "WFSMPLIDDO3CW, "
    sSQL = sSQL & "WFSMPLIDAOICW, "
    sSQL = sSQL & "WFSMPLIDGDCW, "
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    sSQL = sSQL & "EPSMPLIDB1CW, "
    sSQL = sSQL & "EPSMPLIDB2CW, "
    sSQL = sSQL & "EPSMPLIDB3CW, "
    sSQL = sSQL & "EPSMPLIDL1CW, "
    sSQL = sSQL & "EPSMPLIDL2CW, "
    sSQL = sSQL & "EPSMPLIDL3CW, "
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
    sSQL = sSQL & "WFHSGDCW, "
    sSQL = sSQL & "TDAYCW, "
    sSQL = sSQL & "KDAYCW "
    sSQL = sSQL & "from "

    sSQL = sSQL & "(select "
    sSQL = sSQL & "xtalcb as CRYNUM, "
    sSQL = sSQL & "inposcb as INGOTPOS, "
'    sSql = sSql & "rlencb as LENGTH, "
    sSQL = sSQL & "LENCB as LENGTH, "         '理論長さ→長さ　06/11/09 ooba
    sSQL = sSQL & "sxlidcb as SXLID, "
    sSQL = sSQL & "' ' as KRPROCCD, "
    sSQL = sSQL & "gnwkntcb as NOWPROC, "
    sSQL = sSQL & "' ' as LPKRPROCCD, "
    sSQL = sSQL & "newkntcb as LASTPASS, "
    sSQL = sSQL & "livkcb as DELCLS, "
    sSQL = sSQL & "lstccb as LSTATCLS, "
    sSQL = sSQL & "sholdclscb HOLDCLS, "
    sSQL = sSQL & "hinbcb as HINBAN, "
    sSQL = sSQL & "revnumcb as REVNUM, "
    sSQL = sSQL & "factorycb as FACTORY, "
    sSQL = sSQL & "opecb as OPECOND, "
    sSQL = sSQL & "MAICB, "
    sSQL = sSQL & "tdaycb as REGDATE, "
    sSQL = sSQL & "kdaycb as UPDDATE, "
    sSQL = sSQL & "HOLDBCB, "
    sSQL = sSQL & "WFHOLDFLGCB, "
    sSQL = sSQL & "KBLKFLGCB, "           '関連ﾌﾞﾛｯｸﾌﾗｸﾞ　08/01/31 ooba
    sSQL = sSQL & "PLANTCATCB  as PLANTCAT "
    sSQL = sSQL & "from XSDCB "
    sSQL = sSQL & "where GNWKNTCB = '" & inNowPorc & "' "
    sSQL = sSQL & "and livkcb = '0' "

    If sCmbMukesaki <> "ALL" Then
        sSQL = sSQL & "   AND PLANTCATCB      = '" & sCmbMukesaki & "'"
    End If
    sSQL = sSQL & "), "

    sSQL = sSQL & "(select "
    sSQL = sSQL & "SXLIDCW, "
    sSQL = sSQL & "XTALCW, "
    sSQL = sSQL & "INPOSCW, "
    sSQL = sSQL & "TBKBNCW, "
    sSQL = sSQL & "SMPKBNCW, "
    sSQL = sSQL & "REPSMPLIDCW, "
    sSQL = sSQL & "HINBCW, "
    sSQL = sSQL & "REVNUMCW, "
    sSQL = sSQL & "FACTORYCW, "
    sSQL = sSQL & "OPECW, "
    sSQL = sSQL & "KTKBNCW, "
    sSQL = sSQL & "WFINDRSCW, "
    sSQL = sSQL & "WFINDOICW, "
    sSQL = sSQL & "WFINDB1CW, "
    sSQL = sSQL & "WFINDB2CW, "
    sSQL = sSQL & "WFINDB3CW, "
    sSQL = sSQL & "WFINDL1CW, "
    sSQL = sSQL & "WFINDL2CW, "
    sSQL = sSQL & "WFINDL3CW, "
    sSQL = sSQL & "WFINDL4CW, "
    sSQL = sSQL & "WFINDDSCW, "
    sSQL = sSQL & "WFINDDZCW, "
    sSQL = sSQL & "WFINDSPCW, "
    sSQL = sSQL & "WFINDDO1CW, "
    sSQL = sSQL & "WFINDDO2CW, "
    sSQL = sSQL & "WFINDDO3CW, "
    sSQL = sSQL & "WFINDAOICW, "
    sSQL = sSQL & "WFINDGDCW, "
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    sSQL = sSQL & "EPINDB1CW, "
    sSQL = sSQL & "EPINDB2CW, "
    sSQL = sSQL & "EPINDB3CW, "
    sSQL = sSQL & "EPINDL1CW, "
    sSQL = sSQL & "EPINDL2CW, "
    sSQL = sSQL & "EPINDL3CW, "
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
    sSQL = sSQL & "WFRESRS1CW, "
    sSQL = sSQL & "WFRESOICW, "
    sSQL = sSQL & "WFRESB1CW, "
    sSQL = sSQL & "WFRESB2CW, "
    sSQL = sSQL & "WFRESB3CW, "
    sSQL = sSQL & "WFRESL1CW, "
    sSQL = sSQL & "WFRESL2CW, "
    sSQL = sSQL & "WFRESL3CW, "
    sSQL = sSQL & "WFRESL4CW, "
    sSQL = sSQL & "WFRESDSCW, "
    sSQL = sSQL & "WFRESDZCW, "
    sSQL = sSQL & "WFRESSPCW, "
    sSQL = sSQL & "WFRESDO1CW, "
    sSQL = sSQL & "WFRESDO2CW, "
    sSQL = sSQL & "WFRESDO3CW, "
    sSQL = sSQL & "WFRESAOICW, "
    sSQL = sSQL & "WFRESGDCW, "
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    sSQL = sSQL & "EPRESB1CW, "
    sSQL = sSQL & "EPRESB2CW, "
    sSQL = sSQL & "EPRESB3CW, "
    sSQL = sSQL & "EPRESL1CW, "
    sSQL = sSQL & "EPRESL2CW, "
    sSQL = sSQL & "EPRESL3CW, "
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
    sSQL = sSQL & "WFSMPLIDRSCW, "
    sSQL = sSQL & "WFSMPLIDOICW, "
    sSQL = sSQL & "WFSMPLIDB1CW, "
    sSQL = sSQL & "WFSMPLIDB2CW, "
    sSQL = sSQL & "WFSMPLIDB3CW, "
    sSQL = sSQL & "WFSMPLIDL1CW, "
    sSQL = sSQL & "WFSMPLIDL2CW, "
    sSQL = sSQL & "WFSMPLIDL3CW, "
    sSQL = sSQL & "WFSMPLIDL4CW, "
    sSQL = sSQL & "WFSMPLIDDSCW, "
    sSQL = sSQL & "WFSMPLIDDZCW, "
    sSQL = sSQL & "WFSMPLIDSPCW, "
    sSQL = sSQL & "WFSMPLIDDO1CW, "
    sSQL = sSQL & "WFSMPLIDDO2CW, "
    sSQL = sSQL & "WFSMPLIDDO3CW, "
    sSQL = sSQL & "WFSMPLIDAOICW, "
    sSQL = sSQL & "WFSMPLIDGDCW, "
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    sSQL = sSQL & "EPSMPLIDB1CW, "
    sSQL = sSQL & "EPSMPLIDB2CW, "
    sSQL = sSQL & "EPSMPLIDB3CW, "
    sSQL = sSQL & "EPSMPLIDL1CW, "
    sSQL = sSQL & "EPSMPLIDL2CW, "
    sSQL = sSQL & "EPSMPLIDL3CW, "
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
    sSQL = sSQL & "WFHSGDCW, "
    sSQL = sSQL & "TDAYCW, "
    sSQL = sSQL & "KDAYCW "
    sSQL = sSQL & "from XSDCW "
    sSQL = sSQL & "where LIVKCW = '0' "
'Chg Start 2011/03/07 SMPK Miyata
'    sSQL = sSQL & "and TBKBNCW = 'T' "
    sSQL = sSQL & "and (TBKBNCW = 'T' OR TBKBNCW = 'B')"
'Chg End   2011/03/07 SMPK Miyata
'Add Start 2011/03/07 SMPK Miyata
    sSQL = sSQL & "union all "
    sSQL = sSQL & "select "
    sSQL = sSQL & "SXLIDCW, "
    sSQL = sSQL & "XTALCW, "
    sSQL = sSQL & "INPOSCW, "
    sSQL = sSQL & "TBKBNCW, "
    sSQL = sSQL & "SMPKBNCW, "
    sSQL = sSQL & "REPSMPLIDCW, "
    sSQL = sSQL & "HINBCW, "
    sSQL = sSQL & "REVNUMCW, "
    sSQL = sSQL & "FACTORYCW, "
    sSQL = sSQL & "OPECW, "
    sSQL = sSQL & "KTKBNCW, "
    sSQL = sSQL & "WFINDRSCW, "
    sSQL = sSQL & "WFINDOICW, "
    sSQL = sSQL & "WFINDB1CW, "
    sSQL = sSQL & "WFINDB2CW, "
    sSQL = sSQL & "WFINDB3CW, "
    sSQL = sSQL & "WFINDL1CW, "
    sSQL = sSQL & "WFINDL2CW, "
    sSQL = sSQL & "WFINDL3CW, "
    sSQL = sSQL & "WFINDL4CW, "
    sSQL = sSQL & "WFINDDSCW, "
    sSQL = sSQL & "WFINDDZCW, "
    sSQL = sSQL & "WFINDSPCW, "
    sSQL = sSQL & "WFINDDO1CW, "
    sSQL = sSQL & "WFINDDO2CW, "
    sSQL = sSQL & "WFINDDO3CW, "
    sSQL = sSQL & "WFINDAOICW, "
    sSQL = sSQL & "WFINDGDCW, "
    sSQL = sSQL & "EPINDB1CW, "
    sSQL = sSQL & "EPINDB2CW, "
    sSQL = sSQL & "EPINDB3CW, "
    sSQL = sSQL & "EPINDL1CW, "
    sSQL = sSQL & "EPINDL2CW, "
    sSQL = sSQL & "EPINDL3CW, "
    sSQL = sSQL & "WFRESRS1CW, "
    sSQL = sSQL & "WFRESOICW, "
    sSQL = sSQL & "WFRESB1CW, "
    sSQL = sSQL & "WFRESB2CW, "
    sSQL = sSQL & "WFRESB3CW, "
    sSQL = sSQL & "WFRESL1CW, "
    sSQL = sSQL & "WFRESL2CW, "
    sSQL = sSQL & "WFRESL3CW, "
    sSQL = sSQL & "WFRESL4CW, "
    sSQL = sSQL & "WFRESDSCW, "
    sSQL = sSQL & "WFRESDZCW, "
    sSQL = sSQL & "WFRESSPCW, "
    sSQL = sSQL & "WFRESDO1CW, "
    sSQL = sSQL & "WFRESDO2CW, "
    sSQL = sSQL & "WFRESDO3CW, "
    sSQL = sSQL & "WFRESAOICW, "
    sSQL = sSQL & "WFRESGDCW, "
    sSQL = sSQL & "EPRESB1CW, "
    sSQL = sSQL & "EPRESB2CW, "
    sSQL = sSQL & "EPRESB3CW, "
    sSQL = sSQL & "EPRESL1CW, "
    sSQL = sSQL & "EPRESL2CW, "
    sSQL = sSQL & "EPRESL3CW, "
    sSQL = sSQL & "WFSMPLIDRSCW, "
    sSQL = sSQL & "WFSMPLIDOICW, "
    sSQL = sSQL & "WFSMPLIDB1CW, "
    sSQL = sSQL & "WFSMPLIDB2CW, "
    sSQL = sSQL & "WFSMPLIDB3CW, "
    sSQL = sSQL & "WFSMPLIDL1CW, "
    sSQL = sSQL & "WFSMPLIDL2CW, "
    sSQL = sSQL & "WFSMPLIDL3CW, "
    sSQL = sSQL & "WFSMPLIDL4CW, "
    sSQL = sSQL & "WFSMPLIDDSCW, "
    sSQL = sSQL & "WFSMPLIDDZCW, "
    sSQL = sSQL & "WFSMPLIDSPCW, "
    sSQL = sSQL & "WFSMPLIDDO1CW, "
    sSQL = sSQL & "WFSMPLIDDO2CW, "
    sSQL = sSQL & "WFSMPLIDDO3CW, "
    sSQL = sSQL & "WFSMPLIDAOICW, "
    sSQL = sSQL & "WFSMPLIDGDCW, "
    sSQL = sSQL & "EPSMPLIDB1CW, "
    sSQL = sSQL & "EPSMPLIDB2CW, "
    sSQL = sSQL & "EPSMPLIDB3CW, "
    sSQL = sSQL & "EPSMPLIDL1CW, "
    sSQL = sSQL & "EPSMPLIDL2CW, "
    sSQL = sSQL & "EPSMPLIDL3CW, "
    sSQL = sSQL & "WFHSGDCW, "
    sSQL = sSQL & "TDAYCW, "
    sSQL = sSQL & "KDAYCW "
    sSQL = sSQL & "from XSDCW_1 "
    sSQL = sSQL & "where LIVKCW = '0' "
    sSQL = sSQL & "and TBKBNCW = 'C'"
'Add End   2011/03/07 SMPK Miyata
    sSQL = sSQL & ") "
    sSQL = sSQL & "where SXLID = SXLIDCW(+) "

'Del Start 2011/03/07 SMPK Miyata
'    sSQL = sSQL & "union all "
'
'    sSQL = sSQL & "select "
'    sSQL = sSQL & "CRYNUM, "
'    sSQL = sSQL & "INGOTPOS, "
'    sSQL = sSQL & "LENGTH, "
'    sSQL = sSQL & "SXLID, "
'    sSQL = sSQL & "KRPROCCD, "
'    sSQL = sSQL & "NOWPROC, "
'    sSQL = sSQL & "LPKRPROCCD, "
'    sSQL = sSQL & "LASTPASS, "
'    sSQL = sSQL & "DELCLS, "
'    sSQL = sSQL & "LSTATCLS, "
'    sSQL = sSQL & "HOLDCLS, "
'    sSQL = sSQL & "HINBAN, "
'    sSQL = sSQL & "REVNUM, "
'    sSQL = sSQL & "FACTORY, "
'    sSQL = sSQL & "OPECOND, "
'    sSQL = sSQL & "MAICB, "
'    sSQL = sSQL & "REGDATE, "
'    sSQL = sSQL & "UPDDATE, "
'    sSQL = sSQL & "HOLDBCB, "             'ﾎｰﾙﾄﾞ区分　06/02/08 ooba
'    sSQL = sSQL & "WFHOLDFLGCB, "         'WFﾎｰﾙﾄﾞ区分　06/02/08 ooba
'    sSQL = sSQL & "KBLKFLGCB, "           '関連ﾌﾞﾛｯｸﾌﾗｸﾞ　08/01/31 ooba
'    sSQL = sSQL & "PLANTCAT, "            '向先 07/09/05 SPK Tsutsumi Add
'    sSQL = sSQL & "XTALCW, "
'    sSQL = sSQL & "INPOSCW, "
'    sSQL = sSQL & "nvl(TBKBNCW,'B') as TBKBNCW, "
'    sSQL = sSQL & "SMPKBNCW, "
'    sSQL = sSQL & "REPSMPLIDCW, "
'    sSQL = sSQL & "HINBCW, "
'    sSQL = sSQL & "REVNUMCW, "
'    sSQL = sSQL & "FACTORYCW, "
'    sSQL = sSQL & "OPECW, "
'    sSQL = sSQL & "KTKBNCW, "
'    sSQL = sSQL & "WFINDRSCW, "
'    sSQL = sSQL & "WFINDOICW, "
'    sSQL = sSQL & "WFINDB1CW, "
'    sSQL = sSQL & "WFINDB2CW, "
'    sSQL = sSQL & "WFINDB3CW, "
'    sSQL = sSQL & "WFINDL1CW, "
'    sSQL = sSQL & "WFINDL2CW, "
'    sSQL = sSQL & "WFINDL3CW, "
'    sSQL = sSQL & "WFINDL4CW, "
'    sSQL = sSQL & "WFINDDSCW, "
'    sSQL = sSQL & "WFINDDZCW, "
'    sSQL = sSQL & "WFINDSPCW, "
'    sSQL = sSQL & "WFINDDO1CW, "
'    sSQL = sSQL & "WFINDDO2CW, "
'    sSQL = sSQL & "WFINDDO3CW, "
'    sSQL = sSQL & "WFINDAOICW, "
'    sSQL = sSQL & "WFINDGDCW, "
''--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
'    sSQL = sSQL & "EPINDB1CW, "
'    sSQL = sSQL & "EPINDB2CW, "
'    sSQL = sSQL & "EPINDB3CW, "
'    sSQL = sSQL & "EPINDL1CW, "
'    sSQL = sSQL & "EPINDL2CW, "
'    sSQL = sSQL & "EPINDL3CW, "
''--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
'    sSQL = sSQL & "WFRESRS1CW, "
'    sSQL = sSQL & "WFRESOICW, "
'    sSQL = sSQL & "WFRESB1CW, "
'    sSQL = sSQL & "WFRESB2CW, "
'    sSQL = sSQL & "WFRESB3CW, "
'    sSQL = sSQL & "WFRESL1CW, "
'    sSQL = sSQL & "WFRESL2CW, "
'    sSQL = sSQL & "WFRESL3CW, "
'    sSQL = sSQL & "WFRESL4CW, "
'    sSQL = sSQL & "WFRESDSCW, "
'    sSQL = sSQL & "WFRESDZCW, "
'    sSQL = sSQL & "WFRESSPCW, "
'    sSQL = sSQL & "WFRESDO1CW, "
'    sSQL = sSQL & "WFRESDO2CW, "
'    sSQL = sSQL & "WFRESDO3CW, "
'    sSQL = sSQL & "WFRESAOICW, "
'    sSQL = sSQL & "WFRESGDCW, "
''--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
'    sSQL = sSQL & "EPRESB1CW, "
'    sSQL = sSQL & "EPRESB2CW, "
'    sSQL = sSQL & "EPRESB3CW, "
'    sSQL = sSQL & "EPRESL1CW, "
'    sSQL = sSQL & "EPRESL2CW, "
'    sSQL = sSQL & "EPRESL3CW, "
''--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
'    sSQL = sSQL & "WFSMPLIDRSCW, "
'    sSQL = sSQL & "WFSMPLIDOICW, "
'    sSQL = sSQL & "WFSMPLIDB1CW, "
'    sSQL = sSQL & "WFSMPLIDB2CW, "
'    sSQL = sSQL & "WFSMPLIDB3CW, "
'    sSQL = sSQL & "WFSMPLIDL1CW, "
'    sSQL = sSQL & "WFSMPLIDL2CW, "
'    sSQL = sSQL & "WFSMPLIDL3CW, "
'    sSQL = sSQL & "WFSMPLIDL4CW, "
'    sSQL = sSQL & "WFSMPLIDDSCW, "
'    sSQL = sSQL & "WFSMPLIDDZCW, "
'    sSQL = sSQL & "WFSMPLIDSPCW, "
'    sSQL = sSQL & "WFSMPLIDDO1CW, "
'    sSQL = sSQL & "WFSMPLIDDO2CW, "
'    sSQL = sSQL & "WFSMPLIDDO3CW, "
'    sSQL = sSQL & "WFSMPLIDAOICW, "
'    sSQL = sSQL & "WFSMPLIDGDCW, "
''--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
'    sSQL = sSQL & "EPSMPLIDB1CW, "
'    sSQL = sSQL & "EPSMPLIDB2CW, "
'    sSQL = sSQL & "EPSMPLIDB3CW, "
'    sSQL = sSQL & "EPSMPLIDL1CW, "
'    sSQL = sSQL & "EPSMPLIDL2CW, "
'    sSQL = sSQL & "EPSMPLIDL3CW, "
''--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
'    sSQL = sSQL & "WFHSGDCW, "
'    sSQL = sSQL & "TDAYCW, "
'    sSQL = sSQL & "KDAYCW "
'    sSQL = sSQL & "from "
'
'    sSQL = sSQL & "(select "
'    sSQL = sSQL & "xtalcb as CRYNUM, "
'    sSQL = sSQL & "inposcb as INGOTPOS, "
''    sSql = sSql & "rlencb as LENGTH, "
'    sSQL = sSQL & "LENCB as LENGTH, "         '理論長さ→長さ　06/11/09 ooba
'    sSQL = sSQL & "sxlidcb as SXLID, "
'    sSQL = sSQL & "' ' as KRPROCCD, "
'    sSQL = sSQL & "gnwkntcb as NOWPROC, "
'    sSQL = sSQL & "' ' as LPKRPROCCD, "
'    sSQL = sSQL & "newkntcb as LASTPASS, "
'    sSQL = sSQL & "livkcb as DELCLS, "
'    sSQL = sSQL & "lstccb as LSTATCLS, "
'    sSQL = sSQL & "sholdclscb HOLDCLS, "
'    sSQL = sSQL & "hinbcb as HINBAN, "
'    sSQL = sSQL & "revnumcb as REVNUM, "
'    sSQL = sSQL & "factorycb as FACTORY, "
'    sSQL = sSQL & "opecb as OPECOND, "
'    sSQL = sSQL & "MAICB, "
'    sSQL = sSQL & "tdaycb as REGDATE, "
'    sSQL = sSQL & "kdaycb as UPDDATE, "
'    sSQL = sSQL & "HOLDBCB, "
'' 2007/09/04 SPK Tsutsumi Add Start
'    sSQL = sSQL & "WFHOLDFLGCB, "
'    sSQL = sSQL & "KBLKFLGCB, "           '関連ﾌﾞﾛｯｸﾌﾗｸﾞ　08/01/31 ooba
'    sSQL = sSQL & "PLANTCATCB as PLANTCAT "
'' 2007/09/04 SPK Tsutsumi Add End
'    sSQL = sSQL & "from XSDCB "
'    sSQL = sSQL & "where GNWKNTCB = '" & inNowPorc & "' "
'    sSQL = sSQL & "and livkcb = '0' "
'
'' 2007/09/04 SPK Tsutsumi Add Start
'    If sCmbMukesaki <> "ALL" Then
'        sSQL = sSQL & "   AND PLANTCATCB      = '" & sCmbMukesaki & "'"
'    End If
'' 2007/09/04 SPK Tsutsumi Add End
'
'    sSQL = sSQL & "), "
'
'    sSQL = sSQL & "(select "
'    sSQL = sSQL & "SXLIDCW, "
'    sSQL = sSQL & "XTALCW, "
'    sSQL = sSQL & "INPOSCW, "
'    sSQL = sSQL & "TBKBNCW, "
'    sSQL = sSQL & "SMPKBNCW, "
'    sSQL = sSQL & "REPSMPLIDCW, "
'    sSQL = sSQL & "HINBCW, "
'    sSQL = sSQL & "REVNUMCW, "
'    sSQL = sSQL & "FACTORYCW, "
'    sSQL = sSQL & "OPECW, "
'    sSQL = sSQL & "KTKBNCW, "
'    sSQL = sSQL & "WFINDRSCW, "
'    sSQL = sSQL & "WFINDOICW, "
'    sSQL = sSQL & "WFINDB1CW, "
'    sSQL = sSQL & "WFINDB2CW, "
'    sSQL = sSQL & "WFINDB3CW, "
'    sSQL = sSQL & "WFINDL1CW, "
'    sSQL = sSQL & "WFINDL2CW, "
'    sSQL = sSQL & "WFINDL3CW, "
'    sSQL = sSQL & "WFINDL4CW, "
'    sSQL = sSQL & "WFINDDSCW, "
'    sSQL = sSQL & "WFINDDZCW, "
'    sSQL = sSQL & "WFINDSPCW, "
'    sSQL = sSQL & "WFINDDO1CW, "
'    sSQL = sSQL & "WFINDDO2CW, "
'    sSQL = sSQL & "WFINDDO3CW, "
'    sSQL = sSQL & "WFINDAOICW, "
'    sSQL = sSQL & "WFINDGDCW, "
''--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
'    sSQL = sSQL & "EPINDB1CW, "
'    sSQL = sSQL & "EPINDB2CW, "
'    sSQL = sSQL & "EPINDB3CW, "
'    sSQL = sSQL & "EPINDL1CW, "
'    sSQL = sSQL & "EPINDL2CW, "
'    sSQL = sSQL & "EPINDL3CW, "
''--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
'    sSQL = sSQL & "WFRESRS1CW, "
'    sSQL = sSQL & "WFRESOICW, "
'    sSQL = sSQL & "WFRESB1CW, "
'    sSQL = sSQL & "WFRESB2CW, "
'    sSQL = sSQL & "WFRESB3CW, "
'    sSQL = sSQL & "WFRESL1CW, "
'    sSQL = sSQL & "WFRESL2CW, "
'    sSQL = sSQL & "WFRESL3CW, "
'    sSQL = sSQL & "WFRESL4CW, "
'    sSQL = sSQL & "WFRESDSCW, "
'    sSQL = sSQL & "WFRESDZCW, "
'    sSQL = sSQL & "WFRESSPCW, "
'    sSQL = sSQL & "WFRESDO1CW, "
'    sSQL = sSQL & "WFRESDO2CW, "
'    sSQL = sSQL & "WFRESDO3CW, "
'    sSQL = sSQL & "WFRESAOICW, "
'    sSQL = sSQL & "WFRESGDCW, "
''--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
'    sSQL = sSQL & "EPRESB1CW, "
'    sSQL = sSQL & "EPRESB2CW, "
'    sSQL = sSQL & "EPRESB3CW, "
'    sSQL = sSQL & "EPRESL1CW, "
'    sSQL = sSQL & "EPRESL2CW, "
'    sSQL = sSQL & "EPRESL3CW, "
''--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
'    sSQL = sSQL & "WFSMPLIDRSCW, "
'    sSQL = sSQL & "WFSMPLIDOICW, "
'    sSQL = sSQL & "WFSMPLIDB1CW, "
'    sSQL = sSQL & "WFSMPLIDB2CW, "
'    sSQL = sSQL & "WFSMPLIDB3CW, "
'    sSQL = sSQL & "WFSMPLIDL1CW, "
'    sSQL = sSQL & "WFSMPLIDL2CW, "
'    sSQL = sSQL & "WFSMPLIDL3CW, "
'    sSQL = sSQL & "WFSMPLIDL4CW, "
'    sSQL = sSQL & "WFSMPLIDDSCW, "
'    sSQL = sSQL & "WFSMPLIDDZCW, "
'    sSQL = sSQL & "WFSMPLIDSPCW, "
'    sSQL = sSQL & "WFSMPLIDDO1CW, "
'    sSQL = sSQL & "WFSMPLIDDO2CW, "
'    sSQL = sSQL & "WFSMPLIDDO3CW, "
'    sSQL = sSQL & "WFSMPLIDAOICW, "
'    sSQL = sSQL & "WFSMPLIDGDCW, "
''--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
'    sSQL = sSQL & "EPSMPLIDB1CW, "
'    sSQL = sSQL & "EPSMPLIDB2CW, "
'    sSQL = sSQL & "EPSMPLIDB3CW, "
'    sSQL = sSQL & "EPSMPLIDL1CW, "
'    sSQL = sSQL & "EPSMPLIDL2CW, "
'    sSQL = sSQL & "EPSMPLIDL3CW, "
''--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
'    sSQL = sSQL & "WFHSGDCW, "
'    sSQL = sSQL & "TDAYCW, "
'    sSQL = sSQL & "KDAYCW "
'    sSQL = sSQL & "from XSDCW "
'    sSQL = sSQL & "where LIVKCW = '0' "
'    sSQL = sSQL & "and TBKBNCW = 'B' "
'    sSQL = sSQL & ") "
'    sSQL = sSQL & "where SXLID = SXLIDCW(+) "
'Del End   2011/03/07 SMPK Miyata

    sSQL = sSQL & "order by CRYNUM,INGOTPOS,TBKBNCW DESC "

    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_DEFAULT)

    intSxlCount = rs.RecordCount

    'レコード0件時正常終了
    If intSxlCount = 0 Then
        rs.Close
        ReDim sxl(0)
        GoTo proc_exit
    End If

    j = 0
    sNullSXLID = ""
'Add Start 2011/03/07 SMPK Miyata
    lngCmpCnt = 0               'サンプル数 初期化
    ReDim Preserve sxl(0)       'SXL管理    初期化
'Add End   2011/03/07 SMPK Miyata
    
'Chg Start 2011/03/07 SMPK Miyata
'    For i = 1 To intSxlCount Step 2
    For i = 1 To intSxlCount
'Chg End   2011/03/07 SMPK Miyata

'Add Start 2011/03/07 SMPK Miyata
        If rs("SXLID") <> sxl(j).SXLIDCA Then
'Add End   2011/03/07 SMPK Miyata

            lngNullCnt = 0
            j = j + 1
            ReDim Preserve sxl(j)
            ReDim Preserve WFJudgExecOkFlag(j)
            WFJudgExecOkFlag(j) = True '表示色デフォルト設定(緑)
            lngCmpCnt = 0               'サンプル数     'Add 2011/03/07 SMPK Miyata
                    
            With sxl(j)
                If IsNull(rs("CRYNUM")) = False Then .CRYNUMCA = rs("CRYNUM") Else lngNullCnt = lngNullCnt + 1        ' 結晶番号
                If IsNull(rs("INGOTPOS")) = False Then .INPOSCA = rs("INGOTPOS") Else lngNullCnt = lngNullCnt + 1     ' 結晶内開始位置
                If IsNull(rs("LENGTH")) = False Then .GNLCA = rs("LENGTH") Else lngNullCnt = lngNullCnt + 1           ' 長さ
                If IsNull(rs("SXLID")) = False Then .SXLIDCA = rs("SXLID") Else lngNullCnt = lngNullCnt + 1           ' SXLID
                If IsNull(rs("NOWPROC")) = False Then .NOWPROC = rs("NOWPROC") Else lngNullCnt = lngNullCnt + 1       ' 現在工程
                If IsNull(rs("LASTPASS")) = False Then .NEWKNTCA = rs("LASTPASS") Else lngNullCnt = lngNullCnt + 1    ' 最終通過工程
                If IsNull(rs("DELCLS")) = False Then .SAKJCA = rs("DELCLS") Else lngNullCnt = lngNullCnt + 1          ' 削除区分
                If IsNull(rs("LSTATCLS")) = False Then .LSTATBCA = rs("LSTATCLS") Else lngNullCnt = lngNullCnt + 1    ' 最終状態区分
                If IsNull(rs("HOLDCLS")) = False Then .HOLDBCA = rs("HOLDCLS") Else lngNullCnt = lngNullCnt + 1       ' ホールド区分
                If IsNull(rs("HINBAN")) = False Then .HINBCA = rs("HINBAN") Else lngNullCnt = lngNullCnt + 1          ' 品番
                If IsNull(rs("REVNUM")) = False Then .REVNUMCA = rs("REVNUM") Else lngNullCnt = lngNullCnt + 1        ' 製品番号改訂番号
                If IsNull(rs("FACTORY")) = False Then .FACTORYCA = rs("FACTORY") Else lngNullCnt = lngNullCnt + 1     ' 工場
                If IsNull(rs("OPECOND")) = False Then .OPECA = rs("OPECOND") Else lngNullCnt = lngNullCnt + 1         ' 操業条件
                If IsNull(rs("MAICB")) = False Then .MAICB = rs("MAICB") Else lngNullCnt = lngNullCnt + 1             ' 枚数
                If IsNull(rs("REGDATE")) = False Then .TDAYCB = rs("REGDATE") Else lngNullCnt = lngNullCnt + 1        ' 登録日付
                If IsNull(rs("UPDDATE")) = False Then .KDAYCA = rs("UPDDATE") Else lngNullCnt = lngNullCnt + 1        ' 更新日付
                If IsNull(rs("HOLDBCB")) = False Then .HOLDBCB = rs("HOLDBCB") Else .HOLDBCB = " "                  ' ﾎｰﾙﾄﾞ区分　06/02/08 ooba
                If IsNull(rs("WFHOLDFLGCB")) = False Then .WFHOLDFLGCB = rs("WFHOLDFLGCB") Else .WFHOLDFLGCB = " "  ' WFﾎｰﾙﾄﾞ区分　06/02/08 ooba
                If IsNull(rs("KBLKFLGCB")) = False Then .KANREN = rs("KBLKFLGCB") Else .KANREN = " "            ' 関連ﾌﾞﾛｯｸ有無　08/01/31 ooba
    
                ' 向先 07/09/04 SPK Tsutsumi Add Start
                If IsNull(rs("PLANTCAT")) = False Then
                    For k = 0 To UBound(s_MukesakiBase)
                        If s_MukesakiBase(k).sMukeCode = rs("PLANTCAT") Then
                           .PLANTCAT = s_MukesakiBase(k).sMukeName
                           Exit For
                        End If
                    Next k
                Else
                    .PLANTCAT = " "
                End If
                ' 向先 07/09/04 SPK Tsutsumi Add end
            End With
        End If              'Add 2011/03/07 SMPK Miyata

'Add Start 2011/03/07 SMPK Miyata
        If rs("SXLID") = sxl(j).SXLIDCA Then
            lngCmpCnt = lngCmpCnt + 1       'サンプル数カウント     'Add 2011/03/07 SMPK Miyata
'Add End   2011/03/07 SMPK Miyata

'Chg Start 2011/03/07 SMPK Miyata
'Chg End   2011/03/07 SMPK Miyata

'Chg Start 2011/03/07 SMPK Miyata
'        ReDim Preserve SXL(j).WFSMP(2)
'
'        For k = 1 To 2
'            Call Init_SXL_WFSMP(SXL(j).WFSMP(k))
'            With SXL(j).WFSMP(k)
            ReDim Preserve sxl(j).WFSMP(lngCmpCnt)
            With sxl(j).WFSMP(lngCmpCnt)
'Chg End   2011/03/07 SMPK Miyata

                If IsNull(rs("XTALCW")) = False Then .XTALCW = rs("XTALCW") Else lngNullCnt = lngNullCnt + 1                    ' 結晶番号
                If IsNull(rs("INPOSCW")) = False Then .INPOSCW = rs("INPOSCW") Else lngNullCnt = lngNullCnt + 1                 ' 結晶内位置
                If IsNull(rs("SMPKBNCW")) = False Then .SMPKBNCW = rs("SMPKBNCW") Else lngNullCnt = lngNullCnt + 1              ' サンプル区分
                If IsNull(rs("REPSMPLIDCW")) = False Then .REPSMPLIDCW = rs("REPSMPLIDCW") Else lngNullCnt = lngNullCnt + 1     ' サンプルID
                If IsNull(rs("HINBCW")) = False Then .HINBCW = rs("HINBCW") Else lngNullCnt = lngNullCnt + 1                    ' 品番
                If IsNull(rs("REVNUMCW")) = False Then .REVNUMCW = rs("REVNUMCW") Else lngNullCnt = lngNullCnt + 1              ' 製品番号改訂番号
                If IsNull(rs("FACTORYCW")) = False Then .FACTORYCW = rs("FACTORYCW") Else lngNullCnt = lngNullCnt + 1           ' 工場
                If IsNull(rs("OPECW")) = False Then .OPECW = rs("OPECW") Else lngNullCnt = lngNullCnt + 1                       ' 操業条件
                If IsNull(rs("KTKBNCW")) = False Then .KTKBNCW = rs("KTKBNCW") Else lngNullCnt = lngNullCnt + 1                 ' 確定区分

                If IsNull(rs("WFINDRSCW")) = False Then .WFINDRSCW = rs("WFINDRSCW") Else lngNullCnt = lngNullCnt + 1           ' WF検査指示（RS)
                If IsNull(rs("WFINDOICW")) = False Then .WFINDOICW = rs("WFINDOICW") Else lngNullCnt = lngNullCnt + 1           ' WF検査指示（Oi)
                If IsNull(rs("WFINDB1CW")) = False Then .WFINDB1CW = rs("WFINDB1CW") Else lngNullCnt = lngNullCnt + 1           ' WF検査指示（B1)
                If IsNull(rs("WFINDB2CW")) = False Then .WFINDB2CW = rs("WFINDB2CW") Else lngNullCnt = lngNullCnt + 1           ' WF検査指示（B2）
                If IsNull(rs("WFINDB3CW")) = False Then .WFINDB3CW = rs("WFINDB3CW") Else lngNullCnt = lngNullCnt + 1           ' WF検査指示（B3)
                If IsNull(rs("WFINDL1CW")) = False Then .WFINDL1CW = rs("WFINDL1CW") Else lngNullCnt = lngNullCnt + 1           ' WF検査指示（L1)
                If IsNull(rs("WFINDL2CW")) = False Then .WFINDL2CW = rs("WFINDL2CW") Else lngNullCnt = lngNullCnt + 1           ' WF検査指示（L2)
                If IsNull(rs("WFINDL3CW")) = False Then .WFINDL3CW = rs("WFINDL3CW") Else lngNullCnt = lngNullCnt + 1           ' WF検査指示（L3)
                If IsNull(rs("WFINDL4CW")) = False Then .WFINDL4CW = rs("WFINDL4CW") Else lngNullCnt = lngNullCnt + 1           ' WF検査指示（L4)
                If IsNull(rs("WFINDDSCW")) = False Then .WFINDDSCW = rs("WFINDDSCW") Else lngNullCnt = lngNullCnt + 1           ' WF検査指示（DS)
                If IsNull(rs("WFINDDZCW")) = False Then .WFINDDZCW = rs("WFINDDZCW") Else lngNullCnt = lngNullCnt + 1           ' WF検査指示（DZ)
                If IsNull(rs("WFINDSPCW")) = False Then .WFINDSPCW = rs("WFINDSPCW") Else lngNullCnt = lngNullCnt + 1           ' WF検査指示（SP)
                If IsNull(rs("WFINDDO1CW")) = False Then .WFINDDO1CW = rs("WFINDDO1CW") Else lngNullCnt = lngNullCnt + 1        ' WF検査指示（DO1)
                If IsNull(rs("WFINDDO2CW")) = False Then .WFINDDO2CW = rs("WFINDDO2CW") Else lngNullCnt = lngNullCnt + 1        ' WF検査指示（DO2)
                If IsNull(rs("WFINDDO3CW")) = False Then .WFINDDO3CW = rs("WFINDDO3CW") Else lngNullCnt = lngNullCnt + 1        ' WF検査指示（DO3)
                If IsNull(rs("WFINDAOICW")) = False Then .WFINDAOICW = rs("WFINDAOICW") Else lngNullCnt = lngNullCnt + 1        ' WF検査指示 (AOi)
                If IsNull(rs("WFINDGDCW")) = False Then .WFINDGDCW = rs("WFINDGDCW") Else lngNullCnt = lngNullCnt + 1           ' WF検査指示 (GD)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
                If IsNull(rs("EPINDB1CW")) = False Then .EPINDB1CW = rs("EPINDB1CW") Else lngNullCnt = lngNullCnt + 1           ' EP検査指示（B1E)
                If IsNull(rs("EPINDB2CW")) = False Then .EPINDB2CW = rs("EPINDB2CW") Else lngNullCnt = lngNullCnt + 1           ' EP検査指示（B2E）
                If IsNull(rs("EPINDB3CW")) = False Then .EPINDB3CW = rs("EPINDB3CW") Else lngNullCnt = lngNullCnt + 1           ' EP検査指示（B3E)
                If IsNull(rs("EPINDL1CW")) = False Then .EPINDL1CW = rs("EPINDL1CW") Else lngNullCnt = lngNullCnt + 1           ' EP検査指示（L1E)
                If IsNull(rs("EPINDL2CW")) = False Then .EPINDL2CW = rs("EPINDL2CW") Else lngNullCnt = lngNullCnt + 1           ' EP検査指示（L2E)
                If IsNull(rs("EPINDL3CW")) = False Then .EPINDL3CW = rs("EPINDL3CW") Else lngNullCnt = lngNullCnt + 1           ' EP検査指示（L3E)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

                If IsNull(rs("WFRESRS1CW")) = False Then .WFRESRS1CW = rs("WFRESRS1CW") Else lngNullCnt = lngNullCnt + 1        ' WF検査実績（RS)
                If IsNull(rs("WFRESOICW")) = False Then .WFRESOICW = rs("WFRESOICW") Else lngNullCnt = lngNullCnt + 1           ' WF検査実績（Oi)
                If IsNull(rs("WFRESB1CW")) = False Then .WFRESB1CW = rs("WFRESB1CW") Else lngNullCnt = lngNullCnt + 1           ' WF検査実績（B1)
                If IsNull(rs("WFRESB2CW")) = False Then .WFRESB2CW = rs("WFRESB2CW") Else lngNullCnt = lngNullCnt + 1           ' WF検査実績（B2）
                If IsNull(rs("WFRESB3CW")) = False Then .WFRESB3CW = rs("WFRESB3CW") Else lngNullCnt = lngNullCnt + 1           ' WF検査実績（B3)
                If IsNull(rs("WFRESL1CW")) = False Then .WFRESL1CW = rs("WFRESL1CW") Else lngNullCnt = lngNullCnt + 1           ' WF検査実績（L1)
                If IsNull(rs("WFRESL2CW")) = False Then .WFRESL2CW = rs("WFRESL2CW") Else lngNullCnt = lngNullCnt + 1           ' WF検査実績（L2)
                If IsNull(rs("WFRESL3CW")) = False Then .WFRESL3CW = rs("WFRESL3CW") Else lngNullCnt = lngNullCnt + 1           ' WF検査実績（L3)
                If IsNull(rs("WFRESL4CW")) = False Then .WFRESL4CW = rs("WFRESL4CW") Else lngNullCnt = lngNullCnt + 1           ' WF検査実績（L4)
                If IsNull(rs("WFRESDSCW")) = False Then .WFRESDSCW = rs("WFRESDSCW") Else lngNullCnt = lngNullCnt + 1           ' WF検査実績（DS)
                If IsNull(rs("WFRESDZCW")) = False Then .WFRESDZCW = rs("WFRESDZCW") Else lngNullCnt = lngNullCnt + 1           ' WF検査実績（DZ)
                If IsNull(rs("WFRESSPCW")) = False Then .WFRESSPCW = rs("WFRESSPCW") Else lngNullCnt = lngNullCnt + 1           ' WF検査実績（SP)
                If IsNull(rs("WFRESDO1CW")) = False Then .WFRESDO1CW = rs("WFRESDO1CW") Else lngNullCnt = lngNullCnt + 1        ' WF検査実績（DO1)
                If IsNull(rs("WFRESDO2CW")) = False Then .WFRESDO2CW = rs("WFRESDO2CW") Else lngNullCnt = lngNullCnt + 1        ' WF検査実績（DO2)
                If IsNull(rs("WFRESDO3CW")) = False Then .WFRESDO3CW = rs("WFRESDO3CW") Else lngNullCnt = lngNullCnt + 1        ' WF検査実績（DO3)
                If IsNull(rs("WFRESAOICW")) = False Then .WFRESAOICW = rs("WFRESAOICW") Else lngNullCnt = lngNullCnt + 1        ' WF検査実績（AOi)
                If IsNull(rs("WFRESGDCW")) = False Then .WFRESGDCW = rs("WFRESGDCW") Else lngNullCnt = lngNullCnt + 1           ' WF検査実績（GD)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
                If IsNull(rs("EPRESB1CW")) = False Then .EPRESB1CW = rs("EPRESB1CW") Else lngNullCnt = lngNullCnt + 1           ' EP検査実績（B1E)
                If IsNull(rs("EPRESB2CW")) = False Then .EPRESB2CW = rs("EPRESB2CW") Else lngNullCnt = lngNullCnt + 1           ' EP検査実績（B2E）
                If IsNull(rs("EPRESB3CW")) = False Then .EPRESB3CW = rs("EPRESB3CW") Else lngNullCnt = lngNullCnt + 1           ' EP検査実績（B3E)
                If IsNull(rs("EPRESL1CW")) = False Then .EPRESL1CW = rs("EPRESL1CW") Else lngNullCnt = lngNullCnt + 1           ' EP検査実績（L1E)
                If IsNull(rs("EPRESL2CW")) = False Then .EPRESL2CW = rs("EPRESL2CW") Else lngNullCnt = lngNullCnt + 1           ' EP検査実績（L2E)
                If IsNull(rs("EPRESL3CW")) = False Then .EPRESL3CW = rs("EPRESL3CW") Else lngNullCnt = lngNullCnt + 1           ' EP検査実績（L3E)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

                If IsNull(rs("WFSMPLIDRSCW")) = False Then .WFSMPLIDRSCW = rs("WFSMPLIDRSCW") Else lngNullCnt = lngNullCnt + 1      ' WFｻﾝﾌﾟﾙID（RS)
                If IsNull(rs("WFSMPLIDOICW")) = False Then .WFSMPLIDOICW = rs("WFSMPLIDOICW") Else lngNullCnt = lngNullCnt + 1      ' WFｻﾝﾌﾟﾙID（Oi)
                If IsNull(rs("WFSMPLIDB1CW")) = False Then .WFSMPLIDB1CW = rs("WFSMPLIDB1CW") Else lngNullCnt = lngNullCnt + 1      ' WFｻﾝﾌﾟﾙID（B1)
                If IsNull(rs("WFSMPLIDB2CW")) = False Then .WFSMPLIDB2CW = rs("WFSMPLIDB2CW") Else lngNullCnt = lngNullCnt + 1      ' WFｻﾝﾌﾟﾙID（B2）
                If IsNull(rs("WFSMPLIDB3CW")) = False Then .WFSMPLIDB3CW = rs("WFSMPLIDB3CW") Else lngNullCnt = lngNullCnt + 1      ' WFｻﾝﾌﾟﾙID（B3)
                If IsNull(rs("WFSMPLIDL1CW")) = False Then .WFSMPLIDL1CW = rs("WFSMPLIDL1CW") Else lngNullCnt = lngNullCnt + 1      ' WFｻﾝﾌﾟﾙID（L1)
                If IsNull(rs("WFSMPLIDL2CW")) = False Then .WFSMPLIDL2CW = rs("WFSMPLIDL2CW") Else lngNullCnt = lngNullCnt + 1      ' WFｻﾝﾌﾟﾙID（L2)
                If IsNull(rs("WFSMPLIDL3CW")) = False Then .WFSMPLIDL3CW = rs("WFSMPLIDL3CW") Else lngNullCnt = lngNullCnt + 1      ' WFｻﾝﾌﾟﾙID（L3)
                If IsNull(rs("WFSMPLIDL4CW")) = False Then .WFSMPLIDL4CW = rs("WFSMPLIDL4CW") Else lngNullCnt = lngNullCnt + 1      ' WFｻﾝﾌﾟﾙID（L4)
                If IsNull(rs("WFSMPLIDDSCW")) = False Then .WFSMPLIDDSCW = rs("WFSMPLIDDSCW") Else lngNullCnt = lngNullCnt + 1      ' WFｻﾝﾌﾟﾙID（DS)
                If IsNull(rs("WFSMPLIDDZCW")) = False Then .WFSMPLIDDZCW = rs("WFSMPLIDDZCW") Else lngNullCnt = lngNullCnt + 1      ' WFｻﾝﾌﾟﾙID（DZ)
                If IsNull(rs("WFSMPLIDSPCW")) = False Then .WFSMPLIDSPCW = rs("WFSMPLIDSPCW") Else lngNullCnt = lngNullCnt + 1      ' WFｻﾝﾌﾟﾙID（SP)
                If IsNull(rs("WFSMPLIDDO1CW")) = False Then .WFSMPLIDDO1CW = rs("WFSMPLIDDO1CW") Else lngNullCnt = lngNullCnt + 1   ' WFｻﾝﾌﾟﾙID（DO1)
                If IsNull(rs("WFSMPLIDDO2CW")) = False Then .WFSMPLIDDO2CW = rs("WFSMPLIDDO2CW") Else lngNullCnt = lngNullCnt + 1   ' WFｻﾝﾌﾟﾙID（DO2)
                If IsNull(rs("WFSMPLIDDO3CW")) = False Then .WFSMPLIDDO3CW = rs("WFSMPLIDDO3CW") Else lngNullCnt = lngNullCnt + 1   ' WFｻﾝﾌﾟﾙID（DO3)
                If IsNull(rs("WFSMPLIDAOICW")) = False Then .WFSMPLIDAOICW = rs("WFSMPLIDAOICW") Else lngNullCnt = lngNullCnt + 1   ' WFｻﾝﾌﾟﾙID（AOi)
                If IsNull(rs("WFSMPLIDGDCW")) = False Then .WFSMPLIDGDCW = rs("WFSMPLIDGDCW") Else lngNullCnt = lngNullCnt + 1      ' WFｻﾝﾌﾟﾙID（GD)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
                If IsNull(rs("EPSMPLIDB1CW")) = False Then .EPSMPLIDB1CW = rs("EPSMPLIDB1CW") Else lngNullCnt = lngNullCnt + 1      ' EPｻﾝﾌﾟﾙID（B1E)
                If IsNull(rs("EPSMPLIDB2CW")) = False Then .EPSMPLIDB2CW = rs("EPSMPLIDB2CW") Else lngNullCnt = lngNullCnt + 1      ' EPｻﾝﾌﾟﾙID（B2E）
                If IsNull(rs("EPSMPLIDB3CW")) = False Then .EPSMPLIDB3CW = rs("EPSMPLIDB3CW") Else lngNullCnt = lngNullCnt + 1      ' EPｻﾝﾌﾟﾙID（B3E)
                If IsNull(rs("EPSMPLIDL1CW")) = False Then .EPSMPLIDL1CW = rs("EPSMPLIDL1CW") Else lngNullCnt = lngNullCnt + 1      ' EPｻﾝﾌﾟﾙID（L1E)
                If IsNull(rs("EPSMPLIDL2CW")) = False Then .EPSMPLIDL2CW = rs("EPSMPLIDL2CW") Else lngNullCnt = lngNullCnt + 1      ' EPｻﾝﾌﾟﾙID（L2E)
                If IsNull(rs("EPSMPLIDL3CW")) = False Then .EPSMPLIDL3CW = rs("EPSMPLIDL3CW") Else lngNullCnt = lngNullCnt + 1      ' EPｻﾝﾌﾟﾙID（L3E)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
                If IsNull(rs("WFHSGDCW")) = False Then .WFHSGDCW = rs("WFHSGDCW") Else lngNullCnt = lngNullCnt + 1                  ' WF検査保証 (GD)
                If IsNull(rs("TDAYCW")) = False Then .TDAYCW = rs("TDAYCW") Else lngNullCnt = lngNullCnt + 1                        ' 登録日付
                If IsNull(rs("KDAYCW")) = False Then .KDAYCW = rs("KDAYCW") Else lngNullCnt = lngNullCnt + 1                        ' 更新日付
            End With

'Chg Start 2011/03/07 SMPK Miyata
'            If Not rs.EOF Then
'                rs.MoveNext
'            End If
'        Next k
            rs.MoveNext
        End If
'Chg End   2011/03/07 SMPK Miyata

        If lngNullCnt > 0 And sNullSXLID = "" Then
            sNullSXLID = sxl(j).SXLIDCA
        End If

        If lngNullCnt > 0 Then
            WFJudgExecOkFlag(j) = False
        End If
    Next i
    rs.Close
    '仕掛SXL取得SQL変更　06/02/07 ooba END =================================================>

'=================================================================================
' 2011/02/16 tkimura MOD START
' --- 6番の処理をコメントアウトすると高速化が実装される。
Debug.Print "6 " & Now & " 測定結果の受信確認"

    '測定評価結果受信確認
    '指示に対する測定評価結果を受信しているかどうかのチェック
    '受信していれば、WFサンプル管理を更新
'Cng Start 2011/06/16 Y.Hitomi MQ受信時のSIRD反映不具合対応
    If MeasRsltCheck1(sxl()) = FUNCTION_RETURN_FAILURE Then
'    If MeasRsltCheck(SXL()) = FUNCTION_RETURN_FAILURE Then
'Cng End   2011/06/16 Y.Hitomi
        DBDRV_scmzc_fcmlc001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If


Debug.Print "7 " & Now & " 検査完了チェック"

    'すべての検査が完了しているかどうかのチェック
    For i = 1 To UBound(sxl)
        '検査項目チェック
        Select Case UBound(sxl(i).WFSMP)
            Case 1
                ReDim Preserve sxl(i).WFSMP(2) As typ_XSDCW
                    WFJudgExecOkFlag(i) = False
            Case 2
                If Trim(sxl(i).WFSMP(1).XTALCW) = "" And Trim(sxl(i).WFSMP(2).XTALCW) = "" Then
                Else
                    If Not (ChkRslt(sxl(i).WFSMP(1)) And ChkRslt(sxl(i).WFSMP(2))) Then
                        WFJudgExecOkFlag(i) = False
                    End If
                End If
'Add Start 2011/03/07 SMPK Miyata
            Case Else
                For k = 1 To UBound(sxl(i).WFSMP)
                    If Not ChkRslt(sxl(i).WFSMP(k)) Then
                        WFJudgExecOkFlag(i) = False
                        Exit For
                    End If
                Next k
'Add End   2011/03/07 SMPK Miyata
        End Select
    Next
Debug.Print "8 " & Now

    If sNullSXLID <> "" Then
        f_cmbc039_1.lblMsg.Caption = "データ不正 SXLID=" & sNullSXLID
        GoTo proc_exit
    Else
        f_cmbc039_1.lblMsg.Caption = ""
    End If

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    DBDRV_scmzc_fcmlc001b_Disp = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'*******************************************************************************
'*    関数名        : Init_SXL_WFSMP
'*
'*    処理概要      : 1.新サンプル管理テーブルの初期化
'*
'*    パラメータ    : 変数名        ,IO ,型         ,説明
'*                    WFSMP         ,O  ,typ_XSDCW  ,新サンプル管理（SXL）
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Sub Init_SXL_WFSMP(WFSMP As typ_XSDCW)
    With WFSMP
        .FACTORYCW = ""
        .HINBCW = ""
        .INPOSCW = 0
        .KDAYCW = 0
        .KSTAFFCW = ""
        .KTKBNCW = ""
        .LIVKCW = ""
        .OPECW = ""
        .REPSMPLIDCW = ""
        .REVNUMCW = 0
        .SMCRYNUMCW = ""
        .SMPKBNCW = ""
        .SMPLNUMCW = 0
        .SNDDAYCW = 0
        .SNDKCW = ""
        .SXLIDCW = ""
        .TBKBNCW = ""
        .TDAYCW = 0
        .TSTAFFCW = ""
        .WFINDAOICW = ""
        .WFINDB1CW = ""
        .WFINDB1CW = ""
        .WFINDB2CW = ""
        .WFINDB3CW = ""
        .WFINDDO1CW = ""
        .WFINDDO2CW = ""
        .WFINDDO3CW = ""
        .WFINDDSCW = ""
        .WFINDDZCW = ""
        .WFINDL1CW = ""
        .WFINDL2CW = ""
        .WFINDL3CW = ""
        .WFINDL4CW = ""
        .WFINDOICW = ""
        .WFINDOT1CW = ""
        .WFINDOT2CW = ""
        .WFINDRSCW = ""
        .WFINDSPCW = ""
        .WFINDGDCW = ""         '05/02/04 ooba
        .WFRESB1CW = ""
        .WFRESAOICW = ""
        .WFRESB2CW = ""
        .WFRESB3CW = ""
        .WFRESDO1CW = ""
        .WFRESDO2CW = ""
        .WFRESDO3CW = ""
        .WFRESDSCW = ""
        .WFRESDZCW = ""
        .WFRESL1CW = ""
        .WFRESL2CW = ""
        .WFRESL3CW = ""
        .WFRESL4CW = ""
        .WFRESOICW = ""
        .WFRESOT1CW = ""
        .WFRESOT2CW = ""
        .WFRESRS1CW = ""
        .WFRESRS2CW = ""
        .WFRESSPCW = ""
        .WFRESGDCW = ""         '05/02/04 ooba
        .WFSMPLIDAOICW = ""
        .WFSMPLIDB1CW = ""
        .WFSMPLIDB2CW = ""
        .WFSMPLIDB3CW = ""
        .WFSMPLIDDO1CW = ""
        .WFSMPLIDDO2CW = ""
        .WFSMPLIDDO3CW = ""
        .WFSMPLIDDSCW = ""
        .WFSMPLIDDZCW = ""
        .WFSMPLIDL1CW = ""
        .WFSMPLIDL2CW = ""
        .WFSMPLIDL3CW = ""
        .WFSMPLIDL4CW = ""
        .WFSMPLIDOICW = ""
        .WFSMPLIDOT1CW = ""
        .WFSMPLIDOT2CW = ""
        .WFSMPLIDRS1CW = ""
        .WFSMPLIDRS2CW = ""
        .WFSMPLIDRSCW = ""
        .WFSMPLIDSPCW = ""
        .WFSMPLIDGDCW = ""      '05/02/04 ooba
        .WFHSGDCW = ""          '05/02/04 ooba
        .XTALCW = ""
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
        .EPSMPLIDB1CW = ""      'サンプルID(BMD1)
        .EPINDB1CW = ""         '状態FLG(BMD1)
        .EPRESB1CW = ""         '実績FLG(BMD1)
        .EPSMPLIDB2CW = ""      'サンプルID(BMD2)
        .EPINDB2CW = ""         '状態FLG(BMD2)
        .EPRESB2CW = ""         '実績FLG(BMD2)
        .EPSMPLIDB3CW = ""      'サンプルID(BMD3)
        .EPINDB3CW = ""         '状態FLG(BMD3)
        .EPRESB3CW = ""         '実績FLG(BMD3)
        .EPSMPLIDL1CW = ""      'サンプルID(OSF1)
        .EPINDL1CW = ""         '状態FLG(OSF1)
        .EPRESL1CW = ""         '実績FLG(OSF1)
        .EPSMPLIDL2CW = ""      'サンプルID(OSF2)
        .EPINDL2CW = ""         '状態FLG(OSF2)
        .EPRESL2CW = ""         '実績FLG(OSF2)
        .EPSMPLIDL3CW = ""      'サンプルID(OSF3)
        .EPINDL3CW = ""         '状態FLG(OSF3)
        .EPRESL3CW = ""         '実績FLG(OSF3)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
    End With
End Sub

'**************************************************************************************
'*    関数名        : GetSxlidINBlkid
'*
'*    処理概要      : 1.検査実績完了チェック
'*                      (検査指示されている検査が終了しているかチェックする)
'*
'*    パラメータ    : 変数名        ,IO ,型           ,説明
'*                    typ_WfSmp     ,I  ,typ_XSDCW    ,新サンプル管理（SXL）情報構造体
'*
'*    戻り値        : Boolean
'*
'**************************************************************************************
Public Function ChkRslt(typ_WFSmp As typ_XSDCW) As Boolean
    With typ_WFSmp
        If .WFINDRSCW <> "0" And .WFRESRS1CW = "0" Then               ' 状態FLG（Rs)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDOICW <> "0" And .WFRESOICW = "0" Then           ' 状態FLG（Oi)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDB1CW <> "0" And .WFRESB1CW = "0" Then           ' 状態FLG（B1)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDB2CW <> "0" And .WFRESB2CW = "0" Then           ' 状態FLG（B2）
            ChkRslt = False
            Exit Function
        ElseIf .WFINDB3CW <> "0" And .WFRESB3CW = "0" Then           ' 状態FLG（B3)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDL1CW <> "0" And .WFRESL1CW = "0" Then           ' 状態FLG（L1)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDL2CW <> "0" And .WFRESL2CW = "0" Then           ' 状態FLG（L2)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDL3CW <> "0" And .WFRESL3CW = "0" Then           ' 状態FLG（L3)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDL4CW <> "0" And .WFRESL4CW = "0" Then           ' 状態FLG（L4)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDDSCW <> "0" And .WFRESDSCW = "0" Then           ' 状態FLG（DS)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDDZCW <> "0" And .WFRESDZCW = "0" Then           ' 状態FLG（DZ)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDSPCW <> "0" And .WFRESSPCW = "0" Then           ' 状態FLG（SP)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDDO1CW <> "0" And .WFRESDO1CW = "0" Then         ' 状態FLG（DO1)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDDO2CW <> "0" And .WFRESDO2CW = "0" Then         ' 状態FLG（DO2)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDDO3CW <> "0" And .WFRESDO3CW = "0" Then         ' 状態FLG（DO3)
            ChkRslt = False
            Exit Function
        ''残存酸素追加　03/12/15 ooba
        ElseIf .WFINDAOICW <> "0" And .WFRESAOICW = "0" Then         ' 状態FLG (AOi)
            ChkRslt = False
            Exit Function
        'GD追加　05/02/17 ooba
        ElseIf .WFINDGDCW <> "0" And .WFRESGDCW = "0" Then           ' 状態FLG (GD)
            ChkRslt = False
            Exit Function
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
        ElseIf .EPINDB1CW <> "0" And .EPRESB1CW = "0" Then           ' 状態FLG（B1E)
            ChkRslt = False
            Exit Function
        ElseIf .EPINDB2CW <> "0" And .EPRESB2CW = "0" Then           ' 状態FLG（B2E）
            ChkRslt = False
            Exit Function
        ElseIf .EPINDB3CW <> "0" And .EPRESB3CW = "0" Then           ' 状態FLG（B3E)
            ChkRslt = False
            Exit Function
        ElseIf .EPINDL1CW <> "0" And .EPRESL1CW = "0" Then           ' 状態FLG（L1E)
            ChkRslt = False
            Exit Function
        ElseIf .EPINDL2CW <> "0" And .EPRESL2CW = "0" Then           ' 状態FLG（L2E)
            ChkRslt = False
            Exit Function
        ElseIf .EPINDL3CW <> "0" And .EPRESL3CW = "0" Then           ' 状態FLG（L3E)
            ChkRslt = False
            Exit Function
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
        End If
    End With
    ChkRslt = True
End Function

'**************************************************************************************************
'*    関数名        : DBDRV_GetTBCMY013
'*
'*    処理概要      : 1.テーブル「TBCMY013」から条件にあったレコードを抽出する
'*
'*    パラメータ    : 変数名        ,IO ,型           ,説明
'*                    records()     ,O  ,typ_TBCMY013 ,抽出レコード
'*                    sSqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'*                    sSqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'**************************************************************************************************
Public Function DBDRV_GetTBCMY013(records() As typ_TBCMY013, Optional sSqlWhere$ = vbNullString, _
                                   Optional sSqlOrder$ = vbNullString) As FUNCTION_RETURN
    Dim sSQL        As String       'SQL全体
    Dim sSqlBase    As String       'SQL基本部(WHERE節の前まで)
    Dim rs          As OraDynaset   'RecordSet
    Dim lngRecCnt   As Long         'レコード数
    Dim i           As Long

    ''SQLを組み立てる
    sSqlBase = "Select SAMPLEID, OSITEM, MAISU, SPEC, NETSU, ET, MES, DKAN, MESDATA1, MESDATA2, MESDATA3, MESDATA4, MESDATA5," & _
              " MESDATA6, MESDATA7, MESDATA8, MESDATA9, MESDATA10, MESDATA11, MESDATA12, MESDATA13, MESDATA14, MESDATA15," & _
              " TXID, REGDATE, SENDFLAG, SENDDATE "
    sSqlBase = sSqlBase & "From TBCMY013"
    sSQL = sSqlBase
    If (sSqlWhere <> vbNullString) Or (sSqlOrder <> vbNullString) Then
        sSQL = sSQL & " " & sSqlWhere & " " & sSqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMY013 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    lngRecCnt = rs.RecordCount
    ReDim records(lngRecCnt)
    For i = 1 To lngRecCnt
        With records(i)
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
    rs.Close

    DBDRV_GetTBCMY013 = FUNCTION_RETURN_SUCCESS
End Function

'**************************************************************************************************
'*    関数名        : DBDRV_scmzc_fcmlc001c_Disp
'*
'*    処理概要      : 1.WF総合判定 表示用ＤＢドライバ
'*
'*    パラメータ    : 変数名       ,IO  ,型                                    ,説明
'*                    typIn        ,I   ,type_DBDRV_scmzc_fcmlc001c_In         ,入力用
'*                    Siyou        ,O   ,type_DBDRV_scmzc_fcmlc001c_Siyou      ,WF仕様用
'*                    Sokutei      ,O   ,typ_TBCMY013                          ,測定評価結果
'*              　　  sErrMsg 　　 ,O   ,String    　　　　　　　　　　　      ,エラーメッセージ
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'**************************************************************************************************
Public Function DBDRV_scmzc_fcmlc001c_Disp(typIn As type_DBDRV_scmzc_fcmlc001c_In039, _
                                           siyou As type_DBDRV_scmzc_fcmlc001c_Siyou039, _
                                           Sokutei() As typ_TBCMY013, _
                                           sErrMsg As String) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim lngRecCnt   As Integer
    Dim i           As Long
    Dim sDBName     As String

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001c_Disp"

    DBDRV_scmzc_fcmlc001c_Disp = FUNCTION_RETURN_SUCCESS

    'WF仕様取得
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
    sSQL = sSQL & "E021HWFRMCAL, "           ' 品ＷＦ比抵抗面内計算 2001/11/08 S.Sano
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
    sSQL = sSQL & "E025HWFONMCL, "           ' 品ＷＦ酸素濃度面内計算 2001/11/08 S.Sano
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

    sSQL = sSQL & "E029HWFOF1AX, "          ' 品ＷＦＯＳＦ１平均上限
    sSQL = sSQL & "E029HWFOF1MX, "          ' 品ＷＦＯＳＦ１上限
    sSQL = sSQL & "E029HWFOF1SH, "          ' 品ＷＦＯＳＦ１測定位置＿方
    sSQL = sSQL & "E029HWFOF1ST, "          ' 品ＷＦＯＳＦ１測定位置＿点
    sSQL = sSQL & "E029HWFOF1SR, "          ' 品ＷＦＯＳＦ１測定位置＿領
    sSQL = sSQL & "E029HWFOF1HT, "          ' 品ＷＦＯＳＦ１保証方法＿対
    sSQL = sSQL & "E029HWFOF1HS, "          ' 品ＷＦＯＳＦ１保証方法＿処
    sSQL = sSQL & "E029HWFOF2AX, "          ' 品ＷＦＯＳＦ２平均上限
    sSQL = sSQL & "E029HWFOF2MX, "          ' 品ＷＦＯＳＦ２上限
    sSQL = sSQL & "E029HWFOF2SH, "          ' 品ＷＦＯＳＦ２測定位置＿方
    sSQL = sSQL & "E029HWFOF2ST, "          ' 品ＷＦＯＳＦ２測定位置＿点
    sSQL = sSQL & "E029HWFOF2SR, "          ' 品ＷＦＯＳＦ２測定位置＿領
    sSQL = sSQL & "E029HWFOF2HT, "          ' 品ＷＦＯＳＦ２保証方法＿対
    sSQL = sSQL & "E029HWFOF2HS, "          ' 品ＷＦＯＳＦ２保証方法＿処
    sSQL = sSQL & "E029HWFOF3AX, "          ' 品ＷＦＯＳＦ３平均上限
    sSQL = sSQL & "E029HWFOF3MX, "          ' 品ＷＦＯＳＦ３上限
    sSQL = sSQL & "E029HWFOF3SH, "          ' 品ＷＦＯＳＦ３測定位置＿方
    sSQL = sSQL & "E029HWFOF3ST, "          ' 品ＷＦＯＳＦ３測定位置＿点
    sSQL = sSQL & "E029HWFOF3SR, "          ' 品ＷＦＯＳＦ３測定位置＿領
    sSQL = sSQL & "E029HWFOF3HT, "          ' 品ＷＦＯＳＦ３保証方法＿対
    sSQL = sSQL & "E029HWFOF3HS, "          ' 品ＷＦＯＳＦ３保証方法＿処
    sSQL = sSQL & "E029HWFOF4AX, "          ' 品ＷＦＯＳＦ４平均上限
    sSQL = sSQL & "E029HWFOF4MX, "          ' 品ＷＦＯＳＦ４上限
    sSQL = sSQL & "E029HWFOF4SH, "          ' 品ＷＦＯＳＦ４測定位置＿方
    sSQL = sSQL & "E029HWFOF4ST, "          ' 品ＷＦＯＳＦ４測定位置＿点
    sSQL = sSQL & "E029HWFOF4SR, "          ' 品ＷＦＯＳＦ４測定位置＿領
    sSQL = sSQL & "E029HWFOF4HT, "          ' 品ＷＦＯＳＦ４保証方法＿対
    sSQL = sSQL & "E029HWFOF4HS, "          ' 品ＷＦＯＳＦ４保証方法＿処
    sSQL = sSQL & "E029HWFBM1AN, "          ' 品ＷＦＢＭＤ１平均下限
    sSQL = sSQL & "E029HWFBM1AX, "          ' 品ＷＦＢＭＤ１平均上限
    sSQL = sSQL & "E029HWFBM1SH, "          ' 品ＷＦＢＭＤ１測定位置＿方
    sSQL = sSQL & "E029HWFBM1ST, "          ' 品ＷＦＢＭＤ１測定位置＿点
    sSQL = sSQL & "E029HWFBM1SR, "          ' 品ＷＦＢＭＤ１測定位置＿領
    sSQL = sSQL & "E029HWFBM1HT, "          ' 品ＷＦＢＭＤ１保証方法＿対
    sSQL = sSQL & "E029HWFBM1HS, "          ' 品ＷＦＢＭＤ１保証方法＿処
    sSQL = sSQL & "E029HWFBM2AN, "          ' 品ＷＦＢＭＤ２平均下限
    sSQL = sSQL & "E029HWFBM2AX, "          ' 品ＷＦＢＭＤ２平均上限
    sSQL = sSQL & "E029HWFBM2SH, "          ' 品ＷＦＢＭＤ２測定位置＿方
    sSQL = sSQL & "E029HWFBM2ST, "          ' 品ＷＦＢＭＤ２測定位置＿点
    sSQL = sSQL & "E029HWFBM2SR, "          ' 品ＷＦＢＭＤ２測定位置＿領
    sSQL = sSQL & "E029HWFBM2HT, "          ' 品ＷＦＢＭＤ２保証方法＿対
    sSQL = sSQL & "E029HWFBM2HS, "          ' 品ＷＦＢＭＤ２保証方法＿処
    sSQL = sSQL & "E029HWFBM3AN, "          ' 品ＷＦＢＭＤ３平均下限
    sSQL = sSQL & "E029HWFBM3AX, "          ' 品ＷＦＢＭＤ３平均上限
    sSQL = sSQL & "E029HWFBM3SH, "          ' 品ＷＦＢＭＤ３測定位置＿方
    sSQL = sSQL & "E029HWFBM3ST, "          ' 品ＷＦＢＭＤ３測定位置＿点
    sSQL = sSQL & "E029HWFBM3SR, "          ' 品ＷＦＢＭＤ３測定位置＿領
    sSQL = sSQL & "E029HWFBM3HT, "          ' 品ＷＦＢＭＤ３保証方法＿対
    sSQL = sSQL & "E029HWFBM3HS, "          ' 品ＷＦＢＭＤ３保証方法＿処
    sSQL = sSQL & "E029HWFOSF1PTK, "        ' 品ＷＦＯＳＦ１パタン区分　▼2003/05/14 ooba
    sSQL = sSQL & "E029HWFOSF2PTK, "        ' 品ＷＦＯＳＦ２パタン区分
    sSQL = sSQL & "E029HWFOSF3PTK, "        ' 品ＷＦＯＳＦ３パタン区分
    sSQL = sSQL & "E029HWFOSF4PTK, "        ' 品ＷＦＯＳＦ４パタン区分
    sSQL = sSQL & "E029HWFBM1MBP, "         ' 品ＷＦＢＭＤ１面内分布
    sSQL = sSQL & "E029HWFBM2MBP, "         ' 品ＷＦＢＭＤ２面内分布
    sSQL = sSQL & "E029HWFBM3MBP, "         ' 品ＷＦＢＭＤ３面内分布
    sSQL = sSQL & "E029HWFBM1MCL, "         ' 品ＷＦＢＭＤ１面内計算
    sSQL = sSQL & "E029HWFBM2MCL, "         ' 品ＷＦＢＭＤ２面内計算
    sSQL = sSQL & "E029HWFBM3MCL, "         ' 品ＷＦＢＭＤ３面内計算　▲2003/05/14 ooba
    sSQL = sSQL & "E025HWFOS1NS, "          ' 品ＷＦ酸素析出１熱処理法
    sSQL = sSQL & "E025HWFOS2NS, "          ' 品ＷＦ酸素析出２熱処理法
    sSQL = sSQL & "E025HWFOS3NS, "          ' 品ＷＦ酸素析出３熱処理法

    sSQL = sSQL & "E029HWFOF1NS, "          ' 品ＷＦＯＳＦ１熱処理法
    sSQL = sSQL & "E029HWFOF2NS, "          ' 品ＷＦＯＳＦ２熱処理法
    sSQL = sSQL & "E029HWFOF3NS, "          ' 品ＷＦＯＳＦ３熱処理法
    sSQL = sSQL & "E029HWFOF4NS, "          ' 品ＷＦＯＳＦ４熱処理法

    sSQL = sSQL & "E029HWFBM1NS, "          ' 品ＷＦＢＭＤ１熱処理法
    sSQL = sSQL & "E029HWFBM2NS, "          ' 品ＷＦＢＭＤ２熱処理法
    sSQL = sSQL & "E029HWFBM3NS, "          ' 品ＷＦＢＭＤ３熱処理法

    sSQL = sSQL & "E025HWFANTIM, "          ' 品ＷＦＡＮ時間
    sSQL = sSQL & "E025HWFANTNP, "          ' 品ＷＦＡＮ温度

    sSQL = sSQL & "E029HWFOF1ET, "          ' 品ＷＦＯＳＦ１選択ＥＴ代
    sSQL = sSQL & "E029HWFOF2ET, "          ' 品ＷＦＯＳＦ２選択ＥＴ代
    sSQL = sSQL & "E029HWFOF3ET, "          ' 品ＷＦＯＳＦ３選択ＥＴ代
    sSQL = sSQL & "E029HWFOF4ET, "          ' 品ＷＦＯＳＦ４選択ＥＴ代
    sSQL = sSQL & "E029HWFBM1ET, "          ' 品ＷＦＢＭＤ１選択ＥＴ代
    sSQL = sSQL & "E029HWFBM2ET, "          ' 品ＷＦＢＭＤ２選択ＥＴ代
    sSQL = sSQL & "E029HWFBM3ET, "          ' 品ＷＦＢＭＤ３選択ＥＴ代

    sSQL = sSQL & "E029HWFOF1SZ, "          ' 品ＷＦＯＳＦ１測定条件
    sSQL = sSQL & "E029HWFOF2SZ, "          ' 品ＷＦＯＳＦ２測定条件
    sSQL = sSQL & "E029HWFOF3SZ, "          ' 品ＷＦＯＳＦ３測定条件
    sSQL = sSQL & "E029HWFOF4SZ, "          ' 品ＷＦＯＳＦ４測定条件
    sSQL = sSQL & "E029HWFBM1SZ, "          ' 品ＷＦＢＭＤ１測定条件
    sSQL = sSQL & "E029HWFBM2SZ, "          ' 品ＷＦＢＭＤ２測定条件
    sSQL = sSQL & "E029HWFBM3SZ "           ' 品ＷＦＢＭＤ３測定条件

    sSQL = sSQL & " from VECME001"
    sSQL = sSQL & " where E018HINBAN='" & typIn.HIN.hinban & "' and "
    sSQL = sSQL & " E018MNOREVNO=" & typIn.HIN.mnorevno & " and "
    sSQL = sSQL & " E018FACTORY='" & typIn.HIN.factory & "' and "
    sSQL = sSQL & " E018OPECOND='" & typIn.HIN.opecond & "' "

    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    'レコード0件はエラー終了
    If rs.RecordCount = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDBName)
        DBDRV_scmzc_fcmlc001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    With siyou
        .HWFTYPE = rs("E021HWFTYPE")              ' 品ＷＦタイプ
        .HWFCDIR = rs("E022HWFCDIR")              ' 品ＷＦ結晶面方
        .HWFCDOP = rs("E023HWFCDOP")              ' 品ＷＦ結晶ドープ

        .HWFRMIN = rs("E021HWFRMIN")              ' 品ＷＦ比抵抗下限
        .HWFRMAX = rs("E021HWFRMAX")              ' 品ＷＦ比抵抗上限
        .HWFRSPOH = rs("E021HWFRSPOH")            ' 品ＷＦ比抵抗測定位置＿方
        .HWFRSPOT = rs("E021HWFRSPOT")            ' 品ＷＦ比抵抗測定位置＿点
        .HWFRSPOI = rs("E021HWFRSPOI")            ' 品ＷＦ比抵抗測定位置＿位
        .HWFRHWYT = rs("E021HWFRHWYT")            ' 品ＷＦ比抵抗保証方法＿対
        .HWFRHWYS = rs("E021HWFRHWYS")            ' 品ＷＦ比抵抗保証方法＿処
        .HWFRMCAL = rs("E021HWFRMCAL")            ' 品ＷＦ比抵抗面内計算 2001/11/08 S.Sano
        .HWFRAMIN = rs("E021HWFRAMIN")            ' 品ＷＦ比抵抗平均下限
        .HWFRAMAX = rs("E021HWFRAMAX")            ' 品ＷＦ比抵抗平均上限
        .HWFRMBNP = rs("E021HWFRMBNP")            ' 品ＷＦ比抵抗面内分布

        .HWFMKMIN = rs("E024HWFMKMIN")            ' 品ＷＦ無欠陥層下限
        .HWFMKMAX = rs("E024HWFMKMAX")            ' 品ＷＦ無欠陥層上限
        .HWFMKSPH = rs("E024HWFMKSPH")            ' 品ＷＦ無欠陥層測定位置＿方
        .HWFMKSPT = rs("E024HWFMKSPT")            ' 品ＷＦ無欠陥層測定位置＿点
        .HWFMKSPR = rs("E024HWFMKSPR")            ' 品ＷＦ無欠陥層測定位置＿領
        .HWFMKHWT = rs("E024HWFMKHWT")            ' 品ＷＦ無欠陥層保証方法＿対
        .HWFMKHWS = rs("E024HWFMKHWS")            ' 品ＷＦ無欠陥層保証方法＿処

        .HWFONMIN = rs("E025HWFONMIN")            ' 品ＷＦ酸素濃度下限
        .HWFONMAX = rs("E025HWFONMAX")            ' 品ＷＦ酸素濃度上限
        .HWFONSPH = rs("E025HWFONSPH")            ' 品ＷＦ酸素濃度測定位置＿方
        .HWFONSPT = rs("E025HWFONSPT")            ' 品ＷＦ酸素濃度測定位置＿点
        .HWFONSPI = rs("E025HWFONSPI")            ' 品ＷＦ酸素濃度測定位置＿位
        .HWFONHWT = rs("E025HWFONHWT")            ' 品ＷＦ酸素濃度保証方法＿対
        .HWFONHWS = rs("E025HWFONHWS")            ' 品ＷＦ酸素濃度保証方法＿処
        .HWFONMCL = rs("E025HWFONMCL")            ' 品ＷＦ酸素濃度面内計算 2001/11/08 S.Sano
        .HWFONMBP = rs("E025HWFONMBP")            ' 品ＷＦ酸素濃度面内分布
        .HWFONAMN = rs("E025HWFONAMN")            ' 品ＷＦ酸素濃度平均下限
        .HWFONAMX = rs("E025HWFONAMX")            ' 品ＷＦ酸素濃度平均上限

        .HWFOS1MN = rs("E025HWFOS1MN")            ' 品ＷＦ酸素析出１下限
        .HWFOS1MX = rs("E025HWFOS1MX")            ' 品ＷＦ酸素析出１上限
        .HWFOS1SH = rs("E025HWFOS1SH")            ' 品ＷＦ酸素析出１測定位置＿方
        .HWFOS1ST = rs("E025HWFOS1ST")            ' 品ＷＦ酸素析出１測定位置＿点
        .HWFOS1SI = rs("E025HWFOS1SI")            ' 品ＷＦ酸素析出１測定位置＿位
        .HWFOS1HT = rs("E025HWFOS1HT")            ' 品ＷＦ酸素析出１保証方法＿対
        .HWFOS1HS = rs("E025HWFOS1HS")            ' 品ＷＦ酸素析出１保証方法＿処
        .HWFOS2SH = rs("E025HWFOS2SH")            ' 品ＷＦ酸素析出２測定位置＿方
        .HWFOS2ST = rs("E025HWFOS2ST")            ' 品ＷＦ酸素析出２測定位置＿点
        .HWFOS2SI = rs("E025HWFOS2SI")            ' 品ＷＦ酸素析出２測定位置＿位
        .HWFOS2MN = rs("E025HWFOS2MN")            ' 品ＷＦ酸素析出２下限
        .HWFOS2MX = rs("E025HWFOS2MX")            ' 品ＷＦ酸素析出２上限
        .HWFOS2HT = rs("E025HWFOS2HT")            ' 品ＷＦ酸素析出２保証方法＿対
        .HWFOS2HS = rs("E025HWFOS2HS")            ' 品ＷＦ酸素析出２保証方法＿処
        .HWFOS3MN = rs("E025HWFOS3MN")            ' 品ＷＦ酸素析出３下限
        .HWFOS3MX = rs("E025HWFOS3MX")            ' 品ＷＦ酸素析出３上限
        .HWFOS3SH = rs("E025HWFOS3SH")            ' 品ＷＦ酸素析出３測定位置＿方
        .HWFOS3ST = rs("E025HWFOS3ST")            ' 品ＷＦ酸素析出３測定位置＿点
        .HWFOS3SI = rs("E025HWFOS3SI")            ' 品ＷＦ酸素析出３測定位置＿位
        .HWFOS3HT = rs("E025HWFOS3HT")            ' 品ＷＦ酸素析出３保証方法＿対
        .HWFOS3HS = rs("E025HWFOS3HS")            ' 品ＷＦ酸素析出３保証方法＿処

        .HWFDSOMX = rs("E026HWFDSOMX")            ' 品ＷＦＤＳＯＤ上限
        .HWFDSOMN = rs("E026HWFDSOMN")            ' 品ＷＦＤＳＯＤ下限
        .HWFDSOAX = rs("E026HWFDSOAX")            ' 品ＷＦＤＳＯＤ領域上限
        .HWFDSOAN = rs("E026HWFDSOAN")            ' 品ＷＦＤＳＯＤ領域下限
        .HWFDSOHT = rs("E026HWFDSOHT")            ' 品ＷＦＤＳＯＤ保証方法＿対
        .HWFDSOHS = rs("E026HWFDSOHS")            ' 品ＷＦＤＳＯＤ保証方法＿処

        .HWFSPVMX = rs("E028HWFSPVMX")            ' 品ＷＦＳＰＶＦＥ上限
        .HWFSPVSH = rs("E028HWFSPVSH")            ' 品ＷＦＳＰＶＦＥ測定位置＿方
        .HWFSPVST = rs("E028HWFSPVST")            ' 品ＷＦＳＰＶＦＥ測定位置＿点
        .HWFSPVSI = rs("E028HWFSPVSI")            ' 品ＷＦＳＰＶＦＥ測定位置＿位
        .HWFSPVHT = rs("E028HWFSPVHT")            ' 品ＷＦＳＰＶＦＥ保証方法＿対
        .HWFSPVHS = rs("E028HWFSPVHS")            ' 品ＷＦＳＰＶＦＥ保証方法＿処
        .HWFDLSPH = rs("E028HWFDLSPH")            ' 品ＷＦ拡散長測定位置＿方
        .HWFDLSPT = rs("E028HWFDLSPT")            ' 品ＷＦ拡散長測定位置＿点
        .HWFDLSPI = rs("E028HWFDLSPI")            ' 品ＷＦ拡散長測定位置＿位
        .HWFDLHWT = rs("E028HWFDLHWT")            ' 品ＷＦ拡散長保証方法＿対
        .HWFDLHWS = rs("E028HWFDLHWS")            ' 品ＷＦ拡散長保証方法＿処
        .HWFDLMIN = rs("E028HWFDLMIN")            ' 品ＷＦ拡散長下限
        .HWFDLMAX = rs("E028HWFDLMAX")            ' 品ＷＦ拡散長上限

        .HWFOF1AX = rs("E029HWFOF1AX")           ' 品ＷＦＯＳＦ１平均上限
        .HWFOF1MX = rs("E029HWFOF1MX")           ' 品ＷＦＯＳＦ１上限
        .HWFOF1SH = rs("E029HWFOF1SH")           ' 品ＷＦＯＳＦ１測定位置＿方
        .HWFOF1ST = rs("E029HWFOF1ST")           ' 品ＷＦＯＳＦ１測定位置＿点
        .HWFOF1SR = rs("E029HWFOF1SR")           ' 品ＷＦＯＳＦ１測定位置＿領
        .HWFOF1HT = rs("E029HWFOF1HT")           ' 品ＷＦＯＳＦ１保証方法＿対
        .HWFOF1HS = rs("E029HWFOF1HS")           ' 品ＷＦＯＳＦ１保証方法＿処
        .HWFOF2AX = rs("E029HWFOF2AX")           ' 品ＷＦＯＳＦ２平均上限
        .HWFOF2MX = rs("E029HWFOF2MX")           ' 品ＷＦＯＳＦ２上限
        .HWFOF2SH = rs("E029HWFOF2SH")           ' 品ＷＦＯＳＦ２測定位置＿方
        .HWFOF2ST = rs("E029HWFOF2ST")           ' 品ＷＦＯＳＦ２測定位置＿点
        .HWFOF2SR = rs("E029HWFOF2SR")           ' 品ＷＦＯＳＦ２測定位置＿領
        .HWFOF2HT = rs("E029HWFOF2HT")           ' 品ＷＦＯＳＦ２保証方法＿対
        .HWFOF2HS = rs("E029HWFOF2HS")           ' 品ＷＦＯＳＦ２保証方法＿処
        .HWFOF3AX = rs("E029HWFOF3AX")           ' 品ＷＦＯＳＦ３平均上限
        .HWFOF3MX = rs("E029HWFOF3MX")           ' 品ＷＦＯＳＦ３上限
        .HWFOF3SH = rs("E029HWFOF3SH")           ' 品ＷＦＯＳＦ３測定位置＿方
        .HWFOF3ST = rs("E029HWFOF3ST")           ' 品ＷＦＯＳＦ３測定位置＿点
        .HWFOF3SR = rs("E029HWFOF3SR")           ' 品ＷＦＯＳＦ３測定位置＿領
        .HWFOF3HT = rs("E029HWFOF3HT")           ' 品ＷＦＯＳＦ３保証方法＿対
        .HWFOF3HS = rs("E029HWFOF3HS")           ' 品ＷＦＯＳＦ３保証方法＿処
        .HWFOF4AX = rs("E029HWFOF4AX")           ' 品ＷＦＯＳＦ４平均上限
        .HWFOF4MX = rs("E029HWFOF4MX")           ' 品ＷＦＯＳＦ４上限
        .HWFOF4SH = rs("E029HWFOF4SH")           ' 品ＷＦＯＳＦ４測定位置＿方
        .HWFOF4ST = rs("E029HWFOF4ST")           ' 品ＷＦＯＳＦ４測定位置＿点
        .HWFOF4SR = rs("E029HWFOF4SR")           ' 品ＷＦＯＳＦ４測定位置＿領
        .HWFOF4HT = rs("E029HWFOF4HT")           ' 品ＷＦＯＳＦ４保証方法＿対
        .HWFOF4HS = rs("E029HWFOF4HS")           ' 品ＷＦＯＳＦ４保証方法＿処
        If IsNull(rs("E029HWFOSF1PTK")) = False Then .HWFOSF1PTK = rs("E029HWFOSF1PTK")       ' 品ＷＦＯＳＦ１パタン区分　▼2003/05/14 ooba
        If IsNull(rs("E029HWFOSF2PTK")) = False Then .HWFOSF2PTK = rs("E029HWFOSF2PTK")       ' 品ＷＦＯＳＦ２パタン区分
        If IsNull(rs("E029HWFOSF3PTK")) = False Then .HWFOSF3PTK = rs("E029HWFOSF3PTK")       ' 品ＷＦＯＳＦ３パタン区分
        If IsNull(rs("E029HWFOSF4PTK")) = False Then .HWFOSF4PTK = rs("E029HWFOSF4PTK")       ' 品ＷＦＯＳＦ４パタン区分　▲2003/05/14 ooba

        'BMDべき乗数変更対応　2003/05/19 osawa
        .HWFBM1AN = rs("E029HWFBM1AN")           ' 品ＷＦＢＭＤ１平均下限
        .HWFBM1AX = rs("E029HWFBM1AX")           ' 品ＷＦＢＭＤ１平均上限
        .HWFBM1SH = rs("E029HWFBM1SH")           ' 品ＷＦＢＭＤ１測定位置＿方
        .HWFBM1ST = rs("E029HWFBM1ST")           ' 品ＷＦＢＭＤ１測定位置＿点
        .HWFBM1SR = rs("E029HWFBM1SR")           ' 品ＷＦＢＭＤ１測定位置＿領
        .HWFBM1HT = rs("E029HWFBM1HT")           ' 品ＷＦＢＭＤ１保証方法＿対
        .HWFBM1HS = rs("E029HWFBM1HS")           ' 品ＷＦＢＭＤ１保証方法＿処

        'BMDべき乗数変更対応　2003/05/19 osawa
        .HWFBM2AN = rs("E029HWFBM2AN")           ' 品ＷＦＢＭＤ２平均下限
        .HWFBM2AX = rs("E029HWFBM2AX")           ' 品ＷＦＢＭＤ２平均上限
        .HWFBM2SH = rs("E029HWFBM2SH")           ' 品ＷＦＢＭＤ２測定位置＿方
        .HWFBM2ST = rs("E029HWFBM2ST")           ' 品ＷＦＢＭＤ２測定位置＿点
        .HWFBM2SR = rs("E029HWFBM2SR")           ' 品ＷＦＢＭＤ２測定位置＿領
        .HWFBM2HT = rs("E029HWFBM2HT")           ' 品ＷＦＢＭＤ２保証方法＿対
        .HWFBM2HS = rs("E029HWFBM2HS")           ' 品ＷＦＢＭＤ２保証方法＿処

        'BMDべき乗数変更対応　2003/05/19 osawa
        .HWFBM3AN = rs("E029HWFBM3AN")           ' 品ＷＦＢＭＤ３平均下限
        .HWFBM3AX = rs("E029HWFBM3AX")           ' 品ＷＦＢＭＤ３平均上限
        .HWFBM3SH = rs("E029HWFBM3SH")           ' 品ＷＦＢＭＤ３測定位置＿方
        .HWFBM3ST = rs("E029HWFBM3ST")           ' 品ＷＦＢＭＤ３測定位置＿点
        .HWFBM3SR = rs("E029HWFBM3SR")           ' 品ＷＦＢＭＤ３測定位置＿領
        .HWFBM3HT = rs("E029HWFBM3HT")           ' 品ＷＦＢＭＤ３保証方法＿対
        .HWFBM3HS = rs("E029HWFBM3HS")           ' 品ＷＦＢＭＤ３保証方法＿処

        If IsNull(rs("E029HWFBM1MBP")) = True Then    ' 品ＷＦＢＭＤ１面内分布　▼2003/05/14 ooba
            .HWFBM1MBP = -1
        Else
            .HWFBM1MBP = rs("E029HWFBM1MBP")
        End If

        If IsNull(rs("E029HWFBM2MBP")) = True Then    ' 品ＷＦＢＭＤ２面内分布
            .HWFBM2MBP = -1
        Else
            .HWFBM2MBP = rs("E029HWFBM2MBP")
        End If

        If IsNull(rs("E029HWFBM3MBP")) = True Then    ' 品ＷＦＢＭＤ３面内分布
            .HWFBM3MBP = -1
        Else
            .HWFBM3MBP = rs("E029HWFBM3MBP")
        End If

        If IsNull(rs("E029HWFBM1MCL")) = False Then .HWFBM1MCL = rs("E029HWFBM1MCL")         ' 品ＷＦＢＭＤ１面内計算
        If IsNull(rs("E029HWFBM2MCL")) = False Then .HWFBM2MCL = rs("E029HWFBM2MCL")         ' 品ＷＦＢＭＤ２面内計算
        If IsNull(rs("E029HWFBM3MCL")) = False Then .HWFBM3MCL = rs("E029HWFBM3MCL")         ' 品ＷＦＢＭＤ３面内計算　▲2003/05/14 ooba

        .HWFOS1NS = rs("E025HWFOS1NS")           ' 品ＷＦ酸素析出１熱処理法
        .HWFOS2NS = rs("E025HWFOS2NS")           ' 品ＷＦ酸素析出２熱処理法
        .HWFOS3NS = rs("E025HWFOS3NS")           ' 品ＷＦ酸素析出３熱処理法
        .HWFOF1NS = rs("E029HWFOF1NS")           ' 品ＷＦＯＳＦ１熱処理法
        .HWFOF2NS = rs("E029HWFOF2NS")           ' 品ＷＦＯＳＦ２熱処理法
        .HWFOF3NS = rs("E029HWFOF3NS")           ' 品ＷＦＯＳＦ３熱処理法
        .HWFOF4NS = rs("E029HWFOF4NS")           ' 品ＷＦＯＳＦ４熱処理法
        .HWFBM1NS = rs("E029HWFBM1NS")           ' 品ＷＦＢＭＤ１熱処理法
        .HWFBM2NS = rs("E029HWFBM2NS")           ' 品ＷＦＢＭＤ２熱処理法
        .HWFBM3NS = rs("E029HWFBM3NS")           ' 品ＷＦＢＭＤ３熱処理法

        .HWFANTIM = rs("E025HWFANTIM")           ' 品ＷＦＡＮ時間
        .HWFANTNP = rs("E025HWFANTNP")           ' 品ＷＦＡＮ温度

        .HWFOF1ET = rs("E029HWFOF1ET")           ' 品ＷＦＯＳＦ１選択ＥＴ代
        .HWFOF2ET = rs("E029HWFOF2ET")           ' 品ＷＦＯＳＦ２選択ＥＴ代
        .HWFOF3ET = rs("E029HWFOF3ET")           ' 品ＷＦＯＳＦ３選択ＥＴ代
        .HWFOF4ET = rs("E029HWFOF4ET")           ' 品ＷＦＯＳＦ４選択ＥＴ代
        .HWFBM1ET = rs("E029HWFBM1ET")           ' 品ＷＦＢＭＤ１選択ＥＴ代
        .HWFBM2ET = rs("E029HWFBM2ET")           ' 品ＷＦＢＭＤ２選択ＥＴ代
        .HWFBM3ET = rs("E029HWFBM3ET")           ' 品ＷＦＢＭＤ３選択ＥＴ代

        .HWFOF1SZ = rs("E029HWFOF1SZ")           ' 品ＷＦＯＳＦ１測定条件
        .HWFOF2SZ = rs("E029HWFOF2SZ")           ' 品ＷＦＯＳＦ２測定条件
        .HWFOF3SZ = rs("E029HWFOF3SZ")           ' 品ＷＦＯＳＦ３測定条件
        .HWFOF4SZ = rs("E029HWFOF4SZ")           ' 品ＷＦＯＳＦ４測定条件
        .HWFBM1SZ = rs("E029HWFBM1SZ")           ' 品ＷＦＢＭＤ１測定条件
        .HWFBM2SZ = rs("E029HWFBM2SZ")           ' 品ＷＦＢＭＤ２測定条件
        .HWFBM3SZ = rs("E029HWFBM3SZ")           ' 品ＷＦＢＭＤ３測定条件
    End With

    '測定評価結果取得
    sDBName = "Y013"
    If DBDRV_GetTBCMY013(Sokutei(), " where SAMPLEID='" & typIn.SAMPLEID & "' ", "order by OSITEM") = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmlc001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '2001/08/20　ビューからブロックIDを取得するように修正
    'ブロックID取得
    sDBName = "E040"

    'XSDCBを使った場合
    sSQL = "select "
    sSQL = sSQL & " BLOCKID "
    sSQL = sSQL & " from "
    sSQL = sSQL & " VECME013xsb "
    sSQL = sSQL & " where "
    sSQL = sSQL & " CRYNUM = '" & typ_AType.typ_Param.CRYNUMCA & "' "
    sSQL = sSQL & " and INPOSCB = " & typ_AType.typ_Param.INPOSCA & " "
    sSQL = sSQL & " and INGOTPOS >= 0 "

    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    'レコード0件はエラー終了
    If rs.RecordCount = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDBName)
        DBDRV_scmzc_fcmlc001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    lngRecCnt = rs.RecordCount

    ReDim siyou.BLOCKID(lngRecCnt)

    For i = 1 To lngRecCnt
        ' SXL管理
        siyou.BLOCKID(i) = rs("BLOCKID")           ' ブロックID
        rs.MoveNext
    Next
    rs.Close


'2002/01/30 S.Sano Start
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
    sSQL = sSQL & " where K01.HINBAN='" & typIn.HIN.hinban & "' and "
    sSQL = sSQL & " K01.MNOREVNO=" & typIn.HIN.mnorevno & " and "
    sSQL = sSQL & " K01.FACTORY='" & typIn.HIN.factory & "' and "
    sSQL = sSQL & " K01.OPECOND='" & typIn.HIN.opecond & "' and "
    sSQL = sSQL & " K12.HINBAN='" & typIn.HIN.hinban & "' and "
    sSQL = sSQL & " K12.MNOREVNO=" & typIn.HIN.mnorevno & " and "
    sSQL = sSQL & " K12.FACTORY='" & typIn.HIN.factory & "' and "
    sSQL = sSQL & " K12.OPECOND='" & typIn.HIN.opecond & "'"

    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    'レコード0件はエラー終了
    If rs.RecordCount = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDBName)
        DBDRV_scmzc_fcmlc001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    With siyou
        .HWFRSPOT = rs("HSXRSPOT") ' 品ＳＸ比抵抗測定位置＿点
        .HWFRSPOI = rs("HSXRSPOI") ' 品ＳＸ比抵抗測定位置＿位
        .HWFONSPT = rs("HSXONSPT") ' 品ＳＸ酸素濃度測定位置＿点
        .HWFONSPI = rs("HSXONSPI") ' 品ＳＸ酸素濃度測定位置＿位
    End With

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    DBDRV_scmzc_fcmlc001c_Disp = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'***********************************************************************************************
'*    関数名        : DBDRV_scmzc_fcmlc001c_InsWfSougou
'*
'*    処理概要      : 1.WF総合判定 WF総合判定実績挿入用ＤＢドライバ
'*
'*    パラメータ    : 変数名        ,IO  ,型               ,説明
'*                  :WFSougou       ,I   ,typ_TBCMW005     ,WF総合判定実績挿入用
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'***********************************************************************************************
Public Function DBDRV_scmzc_fcmlc001c_InsWfSougou(WfSougou As typ_TBCMW005) As FUNCTION_RETURN
    Dim sSQL As String

    'WF総合判定実績への挿入
    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001c_InsWfSougou"

    sSQL = "insert into TBCMW005 ( "
    sSQL = sSQL & "CRYNUM, "           ' 結晶番号
    sSQL = sSQL & "INGOTPOS, "         ' インゴット内位置
    sSQL = sSQL & "TRANCNT, "          ' 処理回数
    sSQL = sSQL & "CRYLEN, "           ' 長さ
    sSQL = sSQL & "KRPROCCD, "         ' 管理工程コード
    sSQL = sSQL & "PROCCODE, "         ' 工程コード
    sSQL = sSQL & "SXLID, "            ' SXLID
    sSQL = sSQL & "CODE, "             ' 区分コード
    sSQL = sSQL & "TSTAFFID, "         ' 登録社員ID
    sSQL = sSQL & "REGDATE, "          ' 登録日付
    sSQL = sSQL & "KSTAFFID, "         ' 更新社員ID
    sSQL = sSQL & "UPDDATE, "          ' 更新日付
    sSQL = sSQL & "SENDFLAG, "         ' 送信フラグ
    sSQL = sSQL & "SENDDATE, "        ' 送信日付
    sSQL = sSQL & "PLANTCAT) "          ' 向先

    With WfSougou
        sSQL = sSQL & " select "
        sSQL = sSQL & " '" & .CRYNUM & "', "           ' 結晶番号
        sSQL = sSQL & " " & .INGOTPOS & ", "           ' インゴット内位置
        sSQL = sSQL & " nvl(max(TRANCNT),0)+1, "       ' 処理回数
        sSQL = sSQL & " " & .CRYLEN & ", "             ' 長さ
        sSQL = sSQL & " '" & .KRPROCCD & "', "         ' 管理工程コード
        sSQL = sSQL & " '" & .PROCCODE & "', "         ' 工程コード
        sSQL = sSQL & " '" & .SXLID & "', "            ' SXLID
        sSQL = sSQL & " '" & .CODE & "', "             ' 区分コード
        sSQL = sSQL & " '" & .TSTAFFID & "', "         ' 登録社員ID
        sSQL = sSQL & " sysdate, "                     ' 登録日付
        sSQL = sSQL & " '" & .TSTAFFID & "', "         ' 更新社員ID
        sSQL = sSQL & " sysdate, "                     ' 更新日付
        sSQL = sSQL & " '0', "                         ' 送信フラグ
        sSQL = sSQL & " sysdate "                      ' 送信日付
        sSQL = sSQL & " , '" & sCmbMukesaki & "'"      ' 向先 2007/09/04 SPK Tsutsumi Add
        sSQL = sSQL & " from TBCMW005 "
        sSQL = sSQL & " where CRYNUM='" & .CRYNUM & "' "
        sSQL = sSQL & " and INGOTPOS=" & .INGOTPOS
    End With

    If 0 >= OraDB.ExecuteSQL(sSQL) Then
        DBDRV_scmzc_fcmlc001c_InsWfSougou = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_scmzc_fcmlc001c_InsWfSougou = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    DBDRV_scmzc_fcmlc001c_InsWfSougou = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'************************************************************************************************************
'*    関数名        : DBDRV_scmzc_fcmlc001c_UpdGDdata
'*
'*    処理概要      : 1.WF総合判定 WF_GD実績更新用ＤＢドライバ
'*
'*    パラメータ    : 変数名       ,IO  ,型                  ,説明
'*                  : UpdGD        ,I   ,typ_TBCMJ015        ,WF_GD実績更新用
'*                  : sStaffID     ,I   ,sStaffID            ,担当者ｺｰﾄﾞ
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'************************************************************************************************************
Public Function DBDRV_scmzc_fcmlc001c_UpdGDdata(UpdGD As typ_TBCMJ015, sStaffID As String) As FUNCTION_RETURN
    Dim sSQL As String

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001c_UpdGDdata"

    '05/10/25 ooba START ============================================================>
    If UpdGD.MSRSDEN = -1 And UpdGD.MSRSLDL = -1 And UpdGD.MSRSDVD2 = -1 Then
        DBDRV_scmzc_fcmlc001c_UpdGDdata = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If
    '05/10/25 ooba END ==============================================================>

    With UpdGD
        sSQL = "UPDATE TBCMJ015 "
        sSQL = sSQL & "SET "

        If .MSRSDEN <> -1 Then      '05/10/25 ooba
            sSQL = sSQL & "MSRSDEN = " & .MSRSDEN & ", "        ' 測定結果 Den
        End If

        If .MSRSLDL <> -1 Then      '05/10/25 ooba
            sSQL = sSQL & "MSRSLDL = " & .MSRSLDL & ", "        ' 測定結果 L/DL
        End If

        If .MSRSDVD2 <> -1 Then     '05/10/25 ooba
            sSQL = sSQL & "MSRSDVD2 = " & .MSRSDVD2 & ", "      ' 測定結果 DVD2
        End If

'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
        If .MSZEROMN <> -1 Then
            sSQL = sSQL & "MSZEROMN = " & .MSZEROMN & ", "      ' L/DL0連続数最小値
        End If
        If .MSZEROMX <> -1 Then
            sSQL = sSQL & "MSZEROMX = " & .MSZEROMX & ", "      ' L/DL0連続数最大値
        End If
        sSQL = sSQL & "PTNJUDGRES = '" & .PTNJUDGRES & "', "      ' パターン判定結果
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End

        sSQL = sSQL & "KSTAFFID = '" & sStaffID & "', "         ' 更新社員ID
        sSQL = sSQL & "UPDDATE = SYSDATE "                      ' 更新日付
        sSQL = sSQL & "WHERE "
        sSQL = sSQL & "CRYNUM = '" & .CRYNUM & "' "             ' 結晶番号
        sSQL = sSQL & "AND POSITION = " & .POSITION & " "       ' 位置
        sSQL = sSQL & "AND SMPKBN = '" & .SMPKBN & "' "         ' サンプル区分
        sSQL = sSQL & "AND TRANCOND = '" & .TRANCOND & "' "     ' 処理条件
        sSQL = sSQL & "AND TRANCNT = " & .TRANCNT & " "         ' 処理回数
        sSQL = sSQL & "AND HSFLG = '" & .HSFLG & "' "           ' 保証フラグ
        sSQL = sSQL & "AND SMPLNO = '" & .SMPLNO & "' "         ' サンプルＮｏ
    End With

    If 0 >= OraDB.ExecuteSQL(sSQL) Then
        DBDRV_scmzc_fcmlc001c_UpdGDdata = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_scmzc_fcmlc001c_UpdGDdata = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    DBDRV_scmzc_fcmlc001c_UpdGDdata = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'***********************************************************************************************************
'*    関数名        : DBDRV_scmzc_fcmlc001c_UpdSXL1
'*
'*    処理概要      : 1.WF総合判定 SXL管理更新用ＤＢドライバ（現在工程、最終通過工程更新）
'*                      (現在は、Trueを返しているだけ)
'*
'*    パラメータ    : 変数名        ,IO ,型                                  ,説明
'*                    SXL           ,O  ,type_DBDRV_scmzc_fcmlc001c_UpdSXL1  ,SXL管理
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'***********************************************************************************************************
Public Function DBDRV_scmzc_fcmlc001c_UpdSXL1(sxl As type_DBDRV_scmzc_fcmlc001c_UpdSXL1) As FUNCTION_RETURN
''↓追加START SXL管理（E042）→XSDCB機能移行 '06/1/5 SMP岡本
    DBDRV_scmzc_fcmlc001c_UpdSXL1 = FUNCTION_RETURN_SUCCESS
''↑追加END   SXL管理（E042）→XSDCB機能移行 '06/1/5 SMP岡本
End Function

'***********************************************************************************************************
'*    関数名        : GetSxlidINBlkid
'*
'*    処理概要      : 1.WF総合判定 SXL管理更新用ＤＢドライバ（削除区分、最終状態区分更新）
'*
'*    パラメータ    : 変数名       ,IO  ,型                 ,説明
'*                    WFSoku       ,I   ,typ_TBCMW009       ,WFセンター総合判定測定値挿入用
'*                    WFSougou     ,I   ,typ_TBCMW005       ,WF総合判定実績挿入用
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'***********************************************************************************************************
Public Function DBDRV_scmzc_fcmlc001c_UpdSXL2(sxl As type_DBDRV_scmzc_fcmlc001c_UpdSXL2) As FUNCTION_RETURN
    Dim sSQL As String

    'SXL管理の更新

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001c_UpdSXL2"

    sSQL = "update TBCME042 set "
    sSQL = sSQL & " DELCLS='" & sxl.DELCLS & "', "
    sSQL = sSQL & " LSTATCLS='" & sxl.LSTATCLS & "', "
    sSQL = sSQL & " UPDDATE=sysdate, "
    sSQL = sSQL & " SENDFLAG='0' "
    sSQL = sSQL & " where "
    sSQL = sSQL & " CRYNUM='" & sxl.CRYNUM & "' "
    sSQL = sSQL & " and INGOTPOS=" & sxl.INGOTPOS

    If 0 >= OraDB.ExecuteSQL(sSQL) Then
        DBDRV_scmzc_fcmlc001c_UpdSXL2 = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_scmzc_fcmlc001c_UpdSXL2 = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    Resume proc_exit
End Function

'********************************************************************************************
'*    関数名        : GetSxlidINBlkid
'*
'*    処理概要      : 1.WF総合判定 振替廃棄実績挿入用ＤＢドライバ
'*
'*    パラメータ    : 変数名       ,IO ,型                   ,説明
'*                    Hurikae      ,I  ,typ_TBCMW006         ,振替廃棄実績挿入用
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'********************************************************************************************
Public Function DBDRV_scmzc_fcmlc001c_InsHurikae(Hurikae As typ_TBCMW006) As FUNCTION_RETURN
    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001c_InsHurikae"

    '振替廃棄実績への挿入
    If DBDRV_Furikae_Ins(Hurikae) = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmlc001c_InsHurikae = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_scmzc_fcmlc001c_InsHurikae = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'******************************************************************************************
'*    関数名        : DBDRV_scmzc_fcmlc001c_InsSxlKakutei
'*
'*    処理概要      : 1.WF総合判定 SXL確定指示挿入用ＤＢドライバ
'*
'*    パラメータ    : 変数名       ,IO ,型                  ,説明
'*                    Hurikae      ,I  ,typ_TBCMY007        ,SXL確定指示
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*                    (使用していない)
'******************************************************************************************
Public Function DBDRV_scmzc_fcmlc001c_InsSxlKakutei(sxl As typ_TBCMY007) As FUNCTION_RETURN
    Dim sSQL As String

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001c_InsSxlKakutei"

    sSQL = "insert into TBCMY007 ("
    sSQL = sSQL & "SXL_ID, "           ' SXL-ID
    sSQL = sSQL & "SAMPLE_FROM, "      ' サンプルID (From)
    sSQL = sSQL & "SAMPLE_TO, "        ' サンプルID (To)
    sSQL = sSQL & "BLOCKID, "          ' ブロックＩＤ
    sSQL = sSQL & "HINBAN, "           ' 確定品番
    sSQL = sSQL & "KUBUN, "            ' 区分コード
    sSQL = sSQL & "TXID, "             ' トランザクションID
    sSQL = sSQL & "REGDATE, "          ' 登録日付
    sSQL = sSQL & "SUMMITSENDFLAG, "   ' SUMMIT送信フラグ
    sSQL = sSQL & "SENDFLAG, "         ' 送信フラグ
    sSQL = sSQL & "SENDDATE, "                ' 送信日付
    sSQL = sSQL & "PLANTCAT) "                ' 向先

    With sxl
        sSQL = sSQL & "values ("
        sSQL = sSQL & " '" & .SXL_ID & "', "           ' SXL-ID
        sSQL = sSQL & " '" & .SAMPLE_FROM & "', "      ' サンプルID (From)
        sSQL = sSQL & " '" & .SAMPLE_TO & "', "        ' サンプルID (To)
        sSQL = sSQL & " '" & .BLOCKID & "', "          ' ブロックＩＤ
        sSQL = sSQL & " '" & .hinban & "', "           ' 確定品番
        sSQL = sSQL & " '" & .KUBUN & "', "            ' 区分コード
        sSQL = sSQL & " '" & .TXID & "', "             ' トランザクションID
        sSQL = sSQL & " sysdate, "                     ' 登録日付
        sSQL = sSQL & " '0', "                         ' SUMMIT送信フラグ
        sSQL = sSQL & " '0', "                         ' 送信フラグ
        sSQL = sSQL & " sysdate , "                    ' 送信日付
        sSQL = sSQL & " , '" & sCmbMukesaki & "'"      ' 向先 2007/09/04 SPK Tsutsumi Add
    End With

    If 0 >= OraDB.ExecuteSQL(sSQL) Then
        DBDRV_scmzc_fcmlc001c_InsSxlKakutei = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_scmzc_fcmlc001c_InsSxlKakutei = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    DBDRV_scmzc_fcmlc001c_InsSxlKakutei = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'*******************************************************************************
'*    関数名        : GetSxlidINBlkid
'*
'*    処理概要      : 1.SXLの全ブロック入庫チェック
'*
'*    パラメータ    : 変数名        ,IO  ,型             ,説明
'*                    sCryNum       ,I   ,String         ,結晶番号
'*                    pTbcmh004()   ,O   ,typ_TBCMH004   ,引上げ終了実績取得用
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function s_cmmc001db_sSql039(ByVal sCryNum As String, _
                pTbcmh004() As typ_TBCMH004) As Double
    Dim sSQL    As String
    Dim intRet  As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function s_cmmc001db_sSql039"

    sSQL = " where CRYNUM = '" & sCryNum & "' "

    If DBDRV_GetTBCMH004(pTbcmh004, sSQL, "order by CRYNUM") = FUNCTION_RETURN_FAILURE Then
        s_cmmc001db_sSql039 = FUNCTION_RETURN_FAILURE
    Else
        s_cmmc001db_sSql039 = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    s_cmmc001db_sSql039 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'**********************************************************************************************************
'*    関数名        : DBDRV_scmzc_fcmlc001c_UpdWfCrySmp
'*
'*    処理概要      : 1.総合判定 WFサンプル管理更新用（確定区分を１に更新）
'*
'*    パラメータ    : 変数名       ,IO ,型                                       ,説明
'*                    WfCrySmp     ,O  ,type_DBDRV_scmzc_fcmlc001c_UpdWfCrySmp   ,結晶サンプル管理更新用
'*                    sSXLID       ,I  ,String                                   ,SXLID 09/05/25 ooba
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'**********************************************************************************************************
Public Function DBDRV_scmzc_fcmlc001c_UpdWfCrySmp(sSXLID As String) As FUNCTION_RETURN
''Public Function DBDRV_scmzc_fcmlc001c_UpdWfCrySmp(WfCrySmp() As type_DBDRV_scmzc_fcmlc001c_UpdWfCrySmp) _
''                 As FUNCTION_RETURN
    Dim sSQL As String
    Dim i   As Long

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001c_UpdCrySmp"

''    For i = 1 To UBound(WfCrySmp)
        ' WFサンプル管理の更新
        sSQL = "update XSDCW set "
        sSQL = sSQL & "  KTKBNCW='1' "          '確定区分
        sSQL = sSQL & ", KDAYCW=sysdate "
        sSQL = sSQL & ", SNDKCW='0' "
        sSQL = sSQL & " where SXLIDCW = '" & sSXLID & "' "      '条件変更(ｽﾋﾟｰﾄﾞ化) 09/05/25 ooba
''        sSql = sSql & " and XTALCW='" & WfCrySmp(i).CRYNUM & "' "
''        sSql = sSql & " and INPOSCW=" & WfCrySmp(i).INGOTPOS & " "
''        sSql = sSql & " and SMPKBNCW='" & WfCrySmp(i).SMPKBN & "' "

        If 0 >= OraDB.ExecuteSQL(sSQL) Then
            DBDRV_scmzc_fcmlc001c_UpdWfCrySmp = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
''     Next

     DBDRV_scmzc_fcmlc001c_UpdWfCrySmp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    DBDRV_scmzc_fcmlc001c_UpdWfCrySmp = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'*****************************************************************************************************
'*    関数名        : DBDRV_GetTBCMH001039
'*
'*    処理概要      : 1.テーブル「TBCMH001」から条件にあったレコードを抽出する
'*
'*    パラメータ    : 変数名       ,IO ,型           ,説明
'*                   records()     ,O  ,typ_TBCMH001 ,抽出レコード
'*                   sSqlWhere      ,I  ,String      ,抽出条件(SQLのWhere節:省略可能)
'*                   sSqlOrder      ,I  ,String      ,抽出順序(SQLのOrder by節:省略可能)
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*****************************************************************************************************
Public Function DBDRV_GetTBCMH001039(records() As typ_TBCMH001, Optional sSqlWhere$ = vbNullString, _
                                        Optional sSqlOrder$ = vbNullString) As FUNCTION_RETURN
    Dim sSQL        As String       'SQL全体
    Dim sSqlBase    As String       'SQL基本部(WHERE節の前まで)
    Dim rs          As OraDynaset   'RecordSet
    Dim lngRecCnt   As Long         'レコード数
    Dim i           As Long

    ''SQLを組み立てる
    sSqlBase = "Select UPINDNO, KRPROCCD, PROCCODE, MODEL, GOUKI, PGID, CPORGIND, HINBAN, NMNOREVNO, NFACTORY, NOPECOND, NUMNOTE1," & _
              " NUMNOTE2, SEED, SEKIERTB, DPNTCLS, DOPANT, AMRESIST, CRYDOPCL, CRYDOPVL, UPBTCHNM, ADDDOPCL, ADDDOPVL, ADDDOPPT," & _
              " BCNT1COD, BCNT1CMT, BCNT2COD, BCNT2CMT, MTCLS1, MTWGHT1, ESWGHT1, MTCLS2, MTWGHT2, ESWGHT2, MTCLS3, MTWGHT3," & _
              " ESWGHT3, MTCLS4, MTWGHT4, ESWGHT4, MTCLS5, MTWGHT5, ESWGHT5, MTCLS6, MTWGHT6, ESWGHT6, MTCLS7, MTWGHT7, ESWGHT7," & _
              " MTCLS8, MTWGHT8, ESWGHT8, MTCLS9, MTWGHT9, ESWGHT9, MTCLS10, MTWGHT10, ESWGHT10, MTCLS11, MTWGHT11, ESWGHT11," & _
              " MTCLS12, MTWGHT12, ESWGHT12, MTCLS13, MTWGHT13, ESWGHT13, MTCLS14, MTWGHT14, ESWGHT14, MTCLS15, MTWGHT15," & _
              " ESWGHT15, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sSqlBase = sSqlBase & "From TBCMH001"
    sSQL = sSqlBase

    If (sSqlWhere <> vbNullString) Or (sSqlOrder <> vbNullString) Then
        sSQL = sSQL & " " & sSqlWhere & " " & sSqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMH001039 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    lngRecCnt = rs.RecordCount
    ReDim records(lngRecCnt)
    For i = 1 To lngRecCnt
        With records(i)
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

    DBDRV_GetTBCMH001039 = FUNCTION_RETURN_SUCCESS
End Function

'*****************************************************************************************************
'*    関数名        : DBDRV_GetTBCMH004
'*
'*    処理概要      : 1.テーブル「TBCMH004」から条件にあったレコードを抽出する
'*
'*    パラメータ    : 変数名        ,IO ,型                            ,説明
'*                   records()     ,O  ,typ_TBCMH001 ,抽出レコード
'*                   sSqlWhere      ,I  ,String      ,抽出条件(SQLのWhere節:省略可能)
'*                   sSqlOrder      ,I  ,String      ,抽出順序(SQLのOrder by節:省略可能)
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*****************************************************************************************************
Public Function DBDRV_GetTBCMH004(records() As typ_TBCMH004, Optional sSqlWhere$ = vbNullString, _
                                    Optional sSqlOrder$ = vbNullString) As FUNCTION_RETURN
    Dim sSQL        As String       'SQL全体
    Dim sSqlBase    As String       'SQL基本部(WHERE節の前まで)
    Dim rs          As OraDynaset   'RecordSet
    Dim lngRecCnt   As Long         'レコード数
    Dim i           As Long

    ''SQLを組み立てる
    sSqlBase = "Select CRYNUM, KRPROCCD, PROCCODE, LENGTOP, LENGTKDO, LENGTAIL, LENGFREE, DM1, DM2, DM3, WGHTTOP, WGHTTKDO," & _
              " WGHTTAIL, WGHTFREE, WGTOPCUT, UPWEIGHT, CHARGE, SEED, STATCLS, JDGECODE, PWTIME, ADDDPPOS, ADDDPCLS, ADDDPVAL," & _
              " ADDDPNAM, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE "
    sSqlBase = sSqlBase & "From TBCMH004"
    sSQL = sSqlBase
    If (sSqlWhere <> vbNullString) Or (sSqlOrder <> vbNullString) Then
        sSQL = sSQL & " " & sSqlWhere & " " & sSqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMH004 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    lngRecCnt = rs.RecordCount
    ReDim records(lngRecCnt)
    For i = 1 To lngRecCnt
        With records(i)
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

'***********************************************************************************************
'*    関数名        : DBDRV_scmzc_fcmlc001d_DispSiyou
'*
'*    処理概要      : 1.再抜試指示 表示用ＤＢドライバ（WF仕様）
'*
'*    パラメータ    : 変数名       ,IO   ,型                                 ,説明
'*                    typIn        ,I    ,type_DBDRV_scmzc_fcmlc001d_In      ,入力用
'*                    WfSiyou      ,I    ,type_DBDRV_scmzc_fcmlc001d_WfSiyou ,WF仕様取得用
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'***********************************************************************************************
Public Function DBDRV_scmzc_fcmlc001d_DispSiyou(typIn() As type_DBDRV_scmzc_fcmlc001d_In, _
                                            WfSiyou() As type_DBDRV_scmzc_fcmlc001d_WfSiyou, _
                                            sErrMsg As String _
                                            ) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim lngRecCnt   As Long
    Dim i           As Long
    Dim lngInCnt    As Long
    Dim sDBName     As String
    Dim sOT1        As String
    Dim sOT2        As String
    Dim rtn         As FUNCTION_RETURN
    Dim sMAI1       As String
    Dim sMAI2       As String

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001d_DispSiyou"

    lngInCnt = UBound(typIn)

    ReDim WfSiyou(lngInCnt)
    sDBName = "(V001)"

    For i = 1 To lngInCnt
        DoEvents
        ' WF仕様の取得
        sSQL = "select "
        sSQL = sSQL & "E021HWFRMIN, "            ' 品ＷＦ比抵抗下限
        sSQL = sSQL & "E021HWFRMAX, "            ' 品ＷＦ比抵抗上限
        sSQL = sSQL & "E021HWFRHWYS, "           ' 品ＷＦ比抵抗保証方法＿処(Rs)
        sSQL = sSQL & "E025HWFONHWS, "           ' 品ＷＦ酸素濃度保証方法＿処(Oi)
        sSQL = sSQL & "E029HWFBM1HS, "           ' 品ＷＦＢＭＤ１保証方法＿処(B1)
        sSQL = sSQL & "E029HWFBM2HS, "           ' 品ＷＦＢＭＤ２保証方法＿処(B2)
        sSQL = sSQL & "E029HWFBM3HS, "           ' 品ＷＦＢＭＤ３保証方法＿処(B3)
        sSQL = sSQL & "E029HWFOF1HS, "           ' 品ＷＦＯＳＦ１保証方法＿処(L1)
        sSQL = sSQL & "E029HWFOF2HS, "           ' 品ＷＦＯＳＦ２保証方法＿処(L2)
        sSQL = sSQL & "E029HWFOF3HS, "           ' 品ＷＦＯＳＦ３保証方法＿処(L3)
        sSQL = sSQL & "E029HWFOF4HS, "           ' 品ＷＦＯＳＦ４保証方法＿処(L4)
        sSQL = sSQL & "E026HWFDSOHS, "           ' 品ＷＦＤＳＯＤ保証方法＿処(DS)
        sSQL = sSQL & "E024HWFMKHWS, "           ' 品ＷＦ無欠陥層保証方法＿処(DZ)
        sSQL = sSQL & "E028HWFSPVHS, "           ' 品ＷＦＳＰＶＦＥ保証方法＿処(SP)
        sSQL = sSQL & "E028HWFDLHWS, "           ' 品ＷＦ拡散長保証方法＿処(KL)　06/06/08 ooba
        sSQL = sSQL & "E025HWFOS1HS, "           ' 品ＷＦ酸素析出１保証方法＿処(D1)
        sSQL = sSQL & "E025HWFOS2HS, "           ' 品ＷＦ酸素析出２保証方法＿処(D2)
        sSQL = sSQL & "E025HWFOS3HS "            ' 品ＷＦ酸素析出３保証方法＿処(D3)
        sSQL = sSQL & " from VECME001 "
        sSQL = sSQL & " where E018HINBAN='" & typIn(i).HIN.hinban & "' " & _
                    " and E018MNOREVNO=" & typIn(i).HIN.mnorevno & " " & _
                    " and E018FACTORY='" & typIn(i).HIN.factory & "' " & _
                    " and E018OPECOND='" & typIn(i).HIN.opecond & "' "

        Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        'レコード0件時はエラー
        If rs.RecordCount = 0 Then
            DBDRV_scmzc_fcmlc001d_DispSiyou = FUNCTION_RETURN_FAILURE
            ReDim WfSiyou(0)
            sErrMsg = GetMsgStr("EGET") & sDBName
            rs.Close
            GoTo proc_exit
        End If

        With WfSiyou(i)
            .HWFRMIN = fncNullCheck(rs("E021HWFRMIN"))            ' 品ＷＦ比抵抗下限
            .HWFRMAX = fncNullCheck(rs("E021HWFRMAX"))            ' 品ＷＦ比抵抗上限
            .HWFRHWYS = rs("E021HWFRHWYS")          ' 品ＷＦ比抵抗保証方法＿処(Rs)
            .HWFONHWS = rs("E025HWFONHWS")          ' 品ＷＦ酸素濃度保証方法＿処(Oi)
            .HWFBM1HS = rs("E029HWFBM1HS")          ' 品ＷＦＢＭＤ１保証方法＿処(B1)
            .HWFBM2HS = rs("E029HWFBM2HS")          ' 品ＷＦＢＭＤ２保証方法＿処(B2)
            .HWFBM3HS = rs("E029HWFBM3HS")          ' 品ＷＦＢＭＤ３保証方法＿処(B3)
            .HWFOF1HS = rs("E029HWFOF1HS")          ' 品ＷＦＯＳＦ１保証方法＿処(L1)
            .HWFOF2HS = rs("E029HWFOF2HS")          ' 品ＷＦＯＳＦ２保証方法＿処(L2)
            .HWFOF3HS = rs("E029HWFOF3HS")          ' 品ＷＦＯＳＦ３保証方法＿処(L3)
            .HWFOF4HS = rs("E029HWFOF4HS")          ' 品ＷＦＯＳＦ４保証方法＿処(L4)
            .HWFDSOHS = rs("E026HWFDSOHS")          ' 品ＷＦＤＳＯＤ保証方法＿処(DS)
            .HWFMKHWS = rs("E024HWFMKHWS")          ' 品ＷＦ無欠陥層保証方法＿処(DZ)
            .HWFSPVHS = rs("E028HWFSPVHS")          ' 品ＷＦＳＰＶＦＥ保証方法＿処(SP)
            .HWFDLHWS = rs("E028HWFDLHWS")          ' 品ＷＦ拡散長保証方法＿処(KL)　06/06/08 ooba
            .HWFOS1HS = rs("E025HWFOS1HS")          ' 品ＷＦ酸素析出１保証方法＿処(D1)
            .HWFOS2HS = rs("E025HWFOS2HS")          ' 品ＷＦ酸素析出２保証方法＿処(D2)
            .HWFOS3HS = rs("E025HWFOS3HS")          ' 品ＷＦ酸素析出３保証方法＿処(D3)
            'rtn = scmzc_getE036(typIn(i).HIN, sOT1, sOT2)    '03/05/26
            rtn = scmzc_getE036(typIn(i).HIN, sOT1, sOT2, sMAI1, sMAI2)  '04/07/16
            If rtn = FUNCTION_RETURN_FAILURE Then
                rs.Close
                DBDRV_scmzc_fcmlc001d_DispSiyou = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
            .HWFOT1 = sOT1 '### 03/05/26
            .HWFOT2 = sOT2
            .HWFMAI1 = sMAI1  '04/07/16
            .HWFMAI2 = sMAI2
        End With

        ''残存酸素仕様取得追加　03/12/15 ooba START ==============================>
        sSQL = "select HWFZOHWS from TBCME025 "
        sSQL = sSQL & "where HINBAN = '" & typIn(i).HIN.hinban & "' "
        sSQL = sSQL & "and MNOREVNO = " & typIn(i).HIN.mnorevno & " "
        sSQL = sSQL & "and FACTORY = '" & typIn(i).HIN.factory & "' "
        sSQL = sSQL & "and OPECOND = '" & typIn(i).HIN.opecond & "' "

        Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        'レコード0件時はエラー
        If rs.RecordCount = 0 Then
            DBDRV_scmzc_fcmlc001d_DispSiyou = FUNCTION_RETURN_FAILURE
            ReDim WfSiyou(0)
            sErrMsg = GetMsgStr("EGET") & "(E025)"
            rs.Close
            GoTo proc_exit
        End If

        If IsNull(rs("HWFZOHWS")) = False Then WfSiyou(i).HWFZOHWS = rs("HWFZOHWS") Else WfSiyou(i).HWFZOHWS = " "  '品WF残存酸素保証方法_処

        ''残存酸素仕様チェック
        iChkAoi = ChkAoiSiyou(typIn(i).HIN)
        If iChkAoi < 0 Then
            sErrMsg = "残存酸素(AOi)仕様エラー"
            DBDRV_scmzc_fcmlc001d_DispSiyou = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        ''残存酸素仕様取得追加　03/12/15 ooba END ================================>

        '' GD仕様取得　05/02/17 ooba START ==========================================>
        sDBName = "(E026)"
        sSQL = "select "
        sSQL = sSQL & "HWFDENHS, "                    ' 品ＷＦＤｅｎ保証方法＿処(GD)
        sSQL = sSQL & "HWFDVDHS, "                    ' 品ＷＦＤＶＤ２保証方法＿処(GD)
        sSQL = sSQL & "HWFLDLHS "                     ' 品ＷＦＬ／ＤＬ保証方法＿処(GD)
        sSQL = sSQL & "from TBCME026 "
        sSQL = sSQL & "where HINBAN = '" & typIn(i).HIN.hinban & "' "
        sSQL = sSQL & "and MNOREVNO = " & typIn(i).HIN.mnorevno & " "
        sSQL = sSQL & "and FACTORY = '" & typIn(i).HIN.factory & "' "
        sSQL = sSQL & "and OPECOND = '" & typIn(i).HIN.opecond & "' "

        Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        'レコード0件時はエラー
        If rs.RecordCount = 0 Then
            DBDRV_scmzc_fcmlc001d_DispSiyou = FUNCTION_RETURN_FAILURE
            ReDim WfSiyou(0)
            sErrMsg = GetMsgStr("EGET") & sDBName
            rs.Close
            GoTo proc_exit
        End If

        WfSiyou(i).HWFDENHS = rs("HWFDENHS")        ' 品ＷＦＤｅｎ保証方法＿処(GD)
        WfSiyou(i).HWFDVDHS = rs("HWFDVDHS")        ' 品ＷＦＤＶＤ２保証方法＿処(GD)
        WfSiyou(i).HWFLDLHS = rs("HWFLDLHS")        ' 品ＷＦＬ／ＤＬ保証方法＿処(GD)
        '' GD仕様取得　05/02/17 ooba END ============================================>

        '' SPVNr濃度仕様取得　06/06/08 ooba START ===========================>
        sDBName = "E048"
        sSQL = "select "
        sSQL = sSQL & "HWFNRHS ,"         '品WFSPVNR保証方法_処
        sSQL = sSQL & "HWFSIRDHS "        '軸状転位保証方法＿処　Add 2010/01/06 SIRD対応 Y.Hitomi
        sSQL = sSQL & "from TBCME048 "
        sSQL = sSQL & "where HINBAN = '" & typIn(i).HIN.hinban & "' "
        sSQL = sSQL & "and MNOREVNO = " & typIn(i).HIN.mnorevno & " "
        sSQL = sSQL & "and FACTORY = '" & typIn(i).HIN.factory & "' "
        sSQL = sSQL & "and OPECOND = '" & typIn(i).HIN.opecond & "' "

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            DBDRV_scmzc_fcmlc001d_DispSiyou = FUNCTION_RETURN_FAILURE
            ReDim WfSiyou(0)
            sErrMsg = GetMsgStr("EGET") & sDBName
            rs.Close
            GoTo proc_exit
        End If

        If IsNull(rs("HWFNRHS")) = False Then WfSiyou(i).HWFNRHS = rs("HWFNRHS") Else WfSiyou(i).HWFNRHS = " "
        'Add 2010/01/06 SIRD対応 Y.Hitomi
        If IsNull(rs("HWFSIRDHS")) = False Then WfSiyou(i).HWFSIRDHS = rs("HWFSIRDHS") Else WfSiyou(i).HWFSIRDHS = " "

        rs.Close
        '' SPVNr濃度仕様取得　06/06/08 ooba START ===========================>

        '計画長取得
    '2004.09.08 Y.K 紐付け変更
        sSQL = "select "
        sSQL = sSQL & " nvl(SUM(LENGTH),0) as Alllen"
        sSQL = sSQL & " from TBCME039 "
        sSQL = sSQL & " where substr(CRYNUM,1,9) = '" & Mid(typIn(i).CRYNUM, 1, 7) & "0" & Mid(typIn(i).CRYNUM, 9, 1) & "' " & _
                    " and HINBAN='" & typIn(i).HIN.hinban & "' " & _
                    " and REVNUM=" & typIn(i).HIN.mnorevno & " " & _
                    " and FACT='" & typIn(i).HIN.factory & "' " & _
                    " and OPCOND='" & typIn(i).HIN.opecond & "' "

        Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount = 0 Then
            DBDRV_scmzc_fcmlc001d_DispSiyou = FUNCTION_RETURN_FAILURE
            ReDim WfSiyou(0)
            rs.Close
            GoTo proc_exit
        End If

        WfSiyou(i).KEIKAKUL = rs("Alllen")            ' 計画長

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
        '' エピ仕様取得(OSF、BND)
        sDBName = "E050"
        sSQL = "select "
        sSQL = sSQL & "HEPOF1HS, "        '品EPOSF1保証方法_処
        sSQL = sSQL & "HEPOF2HS, "        '品EPOSF2保証方法_処
        sSQL = sSQL & "HEPOF3HS, "        '品EPOSF3保証方法_処
        sSQL = sSQL & "HEPBM1HS, "        '品EPBMD1保証方法_処
        sSQL = sSQL & "HEPBM2HS, "        '品EPBMD2保証方法_処
        sSQL = sSQL & "HEPBM3HS "         '品EPBMD3保証方法_処
        sSQL = sSQL & "from TBCME050 "    '製品仕様エピデータ１
        sSQL = sSQL & "where HINBAN = '" & typIn(i).HIN.hinban & "' "
        sSQL = sSQL & "and MNOREVNO = " & typIn(i).HIN.mnorevno & " "
        sSQL = sSQL & "and FACTORY = '" & typIn(i).HIN.factory & "' "
        sSQL = sSQL & "and OPECOND = '" & typIn(i).HIN.opecond & "' "

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            DBDRV_scmzc_fcmlc001d_DispSiyou = FUNCTION_RETURN_FAILURE
            ReDim WfSiyou(0)
            sErrMsg = GetMsgStr("EGET") & sDBName
            rs.Close
            GoTo proc_exit
        End If

        With WfSiyou(i)
            If IsNull(rs("HEPOF1HS")) = False Then .HEPOF1HS = rs("HEPOF1HS") Else .HEPOF1HS = " "   '品EPOSF1保証方法_処
            If IsNull(rs("HEPOF2HS")) = False Then .HEPOF2HS = rs("HEPOF2HS") Else .HEPOF2HS = " "   '品EPOSF2保証方法_処
            If IsNull(rs("HEPOF3HS")) = False Then .HEPOF3HS = rs("HEPOF3HS") Else .HEPOF3HS = " "   '品EPOSF3保証方法_処
            If IsNull(rs("HEPBM1HS")) = False Then .HEPBM1HS = rs("HEPBM1HS") Else .HEPBM1HS = " "   '品EPBMD1保証方法_処
            If IsNull(rs("HEPBM2HS")) = False Then .HEPBM2HS = rs("HEPBM2HS") Else .HEPBM2HS = " "   '品EPBMD2保証方法_処
            If IsNull(rs("HEPBM3HS")) = False Then .HEPBM3HS = rs("HEPBM3HS") Else .HEPBM3HS = " "   '品EPBMD3保証方法_処
        End With
        rs.Close
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
        '>>>>> 中間抜試規格セット追加 2011/07/15 Marushita
        sDBName = "E036"
        sSQL = "select "
        sSQL = sSQL & "MSMPFLG,     "     '中間抜試フラグ
        sSQL = sSQL & "MSMPTANIMAI, "     '中間抜試単位(枚)
        sSQL = sSQL & "MSMPCONSTMAI "     '中間抜試許容値
        sSQL = sSQL & "from TBCME036 "
        sSQL = sSQL & "where HINBAN = '" & typIn(i).HIN.hinban & "' "
        sSQL = sSQL & "and MNOREVNO = " & typIn(i).HIN.mnorevno & " "
        sSQL = sSQL & "and FACTORY = '" & typIn(i).HIN.factory & "' "
        sSQL = sSQL & "and OPECOND = '" & typIn(i).HIN.opecond & "' "

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            DBDRV_scmzc_fcmlc001d_DispSiyou = FUNCTION_RETURN_FAILURE
            ReDim WfSiyou(0)
            sErrMsg = GetMsgStr("EGET") & sDBName
            rs.Close
            GoTo proc_exit
        End If

        With WfSiyou(i)
            If IsNull(rs("MSMPFLG")) = False Then .CHUFLG = rs("MSMPFLG") Else .CHUFLG = "0"            '中間抜試フラグ
            If IsNull(rs("MSMPTANIMAI")) = False Then .CHUTAN = rs("MSMPTANIMAI") Else .CHUTAN = "0"    '中間抜試単位(枚)
            If IsNull(rs("MSMPCONSTMAI")) = False Then .CHUKYO = rs("MSMPCONSTMAI") Else .CHUKYO = "0"  '中間抜試許容値
        End With
        rs.Close
        '<<<<< 中間抜試規格セット追加 2011/07/15 Marushita
    Next

    DBDRV_scmzc_fcmlc001d_DispSiyou = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    Resume proc_exit
End Function

'*********************************************************************************************
'*    関数名        : DBDRV_scmzc_fcmlc001d_DispSmp
'*
'*    処理概要      : 1.再抜試指示 表示用ＤＢドライバ（WFサンプル管理）
'*
'*    パラメータ    : 変数名       ,IO ,型                                ,説明
'*                    inSXLID      ,I  ,String                            ,入力用SXLID
'*                    WfSmp        ,O  ,type_DBDRV_scmzc_fcmlc001d_WfSmp  ,WFサンプル管理用
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************
Public Function DBDRV_scmzc_fcmlc001d_DispSmp(inSXLID As String, _
                                            WFSMP() As type_DBDRV_scmzc_fcmlc001d_WfSmp, _
                                            sErrMsg As String _
                                            ) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim lngRecCnt   As Long
    Dim i           As Long
    Dim sDBName     As String

    ' WFサンプル管理取得
    ' ビューVECME011(SXL管理を検索し、そのブロックに対するサンプルを表示する)を使用

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001d_DispSmp"

    sDBName = "(XSDCW)"

    sSQL = "select "
    sSQL = sSQL & "XTALCW, "           ' 結晶番号
    sSQL = sSQL & "INPOSCW, "         ' 結晶内位置
    sSQL = sSQL & "SMPKBNCW, "           ' サンプル区分
    sSQL = sSQL & "REPSMPLIDCW, "           ' サンプルID
    sSQL = sSQL & "HINBCW, "           ' 品番
    sSQL = sSQL & "REVNUMCW, "           ' 製品番号改訂番号
    sSQL = sSQL & "FACTORYCW, "          ' 工場
    sSQL = sSQL & "OPECW, "          ' 操業条件
    sSQL = sSQL & "KTKBNCW, "            ' 確定区分
    sSQL = sSQL & "WFINDRSCW, "          ' 状態FLG（Rs)
    sSQL = sSQL & "WFINDOICW, "          ' 状態FLG（Oi)
    sSQL = sSQL & "WFINDB1CW, "          ' 状態FLG（B1)
    sSQL = sSQL & "WFINDB2CW, "          ' 状態FLG（B2）
    sSQL = sSQL & "WFINDB3CW, "          ' 状態FLG（B3)
    sSQL = sSQL & "WFINDL1CW, "          ' 状態FLG（L1)
    sSQL = sSQL & "WFINDL2CW, "          ' 状態FLG（L2)
    sSQL = sSQL & "WFINDL3CW, "          ' 状態FLG（L3)
    sSQL = sSQL & "WFINDL4CW, "          ' 状態FLG（L4)
    sSQL = sSQL & "WFINDDSCW, "          ' 状態FLG（DS)
    sSQL = sSQL & "WFINDDZCW, "          ' 状態FLG（DZ)
    sSQL = sSQL & "WFINDSPCW, "          ' 状態FLG（SP)
    sSQL = sSQL & "WFINDDO1CW, "         ' 状態FLG（DO1)
    sSQL = sSQL & "WFINDDO2CW, "         ' 状態FLG（DO2)
    sSQL = sSQL & "WFINDDO3CW, "         ' 状態FLG（DO3)
    sSQL = sSQL & "WFINDOT1CW, "         ' 状態FLG（OT1)
    sSQL = sSQL & "WFINDOT2CW, "         ' 状態FLG（OT2)
    sSQL = sSQL & "WFINDAOICW, "         ' 状態FLG (AOi)      '残存酸素追加　03/12/15 ooba
    sSQL = sSQL & "WFINDGDCW, "          ' 状態FLG (GD)       'GD追加　05/02/17 ooba
    sSQL = sSQL & "WFHSGDCW "            ' 保証FLG (GD)       'GD追加　05/02/17 ooba
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    sSQL = sSQL & ",EPINDB1CW, "          ' 状態FLG (OSF1E)
    sSQL = sSQL & "EPINDB2CW, "           ' 状態FLG (OSF2E)
    sSQL = sSQL & "EPINDB3CW, "           ' 状態FLG (OSF3E)
    sSQL = sSQL & "EPINDL1CW, "           ' 状態FLG (BMD1E)
    sSQL = sSQL & "EPINDL2CW, "           ' 状態FLG (BMD2E)
    sSQL = sSQL & "EPINDL3CW "            ' 状態FLG (BMD3E)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

    sSQL = sSQL & " from XSDCW"
    sSQL = sSQL & " where SXLIDCW='" & inSXLID & "' " & _
                " and LIVKCW='0' " & _
                " order by INPOSCW "


    Debug.Print sSQL
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount = 0 Then
         DBDRV_scmzc_fcmlc001d_DispSmp = FUNCTION_RETURN_FAILURE
         ReDim WFSMP(0)
         sErrMsg = GetMsgStr("EGET") & sDBName
         rs.Close
         GoTo proc_exit
    End If

    lngRecCnt = rs.RecordCount
    ReDim WFSMP(lngRecCnt)

    For i = 1 To lngRecCnt
        DoEvents
        With WFSMP(i)
            .INGOTPOS = rs("INPOSCW")        ' 結晶内位置
            .SMPLID = rs("REPSMPLIDCW")            ' サンプルID
            .hinban = rs("HINBCW")            ' 品番
            .REVNUM = rs("REVNUMCW")            ' 製品番号改訂番号
            .factory = rs("FACTORYCW")          ' 工場
            .opecond = rs("OPECW")          ' 操業条件
            .WFINDRS = rs("WFINDRSCW")          ' 状態FLG（Rs)
            .WFINDOI = rs("WFINDOICW")          ' 状態FLG（Oi)
            .WFINDB1 = rs("WFINDB1CW")          ' 状態FLG（B1)
            .WFINDB2 = rs("WFINDB2CW")          ' 状態FLG（B2）
            .WFINDB3 = rs("WFINDB3CW")          ' 状態FLG（B3)
            .WFINDL1 = rs("WFINDL1CW")          ' 状態FLG（L1)
            .WFINDL2 = rs("WFINDL2CW")          ' 状態FLG（L2)
            .WFINDL3 = rs("WFINDL3CW")          ' 状態FLG（L3)
            .WFINDL4 = rs("WFINDL4CW")          ' 状態FLG（L4)
            .WFINDDS = rs("WFINDDSCW")          ' 状態FLG（DS)
            .WFINDDZ = rs("WFINDDZCW")          ' 状態FLG（DZ)
            .WFINDSP = rs("WFINDSPCW")          ' 状態FLG（SP)
            .WFINDDO1 = rs("WFINDDO1CW")        ' 状態FLG（DO1)
            .WFINDDO2 = rs("WFINDDO2CW")        ' 状態FLG（DO2)
            .WFINDDO3 = rs("WFINDDO3CW")        ' 状態FLG（DO3)
            .WFINDOTHER1 = rs("WFINDOT1CW")        ' 状態FLG（DO2)
            .WFINDOTHER2 = rs("WFINDOT2CW")        ' 状態FLG（DO3)
            ''残存酸素追加
            If IsNull(rs("WFINDAOICW")) = False Then .WFINDAOI = rs("WFINDAOICW")  ' 状態FLG (AOi)
            .WFINDGD = rs("WFINDGDCW")          ' 状態FLG（GD)      'GD追加　05/02/17 ooba
            .WFHSGD = rs("WFHSGDCW")            ' 保証FLG（GD)      'GD追加　05/02/17 ooba
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
            .EPINDB1CW = rs("EPINDB1CW")
            .EPINDB2CW = rs("EPINDB2CW")
            .EPINDB3CW = rs("EPINDB3CW")
            .EPINDL1CW = rs("EPINDL1CW")
            .EPINDL2CW = rs("EPINDL2CW")
            .EPINDL3CW = rs("EPINDL3CW")
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_scmzc_fcmlc001d_DispSmp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    Resume proc_exit
End Function

'*********************************************************************************************************
'*    関数名        : DBDRV_scmzc_fcmlc001d_Exec
'*
'*    処理概要      : 1.再抜試指示 更新、挿入用ＤＢドライバ
'*
'*    パラメータ    : 変数名       ,IO   ,型               ,説明
'*                    WfSampleGr   ,I    ,typ_WfSampleGr   ,WFサンプル管理用
'*                    SXL          ,O    ,typ_TBCME042     ,SXL管理用
'*                                                          （結晶番号がNullだったら更新、それ以外は挿入）
'*                    WfHantei     ,O    ,typ_TBCMW005     ,WF総合判定実績用
'*                    HuriHai      ,O    ,typ_TBCMW006     ,振替廃棄実績用
'*                    SokuSizi     ,O    ,typ_TBCMY003     ,測定評価方法指示挿入用
'*                    SXLKakuSiji  ,O    ,typ_TBCMY007     ,SXL確定指示
'*      　            pEpMesInd 　 ,I    ,typ_TBCMY020   　,EP測定評価指示
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************************
Public Function DBDRV_scmzc_fcmlc001d_Exec(WfSampleGr() As typ_WfSampleGr, _
                                           sxl() As typ_TBCME042, _
                                           WfHantei As typ_TBCMW005, _
                                           HuriHai() As typ_TBCMW006, _
                                           SokuSizi() As typ_TBCMY003, _
                                           SXLKakuSiji() As typ_TBCMY007, _
                                           pEpMesInd() As typ_TBCMY020, _
                                           sErrMsg As String) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim lngRecCnt   As Long
    Dim i           As Long
    Dim sDBName     As String
    Dim intFromPos  As Integer
    Dim intToPos    As Integer
    Dim vGetFromPos As Variant
    Dim vGetToPos   As Variant
    Dim sTmpSxl()   As String     '仕掛工程再ﾁｪｯｸ用SXLID　06/03/14 ooba

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001d_Exec"

    '' WriteDBLog "Start"

    'WFサンプル管理への挿入(モジュール s_cmzcDBdriverCOM_SQL.bas 使用)
    '↑新サンプル管理への挿入に変更　2003/09/29 iida
    sDBName = "(XSDCW)" '新サンプル管理


    For i = 1 To UBound(WfSampleGr)
        If Trim(WfSampleGr(i).BLOCKID) = "" Then
            If WfSampleGr(i).WFSMP.REPSMPLIDCW <> "" Then  'SXLの先頭が共有の場合、サンプルIDは設定されていない 2003/04/22 okazaki
                With WfSampleGr(i).WFSMP
                    sSQL = "update XSDCW set "
                     sSQL = sSQL & "SXLIDCW = '" & .SXLIDCW & "',"          'SXLID
                    sSQL = sSQL & "HINBCW = '" & .HINBCW & "', "            ' 品番
                    sSQL = sSQL & "REVNUMCW = " & .REVNUMCW & ", "              ' 製品番号改訂番号
                    sSQL = sSQL & "FACTORYCW = '" & .FACTORYCW & "', "          ' 工場
                    sSQL = sSQL & "OPECW = '" & .OPECW & "', "          ' 操業条件

                    ''全振替時の結晶GD引継ぎ対応　05/08/04 ooba START =====================>
                    If (i = 1 And bMotoGDcpyFlg(1)) Or _
                        (i = UBound(WfSampleGr) And bMotoGDcpyFlg(2)) Then

                        sSQL = sSQL & "WFSMPLIDGDCW= '" & .WFSMPLIDGDCW & "', "   'サンプルID(GD)
                        sSQL = sSQL & "WFINDGDCW= '" & .WFINDGDCW & "', "         '状態FLG(GD)
                        sSQL = sSQL & "WFRESGDCW= '" & .WFRESGDCW & "', "         '実績FLG(GD)
                        sSQL = sSQL & "WFHSGDCW= '" & .WFHSGDCW & "', "           '保証FLG(GD)
                    End If
                    ''全振替時の結晶GD引継ぎ対応　05/08/04 ooba END =======================>

                    sSQL = sSQL & "NUKISIFLGCW = '1', "                   ' 抜試指示通過フラグ 09/05/26 ooba
                    sSQL = sSQL & "KDAYCW = sysdate , "                   ' 更新日付
                    sSQL = sSQL & "SNDKCW = '0', "                       ' 送信フラグ
                    sSQL = sSQL & "SNDDAYCW = sysdate ,  "                   ' 送信日付
                    sSQL = sSQL & "KSTAFFCW = '" & .KSTAFFCW & "' "        '更新社員ID

                    sSQL = sSQL & "where XTALCW ='" & .XTALCW & "' and "   ' 結晶番号
                    sSQL = sSQL & "SXLIDCW = '" & tblSXL.SXLID & "' and "      ' SXLID 2003/11/05 条件を追加(最初と最後は元のSXLIDを更新)
                    sSQL = sSQL & "INPOSCW = " & .INPOSCW & " and "      ' 結晶内位置
                    sSQL = sSQL & "SMPKBNCW = '" & .SMPKBNCW & "'"            ' サンプル区分
                End With

                '' WriteDBLog sSql, sDBName
                If 1 <> OraDB.ExecuteSQL(sSQL) Then
                    DBDRV_scmzc_fcmlc001d_Exec = FUNCTION_RETURN_FAILURE
                    sErrMsg = GetMsgStr("EAPLY") & sDBName
                    GoTo proc_exit
                End If
            End If
        Else
            sSQL = "insert into XSDCW ( "
            sSQL = sSQL & "SXLIDCW, "             ' SXLID
            sSQL = sSQL & "SMPKBNCW, "            ' サンプル区分
            sSQL = sSQL & "TBKBNCW, "             ' T/B区分
            sSQL = sSQL & "REPSMPLIDCW, "         ' サンプルID
            sSQL = sSQL & "XTALCW, "              ' 結晶番号
            sSQL = sSQL & "INPOSCW, "             ' 結晶内位置
            sSQL = sSQL & "HINBCW, "              ' 品番
            sSQL = sSQL & "REVNUMCW, "            ' 製品番号改訂番号
            sSQL = sSQL & "FACTORYCW, "           ' 工場
            sSQL = sSQL & "OPECW, "               ' 操業条件
            sSQL = sSQL & "KTKBNCW, "             ' 確定区分
            sSQL = sSQL & "SMCRYNUMCW, "          ' サンプルブロックID
            sSQL = sSQL & "WFSMPLIDRSCW, "        ' サンプルID()
            sSQL = sSQL & "WFSMPLIDRS1CW, "       ' 推定サンプルID1
            sSQL = sSQL & "WFSMPLIDRS2CW, "       ' 推定サンプルID2
            sSQL = sSQL & "WFINDRSCW, "           ' 状態FLG（Rs)
            sSQL = sSQL & "WFRESRS1CW, "           ' 実績FLG1（Rs)
            sSQL = sSQL & "WFRESRS2CW, "           ' 実績FLG2（Rs)
            sSQL = sSQL & "WFSMPLIDOICW, "        ' サンプルID(Oi)
            sSQL = sSQL & "WFINDOICW, "           ' 状態FLG（Oi)
            sSQL = sSQL & "WFRESOICW, "           ' 実績FLG（Oi)
            sSQL = sSQL & "WFSMPLIDB1CW, "        ' サンプルID(B1)
            sSQL = sSQL & "WFINDB1CW, "           ' 状態FLG（B1)
            sSQL = sSQL & "WFRESB1CW, "           ' 実績FLG（B1)
            sSQL = sSQL & "WFSMPLIDB2CW, "        ' サンプルID(B2)
            sSQL = sSQL & "WFINDB2CW, "           ' 状態FLG（B2）
            sSQL = sSQL & "WFRESB2CW, "           ' 実績FLG（B2）
            sSQL = sSQL & "WFSMPLIDB3CW, "        ' サンプルID(B3)
            sSQL = sSQL & "WFINDB3CW, "           ' 状態FLG（B3)
            sSQL = sSQL & "WFRESB3CW, "           ' 実績FLG（B3)
            sSQL = sSQL & "WFSMPLIDL1CW, "        ' サンプルID(L1)
            sSQL = sSQL & "WFINDL1CW, "           ' 状態FLG（L1)
            sSQL = sSQL & "WFRESL1CW, "           ' 実績FLG（L1)
            sSQL = sSQL & "WFSMPLIDL2CW, "        ' サンプルID(L2)
            sSQL = sSQL & "WFINDL2CW, "           ' 状態FLG（L2)
            sSQL = sSQL & "WFRESL2CW, "           ' 実績FLG（L2)
            sSQL = sSQL & "WFSMPLIDL3CW, "        ' サンプルID(L3)
            sSQL = sSQL & "WFINDL3CW, "           ' 状態FLG（L3)
            sSQL = sSQL & "WFRESL3CW, "           ' 実績FLG（L3)
            sSQL = sSQL & "WFSMPLIDL4CW, "        ' サンプルID(L4)
            sSQL = sSQL & "WFINDL4CW, "           ' 状態FLG（L4)
            sSQL = sSQL & "WFRESL4CW, "           ' 実績FLG（L4)
            sSQL = sSQL & "WFSMPLIDDSCW, "        ' サンプルID(DS)
            sSQL = sSQL & "WFINDDSCW, "           ' 状態FLG（DS)
            sSQL = sSQL & "WFRESDSCW, "           ' 実績FLG（DS)
            sSQL = sSQL & "WFSMPLIDDZCW, "        ' サンプルID(DZ)
            sSQL = sSQL & "WFINDDZCW, "           ' 状態FLG（DZ)
            sSQL = sSQL & "WFRESDZCW, "           ' 実績FLG（DZ)
            sSQL = sSQL & "WFSMPLIDSPCW, "        ' サンプルID(SP)
            sSQL = sSQL & "WFINDSPCW, "           ' 状態FLG（SP)
            sSQL = sSQL & "WFRESSPCW, "           ' 実績FLG（SP)
            sSQL = sSQL & "WFSMPLIDDO1CW, "       ' サンプルID(DO1)
            sSQL = sSQL & "WFINDDO1CW, "          ' 状態FLG（DO1)
            sSQL = sSQL & "WFRESDO1CW, "          ' 実績FLG（DO1)
            sSQL = sSQL & "WFSMPLIDDO2CW, "       ' サンプルID(DO2)
            sSQL = sSQL & "WFINDDO2CW, "          ' 状態FLG（DO2)
            sSQL = sSQL & "WFRESDO2CW, "          ' 実績FLG（DO2)
            sSQL = sSQL & "WFSMPLIDDO3CW, "       ' サンプルID(DO3)
            sSQL = sSQL & "WFINDDO3CW, "          ' 状態FLG（DO3)
            sSQL = sSQL & "WFRESDO3CW, "          ' 実績FLG（DO3)
             'add start 2003/05/26 hitec)後藤 -------------------------
            sSQL = sSQL & "WFSMPLIDOT1CW, "       ' サンプルID(OT1)
            sSQL = sSQL & "WFINDOT1CW, "            ' 状態FLG（OT1)
            sSQL = sSQL & "WFRESOT1CW, "          ' 実績FLG（OT1)
            sSQL = sSQL & "WFSMPLIDOT2CW, "       ' サンプルID(OT2)
            sSQL = sSQL & "WFINDOT2CW, "            ' 状態FLG（OT2)
            sSQL = sSQL & "WFRESOT2CW, "          ' 実績FLG（OT2)
            'add end   2003/05/26 hitec)後藤 -------------------------
            sSQL = sSQL & "WFSMPLIDAOICW, "       ' サンプルID(AOi)
            sSQL = sSQL & "WFINDAOICW, "          ' 状態FLG(AOi)
            sSQL = sSQL & "WFRESAOICW, "          ' 実績FLG(AOi)
            '' GD追加　05/02/21 ooba START =====================================>
            sSQL = sSQL & "WFSMPLIDGDCW, "        ' サンプルID (GD)
            sSQL = sSQL & "WFINDGDCW, "           ' 状態FLG (GD)
            sSQL = sSQL & "WFRESGDCW, "           ' 実績FLG (GD)
            sSQL = sSQL & "WFHSGDCW, "            ' 保証FLG (GD)
            '' GD追加　05/02/21 ooba END =======================================>
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
            sSQL = sSQL & "EPSMPLIDL1CW, "    ' サンプルID (OSF1E)
            sSQL = sSQL & "EPINDL1CW, "       ' 状態FLG (OSF1E)
            sSQL = sSQL & "EPRESL1CW, "       ' 実績FLG (OSF1E)
            sSQL = sSQL & "EPSMPLIDL2CW, "    ' サンプルID (OSF2E)
            sSQL = sSQL & "EPINDL2CW, "       ' 状態FLG (OSF2E)
            sSQL = sSQL & "EPRESL2CW, "       ' 実績FLG (OSF2E)
            sSQL = sSQL & "EPSMPLIDL3CW, "    ' サンプルID (OSF3E)
            sSQL = sSQL & "EPINDL3CW, "       ' 状態FLG (OSF3E)
            sSQL = sSQL & "EPRESL3CW, "       ' 実績FLG (OSF3E)
            sSQL = sSQL & "EPSMPLIDB1CW, "    ' サンプルID (BMD1E)
            sSQL = sSQL & "EPINDB1CW, "       ' 状態FLG (BMD1E)
            sSQL = sSQL & "EPRESB1CW, "       ' 実績FLG (BMD1E)
            sSQL = sSQL & "EPSMPLIDB2CW, "    ' サンプルID (BMD2E)
            sSQL = sSQL & "EPINDB2CW, "       ' 状態FLG (BMD2E)
            sSQL = sSQL & "EPRESB2CW, "       ' 実績FLG (BMD2E)
            sSQL = sSQL & "EPSMPLIDB3CW, "    ' サンプルID (BMD3E)
            sSQL = sSQL & "EPINDB3CW, "       ' 状態FLG (BMD3E)
            sSQL = sSQL & "EPRESB3CW, "       ' 実績FLG (BMD3E)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
            sSQL = sSQL & "SMPLNUMCW, "           ' サンプル枚数
            sSQL = sSQL & "SMPLPATCW, "           ' サンプルパターン
            sSQL = sSQL & "LIVKCW, "              ' 生死区分
            sSQL = sSQL & "NUKISIFLGCW, "         ' 抜試指示通過フラグ 09/05/26 ooba
            sSQL = sSQL & "TSTAFFCW, "            ' 登録社員ID
            sSQL = sSQL & "TDAYCW, "              ' 登録日付
            sSQL = sSQL & "KSTAFFCW, "            ' 更新社員ID
            sSQL = sSQL & "KDAYCW, "              ' 更新日付
            sSQL = sSQL & "SNDKCW, "              ' 送信フラグ
            sSQL = sSQL & "SNDDAYCW ) "           ' 送信日付

            With WfSampleGr(i).WFSMP
                sSQL = sSQL & " values ('"
                sSQL = sSQL & .SXLIDCW & "', '"       ' SXLID
                sSQL = sSQL & .SMPKBNCW & "', '"      ' サンプル区分
                sSQL = sSQL & .TBKBNCW & "', '"       ' T/B区分
                sSQL = sSQL & .REPSMPLIDCW & "', '"   ' サンプルID
                sSQL = sSQL & .XTALCW & "', "         ' 結晶番号
                sSQL = sSQL & .INPOSCW & ", '"        ' 結晶内位置
                sSQL = sSQL & .HINBCW & "', "         ' 品番
                sSQL = sSQL & .REVNUMCW & ", '"       ' 製品番号改訂番号
                sSQL = sSQL & .FACTORYCW & "', '"     ' 工場
                sSQL = sSQL & .OPECW & "', '"         ' 操業条件
                sSQL = sSQL & .KTKBNCW & "', '"       ' 確定区分
                sSQL = sSQL & .SMCRYNUMCW & "', '"    ' サンプルブロックID
                sSQL = sSQL & .WFSMPLIDRSCW & "', "  ' サンプルID（Rs）
                sSQL = sSQL & "Null, "               ' 推定サンプルID1（Rs）
                sSQL = sSQL & "Null, '" ' 推定サンプルID2（Rs）
                sSQL = sSQL & .WFINDRSCW & "', '"     ' 状態FLG（Rs)
                sSQL = sSQL & .WFRESRS1CW & "' , "    ' 実績FLG1（Rs)
                sSQL = sSQL & "Null, '"      ' 実績FLG2（Rs)
                sSQL = sSQL & .WFSMPLIDOICW & "', '"  ' サンプルID（Oi）
                sSQL = sSQL & .WFINDOICW & "', '"     ' 状態FLG（Oi)
                sSQL = sSQL & .WFRESOICW & "', '"     ' 実績FLG（Oi)
                sSQL = sSQL & .WFSMPLIDB1CW & "', '"  ' サンプルID（B1）
                sSQL = sSQL & .WFINDB1CW & "', '"     ' 状態FLG（B1)
                sSQL = sSQL & .WFRESB1CW & "', '"     ' 実績FLG（B1)
                sSQL = sSQL & .WFSMPLIDB2CW & "', '"  ' サンプルID（B2）
                sSQL = sSQL & .WFINDB2CW & "', '"     ' 状態FLG（B2)
                sSQL = sSQL & .WFRESB2CW & "', '"     ' 実績FLG（B2)
                sSQL = sSQL & .WFSMPLIDB3CW & "', '"  ' サンプルID（B3）
                sSQL = sSQL & .WFINDB3CW & "', '"     ' 状態FLG（B3)
                sSQL = sSQL & .WFRESB3CW & "', '"     ' 実績FLG（B3)
                sSQL = sSQL & .WFSMPLIDL1CW & "', '"  ' サンプルID（L1）
                sSQL = sSQL & .WFINDL1CW & "', '"     ' 状態FLG（L1)
                sSQL = sSQL & .WFRESL1CW & "', '"     ' 実績FLG（L1)
                sSQL = sSQL & .WFSMPLIDL2CW & "', '"  ' サンプルID（L2）
                sSQL = sSQL & .WFINDL2CW & "', '"     ' 状態FLG（L2)
                sSQL = sSQL & .WFRESL2CW & "', '"     ' 実績FLG（L2)
                sSQL = sSQL & .WFSMPLIDL3CW & "', '"  ' サンプルID（L3）
                sSQL = sSQL & .WFINDL3CW & "', '"     ' 状態FLG（L3)
                sSQL = sSQL & .WFRESL3CW & "', '"     ' 実績FLG（L3)
                sSQL = sSQL & .WFSMPLIDL4CW & "', '"  ' サンプルID（L4）
                sSQL = sSQL & .WFINDL4CW & "', '"     ' 状態FLG（L4)
                sSQL = sSQL & .WFRESL4CW & "', '"     ' 実績FLG（L4)
                sSQL = sSQL & .WFSMPLIDDSCW & "', '"  ' サンプルID（DS）
                sSQL = sSQL & .WFINDDSCW & "', '"     ' 状態FLG（DS)
                sSQL = sSQL & .WFRESDSCW & "', '"     ' 実績FLG（DS)
                sSQL = sSQL & .WFSMPLIDDZCW & "', '"  ' サンプルID（DZ）
                sSQL = sSQL & .WFINDDZCW & "', '"     ' 状態FLG（DZ)
                sSQL = sSQL & .WFRESDZCW & "', '"     ' 実績FLG（DZ)
                sSQL = sSQL & .WFSMPLIDSPCW & "', '"  ' サンプルID（SP）
                sSQL = sSQL & .WFINDSPCW & "', '"     ' 状態FLG（SP)
                sSQL = sSQL & .WFRESSPCW & "', '"     ' 実績FLG（SP)
                sSQL = sSQL & .WFSMPLIDDO1CW & "', '" ' サンプルID（DO1）
                sSQL = sSQL & .WFINDDO1CW & "', '"    ' 状態FLG（DO1)
                sSQL = sSQL & .WFRESDO1CW & "', '"    ' 実績FLG（DO1)
                sSQL = sSQL & .WFSMPLIDDO2CW & "', '" ' サンプルID（DO2）
                sSQL = sSQL & .WFINDDO2CW & "', '"    ' 状態FLG（DO2)
                sSQL = sSQL & .WFRESDO2CW & "', '"    ' 実績FLG（DO2)
                sSQL = sSQL & .WFSMPLIDDO3CW & "', '" ' サンプルID（DO3）
                sSQL = sSQL & .WFINDDO3CW & "', '"    ' 状態FLG（DO3)
                sSQL = sSQL & .WFRESDO3CW & "', '"    ' 実績FLG（DO3)
'                sSQL = sSQL & .WFSMPLIDOT1CW & "', '" ' サンプルID（OT1）
                sSQL = sSQL & "                ', '"   ' サンプルID（OT1）2010/04/08 Y.Hitomi OT1暫定対応
                sSQL = sSQL & .WFINDOT1CW & "', '"    ' 状態FLG（OT1)
                sSQL = sSQL & .WFRESOT1CW & "', '"    ' 実績FLG（OT1)
                sSQL = sSQL & .WFSMPLIDOT2CW & "', '" ' サンプルID（OT2）
                sSQL = sSQL & .WFINDOT2CW & "', '"    ' 状態FLG（OT2)
                sSQL = sSQL & .WFRESOT2CW & "', '"    ' 実績FLG（OT2)
                ''残存酸素追加　03/12/15 ooba START ================================>
                sSQL = sSQL & .WFSMPLIDAOICW & "', '" ' サンプルID（AOi）
                sSQL = sSQL & .WFINDAOICW & "', '"    ' 状態FLG（AOi）
                sSQL = sSQL & .WFRESAOICW & "', '"    ' 実績FLG（AOi）
                ''残存酸素追加　03/12/15 ooba END ==================================>
                '' GD追加　05/02/21 ooba START =====================================>
                sSQL = sSQL & .WFSMPLIDGDCW & "', '"  ' サンプルID (GD)
                sSQL = sSQL & .WFINDGDCW & "', '"     ' 状態FLG (GD)
                sSQL = sSQL & .WFRESGDCW & "', '"     ' 実績FLG (GD)
                sSQL = sSQL & .WFHSGDCW & "', "       ' 保証FLG (GD)
                '' GD追加　05/02/21 ooba END =======================================>
    '--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
                sSQL = sSQL & "'" & .EPSMPLIDL1CW & "', '"  ' サンプルID (OSF1E)
                sSQL = sSQL & .EPINDL1CW & "', '"         ' 状態FLG (OSF1E)
                sSQL = sSQL & .EPRESL1CW & "', '"         ' 実績FLG (OSF1E)
                sSQL = sSQL & .EPSMPLIDL2CW & "', '"      ' サンプルID (OSF2E)
                sSQL = sSQL & .EPINDL2CW & "', '"         ' 状態FLG (OSF2E)
                sSQL = sSQL & .EPRESL2CW & "', '"         ' 実績FLG (OSF2E)
                sSQL = sSQL & .EPSMPLIDL3CW & "', '"      ' サンプルID (OSF3E)
                sSQL = sSQL & .EPINDL3CW & "', '"         ' 状態FLG (OSF3E)
                sSQL = sSQL & .EPRESL3CW & "', '"         ' 実績FLG (OSF3E)
                sSQL = sSQL & .EPSMPLIDB1CW & "', '"      ' サンプルID (BMD1E)
                sSQL = sSQL & .EPINDB1CW & "', '"         ' 状態FLG (BMD1E)
                sSQL = sSQL & .EPRESB1CW & "', '"         ' 実績FLG (BMD1E)
                sSQL = sSQL & .EPSMPLIDB2CW & "', '"      ' サンプルID (BMD2E)
                sSQL = sSQL & .EPINDB2CW & "', '"         ' 状態FLG (BMD2E)
                sSQL = sSQL & .EPRESB2CW & "', '"         ' 実績FLG (BMD2E)
                sSQL = sSQL & .EPSMPLIDB3CW & "', '"      ' サンプルID (BMD3E)
                sSQL = sSQL & .EPINDB3CW & "', '"         ' 状態FLG (BMD3E)
                sSQL = sSQL & .EPRESB3CW & "', "          ' 実績FLG (BMD3E)
                sSQL = sSQL & "NULL, "                ' サンプル枚数
                sSQL = sSQL & "NULL, '"               ' サンプルパターン
                sSQL = sSQL & .LIVKCW & "', '"        ' 生死区分
                sSQL = sSQL & "1', '"                 ' 抜試指示通過フラグ 09/05/26 ooba
                sSQL = sSQL & .TSTAFFCW & "', "       ' 登録社員ID
                sSQL = sSQL & "sysdate, '"            ' 登録日付
                sSQL = sSQL & .KSTAFFCW & "', "       ' 更新社員ID
                sSQL = sSQL & "sysdate, "             ' 更新日付
                sSQL = sSQL & "'0', "                 ' 送信フラグ
                sSQL = sSQL & "sysdate)"              ' 送信日付
            End With

            '' WriteDBLog sSql, sDBName
            If 1 <> OraDB.ExecuteSQL(sSQL) Then
                DBDRV_scmzc_fcmlc001d_Exec = FUNCTION_RETURN_FAILURE
                sErrMsg = GetMsgStr("EAPLY") & sDBName
                GoTo proc_exit
            End If
        End If
    Next

    'Add Start 2011/04/25 SMPK Miyata
    sDBName = "(XSDCW_1)" '新サンプル管理(中間抜試)

    For i = 1 To UBound(sxl)
        sSQL = "update XSDCW_1 set "
        sSQL = sSQL & "SXLIDCW = '" & sxl(i).SXLID & "',"           ' SXLID
        sSQL = sSQL & "HINBCW = '" & sxl(i).hinban & "', "          ' 品番
        sSQL = sSQL & "REVNUMCW = " & sxl(i).REVNUM & ", "          ' 製品番号改訂番号
        sSQL = sSQL & "FACTORYCW = '" & sxl(i).factory & "', "      ' 工場
        sSQL = sSQL & "OPECW = '" & sxl(i).opecond & "', "          ' 操業条件
        sSQL = sSQL & "NUKISIFLGCW = '1', "                         ' 抜試指示通過フラグ 09/05/26 ooba
        sSQL = sSQL & "KDAYCW = sysdate , "                         ' 更新日付
        sSQL = sSQL & "SNDKCW = '0', "                              ' 送信フラグ
        sSQL = sSQL & "SNDDAYCW = sysdate ,  "                      ' 送信日付
        sSQL = sSQL & "KSTAFFCW = '" & WfSampleGr(1).WFSMP.KSTAFFCW & "' "  ' 更新社員ID

        sSQL = sSQL & "where XTALCW ='" & tblSXL.CRYNUM & "' and "  ' 結晶番号
        sSQL = sSQL & "SXLIDCW = '" & tblSXL.SXLID & "' and "       ' SXLID(元のSXLID)
        sSQL = sSQL & "INPOSCW > " & sxl(i).INGOTPOS & " and "      ' 結晶内位置
        sSQL = sSQL & "INPOSCW <= " & sxl(i).INGOTPOS + sxl(i).LENGTH   ' 結晶内位置

        Call OraDB.ExecuteSQL(sSQL)

    Next i

    'Add End   2011/04/25 SMPK Miyata


    '仕掛工程再チェック機能追加　06/03/14 ooba
    ReDim sTmpSxl(1)
    sTmpSxl(1) = Trim(f_cmbc039_3.txtKSXLID.text)
    If DBDRV_CheckCodeXSDCB(sTmpSxl, PROCD_WFC_SOUGOUHANTEI, sErrMsg) = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmlc001d_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ' SXL管理への更新or挿入
    sDBName = "(XSDCB)"
    For i = 1 To UBound(sxl)
        'CRYNUMがNullの時はSXLIDで更新
        If Trim(sxl(i).CRYNUM) = "" Then    '2003/04/20 okazaki

            'サンプルIDの変更が無い場合はSXLIDも、開始位置変えない
            With sxl(i)
                sSQL = "update XSDCB set "
                sSQL = sSQL & "rlencb=" & .LENGTH & ", "          ' 長さ
                sSQL = sSQL & "gnwkntcb='" & .NOWPROC & "', "     ' 現在工程
                sSQL = sSQL & "newkntcb='" & .LASTPASS & "', "    ' 最終通過工程
                sSQL = sSQL & "livkcb='" & .DELCLS & "', "        ' 削除区分
                sSQL = sSQL & "lstccb='" & .LSTATCLS & "', "      ' 最終状態区分
                sSQL = sSQL & "sholdclscb='" & .HOLDCLS & "', "   ' ホールド区分
                sSQL = sSQL & "hinbcb='" & .hinban & "', "        ' 品番
                sSQL = sSQL & "revnumcb=" & .REVNUM & ", "        ' 製品番号改訂番号
                sSQL = sSQL & "factorycb='" & .factory & "', "    ' 工場
                sSQL = sSQL & "opecb='" & .opecond & "', "        ' 操業条件
                sSQL = sSQL & "furyccb='" & .BDCAUS & "', "       ' 不良理由
                sSQL = sSQL & "maicb=" & .COUNT & ", "            ' 枚数
                sSQL = sSQL & "kdaycb=sysdate, "                  ' 更新日付
                sSQL = sSQL & "sndkcb='0' "                       ' 送信FLG
                sSQL = sSQL & " where sxlidcb='" & .SXLID & "' "
            End With

            '' WriteDBLog sSql, sDBName
            If 0 >= OraDB.ExecuteSQL(sSQL) Then
                DBDRV_scmzc_fcmlc001d_Exec = FUNCTION_RETURN_FAILURE
                sErrMsg = GetMsgStr("EAPLY") & sDBName
                GoTo proc_exit
            End If
        Else
            sSQL = "insert into XSDCB ( "
            sSQL = sSQL & "xtalcb, "          ' 結晶番号
            sSQL = sSQL & "inposcb, "         ' 結晶内開始位置
            sSQL = sSQL & "rlencb, "          ' 長さ
            sSQL = sSQL & "sxlidcb, "         ' SXLID
            sSQL = sSQL & "gnwkntcb, "        ' 現在工程
            sSQL = sSQL & "newkntcb, "        ' 最終通過工程
            sSQL = sSQL & "livkcb, "          ' 削除区分
            sSQL = sSQL & "lstccb, "          ' 最終状態区分
            sSQL = sSQL & "sholdclscb, "      ' ホールド区分
            sSQL = sSQL & "hinbcb, "          ' 品番
            sSQL = sSQL & "revnumcb, "        ' 製品番号改訂番号
            sSQL = sSQL & "factorycb, "       ' 工場
            sSQL = sSQL & "opecb, "           ' 操業条件
            sSQL = sSQL & "furyccb, "         ' 不良理由
            sSQL = sSQL & "maicb, "           ' 枚数
            sSQL = sSQL & "tdaycb, "          ' 登録日付
            sSQL = sSQL & "kdaycb, "          ' 更新日付
            sSQL = sSQL & "sndkcb, "          ' 送信フラグ
            sSQL = sSQL & "WSRMAICB, "        ' WS洗後枚数
            sSQL = sSQL & "WSNMAICB, "        ' WS洗浄欠落枚数
            sSQL = sSQL & "WFCMAICB, "        ' 受入枚数
            sSQL = sSQL & "SXLRMAICB, "       ' SXL指示(良品)
            sSQL = sSQL & "WFCNMAICB, "       ' WFC内欠落枚数
            sSQL = sSQL & "SXLEMAICB, "       ' SXL確定枚数
            sSQL = sSQL & "SRMAICB, "         ' サンプル抜指示枚数
            sSQL = sSQL & "SNMAICB, "         ' サンプル抜指示不良枚数
            sSQL = sSQL & "STMAICB, "         ' サンプル枚数
            sSQL = sSQL & "FURIMAICB, "       ' 振替枚数
            sSQL = sSQL & "XTWORKCB, "        ' 製造工場
            sSQL = sSQL & "WFWORKCB, "        ' ウェーハ製造
            sSQL = sSQL & "LUFRCCB, "         ' 格上コード
            sSQL = sSQL & "LUFRBCB, "         ' 格上区分
            sSQL = sSQL & "LDERCCB, "         ' 格下コード
            sSQL = sSQL & "HOLDCCB, "         ' ホールドコード
            sSQL = sSQL & "HOLDBCB, "         ' ホールド区分
            sSQL = sSQL & "EXKUBCB, "         ' 例外区分
            sSQL = sSQL & "HENPKCB, "         ' 返品区分
            sSQL = sSQL & "KANKCB, "          ' 完了区分
            sSQL = sSQL & "NFCB, "            ' 入庫区分
            sSQL = sSQL & "SAKJCB, "          ' 削除区分
            sSQL = sSQL & "SUMITCB "          ' SUMIT送信フラグ
            sSQL = sSQL & " ) "

            With sxl(i)
                sSQL = sSQL & " values ( "
                sSQL = sSQL & " '" & .CRYNUM & "', "           ' 結晶番号
                sSQL = sSQL & " " & .INGOTPOS & ", "           ' 結晶内開始位置
                sSQL = sSQL & " " & .LENGTH & ", "             ' 長さ
                sSQL = sSQL & " '" & .SXLID & "', "            ' SXLID
                sSQL = sSQL & " '" & .NOWPROC & "', "          ' 現在工程
                sSQL = sSQL & " '" & .LASTPASS & "', "         ' 最終通過工程
                sSQL = sSQL & " '" & .DELCLS & "', "           ' 削除区分
                sSQL = sSQL & " '" & .LSTATCLS & "', "         ' 最終状態区分
                sSQL = sSQL & " '" & .HOLDCLS & "', "          ' ホールド区分
                sSQL = sSQL & " '" & .hinban & "', "           ' 品番
                sSQL = sSQL & " " & .REVNUM & ", "             ' 製品番号改訂番号
                sSQL = sSQL & " '" & .factory & "', "          ' 工場
                sSQL = sSQL & " '" & .opecond & "', "          ' 操業条件
                sSQL = sSQL & " '" & .BDCAUS & "', "           ' 不良理由
                sSQL = sSQL & " " & .COUNT & ", "              ' 枚数
                sSQL = sSQL & " sysdate, "                     ' 登録日付
                sSQL = sSQL & " sysdate, "                     ' 更新日付
                sSQL = sSQL & " '0', "                         ' 送信フラグ
                sSQL = sSQL & " '0', "                         ' WS洗後枚数
                sSQL = sSQL & " '0', "                         ' WS洗浄欠落枚数
                sSQL = sSQL & " '0', "                         ' 受入枚数
                sSQL = sSQL & " '0', "                         ' SXL指示(良品)
                sSQL = sSQL & " '0', "                         ' WFC内欠落枚数
                sSQL = sSQL & " '0', "                         ' SXL確定枚数
                sSQL = sSQL & " '0', "                         ' サンプル抜指示枚数
                sSQL = sSQL & " '0', "                         ' サンプル抜指示不良枚数
                sSQL = sSQL & " '0', "                         ' サンプル枚数
                sSQL = sSQL & " '0', "                         ' 振替枚数
                sSQL = sSQL & " '42', "                        ' 製造工場
                sSQL = sSQL & " '  ', "                        ' ウェーハ製造
                sSQL = sSQL & " '   ', "                       ' 格上コード
                sSQL = sSQL & " ' ', "                         ' 格上区分
                sSQL = sSQL & " '   ', "                       ' 格下コード
                sSQL = sSQL & " '   ', "                       ' ホールドコード
                sSQL = sSQL & " '0', "                         ' ホールド区分
                sSQL = sSQL & " ' ', "                         ' 例外区分
                sSQL = sSQL & " ' ', "                         ' 返品区分
                sSQL = sSQL & " '0', "                         ' 完了区分
                sSQL = sSQL & " '0', "                         ' 入庫区分
                sSQL = sSQL & " '0', "                         ' 削除区分
                sSQL = sSQL & " '0' "                          ' SUMIT送信フラグ
                sSQL = sSQL & " ) "
            End With

            '' WriteDBLog sSql, sDBName
            If 0 >= OraDB.ExecuteSQL(sSQL) Then
                DBDRV_scmzc_fcmlc001d_Exec = FUNCTION_RETURN_FAILURE
                sErrMsg = GetMsgStr("EAPLY") & sDBName
                GoTo proc_exit
            End If
        End If
    Next

    sDBName = "(W005)"

    ' WF総合判定実績への挿入
    sSQL = "insert into TBCMW005 ( "
    sSQL = sSQL & "CRYNUM, "           ' 結晶番号
    sSQL = sSQL & "INGOTPOS, "         ' インゴット位置
    sSQL = sSQL & "TRANCNT, "          ' 処理回数
    sSQL = sSQL & "CRYLEN, "           ' 長さ
    sSQL = sSQL & "KRPROCCD, "         ' 管理工程コード
    sSQL = sSQL & "PROCCODE, "         ' 工程コード
    sSQL = sSQL & "SXLID, "            ' SXLID
    sSQL = sSQL & "CODE, "             ' 区分コード
    sSQL = sSQL & "TSTAFFID, "         ' 登録社員ID
    sSQL = sSQL & "REGDATE, "          ' 登録日付
    sSQL = sSQL & "KSTAFFID, "         ' 更新社員ID
    sSQL = sSQL & "UPDDATE, "          ' 更新日付
    sSQL = sSQL & "SENDFLAG, "         ' 送信フラグ
    sSQL = sSQL & "SENDDATE, "        ' 送信日付
    sSQL = sSQL & "PLANTCAT) "          ' 向先

    With WfHantei
        sSQL = sSQL & " select "
        sSQL = sSQL & " '" & .CRYNUM & "', "           ' 結晶番号
        sSQL = sSQL & " " & .INGOTPOS & ", "           ' インゴット位置
        sSQL = sSQL & " nvl(max(TRANCNT),0)+1, "       ' 処理回数
        sSQL = sSQL & " " & .CRYLEN & ", "             ' 長さ
        sSQL = sSQL & " '" & .KRPROCCD & "', "         ' 管理工程コード
        sSQL = sSQL & " '" & .PROCCODE & "', "         ' 工程コード
        sSQL = sSQL & " '" & .SXLID & "', "            ' SXLID
        sSQL = sSQL & " '" & .CODE & "', "             ' 区分コード
        sSQL = sSQL & " '" & .TSTAFFID & "', "         ' 登録社員ID
        sSQL = sSQL & " sysdate, "                     ' 登録日付
        sSQL = sSQL & " '" & .KSTAFFID & "', "         ' 更新社員ID
        sSQL = sSQL & " sysdate, "                     ' 更新日付
        sSQL = sSQL & " '0', "                         ' 送信フラグ
        sSQL = sSQL & " sysdate "                      ' 送信日付
        sSQL = sSQL & " , '" & sCmbMukesaki & "'"      ' 向先 2007/09/04 SPK Tsutsumi Add
        sSQL = sSQL & " from TBCMW005 "
        sSQL = sSQL & " where CRYNUM='" & .CRYNUM & "' " & _
                    " and INGOTPOS=" & .INGOTPOS
    End With

    '' WriteDBLog sSql, sDBName
    If 0 >= OraDB.ExecuteSQL(sSQL) Then
        DBDRV_scmzc_fcmlc001d_Exec = FUNCTION_RETURN_FAILURE
        sErrMsg = GetMsgStr("EAPLY") & sDBName
        GoTo proc_exit
    End If

    sDBName = "(W006)"
    For i = 1 To UBound(HuriHai)
        ' 振替廃棄実績への挿入 (モジュール s_cmzcDBdriverCOM_SQL.bas 使用)
        If DBDRV_Furikae_Ins(HuriHai(i)) = FUNCTION_RETURN_FAILURE Then
            DBDRV_scmzc_fcmlc001d_Exec = FUNCTION_RETURN_FAILURE
            sErrMsg = GetMsgStr("EAPLY") & sDBName
            GoTo proc_exit
        End If
    Next

    sDBName = "(Y003)"
    ' 測定評価方法指示への挿入 (モジュール s_cmzcDBdriverCOM_SQL.bas 使用)
    If DBDRV_SokuSizi_Ins(SokuSizi()) = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmlc001d_Exec = FUNCTION_RETURN_FAILURE
        sErrMsg = GetMsgStr("EAPLY") & sDBName
        GoTo proc_exit
    End If

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    '' エピ測定評価指示情報の挿入(s_cmzcDBdriverCOM_SQL.bas が必要)
    sDBName = "(Y020)"
    If DBDRV_SokuSizi_EP_Ins(pEpMesInd()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EAPLY") & sDBName
        GoTo proc_exit
    End If
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

    sDBName = "(Y007)"
    '2001/08/04　追加
    For i = 1 To UBound(SXLKakuSiji)
        ' SXL確定指示への挿入
        sSQL = "insert into TBCMY007 ("
        sSQL = sSQL & "SXL_ID, "           ' SXL-ID
        sSQL = sSQL & "SAMPLE_FROM, "      ' サンプルID (From)
        sSQL = sSQL & "SAMPLE_TO, "        ' サンプルID (To)
        sSQL = sSQL & "BLOCKID, "          ' ブロックＩＤ
        sSQL = sSQL & "HINBAN, "           ' 確定品番
        sSQL = sSQL & "KUBUN, "            ' 区分コード
        sSQL = sSQL & "TXID, "             ' トランザクションID
        sSQL = sSQL & "REGDATE, "          ' 登録日付
        sSQL = sSQL & "SUMMITSENDFLAG, "   ' SUMMIT送信フラグ
        sSQL = sSQL & "SENDFLAG, "         ' 送信フラグ
        sSQL = sSQL & "SENDDATE, "         ' 送信日付
        sSQL = sSQL & "PLANTCAT, "         ' 向先 2007/09/04 SPK Tsutsumi Add
        sSQL = sSQL & "MESDATA1TOP, "      ' 測定値１(Top)  center        '04/02/12 ooba START ====>
        sSQL = sSQL & "MESDATA2TOP, "      ' 測定値２(Top)  R/2
        sSQL = sSQL & "MESDATA3TOP, "      ' 測定値３(Top)  Inside 10mm
        sSQL = sSQL & "MESDATA4TOP, "      ' 測定値４(Top)  Inside   6mm
        sSQL = sSQL & "MESDATA5TOP, "      ' 測定値５(Top)  Inside   3mm
        sSQL = sSQL & "MESDATA1BOT, "      ' 測定値１(Tail)  center
        sSQL = sSQL & "MESDATA2BOT, "      ' 測定値２(Tail)  R/2
        sSQL = sSQL & "MESDATA3BOT, "      ' 測定値３(Tail)  Inside 10mm
        sSQL = sSQL & "MESDATA4BOT, "      ' 測定値４(Tail)  Inside   6mm
        sSQL = sSQL & "MESDATA5BOT )"      ' 測定値５(Tail)  Inside   3mm '04/02/12 ooba END ======>

        With SXLKakuSiji(i)
            sSQL = sSQL & "values ("
            sSQL = sSQL & " '" & .SXL_ID & "', "           ' SXL-ID
            sSQL = sSQL & " '" & .SAMPLE_FROM & "', "      ' サンプルID (From)
            sSQL = sSQL & " '" & .SAMPLE_TO & "', "        ' サンプルID (To)
            sSQL = sSQL & " '" & .BLOCKID & "', "          ' ブロックＩＤ
            sSQL = sSQL & " '" & .hinban & "', "           ' 確定品番
            sSQL = sSQL & " '" & .KUBUN & "', "            ' 区分コード
            sSQL = sSQL & " '" & .TXID & "', "             ' トランザクションID
            sSQL = sSQL & " sysdate, "                     ' 登録日付
            sSQL = sSQL & " '0', "                         ' SUMMIT送信フラグ
            sSQL = sSQL & " '3', "                         ' 送信フラグ   'upd 2003/06/05 hitec)matsumoto 送信フラグを3に変更
            sSQL = sSQL & " sysdate, "                     ' 送信日付
            sSQL = sSQL & " '" & sCmbMukesaki & "', "      ' 向先 2007/09/04 SPK Tsutsumi Add
            sSQL = sSQL & " '" & .MESDATA1TOP & "', "      ' 測定値１(Top)  center        '04/02/12 ooba START ====>
            sSQL = sSQL & " '" & .MESDATA2TOP & "', "      ' 測定値２(Top)  R/2
            sSQL = sSQL & " '" & .MESDATA3TOP & "', "      ' 測定値３(Top)  Inside 10mm
            sSQL = sSQL & " '" & .MESDATA4TOP & "', "      ' 測定値４(Top)  Inside   6mm
            sSQL = sSQL & " '" & .MESDATA5TOP & "', "      ' 測定値５(Top)  Inside   3mm
            sSQL = sSQL & " '" & .MESDATA1BOT & "', "      ' 測定値１(Tail)  center
            sSQL = sSQL & " '" & .MESDATA2BOT & "', "      ' 測定値２(Tail)  R/2
            sSQL = sSQL & " '" & .MESDATA3BOT & "', "      ' 測定値３(Tail)  Inside 10mm
            sSQL = sSQL & " '" & .MESDATA4BOT & "', "      ' 測定値４(Tail)  Inside   6mm
            sSQL = sSQL & " '" & .MESDATA5BOT & "' ) "     ' 測定値５(Tail)  Inside   3mm '04/02/12 ooba END ======>
        End With

        '' WriteDBLog sSql, sDBName
        If 0 >= OraDB.ExecuteSQL(sSQL) Then
            DBDRV_scmzc_fcmlc001d_Exec = FUNCTION_RETURN_FAILURE
            sErrMsg = GetMsgStr("EAPLY") & sDBName
            GoTo proc_exit
        End If

    Next

    '関連ブロック情報登録停止　08/01/23 ooba
''    '関連ﾌﾞﾛｯｸ情報登録　07/08/06 ooba START =====================================>
''    If UBound(tSXLID) > 0 Then
''        sDbName = "(Y023)"
''        If DBDRV_KanrenBlk(WfHantei.CRYNUM, tSXLID(), _
''                            SIngotP, EIngotP) = FUNCTION_RETURN_FAILURE Then
''
''            sErrMsg = GetMsgStr("EAPLY") & sDbName
''            GoTo proc_exit
''        End If
''    End If
''    '関連ﾌﾞﾛｯｸ情報登録　07/08/06 ooba END =======================================>

    DBDRV_scmzc_fcmlc001d_Exec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    Resume proc_exit
End Function

'*******************************************************************************
'*    関数名        : GetSxlidINBlkid
'*
'*    処理概要      : 1.再抜試指示 表示用ＤＢドライバ（WFサンプル管理）
'*                      (欠落情報の取得)
'*
'*    パラメータ    : 変数名       ,IO   ,型                ,説明
'*                    CRYNUM       ,I    ,String            ,結晶番号
'*                    pLackWaf     ,O    ,typ_LackWaf       ,欠落情報
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function DBDRV_scmzc_fcmlc001d_LostInfo(CRYNUM As String, _
                                            pLackWaf() As typ_LackWaf _
                                            ) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim lngRecCnt   As Long
    Dim i           As Long
    Dim rs          As OraDynaset

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001d_LostInfo"

     '' 欠落ウェハー情報の取得
    sSQL = "select distinct LOTID as BLOCKID, REJWFFROM as WAFERNO, REJFROM as TOP_POS, REJTO as TAIL_POS "
    sSQL = sSQL & "from VECMW005 "
    sSQL = sSQL & "where (REJCAT='A' or ALLSCRAP='Y') "
    sSQL = sSQL & "  and LOTID like '" & left$(CRYNUM, 9) & "%'"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    lngRecCnt = rs.RecordCount

    ReDim pLackWaf(lngRecCnt)

    For i = 1 To lngRecCnt
        DoEvents
        With pLackWaf(i)
            .BLOCKID = rs("BLOCKID")    ' ブロックID
            .WAFERNO = rs("WAFERNO")    ' ウェハー連番
            .TOP_POS = rs("TOP_POS")    ' ウェハー開始位置
            .TAIL_POS = rs("TAIL_POS")  ' ウェハー終了位置
        End With
        rs.MoveNext
    Next i
    rs.Close

    DBDRV_scmzc_fcmlc001d_LostInfo = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    '' WriteDBLog " ", "End"
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    Resume proc_exit
End Function

'********************************************************************************************
'*    関数名        : DBDRV_GetDSODSpec
'*
'*    処理概要      : 1.品WFDSOD検査情報取得
'*                      (使用していない)
'*    パラメータ    : 変数名        ,IO ,型                      ,説明
'*                    HIN           ,O  ,tFullHinban　           ,品番情報
'*                    HWFDSOKE      ,O  ,String     　           ,品WFDSOD検査
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'********************************************************************************************
Public Function DBDRV_GetDSODSpec(HIN As tFullHinban, HWFDSOKE As String) As FUNCTION_RETURN
    Dim sSQL    As String
    Dim rs      As OraDynaset

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_GetDSODSpec"
    DBDRV_GetDSODSpec = FUNCTION_RETURN_FAILURE

    sSQL = "select HWFDSOKE "
    sSQL = sSQL & "from TBCME026 "
    sSQL = sSQL & "where HINBAN = '" & HIN.hinban & "' and "
    sSQL = sSQL & "MNOREVNO = " & HIN.mnorevno & " and "
    sSQL = sSQL & "FACTORY = '" & HIN.factory & "' and "
    sSQL = sSQL & "OPECOND = '" & HIN.opecond & "'"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount = 0 Then
        rs.Close
        GoTo proc_exit
    End If

    HWFDSOKE = rs("HWFDSOKE") ' 品WFDSOD検査

    rs.Close

    DBDRV_GetDSODSpec = FUNCTION_RETURN_SUCCESS

proc_exit:
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    Resume proc_exit
End Function

'*******************************************************************************************
'*    関数名        : DBDRV_GetDZSpec
'*
'*    処理概要      : 1.品WFDSOD検査情報取得
'*                      (使用していない)
'*
'*    パラメータ    : 変数名        ,IO ,型                      ,説明
'*                    HIN           ,O  ,tFullHinban　           ,品番情報
'*                    HWFMKSZY      ,O  ,String     　           ,品WF無欠陥層測定条件
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************
Public Function DBDRV_GetDZSpec(HIN As tFullHinban, HWFMKSZY As String) As FUNCTION_RETURN
    Dim sql As String
    Dim rs As OraDynaset

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_GetDZSpec"
    DBDRV_GetDZSpec = FUNCTION_RETURN_FAILURE

    sql = "select HWFDSOKE "
    sql = sql & "from TBCME026 "
    sql = sql & "where HINBAN = '" & HIN.hinban & "' and "
    sql = sql & "MNOREVNO = " & HIN.mnorevno & " and "
    sql = sql & "FACTORY = '" & HIN.factory & "' and "
    sql = sql & "OPECOND = '" & HIN.opecond & "'"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount = 0 Then
        rs.Close
        GoTo proc_exit
    End If

    HWFMKSZY = rs("HWFMKSZY") ' 品WF無欠陥層測定条件

    rs.Close

    DBDRV_GetDZSpec = FUNCTION_RETURN_SUCCESS

proc_exit:
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'*****************************************************************************************************
'*    関数名        : DBDRV_GetNoTestHinInfo
'*
'*    処理概要      : 1.SXLの全ブロック入庫チェック
'*
'*    パラメータ    : 変数名        ,IO ,型             ,説明
'*                    HIN           ,I  ,tFullHinban    ,品番情報
'*                    Inf           ,O  ,NoTest_Info    ,
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*****************************************************************************************************
Public Function DBDRV_GetNoTestHinInfo(HIN() As tFullHinban, Inf() As NoTest_Info) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim lngRecCnt   As Long
    Dim c0          As Long

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_GetNoTestHinInfo"
    DBDRV_GetNoTestHinInfo = FUNCTION_RETURN_FAILURE

    For c0 = 0 To 1
        sSQL = "select "
        sSQL = sSQL & "HWFRHWYS " '品WF比抵抗保証方法＿処
        sSQL = sSQL & "from TBCME021 "
        sSQL = sSQL & "where "
        sSQL = sSQL & "HINBAN = '" & HIN(c0).hinban & "' and "
        sSQL = sSQL & "MNOREVNO = " & HIN(c0).mnorevno & " and "
        sSQL = sSQL & "FACTORY = '" & HIN(c0).factory & "' and "
        sSQL = sSQL & "OPECOND = '" & HIN(c0).opecond & "'"

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        lngRecCnt = rs.RecordCount

        If lngRecCnt = 0 Then
            rs.Close
            GoTo proc_exit
        End If

        Inf(c0).Res.HWFRHWYS = rs("HWFRHWYS")  '品WF比抵抗保証方法＿処
        rs.Close

        sSQL = "select "
        sSQL = sSQL & "HWFMKHWS, " '品WF無欠陥層保証方法＿処
        sSQL = sSQL & "HWFMKSZY, " '品WF無欠陥層測定条件
        sSQL = sSQL & "HWFMKSPH, " '品WF無欠陥層測定位置_方
        sSQL = sSQL & "HWFMKSPT, " '品WF無欠陥層測定位置_点
        sSQL = sSQL & "HWFMKSPR " '品WF無欠陥層測定位置_領
        sSQL = sSQL & "from TBCME024 "
        sSQL = sSQL & "where "
        sSQL = sSQL & "HINBAN = '" & HIN(c0).hinban & "' and "
        sSQL = sSQL & "MNOREVNO = " & HIN(c0).mnorevno & " and "
        sSQL = sSQL & "FACTORY = '" & HIN(c0).factory & "' and "
        sSQL = sSQL & "OPECOND = '" & HIN(c0).opecond & "'"

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        lngRecCnt = rs.RecordCount

        If lngRecCnt = 0 Then
            rs.Close
            GoTo proc_exit
        End If

        Inf(c0).DZ.HWFMKHWS = rs("HWFMKHWS") '品WF無欠陥層保証方法＿処
        Inf(c0).DZ.HWFMKSZY = rs("HWFMKSZY") '品WF無欠陥層測定条件
        Inf(c0).DZ.HWFMKSPH = rs("HWFMKSPH") '品WF無欠陥層測定位置_方
        Inf(c0).DZ.HWFMKSPT = rs("HWFMKSPT") '品WF無欠陥層測定位置_点
        Inf(c0).DZ.HWFMKSPR = rs("HWFMKSPR") '品WF無欠陥層測定位置_領
        rs.Close

        sSQL = "select "
        sSQL = sSQL & "HWFONHWS, " '品WF酸素濃度保証方法＿処
        sSQL = sSQL & "HWFOS1HS, " '品WF酸素析出1保証方法＿処
        sSQL = sSQL & "HWFOS1NS, " '品WF酸素析出1熱処理法
        sSQL = sSQL & "HWFOS1SH, " '品WF酸素析出1測定位置_方
        sSQL = sSQL & "HWFOS1ST, " '品WF酸素析出1測定位置_点
        sSQL = sSQL & "HWFOS1SI, " '品WF酸素析出1測定位置＿位
        sSQL = sSQL & "HWFOS2HS, " '品WF酸素析出2保証方法＿処
        sSQL = sSQL & "HWFOS2NS, " '品WF酸素析出2熱処理法
        sSQL = sSQL & "HWFOS2SH, " '品WF酸素析出2測定位置_方
        sSQL = sSQL & "HWFOS2ST, " '品WF酸素析出2測定位置_点
        sSQL = sSQL & "HWFOS2SI, " '品WF酸素析出2測定位置＿位
        sSQL = sSQL & "HWFOS3HS, " '品WF酸素析出3保証方法＿処
        sSQL = sSQL & "HWFOS3NS, " '品WF酸素析出3熱処理法
        sSQL = sSQL & "HWFOS3SH, " '品WF酸素析出3測定位置_方
        sSQL = sSQL & "HWFOS3ST, " '品WF酸素析出3測定位置_点
        sSQL = sSQL & "HWFOS3SI " '品WF酸素析出3測定位置＿位
        sSQL = sSQL & "from TBCME025 "
        sSQL = sSQL & "where "
        sSQL = sSQL & "HINBAN = '" & HIN(c0).hinban & "' and "
        sSQL = sSQL & "MNOREVNO = " & HIN(c0).mnorevno & " and "
        sSQL = sSQL & "FACTORY = '" & HIN(c0).factory & "' and "
        sSQL = sSQL & "OPECOND = '" & HIN(c0).opecond & "'"

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        lngRecCnt = rs.RecordCount

        If lngRecCnt = 0 Then
            rs.Close
            GoTo proc_exit
        End If

        Inf(c0).Oi.HWFONHWS = rs("HWFONHWS") '品WF酸素濃度保証方法＿処
        Inf(c0).Doi(0).HWFOSxHS = rs("HWFOS1HS")  '品WF酸素析出1保証方法＿処
        Inf(c0).Doi(0).HWFOSxNS = rs("HWFOS1NS")  '品WF酸素析出1熱処理法
        Inf(c0).Doi(0).HWFOSxSH = rs("HWFOS1SH")  '品WF酸素析出1測定位置_方
        Inf(c0).Doi(0).HWFOSxST = rs("HWFOS1ST")  '品WF酸素析出1測定位置_点
        Inf(c0).Doi(0).HWFOSxSI = rs("HWFOS1SI")  '品WF酸素析出1測定位置＿位
        Inf(c0).Doi(1).HWFOSxHS = rs("HWFOS2HS")  '品WF酸素析出2保証方法＿処
        Inf(c0).Doi(1).HWFOSxNS = rs("HWFOS2NS")  '品WF酸素析出2熱処理法
        Inf(c0).Doi(1).HWFOSxSH = rs("HWFOS2SH")  '品WF酸素析出2測定位置_方
        Inf(c0).Doi(1).HWFOSxST = rs("HWFOS2ST")  '品WF酸素析出2測定位置_点
        Inf(c0).Doi(1).HWFOSxSI = rs("HWFOS2SI")  '品WF酸素析出2測定位置＿位
        Inf(c0).Doi(2).HWFOSxHS = rs("HWFOS3HS")  '品WF酸素析出3保証方法＿処
        Inf(c0).Doi(2).HWFOSxNS = rs("HWFOS3NS")  '品WF酸素析出3熱処理法
        Inf(c0).Doi(2).HWFOSxSH = rs("HWFOS3SH")  '品WF酸素析出3測定位置_方
        Inf(c0).Doi(2).HWFOSxST = rs("HWFOS3ST")  '品WF酸素析出3測定位置_点
        Inf(c0).Doi(2).HWFOSxSI = rs("HWFOS3SI")  '品WF酸素析出3測定位置＿位
        rs.Close


        sSQL = "select "
        sSQL = sSQL & "HWFDSOHS, " '品WFDSOD保証方法＿処
        sSQL = sSQL & "HWFDSOKE " '品WFDSOD検査"
        sSQL = sSQL & "from TBCME026 "
        sSQL = sSQL & "where "
        sSQL = sSQL & "HINBAN = '" & HIN(c0).hinban & "' and "
        sSQL = sSQL & "MNOREVNO = " & HIN(c0).mnorevno & " and "
        sSQL = sSQL & "FACTORY = '" & HIN(c0).factory & "' and "
        sSQL = sSQL & "OPECOND = '" & HIN(c0).opecond & "'"

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        lngRecCnt = rs.RecordCount

        If lngRecCnt = 0 Then
            rs.Close
            GoTo proc_exit
        End If

        Inf(c0).Dsod.HWFDSOHS = rs("HWFDSOHS") '品WFDSOD保証方法＿処
        Inf(c0).Dsod.HWFDSOKE = rs("HWFDSOKE") '品WFDSOD検査"
        rs.Close

        sSQL = "select "
        sSQL = sSQL & "HWFSPVHS, " '品WFSPVFE保証方法＿処
        sSQL = sSQL & "HWFSPVSH, " '品WFSPVFE測定位置_方
        sSQL = sSQL & "HWFSPVST, " '品WFSPVFE測定位置_点
        sSQL = sSQL & "HWFSPVSI, " '品WFSPVFE測定位置＿位
        sSQL = sSQL & "HWFDLHWS, " '品WF拡散長保証方法＿処
        sSQL = sSQL & "HWFDLSPH, " '品WF拡散長測定位置_方
        sSQL = sSQL & "HWFDLSPT, " '品WF拡散長測定位置_点
        sSQL = sSQL & "HWFDLSPI " '品WF拡散長測定位置＿位
        sSQL = sSQL & "from TBCME028 "
        sSQL = sSQL & "where "
        sSQL = sSQL & "HINBAN = '" & HIN(c0).hinban & "' and "
        sSQL = sSQL & "MNOREVNO = " & HIN(c0).mnorevno & " and "
        sSQL = sSQL & "FACTORY = '" & HIN(c0).factory & "' and "
        sSQL = sSQL & "OPECOND = '" & HIN(c0).opecond & "'"

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        lngRecCnt = rs.RecordCount

        If lngRecCnt = 0 Then
            rs.Close
            GoTo proc_exit
        End If

        Inf(c0).SpvFe.HWFSPVHS = rs("HWFSPVHS") '品WFSPVFE保証方法＿処
        Inf(c0).SpvFe.HWFSPVSH = rs("HWFSPVSH") '品WFSPVFE測定位置_方
        Inf(c0).SpvFe.HWFSPVST = rs("HWFSPVST") '品WFSPVFE測定位置_点
        Inf(c0).SpvFe.HWFSPVSI = rs("HWFSPVSI") '品WFSPVFE測定位置＿位
        Inf(c0).Spv.HWFDLHWS = rs("HWFDLHWS") '品WF拡散長保証方法＿処
        Inf(c0).Spv.HWFDLSPH = rs("HWFDLSPH")  '品WF拡散長測定位置_方
        Inf(c0).Spv.HWFDLSPT = rs("HWFDLSPT") '品WF拡散長測定位置_点
        Inf(c0).Spv.HWFDLSPI = rs("HWFDLSPI") '品WF拡散長測定位置＿位
        rs.Close

        sSQL = "select "
        sSQL = sSQL & "HWFBM1HS, " '品WFBMD1保証方法＿処
        sSQL = sSQL & "HWFBM1ET, " '品WFBMD1選択ET代
        sSQL = sSQL & "HWFBM1NS, " '品WFBMD1熱処理法
        sSQL = sSQL & "HWFBM1SZ, " '品WFBMD1測定条件
        sSQL = sSQL & "HWFBM1SH, " '品WFBMD1測定位置_方
        sSQL = sSQL & "HWFBM1ST, " '品WFBMD1測定位置_点
        sSQL = sSQL & "HWFBM1SR, " '品WFBMD1測定位置_領
        sSQL = sSQL & "HWFBM2HS, " '品WFBMD2保証方法＿処
        sSQL = sSQL & "HWFBM2ET, " '品WFBMD2選択ET代
        sSQL = sSQL & "HWFBM2NS, " '品WFBMD2熱処理法
        sSQL = sSQL & "HWFBM2SZ, " '品WFBMD2測定条件
        sSQL = sSQL & "HWFBM2SH, " '品WFBMD2測定位置_方
        sSQL = sSQL & "HWFBM2ST, " '品WFBMD2測定位置_点
        sSQL = sSQL & "HWFBM2SR, " '品WFBMD2測定位置_領
        sSQL = sSQL & "HWFBM3HS, " '品WFBMD3保証方法＿処
        sSQL = sSQL & "HWFBM3ET, " '品WFBMD3選択ET代
        sSQL = sSQL & "HWFBM3NS, " '品WFBMD3熱処理法
        sSQL = sSQL & "HWFBM3SZ, " '品WFBMD3測定条件
        sSQL = sSQL & "HWFBM3SH, " '品WFBMD3測定位置_方
        sSQL = sSQL & "HWFBM3ST, " '品WFBMD3測定位置_点
        sSQL = sSQL & "HWFBM3SR, " '品WFBMD3測定位置_領
        sSQL = sSQL & "HWFOF1HS, " '品WFOSF1保証方法＿処
        sSQL = sSQL & "HWFOF1ET, " '品WFOSF1選択ET代
        sSQL = sSQL & "HWFOF1NS, " '品WFOSF1熱処理法
        sSQL = sSQL & "HWFOF1SZ, " '品WFOSF1測定条件
        sSQL = sSQL & "HWFOF1SH, " '品WFOSF1測定位置_方
        sSQL = sSQL & "HWFOF1ST, " '品WFOSF1測定位置_点
        sSQL = sSQL & "HWFOF1SR, " '品WFOSF1測定位置_領
        sSQL = sSQL & "HWFOF2HS, " '品WFOSF2保証方法＿処
        sSQL = sSQL & "HWFOF2ET, " '品WFOSF2選択ET代
        sSQL = sSQL & "HWFOF2NS, " '品WFOSF2熱処理法
        sSQL = sSQL & "HWFOF2SZ, " '品WFOSF2測定条件
        sSQL = sSQL & "HWFOF2SH, " '品WFOSF2測定位置_方
        sSQL = sSQL & "HWFOF2ST, " '品WFOSF2測定位置_点
        sSQL = sSQL & "HWFOF2SR, " '品WFOSF2測定位置_領
        sSQL = sSQL & "HWFOF3HS, " '品WFOSF3保証方法＿処
        sSQL = sSQL & "HWFOF3ET, " '品WFOSF3選択ET代
        sSQL = sSQL & "HWFOF3NS, " '品WFOSF3熱処理法
        sSQL = sSQL & "HWFOF3SZ, " '品WFOSF3測定条件
        sSQL = sSQL & "HWFOF3SH, " '品WFOSF3測定位置_方
        sSQL = sSQL & "HWFOF3ST, " '品WFOSF3測定位置_点
        sSQL = sSQL & "HWFOF3SR, " '品WFOSF3測定位置_領
        sSQL = sSQL & "HWFOF4HS, " '品WFOSF4保証方法＿処
        sSQL = sSQL & "HWFOF4ET, " '品WFOSF4選択ET代
        sSQL = sSQL & "HWFOF4NS, " '品WFOSF4熱処理法
        sSQL = sSQL & "HWFOF4SZ, " '品WFOSF4測定条件
        sSQL = sSQL & "HWFOF4SH, " '品WFOSF4測定位置_方
        sSQL = sSQL & "HWFOF4ST, " '品WFOSF4測定位置_点
        sSQL = sSQL & "HWFOF4SR " '品WFOSF4測定位置_領
        sSQL = sSQL & "from TBCME029 "
        sSQL = sSQL & "where "
        sSQL = sSQL & "HINBAN = '" & HIN(c0).hinban & "' and "
        sSQL = sSQL & "MNOREVNO = " & HIN(c0).mnorevno & " and "
        sSQL = sSQL & "FACTORY = '" & HIN(c0).factory & "' and "
        sSQL = sSQL & "OPECOND = '" & HIN(c0).opecond & "'"

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        lngRecCnt = rs.RecordCount

        If lngRecCnt = 0 Then
            rs.Close
            GoTo proc_exit
        End If

        Inf(c0).BMD(0).HWFBMxHS = rs("HWFBM1HS") '品WFBMD1保証方法＿処
        Inf(c0).BMD(0).HWFBMxET = rs("HWFBM1ET") '品WFBMD1選択ET代
        Inf(c0).BMD(0).HWFBMxNS = rs("HWFBM1NS") '品WFBMD1熱処理法
        Inf(c0).BMD(0).HWFBMxSZ = rs("HWFBM1SZ") '品WFBMD1測定条件
        Inf(c0).BMD(0).HWFBMxSH = rs("HWFBM1SH") '品WFBMD1測定位置_方
        Inf(c0).BMD(0).HWFBMxST = rs("HWFBM1ST") '品WFBMD1測定位置_点
        Inf(c0).BMD(0).HWFBMxSR = rs("HWFBM1SR") '品WFBMD1測定位置_領
        Inf(c0).BMD(1).HWFBMxHS = rs("HWFBM2HS") '品WFBMD2保証方法＿処
        Inf(c0).BMD(1).HWFBMxET = rs("HWFBM2ET") '品WFBMD2選択ET代
        Inf(c0).BMD(1).HWFBMxNS = rs("HWFBM2NS") '品WFBMD2熱処理法
        Inf(c0).BMD(1).HWFBMxSZ = rs("HWFBM2SZ") '品WFBMD2測定条件
        Inf(c0).BMD(1).HWFBMxSH = rs("HWFBM2SH") '品WFBMD2測定位置_方
        Inf(c0).BMD(1).HWFBMxST = rs("HWFBM2ST") '品WFBMD2測定位置_点
        Inf(c0).BMD(1).HWFBMxSR = rs("HWFBM2SR") '品WFBMD2測定位置_領
        Inf(c0).BMD(2).HWFBMxHS = rs("HWFBM3HS") '品WFBMD3保証方法＿処
        Inf(c0).BMD(2).HWFBMxET = rs("HWFBM3ET") '品WFBMD3選択ET代
        Inf(c0).BMD(2).HWFBMxNS = rs("HWFBM3NS") '品WFBMD3熱処理法
        Inf(c0).BMD(2).HWFBMxSZ = rs("HWFBM3SZ") '品WFBMD3測定条件
        Inf(c0).BMD(2).HWFBMxSH = rs("HWFBM3SH") '品WFBMD3測定位置_方
        Inf(c0).BMD(2).HWFBMxST = rs("HWFBM3ST") '品WFBMD3測定位置_点
        Inf(c0).BMD(2).HWFBMxSR = rs("HWFBM3SR") '品WFBMD3測定位置_領
        Inf(c0).OSF(0).HWFOFxHS = rs("HWFOF1HS") '品WFOSF1保証方法＿処
        Inf(c0).OSF(0).HWFOFxET = rs("HWFOF1ET") '品WFOSF1選択ET代
        Inf(c0).OSF(0).HWFOFxNS = rs("HWFOF1NS") '品WFOSF1熱処理法
        Inf(c0).OSF(0).HWFOFxSZ = rs("HWFOF1SZ") '品WFOSF1測定条件
        Inf(c0).OSF(0).HWFOFxSH = rs("HWFOF1SH") '品WFOSF1測定位置_方
        Inf(c0).OSF(0).HWFOFxST = rs("HWFOF1ST") '品WFOSF1測定位置_点
        Inf(c0).OSF(0).HWFOFxSR = rs("HWFOF1SR") '品WFOSF1測定位置_領
        Inf(c0).OSF(1).HWFOFxHS = rs("HWFOF2HS") '品WFOSF2保証方法＿処
        Inf(c0).OSF(1).HWFOFxET = rs("HWFOF2ET") '品WFOSF2選択ET代
        Inf(c0).OSF(1).HWFOFxNS = rs("HWFOF2NS") '品WFOSF2熱処理法
        Inf(c0).OSF(1).HWFOFxSZ = rs("HWFOF2SZ") '品WFOSF2測定条件
        Inf(c0).OSF(1).HWFOFxSH = rs("HWFOF2SH") '品WFOSF2測定位置_方
        Inf(c0).OSF(1).HWFOFxST = rs("HWFOF2ST") '品WFOSF2測定位置_点
        Inf(c0).OSF(1).HWFOFxSR = rs("HWFOF2SR") '品WFOSF2測定位置_領
        Inf(c0).OSF(2).HWFOFxHS = rs("HWFOF3HS") '品WFOSF3保証方法＿処
        Inf(c0).OSF(2).HWFOFxET = rs("HWFOF3ET") '品WFOSF3選択ET代
        Inf(c0).OSF(2).HWFOFxNS = rs("HWFOF3NS") '品WFOSF3熱処理法
        Inf(c0).OSF(2).HWFOFxSZ = rs("HWFOF3SZ") '品WFOSF3測定条件
        Inf(c0).OSF(2).HWFOFxSH = rs("HWFOF3SH") '品WFOSF3測定位置_方
        Inf(c0).OSF(2).HWFOFxST = rs("HWFOF3ST") '品WFOSF3測定位置_点
        Inf(c0).OSF(2).HWFOFxSR = rs("HWFOF3SR") '品WFOSF3測定位置_領
        Inf(c0).OSF(3).HWFOFxHS = rs("HWFOF4HS") '品WFOSF4保証方法＿処
        Inf(c0).OSF(3).HWFOFxET = rs("HWFOF4ET") '品WFOSF4選択ET代
        Inf(c0).OSF(3).HWFOFxNS = rs("HWFOF4NS") '品WFOSF4熱処理法
        Inf(c0).OSF(3).HWFOFxSZ = rs("HWFOF4SZ") '品WFOSF4測定条件
        Inf(c0).OSF(3).HWFOFxSH = rs("HWFOF4SH") '品WFOSF4測定位置_方
        Inf(c0).OSF(3).HWFOFxST = rs("HWFOF4ST") '品WFOSF4測定位置_点
        Inf(c0).OSF(3).HWFOFxSR = rs("HWFOF4SR") '品WFOSF4測定位置_領
        rs.Close
    Next

    DBDRV_GetNoTestHinInfo = FUNCTION_RETURN_SUCCESS

proc_exit:
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    Resume proc_exit
End Function

'*******************************************************************************
'*    関数名        : scmzc_getWF
'*
'*    処理概要      : 1.製品仕様WFデータの取得ドライバ
'*
'*    パラメータ    : 変数名        ,IO ,型               ,説明
'*　　                pSpWFSamp　　 ,IO ,typ_SpWFSamp   　,WFサンプル仕様
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function scmzc_getWF(pSpWFSamp As typ_SpWFSamp) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sSQL    As String
    Dim sOT1    As String
    Dim sOT2    As String
    Dim sMAI1   As String     '04/07/16
    Dim sMAI2   As String
    Dim rtn     As FUNCTION_RETURN

     '' エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function scmzc_getWF"

    '' 製品仕様の取得
    'DK温度追加      08/08/25 Systech
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''    sSql = "select " & _
'''          "E021HWFRSPOH, E021HWFRSPOT, E021HWFRSPOI, E021HWFRHWYS, E024HWFMKSPH, " & _
'''          "E024HWFMKSPT, E024HWFMKSPR, E024HWFMKHWS, E024HWFMKSZY, E024HWFMKNSW, " & _
'''          "E024HWFMKCET, E025HWFONSPH, E025HWFONSPT, E025HWFONSPI, E025HWFONHWS, " & _
'''          "E025HWFONKWY, E025HWFOS1NS, E025HWFOS1SH, E025HWFOS1ST, E025HWFOS1SI, " & _
'''          "E025HWFOS1HS, E025HWFOS2NS, E025HWFOS2SH, E025HWFOS2ST, E025HWFOS2SI, " & _
'''          "E025HWFOS2HS, E025HWFOS3NS, E025HWFOS3SH, E025HWFOS3ST, E025HWFOS3SI, " & _
'''          "E025HWFOS3HS, E025HWFANTNP, E025HWFANTIM, E026HWFDSOHS, E028HWFSPVSH, " & _
'''          "E028HWFSPVST, E028HWFSPVSI, E028HWFSPVHS, E028HWFDLSPH, E028HWFDLSPT, " & _
'''          "E028HWFDLSPI, E028HWFDLHWS, E029HWFOF1ET, E029HWFOF1NS, E029HWFOF1SZ, " & _
'''          "E029HWFOF1SH, E029HWFOF1ST, E029HWFOF1SR, E029HWFOF1HS, E029HWFOF2ET, " & _
'''          "E029HWFOF2NS, E029HWFOF2SZ, E029HWFOF2SH, E029HWFOF2ST, E029HWFOF2SR, " & _
'''          "E029HWFOF2HS, E029HWFOF3ET, E029HWFOF3NS, E029HWFOF3SZ, E029HWFOF3SH, " & _
'''          "E029HWFOF3ST, E029HWFOF3SR, E029HWFOF3HS, E029HWFOF4ET, E029HWFOF4NS, " & _
'''          "E029HWFOF4SZ, E029HWFOF4SH, E029HWFOF4ST, E029HWFOF4SR, E029HWFOF4HS, " & _
'''          "E029HWFBM1ET, E029HWFBM1NS, E029HWFBM1SZ, E029HWFBM1SH, E029HWFBM1ST, " & _
'''          "E029HWFBM1SR, E029HWFBM1HS, E029HWFBM2ET, E029HWFBM2NS, E029HWFBM2SZ, " & _
'''          "E029HWFBM2SH, E029HWFBM2ST, E029HWFBM2SR, E029HWFBM2HS, E029HWFBM3ET, " & _
'''          "E029HWFBM3NS, E029HWFBM3SZ, E029HWFBM3SH, E029HWFBM3ST, E029HWFBM3SR, E029HWFBM3HS" & _
'''          ", NVL(U.HSXDKTMP, ' ') as HSXDKTMP" & _
'''          " from  VECME001,TBCME036 U" & _
'''          " where E018HINBAN='" & pSpWFSamp.HIN.hinban & "' and E018MNOREVNO=" & pSpWFSamp.HIN.mnorevno & _
'''          " and E018FACTORY='" & pSpWFSamp.HIN.factory & "' and E018OPECOND='" & pSpWFSamp.HIN.opecond & "'" & _
'''          " and U.HINBAN = E018HINBAN and U.MNOREVNO = E018MNOREVNO and U.FACTORY = E018FACTORY and U.OPECOND = E018OPECOND"

    sSQL = "select " & _
          "E021HWFRSPOH, E021HWFRSPOT, E021HWFRSPOI, E021HWFRHWYS, E024HWFMKSPH, " & _
          "E024HWFMKSPT, E024HWFMKSPR, E024HWFMKHWS, E024HWFMKSZY, E024HWFMKNSW, " & _
          "E024HWFMKCET, E025HWFONSPH, E025HWFONSPT, E025HWFONSPI, E025HWFONHWS, " & _
          "E025HWFONKWY, E025HWFOS1NS, E025HWFOS1SH, E025HWFOS1ST, E025HWFOS1SI, " & _
          "E025HWFOS1HS, E025HWFOS2NS, E025HWFOS2SH, E025HWFOS2ST, E025HWFOS2SI, " & _
          "E025HWFOS2HS, E025HWFOS3NS, E025HWFOS3SH, E025HWFOS3ST, E025HWFOS3SI, " & _
          "E025HWFOS3HS, E025HWFANTNP, E025HWFANTIM, E026HWFDSOHS, E028HWFSPVSH, " & _
          "E028HWFSPVST, E028HWFSPVSI, E028HWFSPVHS, E028HWFDLSPH, E028HWFDLSPT, " & _
          "E028HWFDLSPI, E028HWFDLHWS, E029HWFOF1ET, E029HWFOF1NS, E029HWFOF1SZ, " & _
          "E029HWFOF1SH, E029HWFOF1ST, E029HWFOF1SR, E029HWFOF1HS, E029HWFOF2ET, " & _
          "E029HWFOF2NS, E029HWFOF2SZ, E029HWFOF2SH, E029HWFOF2ST, E029HWFOF2SR, " & _
          "E029HWFOF2HS, E029HWFOF3ET, E029HWFOF3NS, E029HWFOF3SZ, E029HWFOF3SH, " & _
          "E029HWFOF3ST, E029HWFOF3SR, E029HWFOF3HS, " & _
          "E029HWFBM1ET, E029HWFBM1NS, E029HWFBM1SZ, E029HWFBM1SH, E029HWFBM1ST, " & _
          "E029HWFBM1SR, E029HWFBM1HS, E029HWFBM2ET, E029HWFBM2NS, E029HWFBM2SZ, " & _
          "E029HWFBM2SH, E029HWFBM2ST, E029HWFBM2SR, E029HWFBM2HS, E029HWFBM3ET, " & _
          "E029HWFBM3NS, E029HWFBM3SZ, E029HWFBM3SH, E029HWFBM3ST, E029HWFBM3SR, E029HWFBM3HS" & _
          ", NVL(U.HSXDKTMP, ' ') as HSXDKTMP" & _
          " from  VECME001,TBCME036 U" & _
          " where E018HINBAN='" & pSpWFSamp.HIN.hinban & "' and E018MNOREVNO=" & pSpWFSamp.HIN.mnorevno & _
          " and E018FACTORY='" & pSpWFSamp.HIN.factory & "' and E018OPECOND='" & pSpWFSamp.HIN.opecond & "'" & _
          " and U.HINBAN = E018HINBAN and U.MNOREVNO = E018MNOREVNO and U.FACTORY = E018FACTORY and U.OPECOND = E018OPECOND"
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    With pSpWFSamp
        .HWFRSPOH = rs("E021HWFRSPOH")
        .HWFRSPOT = rs("E021HWFRSPOT")
        .HWFRSPOI = rs("E021HWFRSPOI")
        .HWFRHWYS = rs("E021HWFRHWYS")
        .HWFMKSPH = rs("E024HWFMKSPH")
        .HWFMKSPT = rs("E024HWFMKSPT")
        .HWFMKSPR = rs("E024HWFMKSPR")
        .HWFMKHWS = rs("E024HWFMKHWS")
        .HWFMKSZY = rs("E024HWFMKSZY")
        .HWFMKNSW = rs("E024HWFMKNSW")
        .HWFMKCET = fncNullCheck(rs("E024HWFMKCET"))
        .HWFONSPH = rs("E025HWFONSPH")
        .HWFONSPT = rs("E025HWFONSPT")
        .HWFONSPI = rs("E025HWFONSPI")
        .HWFONHWS = rs("E025HWFONHWS")
        .HWFONKWY = rs("E025HWFONKWY")
        .HWFOS1NS = rs("E025HWFOS1NS")
        .HWFOS1HS = rs("E025HWFOS1HS")
        .HWFOS1SH = rs("E025HWFOS1SH")
        .HWFOS1ST = rs("E025HWFOS1ST")
        .HWFOS1SI = rs("E025HWFOS1SI")
        .HWFOS2NS = rs("E025HWFOS2NS")
        .HWFOS2SH = rs("E025HWFOS2SH")
        .HWFOS2ST = rs("E025HWFOS2ST")
        .HWFOS2SI = rs("E025HWFOS2SI")
        .HWFOS2HS = rs("E025HWFOS2HS")
        .HWFOS3NS = rs("E025HWFOS3NS")
        .HWFOS3SH = rs("E025HWFOS3SH")
        .HWFOS3ST = rs("E025HWFOS3ST")
        .HWFOS3SI = rs("E025HWFOS3SI")
        .HWFOS3HS = rs("E025HWFOS3HS")
        .HWFANTNP = fncNullCheck(rs("E025HWFANTNP"))
        .HWFANTIM = fncNullCheck(rs("E025HWFANTIM"))
        .HWFDSOHS = rs("E026HWFDSOHS")
        .HWFSPVSH = rs("E028HWFSPVSH")
        .HWFSPVST = rs("E028HWFSPVST")
        .HWFSPVSI = rs("E028HWFSPVSI")
        .HWFSPVHS = rs("E028HWFSPVHS")
        .HWFDLSPH = rs("E028HWFDLSPH")
        .HWFDLSPT = rs("E028HWFDLSPT")
        .HWFDLSPI = rs("E028HWFDLSPI")
        .HWFDLHWS = rs("E028HWFDLHWS")
        .HWFOF1ET = fncNullCheck(rs("E029HWFOF1ET"))
        .HWFOF1NS = rs("E029HWFOF1NS")
        .HWFOF1SZ = rs("E029HWFOF1SZ")
        .HWFOF1SH = rs("E029HWFOF1SH")
        .HWFOF1ST = rs("E029HWFOF1ST")
        .HWFOF1SR = rs("E029HWFOF1SR")
        .HWFOF1HS = rs("E029HWFOF1HS")
        .HWFOF2ET = fncNullCheck(rs("E029HWFOF2ET"))
        .HWFOF2NS = rs("E029HWFOF2NS")
        .HWFOF2SZ = rs("E029HWFOF2SZ")
        .HWFOF2SH = rs("E029HWFOF2SH")
        .HWFOF2ST = rs("E029HWFOF2ST")
        .HWFOF2SR = rs("E029HWFOF2SR")
        .HWFOF2HS = rs("E029HWFOF2HS")
        .HWFOF3ET = fncNullCheck(rs("E029HWFOF3ET"))
        .HWFOF3NS = rs("E029HWFOF3NS")
        .HWFOF3SZ = rs("E029HWFOF3SZ")
        .HWFOF3SH = rs("E029HWFOF3SH")
        .HWFOF3ST = rs("E029HWFOF3ST")
        .HWFOF3SR = rs("E029HWFOF3SR")
        .HWFOF3HS = rs("E029HWFOF3HS")
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''        .HWFOF4ET = fncNullCheck(rs("E029HWFOF4ET"))
'''        .HWFOF4NS = rs("E029HWFOF4NS")
'''        .HWFOF4SZ = rs("E029HWFOF4SZ")
'''        .HWFOF4SH = rs("E029HWFOF4SH")
'''        .HWFOF4ST = rs("E029HWFOF4ST")
'''        .HWFOF4SR = rs("E029HWFOF4SR")
'''        .HWFOF4HS = rs("E029HWFOF4HS")
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
        .HWFBM1ET = fncNullCheck(rs("E029HWFBM1ET"))
        .HWFBM1NS = rs("E029HWFBM1NS")
        .HWFBM1SZ = rs("E029HWFBM1SZ")
        .HWFBM1SH = rs("E029HWFBM1SH")
        .HWFBM1ST = rs("E029HWFBM1ST")
        .HWFBM1SR = rs("E029HWFBM1SR")
        .HWFBM1HS = rs("E029HWFBM1HS")
        .HWFBM2ET = fncNullCheck(rs("E029HWFBM2ET"))
        .HWFBM2NS = rs("E029HWFBM2NS")
        .HWFBM2SZ = rs("E029HWFBM2SZ")
        .HWFBM2SH = rs("E029HWFBM2SH")
        .HWFBM2ST = rs("E029HWFBM2ST")
        .HWFBM2SR = rs("E029HWFBM2SR")
        .HWFBM2HS = rs("E029HWFBM2HS")
        .HWFBM3ET = fncNullCheck(rs("E029HWFBM3ET"))
        .HWFBM3NS = rs("E029HWFBM3NS")
        .HWFBM3SZ = rs("E029HWFBM3SZ")
        .HWFBM3SH = rs("E029HWFBM3SH")
        .HWFBM3ST = rs("E029HWFBM3ST")
        .HWFBM3SR = rs("E029HWFBM3SR")
        .HWFBM3HS = rs("E029HWFBM3HS")
'--------------- 2008/08/25 INSERT START  By Systech ---------------
        .HSXDKTMP = rs("HSXDKTMP")
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
        rtn = scmzc_getE036(pSpWFSamp.HIN, sOT1, sOT2, sMAI1, sMAI2)
        If rtn = FUNCTION_RETURN_FAILURE Then
            rs.Close
            scmzc_getWF = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        .HWOTHER1 = sOT1 '### 03/05/26
        .HWOTHER2 = sOT2
        .HWOTHER1MAI = sMAI1    '04/07/16
        .HWOTHER2MAI = sMAI2
    End With
    rs.Close

    '検査頻度_抜ﾃﾞｰﾀ取得　04/04/13 ooba START =================================================>
    sSQL = "select "
    sSQL = sSQL & "TBCME024.HWFANGZY, "               '品WF高温ANｶﾞｽ条件　04/07/29 ooba
    sSQL = sSQL & "TBCME021.HWFRKHNN, "
    sSQL = sSQL & "TBCME025.HWFONKHN, "
    sSQL = sSQL & "TBCME029.HWFOF1KN, "
    sSQL = sSQL & "TBCME029.HWFOF2KN, "
    sSQL = sSQL & "TBCME029.HWFOF3KN, "
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''    sSql = sSql & "TBCME029.HWFOF4KN, "
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
    sSQL = sSQL & "TBCME029.HWFBM1KN, "
    sSQL = sSQL & "TBCME029.HWFBM2KN, "
    sSQL = sSQL & "TBCME029.HWFBM3KN, "
    sSQL = sSQL & "TBCME025.HWFOS1KN, "
    sSQL = sSQL & "TBCME025.HWFOS2KN, "
    sSQL = sSQL & "TBCME025.HWFOS3KN, "
    sSQL = sSQL & "TBCME026.HWFDSOKN, "
    sSQL = sSQL & "TBCME024.HWFMKKHN, "
    sSQL = sSQL & "TBCME028.HWFSPVKN, "
    sSQL = sSQL & "TBCME028.HWFDLKHN, "
    sSQL = sSQL & "TBCME025.HWFZOKHN, "
    sSQL = sSQL & "TBCME026.HWFGDKHN "                '検査頻度_抜(GD)　05/02/18 ooba
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
    sSQL = sSQL & "and TBCME021.HINBAN = '" & pSpWFSamp.HIN.hinban & "' "
    sSQL = sSQL & "and TBCME021.MNOREVNO = " & pSpWFSamp.HIN.mnorevno & " "
    sSQL = sSQL & "and TBCME021.FACTORY = '" & pSpWFSamp.HIN.factory & "' "
    sSQL = sSQL & "and TBCME021.OPECOND = '" & pSpWFSamp.HIN.opecond & "' "

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    With pSpWFSamp
        If IsNull(rs("HWFANGZY")) = False Then .HWFANGZY = rs("HWFANGZY") Else .HWFANGZY = " "  '品WF高温ANｶﾞｽ条件　04/07/29 ooba
        If IsNull(rs("HWFRKHNN")) = False Then .HWFRKHNN = rs("HWFRKHNN") Else .HWFRKHNN = " "
        If IsNull(rs("HWFONKHN")) = False Then .HWFONKHN = rs("HWFONKHN") Else .HWFONKHN = " "
        If IsNull(rs("HWFOF1KN")) = False Then .HWFOF1KN = rs("HWFOF1KN") Else .HWFOF1KN = " "
        If IsNull(rs("HWFOF2KN")) = False Then .HWFOF2KN = rs("HWFOF2KN") Else .HWFOF2KN = " "
        If IsNull(rs("HWFOF3KN")) = False Then .HWFOF3KN = rs("HWFOF3KN") Else .HWFOF3KN = " "
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''        If IsNull(rs("HWFOF4KN")) = False Then .HWFOF4KN = rs("HWFOF4KN") Else .HWFOF4KN = " "
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
        If IsNull(rs("HWFBM1KN")) = False Then .HWFBM1KN = rs("HWFBM1KN") Else .HWFBM1KN = " "
        If IsNull(rs("HWFBM2KN")) = False Then .HWFBM2KN = rs("HWFBM2KN") Else .HWFBM2KN = " "
        If IsNull(rs("HWFBM3KN")) = False Then .HWFBM3KN = rs("HWFBM3KN") Else .HWFBM3KN = " "
        If IsNull(rs("HWFOS1KN")) = False Then .HWFOS1KN = rs("HWFOS1KN") Else .HWFOS1KN = " "
        If IsNull(rs("HWFOS2KN")) = False Then .HWFOS2KN = rs("HWFOS2KN") Else .HWFOS2KN = " "
        If IsNull(rs("HWFOS3KN")) = False Then .HWFOS3KN = rs("HWFOS3KN") Else .HWFOS3KN = " "
        If IsNull(rs("HWFDSOKN")) = False Then .HWFDSOKN = rs("HWFDSOKN") Else .HWFDSOKN = " "
        If IsNull(rs("HWFMKKHN")) = False Then .HWFMKKHN = rs("HWFMKKHN") Else .HWFMKKHN = " "
        If IsNull(rs("HWFSPVKN")) = False Then .HWFSPVKN = rs("HWFSPVKN") Else .HWFSPVKN = " "
        If IsNull(rs("HWFDLKHN")) = False Then .HWFDLKHN = rs("HWFDLKHN") Else .HWFDLKHN = " "
        If IsNull(rs("HWFZOKHN")) = False Then .HWFZOKHN = rs("HWFZOKHN") Else .HWFZOKHN = " "
        If IsNull(rs("HWFGDKHN")) = False Then .HWFGDKHN = rs("HWFGDKHN") Else .HWFGDKHN = " "  '検査頻度_抜(GD)　05/02/18 ooba
    End With
    rs.Close
    '検査頻度_抜ﾃﾞｰﾀ取得　04/04/13 ooba END ===================================================>

    ''残存酸素仕様取得追加　03/12/15 ooba START ================================================>
    sSQL = "select HWFZOHWS, HWFZOSPH, HWFZOSPT, HWFZOSPI, HWFZONSW from TBCME025 "
    sSQL = sSQL & "where HINBAN = '" & pSpWFSamp.HIN.hinban & "' "
    sSQL = sSQL & "and MNOREVNO = " & pSpWFSamp.HIN.mnorevno & " "
    sSQL = sSQL & "and FACTORY = '" & pSpWFSamp.HIN.factory & "' "
    sSQL = sSQL & "and OPECOND = '" & pSpWFSamp.HIN.opecond & "' "

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    If IsNull(rs("HWFZOHWS")) = False Then pSpWFSamp.HWFZOHWS = rs("HWFZOHWS") Else pSpWFSamp.HWFZOHWS = " "  ' 処理方法(AO)
    If IsNull(rs("HWFZOSPH")) = False Then pSpWFSamp.HWFZOSPH = rs("HWFZOSPH") Else pSpWFSamp.HWFZOSPH = " "  ' 測定方法(AO)
    If IsNull(rs("HWFZOSPT")) = False Then pSpWFSamp.HWFZOSPT = rs("HWFZOSPT") Else pSpWFSamp.HWFZOSPT = " "  ' 測定点数(AO)
    If IsNull(rs("HWFZOSPI")) = False Then pSpWFSamp.HWFZOSPI = rs("HWFZOSPI") Else pSpWFSamp.HWFZOSPI = " "  ' 測定位置(AO)
    If IsNull(rs("HWFZONSW")) = False Then pSpWFSamp.HWFZONSW = rs("HWFZONSW") Else pSpWFSamp.HWFZONSW = " "  ' 熱処理法(AO)

    rs.Close
    ''残存酸素仕様取得追加　03/12/15 ooba END ==================================================>

    '' GD仕様取得　05/02/18 ooba START ========================================================>
''Upd start (TCS)T.Terauchi 2005/10/05  抜試指示4.5ﾗｲﾝ対応
    sSQL = "select "
    sSQL = sSQL & "T1.HWFGDSPH AS HWFGDSPH, "         '測定方法(GD)　05/10/25 ooba
    sSQL = sSQL & "T1.HWFGDSPT AS HWFGDSPT, "         '測定点数(GD)　05/10/25 ooba
    sSQL = sSQL & "T1.HWFGDZAR AS HWFGDZAR, "         '除外領域(GD)　05/10/25 ooba
    sSQL = sSQL & "T1.HWFDENHS AS HWFDENHS, "         '処理方法(GD/DEN)
    sSQL = sSQL & "T1.HWFLDLHS AS HWFLDLHS, "         '処理方法(GD/LDL)
    sSQL = sSQL & "T1.HWFDVDHS AS HWFDVDHS"           '処理方法(GD/DVD2)
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    sSQL = sSQL & ",T1.HWFGDSZY AS HWFGDSZY"          '測定条件(GD)
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
    sSQL = sSQL & ",T2.HWFGDLINE AS HWFGDLINE "       'ﾗｲﾝ数
    sSQL = sSQL & "from TBCME026 T1,TBCME036 T2 "
    sSQL = sSQL & "where T1.HINBAN = '" & pSpWFSamp.HIN.hinban & "' "
    sSQL = sSQL & "and T1.MNOREVNO = " & pSpWFSamp.HIN.mnorevno & " "
    sSQL = sSQL & "and T1.FACTORY = '" & pSpWFSamp.HIN.factory & "' "
    sSQL = sSQL & "and T1.OPECOND = '" & pSpWFSamp.HIN.opecond & "' "
    sSQL = sSQL & "and T1.HINBAN = T2.HINBAN "
    sSQL = sSQL & "and T1.MNOREVNO = T2.MNOREVNO "
    sSQL = sSQL & "and T1.FACTORY = T2.FACTORY "
    sSQL = sSQL & "and T1.OPECOND = T2.OPECOND "
''Upd end   (TCS)T.Terauchi 2005/10/05  抜試指示4.5ﾗｲﾝ対応

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    If IsNull(rs("HWFGDSPH")) = False Then pSpWFSamp.HWFGDSPH = rs("HWFGDSPH") Else pSpWFSamp.HWFGDSPH = " "  '05/10/25 ooba
    If IsNull(rs("HWFGDSPT")) = False Then pSpWFSamp.HWFGDSPT = rs("HWFGDSPT") Else pSpWFSamp.HWFGDSPT = " "  '05/10/25 ooba
    If IsNull(rs("HWFGDZAR")) = False Then pSpWFSamp.HWFGDZAR = rs("HWFGDZAR") Else pSpWFSamp.HWFGDZAR = " "  '05/10/25 ooba
    If IsNull(rs("HWFDENHS")) = False Then pSpWFSamp.HWFDENHS = rs("HWFDENHS") Else pSpWFSamp.HWFDENHS = " "  '処理方法(GD/DEN)
    If IsNull(rs("HWFLDLHS")) = False Then pSpWFSamp.HWFLDLHS = rs("HWFLDLHS") Else pSpWFSamp.HWFLDLHS = " "  '処理方法(GD/LDL)
    If IsNull(rs("HWFDVDHS")) = False Then pSpWFSamp.HWFDVDHS = rs("HWFDVDHS") Else pSpWFSamp.HWFDVDHS = " "  '処理方法(GD/DVD2)
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    If IsNull(rs("HWFGDSZY")) = False Then pSpWFSamp.HWFGDSZY = rs("HWFGDSZY") Else pSpWFSamp.HWFGDSZY = " "
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
''Upd Start (TCS)T.Terauchi 2005/10/05  抜試指示4.5ﾗｲﾝ対応
    If IsNull(rs("HWFGDLINE")) = False Then pSpWFSamp.HWFGDLINE = CStr(rs("HWFGDLINE"))
''Upd End   (TCS)T.Terauchi 2005/10/05  抜試指示4.5ﾗｲﾝ対応

    rs.Close
    '' GD仕様取得　05/02/18 ooba END ==========================================================>

    '' SPV仕様取得　06/06/08 ooba START ===============================================>
    sSQL = "select HWFNRHS, "                    '品WFSPVNR保証方法_処
    sSQL = sSQL & "HWFNRSH, "                     '品WFSPVNR測定位置_方
    sSQL = sSQL & "HWFNRST, "                     '品WFSPVNR測定位置_点
    sSQL = sSQL & "HWFNRSI, "                     '品WFSPVNR測定位置_位
    sSQL = sSQL & "HWFNRKN, "                     '品WFSPVNR検査頻度_抜
    sSQL = sSQL & "HWFSPVPUG, "                   '品WFSPVFEPUA限
    sSQL = sSQL & "HWFSPVPUR, "                   '品WFSPVFEPUA率
    sSQL = sSQL & "HWFSPVSTD, "                   '品WFSPVFE標準偏差
    sSQL = sSQL & "HWFDLPUG, "                    '品WF拡散長PUA限
    sSQL = sSQL & "HWFDLPUR, "                    '品WF拡散長PUA率
    sSQL = sSQL & "HWFNRPUG, "                    '品WFSPVNRPUA限
    sSQL = sSQL & "HWFNRPUR, "                    '品WFSPVNRPUA率
    sSQL = sSQL & "HWFNRSTD "                     '品WFSPVNR標準偏差
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
    sSQL = sSQL & ",HWFSIRDMX, "                  '軸状転位上限
    sSQL = sSQL & "HWFSIRDSZ, "                   '軸状転位測定条件
    sSQL = sSQL & "HWFSIRDHT, "                   '軸状転位保証方法＿対
    sSQL = sSQL & "HWFSIRDHS, "                   '軸状転位保証方法_処
    sSQL = sSQL & "HWFSIRDKM, "                   '軸状転位検査頻度＿枚
    sSQL = sSQL & "HWFSIRDKN, "                   '軸状転位検査頻度_抜
    sSQL = sSQL & "HWFSIRDKH, "                   '軸状転位検査頻度＿保
    sSQL = sSQL & "HWFSIRDKU, "                   '軸状転位検査頻度＿ウ
    sSQL = sSQL & "HWFSIRDPS  "                   '軸状転位TB保証位置
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
    sSQL = sSQL & "from TBCME048 "
    sSQL = sSQL & "where HINBAN = '" & pSpWFSamp.HIN.hinban & "' "
    sSQL = sSQL & "and MNOREVNO = " & pSpWFSamp.HIN.mnorevno & " "
    sSQL = sSQL & "and FACTORY = '" & pSpWFSamp.HIN.factory & "' "
    sSQL = sSQL & "and OPECOND = '" & pSpWFSamp.HIN.opecond & "' "

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    If IsNull(rs("HWFNRHS")) = False Then pSpWFSamp.HWFNRHS = rs("HWFNRHS") Else pSpWFSamp.HWFNRHS = " "
    If IsNull(rs("HWFNRSH")) = False Then pSpWFSamp.HWFNRSH = rs("HWFNRSH") Else pSpWFSamp.HWFNRSH = " "
    If IsNull(rs("HWFNRST")) = False Then pSpWFSamp.HWFNRST = rs("HWFNRST") Else pSpWFSamp.HWFNRST = " "
    If IsNull(rs("HWFNRSI")) = False Then pSpWFSamp.HWFNRSI = rs("HWFNRSI") Else pSpWFSamp.HWFNRSI = " "
    If IsNull(rs("HWFNRKN")) = False Then pSpWFSamp.HWFNRKN = rs("HWFNRKN") Else pSpWFSamp.HWFNRKN = " "
    If IsNull(rs("HWFSPVPUG")) = False Then pSpWFSamp.HWFSPVPUG = rs("HWFSPVPUG") Else pSpWFSamp.HWFSPVPUG = " "
    If IsNull(rs("HWFSPVPUR")) = False Then pSpWFSamp.HWFSPVPUR = rs("HWFSPVPUR") Else pSpWFSamp.HWFSPVPUR = " "
    If IsNull(rs("HWFSPVSTD")) = False Then pSpWFSamp.HWFSPVSTD = rs("HWFSPVSTD") Else pSpWFSamp.HWFSPVSTD = " "
    If IsNull(rs("HWFDLPUG")) = False Then pSpWFSamp.HWFDLPUG = rs("HWFDLPUG") Else pSpWFSamp.HWFDLPUG = " "
    If IsNull(rs("HWFDLPUR")) = False Then pSpWFSamp.HWFDLPUR = rs("HWFDLPUR") Else pSpWFSamp.HWFDLPUR = " "
    If IsNull(rs("HWFNRPUG")) = False Then pSpWFSamp.HWFNRPUG = rs("HWFNRPUG") Else pSpWFSamp.HWFNRPUG = " "
    If IsNull(rs("HWFNRPUR")) = False Then pSpWFSamp.HWFNRPUR = rs("HWFNRPUR") Else pSpWFSamp.HWFNRPUR = " "
    If IsNull(rs("HWFNRSTD")) = False Then pSpWFSamp.HWFNRSTD = rs("HWFNRSTD") Else pSpWFSamp.HWFNRSTD = " "
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
    If IsNull(rs("HWFSIRDMX")) = False Then pSpWFSamp.HWFSIRDMX = rs("HWFSIRDMX") Else pSpWFSamp.HWFSIRDMX = "0"    '軸状転位上限
    If IsNull(rs("HWFSIRDSZ")) = False Then pSpWFSamp.HWFSIRDSZ = rs("HWFSIRDSZ") Else pSpWFSamp.HWFSIRDSZ = " "    '軸状転位測定条件
    If IsNull(rs("HWFSIRDHT")) = False Then pSpWFSamp.HWFSIRDHT = rs("HWFSIRDHT") Else pSpWFSamp.HWFSIRDHT = " "    '軸状転位保証方法＿対
    If IsNull(rs("HWFSIRDHS")) = False Then pSpWFSamp.HWFSIRDHS = rs("HWFSIRDHS") Else pSpWFSamp.HWFSIRDHS = " "    '軸状転位保証方法＿処
    If IsNull(rs("HWFSIRDKM")) = False Then pSpWFSamp.HWFSIRDKM = rs("HWFSIRDKM") Else pSpWFSamp.HWFSIRDKM = " "    '軸状転位検査頻度＿枚
    If IsNull(rs("HWFSIRDKN")) = False Then pSpWFSamp.HWFSIRDKN = rs("HWFSIRDKN") Else pSpWFSamp.HWFSIRDKN = " "    '軸状転位検査頻度＿抜
    If IsNull(rs("HWFSIRDKH")) = False Then pSpWFSamp.HWFSIRDKH = rs("HWFSIRDKH") Else pSpWFSamp.HWFSIRDKH = " "    '軸状転位検査頻度＿保
    If IsNull(rs("HWFSIRDKU")) = False Then pSpWFSamp.HWFSIRDKU = rs("HWFSIRDKU") Else pSpWFSamp.HWFSIRDKU = " "    '軸状転位検査頻度＿ウ
    If IsNull(rs("HWFSIRDPS")) = False Then pSpWFSamp.HWFSIRDPS = Trim(rs("HWFSIRDPS")) Else pSpWFSamp.HWFSIRDPS = " "    '軸状転位TB保証位置
    
    '「軸状転位TB保証位置」を判定し、「軸状転位検査頻度＿抜」に編集（仮対応）
    Select Case Trim(pSpWFSamp.HWFSIRDPS)
    Case "T"
        pSpWFSamp.HWFSIRDKN = "3"
    Case "B"
        pSpWFSamp.HWFSIRDKN = "4"
    Case "TB"
        pSpWFSamp.HWFSIRDKN = "6"
    Case Else
        pSpWFSamp.HWFSIRDKN = " "
    End Select
    
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)

    rs.Close
    '' SPV仕様取得　06/06/08 ooba END =================================================>

    '' 製品仕様管理の取得
    sSQL = "select HWFIGKBN from TBCME017" & _
          " where HINBAN='" & pSpWFSamp.HIN.hinban & "' and MNOREVNO=" & pSpWFSamp.HIN.mnorevno & _
          " and FACTORY='" & pSpWFSamp.HIN.factory & "'"
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    pSpWFSamp.HWFIGKBN = rs("HWFIGKBN")
    rs.Close

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    '' エピ仕様取得(BMD1E～BMD3E,OSF1E～OSF3E)
    sSQL = "select HEPOF1NS, "                   ' 品熱処理法(OSF1E)
    sSQL = sSQL & "HEPOF1SZ, "                    ' 品測定条件(OSF1E)
    sSQL = sSQL & "HEPOF1ET, "                    ' 品選択ET代(OSF1E)
    sSQL = sSQL & "HEPOF1HS, "                    ' 品保証方法_処(OSF1E)
    sSQL = sSQL & "HEPOF1SH, "                    ' 品測定位置_方(OSF1E)
    sSQL = sSQL & "HEPOF1ST, "                    ' 品測定位置_点(OSF1E)
    sSQL = sSQL & "HEPOF1SR, "                    ' 品測定位置_領(OSF1E)
    sSQL = sSQL & "HEPOF2NS, "                    ' 品熱処理法(OSF2E)
    sSQL = sSQL & "HEPOF2SZ, "                    ' 品測定条件(OSF2E)
    sSQL = sSQL & "HEPOF2ET, "                    ' 品選択ET代(OSF2E)
    sSQL = sSQL & "HEPOF2HS, "                    ' 品保証方法_処(OSF2E)
    sSQL = sSQL & "HEPOF2SH, "                    ' 品測定位置_方(OSF2E)
    sSQL = sSQL & "HEPOF2ST, "                    ' 品測定位置_点(OSF2E)
    sSQL = sSQL & "HEPOF2SR, "                    ' 品測定位置_領(OSF2E)
    sSQL = sSQL & "HEPOF3NS, "                    ' 品熱処理法(OSF3E)
    sSQL = sSQL & "HEPOF3SZ, "                    ' 品測定条件(OSF3E)
    sSQL = sSQL & "HEPOF3ET, "                    ' 品選択ET代(OSF3E)
    sSQL = sSQL & "HEPOF3HS, "                    ' 品保証方法_処(OSF3E)
    sSQL = sSQL & "HEPOF3SH, "                    ' 品測定位置_方(OSF3E)
    sSQL = sSQL & "HEPOF3ST, "                    ' 品測定位置_点(OSF3E)
    sSQL = sSQL & "HEPOF3SR, "                    ' 品測定位置_領(OSF3E)
    sSQL = sSQL & "HEPBM1NS, "                    ' 品熱処理法(BMD1E)
    sSQL = sSQL & "HEPBM1SZ, "                    ' 品測定条件(BMD1E)
    sSQL = sSQL & "HEPBM1ET, "                    ' 品選択ET代(BMD1E)
    sSQL = sSQL & "HEPBM1HS, "                    ' 品保証方法_処(BMD1E)
    sSQL = sSQL & "HEPBM1SH, "                    ' 品測定位置_方(BMD1E)
    sSQL = sSQL & "HEPBM1ST, "                    ' 品測定位置_点(BMD1E)
    sSQL = sSQL & "HEPBM1SR, "                    ' 品測定位置_領(BMD1E)
    sSQL = sSQL & "HEPBM2NS, "                    ' 品熱処理法(BMD2E)
    sSQL = sSQL & "HEPBM2SZ, "                    ' 品測定条件(BMD2E)
    sSQL = sSQL & "HEPBM2ET, "                    ' 品選択ET代(BMD2E)
    sSQL = sSQL & "HEPBM2HS, "                    ' 品保証方法_処(BMD1E)
    sSQL = sSQL & "HEPBM2SH, "                    ' 品測定位置_方(BMD2E)
    sSQL = sSQL & "HEPBM2ST, "                    ' 品測定位置_点(BMD2E)
    sSQL = sSQL & "HEPBM2SR, "                    ' 品測定位置_領(BMD2E)
    sSQL = sSQL & "HEPBM3NS, "                    ' 品熱処理法(BMD3E)
    sSQL = sSQL & "HEPBM3SZ, "                    ' 品測定条件(BMD3E)
    sSQL = sSQL & "HEPBM3ET, "                    ' 品選択ET代(BMD3E)
    sSQL = sSQL & "HEPBM3HS, "                    ' 品保証方法_処(BMD1E)
    sSQL = sSQL & "HEPBM3SH, "                    ' 品測定位置_方(BMD3E)
    sSQL = sSQL & "HEPBM3ST, "                    ' 品測定位置_点(BMD3E)
    sSQL = sSQL & "HEPBM3SR, "                    ' 品測定位置_領(BMD3E)
    sSQL = sSQL & "HEPACEN, "                     ' 品E1厚中心
    sSQL = sSQL & "HEPANTNP, "                    ' 品EPAN温度
    sSQL = sSQL & "HEPANTIM, "                    ' 品EPAN時間
    sSQL = sSQL & "HEPIGKBN, "                    ' 品EPIG区分
    sSQL = sSQL & "HEPANGZY "                     ' 品EP高温ANガス条件
    sSQL = sSQL & "from TBCME050 "
    sSQL = sSQL & "where HINBAN = '" & pSpWFSamp.HIN.hinban & "' "
    sSQL = sSQL & "and MNOREVNO = " & pSpWFSamp.HIN.mnorevno & " "
    sSQL = sSQL & "and FACTORY = '" & pSpWFSamp.HIN.factory & "' "
    sSQL = sSQL & "and OPECOND = '" & pSpWFSamp.HIN.opecond & "' "

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    If IsNull(rs("HEPOF1NS")) = False Then pSpWFSamp.HEPOF1NS = rs("HEPOF1NS") Else pSpWFSamp.HEPOF1NS = " "
    If IsNull(rs("HEPOF1SZ")) = False Then pSpWFSamp.HEPOF1SZ = rs("HEPOF1SZ") Else pSpWFSamp.HEPOF1SZ = " "
    pSpWFSamp.HEPOF1ET = fncNullCheck(rs("HEPOF1ET"))
    If IsNull(rs("HEPOF1HS")) = False Then pSpWFSamp.HEPOF1HS = rs("HEPOF1HS") Else pSpWFSamp.HEPOF1HS = " "
    If IsNull(rs("HEPOF1SH")) = False Then pSpWFSamp.HEPOF1SH = rs("HEPOF1SH") Else pSpWFSamp.HEPOF1SH = " "
    If IsNull(rs("HEPOF1ST")) = False Then pSpWFSamp.HEPOF1ST = rs("HEPOF1ST") Else pSpWFSamp.HEPOF1ST = " "
    If IsNull(rs("HEPOF1SR")) = False Then pSpWFSamp.HEPOF1SR = rs("HEPOF1SR") Else pSpWFSamp.HEPOF1SR = " "
    If IsNull(rs("HEPOF2NS")) = False Then pSpWFSamp.HEPOF2NS = rs("HEPOF2NS") Else pSpWFSamp.HEPOF2NS = " "
    If IsNull(rs("HEPOF2SZ")) = False Then pSpWFSamp.HEPOF2SZ = rs("HEPOF2SZ") Else pSpWFSamp.HEPOF2SZ = " "
    pSpWFSamp.HEPOF2ET = fncNullCheck(rs("HEPOF2ET"))
    If IsNull(rs("HEPOF2HS")) = False Then pSpWFSamp.HEPOF2HS = rs("HEPOF2HS") Else pSpWFSamp.HEPOF2HS = " "
    If IsNull(rs("HEPOF2SH")) = False Then pSpWFSamp.HEPOF2SH = rs("HEPOF2SH") Else pSpWFSamp.HEPOF2SH = " "
    If IsNull(rs("HEPOF2ST")) = False Then pSpWFSamp.HEPOF2ST = rs("HEPOF2ST") Else pSpWFSamp.HEPOF2ST = " "
    If IsNull(rs("HEPOF2SR")) = False Then pSpWFSamp.HEPOF2SR = rs("HEPOF2SR") Else pSpWFSamp.HEPOF2SR = " "
    If IsNull(rs("HEPOF3NS")) = False Then pSpWFSamp.HEPOF3NS = rs("HEPOF3NS") Else pSpWFSamp.HEPOF3NS = " "
    If IsNull(rs("HEPOF3SZ")) = False Then pSpWFSamp.HEPOF3SZ = rs("HEPOF3SZ") Else pSpWFSamp.HEPOF3SZ = " "
    pSpWFSamp.HEPOF3ET = fncNullCheck(rs("HEPOF3ET"))
    If IsNull(rs("HEPOF3HS")) = False Then pSpWFSamp.HEPOF3HS = rs("HEPOF3HS") Else pSpWFSamp.HEPOF3HS = " "
    If IsNull(rs("HEPOF3SH")) = False Then pSpWFSamp.HEPOF3SH = rs("HEPOF3SH") Else pSpWFSamp.HEPOF3SH = " "
    If IsNull(rs("HEPOF3ST")) = False Then pSpWFSamp.HEPOF3ST = rs("HEPOF3ST") Else pSpWFSamp.HEPOF3ST = " "
    If IsNull(rs("HEPOF3SR")) = False Then pSpWFSamp.HEPOF3SR = rs("HEPOF3SR") Else pSpWFSamp.HEPOF3SR = " "
    If IsNull(rs("HEPBM1NS")) = False Then pSpWFSamp.HEPBM1NS = rs("HEPBM1NS") Else pSpWFSamp.HEPBM1NS = " "
    If IsNull(rs("HEPBM1SZ")) = False Then pSpWFSamp.HEPBM1SZ = rs("HEPBM1SZ") Else pSpWFSamp.HEPBM1SZ = " "
    pSpWFSamp.HEPBM1ET = fncNullCheck(rs("HEPBM1ET"))
    If IsNull(rs("HEPBM1HS")) = False Then pSpWFSamp.HEPBM1HS = rs("HEPBM1HS") Else pSpWFSamp.HEPBM1HS = " "
    If IsNull(rs("HEPBM1SH")) = False Then pSpWFSamp.HEPBM1SH = rs("HEPBM1SH") Else pSpWFSamp.HEPBM1SH = " "
    If IsNull(rs("HEPBM1ST")) = False Then pSpWFSamp.HEPBM1ST = rs("HEPBM1ST") Else pSpWFSamp.HEPBM1ST = " "
    If IsNull(rs("HEPBM1SR")) = False Then pSpWFSamp.HEPBM1SR = rs("HEPBM1SR") Else pSpWFSamp.HEPBM1SR = " "
    If IsNull(rs("HEPBM2NS")) = False Then pSpWFSamp.HEPBM2NS = rs("HEPBM2NS") Else pSpWFSamp.HEPBM2NS = " "
    If IsNull(rs("HEPBM2SZ")) = False Then pSpWFSamp.HEPBM2SZ = rs("HEPBM2SZ") Else pSpWFSamp.HEPBM2SZ = " "
    pSpWFSamp.HEPBM2ET = fncNullCheck(rs("HEPBM2ET"))
    If IsNull(rs("HEPBM2HS")) = False Then pSpWFSamp.HEPBM2HS = rs("HEPBM2HS") Else pSpWFSamp.HEPBM2HS = " "
    If IsNull(rs("HEPBM2SH")) = False Then pSpWFSamp.HEPBM2SH = rs("HEPBM2SH") Else pSpWFSamp.HEPBM2SH = " "
    If IsNull(rs("HEPBM2ST")) = False Then pSpWFSamp.HEPBM2ST = rs("HEPBM2ST") Else pSpWFSamp.HEPBM2ST = " "
    If IsNull(rs("HEPBM2SR")) = False Then pSpWFSamp.HEPBM2SR = rs("HEPBM2SR") Else pSpWFSamp.HEPBM2SR = " "
    If IsNull(rs("HEPBM3NS")) = False Then pSpWFSamp.HEPBM3NS = rs("HEPBM3NS") Else pSpWFSamp.HEPBM3NS = " "
    If IsNull(rs("HEPBM3SZ")) = False Then pSpWFSamp.HEPBM3SZ = rs("HEPBM3SZ") Else pSpWFSamp.HEPBM3SZ = " "
    pSpWFSamp.HEPBM3ET = fncNullCheck(rs("HEPBM3ET"))
    If IsNull(rs("HEPBM3HS")) = False Then pSpWFSamp.HEPBM3HS = rs("HEPBM3HS") Else pSpWFSamp.HEPBM3HS = " "
    If IsNull(rs("HEPBM3SH")) = False Then pSpWFSamp.HEPBM3SH = rs("HEPBM3SH") Else pSpWFSamp.HEPBM3SH = " "
    If IsNull(rs("HEPBM3ST")) = False Then pSpWFSamp.HEPBM3ST = rs("HEPBM3ST") Else pSpWFSamp.HEPBM3ST = " "
    If IsNull(rs("HEPBM3SR")) = False Then pSpWFSamp.HEPBM3SR = rs("HEPBM3SR") Else pSpWFSamp.HEPBM3SR = " "
    pSpWFSamp.HEPACEN = fncNullCheck(rs("HEPACEN"))
    pSpWFSamp.HEPANTNP = fncNullCheck(rs("HEPANTNP"))
    pSpWFSamp.HEPANTIM = fncNullCheck(rs("HEPANTIM"))
    If IsNull(rs("HEPIGKBN")) = False Then pSpWFSamp.HEPIGKBN = rs("HEPIGKBN") Else pSpWFSamp.HEPIGKBN = " "
    If IsNull(rs("HEPANGZY")) = False Then pSpWFSamp.HEPANGZY = rs("HEPANGZY") Else pSpWFSamp.HEPANGZY = " "
    rs.Close
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

    scmzc_getWF = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    scmzc_getWF = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'***********************************************************************************
'*    関数名        : MakeParameter
'*
'*    処理概要      : 1.新DB書込み処理
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    intFormID     ,I  ,Integer  ,（1:WFセンタ総合判定　2:再抜試
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'***********************************************************************************
Public Function MakeParameter(ByVal intFormID As Integer) As FUNCTION_RETURN
    Dim sErrTbl             As String
    Dim vBeforeBlock        As Variant  '現在行のブロックID
    Dim vAfterBlock         As Variant  '次の行のブロックID
    Dim intSprCnt           As Integer  'スプレッドループカウント
    Dim sErrMsg             As String
    Dim vIngotpos           As Variant
    Dim lngBeginIngotpos    As Long
    Dim lngEndIngotpos      As Long
    Dim vBeginSeq           As Variant
    Dim lngWfBeginSeq       As Long
    Dim lngWfEndSeq         As Long
    Dim sCryNum             As String
    Dim sSXLID              As String

    If intFormID = 1 Then 'WFセンター総合判定実行
        '構造体作成
        If cmbc039_2_CreateTable(sErrMsg) = FUNCTION_RETURN_FAILURE Then
            MakeParameter = FUNCTION_RETURN_FAILURE
            f_cmbc039_2.lblMsg.Caption = sErrMsg
            Exit Function
        End If
    ElseIf intFormID = 2 Then '再抜試指示実行
        sSXLID = Trim(f_cmbc039_3.txtKSXLID.text)
        '品番を1列追加したことによる列の変更-------start iida 2003/09/06
        With f_cmbc039_3.sprExamine
            lngBeginIngotpos = SIngotP  '2003/04/22 okazaki
            lngEndIngotpos = EIngotP  '2003/05/01 hitec)matsumoto
            .GetText 6, 1, vBeginSeq
            '既存・新規ブロック位置の修正 2003/04/22
            lngWfBeginSeq = CInt(Trim(vBeginSeq))
            .GetText 6, .MaxRows, vBeginSeq
            lngWfEndSeq = CInt(Trim(vBeginSeq))
        End With

        '構造体作成
        intSprCnt = 0
        'テーブル展開処理
        If cmbc039_3_CreateTable(sSXLID, lngBeginIngotpos, lngEndIngotpos, lngWfBeginSeq, lngWfEndSeq, sErrMsg) = FUNCTION_RETURN_FAILURE Then 'upd 2003/03/29 hitec)matsumoto lngWfBeginSeq,lngWfEndSeq追加
            MakeParameter = FUNCTION_RETURN_FAILURE
            f_cmbc039_3.lblMsg.Caption = sErrMsg
            Exit Function
        End If
    End If
    MakeParameter = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

End Function

'***********************************************************************************************
'*    関数名        : cmbc039_2_CreateXSDC2
'*
'*    処理概要      : 1.分割結晶（ブロック）前工程実績取得＆構造体作成
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    intBlockCnt   ,I  ,Integer  ,ブロック数
'*                    bNoData　　   ,I  ,Boolean  ,データ有無フラグ
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'***********************************************************************************************
Public Function cmbc039_2_CreateXSDC2(ByVal intBlockCnt As Integer, ByRef bNoData As Boolean) _
                                        As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sSQL        As String
    Dim intProcNo   As Integer
    Dim dblDiameter As Double
    Dim intNum      As Integer

    '' エラーハンドラの設定
    On Error GoTo proc_err

    intProcNo = 0
    bNoData = False

    'ブロックIDを得る
    sSQL = "SELECT * from XSDC2 "
    sSQL = sSQL & " WHERE CRYNUMC2='" & strBlockID(intBlockCnt) & "'"
    sSQL = sSQL & "   AND LIVKC2= '0'"   '生死区分

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        bNoData = True
        cmbc039_2_CreateXSDC2 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    rs.MoveFirst
    If rs.EOF = False Then
        With BlkOld
            If IsNull(rs.Fields("CRYNUMC2")) = False Then .CRYNUMC2 = rs.Fields("CRYNUMC2")
            If IsNull(rs.Fields("KCNTC2")) = False Then .KCNTC2 = rs.Fields("KCNTC2")       '工程連番
            If IsNull(rs.Fields("XTALC2")) = False Then .XTALC2 = rs.Fields("XTALC2")
            If IsNull(rs.Fields("INPOSC2")) = False Then .INPOSC2 = rs.Fields("INPOSC2")
            If IsNull(rs.Fields("NEKKNTC2")) = False Then .NEKKNTC2 = rs.Fields("NEKKNTC2")
            If IsNull(rs.Fields("NEWKNTC2")) = False Then .NEWKNTC2 = rs.Fields("NEWKNTC2")
            If IsNull(rs.Fields("NEWKKBC2")) = False Then .NEWKKBC2 = rs.Fields("NEWKKBC2")
            If IsNull(rs.Fields("NEMACOC2")) = False Then .NEMACOC2 = rs.Fields("NEMACOC2")
            If IsNull(rs.Fields("GNKKNTC2")) = False Then .GNKKNTC2 = rs.Fields("GNKKNTC2")
            If IsNull(rs.Fields("GNWKNTC2")) = False Then .GNWKNTC2 = rs.Fields("GNWKNTC2")
            If IsNull(rs.Fields("GNWKKBC2")) = False Then .GNWKKBC2 = rs.Fields("GNWKKBC2")
            If IsNull(rs.Fields("GNMACOC2")) = False Then .GNMACOC2 = rs.Fields("GNMACOC2")
            If IsNull(rs.Fields("GNDAYC2")) = False Then .GNDAYC2 = rs.Fields("GNDAYC2")
            If IsNull(rs.Fields("GNLC2")) = False Then .GNLC2 = rs.Fields("GNLC2")          '現在長さ（前工程長さ）
            If IsNull(rs.Fields("GNWC2")) = False Then .GNWC2 = rs.Fields("GNWC2")          '現在重量
            If IsNull(rs.Fields("GNMC2")) = False Then .GNMC2 = rs.Fields("GNMC2")          '現在枚数（前工程枚数）
            If IsNull(rs.Fields("SUMITLC2")) = False Then .SUMITLC2 = rs.Fields("SUMITLC2")
            If IsNull(rs.Fields("SUMITWC2")) = False Then .SUMITWC2 = rs.Fields("SUMITWC2")
            If IsNull(rs.Fields("SUMITMC2")) = False Then .SUMITMC2 = rs.Fields("SUMITMC2")
            If IsNull(rs.Fields("CHGC2")) = False Then .CHGC2 = rs.Fields("CHGC2")
            If IsNull(rs.Fields("KAKOUBC2")) = False Then .KAKOUBC2 = rs.Fields("KAKOUBC2")
            If IsNull(rs.Fields("KEIDAYC2")) = False Then .KEIDAYC2 = rs.Fields("KEIDAYC2")
            If IsNull(rs.Fields("GNTKUBC2")) = False Then .GNTKUBC2 = rs.Fields("GNTKUBC2")
            If IsNull(rs.Fields("GNTNOC2")) = False Then .GNTNOC2 = rs.Fields("GNTNOC2")
            If IsNull(rs.Fields("XTWORKC2")) = False Then .XTWORKC2 = rs.Fields("XTWORKC2")
            If IsNull(rs.Fields("WFWORKC2")) = False Then .WFWORKC2 = rs.Fields("WFWORKC2")
            If IsNull(rs.Fields("LSTATBC2")) = False Then .LSTATBC2 = rs.Fields("LSTATBC2")
            If IsNull(rs.Fields("RSTATBC2")) = False Then .RSTATBC2 = rs.Fields("RSTATBC2")
            If IsNull(rs.Fields("LUFRCC2")) = False Then .LUFRCC2 = rs.Fields("LUFRCC2")
            If IsNull(rs.Fields("LUFRBC2")) = False Then .LUFRBC2 = rs.Fields("LUFRBC2")
            If IsNull(rs.Fields("LDFRCC2")) = False Then .LDFRCC2 = rs.Fields("LDFRCC2")
            If IsNull(rs.Fields("LDFRBC2")) = False Then .LDFRBC2 = rs.Fields("LDFRBC2")
            If IsNull(rs.Fields("HOLDCC2")) = False Then .HOLDCC2 = rs.Fields("HOLDCC2")
            If IsNull(rs.Fields("HOLDBC2")) = False Then .HOLDBC2 = rs.Fields("HOLDBC2")
            If IsNull(rs.Fields("EXKUBC2")) = False Then .EXKUBC2 = rs.Fields("EXKUBC2")
            If IsNull(rs.Fields("HENPKC2")) = False Then .HENPKC2 = rs.Fields("HENPKC2")
            If IsNull(rs.Fields("LIVKC2")) = False Then .LIVKC2 = rs.Fields("LIVKC2")
            If IsNull(rs.Fields("KANKC2")) = False Then .KANKC2 = rs.Fields("KANKC2")
            If IsNull(rs.Fields("NFC2")) = False Then .NFC2 = rs.Fields("NFC2")
            If IsNull(rs.Fields("SAKJC2")) = False Then .SAKJC2 = rs.Fields("SAKJC2")
            If IsNull(rs.Fields("TDAYC2")) = False Then .TDAYC2 = rs.Fields("TDAYC2")
            If IsNull(rs.Fields("KDAYC2")) = False Then .KDAYC2 = rs.Fields("KDAYC2")
            If IsNull(rs.Fields("SUMITBC2")) = False Then .SUMITBC2 = rs.Fields("SUMITBC2")
            If IsNull(rs.Fields("SNDKC2")) = False Then .SNDKC2 = rs.Fields("SNDKC2")
            If IsNull(rs.Fields("SNDDAYC2")) = False Then .SNDDAYC2 = rs.Fields("SNDDAYC2")
            If IsNull(rs.Fields("PLANTCATC2")) = False Then .PLANTCATC2 = rs.Fields("PLANTCATC2")   ' 2007/09/04 SPK Tsutsumi Add
        End With

        '前工程の構造体を現在工程の構造体にコピー
        BlkNow = BlkOld

        '現在工程の工程連番を修正
        With BlkNow
            .KCNTC2 = CInt(.KCNTC2) + 1     '工程連番
            'Cng Start  2010/09/02 Y.Hitomi
            'ブロック内SXLが1つでも完了していた場合、工程コードを更新しないようにする。
            If (.GNWKNTC2 <> "     " Or _
                .GNWKNTC2 <> "CW800" Or _
                .GNWKNTC2 <> "TX860") Then
            
                .NEWKNTC2 = Kihon.NOWPROC       '前工程コードを最終通過工程にセット
                .GNWKNTC2 = Kihon.NEWPROC       '現在工程コードを現在工程へセット
            End If
            '  .NEWKNTC2 = Kihon.NOWPROC       '前工程コードを最終通過工程にセット
            '  .GNWKNTC2 = Kihon.NEWPROC       '現在工程コードを現在工程へセット
            'Cng End    2010/09/02 Y.Hitomi
            
            .SUMITBC2 = "0"
            .SUMITLC2 = "0"
            .SUMITMC2 = "0"
            .SUMITWC2 = "0"
            '現在重量を求める
            If GetDiameter(strBlockID(intBlockCnt), dblDiameter) = FUNCTION_RETURN_FAILURE Then  '直径を求める
                dblDiameter = 0
            End If
            '基本情報の直径をセット
            Kihon.DIAMETER = dblDiameter
        End With
    End If

    rs.Close
    cmbc039_2_CreateXSDC2 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    cmbc039_2_CreateXSDC2 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'********************************************************************************************
'*    関数名        : cmbc039_2_CreateXSDCA
'*
'*    処理概要      : 1.分割結晶（品番）前工程実績取得＆構造体作成
'*
'*    パラメータ    : 変数名        ,IO ,型                            ,説明
'*                    intBlockCnt   ,I  ,Integer  ,ブロック数
'*                    bNoData　　   ,I  ,Boolean  ,データ有無フラグ
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'********************************************************************************************
'Cng Start 2010/10/03 Y.Hitomi
Public Function cmbc039_2_CreateXSDCA(ByVal intBlockCnt As Integer, ByRef bNoData As Boolean, ByVal strSXLID As String) _
                                        As FUNCTION_RETURN
'Public Function cmbc039_2_CreateXSDCA(ByVal intBlockCnt As Integer, ByRef bNoData As Boolean) _
                                        As FUNCTION_RETURN
'Cng End 2010/10/03 Y.Hitomi
    
    Dim intLoopCnt  As Integer
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim intProcNo   As Integer
    Dim dblDiameter As Double
    Dim intNum      As Integer

    '' エラーハンドラの設定
    On Error GoTo proc_err

    intProcNo = 0
    bNoData = False

    'ブロックIDを得る
    sql = "SELECT * from XSDCA"
    sql = sql & " WHERE CRYNUMCA='" & strBlockID(intBlockCnt) & "'"
    sql = sql & "   AND LIVKCA= '0'"   '生死区分

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        bNoData = True
        cmbc039_2_CreateXSDCA = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    rs.MoveFirst
    intLoopCnt = 0

    Do While Not rs.EOF
        ReDim Preserve HinOld(intLoopCnt)
        ReDim Preserve HinNow(intLoopCnt)
        Kihon.CNTHINOLD = intLoopCnt + 1
        Kihon.CNTHINNOW = intLoopCnt + 1
        With HinOld(intLoopCnt)
            If IsNull(rs.Fields("CRYNUMCA")) = False Then .CRYNUMCA = rs.Fields("CRYNUMCA")
            If IsNull(rs.Fields("HINBCA")) = False Then .HINBCA = rs.Fields("HINBCA")
            If IsNull(rs.Fields("INPOSCA")) = False Then .INPOSCA = rs.Fields("INPOSCA")
            If IsNull(rs.Fields("REVNUMCA")) = False Then .REVNUMCA = rs.Fields("REVNUMCA")
            If IsNull(rs.Fields("FACTORYCA")) = False Then .FACTORYCA = rs.Fields("FACTORYCA")
            If IsNull(rs.Fields("OPECA")) = False Then .OPECA = rs.Fields("OPECA")
            If IsNull(rs.Fields("KCKNTCA")) = False Then .KCKNTCA = rs.Fields("KCKNTCA")
            If IsNull(rs.Fields("SXLIDCA")) = False Then .SXLIDCA = rs.Fields("SXLIDCA")
            If IsNull(rs.Fields("XTALCA")) = False Then .XTALCA = rs.Fields("XTALCA")
            If IsNull(rs.Fields("NEKKNTCA")) = False Then .NEKKNTCA = rs.Fields("NEKKNTCA")
            If IsNull(rs.Fields("NEWKNTCA")) = False Then .NEWKNTCA = rs.Fields("NEWKNTCA")
            If IsNull(rs.Fields("NEWKKBCA")) = False Then .NEWKKBCA = rs.Fields("NEWKKBCA")
            If IsNull(rs.Fields("NEMACOCA")) = False Then .NEMACOCA = rs.Fields("NEMACOCA")
            If IsNull(rs.Fields("GNKKNTCA")) = False Then .GNKKNTCA = rs.Fields("GNKKNTCA")
            If IsNull(rs.Fields("GNWKNTCA")) = False Then .GNWKNTCA = rs.Fields("GNWKNTCA")
            If IsNull(rs.Fields("GNWKKBCA")) = False Then .GNWKKBCA = rs.Fields("GNWKKBCA")
            If IsNull(rs.Fields("GNMACOCA")) = False Then .GNMACOCA = rs.Fields("GNMACOCA")
            If IsNull(rs.Fields("GNDAYCA")) = False Then .GNDAYCA = rs.Fields("GNDAYCA")
            If IsNull(rs.Fields("GNLCA")) = False Then .GNLCA = rs.Fields("GNLCA")
            If IsNull(rs.Fields("GNWCA")) = False Then .GNWCA = rs.Fields("GNWCA")
            If IsNull(rs.Fields("GNMCA")) = False Then .GNMCA = rs.Fields("GNMCA")
            If IsNull(rs.Fields("SUMITLCA")) = False Then .SUMITLCA = rs.Fields("SUMITLCA")
            If IsNull(rs.Fields("SUMITWCA")) = False Then .SUMITWCA = rs.Fields("SUMITWCA")
            If IsNull(rs.Fields("SUMITMCA")) = False Then .SUMITMCA = rs.Fields("SUMITMCA")
            If IsNull(rs.Fields("CHGCA")) = False Then .CHGCA = rs.Fields("CHGCA")
            If IsNull(rs.Fields("KAKOUBCA")) = False Then .KAKOUBCA = rs.Fields("KAKOUBCA")
            If IsNull(rs.Fields("KEIDAYCA")) = False Then .KEIDAYCA = rs.Fields("KEIDAYCA")
            If IsNull(rs.Fields("GNTKUBCA")) = False Then .GNTKUBCA = rs.Fields("GNTKUBCA")
            If IsNull(rs.Fields("GNTNOCA")) = False Then .GNTNOCA = rs.Fields("GNTNOCA")
            If IsNull(rs.Fields("XTWORKCA")) = False Then .XTWORKCA = rs.Fields("XTWORKCA")
            If IsNull(rs.Fields("WFWORKCA")) = False Then .WFWORKCA = rs.Fields("WFWORKCA")
            If IsNull(rs.Fields("LSTATBCA")) = False Then .LSTATBCA = rs.Fields("LSTATBCA")
            If IsNull(rs.Fields("RSTATBCA")) = False Then .RSTATBCA = rs.Fields("RSTATBCA")
            If IsNull(rs.Fields("LUFRCCA")) = False Then .LUFRCCA = rs.Fields("LUFRCCA")
            If IsNull(rs.Fields("LUFRBCA")) = False Then .LUFRBCA = rs.Fields("LUFRBCA")
            If IsNull(rs.Fields("LDFRCCA")) = False Then .LDFRCCA = rs.Fields("LDFRCCA")
            If IsNull(rs.Fields("LDFRBCA")) = False Then .LDFRBCA = rs.Fields("LDFRBCA")
            If IsNull(rs.Fields("HOLDCCA")) = False Then .HOLDCCA = rs.Fields("HOLDCCA")
            If IsNull(rs.Fields("HOLDBCA")) = False Then .HOLDBCA = rs.Fields("HOLDBCA")
            If IsNull(rs.Fields("EXKUBCA")) = False Then .EXKUBCA = rs.Fields("EXKUBCA")
            If IsNull(rs.Fields("HENPKCA")) = False Then .HENPKCA = rs.Fields("HENPKCA")
            If IsNull(rs.Fields("LIVKCA")) = False Then .LIVKCA = rs.Fields("LIVKCA")
            If IsNull(rs.Fields("KANKCA")) = False Then .KANKCA = rs.Fields("KANKCA")
            If IsNull(rs.Fields("NFCA")) = False Then .NFCA = rs.Fields("NFCA")
            If IsNull(rs.Fields("SAKJCA")) = False Then .SAKJCA = rs.Fields("SAKJCA")
            If IsNull(rs.Fields("TDAYCA")) = False Then .TDAYCA = rs.Fields("TDAYCA")
            If IsNull(rs.Fields("KDAYCA")) = False Then .KDAYCA = rs.Fields("KDAYCA")
            If IsNull(rs.Fields("SUMITBCA")) = False Then .SUMITBCA = rs.Fields("SUMITBCA")
            If IsNull(rs.Fields("SNDKCA")) = False Then .SNDKCA = rs.Fields("SNDKCA")
            If IsNull(rs.Fields("SNDDAYCA")) = False Then .SNDDAYCA = rs.Fields("SNDDAYCA")
            If IsNull(rs.Fields("PLANTCATCA")) = False Then .PLANTCATCA = rs.Fields("PLANTCATCA")   '2007/09/04 SPK Tsutsumi Add
        End With

        '前工程の構造体を現在工程の構造体へコピー
        HinNow(intLoopCnt) = HinOld(intLoopCnt)

        With HinNow(intLoopCnt)
            .KCKNTCA = CInt(.KCKNTCA) + 1
            'Cng Start 2010/10/03 Y.Hitomi
            '実行指示SXLIDのみ工程コードを変更し、それ以外は、前工程を引き継ぐ
            If strSXLID = .SXLIDCA Then
                .NEWKNTCA = Kihon.NOWPROC             '前工程コードを最終通過工程にセット
                .GNWKNTCA = Kihon.NEWPROC             '現在工程コードを現在工程へセット
            Else
                .NEWKNTCA = rs.Fields("NEWKNTCA")     '前工程コードを最終通過工程にセット
                .GNWKNTCA = rs.Fields("GNWKNTCA")     '現在工程コードを現在工程へセット
            End If
            '.NEWKNTCA = Kihon.NOWPROC             '前工程コードを最終通過工程にセット
            '.GNWKNTCA = Kihon.NEWPROC             '現在工程コードを現在工程へセット
            'Cng End   2010/10/03 Y.Hitomi
            .SUMITBCA = "0"
            .SUMITLCA = HinOld(intLoopCnt).SUMITLCA   ''03/05/13 後藤
            .SUMITMCA = HinOld(intLoopCnt).SUMITMCA   ''03/05/13 後藤
            .SUMITWCA = HinOld(intLoopCnt).SUMITWCA   ''03/05/13 後藤
            '現在重量を求める
            If GetDiameter(strBlockID(intBlockCnt), dblDiameter) = FUNCTION_RETURN_FAILURE Then  '直径を求める
                dblDiameter = 0
            End If
        End With

        intLoopCnt = intLoopCnt + 1
        rs.MoveNext
    Loop

    rs.Close
    cmbc039_2_CreateXSDCA = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    cmbc039_2_CreateXSDCA = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'*************************************************************************************************************
'*    関数名        : cmbc039_3_CreateTable
'*
'*    処理概要      : 1.構造体作成処理
'*
'*    パラメータ    : 変数名           ,IO ,型      ,説明
'*                    strSXLID         ,I  ,String  ,SXL-ID
'*                    lngBeginIngotpos ,I  ,Long    ,ブロック管理データの長さ
'*                    strErrMsg        ,O  ,String  ,ErrMsg格納
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*************************************************************************************************************
Public Function cmbc039_3_CreateTable(ByVal strSXLID As String, ByVal lngBeginIngotpos As Long, _
                                      ByVal lngEndIngotpos As Long, ByVal lngWfBeginSeq As Long, _
                                      ByVal lngWfEndSeq As Long, ByRef strErrMsg As String) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sErrTbl     As String
    Dim intSprCnt   As Integer
    Dim intRowCnt   As Integer
    Dim sDBName     As String
    Dim blNoData    As Boolean
    Dim sSQL        As String
    Dim intLoopCnt  As Integer

    blNoData = False

    giInpos = 9000  '在庫減、振替情報の位置を初期化

    'ブロック管理からブロックＩＤを取得
    sSQL = "SELECT DISTINCT(CRYNUMCA) "
    sSQL = sSQL & " FROM XSDCA"
    sSQL = sSQL & " WHERE CRYNUMCA like '" & left(strSXLID, 9) & "%'"   'ｲﾝﾃﾞｯｸｽ項目追加 09/05/25 ooba
    sSQL = sSQL & "   AND SXLIDCA = '" & strSXLID & "'"
    sSQL = sSQL & "   AND LIVKCA = '0'"   'add 2003/05/19 hitec)matsumoto

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    'ブロックIDを取得
    intLoopCnt = 0
    Do While Not rs.EOF
        ReDim Preserve strBlockID(intLoopCnt) As String
        If IsNull(rs("CRYNUMCA")) = True Then
            strBlockID(intLoopCnt) = ""
        Else
            strBlockID(intLoopCnt) = rs("CRYNUMCA")            'ブロックID
        End If
        '基本情報構造体
        With Kihon
            .STAFFID = Trim(f_cmbc039_3.txtStaffID.text)
            .NEWPROC = PROCD_WFC_SOUGOUHANTEI
            .NOWPROC = PROCD_WFC_SAINUKISI
            .DIAMETER = 0
        End With

        '分割結晶（ブロック）から前工程実績取得
        sDBName = "XSDC2"
        If cmbc039_3_CreateXSDC2(strBlockID(intLoopCnt), blNoData) = FUNCTION_RETURN_FAILURE Then
            If blNoData = True Then
                cmbc039_3_CreateTable = FUNCTION_RETURN_SUCCESS
                GoTo proc_exit
            Else
                cmbc039_3_CreateTable = FUNCTION_RETURN_FAILURE
                strErrMsg = GetMsgStr("EGET") & sDBName
                GoTo proc_exit
            End If
        End If

        '分割結晶（品番）から前工程実績取得
        sDBName = "XSDCA"
        If cmbc039_3_CreateXSDCA(strBlockID(intLoopCnt), blNoData) = FUNCTION_RETURN_FAILURE Then
            If blNoData = True Then
                cmbc039_3_CreateTable = FUNCTION_RETURN_SUCCESS
                GoTo proc_exit
            Else
                cmbc039_3_CreateTable = FUNCTION_RETURN_FAILURE
                strErrMsg = GetMsgStr("EGET") & sDBName
                GoTo proc_exit
            End If
        End If

        '現在工程実績作成
        sDBName = "XSDC2,XSDCA"
        strErrMsg = GetMsgStr("EAPLY") & sDBName
        
        'Cng Start 2010/10/03 Y.Hitomi
        'If cmbc039_3_CreateNowProc(strBlockID(intLoopCnt), lngBeginIngotpos, lngEndIngotpos, lngWfBeginSeq, lngWfEndSeq, strErrMsg) = FUNCTION_RETURN_FAILURE Then   'upd 2003/03/29 hitec)matsumoto 結晶内位置を使用していたが、マップ位置に変更
        If cmbc039_3_CreateNowProc(strBlockID(intLoopCnt), lngBeginIngotpos, lngEndIngotpos, lngWfBeginSeq, lngWfEndSeq, strErrMsg, strSXLID) = FUNCTION_RETURN_FAILURE Then  'upd 2003/03/29 hitec)matsumoto 結晶内位置を使用していたが、マップ位置に変更
        'Cng End  2010/10/03 Y.Hitomi
            cmbc039_3_CreateTable = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        strErrMsg = ""

        '基本処理
        If KihonProc = FUNCTION_RETURN_FAILURE Then
            cmbc039_3_CreateTable = FUNCTION_RETURN_FAILURE
            strErrMsg = GetMsgStr("EAPLY")
            GoTo proc_exit
        End If

        rs.MoveNext
        intLoopCnt = intLoopCnt + 1
    Loop

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

End Function

'**********************************************************************************
'*    関数名        : cmbc039_2_CreateTable
'*
'*    処理概要      : 1.構造体作成処理
'*
'*    パラメータ    : 変数名        ,IO ,型      ,説明
'*                    strErrMsg     ,O  ,String  ,ErrMsg格納
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'**********************************************************************************
Public Function cmbc039_2_CreateTable(ByRef strErrMsg As String) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rsMain      As OraDynaset
    Dim sErrTbl     As String
    Dim intBlockCnt As Integer
    Dim sDBName     As String
    Dim blNoData    As Boolean
    Dim sTmpSxl()   As String       '仕掛工程再ﾁｪｯｸ用SXLID　06/03/14 ooba
    Dim blKouteiChk As Boolean      '工程ﾁｪｯｸﾌﾗｸﾞ　06/03/14 ooba

    On Error GoTo proc_err

    blNoData = False
    'ブロックID取得
    sDBName = "XSDCA"
    'ｲﾝﾃﾞｯｸｽ項目(CRYNUMCA)追加 09/05/25 ooba
    sSQL = "select DISTINCT(CRYNUMCA) from XSDCA " & _
          "where CRYNUMCA like '" & left(SelectSxlID039, 9) & "%' " & _
          "  and SXLIDCA='" & SelectSxlID039 & "' " & _
          "  and LIVKCA= '0'"
    Set rsMain = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rsMain.RecordCount = 0 Then
        Debug.Print "XSDCA：前工程実績なし"
        Debug.Print sSQL
        rsMain.Close
        cmbc039_2_CreateTable = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If

    intBlockCnt = 0
    blNoData = False
    blKouteiChk = False      '06/03/14 ooba

    Do While Not rsMain.EOF
        intBlockCnt = intBlockCnt + 1
        ReDim Preserve strBlockID(intBlockCnt)
        strBlockID(intBlockCnt) = rsMain("CRYNUMCA")

        With Kihon
            .STAFFID = Trim(f_cmbc039_2.txtStaffID.text)
            .NOWPROC = PROCD_WFC_SOUGOUHANTEI
            .NEWPROC = PROCD_SXL_KAKUTEI
            .DIAMETER = 0       '--------------保留
            .ALLSCRAP = "N" '全数スクラップ
            .FURYOUMU = "N"   '不良無し
        End With

        '分割結晶（ブロック）から前工程実績取得
        sDBName = "XSDC2"
        If cmbc039_2_CreateXSDC2(intBlockCnt, blNoData) = FUNCTION_RETURN_FAILURE Then
            If blNoData = True Then
                rsMain.Close
                Debug.Print "cmbc039_2_CreateXSDC2(" & intBlockCnt & "," & blNoData & "):XSDC2前工程実績なし"
                cmbc039_2_CreateTable = FUNCTION_RETURN_SUCCESS
                Exit Function
            Else
                rsMain.Close
                Debug.Print "cmbc039_2_CreateXSDC2(" & intBlockCnt & "," & blNoData & "):XSDC2前工程実績読込みエラー"
                cmbc039_2_CreateTable = FUNCTION_RETURN_FAILURE
                strErrMsg = GetMsgStr("EGET") & sDBName
                Exit Function
            End If
        End If

        sDBName = "XSDCA"
        '分割結晶（品番）から前工程実績取得
'2010/10/03 Cng Start Y.Hitomi
'        If cmbc039_2_CreateXSDCA(intBlockCnt, blNoData) = FUNCTION_RETURN_FAILURE Then
        If cmbc039_2_CreateXSDCA(intBlockCnt, blNoData, SelectSxlID039) = FUNCTION_RETURN_FAILURE Then
'2010/10/03 Cng End Y.Hitomi
            If blNoData = True Then
                rsMain.Close
                Debug.Print "cmbc039_2_CreateXSDCA(" & intBlockCnt & "," & blNoData & "):XSDCA前工程実績なし"
                cmbc039_2_CreateTable = FUNCTION_RETURN_SUCCESS
                Exit Function
            Else
                rsMain.Close
                Debug.Print "cmbc039_2_CreateXSDCA(" & intBlockCnt & "," & blNoData & "):XSDCA前工程実績読込みエラー"
                cmbc039_2_CreateTable = FUNCTION_RETURN_FAILURE
                strErrMsg = GetMsgStr("EGET") & sDBName
                Exit Function
            End If
        End If

        '仕掛工程再チェック機能追加　06/03/14 ooba
        ReDim sTmpSxl(1)
        sTmpSxl(1) = SelectSxlID039
        If Not blKouteiChk Then
            If DBDRV_CheckCodeXSDCB(sTmpSxl, Kihon.NOWPROC, strErrMsg) = FUNCTION_RETURN_FAILURE Then
                cmbc039_2_CreateTable = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
            blKouteiChk = True
        End If

        '基本処理
        If KihonProc = FUNCTION_RETURN_FAILURE Then
            rsMain.Close
            Debug.Print "KihonProc()：基本処理異常終了"
            cmbc039_2_CreateTable = FUNCTION_RETURN_FAILURE
            strErrMsg = GetMsgStr("EAPLY")
            Exit Function
        End If

        rsMain.MoveNext
    Loop
    rsMain.Close
    cmbc039_2_CreateTable = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    cmbc039_2_CreateTable = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'*******************************************************************************************
'*    関数名        : cmbc039_3_CreateXSDCA
'*
'*    処理概要      : 1.分割結晶（品番）前工程実績取得＆構造体作成
'*
'*    パラメータ    : 変数名        ,IO ,型      ,説明
'*                    strBlockID    ,O  ,String  ,ブロックID
'*                    bNoData　　   ,I  ,Boolean ,データ有無フラグ
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************
Public Function cmbc039_3_CreateXSDCA(ByVal strBlockID As String, ByRef bNoData As Boolean) _
                                        As FUNCTION_RETURN
    Dim intLoopCnt  As Integer
    Dim rs          As OraDynaset
    Dim sSQL        As String
    Dim intProcNo   As Integer

    '' エラーハンドラの設定
    On Error GoTo proc_err

    intProcNo = 0

    'ブロックIDを得る
    sSQL = "SELECT * from XSDCA"
    sSQL = sSQL & " WHERE CRYNUMCA='" & strBlockID & "'"
    sSQL = sSQL & "   AND LIVKCA= '0'"   '生死区分

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        cmbc039_3_CreateXSDCA = FUNCTION_RETURN_FAILURE
        rs.Close
        GoTo proc_exit
    End If

    rs.MoveFirst
    intLoopCnt = 0
    BlkOld.GNLC2 = 0
    BlkOld.GNWC2 = 0
    BlkOld.GNMC2 = 0

    Do While Not rs.EOF
        ReDim Preserve HinOld(intLoopCnt)
        ReDim Preserve HinNow(intLoopCnt)
        With HinOld(intLoopCnt)
            If IsNull(rs.Fields("CRYNUMCA")) = False Then .CRYNUMCA = rs.Fields("CRYNUMCA")
            If IsNull(rs.Fields("HINBCA")) = False Then .HINBCA = rs.Fields("HINBCA")
            If IsNull(rs.Fields("INPOSCA")) = False Then .INPOSCA = rs.Fields("INPOSCA")
            If IsNull(rs.Fields("REVNUMCA")) = False Then .REVNUMCA = rs.Fields("REVNUMCA")
            If IsNull(rs.Fields("FACTORYCA")) = False Then .FACTORYCA = rs.Fields("FACTORYCA")
            If IsNull(rs.Fields("OPECA")) = False Then .OPECA = rs.Fields("OPECA")
            If IsNull(rs.Fields("KCKNTCA")) = False Then .KCKNTCA = rs.Fields("KCKNTCA")
            If IsNull(rs.Fields("SXLIDCA")) = False Then .SXLIDCA = rs.Fields("SXLIDCA")
            If IsNull(rs.Fields("XTALCA")) = False Then .XTALCA = rs.Fields("XTALCA")
            If IsNull(rs.Fields("NEKKNTCA")) = False Then .NEKKNTCA = rs.Fields("NEKKNTCA")
            If IsNull(rs.Fields("NEWKNTCA")) = False Then .NEWKNTCA = rs.Fields("NEWKNTCA")
            If IsNull(rs.Fields("NEWKKBCA")) = False Then .NEWKKBCA = rs.Fields("NEWKKBCA")
            If IsNull(rs.Fields("NEMACOCA")) = False Then .NEMACOCA = rs.Fields("NEMACOCA")
            If IsNull(rs.Fields("GNKKNTCA")) = False Then .GNKKNTCA = rs.Fields("GNKKNTCA")
            If IsNull(rs.Fields("GNWKNTCA")) = False Then .GNWKNTCA = rs.Fields("GNWKNTCA")
            If IsNull(rs.Fields("GNWKKBCA")) = False Then .GNWKKBCA = rs.Fields("GNWKKBCA")
            If IsNull(rs.Fields("GNMACOCA")) = False Then .GNMACOCA = rs.Fields("GNMACOCA")
            If IsNull(rs.Fields("GNDAYCA")) = False Then .GNDAYCA = rs.Fields("GNDAYCA")
            If IsNull(rs.Fields("GNLCA")) = False Then .GNLCA = rs.Fields("GNLCA")
            If IsNull(rs.Fields("GNWCA")) = False Then .GNWCA = rs.Fields("GNWCA")
            If IsNull(rs.Fields("GNMCA")) = False Then .GNMCA = rs.Fields("GNMCA")
            BlkOld.GNLC2 = CLng(BlkOld.GNLC2) + CLng(.GNLCA)
            BlkOld.GNWC2 = CLng(BlkOld.GNWC2) + CLng(.GNWCA)
            BlkOld.GNMC2 = CLng(BlkOld.GNMC2) + CLng(.GNMCA)
            If IsNull(rs.Fields("SUMITLCA")) = False Then .SUMITLCA = rs.Fields("SUMITLCA")
            If IsNull(rs.Fields("SUMITWCA")) = False Then .SUMITWCA = rs.Fields("SUMITWCA")
            If IsNull(rs.Fields("SUMITMCA")) = False Then .SUMITMCA = rs.Fields("SUMITMCA")
            If IsNull(rs.Fields("CHGCA")) = False Then .CHGCA = rs.Fields("CHGCA")
            If IsNull(rs.Fields("KAKOUBCA")) = False Then .KAKOUBCA = rs.Fields("KAKOUBCA")
            If IsNull(rs.Fields("KEIDAYCA")) = False Then .KEIDAYCA = rs.Fields("KEIDAYCA")
            If IsNull(rs.Fields("GNTKUBCA")) = False Then .GNTKUBCA = rs.Fields("GNTKUBCA")
            If IsNull(rs.Fields("GNTNOCA")) = False Then .GNTNOCA = rs.Fields("GNTNOCA")
            If IsNull(rs.Fields("XTWORKCA")) = False Then .XTWORKCA = rs.Fields("XTWORKCA")
            If IsNull(rs.Fields("WFWORKCA")) = False Then .WFWORKCA = rs.Fields("WFWORKCA")
            If IsNull(rs.Fields("LSTATBCA")) = False Then .LSTATBCA = rs.Fields("LSTATBCA")
            If IsNull(rs.Fields("RSTATBCA")) = False Then .RSTATBCA = rs.Fields("RSTATBCA")
            If IsNull(rs.Fields("LUFRCCA")) = False Then .LUFRCCA = rs.Fields("LUFRCCA")
            If IsNull(rs.Fields("LUFRBCA")) = False Then .LUFRBCA = rs.Fields("LUFRBCA")
            If IsNull(rs.Fields("LDFRCCA")) = False Then .LDFRCCA = rs.Fields("LDFRCCA")
            If IsNull(rs.Fields("LDFRBCA")) = False Then .LDFRBCA = rs.Fields("LDFRBCA")
            If IsNull(rs.Fields("HOLDCCA")) = False Then .HOLDCCA = rs.Fields("HOLDCCA")
            If IsNull(rs.Fields("HOLDBCA")) = False Then .HOLDBCA = rs.Fields("HOLDBCA")
            If IsNull(rs.Fields("EXKUBCA")) = False Then .EXKUBCA = rs.Fields("EXKUBCA")
            If IsNull(rs.Fields("HENPKCA")) = False Then .HENPKCA = rs.Fields("HENPKCA")
            If IsNull(rs.Fields("LIVKCA")) = False Then .LIVKCA = rs.Fields("LIVKCA")
            If IsNull(rs.Fields("KANKCA")) = False Then .KANKCA = rs.Fields("KANKCA")
            If IsNull(rs.Fields("NFCA")) = False Then .NFCA = rs.Fields("NFCA")
            If IsNull(rs.Fields("SAKJCA")) = False Then .SAKJCA = rs.Fields("SAKJCA")
            If IsNull(rs.Fields("TDAYCA")) = False Then .TDAYCA = rs.Fields("TDAYCA")
            If IsNull(rs.Fields("KDAYCA")) = False Then .KDAYCA = rs.Fields("KDAYCA")
            If IsNull(rs.Fields("SUMITBCA")) = False Then .SUMITBCA = rs.Fields("SUMITBCA")
            If IsNull(rs.Fields("SNDKCA")) = False Then .SNDKCA = rs.Fields("SNDKCA")
            If IsNull(rs.Fields("SNDDAYCA")) = False Then .SNDDAYCA = rs.Fields("SNDDAYCA")
            If IsNull(rs.Fields("PLANTCATCA")) = False Then .PLANTCATCA = rs.Fields("PLANTCATCA")   '2007/09/04 SPK Tsutsumi Add
        End With
        '基本情報にデータ件数をセット
        With Kihon
            .CNTHINOLD = intLoopCnt + 1
        End With
        intLoopCnt = intLoopCnt + 1
        rs.MoveNext
    Loop

    rs.Close
    cmbc039_3_CreateXSDCA = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    cmbc039_3_CreateXSDCA = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'*******************************************************************************************
'*    関数名        : cmbc039_3_CreateXSDC2
'*
'*    処理概要      : 1.分割結晶（ブロック）前工程実績取得＆構造体作成
'*
'*    パラメータ    : 変数名        ,IO ,型      ,説明
'*                    strBlockID    ,O  ,String  ,ブロックID
'*                    bNoData　　   ,I  ,Boolean ,データ有無フラグ
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************
Public Function cmbc039_3_CreateXSDC2(ByVal strBlockID As String, ByRef bNoData As Boolean) _
                                        As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sSQL        As String
    Dim intProcNo   As Integer

    '' エラーハンドラの設定
    On Error GoTo proc_err

    intProcNo = 0
    bNoData = False

    'ブロックIDを得る
    sSQL = "SELECT * from XSDC2 "
    sSQL = sSQL & " WHERE CRYNUMC2='" & strBlockID & "'"
    sSQL = sSQL & "   AND LIVKC2= '0'"   '生死区分

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        bNoData = True
        cmbc039_3_CreateXSDC2 = FUNCTION_RETURN_FAILURE
        rs.Close
        GoTo proc_exit
    End If

    rs.MoveFirst
    If rs.EOF = False Then
        With BlkOld
            If IsNull(rs.Fields("CRYNUMC2")) = False Then .CRYNUMC2 = rs.Fields("CRYNUMC2")
            If IsNull(rs.Fields("KCNTC2")) = False Then .KCNTC2 = rs.Fields("KCNTC2")       '工程連番
            If IsNull(rs.Fields("XTALC2")) = False Then .XTALC2 = rs.Fields("XTALC2")
            If IsNull(rs.Fields("INPOSC2")) = False Then .INPOSC2 = rs.Fields("INPOSC2")
            If IsNull(rs.Fields("NEKKNTC2")) = False Then .NEKKNTC2 = rs.Fields("NEKKNTC2")
            If IsNull(rs.Fields("NEWKNTC2")) = False Then .NEWKNTC2 = rs.Fields("NEWKNTC2")
            If IsNull(rs.Fields("NEWKKBC2")) = False Then .NEWKKBC2 = rs.Fields("NEWKKBC2")
            If IsNull(rs.Fields("NEMACOC2")) = False Then .NEMACOC2 = rs.Fields("NEMACOC2")
            If IsNull(rs.Fields("GNKKNTC2")) = False Then .GNKKNTC2 = rs.Fields("GNKKNTC2")
            If IsNull(rs.Fields("GNWKNTC2")) = False Then .GNWKNTC2 = rs.Fields("GNWKNTC2")
            If IsNull(rs.Fields("GNWKKBC2")) = False Then .GNWKKBC2 = rs.Fields("GNWKKBC2")
            If IsNull(rs.Fields("GNMACOC2")) = False Then .GNMACOC2 = rs.Fields("GNMACOC2")
            If IsNull(rs.Fields("GNDAYC2")) = False Then .GNDAYC2 = rs.Fields("GNDAYC2")
            If IsNull(rs.Fields("GNLC2")) = False Then .GNLC2 = rs.Fields("GNLC2")          '現在長さ
            If IsNull(rs.Fields("GNWC2")) = False Then .GNWC2 = rs.Fields("GNWC2")          '現在重量
            If IsNull(rs.Fields("GNMC2")) = False Then .GNMC2 = rs.Fields("GNMC2")          '現在枚数
            If IsNull(rs.Fields("SUMITLC2")) = False Then .SUMITLC2 = rs.Fields("SUMITLC2")
            If IsNull(rs.Fields("SUMITWC2")) = False Then .SUMITWC2 = rs.Fields("SUMITWC2")
            If IsNull(rs.Fields("SUMITMC2")) = False Then .SUMITMC2 = rs.Fields("SUMITMC2")
            If IsNull(rs.Fields("CHGC2")) = False Then .CHGC2 = rs.Fields("CHGC2")
            If IsNull(rs.Fields("KAKOUBC2")) = False Then .KAKOUBC2 = rs.Fields("KAKOUBC2")
            If IsNull(rs.Fields("KEIDAYC2")) = False Then .KEIDAYC2 = rs.Fields("KEIDAYC2")
            If IsNull(rs.Fields("GNTKUBC2")) = False Then .GNTKUBC2 = rs.Fields("GNTKUBC2")
            If IsNull(rs.Fields("GNTNOC2")) = False Then .GNTNOC2 = rs.Fields("GNTNOC2")
            If IsNull(rs.Fields("XTWORKC2")) = False Then .XTWORKC2 = rs.Fields("XTWORKC2")
            If IsNull(rs.Fields("WFWORKC2")) = False Then .WFWORKC2 = rs.Fields("WFWORKC2")
            If IsNull(rs.Fields("LSTATBC2")) = False Then .LSTATBC2 = rs.Fields("LSTATBC2")
            If IsNull(rs.Fields("RSTATBC2")) = False Then .RSTATBC2 = rs.Fields("RSTATBC2")
            If IsNull(rs.Fields("LUFRCC2")) = False Then .LUFRCC2 = rs.Fields("LUFRCC2")
            If IsNull(rs.Fields("LUFRBC2")) = False Then .LUFRBC2 = rs.Fields("LUFRBC2")
            If IsNull(rs.Fields("LDFRCC2")) = False Then .LDFRCC2 = rs.Fields("LDFRCC2")
            If IsNull(rs.Fields("LDFRBC2")) = False Then .LDFRBC2 = rs.Fields("LDFRBC2")
            If IsNull(rs.Fields("HOLDCC2")) = False Then .HOLDCC2 = rs.Fields("HOLDCC2")
            If IsNull(rs.Fields("HOLDBC2")) = False Then .HOLDBC2 = rs.Fields("HOLDBC2")
            If IsNull(rs.Fields("EXKUBC2")) = False Then .EXKUBC2 = rs.Fields("EXKUBC2")
            If IsNull(rs.Fields("HENPKC2")) = False Then .HENPKC2 = rs.Fields("HENPKC2")
            If IsNull(rs.Fields("LIVKC2")) = False Then .LIVKC2 = rs.Fields("LIVKC2")
            If IsNull(rs.Fields("KANKC2")) = False Then .KANKC2 = rs.Fields("KANKC2")
            If IsNull(rs.Fields("NFC2")) = False Then .NFC2 = rs.Fields("NFC2")
            If IsNull(rs.Fields("SAKJC2")) = False Then .SAKJC2 = rs.Fields("SAKJC2")
            If IsNull(rs.Fields("TDAYC2")) = False Then .TDAYC2 = rs.Fields("TDAYC2")
            If IsNull(rs.Fields("KDAYC2")) = False Then .KDAYC2 = rs.Fields("KDAYC2")
            If IsNull(rs.Fields("SUMITBC2")) = False Then .SUMITBC2 = rs.Fields("SUMITBC2")
            If IsNull(rs.Fields("SNDKC2")) = False Then .SNDKC2 = rs.Fields("SNDKC2")
            If IsNull(rs.Fields("SNDDAYC2")) = False Then .SNDDAYC2 = rs.Fields("SNDDAYC2")
            If IsNull(rs.Fields("PLANTCATC2")) = False Then .PLANTCATC2 = rs.Fields("PLANTCATC2")   '2007/09/04 SPK Tsutsumi Add
        End With
    End If

    rs.Close
    cmbc039_3_CreateXSDC2 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    cmbc039_3_CreateXSDC2 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'***************************************************************************************************
'*    関数名        : cmbc039_3_CreateNowProc
'*
'*    処理概要      : 1.現在工程構造体作成
'*
'*    パラメータ    : 変数名           ,IO ,型      ,説明
'*                    strSXLID         ,I  ,String  ,SXL-ID
'*                    lngBeginIngotpos ,I  ,Long    ,ブロック管理データの長さ
'*                    strErrMsg        ,O  ,String  ,ErrMsg格納
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'***************************************************************************************************
'Cng Start 2010/10/03 Y.Hitomi
Public Function cmbc039_3_CreateNowProc(ByVal strBlockID As String, ByVal lngBeginIngotpos As Long, _
                                        ByVal lngEndIngotpos As Long, ByVal lngWfBeginSeq As Long, _
                                        ByVal lngWfEndSeq As Long, _
                                        ByRef strErrMsg As String, ByVal strSXLID As String) As FUNCTION_RETURN
'Public Function cmbc039_3_CreateNowProc(ByVal strBlockID As String, ByVal lngBeginIngotpos As Long, _
                                        ByVal lngEndIngotpos As Long, ByVal lngWfBeginSeq As Long, _
                                        ByVal lngWfEndSeq As Long, _
                                        ByRef strErrMsg As String) As FUNCTION_RETURN
'Cng End 2010/10/03 Y.Hitomi

    Dim rs              As OraDynaset
    Dim rs2             As OraDynaset
    Dim sSQL            As String
    Dim intProcNo       As Integer
    Dim intHinOldCnt    As Integer
    Dim intLengthCnt    As Integer
    Dim intLoopCnt      As Integer
    Dim dblDiameter     As Double
    Dim intNum          As Integer
    Dim sCryNum         As String
    Dim intBlkLength    As Integer  'ブロック管理データの長さ
    Dim intBlkIngotPos  As Integer  'ブロック管理データの位置
    Dim intSxlLength    As Integer  'シングル管理データの長さ
    Dim intSxlIngotPos  As Integer  'シングル管理データの位置
    Dim blFlg           As Boolean
    Dim intSP           As Integer  '長さ判定用
    Dim intEP           As Integer  '長さ判定用
    Dim intSBP          As Integer  '長さ判定用
    Dim intEBP          As Integer  '長さ判定用
    Dim intLength       As Integer  '長さ
    Dim intIngotpos     As Integer  '位置
    Dim blRtn           As Boolean  '戻り値
    Dim intWFcnt        As Integer  'WFマップ枚数 add 2003/03/29 hitec)matsumoto
    Dim sMotoHinban     As String

    '' エラーハンドラの設定
    On Error GoTo proc_err

    intProcNo = 0

    'ブロック管理から情報取得
    sSQL = "SELECT * from TBCME040 "
    sSQL = sSQL & " WHERE BLOCKID='" & strBlockID & "'"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        cmbc039_3_CreateNowProc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    intLoopCnt = 0
    If rs.EOF = False Then
        If IsNull(rs("CRYNUM")) = False Then sCryNum = rs("CRYNUM")           '結晶番号
        If IsNull(rs("LENGTH")) = False Then intBlkLength = rs("LENGTH")        '長さ
        If IsNull(rs("INGOTPOS")) = False Then intBlkIngotPos = rs("ingotpos")  '位置
    End If

    rs.Close

    'upd start 2003/03/29 hitec)matsumoto 全数廃棄かは、WFマップを見る----------

    'ブロックIDを得る       'del 2003/03/29 hitec)matsumoto 下へ移動
    sSQL = "SELECT LOTID from TBCMY011 "
    sSQL = sSQL & " WHERE LOTID='" & strBlockID & "'"     '2003/04/03 hitec)matsumoto 全数スクラップ="Y"はブロック単位なので、シングル範囲で取れない
    sSQL = sSQL & "   AND TO_NUMBER(WFSTA) <= 1"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount = 0 Then
        '前工程実績を、現在工程実績にコピー
        BlkNow = BlkOld
        BlkNow.GNLC2 = "0"
        BlkNow.GNWC2 = "0"
        BlkNow.GNMC2 = "0"
        BlkNow.GNWKNTC2 = Kihon.NEWPROC
        BlkNow.NEWKNTC2 = Kihon.NOWPROC
        For intHinOldCnt = 0 To Kihon.CNTHINOLD - 1
            ReDim Preserve HinNow(intHinOldCnt) As typ_XSDCA_Update
            HinNow(intHinOldCnt) = HinOld(intHinOldCnt)
            HinNow(intHinOldCnt).GNLCA = "0"    '全数スクラップ=長さが0
            HinNow(intHinOldCnt).GNWCA = "0"    '重量 = 0
            HinNow(intHinOldCnt).GNMCA = "0"    '枚数 = 0
            HinNow(intHinOldCnt).GNWKNTCA = Kihon.NEWPROC
            HinNow(intHinOldCnt).NEWKNTCA = Kihon.NOWPROC
        Next
        Kihon.CNTHINNOW = 1
        Kihon.ALLSCRAP = "Y"

    '前工程の長さと現在工程の長さをくらべ、不良が存在するか判定 'upd 2003/03/29 hitec)matsumoto 長さではなく枚数で比べる
        If CInt(BlkNow.GNMC2) = CInt(BlkOld.GNMC2) Then '不良なし
            '基本情報構造体
            With Kihon
                .FURYOUMU = "N"
            End With
        Else
            rs.Close
            strErrMsg = GetMsgStr("EWFM5", "前工程=" & BlkOld.GNMC2 & "：現在工程=" & BlkNow.GNMC2) '03/06/06 後藤
            cmbc039_3_CreateNowProc = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        rs.Close
        cmbc039_3_CreateNowProc = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If
    rs.Close

    '前工程の構造体を現在工程の構造体へコピー
    BlkNow = BlkOld
    '工程連番に＋１する
    With BlkNow
        If BlkNow.KCNTC2 = "" Then BlkNow.KCNTC2 = "0"
        .KCNTC2 = CInt(BlkNow.KCNTC2) + 1   '工程連番
        'Cng Start 2010/09/02 Y.Hitomi
        ''ブロック内SXLが1つでも完了していた場合、工程コードを更新しないようにする
        If (.GNWKNTC2 <> "     " Or _
            .GNWKNTC2 <> "CW800" Or _
            .GNWKNTC2 <> "TX860") Then
            
            .NEWKNTC2 = Kihon.NOWPROC           '前工程
            .GNWKNTC2 = Kihon.NEWPROC           '現在工程
        End If
'            .NEWKNTC2 = Kihon.NOWPROC           '前工程
'            .GNWKNTC2 = Kihon.NEWPROC           '現在工程
        'Cng End 2010/09/02 Y.Hitomi
        
        .SUMITBC2 = "0"
        .SUMITLC2 = "0"
        .SUMITMC2 = "0"
        .SUMITWC2 = "0"
    End With

    sSQL = ""
    sSQL = sSQL & " SELECT"
    sSQL = sSQL & "  xtalcb as CRYNUM"        ' 結晶番号
    sSQL = sSQL & " ,inposcb as INGOTPOS"     ' 結晶内開始位置
    sSQL = sSQL & " ,rlencb as LENGTH"        ' 長さ
    sSQL = sSQL & " ,sxlidcb as SXLID"        ' SXLID
    sSQL = sSQL & " ,' ' as KRPROCCD"         ' 管理工程
    sSQL = sSQL & " ,gnwkntcb as NOWPROC"     ' 現在工程
    sSQL = sSQL & " ,' ' as LPKRPROCCD"       ' 最終通過管理工程
    sSQL = sSQL & " ,newkntcb as LASTPASS"    ' 最終通過工程
    sSQL = sSQL & " ,livkcb as DELCLS"        ' 削除区分
    sSQL = sSQL & " ,lstccb as LSTATCLS"      ' 最終状態区分
    sSQL = sSQL & " ,sholdclscb HOLDCLS"      ' ホールド区分
    sSQL = sSQL & " ,hinbcb as HINBAN"        ' 品番
    sSQL = sSQL & " ,revnumcb as REVNUM"      ' 製品番号改訂番号
    sSQL = sSQL & " ,factorycb as FACTORY"    ' 工場
    sSQL = sSQL & " ,opecb as OPECOND"        ' 操業条件
    sSQL = sSQL & " ,maicb"                   ' 枚数
    sSQL = sSQL & " ,tdaycb as REGDATE"       ' 登録日付
    sSQL = sSQL & " ,kdaycb as UPDDATE"       ' 更新日付
    sSQL = sSQL & " ,' ' as SUMMITSENDFLAG"   ' SUMMIT送信フラグ
    sSQL = sSQL & " ,sndkcb as SENDFLAG"      ' 送信フラグ
    sSQL = sSQL & " ,sndaycb as SENDDATE"     ' 送信日付
    sSQL = sSQL & " ,plantcatcb as PLANTCAT"  ' 向先 2007/09/04 SPK Tsutsumi Add
    sSQL = sSQL & " FROM XSDCB"
    sSQL = sSQL & " WHERE sxlidcb like '" & left(sCryNum, 9) & "%'"     'ｲﾝﾃﾞｯｸｽ項目追加 09/05/25 ooba
    sSQL = sSQL & "   AND xtalcb = '" & sCryNum & "'"
    sSQL = sSQL & "   AND ((inposcb >=" & lngBeginIngotpos & ") And (inposcb + rlencb <= " & lngEndIngotpos & "))"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    intLoopCnt = 0
    BlkNow.GNMC2 = 0    '現在工程（ブロック）の枚数をクリアしておく
    Do While Not rs.EOF
        ReDim Preserve HinNow(intLoopCnt) As typ_XSDCA_Update

        If IsNull(rs("CRYNUM")) = False Then sCryNum = rs("CRYNUM")
        If IsNull(rs("LENGTH")) = False Then intSxlLength = rs("LENGTH")
        If IsNull(rs("INGOTPOS")) = False Then intSxlIngotPos = rs("INGOTPOS")

        '-- ブロックとシングルの位置関係を判定し、長さを算出 --------
        intSP = intSxlIngotPos         'シングル開始位置
        intEP = intSP + intSxlLength      'シングル終端位置
        intSBP = intBlkIngotPos        'ブロック開始位置
        intEBP = intSBP + intBlkLength    'ブロック終端位置

        '' ブロックがSXLの中に完全に含まれている場合 ---------
        If intSP <= intSBP And intEP >= intEBP Then

            intLength = intBlkLength                    'ブロック管理の長さを使用
            intIngotpos = intBlkIngotPos

        '' ブロックがSXLの開始位置より上にあり、かつ終端位置よりも長い場合 ---------
        ElseIf intSP >= intSBP And intEP <= intEBP Then

            intLength = intSxlLength                  'シングル管理の長さを使用
            intIngotpos = intSxlIngotPos

        '' ブロックが一部SXLにかかっている場合
        '' (ブロックが上側。ただしブロックの終端とSXLの開始位置が一致しないこと) ------------
        ElseIf intSP > intSBP And intSP < intEBP And intSP <> intEBP Then

            intLength = intEBP - intSP                        'ブロックの終端位置 - シングルの開始位置
            intIngotpos = intSxlIngotPos

        '' ブロックが一部SXLにかかっている場合
        '' (ブロックが下側。ただしSXLの終端とブロックの開始位置が一致しないこと) ----------
        ElseIf intSP < intSBP And intEP > intSBP And intEP <> intSBP Then

            intLength = intEP - intSBP                        'シングルの終端位置 - ブロックの開始位置
            intIngotpos = intBlkIngotPos

        Else
            GoTo LoopNext
        End If
        '----------------------------------------------------

        '現在工程編集
        With HinNow(intLoopCnt)
            .CRYNUMCA = strBlockID       'ブロックID
            If IsNull(rs("HINBAN")) = False Then .HINBCA = rs("HINBAN")         '品番
            If IsNull(rs("SXLID")) = False Then .SXLIDCA = rs("SXLID")          'シングルID
            If IsNull(rs("REVNUM")) = False Then .REVNUMCA = rs("REVNUM")       '製品番号改訂番号
            If IsNull(rs("FACTORY")) = False Then .FACTORYCA = rs("FACTORY")    '工場
            If IsNull(rs("OPECOND")) = False Then .OPECA = rs("OPECOND")        '操業条件
            If IsNull(rs("CRYNUM")) = False Then .XTALCA = rs("CRYNUM")         '結晶番号
            'add start 2003/04/27 hitec)matsumoto Z品番のものは、良品とさせるため、元品番を取得する------------------
            If Trim(.HINBCA) = "Z" Then
'Cng Start 2012/4/26 Y.Hitomi
'                If GetZMotoHinban(.SXLIDCA, sMotoHinban) = FUNCTION_RETURN_FAILURE Then
'                    cmbc039_3_CreateNowProc = FUNCTION_RETURN_FAILURE
'                    rs.Close
'                    GoTo proc_exit
'                End If
'                If sMotoHinban = vbNullString Then
'                    cmbc039_3_CreateNowProc = FUNCTION_RETURN_FAILURE
'                    rs.Close
'                    GoTo proc_exit
'                End If
'                .HINBCA = sMotoHinban
                .HINBCA = tblSXL.hinban '元品番情報
'Cng End 2012/4/26 Y.Hitomi
                .FACTORYCA = tblSXL.factory   '工場
                .OPECA = tblSXL.opecond       '操業条件
                .REVNUMCA = tblSXL.REVNUM     '結晶番号
                'Add Start 2010/10/14 Y.Hitomi
                .GNWKNTCA = PROCD_SXL_MAP     '現在工程
                'Add End   2010/10/14 Y.Hitomi
            Else
                If IsNull(rs("FACTORY")) = False Then .FACTORYCA = rs("FACTORY")    '工場
                If IsNull(rs("OPECOND")) = False Then .OPECA = rs("OPECOND")        '操業条件
                If IsNull(rs("REVNUM")) = False Then .REVNUMCA = rs("REVNUM")       '製品番号改訂番号
                If IsNull(rs("PLANTCAT")) = False Then .PLANTCATCA = rs("PLANTCAT")     '向先 2007/09/04 SPK Tsutsumi Add
                'Add Start  2010/10/14 Y.Hitomi
                .GNWKNTCA = Kihon.NEWPROC     '現在工程
                'Add End    2010/10/14 Y.Hitomi
            End If
            'add end   2003/04/27 hitec)matsumoto ------------------
        
            .NEWKNTCA = Kihon.NOWPROC   '前工程
            'Cng Start 2010/10/14 Y.Hitomi
            '.GNWKNTCA = Kihon.NEWPROC   '現在工程
            'Cng End   2010/10/14 Y.Hitomi
            
            .KCKNTCA = BlkNow.KCNTC2    '工程連番
            .NEMACOCA = BlkNow.NEMACOC2 '最終通過処理回数
            .GNMACOCA = BlkNow.GNMACOC2 '現在処理回数
            .SUMITBCA = "0"
            .SUMITLCA = "0"
            .SUMITMCA = "0"
            .SUMITWCA = "0"
            .INPOSCA = intIngotpos      '結晶内開始位置
            .GNLCA = intLength          '長さ

            '現在重量を求める
            If GetDiameter(strBlockID, dblDiameter) = FUNCTION_RETURN_FAILURE Then  '直径を求める
                dblDiameter = 0
            End If
            '取得した直径を元に重量を求める
            HinNow(intLoopCnt).GNWCA = CStr(WeightOfCylinder(dblDiameter, CDbl(.GNLCA)))

            sSQL = "SELECT LOTID from TBCMY011 "         'upd 2003/04/29 hitec)matsumoto シングルを条件に取得するのではなく、ブロック単位で枚数が取得できるよう、条件を変更
            sSQL = sSQL & " WHERE LOTID = '" & .CRYNUMCA & "'"
            sSQL = sSQL & " AND MSXLID ='" & .SXLIDCA & "'"    ''' 03/04/27 修正 後藤
            sSQL = sSQL & "   AND TO_NUMBER(WFSTA) <= 1"

            Set rs2 = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
            intWFcnt = 0
            Do While Not rs2.EOF
                intWFcnt = intWFcnt + 1
                rs2.MoveNext
            Loop
            rs2.Close

            .GNMCA = intWFcnt   'add 2003/03/29 hitec)matsumoto 上でWFマップテーブルから枚数カウント取得しているので、それを良品枚数とする
            .SUMITLCA = .GNLCA ''' 03/05/13 後藤
            .SUMITMCA = .GNMCA
            .SUMITWCA = .GNWCA
        End With

        With BlkNow
            '現在重量を求める
            If GetDiameter(strBlockID, dblDiameter) = FUNCTION_RETURN_FAILURE Then  '直径を求める
                dblDiameter = 0
            End If
            '基本情報の直径セット
            Kihon.DIAMETER = dblDiameter
            '取得した直径を元に重量を求める
            .GNMC2 = CStr(CLng(BlkNow.GNMC2) + CLng(HinNow(intLoopCnt).GNMCA))  '枚数 'upd 2003/03/29 hitec)matsumoto 枚数再計算はしない
        End With
        intLoopCnt = intLoopCnt + 1
        '良品件数セット
        With Kihon
            .CNTHINNOW = intLoopCnt
        End With

LoopNext:

        rs.MoveNext
    Loop

    rs.Close

    '現在工程で作られたもの以外のデータが前工程に存在した場合取得
    If GetOtherData(intBlkLength, intBlkIngotPos) = FUNCTION_RETURN_FAILURE Then
        cmbc039_3_CreateNowProc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '前工程の長さと現在工程の長さをくらべ、不良が存在するか判定
    If BlkNow.GNLC2 = "" Then BlkNow.GNLC2 = "0"
    If BlkOld.GNLC2 = "" Then BlkOld.GNLC2 = "0"
    If BlkNow.GNMC2 = "" Then BlkNow.GNMC2 = "0"
    If BlkOld.GNMC2 = "" Then BlkOld.GNMC2 = "0"
    If CInt(BlkNow.GNMC2) = CInt(BlkOld.GNMC2) Then '不良なし
        '基本情報構造体
        With Kihon
            .FURYOUMU = "N"
        End With
    Else                                            '不良あり
        rs.Close
        strErrMsg = GetMsgStr("EWFM5", "前工程=" & BlkOld.GNMC2 & "：現在工程=" & BlkNow.GNMC2) '03/06/06 後藤
        cmbc039_3_CreateNowProc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    cmbc039_3_CreateNowProc = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    cmbc039_3_CreateNowProc = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*********************************************************************************************
'*    関数名        : GetOtherData
'*
'*    処理概要      : 1.SXLの全ブロック入庫チェック
'*
'*    パラメータ    : 変数名           ,IO ,型      ,説明
'*                    intBlkLength     ,I  ,Integer ,ブロック管理データの長さ
'*                    lngBeginIngotpos ,I  ,Integer ,ブロック管理データの長さ
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************
Private Function GetOtherData(ByVal intBlkLength As Integer, ByVal intBlkIngotPos As Integer) _
                                As FUNCTION_RETURN
    Dim rs              As OraDynaset
    Dim sSQL            As String
    Dim intHinOldCnt    As Integer
    Dim intHinNowCnt    As Integer
    Dim intHinCnt       As Integer
    Dim blUpdFlg        As Boolean
    Dim intFuryouCnt    As Integer
    Dim intWFcnt        As Integer

    intHinCnt = Kihon.CNTHINNOW

    For intHinOldCnt = 0 To Kihon.CNTHINOLD - 1
        blUpdFlg = False
        For intHinNowCnt = 0 To intHinCnt - 1
            If (HinOld(intHinOldCnt).XTALCA = HinNow(intHinNowCnt).XTALCA) _
                And (HinOld(intHinOldCnt).INPOSCA = HinNow(intHinNowCnt).INPOSCA) Then

                blUpdFlg = True
            End If
        Next

        If blUpdFlg = False Then
            '前工程品番にあって、現在工程品番にないものは、前工程品番を現在工程品番にコピー
            ReDim Preserve HinNow(Kihon.CNTHINNOW) As typ_XSDCA_Update
            '前工程品番をコピー
            HinNow(Kihon.CNTHINNOW) = HinOld(intHinOldCnt)
            With HinNow(Kihon.CNTHINNOW)
                .KCKNTCA = BlkNow.KCNTC2
                'Cng Start 2010/10/14 Y.Hitomi
                .NEWKNTCA = HinOld(intHinOldCnt).NEWKNTCA
                .GNWKNTCA = HinOld(intHinOldCnt).GNWKNTCA
'                .NEWKNTCA = Kihon.NOWPROC
'                .GNWKNTCA = Kihon.NEWPROC
                'Cng End 2010/10/14 Y.Hitomi
                .SUMITBCA = "0"

                sSQL = "SELECT LOTID from TBCMY011 "
                sSQL = sSQL & " WHERE LOTID = '" & .CRYNUMCA & "'"
                sSQL = sSQL & "   AND MSXLID ='" & .SXLIDCA & "'"
                sSQL = sSQL & "   AND TO_NUMBER(WFSTA) <= 1"

                Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
                intWFcnt = 0
                Do While Not rs.EOF
                    intWFcnt = intWFcnt + 1
                    rs.MoveNext
                Loop
                rs.Close
                .GNMCA = intWFcnt

                BlkNow.GNLC2 = CLng(BlkNow.GNLC2) + val(.GNLCA)
                BlkNow.GNMC2 = CLng(BlkNow.GNMC2) + val(.GNMCA)
                BlkNow.GNWC2 = CLng(BlkNow.GNWC2) + val(.GNWCA)
            End With
            Kihon.CNTHINNOW = Kihon.CNTHINNOW + 1
        End If
    Next
    GetOtherData = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function
End Function

'*******************************************************************************
'*    関数名        : Shikibetsu
'*
'*    処理概要      : 1.サンプルデータ作成識別
'*                    (抵抗保証フラグ(HSXRHWYS)、酸素保証フラグ(HSXONHWS))
'*　　　　　　　　　　(どちらかでも「Ｈ」だったらTrue)
'*
'*    パラメータ    : 変数名        ,IO ,型      ,説明
'*                    hinb          ,I  ,String  ,品番
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function Shikibetsu(ByVal hinb As String) As Boolean
    Dim blDbIsMine      As Boolean
    Dim rs              As OraDynaset
    Dim sSQL            As String
    Dim i               As Integer
    Dim intRecCnt       As Integer
    Dim sWork1, sWork2  As String

    Shikibetsu = False
    If OraDB Is Nothing Then
        blDbIsMine = True
        OraDBOpen
    End If

    ''汎用コードマスタから、コードNOに対応するコードの一覧を得る
    sSQL = "select E18.HSXRHWYS, E19.HSXONHWS"
    sSQL = sSQL & " from TBCME018 E18 ,TBCME019 E19 "
    sSQL = sSQL & " where rtrim(E18.HINBAN) = '" & Trim$(hinb) & "' "
    sSQL = sSQL & " and E18.HINBAN = E19.HINBAN "
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_DEFAULT)
    If rs.EOF <> True Then
        sWork1 = rs("HSXRHWYS")
        sWork2 = rs("HSXONHWS")
        If Trim(sWork1) = "H" Or Trim(sWork2) = "H" Then
            Shikibetsu = True
        End If
    End If
    rs.Close

    If blDbIsMine Then
        OraDBClose
    End If
End Function

'****************************************************************************************
'*    関数名        : DBDRV_MIN_MAX_SEQGET
'*
'*    処理概要      : 1.抜試指示 MIN,MAX値を取得
'*                      (SXLID,BLOCKID→最大、最小（ブロックＰで判定）のデータを取得する)
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    iWfNum        ,O  ,Integer  ,SXL管理
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*****************************************************************************************
Public Function DBDRV_MIN_MAX_SEQGET(ByRef iWfNum As Integer) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim i           As Long
    Dim lngCnt      As Long
    Dim sDBName     As String
    Dim intUCount   As Integer
    Dim dblWFLen    As Double  '2003/04/25 hitec)okazaki
    Dim iRtn        As FUNCTION_RETURN
    Dim dblEPS      As Double

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_MIN_MAX_SEQGET"

    dblEPS = 0.000001        'εの設定
    iWfNum = 0
    sDBName = "(V001)"
    i = 0

    ' SXLIDの取得
    For i = 0 To UBound(tSXLID())
        sSQL = "select "
        sSQL = sSQL & "LOTID,"                ' ブロックID"
        sSQL = sSQL & "MSXLID,"                ' SXLID"
        sSQL = sSQL & "blockseq,"             ' ブロック内連番"
        sSQL = sSQL & "WFSTA,"                ' WF状態"
        sSQL = sSQL & "MHINBAN,"               ' 品番"
        sSQL = sSQL & "RTOP_POS,"             ' 論理ブロック内位置"
        sSQL = sSQL & "RITOP_POS,"            ' 論理結晶内位置"
        sSQL = sSQL & "MSMPLEID,"              ' 抜試位置"
        sSQL = sSQL & "SHAFLAG,"               ' サンプルフラグ"
        sSQL = sSQL & "INDTM,"
        sSQL = sSQL & "BASKETID,"
        sSQL = sSQL & "SLOTNO,"
        sSQL = sSQL & "CURRWPCS,"
        sSQL = sSQL & "EXISTFLG,"
        sSQL = sSQL & "TOP_POS,"
        sSQL = sSQL & "REJCAT,"
        sSQL = sSQL & "TXID,"
        sSQL = sSQL & "REGDATE,"
        sSQL = sSQL & "SUMMITSENDFLAG,"
        sSQL = sSQL & "SENDFLAG,"
        sSQL = sSQL & "SENDDATE,"
        sSQL = sSQL & "HREJCODE,"
        sSQL = sSQL & "UPDPROC,"
        sSQL = sSQL & "UPDDATE,"
        sSQL = sSQL & "MREVNUM,"
        sSQL = sSQL & "MFACTORY,"
        sSQL = sSQL & "MOPECOND,"
        sSQL = sSQL & "kankbn,"
        sSQL = sSQL & "NREJCODE"
        sSQL = sSQL & " from TBCMY011 "
        sSQL = sSQL & " where LOTID ='" & tSXLID(i).LOTID & "'"
        sSQL = sSQL & "   and MSXLID ='" & tSXLID(i).SXLID & "'"
        sSQL = sSQL & " ORDER BY blockseq ASC"

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        iWfNum = 0
        Do While Not rs.EOF
            If CInt(rs.Fields("WFSTA")) <= 1 Then
                iWfNum = iWfNum + 1
            End If
            rs.MoveNext
        Loop
        If rs.RecordCount = 0 Then
            DBDRV_MIN_MAX_SEQGET = FUNCTION_RETURN_FAILURE
            f_cmbc039_3.lblMsg.Caption = GetMsgStr("EWFM6", "(Y011)")
            Exit Function
        End If
        rs.MoveFirst    '先頭ﾚｺｰﾄﾞに移動

        Do While Not rs.EOF
            '先頭ﾚｺｰﾄﾞ
            ReDim Preserve tExamine(intUCount)   '配列の再定義
            With tExamine(intUCount)
                If IsNull(rs!LOTID) = True Then
                    .LOTID = vbNullString           ' ブロックID
                 Else
                    .LOTID = rs!LOTID
                End If
                If IsNull(rs!MSXLID) = True Then
                    .SXLID = vbNullString
                Else
                    .SXLID = rs!MSXLID               ' SXLID
                End If
                .MinMax = 0                         ' 0:MIN 1:MAX
                If IsNull(rs!BLOCKSEQ) = True Then
                    .BLOCKSEQ = vbNullString
                Else
                    .BLOCKSEQ = rs!BLOCKSEQ         ' ブロック内連番
                End If
                If IsNull(rs!WFSTA) = True Then
                    .WFSTA = vbNullString
                Else
                    .WFSTA = rs!WFSTA               ' WF状態
                End If
                If IsNull(rs!mhinban) = True Then
                    .hinban = vbNullString
                Else
                    .hinban = rs!mhinban             ' 品番
                End If
                If IsNull(rs!MSMPLEID) = True Then
                    .SMPLEID = vbNullString
                Else
                    .SMPLEID = rs!MSMPLEID          ' 抜試位置
                End If
                If IsNull(rs!RTOP_POS) = True Then
                    .RTOP_POS = 0
                Else
                    'WF一枚の長さ取得                                   '2003/04/25 hitec)okazaki
                    iRtn = DBDRV_WFLENGET(tSXLID(i).LOTID, dblWFLen)
                    'ブロック先頭の表示位置はWF一枚の長さを引いたもの   '2003/04/25 hitec)okazaki
                    If Right(.SMPLEID, 1) <> "D" Then
                        .RTOP_POS = Fix(CDbl(rs.Fields("RTOP_POS")) - dblWFLen + 0.99999)       ' 論理ブロック内位置  'upd 2003/08/06 hitec)matsumoto
                    Else
                        .RTOP_POS = Int(CDbl(rs.Fields("RTOP_POS")) + dblEPS)      'Dの場合切り捨て(WF幅内に整数位置が2つ以上ある場合の対応)　06/10/27 ooba
                    End If
                End If
                If IsNull(rs!RITOP_POS) = True Then
                    .RITOP_POS = 0
                Else
                    'ブロック先頭の表示位置はWF一枚の長さを引いたもの   '2003/04/25 hitec)okazaki
                    If Right(.SMPLEID, 1) <> "D" Then
                        .RITOP_POS = Fix(CDbl(rs.Fields("RITOP_POS")) - dblWFLen + 0.99999)        ' 論理結晶内位置   'upd 2003/08/06 hitec)matsumoto
                    Else
                        .RITOP_POS = Int(CDbl(rs.Fields("RITOP_POS")) + dblEPS)    'Dの場合切り捨て(WF幅内に整数位置が2つ以上ある場合の対応)　06/10/27 ooba
                    End If
                End If
                If IsNull(rs!SHAFLAG) = True Then
                    .SHAFLAG = vbNullString
                Else
                    .SHAFLAG = rs!SHAFLAG           ' サンプルフラグ
                    If Trim(.SHAFLAG) = "1" Then
                        If Trim(.SMPLEID) = vbNullString Then   'add 2003/06/24 hitec)matsumoto サンプルフラグが
                            f_cmbc039_3.cmdF(6).Enabled = False
                            f_cmbc039_3.cmdF(7).Enabled = False
                            f_cmbc039_3.cmdF(8).Enabled = False
                            f_cmbc039_3.cmdF(9).Enabled = False
                            f_cmbc039_3.cmdF(10).Enabled = False
                            f_cmbc039_3.cmdF(12).Enabled = False
                            DBDRV_MIN_MAX_SEQGET = FUNCTION_RETURN_FAILURE
                            f_cmbc039_3.lblMsg.Caption = GetMsgStr("ENSP4", "Y011")
                            rs.Close
                            Exit Function
                        End If
                    End If
                End If
                If IsNull(rs!INDTM) = True Then
                Else
                    .INDTM = rs!INDTM               ' ウェハーセンター入庫日時
                End If
                If IsNull(rs!BASKETID) = True Then
                    .BASKETID = vbNullString
                Else
                    .BASKETID = rs!BASKETID         ' バスケットID
                End If
                If IsNull(rs!SLOTNO) = True Then
                    .SLOTNO = vbNullString
                Else
                    .SLOTNO = rs!SLOTNO             ' スロットNO
                End If
                If IsNull(rs!CURRWPCS) = True Then
                    .CURRWPCS = vbNullString
                Else
                    .CURRWPCS = iWfNum              ' ウェハー枚数
                End If
                If IsNull(rs!EXISTFLG) = True Then
                    .EXISTFLG = vbNullString
                Else
                    .EXISTFLG = rs!EXISTFLG         ' 存在フラグ
                End If
                If IsNull(rs!TOP_POS) = True Then
                    .TOP_POS = 0
                Else
                    .TOP_POS = Int(CDbl(rs!TOP_POS) / 10 + dblEPS)         ' ブロックのTOPからの位置   'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                End If
                If IsNull(rs!REJCAT) = True Then
                    .REJCAT = vbNullString
                Else
                    .REJCAT = rs!REJCAT             ' 欠落理由
                End If
                If IsNull(rs!TXID) = True Then
                    .TXID = vbNullString
                Else
                    .TXID = rs!TXID                 ' トランザクションID
                End If
                If IsNull(rs!REGDATE) = True Then
                Else
                    .REGDATE = rs!REGDATE           ' 登録日付
                End If
                If IsNull(rs!SUMMITSENDFLAG) = True Then
                    .SUMMITSENDFLAG = vbNullString
                Else
                    .SUMMITSENDFLAG = rs!SUMMITSENDFLAG ' SUMIT送信フラグ
                End If
                If IsNull(rs!SENDFLAG) = True Then
                    .SENDFLAG = vbNullString
                Else
                    .SENDFLAG = rs!SENDFLAG         ' 送信フラグ
                End If
                If IsNull(rs!SENDDATE) = True Then
                Else
                    .SENDDATE = rs!SENDDATE         ' 送信日付
                End If
                If IsNull(rs!HREJCODE) = True Then
                    .HREJCODE = vbNullString
                Else
                    .HREJCODE = rs!HREJCODE         ' 不良理由コード
                End If
                If IsNull(rs!UPDPROC) = True Then
                    .UPDPROC = vbNullString
                Else
                    .UPDPROC = rs!UPDPROC           ' 更新工程
                End If
                If IsNull(rs!UPDDATE) = True Then
                Else
                    .UPDDATE = rs!UPDDATE           ' 更新日付
                End If
                If IsNull(rs!MREVNUM) = True Then
                    .REVNUM = 0
                Else
                    .REVNUM = rs!MREVNUM             ' 製品番号改訂番号
                End If
                If IsNull(rs!Mfactory) = True Then
                    .factory = vbNullString
                Else
                    .factory = rs!Mfactory           ' 工場
                End If
                If IsNull(rs!Mopecond) = True Then
                    .opecond = vbNullString
                Else
                    .opecond = rs!Mopecond           ' 操業条件
                End If
                If IsNull(rs!KANKBN) = True Then
                    .KANKBN = vbNullString
                Else
                    .KANKBN = rs!KANKBN             ' 完了区分
                End If
                If IsNull(rs!NREJCODE) = True Then
                    .NREJCODE = vbNullString
                Else
                    .NREJCODE = rs!NREJCODE         ' 抜試返答理由コード
                End If
            End With

            '最終ﾚｺｰﾄﾞ
            rs.MoveLast                             '最終ﾚｺｰﾄﾞに移動
            intUCount = intUCount + 1
            ReDim Preserve tExamine(intUCount)    '配列の再定義
            With tExamine(intUCount)
                If IsNull(rs!LOTID) = True Then
                    .LOTID = vbNullString           ' ブロックID
                Else
                    .LOTID = rs!LOTID
                End If
                If IsNull(rs!MSXLID) = True Then
                    .SXLID = vbNullString
                Else
                    .SXLID = rs!MSXLID               ' SXLID
                End If
                .MinMax = 1                          ' 0:MIN 1:MAX
                If IsNull(rs!BLOCKSEQ) = True Then
                    .BLOCKSEQ = vbNullString
                Else
                    .BLOCKSEQ = rs!BLOCKSEQ         ' ブロック内連番
                End If
                If IsNull(rs!WFSTA) = True Then
                    .WFSTA = vbNullString
                Else
                    .WFSTA = rs!WFSTA               ' WF状態
                End If
                If IsNull(rs!mhinban) = True Then
                    .hinban = vbNullString
                Else
                    .hinban = rs!mhinban             ' 品番
                End If
                If IsNull(rs!RTOP_POS) = True Then
                    .RTOP_POS = 0
                Else
                    .RTOP_POS = Int(CDbl(rs.Fields("RTOP_POS")) + dblEPS)           ' 論理ブロック内位置   'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                End If
                If IsNull(rs!RITOP_POS) = True Then
                    .RITOP_POS = 0
                Else
                    .RITOP_POS = Int(CDbl(rs.Fields("RITOP_POS")) + dblEPS)        ' 論理結晶内位置    'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                End If
                If IsNull(rs!MSMPLEID) = True Then
                    .SMPLEID = vbNullString
                Else
                    .SMPLEID = rs!MSMPLEID           ' 抜試位置
                    .SMPLEID = tblsmp(2).SMPLID      ' 抜試位置    2003/10/26 SystemBrain
                End If
                If IsNull(rs!SHAFLAG) = True Then
                    .SHAFLAG = vbNullString
                Else
                    .SHAFLAG = rs!SHAFLAG            ' サンプルフラグ
                    If Trim(.SHAFLAG) = "1" Then
                        If Trim(.SMPLEID) = vbNullString Then   'add 2003/06/24 hitec)matsumoto サンプルフラグが
                            f_cmbc039_3.cmdF(6).Enabled = False
                            f_cmbc039_3.cmdF(7).Enabled = False
                            f_cmbc039_3.cmdF(8).Enabled = False
                            f_cmbc039_3.cmdF(9).Enabled = False
                            f_cmbc039_3.cmdF(10).Enabled = False
                            f_cmbc039_3.cmdF(12).Enabled = False
                            DBDRV_MIN_MAX_SEQGET = FUNCTION_RETURN_FAILURE
                            f_cmbc039_3.lblMsg.Caption = GetMsgStr("ENSP4", "Y011")
                            rs.Close
                            Exit Function
                        End If
                    End If
                End If
                If IsNull(rs!INDTM) = True Then
                Else
                    .INDTM = rs!INDTM               ' ウェハーセンター入庫日時
                End If
                If IsNull(rs!BASKETID) = True Then
                    .BASKETID = vbNullString
                Else
                    .BASKETID = rs!BASKETID         ' バスケットID
                End If
                If IsNull(rs!SLOTNO) = True Then
                    .SLOTNO = vbNullString
                Else
                    .SLOTNO = rs!SLOTNO             ' スロットNO
                End If
                If IsNull(rs!CURRWPCS) = True Then
                    .CURRWPCS = vbNullString
                Else
                    .CURRWPCS = rs!CURRWPCS         ' ウェハー枚数
                End If
                If IsNull(rs!EXISTFLG) = True Then
                    .EXISTFLG = vbNullString
                Else
                    .EXISTFLG = rs!EXISTFLG         ' 存在フラグ
                End If
                If IsNull(rs!TOP_POS) = True Then
                    .TOP_POS = vbNullString
                Else
                    .TOP_POS = Int(CDbl(rs!TOP_POS) / 10 + 0.9 + dblEPS)        ' ブロックのTOPからの位置  'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                End If
                If IsNull(rs!REJCAT) = True Then
                    .REJCAT = vbNullString
                Else
                    .REJCAT = rs!REJCAT             ' 欠落理由
                End If
                If IsNull(rs!TXID) = True Then
                    .TXID = vbNullString
                Else
                    .TXID = rs!TXID                 ' トランザクションID
                End If
                If IsNull(rs!REGDATE) = True Then
                Else
                    .REGDATE = rs!REGDATE           ' 登録日付
                End If
                If IsNull(rs!SUMMITSENDFLAG) = True Then
                    .SUMMITSENDFLAG = vbNullString
                Else
                    .SUMMITSENDFLAG = rs!SUMMITSENDFLAG ' SUMIT送信フラグ
                End If
                If IsNull(rs!SENDFLAG) = True Then
                    .SENDFLAG = vbNullString
                Else
                    .SENDFLAG = rs!SENDFLAG         ' 送信フラグ
                End If
                If IsNull(rs!SENDDATE) = True Then
                Else
                    .SENDDATE = rs!SENDDATE         ' 送信日付
                End If
                If IsNull(rs!HREJCODE) = True Then
                    .HREJCODE = vbNullString
                Else
                    .HREJCODE = rs!HREJCODE         ' 不良理由コード
                End If
                If IsNull(rs!UPDPROC) = True Then
                    .UPDPROC = vbNullString
                Else
                    .UPDPROC = rs!UPDPROC           ' 更新工程
                End If
                If IsNull(rs!UPDDATE) = True Then
                Else
                    .UPDDATE = rs!UPDDATE           ' 更新日付
                End If
                If IsNull(rs!MREVNUM) = True Then
                    .REVNUM = 0
                Else
                    .REVNUM = rs!MREVNUM             ' 製品番号改訂番号
                End If
                If IsNull(rs!Mfactory) = True Then
                    .factory = vbNullString
                Else
                    .factory = rs!Mfactory           ' 工場
                End If
                If IsNull(rs!Mopecond) = True Then
                    .opecond = vbNullString
                Else
                    .opecond = rs!Mopecond           ' 操業条件
                End If
                If IsNull(rs!KANKBN) = True Then
                    .KANKBN = vbNullString
                Else
                    .KANKBN = rs!KANKBN             ' 完了区分
                End If
                If IsNull(rs!NREJCODE) = True Then
                    .NREJCODE = vbNullString
                Else
                    .NREJCODE = rs!NREJCODE         ' 抜試返答理由コード
                End If
            End With
            intUCount = intUCount + 1
            rs.MoveNext
        Loop
    Next
    '’ループ終了

    DBDRV_MIN_MAX_SEQGET = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    DBDRV_MIN_MAX_SEQGET = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    関数名        : DBDRV_BLOCKIDGET
'*
'*    処理概要      : 1.抜試指示 ブロックＩＤ(SXLが他のブロックに跨る場合）を取得
'*
'*    パラメータ    : 変数名        ,IO ,型                            ,説明
'*                    なし
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function DBDRV_BLOCKIDGET() As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim i           As Long
    Dim lngCnt      As Long
    Dim sDBName     As String
    Dim intUCount   As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_BLOCKIDGET"

    sDBName = "(V001)"

    ' SXLIDの取得
    sSQL = "select"
    sSQL = sSQL & " CRYNUMCA,SXLIDCA"
    sSQL = sSQL & " from XSDCA "
    sSQL = sSQL & " where SXLIDCA ='" & tSXLID(0).SXLID & "'"
    sSQL = sSQL & "   and LIVKCA = '0'"
    sSQL = sSQL & "   ORDER BY CRYNUMCA,SXLIDCA"  'add 2003/04/30 hitec)matsumoto

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    '''抽出レコードが存在ならば該当
    If rs.EOF Then
        DBDRV_BLOCKIDGET = FUNCTION_RETURN_FAILURE
        Exit Function
    Else
        '’抽出レコードをすべて取得（ループ）
        rs.MoveFirst  '先頭ﾚｺｰﾄﾞに移動
        intUCount = 0
        Do While Not rs.EOF
            '’配列にその組み合わせを追加する
            If intUCount = 0 Then  '0（まだロットを入れていない状態）
                With tSXLID(intUCount)
                    .SXLID = rs.Fields("SXLIDCA")         'SXLIDCA
                    .LOTID = rs.Fields("CRYNUMCA")
                End With
            Else    '対象ロットが複数あった時
                ReDim Preserve tSXLID(intUCount)  '配列の再定義
                With tSXLID(intUCount)
                    .SXLID = rs.Fields("SXLIDCA")         'SXLIDCA
                    .LOTID = rs.Fields("CRYNUMCA")
                End With
            End If
            intUCount = intUCount + 1
            rs.MoveNext
        Loop
        rs.Close
    End If

    DBDRV_BLOCKIDGET = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    Resume proc_exit
End Function

'*******************************************************************************
'*    関数名        : DVDRV_KENSA_KOUMOKU
'*
'*    処理概要      : 1.抜試指示　検査項目を取得
'*                      (SXLID,BLOCKID→最大、最小（ブロックＰで判定）のデータを取得する)
'*    パラメータ    : 変数名        ,IO ,型               ,説明
'*                    tKensa        ,I  ,typ_XSDCW        ,新サンプル管理（SXL）
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function DVDRV_KENSA_KOUMOKU(tKensa() As typ_XSDCW) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim i           As Long
    Dim sDBName     As String
    Dim intUCount   As Integer
    Dim tHIN        As tFullHinban
    Dim sOT1        As String
    Dim sOT2        As String
    Dim sMAI1       As String      '04/07/16
    Dim sMAI2       As String
    Dim rtn         As FUNCTION_RETURN
    Dim intIdx      As Integer
    Dim intCnt      As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DVDRV_KENSA_KOUMOKU"

    sDBName = "(V001)"
    intUCount = UBound(tSXLID)
    ReDim tKensa((intUCount * 2) + 1)            '領域再定義
    intIdx = 0

    '’ループ開始
    For i = 0 To intUCount
        ' SXLIDの取得
        If Trim(tSXLID(i).SXLID) <> "" Then
            sSQL = "select "
            sSQL = sSQL & "SXLIDCW,"
            sSQL = sSQL & "SMPKBNCW,"
            sSQL = sSQL & "TBKBNCW,"
            sSQL = sSQL & "REPSMPLIDCW,"
            sSQL = sSQL & "XTALCW,"
            sSQL = sSQL & "INPOSCW,"
            sSQL = sSQL & "HINBCW,"
            sSQL = sSQL & "REVNUMCW,"
            sSQL = sSQL & "FACTORYCW,"
            sSQL = sSQL & "OPECW,"
            sSQL = sSQL & "KTKBNCW,"
            sSQL = sSQL & "SMCRYNUMCW,"
            sSQL = sSQL & "WFSMPLIDRSCW,"
            sSQL = sSQL & "WFSMPLIDRS1CW,"
            sSQL = sSQL & "WFSMPLIDRS2CW,"
            sSQL = sSQL & "WFINDRSCW,"
            sSQL = sSQL & "WFRESRS1CW,"
            sSQL = sSQL & "WFSMPLIDOICW,"
            sSQL = sSQL & "WFINDOICW,"
            sSQL = sSQL & "WFRESOICW,"
            sSQL = sSQL & "WFSMPLIDB1CW,"
            sSQL = sSQL & "WFINDB1CW,"
            sSQL = sSQL & "WFRESB1CW,"
            sSQL = sSQL & "WFSMPLIDB2CW,"
            sSQL = sSQL & "WFINDB2CW,"
            sSQL = sSQL & "WFRESB2CW,"
            sSQL = sSQL & "WFSMPLIDB3CW,"
            sSQL = sSQL & "WFINDB3CW,"
            sSQL = sSQL & "WFRESB3CW,"
            sSQL = sSQL & "WFSMPLIDL1CW,"
            sSQL = sSQL & "WFINDL1CW,"
            sSQL = sSQL & "WFRESL1CW,"
            sSQL = sSQL & "WFSMPLIDL2CW,"
            sSQL = sSQL & "WFINDL2CW,"
            sSQL = sSQL & "WFRESL2CW,"
            sSQL = sSQL & "WFSMPLIDL3CW,"
            sSQL = sSQL & "WFINDL3CW,"
            sSQL = sSQL & "WFRESL3CW,"
            sSQL = sSQL & "WFSMPLIDL4CW,"
            sSQL = sSQL & "WFINDL4CW,"
            sSQL = sSQL & "WFRESL4CW,"
            sSQL = sSQL & "WFSMPLIDDSCW,"
            sSQL = sSQL & "WFINDDSCW,"
            sSQL = sSQL & "WFRESDSCW,"
            sSQL = sSQL & "WFSMPLIDDZCW,"
            sSQL = sSQL & "WFINDDZCW,"
            sSQL = sSQL & "WFRESDZCW,"
            sSQL = sSQL & "WFSMPLIDSPCW,"
            sSQL = sSQL & "WFINDSPCW,"
            sSQL = sSQL & "WFRESSPCW,"
            sSQL = sSQL & "WFSMPLIDDO1CW,"
            sSQL = sSQL & "WFINDDO1CW,"
            sSQL = sSQL & "WFRESDO1CW,"
            sSQL = sSQL & "WFSMPLIDDO2CW,"
            sSQL = sSQL & "WFINDDO2CW,"
            sSQL = sSQL & "WFRESDO2CW,"
            sSQL = sSQL & "WFSMPLIDDO3CW,"
            sSQL = sSQL & "WFINDDO3CW,"
            sSQL = sSQL & "WFRESDO3CW,"
            sSQL = sSQL & "WFSMPLIDOT1CW,"
            sSQL = sSQL & "WFSMPLIDOT2CW,"
            sSQL = sSQL & "NVL(WFINDOT1CW,'0') as DOT1,"            ' 状態FLG（OT1)
            sSQL = sSQL & "NVL(WFRESOT1CW,'0') as SOT1,"            ' 実績FLG（OT1)
            sSQL = sSQL & "NVL(WFINDOT2CW,'0') as DOT2,"            ' 状態FLG（OT2)
            sSQL = sSQL & "NVL(WFRESOT2CW,'0') as SOT2,"            ' 実績FLG（OT2)
            sSQL = sSQL & "WFSMPLIDAOICW,"
            sSQL = sSQL & "WFINDAOICW,"
            sSQL = sSQL & "WFRESAOICW,"
            sSQL = sSQL & "SMPLNUMCW,"
            sSQL = sSQL & "SMPLPATCW,"
            sSQL = sSQL & "TSTAFFCW,"
            sSQL = sSQL & "TDAYCW,"
            sSQL = sSQL & "KSTAFFCW,"
            sSQL = sSQL & "KDAYCW,"
            sSQL = sSQL & "SNDKCW,"
            sSQL = sSQL & "SNDDAYCW,"
            sSQL = sSQL & "WFSMPLIDGDCW,"     '' GD追加　05/02/17 ooba START ===========>
            sSQL = sSQL & "WFINDGDCW,"
            sSQL = sSQL & "WFRESGDCW,"
            sSQL = sSQL & "WFHSGDCW"          '' GD追加　05/02/17 ooba END =============>
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
            sSQL = sSQL & ",EPSMPLIDB1CW, "
            sSQL = sSQL & "EPINDB1CW, "
            sSQL = sSQL & "EPRESB1CW, "
            sSQL = sSQL & "EPSMPLIDB2CW, "
            sSQL = sSQL & "EPINDB2CW, "
            sSQL = sSQL & "EPRESB2CW, "
            sSQL = sSQL & "EPSMPLIDB3CW, "
            sSQL = sSQL & "EPINDB3CW, "
            sSQL = sSQL & "EPRESB3CW, "
            sSQL = sSQL & "EPSMPLIDL1CW, "
            sSQL = sSQL & "EPINDL1CW, "
            sSQL = sSQL & "EPRESL1CW, "
            sSQL = sSQL & "EPSMPLIDL2CW, "
            sSQL = sSQL & "EPINDL2CW, "
            sSQL = sSQL & "EPRESL2CW, "
            sSQL = sSQL & "EPSMPLIDL3CW, "
            sSQL = sSQL & "EPINDL3CW, "
            sSQL = sSQL & "EPRESL3CW "
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
            sSQL = sSQL & " from XSDCW "
            sSQL = sSQL & " where SXLIDCW ='" & tSXLID(i).SXLID & "'"
            sSQL = sSQL & "   and LIVKCW  ='0'"                           ' 生死区分は必ず確認する事
            sSQL = sSQL & " order by INPOSCW"

            Debug.Print sSQL
            Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

            '''抽出レコードが存在ならば該当
            If Not rs.EOF Then
                intCnt = 0
                Do While Not rs.EOF
                    intCnt = intCnt + 1
                    ' ３件目以降が存在する場合エラー
                    If intCnt > 2 Then
                        Exit Do
                    End If

                    With tKensa(intIdx)
                        .SXLIDCW = rs("SXLIDCW")
                        .SMPKBNCW = rs!SMPKBNCW
                        .TBKBNCW = rs!TBKBNCW
                        .REPSMPLIDCW = rs!REPSMPLIDCW
                        .XTALCW = rs!XTALCW
                        .INPOSCW = rs!INPOSCW
                        .HINBCW = rs!HINBCW
                        .REVNUMCW = rs!REVNUMCW
                        .FACTORYCW = rs!FACTORYCW
                        .OPECW = rs!OPECW
                        .KTKBNCW = rs!KTKBNCW
                        .SMCRYNUMCW = rs!SMCRYNUMCW
                        If IsNull(rs!WFSMPLIDRSCW) Then
                            .WFSMPLIDRSCW = ""
                        Else
                            .WFSMPLIDRSCW = rs!WFSMPLIDRSCW
                        End If
                        If IsNull(rs("WFSMPLIDRS1CW")) Then
                            .WFSMPLIDRS1CW = ""
                        Else
                            .WFSMPLIDRS1CW = rs("WFSMPLIDRS1CW")
                        End If
                        If IsNull(rs!WFSMPLIDRS2CW) Then
                            .WFSMPLIDRS2CW = ""
                        Else
                            .WFSMPLIDRS2CW = rs!WFSMPLIDRS2CW
                        End If
                        If IsNull(rs!WFINDRSCW) Then
                            .WFINDRSCW = ""
                        Else
                            .WFINDRSCW = rs!WFINDRSCW
                        End If
                        If IsNull(rs!WFRESRS1CW) Then
                            .WFRESRS1CW = ""
                        Else
                            .WFRESRS1CW = rs!WFRESRS1CW
                        End If
                        If IsNull(rs!WFSMPLIDOICW) Then
                            .WFSMPLIDOICW = ""
                        Else
                            .WFSMPLIDOICW = rs!WFSMPLIDOICW
                        End If
                        If IsNull(rs!WFINDOICW) Then
                            .WFINDOICW = ""
                        Else
                            .WFINDOICW = rs!WFINDOICW
                        End If
                        If IsNull(rs!WFRESOICW) Then
                            .WFRESOICW = ""
                        Else
                            .WFRESOICW = rs!WFRESOICW
                        End If
                        If IsNull(rs!WFSMPLIDB1CW) Then
                            .WFSMPLIDB1CW = ""
                        Else
                            .WFSMPLIDB1CW = rs!WFSMPLIDB1CW
                        End If
                        If IsNull(rs!WFINDB1CW) Then
                            .WFINDB1CW = ""
                        Else
                            .WFINDB1CW = rs!WFINDB1CW
                        End If
                        If IsNull(rs!WFRESB1CW) Then
                            .WFRESB1CW = ""
                        Else
                            .WFRESB1CW = rs!WFRESB1CW
                        End If
                        If IsNull(rs!WFSMPLIDB2CW) Then
                            .WFSMPLIDB2CW = ""
                        Else
                            .WFSMPLIDB2CW = rs!WFSMPLIDB2CW
                        End If
                        If IsNull(rs!WFINDB2CW) Then
                            .WFINDB2CW = ""
                        Else
                            .WFINDB2CW = rs!WFINDB2CW
                        End If
                        If IsNull(rs!WFRESB2CW) Then
                            .WFRESB2CW = ""
                        Else
                            .WFRESB2CW = rs!WFRESB2CW
                        End If
                        If IsNull(rs!WFSMPLIDB3CW) Then
                            .WFSMPLIDB3CW = ""
                        Else
                            .WFSMPLIDB3CW = rs!WFSMPLIDB3CW
                        End If
                        If IsNull(rs!WFINDB3CW) Then
                            .WFINDB3CW = ""
                        Else
                            .WFINDB3CW = rs!WFINDB3CW
                        End If
                        If IsNull(rs!WFRESB3CW) Then
                            .WFRESB3CW = ""
                        Else
                            .WFRESB3CW = rs!WFRESB3CW
                        End If
                        If IsNull(rs!WFSMPLIDL1CW) Then
                            .WFSMPLIDL1CW = ""
                        Else
                            .WFSMPLIDL1CW = rs!WFSMPLIDL1CW
                        End If
                        If IsNull(rs!WFINDL1CW) Then
                            .WFINDL1CW = ""
                        Else
                            .WFINDL1CW = rs!WFINDL1CW
                        End If
                        If IsNull(rs!WFRESL1CW) Then
                            .WFRESL1CW = ""
                        Else
                            .WFRESL1CW = rs!WFRESL1CW
                        End If
                        If IsNull(rs!WFSMPLIDL2CW) Then
                            .WFSMPLIDL2CW = ""
                        Else
                            .WFSMPLIDL2CW = rs!WFSMPLIDL2CW
                        End If
                        If IsNull(rs!WFINDL2CW) Then
                            .WFINDL2CW = ""
                        Else
                            .WFINDL2CW = rs!WFINDL2CW
                        End If
                        If IsNull(rs!WFRESL2CW) Then
                            .WFRESL2CW = ""
                        Else
                            .WFRESL2CW = rs!WFRESL2CW
                        End If
                        If IsNull(rs!WFSMPLIDL3CW) Then
                            .WFSMPLIDL3CW = ""
                        Else
                            .WFSMPLIDL3CW = rs!WFSMPLIDL3CW
                        End If
                        If IsNull(rs!WFINDL3CW) Then
                            .WFINDL3CW = ""
                        Else
                            .WFINDL3CW = rs!WFINDL3CW
                        End If
                        If IsNull(rs!WFRESL3CW) Then
                            .WFRESL3CW = ""
                        Else
                            .WFRESL3CW = rs!WFRESL3CW
                        End If
                        If IsNull(rs!WFSMPLIDL4CW) Then
                            .WFSMPLIDL4CW = ""
                        Else
                            .WFSMPLIDL4CW = rs!WFSMPLIDL4CW
                        End If
                        If IsNull(rs!WFINDL4CW) Then
                            .WFINDL4CW = ""
                        Else
                            .WFINDL4CW = rs!WFINDL4CW
                        End If
                        If IsNull(rs!WFRESL4CW) Then
                            .WFRESL4CW = ""
                        Else
                            .WFRESL4CW = rs!WFRESL4CW
                        End If
                        If IsNull(rs!WFSMPLIDDSCW) Then
                            .WFSMPLIDDSCW = ""
                        Else
                            .WFSMPLIDDSCW = rs!WFSMPLIDDSCW
                        End If
                        If IsNull(rs!WFINDDSCW) Then
                            .WFINDDSCW = ""
                        Else
                            .WFINDDSCW = rs!WFINDDSCW
                        End If
                        If IsNull(rs!WFRESDSCW) Then
                            .WFRESDSCW = ""
                        Else
                            .WFRESDSCW = rs!WFRESDSCW
                        End If
                        If IsNull(rs!WFSMPLIDDZCW) Then
                            .WFSMPLIDDZCW = ""
                        Else
                            .WFSMPLIDDZCW = rs!WFSMPLIDDZCW
                        End If
                        If IsNull(rs!WFINDDZCW) Then
                            .WFINDDZCW = ""
                        Else
                            .WFINDDZCW = rs!WFINDDZCW
                        End If
                        If IsNull(rs!WFRESDZCW) Then
                            .WFRESDZCW = ""
                        Else
                            .WFRESDZCW = rs!WFRESDZCW
                        End If
                        If IsNull(rs!WFSMPLIDSPCW) Then
                            .WFSMPLIDSPCW = ""
                        Else
                            .WFSMPLIDSPCW = rs!WFSMPLIDSPCW
                        End If
                        If IsNull(rs!WFINDSPCW) Then
                            .WFINDSPCW = ""
                        Else
                            .WFINDSPCW = rs!WFINDSPCW
                        End If
                        If IsNull(rs!WFRESSPCW) Then
                            .WFRESSPCW = ""
                        Else
                            .WFRESSPCW = rs!WFRESSPCW
                        End If
                        If IsNull(rs!WFSMPLIDDO1CW) Then
                            .WFSMPLIDDO1CW = ""
                        Else
                            .WFSMPLIDDO1CW = rs!WFSMPLIDDO1CW
                        End If
                        If IsNull(rs!WFINDDO1CW) Then
                            .WFINDDO1CW = ""
                        Else
                            .WFINDDO1CW = rs!WFINDDO1CW
                        End If
                        If IsNull(rs!WFRESDO1CW) Then
                            .WFRESDO1CW = ""
                        Else
                            .WFRESDO1CW = rs!WFRESDO1CW
                        End If
                        If IsNull(rs!WFSMPLIDDO2CW) Then
                            .WFSMPLIDDO2CW = ""
                        Else
                            .WFSMPLIDDO2CW = rs!WFSMPLIDDO2CW
                        End If
                        If IsNull(rs!WFINDDO2CW) Then
                            .WFINDDO2CW = ""
                        Else
                            .WFINDDO2CW = rs!WFINDDO2CW
                        End If
                        If IsNull(rs!WFRESDO2CW) Then
                            .WFRESDO2CW = ""
                        Else
                            .WFRESDO2CW = rs!WFRESDO2CW
                        End If
                        If IsNull(rs!WFSMPLIDDO3CW) Then
                            .WFSMPLIDDO3CW = ""
                        Else
                            .WFSMPLIDDO3CW = rs!WFSMPLIDDO3CW
                        End If
                        If IsNull(rs!WFINDDO3CW) Then
                            .WFINDDO3CW = ""
                        Else
                            .WFINDDO3CW = rs!WFINDDO3CW
                        End If
                        If IsNull(rs!WFRESDO3CW) Then
                            .WFRESDO3CW = ""
                        Else
                            .WFRESDO3CW = rs!WFRESDO3CW
                        End If
                        If IsNull(rs!WFSMPLIDOT1CW) Then
                            .WFSMPLIDOT1CW = ""
                        Else
                            .WFSMPLIDOT1CW = rs!WFSMPLIDOT1CW
                        End If
                        If IsNull(rs!sOT1) Then
                            .WFRESOT1CW = ""
                        Else
                            .WFRESOT1CW = rs!sOT1
                        End If
                        If IsNull(rs!WFSMPLIDOT2CW) Then
                            .WFSMPLIDOT2CW = ""
                        Else
                            .WFSMPLIDOT2CW = rs!WFSMPLIDOT2CW
                        End If
                        If IsNull(rs!sOT2) Then
                            .WFRESOT2CW = ""
                        Else
                            .WFRESOT2CW = rs!sOT2
                        End If

                        tHIN.hinban = .HINBCW
                        tHIN.factory = .FACTORYCW
                        tHIN.mnorevno = .REVNUMCW
                        tHIN.opecond = .OPECW

                        rtn = scmzc_getE036(tHIN, sOT1, sOT2, sMAI1, sMAI2)
                        If rtn = FUNCTION_RETURN_FAILURE Then
                            rs.Close
                            DVDRV_KENSA_KOUMOKU = FUNCTION_RETURN_FAILURE
                            GoTo proc_exit
                        End If
                        If sOT1 = "1" Then
                            .WFINDOT1CW = rs!DOT1
                        Else
                            .WFINDOT1CW = 0
                        End If
                        If sOT2 = "1" Then
                            .WFINDOT2CW = rs!DOT2
                        Else
                            .WFINDOT2CW = 0
                        End If

                        If IsNull(rs!WFSMPLIDAOICW) Then
                            .WFSMPLIDAOICW = ""
                        Else
                            .WFSMPLIDAOICW = rs!WFSMPLIDAOICW
                        End If
                        If IsNull(rs!WFINDAOICW) Then
                            .WFINDAOICW = ""
                        Else
                            .WFINDAOICW = rs!WFINDAOICW
                        End If
                        If IsNull(rs!WFRESAOICW) Then
                            .WFRESAOICW = ""
                        Else
                            .WFRESAOICW = rs!WFRESAOICW
                        End If
                        If IsNull(rs!SMPLNUMCW) Then
                            .SMPLNUMCW = 0
                        Else
                            .SMPLNUMCW = rs!SMPLNUMCW
                        End If
                        If IsNull(rs!SMPLPATCW) Then
                            .SMPLPATCW = ""
                        Else
                            .SMPLPATCW = rs!SMPLPATCW
                        End If
                        If IsNull(rs!TSTAFFCW) Then
                            .TSTAFFCW = ""
                        Else
                            .TSTAFFCW = rs!TSTAFFCW
                        End If
                        If IsNull(rs!TDAYCW) Then
                            .TDAYCW = "2003/10/3"
                        Else
                            .TDAYCW = rs!TDAYCW
                        End If
                        If IsNull(rs!KSTAFFCW) Then
                            .KSTAFFCW = ""
                        Else
                            .KSTAFFCW = rs!KSTAFFCW
                        End If
                        If IsNull(rs!KDAYCW) Then
                            .KDAYCW = "2003/10/3"
                        Else
                            .KDAYCW = rs!KDAYCW
                        End If
                        If IsNull(rs!SNDKCW) Then
                            .SNDKCW = ""
                        Else
                            .SNDKCW = rs!SNDKCW
                        End If
                        If IsNull(rs!SNDDAYCW) Then
                            .SNDDAYCW = "2003/10/3"
                        Else
                            .SNDDAYCW = rs!SNDDAYCW
                        End If
                        '' GD追加　05/02/17 ooba START ==================>
                        If IsNull(rs!WFSMPLIDGDCW) Then
                            .WFSMPLIDGDCW = ""
                        Else
                            .WFSMPLIDGDCW = rs!WFSMPLIDGDCW
                        End If
                        If IsNull(rs!WFINDGDCW) Then
                            .WFINDGDCW = ""
                        Else
                            .WFINDGDCW = rs!WFINDGDCW
                        End If
                        If IsNull(rs!WFRESGDCW) Then
                            .WFRESGDCW = ""
                        Else
                            .WFRESGDCW = rs!WFRESGDCW
                        End If
                        If IsNull(rs!WFHSGDCW) Then
                            .WFHSGDCW = ""
                        Else
                            .WFHSGDCW = rs!WFHSGDCW
                        End If
                        '' GD追加　05/02/17 ooba END ====================>

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
                        ' BMD1E
                        If IsNull(rs!EPSMPLIDB1CW) Then
                            .EPSMPLIDB1CW = ""
                        Else
                            .EPSMPLIDB1CW = rs!EPSMPLIDB1CW
                        End If
                        If IsNull(rs!EPINDB1CW) Then
                            .EPINDB1CW = ""
                        Else
                            .EPINDB1CW = rs!EPINDB1CW
                        End If
                        If IsNull(rs!EPRESB1CW) Then
                            .EPRESB1CW = ""
                        Else
                            .EPRESB1CW = rs!EPRESB1CW
                        End If
                        ' BMD2E
                        If IsNull(rs!EPSMPLIDB2CW) Then
                            .EPSMPLIDB2CW = ""
                        Else
                            .EPSMPLIDB2CW = rs!EPSMPLIDB2CW
                        End If
                        If IsNull(rs!EPINDB2CW) Then
                            .EPINDB2CW = ""
                        Else
                            .EPINDB2CW = rs!EPINDB2CW
                        End If
                        If IsNull(rs!EPRESB2CW) Then
                            .EPRESB2CW = ""
                        Else
                            .EPRESB2CW = rs!EPRESB2CW
                        End If
                        ' BMD3E
                        If IsNull(rs!EPSMPLIDB3CW) Then
                            .EPSMPLIDB3CW = ""
                        Else
                            .EPSMPLIDB3CW = rs!EPSMPLIDB3CW
                        End If
                        If IsNull(rs!EPINDB3CW) Then
                            .EPINDB3CW = ""
                        Else
                            .EPINDB3CW = rs!EPINDB3CW
                        End If
                        If IsNull(rs!EPRESB3CW) Then
                            .EPRESB3CW = ""
                        Else
                            .EPRESB3CW = rs!EPRESB3CW
                        End If
                        ' OSF1E
                        If IsNull(rs!EPSMPLIDL1CW) Then
                            .EPSMPLIDL1CW = ""
                        Else
                            .EPSMPLIDL1CW = rs!EPSMPLIDL1CW
                        End If
                        If IsNull(rs!EPINDL1CW) Then
                            .EPINDL1CW = ""
                        Else
                            .EPINDL1CW = rs!EPINDL1CW
                        End If
                        If IsNull(rs!EPRESL1CW) Then
                            .EPRESL1CW = ""
                        Else
                            .EPRESL1CW = rs!EPRESL1CW
                        End If
                        ' OSF2E
                        If IsNull(rs!EPSMPLIDL2CW) Then
                            .EPSMPLIDL2CW = ""
                        Else
                            .EPSMPLIDL2CW = rs!EPSMPLIDL2CW
                        End If
                        If IsNull(rs!EPINDL2CW) Then
                            .EPINDL2CW = ""
                        Else
                            .EPINDL2CW = rs!EPINDL2CW
                        End If
                        If IsNull(rs!EPRESL2CW) Then
                            .EPRESL2CW = ""
                        Else
                            .EPRESL2CW = rs!EPRESL2CW
                        End If
                        ' OSF3E
                        If IsNull(rs!EPSMPLIDL3CW) Then
                            .EPSMPLIDL3CW = ""
                        Else
                            .EPSMPLIDL3CW = rs!EPSMPLIDL3CW
                        End If
                        If IsNull(rs!EPINDL3CW) Then
                            .EPINDL3CW = ""
                        Else
                            .EPINDL3CW = rs!EPINDL3CW
                        End If
                        If IsNull(rs!EPRESL3CW) Then
                            .EPRESL3CW = ""
                        Else
                            .EPRESL3CW = rs!EPRESL3CW
                        End If
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
                    End With

                    intIdx = intIdx + 1
                    rs.MoveNext
                Loop
                rs.Close

                ' 取得件数が２件でない場合エラー
                If intCnt <> 2 Then
                    f_cmbc039_3.lblMsg.Caption = GetMsgStr("ENSP2")
                    DVDRV_KENSA_KOUMOKU = FUNCTION_RETURN_FAILURE
                    Exit Function
                End If
            Else
                 f_cmbc039_3.lblMsg.Caption = GetMsgStr("ENSP2")
                DVDRV_KENSA_KOUMOKU = FUNCTION_RETURN_FAILURE
                Exit Function
            End If
        End If
    Next i
    '’ループ終了

    DVDRV_KENSA_KOUMOKU = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    GoTo proc_exit
End Function

'*****************************************************************************************************************
'*    関数名        : DBDRV_GET_WFMAP
'*
'*    処理概要      : 1.抜試指示 入力したブロックＰから、該当ＷＦを検索
'*
'*    パラメータ    : 変数名        ,IO ,型      ,説明
'*                    sBlkId        ,I  ,String  ,ブロックID
'*                    sSXLID        ,I  ,String  ,SXL-ID
'*                    iBlkP         ,I  ,Integer ,ブロックP
'*                    sBlkP         ,I  ,Variant ,Spreadに記載されているブロックP
'*                    sKessyoP      ,O  ,Variant ,論理結晶内位置
'*                    sNextIngotP   ,O  ,String  ,次結晶位置
'*                    sBlkSeq       ,O  ,Variant ,ブロック内連番
'*                    sBlkSeq2      ,O  ,Variant ,ブロック内連番
'*                    sSmpId1       ,O  ,Variant ,サンプルID
'*                    sSmpId2       ,O  ,Variant ,サンプルID
'*                    iNextBlkP     ,O  ,Integer ,次ブロックP
'*                    vWfNum        ,O  ,Variant ,Wafer数
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*****************************************************************************************************************
Public Function DBDRV_GET_WFMAP(ByVal sBlkId As String, ByVal sSXLID As String, ByVal iBlkP As Integer, _
                                ByRef sBlkP As Variant, ByRef sKessyoP As Variant, ByRef sNextIngotP As String, _
                                ByRef sBlkSeq As Variant, ByRef sBlkSeq2 As Variant, ByRef sSmpId1 As Variant, _
                                ByRef sSmpId2 As Variant, ByRef iNextBlkP As Integer, ByRef vWfNum As Variant) _
                                    As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim i           As Long
    Dim lngCnt      As Long
    Dim sDBName     As String
    Dim intLoopCnt  As Integer
    Dim dblChkBlkP  As Double
    Dim intChkBlkP  As Integer
    Dim intTopPos   As Integer
    Dim sAddSmpId1  As String
    Dim sAddSmpId2  As String
    Dim vBlkId      As Variant
    Dim intBlkflg   As Integer
    Dim dblWFLen    As Double
    Dim iRtn        As FUNCTION_RETURN
    Dim intSearchWf As Integer
    Dim dblEPS      As Double

    dblEPS = 0.000001        'εの設定 'add 2003/06/13 hitec)matsumoto

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_GET_WFMAP"

    sDBName = "(Y011)"
    i = 0

    sSQL = "select "
    sSQL = sSQL & "LOTID,"                ' ブロックID"
    sSQL = sSQL & "MSXLID,"                ' SXLID"
    sSQL = sSQL & "blockseq,"             ' ブロック内連番"
    sSQL = sSQL & "WFSTA,"                ' WF状態"
    sSQL = sSQL & "RTOP_POS,"             ' 論理ブロック内位置"
    sSQL = sSQL & "RITOP_POS,"            ' 論理結晶内位置"
    sSQL = sSQL & "MSMPLEID,"              ' 抜試位置"
    sSQL = sSQL & "SHAFLAG,"              ' サンプルフラグ"
    sSQL = sSQL & "TOP_POS"               ' ブロック内位置
    sSQL = sSQL & " from TBCMY011 "
    sSQL = sSQL & " where MSXLID ='" & sSXLID & "'"
    sSQL = sSQL & "   AND LOTID ='" & sBlkId & "'"
    sSQL = sSQL & "   AND TO_NUMBER(WFSTA) <= 1"  'del 2003/04/28 hitec)matsumoto 状態を見ない
    sSQL = sSQL & " ORDER BY BLOCKSEQ ASC"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    intLoopCnt = 0
    vWfNum = 0
    Do While Not rs.EOF
        If IsNull(rs.Fields("RTOP_POS")) = True Then
            dblChkBlkP = 0
        Else
            dblChkBlkP = CDbl(rs.Fields("RTOP_POS"))
        End If
        If (iBlkP < dblChkBlkP) And (dblChkBlkP <= iNextBlkP) Then
            vWfNum = CInt(vWfNum) + 1
        End If
        rs.MoveNext
    Loop
    rs.Close

    sSQL = "select "
    sSQL = sSQL & "LOTID,"                ' ブロックID"
    sSQL = sSQL & "MSXLID,"                ' SXLID"
    sSQL = sSQL & "blockseq,"             ' ブロック内連番"
    sSQL = sSQL & "WFSTA,"                ' WF状態"
    sSQL = sSQL & "RTOP_POS,"             ' 論理ブロック内位置"
    sSQL = sSQL & "RITOP_POS,"            ' 論理結晶内位置"
    sSQL = sSQL & "MSMPLEID,"              ' 抜試位置"
    sSQL = sSQL & "SHAFLAG,"              ' サンプルフラグ"
    sSQL = sSQL & "TOP_POS"               ' ブロック内位置
    sSQL = sSQL & " from TBCMY011 "
    sSQL = sSQL & " where MSXLID ='" & sSXLID & "'"
    sSQL = sSQL & "   AND LOTID ='" & sBlkId & "'"
    sSQL = sSQL & " ORDER BY BLOCKSEQ ASC"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    intLoopCnt = 0
    rs.MoveFirst
    Do While Not rs.EOF
        Select Case Right(sSmpId1, 1)
            Case "T"
                rs.MoveFirst
                'WFの欠落を判定（CW740ではほぼありえない)   'add 2003/05/05 hitec)matsumoto
                If CStr(rs.Fields("WFSTA")) = "4" Then
                    rs.Close
                    DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                    Exit Function
                End If
                If IsNull(rs.Fields("RTOP_POS")) = False Then
                    sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + dblEPS)  'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                End If
                If IsNull(rs.Fields("RITOP_POS")) = False Then
                    sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")) + dblEPS)  'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                End If
                sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
                intTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + dblEPS)  '切り捨て 'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(intTopPos), "000") & "T"
                Exit Do
            Case "U"
                If IsNull(rs.Fields("RTOP_POS")) = False Then
                    dblChkBlkP = CDbl(rs.Fields("RTOP_POS"))
                End If
                If dblChkBlkP > iBlkP Or dblChkBlkP = iBlkP Then
                    If dblChkBlkP > iBlkP Then
                        rs.MovePrevious
                        If IsNull(rs.Fields("BLOCKSEQ")) = True Then    'add 2003/04/28 hitec)matsumoto  NULLの場合（該当WF無し）は、下に検索する
                            Do
                                rs.MoveNext
                                If IsNull(rs.Fields("RTOP_POS")) = False Then
                                    Exit Do
                                End If
                            Loop
                        End If
                        'WFの欠落を判定（CW740ではほぼありえない)   'add 2003/05/05 hitec)matsumoto
                        If CStr(rs.Fields("WFSTA")) = "4" Then
                            rs.Close
                            DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                            Exit Function
                        End If
                        sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + 0.9 + dblEPS)    'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                        sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")) + 0.9 + dblEPS)    'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                        sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
                        intTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + 0.9 + dblEPS) '切り上げ    'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                        sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(intTopPos), "000") & "U"
                        rs.MoveNext
                    ElseIf dblChkBlkP = iBlkP Then
                        'WFの欠落を判定（CW740ではほぼありえない)   'add 2003/05/05 hitec)matsumoto
                        If CStr(rs.Fields("WFSTA")) = "4" Then
                            rs.Close
                            DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                            Exit Function
                        End If
                        sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + 0.9 + dblEPS)    'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                        sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")) + 0.9 + dblEPS)    'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                        sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
                        intTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + 0.9 + dblEPS) '切り上げ    'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                        sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(intTopPos), "000") & "U"
                        rs.MoveNext
                    End If

                    If Not rs.EOF Then
                        If sSmpId2 <> vbNullString Then 'Dのサンプルを作成
                            '0以外は0.1mm引いて切捨て(WF操業:Dは該当位置を含まずに下方向抜取り) 08/11/06 ooba
                            If rs.Fields("TOP_POS") > 0 Then
                                intTopPos = Int(CDbl(rs.Fields("TOP_POS") - 1) / 10 + dblEPS)   '0.1mm引いて切捨て
                            Else
                                intTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + dblEPS)  '切り捨て 'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                            End If
                            sAddSmpId2 = Mid(sBlkId, 1, 12) & Format(CStr(intTopPos), "000") & "D"
                        End If
                        sBlkSeq2 = CStr(rs.Fields("BLOCKSEQ"))
                        sNextIngotP = Int(CDbl(rs.Fields("RITOP_POS")) + dblEPS)   'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                        Exit Do
                    Else
                        '現在のブロックIDの次のブロックIDを取得
                        With f_cmbc039_3.sprExamine
                            intBlkflg = 0
                            For i = 1 To .MaxRows
                                .GetText 1, i, vBlkId
                                If vBlkId <> "" And Len(vBlkId) <> 1 Then    '03/05/31
                                    If intBlkflg = 1 Then
                                        sBlkId = left(sBlkId, 9) & CStr(vBlkId) '次のBLID取得
                                        Exit For
                                    ElseIf Right(sBlkId, 3) = vBlkId Then
                                        intBlkflg = 1
                                    End If
                                End If
                            Next i
                        End With
                        rs.Close

                        sSQL = "select "
                        sSQL = sSQL & "LOTID,"                ' ブロックID"
                        sSQL = sSQL & "MSXLID,"                ' SXLID"
                        sSQL = sSQL & "blockseq,"             ' ブロック内連番"
                        sSQL = sSQL & "WFSTA,"                ' WF状態"
                        sSQL = sSQL & "RTOP_POS,"             ' 論理ブロック内位置"
                        sSQL = sSQL & "RITOP_POS,"            ' 論理結晶内位置"
                        sSQL = sSQL & "MSMPLEID,"              ' 抜試位置"
                        sSQL = sSQL & "SHAFLAG,"              ' サンプルフラグ"
                        sSQL = sSQL & "TOP_POS"               ' ブロック内位置
                        sSQL = sSQL & " from TBCMY011 "
                        sSQL = sSQL & " where MSXLID ='" & sSXLID & "'"
                        sSQL = sSQL & "   AND LOTID ='" & sBlkId & "'"
                        sSQL = sSQL & " ORDER BY BLOCKSEQ ASC"

                        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

                        intLoopCnt = 0
                        rs.MoveFirst
                        'WFの欠落を判定（CW740ではほぼありえない)   'add 2003/05/05 hitec)matsumoto
                        If CStr(rs.Fields("WFSTA")) = "4" Then
                            rs.Close
                            DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                            Exit Function
                        End If
                        If IsNull(rs.Fields("RTOP_POS")) = False Then
                            'WF一枚の長さ取得                                   '2003/04/25 hitec)okazaki
                            iRtn = DBDRV_WFLENGET(sBlkId, dblWFLen)
                            'ブロック先頭の表示位置はWF一枚の長さを引いたもの   '2003/04/25 hitec)okazaki
                            sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) - dblWFLen + dblEPS)   'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                        End If
                        If IsNull(rs.Fields("RITOP_POS")) = False Then
                            sNextIngotP = Int(CDbl(rs.Fields("RITOP_POS")) + dblEPS) 'upd 2003/06/13 hitec)matsumoto
                        End If
                        If sSmpId2 <> vbNullString Then 'Dのサンプルを作成
                            '0以外は0.1mm引いて切捨て(WF操業:Dは該当位置を含まずに下方向抜取り) 08/11/06 ooba
                            If rs.Fields("TOP_POS") > 0 Then
                                intTopPos = Int(CDbl(rs.Fields("TOP_POS") - 1) / 10 + dblEPS)   '0.1mm引いて切捨て
                            Else
                                intTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + dblEPS)  '切り捨て 'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                            End If
                            sAddSmpId2 = Mid(sBlkId, 1, 12) & Format(CStr(intTopPos), "000") & "D"
                        End If
                        sBlkSeq2 = CStr(rs.Fields("BLOCKSEQ"))
                        Exit Do
                    End If
                End If
            Case "D"
                If IsNull(rs.Fields("RTOP_POS")) = False Then
                    dblChkBlkP = CDbl(rs.Fields("RTOP_POS"))
                End If
                If dblChkBlkP > iBlkP Then
                    sNextIngotP = Int(CDbl(rs.Fields("RITOP_POS")) + dblEPS)   'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                    sBlkSeq2 = CStr(rs.Fields("BLOCKSEQ"))
                    '0以外は0.1mm引いて切捨て(WF操業:Dは該当位置を含まずに下方向抜取り) 08/11/06 ooba
                    If rs.Fields("TOP_POS") > 0 Then
                        intTopPos = Int(CDbl(rs.Fields("TOP_POS") - 1) / 10 + dblEPS)   '0.1mm引いて切捨て
                    Else
                        intTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + dblEPS)  '切り捨て 'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                    End If
                    sAddSmpId2 = Mid(sBlkId, 1, 12) & Format(CStr(intTopPos), "000") & "D"   'DなのでsAddSmpId2に入れる
                    rs.MovePrevious
                    'WFの欠落を判定（CW740ではほぼありえない)   'add 2003/05/05 hitec)matsumoto
                    If CStr(rs.Fields("WFSTA")) = "4" Then
                        rs.Close
                        DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                        Exit Function
                    End If
                    If sSmpId2 <> vbNullString Then 'Uのサンプルを作成
                        intTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + 0.9 + dblEPS)   '切り上げ  'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                        sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(intTopPos), "000") & "U"   '"U"なのでsAddSmpId1に入れる
                    End If
                    sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + 0.9 + dblEPS)    'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                    sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
                    sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")) + 0.9 + dblEPS)    'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                    Exit Do
                End If
            Case "B"
                rs.MoveLast
                'WFの欠落を判定（CW740ではほぼありえない)   'add 2003/05/05 hitec)matsumoto
                If CStr(rs.Fields("WFSTA")) = "4" Then
                    rs.Close
                    DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                    Exit Function
                End If
                If IsNull(rs.Fields("RTOP_POS")) = False Then
                    sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + 0.9 + dblEPS)    'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                End If
                If IsNull(rs.Fields("RITOP_POS")) = False Then
                    sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")) + 0.9 + dblEPS)    'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                End If
                sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
                sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
                intTopPos = Int(CDbl(rs.Fields("TOP_POS")) + 0.9 + dblEPS) '切り上げ 'add 2003/06/13 hitec)matsumoto [+ dblEPS]追加
                sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(intTopPos), "000") & "B"
                Exit Do
        End Select
        rs.MoveNext
    Loop
    sSmpId1 = sAddSmpId1
    sSmpId2 = sAddSmpId2
    rs.Close

    DBDRV_GET_WFMAP = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
'    gErr.HandleError
    DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    関数名        : DBDRV_UPD_WFMap
'*
'*    処理概要      : 1.WFマップテーブル更新
'*                    (WFマップテーブル(TBCMY011)を更新する)
'*
'*    パラメータ    : 変数名        ,IO ,型                            ,説明
'*                    SXL           ,O  ,DBDRV_scmzc_fcmlc001b_SXL039  ,SXL管理
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function DBDRV_UPD_WFMap() As FUNCTION_RETURN
    Dim sSQL            As String
    Dim rs              As OraDynaset
    Dim i               As Long
    Dim lngLoopCnt      As Long
    Dim sDBName         As String
    Dim intUCount       As Integer
    Dim dtmNowtime      As Date
    Dim vGetMaxSeq      As Variant
    Dim sGetSXLid       As String
    Dim intNowIngotPos  As Integer
    Dim intGetSmplLoop  As Integer
    Dim intFromBlkSeq   As Integer
    Dim intToBlkSeq     As Integer
    Dim intNextLoopCnt  As Integer
    Dim vGetSample      As Variant
    Dim m               As Integer
    Dim intGetNextSeq   As Integer
    Dim vGetHinban      As Variant
    Dim intAllScrapCnt  As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_UPD_WFMap"

    sDBName = "(Y011)"
    With f_cmbc039_3.sprExamine
        m = .MaxRows
        For lngLoopCnt = 1 To m Step 2
            intFromBlkSeq = gtSprWfMap(lngLoopCnt).BLOCKSEQ 'ブロックSEQを取得
            intToBlkSeq = gtSprWfMap(lngLoopCnt + 1).BLOCKSEQ 'ブロックSEQを取得
            .row = lngLoopCnt
            .col = 10
            If (Len(Trim(.text)) > 0) Or (gtSprWfMap(lngLoopCnt).hinban = "Z") Then    'サンプル行の場合
                If (gtSprWfMap(lngLoopCnt).hinban = "Z") And (gtSprWfMap(lngLoopCnt - 1).hinban = "Z") Then 'Zが連続してあった場合は、上側のSXLIDをつける
                    'SXLIDはつけない
                    If CheckGetSampleID(lngLoopCnt - 1) = True Then
                        .GetText 5, lngLoopCnt, gtSprWfMap(lngLoopCnt).KESSYOUP                 '2003/06/01 add (上SXLの終了結晶位置と下SXLの開始位置が異なるケースが存在するため）
                        sGetSXLid = Mid(gtSprWfMap(lngLoopCnt).LOTID, 1, 10) & GetWafPos(CInt(gtSprWfMap(lngLoopCnt).KESSYOUP))
                    End If
                Else
                    If lngLoopCnt = 1 Then
                        sGetSXLid = tblSXL.SXLID 'upd 2003/05/19 hitec)matsumoto 位置からSXLIDを作らない
                    Else
                        .GetText 5, lngLoopCnt, gtSprWfMap(lngLoopCnt).KESSYOUP                 '2003/06/01 add (上SXLの終了結晶位置と下SXLの開始位置が異なるケースが存在するため）
                        sGetSXLid = Mid(gtSprWfMap(lngLoopCnt).LOTID, 1, 10) & GetWafPos(CInt(gtSprWfMap(lngLoopCnt).KESSYOUP))
                    End If
                End If
            End If
            If gtSprWfMap(lngLoopCnt).hinban = "Z" Then
                sSQL = "UPDATE TBCMY011 SET"
                dtmNowtime = Format(Now, "yyyy/mm/dd hh:mm:ss")

                '先頭サンプルIDは変わらないためSXLIDも既存の方式で作られたままとする　2003/04/22
                If lngLoopCnt = 1 Then
                    intNowIngotPos = SIngotP
                Else
                    intNowIngotPos = gtSprWfMap(lngLoopCnt).KESSYOUP
                End If

                sSQL = sSQL & " MSXLID = '" & sGetSXLid & "'"

                sSQL = sSQL & ",UPDPROC = 'CW760'"             ' 更新工程
                sSQL = sSQL & ",UPDDATE = sysdate"    'upd 2003/05/03 hitec)matsumoto
                sSQL = sSQL & " WHERE LOTID ='" & gtSprWfMap(lngLoopCnt).LOTID & "'" ' ブロックID"
                If intFromBlkSeq <= intToBlkSeq Then
                    sSQL = sSQL & "   AND ((BLOCKSEQ >= " & intFromBlkSeq & ")"    ' ブロック内連番"
                    sSQL = sSQL & "       AND (BLOCKSEQ <= " & intToBlkSeq & "))"
                Else
                    sSQL = sSQL & "   AND (BLOCKSEQ >= " & intFromBlkSeq & ")"    ' ブロック内連番"
                End If
                '' WriteDBLog sSql
                If 0 >= OraDB.ExecuteSQL(sSQL) Then
                    DBDRV_UPD_WFMap = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
            Else
                sSQL = "UPDATE TBCMY011 SET"
                sSQL = sSQL & " mhinban = '" & gtSprWfMap(lngLoopCnt).hinban & "'"    ' 品番"
                dtmNowtime = Format(Now, "yyyy/mm/dd hh:mm:ss")
                '先頭サンプルIDは変わらないためSXLIDも既存の方式で作られたままとする　2003/04/22
                If lngLoopCnt = 1 Then
                    intNowIngotPos = SIngotP
                Else
                    intNowIngotPos = gtSprWfMap(lngLoopCnt).KESSYOUP
                End If
                sSQL = sSQL & ",MSXLID = '" & sGetSXLid & "'"

                sSQL = sSQL & ",UPDPROC = 'CW760'"             ' 更新工程
                sSQL = sSQL & ",UPDDATE = sysdate"    'upd 2003/05/03 hitec)matsumoto
                sSQL = sSQL & ",MREVNUM = " & gtSprWfMap(lngLoopCnt).REVNUM          ' 製品番号改訂番号
                sSQL = sSQL & ",MFACTORY = '" & gtSprWfMap(lngLoopCnt).factory & "'" ' 工場
                sSQL = sSQL & ",MOPECOND = '" & gtSprWfMap(lngLoopCnt).opecond & "'" ' 操業条件
                sSQL = sSQL & " WHERE LOTID ='" & gtSprWfMap(lngLoopCnt).LOTID & "'"                   ' ブロックID"
                If (intFromBlkSeq <= intToBlkSeq) Then

                    sSQL = sSQL & "   AND ((BLOCKSEQ >= " & intFromBlkSeq & ")"    ' ブロック内連番"
                    sSQL = sSQL & "       AND (BLOCKSEQ <= " & intToBlkSeq & "))"
                Else
                    sSQL = sSQL & "   AND (BLOCKSEQ >= " & intFromBlkSeq & ")"    ' ブロック内連番"

                End If
                '' WriteDBLog sSql

                If 0 >= OraDB.ExecuteSQL(sSQL) Then
                    DBDRV_UPD_WFMap = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If

                sSQL = "UPDATE TBCMY011 SET"
                    sSQL = sSQL & " SHAFLAG = '0'"             ' サンプルフラグ"
                    sSQL = sSQL & ",WFSTA = '0'"               ' WF状態
                sSQL = sSQL & " WHERE LOTID ='" & gtSprWfMap(lngLoopCnt).LOTID & "'"                   ' ブロックID"
                If (intFromBlkSeq <= intToBlkSeq) Then

                    sSQL = sSQL & "   AND ((BLOCKSEQ >= " & intFromBlkSeq & ")"    ' ブロック内連番"
                    sSQL = sSQL & "       AND (BLOCKSEQ <= " & intToBlkSeq & "))"

                Else
                    sSQL = sSQL & "   AND (BLOCKSEQ >= " & intFromBlkSeq & ")"    ' ブロック内連番"

                End If
                sSQL = sSQL & "  AND ( WFSTA <> '0'"
                sSQL = sSQL & "  AND  WFSTA <> '4')"
                '' WriteDBLog sSql

                If 0 >= OraDB.ExecuteSQL(sSQL) Then
                End If
            End If
        Next

        For lngLoopCnt = 1 To UBound(gtSprWfMap())
            If 0 = lngLoopCnt Mod 2 Then
                .GetText 2, lngLoopCnt - 1, vGetHinban
            Else
                .GetText 2, lngLoopCnt, vGetHinban
            End If
            If Trim(vGetHinban) <> "Z" Then
                .GetText 10, lngLoopCnt, vGetSample
                If (vGetSample <> vbNullString) Then
                    sSQL = "UPDATE TBCMY011 SET"
                    .GetText 10, lngLoopCnt, vGetSample
                    If vGetSample = gsWF_SMPL_JOINT Then    '共有
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh
                        .GetText 38, lngLoopCnt, vGetSample
                        '############################# 2003/05/23 end

                        Call Cnv_GetSample(vGetSample)      '2004/01/29 ooba

                        sSQL = sSQL & " MSMPLEID = '" & vGetSample & "'" ' 抜試位置"
                        sSQL = sSQL & ",SHAFLAG = '1'"             ' サンプルフラグ"
                        sSQL = sSQL & ",WFSTA = '1'"    ' WF状態サンプル
                    Else
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh
                        .GetText 38, lngLoopCnt, vGetSample

                        Call Cnv_GetSample(vGetSample)      '2004/01/29 ooba

                        sSQL = sSQL & " MSMPLEID = '" & vGetSample & "'" ' 抜試位置"
                        If lngLoopCnt <> 1 And lngLoopCnt <> UBound(gtSprWfMap()) Then  'upd hitec)matsumoto 初期表示サンプルのフラグは更新しない
                            sSQL = sSQL & ",WFSTA = '0'"    ' WF状態サンプル
                            sSQL = sSQL & ",SHAFLAG = '1'"             ' サンプルフラグ"
                        End If
                    End If

                    dtmNowtime = Format(Now, "yyyy/mm/dd hh:mm:ss")

                    sSQL = sSQL & ",UPDPROC = 'CW760'"             ' 更新工程
                    sSQL = sSQL & ",UPDDATE = sysdate"    'upd 2003/05/03 hitec)matsumoto

                    sSQL = sSQL & " WHERE LOTID ='" & gtSprWfMap(lngLoopCnt).LOTID & "'"                   ' ブロックID"
                    sSQL = sSQL & " AND BLOCKSEQ = " & gtSprWfMap(lngLoopCnt).BLOCKSEQ              ' ブロック内連番"
                    '' WriteDBLog sSql
                    If 0 >= OraDB.ExecuteSQL(sSQL) Then
                        DBDRV_UPD_WFMap = FUNCTION_RETURN_FAILURE
                        GoTo proc_exit
                    End If
                End If
            End If
        Next
    End With

    DBDRV_UPD_WFMap = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    DBDRV_UPD_WFMap = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'*******************************************************************************
'*    関数名        : DBDRV_WFLENGET
'*
'*    処理概要      : 1.該当ブロックのWF１枚の長さ（計算長）を取得
'*
'*    パラメータ    : 変数名        ,IO ,型                    ,説明
'*                    BLOCKID       ,I  ,STRING                ,ブロックＩＤ
'*                    dblWFLen      ,O  ,DOUBLE        　　    ,WF1枚の計算長さ
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function DBDRV_WFLENGET(ByVal strBlockID As String, _
                                ByRef dblWFLen As Double) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim intRealLen  As Integer
    Dim intWFcnt    As Integer
    Dim rs          As OraDynaset
    Dim intKetuFrom As Integer
    Dim intKetuTo   As Integer
    Dim intKetuLen  As Integer
    Dim sDBName     As String

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_WFLENGET"

    '実長さ、WF枚数取得
    sDBName = "(Y011)"

    sSQL = "select e40.blockid,e40.reallen,y11.cnt"
    sSQL = sSQL & " from tbcme040 e40,"
    sSQL = sSQL & " xsdca xa,"
    sSQL = sSQL & " (select lotid,count(lotid) cnt"
    sSQL = sSQL & " from tbcmy011"
    sSQL = sSQL & " where lotid ='" & strBlockID & "'"
    sSQL = sSQL & " group by lotid  ) y11"
    sSQL = sSQL & " where e40.blockid = xa.CRYNUMCA"
    sSQL = sSQL & " and   y11.lotid   = xa.CRYNUMCA"
    sSQL = sSQL & " and   y11.lotid  = '" & strBlockID & "'"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If Not rs.EOF Then
           intRealLen = CInt(rs!REALLEN)
           intWFcnt = CInt(rs!cnt)
    Else
        rs.Close
        DBDRV_WFLENGET = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    rs.Close

    '欠落長さ取得
    sDBName = "(Y012)"
    sSQL = "SELECT DISTINCT LENFROM,LENTO FROM TBCMY012"
    sSQL = sSQL & " Where "
    sSQL = sSQL & " LOTID   = '" & strBlockID & "'"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    intKetuLen = 0
    Do While Not rs.EOF
        If (IsNull(rs.Fields("LENFROM")) = True) Or rs.Fields("LENFROM") = -1 Or _
            (IsNull(rs.Fields("LENTO")) = True) Or rs.Fields("LENTO") = -1 Then
        Else
            intKetuFrom = CInt(rs.Fields("LENFROM"))
            intKetuTo = CInt(rs.Fields("LENTO"))
            intKetuLen = intKetuLen + intKetuTo - intKetuFrom
        End If
        rs.MoveNext
    Loop
    rs.Close

    'WF長さ計算
    dblWFLen = (intRealLen - intKetuLen) / intWFcnt

    DBDRV_WFLENGET = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    DBDRV_WFLENGET = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'**********************************************************************************************************
'*    関数名        : GetZMotoHinban
'*
'*    処理概要      : 1.TBCMY007からZ品番の元品番を取得する
'*
'*    パラメータ    : 変数名        ,IO ,型      ,説明
'*                    strSXLID      ,I  ,String  ,SXL-ID
'*                    strMotoHinban ,O  ,String  ,元品番
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'**********************************************************************************************************
Public Function GetZMotoHinban(ByVal strSXLID As String, ByRef strMotoHinban As String) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rsMain      As OraDynaset
    Dim sErrTbl     As String
    Dim strDBName   As String

    On Error GoTo proc_err

    'ブロックID取得
    strDBName = "Y007"
    sSQL = " select HINBAN from TBCMY007"
    sSQL = sSQL & "  where SXL_ID='" & strSXLID & "'"

    Set rsMain = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rsMain.RecordCount = 0 Then
        rsMain.Close
        GetZMotoHinban = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    Do While Not rsMain.EOF
        If IsNull(rsMain.Fields("HINBAN")) = True Then
            strMotoHinban = vbNullString
        Else
            strMotoHinban = Mid(Trim(rsMain.Fields("HINBAN")), 1, 8)
        End If
        rsMain.MoveNext
    Loop
    rsMain.Close
    GetZMotoHinban = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    GetZMotoHinban = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'*******************************************************************************************************
'*    関数名        : scmzc_getE036
'*
'*    処理概要      : 1.製品仕様WFデータ（OT１、OT2)の取得ドライバ
'*
'*    パラメータ    : 変数名        ,IO ,型           ,説明
'*                    pHIN          ,I  ,tFullHinban  ,品番情報
'*                    strOT1        ,O  ,String       ,その他サンプル1
'*                    strOT2        ,O  ,String       ,その他サンプル2
'*                    strMAI1       ,O  ,String       ,その他サンプル1枚数
'*                    strMAI2       ,O  ,String       ,その他サンプル2枚数
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************************
Public Function scmzc_getE036(pHIN As tFullHinban, strOT1 As String, strOT2 As String, strMAI1, strMAI2) _
                                As FUNCTION_RETURN
    Dim sSQL As String
    Dim rs  As OraDynaset

    '' エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmzcF_cmkc001WF.bas -- Function scmzc_getE036"

    sSQL = ""
    sSQL = sSQL & " SELECT"
    sSQL = sSQL & "   a.ot1 AS other1"
    sSQL = sSQL & "  ,a.ot1m AS other1mai"
    sSQL = sSQL & "  ,b.ot2 AS other2"
    sSQL = sSQL & "  ,b.ot2m AS other2mai"
    sSQL = sSQL & " FROM"
    sSQL = sSQL & "   ("
    sSQL = sSQL & "    SELECT"
    sSQL = sSQL & "      COUNT(other1)"
    sSQL = sSQL & "     ,MAX(other1) AS ot1"
    sSQL = sSQL & "     ,MAX(other1mai) AS ot1m"
    sSQL = sSQL & "    FROM"
    sSQL = sSQL & "      tbcme036"
    sSQL = sSQL & "    WHERE hinban   = '" & pHIN.hinban & "'"
    sSQL = sSQL & "      AND mnorevno = " & pHIN.mnorevno
    sSQL = sSQL & "      AND factory  = '" & pHIN.factory & "'"
    sSQL = sSQL & "      AND opecond  = '" & pHIN.opecond & "'"
    sSQL = sSQL & "      AND othertime > SYSDATE"
    sSQL = sSQL & "   ) a"
    sSQL = sSQL & "  ,("
    sSQL = sSQL & "    SELECT"
    sSQL = sSQL & "      COUNT(other2)"
    sSQL = sSQL & "     ,MAX(other2) AS ot2"
    sSQL = sSQL & "     ,MAX(other2mai) AS ot2m"
    sSQL = sSQL & "    FROM"
    sSQL = sSQL & "      tbcme036"
    sSQL = sSQL & "    WHERE hinban   = '" & pHIN.hinban & "'"
    sSQL = sSQL & "      AND mnorevno = " & pHIN.mnorevno
    sSQL = sSQL & "      AND factory  = '" & pHIN.factory & "'"
    sSQL = sSQL & "      AND opecond  = '" & pHIN.opecond & "'"
    sSQL = sSQL & "      AND othertime2 > SYSDATE"
    sSQL = sSQL & "   ) b"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount = 0 Then
        rs.Close
        strOT1 = "0"
        strOT2 = "0"
        strMAI1 = "0"
        strMAI2 = "0"
        GoTo proc_exit
    End If
    If IsNull(rs("OTHER1")) = True Then
        strOT1 = "0"
    Else
        strOT1 = rs("OTHER1")
    End If
    If IsNull(rs("OTHER2")) = True Then
        strOT2 = "0"
    Else
        strOT2 = rs("OTHER2")
    End If
    If IsNull(rs("OTHER1MAI")) = True Then
        strMAI1 = "0"
    Else
        strMAI1 = rs("OTHER1MAI")
    End If
    If IsNull(rs("OTHER2MAI")) = True Then
        strMAI2 = "0"
    Else
        strMAI2 = rs("OTHER2MAI")
    End If
    scmzc_getE036 = FUNCTION_RETURN_SUCCESS
    rs.Close

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    scmzc_getE036 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    関数名        : Pic_Disp
'*
'*    処理概要      : 1.SXLチェックボックス詳細の表示
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    iIndex        ,I  ,Integer  ,１：表示 / ０：非表示
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Public Sub Pic_Disp(iIndex As Integer)
    Dim intCnt    As Integer

    With f_cmbc039_3
        If iIndex = 0 Then
            For intCnt = 0 To 2
            .lbl_check(intCnt).Visible = False
            Next
            .pic_check(0).Visible = False
            .pic_check(1).Visible = False
        ElseIf iIndex = 1 Then
            For intCnt = 0 To 2
            .lbl_check(intCnt).Visible = True
            Next
            .pic_check(0).Visible = True
            .pic_check(1).Visible = True
        End If
    End With
End Sub

'*******************************************************************************
'*    関数名        : CheckGetSampleID
'*
'*    処理概要      : 1.サンプルＩＤの取得判定
'*
'*    パラメータ    : 変数名       ,IO ,型       ,説明
'*                    iWafPos      ,I  ,Integer　,抜試指示テーブル位置
'*
'*    戻り値        : Boolean(選択の有無)
'*
'*******************************************************************************
Public Function CheckGetSampleID(iWafPos As Integer) As Boolean
    Dim vNowhinban As Variant
    Dim vUDhinban  As Variant
    Dim vFlg       As Variant
    Dim sSampID    As Variant
    Dim vSampleID  As Variant
    Dim intPointer As Integer
    Dim vOldHinban As Variant
    Dim blCheckbox As Boolean       'チェックボックスフラグ
    Dim lngRow     As Long

    CheckGetSampleID = False

    With f_cmbc039_3.sprExamine
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh
        .GetText 37, iWafPos, vFlg
        If vFlg = "1" Then  '先頭と最終行（先頭はこの関数には来ないはず）
            Exit Function

        ElseIf vFlg = "3" And iWafPos Mod 2 = 0 Then    '初期表示でサンプルの無い行（ブロックの境）
            .GetText 2, iWafPos - 1, vNowhinban         '現在の品番
            .GetText 2, iWafPos + 1, vUDhinban  '下品番
            If vUDhinban <> vNowhinban Then
                'チェックボックス消去
                .col = 1
                .row = iWafPos
                .CellType = CellTypeEdit
                .text = ""
                Call Pic_Disp(0) '03/05/31
                CheckGetSampleID = True
            Else
                If ADD_CHECKBOX(iWafPos, blCheckbox) = FUNCTION_RETURN_FAILURE Then
                    CheckGetSampleID = True
                End If
            End If
        ElseIf iWafPos Mod 2 = 0 Then   '偶数行
            .GetText 2, iWafPos - 1, vNowhinban
            .GetText 2, iWafPos + 1, vUDhinban  '下品番
            If vNowhinban = vUDhinban Then
                If vFlg = "2" Or vFlg = "" Or vFlg = "0" Then
                    CheckGetSampleID = True
                End If
            Else
                CheckGetSampleID = True
            End If
        End If
    End With
End Function

'*****************************************************************************************************************
'*    関数名        : GetSxlidINBlkid
'*
'*    処理概要      : 1.同一SXL同一品番のブロック境界に抜試有無のチェックボックス表示、チェックボックスの内容判定
'*                      CW740の場合、A欠落を飛ばしてブロック最終行にチェックボックスを配置する
'*
'*    パラメータ    : 変数名        ,IO ,型         ,説明
'*　　                iWafPos       ,I  ,Integer　  ,Spread行
'*                    bCheckbox     ,IO ,Boolean    ,チェックボックス表示フラグ
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*****************************************************************************************************************
Private Function ADD_CHECKBOX(ByVal iWafPos As Integer, bCheckbox As Boolean) As FUNCTION_RETURN
    Dim j           As Integer
    Dim intRowCnt   As Integer
    Dim vGetLot     As Variant   'CW740用
    Dim vGetLot2    As Variant   'CW740用
    Dim intCnt      As Integer
    Dim vBackColor  As Variant

    ADD_CHECKBOX = FUNCTION_RETURN_SUCCESS

    intRowCnt = iWafPos

    With f_cmbc039_3.sprExamine
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh
        .GetText 39, iWafPos, vGetLot
        .GetText 39, iWafPos + 1, vGetLot2
        If vGetLot = vGetLot2 Then
            Exit Function
        End If

        .col = 1
        .row = intRowCnt
        If .CellType <> CellTypeCheckBox Then
            bCheckbox = True
        ElseIf .text = "1" Then
            ADD_CHECKBOX = FUNCTION_RETURN_FAILURE
            Exit Function
        Else
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh
            For j = 10 To 35
                If j <> 27 And j <> 35 Then
                    .SetText j, intRowCnt, vbNullString
                    .SetText j, intRowCnt + 1, vbNullString
                End If
            Next j

            .col = 2
            .row = intRowCnt - 1
            If .text <> "Z" Then
                .col = 11
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh
                .col2 = 35
                .row = intRowCnt
                .row2 = intRowCnt
                .Lock = True
                .BlockMode = True
                .backColor = vbWhite
                .BlockMode = False
            Else
                .col = 11
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh
                .col2 = 35
                .row = intRowCnt
                .row2 = intRowCnt
                .Lock = True
                .BlockMode = True
                .backColor = &H8080FF
                .BlockMode = False
            End If
            .col = 2
            .row = intRowCnt + 1
            If .text <> "Z" Then
                .col = 11
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh
                .col2 = 35
                .row = intRowCnt + 1
                .row2 = intRowCnt + 1
                .BlockMode = True
                .backColor = vbWhite
                .BlockMode = False
                .Lock = True
            Else
                .col = 11
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh
                .col2 = 35
                .row = intRowCnt + 1
                .row2 = intRowCnt + 1
                .BlockMode = True
                .backColor = &H8080FF
                .BlockMode = False
                .Lock = True
            End If
        End If
        If bCheckbox = True Then
            .col = 1
            .row = intRowCnt
            .Lock = False
            .CellType = CellTypeCheckBox
            Call Pic_Disp(1) '03/05/31
            .TypeCheckTextAlign = TypeCheckTextAlignLeft
            .TypeCheckType = TypeCheckTypeNormal
            .TypeCheckCenter = False
        End If
    End With
End Function

'*******************************************************************************
'*    関数名        : Cnv_GetSample
'*
'*    処理概要      : 1.サンプルIDの変換処理
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    vGetSample    ,I  ,Variant  ,SXL管理
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Public Sub Cnv_GetSample(ByRef vGetSample As Variant)
    Dim i       As Integer
    Dim sKbn    As String

    For i = 1 To UBound(CngSmpID_UD)
        If CngSmpID_UD(i) = vGetSample Then
           sKbn = Cnv_Smp_KB(Right(vGetSample, 1))
           vGetSample = left(vGetSample, Len(vGetSample) - 1) + sKbn
           Exit Sub
        End If
    Next
End Sub

'*******************************************************************************
'*    関数名        : Cnv_Smp_KB
'*
'*    処理概要      : 1.サンプル区分の変換処理
'*　　　　　　　　　　(Ｕ⇒Ｂ　Ｄ⇒Ｔに変換)
'*    パラメータ    : 変数名        ,IO ,型      ,説明
'*                    SmpKb         ,I  ,String  ,サンプル区分
'*
'*    戻り値        : String（サンプル区分）
'*
'*
'*******************************************************************************
Public Function Cnv_Smp_KB(SmpKb As String) As String
    If SmpKb = "U" Then
        Cnv_Smp_KB = "B"
        Exit Function
    End If

    If SmpKb = "D" Then
        Cnv_Smp_KB = "T"
        Exit Function
    End If
End Function

'********************************************************************************************************
'*    関数名        : chkComSAMPL
'*
'*    処理概要      : 1.共有サンプルチェック処理
'*                    (指定されたｻﾝﾌﾟﾙIDが全共有かどうかをﾁｪｯｸし、全共有の場合、共有ｻﾝﾌﾟﾙIDを取得し返す)
'*
'*    パラメータ    : 変数名        ,IO ,型         ,説明
'*                    inSXLID       ,I  ,String     , SXL-ID
'*                    inSMPLID      ,I  ,String     , ｻﾝﾌﾟﾙID
'*                    outSMPLID     ,O  ,String     , 共有ｻﾝﾌﾟﾙID(共有でない場合、inSMPLIDを返す)
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************************
Public Function chkComSAMPL(inSXLID As String, inSMPLID As String, outSMPLID As String) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sSQL        As String
    Dim sXTALCW     As String
    Dim sINPOSCW    As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function chkComSAMPL"

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    chkComSAMPL = FUNCTION_RETURN_SUCCESS
    outSMPLID = inSMPLID

    '-------------------- 全共有確認(XSDCW) ----------------------------------------
    sSQL = "select XTALCW, INPOSCW from XSDCW "
    sSQL = sSQL & "where SXLIDCW = '" & inSXLID & "' and "
    sSQL = sSQL & "      REPSMPLIDCW = '" & inSMPLID & "' and "
    sSQL = sSQL & "      (WFINDRSCW = '2' or WFINDRSCW = '0' or WFINDRSCW = ' ' or WFINDRSCW is null) and "
    sSQL = sSQL & "      (WFINDOICW = '2' or WFINDOICW = '0' or WFINDOICW = ' ' or WFINDOICW is null) and "
    sSQL = sSQL & "      (WFINDB1CW = '2' or WFINDB1CW = '0' or WFINDB1CW = ' ' or WFINDB1CW is null) and "
    sSQL = sSQL & "      (WFINDB2CW = '2' or WFINDB2CW = '0' or WFINDB2CW = ' ' or WFINDB2CW is null) and "
    sSQL = sSQL & "      (WFINDB2CW = '2' or WFINDB3CW = '0' or WFINDB3CW = ' ' or WFINDB3CW is null) and "
    sSQL = sSQL & "      (WFINDL1CW = '2' or WFINDL1CW = '0' or WFINDL1CW = ' ' or WFINDL1CW is null) and "
    sSQL = sSQL & "      (WFINDL2CW = '2' or WFINDL2CW = '0' or WFINDL2CW = ' ' or WFINDL2CW is null) and "
    sSQL = sSQL & "      (WFINDL3CW = '2' or WFINDL3CW = '0' or WFINDL3CW = ' ' or WFINDL3CW is null) and "
    sSQL = sSQL & "      (WFINDL4CW = '2' or WFINDL4CW = '0' or WFINDL4CW = ' ' or WFINDL4CW is null) and "
    sSQL = sSQL & "      (WFINDDSCW = '2' or WFINDDSCW = '0' or WFINDDSCW = ' ' or WFINDDSCW is null) and "
    sSQL = sSQL & "      (WFINDDZCW = '2' or WFINDDZCW = '0' or WFINDDZCW = ' ' or WFINDDZCW is null) and "
    sSQL = sSQL & "      (WFINDSPCW = '2' or WFINDSPCW = '0' or WFINDSPCW = ' ' or WFINDSPCW is null) and "
    sSQL = sSQL & "      (WFINDDO1CW = '2' or WFINDDO1CW = '0' or WFINDDO1CW = ' ' or WFINDDO1CW is null) and "
    sSQL = sSQL & "      (WFINDDO2CW = '2' or WFINDDO2CW = '0' or WFINDDO2CW = ' ' or WFINDDO2CW is null) and "
    sSQL = sSQL & "      (WFINDDO3CW = '2' or WFINDDO3CW = '0' or WFINDDO3CW = ' ' or WFINDDO3CW is null) and "
    sSQL = sSQL & "      (WFINDOT1CW = '2' or WFINDOT1CW = '0' or WFINDOT1CW = ' ' or WFINDOT1CW is null) and "
    sSQL = sSQL & "      (WFINDOT2CW = '2' or WFINDOT2CW = '0' or WFINDOT2CW = ' ' or WFINDOT2CW is null) and "
    sSQL = sSQL & "      (WFINDAOICW = '2' or WFINDAOICW = '0' or WFINDAOICW = ' ' or WFINDAOICW is null) and "   '残存酸素追加　03/12/19 ooba
    sSQL = sSQL & "      (WFINDGDCW = '2' or WFINDGDCW = '0' or WFINDGDCW = ' ' or WFINDGDCW is null or (WFINDGDCW = '1' and WFHSGDCW = '1')) "   'GD追加　05/02/24 ooba
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    sSQL = sSQL & "  and (EPINDB1CW = '2' or EPINDB1CW = '0' or EPINDB1CW = ' ' or EPINDB1CW is null) and "
    sSQL = sSQL & "      (EPINDB2CW = '2' or EPINDB2CW = '0' or EPINDB2CW = ' ' or EPINDB2CW is null) and "
'--- 2009/07/30 Change Y.Hitomi
'    sSql = sSql & "      (EPINDB2CW = '2' or EPINDB3CW = '0' or EPINDB3CW = ' ' or EPINDB3CW is null) and "
    sSQL = sSQL & "      (EPINDB3CW = '2' or EPINDB3CW = '0' or EPINDB3CW = ' ' or EPINDB3CW is null) and "
    sSQL = sSQL & "      (EPINDL1CW = '2' or EPINDL1CW = '0' or EPINDL1CW = ' ' or EPINDL1CW is null) and "
    sSQL = sSQL & "      (EPINDL2CW = '2' or EPINDL2CW = '0' or EPINDL2CW = ' ' or EPINDL2CW is null) and "
    sSQL = sSQL & "      (EPINDL3CW = '2' or EPINDL3CW = '0' or EPINDL3CW = ' ' or EPINDL3CW is null) "
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GoTo proc_exit
    End If
    sXTALCW = rs("XTALCW")      '結晶番号
    sINPOSCW = rs("INPOSCW")    '結晶内位置
    Set rs = Nothing

    '-------------------- 共有ｻﾝﾌﾟﾙIDの取得(XSDCW) ----------------------------------------
    sSQL = "select REPSMPLIDCW from XSDCW "
    sSQL = sSQL & "where SXLIDCW like '" & left(sXTALCW, 9) & "%' and "     'ｲﾝﾃﾞｯｸｽ項目追加 09/05/25 ooba
    sSQL = sSQL & "      XTALCW = '" & sXTALCW & "' and "
    sSQL = sSQL & "      INPOSCW = '" & sINPOSCW & "' and "
    sSQL = sSQL & "      SXLIDCW != '" & inSXLID & "' and "
    sSQL = sSQL & "      REPSMPLIDCW != '" & inSMPLID & "' "
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GoTo proc_exit
    End If
    outSMPLID = rs("REPSMPLIDCW")       '代表ｻﾝﾌﾟﾙID(共有)
    Set rs = Nothing

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    chkComSAMPL = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'***********************************************************************************************
'*    関数名        : cmbc039_GetSxlRsData
'*
'*    処理概要      : 1.SXL確定指示(TBCMY007)ﾃｰﾌﾞﾙにｾｯﾄするSXLの比抵抗ﾃﾞｰﾀを取得する。
'*
'*    パラメータ    : 変数名        ,IO  ,型                ,説明
'*                    oldSXLID      ,I   ,String            旧SXLID
'*                    newSXLID      ,I   ,String            新SXLID
'*                    iRow          ,I   ,String            画面ｽﾌﾟﾚｯﾄﾞ行数
'*                    sDataPattern  ,I   ,String            比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝ
'*                                                            ●ﾊﾟﾀｰﾝA : WF実績ﾃﾞｰﾀ取得
'*                                                            ●ﾊﾟﾀｰﾝB : 結晶実績ﾃﾞｰﾀ取得
'*                                                            ●ﾊﾟﾀｰﾝC : 取得ﾃﾞｰﾀなし
'*                    iSxlPattern   ,I   ,String            登録SXLﾊﾟﾀｰﾝ
'*                                                            ●ﾊﾟﾀｰﾝ1 : 全廃棄SXL
'*                                                            ●ﾊﾟﾀｰﾝ2 : 上追込みSXL
'*                                                            ●ﾊﾟﾀｰﾝ3 : 下追込みSXL
'*                                                            ●ﾊﾟﾀｰﾝ4 : SXLの間をZ
'*                                                            ●ﾊﾟﾀｰﾝ0 : 取得ﾃﾞｰﾀなし
'*                    mesdata()     ,O   ,String            比抵抗ﾃﾞｰﾀ
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'***********************************************************************************************
Public Function cmbc039_GetSxlRsData(oldSXLID As String, newSXLID As String, IRow As Integer, _
                                        sDataPattern As String, iSxlPattern As Integer, _
                                        mesdata() As String) As FUNCTION_RETURN
    Dim sTBkbn      As String        'T/B区分
    Dim sBlkId      As String        'ｻﾝﾌﾟﾙﾌﾞﾛｯｸID
    Dim sSmpId      As String        'ｻﾝﾌﾟﾙID(Rs)
    Dim i           As Integer
    Dim j           As Integer
    Dim intChkRow   As Integer
    Dim sSQL        As String
    Dim rs          As OraDynaset

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function cmbc039_GetSxlRsData"
    cmbc039_GetSxlRsData = FUNCTION_RETURN_FAILURE

    '比抵抗ﾃﾞｰﾀ初期化
    For i = 1 To 10
        mesdata(i) = ""
    Next

    '比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝが『A』の場合、WF実績ﾃﾞｰﾀ(TBCMY013)を取得する。
    If sDataPattern = "A" Then
        For i = 1 To 2
            If i = 1 Then sTBkbn = "T" Else sTBkbn = "B"

            '登録SXLが『ﾊﾟﾀｰﾝ1』、『ﾊﾟﾀｰﾝ2』TOP側、『ﾊﾟﾀｰﾝ3』BOT側はﾃｰﾌﾞﾙからｻﾝﾌﾟﾙID(Rs)を取得する。
            If iSxlPattern = 1 Or (iSxlPattern = 2 And sTBkbn = "T") Or _
                (iSxlPattern = 3 And sTBkbn = "B") Then

                '該当SXLより、新ｻﾝﾌﾟﾙ管理-WF<XSDCW>のｻﾝﾌﾟﾙID_Rsを取得。
                sSQL = "select WFSMPLIDRSCW "
                sSQL = sSQL & "from XSDCW "
                sSQL = sSQL & "where TBKBNCW = '" & sTBkbn & "' "
                sSQL = sSQL & "and SXLIDCW = '" & oldSXLID & "' "

                Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

                If rs.RecordCount = 1 Then
                    sSmpId = rs("WFSMPLIDRSCW")
                End If
                Set rs = Nothing

            '登録SXLが『ﾊﾟﾀｰﾝ2』BOT側、『ﾊﾟﾀｰﾝ3』TOP側、『ﾊﾟﾀｰﾝ4』は内部ﾃﾞｰﾀからｻﾝﾌﾟﾙID(Rs)を取得する。
            ElseIf (iSxlPattern = 2 And sTBkbn = "B") Or (iSxlPattern = 3 And sTBkbn = "T") Or _
                    iSxlPattern = 4 Then

                If f_cmbc039_3.sprExamine.MaxRows = UBound(tblWfSample) Then
                    If sTBkbn = "T" Then
                        sSmpId = tblWfSample(IRow).WFSMP.WFSMPLIDRSCW
                    ElseIf sTBkbn = "B" Then
                        sSmpId = tblWfSample(IRow + 1).WFSMP.WFSMPLIDRSCW
                    End If

                '1SXL複数ﾌﾞﾛｯｸをSXL分割しない場合
                Else
                    If IRow > UBound(tblWfSample) Then
                        intChkRow = UBound(tblWfSample)
                    Else
                        intChkRow = IRow + 1
                    End If
                    For j = intChkRow To 1 Step -1
                        If tblWfSample(j).WFSMP.SXLIDCW = newSXLID Then
                            'TOP側は奇数行、BOT側は偶数行
                            If (sTBkbn = "T" And j Mod 2 = 1) Or (sTBkbn = "B" And j Mod 2 = 0) Then
                                sSmpId = tblWfSample(j).WFSMP.WFSMPLIDRSCW
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If

            If Trim(sSmpId) <> "" Then
                'ｻﾝﾌﾟﾙID_Rsから、測定評価結果<TBCMY013>の比抵抗実績ﾃﾞｰﾀ(TOP側/BOT側)を取得する。
                sSQL = "select MESDATA1, MESDATA2, MESDATA3, MESDATA4, MESDATA5 "
                sSQL = sSQL & "from TBCMY013 "
                sSQL = sSQL & "where OSITEM = 'RES' "
                sSQL = sSQL & "and SAMPLEID = '" & sSmpId & "' "

                Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

                If rs.RecordCount = 1 Then
                    'TOP側実績ﾃﾞｰﾀ
                    If sTBkbn = "T" Then
                        mesdata(1) = rs("MESDATA1")
                        mesdata(2) = rs("MESDATA2")
                        mesdata(3) = rs("MESDATA3")
                        mesdata(4) = rs("MESDATA4")
                        mesdata(5) = rs("MESDATA5")
                    'BOT側実績ﾃﾞｰﾀ
                    ElseIf sTBkbn = "B" Then
                        mesdata(6) = rs("MESDATA1")
                        mesdata(7) = rs("MESDATA2")
                        mesdata(8) = rs("MESDATA3")
                        mesdata(9) = rs("MESDATA4")
                        mesdata(10) = rs("MESDATA5")
                    End If
                End If
                Set rs = Nothing
            End If
        Next
    '比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝが『B』の場合、結晶実績ﾃﾞｰﾀ(TBCMJ002)を取得する。
    ElseIf sDataPattern = "B" Then
        For i = 1 To 2
            If i = 1 Then sTBkbn = "T" Else sTBkbn = "B"

            '登録SXLが『ﾊﾟﾀｰﾝ1』、『ﾊﾟﾀｰﾝ2』TOP側、『ﾊﾟﾀｰﾝ3』BOT側はﾃｰﾌﾞﾙからｻﾝﾌﾟﾙﾌﾞﾛｯｸIDを取得する。
            If iSxlPattern = 1 Or (iSxlPattern = 2 And sTBkbn = "T") Or _
                (iSxlPattern = 3 And sTBkbn = "B") Then

                '該当SXLより、新ｻﾝﾌﾟﾙ管理-WF<XSDCW>のｻﾝﾌﾟﾙﾌﾞﾛｯｸIDを取得
                sSQL = "select SMCRYNUMCW "
                sSQL = sSQL & "from XSDCW "
                sSQL = sSQL & "where TBKBNCW = '" & sTBkbn & "' "
                sSQL = sSQL & "and SXLIDCW = '" & oldSXLID & "' "

                Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

                If rs.RecordCount = 1 Then
                    sBlkId = rs("SMCRYNUMCW")
                End If
                Set rs = Nothing
            '登録SXLが『ﾊﾟﾀｰﾝ2』BOT側、『ﾊﾟﾀｰﾝ3』TOP側、『ﾊﾟﾀｰﾝ4』は内部ﾃﾞｰﾀからｻﾝﾌﾟﾙﾌﾞﾛｯｸIDを取得する。
            ElseIf (iSxlPattern = 2 And sTBkbn = "B") Or (iSxlPattern = 3 And sTBkbn = "T") Or _
                    iSxlPattern = 4 Then

                If f_cmbc039_3.sprExamine.MaxRows = UBound(tblWfSample) Then
                    If sTBkbn = "T" Then
                        sBlkId = tblWfSample(IRow).WFSMP.SMCRYNUMCW
                    ElseIf sTBkbn = "B" Then
                        sBlkId = tblWfSample(IRow + 1).WFSMP.SMCRYNUMCW
                    End If
                '1SXL複数ﾌﾞﾛｯｸをSXL分割しない場合
                Else
                    If IRow > UBound(tblWfSample) Then
                        intChkRow = UBound(tblWfSample)
                    Else
                        intChkRow = IRow + 1
                    End If
                    For j = intChkRow To 1 Step -1
                        If tblWfSample(j).WFSMP.SXLIDCW = newSXLID Then
                            'TOP側は奇数行、BOT側は偶数行
                            If (sTBkbn = "T" And j Mod 2 = 1) Or (sTBkbn = "B" And j Mod 2 = 0) Then
                                sBlkId = tblWfSample(j).WFSMP.SMCRYNUMCW
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If

            If Trim(sBlkId) <> "" Then
                'T/B区分、ｻﾝﾌﾟﾙﾌﾞﾛｯｸIDから、新ｻﾝﾌﾟﾙ管理-ﾌﾞﾛｯｸ<XSDCS>の結晶番号、ｻﾝﾌﾟﾙID_Rsを取得。
                '結晶番号、ｻﾝﾌﾟﾙID_Rsから、結晶抵抗実績<TBCMJ002>の比抵抗実績ﾃﾞｰﾀ(TOP側/BOT側)を取得する。
                sSQL = "select MEAS1, MEAS2, MEAS3, MEAS4, MEAS5 "
                sSQL = sSQL & "from TBCMJ002 "
                sSQL = sSQL & "where (CRYNUM, SMPLNO) in ( "
                sSQL = sSQL & "         select XTALCS, CRYSMPLIDRSCS "
                sSQL = sSQL & "         from XSDCS "
                sSQL = sSQL & "         where TBKBNCS = '" & sTBkbn & "' "
                sSQL = sSQL & "         and CRYNUMCS = '" & sBlkId & "') "
                sSQL = sSQL & "and TRANCNT = ( "
                sSQL = sSQL & "         select max(TRANCNT) "
                sSQL = sSQL & "         from TBCMJ002 "
                sSQL = sSQL & "         where (CRYNUM, SMPLNO) in ( "
                sSQL = sSQL & "                  select XTALCS, CRYSMPLIDRSCS "
                sSQL = sSQL & "                  from XSDCS "
                sSQL = sSQL & "                  where TBKBNCS = '" & sTBkbn & "' "
                sSQL = sSQL & "                  and CRYNUMCS = '" & sBlkId & "')) "

                Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

                If rs.RecordCount = 1 Then
                    'TOP側実績ﾃﾞｰﾀ
                    If sTBkbn = "T" Then
                        mesdata(1) = CStr(rs("MEAS1"))
                        mesdata(2) = CStr(rs("MEAS2"))
                        mesdata(3) = CStr(rs("MEAS3"))
                        mesdata(4) = CStr(rs("MEAS4"))
                        mesdata(5) = CStr(rs("MEAS5"))
                    'BOT側実績ﾃﾞｰﾀ
                    ElseIf sTBkbn = "B" Then
                        mesdata(6) = CStr(rs("MEAS1"))
                        mesdata(7) = CStr(rs("MEAS2"))
                        mesdata(8) = CStr(rs("MEAS3"))
                        mesdata(9) = CStr(rs("MEAS4"))
                        mesdata(10) = CStr(rs("MEAS5"))
                    End If
                End If
                Set rs = Nothing
            End If
        Next
    '比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝが『C』の場合、取得実績ﾃﾞｰﾀなし。
    ElseIf sDataPattern = "C" Then
    End If

    '取得ﾃﾞｰﾀが空白/-1/NULLの時はｽﾍﾟｰｽをｾｯﾄする。
    For i = 1 To 10
        If mesdata(i) = "" Or mesdata(i) = "-1" Or mesdata(i) = vbNullString Then
            mesdata(i) = " "
        End If
    Next

    cmbc039_GetSxlRsData = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    cmbc039_GetSxlRsData = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'***********************************************************************************************
'*    関数名        : DBDRV_GetTBCMJ015Cnt
'*
'*    処理概要      : 1.GD実績存在ﾁｪｯｸ
'*
'*    パラメータ    : 変数名        ,IO ,型              ,説明
'*      　　          sSampleid　　 ,I  ,String        　,ｻﾝﾌﾟﾙID
'*      　　          iRecCnt　　   ,O  ,Integer         ,ﾚｺｰﾄﾞ数
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'***********************************************************************************************
Public Function DBDRV_GetTBCMJ015Cnt(sSampleid As String, iRecCnt As Integer) As FUNCTION_RETURN
    Dim sSQL    As String
    Dim rs      As OraDynaset

    'ｻﾝﾌﾟﾙIDと保証ﾌﾗｸﾞを元にGD実績を取得
    sSQL = "SELECT CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, HSFLG, SMPLNO "
    sSQL = sSQL & "FROM TBCMJ015 "
    sSQL = sSQL & "WHERE SMPLNO = '" & sSampleid & "' "
    sSQL = sSQL & "AND HSFLG = '1' "

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)
    If rs Is Nothing Then
        iRecCnt = 0
        DBDRV_GetTBCMJ015Cnt = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    '抽出結果のﾚｺｰﾄﾞ数を登録
    iRecCnt = rs.RecordCount
    rs.Close

    DBDRV_GetTBCMJ015Cnt = FUNCTION_RETURN_SUCCESS
End Function

'*****************************************************************************************
'*    関数名        : ChkAoiSiyou
'*
'*    処理概要      : 1.酸素析出と残存酸素の仕様チェック
'*                    (酸素析出(Δoi)と残存酸素の両方に仕様が立っていた場合エラーを返す)
'*
'*    パラメータ    : 変数名        ,IO ,型              ,説明
'*                　　pHin　　    　,I  ,tFullHinban   　,品番
'*
'*    戻り値        : 仕様チェック結果(-1:ｴﾗｰ，0:AOi仕様無，1:AOi仕様有)
'*
'*****************************************************************************************
Public Function ChkAoiSiyou(pHIN As tFullHinban) As Integer
    Dim sSQL            As String
    Dim rs              As OraDynaset
    Dim sDoiSiyou(2)    As String       '検査有無(DOi1～3)
    Dim sAoiSiyou       As String       '検査有無(AOi)
    Dim intCnt          As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmzcF_cmkc001WF.bas -- Function ChkAoiSiyou"

    sSQL = "select HWFOS1HS, HWFOS2HS, HWFOS3HS, HWFZOHWS from TBCME025 "
    sSQL = sSQL & "where HINBAN = '" & pHIN.hinban & "' "
    sSQL = sSQL & "and MNOREVNO = " & pHIN.mnorevno & " "
    sSQL = sSQL & "and FACTORY = '" & pHIN.factory & "' "
    sSQL = sSQL & "and OPECOND = '" & pHIN.opecond & "' "

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount = 0 Then
        rs.Close
        ChkAoiSiyou = -1
        GoTo proc_exit
    End If

    If IsNull(rs("HWFOS1HS")) = False Then sDoiSiyou(0) = rs("HWFOS1HS") '品WF酸素析出1保証方法_処
    If IsNull(rs("HWFOS2HS")) = False Then sDoiSiyou(1) = rs("HWFOS2HS") '品WF酸素析出2保証方法_処
    If IsNull(rs("HWFOS3HS")) = False Then sDoiSiyou(2) = rs("HWFOS3HS") '品WF酸素析出3保証方法_処
    If IsNull(rs("HWFZOHWS")) = False Then sAoiSiyou = rs("HWFZOHWS")    '品WF残存酸素保証方法_処

    '酸素析出と残存酸素の仕様チェック
    ChkAoiSiyou = 0
    For intCnt = 0 To 2
        If sDoiSiyou(intCnt) = "H" Or sDoiSiyou(intCnt) = "S" Then
            '酸素析出(Δoi)と残存酸素の両方に仕様が立っていた場合はエラー
            If sAoiSiyou = "H" Or sAoiSiyou = "S" Then
                ChkAoiSiyou = -1
                Exit For
            End If
        Else
            If sAoiSiyou = "H" Or sAoiSiyou = "S" Then
                ChkAoiSiyou = 1
            End If
        End If
    Next

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    ChkAoiSiyou = -1
    Resume proc_exit
End Function

'***********************************************************************************************
'*    関数名        : DBDRV_GetTBCMJ016Cnt
'*
'*    処理概要      : 1.SPV実績存在ﾁｪｯｸ
'*
'*    パラメータ    : 変数名        ,IO ,型              ,説明
'*      　　          sSampleid　　 ,I  ,String        　,ｻﾝﾌﾟﾙID
'*      　　          iRecCnt　　   ,O  ,Integer         ,ﾚｺｰﾄﾞ数
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'***********************************************************************************************
Public Function DBDRV_GetTBCMJ016Cnt(sSampleid As String, iRecCnt As Integer) As FUNCTION_RETURN
    Dim sSQL    As String
    Dim rs      As OraDynaset

    'ｻﾝﾌﾟﾙIDと保証ﾌﾗｸﾞを元にSPV実績を取得
    sSQL = "SELECT CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, HSFLG, SMPLNO "
    sSQL = sSQL & "FROM TBCMJ016 "
    sSQL = sSQL & "WHERE SMPLNO = '" & sSampleid & "' "
    sSQL = sSQL & "AND HSFLG = '1' "

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)
    If rs Is Nothing Then
        iRecCnt = 0
        DBDRV_GetTBCMJ016Cnt = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    '抽出結果のﾚｺｰﾄﾞ数を登録
    iRecCnt = rs.RecordCount
    rs.Close

    DBDRV_GetTBCMJ016Cnt = FUNCTION_RETURN_SUCCESS

End Function

'***************************************************************************************
'*    関数名        : DBDRV_WARPMAPGET
'*
'*    処理概要      : 1.WFﾏｯﾌﾟﾃﾞｰﾀ取得(Warp判定用)
'*
'*    パラメータ    : 変数名        ,IO ,型                  ,説明
'*          　　      tWarpMapTmp() ,I  ,type_DBDRV_Nukisi   ,WFﾏｯﾌﾟﾃﾞｰﾀ
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'***************************************************************************************
Public Function DBDRV_WARPMAPGET(tWarpMapTmp() As type_DBDRV_Nukisi) As FUNCTION_RETURN
    Dim i, j, k, m, n   As Integer
    Dim sSQL            As String
    Dim rs              As OraDynaset

    'エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_WARPMAPGET"

    m = 0
    ReDim sWrpLOTID(0)
    ReDim iWrpBLOCKSEQ(0)
    ReDim tWarpMapTmp(0)

    For i = 0 To UBound(tSXLID)
        sSQL = "select "
        sSQL = sSQL & "LOTID, "
        sSQL = sSQL & "BLOCKSEQ, "
        sSQL = sSQL & "MSXLID, "
        sSQL = sSQL & "MHINBAN, "
        sSQL = sSQL & "MREVNUM, "
        sSQL = sSQL & "MFACTORY, "
        sSQL = sSQL & "MOPECOND, "
        sSQL = sSQL & "SHAFLAG, "
        sSQL = sSQL & "MSMPLEID "
        sSQL = sSQL & "from TBCMY011 "
        sSQL = sSQL & "where LOTID = '" & tSXLID(i).LOTID & "' "
        sSQL = sSQL & "and MSXLID = '" & tSXLID(i).SXLID & "' "
        sSQL = sSQL & "order by BLOCKSEQ "

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount = 0 Then
            rs.Close
            DBDRV_WARPMAPGET = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

        k = UBound(sWrpLOTID)
        n = rs.RecordCount
        j = 0
        ReDim Preserve sWrpLOTID(k + n)         'ﾌﾞﾛｯｸID
        ReDim Preserve iWrpBLOCKSEQ(k + n)      'ﾌﾞﾛｯｸ内連番

        Do While Not rs.EOF
            j = j + 1
            'ﾌﾞﾛｯｸID
            If IsNull(rs("LOTID")) Then
                sWrpLOTID(k + j) = ""
            Else
                sWrpLOTID(k + j) = rs("LOTID")
            End If
            'ﾌﾞﾛｯｸ内連番
            If IsNull(rs("BLOCKSEQ")) Then
                iWrpBLOCKSEQ(k + j) = 0
            Else
                iWrpBLOCKSEQ(k + j) = rs("BLOCKSEQ")
            End If
            rs.MoveNext
        Loop

        'SXLのTOP側
        rs.MoveFirst
        m = m + 1
        ReDim Preserve tWarpMapTmp(m)
        With tWarpMapTmp(m)
            'ﾌﾞﾛｯｸID
            If IsNull(rs("LOTID")) = False Then .LOTID = rs("LOTID") Else .LOTID = vbNullString
            'ﾌﾞﾛｯｸ内連番
            If IsNull(rs("BLOCKSEQ")) = False Then .BLOCKSEQ = rs("BLOCKSEQ") Else .BLOCKSEQ = "0"
            'SXLID
            If IsNull(rs("MSXLID")) = False Then .SXLID = rs("MSXLID") Else .SXLID = vbNullString
            '品番
            If IsNull(rs("MHINBAN")) = False Then .hinban = rs("MHINBAN") Else .hinban = vbNullString
            '製品番号改訂番号
            If IsNull(rs("MREVNUM")) = False Then .REVNUM = rs("MREVNUM") Else .REVNUM = 0
            '工場
            If IsNull(rs("MFACTORY")) = False Then .factory = rs("MFACTORY") Else .factory = vbNullString
            '操業条件
            If IsNull(rs("MOPECOND")) = False Then .opecond = rs("MOPECOND") Else .opecond = vbNullString
            '抜試位置(ｻﾝﾌﾟﾙID)
            If IsNull(rs("MSMPLEID")) = False Then .SMPLEID = rs("MSMPLEID") Else .SMPLEID = vbNullString
            'ｻﾝﾌﾟﾙﾌﾗｸﾞ
            If IsNull(rs("SHAFLAG")) = False Then .SHAFLAG = rs("SHAFLAG") Else .SHAFLAG = vbNullString
            If Trim(.SHAFLAG) = "1" Then
                If Trim(.SMPLEID) = vbNullString Then
                    DBDRV_WARPMAPGET = FUNCTION_RETURN_FAILURE
                    rs.Close
                    GoTo proc_exit
                End If
            End If
        End With

        'SXLのBOT側
        rs.MoveLast
        m = m + 1
        ReDim Preserve tWarpMapTmp(m)
        With tWarpMapTmp(m)
            'ﾌﾞﾛｯｸID
            If IsNull(rs("LOTID")) = False Then .LOTID = rs("LOTID") Else .LOTID = vbNullString
            'ﾌﾞﾛｯｸ内連番
            If IsNull(rs("BLOCKSEQ")) = False Then .BLOCKSEQ = rs("BLOCKSEQ") Else .BLOCKSEQ = "0"
            'SXLID
            If IsNull(rs("MSXLID")) = False Then .SXLID = rs("MSXLID") Else .SXLID = vbNullString
            '品番
            If IsNull(rs("MHINBAN")) = False Then .hinban = rs("MHINBAN") Else .hinban = vbNullString
            '製品番号改訂番号
            If IsNull(rs("MREVNUM")) = False Then .REVNUM = rs("MREVNUM") Else .REVNUM = 0
            '工場
            If IsNull(rs("MFACTORY")) = False Then .factory = rs("MFACTORY") Else .factory = vbNullString
            '操業条件
            If IsNull(rs("MOPECOND")) = False Then .opecond = rs("MOPECOND") Else .opecond = vbNullString
            '抜試位置(ｻﾝﾌﾟﾙID)
            If IsNull(rs("MSMPLEID")) = False Then .SMPLEID = rs("MSMPLEID") Else .SMPLEID = vbNullString
            'ｻﾝﾌﾟﾙﾌﾗｸﾞ
            If IsNull(rs("SHAFLAG")) = False Then .SHAFLAG = rs("SHAFLAG") Else .SHAFLAG = vbNullString
            If Trim(.SHAFLAG) = "1" Then
                If Trim(.SMPLEID) = vbNullString Then
                    DBDRV_WARPMAPGET = FUNCTION_RETURN_FAILURE
                    rs.Close
                    GoTo proc_exit
                End If
            End If
        End With
        rs.Close
    Next i

    DBDRV_WARPMAPGET = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    DBDRV_WARPMAPGET = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'***************************************************************************************
'*    関数名        : DBDRV_KanrenBlk
'*
'*    処理概要      : 1.関連ﾌﾞﾛｯｸ紐付紐切(TBCMY023)登録
'*
'*    パラメータ    : 変数名      ,IO ,型                 ,説明
'*      　　      　　sCrynum     ,I  ,String         　  ,結晶番号
'*      　　      　　sKblockid() ,I  ,type_DBDRV_LOTSXL  ,関連ﾌﾞﾛｯｸ
'*      　　      　　iSpos       ,I  ,Integer        　  ,結晶内開始位置
'*                　　iEpos       ,I  ,Integer        　  ,結晶内終了位置
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'***************************************************************************************
Public Function DBDRV_KanrenBlk(sCryNum As String, sKblockid() As type_DBDRV_LOTSXL, _
                                iSpos As Integer, iEpos As Integer) As FUNCTION_RETURN
    Dim sSQL            As String
    Dim i, j            As Long
    Dim rs              As OraDynaset
    Dim lngRecCnt       As Long             'ﾚｺｰﾄﾞ数
    Dim sLotid          As String           'ﾌﾞﾛｯｸID(WFﾏｯﾌﾟ)
    Dim sSXLID          As String           'SXLID(WFﾏｯﾌﾟ)
    Dim udtKanrenData() As typ_TBCMY023     '関連ﾌﾞﾛｯｸ紐付紐切ﾃﾞｰﾀ
    Dim blCutFlg        As Boolean          '関連ﾌﾞﾛｯｸ紐切りﾌﾗｸﾞ
    Dim intTrnCnt       As Integer          '処理回数

    '' エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_KanrenBlk"

    DBDRV_KanrenBlk = FUNCTION_RETURN_FAILURE

    '処理回数取得
    sSQL = "SELECT NVL(MAX(TRANCNT),0) MAXCNT FROM TBCMY023"
    sSQL = sSQL & " WHERE CRYNUM = '" & sCryNum & "'"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    If rs.RecordCount = 0 Then
        intTrnCnt = 1
    Else
        intTrnCnt = rs("MAXCNT") + 1          '処理回数(最大) + 1
    End If
    rs.Close

    lngRecCnt = 0             '登録ﾚｺｰﾄﾞ数
    blCutFlg = False         '関連ﾌﾞﾛｯｸ紐切りﾌﾗｸﾞ(False:紐切り無)

    '関連ﾌﾞﾛｯｸ紐切ﾃﾞｰﾀｾｯﾄ
    For i = 0 To UBound(sKblockid)
        lngRecCnt = lngRecCnt + 1
        ReDim Preserve udtKanrenData(lngRecCnt)
        With udtKanrenData(lngRecCnt)
            .CRYNUM = sCryNum               '結晶番号
            .TRANCNT = intTrnCnt              '処理回数
            .BLOCKID = sKblockid(i).LOTID   'ﾌﾞﾛｯｸID
            .PROCCAT = "D"                  '処理区分(D:紐切)
            .TXID = "TX879I"                'ﾄﾗﾝｻﾞｸｼｮﾝID
        End With
    Next i

    'WFﾏｯﾌﾟよりﾌﾞﾛｯｸID,SXLIDを取得
    sSQL = "SELECT LOTID, MSXLID FROM TBCMY011"
    sSQL = sSQL & " WHERE LOTID LIKE '" & left(sCryNum, 9) & "%'"
    sSQL = sSQL & " AND (WFSTA = '0' OR WFSTA = '1')"
    sSQL = sSQL & " AND RITOP_POS > " & iSpos
    sSQL = sSQL & " AND RITOP_POS <= " & iEpos
    sSQL = sSQL & " AND MSXLID IS NOT NULL"
    sSQL = sSQL & " GROUP BY LOTID, MSXLID"
    sSQL = sSQL & " ORDER BY LOTID, MAX(BLOCKSEQ)"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    'ﾃﾞｰﾀなし
    If rs.RecordCount = 0 Then
        rs.Close
        GoTo proc_exit
    End If

    '関連ﾌﾞﾛｯｸ紐付ﾃﾞｰﾀｾｯﾄ
    For i = 1 To rs.RecordCount
        If i > 1 Then
            '別ﾌﾞﾛｯｸで同一SXL(関連ﾌﾞﾛｯｸ○)
            If sLotid <> rs("LOTID") And sSXLID = rs("MSXLID") Then
                '関連ﾌﾞﾛｯｸ(上)
                If udtKanrenData(lngRecCnt).BLOCKID <> sLotid Then
                    intTrnCnt = intTrnCnt + 1       '処理回数
                    lngRecCnt = lngRecCnt + 1
                    ReDim Preserve udtKanrenData(lngRecCnt)
                    With udtKanrenData(lngRecCnt)
                        .CRYNUM = sCryNum               '結晶番号
                        .TRANCNT = intTrnCnt              '処理回数
                        .BLOCKID = sLotid               'ﾌﾞﾛｯｸID
                        .PROCCAT = "C"                  '処理区分(C:付替え)
                        .TXID = "TX879I"                'ﾄﾗﾝｻﾞｸｼｮﾝID
                    End With
                End If
                '関連ﾌﾞﾛｯｸ(下)
                lngRecCnt = lngRecCnt + 1
                ReDim Preserve udtKanrenData(lngRecCnt)
                With udtKanrenData(lngRecCnt)
                    .CRYNUM = sCryNum                   '結晶番号
                    .TRANCNT = intTrnCnt                  '処理回数
                    .BLOCKID = rs("LOTID")              'ﾌﾞﾛｯｸID
                    .PROCCAT = "C"                      '処理区分(C:付替え)
                    .TXID = "TX879I"                    'ﾄﾗﾝｻﾞｸｼｮﾝID
                End With

            '別ﾌﾞﾛｯｸで別SXL(関連ﾌﾞﾛｯｸ×)
            ElseIf sLotid <> rs("LOTID") And sSXLID <> rs("MSXLID") Then
                blCutFlg = True          '関連ﾌﾞﾛｯｸ紐切りﾌﾗｸﾞ(True:紐切り有)
            End If
        End If
        sLotid = rs("LOTID")        'ﾌﾞﾛｯｸID
        sSXLID = rs("MSXLID")       'SXLID
        rs.MoveNext
    Next i
    rs.Close

    '関連ﾌﾞﾛｯｸ紐切りが発生した場合、関連ﾌﾞﾛｯｸ紐付紐切(TBCMY023)に登録
    If blCutFlg Then
        For i = 1 To UBound(udtKanrenData)
            With udtKanrenData(i)
                sSQL = "INSERT INTO TBCMY023"
                sSQL = sSQL & " (CRYNUM,"
                sSQL = sSQL & " TRANCNT,"
                sSQL = sSQL & " BLOCKID,"
                sSQL = sSQL & " PROCCAT,"
                sSQL = sSQL & " TXID,"
                sSQL = sSQL & " REGDATE,"
                sSQL = sSQL & " SUMITFLAG,"               '07/12/21 ooba
                sSQL = sSQL & " SUMITSND,"                '07/12/21 ooba
                sSQL = sSQL & " SSENDNO,"                 '07/12/21 ooba
                sSQL = sSQL & " SENDFLAG,"
                sSQL = sSQL & " SENDDATE, "
                sSQL = sSQL & " PLANTCAT) "
                sSQL = sSQL & " VALUES"
                sSQL = sSQL & " ('" & .CRYNUM & "',"      '結晶番号
                sSQL = sSQL & .TRANCNT & ","              '処理回数
                sSQL = sSQL & " '" & .BLOCKID & "',"      'ﾌﾞﾛｯｸID
                sSQL = sSQL & " '" & .PROCCAT & "',"      '処理区分
                sSQL = sSQL & " '" & .TXID & "',"         'ﾄﾗﾝｻﾞｸｼｮﾝID
                sSQL = sSQL & " SYSDATE,"                 '登録日付
                sSQL = sSQL & " '0',"                     'SUMIT送信ﾌﾗｸﾞ  07/12/21 ooba
                sSQL = sSQL & " NULL,"                    'SUMIT送信日付  07/12/21 ooba
                sSQL = sSQL & " NULL,"                    '送信順連番  07/12/21 ooba
                sSQL = sSQL & " '0',"                     '送信ﾌﾗｸﾞ
                sSQL = sSQL & " NULL, "                    '送信日付
                sSQL = sSQL & "  '" & sCmbMukesaki & "') "  '向先
            End With

            '登録失敗
            If OraDB.ExecuteSQL(sSQL) <= 0 Then
                GoTo proc_exit
            End If
        Next i
    End If

    DBDRV_KanrenBlk = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    Resume proc_exit
End Function
' add SETkimizuka Start 09/03/17
'***************************************************************************************
'*    関数名        : DBDRV_XODY4GET
'*
'*    処理概要      : 流動停止項目取得
'*
'*    パラメータ    : 変数名      ,IO ,型                 ,説明
'*      　　      　　sCrynum     ,I  ,String         　  ,結晶番号
'*      　　      　　sKblockid() ,I  ,type_DBDRV_LOTSXL  ,関連ﾌﾞﾛｯｸ
'*      　　      　　iSpos       ,I  ,Integer        　  ,結晶内開始位置
'*                　　iEpos       ,I  ,Integer        　  ,結晶内終了位置
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'***************************************************************************************
Public Function DBDRV_XODY4GET(udt_ww() As DBDRV_scmzc_fcmlc001b_SXL039) As FUNCTION_RETURN
    Dim sql             As String
    Dim rs              As OraDynaset
    Dim sOldID          As String
    Dim iCnt            As Integer
    Dim sSxl            As String
    
    sSxl = "("
    For iCnt = 1 To UBound(udt_ww)
        sSxl = sSxl & "'" & udt_ww(iCnt).SXLIDCA & "'"
        If iCnt < UBound(udt_ww) Then
            sSxl = sSxl & ","
        End If
    Next
    sSxl = sSxl & ")"
    
    
    ' 流動監視SQL修正 upd SETkimizuka Start  09/06/30
    sql = "SELECT "
    sql = sql & "   NVL(SXLIDY3,' ') as SXLIDY4"          '
    sql = sql & " , NVL(TO_CHAR(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4)),' ') as AGRSTATUSY4 "
    sql = sql & " , DECODE(CAUSEY4,NULL,' ',TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSEY4"        '
    sql = sql & " , NVL(Y4.PRINTKINDY4 || Y4.PRINTNOY4,' ') as PRINTNOY4"
    sql = sql & " , NVL(STOPY4,'0') as STOP "
    sql = sql & " , NVL(WKKTY4,' ') as WKKTY4 "
    sql = sql & " FROM XODY3 Y3,XODY4 Y4,KODA9 A9 "
    sql = sql & " WHERE  "
    sql = sql & "  SXLIDY3 IN " & sSxl
    sql = sql & " AND Y3.XTALNOY3 = Y4.XTALNOY4(+) "
    sql = sql & " AND Y3.RCNTY3 = Y4.RCNTY4(+) "
    sql = sql & " AND Y3.LIVKY3(+) = '0' "
    sql = sql & " AND Y4.LIVKY4(+) = '0' "
    sql = sql & " AND A9.SYSCA9(+) = 'X' AND A9.SHUCA9(+) = '30' AND Y4.CAUSEY4 = A9.CODEA9(+) "
'    sql = sql & " GROUP BY SXLIDY3,STOPY4,CAUSEY4,Y4.PRINTNOY4,Y4.PRINTKINDY4,NAMEJA9,AGRSTATUSY4,WKKTY4 "
    
    sql = sql & " UNION SELECT "
    sql = sql & "   NVL(SXLIDY3,' ') as SXLIDY4"          '
    sql = sql & " , NVL(TO_CHAR(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4)),' ') as AGRSTATUSY4 "
    sql = sql & " , DECODE(CAUSEY4,NULL,' ',TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSEY4"        '
    sql = sql & " , NVL(Y4.PRINTKINDY4 || Y4.PRINTNOY4,' ') as PRINTNOY4"
    sql = sql & " , NVL(STOPY4,'0') as STOP "
    sql = sql & " , NVL(WKKTY4,' ') as WKKTY4 "
    sql = sql & " FROM XODY3 Y3,XODY4 Y4,KODA9 A9 "
    sql = sql & " WHERE  "
    sql = sql & "  SXLIDY3 IN " & sSxl
    sql = sql & " AND Y3.XTALNOY3 = Y4.XTALNOY4(+) "
    sql = sql & " AND Y4.WKKTY4(+) = 'CW000'"
    sql = sql & " AND Y3.LIVKY3(+) = '0' "
    sql = sql & " AND Y4.LIVKY4(+) = '0' "
    sql = sql & " AND A9.SYSCA9(+) = 'X' AND A9.SHUCA9(+) = '30' AND Y4.CAUSEY4 = A9.CODEA9(+) "
'    sql = sql & " GROUP BY SXLIDY3,STOPY4,CAUSEY4,Y4.PRINTNOY4,Y4.PRINTKINDY4,NAMEJA9,AGRSTATUSY4,WKKTY4 "
    
'    sql = "SELECT "
'    sql = sql & "   NVL(SXLIDY3,' ') as SXLIDY4"          '
'    sql = sql & " , NVL(TO_CHAR(MIN(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4))),' ') as AGRSTATUSY4 "
'    sql = sql & " , DECODE(CAUSEY4,NULL,' ',TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSEY4"        '
'    sql = sql & " , NVL(Y5.PRINTKIND || Y5.PRINTNO,' ') as PRINTNOY4"
'    sql = sql & " , NVL(STOPY4,'0') as STOP "
'   sql = sql & "      FROM XODY3  "
'    sql = sql & "           LEFT OUTER JOIN (SELECT XODY4.AGRSTATUSY4,XODY4.CAUSEY4,XODY4.STOPY4,XODY4.XTALNOY4 FROM XODY3,XODY4 "
'    sql = sql & "           INNER JOIN (SELECT MIN(DECODE(WKKTY4,'CW750',3,'CW760',2,'CW000',1,9)) as WKKTY4 ,XTALNOY4 FROM XODY4  "
'    sql = sql & "              WHERE STOPY4 <> '2' AND LIVKY4 = '0' AND WKKTY4 ='CW000'"
'    sql = sql & "              GROUP BY XTALNOY4) Y4_WKKT ON (XODY4.XTALNOY4 = Y4_WKKT.XTALNOY4 AND XODY4.WKKTY4    = DECODE(Y4_WKKT.WKKTY4,3,'CW750',2,'CW760',1,'CW000',' ') ) "
'    sql = sql & "           WHERE XTALNOY3 = XODY4.XTALNOY4 AND LIVKY3 = '0' AND XODY4.STOPY4 <> '2' AND XODY4.LIVKY4 = '0' AND XODY4.WKKTY4 ='CW000'"
'    sql = sql & "           GROUP BY XODY4.AGRSTATUSY4,XODY4.CAUSEY4,XODY4.STOPY4,XODY4.XTALNOY4 ) XODY4  on ( XTALNOY3 = XTALNOY4" & ")"
'    sql = sql & "           LEFT OUTER JOIN KODA9 on ( SYSCA9 = 'X' AND SHUCA9 = '30' AND CAUSEY4 = CODEA9 ) "
'    sql = sql & "           LEFT OUTER JOIN (SELECT XTALNOY4 as XTALNO,SXLIDY4 as SXLID,PRINTNOY5 as PRINTNO,PRINTKINDY5 as PRINTKIND "
'    sql = sql & "                FROM XODY3,XODY4,XODY5 "
'    sql = sql & "              WHERE XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND LIVKY3 = '0' "
'    sql = sql & "                AND PRINTKINDY4 = PRINTKINDY5 AND PRINTNOY4 = PRINTNOY5  "
'    sql = sql & "                AND HKBNY5 ='0' GROUP BY XTALNOY4,SXLIDY4,PRINTNOY5,PRINTKINDY5) Y5 ON (XTALNOY3 = XTALNO AND SXLIDY3 = SXLID ) "
'    sql = sql & "      WHERE  "
'    sql = sql & "       ( AGRSTATUSY4 IS NOT NULL OR Y5.PRINTKIND IS NOT NULL) AND LIVKY3    = '0' AND SXLIDY3 IN " & sSxl
'    sql = sql & " GROUP BY SXLIDY3,STOPY4,CAUSEY4,Y5.PRINTNO,Y5.PRINTKIND,NAMEJA9,AGRSTATUSY4 "
'
'    sql = sql & " UNION SELECT "
'    sql = sql & "   NVL(SXLIDY3,' ') as SXLIDY4"          '
'    sql = sql & " , NVL(TO_CHAR(MIN(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4))),' ') as AGRSTATUSY4 "
'    sql = sql & " , DECODE(CAUSEY4,NULL,' ',TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSEY4"        '
'    sql = sql & " , NVL(Y5.PRINTKIND || Y5.PRINTNO,' ') as PRINTNOY4"
'    sql = sql & " , NVL(STOPY4,'0') as STOP "
'    sql = sql & "      FROM XODY3  "
'    sql = sql & "           LEFT OUTER JOIN (SELECT XODY4.AGRSTATUSY4,XODY4.CAUSEY4,XODY4.STOPY4,XODY4.SXLIDY4 FROM XODY3,XODY4 "
'    sql = sql & "           INNER JOIN (SELECT MIN(DECODE(WKKTY4,'CW750',3,'CW760',2,'CW000',1,9)) as WKKTY4 ,XTALNOY4,SXLIDY4 FROM XODY4  "
'    sql = sql & "              WHERE STOPY4 <> '2' AND LIVKY4 = '0' AND WKKTY4 in ('CW750','CW760')"
'    sql = sql & "              GROUP BY XTALNOY4,SXLIDY4) Y4_WKKT ON (XODY4.XTALNOY4 = Y4_WKKT.XTALNOY4 AND XODY4.SXLIDY4 = Y4_WKKT.SXLIDY4 AND XODY4.WKKTY4    = DECODE(Y4_WKKT.WKKTY4,3,'CW750',2,'CW760',1,'CW000',' ') ) "
'    sql = sql & "           WHERE XTALNOY3 = XODY4.XTALNOY4 AND RCNTY3 = XODY4.RCNTY4 AND LIVKY3 = '0' AND XODY4.STOPY4 <> '2' AND XODY4.LIVKY4 = '0' AND XODY4.WKKTY4 IN ('CW750','CW760')"
'    sql = sql & "           GROUP BY XODY4.AGRSTATUSY4,XODY4.CAUSEY4,XODY4.STOPY4,XODY4.SXLIDY4 ) XODY4  on ( SXLIDY3 = SXLIDY4" & ")"
'    sql = sql & "           LEFT OUTER JOIN KODA9 on ( SYSCA9 = 'X' AND SHUCA9 = '30' AND CAUSEY4 = CODEA9 ) "
'    sql = sql & "           LEFT OUTER JOIN (SELECT XTALNOY4 as XTALNO,SXLIDY4 as SXLID,PRINTNOY5 as PRINTNO,PRINTKINDY5 as PRINTKIND "
'    sql = sql & "                FROM XODY3,XODY4,XODY5 "
'    sql = sql & "              WHERE XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND LIVKY3 = '0' "
'    sql = sql & "                AND PRINTKINDY4 = PRINTKINDY5 AND PRINTNOY4 = PRINTNOY5  "
'    sql = sql & "                AND HKBNY5 ='0' GROUP BY XTALNOY4,SXLIDY4,PRINTNOY5,PRINTKINDY5) Y5 ON (XTALNOY3 = XTALNO AND SXLIDY3 = SXLID ) "
'    sql = sql & "      WHERE  "
'    sql = sql & "       LIVKY3    = '0' AND SXLIDY3 IN " & sSxl
'    sql = sql & " GROUP BY SXLIDY3,STOPY4,CAUSEY4,Y5.PRINTNO,Y5.PRINTKIND,NAMEJA9,AGRSTATUSY4 "
'    sql = sql & " ORDER BY SXLIDY3"
    ' 流動監視SQL修正 upd SETkimizuka End  09/06/30
'Debug.Print sql
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    Do While Not rs.EOF
        
        For iCnt = 1 To UBound(udt_ww)
            If udt_ww(iCnt).SXLIDCA = rs("SXLIDY4") Then
                If rs("STOP") <> "2" And (rs("WKKTY4") = "CW750" Or rs("WKKTY4") = "CW760" Or rs("WKKTY4") = "CW000") Then
                    ' 流動監視SQL修正 upd SETkimizuka Start  09/06/30
                    'udt_ww(iCnt).STOP = rs("STOP")
                    'udt_ww(iCnt).AGRSTATUS = rs("AGRSTATUSY4")
                    If Trim(udt_ww(iCnt).AGRSTATUS) = "" Or rs("AGRSTATUSY4") < udt_ww(iCnt).AGRSTATUS Then
                        udt_ww(iCnt).STOP = rs("STOP")
                        udt_ww(iCnt).AGRSTATUS = rs("AGRSTATUSY4")
                    End If
                    ' 流動監視SQL修正 upd SETkimizuka End  09/06/30
                    If Trim(rs("CAUSEY4")) <> "" And InStr(udt_ww(iCnt).CAUSE, rs("CAUSEY4")) = 0 Then
                        udt_ww(iCnt).CAUSE = udt_ww(iCnt).CAUSE & rs("CAUSEY4") & vbTab
                    End If
                End If
                If Trim(rs("PRINTNOY4")) <> "" And InStr(udt_ww(iCnt).PRINTNO, rs("PRINTNOY4")) = 0 Then
                    udt_ww(iCnt).PRINTNO = udt_ww(iCnt).PRINTNO & rs("PRINTNOY4") & vbTab
                End If
                Exit For
            End If
        Next
        rs.MoveNext
    Loop

    rs.Close
    
proc_exit:
    '' 終了
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function
' add SETkimizuka End 09/03/17

' add SPK_Hitomi Start 09/10/20
'********************************************************************************************************
'*    関数名        : ChkHosho
'*
'*    処理概要      : 1.保証方法判定
'*                    (XSDCWの確定区分より、ﾌﾞﾛｯｸ,WF保証を判定する)
'*
'*    パラメータ    : 変数名        ,IO ,型         ,説明
'*                    inSXLID       ,I  ,String     , SXL-ID
'*                    outHosho  　  ,O  ,String     , 保証方法(1:ﾌﾞﾛｯｸ保証,2:WF保証)
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************************
Public Function ChkHosho(inSXLID As String, outHosho As String) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sSQL        As String


    'エラーハンドラの設定
    On Error GoTo proc_err

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    ChkHosho = FUNCTION_RETURN_SUCCESS

    sSQL = "select REPSMPLIDCW,WFSMPLIDGDCW from XSDCW where SXLIDCW = '" & inSXLID & "' and KTKBNCW = '9' and LIVKCW = '0'"
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    If rs.RecordCount > 0 Then
        outHosho = 1 'ﾌﾞﾛｯｸ保証
    ElseIf rs.RecordCount = 0 Then
        outHosho = 2 'WF保証
    End If
    
    Set rs = Nothing

proc_exit:

    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    ChkHosho = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START
'---------------------------------------------------------------------------
'概要      :結晶番号上９桁よりTBCMJ022を検索し、SIRD検査情報を返す
'---------------------------------------------------------------------------
'ﾊﾟﾗﾒｰﾀ    :変数名      ,IO     ,型                     ,説明
'          :pCRYNUM     ,I  　　,String                 ,結晶番号
'          :pflgSird    ,O  　　,Boolean                ,SIRD検査情報の有無(True:有、False：無)
'          :pSMPLID     ,O  　　,String                 ,SIRD検査情報の代表ｻﾝﾌﾟﾙID
'          :戻り値      ,O      ,Boolean                ,[True:OK／False:NG]
'---------------------------------------------------------------------------
Public Function fncGetSirdSample(ByVal pCRYNUM As String, ByRef pflgSird As Boolean, ByRef pSMPLID As String) As Boolean

    Dim sql         As String           'SQL全体
    Dim rs          As OraDynaset       'RecordSet

    '--ｴﾗｰﾊﾝﾄﾞﾗの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function fncGetSirdSample"
    
    '--初期化
    fncGetSirdSample = False: pflgSird = False: pSMPLID = ""
    Set rs = Nothing      'Oracle RecordSet Free

    '--SQL文生成
    sql = "select SMPLNO  "
    sql = sql & "from TBCMJ022 " & vbCrLf
    sql = sql & "where" & vbCrLf
'Cng Start 2011/10/06 Y.Hitomi
    sql = sql & "     substr(CRYNUM,1,9) = '" & left(pCRYNUM, 9) & "'" & vbCrLf     '結晶番号(上9桁)
'    sql = sql & "     substr(CRYNUM,1,7) = '" & left(pCRYNUM, 7) & "'" & vbCrLf     '結晶番号(上7桁)
'Cng Start 2011/10/06 Y.Hitomi
    sql = sql & " and TRANCNT   = 0" & vbCrLf                                       '処理回数
    

    '--ﾃﾞｰﾀを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_READONLY Or ORADYN_NOCACHE)
    If rs Is Nothing Then
        GoTo proc_exit
    End If
    
    '--抽出結果参照
    If Not (rs.EOF) Then
        '<< ﾃﾞｰﾀ有り >>
        rs.MoveFirst
        pflgSird = True                 '[SIRD検査情報有り]
        pSMPLID = rs("SMPLNO")          '[代表ｻﾝﾌﾟﾙID]
    End If
    
    fncGetSirdSample = True

proc_exit:
    '<< 終了 >>
    'Oracle RecordSet Free
    If Not (rs Is Nothing) Then
        rs.Close
    End If
    Set rs = Nothing
    
    gErr.Pop
    Exit Function

proc_err:
    '<< ｴﾗｰﾊﾝﾄﾞﾗ >>
    'Oracle RecordSet Free
    If Not (rs Is Nothing) Then
        rs.Close
    End If
    Set rs = Nothing

    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    
    gErr.HandleError
    Resume proc_exit

End Function
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD END

'概要      :中間抜試単位(mm)の取得
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO   ,型                ,説明
'　　      :HIN  　　 　,I    ,tFullHinban 　    ,12桁品番
'　　      :iMSMPTANI　 ,O    ,Integer 　        ,中間抜試単位(mm)
'      　　:戻り値      ,O    ,FUNCTION_RETURN   ,抽出の成否
'説明      :
'履歴      :2011/06/30 Marushita
Public Function getMSMPTANI(HIN As tFullHinban, iMSMPTANI As Integer) As FUNCTION_RETURN

    Dim sSQL As String
    Dim rs As OraDynaset
    
    getMSMPTANI = FUNCTION_RETURN_FAILURE
        
    iMSMPTANI = 0
    If Trim(HIN.hinban) = "Z" Or Trim(HIN.hinban) = "G" Or Trim(HIN.hinban) = "" Then
        Exit Function
    End If
    
    sSQL = "SELECT MSMPTANI"
    sSQL = sSQL & " FROM TBCME036"
    sSQL = sSQL & " WHERE"
    sSQL = sSQL & " HINBAN = '" & HIN.hinban & "'"
    sSQL = sSQL & " AND MNOREVNO = " & HIN.mnorevno
    sSQL = sSQL & " AND FACTORY = '" & HIN.factory & "'"
    sSQL = sSQL & " AND OPECOND = '" & HIN.opecond & "'"
    
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    If rs.RecordCount > 0 Then
        If IsNull(rs.Fields("MSMPTANI")) = False Then iMSMPTANI = rs.Fields("MSMPTANI") Else iMSMPTANI = 0
        getMSMPTANI = FUNCTION_RETURN_SUCCESS
    End If
    rs.Close
    
End Function
'概要      :シングル確定可否のチェック
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO   ,型                ,説明
'　　      :HIN  　　 　,I    ,tFullHinban 　    ,12桁品番
'      　　:sSXLIDFLG   ,O    ,String   　　　　 ,SXLID確定可否フラグ
'      　　:戻り値      ,O    ,FUNCTION_RETURN   ,抽出の成否
'説明      :
'履歴      :2011/09/29 Y.Hitomi
Public Function getSXLIDFLG(HIN As tFullHinban, sSXLIDFLG) As FUNCTION_RETURN

    Dim sSQL As String
    Dim rs As OraDynaset
    
    getSXLIDFLG = FUNCTION_RETURN_FAILURE
        
    If Trim(HIN.hinban) = "Z" Or Trim(HIN.hinban) = "G" Or Trim(HIN.hinban) = "" Then
        Exit Function
    End If
    
    sSQL = "SELECT NVL(SXLIDFLG,'0') as SXLIDFLG "
    sSQL = sSQL & " FROM TBCME036"
    sSQL = sSQL & " WHERE"
    sSQL = sSQL & " HINBAN = '" & HIN.hinban & "'"
    sSQL = sSQL & " AND MNOREVNO = " & HIN.mnorevno
    sSQL = sSQL & " AND FACTORY = '" & HIN.factory & "'"
    sSQL = sSQL & " AND OPECOND = '" & HIN.opecond & "'"
    
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)
'Add Start 2011/10/03 Y.Hitomi
    If rs.RecordCount > 0 And IsNull(rs.Fields("SXLIDFLG")) = False Then
'    If rs.RecordCount > 0 Then
'Add End 2011/10/03 Y.Hitomi
        sSXLIDFLG = rs.Fields("SXLIDFLG")
        getSXLIDFLG = FUNCTION_RETURN_SUCCESS
    Else
        sSXLIDFLG = "0"
    End If
           
    rs.Close
    
End Function

