Attribute VB_Name = "s_cmzcCrystal"
Option Explicit
'                                     2001/05/31
'================================================
' クラス・ユーザ定義型の変換プロシージャ
' 定義内容: 060207_結晶管理
'================================================


'------ テーブル名:TBCME037    ---- 結晶情報

' ユーザ定義型をクラスに変換する(結晶情報)
Public Function u2c_TBCME037(data As typ_TBCME037) As c_TBCME037
Dim cls As New c_TBCME037       '変換先クラス

    cls.CRYNUM = data.CRYNUM                '結晶番号
    cls.KRPROCCD = data.KRPROCCD            '管理工程コード
    cls.PROCCD = data.PROCCD                '工程コード
    cls.RPHINBAN = data.RPHINBAN            'ねらい品番
    cls.RPREVNUM = data.RPREVNUM            'ねらい品番製品番号改訂番号
    cls.RPFACT = data.RPFACT                'ねらい品番工場
    cls.RPOPCOND = data.RPOPCOND            'ねらい品番操業条件
    cls.PRODCOND = data.PRODCOND            '製作条件
    cls.PGID = data.PGID                    'ＰＧ－ＩＤ
    cls.UPLENGTH = data.UPLENGTH            '引上げ長さ
    cls.TOPLENG = data.TOPLENG              'ＴＯＰ長さ
    cls.BODYLENG = data.BODYLENG            '直胴長さ
    cls.BOTLENG = data.BOTLENG              'ＢＯＴ長さ
    cls.FREELENG = data.FREELENG            'フリー長
    cls.DIAMETER = data.DIAMETER            '直径
    cls.CHARGE = data.CHARGE                'チャージ量
    cls.SEED = data.SEED                    'シード
    cls.ADDDPCLS = data.ADDDPCLS            '追加ドープ種類
    cls.ADDDPPOS = data.ADDDPPOS            '追加ドープ位置
    cls.ADDDPVAL = data.ADDDPVAL            '追加ドープ量
    'cls.REGDATE = data.REGDATE              '登録日付
    'cls.UPDDATE = data.UPDDATE              '更新日付
    'cls.SENDFLAG = data.SENDFLAG            '送信フラグ
    'cls.SENDDATE = data.SENDDATE            '送信日付

    Set u2c_TBCME037 = cls
End Function


' クラスをユーザ定義型に変換する(結晶情報)
Public Function c2u_TBCME037(cls As c_TBCME037) As typ_TBCME037
Dim data As typ_TBCME037        '変換先ユーザ定義型

    data.CRYNUM = cls.CRYNUM                '結晶番号
    data.KRPROCCD = cls.KRPROCCD            '管理工程コード
    data.PROCCD = cls.PROCCD                '工程コード
    data.RPHINBAN = cls.RPHINBAN            'ねらい品番
    data.RPREVNUM = cls.RPREVNUM            'ねらい品番製品番号改訂番号
    data.RPFACT = cls.RPFACT                'ねらい品番工場
    data.RPOPCOND = cls.RPOPCOND            'ねらい品番操業条件
    data.PRODCOND = cls.PRODCOND            '製作条件
    data.PGID = cls.PGID                    'ＰＧ－ＩＤ
    data.UPLENGTH = cls.UPLENGTH            '引上げ長さ
    data.TOPLENG = cls.TOPLENG              'ＴＯＰ長さ
    data.BODYLENG = cls.BODYLENG            '直胴長さ
    data.BOTLENG = cls.BOTLENG              'ＢＯＴ長さ
    data.FREELENG = cls.FREELENG            'フリー長
    data.DIAMETER = cls.DIAMETER            '直径
    data.CHARGE = cls.CHARGE                'チャージ量
    data.SEED = cls.SEED                    'シード
    data.ADDDPCLS = cls.ADDDPCLS            '追加ドープ種類
    data.ADDDPPOS = cls.ADDDPPOS            '追加ドープ位置
    data.ADDDPVAL = cls.ADDDPVAL            '追加ドープ量
    'data.REGDATE = cls.REGDATE              '登録日付
    'data.UPDDATE = cls.UPDDATE              '更新日付
    'data.SENDFLAG = cls.SENDFLAG            '送信フラグ
    'data.SENDDATE = cls.SENDDATE            '送信日付

    c2u_TBCME037 = data
End Function


'------ テーブル名:TBCME038    ---- ブロック設計

' ユーザ定義型をクラスに変換する(結晶情報)
Public Function u2c_TBCME038(data As typ_TBCME038) As c_TBCME038
Dim cls As New c_TBCME038       '変換先クラス

    cls.CRYNUM = data.CRYNUM                '結晶番号
    cls.IngotPos = data.IngotPos            '結晶内開始位置
    cls.LENGTH = data.LENGTH                '長さ
    cls.USECLASS = data.USECLASS            '使用区分
    'cls.REGDATE = data.REGDATE              '登録日付
    'cls.UPDDATE = data.UPDDATE              '更新日付
    'cls.SENDFLAG = data.SENDFLAG            '送信フラグ
    'cls.SENDDATE = data.SENDDATE            '送信日付

    Set u2c_TBCME038 = cls
End Function


' クラスをユーザ定義型に変換する(ブロック設計)
Public Function c2u_TBCME038(cls As c_TBCME038) As typ_TBCME038
Dim data As typ_TBCME038        '変換先ユーザ定義型

    data.CRYNUM = cls.CRYNUM                '結晶番号
    data.IngotPos = cls.IngotPos            '結晶内開始位置
    data.LENGTH = cls.LENGTH                '長さ
    data.USECLASS = cls.USECLASS            '使用区分
    'data.REGDATE = cls.REGDATE              '登録日付
    'data.UPDDATE = cls.UPDDATE              '更新日付
    'data.SENDFLAG = cls.SENDFLAG            '送信フラグ
    'data.SENDDATE = cls.SENDDATE            '送信日付

    c2u_TBCME038 = data
End Function


'------ テーブル名:TBCME039    ---- 品番設計

' ユーザ定義型をクラスに変換する(結晶情報)
Public Function u2c_TBCME039(data As typ_TBCME039) As c_TBCME039
Dim cls As New c_TBCME039       '変換先クラス

    cls.CRYNUM = data.CRYNUM                '結晶番号
    cls.IngotPos = data.IngotPos            '結晶内開始位置
    cls.HINBAN = data.HINBAN                '品番
    cls.REVNUM = data.REVNUM                '改訂番号
    cls.FACT = data.FACT                    '工場
    cls.OPCOND = data.OPCOND                '操業条件
    cls.LENGTH = data.LENGTH                '長さ
    cls.USECLASS = data.USECLASS            '使用区分
    'cls.REGDATE = data.REGDATE              '登録日付
    'cls.UPDDATE = data.UPDDATE              '更新日付
    'cls.SENDFLAG = data.SENDFLAG            '送信フラグ
    'cls.SENDDATE = data.SENDDATE            '送信日付

    Set u2c_TBCME039 = cls
End Function


' クラスをユーザ定義型に変換する(品番設計)
Public Function c2u_TBCME039(cls As c_TBCME039) As typ_TBCME039
Dim data As typ_TBCME039        '変換先ユーザ定義型

    data.CRYNUM = cls.CRYNUM                '結晶番号
    data.IngotPos = cls.IngotPos            '結晶内開始位置
    data.HINBAN = cls.HINBAN                '品番
    data.REVNUM = cls.REVNUM                '改訂番号
    data.FACT = cls.FACT                    '工場
    data.OPCOND = cls.OPCOND                '操業条件
    data.LENGTH = cls.LENGTH                '長さ
    data.USECLASS = cls.USECLASS            '使用区分
    'data.REGDATE = cls.REGDATE              '登録日付
    'data.UPDDATE = cls.UPDDATE              '更新日付
    'data.SENDFLAG = cls.SENDFLAG            '送信フラグ
    'data.SENDDATE = cls.SENDDATE            '送信日付

    c2u_TBCME039 = data
End Function


'------ テーブル名:TBCME040    ---- ブロック管理

' ユーザ定義型をクラスに変換する(結晶情報)
Public Function u2c_cmzc001b(data As typ_TBCME040) As c_cmzc001b
Dim cls As New c_cmzc001b       '変換先クラス

    cls.CRYNUM = data.CRYNUM                '結晶番号
    cls.IngotPos = data.IngotPos            '結晶内開始位置
    cls.LENGTH = data.LENGTH                '長さ
    cls.BLOCKID = data.BLOCKID              'ブロックID
    cls.KRPROCCD = data.KRPROCCD            '現在管理工程
    cls.NOWPROC = data.NOWPROC              '現在工程
    cls.LPKRPROCCD = data.LPKRPROCCD        '最終通過管理工程
    cls.LASTPASS = data.LASTPASS            '最終通過工程
    cls.DELCLS = data.DELCLS                '削除区分
    cls.LSTATCLS = data.LSTATCLS            '最終状態区分
    cls.RSTATCLS = data.RSTATCLS            '流動状態区分
    cls.HOLDCLS = data.HOLDCLS              'ホールド区分
    cls.BDCAUS = data.BDCAUS                '不良理由
    'cls.REGDATE = data.REGDATE              '登録日付
    'cls.UPDDATE = data.UPDDATE              '更新日付
    'cls.SUMMITSENDFLAG = data.SUMMITSENDFLAG 'SUMMIT送信フラグ
    'cls.SENDFLAG = data.SENDFLAG            '送信フラグ
    'cls.SENDDATE = data.SENDDATE            '送信日付

    Set u2c_cmzc001b = cls
End Function


' クラスをユーザ定義型に変換する(ブロック管理)
Public Function c2u_TBCME040(cls As c_cmzc001b) As typ_TBCME040
Dim data As typ_TBCME040        '変換先ユーザ定義型

    data.CRYNUM = cls.CRYNUM                '結晶番号
    data.IngotPos = cls.IngotPos            '結晶内開始位置
    data.LENGTH = cls.LENGTH                '長さ
    data.BLOCKID = cls.BLOCKID              'ブロックID
    data.KRPROCCD = cls.KRPROCCD            '現在管理工程
    data.NOWPROC = cls.NOWPROC              '現在工程
    data.LPKRPROCCD = cls.LPKRPROCCD        '最終通過管理工程
    data.LASTPASS = cls.LASTPASS            '最終通過工程
    data.DELCLS = cls.DELCLS                '削除区分
    data.LSTATCLS = cls.LSTATCLS            '最終状態区分
    data.RSTATCLS = cls.RSTATCLS            '流動状態区分
    data.HOLDCLS = cls.HOLDCLS              'ホールド区分
    data.BDCAUS = cls.BDCAUS                '不良理由
    'data.REGDATE = cls.REGDATE              '登録日付
    'data.UPDDATE = cls.UPDDATE              '更新日付
    'data.SUMMITSENDFLAG = cls.SUMMITSENDFLAG 'SUMMIT送信フラグ
    'data.SENDFLAG = cls.SENDFLAG            '送信フラグ
    'data.SENDDATE = cls.SENDDATE            '送信日付

    c2u_TBCME040 = data
End Function


'------ テーブル名:TBCME041    ---- 品番管理

' ユーザ定義型をクラスに変換する(結晶情報)
Public Function u2c_cmzc001d(data As typ_TBCME041) As c_cmzc001d
Dim cls As New c_cmzc001d       '変換先クラス

    cls.CRYNUM = data.CRYNUM                '結晶番号
    cls.IngotPos = data.IngotPos            '結晶内開始位置
    cls.HINBAN = data.HINBAN                '品番
    cls.REVNUM = data.REVNUM                '製品番号改訂番号
    cls.factory = data.factory              '工場
    cls.opecond = data.opecond              '操業条件
    cls.LENGTH = data.LENGTH                '長さ
    'cls.REGDATE = data.REGDATE              '登録日付
    'cls.UPDDATE = data.UPDDATE              '更新日付
    'cls.SENDFLAG = data.SENDFLAG            '送信フラグ
    'cls.SENDDATE = data.SENDDATE            '送信日付

    Set u2c_cmzc001d = cls
End Function


' クラスをユーザ定義型に変換する(品番管理)
Public Function c2u_TBCME041(cls As c_cmzc001d) As typ_TBCME041
Dim data As typ_TBCME041        '変換先ユーザ定義型

    data.CRYNUM = cls.CRYNUM                '結晶番号
    data.IngotPos = cls.IngotPos            '結晶内開始位置
    data.HINBAN = cls.HINBAN                '品番
    data.REVNUM = cls.REVNUM                '製品番号改訂番号
    data.factory = cls.factory              '工場
    data.opecond = cls.opecond              '操業条件
    data.LENGTH = cls.LENGTH                '長さ
    'data.REGDATE = cls.REGDATE              '登録日付
    'data.UPDDATE = cls.UPDDATE              '更新日付
    'data.SENDFLAG = cls.SENDFLAG            '送信フラグ
    'data.SENDDATE = cls.SENDDATE            '送信日付

    c2u_TBCME041 = data
End Function


'------ テーブル名:TBCME042    ---- SXL管理

' ユーザ定義型をクラスに変換する(結晶情報)
Public Function u2c_TBCME042(data As typ_TBCME042) As c_TBCME042
Dim cls As New c_TBCME042       '変換先クラス

    cls.CRYNUM = data.CRYNUM                '結晶番号
    cls.IngotPos = data.IngotPos            '結晶内開始位置
    cls.LENGTH = data.LENGTH                '長さ
    cls.SXLID = data.SXLID                  'SXLID
    cls.KRPROCCD = data.KRPROCCD            '管理工程
    cls.NOWPROC = data.NOWPROC              '現在工程
    cls.LPKRPROCCD = data.LPKRPROCCD        '最終通過管理工程
    cls.LASTPASS = data.LASTPASS            '最終通過工程
    cls.DELCLS = data.DELCLS                '削除区分
    cls.LSTATCLS = data.LSTATCLS            '最終状態区分
    cls.HOLDCLS = data.HOLDCLS              'ホールド区分
    cls.HINBAN = data.HINBAN                '品番
    cls.REVNUM = data.REVNUM                '製品番号改訂番号
    cls.factory = data.factory              '工場
    cls.opecond = data.opecond              '操業条件
    cls.Count = data.Count                  '枚数
    'cls.REGDATE = data.REGDATE              '登録日付
    'cls.UPDDATE = data.UPDDATE              '更新日付
    'cls.SUMMITSENDFLAG = data.SUMMITSENDFLAG 'SUMMIT送信フラグ
    'cls.SENDFLAG = data.SENDFLAG            '送信フラグ
    'cls.SENDDATE = data.SENDDATE            '送信日付

    Set u2c_TBCME042 = cls
End Function


' クラスをユーザ定義型に変換する(SXL管理)
Public Function c2u_TBCME042(cls As c_TBCME042) As typ_TBCME042
Dim data As typ_TBCME042        '変換先ユーザ定義型

    data.CRYNUM = cls.CRYNUM                '結晶番号
    data.IngotPos = cls.IngotPos            '結晶内開始位置
    data.LENGTH = cls.LENGTH                '長さ
    data.SXLID = cls.SXLID                  'SXLID
    data.KRPROCCD = cls.KRPROCCD            '管理工程
    data.NOWPROC = cls.NOWPROC              '現在工程
    data.LPKRPROCCD = cls.LPKRPROCCD        '最終通過管理工程
    data.LASTPASS = cls.LASTPASS            '最終通過工程
    data.DELCLS = cls.DELCLS                '削除区分
    data.LSTATCLS = cls.LSTATCLS            '最終状態区分
    data.HOLDCLS = cls.HOLDCLS              'ホールド区分
    data.HINBAN = cls.HINBAN                '品番
    data.REVNUM = cls.REVNUM                '製品番号改訂番号
    data.factory = cls.factory              '工場
    data.opecond = cls.opecond              '操業条件
    data.Count = cls.Count                  '枚数
    'data.REGDATE = cls.REGDATE              '登録日付
    'data.UPDDATE = cls.UPDDATE              '更新日付
    'data.SUMMITSENDFLAG = cls.SUMMITSENDFLAG 'SUMMIT送信フラグ
    'data.SENDFLAG = cls.SENDFLAG            '送信フラグ
    'data.SENDDATE = cls.SENDDATE            '送信日付

    c2u_TBCME042 = data
End Function


'------ テーブル名:XSDCS    ---- 新サンプル管理（ブロック）

' ユーザ定義型をクラスに変換する(結晶情報)
Public Function u2c_cmzc001e(data As typ_XSDCS) As c_cmzc001e
'Public Function u2c_cmzc001e(data As typ_TBCME043) As c_cmzc001e
Dim cls As New c_cmzc001e       '変換先クラス
    
    cls.CRYNUMCS = data.CRYNUMCS                'ブロックID
    cls.SMPKBNCS = data.SMPKBNCS                'サンプル区分
    cls.TBKBNCS = data.TBKBNCS                  'T/B区分
    cls.REPSMPLIDCS = data.REPSMPLIDCS          'サンプルNo
    cls.XTALCS = data.XTALCS                    '結晶番号
    cls.INPOSCS = data.INPOSCS                  '結晶内位置
    cls.HINBCS = data.HINBCS                    '品番
    cls.REVNUMCS = data.REVNUMCS                '製品番号改訂番号
    cls.FACTORYCS = data.FACTORYCS              '工場
    cls.OPECS = data.OPECS                      '操業条件
    cls.KTKBNCS = data.KTKBNCS                  '確定区分
    cls.SMPLUMU = data.SMPLUMU                  'サンプル有無区分
    cls.BLKKTFLAGCS = data.BLKKTFLAGCS          'ブロック確定フラグ
    cls.CRYSMPLIDRSCS = data.CRYSMPLIDRSCS      'サンプルID
    cls.CRYSMPLIDRS1CS = data.CRYSMPLIDRS1CS    '推定サンプルID1
    cls.CRYSMPLIDRS2CS = data.CRYSMPLIDRS2CS    '推定サンプルID2
    cls.CRYINDRSCS = data.CRYINDRSCS            '状態FLG(Rs)
    cls.CRYRESRS1CS = data.CRYRESRS1CS          '実績FLG1(Rs)
    cls.CRYRESRS2CS = data.CRYRESRS2CS          '実績FLG2(Rs)
    cls.CRYSMPLIDOICS = data.CRYSMPLIDOICS      'サンプルID(Oi)
    cls.CRYINDOICS = data.CRYINDOICS            '状態FLG(Oi)
    cls.CRYRESOICS = data.CRYRESOICS            '実績FLG(Oi)
    cls.CRYSMPLIDB1CS = data.CRYSMPLIDB1CS      'サンプルID(B1)
    cls.CRYINDB1CS = data.CRYINDB1CS            '状態FLG(B1)
    cls.CRYRESB1CS = data.CRYRESB1CS            '実績FLG(B1)
    cls.CRYSMPLIDB2CS = data.CRYSMPLIDB2CS      'サンプルID(B2)
    cls.CRYINDB2CS = data.CRYINDB2CS            '状態FLG(B2)
    cls.CRYRESB2CS = data.CRYRESB2CS            '実績FLG(B2)
    cls.CRYSMPLIDB3CS = data.CRYSMPLIDB3CS      'サンプルID(B3)
    cls.CRYINDB3CS = data.CRYINDB3CS            '状態FLG(B3)
    cls.CRYRESB3CS = data.CRYRESB3CS            '実績FLG(B3)
    cls.CRYSMPLIDL1CS = data.CRYSMPLIDL1CS      'サンプルID(L1)
    cls.CRYINDL1CS = data.CRYINDL1CS            '状態FLG(L1)
    cls.CRYRESL1CS = data.CRYRESL1CS            '実績FLG(L1)
    cls.CRYSMPLIDL2CS = data.CRYSMPLIDL2CS      'サンプルID(L2)
    cls.CRYINDL2CS = data.CRYINDL2CS            '状態FLG(L2)
    cls.CRYRESL2CS = data.CRYRESL2CS            '実績FLG(L2)
    cls.CRYSMPLIDL3CS = data.CRYSMPLIDL3CS      'サンプルID(L3)
    cls.CRYINDL3CS = data.CRYINDL3CS            '状態FLG(L3)
    cls.CRYRESL3CS = data.CRYRESL3CS            '実績FLG(L3)
    cls.CRYSMPLIDL4CS = data.CRYSMPLIDL4CS      'サンプルID(L4)
    cls.CRYINDL4CS = data.CRYINDL4CS            '状態FLG(L4)
    cls.CRYRESL4CS = data.CRYRESL4CS            '実績FLG(L4)
    cls.CRYSMPLIDCSCS = data.CRYSMPLIDCSCS      'サンプルID(Cs)
    cls.CRYINDCSCS = data.CRYINDCSCS            '状態FLG(Cs)
    cls.CRYRESCSCS = data.CRYRESCSCS            '実績FLG(Cs)
    cls.CRYSMPLIDGDCS = data.CRYSMPLIDGDCS      'サンプルID(GD)
    cls.CRYINDGDCS = data.CRYINDGDCS            '状態FLG(GD)
    cls.CRYRESGDCS = data.CRYRESGDCS            '実績FLG(GD)
    cls.CRYSMPLIDTCS = data.CRYSMPLIDTCS        'サンプルID(T)
    cls.CRYINDTCS = data.CRYINDTCS              '状態FLG(T)
    cls.CRYRESTCS = data.CRYRESTCS              '実績FLG(T)
    cls.CRYSMPLIDEPCS = data.CRYSMPLIDEPCS      'サンプルID(EPD)
    cls.CRYINDEPCS = data.CRYINDEPCS            '状態FLG(EPD)
    cls.CRYRESEPCS = data.CRYRESEPCS            '実績FLG(EPD)
    cls.SMPLNUMCS = data.SMPLNUMCS              'サンプル枚数
    cls.SMPLPATCS = data.SMPLPATCS              'サンプルパターン
    'cls.REGDATE = data.REGDATE              '登録日付
    'cls.UPDDATE = data.UPDDATE              '更新日付
    'cls.SENDFLAG = data.SENDFLAG            '送信フラグ
    'cls.SENDDATE = data.SENDDATE            '送信日付

    Set u2c_cmzc001e = cls
End Function


' クラスをユーザ定義型に変換する(結晶サンプル管理)
'Public Function c2u_TBCME043(cls As c_cmzc001e) As typ_TBCME043
'Dim data As typ_TBCME043        '変換先ユーザ定義型
'
'    data.CRYNUM = cls.CRYNUM                '結晶番号
'    data.IngotPos = cls.IngotPos            '結晶内位置
'    data.SMPKBN = cls.SMPKBN                'サンプル区分
'    data.SMPLNO = cls.SMPLNO                'サンプルNo
'    data.HINBAN = cls.HINBAN                '品番
'    data.REVNUM = cls.REVNUM                '製品番号改訂番号
'    data.factory = cls.factory              '工場
'    data.opecond = cls.opecond              '操業条件
'    data.SMPLUMU = cls.SMPLUMU              'サンプル有無区分
'    data.CRYINDRS = cls.CRYINDRS            '結晶検査指示（Rs)
'    data.CRYINDOI = cls.CRYINDOI            '結晶検査指示（Oi)
'    data.CRYINDB1 = cls.CRYINDB1            '結晶検査指示（B1)
'    data.CRYINDB2 = cls.CRYINDB2            '結晶検査指示（B2）
'    data.CRYINDB3 = cls.CRYINDB3            '結晶検査指示（B3)
'    data.CRYINDL1 = cls.CRYINDL1            '結晶検査指示（L1)
'    data.CRYINDL2 = cls.CRYINDL2            '結晶検査指示（L2)
'    data.CRYINDL3 = cls.CRYINDL3            '結晶検査指示（L3)
'    data.CRYINDL4 = cls.CRYINDL4            '結晶検査指示（L4)
'    data.CRYINDCS = cls.CRYINDCS            '結晶検査指示（Cs)
'    data.CRYINDGD = cls.CRYINDGD            '結晶検査指示（GD)
'    data.CRYINDT = cls.CRYINDT              '結晶検査指示（T)
'    data.CRYINDEP = cls.CRYINDEP            '結晶検査指示（EPD)
'    data.CRYRESRS = cls.CRYRESRS            '結晶検査実績（Rs)
'    data.CRYRESOI = cls.CRYRESOI            '結晶検査実績（Oi)
'    data.CRYRESB1 = cls.CRYRESB1            '結晶検査実績（B1)
'    data.CRYRESB2 = cls.CRYRESB2            '結晶検査実績（B2）
'    data.CRYRESB3 = cls.CRYRESB3            '結晶検査実績（B3)
'    data.CRYRESL1 = cls.CRYRESL1            '結晶検査実績（L1)
'    data.CRYRESL2 = cls.CRYRESL2            '結晶検査実績（L2)
'    data.CRYRESL3 = cls.CRYRESL3            '結晶検査実績（L3)
'    data.CRYRESL4 = cls.CRYRESL4            '結晶検査実績（L4)
'    data.CRYRESCS = cls.CRYRESCS            '結晶検査実績（Cs)
'    data.CRYRESGD = cls.CRYRESGD            '結晶検査実績（GD)
'    data.CRYREST = cls.CRYREST              '結晶検査実績（T)
'    data.CRYRESEP = cls.CRYRESEP            '結晶検査実績（EPD)
'    data.SMPLNUM = cls.SMPLNUM              'サンプル枚数
'    data.SMPLPAT = cls.SMPLPAT              'サンプルパターン
'    'data.REGDATE = cls.REGDATE              '登録日付
'    'data.UPDDATE = cls.UPDDATE              '更新日付
'    'data.SENDFLAG = cls.SENDFLAG            '送信フラグ
'    'data.SENDDATE = cls.SENDDATE            '送信日付
'
'    c2u_TBCME043 = data
'End Function

Public Function c2u_XSDCS(cls As c_cmzc001e) As typ_XSDCS
Dim data As typ_XSDCS        '変換先ユーザ定義型

    data.CRYNUMCS = cls.CRYNUMCS                'ブロックID
    data.SMPKBNCS = cls.SMPKBNCS                'サンプル区分
    data.TBKBNCS = cls.TBKBNCS                  'T/B区分
    data.REPSMPLIDCS = cls.REPSMPLIDCS          'サンプルNo
    data.XTALCS = cls.XTALCS                    '結晶番号
    data.INPOSCS = cls.INPOSCS                  '結晶内位置
    data.HINBCS = cls.HINBCS                    '品番
    data.REVNUMCS = cls.REVNUMCS                '製品番号改訂番号
    data.FACTORYCS = cls.FACTORYCS              '工場
    data.OPECS = cls.OPECS                      '操業条件
    data.KTKBNCS = cls.KTKBNCS                  '確定区分
    data.SMPLUMU = cls.SMPLUMU                  'サンプル有無区分
    data.BLKKTFLAGCS = cls.BLKKTFLAGCS          'ブロック確定フラグ
    data.CRYSMPLIDRSCS = cls.CRYSMPLIDRSCS      'サンプルID
    data.CRYSMPLIDRS1CS = cls.CRYSMPLIDRS1CS    '推定サンプルID1
    data.CRYSMPLIDRS2CS = cls.CRYSMPLIDRS2CS    '推定サンプルID2
    data.CRYINDRSCS = cls.CRYINDRSCS            '状態FLG(Rs)
    data.CRYRESRS1CS = cls.CRYRESRS1CS          '実績FLG1(Rs)
    data.CRYRESRS2CS = cls.CRYRESRS2CS          '実績FLG2(Rs)
    data.CRYSMPLIDOICS = cls.CRYSMPLIDOICS      'サンプルID(Oi)
    data.CRYINDOICS = cls.CRYINDOICS            '状態FLG(Oi)
    data.CRYRESOICS = cls.CRYRESOICS            '実績FLG(Oi)
    data.CRYSMPLIDB1CS = cls.CRYSMPLIDB1CS      'サンプルID(B1)
    data.CRYINDB1CS = cls.CRYINDB1CS            '状態FLG(B1)
    data.CRYRESB1CS = cls.CRYRESB1CS            '実績FLG(B1)
    data.CRYSMPLIDB2CS = cls.CRYSMPLIDB2CS      'サンプルID(B2)
    data.CRYINDB2CS = cls.CRYINDB2CS            '状態FLG(B2)
    data.CRYRESB2CS = cls.CRYRESB2CS            '実績FLG(B2)
    data.CRYSMPLIDB3CS = cls.CRYSMPLIDB3CS      'サンプルID(B3)
    data.CRYINDB3CS = cls.CRYINDB3CS            '状態FLG(B3)
    data.CRYRESB3CS = cls.CRYRESB3CS            '実績FLG(B3)
    data.CRYSMPLIDL1CS = cls.CRYSMPLIDL1CS      'サンプルID(L1)
    data.CRYINDL1CS = cls.CRYINDL1CS            '状態FLG(L1)
    data.CRYRESL1CS = cls.CRYRESL1CS            '実績FLG(L1)
    data.CRYSMPLIDL2CS = cls.CRYSMPLIDL2CS      'サンプルID(L2)
    data.CRYINDL2CS = cls.CRYINDL2CS            '状態FLG(L2)
    data.CRYRESL2CS = cls.CRYRESL2CS            '実績FLG(L2)
    data.CRYSMPLIDL3CS = cls.CRYSMPLIDL3CS      'サンプルID(L3)
    data.CRYINDL3CS = cls.CRYINDL3CS            '状態FLG(L3)
    data.CRYRESL3CS = cls.CRYRESL3CS            '実績FLG(L3)
    data.CRYSMPLIDL4CS = cls.CRYSMPLIDL4CS      'サンプルID(L4)
    data.CRYINDL4CS = cls.CRYINDL4CS            '状態FLG(L4)
    data.CRYRESL4CS = cls.CRYRESL4CS            '実績FLG(L4)
    data.CRYSMPLIDCSCS = cls.CRYSMPLIDCSCS      'サンプルID(Cs)
    data.CRYINDCSCS = cls.CRYINDCSCS            '状態FLG(Cs)
    data.CRYRESCSCS = cls.CRYRESCSCS            '実績FLG(Cs)
    data.CRYSMPLIDGDCS = cls.CRYSMPLIDGDCS      'サンプルID(GD)
    data.CRYINDGDCS = cls.CRYINDGDCS            '状態FLG(GD)
    data.CRYRESGDCS = cls.CRYRESGDCS            '実績FLG(GD)
    data.CRYSMPLIDTCS = cls.CRYSMPLIDTCS        'サンプルID(T)
    data.CRYINDTCS = cls.CRYINDTCS              '状態FLG(T)
    data.CRYRESTCS = cls.CRYRESTCS              '実績FLG(T)
    data.CRYSMPLIDEPCS = cls.CRYSMPLIDEPCS      'サンプルID(EPD)
    data.CRYINDEPCS = cls.CRYINDEPCS            '状態FLG(EPD)
    data.CRYRESEPCS = cls.CRYRESEPCS            '実績FLG(EPD)
    data.SMPLNUMCS = cls.SMPLNUMCS              'サンプル枚数
    data.SMPLPATCS = cls.SMPLPATCS              'サンプルパターン
    'data.REGDATE = cls.REGDATE                 '登録日付
    'data.UPDDATE = cls.UPDDATE                 '更新日付
    'data.SENDFLAG = cls.SENDFLAG               '送信フラグ
    'data.SENDDATE = cls.SENDDATE               '送信日付

    c2u_XSDCS = data
End Function

''2003/09/02 SystemBrain サンプル管理変更
''------ テーブル名:TBCME044    ---- WFサンプル管理
'
'' ユーザ定義型をクラスに変換する(結晶情報)
'Public Function u2c_cmzc001f(data As typ_TBCME044) As c_cmzc001f
'Dim cls As New c_cmzc001f       '変換先クラス
'
'    cls.CRYNUM = data.CRYNUM                '結晶番号
'    cls.INGOTPOS = data.INGOTPOS            '結晶内位置
'    cls.SMPKBN = data.SMPKBN                'サンプル区分
'    cls.SMPLID = data.SMPLID                'サンプルID
'    cls.hinban = data.hinban                '品番
'    cls.REVNUM = data.REVNUM                '製品番号改訂番号
'    cls.FACTORY = data.FACTORY              '工場
'    cls.OPECOND = data.OPECOND              '操業条件
'    cls.SMPLUMU = data.SMPLUMU              'サンプル有無区分
'    cls.WFINDRS = data.WFINDRS              'WF検査指示（Rs)
'    cls.WFINDOI = data.WFINDOI              'WF検査指示（Oi)
'    cls.WFINDB1 = data.WFINDB1              'WF検査指示（B1)
'    cls.WFINDB2 = data.WFINDB2              'WF検査指示（B2）
'    cls.WFINDB3 = data.WFINDB3              'WF検査指示（B3)
'    cls.WFINDL1 = data.WFINDL1              'WF検査指示（L1)
'    cls.WFINDL2 = data.WFINDL2              'WF検査指示（L2)
'    cls.WFINDL3 = data.WFINDL3              'WF検査指示（L3)
'    cls.WFINDL4 = data.WFINDL4              'WF検査指示（L4)
'    cls.WFINDDS = data.WFINDDS              'WF検査指示（DS)
'    cls.WFINDDZ = data.WFINDDZ              'WF検査指示（DZ)
'    cls.WFINDSP = data.WFINDSP              'WF検査指示（SP)
'    cls.WFINDDO1 = data.WFINDDO1            'WF検査指示（DO1)
'    cls.WFINDDO2 = data.WFINDDO2            'WF検査指示（DO2)
'    cls.WFINDDO3 = data.WFINDDO3            'WF検査指示（DO3)
'    cls.WFRESRS = data.WFRESRS              'WF検査実績（Rs)
'    cls.WFRESOI = data.WFRESOI              'WF検査実績（Oi)
'    cls.WFRESB1 = data.WFRESB1              'WF検査実績（B1)
'    cls.WFRESB2 = data.WFRESB2              'WF検査実績（B2）
'    cls.WFRESB3 = data.WFRESB3              'WF検査実績（B3)
'    cls.WFRESL1 = data.WFRESL1              'WF検査実績（L1)
'    cls.WFRESL2 = data.WFRESL2              'WF検査実績（L2)
'    cls.WFRESL3 = data.WFRESL3              'WF検査実績（L3)
'    cls.WFRESL4 = data.WFRESL4              'WF検査実績（L4)
'    cls.WFRESDS = data.WFRESDS              'WF検査実績（DS)
'    cls.WFRESDZ = data.WFRESDZ              'WF検査実績（DZ)
'    cls.WFRESSP = data.WFRESSP              'WF検査実績（SP)
'    cls.WFRESDO1 = data.WFRESDO1            'WF検査実績（DO1)
'    cls.WFRESDO2 = data.WFRESDO2            'WF検査実績（DO2)
'    cls.WFRESDO3 = data.WFRESDO3            'WF検査実績（DO3)
'    'cls.REGDATE = data.REGDATE              '登録日付
'    'cls.UPDDATE = data.UPDDATE              '更新日付
'    'cls.SENDFLAG = data.SENDFLAG            '送信フラグ
'    'cls.SENDDATE = data.SENDDATE            '送信日付
'
'    Set u2c_cmzc001f = cls
'End Function

'------ テーブル名:XSDCW    ---- 新サンプル管理(SXL)

' ユーザ定義型をクラスに変換する(結晶情報)
Public Function u2c_cmzc001f(data As typ_XSDCW) As c_cmzc001f
Dim cls As New c_cmzc001f       '変換先クラス

    cls.SXLIDCW = data.SXLIDCW              'SXLID
    cls.SMPKBNCW = data.SMPKBNCW            'サンプル区分
    cls.TBKBNCW = data.TBKBNCW              'T/B区分
    cls.REPSMPLIDCW = data.REPSMPLIDCW      '代表サンプルID
    cls.XTALCW = data.XTALCW                '結晶番号
    cls.INPOSCW = data.INPOSCW              '結晶内位置
    cls.HINBCW = data.HINBCW                '品番
    cls.REVNUMCW = data.REVNUMCW            '製品番号改訂番号
    cls.FACTORYCW = data.FACTORYCW          '工場
    cls.OPECW = data.OPECW                  '操業番号
    cls.KTKBNCW = data.KTKBNCW              '確定区分
    cls.SMPLUMU = data.SMPLUMU              'サンプル有無区分
    cls.SMCRYNUMCW = data.SMCRYNUMCW        'サンプルブロックID
    cls.WFSMPLIDRSCW = data.WFSMPLIDRSCW    'サンプルID(Rs)
    cls.WFSMPLIDRS1CW = data.WFSMPLIDRS1CW  '推定サンプルID1(Rs)
    cls.WFSMPLIDRS2CW = data.WFSMPLIDRS2CW  '推定サンプルID2(Rs)
    cls.WFINDRSCW = data.WFINDRSCW          '状態FLG(Rs)
    cls.WFRESRS1CW = data.WFRESRS1CW        '実績FLG1(Rs)
    cls.WFRESRS2CW = data.WFRESRS2CW        '実績FLG2(Rs)
    cls.WFSMPLIDOICW = data.WFSMPLIDOICW    'サンプルID(Oi)
    cls.WFINDOICW = data.WFINDOICW          '状態FLG(Oi)
    cls.WFRESOICW = data.WFRESOICW          '実績FLG(Oi)
    cls.WFSMPLIDB1CW = data.WFSMPLIDB1CW    'サンプルID(B1)
    cls.WFINDB1CW = data.WFINDB1CW          '状態FLG(B1)
    cls.WFRESB1CW = data.WFRESB1CW          '実績FLG(B1)
    cls.WFSMPLIDB2CW = data.WFSMPLIDB2CW    'サンプルID(B2)
    cls.WFINDB2CW = data.WFINDB2CW          '状態FLG(B2)
    cls.WFRESB2CW = data.WFRESB2CW          '実績FLG(B2)
    cls.WFSMPLIDB3CW = data.WFSMPLIDB3CW    'サンプルID(B3)
    cls.WFINDB3CW = data.WFINDB3CW          '状態FLG(B3)
    cls.WFRESB3CW = data.WFRESB3CW          '実績FLG(B3)
    cls.WFSMPLIDL1CW = data.WFSMPLIDL1CW    'サンプルID(L1)
    cls.WFINDL1CW = data.WFINDL1CW          '状態FLG(L1)
    cls.WFRESL1CW = data.WFRESL1CW          '実績FLG(L1)
    cls.WFSMPLIDL2CW = data.WFSMPLIDL2CW    'サンプルID(L2)
    cls.WFINDL2CW = data.WFINDL2CW          '状態FLG(L2)
    cls.WFRESL2CW = data.WFRESL2CW          '実績FLG(L2)
    cls.WFSMPLIDL3CW = data.WFSMPLIDL3CW    'サンプルID(L3)
    cls.WFINDL3CW = data.WFINDL3CW          '状態FLG(L3)
    cls.WFRESL3CW = data.WFRESL3CW          '実績FLG(L3)
    cls.WFSMPLIDL4CW = data.WFSMPLIDL4CW    'サンプルID(L4)
    cls.WFINDL4CW = data.WFINDL4CW          '状態FLG(L4)
    cls.WFRESL4CW = data.WFRESL4CW          '実績FLG(L4)
    cls.WFSMPLIDDSCW = data.WFSMPLIDDSCW    'サンプルID(DS)
    cls.WFINDDSCW = data.WFINDDSCW          '状態FLG(DS)
    cls.WFRESDSCW = data.WFRESDSCW          '実績FLG(DS)
    cls.WFSMPLIDDZCW = data.WFSMPLIDDZCW    'サンプルID(DZ)
    cls.WFINDDZCW = data.WFINDDZCW          '状態FLG(DZ)
    cls.WFRESDZCW = data.WFRESDZCW          '実績FLG(DZ)
    cls.WFSMPLIDSPCW = data.WFSMPLIDSPCW    'サンプルID(SP)
    cls.WFINDSPCW = data.WFINDSPCW          '状態FLG(SP)
    cls.WFRESSPCW = data.WFRESSPCW          '実績FLG(SP)
    cls.WFSMPLIDDO1CW = data.WFSMPLIDDO1CW  'サンプルID(DO1)
    cls.WFINDDO1CW = data.WFINDDO1CW        '状態FLG(DO1)
    cls.WFRESDO1CW = data.WFRESDO1CW        '実績FLG(DO1)
    cls.WFSMPLIDDO2CW = data.WFSMPLIDDO2CW  'サンプルID(DO2)
    cls.WFINDDO2CW = data.WFINDDO2CW        '状態FLG(DO2)
    cls.WFRESDO2CW = data.WFRESDO2CW        '実績FLG(DO2)
    cls.WFSMPLIDDO3CW = data.WFSMPLIDDO3CW  'サンプルID(DO3)
    cls.WFINDDO3CW = data.WFINDDO3CW        '状態FLG(DO3)
    cls.WFRESDO3CW = data.WFRESDO3CW        '実績FLG(DO3)

    'cls.REGDATE = data.REGDATE              '登録日付
    'cls.UPDDATE = data.UPDDATE              '更新日付
    'cls.SENDFLAG = data.SENDFLAG            '送信フラグ
    'cls.SENDDATE = data.SENDDATE            '送信日付

    Set u2c_cmzc001f = cls
End Function

''2003/09/02 SystemBrain サンプル管理変更
' クラスをユーザ定義型に変換する(WFサンプル管理)
'Public Function c2u_TBCME044(cls As c_cmzc001f) As typ_TBCME044
'Dim data As typ_TBCME044        '変換先ユーザ定義型
'
'    data.CRYNUM = cls.CRYNUM                '結晶番号
'    data.INGOTPOS = cls.INGOTPOS            '結晶内位置
'    data.SMPKBN = cls.SMPKBN                'サンプル区分
'    data.SMPLID = cls.SMPLID                'サンプルID
'    data.hinban = cls.hinban                '品番
'    data.REVNUM = cls.REVNUM                '製品番号改訂番号
'    data.FACTORY = cls.FACTORY              '工場
'    data.OPECOND = cls.OPECOND              '操業条件
'    data.SMPLUMU = cls.SMPLUMU              'サンプル有無区分
'    data.WFINDRS = cls.WFINDRS              'WF検査指示（Rs)
'    data.WFINDOI = cls.WFINDOI              'WF検査指示（Oi)
'    data.WFINDB1 = cls.WFINDB1              'WF検査指示（B1)
'    data.WFINDB2 = cls.WFINDB2              'WF検査指示（B2）
'    data.WFINDB3 = cls.WFINDB3              'WF検査指示（B3)
'    data.WFINDL1 = cls.WFINDL1              'WF検査指示（L1)
'    data.WFINDL2 = cls.WFINDL2              'WF検査指示（L2)
'    data.WFINDL3 = cls.WFINDL3              'WF検査指示（L3)
'    data.WFINDL4 = cls.WFINDL4              'WF検査指示（L4)
'    data.WFINDDS = cls.WFINDDS              'WF検査指示（DS)
'    data.WFINDDZ = cls.WFINDDZ              'WF検査指示（DZ)
'    data.WFINDSP = cls.WFINDSP              'WF検査指示（SP)
'    data.WFINDDO1 = cls.WFINDDO1            'WF検査指示（DO1)
'    data.WFINDDO2 = cls.WFINDDO2            'WF検査指示（DO2)
'    data.WFINDDO3 = cls.WFINDDO3            'WF検査指示（DO3)
'    data.WFRESRS = cls.WFRESRS              'WF検査実績（Rs)
'    data.WFRESOI = cls.WFRESOI              'WF検査実績（Oi)
'    data.WFRESB1 = cls.WFRESB1              'WF検査実績（B1)
'    data.WFRESB2 = cls.WFRESB2              'WF検査実績（B2）
'    data.WFRESB3 = cls.WFRESB3              'WF検査実績（B3)
'    data.WFRESL1 = cls.WFRESL1              'WF検査実績（L1)
'    data.WFRESL2 = cls.WFRESL2              'WF検査実績（L2)
'    data.WFRESL3 = cls.WFRESL3              'WF検査実績（L3)
'    data.WFRESL4 = cls.WFRESL4              'WF検査実績（L4)
'    data.WFRESDS = cls.WFRESDS              'WF検査実績（DS)
'    data.WFRESDZ = cls.WFRESDZ              'WF検査実績（DZ)
'    data.WFRESSP = cls.WFRESSP              'WF検査実績（SP)
'    data.WFRESDO1 = cls.WFRESDO1            'WF検査実績（DO1)
'    data.WFRESDO2 = cls.WFRESDO2            'WF検査実績（DO2)
'    data.WFRESDO3 = cls.WFRESDO3            'WF検査実績（DO3)
'    'data.REGDATE = cls.REGDATE              '登録日付
'    'data.UPDDATE = cls.UPDDATE              '更新日付
'    'data.SENDFLAG = cls.SENDFLAG            '送信フラグ
'    'data.SENDDATE = cls.SENDDATE            '送信日付
'
'    c2u_TBCME044 = data
'End Function
Public Function c2u_XSDCW(cls As c_cmzc001f) As typ_XSDCW
Dim data As typ_XSDCW        '変換先ユーザ定義型

    data.SXLIDCW = cls.SXLIDCW              'SXLID
    data.SMPKBNCW = cls.SMPKBNCW            'サンプル区分
    data.TBKBNCW = cls.TBKBNCW              'T/B区分
    data.REPSMPLIDCW = cls.REPSMPLIDCW      '代表サンプルID
    data.XTALCW = cls.XTALCW                '結晶番号
    data.INPOSCW = cls.INPOSCW              '結晶内位置
    data.HINBCW = cls.HINBCW                '品番
    data.REVNUMCW = cls.REVNUMCW            '製品番号改訂番号
    data.FACTORYCW = cls.FACTORYCW          '工場
    data.OPECW = cls.OPECW                  '操業番号
    data.KTKBNCW = cls.KTKBNCW              '確定区分
    data.SMPLUMU = cls.SMPLUMU              'サンプル有無区分
    data.SMCRYNUMCW = cls.SMCRYNUMCW        'サンプルブロックID
    data.WFSMPLIDRSCW = cls.WFSMPLIDRSCW    'サンプルID(Rs)
    data.WFSMPLIDRS1CW = cls.WFSMPLIDRS1CW  '推定サンプルID1(Rs)
    data.WFSMPLIDRS2CW = cls.WFSMPLIDRS2CW  '推定サンプルID2(Rs)
    data.WFINDRSCW = cls.WFINDRSCW          '状態FLG(Rs)
    data.WFRESRS1CW = cls.WFRESRS1CW        '実績FLG1(Rs)
    data.WFRESRS2CW = cls.WFRESRS2CW        '実績FLG2(Rs)
    data.WFSMPLIDOICW = cls.WFSMPLIDOICW    'サンプルID(Oi)
    data.WFINDOICW = cls.WFINDOICW          '状態FLG(Oi)
    data.WFRESOICW = cls.WFRESOICW          '実績FLG(Oi)
    data.WFSMPLIDB1CW = cls.WFSMPLIDB1CW    'サンプルID(B1)
    data.WFINDB1CW = cls.WFINDB1CW          '状態FLG(B1)
    data.WFRESB1CW = cls.WFRESB1CW          '実績FLG(B1)
    data.WFSMPLIDB2CW = cls.WFSMPLIDB2CW    'サンプルID(B2)
    data.WFINDB2CW = cls.WFINDB2CW          '状態FLG(B2)
    data.WFRESB2CW = cls.WFRESB2CW          '実績FLG(B2)
    data.WFSMPLIDB3CW = cls.WFSMPLIDB3CW    'サンプルID(B3)
    data.WFINDB3CW = cls.WFINDB3CW          '状態FLG(B3)
    data.WFRESB3CW = cls.WFRESB3CW          '実績FLG(B3)
    data.WFSMPLIDL1CW = cls.WFSMPLIDL1CW    'サンプルID(L1)
    data.WFINDL1CW = cls.WFINDL1CW          '状態FLG(L1)
    data.WFRESL1CW = cls.WFRESL1CW          '実績FLG(L1)
    data.WFSMPLIDL2CW = cls.WFSMPLIDL2CW    'サンプルID(L2)
    data.WFINDL2CW = cls.WFINDL2CW          '状態FLG(L2)
    data.WFRESL2CW = cls.WFRESL2CW          '実績FLG(L2)
    data.WFSMPLIDL3CW = cls.WFSMPLIDL3CW    'サンプルID(L3)
    data.WFINDL3CW = cls.WFINDL3CW          '状態FLG(L3)
    data.WFRESL3CW = cls.WFRESL3CW          '実績FLG(L3)
    data.WFSMPLIDL4CW = cls.WFSMPLIDL4CW    'サンプルID(L4)
    data.WFINDL4CW = cls.WFINDL4CW          '状態FLG(L4)
    data.WFRESL4CW = cls.WFRESL4CW          '実績FLG(L4)
    data.WFSMPLIDDSCW = cls.WFSMPLIDDSCW    'サンプルID(DS)
    data.WFINDDSCW = cls.WFINDDSCW          '状態FLG(DS)
    data.WFRESDSCW = cls.WFRESDSCW          '実績FLG(DS)
    data.WFSMPLIDDZCW = cls.WFSMPLIDDZCW    'サンプルID(DZ)
    data.WFINDDZCW = cls.WFINDDZCW          '状態FLG(DZ)
    data.WFRESDZCW = cls.WFRESDZCW          '実績FLG(DZ)
    data.WFSMPLIDSPCW = cls.WFSMPLIDSPCW    'サンプルID(SP)
    data.WFINDSPCW = cls.WFINDSPCW          '状態FLG(SP)
    data.WFRESSPCW = cls.WFRESSPCW          '実績FLG(SP)
    data.WFSMPLIDDO1CW = cls.WFSMPLIDDO1CW  'サンプルID(DO1)
    data.WFINDDO1CW = cls.WFINDDO1CW        '状態FLG(DO1)
    data.WFRESDO1CW = cls.WFRESDO1CW        '実績FLG(DO1)
    data.WFSMPLIDDO2CW = cls.WFSMPLIDDO2CW  'サンプルID(DO2)
    data.WFINDDO2CW = cls.WFINDDO2CW        '状態FLG(DO2)
    data.WFRESDO2CW = cls.WFRESDO2CW        '実績FLG(DO2)
    data.WFSMPLIDDO3CW = cls.WFSMPLIDDO3CW  'サンプルID(DO3)
    data.WFINDDO3CW = cls.WFINDDO3CW        '状態FLG(DO3)
    data.WFRESDO3CW = cls.WFRESDO3CW        '実績FLG(DO3)

    'data.REGDATE = cls.REGDATE              '登録日付
    'data.UPDDATE = cls.UPDDATE              '更新日付
    'data.SENDFLAG = cls.SENDFLAG            '送信フラグ
    'data.SENDDATE = cls.SENDDATE            '送信日付

    c2u_XSDCW = data
End Function


'------ テーブル名:TBCME045    ---- 切断指示

' ユーザ定義型をクラスに変換する(結晶情報)
Public Function u2c_cmzc001c(data As typ_TBCME045) As c_cmzc001c
Dim cls As New c_cmzc001c       '変換先クラス

    cls.CRYNUM = data.CRYNUM                '結晶番号
    cls.IngotPos = data.IngotPos            '結晶内開始位置
    cls.TRANCNT = data.TRANCNT              '処理回数
    cls.LENGTH = data.LENGTH                '長さ
    cls.PROCCODE = data.PROCCODE            '工程コード
    cls.StaffID = data.StaffID              '社員ID
    cls.HINBAN = data.HINBAN                '上品番
    cls.REVNUM = data.REVNUM                '上品番製品番号改訂番号
    cls.factory = data.factory              '上品番工場
    cls.opecond = data.opecond              '上品番操業条件
    cls.BDCAUS = data.BDCAUS                '区分コード
    cls.STATCLS = data.STATCLS              '状態区分
    cls.BLOCKID = data.BLOCKID              'ブロックID
    'cls.REGDATE = data.REGDATE              '登録日付
    'cls.UPDDATE = data.UPDDATE              '更新日付
    'cls.SENDFLAG = data.SENDFLAG            '送信フラグ
    'cls.SENDDATE = data.SENDDATE            '送信日付

    Set u2c_cmzc001c = cls
End Function


' クラスをユーザ定義型に変換する(切断指示)
Public Function c2u_TBCME045(cls As c_cmzc001c) As typ_TBCME045
Dim data As typ_TBCME045        '変換先ユーザ定義型

    data.CRYNUM = cls.CRYNUM                '結晶番号
    data.IngotPos = cls.IngotPos            '結晶内開始位置
    data.TRANCNT = cls.TRANCNT              '処理回数
    data.LENGTH = cls.LENGTH                '長さ
    data.PROCCODE = cls.PROCCODE            '工程コード
    data.StaffID = cls.StaffID              '社員ID
    data.HINBAN = cls.HINBAN                '上品番
    data.REVNUM = cls.REVNUM                '上品番製品番号改訂番号
    data.factory = cls.factory              '上品番工場
    data.opecond = cls.opecond              '上品番操業条件
    data.BDCAUS = cls.BDCAUS                '区分コード
    data.STATCLS = cls.STATCLS              '状態区分
    data.BLOCKID = cls.BLOCKID              'ブロックID
    'data.REGDATE = cls.REGDATE              '登録日付
    'data.UPDDATE = cls.UPDDATE              '更新日付
    'data.SENDFLAG = cls.SENDFLAG            '送信フラグ
    'data.SENDDATE = cls.SENDDATE            '送信日付

    c2u_TBCME045 = data
End Function
