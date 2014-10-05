Attribute VB_Name = "s_PngOcxAddon"
'///////////////////////////////////////////////////
' @(S)
'       LEADTOOLS 14.0J API関数
'
' @(h)  s_pngOcxAddon.bas ver 1.0 ( 2007.09.11 宮武 佳子 )
'
'///////////////////////////////////////////////////

'===============================================================================
'   構造体定義
'===============================================================================
' LEAD BitmapHandle構造体 画像の操作および変換を行う領域として使用
Type LEADBitmapHandle
 Reserved1 As Long
 Reserved2 As Long
 
 uStructSize As Long
 Width As Long
 Height As Long
 BitsPerPixel As Long
 BytesPerLine As Long
 CurrentPos As Long
 CurrentRow As Long

 Flags As Long

 nColors As Long

 hPalette As Long
 
 ViewPerspective As Long
 Order As Long
 Size As Long
 pDitherData As Long
 DitheringMethod As Long
 pRgnInfo As Long             ' 内部使用
 XResolution As Long
 YResolution As Long
 left As Long                 ' アニメーション再生用の左端オフセット
 top As Long                  ' ビューの視点と方向がTOP_LEFTに基づく、アニメーション再生用の上端オフセット
 Delay As Long                ' アニメーション再生用の遅延（ミリ秒単位）
 Background As Long           ' アニメーション再生用の背景色(COLORREF)
 Transparency As Long         ' アニメーション再生用の透過色 (COLORREF)
 DisposalMethod As Long       ' アニメーション再生用の破棄モード
 LowBit As Long               ' ウィンドウレベル調整に使用する低位ビットを示す値
 HighBit As Long              ' ウィンドウレベル調整に使用する高位ビットを示す値
 pLUT As Long                 ' LUT set by L_WindowLevel
 LUTLength As Long            ' no of entries pointed by LUT
 DataOffset As Long           ' Internal use
 MinVal As Long               ' グレースケール画像に含まれるピクセルの明度の最小レベル
 MaxVal As Long               ' グレースケール画像含まれるピクセルの明度の最大レベル
 PaintLowBit As Long          ' low bit set by L_WindowLevelBitmap
 PaintHighBit As Long         ' high bit set by L_WindowLevelBitmap
 PaintGamma As Long
 PaintContrast As Long
 PaintIntensity As Long
 pPaintLUT As Long
 pOverlaysInfo As Long
 Reserved As Long
End Type

' 拡張ファイルオプション構造体 画像ファイルのロードに関する拡張オプションを設定
Type LOADFILEOPTION
    uStructSize As Long
    XResolution As Long
    YResolution As Long
    Flags As Long
    Passes As Long
    PageNumber As Long
    GlobalLoop As Long
    IFD As Long
    uRedScan As Long
    uGreenScan As Long
    uBlueScan As Long
End Type

' FileInfo構造体 指定されたファイルの情報を提供
Type FileInfo
 uStructSize As Long       'この構造体のサイズ（バイト単位）
 Format As Long            'ファイル形式
 Name As String * 16       'NULLで終了するDOSのファイル名を含む16キャラクタの配列。WIN32では、名前が15キャラクタを超える場合、超過部分が切り詰められます。
 Width As Long             '画像の幅（ピクセル単位）
 Height As Long            '画像の高さ（ピクセル単位）
 BitsPerPixel As Long      '１ピクセルあたりのビット数
 SizeDisk As Long          'ディスク上のファイルサイズ（バイト単位）
 SizeMem As Long           'ファイルのロードに必要なメモリ量（バイト単位）
 Compression As String * 20 'NULLで終了する圧縮形式を説明する文字列を含む20キャラクタの配列
 ViewPerspective As Long   '画像のはじまりが保存されている位置
 Order As Long             'ビットマップのカラーオーダー。有効な値は、ORDER_RGB、ORDER_BGRおよびORDER_GRAY
 PageNumber As Long        '複数ページのファイルのページ番号
 TotalPages As Long        'ファイル内にあるページの総数
 XResolution As Long       '水平解像度（DPI：インチあたりのドット数）
 YResolution As Long       '垂直解像度（DPI）
 uFlags As Long            'ファイルについての説明
 GlobalLoop As Long        'グローバルなアニメーション・ループカウント
 GlobalWidth As Long       'アニメーション再生のためのグローバルな幅
 GlobalHeight As Long      'アニメーション再生のためのグローバルな高さ
 GlobalBackground As Long  'アニメーション再生のためのグローバルな背景色
 GlobalPalette(255) As Long 'アニメーション再生のためのグローバルなパレット
 IFD As Long               'ページが開始される、TIFFファイル内のオフセット
 Layers As Long            'ファイルに含まれるレイヤの数。この値が０の場合、ファイルにはレイヤが含まれていません。レイヤを持つことができるのはPSDファイルのみです。
End Type

' SAVEFILEOPTION構造体 画像ファイルへの書き込みに関する拡張オプションを設定
Type SAVEFILEOPTION
 uStructSize As Long
 Reserved1 As Long
 Reserved2 As Long
 Flags As Long
 Passes As Long
 PageNumber As Long
 GlobalWidth As Long
 GlobalHeight As Long
 GlobalLoop As Long
 GlobalBackground As Long
 GlobalPalette(255) As Long
 StampWidth As Long
 StampHeight As Long
 StampBits As Long
 Constrained As Long
 FieldPic As Long
 FrameRate As Long
 FramesGOP As Long
 PelAspectRatio As Long
 ProfileID As Long
 LevelID As Long
 ChromaFormat As Long
 VideoFormat As Long
 IntraDcPrec As Long
 BitRate As Long
 UserInfo As String * 255
 IFD As Long
 szPassword As String * 255
 PhotometricInterpretation As Long
 TileWidth As Long
 TileHeight As Long
 Flags2 As Long
End Type


'===============================================================================
'   定数定義
'===============================================================================
' ビットマップ用のカラーオーダー
Global Const ORDER_RGB = 0                    ' メモリ内で緑、青、赤の順
Global Const ORDER_BGR = 1                    ' メモリ内で青、緑、赤の順
'Global Const ORDER_SOURCE = -1
Global Const ORDER_GRAY = 2                   ' 12ビットまたは16ビットのグレースケール画像
' 出力ファイル形式
Global Const FILE_PNG = 75
' 構造体サイズ
Global Const SIZEOFLEADBITMAPHANDLE = 164
Global Const SIZEOFSAVEFILEOPTION = 1650


'===============================================================================
'   DLL関数
'===============================================================================
'' 画像ファイルをロードしてビットマップを作成
Declare Function L_LoadBitmap Lib "LTFIL14N" (ByVal pFile$, hBitmap As LEADBitmapHandle, ByVal uStructSize&, ByVal nBits&, ByVal nOrder&, pLoadOptions As LOADFILEOPTION, pFileInfo As FileInfo) As Long
'' ファイルを保存するときに使用する拡張オプションのデフォルト値を取得
Declare Function L_GetDefaultSaveFileOption Lib "LTFIL14N" (ByRef pOption As SAVEFILEOPTION, ByVal uStructSize&) As Long
'' ビットマップに含まれる画像をファイル（圧縮または非圧縮）に保存
Declare Function L_SaveBitmap Lib "LTFIL14N" (ByVal pFile$, hBitmap As LEADBitmapHandle, ByVal nFormat&, ByVal nBits&, ByVal nQFactor&, pSaveOptions As SAVEFILEOPTION) As Long
