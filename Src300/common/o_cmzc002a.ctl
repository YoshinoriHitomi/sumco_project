VERSION 5.00
Begin VB.UserControl o_cmzc002a 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  '実線
   CanGetFocus     =   0   'False
   ClientHeight    =   5220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   ScaleHeight     =   5220
   ScaleWidth      =   3855
End
Attribute VB_Name = "o_cmzc002a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'                                     2001/05/17
'======================================================
' 結晶図コントロール
' 概要    : 与えられた結晶クラスの内容を図示する
' 参照    : 結晶クラス      (c_cmzcXl.ctl)
'         : 結晶レコード保持クラス      (c_cmzcBlk.cls～c_cmzc001g.cls)
'         : 結晶レコードコレクション    (c_cmzc001h.cls～c_cmzc001m.cls)
'======================================================

'内部使用の定数
Const SMP_WIDTH = 80                'サンプルマークの幅
Const SMP_HEIGHT = 60               'サンプルマークの高さ
Const MARGIN_CENTER = 160           'サンプルマーク用の空き幅

'内部変数
Dim m_Xl As c_cmzcXl                '描画情報となる結晶クラス
Dim pxXL_Left As Long               '結晶図左端
Dim pxXL_Center As Long             '結晶図中心
Dim pxXL_Right As Long              '結晶図右端
Dim pyXL_Top As Long                '結晶図Top位置
Dim pyXL_Zero As Long               '結晶図直胴Top端位置
Dim pyXL_Bot As Long                '結晶図直胴Tail端位置
Dim pyXL_Tail As Long               '結晶図Tail位置


'警告! 以下のｺﾒﾝﾄ行を変更または削除しないでください !
'MemberInfo=7
'概要      :Clearメソッド
'説明      :内部データと表示を初期化する
'履歴      :2001/05/17 作成  野村
Public Function Clear() As Integer
Attribute Clear.VB_Description = "結晶図を初期化します"

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "o_cmzc002a.ctl -- Function Clear"

    '' 内部変数を初期化する
    Set m_Xl = New c_cmzcXl
    
    '' 初期表示に関する既定値を設定する
    m_Xl.TOPLENG = 200
    m_Xl.BODYLENG = 1500
    m_Xl.BOTLENG = 400
    
    '' 初期状態で描画する
    UserControl_Resize

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Function

'警告! 以下のｺﾒﾝﾄ行を変更または削除しないでください !
'MemberInfo=7
'概要      :Drawメソッド
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型         ,説明
'          :Xl            ,I  ,c_cmzcXl ,描画対象の結晶クラスオブジェクト
'          :戻り値        ,O  ,Integer    ,
'説明      :与えられた結晶クラスオブジェクトの内容を元に描画する
'履歴      :2001/05/17 作成  野村
Public Function Draw(Xl As c_cmzcXl) As Integer
Attribute Draw.VB_Description = "結晶クラスの情報で、結晶図を描画します"
Dim pos As Integer
Dim Cut As c_cmzcCut    '切断指示
Dim blk As c_cmzcBlk    'ブロック
Dim HIN As c_cmzcHin    '品番
Dim SXL As c_cmzcSxl    'SXL
Dim n As Integer
Dim wk As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "o_cmzc002a.ctl -- Function Draw"

    '' 描画対象となる結晶クラスの内容を複写する
    Set m_Xl = Xl.Clone
    
    '' 描画のための調整を行う
    With m_Xl
        ''引上げ長を超えるブロックの長さを調整する
        pos = .Blks.LowerArea(CStr(.BODYLENG + 1))
        If pos < 9999 Then
            Set blk = .Blks(CStr(pos))
            blk.LENGTH = .BODYLENG - blk.INGOTPOS
        End If
    
        ''切断指示を含むブロック開始位置に、切断指示を追加する
        For Each Cut In .Cuts
            pos = .Blks.LowerArea(Cut.INGOTPOS)
            If (0 < pos) And (pos < Cut.INGOTPOS) Then
                If Not .Cuts.Exist(pos) Then
                    .AddCut pos, Cut.INGOTPOS - pos
                End If
            End If
        Next
        
        ''SXLを品番区切りとして設定する
        For Each SXL In .Sxls
            pos = .Hins.LowerArea(SXL.INGOTPOS)
            If pos <> SXL.INGOTPOS Then
                '品番区切り位置でなかったら、そこを区切りとしてSXLの品番を設定する
                .Hins.AddLine SXL.INGOTPOS
            End If
            pos = .Hins.LowerArea(SXL.INGOTPOS + SXL.LENGTH)
            If pos <> SXL.INGOTPOS + SXL.LENGTH Then
                '品番区切り位置でなかったら、そこを区切りとしてSXLの品番を設定する
                .Hins.AddLine SXL.INGOTPOS + SXL.LENGTH
            End If
            'SXLの品番を設定する
            With .Hins(CStr(SXL.INGOTPOS))
                .hinban = SXL.hinban
                .REVNUM = SXL.REVNUM
                .factory = SXL.factory
                .opecond = SXL.opecond
            End With
        Next
    
        ''引上げ長を超える描画内容を調整する
        .BlkPlans.LimitByIngotPos .BODYLENG
        .Blks.LimitByIngotPos .BODYLENG
        .HinPlans.LimitByIngotPos .BODYLENG
        .Hins.LimitByIngotPos .BODYLENG
        
        ''無切断ブロックを削除する
        If .Blks.COUNT Then
            n = .Blks.COUNT
            If (.Blks(n).INGOTPOS + .Blks(n).LENGTH = .BODYLENG) Then
                If Mid$(.Blks(n).BLOCKID, 10, 3) = "0$2" Then
                    .Blks.Remove n, False
                End If
            End If
            If .Blks.COUNT Then
                If (.Blks(1).INGOTPOS = 0) Then
                    If Mid$(.Blks(1).BLOCKID, 10, 3) = "0$1" Then
                        .Blks.Remove 1, False
                    End If
                End If
            End If
        End If
        
        If .Blks.COUNT Then
            .Hins.LimitByIngotArea .Blks(1).INGOTPOS, .Blks(.Blks.COUNT).INGOTPOS + .Blks(.Blks.COUNT).LENGTH
        End If
    End With
    
    '' 描画する
    UserControl_Resize

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Function

'概要      :コントロール Initialize時処理
'説明      :
'履歴      :2001/05/17 作成  野村
Private Sub UserControl_Initialize()

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "o_cmzc002a.ctl -- Sub UserControl_Initialize"

    '' 内部変数を初期化する
    Clear

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Sub

'概要      :描画処理
'説明      :内部データに基づいて描画を行う
'履歴      :2001/05/17 作成  野村
Private Sub UserControl_Paint()
    Dim pos As Integer
    Dim pos1 As Integer
    Dim pos2 As Integer
    Dim px As Long
    Dim py As Long
    Dim py2 As Long
    Dim pBefore As Long
    Dim s As String
    Dim smpWidth As Long
    Dim smpRight As Long
    Dim margin As Long
    Dim blk As c_cmzcBlk        '描画対象の Blk
    Dim Cut As c_cmzcCut        '描画対象の Cut
    Dim HIN As c_cmzcHin        '描画対象の Hin
    Dim SXL As c_cmzcSxl        '描画対象の Sxl
    Dim XlSmp As c_cmzcXlSmp    '描画対象の XlSmp
    Dim WFSMP As c_cmzcWfSmp    '描画対象の WfSmp
    Dim Rej As c_cmzcRej        '描画対象の rej
    Dim drawTarget As Boolean   'それを描画するか

'   Debug.Print "Ctl:Paint"

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "o_cmzc002a.ctl -- Sub UserControl_Paint"

    Cls     '最初に消す
    
    With m_Xl
        '' 結晶図の外枠を描画する
        DrawStyle = vbSolid
        Line (pxXL_Center, pyXL_Top)-(pxXL_Left, pyXL_Zero), vbBlack
        Line (pxXL_Center, pyXL_Top)-(pxXL_Right, pyXL_Zero), vbBlack
        Line (pxXL_Left, pyXL_Bot)-(pxXL_Center, pyXL_Tail), vbBlack
        Line (pxXL_Right, pyXL_Bot)-(pxXL_Center, pyXL_Tail), vbBlack
        
        '' 現ブロックの背景色を変える(現ブロック指定がない場合は全域白)
        If .CurrentBlock <> vbNullString Then
            py = pyXL_Zero
            For Each blk In .Blks
                py = GetY(blk.INGOTPOS)
                py2 = GetY(blk.INGOTPOS + blk.LENGTH)
                
                If (Right$(blk.BLOCKID, 3) = Right$(.CurrentBlock, 3)) Then
                    Line (pxXL_Left, py)-(pxXL_Right, py2), vbWhite, BF
                Else
                    Line (pxXL_Left, py)-(pxXL_Right, py2), BackColor, BF
                End If
            Next
        Else
            Line (pxXL_Left, pyXL_Zero)-(pxXL_Right, pyXL_Bot), vbWhite, BF
        End If
        
        '' 追加ドープ位置を描画する
        If .ADDDPPOS Then
            py = GetY(.ADDDPPOS)
            DrawStyle = vbSolid
            'px = Width - (Width - pxXL_Right) / 2
            px = pxXL_Right + TextWidth("0000")
            Line (pxXL_Center, py)-(px, py), vbMagenta
            s = "追加Dope"
            FontSize = 8
            CurrentY = py - TextHeight(s) / 2
            CurrentX = px + 20
            ForeColor = vbMagenta
            Print s;
            ForeColor = vbBlack
        End If
        
        '' 欠落位置を描画する
        For Each Rej In .Rejs
            If m_Xl.GetIngotPos(Rej.LOTID, Rej.LENFROM, pos) = FUNCTION_RETURN_SUCCESS Then
                py = GetY(pos)
                If m_Xl.GetIngotPos(Rej.LOTID, Rej.LENTO, pos) = FUNCTION_RETURN_SUCCESS Then
                    pBefore = GetY(pos)
                    
                    DrawStyle = vbSolid
                    FillStyle = vbDiagonalCross
                    FillColor = vbGreen
                    Line (pxXL_Left, py)-(pxXL_Right, pBefore), vbGreen, B
                    FillStyle = vbFSSolid
                    ForeColor = vbBlack
                End If
            End If
        Next
    
        '' 結晶図の両脇の縦線を描画する
        Line (pxXL_Left, pyXL_Zero)-(pxXL_Left, pyXL_Bot), vbBlack
        Line (pxXL_Right, pyXL_Zero)-(pxXL_Right, pyXL_Bot), vbBlack
    
        '' サンプルがあるときとないときで、切断位置・品番区切位置の長さを変える
        If .WfSmps.COUNT Then
            margin = MARGIN_CENTER
        Else
            margin = 0
        End If
    
        '' ブロックを描画する
        py = GetY(0)
        For Each blk In .Blks
            pos1 = blk.INGOTPOS                 'ブロック上端
            pos2 = blk.INGOTPOS + blk.LENGTH    'ブロック下端
            py = GetY(pos1)
            py2 = GetY(pos2)
            
            ''ブロック開始位置の描画
            Line (pxXL_Left, py)-(pxXL_Center - margin, py), vbBlack
            s = Str(pos1)
            FontSize = 8
            CurrentY = py - TextHeight(s) / 2
            CurrentX = pxXL_Left - TextWidth(s) - 20
            Print s;
            
            ''ブロック終了位置の描画
            Line (pxXL_Left, py2)-(pxXL_Center - margin, py2), vbBlack
            s = Str(pos2)
            FontSize = 8
            CurrentY = py2 - TextHeight(s) / 2
            CurrentX = pxXL_Left - TextWidth(s) - 20
            Print s;
            
            ''ブロックIDの描画
            drawTarget = True
            s = blk.BLOCKID
            If Len(s) = 0 Then              '' ブロックIDが空なら、描画対象外
                drawTarget = False
            ElseIf pos2 - pos1 <= 1 Then
                drawTarget = False
            ElseIf .Cuts.ExistInArea(blk.INGOTPOS + 1, blk.LENGTH - 1) Then '' ブロックTop/Botを除いた間に切断指示を含んでいたら描画対象外
                drawTarget = False
            Else
                s = Right$(s, 3)
                If Mid$(s, 2, 1) = "$" Then '' 「$」を含むブロックIDは、描画対象外
                    drawTarget = False
                End If
            End If
            If drawTarget Then
                FontSize = 9
                CurrentY = (py + py2 - TextHeight(s)) / 2
                CurrentX = (pxXL_Left + pxXL_Center - TextWidth(s)) / 2
                Select Case pos2 - pos1
                    Case Is < 100
                        ForeColor = vbRed
                    Case Is > 400
                        ForeColor = vbRed
                    Case Else
                        ForeColor = vbBlack
                End Select
                Print s;
                ForeColor = vbBlack
            End If
            
            ''工程コードの描画
            'drawTarget = True
            s = blk.NOWPROC
            If (s = vbNullString) Then              '' 工程コード未登録の場合は描画対象外
                drawTarget = False
            ElseIf blk.DELCLS = "1" Then
                Select Case blk.LSTATCLS
                  Case "R"
                    s = "ﾘﾒﾙﾄ"
                  Case "H"
                    s = "ﾊｲｷ"
                  Case "W"
                    s = "WF出荷"
                  Case "B"
                    s = "BAR出荷"
                  Case "V"
                    s = "外販"
                End Select
            ElseIf blk.HOLDCLS = "1" Then
                s = "ﾎｰﾙﾄﾞ"
            Else
                s = blk.NOWPROC
                If (s = "CB320") Then               '' クリスタルカタログのときは<ｶﾀﾛｸﾞ>と描画
                    s = "ｶﾀﾛｸﾞ"
                'ElseIf (Left$(s, 2) = "CB") Then    '' その他原料系の時は描画対象外(リメルト・廃棄)
                '    drawTarget = False
                End If
            End If
            If (drawTarget) Then
                s = "<" & s & ">"
                FontSize = 9
                CurrentY = (py + py2 - TextHeight(s)) / 2
                CurrentX = (pxXL_Left - TextWidth(s & "    ")) / 2
                ForeColor = vbBlack
                Print s;
                ForeColor = vbBlack
            End If
        Next
        
        '' 切断指示を描画する
        py = GetY(0)
        For Each Cut In .Cuts
            pos1 = Cut.INGOTPOS                 'ブロック上端
            pos2 = Cut.INGOTPOS + Cut.LENGTH    'ブロック下端
            py = GetY(pos1)
            py2 = GetY(pos2)
            
            '' 切断指示自体の描画
            DrawStyle = vbDot
            Line (pxXL_Left, py)-(pxXL_Center - margin, py), vbBlack
            s = CStr(Cut.INGOTPOS)
            FontSize = 8
            CurrentY = py - TextHeight(s) / 2
            CurrentX = pxXL_Left - TextWidth(s) - 20
            ForeColor = vbBlue
            Print s;
            ForeColor = vbBlack
                        
            If Cut.LENGTH > 1 Then
                '' ブロックIDの描画
                s = Right$(Cut.BLOCKID, 3)
                If Mid$(s, 2, 1) <> "$" Then
                    FontSize = 9
                    CurrentY = (py + py2 - TextHeight(s)) / 2
                    CurrentX = (pxXL_Left + pxXL_Center - TextWidth(s)) / 2
                    Select Case pos2 - pos1
                        Case Is < 100
                            ForeColor = vbRed
                        Case Is > 400
                            ForeColor = vbRed
                        Case Else
                            ForeColor = vbBlack
                    End Select
                    Print s;
                    ForeColor = vbBlack
                End If
            End If
        Next
        
        '' 品番区切位置を描画する
        py = GetY(0)
        For Each HIN In .Hins
            py = GetY(HIN.INGOTPOS)
            py2 = GetY(HIN.INGOTPOS + HIN.LENGTH)
            
            '' 品番開始位置の描画
            DrawStyle = vbDot
            Line (pxXL_Center + margin, py)-(pxXL_Right, py), vbBlack
            s = " " & CStr(HIN.INGOTPOS)
            FontSize = 8
            CurrentY = py - TextHeight(s) / 2
            CurrentX = pxXL_Right + 20
            Print s;
            
            '' 品番終了位置の描画
            DrawStyle = vbDot
            Line (pxXL_Center + margin, py2)-(pxXL_Right, py2), vbBlack
            s = " " & CStr(HIN.INGOTPOS + HIN.LENGTH)
            FontSize = 8
            CurrentY = py2 - TextHeight(s) / 2
            CurrentX = pxXL_Right + 20
            Print s;
            
            '' 品番の描画
            s = Trim$(HIN.hinban)
            FontSize = 8
            CurrentY = (py + py2 - TextHeight(s)) / 2
            CurrentX = (pxXL_Right + pxXL_Center - TextWidth(s)) / 2
            ForeColor = vbBlack
            Print s;
            ForeColor = vbBlack
        Next
    
        '' WFサンプル位置を描画する
        For Each WFSMP In .WfSmps
            py = GetY(WFSMP.INGOTPOS)
            
            DrawStyle = vbSolid
            Line (pxXL_Center - SMP_WIDTH, py)-(pxXL_Center, py - SMP_HEIGHT), vbBlack
            Line (pxXL_Center, py - SMP_HEIGHT)-(pxXL_Center + SMP_WIDTH, py), vbBlack
            Line (pxXL_Center + SMP_WIDTH, py)-(pxXL_Center, py + SMP_HEIGHT), vbBlack
            Line (pxXL_Center, py + SMP_HEIGHT)-(pxXL_Center - SMP_WIDTH, py), vbBlack
        Next
    End With

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Sub

'概要      :Resize時処理
'説明      :図の基本サイズを計算し、描画する
'履歴      :2001/05/17 作成  野村
Private Sub UserControl_Resize()
    Dim totalLen As Long
    Dim totalHeight As Long
    Dim zoom As Double

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "o_cmzc002a.ctl -- Sub UserControl_Resize"
    
    ''コントロールの大きさに合わせ、図の基本サイズを計算する
    With m_Xl
        If (.TOPLENG + .BODYLENG + .BOTLENG = 0) Then GoTo proc_exit
        
        pxXL_Left = Width / 4                       ''結晶図左端位置
        pxXL_Right = Width - Width / 4              ''結晶図右端位置
        pxXL_Center = (pxXL_Left + pxXL_Right) / 2  ''結晶図中心位置
        pyXL_Top = 200                              ''結晶図Top位置
        pyXL_Tail = Height - 200                    ''結晶図Tail位置
        
        totalLen = .TOPLENG + .BODYLENG + .BOTLENG
        totalHeight = pyXL_Tail - pyXL_Top
        If totalLen = 0 Then
            zoom = totalHeight
        Else
            zoom = totalHeight * 1# / totalLen
        End If
        pyXL_Zero = pyXL_Top + .TOPLENG * zoom      ''結晶図直胴Top端位置
        pyXL_Bot = pyXL_Tail - .BOTLENG * zoom      ''結晶図直胴Tail端位置
        
        'Debug.Print pyXL_Top, pyXL_Zero, pyXL_Bot, pyXL_Tail, zoom
    End With
    UserControl_Paint

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Sub

'概要      :Terminate時処理
'説明      :
'履歴      :2001/05/17 作成  野村
Private Sub UserControl_Terminate()

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "o_cmzc002a.ctl -- Sub UserControl_Terminate"

    '' 内部クラスを解放する
    Set m_Xl = Nothing

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Sub




'概要      :インゴット内位置に対応するコントロール内座標(Y)を得る
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,インゴット内位置
'          :戻り値        ,O  ,Long      ,Y座標
'説明      :
'履歴      :2001/05/17 作成  野村
Private Function GetY(pos As Integer) As Long
    Dim bodyHeight As Long

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "o_cmzc002a.ctl -- Function GetY"
    
    ''対応するコントロール内座標(Y)を計算する
    bodyHeight = pyXL_Bot - pyXL_Zero
    If m_Xl.BODYLENG = 0 Then
        GetY = pyXL_Zero + bodyHeight
    Else
        GetY = pyXL_Zero + bodyHeight * (pos * 1# / m_Xl.BODYLENG)
    End If

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Function


'概要      :より小さい値を選択する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :value1        ,I  ,Long      ,値1
'          :value2        ,I  ,Long      ,値2
'          :戻り値        ,O  ,Long      ,小さい値
'説明      :
'履歴      :2001/05/17 作成  野村
Private Function LowerValue(value1 As Long, value2 As Long) As Long

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "o_cmzc002a.ctl -- Function LowerValue"

    ''与えられた値の内、小さい方を返す
    If value1 < value2 Then
        LowerValue = value1
    Else
        LowerValue = value2
    End If

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Function


'概要      :より大きい値を選択する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :value1        ,I  ,Long      ,値1
'          :value2        ,I  ,Long      ,値2
'          :戻り値        ,O  ,Long      ,大きい値
'説明      :
'履歴      :2001/05/17 作成  野村
Private Function HigherValue(value1 As Long, value2 As Long) As Long

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "o_cmzc002a.ctl -- Function HigherValue"

    If value1 > value2 Then
        HigherValue = value1
    Else
        HigherValue = value2
    End If

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Function
