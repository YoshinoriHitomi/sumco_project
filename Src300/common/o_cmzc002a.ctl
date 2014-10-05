VERSION 5.00
Begin VB.UserControl o_cmzc002a 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  '����
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
' �����}�R���g���[��
' �T�v    : �^����ꂽ�����N���X�̓��e��}������
' �Q��    : �����N���X      (c_cmzcXl.ctl)
'         : �������R�[�h�ێ��N���X      (c_cmzcBlk.cls�`c_cmzc001g.cls)
'         : �������R�[�h�R���N�V����    (c_cmzc001h.cls�`c_cmzc001m.cls)
'======================================================

'�����g�p�̒萔
Const SMP_WIDTH = 80                '�T���v���}�[�N�̕�
Const SMP_HEIGHT = 60               '�T���v���}�[�N�̍���
Const MARGIN_CENTER = 160           '�T���v���}�[�N�p�̋󂫕�

'�����ϐ�
Dim m_Xl As c_cmzcXl                '�`����ƂȂ錋���N���X
Dim pxXL_Left As Long               '�����}���[
Dim pxXL_Center As Long             '�����}���S
Dim pxXL_Right As Long              '�����}�E�[
Dim pyXL_Top As Long                '�����}Top�ʒu
Dim pyXL_Zero As Long               '�����}����Top�[�ʒu
Dim pyXL_Bot As Long                '�����}����Tail�[�ʒu
Dim pyXL_Tail As Long               '�����}Tail�ʒu


'�x��! �ȉ��̺��čs��ύX�܂��͍폜���Ȃ��ł������� !
'MemberInfo=7
'�T�v      :Clear���\�b�h
'����      :�����f�[�^�ƕ\��������������
'����      :2001/05/17 �쐬  �쑺
Public Function Clear() As Integer
Attribute Clear.VB_Description = "�����}�����������܂�"

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "o_cmzc002a.ctl -- Function Clear"

    '' �����ϐ�������������
    Set m_Xl = New c_cmzcXl
    
    '' �����\���Ɋւ������l��ݒ肷��
    m_Xl.TOPLENG = 200
    m_Xl.BODYLENG = 1500
    m_Xl.BOTLENG = 400
    
    '' ������Ԃŕ`�悷��
    UserControl_Resize

proc_exit:
    '�I��
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Function

'�x��! �ȉ��̺��čs��ύX�܂��͍폜���Ȃ��ł������� !
'MemberInfo=7
'�T�v      :Draw���\�b�h
'���Ұ�    :�ϐ���        ,IO ,�^         ,����
'          :Xl            ,I  ,c_cmzcXl ,�`��Ώۂ̌����N���X�I�u�W�F�N�g
'          :�߂�l        ,O  ,Integer    ,
'����      :�^����ꂽ�����N���X�I�u�W�F�N�g�̓��e�����ɕ`�悷��
'����      :2001/05/17 �쐬  �쑺
Public Function Draw(Xl As c_cmzcXl) As Integer
Attribute Draw.VB_Description = "�����N���X�̏��ŁA�����}��`�悵�܂�"
Dim pos As Integer
Dim Cut As c_cmzcCut    '�ؒf�w��
Dim blk As c_cmzcBlk    '�u���b�N
Dim HIN As c_cmzcHin    '�i��
Dim SXL As c_cmzcSxl    'SXL
Dim n As Integer
Dim wk As Integer

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "o_cmzc002a.ctl -- Function Draw"

    '' �`��ΏۂƂȂ錋���N���X�̓��e�𕡎ʂ���
    Set m_Xl = Xl.Clone
    
    '' �`��̂��߂̒������s��
    With m_Xl
        ''���グ���𒴂���u���b�N�̒����𒲐�����
        pos = .Blks.LowerArea(CStr(.BODYLENG + 1))
        If pos < 9999 Then
            Set blk = .Blks(CStr(pos))
            blk.LENGTH = .BODYLENG - blk.INGOTPOS
        End If
    
        ''�ؒf�w�����܂ރu���b�N�J�n�ʒu�ɁA�ؒf�w����ǉ�����
        For Each Cut In .Cuts
            pos = .Blks.LowerArea(Cut.INGOTPOS)
            If (0 < pos) And (pos < Cut.INGOTPOS) Then
                If Not .Cuts.Exist(pos) Then
                    .AddCut pos, Cut.INGOTPOS - pos
                End If
            End If
        Next
        
        ''SXL��i�ԋ�؂�Ƃ��Đݒ肷��
        For Each SXL In .Sxls
            pos = .Hins.LowerArea(SXL.INGOTPOS)
            If pos <> SXL.INGOTPOS Then
                '�i�ԋ�؂�ʒu�łȂ�������A��������؂�Ƃ���SXL�̕i�Ԃ�ݒ肷��
                .Hins.AddLine SXL.INGOTPOS
            End If
            pos = .Hins.LowerArea(SXL.INGOTPOS + SXL.LENGTH)
            If pos <> SXL.INGOTPOS + SXL.LENGTH Then
                '�i�ԋ�؂�ʒu�łȂ�������A��������؂�Ƃ���SXL�̕i�Ԃ�ݒ肷��
                .Hins.AddLine SXL.INGOTPOS + SXL.LENGTH
            End If
            'SXL�̕i�Ԃ�ݒ肷��
            With .Hins(CStr(SXL.INGOTPOS))
                .hinban = SXL.hinban
                .REVNUM = SXL.REVNUM
                .factory = SXL.factory
                .opecond = SXL.opecond
            End With
        Next
    
        ''���グ���𒴂���`����e�𒲐�����
        .BlkPlans.LimitByIngotPos .BODYLENG
        .Blks.LimitByIngotPos .BODYLENG
        .HinPlans.LimitByIngotPos .BODYLENG
        .Hins.LimitByIngotPos .BODYLENG
        
        ''���ؒf�u���b�N���폜����
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
    
    '' �`�悷��
    UserControl_Resize

proc_exit:
    '�I��
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�R���g���[�� Initialize������
'����      :
'����      :2001/05/17 �쐬  �쑺
Private Sub UserControl_Initialize()

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "o_cmzc002a.ctl -- Sub UserControl_Initialize"

    '' �����ϐ�������������
    Clear

proc_exit:
    '�I��
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Sub

'�T�v      :�`�揈��
'����      :�����f�[�^�Ɋ�Â��ĕ`����s��
'����      :2001/05/17 �쐬  �쑺
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
    Dim blk As c_cmzcBlk        '�`��Ώۂ� Blk
    Dim Cut As c_cmzcCut        '�`��Ώۂ� Cut
    Dim HIN As c_cmzcHin        '�`��Ώۂ� Hin
    Dim SXL As c_cmzcSxl        '�`��Ώۂ� Sxl
    Dim XlSmp As c_cmzcXlSmp    '�`��Ώۂ� XlSmp
    Dim WFSMP As c_cmzcWfSmp    '�`��Ώۂ� WfSmp
    Dim Rej As c_cmzcRej        '�`��Ώۂ� rej
    Dim drawTarget As Boolean   '�����`�悷�邩

'   Debug.Print "Ctl:Paint"

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "o_cmzc002a.ctl -- Sub UserControl_Paint"

    Cls     '�ŏ��ɏ���
    
    With m_Xl
        '' �����}�̊O�g��`�悷��
        DrawStyle = vbSolid
        Line (pxXL_Center, pyXL_Top)-(pxXL_Left, pyXL_Zero), vbBlack
        Line (pxXL_Center, pyXL_Top)-(pxXL_Right, pyXL_Zero), vbBlack
        Line (pxXL_Left, pyXL_Bot)-(pxXL_Center, pyXL_Tail), vbBlack
        Line (pxXL_Right, pyXL_Bot)-(pxXL_Center, pyXL_Tail), vbBlack
        
        '' ���u���b�N�̔w�i�F��ς���(���u���b�N�w�肪�Ȃ��ꍇ�͑S�攒)
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
        
        '' �ǉ��h�[�v�ʒu��`�悷��
        If .ADDDPPOS Then
            py = GetY(.ADDDPPOS)
            DrawStyle = vbSolid
            'px = Width - (Width - pxXL_Right) / 2
            px = pxXL_Right + TextWidth("0000")
            Line (pxXL_Center, py)-(px, py), vbMagenta
            s = "�ǉ�Dope"
            FontSize = 8
            CurrentY = py - TextHeight(s) / 2
            CurrentX = px + 20
            ForeColor = vbMagenta
            Print s;
            ForeColor = vbBlack
        End If
        
        '' �����ʒu��`�悷��
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
    
        '' �����}�̗��e�̏c����`�悷��
        Line (pxXL_Left, pyXL_Zero)-(pxXL_Left, pyXL_Bot), vbBlack
        Line (pxXL_Right, pyXL_Zero)-(pxXL_Right, pyXL_Bot), vbBlack
    
        '' �T���v��������Ƃ��ƂȂ��Ƃ��ŁA�ؒf�ʒu�E�i�ԋ�؈ʒu�̒�����ς���
        If .WfSmps.COUNT Then
            margin = MARGIN_CENTER
        Else
            margin = 0
        End If
    
        '' �u���b�N��`�悷��
        py = GetY(0)
        For Each blk In .Blks
            pos1 = blk.INGOTPOS                 '�u���b�N��[
            pos2 = blk.INGOTPOS + blk.LENGTH    '�u���b�N���[
            py = GetY(pos1)
            py2 = GetY(pos2)
            
            ''�u���b�N�J�n�ʒu�̕`��
            Line (pxXL_Left, py)-(pxXL_Center - margin, py), vbBlack
            s = Str(pos1)
            FontSize = 8
            CurrentY = py - TextHeight(s) / 2
            CurrentX = pxXL_Left - TextWidth(s) - 20
            Print s;
            
            ''�u���b�N�I���ʒu�̕`��
            Line (pxXL_Left, py2)-(pxXL_Center - margin, py2), vbBlack
            s = Str(pos2)
            FontSize = 8
            CurrentY = py2 - TextHeight(s) / 2
            CurrentX = pxXL_Left - TextWidth(s) - 20
            Print s;
            
            ''�u���b�NID�̕`��
            drawTarget = True
            s = blk.BLOCKID
            If Len(s) = 0 Then              '' �u���b�NID����Ȃ�A�`��ΏۊO
                drawTarget = False
            ElseIf pos2 - pos1 <= 1 Then
                drawTarget = False
            ElseIf .Cuts.ExistInArea(blk.INGOTPOS + 1, blk.LENGTH - 1) Then '' �u���b�NTop/Bot���������Ԃɐؒf�w�����܂�ł�����`��ΏۊO
                drawTarget = False
            Else
                s = Right$(s, 3)
                If Mid$(s, 2, 1) = "$" Then '' �u$�v���܂ރu���b�NID�́A�`��ΏۊO
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
            
            ''�H���R�[�h�̕`��
            'drawTarget = True
            s = blk.NOWPROC
            If (s = vbNullString) Then              '' �H���R�[�h���o�^�̏ꍇ�͕`��ΏۊO
                drawTarget = False
            ElseIf blk.DELCLS = "1" Then
                Select Case blk.LSTATCLS
                  Case "R"
                    s = "����"
                  Case "H"
                    s = "ʲ�"
                  Case "W"
                    s = "WF�o��"
                  Case "B"
                    s = "BAR�o��"
                  Case "V"
                    s = "�O��"
                End Select
            ElseIf blk.HOLDCLS = "1" Then
                s = "ΰ���"
            Else
                s = blk.NOWPROC
                If (s = "CB320") Then               '' �N���X�^���J�^���O�̂Ƃ���<��۸�>�ƕ`��
                    s = "��۸�"
                'ElseIf (Left$(s, 2) = "CB") Then    '' ���̑������n�̎��͕`��ΏۊO(�������g�E�p��)
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
        
        '' �ؒf�w����`�悷��
        py = GetY(0)
        For Each Cut In .Cuts
            pos1 = Cut.INGOTPOS                 '�u���b�N��[
            pos2 = Cut.INGOTPOS + Cut.LENGTH    '�u���b�N���[
            py = GetY(pos1)
            py2 = GetY(pos2)
            
            '' �ؒf�w�����̂̕`��
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
                '' �u���b�NID�̕`��
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
        
        '' �i�ԋ�؈ʒu��`�悷��
        py = GetY(0)
        For Each HIN In .Hins
            py = GetY(HIN.INGOTPOS)
            py2 = GetY(HIN.INGOTPOS + HIN.LENGTH)
            
            '' �i�ԊJ�n�ʒu�̕`��
            DrawStyle = vbDot
            Line (pxXL_Center + margin, py)-(pxXL_Right, py), vbBlack
            s = " " & CStr(HIN.INGOTPOS)
            FontSize = 8
            CurrentY = py - TextHeight(s) / 2
            CurrentX = pxXL_Right + 20
            Print s;
            
            '' �i�ԏI���ʒu�̕`��
            DrawStyle = vbDot
            Line (pxXL_Center + margin, py2)-(pxXL_Right, py2), vbBlack
            s = " " & CStr(HIN.INGOTPOS + HIN.LENGTH)
            FontSize = 8
            CurrentY = py2 - TextHeight(s) / 2
            CurrentX = pxXL_Right + 20
            Print s;
            
            '' �i�Ԃ̕`��
            s = Trim$(HIN.hinban)
            FontSize = 8
            CurrentY = (py + py2 - TextHeight(s)) / 2
            CurrentX = (pxXL_Right + pxXL_Center - TextWidth(s)) / 2
            ForeColor = vbBlack
            Print s;
            ForeColor = vbBlack
        Next
    
        '' WF�T���v���ʒu��`�悷��
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
    '�I��
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Sub

'�T�v      :Resize������
'����      :�}�̊�{�T�C�Y���v�Z���A�`�悷��
'����      :2001/05/17 �쐬  �쑺
Private Sub UserControl_Resize()
    Dim totalLen As Long
    Dim totalHeight As Long
    Dim zoom As Double

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "o_cmzc002a.ctl -- Sub UserControl_Resize"
    
    ''�R���g���[���̑傫���ɍ��킹�A�}�̊�{�T�C�Y���v�Z����
    With m_Xl
        If (.TOPLENG + .BODYLENG + .BOTLENG = 0) Then GoTo proc_exit
        
        pxXL_Left = Width / 4                       ''�����}���[�ʒu
        pxXL_Right = Width - Width / 4              ''�����}�E�[�ʒu
        pxXL_Center = (pxXL_Left + pxXL_Right) / 2  ''�����}���S�ʒu
        pyXL_Top = 200                              ''�����}Top�ʒu
        pyXL_Tail = Height - 200                    ''�����}Tail�ʒu
        
        totalLen = .TOPLENG + .BODYLENG + .BOTLENG
        totalHeight = pyXL_Tail - pyXL_Top
        If totalLen = 0 Then
            zoom = totalHeight
        Else
            zoom = totalHeight * 1# / totalLen
        End If
        pyXL_Zero = pyXL_Top + .TOPLENG * zoom      ''�����}����Top�[�ʒu
        pyXL_Bot = pyXL_Tail - .BOTLENG * zoom      ''�����}����Tail�[�ʒu
        
        'Debug.Print pyXL_Top, pyXL_Zero, pyXL_Bot, pyXL_Tail, zoom
    End With
    UserControl_Paint

proc_exit:
    '�I��
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Sub

'�T�v      :Terminate������
'����      :
'����      :2001/05/17 �쐬  �쑺
Private Sub UserControl_Terminate()

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "o_cmzc002a.ctl -- Sub UserControl_Terminate"

    '' �����N���X���������
    Set m_Xl = Nothing

proc_exit:
    '�I��
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Sub




'�T�v      :�C���S�b�g���ʒu�ɑΉ�����R���g���[�������W(Y)�𓾂�
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :pos           ,I  ,Integer   ,�C���S�b�g���ʒu
'          :�߂�l        ,O  ,Long      ,Y���W
'����      :
'����      :2001/05/17 �쐬  �쑺
Private Function GetY(pos As Integer) As Long
    Dim bodyHeight As Long

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "o_cmzc002a.ctl -- Function GetY"
    
    ''�Ή�����R���g���[�������W(Y)���v�Z����
    bodyHeight = pyXL_Bot - pyXL_Zero
    If m_Xl.BODYLENG = 0 Then
        GetY = pyXL_Zero + bodyHeight
    Else
        GetY = pyXL_Zero + bodyHeight * (pos * 1# / m_Xl.BODYLENG)
    End If

proc_exit:
    '�I��
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :��菬�����l��I������
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :value1        ,I  ,Long      ,�l1
'          :value2        ,I  ,Long      ,�l2
'          :�߂�l        ,O  ,Long      ,�������l
'����      :
'����      :2001/05/17 �쐬  �쑺
Private Function LowerValue(value1 As Long, value2 As Long) As Long

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "o_cmzc002a.ctl -- Function LowerValue"

    ''�^����ꂽ�l�̓��A����������Ԃ�
    If value1 < value2 Then
        LowerValue = value1
    Else
        LowerValue = value2
    End If

proc_exit:
    '�I��
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :���傫���l��I������
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :value1        ,I  ,Long      ,�l1
'          :value2        ,I  ,Long      ,�l2
'          :�߂�l        ,O  ,Long      ,�傫���l
'����      :
'����      :2001/05/17 �쐬  �쑺
Private Function HigherValue(value1 As Long, value2 As Long) As Long

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "o_cmzc002a.ctl -- Function HigherValue"

    If value1 > value2 Then
        HigherValue = value1
    Else
        HigherValue = value2
    End If

proc_exit:
    '�I��
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Function
