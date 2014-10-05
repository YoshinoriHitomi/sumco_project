Attribute VB_Name = "s_cmzcinpchk"
Option Explicit

''==========================================
'' ���̓`�F�b�N�֐��Q
''==========================================


'' ���̓`�F�b�N�֐��̖߂�l
Public Enum CHK_RESULT
    CHK_OK          '' ����
    CHK_NG          '' �ُ�
    CHK_NULL        '' ������
End Enum

Public Enum CHK_TYPE
    CHK_NUMBER      '' ���l
    CHK_NUMSTR      '' ������
    CHK_STRING      '' ������
End Enum

Public Enum CHK_NUMTYPE
    NUMTYPE_ALL         ''+/0/- �S��OK
    NUMTYPE_PLUS        ''+ �̂�OK
    NUMTYPE_ZEROPLUS    ''+/0 �̂�OK
End Enum
    

'���l�t�H�[�}�b�g(�����_�ȉ��͂P���ȏ�œ����Ă���Ƃ���܂�)
'�����݂̂̏ꍇ
Public Const FMT_U0 = "#,##0; ; "
Public Const FMT_M0 = "#,##0;-#,##0; "
'���̐��̂ݕ\������ꍇ
Public Const FMT_U1 = "0.0; ; "
Public Const FMT_U2 = "0.0#; ; "
Public Const FMT_U3 = "0.0##; ; "
Public Const FMT_U4 = "0.0###; ; "
Public Const FMT_U5 = "0.0####; ; "
Public Const FMT_U6 = "0.0#####; ; "
'���̐����\������ꍇ
Public Const FMT_M1 = "0.0;-0.0; "
Public Const FMT_M2 = "0.0#;-0.0#; "
Public Const FMT_M3 = "0.0##;-0.0##; "
Public Const FMT_M4 = "0.0###;-0.0###; "
Public Const FMT_M5 = "0.0####;-0.0####; "
Public Const FMT_M6 = "0.0#####;-0.0#####; "

'�T�v      :���l�^���͂̃`�F�b�N
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :s             ,I  ,String    ,�]���Ώە�����
'          :upperLen      ,I  ,Integer   ,����������
'          :lowerLen      ,I  ,Integer   ,�����_�ȉ�����
'          :�߂�l        ,O  ,CHK_RESULT,�`�F�b�N����
'����      :�����E�������܂߂����l�̌����`�F�b�N���s���A����Ȃ��CHK_OK�Ƃ���B
'����      :2001/06/20(wed) ����  �쐬
Public Function ChkNumber(ByVal s$, ByVal upperLen%, ByVal lowerLen%, Optional numType As CHK_NUMTYPE = NUMTYPE_ALL) As CHK_RESULT
Dim Txt_Str     As String               '�]���Ώە�����
Dim Str_Num()   As String               '�Ώە����z��
Dim Num         As String
    

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcinpchk.bas -- Function ChkNumber"

    ChkNumber = CHK_NG                      '�X�e�[�^�X�����l�Z�b�g(error)
    Txt_Str = s                             '����s�Z�b�g

    '*** Null�`�F�b�N ***
    If Trim$(Txt_Str) = vbNullString Then '����s��Null�l�̏ꍇ
        ChkNumber = CHK_NULL                'Null�X�e�[�^�X
        GoTo proc_exit
    End If
    
    '*** �����`�F�b�N ***
    If upperLen <= 0 Then                   '�������������O�ȉ��̏ꍇ
        GoTo proc_exit
    End If

    '*** ���l�`�F�b�N ***
    If IsNumeric(Txt_Str) = False Then      '�Ώە����񂪐��l�łȂ��ꍇ
        GoTo proc_exit
    End If
    If (Right$(Txt_Str, 1) = "+") Or (Right$(Txt_Str, 1) = "-") Then    '�u10-�v����isnumeric()�͒ʂ�̂�
        GoTo proc_exit
    End If

    '*** �����`�F�b�N ***
    Str_Num = Split(Txt_Str, ".")           '����s�������_��؂�Ŕz��ɃZ�b�g
    If Str_Num(0) = "" Then                 '�������������ꍇ
        Num = 0
    Else                                    '������������ꍇ
        Num = Abs(CDbl(Str_Num(0)))         '�����Ƣ�C��L��������
    End If
        '�������`�F�b�N
    If Len(Num) > upperLen Then
        GoTo proc_exit
    End If
        '�����l�̃`�F�b�N
    If UBound(Str_Num) > 0 Then
        If Len(Str_Num(1)) > lowerLen Then
            GoTo proc_exit
        End If
    End If

    '*** �͈̓`�F�b�N ***
    If numType = NUMTYPE_PLUS Then          '+ �̂݉�
        If val(Txt_Str) <= 0# Then
            GoTo proc_exit
        End If
    ElseIf numType = NUMTYPE_ZEROPLUS Then  '+/0 �̂݉�
        If val(Txt_Str) < 0# Then
            GoTo proc_exit
        End If
    End If

    ChkNumber = CHK_OK                      '����X�e�[�^�X


proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function



'�T�v      :��������͂̃`�F�b�N
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :s             ,I  ,String    ,�]���Ώە�����
'          :sLen          ,I  ,Integer   ,�L��������
'          :�߂�l        ,O  ,CHK_RESULT,�`�F�b�N����
'����      :�L�����������傤�ǂ̐�����݂̂� CHK_OK �Ƃ���
'����      :2001/06/20 �쐬  �쑺
Public Function ChkNumStr(ByVal s$, ByVal sLen%) As CHK_RESULT

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcinpchk.bas -- Function ChkNumStr"

    If s = vbNullString Then
        ChkNumStr = CHK_NULL
    ElseIf s Like String(sLen, "#") Then
        ChkNumStr = CHK_OK
    Else
        ChkNumStr = CHK_NG
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :��������͂̃`�F�b�N
'���Ұ�    :�ϐ���        ,IO ,�^            ,����
'          :s             ,I  ,String        ,�]���Ώە�����
'          :suLen         ,I  ,Integer       ,�L��������(���)
'          :slLen         ,I  ,Integer       ,�L��������(����)
'          :�߂�l        ,O  ,CHK_RESULT,�`�F�b�N����
'����      :�w�蕶�����ȓ��Ȃ� CHK_OK �Ƃ���
'����      :2001/06/20 �쐬  �쑺
Public Function ChkString(ByVal s$, ByVal suLen%, ByVal slLen%) As CHK_RESULT
Dim chkS As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcinpchk.bas -- Function ChkString"

    chkS = StrConv(s, vbFromUnicode)
    If s = vbNullString Then
        ChkString = CHK_NULL
    ElseIf (LenB(chkS) >= slLen) And (LenB(chkS) <= suLen) Then
        ChkString = CHK_OK
    Else
        ChkString = CHK_NG
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function



'�T�v      :TextBox�̓��͓��e���`�F�b�N����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :txt           ,I  ,TextBox   ,�`�F�b�N�Ώۂ̃e�L�X�g�{�b�N�X
'          :chkType       ,I  ,CHK_TYPE  ,���̓`�F�b�N�̃^�C�v
'          :upperLen      ,I  ,Integer   ,�����_����̌����i�������j
'          :[lowerLen]    ,I  ,Integer   ,�����_��艺�̌���
'          :[outFmt]      ,I  ,String    ,�\�������w��
'          :[nullOK]      ,I  ,Boolean   ,Null����
'          :[numType]     ,I  ,CHK_NUMTYPE   ,���l�̗L���͈�
'          :�߂�l        ,O  ,FUNCTION_RETURN, ����OK/NG
'����      :
'����      :2001/06/20 �쐬  �쑺
Public Function ChkTextBox(txt As TextBox, chkType As CHK_TYPE, upperLen%, Optional lowerLen% = 0, Optional outFmt$ = vbNullString, Optional nullOK As Boolean = False, Optional numType As CHK_NUMTYPE = NUMTYPE_ALL) As FUNCTION_RETURN
Dim chkTxt As String
Dim chkFmt As String
Dim chkResult As CHK_RESULT
Dim RET As FUNCTION_RETURN
    

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcinpchk.bas -- Function ChkTextBox"

    chkTxt = Trim$(txt.Text)
    
    ''chkType �ɏ]���A���̓`�F�b�N�֐����Ăяo��
    Select Case chkType
      Case CHK_NUMBER
            chkResult = ChkNumber(chkTxt, upperLen, lowerLen, numType)
      Case CHK_NUMSTR
            chkResult = ChkNumStr(chkTxt, upperLen)
      Case CHK_STRING
            chkResult = ChkString(chkTxt, upperLen, lowerLen)
    End Select
    
    ''Null�̋���/�s���𓥂܂��ă`�F�b�N���ʂ�]������
    If nullOK Then
        If chkResult = CHK_NG Then
            RET = FUNCTION_RETURN_FAILURE
        Else
            RET = FUNCTION_RETURN_SUCCESS
        End If
    Else
        If chkResult = CHK_OK Then
            RET = FUNCTION_RETURN_SUCCESS
        Else
            RET = FUNCTION_RETURN_FAILURE
        End If
    End If
    
    '�`�F�b�N���ʂɂ���ĉ�ʂɔ��f����
    If RET = FUNCTION_RETURN_SUCCESS Then
        ''����OK�Ȃ�A�e�L�X�g�{�b�N�X�̔w�i�F�� COLOR_OK �ɐݒ肷��
        txt.BackColor = COLOR_OK
        '�����w�肪����ꍇ�A���`����
        If outFmt <> vbNullString Then
            txt.Text = Format$(chkTxt, outFmt)
        End If
    Else
        ''����NG�Ȃ�A�e�L�X�g�{�b�N�X�̔w�i�F�� COLOR_NG �ɐݒ肷��
        txt.BackColor = COLOR_NG
        ''�t�H�[�J�X�����̃e�L�X�g�{�b�N�X�Ɉڂ�
        txt.SetFocus
    End If
    
    ''�`�F�b�N���ʂ�Ԃ�
    ChkTextBox = RET

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :���l���w�茅�܂łɐ؂�̂Ă�
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :v             ,I  ,Double    ,���̒l
'          :col           ,I  ,Integer   ,���ʂ̏����_�ȉ�����
'          :�߂�l        ,O  ,Double    ,����
'����      :�����̐؂�̂Ăɂ��Ă̒�`�́AExcel��Trunc�֐����Q�l�ɂ���
'          :0�ȉ��̌������w�肳�ꂽ�Ƃ��́A�����l�ɐ؂�̂Ă�
'����      :2002/03/22 �쑺 �쐬
Function RoundDown(ByVal v As Double, ByVal col As Integer) As Double
Dim s As String

    If col <= 0 Then
        RoundDown = Fix(v)
    Else
        s = Format$(Abs(v), "0." & String(col + 1, "0"))
        s = Left$(s, Len(s) - 1)
        RoundDown = Sgn(v) * val(s)
    End If
End Function

'�T�v      :���l���w�茅�܂łɐ؂�グ��
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :v             ,I  ,Double    ,���̒l
'          :col           ,I  ,Integer   ,���ʂ̏����_�ȉ�����
'          :�߂�l        ,O  ,Double    ,����
'����      :�����̐؏グ�ɂ��Ă̒�`�́AExcel��RoundUp�֐����Q�l�ɂ���
'          :0�ȉ��̌������w�肳�ꂽ�Ƃ��́A�����l�ɐ؂�グ��
'����      :2002/03/22 �쑺 �쐬
Function RoundUp(ByVal v As Double, ByVal col As Integer) As Double
Dim d As Double

    If col < 0 Then col = 0
    d = Abs(RoundDown(v, col))
    If d < Abs(v) Then
        If col > 0 Then
            RoundUp = Sgn(v) * (d + val("0." & String(col - 1, "0") & "1"))
        Else
            RoundUp = Sgn(v) * (d + 1)
        End If
    Else
        RoundUp = Sgn(v) * d
    End If
End Function
