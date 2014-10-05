Attribute VB_Name = "kpkcommon"
' ���Y�Ǘ��p
Public SeisanOraDB As OraDatabase 'oracle db object
Public SeisanOraSess As OraSession 'oracle session object

'TextComboButtonChenge��TYPE��`
Public Const TXT_CHENGE As Long = 1     '�Ώۂ̓e�L�X�g�̂�
Public Const COM_CHENGE As Long = 2     '�Ώۂ̓R���{�̂�
Public Const BTN_CHENGE As Long = 4     '�Ώۂ̓{�^���̂�
Public Const T_C_CHENGE As Long = TXT_CHENGE + COM_CHENGE   '�Ώۂ̓e�L�X�g�A�R���{
Public Const T_B_CHENGE As Long = TXT_CHENGE + BTN_CHENGE   '�Ώۂ̓e�L�X�g�A�{�^��
Public Const C_B_CHENGE As Long = COM_CHENGE + BTN_CHENGE   '�Ώۂ̓R���{�A�{�^��
Public Const T_C_B_CHENGE As Long = TXT_CHENGE + COM_CHENGE + BTN_CHENGE  '�Ώۂ̓e�L�X�g�A�R���{�A�{�^��

Public l_OverDay As Long '10000�ȏ�̃f�[�^������ꍇ�ɉ��Z����
' @(f)
' �@�\    : <�J�n���Ǝ���><�����߂�����t�Ǝ���>��Ԃ�
'
' �Ԃ�l  : ����=True�@�ُ�=False
'
' ������:
'           sStDate  : �J�n��
'           sEndDate : �I����
'
Public Function HidimeKeisan(sStDate As String, SENDDATE As String) As Boolean
    Dim bErrFlg As Boolean
    Dim sTime As String
    Dim sWk_Date As String
    Dim dateWk   As Date
    Dim lNumber As Long
    
    lNumber = 86400 - 1  '�Q�S���ԁi�b�j�|�P�b
    bErrFlg = True
        
    sTime = F_DbConectEndTime("X", "80", "1")
    sStDate = Mid(sStDate, 1, 4) & "/" & Mid(sStDate, 5, 2) & "/" & Mid(sStDate, 7, 2) & " " & sTime
    SENDDATE = Mid(SENDDATE, 1, 4) & "/" & Mid(SENDDATE, 5, 2) & "/" & Mid(SENDDATE, 7, 2) & " " & sTime
    dateWk = SENDDATE
    SENDDATE = DateAdd("S", lNumber, dateWk)

    HidimeKeisan = bErrFlg
End Function

' @(f)
' �@�\    : �����߂��鎞�Ԃ�TB���甲���o��
'
' �Ԃ�l  : "7:00:00"or �eTB�f�[�^
'
' ������:   s_SYSCA9   : SYSCA9�̌ďo������
'           s_SHUCA9   : SHUCA9�̌ďo������
'           s_CODEA9   : CODEA9�̌ďo������
'
' �@�\����: Table:KODA9��������ďo
'
Public Function F_DbConectEndTime(s_SYSCA9 As String, _
                                  s_SHUCA9 As String, _
                                  s_CODEA9 As String) As String
  Dim dynOraDyn As OraDynaset
  Dim wk_koteiCdName As String
  Dim s_SQL As String
  Dim i_Lp As Integer
  
  '�����l
  F_DbConectEndTime = "07:00:00"
  
            s_SQL = "SELECT"
    s_SQL = s_SQL + " KCODE01A9 "
    s_SQL = s_SQL + "FROM"
    s_SQL = s_SQL + " KODA9 "
    s_SQL = s_SQL + "WHERE"
    s_SQL = s_SQL + " SYSCA9 = '" + s_SYSCA9 + "' AND"
    s_SQL = s_SQL + " SHUCA9 = '" + s_SHUCA9 + "' AND"
    s_SQL = s_SQL + " CODEA9 = '" + s_CODEA9 + "'"
    
    '�I���N���ڑ�
    If DynSet2(dynOraDyn, s_SQL) = False Then
    
        ''�_�C�i�Z�b�g�쐬���s
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
' �@�\    : ComBox��DB����擾�����l������B
'
' �Ԃ�l  : True �� ����   False ���@�ُ�
'
' ������:   ComBoxName : �������ރR���{Box   Null�l�s��
'           s_SYSCA9   : SYSCA9�̌ďo������  NULL�l�s��
'           s_SHUCA9   : SHUCA9�̌ďo������  NULL�l�s��
'           s_CODEA9   : CODEA9�̌ďo������  �ȗ���
'           s_CTR01A9  : CTR01A9�̌ďo������ �ȗ���
'
' �@�\����: Table:KODA9����NAMESJA9�������Ōďo
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
    
    '�I���N���ڑ�
    If DynSet2(dynOraDyn, s_SQL) = False Then
    ''�_�C�i�Z�b�g�쐬���s
        Call MsgOut(100, "", ERR_DISP_LOG, "TBCMB002")
        Exit Function
    End If
            
    '�R���{�{�b�N�X�ɍ��ڂ�\������
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
' �@�\    : ���C���t�H�[���̃R���{Box��Null�l�������ꍇ�A�擾����Index��\������
'
' �Ԃ�l  : �Ȃ�
'
' ������ : �R���{Box
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
' �@�\      : �J�n������I�����܂ł̎��Ԃ��X�g�����O�ŕԂ�
'
' �Ԃ�l    :�@"0000:00.0"
'
'       StatDay �F �J�n���iDate�^�ɕϊ��ł���`���j
'       EndDay  �F �I�����iDate�^�ɕϊ��ł���`���j
'
Function F_TimeStatEnd(StatDay, EndDay) As String
  Dim l_Day  As Long    '���t
  Dim d_Time As Double  '����
  Dim d_Min  As Double
  Dim s_Day  As String
  Dim i_Canma As Integer
    
    F_TimeStatEnd = "0000$00.0"
    
    '���������ꍇ�͏����𔲂���
    If StatDay = "" Or EndDay = "" Then _
        Exit Function
    
    '�J�n�����I���������ɂȂ����ꍇ�͏����𔲂���
    If CDate(StatDay) > CDate(EndDay) Then Exit Function
    
    d_Min = DateDiff("n", StatDay, EndDay)
    '1���ȏ�̏ꍇ
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
    
    '�����_�ȉ����ʂ܂ŕ\������
    d_Time = CInt(d_Time * 10)
    d_Time = d_Time * 0.1
            
    F_TimeStatEnd = F_DayTimeAr(l_Day, d_Time)

End Function
' @(f)
'
' �@�\      : ���v���Ԃ��畽�ς�����o��
'
' �Ԃ�l    :�@"0000:00.0"   �f�[�^������ "0000$00.0"
'
'       SumTime  �F ���v����(0000:00.0)
'       i_AvgCnt �F ���� (����)
'
Function F_DayTimeAvg(SumTime As String, i_AvgCnt As Integer, Optional OverDay As Long) As String
  Dim l_Day   As Long
  Dim l_Day_1 As Double
  Dim d_Time  As Double
  Dim d_Time2 As Double
  Dim d_Time3 As Double
  Dim i_Time  As Integer
  
    '�f�[�^�������̂ŏ������s��Ȃ�
    If SumTime = "0000$00.0" Then Exit Function
    If SumTime = "" Then Exit Function
    
    '���t�̕���
    l_Day_1 = Left(SumTime, 4) / i_AvgCnt
    If l_Day_1 >= 1 Then
        l_Day = l_Day_1
    Else
        l_Day = 0
    End If
    '���t�̏����_�������Ԋ��Z
    d_Time3 = (l_Day_1 - l_Day) * 24
    If d_Time3 < 0 Then
        l_Day = l_Day - 1
        d_Time3 = (l_Day_1 - (l_Day)) * 24
    End If
    '���Ԃ̌v�Z
    d_Time = (CDbl(Right(SumTime, 4)) / i_AvgCnt) + d_Time3
    
    If d_Time >= 24 Then
        l_Day = (l_Day + 1) / i_AvgCnt
        d_Time = d_Time - 24
    End If
    
    '�����_�ȉ����ʂ܂ŕ\������
    i_Time = CInt(d_Time * 10)
    d_Time = i_Time * 0.1
    
    If OverDay <> 0 Then _
        l_Day = l_Day + (OverDay * 10000)
    
    F_DayTimeAvg = F_DayTimeAr(l_Day, d_Time)

End Function
' @(f)
'
' �@�\      : ���v���Z�o
'
' �Ԃ�l    :�@"0000:00.0"
'       Sumtime_1 �F ���v�����̐��̈��'
'       Sumtime_2 �F ���v�����̐��̈��
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
    
    '9999�����������f�[�^�̓G���[��\������(�ő�łQ�T�����ȏ�j
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
' �@�\      : �\���`���ϊ�
'
' �Ԃ�l    :�@"####:##.#"
'
Function F_DayTimeAr(l_Day As Long, d_Time As Double) As String
  Dim s_Sp1 As String
  Dim s_Sp2 As String
  Dim s_Sp3 As String

    F_DayTimeAr = "0000:00.0"
    
    If InStr(1, CStr(d_Time), ".") = 0 Then
        s_Sp3 = ".0"
    End If

    '���X�y�[�X
    Select Case l_Day
      Case 0 To 9: s_Sp1 = "000"
      Case 10 To 99: s_Sp1 = "00"
      Case 100 To 999: s_Sp1 = "0"
      Case 1000 To 9999: s_Sp1 = ""
    End Select
    '���ԃX�y�[�X
    Select Case d_Time
      Case 0 To 9.9: s_Sp2 = "0"
      Case 10 To 24: s_Sp2 = ""
    End Select
    
    F_DayTimeAr = s_Sp1 & CStr(l_Day) & ":" & s_Sp2 & CStr(d_Time) & s_Sp3

End Function
' @(f)
'
' �@�\      : �\���`���ϊ�
'
' �Ԃ�l    :�@"###��:##.#"
'
Function F_DispDayTime(s_DayTime As String) As String
  Dim s_Sp1 As String
  Dim s_Sp2 As String
  Dim s_Sp3 As String
  Dim l_Day As Long
  Dim d_Time As Double

    F_DispDayTime = "  0�� 0.0"
    If s_DayTime = "" Then Exit Function
    
    l_Day = CInt(Left(s_DayTime, 4))
    d_Time = CDbl(Right(s_DayTime, 4))
    
    
    If InStr(1, CStr(d_Time), ".") = 0 Then
        s_Sp3 = ".0"
    End If

    '���X�y�[�X
    Select Case l_Day
      Case 0 To 9: s_Sp1 = "  "
      Case 10 To 99: s_Sp1 = " "
      Case 100 To 999: s_Sp1 = ""
    End Select
    '���ԃX�y�[�X
    Select Case d_Time
      Case 0 To 9.9: s_Sp2 = " "
      Case 10 To 24: s_Sp2 = ""
    End Select
    
    F_DispDayTime = s_Sp1 & CStr(l_Day) & "��" & s_Sp2 & CStr(d_Time) & s_Sp3

End Function
' @(f)
'
' �@�\      : �\���`���ϊ�
'
' �Ԃ�l    :�@"###��##.#" �� �� "######.#"
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
' �@�\      : ��������d�ʂ����߂�v�Z����
'
' �Ԃ�l    :�@�d��
'
Function fncNagaWeightChg(lNagasa As Long) As Long
    fncNagaWeightChg = (301 / 2) ^ 2 * 3.1416 * lNagasa * 2.33 / 1000
End Function

' @(f)
'
' �@�\      : ���Y�Ǘ��c�a �n�o�d�m
'
' �Ԃ�l    :�@�d��
'

Public Function OraDBSeisanOpen() As Boolean
    'Oracle Session Object
        Dim sDbName As String
    Dim sUID As String
    Dim sPWD As String
    
'    Select Case gsFactryCd
'    Case "42"               '�f�R�O�O����
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

'�T�v      :Oracle�̃Z�b�V���������(���Y�Ǘ��c�a)
'����      :�A�v���P�[�V�����̏I�����ɌĂ�
'����      :
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
' �@�\    :�I���N���_�C�i�Z�b�g�̍쐬(���Y�Ǘ��c�a)
'
' �Ԃ�l  : ���� - true
'           �ُ� - false
'
' ������  : ARG1 - �_�C�i�Z�b�g�Z�b�g�I�u�W�F�N�g
'           ARG2 - SQL��
'           ARG3 - �_�C�i�Z�b�g�I�v�V����
'
' �@�\����: �I���N���_�C�i�Z�b�g�쐬
'
'///////////////////////////////////////////////////
Public Function DynSetSeisan(ByRef objOraDynaset As Object, sSqlStmt As String, Optional vOpt = &H4&) As Boolean
    On Error GoTo DynErr
    
    ''�I���N���_�C�i�Z�b�g�쐬
    Set objOraDynaset = SeisanOraDB.CreateDynaset(sSqlStmt, vOpt)
    DynSetSeisan = True
    Exit Function
    
DynErr:
    DynSetSeisan = False
End Function

'///////////////////////////////////////////////////
' @(f)
' �@�\    : �t�H�[����̃R���g���[�����g�p�s�ɂ���
'
' �Ԃ�l  :
'
' ������  : �t�H�[��
'        �F �Ώ�(lType)  1:�e�L�X�g 2:�R���{�{�b�N�X 4:�{�^��
'        �F �i�[�l(bSet) true/false
'
' �@�\����: �w�肵���t�H�[���́u[�e�P]Ҳ��ƭ��v�{�^���ȊO
'           �̃R���g���[�����g�p�s�ɂ���
'
'///////////////////////////////////////////////////
Public Sub TextComboButtonChenge(frmForm As Form, lType As Long, bSet As Boolean)
    Dim iIdx As Integer
    Dim ctlControl As Control
    
    
    ''�t�H�[����̃R���g���[����S�Ďg�p�s�ɂ���
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
    
    ''�u[�e�P]Ҳ��ƭ��v�{�^�����g�p�\�ɂ���
    If ((lType And BTN_CHENGE) = BTN_CHENGE) Then
        frmForm.cmdF(1).Enabled = True
    End If
End Sub


