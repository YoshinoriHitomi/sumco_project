Attribute VB_Name = "s_kensa2_SQL"
'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------
'�t�B�[���h�������p
Dim fldNames() As String    '��rs�Ɋ܂܂��t�B�[���h���ێ��z��
Dim fldCnt As Integer       '��rs�Ɋ܂܂��t�B�[���h��

'�T�v      :�e�[�u���uTBCME019�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME019 ,���o���R�[�h
'          :formID        ,I  ,String       ,�g�p�t�H�[��ID
'          :sqlOrder      ,I  ,tFullHinban  ,���o�i�ԁi�z��j
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :�����L�[�́A�HINBAN�+�uMNOREVNO�v+�uFACTORY�v+�uOPECOND�v�̕�����Ƃ���
'����      :2001/06/27�쐬�@���� (2002/07 s_cmzcF_TBCME019_SQL.bas���ړ�)

Public Function DBDRV_GetTBCME019(records() As typ_TBCME019, formID$, hin() As tFullHinban) As FUNCTION_RETURN
Dim sql         As String           'SQL�S��
Dim sqlBase     As String           'SQL��{��(WHERE�߂̑O�܂�)
Dim sqlWhere    As String           'SQLWhere��
Dim rs          As OraDynaset       'RecordSet
Dim recCnt      As Long             '���R�[�h��
Dim key         As String           '����KEY
Dim i           As Long             'ٰ�߶���


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_TBCME019_SQL.bas -- Function DBDRV_GetTBCME019"

 Select Case formID
        Case "f_cmbc021_1"           '�uFTIR(Oi,Cs)���ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXTMMAXN, HSXTMSPH, HSXTMSPT," & _
              " HSXTMSPR, HSXTMKHM, HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT," & _
              " HSXLTHWS, HSXLTKWY, HSXLTNSW, HSXLTKHM, HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
              " HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
              " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS," & _
              " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH," & _
              " HSXOS1ST, HSXOS1SI, HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS, HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
              " HSXOS2SH, HSXOS2ST, HSXOS2SI, HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG "
        
        Case "f_cmbc022_1"           '�uGFA(Oi)���ѓ��́v
             sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXTMMAXN, HSXTMSPH, HSXTMSPT," & _
              " HSXTMSPR, HSXTMKHM, HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT," & _
              " HSXLTHWS, HSXLTKWY, HSXLTNSW, HSXLTKHM, HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
              " HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
              " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS," & _
              " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH," & _
              " HSXOS1ST, HSXOS1SI, HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS, HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
              " HSXOS2SH, HSXOS2ST, HSXOS2SI, HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG "
        
        Case "f_cmbc023_1"           '�u��R���ѓ��́v
           sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXTMMAXN, HSXTMSPH, HSXTMSPT," & _
              " HSXTMSPR, HSXTMKHM, HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT," & _
              " HSXLTHWS, HSXLTKWY, HSXLTNSW, HSXLTKHM, HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
              " HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
              " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS," & _
              " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH," & _
              " HSXOS1ST, HSXOS1SI, HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS, HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
              " HSXOS2SH, HSXOS2ST, HSXOS2SI, HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG "
        
        Case "f_cmbc024_1"           '�uBMD���ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXTMMAXN, HSXTMSPH, HSXTMSPT," & _
              " HSXTMSPR, HSXTMKHM, HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT," & _
              " HSXLTHWS, HSXLTKWY, HSXLTNSW, HSXLTKHM, HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
              " HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
              " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS," & _
              " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH," & _
              " HSXOS1ST, HSXOS1SI, HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS, HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
              " HSXOS2SH, HSXOS2ST, HSXOS2SI, HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG "
        
        Case "f_cmec030_1"           '�uBMD���ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXTMMAXN, HSXTMSPH, HSXTMSPT," & _
              " HSXTMSPR, HSXTMKHM, HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT," & _
              " HSXLTHWS, HSXLTKWY, HSXLTNSW, HSXLTKHM, HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
              " HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
              " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS," & _
              " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH," & _
              " HSXOS1ST, HSXOS1SI, HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS, HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
              " HSXOS2SH, HSXOS2ST, HSXOS2SI, HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG "
        
        Case "f_cmbc025_1"           '�uOSF���ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXTMMAXN, HSXTMSPH, HSXTMSPT," & _
              " HSXTMSPR, HSXTMKHM, HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT," & _
              " HSXLTHWS, HSXLTKWY, HSXLTNSW, HSXLTKHM, HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
              " HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
              " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS," & _
              " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH," & _
              " HSXOS1ST, HSXOS1SI, HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS, HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
              " HSXOS2SH, HSXOS2ST, HSXOS2SI, HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG "
        
        Case "f_cmec031_1"           '�uOSF���ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXTMMAXN, HSXTMSPH, HSXTMSPT," & _
              " HSXTMSPR, HSXTMKHM, HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT," & _
              " HSXLTHWS, HSXLTKWY, HSXLTNSW, HSXLTKHM, HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
              " HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
              " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS," & _
              " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH," & _
              " HSXOS1ST, HSXOS1SI, HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS, HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
              " HSXOS2SH, HSXOS2ST, HSXOS2SI, HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG "
        
        Case "f_cmbc026_1"           '�uGD���ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXTMMAXN, HSXTMSPH, HSXTMSPT," & _
              " HSXTMSPR, HSXTMKHM, HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT," & _
              " HSXLTHWS, HSXLTKWY, HSXLTNSW, HSXLTKHM, HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
              " HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
              " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS," & _
              " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH," & _
              " HSXOS1ST, HSXOS1SI, HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS, HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
              " HSXOS2SH, HSXOS2ST, HSXOS2SI, HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG "
        
        Case "f_cmbc027_1"           '�u���C�t�^�C�����ѓ��́v
           sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXTMMAXN, HSXTMSPH, HSXTMSPT," & _
              " HSXTMSPR, HSXTMKHM, HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT," & _
              " HSXLTHWS, HSXLTKWY, HSXLTNSW, HSXLTKHM, HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
              " HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
              " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS," & _
              " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH," & _
              " HSXOS1ST, HSXOS1SI, HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS, HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
              " HSXOS2SH, HSXOS2ST, HSXOS2SI, HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG "
        
        Case "f_cmbc028_1i"           '�uFPD���ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXTMMAXN, HSXTMSPH, HSXTMSPT," & _
              " HSXTMSPR, HSXTMKHM, HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT," & _
              " HSXLTHWS, HSXLTKWY, HSXLTNSW, HSXLTKHM, HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
              " HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
              " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS," & _
              " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH," & _
              " HSXOS1ST, HSXOS1SI, HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS, HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
              " HSXOS2SH, HSXOS2ST, HSXOS2SI, HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG "
                
        Case "f_cmbc029_1"           '�uGFA�Z�����ݒ�v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXTMMAXN, HSXTMSPH, HSXTMSPT," & _
              " HSXTMSPR, HSXTMKHM, HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT," & _
              " HSXLTHWS, HSXLTKWY, HSXLTNSW, HSXLTKHM, HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
              " HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
              " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS," & _
              " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH," & _
              " HSXOS1ST, HSXOS1SI, HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS, HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
              " HSXOS2SH, HSXOS2ST, HSXOS2SI, HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG "
    
    End Select
    
    sqlBase = sqlBase & "From TBCME019"
    
    '''SQL��Where���쐬
    For i = 0 To UBound(hin)
        With hin(i)
            key = key & "'" & .hinban & Format(.mnorevno, "00000") & .factory & .opecond & "'"
            If i <> UBound(hin) Then
                key = key & ", "
            End If
        End With
    Next
    sqlWhere = " Where(HINBAN||TO_CHAR(MNOREVNO, 'FM00000')||FACTORY||OPECOND in(" & key & "))"
    sql = sqlBase & sqlWhere
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME019 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''�t�B�[���h����o�^����
    fldCnt = rs.Fields.COUNT
    ReDim fldNames(fldCnt)
    For i = 1 To fldCnt
        fldNames(i) = rs.FieldName(i - 1)
    Next
   
     ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
         With records(i)
            If fldNameExist("HINBAN") Then .hinban = rs("HINBAN")               ' �i��
            If fldNameExist("MNOREVNO") Then .mnorevno = rs("MNOREVNO")         ' ���i�ԍ������ԍ�
            If fldNameExist("FACTORY") Then .factory = rs("FACTORY")            ' �H��
            If fldNameExist("OPECOND") Then .opecond = rs("OPECOND")            ' ���Ə���
            If fldNameExist("HMGSTRRNO") Then .HMGSTRRNO = rs("HMGSTRRNO")      ' �i�Ǘ��d�l�o�^�˗��ԍ�
            If fldNameExist("HMGSTFNO") Then .HMGSTFNO = rs("HMGSTFNO")         ' �i�Ǘ��Ј��m��
            If fldNameExist("HMGSXSNO") Then .HMGSXSNO = rs("HMGSXSNO")         ' �i�Ǘ��r�w���i�ԍ�
            If fldNameExist("HMGSXSNE") Then .HMGSXSNE = fncNullCheck(rs("HMGSXSNE"))         ' �i�Ǘ��r�w���i�ԍ��}��
            If fldNameExist("HSXTMMAXN") Then .HSXTMMAX = fncNullCheck(rs("HSXTMMAXN"))        ' �i�r�w�]�ʖ��x���    �v�e�T���v�������ύX 2003.05.20 yakimura
            If fldNameExist("HSXTMSPH") Then .HSXTMSPH = rs("HSXTMSPH")         ' �i�r�w�]�ʖ��x����ʒu�Q��
            If fldNameExist("HSXTMSPT") Then .HSXTMSPT = rs("HSXTMSPT")         ' �i�r�w�]�ʖ��x����ʒu�Q�_
            If fldNameExist("HSXTMSPR") Then .HSXTMSPR = rs("HSXTMSPR")         ' �i�r�w�]�ʖ��x����ʒu�Q��
            If fldNameExist("HSXTMKHM") Then .HSXTMKHM = rs("HSXTMKHM")         ' �i�r�w�]�ʖ��x�����p�x�Q��
            If fldNameExist("HSXTMKHI") Then .HSXTMKHI = rs("HSXTMKHI")         ' �i�r�w�]�ʖ��x�����p�x�Q��
            If fldNameExist("HSXTMKHH") Then .HSXTMKHH = rs("HSXTMKHH")         ' �i�r�w�]�ʖ��x�����p�x�Q��
            If fldNameExist("HSXTMKHS") Then .HSXTMKHS = rs("HSXTMKHS")         ' �i�r�w�]�ʖ��x�����p�x�Q��
            If fldNameExist("HSXLTMIN") Then .HSXLTMIN = fncNullCheck(rs("HSXLTMIN"))         ' �i�r�w�k�^�C������ 'NULL�Ή�
            If fldNameExist("HSXLTMAX") Then .HSXLTMAX = fncNullCheck(rs("HSXLTMAX"))         ' �i�r�w�k�^�C����� 'NULL�Ή�
            If fldNameExist("HSXLTSPH") Then .HSXLTSPH = rs("HSXLTSPH")         ' �i�r�w�k�^�C������ʒu�Q��
            If fldNameExist("HSXLTSPT") Then .HSXLTSPT = rs("HSXLTSPT")         ' �i�r�w�k�^�C������ʒu�Q�_
            If fldNameExist("HSXLTSPI") Then .HSXLTSPI = rs("HSXLTSPI")         ' �i�r�w�k�^�C������ʒu�Q��
            If fldNameExist("HSXLTHWT") Then .HSXLTHWT = rs("HSXLTHWT")         ' �i�r�w�k�^�C���ۏؕ��@�Q��
            If fldNameExist("HSXLTHWS") Then .HSXLTHWS = rs("HSXLTHWS")         ' �i�r�w�k�^�C���ۏؕ��@�Q��
            If fldNameExist("HSXLTKWY") Then .HSXLTKWY = rs("HSXLTKWY")         ' �i�r�w�k�^�C���������@
            If fldNameExist("HSXLTNSW") Then .HSXLTNSW = rs("HSXLTNSW")         ' �i�r�w�k�^�C���M�����@
            If fldNameExist("HSXLTKHM") Then .HSXLTKHM = rs("HSXLTKHM")         ' �i�r�w�k�^�C�������p�x�Q��
            If fldNameExist("HSXLTKHI") Then .HSXLTKHI = rs("HSXLTKHI")         ' �i�r�w�k�^�C�������p�x�Q��
            If fldNameExist("HSXLTKHH") Then .HSXLTKHH = rs("HSXLTKHH")         ' �i�r�w�k�^�C�������p�x�Q��
            If fldNameExist("HSXLTKHS") Then .HSXLTKHS = rs("HSXLTKHS")         ' �i�r�w�k�^�C�������p�x�Q��
            If fldNameExist("HSXLTMBP") Then .HSXLTMBP = fncNullCheck(rs("HSXLTMBP"))         ' �i�r�w�k�^�C���ʓ����z
            If fldNameExist("HSXLTMCL") Then .HSXLTMCL = rs("HSXLTMCL")         ' �i�r�w�k�^�C���ʓ��v�Z
            If fldNameExist("HSXCNMIN") Then .HSXCNMIN = fncNullCheck(rs("HSXCNMIN"))         ' �i�r�w�Y�f�Z�x����
            If fldNameExist("HSXCNMAX") Then .HSXCNMAX = fncNullCheck(rs("HSXCNMAX"))         ' �i�r�w�Y�f�Z�x���
            If fldNameExist("HSXCNSPH") Then .HSXCNSPH = rs("HSXCNSPH")         ' �i�r�w�Y�f�Z�x����ʒu�Q��
            If fldNameExist("HSXCNSPT") Then .HSXCNSPT = rs("HSXCNSPT")         ' �i�r�w�Y�f�Z�x����ʒu�Q�_
            If fldNameExist("HSXCNSPI") Then .HSXCNSPI = rs("HSXCNSPI")         ' �i�r�w�Y�f�Z�x����ʒu�Q��
            If fldNameExist("HSXCNHWT") Then .HSXCNHWT = rs("HSXCNHWT")         ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
            If fldNameExist("HSXCNHWS") Then .HSXCNHWS = rs("HSXCNHWS")         ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
            If fldNameExist("HSXCNKWY") Then .HSXCNKWY = rs("HSXCNKWY")         ' �i�r�w�Y�f�Z�x�������@
            If fldNameExist("HSXCNKHM") Then .HSXCNKHM = rs("HSXCNKHM")         ' �i�r�w�Y�f�Z�x�����p�x�Q��
            If fldNameExist("HSXCNKHI") Then .HSXCNKHI = rs("HSXCNKHI")         ' �i�r�w�Y�f�Z�x�����p�x�Q��
            If fldNameExist("HSXCNKHH") Then .HSXCNKHH = rs("HSXCNKHH")         ' �i�r�w�Y�f�Z�x�����p�x�Q��
            If fldNameExist("HSXCNKHS") Then .HSXCNKHS = rs("HSXCNKHS")         ' �i�r�w�Y�f�Z�x�����p�x�Q��
            If fldNameExist("HSXONMIN") Then .HSXONMIN = fncNullCheck(rs("HSXONMIN"))         ' �i�r�w�_�f�Z�x����
            If fldNameExist("HSXONMAX") Then .HSXONMAX = fncNullCheck(rs("HSXONMAX"))         ' �i�r�w�_�f�Z�x���
            If fldNameExist("HSXONSPH") Then .HSXONSPH = rs("HSXONSPH")         ' �i�r�w�_�f�Z�x����ʒu�Q��
            If fldNameExist("HSXONSPT") Then .HSXONSPT = rs("HSXONSPT")         ' �i�r�w�_�f�Z�x����ʒu�Q�_
            If fldNameExist("HSXONSPI") Then .HSXONSPI = rs("HSXONSPI")         ' �i�r�w�_�f�Z�x����ʒu�Q��
            If fldNameExist("HSXONHWT") Then .HSXONHWT = rs("HSXONHWT")         ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
            If fldNameExist("HSXONHWS") Then .HSXONHWS = rs("HSXONHWS")         ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
            If fldNameExist("HSXONKWY") Then .HSXONKWY = rs("HSXONKWY")         ' �i�r�w�_�f�Z�x�������@
            If fldNameExist("HSXONKHM") Then .HSXONKHM = rs("HSXONKHM")         ' �i�r�w�_�f�Z�x�����p�x�Q��
            If fldNameExist("HSXONKHI") Then .HSXONKHI = rs("HSXONKHI")         ' �i�r�w�_�f�Z�x�����p�x�Q��
            If fldNameExist("HSXONKHH") Then .HSXONKHH = rs("HSXONKHH")         ' �i�r�w�_�f�Z�x�����p�x�Q��
            If fldNameExist("HSXONKHS") Then .HSXONKHS = rs("HSXONKHS")         ' �i�r�w�_�f�Z�x�����p�x�Q��
            If fldNameExist("HSXONMBP") Then .HSXONMBP = fncNullCheck(rs("HSXONMBP"))         ' �i�r�w�_�f�Z�x�ʓ����z
            If fldNameExist("HSXONMCL") Then .HSXONMCL = rs("HSXONMCL")         ' �i�r�w�_�f�Z�x�ʓ��v�Z
            If fldNameExist("HSXONLTB") Then .HSXONLTB = fncNullCheck(rs("HSXONLTB"))         ' �i�r�w�_�f�Z�x�k�s���z
            If fldNameExist("HSXONLTC") Then .HSXONLTC = rs("HSXONLTC")         ' �i�r�w�_�f�Z�x�k�s�v�Z
            If fldNameExist("HSXONSDV") Then .HSXONSDV = fncNullCheck(rs("HSXONSDV"))         ' �i�r�w�_�f�Z�x�W���΍�
            If fldNameExist("HSXONAMN") Then .HSXONAMN = fncNullCheck(rs("HSXONAMN"))         ' �i�r�w�_�f�Z�x���ω���
            If fldNameExist("HSXONAMX") Then .HSXONAMX = fncNullCheck(rs("HSXONAMX"))         ' �i�r�w�_�f�Z�x���Ϗ��
            If fldNameExist("HSXOS1MN") Then .HSXOS1MN = fncNullCheck(rs("HSXOS1MN"))         ' �i�r�w�_�f�͏o�P����
            If fldNameExist("HSXOS1MX") Then .HSXOS1MX = fncNullCheck(rs("HSXOS1MX"))         ' �i�r�w�_�f�͏o�P���
            If fldNameExist("HSXOS1NS") Then .HSXOS1NS = rs("HSXOS1NS")         ' �i�r�w�_�f�͏o�P�M�����@
            If fldNameExist("HSXOS1SH") Then .HSXOS1SH = rs("HSXOS1SH")         ' �i�r�w�_�f�͏o�P����ʒu�Q��
            If fldNameExist("HSXOS1ST") Then .HSXOS1ST = rs("HSXOS1ST")         ' �i�r�w�_�f�͏o�P����ʒu�Q�_
            If fldNameExist("HSXOS1SI") Then .HSXOS1SI = rs("HSXOS1SI")         ' �i�r�w�_�f�͏o�P����ʒu�Q��
            If fldNameExist("HSXOS1HT") Then .HSXOS1HT = rs("HSXOS1HT")         ' �i�r�w�_�f�͏o�P�ۏؕ��@�Q��
            If fldNameExist("HSXOS1HS") Then .HSXOS1HS = rs("HSXOS1HS")         ' �i�r�w�_�f�͏o�P�ۏؕ��@�Q��
            If fldNameExist("HSXOS1HM") Then .HSXOS1HM = rs("HSXOS1HM")         ' �i�r�w�_�f�͏o�P�����p�x�Q��
            If fldNameExist("HSXOS1KI") Then .HSXOS1KI = rs("HSXOS1KI")         ' �i�r�w�_�f�͏o�P�����p�x�Q��
            If fldNameExist("HSXOS1KH") Then .HSXOS1KH = rs("HSXOS1KH")         ' �i�r�w�_�f�͏o�P�����p�x�Q��
            If fldNameExist("HSXOS1KS") Then .HSXOS1KS = rs("HSXOS1KS")         ' �i�r�w�_�f�͏o�P�����p�x�Q��
            If fldNameExist("HSXOS2MN") Then .HSXOS2MN = fncNullCheck(rs("HSXOS2MN"))         ' �i�r�w�_�f�͏o�Q����
            If fldNameExist("HSXOS2MX") Then .HSXOS2MX = fncNullCheck(rs("HSXOS2MX"))         ' �i�r�w�_�f�͏o�Q���
            If fldNameExist("HSXOS2NS") Then .HSXOS2NS = rs("HSXOS2NS")         ' �i�r�w�_�f�͏o�Q�M�����@
            If fldNameExist("HSXOS2SH") Then .HSXOS2SH = rs("HSXOS2SH")         ' �i�r�w�_�f�͏o�Q����ʒu�Q��
            If fldNameExist("HSXOS2ST") Then .HSXOS2ST = rs("HSXOS2ST")         ' �i�r�w�_�f�͏o�Q����ʒu�Q�_
            If fldNameExist("HSXOS2SI") Then .HSXOS2SI = rs("HSXOS2SI")         ' �i�r�w�_�f�͏o�Q����ʒu�Q��
            If fldNameExist("HSXOS2HT") Then .HSXOS2HT = rs("HSXOS2HT")         ' �i�r�w�_�f�͏o�Q�ۏؕ��@�Q��
            If fldNameExist("HSXOS2HS") Then .HSXOS2HS = rs("HSXOS2HS")         ' �i�r�w�_�f�͏o�Q�ۏؕ��@�Q��
            If fldNameExist("HSXOS2KM") Then .HSXOS2KM = rs("HSXOS2KM")         ' �i�r�w�_�f�͏o�Q�����p�x�Q��
            If fldNameExist("HSXOS2KN") Then .HSXOS2KN = rs("HSXOS2KN")         ' �i�r�w�_�f�͏o�Q�����p�x�Q��
            If fldNameExist("HSXOS2KH") Then .HSXOS2KH = rs("HSXOS2KH")         ' �i�r�w�_�f�͏o�Q�����p�x�Q��
            If fldNameExist("HSXOS2KU") Then .HSXOS2KU = rs("HSXOS2KU")         ' �i�r�w�_�f�͏o�Q�����p�x�Q��
            If fldNameExist("IFKBN") Then .IFKBN = rs("IFKBN")                  ' �h�^�e�敪
            If fldNameExist("SYORIKBN") Then .SYORIKBN = rs("SYORIKBN")         ' �����敪
            If fldNameExist("SPECRRNO") Then .SPECRRNO = rs("SPECRRNO")         ' �d�l�o�^�˗��ԍ�
            If fldNameExist("SXLMCNO") Then .SXLMCNO = rs("SXLMCNO")            ' �r�w�k��������ԍ�
            If fldNameExist("WFMCNO") Then .WFMCNO = rs("WFMCNO")               ' �v�e��������ԍ�
            If fldNameExist("STAFFID") Then .StaffID = rs("STAFFID")            ' �Ј�ID
            If fldNameExist("REGDATE") Then .REGDATE = rs("REGDATE")            ' �o�^���t
            If fldNameExist("UPDDATE") Then .UPDDATE = rs("UPDDATE")            ' �X�V���t
            If fldNameExist("SENDFLAG") Then .SENDFLAG = rs("SENDFLAG")         ' ���M�t���O
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME019 = FUNCTION_RETURN_SUCCESS
  

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�e�[�u���uTBCME020�v��������ɂ��������R�[�h�𒊏o����
'          :records()     ,O  ,typ_TBCME020 ,���o���R�[�h
'          :formID        ,I  ,String       ,�g�p�t�H�[��ID
'          :sqlOrder      ,I  ,tFullHinban  ,���o�i�ԁi�z��j
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :�����L�[�́A�HINBAN�+�uMNOREVNO�v+�uFACTORY�v+�uOPECOND�v�̕�����Ƃ���
'����      :2001/06/27�쐬�@����

Public Function DBDRV_GetTBCME020(records() As typ_TBCME020, formID$, hin() As tFullHinban) As FUNCTION_RETURN
Dim sql         As String           'SQL�S��
Dim sqlBase     As String           'SQL��{��(WHERE�߂̑O�܂�)
Dim sqlWhere    As String           'SQLWhere��
Dim rs          As OraDynaset       'RecordSet
Dim recCnt      As Long             '���R�[�h��
Dim key         As String           '����KEY
Dim i           As Long             'ٰ�߶���
Dim j           As Long             'ٰ�߶���2


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_TBCME020_SQL.bas -- Function DBDRV_GetTBCME020"

   Select Case formID
        Case "f_cmbc021_1"           '�uFTIR(Oi,Cs)���ѓ��́v
             sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXDENKU, HSXDENMX, HSXDENMN," & _
              " HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXLDLKU, HSXLDLMX, HSXLDLMN, HSXLDLHT," & _
              " HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH, HSXGDKHS, HSXDSOKE," & _
              " HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS, HSXLIFTW," & _
              " HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH," & _
              " HSXOF1KS, HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ," & _
              " HSXOF2KM, HSXOF2KI, HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR,"
            sqlBase = sqlBase & " HSXOF3HT, HSXOF3HS, HSXOF3SZ, HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX," & _
              " HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS," & _
              " HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI," & _
              " HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS," & _
              " HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET, HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST," & _
              " HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS, HSXBM3NS, HSXBM3ET, HSXNOTE,"
              For i = 1 To 10
                sqlBase = sqlBase & "HSXRS" & i & "N, "
                sqlBase = sqlBase & "HSXRS" & i & "Y, "
              Next
              sqlBase = sqlBase & "SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc022_1"           '�uGFA(Oi)���ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXDENKU, HSXDENMX, HSXDENMN," & _
              " HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXLDLKU, HSXLDLMX, HSXLDLMN, HSXLDLHT," & _
              " HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH, HSXGDKHS, HSXDSOKE," & _
              " HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS, HSXLIFTW," & _
              " HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH," & _
              " HSXOF1KS, HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ," & _
              " HSXOF2KM, HSXOF2KI, HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR,"
            sqlBase = sqlBase & " HSXOF3HT, HSXOF3HS, HSXOF3SZ, HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX," & _
              " HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS," & _
              " HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI," & _
              " HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS," & _
              " HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET, HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST," & _
              " HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS, HSXBM3NS, HSXBM3ET, HSXNOTE,"
            For i = 1 To 10
                sqlBase = sqlBase & "HSXRS" & i & "N, "
                sqlBase = sqlBase & "HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc023_1"           '�u��R���ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXDENKU, HSXDENMX, HSXDENMN," & _
              " HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXLDLKU, HSXLDLMX, HSXLDLMN, HSXLDLHT," & _
              " HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH, HSXGDKHS, HSXDSOKE," & _
              " HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS, HSXLIFTW," & _
              " HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH," & _
              " HSXOF1KS, HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ," & _
              " HSXOF2KM, HSXOF2KI, HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR,"
            sqlBase = sqlBase & " HSXOF3HT, HSXOF3HS, HSXOF3SZ, HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX," & _
              " HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS," & _
              " HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI," & _
              " HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS," & _
              " HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET, HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST," & _
              " HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS, HSXBM3NS, HSXBM3ET, HSXNOTE,"
            For i = 1 To 10
                sqlBase = sqlBase & "HSXRS" & i & "N, "
                sqlBase = sqlBase & "HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
              
        Case "f_cmbc024_1"           '�uBMD���ѓ��́v
             sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXDENKU, HSXDENMX, HSXDENMN," & _
              " HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXLDLKU, HSXLDLMX, HSXLDLMN, HSXLDLHT," & _
              " HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH, HSXGDKHS, HSXDSOKE," & _
              " HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS, HSXLIFTW," & _
              " HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH," & _
              " HSXOF1KS, HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ," & _
              " HSXOF2KM, HSXOF2KI, HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR,"
            sqlBase = sqlBase & " HSXOF3HT, HSXOF3HS, HSXOF3SZ, HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX," & _
              " HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS," & _
              " HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI," & _
              " HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS," & _
              " HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET, HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST," & _
              " HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS, HSXBM3NS, HSXBM3ET, HSXNOTE," & _
              " HSXBMD1MBP, HSXBMD2MBP, HSXBMD3MBP,"
' OSF�CBMD���ڒǉ��Ή�  ���@1�s���@2002.04.02 yakimura
            For i = 1 To 10
                sqlBase = sqlBase & "HSXRS" & i & "N, "
                sqlBase = sqlBase & "HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmec030_1"           '�uBMD���ѓ��́v
             sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXDENKU, HSXDENMX, HSXDENMN," & _
              " HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXLDLKU, HSXLDLMX, HSXLDLMN, HSXLDLHT," & _
              " HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH, HSXGDKHS, HSXDSOKE," & _
              " HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS, HSXLIFTW," & _
              " HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH," & _
              " HSXOF1KS, HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ," & _
              " HSXOF2KM, HSXOF2KI, HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR,"
            sqlBase = sqlBase & " HSXOF3HT, HSXOF3HS, HSXOF3SZ, HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX," & _
              " HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS," & _
              " HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI," & _
              " HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS," & _
              " HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET, HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST," & _
              " HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS, HSXBM3NS, HSXBM3ET, HSXNOTE," & _
              " HSXBMD1MBP, HSXBMD2MBP, HSXBMD3MBP,"
' OSF�CBMD���ڒǉ��Ή�  ���@1�s���@2002.04.02 yakimura
            For i = 1 To 10
                sqlBase = sqlBase & "HSXRS" & i & "N, "
                sqlBase = sqlBase & "HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc025_1"           '�uOSF���ѓ��́v
             sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXDENKU, HSXDENMX, HSXDENMN," & _
              " HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXLDLKU, HSXLDLMX, HSXLDLMN, HSXLDLHT," & _
              " HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH, HSXGDKHS, HSXDSOKE," & _
              " HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS, HSXLIFTW," & _
              " HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH," & _
              " HSXOF1KS, HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ," & _
              " HSXOF2KM, HSXOF2KI, HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR,"
            sqlBase = sqlBase & " HSXOF3HT, HSXOF3HS, HSXOF3SZ, HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX," & _
              " HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS," & _
              " HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI," & _
              " HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS," & _
              " HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET, HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST," & _
              " HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS, HSXBM3NS, HSXBM3ET, HSXNOTE," & _
              " HSXOSF1PTK, HSXOSF2PTK, HSXOSF3PTK, HSXOSF4PTK,"
' OSF�CBMD���ڒǉ��Ή�  ���@1�s���@2002.04.02 yakimura
            For i = 1 To 10
                sqlBase = sqlBase & "HSXRS" & i & "N, "
                sqlBase = sqlBase & "HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmec031_1"           '�uOSF���ѓ��́v
             sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXDENKU, HSXDENMX, HSXDENMN," & _
              " HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXLDLKU, HSXLDLMX, HSXLDLMN, HSXLDLHT," & _
              " HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH, HSXGDKHS, HSXDSOKE," & _
              " HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS, HSXLIFTW," & _
              " HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH," & _
              " HSXOF1KS, HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ," & _
              " HSXOF2KM, HSXOF2KI, HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR,"
            sqlBase = sqlBase & " HSXOF3HT, HSXOF3HS, HSXOF3SZ, HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX," & _
              " HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS," & _
              " HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI," & _
              " HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS," & _
              " HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET, HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST," & _
              " HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS, HSXBM3NS, HSXBM3ET, HSXNOTE," & _
              " HSXOSF1PTK, HSXOSF2PTK, HSXOSF3PTK, HSXOSF4PTK,"
' OSF�CBMD���ڒǉ��Ή�  ���@1�s���@2002.04.02 yakimura
            For i = 1 To 10
                sqlBase = sqlBase & "HSXRS" & i & "N, "
                sqlBase = sqlBase & "HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc026_1"           '�uGD���ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXDENKU, HSXDENMX, HSXDENMN," & _
              " HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXLDLKU, HSXLDLMX, HSXLDLMN, HSXLDLHT," & _
              " HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH, HSXGDKHS, HSXDSOKE," & _
              " HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS, HSXLIFTW," & _
              " HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH," & _
              " HSXOF1KS, HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ," & _
              " HSXOF2KM, HSXOF2KI, HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR,"
            sqlBase = sqlBase & " HSXOF3HT, HSXOF3HS, HSXOF3SZ, HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX," & _
              " HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS," & _
              " HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI," & _
              " HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS," & _
              " HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET, HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST," & _
              " HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS, HSXBM3NS, HSXBM3ET, HSXNOTE,"
            For i = 1 To 10
                sqlBase = sqlBase & "HSXRS" & i & "N, "
                sqlBase = sqlBase & "HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc027_1"           '�u���C�t�^�C�����ѓ��́v
             sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXDENKU, HSXDENMX, HSXDENMN," & _
              " HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXLDLKU, HSXLDLMX, HSXLDLMN, HSXLDLHT," & _
              " HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH, HSXGDKHS, HSXDSOKE," & _
              " HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS, HSXLIFTW," & _
              " HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH," & _
              " HSXOF1KS, HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ," & _
              " HSXOF2KM, HSXOF2KI, HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR,"
            sqlBase = sqlBase & " HSXOF3HT, HSXOF3HS, HSXOF3SZ, HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX," & _
              " HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS," & _
              " HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI," & _
              " HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS," & _
              " HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET, HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST," & _
              " HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS, HSXBM3NS, HSXBM3ET, HSXNOTE,"
            For i = 1 To 10
                sqlBase = sqlBase & "HSXRS" & i & "N, "
                sqlBase = sqlBase & "HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc028_1"           '�uFPD���ѓ��́v
             sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXDENKU, HSXDENMX, HSXDENMN," & _
              " HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXLDLKU, HSXLDLMX, HSXLDLMN, HSXLDLHT," & _
              " HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH, HSXGDKHS, HSXDSOKE," & _
              " HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS, HSXLIFTW," & _
              " HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH," & _
              " HSXOF1KS, HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ," & _
              " HSXOF2KM, HSXOF2KI, HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR,"
             sqlBase = sqlBase & " HSXOF3HT, HSXOF3HS, HSXOF3SZ, HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX," & _
              " HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS," & _
              " HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI," & _
              " HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS," & _
              " HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET, HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST," & _
              " HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS, HSXBM3NS, HSXBM3ET, HSXNOTE,"
            For i = 1 To 10
                sqlBase = sqlBase & "HSXRS" & i & "N, "
                sqlBase = sqlBase & "HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc029_1"           '�uGFA�Z�����ݒ�v
             sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXDENKU, HSXDENMX, HSXDENMN," & _
              " HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXLDLKU, HSXLDLMX, HSXLDLMN, HSXLDLHT," & _
              " HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH, HSXGDKHS, HSXDSOKE," & _
              " HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS, HSXLIFTW," & _
              " HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH," & _
              " HSXOF1KS, HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ," & _
              " HSXOF2KM, HSXOF2KI, HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR,"
            sqlBase = sqlBase & " HSXOF3HT, HSXOF3HS, HSXOF3SZ, HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX," & _
              " HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS," & _
              " HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI," & _
              " HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS," & _
              " HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET, HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST," & _
              " HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS, HSXBM3NS, HSXBM3ET, HSXNOTE,"
            For i = 1 To 10
                sqlBase = sqlBase & "HSXRS" & i & "N, "
                sqlBase = sqlBase & "HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    
    End Select
       
    sqlBase = sqlBase & "From TBCME020"
    
    '''SQL��Where���쐬
    For i = 0 To UBound(hin)
        With hin(i)
            key = key & "'" & .hinban & Format(.mnorevno, "00000") & .factory & .opecond & "'"
            If i <> UBound(hin) Then
                key = key & ", "
            End If
        End With
    Next
    sqlWhere = " Where(HINBAN||TO_CHAR(MNOREVNO, 'FM00000')||FACTORY||OPECOND in(" & key & "))"
    sql = sqlBase & sqlWhere
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME020 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''�t�B�[���h����o�^����
    fldCnt = rs.Fields.COUNT
    ReDim fldNames(fldCnt)
    For i = 1 To fldCnt
        fldNames(i) = rs.FieldName(i - 1)
    Next
     
    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            If fldNameExist("HINBAN") Then .hinban = rs("HINBAN")             ' �i��
            If fldNameExist("MNOREVNO") Then .mnorevno = rs("MNOREVNO")       ' ���i�ԍ������ԍ�
            If fldNameExist("FACTORY") Then .factory = rs("FACTORY")          ' �H��
            If fldNameExist("OPECOND") Then .opecond = rs("OPECOND")          ' ���Ə���
            If fldNameExist("HMGSTRRNO") Then .HMGSTRRNO = rs("HMGSTRRNO")    ' �i�Ǘ��d�l�o�^�˗��ԍ�
            If fldNameExist("HMGSTFNO") Then .HMGSTFNO = rs("HMGSTFNO")       ' �i�Ǘ��Ј��m��
            If fldNameExist("HMGSXSNO") Then .HMGSXSNO = rs("HMGSXSNO")       ' �i�Ǘ��r�w���i�ԍ�
            If fldNameExist("HMGSXSNE") Then .HMGSXSNE = fncNullCheck(rs("HMGSXSNE"))  ' �i�Ǘ��r�w���i�ԍ��}��
            If fldNameExist("HSXDENKU") Then .HSXDENKU = rs("HSXDENKU")       ' �i�r�w�c���������L��
            If fldNameExist("HSXDENMX") Then .HSXDENMX = fncNullCheck(rs("HSXDENMX"))  ' �i�r�w�c�������
            If fldNameExist("HSXDENMN") Then .HSXDENMN = fncNullCheck(rs("HSXDENMN"))  ' �i�r�w�c��������
            If fldNameExist("HSXDENHT") Then .HSXDENHT = rs("HSXDENHT")       ' �i�r�w�c�����ۏؕ��@�Q��
            If fldNameExist("HSXDENHS") Then .HSXDENHS = rs("HSXDENHS")       ' �i�r�w�c�����ۏؕ��@�Q��
            If fldNameExist("HSXDVDKU") Then .HSXDVDKU = rs("HSXDVDKU")       ' �i�r�w�c�u�c�Q�����L��
            If fldNameExist("HSXDVDMXN") Then .HSXDVDMX = fncNullCheck(rs("HSXDVDMXN")) ' �i�r�w�c�u�c�Q���    �v�e�T���v�������ύX 2003.05.20 yakimura
            If fldNameExist("HSXDVDMNN") Then .HSXDVDMN = fncNullCheck(rs("HSXDVDMNN")) ' �i�r�w�c�u�c�Q����    �v�e�T���v�������ύX 2003.05.20 yakimura
            If fldNameExist("HSXDVDHT") Then .HSXDVDHT = rs("HSXDVDHT")       ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
            If fldNameExist("HSXDVDHS") Then .HSXDVDHS = rs("HSXDVDHS")       ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
            If fldNameExist("HSXLDLKU") Then .HSXLDLKU = rs("HSXLDLKU")       ' �i�r�w�k�^�c�k�����L��
            If fldNameExist("HSXLDLMX") Then .HSXLDLMX = fncNullCheck(rs("HSXLDLMX"))   ' �i�r�w�k�^�c�k���
            If fldNameExist("HSXLDLMN") Then .HSXLDLMN = fncNullCheck(rs("HSXLDLMN"))   ' �i�r�w�k�^�c�k����
            If fldNameExist("HSXLDLHT") Then .HSXLDLHT = rs("HSXLDLHT")       ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
            If fldNameExist("HSXLDLHS") Then .HSXLDLHS = rs("HSXLDLHS")       ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
            If fldNameExist("HSXGDSZY") Then .HSXGDSZY = rs("HSXGDSZY")       ' �i�r�w�f�c�������
            If fldNameExist("HSXGDSPH") Then .HSXGDSPH = rs("HSXGDSPH")       ' �i�r�w�f�c����ʒu�Q��
            If fldNameExist("HSXGDSPT") Then .HSXGDSPT = rs("HSXGDSPT")       ' �i�r�w�f�c����ʒu�Q�_
            If fldNameExist("HSXGDSPR") Then .HSXGDSPR = rs("HSXGDSPR")       ' �i�r�w�f�c����ʒu�Q��
            If fldNameExist("HSXGDZAR") Then .HSXGDZAR = fncNullCheck(rs("HSXGDZAR"))   ' �i�r�w�f�c���O�̈�
            If fldNameExist("HSXGDKHM") Then .HSXGDKHM = rs("HSXGDKHM")       ' �i�r�w�f�c�����p�x�Q��
            If fldNameExist("HSXGDKHI") Then .HSXGDKHI = rs("HSXGDKHI")       ' �i�r�w�f�c�����p�x�Q��
            If fldNameExist("HSXGDKHH") Then .HSXGDKHH = rs("HSXGDKHH")       ' �i�r�w�f�c�����p�x�Q��
            If fldNameExist("HSXGDKHS") Then .HSXGDKHS = rs("HSXGDKHS")       ' �i�r�w�f�c�����p�x�Q��
            If fldNameExist("HSXDSOKE") Then .HSXDSOKE = rs("HSXDSOKE")       ' �i�r�w�c�r�n�c����
            If fldNameExist("HSXDSOMX") Then .HSXDSOMX = fncNullCheck(rs("HSXDSOMX"))  ' �i�r�w�c�r�n�c���
            If fldNameExist("HSXDSOMN") Then .HSXDSOMN = fncNullCheck(rs("HSXDSOMN"))  ' �i�r�w�c�r�n�c����
            If fldNameExist("HSXDSOAX") Then .HSXDSOAX = fncNullCheck(rs("HSXDSOAX"))  ' �i�r�w�c�r�n�c�̈���
            If fldNameExist("HSXDSOAN") Then .HSXDSOAN = fncNullCheck(rs("HSXDSOAN"))  ' �i�r�w�c�r�n�c�̈扺��
            If fldNameExist("HSXDSOHT") Then .HSXDSOHT = rs("HSXDSOHT")       ' �i�r�w�c�r�n�c�ۏؕ��@�Q��
            If fldNameExist("HSXDSOHS") Then .HSXDSOHS = rs("HSXDSOHS")       ' �i�r�w�c�r�n�c�ۏؕ��@�Q��
            If fldNameExist("HSXDSOKM") Then .HSXDSOKM = rs("HSXDSOKM")       ' �i�r�w�c�r�n�c�����p�x�Q��
            If fldNameExist("HSXDSOKI") Then .HSXDSOKI = rs("HSXDSOKI")       ' �i�r�w�c�r�n�c�����p�x�Q��
            If fldNameExist("HSXDSOKH") Then .HSXDSOKH = rs("HSXDSOKH")       ' �i�r�w�c�r�n�c�����p�x�Q��
            If fldNameExist("HSXDSOKS") Then .HSXDSOKS = rs("HSXDSOKS")       ' �i�r�w�c�r�n�c�����p�x�Q��
            If fldNameExist("HSXLIFTW") Then .HSXLIFTW = rs("HSXLIFTW")       ' �i�r�w������@
            If fldNameExist("HSXSDSLP") Then .HSXSDSLP = rs("HSXSDSLP")       ' �i�r�w�V�[�h�X
            If fldNameExist("HSXGKKNO") Then .HSXGKKNO = rs("HSXGKKNO")       ' �i�r�w�O�ϋK�i�m��
            If fldNameExist("HSXCDOP") Then .HSXCDOP = rs("HSXCDOP")         ' �i�r�w�����h�[�v
            If fldNameExist("HSXCDOPN") Then .HSXCDOPN = fncNullCheck(rs("HSXCDOPN"))       ' �i�r�w�����h�[�v�Z�x
            If fldNameExist("HSXCDPNI") Then .HSXCDPNI = rs("HSXCDPNI")       ' �i�r�w�����h�[�v�Z�x�w��
            If fldNameExist("HSXGSFIN") Then .HSXGSFIN = rs("HSXGSFIN")       ' �i�r�w�O���d�グ
            If fldNameExist("HSXCLMIN") Then .HSXCLMIN = fncNullCheck(rs("HSXCLMIN"))  ' �i�r�w����������
            If fldNameExist("HSXCLMAX") Then .HSXCLMAX = fncNullCheck(rs("HSXCLMAX"))  ' �i�r�w���������
            If fldNameExist("HSXCLPMN") Then .HSXCLPMN = fncNullCheck(rs("HSXCLPMN"))  ' �i�r�w���������e����
            If fldNameExist("HSXCLPR") Then .HSXCLPR = fncNullCheck(rs("HSXCLPR"))     ' �i�r�w���������e�䗦
            If fldNameExist("HSXWFWAR") Then .HSXWFWAR = rs("HSXWFWAR")       ' �i�r�w�v�e�v�����������N
#If False Then  '�e�[�u���̌^��`��s_cmzcTableDefs.bas�ňقȂ邽�߂̑Ή�
            For j = 1 To 4
                If fldNameExist("HSXOF" & j & "AX") Then .HSXOF_AX(j) = fncNullCheck(rs("HSXOF" & j & "AX"))  ' �i�r�w�n�r�e(n)���Ϗ��
                If fldNameExist("HSXOF" & j & "MX") Then .HSXOF_MX(j) = fncNullCheck(rs("HSXOF" & j & "MX"))  ' �i�r�w�n�r�e(n)���
                If fldNameExist("HSXOF" & j & "SH") Then .HSXOF_SH(j) = rs("HSXOF" & j & "SH")  ' �i�r�w�n�r�e(n)����ʒu�Q��
                If fldNameExist("HSXOF" & j & "ST") Then .HSXOF_ST(j) = rs("HSXOF" & j & "ST")  ' �i�r�w�n�r�e(n)����ʒu�Q�_
                If fldNameExist("HSXOF" & j & "SR") Then .HSXOF_SR(j) = rs("HSXOF" & j & "SR")  ' �i�r�w�n�r�e(n)����ʒu�Q��
                If fldNameExist("HSXOF" & j & "HT") Then .HSXOF_HT(j) = rs("HSXOF" & j & "HT")  ' �i�r�w�n�r�e(n)�ۏؕ��@�Q��
                If fldNameExist("HSXOF" & j & "HS") Then .HSXOF_HS(j) = rs("HSXOF" & j & "HS")  ' �i�r�w�n�r�e(n)�ۏؕ��@�Q��
                If fldNameExist("HSXOF" & j & "SZ") Then .HSXOF_SZ(j) = rs("HSXOF" & j & "SZ")  ' �i�r�w�n�r�e(n)�������
                If fldNameExist("HSXOF" & j & "KM") Then .HSXOF_KM(j) = rs("HSXOF" & j & "KM")  ' �i�r�w�n�r�e(n)�����p�x�Q��
                If fldNameExist("HSXOF" & j & "KI") Then .HSXOF_KI(j) = rs("HSXOF" & j & "KI")  ' �i�r�w�n�r�e(n)�����p�x�Q��
                If fldNameExist("HSXOF" & j & "KH") Then .HSXOF_KH(j) = rs("HSXOF" & j & "KH")  ' �i�r�w�n�r�e(n)�����p�x�Q��
                If fldNameExist("HSXOF" & j & "KS") Then .HSXOF_KS(j) = rs("HSXOF" & j & "KS")  ' �i�r�w�n�r�e(n)�����p�x�Q��
                If fldNameExist("HSXOF" & j & "NS") Then .HSXOF_NS(j) = rs("HSXOF" & j & "NS")  ' �i�r�w�n�r�e(n)�M�����@
                If fldNameExist("HSXOF" & j & "ET") Then .HSXOF_ET(j) = fncNullCheck(rs("HSXOF" & j & "ET"))  ' �i�r�w�n�r�e(n)�I���d�s��
                'NULL�Ή�
                If fldNameExist("HSXOSF" & j & "PTK") Then                       ' �i�r�w�n�r�e(n)�p�^���敪
                   If IsNull(rs("HSXOSF" & j & "PTK")) = False Then .HSXOSF_PTK(j) = rs("HSXOSF" & j & "PTK")
                End If
            Next
            For j = 1 To 3
                If fldNameExist("HSXBM" & j & "AN") Then .HSXBM_AN(j) = fncNullCheck(rs("HSXBM" & j & "AN"))  ' �i�r�w�a�l�c(n)���ω���
                If fldNameExist("HSXBM" & j & "AX") Then .HSXBM_AX(j) = fncNullCheck(rs("HSXBM" & j & "AX"))  ' �i�r�w�a�l�c(n)���Ϗ��
                If fldNameExist("HSXBM" & j & "SH") Then .HSXBM_SH(j) = rs("HSXBM" & j & "SH")  ' �i�r�w�a�l�c(n)����ʒu�Q��
                If fldNameExist("HSXBM" & j & "ST") Then .HSXBM_ST(j) = rs("HSXBM" & j & "ST")  ' �i�r�w�a�l�c(n)����ʒu�Q�_
                If fldNameExist("HSXBM" & j & "SR") Then .HSXBM_SR(j) = rs("HSXBM" & j & "SR")  ' �i�r�w�a�l�c(n)����ʒu�Q��
                If fldNameExist("HSXBM" & j & "HT") Then .HSXBM_HT(j) = rs("HSXBM" & j & "HT")  ' �i�r�w�a�l�c(n)�ۏؕ��@�Q��
                If fldNameExist("HSXBM" & j & "HS") Then .HSXBM_HS(j) = rs("HSXBM" & j & "HS")  ' �i�r�w�a�l�c(n)�ۏؕ��@�Q��
                If fldNameExist("HSXBM" & j & "SZ") Then .HSXBM_SZ(j) = rs("HSXBM" & j & "SZ")  ' �i�r�w�a�l�c(n)�������
                If fldNameExist("HSXBM" & j & "KM") Then .HSXBM_KM(j) = rs("HSXBM" & j & "KM")  ' �i�r�w�a�l�c(n)�����p�x�Q��
                If fldNameExist("HSXBM" & j & "KI") Then .HSXBM_KI(j) = rs("HSXBM" & j & "KI")  ' �i�r�w�a�l�c(n)�����p�x�Q��
                If fldNameExist("HSXBM" & j & "KH") Then .HSXBM_KH(j) = rs("HSXBM" & j & "KH")  ' �i�r�w�a�l�c(n)�����p�x�Q��
                If fldNameExist("HSXBM" & j & "KS") Then .HSXBM_KS(j) = rs("HSXBM" & j & "KS")  ' �i�r�w�a�l�c(n)�����p�x�Q��
                If fldNameExist("HSXBM" & j & "NS") Then .HSXBM_NS(j) = rs("HSXBM" & j & "NS")  ' �i�r�w�a�l�c(n)�M�����@
                If fldNameExist("HSXBM" & j & "ET") Then .HSXBM_ET(j) = fncNullCheck(rs("HSXBM" & j & "ET"))  ' �i�r�w�a�l�c(n)�I���d�s��
                'NULL�Ή�
                If fldNameExist("HSXBMD" & j & "MBP") Then                      ' �i�r�w�a�l�c(n)�ʓ����z
                   If IsNull(rs("HSXBMD" & j & "MBP")) = False Then .HSXBMD_MBP(j) = fncNullCheck(rs("HSXBMD" & j & "MBP"))
                End If
            Next
#Else
                If fldNameExist("HSXOF1AX") Then .HSXOF1AX = fncNullCheck(rs("HSXOF1AX"))  ' �i�r�w�n�r�e1���Ϗ��
                If fldNameExist("HSXOF1MX") Then .HSXOF1MX = fncNullCheck(rs("HSXOF1MX"))  ' �i�r�w�n�r�e1���
                If fldNameExist("HSXOF1SH") Then .HSXOF1SH = rs("HSXOF1SH")  ' �i�r�w�n�r�e1����ʒu�Q��
                If fldNameExist("HSXOF1ST") Then .HSXOF1ST = rs("HSXOF1ST")  ' �i�r�w�n�r�e1����ʒu�Q�_
                If fldNameExist("HSXOF1SR") Then .HSXOF1SR = rs("HSXOF1SR")  ' �i�r�w�n�r�e1����ʒu�Q��
                If fldNameExist("HSXOF1HT") Then .HSXOF1HT = rs("HSXOF1HT")  ' �i�r�w�n�r�e1�ۏؕ��@�Q��
                If fldNameExist("HSXOF1HS") Then .HSXOF1HS = rs("HSXOF1HS")  ' �i�r�w�n�r�e1�ۏؕ��@�Q��
                If fldNameExist("HSXOF1SZ") Then .HSXOF1SZ = rs("HSXOF1SZ")  ' �i�r�w�n�r�e1�������
                If fldNameExist("HSXOF1KM") Then .HSXOF1KM = rs("HSXOF1KM")  ' �i�r�w�n�r�e1�����p�x�Q��
                If fldNameExist("HSXOF1KI") Then .HSXOF1KI = rs("HSXOF1KI")  ' �i�r�w�n�r�e1�����p�x�Q��
                If fldNameExist("HSXOF1KH") Then .HSXOF1KH = rs("HSXOF1KH")  ' �i�r�w�n�r�e1�����p�x�Q��
                If fldNameExist("HSXOF1KS") Then .HSXOF1KS = rs("HSXOF1KS")  ' �i�r�w�n�r�e1�����p�x�Q��
                If fldNameExist("HSXOF1NS") Then .HSXOF1NS = rs("HSXOF1NS")  ' �i�r�w�n�r�e1�M�����@
                If fldNameExist("HSXOF1ET") Then .HSXOF1ET = fncNullCheck(rs("HSXOF1ET"))  ' �i�r�w�n�r�e1�I���d�s��
                If fldNameExist("HSXOSF1PTK") Then                           ' �i�r�w�n�r�e1�p�^���敪
                   If IsNull(rs("HSXOSF1PTK")) = False Then .HSXOSF1PTK = rs("HSXOSF1PTK")
                End If
                If fldNameExist("HSXOF2AX") Then .HSXOF2AX = fncNullCheck(rs("HSXOF2AX"))  ' �i�r�w�n�r�e2���Ϗ��
                If fldNameExist("HSXOF2MX") Then .HSXOF2MX = fncNullCheck(rs("HSXOF2MX"))  ' �i�r�w�n�r�e2���
                If fldNameExist("HSXOF2SH") Then .HSXOF2SH = rs("HSXOF2SH")  ' �i�r�w�n�r�e2����ʒu�Q��
                If fldNameExist("HSXOF2ST") Then .HSXOF2ST = rs("HSXOF2ST")  ' �i�r�w�n�r�e2����ʒu�Q�_
                If fldNameExist("HSXOF2SR") Then .HSXOF2SR = rs("HSXOF2SR")  ' �i�r�w�n�r�e2����ʒu�Q��
                If fldNameExist("HSXOF2HT") Then .HSXOF2HT = rs("HSXOF2HT")  ' �i�r�w�n�r�e2�ۏؕ��@�Q��
                If fldNameExist("HSXOF2HS") Then .HSXOF2HS = rs("HSXOF2HS")  ' �i�r�w�n�r�e2�ۏؕ��@�Q��
                If fldNameExist("HSXOF2SZ") Then .HSXOF2SZ = rs("HSXOF2SZ")  ' �i�r�w�n�r�e2�������
                If fldNameExist("HSXOF2KM") Then .HSXOF2KM = rs("HSXOF2KM")  ' �i�r�w�n�r�e2�����p�x�Q��
                If fldNameExist("HSXOF2KI") Then .HSXOF2KI = rs("HSXOF2KI")  ' �i�r�w�n�r�e2�����p�x�Q��
                If fldNameExist("HSXOF2KH") Then .HSXOF2KH = rs("HSXOF2KH")  ' �i�r�w�n�r�e2�����p�x�Q��
                If fldNameExist("HSXOF2KS") Then .HSXOF2KS = rs("HSXOF2KS")  ' �i�r�w�n�r�e2�����p�x�Q��
                If fldNameExist("HSXOF2NS") Then .HSXOF2NS = rs("HSXOF2NS")  ' �i�r�w�n�r�e2�M�����@
                If fldNameExist("HSXOF2ET") Then .HSXOF2ET = fncNullCheck(rs("HSXOF2ET"))  ' �i�r�w�n�r�e2�I���d�s��
                If fldNameExist("HSXOSF2PTK") Then                           ' �i�r�w�n�r�e2�p�^���敪
                   If IsNull(rs("HSXOSF2PTK")) = False Then .HSXOSF2PTK = rs("HSXOSF2PTK")
                End If
                If fldNameExist("HSXOF3AX") Then .HSXOF3AX = fncNullCheck(rs("HSXOF3AX"))  ' �i�r�w�n�r�e3���Ϗ��
                If fldNameExist("HSXOF3MX") Then .HSXOF3MX = fncNullCheck(rs("HSXOF3MX"))  ' �i�r�w�n�r�e3���
                If fldNameExist("HSXOF3SH") Then .HSXOF3SH = rs("HSXOF3SH")  ' �i�r�w�n�r�e3����ʒu�Q��
                If fldNameExist("HSXOF3ST") Then .HSXOF3ST = rs("HSXOF3ST")  ' �i�r�w�n�r�e3����ʒu�Q�_
                If fldNameExist("HSXOF3SR") Then .HSXOF3SR = rs("HSXOF3SR")  ' �i�r�w�n�r�e3����ʒu�Q��
                If fldNameExist("HSXOF3HT") Then .HSXOF3HT = rs("HSXOF3HT")  ' �i�r�w�n�r�e3�ۏؕ��@�Q��
                If fldNameExist("HSXOF3HS") Then .HSXOF3HS = rs("HSXOF3HS")  ' �i�r�w�n�r�e3�ۏؕ��@�Q��
                If fldNameExist("HSXOF3SZ") Then .HSXOF3SZ = rs("HSXOF3SZ")  ' �i�r�w�n�r�e3�������
                If fldNameExist("HSXOF3KM") Then .HSXOF3KM = rs("HSXOF3KM")  ' �i�r�w�n�r�e3�����p�x�Q��
                If fldNameExist("HSXOF3KI") Then .HSXOF3KI = rs("HSXOF3KI")  ' �i�r�w�n�r�e3�����p�x�Q��
                If fldNameExist("HSXOF3KH") Then .HSXOF3KH = rs("HSXOF3KH")  ' �i�r�w�n�r�e3�����p�x�Q��
                If fldNameExist("HSXOF3KS") Then .HSXOF3KS = rs("HSXOF3KS")  ' �i�r�w�n�r�e3�����p�x�Q��
                If fldNameExist("HSXOF3NS") Then .HSXOF3NS = rs("HSXOF3NS")  ' �i�r�w�n�r�e3�M�����@
                If fldNameExist("HSXOF3ET") Then .HSXOF3ET = fncNullCheck(rs("HSXOF3ET"))  ' �i�r�w�n�r�e3�I���d�s��
                If fldNameExist("HSXOSF3PTK") Then                           ' �i�r�w�n�r�e3�p�^���敪
                   If IsNull(rs("HSXOSF3PTK")) = False Then .HSXOSF3PTK = rs("HSXOSF3PTK")
                End If
                If fldNameExist("HSXOF4AX") Then .HSXOF4AX = fncNullCheck(rs("HSXOF4AX"))  ' �i�r�w�n�r�e4���Ϗ��
                If fldNameExist("HSXOF4MX") Then .HSXOF4MX = fncNullCheck(rs("HSXOF4MX"))  ' �i�r�w�n�r�e4���
                If fldNameExist("HSXOF4SH") Then .HSXOF4SH = rs("HSXOF4SH")  ' �i�r�w�n�r�e4����ʒu�Q��
                If fldNameExist("HSXOF4ST") Then .HSXOF4ST = rs("HSXOF4ST")  ' �i�r�w�n�r�e4����ʒu�Q�_
                If fldNameExist("HSXOF4SR") Then .HSXOF4SR = rs("HSXOF4SR")  ' �i�r�w�n�r�e4����ʒu�Q��
                If fldNameExist("HSXOF4HT") Then .HSXOF4HT = rs("HSXOF4HT")  ' �i�r�w�n�r�e4�ۏؕ��@�Q��
                If fldNameExist("HSXOF4HS") Then .HSXOF4HS = rs("HSXOF4HS")  ' �i�r�w�n�r�e4�ۏؕ��@�Q��
                If fldNameExist("HSXOF4SZ") Then .HSXOF4SZ = rs("HSXOF4SZ")  ' �i�r�w�n�r�e4�������
                If fldNameExist("HSXOF4KM") Then .HSXOF4KM = rs("HSXOF4KM")  ' �i�r�w�n�r�e4�����p�x�Q��
                If fldNameExist("HSXOF4KI") Then .HSXOF4KI = rs("HSXOF4KI")  ' �i�r�w�n�r�e4�����p�x�Q��
                If fldNameExist("HSXOF4KH") Then .HSXOF4KH = rs("HSXOF4KH")  ' �i�r�w�n�r�e4�����p�x�Q��
                If fldNameExist("HSXOF4KS") Then .HSXOF4KS = rs("HSXOF4KS")  ' �i�r�w�n�r�e4�����p�x�Q��
                If fldNameExist("HSXOF4NS") Then .HSXOF4NS = rs("HSXOF4NS")  ' �i�r�w�n�r�e4�M�����@
                If fldNameExist("HSXOF4ET") Then .HSXOF4ET = fncNullCheck(rs("HSXOF4ET"))  ' �i�r�w�n�r�e4�I���d�s��
                If fldNameExist("HSXOSF4PTK") Then                           ' �i�r�w�n�r�e4�p�^���敪
                   If IsNull(rs("HSXOSF4PTK")) = False Then .HSXOSF4PTK = rs("HSXOSF4PTK")
                End If
                If fldNameExist("HSXBM1AN") Then .HSXBM1AN = fncNullCheck(rs("HSXBM1AN"))  ' �i�r�w�a�l�c1���ω���
                If fldNameExist("HSXBM1AX") Then .HSXBM1AX = fncNullCheck(rs("HSXBM1AX"))  ' �i�r�w�a�l�c1���Ϗ��
                If fldNameExist("HSXBM1SH") Then .HSXBM1SH = rs("HSXBM1SH")  ' �i�r�w�a�l�c1����ʒu�Q��
                If fldNameExist("HSXBM1ST") Then .HSXBM1ST = rs("HSXBM1ST")  ' �i�r�w�a�l�c1����ʒu�Q�_
                If fldNameExist("HSXBM1SR") Then .HSXBM1SR = rs("HSXBM1SR")  ' �i�r�w�a�l�c1����ʒu�Q��
                If fldNameExist("HSXBM1HT") Then .HSXBM1HT = rs("HSXBM1HT")  ' �i�r�w�a�l�c1�ۏؕ��@�Q��
                If fldNameExist("HSXBM1HS") Then .HSXBM1HS = rs("HSXBM1HS")  ' �i�r�w�a�l�c1�ۏؕ��@�Q��
                If fldNameExist("HSXBM1SZ") Then .HSXBM1SZ = rs("HSXBM1SZ")  ' �i�r�w�a�l�c1�������
                If fldNameExist("HSXBM1KM") Then .HSXBM1KM = rs("HSXBM1KM")  ' �i�r�w�a�l�c1�����p�x�Q��
                If fldNameExist("HSXBM1KI") Then .HSXBM1KI = rs("HSXBM1KI")  ' �i�r�w�a�l�c1�����p�x�Q��
                If fldNameExist("HSXBM1KH") Then .HSXBM1KH = rs("HSXBM1KH")  ' �i�r�w�a�l�c1�����p�x�Q��
                If fldNameExist("HSXBM1KS") Then .HSXBM1KS = rs("HSXBM1KS")  ' �i�r�w�a�l�c1�����p�x�Q��
                If fldNameExist("HSXBM1NS") Then .HSXBM1NS = rs("HSXBM1NS")  ' �i�r�w�a�l�c1�M�����@
                If fldNameExist("HSXBM1ET") Then .HSXBM1ET = fncNullCheck(rs("HSXBM1ET"))  ' �i�r�w�a�l�c1�I���d�s��
                'NULL�Ή�
                If fldNameExist("HSXBMD1MBP") Then                           ' �i�r�w�a�l�c1�ʓ����z
                   If IsNull(rs("HSXBMD1MBP")) = False Then .HSXBMD1MBP = fncNullCheck(rs("HSXBMD1MBP"))
                End If
                If fldNameExist("HSXBM2AN") Then .HSXBM2AN = fncNullCheck(rs("HSXBM2AN"))  ' �i�r�w�a�l�c2���ω���
                If fldNameExist("HSXBM2AX") Then .HSXBM2AX = fncNullCheck(rs("HSXBM2AX"))  ' �i�r�w�a�l�c2���Ϗ��
                If fldNameExist("HSXBM2SH") Then .HSXBM2SH = rs("HSXBM2SH")  ' �i�r�w�a�l�c2����ʒu�Q��
                If fldNameExist("HSXBM2ST") Then .HSXBM2ST = rs("HSXBM2ST")  ' �i�r�w�a�l�c2����ʒu�Q�_
                If fldNameExist("HSXBM2SR") Then .HSXBM2SR = rs("HSXBM2SR")  ' �i�r�w�a�l�c2����ʒu�Q��
                If fldNameExist("HSXBM2HT") Then .HSXBM2HT = rs("HSXBM2HT")  ' �i�r�w�a�l�c2�ۏؕ��@�Q��
                If fldNameExist("HSXBM2HS") Then .HSXBM2HS = rs("HSXBM2HS")  ' �i�r�w�a�l�c2�ۏؕ��@�Q��
                If fldNameExist("HSXBM2SZ") Then .HSXBM2SZ = rs("HSXBM2SZ")  ' �i�r�w�a�l�c2�������
                If fldNameExist("HSXBM2KM") Then .HSXBM2KM = rs("HSXBM2KM")  ' �i�r�w�a�l�c2�����p�x�Q��
                If fldNameExist("HSXBM2KI") Then .HSXBM2KI = rs("HSXBM2KI")  ' �i�r�w�a�l�c2�����p�x�Q��
                If fldNameExist("HSXBM2KH") Then .HSXBM2KH = rs("HSXBM2KH")  ' �i�r�w�a�l�c2�����p�x�Q��
                If fldNameExist("HSXBM2KS") Then .HSXBM2KS = rs("HSXBM2KS")  ' �i�r�w�a�l�c2�����p�x�Q��
                If fldNameExist("HSXBM2NS") Then .HSXBM2NS = rs("HSXBM2NS")  ' �i�r�w�a�l�c2�M�����@
                If fldNameExist("HSXBM2ET") Then .HSXBM2ET = fncNullCheck(rs("HSXBM2ET"))  ' �i�r�w�a�l�c2�I���d�s��
                'NULL�Ή�
                If fldNameExist("HSXBMD2MBP") Then                           ' �i�r�w�a�l�c2�ʓ����z
                   If IsNull(rs("HSXBMD2MBP")) = False Then .HSXBMD2MBP = rs("HSXBMD2MBP")
                End If
                If fldNameExist("HSXBM3AN") Then .HSXBM3AN = fncNullCheck(rs("HSXBM3AN"))  ' �i�r�w�a�l�c3���ω���
                If fldNameExist("HSXBM3AX") Then .HSXBM3AX = fncNullCheck(rs("HSXBM3AX"))  ' �i�r�w�a�l�c3���Ϗ��
                If fldNameExist("HSXBM3SH") Then .HSXBM3SH = rs("HSXBM3SH")  ' �i�r�w�a�l�c3����ʒu�Q��
                If fldNameExist("HSXBM3ST") Then .HSXBM3ST = rs("HSXBM3ST")  ' �i�r�w�a�l�c3����ʒu�Q�_
                If fldNameExist("HSXBM3SR") Then .HSXBM3SR = rs("HSXBM3SR")  ' �i�r�w�a�l�c3����ʒu�Q��
                If fldNameExist("HSXBM3HT") Then .HSXBM3HT = rs("HSXBM3HT")  ' �i�r�w�a�l�c3�ۏؕ��@�Q��
                If fldNameExist("HSXBM3HS") Then .HSXBM3HS = rs("HSXBM3HS")  ' �i�r�w�a�l�c3�ۏؕ��@�Q��
                If fldNameExist("HSXBM3SZ") Then .HSXBM3SZ = rs("HSXBM3SZ")  ' �i�r�w�a�l�c3�������
                If fldNameExist("HSXBM3KM") Then .HSXBM3KM = rs("HSXBM3KM")  ' �i�r�w�a�l�c3�����p�x�Q��
                If fldNameExist("HSXBM3KI") Then .HSXBM3KI = rs("HSXBM3KI")  ' �i�r�w�a�l�c3�����p�x�Q��
                If fldNameExist("HSXBM3KH") Then .HSXBM3KH = rs("HSXBM3KH")  ' �i�r�w�a�l�c3�����p�x�Q��
                If fldNameExist("HSXBM3KS") Then .HSXBM3KS = rs("HSXBM3KS")  ' �i�r�w�a�l�c3�����p�x�Q��
                If fldNameExist("HSXBM3NS") Then .HSXBM3NS = rs("HSXBM3NS")  ' �i�r�w�a�l�c3�M�����@
                If fldNameExist("HSXBM3ET") Then .HSXBM3ET = fncNullCheck(rs("HSXBM3ET"))  ' �i�r�w�a�l�c3�I���d�s��
                'NULL�Ή�
                If fldNameExist("HSXBMD3MBP") Then                           ' �i�r�w�a�l�c3�ʓ����z
                    If IsNull(rs("HSXBMD3MBP")) = False Then .HSXBMD3MBP = fncNullCheck(rs("HSXBMD3MBP"))
                End If
#End If
            If fldNameExist("HSXNOTE") Then .HSXNOTE = rs("HSXNOTE")         ' �i�r�w���L
#If False Then  '�e�[�u���̌^��`��s_cmzcTableDefs.bas�ňႤ���ߖ����Ƃ���
            For j = 1 To 10
                If fldNameExist("HSXRS" & j & "N") Then .HSXRS_N(j) = rs("HSXRS" & j & "N")     ' �i�r�w�\��(n)�Q��
                If fldNameExist("HSXRS" & j & "Y") Then .HSXRS_Y(j) = rs("HSXRS" & j & "Y")     ' �i�r�w�\��(n)�Q�p
            Next
#Else
                If fldNameExist("HSXRS1N") Then .HSXRS1N = rs("HSXRS1N")     ' �i�r�w�\��1�Q��
                If fldNameExist("HSXRS2N") Then .HSXRS2N = rs("HSXRS2N")     ' �i�r�w�\��2�Q��
                If fldNameExist("HSXRS3N") Then .HSXRS3N = rs("HSXRS3N")     ' �i�r�w�\��3�Q��
                If fldNameExist("HSXRS4N") Then .HSXRS4N = rs("HSXRS4N")     ' �i�r�w�\��4�Q��
                If fldNameExist("HSXRS5N") Then .HSXRS5N = rs("HSXRS5N")     ' �i�r�w�\��5�Q��
                If fldNameExist("HSXRS6N") Then .HSXRS6N = rs("HSXRS6N")     ' �i�r�w�\��6�Q��
                If fldNameExist("HSXRS7N") Then .HSXRS7N = rs("HSXRS7N")     ' �i�r�w�\��7�Q��
                If fldNameExist("HSXRS8N") Then .HSXRS8N = rs("HSXRS8N")     ' �i�r�w�\��8�Q��
                If fldNameExist("HSXRS9N") Then .HSXRS9N = rs("HSXRS9N")     ' �i�r�w�\��9�Q��
                If fldNameExist("HSXRS10N") Then .HSXRS10N = rs("HSXRS10N")  ' �i�r�w�\��10�Q��
                If fldNameExist("HSXRS1Y") Then .HSXRS1Y = rs("HSXRS1Y")     ' �i�r�w�\��1�Q�p
                If fldNameExist("HSXRS2Y") Then .HSXRS2Y = rs("HSXRS2Y")     ' �i�r�w�\��2�Q�p
                If fldNameExist("HSXRS3Y") Then .HSXRS3Y = rs("HSXRS3Y")     ' �i�r�w�\��3�Q�p
                If fldNameExist("HSXRS4Y") Then .HSXRS4Y = rs("HSXRS4Y")     ' �i�r�w�\��4�Q�p
                If fldNameExist("HSXRS5Y") Then .HSXRS5Y = rs("HSXRS5Y")     ' �i�r�w�\��5�Q�p
                If fldNameExist("HSXRS6Y") Then .HSXRS6Y = rs("HSXRS6Y")     ' �i�r�w�\��6�Q�p
                If fldNameExist("HSXRS7Y") Then .HSXRS7Y = rs("HSXRS7Y")     ' �i�r�w�\��7�Q�p
                If fldNameExist("HSXRS8Y") Then .HSXRS8Y = rs("HSXRS8Y")     ' �i�r�w�\��8�Q�p
                If fldNameExist("HSXRS9Y") Then .HSXRS9Y = rs("HSXRS9Y")     ' �i�r�w�\��9�Q�p
                If fldNameExist("HSXRS1YN") Then .HSXRS10Y = rs("HSXRS10Y")     ' �i�r�w�\��10�Q�p
#End If
            If fldNameExist("SPECRRNO") Then .SPECRRNO = rs("SPECRRNO")     ' �d�l�o�^�˗��ԍ�
            If fldNameExist("SXLMCNO") Then .SXLMCNO = rs("SXLMCNO")        ' �r�w�k��������ԍ�
            If fldNameExist("WFMCNO") Then .WFMCNO = rs("WFMCNO")           ' �v�e��������ԍ�
            If fldNameExist("STAFFID") Then .StaffID = rs("STAFFID")        ' �Ј�ID
            If fldNameExist("REGDATE") Then .REGDATE = rs("REGDATE")        ' �o�^���t
            If fldNameExist("UPDDATE") Then .UPDDATE = rs("UPDDATE")        ' �X�V���t
            If fldNameExist("SENDFLAG") Then .SENDFLAG = rs("SENDFLAG")     ' ���M�t���O
            If fldNameExist("SENDDATE") Then .SENDDATE = rs("SENDDATE")     ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME020 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function
'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

Private Function fldNameExist(fldName As String) As Boolean
    Dim sql         As String           'SQL�S��
    Dim i As Integer                    'ٰ�߶���


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_TBCME***_SQL.bas -- Function fldNameExist"

    fldNameExist = False                '�װ�ð���i�����l�j���
    
    For i = 1 To fldCnt                 '̨���ސ���ٰ��
        If fldName = fldNames(i) Then   '������̨���ޖ��ƈ�v������̂��������ꍇ
            fldNameExist = True         '����ð�����
            Exit For                    'ٰ�߂𔲂���
        End If
    Next
    

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function
'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME018�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME018 ,���o���R�[�h
'          :formID        ,I  ,String       ,�g�p�t�H�[��ID
'          :sqlOrder      ,I  ,tFullHinban  ,���o�i�ԁi�z��j
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :�����L�[�́A�HINBAN�+�uMNOREVNO�v+�uFACTORY�v+�uOPECOND�v�̕�����Ƃ���
'����      :2001/06/27�쐬�@����

Public Function DBDRV_GetTBCME018(records() As typ_TBCME018, formID$, hin() As tFullHinban) As FUNCTION_RETURN
Dim sql         As String           'SQL�S��
Dim sqlBase     As String           'SQL��{��(WHERE�߂̑O�܂�)
Dim sqlWhere    As String           'SQLWhere��
Dim rs          As OraDynaset       'RecordSet
Dim recCnt      As Long             '���R�[�h��
Dim key         As String           '����KEY
Dim i           As Long             'ٰ�߶���
    

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_TBCME018_SQL.bas -- Function DBDRV_GetTBCME018"

    Select Case formID
        Case "f_cmbc021_1"           '�uFTIR(Oi,Cs)���ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc022_1"           '�uGFA(Oi)���ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc023_1"           '�u��R���ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc024_1"           '�uBMD���ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmec030_1"           '�uBMD���ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc025_1"           '�uOSF���ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmec031_1"           '�uOSF���ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc026_1"           '�uGD���ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc027_1"           '�u���C�t�^�C�����ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc028_1"           '�uFPD���ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc029_1"           '�uGFA�Z�����ݒ�v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    
    End Select
    
    sqlBase = sqlBase & "From TBCME018"
    
    '''SQL��Where���쐬
    For i = 0 To UBound(hin)
        With hin(i)
            key = key & "'" & .hinban & Format(.mnorevno, "00000") & .factory & .opecond & "'"
            If i <> UBound(hin) Then
                key = key & ", "
            End If
        End With
    Next
    sqlWhere = " Where(HINBAN||TO_CHAR(MNOREVNO, 'FM00000')||FACTORY||OPECOND in(" & key & "))"
    sql = sqlBase & sqlWhere
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME018 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''�t�B�[���h����o�^����
    fldCnt = rs.Fields.COUNT
    ReDim fldNames(fldCnt)
    For i = 1 To fldCnt
        fldNames(i) = rs.FieldName(i - 1)
    Next
    
    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            If fldNameExist("HINBAN") Then .hinban = rs("HINBAN")           ' �i��
            If fldNameExist("MNOREVNO") Then .mnorevno = rs("MNOREVNO")     ' ���i�ԍ������ԍ�
            If fldNameExist("FACTORY") Then .factory = rs("FACTORY")        ' �H��
            If fldNameExist("OPECOND") Then .opecond = rs("OPECOND")        ' ���Ə���
            If fldNameExist("HMGSTRRNO") Then .HMGSTRRNO = rs("HMGSTRRNO")  ' �i�Ǘ��d�l�o�^�˗��ԍ�
            If fldNameExist("HMGSTFNO") Then .HMGSTFNO = rs("HMGSTFNO")     ' �i�Ǘ��Ј��m��
            If fldNameExist("HMGSXSNO") Then .HMGSXSNO = rs("HMGSXSNO")     ' �i�Ǘ��r�w���i�ԍ�
            If fldNameExist("HMGSXSNE") Then .HMGSXSNE = fncNullCheck(rs("HMGSXSNE"))     ' �i�Ǘ��r�w���i�ԍ��}��
            If fldNameExist("CONFLAG") Then .CONFLAG = rs("CONFLAG")        ' �m�F�t���O
            If fldNameExist("REINFLAG") Then .REINFLAG = rs("REINFLAG")     ' �ĕt�^�t���O
            If fldNameExist("HSXTRWKB") Then .HSXTRWKB = rs("HSXTRWKB")     ' �i�r�w�����ۋ敪
            If fldNameExist("HSXTYPE") Then .HSXTYPE = rs("HSXTYPE")        ' �i�r�w�^�C�v
            If fldNameExist("KSXTYPKW") Then .KSXTYPKW = rs("KSXTYPKW")     ' �i�r�w�^�C�v�������@
            If fldNameExist("HSXDOP") Then .HSXDOP = rs("HSXDOP")           ' �i�r�w�h�[�p���g
            If fldNameExist("HSXRMIN") Then .HSXRMIN = fncNullCheck(rs("HSXRMIN"))        ' �i�r�w���R����
            If fldNameExist("HSXRMAX") Then .HSXRMAX = fncNullCheck(rs("HSXRMAX"))        ' �i�r�w���R���
            If fldNameExist("HSXRSPOH") Then .HSXRSPOH = rs("HSXRSPOH")     ' �i�r�w���R����ʒu�Q��
            If fldNameExist("HSXRSPOT") Then .HSXRSPOT = rs("HSXRSPOT")     ' �i�r�w���R����ʒu�Q�_
            If fldNameExist("HSXRSPOI") Then .HSXRSPOI = rs("HSXRSPOI")     ' �i�r�w���R����ʒu�Q��
            If fldNameExist("HSXRHWYT") Then .HSXRHWYT = rs("HSXRHWYT")     ' �i�r�w���R�ۏؕ��@�Q��
            If fldNameExist("HSXRHWYS") Then .HSXRHWYS = rs("HSXRHWYS")     ' �i�r�w���R�ۏؕ��@�Q��
            If fldNameExist("HSXRKWAY") Then .HSXRKWAY = rs("HSXRKWAY")     ' �i�r�w���R�������@
            If fldNameExist("HSXRKHNM") Then .HSXRKHNM = rs("HSXRKHNM")     ' �i�r�w���R�����p�x�Q��
            If fldNameExist("HSXRKHNI") Then .HSXRKHNI = rs("HSXRKHNI")     ' �i�r�w���R�����p�x�Q��
            If fldNameExist("HSXRKHNH") Then .HSXRKHNH = rs("HSXRKHNH")     ' �i�r�w���R�����p�x�Q��
            If fldNameExist("HSXRKHNS") Then .HSXRKHNS = rs("HSXRKHNS")     ' �i�r�w���R�����p�x�Q��
            If fldNameExist("HSXRMCAL") Then .HSXRMCAL = rs("HSXRMCAL")     ' �i�r�w���R�ʓ��v�Z
            If fldNameExist("HSXRMBNP") Then .HSXRMBNP = fncNullCheck(rs("HSXRMBNP"))     ' �i�r�w���R�ʓ����z
            If fldNameExist("HSXRMCL2") Then .HSXRMCL2 = rs("HSXRMCL2")     ' �i�r�w���R�ʓ��v�Z�Q
            If fldNameExist("HSXRMBP2") Then .HSXRMBP2 = fncNullCheck(rs("HSXRMBP2"))     ' �i�r�w���R�ʓ����z�Q
            If fldNameExist("HSXRSDEV") Then .HSXRSDEV = fncNullCheck(rs("HSXRSDEV"))     ' �i�r�w���R�W���΍�
            If fldNameExist("HSXRAMIN") Then .HSXRAMIN = fncNullCheck(rs("HSXRAMIN"))     ' �i�r�w���R���ω���
            If fldNameExist("HSXRAMAX") Then .HSXRAMAX = fncNullCheck(rs("HSXRAMAX"))     ' �i�r�w���R���Ϗ��
            If fldNameExist("HSXFORM") Then .HSXFORM = rs("HSXFORM")        ' �i�r�w�`��
            If fldNameExist("HSXD1CEN") Then .HSXD1CEN = fncNullCheck(rs("HSXD1CEN"))     ' �i�r�w���a�P���S
            If fldNameExist("HSXD1MIN") Then .HSXD1MIN = fncNullCheck(rs("HSXD1MIN"))     ' �i�r�w���a�P����
            If fldNameExist("HSXD1MAX") Then .HSXD1MAX = fncNullCheck(rs("HSXD1MAX"))     ' �i�r�w���a�P���
            If fldNameExist("HSXD2CEN") Then .HSXD2CEN = fncNullCheck(rs("HSXD2CEN"))     ' �i�r�w���a�Q���S
            If fldNameExist("HSXD2MIN") Then .HSXD2MIN = fncNullCheck(rs("HSXD2MIN"))     ' �i�r�w���a�Q����
            If fldNameExist("HSXD2MAX") Then .HSXD2MAX = fncNullCheck(rs("HSXD2MAX"))     ' �i�r�w���a�Q���
            If fldNameExist("HSXCDIR") Then .HSXCDIR = rs("HSXCDIR")        ' �i�r�w�����ʕ���
            If fldNameExist("HSXCSCEN") Then .HSXCSCEN = fncNullCheck(rs("HSXCSCEN"))     ' �i�r�w�����ʌX���S
            If fldNameExist("HSXCSMIN") Then .HSXCSMIN = fncNullCheck(rs("HSXCSMIN"))     ' �i�r�w�����ʌX����
            If fldNameExist("HSXCSMAX") Then .HSXCSMAX = fncNullCheck(rs("HSXCSMAX"))     ' �i�r�w�����ʌX���
            If fldNameExist("HSXCKWAY") Then .HSXCKWAY = rs("HSXCKWAY")     ' �i�r�w�����ʌ������@
            If fldNameExist("HSXCKHNM") Then .HSXCKHNM = rs("HSXCKHNM")     ' �i�r�w�����ʌ����p�x�Q��
            If fldNameExist("HSXCKHNI") Then .HSXCKHNI = rs("HSXCKHNI")     ' �i�r�w�����ʌ����p�x�Q��
            If fldNameExist("HSXCKHNH") Then .HSXCKHNH = rs("HSXCKHNH")     ' �i�r�w�����ʌ����p�x�Q��
            If fldNameExist("HSXCKHNS") Then .HSXCKHNS = rs("HSXCKHNS")     ' �i�r�w�����ʌ����p�x�Q��
            If fldNameExist("HSXCSDIR") Then .HSXCSDIR = rs("HSXCSDIR")     ' �i�r�w�����ʌX����
            If fldNameExist("HSXCSDIS") Then .HSXCSDIS = rs("HSXCSDIS")     ' �i�r�w�����ʌX���ʎw��
            If fldNameExist("HSXCTDIR") Then .HSXCTDIR = rs("HSXCTDIR")     ' �i�r�w�����ʌX�c����
            If fldNameExist("HSXCTCEN") Then .HSXCTCEN = fncNullCheck(rs("HSXCTCEN"))   ' �i�r�w�����ʌX�c���S
            If fldNameExist("HSXCTMIN") Then .HSXCTMIN = fncNullCheck(rs("HSXCTMIN"))     ' �i�r�w�����ʌX�c����
            If fldNameExist("HSXCTMAX") Then .HSXCTMAX = fncNullCheck(rs("HSXCTMAX"))     ' �i�r�w�����ʌX�c���
            If fldNameExist("HSXCYDIR") Then .HSXCYDIR = rs("HSXCYDIR")     ' �i�r�w�����ʌX������
            If fldNameExist("HSXCYCEN") Then .HSXCYCEN = fncNullCheck(rs("HSXCYCEN"))     ' �i�r�w�����ʌX�����S
            If fldNameExist("HSXCYMIN") Then .HSXCYMIN = fncNullCheck(rs("HSXCYMIN"))     ' �i�r�w�����ʌX������
            If fldNameExist("HSXCYMAX") Then .HSXCYMAX = fncNullCheck(rs("HSXCYMAX"))     ' �i�r�w�����ʌX�����
            If fldNameExist("HSXOF1PD") Then .HSXOF1PD = rs("HSXOF1PD")     ' �i�r�w�n�e�P�ʒu����
            If fldNameExist("HSXOF1PN") Then .HSXOF1PN = fncNullCheck(rs("HSXOF1PN"))     ' �i�r�w�n�e�P�ʒu����
            If fldNameExist("HSXOF1PX") Then .HSXOF1PX = fncNullCheck(rs("HSXOF1PX"))     ' �i�r�w�n�e�P�ʒu���
            If fldNameExist("HSXOF1PW") Then .HSXOF1PW = rs("HSXOF1PW")     ' �i�r�w�n�e�P�ʒu�������@
            If fldNameExist("HSXOF1LC") Then .HSXOF1LC = fncNullCheck(rs("HSXOF1LC"))     ' �i�r�w�n�e�P�����S
            If fldNameExist("HSXOF1LN") Then .HSXOF1LN = fncNullCheck(rs("HSXOF1LN"))     ' �i�r�w�n�e�P������
            If fldNameExist("HSXOF1LX") Then .HSXOF1LX = fncNullCheck(rs("HSXOF1LX"))     ' �i�r�w�n�e�P�����
            If fldNameExist("HSXOF1DC") Then .HSXOF1DC = fncNullCheck(rs("HSXOF1DC"))     ' �i�r�w�n�e�P���a���S
            If fldNameExist("HSXOF1DN") Then .HSXOF1DN = fncNullCheck(rs("HSXOF1DN"))     ' �i�r�w�n�e�P���a����
            If fldNameExist("HSXOF1DX") Then .HSXOF1DX = fncNullCheck(rs("HSXOF1DX"))     ' �i�r�w�n�e�P���a���
            If fldNameExist("HSXDFORM") Then .HSXDFORM = rs("HSXDFORM")     ' �i�r�w�a�`��
            If fldNameExist("HSXDPDRC") Then .HSXDPDRC = rs("HSXDPDRC")     ' �i�r�w�a�ʒu����
            If fldNameExist("HSXDPACN") Then .HSXDPACN = fncNullCheck(rs("HSXDPACN"))     ' �i�r�w�a�ʒu�p�x���S
            If fldNameExist("HSXDPAMN") Then .HSXDPAMN = fncNullCheck(rs("HSXDPAMN"))     ' �i�r�w�a�ʒu�p�x����
            If fldNameExist("HSXDPAMX") Then .HSXDPAMX = fncNullCheck(rs("HSXDPAMX"))     ' �i�r�w�a�ʒu�p�x���
            If fldNameExist("HSXDPKWY") Then .HSXDPKWY = rs("HSXDPKWY")     ' �i�r�w�a�ʒu�������@
            If fldNameExist("HSXDPDIR") Then .HSXDPDIR = rs("HSXDPDIR")     ' �i�r�w�a�ʒu����
            If fldNameExist("HSXDPMIN") Then .HSXDPMIN = fncNullCheck(rs("HSXDPMIN"))     ' �i�r�w�a�ʒu����
            If fldNameExist("HSXDPMAX") Then .HSXDPMAX = fncNullCheck(rs("HSXDPMAX"))     ' �i�r�w�a�ʒu���
            If fldNameExist("HSXDWCEN") Then .HSXDWCEN = fncNullCheck(rs("HSXDWCEN"))     ' �i�r�w�a�В��S
            If fldNameExist("HSXDWMIN") Then .HSXDWMIN = fncNullCheck(rs("HSXDWMIN"))     ' �i�r�w�a�Љ���
            If fldNameExist("HSXDWMAX") Then .HSXDWMAX = fncNullCheck(rs("HSXDWMAX"))     ' �i�r�w�a�Џ��
            If fldNameExist("HSXDDCEN") Then .HSXDDCEN = fncNullCheck(rs("HSXDDCEN"))     ' �i�r�w�a�[���S
            If fldNameExist("HSXDDMIN") Then .HSXDDMIN = fncNullCheck(rs("HSXDDMIN"))     ' �i�r�w�a�[����
            If fldNameExist("HSXDDMAX") Then .HSXDDMAX = fncNullCheck(rs("HSXDDMAX"))     ' �i�r�w�a�[���
            If fldNameExist("HSXDACEN") Then .HSXDACEN = fncNullCheck(rs("HSXDACEN"))     ' �i�r�w�a�p�x���S
            If fldNameExist("HSXDAMIN") Then .HSXDAMIN = fncNullCheck(rs("HSXDAMIN"))     ' �i�r�w�a�p�x����
            If fldNameExist("HSXDAMAX") Then .HSXDAMAX = fncNullCheck(rs("HSXDAMAX"))     ' �i�r�w�a�p�x���
            If fldNameExist("IFKBN") Then .IFKBN = rs("IFKBN")              ' �h�^�e�敪
            If fldNameExist("SYORIKBN") Then .SYORIKBN = rs("SYORIKBN")     ' �����敪
            If fldNameExist("SPECRRNO") Then .SPECRRNO = rs("SPECRRNO")     ' �d�l�o�^�˗��ԍ�
            If fldNameExist("SXLMCNO") Then .SXLMCNO = rs("SXLMCNO")        ' �r�w�k��������ԍ�
            If fldNameExist("WFMCNO") Then .WFMCNO = rs("WFMCNO")           ' �v�e��������ԍ�
            If fldNameExist("STAFFID") Then .StaffID = rs("STAFFID")        ' �Ј�ID
            If fldNameExist("REGDATE") Then .REGDATE = rs("REGDATE")        ' �o�^���t
            If fldNameExist("UPDDATE") Then .UPDDATE = rs("UPDDATE")        ' �X�V���t
            If fldNameExist("SENDFLAG") Then .SENDFLAG = rs("SENDFLAG")     ' ���M�t���O
            If fldNameExist("SENDDATE") Then .SENDDATE = rs("SENDDATE")     ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME018 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMJ002�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMJ002 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺
Public Function DBDRV_GetTBCMJ002(records() As typ_TBCMJ002, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, KRPROCCD, PROCCODE, HINBAN, REVNUM, FACTORY," & _
              " OPECOND, GOUKI, TYPE, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, EFEHS, RRG, JUDGDATA, TSTAFFID, REGDATE, KSTAFFID," & _
              " UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMJ002"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMJ002 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .POSITION = rs("POSITION")       ' �ʒu
            .SMPKBN = rs("SMPKBN")           ' �T���v���敪
            .TRANCOND = rs("TRANCOND")       ' ��������
            .TRANCNT = rs("TRANCNT")         ' ������
            .SMPLNO = rs("SMPLNO")           ' �T���v���m��
            .SMPLUMU = rs("SMPLUMU")         ' �T���v���L��
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
            .PROCCODE = rs("PROCCODE")       ' �H���R�[�h
            .hinban = rs("HINBAN")           ' �i��
            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
            .factory = rs("FACTORY")         ' �H��
            .opecond = rs("OPECOND")         ' ���Ə���
            .GOUKI = rs("GOUKI")             ' ���@
            .TYPE = rs("TYPE")               ' �^�C�v
            .MEAS1 = rs("MEAS1")             ' ����l�P
            .MEAS2 = rs("MEAS2")             ' ����l�Q
            .MEAS3 = rs("MEAS3")             ' ����l�R
            .MEAS4 = rs("MEAS4")             ' ����l�S
            .MEAS5 = rs("MEAS5")             ' ����l�T
            .EFEHS = rs("EFEHS")             ' �����ΐ�
            .RRG = rs("RRG")                 ' �q�q�f
            .JudgData = rs("JUDGDATA")       ' �����Ώےl
            .TSTAFFID = rs("TSTAFFID")       ' �o�^�Ј�ID
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .KSTAFFID = rs("KSTAFFID")       ' �X�V�Ј�ID
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMJ002 = FUNCTION_RETURN_SUCCESS
End Function
'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMH004�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMH004 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺
Public Function DBDRV_GetTBCMH004(records() As typ_TBCMH004, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, KRPROCCD, PROCCODE, LENGTOP, LENGTKDO, LENGTAIL, LENGFREE, DM1, DM2, DM3, WGHTTOP, WGHTTKDO," & _
              " WGHTTAIL, WGHTFREE, WGTOPCUT, UPWEIGHT, CHARGE, SEED, STATCLS, JDGECODE, PWTIME, ADDDPPOS, ADDDPCLS, ADDDPVAL," & _
              " ADDDPNAM, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMH004"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMH004 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
            .PROCCODE = rs("PROCCODE")       ' �H���R�[�h
            .LENGTOP = rs("LENGTOP")         ' �����iTOP�j
            .LENGTKDO = rs("LENGTKDO")       ' �����i�����j
            .LENGTAIL = rs("LENGTAIL")       ' �����iTAIL�j
            .LENGFREE = rs("LENGFREE")       ' �t���[����
            .DM1 = rs("DM1")                 ' �������a�P
            .DM2 = rs("DM2")                 ' �������a�Q
            .DM3 = rs("DM3")                 ' �������a�R
            .WGHTTOP = rs("WGHTTOP")         ' �d�ʁiTOP�j
            .WGHTTKDO = rs("WGHTTKDO")       ' �d�ʁi�����j
            .WGHTTAIL = rs("WGHTTAIL")       ' �d�ʁiTAIL)
            .WGHTFREE = rs("WGHTFREE")       ' �d�ʁi�t���[�����j
            .WGTOPCUT = rs("WGTOPCUT")       ' �g�b�v�J�b�g�d��
            .UPWEIGHT = rs("UPWEIGHT")       ' ���グ�d��
            .CHARGE = rs("CHARGE")           ' �`���[�W��
            .SEED = rs("SEED")               ' �V�[�h
            .STATCLS = rs("STATCLS")         ' BOT�󋵋敪
            .JDGECODE = rs("JDGECODE")       ' ����R�[�h
            .PWTIME = rs("PWTIME")           ' �p���[����
            .ADDDPPOS = rs("ADDDPPOS")       ' �ǉ��h�[�v�ʒu
            .ADDDPCLS = rs("ADDDPCLS")       ' �ǉ��h�[�p���g���
            .ADDDPVAL = rs("ADDDPVAL")       ' �ǉ��h�[�v��
            .ADDDPNAM = rs("ADDDPNAM")       ' �ǉ��h�[�v��
            .TSTAFFID = rs("TSTAFFID")       ' �o�^�Ј�ID
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .KSTAFFID = rs("KSTAFFID")       ' �X�V�Ј�ID
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMH004 = FUNCTION_RETURN_SUCCESS
End Function
'�T�v      :�e�[�u���uXSDCS�v�̏����ɂ��������R�[�h���X�V����
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :records       ,I   ,typ_XSDCS   ,�X�V���R�[�h
'          :[sqlWhere]    ,I   ,String         ,�X�V����(SQL��Where��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,�X�V�̐���
'����      :
'����      :2001/07/13�쐬�@�ɓ�
Public Function DBDRV_UpdateTBCME043(records As typ_XSDCS, Optional sqlWhere$ = vbNullString) As FUNCTION_RETURN
    Dim sql As String
    
    DBDRV_UpdateTBCME043 = FUNCTION_RETURN_FAILURE

    With records
'        sql = "update TBCME043 set "
''        sql = sql & "HINBAN='" & .HINBAN & "', "              ' �i��
''        sql = sql & "REVNUM=" & .REVNUM & ", "                ' ���i�ԍ������ԍ�
''        sql = sql & "FACTORY='" & .FACTORY & "', "            ' �H��
''        sql = sql & "OPECOND='" & .OPECOND & "', "            ' ���Ə���
''        sql = sql & "KTKBN='" & .KTKBN & "', "                ' �m��敪
''        sql = sql & "CRYINDRS='" & .CRYINDRS & "', "          ' ���������w���iRs)
''        sql = sql & "CRYINDOI='" & .CRYINDOI & "', "          ' ���������w���iOi)
''        sql = sql & "CRYINDB1='" & .CRYINDB1 & "', "          ' ���������w���iB1)
''        sql = sql & "CRYINDB2='" & .CRYINDB2 & "', "          ' ���������w���iB2�j
''        sql = sql & "CRYINDB3='" & .CRYINDB3 & "', "          ' ���������w���iB3)
''        sql = sql & "CRYINDL1='" & .CRYINDL1 & "', "          ' ���������w���iL1)
''        sql = sql & "CRYINDL2='" & .CRYINDL2 & "', "          ' ���������w���iL2)
''        sql = sql & "CRYINDL3='" & .CRYINDL3 & "', "          ' ���������w���iL3)
''        sql = sql & "CRYINDL4='" & .CRYINDL4 & "', "          ' ���������w���iL4)
''        sql = sql & "CRYINDCS='" & .CRYINDCS & "', "          ' ���������w���iCs)
''        sql = sql & "CRYINDGD='" & .CRYINDGD & "', "          ' ���������w���iGD)
''        sql = sql & "CRYINDT='" & .CRYINDT & "', "            ' ���������w���iT)
''        sql = sql & "CRYINDEP='" & .CRYINDEP & "', "          ' ���������w���iEPD)
'        sql = sql & "CRYRESRS='" & .CRYRESRS & "', "          ' �����������сiRs)
'        sql = sql & "CRYRESOI='" & .CRYRESOI & "', "          ' �����������сiOi)
'        sql = sql & "CRYRESB1='" & .CRYRESB1 & "', "          ' �����������сiB1)
'        sql = sql & "CRYRESB2='" & .CRYRESB2 & "', "          ' �����������сiB2�j
'        sql = sql & "CRYRESB3='" & .CRYRESB3 & "', "          ' �����������сiB3)
'        sql = sql & "CRYRESL1='" & .CRYRESL1 & "', "          ' �����������сiL1)
'        sql = sql & "CRYRESL2='" & .CRYRESL2 & "', "          ' �����������сiL2)
'        sql = sql & "CRYRESL3='" & .CRYRESL3 & "', "          ' �����������сiL3)
'        sql = sql & "CRYRESL4='" & .CRYRESL4 & "', "          ' �����������сiL4)
'        sql = sql & "CRYRESCS='" & .CRYRESCS & "', "          ' �����������сiCs)
'        sql = sql & "CRYRESGD='" & .CRYRESGD & "', "          ' �����������сiGD)
'        sql = sql & "CRYREST='" & .CRYREST & "', "            ' �����������сiT)
'        sql = sql & "CRYRESEP='" & .CRYRESEP & "', "          ' �����������сiEPD)
''        sql = sql & "SMPLNUM=" & .SMPLNUM & ", "              ' �T���v������
''        sql = sql & "SMPLPAT='" & .SMPLPAT & "', "            ' �T���v���p�^�[��
'        sql = sql & "UPDDATE=sysdate, "                       ' �X�V���t
'        sql = sql & "SENDFLAG='0'"                            ' ���M�t���O


        sql = "update XSDCS set "
        sql = sql & "CRYRESRS1CS='" & .CRYRESRS1CS & "', "          ' �����������сiRs)
        sql = sql & "CRYRESRS2CS='" & .CRYRESRS2CS & "', "          ' �����������сiRs)
        sql = sql & "CRYRESOICS='" & .CRYRESOICS & "', "          ' �����������сiOi)
        sql = sql & "CRYRESB1CS='" & .CRYRESB1CS & "', "          ' �����������сiB1)
        sql = sql & "CRYRESB2CS='" & .CRYRESB2CS & "', "          ' �����������сiB2�j
        sql = sql & "CRYRESB3CS='" & .CRYRESB3CS & "', "          ' �����������сiB3)
        sql = sql & "CRYRESL1CS='" & .CRYRESL1CS & "', "          ' �����������сiL1)
        sql = sql & "CRYRESL2CS='" & .CRYRESL2CS & "', "          ' �����������сiL2)
        sql = sql & "CRYRESL3CS='" & .CRYRESL3CS & "', "          ' �����������сiL3)
        sql = sql & "CRYRESL4CS='" & .CRYRESL4CS & "', "          ' �����������сiL4)
        sql = sql & "CRYRESCSCS='" & .CRYRESCSCS & "', "          ' �����������сiCs)
        sql = sql & "CRYRESGDCS='" & .CRYRESGDCS & "', "          ' �����������сiGD)
        sql = sql & "CRYRESTCS='" & .CRYRESTCS & "', "            ' �����������сiT)
        sql = sql & "CRYRESEPCS='" & .CRYRESEPCS & "', "          ' �����������сiEPD)
        sql = sql & "KDAYCS=sysdate, "                       ' �X�V���t
        sql = sql & "SNDKCS='0'"                            ' ���M�t���O

    End With

    If sqlWhere <> vbNullString Then
        sql = sql & " " & sqlWhere
    End If

    If OraDB.ExecuteSQL(sql) <= 0 Then
        Exit Function
    End If

    DBDRV_UpdateTBCME043 = FUNCTION_RETURN_SUCCESS

End Function

'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uXSDCS�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_XSDCS ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺
Public Function DBDRV_GetTBCME043(records() As typ_XSDCS, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
'    sqlBase = "Select CRYNUMCS, SMPKBNCS, TBKBNCS, REPSMPLIDCS, XTALCS, INPOSCS, HINBCS, REVNUMCS, FACTORYCS, OPECS, KTKBNCS, BLKKTFLAGCS, " & _
'              " CRYSMPLIDRSCS, CRYSMPLIDRS1CS, CRYSMPLIDRS2CS, CRYINDRSCS, CRYRESRS1CS, CRYRESRS2CS, CRYSMPLIDOICS, CRYINDOICS, CRYRESOICS, " & _
'              " CRYSMPLIDB1CS, CRYINDB1CS, CRYRESB1CS, CRYSMPLIDB2CS, CRYINDB2CS, CRYRESB2CS, CRYSMPLIDB3CS, CRYINDB3CS, CRYRESB3CS, CRYSMPLIDL1CS, " & _
'              " CRYINDL1CS, CRYRESL1CS, CRYSMPLIDL2CS, CRYINDL2CS, CRYRESL2CS, CRYSMPLIDL3CS, CRYINDL3CS, CRYRESL3CS, CRYSMPLIDL4CS, CRYINDL4CS, " & _
'              " CRYRESL4CS, CRYSMPLIDCSCS, CRYINDCSCS, CRYRESCSCS, CRYSMPLIDGDCS, CRYINDGDCS, CRYRESGDCS, CRYSMPLIDTCS, CRYINDTCS, CRYRESTCS, " & _
'              " CRYSMPLIDEPCS, CRYINDEPCS, CRYRESEPCS, SMPLNUMCS, SMPLPATCS, TSTAFFCS, TDAYCS, KSTAFFCS, KDAYCS, SNDKCS, SNDDAYCS "
    sqlBase = "Select CRYNUMCS, SMPKBNCS, TBKBNCS, REPSMPLIDCS, XTALCS, INPOSCS, HINBCS, REVNUMCS, FACTORYCS, OPECS, KTKBNCS, BLKKTFLAGCS, " & _
              " CRYSMPLIDRSCS, nvl(CRYSMPLIDRS1CS, 0) as CRYSMPLIDRS1CS, nvl(CRYSMPLIDRS2CS, 0) as CRYSMPLIDRS2CS, CRYINDRSCS, CRYRESRS1CS, nvl(CRYRESRS2CS, ' ') as CRYRESRS2CS, CRYSMPLIDOICS, CRYINDOICS, CRYRESOICS, " & _
              " CRYSMPLIDB1CS, CRYINDB1CS, CRYRESB1CS, CRYSMPLIDB2CS, CRYINDB2CS, CRYRESB2CS, CRYSMPLIDB3CS, CRYINDB3CS, CRYRESB3CS, CRYSMPLIDL1CS, " & _
              " CRYINDL1CS, CRYRESL1CS, CRYSMPLIDL2CS, CRYINDL2CS, CRYRESL2CS, CRYSMPLIDL3CS, CRYINDL3CS, CRYRESL3CS, CRYSMPLIDL4CS, CRYINDL4CS, " & _
              " CRYRESL4CS, CRYSMPLIDCSCS, CRYINDCSCS, CRYRESCSCS, CRYSMPLIDGDCS, CRYINDGDCS, CRYRESGDCS, CRYSMPLIDTCS, CRYINDTCS, CRYRESTCS, " & _
              " CRYSMPLIDEPCS, CRYINDEPCS, CRYRESEPCS, SMPLNUMCS, SMPLPATCS, nvl(TSTAFFCS, ' ') as TSTAFFCS, TDAYCS, nvl(KSTAFFCS, ' ') as KSTAFFCS, KDAYCS, nvl(SNDKCS, ' ') as SNDKCS, nvl(SNDDAYCS, sysdate) as SNDDAYCS "
    sqlBase = sqlBase & "From XSDCS"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME043 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
'            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
'            .IngotPos = rs("INGOTPOS")       ' �������ʒu
'            .SMPKBN = rs("SMPKBN")           ' �T���v���敪
'            .SMPLNO = rs("SMPLNO")           ' �T���v��No
'            .hinban = rs("HINBAN")           ' �i��
'            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
'            .factory = rs("FACTORY")         ' �H��
'            .opecond = rs("OPECOND")         ' ���Ə���
'            .KTKBN = rs("KTKBN")             ' �m��敪
'            .CRYINDRS = rs("CRYINDRS")       ' ���������w���iRs)
'            .CRYINDOI = rs("CRYINDOI")       ' ���������w���iOi)
'            .CRYINDB1 = rs("CRYINDB1")       ' ���������w���iB1)
'            .CRYINDB2 = rs("CRYINDB2")       ' ���������w���iB2�j
'            .CRYINDB3 = rs("CRYINDB3")       ' ���������w���iB3)
'            .CRYINDL1 = rs("CRYINDL1")       ' ���������w���iL1)
'            .CRYINDL2 = rs("CRYINDL2")       ' ���������w���iL2)
'            .CRYINDL3 = rs("CRYINDL3")       ' ���������w���iL3)
'            .CRYINDL4 = rs("CRYINDL4")       ' ���������w���iL4)
'            .CRYINDCS = rs("CRYINDCS")       ' ���������w���iCs)
'            .CRYINDGD = rs("CRYINDGD")       ' ���������w���iGD)
'            .CRYINDT = rs("CRYINDT")         ' ���������w���iT)
'            .CRYINDEP = rs("CRYINDEP")       ' ���������w���iEPD)
'            .CRYRESRS = rs("CRYRESRS")       ' �����������сiRs)
'            .CRYRESOI = rs("CRYRESOI")       ' �����������сiOi)
'            .CRYRESB1 = rs("CRYRESB1")       ' �����������сiB1)
'            .CRYRESB2 = rs("CRYRESB2")       ' �����������сiB2�j
'            .CRYRESB3 = rs("CRYRESB3")       ' �����������сiB3)
'            .CRYRESL1 = rs("CRYRESL1")       ' �����������сiL1)
'            .CRYRESL2 = rs("CRYRESL2")       ' �����������сiL2)
'            .CRYRESL3 = rs("CRYRESL3")       ' �����������сiL3)
'            .CRYRESL4 = rs("CRYRESL4")       ' �����������сiL4)
'            .CRYRESCS = rs("CRYRESCS")       ' �����������сiCs)
'            .CRYRESGD = rs("CRYRESGD")       ' �����������сiGD)
'            .CRYREST = rs("CRYREST")         ' �����������сiT)
'            .CRYRESEP = rs("CRYRESEP")       ' �����������сiEPD)
'            .SMPLNUM = rs("SMPLNUM")         ' �T���v������
'            .SMPLPAT = rs("SMPLPAT")         ' �T���v���p�^�[��
'            .REGDATE = rs("REGDATE")         ' �o�^���t
'            .UPDDATE = rs("UPDDATE")         ' �X�V���t
'            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
'            .SENDDATE = rs("SENDDATE")       ' ���M���t

            If IsNull(rs("CRYNUMCS")) = False Then .CRYNUMCS = rs("CRYNUMCS")                   ' �u���b�NID
            If IsNull(rs("SMPKBNCS")) = False Then .SMPKBNCS = rs("SMPKBNCS")                   ' �T���v���敪
            If IsNull(rs("TBKBNCS")) = False Then .TBKBNCS = rs("TBKBNCS")                      ' T/B�敪
            If IsNull(rs("REPSMPLIDCS")) = False Then .REPSMPLIDCS = rs("REPSMPLIDCS")          ' ��\�T���v��ID
            If IsNull(rs("XTALCS")) = False Then .XTALCS = rs("XTALCS")                         ' �����ԍ�
            If IsNull(rs("INPOSCS")) = False Then .INPOSCS = rs("INPOSCS")                      ' �������ʒu
            If IsNull(rs("HINBCS")) = False Then .HINBCS = rs("HINBCS")                         ' �i��
            If IsNull(rs("REVNUMCS")) = False Then .REVNUMCS = rs("REVNUMCS")                   ' ���i�ԍ������ԍ�
            If IsNull(rs("FACTORYCS")) = False Then .FACTORYCS = rs("FACTORYCS")                ' �H��
            If IsNull(rs("OPECS")) = False Then .OPECS = rs("OPECS")                            ' ���Ə���
            If IsNull(rs("KTKBNCS")) = False Then .KTKBNCS = rs("KTKBNCS")                      ' �m��敪
            If IsNull(rs("BLKKTFLAGCS")) = False Then .BLKKTFLAGCS = rs("BLKKTFLAGCS")          ' �u���b�N�m��t���O
            If IsNull(rs("CRYSMPLIDRSCS")) = False Then .CRYSMPLIDRSCS = rs("CRYSMPLIDRSCS")    ' �T���v��ID(Rs)
            If IsNull(rs("CRYSMPLIDRS1CS")) = False Then .CRYSMPLIDRS1CS = rs("CRYSMPLIDRS1CS") ' ����T���v��ID1(Rs)
            If IsNull(rs("CRYSMPLIDRS2CS")) = False Then .CRYSMPLIDRS2CS = rs("CRYSMPLIDRS2CS") ' ����T���v��ID2(Rs)
            If IsNull(rs("CRYINDRSCS")) = False Then .CRYINDRSCS = rs("CRYINDRSCS")             ' ���FLG(Rs)
            If IsNull(rs("CRYRESRS1CS")) = False Then .CRYRESRS1CS = rs("CRYRESRS1CS")          ' ����FLG1(Rs)
            If IsNull(rs("CRYRESRS2CS")) = False Then .CRYRESRS2CS = rs("CRYRESRS2CS")          ' ����FLG2(Rs)
            If IsNull(rs("CRYSMPLIDOICS")) = False Then .CRYSMPLIDOICS = rs("CRYSMPLIDOICS")    ' �T���v��ID(Oi)
            If IsNull(rs("CRYINDOICS")) = False Then .CRYINDOICS = rs("CRYINDOICS")             ' ���FLG(Oi)
            If IsNull(rs("CRYRESOICS")) = False Then .CRYRESOICS = rs("CRYRESOICS")             ' ����FLG(Oi)
            If IsNull(rs("CRYSMPLIDB1CS")) = False Then .CRYSMPLIDB1CS = rs("CRYSMPLIDB1CS")    ' �T���v��ID(B1)
            If IsNull(rs("CRYINDB1CS")) = False Then .CRYINDB1CS = rs("CRYINDB1CS")             ' ���FLG(B1)
            If IsNull(rs("CRYRESB1CS")) = False Then .CRYRESB1CS = rs("CRYRESB1CS")             ' ����FLG(B1)
            If IsNull(rs("CRYSMPLIDB2CS")) = False Then .CRYSMPLIDB2CS = rs("CRYSMPLIDB2CS")    ' �T���v��ID(B2)
            If IsNull(rs("CRYINDB2CS")) = False Then .CRYINDB2CS = rs("CRYINDB2CS")             ' ���FLG(B2)
            If IsNull(rs("CRYRESB2CS")) = False Then .CRYRESB2CS = rs("CRYRESB2CS")             ' ����FLG(B2)
            If IsNull(rs("CRYSMPLIDB3CS")) = False Then .CRYSMPLIDB3CS = rs("CRYSMPLIDB3CS")    ' �T���v��ID(B3)
            If IsNull(rs("CRYINDB3CS")) = False Then .CRYINDB3CS = rs("CRYINDB3CS")             ' ���FLG(B3)
            If IsNull(rs("CRYRESB3CS")) = False Then .CRYRESB3CS = rs("CRYRESB3CS")             ' ����FLG(B3)
            If IsNull(rs("CRYSMPLIDL1CS")) = False Then .CRYSMPLIDL1CS = rs("CRYSMPLIDL1CS")    ' �T���v��ID(L1)
            If IsNull(rs("CRYINDL1CS")) = False Then .CRYINDL1CS = rs("CRYINDL1CS")             ' ���FLG(L1)
            If IsNull(rs("CRYRESL1CS")) = False Then .CRYRESL1CS = rs("CRYRESL1CS")             ' ����FLG(L1)
            If IsNull(rs("CRYSMPLIDL2CS")) = False Then .CRYSMPLIDL2CS = rs("CRYSMPLIDL2CS")    ' �T���v��ID(L2)
            If IsNull(rs("CRYINDL2CS")) = False Then .CRYINDL2CS = rs("CRYINDL2CS")             ' ���FLG(L2)
            If IsNull(rs("CRYRESL2CS")) = False Then .CRYRESL2CS = rs("CRYRESL2CS")             ' ����FLG(L2)
            If IsNull(rs("CRYSMPLIDL3CS")) = False Then .CRYSMPLIDL3CS = rs("CRYSMPLIDL3CS")    ' �T���v��ID(L3)
            If IsNull(rs("CRYINDL3CS")) = False Then .CRYINDL3CS = rs("CRYINDL3CS")             ' ���FLG(L3)
            If IsNull(rs("CRYRESL3CS")) = False Then .CRYRESL3CS = rs("CRYRESL3CS")             ' ����FLG(L3)
            If IsNull(rs("CRYSMPLIDL4CS")) = False Then .CRYSMPLIDL4CS = rs("CRYSMPLIDL4CS")    ' �T���v��ID(L4)
            If IsNull(rs("CRYINDL4CS")) = False Then .CRYINDL4CS = rs("CRYINDL4CS")             ' ���FLG(L4)
            If IsNull(rs("CRYRESL4CS")) = False Then .CRYRESL4CS = rs("CRYRESL4CS")             ' ����FLG(L4)
            If IsNull(rs("CRYSMPLIDCSCS")) = False Then .CRYSMPLIDCSCS = rs("CRYSMPLIDCSCS")    ' �T���v��ID(Cs)
            If IsNull(rs("CRYINDCSCS")) = False Then .CRYINDCSCS = rs("CRYINDCSCS")             ' ���FLG(Cs)
            If IsNull(rs("CRYRESCSCS")) = False Then .CRYRESCSCS = rs("CRYRESCSCS")             ' ����FLG(Cs)
            If IsNull(rs("CRYSMPLIDGDCS")) = False Then .CRYSMPLIDGDCS = rs("CRYSMPLIDGDCS")    ' �T���v��ID(GD)
            If IsNull(rs("CRYINDGDCS")) = False Then .CRYINDGDCS = rs("CRYINDGDCS")             ' ���FLG(GD)
            If IsNull(rs("CRYRESGDCS")) = False Then .CRYRESGDCS = rs("CRYRESGDCS")             ' ����FLG(GD)
            If IsNull(rs("CRYSMPLIDTCS")) = False Then .CRYSMPLIDTCS = rs("CRYSMPLIDTCS")       ' �T���v��ID(T)
            If IsNull(rs("CRYINDTCS")) = False Then .CRYINDTCS = rs("CRYINDTCS")                ' ���FLG(T)
            If IsNull(rs("CRYRESTCS")) = False Then .CRYRESTCS = rs("CRYRESTCS")                ' ����FLG(T)
            If IsNull(rs("CRYSMPLIDEPCS")) = False Then .CRYSMPLIDEPCS = rs("CRYSMPLIDEPCS")    ' �T���v��ID(EPD)
            If IsNull(rs("CRYINDEPCS")) = False Then .CRYINDEPCS = rs("CRYINDEPCS")             ' ���FLG(EPD)
            If IsNull(rs("CRYRESEPCS")) = False Then .CRYRESEPCS = rs("CRYRESEPCS")             ' ����FLG(EPD)
            If IsNull(rs("SMPLNUMCS")) = False Then .SMPLNUMCS = rs("SMPLNUMCS")                ' �T���v������
            If IsNull(rs("SMPLPATCS")) = False Then .SMPLPATCS = rs("SMPLPATCS")                ' �T���v���p�^�[��
            If IsNull(rs("TSTAFFCS")) = False Then .TSTAFFCS = rs("TSTAFFCS")                   ' �o�^�Ј�ID
            If IsNull(rs("TDAYCS")) = False Then .TDAYCS = rs("TDAYCS")                         ' �o�^���t
            If IsNull(rs("KSTAFFCS")) = False Then .KSTAFFCS = rs("KSTAFFCS")                   ' �X�V�Ј�ID
            If IsNull(rs("KDAYCS")) = False Then .KDAYCS = rs("KDAYCS")                         ' �X�V���t
            If IsNull(rs("SNDKCS")) = False Then .SNDKCS = rs("SNDKCS")                         ' ���M�t���O
            If IsNull(rs("SNDDAYCS")) = False Then .SNDDAYCS = rs("SNDDAYCS")                   ' ���M���t

        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME043 = FUNCTION_RETURN_SUCCESS
End Function

'�T�v      :�e�[�u���uXSDCS�v�̏����ɂ��������R�[�h���X�V����
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :records       ,I   ,typ_XSDCS   ,�X�V���R�[�h
'          :[sqlWhere]    ,I   ,String         ,�X�V����(SQL��Where��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,�X�V�̐���
'����      :
'����      :2001/07/13�쐬�@�ɓ�
Public Function DBDRV_UpdateXSDCS(sqlUpdate As String) As FUNCTION_RETURN
    
    DBDRV_UpdateXSDCS = FUNCTION_RETURN_FAILURE

    If OraDB.ExecuteSQL(sqlUpdate) <= 0 Then
        Exit Function
    End If

    DBDRV_UpdateXSDCS = FUNCTION_RETURN_SUCCESS

End Function
