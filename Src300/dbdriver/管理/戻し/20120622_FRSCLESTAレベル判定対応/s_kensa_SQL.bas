Attribute VB_Name = "s_kensa_SQL"
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

Public Function DBDRV_GetTBCME019(records() As typ_TBCME019, formID$, HIN() As tFullHinban) As FUNCTION_RETURN
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
        
        
        Case "f_cmbc053_1i"           '�u�w������ ���ѓ��́v
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
    For i = 0 To UBound(HIN)
        With HIN(i)
            key = key & "'" & .hinban & Format(.mnorevno, "00000") & .factory & .opecond & "'"
            If i <> UBound(HIN) Then
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

Public Function DBDRV_GetTBCME020(records() As typ_TBCME020, formID$, HIN() As tFullHinban) As FUNCTION_RETURN
Dim sql         As String           'SQL�S��
Dim sqlBase     As String           'SQL��{��(WHERE�߂̑O�܂�)
Dim sqlWhere    As String           'SQLWhere��
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga START ---
Dim sqlAnd      As String           'SQLAnd��
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga End   ---
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
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga START ---
             sqlBase = "Select T.HINBAN, T.MNOREVNO, T.FACTORY, T.OPECOND, T.HMGSTRRNO, T.HMGSTFNO, T.HMGSXSNO, T.HMGSXSNE, T.HSXDENKU, T.HSXDENMX, T.HSXDENMN," & _
              " T.HSXDENHT, T.HSXDENHS, T.HSXDVDKU, T.HSXDVDMXN, T.HSXDVDMNN, T.HSXDVDHT, T.HSXDVDHS, T.HSXLDLKU, T.HSXLDLMX, T.HSXLDLMN, T.HSXLDLHT," & _
              " T.HSXLDLHS, T.HSXGDSZY, T.HSXGDSPH, T.HSXGDSPT, T.HSXGDSPR, T.HSXGDZAR, T.HSXGDKHM, T.HSXGDKHI, T.HSXGDKHH, T.HSXGDKHS, T.HSXDSOKE," & _
              " T.HSXDSOMX, T.HSXDSOMN, T.HSXDSOAX, T.HSXDSOAN, T.HSXDSOHT, T.HSXDSOHS, T.HSXDSOKM, T.HSXDSOKI, T.HSXDSOKH, T.HSXDSOKS, T.HSXLIFTW," & _
              " T.HSXSDSLP, T.HSXGKKNO, T.HSXCDOP, T.HSXCDOPN, T.HSXCDPNI, T.HSXGSFIN, T.HSXCLMIN, T.HSXCLMAX, T.HSXCLPMN, T.HSXCLPR, T.HSXWFWAR," & _
              " T.HSXOF1AX, T.HSXOF1MX, T.HSXOF1SH, T.HSXOF1ST, T.HSXOF1SR, T.HSXOF1HT, T.HSXOF1HS, T.HSXOF1SZ, T.HSXOF1KM, T.HSXOF1KI, T.HSXOF1KH," & _
              " T.HSXOF1KS, T.HSXOF1NS, T.HSXOF1ET, T.HSXOF2AX, T.HSXOF2MX, T.HSXOF2SH, T.HSXOF2ST, T.HSXOF2SR, T.HSXOF2HT, T.HSXOF2HS, T.HSXOF2SZ," & _
              " T.HSXOF2KM, T.HSXOF2KI, T.HSXOF2KH, T.HSXOF2KS, T.HSXOF2NS, T.HSXOF2ET, T.HSXOF3AX, T.HSXOF3MX, T.HSXOF3SH, T.HSXOF3ST, T.HSXOF3SR,"
            sqlBase = sqlBase & " T.HSXOF3HT, T.HSXOF3HS, T.HSXOF3SZ, T.HSXOF3KM, T.HSXOF3KI, T.HSXOF3KH, T.HSXOF3KS, T.HSXOF3NS, T.HSXOF3ET, T.HSXOF4AX, T.HSXOF4MX," & _
              " T.HSXOF4SH, T.HSXOF4ST, T.HSXOF4SR, T.HSXOF4HT, T.HSXOF4HS, T.HSXOF4SZ, T.HSXOF4KM, T.HSXOF4KI, T.HSXOF4KH, T.HSXOF4KS, T.HSXOF4NS," & _
              " T.HSXOF4ET, T.HSXBM1AN, T.HSXBM1AX, T.HSXBM1SH, T.HSXBM1ST, T.HSXBM1SR, T.HSXBM1HT, T.HSXBM1HS, T.HSXBM1SZ, T.HSXBM1KM, T.HSXBM1KI," & _
              " T.HSXBM1KH, T.HSXBM1KS, T.HSXBM1NS, T.HSXBM1ET, T.HSXBM2AN, T.HSXBM2AX, T.HSXBM2SH, T.HSXBM2ST, T.HSXBM2SR, T.HSXBM2HT, T.HSXBM2HS," & _
              " T.HSXBM2SZ, T.HSXBM2KM, T.HSXBM2KI, T.HSXBM2KH, T.HSXBM2KS, T.HSXBM2NS, T.HSXBM2ET, T.HSXBM3AN, T.HSXBM3AX, T.HSXBM3SH, T.HSXBM3ST," & _
              " T.HSXBM3SR, T.HSXBM3HT, T.HSXBM3HS, T.HSXBM3SZ, T.HSXBM3KM, T.HSXBM3KI, T.HSXBM3KH, T.HSXBM3KS, T.HSXBM3NS, T.HSXBM3ET, T.HSXNOTE," & _
              " T.HSXOSF1PTK, T.HSXOSF2PTK, T.HSXOSF3PTK, T.HSXOSF4PTK,"
' OSF�CBMD���ڒǉ��Ή�  ���@1�s���@2002.04.02 yakimura
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga END   ---
            For i = 1 To 10
                sqlBase = sqlBase & "T.HSXRS" & i & "N, "
                sqlBase = sqlBase & "T.HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "T.SPECRRNO, T.SXLMCNO, T.WFMCNO, T.STAFFID, T.REGDATE, T.UPDDATE, T.SENDFLAG, T.SENDDATE, U.COSF3FLAG, T.HSXCOSF3NS "
        
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
            sqlBase = sqlBase & ", HSXGDPTK "   '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
            
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
    
    
        Case "f_cmbc053_1"           '�uX������ ���ѓ��́v
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

        'Add Start 2010/12/17 SMPK Miyata
        Case "f_cmbc054_1"           '�uCu-deco���ѓ��́v
             sqlBase = "Select T.HINBAN, T.MNOREVNO, T.FACTORY, T.OPECOND, T.HMGSTRRNO, T.HMGSTFNO, T.HMGSXSNO, T.HMGSXSNE, T.HSXDENKU, T.HSXDENMX, T.HSXDENMN," & _
              " T.HSXDENHT, T.HSXDENHS, T.HSXDVDKU, T.HSXDVDMXN, T.HSXDVDMNN, T.HSXDVDHT, T.HSXDVDHS, T.HSXLDLKU, T.HSXLDLMX, T.HSXLDLMN, T.HSXLDLHT," & _
              " T.HSXLDLHS, T.HSXGDSZY, T.HSXGDSPH, T.HSXGDSPT, T.HSXGDSPR, T.HSXGDZAR, T.HSXGDKHM, T.HSXGDKHI, T.HSXGDKHH, T.HSXGDKHS, T.HSXDSOKE," & _
              " T.HSXDSOMX, T.HSXDSOMN, T.HSXDSOAX, T.HSXDSOAN, T.HSXDSOHT, T.HSXDSOHS, T.HSXDSOKM, T.HSXDSOKI, T.HSXDSOKH, T.HSXDSOKS, T.HSXLIFTW," & _
              " T.HSXSDSLP, T.HSXGKKNO, T.HSXCDOP, T.HSXCDOPN, T.HSXCDPNI, T.HSXGSFIN, T.HSXCLMIN, T.HSXCLMAX, T.HSXCLPMN, T.HSXCLPR, T.HSXWFWAR," & _
              " T.HSXOF1AX, T.HSXOF1MX, T.HSXOF1SH, T.HSXOF1ST, T.HSXOF1SR, T.HSXOF1HT, T.HSXOF1HS, T.HSXOF1SZ, T.HSXOF1KM, T.HSXOF1KI, T.HSXOF1KH," & _
              " T.HSXOF1KS, T.HSXOF1NS, T.HSXOF1ET, T.HSXOF2AX, T.HSXOF2MX, T.HSXOF2SH, T.HSXOF2ST, T.HSXOF2SR, T.HSXOF2HT, T.HSXOF2HS, T.HSXOF2SZ," & _
              " T.HSXOF2KM, T.HSXOF2KI, T.HSXOF2KH, T.HSXOF2KS, T.HSXOF2NS, T.HSXOF2ET, T.HSXOF3AX, T.HSXOF3MX, T.HSXOF3SH, T.HSXOF3ST, T.HSXOF3SR,"
            sqlBase = sqlBase & " T.HSXOF3HT, T.HSXOF3HS, T.HSXOF3SZ, T.HSXOF3KM, T.HSXOF3KI, T.HSXOF3KH, T.HSXOF3KS, T.HSXOF3NS, T.HSXOF3ET, T.HSXOF4AX, T.HSXOF4MX," & _
              " T.HSXOF4SH, T.HSXOF4ST, T.HSXOF4SR, T.HSXOF4HT, T.HSXOF4HS, T.HSXOF4SZ, T.HSXOF4KM, T.HSXOF4KI, T.HSXOF4KH, T.HSXOF4KS, T.HSXOF4NS," & _
              " T.HSXOF4ET, T.HSXBM1AN, T.HSXBM1AX, T.HSXBM1SH, T.HSXBM1ST, T.HSXBM1SR, T.HSXBM1HT, T.HSXBM1HS, T.HSXBM1SZ, T.HSXBM1KM, T.HSXBM1KI," & _
              " T.HSXBM1KH, T.HSXBM1KS, T.HSXBM1NS, T.HSXBM1ET, T.HSXBM2AN, T.HSXBM2AX, T.HSXBM2SH, T.HSXBM2ST, T.HSXBM2SR, T.HSXBM2HT, T.HSXBM2HS," & _
              " T.HSXBM2SZ, T.HSXBM2KM, T.HSXBM2KI, T.HSXBM2KH, T.HSXBM2KS, T.HSXBM2NS, T.HSXBM2ET, T.HSXBM3AN, T.HSXBM3AX, T.HSXBM3SH, T.HSXBM3ST," & _
              " T.HSXBM3SR, T.HSXBM3HT, T.HSXBM3HS, T.HSXBM3SZ, T.HSXBM3KM, T.HSXBM3KI, T.HSXBM3KH, T.HSXBM3KS, T.HSXBM3NS, T.HSXBM3ET, T.HSXNOTE," & _
              " T.HSXOSF1PTK, T.HSXOSF2PTK, T.HSXOSF3PTK, T.HSXOSF4PTK,"
            sqlBase = sqlBase & "T.HSXCPK, T.HSXCSZ, T.HSXCHT, T.HSXCHS, T.HSXCJPK, T.HSXCJNS, T.HSXCJHT, T.HSXCJHS, " & _
              " T.HSXCJLTPK, T.HSXCJLTNS, T.HSXCJLTHT, T.HSXCJLTHS, T.HSXCJ2PK, T.HSXCJ2NS, T.HSXCJ2HT, T.HSXCJ2HS, "
            For i = 1 To 10
                sqlBase = sqlBase & "T.HSXRS" & i & "N, "
                sqlBase = sqlBase & "T.HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "T.SPECRRNO, T.SXLMCNO, T.WFMCNO, T.STAFFID, T.REGDATE, T.UPDDATE, T.SENDFLAG, T.SENDDATE, U.COSF3FLAG "
        'Add End   2010/12/17 SMPK Miyata

    End Select
       
       
'Chg Start 2010/12/17 SMPK Miyata
'    If Trim(formID) = "f_cmbc025_1" Then
    If Trim(formID) = "f_cmbc025_1" Or Trim(formID) = "f_cmbc054_1" Then
'Chg End   2010/12/17 SMPK Miyata
        sqlBase = sqlBase & "From TBCME020 T , TBCME036 U"
    Else
        sqlBase = sqlBase & "From TBCME020"
    End If
    
    '''SQL��Where���쐬
    For i = 0 To UBound(HIN)
        With HIN(i)
            key = key & "'" & .hinban & Format(.mnorevno, "00000") & .factory & .opecond & "'"
            If i <> UBound(HIN) Then
                key = key & ", "
            End If
        End With
    Next
    
'Chg Start 2010/12/17 SMPK Miyata
''C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga START ---
'    If Trim(formID) = "f_cmbc025_1" Then
    If Trim(formID) = "f_cmbc025_1" Or Trim(formID) = "f_cmbc054_1" Then
'Chg End   2010/12/17 SMPK Miyata
        sqlWhere = " Where(T.HINBAN||TO_CHAR(T.MNOREVNO, 'FM00000')||T.FACTORY||T.OPECOND in(" & key & "))"
        sqlAnd = " And(U.HINBAN||TO_CHAR(U.MNOREVNO, 'FM00000')||U.FACTORY||U.OPECOND in(" & key & "))"
        sql = sqlBase & sqlWhere & sqlAnd
    Else
        sqlWhere = " Where(HINBAN||TO_CHAR(MNOREVNO, 'FM00000')||FACTORY||OPECOND in(" & key & "))"
        sql = sqlBase & sqlWhere
    End If
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga END   ---
    
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
            If fldNameExist("HINBAN") Then .hinban = rs("HINBAN")           ' �i��
            If fldNameExist("MNOREVNO") Then .mnorevno = rs("MNOREVNO")       ' ���i�ԍ������ԍ�
            If fldNameExist("FACTORY") Then .factory = rs("FACTORY")         ' �H��
            If fldNameExist("OPECOND") Then .opecond = rs("OPECOND")         ' ���Ə���
            If fldNameExist("HMGSTRRNO") Then .HMGSTRRNO = rs("HMGSTRRNO")     ' �i�Ǘ��d�l�o�^�˗��ԍ�
            If fldNameExist("HMGSTFNO") Then .HMGSTFNO = rs("HMGSTFNO")       ' �i�Ǘ��Ј��m��
            If fldNameExist("HMGSXSNO") Then .HMGSXSNO = rs("HMGSXSNO")       ' �i�Ǘ��r�w���i�ԍ�
            If fldNameExist("HMGSXSNE") Then .HMGSXSNE = fncNullCheck(rs("HMGSXSNE"))       ' �i�Ǘ��r�w���i�ԍ��}��
            If fldNameExist("HSXDENKU") Then .HSXDENKU = rs("HSXDENKU")       ' �i�r�w�c���������L��
            If fldNameExist("HSXDENMX") Then .HSXDENMX = fncNullCheck(rs("HSXDENMX"))       ' �i�r�w�c�������
            If fldNameExist("HSXDENMN") Then .HSXDENMN = fncNullCheck(rs("HSXDENMN"))       ' �i�r�w�c��������
            If fldNameExist("HSXDENHT") Then .HSXDENHT = rs("HSXDENHT")       ' �i�r�w�c�����ۏؕ��@�Q��
            If fldNameExist("HSXDENHS") Then .HSXDENHS = rs("HSXDENHS")       ' �i�r�w�c�����ۏؕ��@�Q��
            If fldNameExist("HSXDVDKU") Then .HSXDVDKU = rs("HSXDVDKU")       ' �i�r�w�c�u�c�Q�����L��
            If fldNameExist("HSXDVDMXN") Then .HSXDVDMX = fncNullCheck(rs("HSXDVDMXN"))       ' �i�r�w�c�u�c�Q���    �v�e�T���v�������ύX 2003.05.20 yakimura
            If fldNameExist("HSXDVDMNN") Then .HSXDVDMN = fncNullCheck(rs("HSXDVDMNN"))       ' �i�r�w�c�u�c�Q����    �v�e�T���v�������ύX 2003.05.20 yakimura
            If fldNameExist("HSXDVDHT") Then .HSXDVDHT = rs("HSXDVDHT")       ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
            If fldNameExist("HSXDVDHS") Then .HSXDVDHS = rs("HSXDVDHS")       ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
            If fldNameExist("HSXLDLKU") Then .HSXLDLKU = rs("HSXLDLKU")       ' �i�r�w�k�^�c�k�����L��
            If fldNameExist("HSXLDLMX") Then .HSXLDLMX = fncNullCheck(rs("HSXLDLMX"))       ' �i�r�w�k�^�c�k���
            If fldNameExist("HSXLDLMN") Then .HSXLDLMN = fncNullCheck(rs("HSXLDLMN"))       ' �i�r�w�k�^�c�k����
            If fldNameExist("HSXLDLHT") Then .HSXLDLHT = rs("HSXLDLHT")       ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
            If fldNameExist("HSXLDLHS") Then .HSXLDLHS = rs("HSXLDLHS")       ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
            If fldNameExist("HSXGDSZY") Then .HSXGDSZY = rs("HSXGDSZY")       ' �i�r�w�f�c�������
            If fldNameExist("HSXGDSPH") Then .HSXGDSPH = rs("HSXGDSPH")       ' �i�r�w�f�c����ʒu�Q��
            If fldNameExist("HSXGDSPT") Then .HSXGDSPT = rs("HSXGDSPT")       ' �i�r�w�f�c����ʒu�Q�_
            If fldNameExist("HSXGDSPR") Then .HSXGDSPR = rs("HSXGDSPR")       ' �i�r�w�f�c����ʒu�Q��
            If fldNameExist("HSXGDZAR") Then .HSXGDZAR = fncNullCheck(rs("HSXGDZAR"))       ' �i�r�w�f�c���O�̈�
            If fldNameExist("HSXGDKHM") Then .HSXGDKHM = rs("HSXGDKHM")       ' �i�r�w�f�c�����p�x�Q��
            If fldNameExist("HSXGDKHI") Then .HSXGDKHI = rs("HSXGDKHI")       ' �i�r�w�f�c�����p�x�Q��
            If fldNameExist("HSXGDKHH") Then .HSXGDKHH = rs("HSXGDKHH")       ' �i�r�w�f�c�����p�x�Q��
            If fldNameExist("HSXGDKHS") Then .HSXGDKHS = rs("HSXGDKHS")       ' �i�r�w�f�c�����p�x�Q��
            If fldNameExist("HSXDSOKE") Then .HSXDSOKE = rs("HSXDSOKE")       ' �i�r�w�c�r�n�c����
            If fldNameExist("HSXDSOMX") Then .HSXDSOMX = fncNullCheck(rs("HSXDSOMX"))       ' �i�r�w�c�r�n�c���
            If fldNameExist("HSXDSOMN") Then .HSXDSOMN = fncNullCheck(rs("HSXDSOMN"))       ' �i�r�w�c�r�n�c����
            If fldNameExist("HSXDSOAX") Then .HSXDSOAX = fncNullCheck(rs("HSXDSOAX"))       ' �i�r�w�c�r�n�c�̈���
            If fldNameExist("HSXDSOAN") Then .HSXDSOAN = fncNullCheck(rs("HSXDSOAN"))       ' �i�r�w�c�r�n�c�̈扺��
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
            If fldNameExist("HSXCLMIN") Then .HSXCLMIN = fncNullCheck(rs("HSXCLMIN"))       ' �i�r�w����������
            If fldNameExist("HSXCLMAX") Then .HSXCLMAX = fncNullCheck(rs("HSXCLMAX"))       ' �i�r�w���������
            If fldNameExist("HSXCLPMN") Then .HSXCLPMN = fncNullCheck(rs("HSXCLPMN"))       ' �i�r�w���������e����
            If fldNameExist("HSXCLPR") Then .HSXCLPR = fncNullCheck(rs("HSXCLPR"))         ' �i�r�w���������e�䗦
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
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
                If fldNameExist("HSXOSF" & j & "PTK") Then                       ' �i�r�w�n�r�e(n)�p�^���敪
                   If IsNull(rs("HSXOSF" & j & "PTK")) = False Then .HSXOSF_PTK(j) = rs("HSXOSF" & j & "PTK")
                   End If
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
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
'NULL�Ή� 2003/12/21
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'                If fldNameExist("HSXBMD" & j & "MBP") Then                      ' �i�r�w�a�l�c(n)�ʓ����z
'                   If IsNull(rs("HSXBMD" & j & "MBP")) = False Then .HSXBMD_MBP(j) = fncNullCheck(rs("HSXBMD" & j & "MBP"))
'                   End If
                If fldNameExist("HSXBMD" & j & "MBP") Then .HSXBMD_MBP(j) = fncNullCheck(rs("HSXBMD" & j & "MBP")) ' �i�r�w�a�l�c(n)�ʓ����z
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'NULL�Ή� 2003/12/21
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
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
                If fldNameExist("HSXOSF1PTK") Then                           ' �i�r�w�n�r�e1�p�^���敪
                   If IsNull(rs("HSXOSF1PTK")) = False Then .HSXOSF1PTK = rs("HSXOSF1PTK")
                   End If
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
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
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
                If fldNameExist("HSXOSF2PTK") Then                           ' �i�r�w�n�r�e2�p�^���敪
                   If IsNull(rs("HSXOSF2PTK")) = False Then .HSXOSF2PTK = rs("HSXOSF2PTK")
                   End If
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
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
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
                If fldNameExist("HSXOSF3PTK") Then                           ' �i�r�w�n�r�e3�p�^���敪
                   If IsNull(rs("HSXOSF3PTK")) = False Then .HSXOSF3PTK = rs("HSXOSF3PTK")
                   End If
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
                If fldNameExist("HSXOF4AX") Then .HSXOF4AX = fncNullCheck(rs("HSXOF4AX"))  ' �i�r�w�n�r�e4���Ϗ��
                If fldNameExist("HSXOF4MX") Then .HSXOF4MX = fncNullCheck(rs("HSXOF4MX"))  ' �i�r�w�n�r�e4���
                If fldNameExist("HSXOF4SH") Then .HSXOF4SH = rs("HSXOF4SH")  ' �i�r�w�n�r�e4����ʒu�Q��
                If fldNameExist("HSXOF4ST") Then .HSXOF4ST = rs("HSXOF4ST")  ' �i�r�w�n�r�e4����ʒu�Q�_
                If fldNameExist("HSXOF4SR") Then .HSXOF4SR = rs("HSXOF4SR")  ' �i�r�w�n�r�e4����ʒu�Q��
                If fldNameExist("HSXOF4HT") Then .HSXOF4HT = rs("HSXOF4HT")  ' �i�r�w�n�r�e4�ۏؕ��@�Q��
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga START ---
                If fldNameExist("COSF3FLAG") Then
                    If IsNull(rs("COSF3FLAG")) = False Then .HSXOF4HS = rs("COSF3FLAG") Else .HSXOF4HS = " "
                End If
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga END   ---
                If fldNameExist("HSXOF4SZ") Then .HSXOF4SZ = rs("HSXOF4SZ")  ' �i�r�w�n�r�e4�������
                If fldNameExist("HSXOF4KM") Then .HSXOF4KM = rs("HSXOF4KM")  ' �i�r�w�n�r�e4�����p�x�Q��
                If fldNameExist("HSXOF4KI") Then .HSXOF4KI = rs("HSXOF4KI")  ' �i�r�w�n�r�e4�����p�x�Q��
                If fldNameExist("HSXOF4KH") Then .HSXOF4KH = rs("HSXOF4KH")  ' �i�r�w�n�r�e4�����p�x�Q��
                If fldNameExist("HSXOF4KS") Then .HSXOF4KS = rs("HSXOF4KS")  ' �i�r�w�n�r�e4�����p�x�Q��
                If fldNameExist("HSXOF4NS") Then .HSXOF4NS = rs("HSXOF4NS")  ' �i�r�w�n�r�e4�M�����@
                If fldNameExist("HSXOF4ET") Then .HSXOF4ET = fncNullCheck(rs("HSXOF4ET"))  ' �i�r�w�n�r�e4�I���d�s��
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
                If fldNameExist("HSXOSF4PTK") Then                           ' �i�r�w�n�r�e4�p�^���敪
                   If IsNull(rs("HSXOSF4PTK")) = False Then .HSXOSF4PTK = rs("HSXOSF4PTK")
                   End If
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura

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
'NULL�Ή� 2003/12/21
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'                If fldNameExist("HSXBMD1MBP") Then                           ' �i�r�w�a�l�c1�ʓ����z
'                   If IsNull(rs("HSXBMD1MBP")) = False Then .HSXBMD1MBP = fncNullCheck(rs("HSXBMD1MBP"))
'                   End If
                If fldNameExist("HSXBMD1MBP") Then .HSXBMD1MBP = fncNullCheck(rs("HSXBMD1MBP")) ' �i�r�w�a�l�c1�ʓ����z
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'NULL�Ή� 2003/12/21
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
'NULL�Ή� 2003/12/21
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'                If fldNameExist("HSXBMD2MBP") Then                           ' �i�r�w�a�l�c2�ʓ����z
'                   If IsNull(rs("HSXBMD2MBP")) = False Then .HSXBMD2MBP = rs("HSXBMD2MBP")
'                   End If
                If fldNameExist("HSXBMD2MBP") Then .HSXBMD2MBP = rs("HSXBMD2MBP") ' �i�r�w�a�l�c2�ʓ����z
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'NULL�Ή� 2003/12/21
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
'NULL�Ή� 2003/12/21
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'                If fldNameExist("HSXBMD3MBP") Then                           ' �i�r�w�a�l�c3�ʓ����z
'                   If IsNull(rs("HSXBMD3MBP")) = False Then .HSXBMD3MBP = fncNullCheck(rs("HSXBMD3MBP"))
'                   End If
                If fldNameExist("HSXBMD3MBP") Then .HSXBMD3MBP = fncNullCheck(rs("HSXBMD3MBP")) ' �i�r�w�a�l�c3�ʓ����z
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'NULL�Ή� 2003/12/21
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
                If fldNameExist("HSXRS10N") Then .HSXRS10N = rs("HSXRS10N")     ' �i�r�w�\��10�Q��

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
            
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
                If fldNameExist("HSXGDPTK") Then         ' �i�r�w�f�c�p�^���敪
                If IsNull(rs("HSXGDPTK")) = False Then .HSXGDPTK = rs("HSXGDPTK") Else .HSXGDPTK = " "
            End If
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End

            'Add Start 2011/01/26 SMPK Miyata
            If fldNameExist("HSXCPK") Then
                If IsNull(rs("HSXCPK")) = False Then .HSXCPK = rs("HSXCPK")         '�i�r�w�b�p�^�[���敪
            End If
            If fldNameExist("HSXCSZ") Then
                If IsNull(rs("HSXCSZ")) = False Then .HSXCSZ = rs("HSXCSZ")         '�i�r�w�b�������
            End If
            If fldNameExist("HSXCHT") Then
                If IsNull(rs("HSXCHT")) = False Then .HSXCHT = rs("HSXCHT")         '�i�r�w�b�ۏؕ��@�Q��
            End If
            If fldNameExist("HSXCHS") Then
                If IsNull(rs("HSXCHS")) = False Then .HSXCHS = rs("HSXCHS")         '�i�r�w�b�ۏؕ��@�Q��
            End If
            If fldNameExist("HSXCJPK") Then
                If IsNull(rs("HSXCJPK")) = False Then .HSXCJPK = rs("HSXCJPK")       '�i�r�w�b�i�p�^�[���敪
            End If
            If fldNameExist("HSXCJNS") Then
                If IsNull(rs("HSXCJNS")) = False Then .HSXCJNS = rs("HSXCJNS")       '�i�r�w�b�i�M�����@
            End If
            If fldNameExist("HSXCJHT") Then
                If IsNull(rs("HSXCJHT")) = False Then .HSXCJHT = rs("HSXCJHT")       '�i�r�w�b�i�ۏؕ��@�Q��
            End If
            If fldNameExist("HSXCJHS") Then
                If IsNull(rs("HSXCJHS")) = False Then .HSXCJHS = rs("HSXCJHS")       '�i�r�w�b�i�ۏؕ��@�Q��
            End If
            If fldNameExist("HSXCJLTPK") Then
                If IsNull(rs("HSXCJLTPK")) = False Then .HSXCJLTPK = rs("HSXCJLTPK")   '�i�r�w�b�i�k�s�p�^�[���敪
            End If
            If fldNameExist("HSXCJLTNS") Then
                If IsNull(rs("HSXCJLTNS")) = False Then .HSXCJLTNS = rs("HSXCJLTNS")   '�i�r�w�b�i�k�s�M�����@
            End If
            If fldNameExist("HSXCJLTHT") Then
                If IsNull(rs("HSXCJLTHT")) = False Then .HSXCJLTHT = rs("HSXCJLTHT")   '�i�r�w�b�i�k�s�ۏؕ��@�Q��
            End If
            If fldNameExist("HSXCJLTHS") Then
                If IsNull(rs("HSXCJLTHS")) = False Then .HSXCJLTHS = rs("HSXCJLTHS")   '�i�r�w�b�i�k�s�ۏؕ��@�Q��
            End If
            If fldNameExist("HSXCJ2PK") Then
                If IsNull(rs("HSXCJ2PK")) = False Then .HSXCJ2PK = rs("HSXCJ2PK")     '�i�r�w�b�i�Q�p�^�[���敪
            End If
            If fldNameExist("HSXCJ2NS") Then
                If IsNull(rs("HSXCJ2NS")) = False Then .HSXCJ2NS = rs("HSXCJ2NS")     '�i�r�w�b�i�Q�M�����@
            End If
            If fldNameExist("HSXCJ2HT") Then
                If IsNull(rs("HSXCJ2HT")) = False Then .HSXCJ2HT = rs("HSXCJ2HT")     '�i�r�w�b�i�Q�ۏؕ��@�Q��
            End If
            If fldNameExist("HSXCJ2HS") Then
                If IsNull(rs("HSXCJ2HS")) = False Then .HSXCJ2HS = rs("HSXCJ2HS")     '�i�r�w�b�i�Q�ۏؕ��@�Q��
            End If
            'Add End   2011/01/26 SMPK Miyata
            
            'Add Start 2011/02/17 Y.Hitomi
            If fldNameExist("HSXCOSF3NS") Then
                If IsNull(rs("HSXCOSF3NS")) = False Then .HSXCOSF3NS = rs("HSXCOSF3NS")     '�i�r�w�b�i�Q�ۏؕ��@�Q��
            End If
            'Add End   2011/02/17 Y.Hitomi
        
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

'�T�v      :�e�[�u���uTBCME021�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :records()     ,O  ,typ_TBCME021    ,���o���R�[�h
'          :formID        ,I  ,String          ,�g�p�t�H�[��ID
'          :sqlOrder      ,I  ,tFullHinban     ,���o�i�ԁi�z��j
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,���o�̐���
'����      :�����L�[�́A�HINBAN�+�uMNOREVNO�v+�uFACTORY�v+�uOPECOND�v�̕�����Ƃ���
'����      :05/03/01 ooba
Public Function DBDRV_GetTBCME021(records() As typ_TBCME021, formID$, HIN() As tFullHinban) As FUNCTION_RETURN

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
    gErr.Push "s_kensa_SQL.bas -- Function DBDRV_GetTBCME021"

    Select Case formID
        Case "f_cmbc026_1"           '�uGD���ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGWFSNO, HMGWFSNE, CONFLAG, " & _
              "REINFLAG, HWFTRWKB, HWFFACES, HWFBACKS, HWFBDSWY, HWFTYPE, HWFTYPKW, HWFDOP, HWFFKBWK, " & _
              "HWFFKBWS, HWFRMIN, HWFRMAX, HWFRSPOH, HWFRSPOT, HWFRSPOI, HWFRHWYT, HWFRHWYS, HWFRKWAY, " & _
              "HWFRKHNM, HWFRKHNN, HWFRKHNH, HWFRKHNU, HWFRSDEV, HWFRAMIN, HWFRAMAX, HWFRMBNP, HWFRMCAL, " & _
              "HWFRMBP2, HWFRMCL2, HWFRKBSH, HWFRKBST, HWFRKBSI, HWFRKBHT, HWFRKBHS, HWFSTMAX, HWFSTSPH, " & _
              "HWFSTSPT, HWFSTSPI, HWFSTHWT, HWFSTHWS, HWFSTKWY, HWFSTKHM, HWFSTKHN, HWFSTKHH, HWFSTKHU, "
            sqlBase = sqlBase & "HWFACEN, HWFAMIN, HWFAMAX, HWFASPOH, HWFASPOT, HWFASPOI, HWFAHWYT, HWFAHWYS, HWFAKWAY, " & _
              "HWFAKHNM, HWFAKHNN, HWFAKHNH, HWFAKHNU, HWFASDEV, HWFAAMIN, HWFAAMAX, HWFAMBNP, HWFAMCAL, " & _
              "HWFALTBP, HWFALTCL, HWFALTRA, HWFAMRAN, HWFDIVS, HWFAKBSH, HWFAKBST, HWFAKBSI, HWFAKBHT, " & _
              "HWFAKBHS, HWFWFORM, HWFD1CEN, HWFD1MIN, HWFD1MAX, HWFD2CEN, HWFD2MIN, HWFD2MAX, HWFDKHNM, " & _
              "HWFDKHNN, HWFDKHNH, HWFDKHNU, HWFLPMNP, HWFSGMNP, HWFETMNP, HWFMPMNP, HWFLPKS1, HWFLPKS2, " & _
              "HWFLPKZ1, HWFLPKZ2, IFKBN, SYORIKBN, SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        '�ǉ� 2005/06/15 ffc)tanabe start
        Case "f_cmec067_1"           '�uSPV���юQ�Ɓv
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HWFTYPE, HWFD1CEN "
        '�ǉ� 2005/06/15 ffc)tanabe end
        
    End Select
       
    sqlBase = sqlBase & "From TBCME021"
    
    '''SQL��Where���쐬
    For i = 0 To UBound(HIN)
        With HIN(i)
            key = key & "'" & .hinban & Format(.mnorevno, "00000") & .factory & .opecond & "'"
            If i <> UBound(HIN) Then
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
        DBDRV_GetTBCME021 = FUNCTION_RETURN_FAILURE
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
            If fldNameExist("HINBAN") Then .hinban = rs("HINBAN") '�i��
            If fldNameExist("MNOREVNO") Then .mnorevno = rs("MNOREVNO") '���i�ԍ������ԍ�
            If fldNameExist("FACTORY") Then .factory = rs("FACTORY") '�H��
            If fldNameExist("OPECOND") Then .opecond = rs("OPECOND") '���Ə���
            If fldNameExist("HMGSTRRNO") Then .HMGSTRRNO = rs("HMGSTRRNO") '�i�Ǘ��d�l�o�^�˗��ԍ�
            If fldNameExist("HMGSTFNO") Then .HMGSTFNO = rs("HMGSTFNO") '�i�Ǘ��Ј��m��
            If fldNameExist("HMGWFSNO") Then .HMGWFSNO = rs("HMGWFSNO") '�i�Ǘ��v�e���i�ԍ�
            If fldNameExist("HMGWFSNE") Then .HMGWFSNE = fncNullCheck(rs("HMGWFSNE")) '�i�Ǘ��v�e���i�ԍ��}��
            If fldNameExist("CONFLAG") Then .CONFLAG = rs("CONFLAG") '�m�F�t���O
            If fldNameExist("REINFLAG") Then .REINFLAG = rs("REINFLAG") '�ĕt�^�t���O
            If fldNameExist("HWFTRWKB") Then .HWFTRWKB = rs("HWFTRWKB") '�i�v�e�����ۋ敪
            If fldNameExist("HWFFACES") Then .HWFFACES = rs("HWFFACES") '�i�v�e�\�ʎd�グ
            If fldNameExist("HWFBACKS") Then .HWFBACKS = rs("HWFBACKS") '�i�v�e���d�グ
            If fldNameExist("HWFBDSWY") Then .HWFBDSWY = rs("HWFBDSWY") '�i�v�e�a�c�������@
            If fldNameExist("HWFTYPE") Then .HWFTYPE = rs("HWFTYPE") '�i�v�e�^�C�v
            If fldNameExist("HWFTYPKW") Then .HWFTYPKW = rs("HWFTYPKW") '�i�v�e�^�C�v�������@
            If fldNameExist("HWFDOP") Then .HWFDOP = rs("HWFDOP") '�i�v�e�h�[�p���g
            If fldNameExist("HWFFKBWK") Then .HWFFKBWK = rs("HWFFKBWK") '�i�v�e�\�ʋ敪���@�Q��
            If fldNameExist("HWFFKBWS") Then .HWFFKBWS = rs("HWFFKBWS") '�i�v�e�\�ʋ敪���@�Q�w
            If fldNameExist("HWFRMIN") Then .HWFRMIN = fncNullCheck(rs("HWFRMIN")) '�i�v�e���R����
            If fldNameExist("HWFRMAX") Then .HWFRMAX = fncNullCheck(rs("HWFRMAX")) '�i�v�e���R���
            If fldNameExist("HWFRSPOH") Then .HWFRSPOH = rs("HWFRSPOH") '�i�v�e���R����ʒu�Q��
            If fldNameExist("HWFRSPOT") Then .HWFRSPOT = rs("HWFRSPOT") '�i�v�e���R����ʒu�Q�_
            If fldNameExist("HWFRSPOI") Then .HWFRSPOI = rs("HWFRSPOI") '�i�v�e���R����ʒu�Q��
            If fldNameExist("HWFRHWYT") Then .HWFRHWYT = rs("HWFRHWYT") '�i�v�e���R�ۏؕ��@�Q��
            If fldNameExist("HWFRHWYS") Then .HWFRHWYS = rs("HWFRHWYS") '�i�v�e���R�ۏؕ��@�Q��
            If fldNameExist("HWFRKWAY") Then .HWFRKWAY = rs("HWFRKWAY") '�i�v�e���R�������@
            If fldNameExist("HWFRKHNM") Then .HWFRKHNM = rs("HWFRKHNM") '�i�v�e���R�����p�x�Q��
            If fldNameExist("HWFRKHNN") Then .HWFRKHNN = rs("HWFRKHNN") '�i�v�e���R�����p�x�Q��
            If fldNameExist("HWFRKHNH") Then .HWFRKHNH = rs("HWFRKHNH") '�i�v�e���R�����p�x�Q��
            If fldNameExist("HWFRKHNU") Then .HWFRKHNU = rs("HWFRKHNU") '�i�v�e���R�����p�x�Q�E
            If fldNameExist("HWFRSDEV") Then .HWFRSDEV = fncNullCheck(rs("HWFRSDEV")) '�i�v�e���R�W���΍�
            If fldNameExist("HWFRAMIN") Then .HWFRAMIN = fncNullCheck(rs("HWFRAMIN")) '�i�v�e���R���ω���
            If fldNameExist("HWFRAMAX") Then .HWFRAMAX = fncNullCheck(rs("HWFRAMAX")) '�i�v�e���R���Ϗ��
            If fldNameExist("HWFRMBNP") Then .HWFRMBNP = fncNullCheck(rs("HWFRMBNP")) '�i�v�e���R�ʓ����z
            If fldNameExist("HWFRMCAL") Then .HWFRMCAL = rs("HWFRMCAL") '�i�v�e���R�ʓ��v�Z
            If fldNameExist("HWFRMBP2") Then .HWFRMBP2 = fncNullCheck(rs("HWFRMBP2")) '�i�v�e���R�ʓ����z�Q
            If fldNameExist("HWFRMCL2") Then .HWFRMCL2 = rs("HWFRMCL2") '�i�v�e���R�ʓ��v�Z�Q
            If fldNameExist("HWFRKBSH") Then .HWFRKBSH = rs("HWFRKBSH") '�i�v�e���R�U�敪����ʒu�Q��
            If fldNameExist("HWFRKBST") Then .HWFRKBST = rs("HWFRKBST") '�i�v�e���R�U�敪����ʒu�Q�_
            If fldNameExist("HWFRKBSI") Then .HWFRKBSI = rs("HWFRKBSI") '�i�v�e���R�U�敪����ʒu�Q��
            If fldNameExist("HWFRKBHT") Then .HWFRKBHT = rs("HWFRKBHT") '�i�v�e���R�U�敪�ۏؕ��@�Q��
            If fldNameExist("HWFRKBHS") Then .HWFRKBHS = rs("HWFRKBHS") '�i�v�e���R�U�敪�ۏؕ��@�Q��
            If fldNameExist("HWFSTMAX") Then .HWFSTMAX = fncNullCheck(rs("HWFSTMAX")) '�i�v�e�X�g���G���
            If fldNameExist("HWFSTSPH") Then .HWFSTSPH = rs("HWFSTSPH") '�i�v�e�X�g���G����ʒu�Q��
            If fldNameExist("HWFSTSPT") Then .HWFSTSPT = rs("HWFSTSPT") '�i�v�e�X�g���G����ʒu�Q�_
            If fldNameExist("HWFSTSPI") Then .HWFSTSPI = rs("HWFSTSPI") '�i�v�e�X�g���G����ʒu�Q��
            If fldNameExist("HWFSTHWT") Then .HWFSTHWT = rs("HWFSTHWT") '�i�v�e�X�g���G�ۏؕ��@�Q��
            If fldNameExist("HWFSTHWS") Then .HWFSTHWS = rs("HWFSTHWS") '�i�v�e�X�g���G�ۏؕ��@�Q��
            If fldNameExist("HWFSTKWY") Then .HWFSTKWY = rs("HWFSTKWY") '�i�v�e�X�g���G�������@
            If fldNameExist("HWFSTKHM") Then .HWFSTKHM = rs("HWFSTKHM") '�i�v�e�X�g���G�����p�x�Q��
            If fldNameExist("HWFSTKHN") Then .HWFSTKHN = rs("HWFSTKHN") '�i�v�e�X�g���G�����p�x�Q��
            If fldNameExist("HWFSTKHH") Then .HWFSTKHH = rs("HWFSTKHH") '�i�v�e�X�g���G�����p�x�Q��
            If fldNameExist("HWFSTKHU") Then .HWFSTKHU = rs("HWFSTKHU") '�i�v�e�X�g���G�����p�x�Q�E
            If fldNameExist("HWFACEN") Then .HWFACEN = fncNullCheck(rs("HWFACEN")) '�i�v�e�����S
            If fldNameExist("HWFAMIN") Then .HWFAMIN = fncNullCheck(rs("HWFAMIN")) '�i�v�e������
            If fldNameExist("HWFAMAX") Then .HWFAMAX = fncNullCheck(rs("HWFAMAX")) '�i�v�e�����
            If fldNameExist("HWFASPOH") Then .HWFASPOH = rs("HWFASPOH") '�i�v�e������ʒu�Q��
            If fldNameExist("HWFASPOT") Then .HWFASPOT = rs("HWFASPOT") '�i�v�e������ʒu�Q�_
            If fldNameExist("HWFASPOI") Then .HWFASPOI = rs("HWFASPOI") '�i�v�e������ʒu�Q��
            If fldNameExist("HWFAHWYT") Then .HWFAHWYT = rs("HWFAHWYT") '�i�v�e���ۏؕ��@�Q��
            If fldNameExist("HWFAHWYS") Then .HWFAHWYS = rs("HWFAHWYS") '�i�v�e���ۏؕ��@�Q��
            If fldNameExist("HWFAKWAY") Then .HWFAKWAY = rs("HWFAKWAY") '�i�v�e���������@
            If fldNameExist("HWFAKHNM") Then .HWFAKHNM = rs("HWFAKHNM") '�i�v�e�������p�x�Q��
            If fldNameExist("HWFAKHNN") Then .HWFAKHNN = rs("HWFAKHNN") '�i�v�e�������p�x�Q��
            If fldNameExist("HWFAKHNH") Then .HWFAKHNH = rs("HWFAKHNH") '�i�v�e�������p�x�Q��
            If fldNameExist("HWFAKHNU") Then .HWFAKHNU = rs("HWFAKHNU") '�i�v�e�������p�x�Q�E
            If fldNameExist("HWFASDEV") Then .HWFASDEV = fncNullCheck(rs("HWFASDEV")) '�i�v�e���W���΍�
            If fldNameExist("HWFAAMIN") Then .HWFAAMIN = fncNullCheck(rs("HWFAAMIN")) '�i�v�e�����ω���
            If fldNameExist("HWFAAMAX") Then .HWFAAMAX = fncNullCheck(rs("HWFAAMAX")) '�i�v�e�����Ϗ��
            If fldNameExist("HWFAMBNP") Then .HWFAMBNP = fncNullCheck(rs("HWFAMBNP")) '�i�v�e���ʓ����z
            If fldNameExist("HWFAMCAL") Then .HWFAMCAL = rs("HWFAMCAL") '�i�v�e���ʓ��v�Z
            If fldNameExist("HWFALTBP") Then .HWFALTBP = fncNullCheck(rs("HWFALTBP")) '�i�v�e���k�s���z
            If fldNameExist("HWFALTCL") Then .HWFALTCL = rs("HWFALTCL") '�i�v�e���k�s�v�Z
            If fldNameExist("HWFALTRA") Then .HWFALTRA = fncNullCheck(rs("HWFALTRA")) '�i�v�e���k�s�͈�
            If fldNameExist("HWFAMRAN") Then .HWFAMRAN = fncNullCheck(rs("HWFAMRAN")) '�i�v�e���ʓ��͈�
            If fldNameExist("HWFDIVS") Then .HWFDIVS = fncNullCheck(rs("HWFDIVS")) '�i�v�e������
            If fldNameExist("HWFAKBSH") Then .HWFAKBSH = rs("HWFAKBSH") '�i�v�e���U�敪����ʒu�Q��
            If fldNameExist("HWFAKBST") Then .HWFAKBST = rs("HWFAKBST") '�i�v�e���U�敪����ʒu�Q�_
            If fldNameExist("HWFAKBSI") Then .HWFAKBSI = rs("HWFAKBSI") '�i�v�e���U�敪����ʒu�Q��
            If fldNameExist("HWFAKBHT") Then .HWFAKBHT = rs("HWFAKBHT") '�i�v�e���U�敪�ۏؕ��@�Q��
            If fldNameExist("HWFAKBHS") Then .HWFAKBHS = rs("HWFAKBHS") '�i�v�e���U�敪�ۏؕ��@�Q��
            If fldNameExist("HWFWFORM") Then .HWFWFORM = rs("HWFWFORM") '�i�v�e�E�F�[�n�`��
            If fldNameExist("HWFD1CEN") Then .HWFD1CEN = fncNullCheck(rs("HWFD1CEN")) '�i�v�e���a�P���S
            If fldNameExist("HWFD1MIN") Then .HWFD1MIN = fncNullCheck(rs("HWFD1MIN")) '�i�v�e���a�P����
            If fldNameExist("HWFD1MAX") Then .HWFD1MAX = fncNullCheck(rs("HWFD1MAX")) '�i�v�e���a�P���
            If fldNameExist("HWFD2CEN") Then .HWFD2CEN = fncNullCheck(rs("HWFD2CEN")) '�i�v�e���a�Q���S
            If fldNameExist("HWFD2MIN") Then .HWFD2MIN = fncNullCheck(rs("HWFD2MIN")) '�i�v�e���a�Q����
            If fldNameExist("HWFD2MAX") Then .HWFD2MAX = fncNullCheck(rs("HWFD2MAX")) '�i�v�e���a�Q���
            If fldNameExist("HWFDKHNM") Then .HWFDKHNM = rs("HWFDKHNM") '�i�v�e���a�����p�x�Q��
            If fldNameExist("HWFDKHNN") Then .HWFDKHNN = rs("HWFDKHNN") '�i�v�e���a�����p�x�Q��
            If fldNameExist("HWFDKHNH") Then .HWFDKHNH = rs("HWFDKHNH") '�i�v�e���a�����p�x�Q��
            If fldNameExist("HWFDKHNU") Then .HWFDKHNU = rs("HWFDKHNU") '�i�v�e���a�����p�x�Q�E
            If fldNameExist("HWFLPMNP") Then .HWFLPMNP = fncNullCheck(rs("HWFLPMNP")) '�i�v�e�k�o���ŏ����H��
            If fldNameExist("HWFSGMNP") Then .HWFSGMNP = fncNullCheck(rs("HWFSGMNP")) '�i�v�e�r�f���ŏ����H��
            If fldNameExist("HWFETMNP") Then .HWFETMNP = fncNullCheck(rs("HWFETMNP")) '�i�v�e�d�s���ŏ����H��
            If fldNameExist("HWFMPMNP") Then .HWFMPMNP = fncNullCheck(rs("HWFMPMNP")) '�i�v�e�l�o���ŏ����H��
            If fldNameExist("HWFLPKS1") Then .HWFLPKS1 = rs("HWFLPKS1") '�i�v�e�k�o�����ގ�P
            If fldNameExist("HWFLPKS2") Then .HWFLPKS2 = rs("HWFLPKS2") '�i�v�e�k�o�����ގ�Q
            If fldNameExist("HWFLPKZ1") Then .HWFLPKZ1 = rs("HWFLPKZ1") '�i�v�e�k�o�����ޗ��x��P
            If fldNameExist("HWFLPKZ2") Then .HWFLPKZ2 = rs("HWFLPKZ2") '�i�v�e�k�o�����ޗ��x��Q
            If fldNameExist("IFKBN") Then .IFKBN = rs("IFKBN") '�h�^�e�敪
            If fldNameExist("SYORIKBN") Then .SYORIKBN = rs("SYORIKBN") '�����敪
            If fldNameExist("SPECRRNO") Then .SPECRRNO = rs("SPECRRNO") '�d�l�o�^�˗��ԍ�
            If fldNameExist("SXLMCNO") Then .SXLMCNO = rs("SXLMCNO") '�r�w�k��������ԍ�
            If fldNameExist("WFMCNO") Then .WFMCNO = rs("WFMCNO") '�v�e��������ԍ�
            If fldNameExist("STAFFID") Then .StaffID = rs("STAFFID") '�Ј�ID
            If fldNameExist("REGDATE") Then .REGDATE = rs("REGDATE") '�o�^���t
            If fldNameExist("UPDDATE") Then .UPDDATE = rs("UPDDATE") '�X�V���t
            If fldNameExist("SENDFLAG") Then .SENDFLAG = rs("SENDFLAG") '���M�t���O
            If fldNameExist("SENDDATE") Then .SENDDATE = rs("SENDDATE") '���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME021 = FUNCTION_RETURN_SUCCESS

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

'�T�v      :�e�[�u���uTBCME022�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :records()     ,O  ,typ_TBCME022    ,���o���R�[�h
'          :formID        ,I  ,String          ,�g�p�t�H�[��ID
'          :sqlOrder      ,I  ,tFullHinban     ,���o�i�ԁi�z��j
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,���o�̐���
'����      :�����L�[�́A�HINBAN�+�uMNOREVNO�v+�uFACTORY�v+�uOPECOND�v�̕�����Ƃ���
'����      :05/03/01 ooba
Public Function DBDRV_GetTBCME022(records() As typ_TBCME022, formID$, HIN() As tFullHinban) As FUNCTION_RETURN

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
    gErr.Push "s_kensa_SQL.bas -- Function DBDRV_GetTBCME022"

    Select Case formID
        Case "f_cmbc026_1"           '�uGD���ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGWFSNO, HMGWFSNE, HWFCDIR, " & _
              "HWFCSCEN, HWFCSMIN, HWFCSMAX, HWFCSDIS, HWFCSDIR, HWFCKWAY, HWFCKHNM, HWFCKHNN, HWFCKHNH, " & _
              "HWFCKHNU, HWFCTDIR, HWFCTCEN, HWFCTMIN, HWFCTMAX, HWFCYDIR, HWFCYCEN, HWFCYMIN, HWFCYMAX, " & _
              "HWFKPTNN, HWFOFPKM, HWFOFPKN, HWFOFPKH, HWFOFPKU, HWFOFLKM, HWFOFLKN, HWFOFLKH, HWFOFLKU, " & _
              "HWFOF1PD, HWFOF1PN, HWFOF1PX, HWFOF1PW, HWFOF1LC, HWFOF1LN, HWFOF1LX, HWFOF1RF, HWFOFRRC, " & _
              "HWFOFRRN, HWFOFRRX, HWFOFRLC, HWFOFRLN, HWFOFRLX, HWFOF1DC, HWFOF1DN, HWFOF1DX, HWFZFORM, "
            sqlBase = sqlBase & "HWFD3CEN, HWFD3MIN, HWFD3MAX, HWFDFKJ, HWFDFKHM, HWFDFKHN, HWFDFKHH, " & _
              "HWFDFKHU, HWFDPDRC, HWFDPACN, HWFDPAMN, HWFDPAMX, HWFDPDIR, HWFDPMIN, HWFDPMAX, HWFDPKWY, " & _
              "HWFDPKHM, HWFDPKHB, HWFDPKHH, HWFDPKHU, HWFDACEN, HWFDAMIN, HWFDAMAX, HWFDWCEN, HWFDWMIN, " & _
              "HWFDWMAX, HWFDDCEN, HWFDDMIN, HWFDDMAX, HWFDBRCN, HWFDBRMN, HWFDBRMX, HWFDRRCN, HWFDRRMN, " & _
              "HWFDRRMX, HWFDLRCN, HWFDLRMN, HWFDLRMX, IFKBN, SYORIKBN, SPECRRNO, SXLMCNO, WFMCNO, " & _
              "STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        '�ǉ� 2005/06/15 ffc)tanabe start
        Case "f_cmec067_1"           '�uSPV���юQ�Ɓv
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HWFCDIR "
        '�ǉ� 2005/06/15 ffc)tanabe end

    End Select
       
    sqlBase = sqlBase & "From TBCME022"
    
    '''SQL��Where���쐬
    For i = 0 To UBound(HIN)
        With HIN(i)
            key = key & "'" & .hinban & Format(.mnorevno, "00000") & .factory & .opecond & "'"
            If i <> UBound(HIN) Then
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
        DBDRV_GetTBCME022 = FUNCTION_RETURN_FAILURE
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
            If fldNameExist("HINBAN") Then .hinban = rs("HINBAN") '�i��
            If fldNameExist("MNOREVNO") Then .mnorevno = rs("MNOREVNO") '���i�ԍ������ԍ�
            If fldNameExist("FACTORY") Then .factory = rs("FACTORY") '�H��
            If fldNameExist("OPECOND") Then .opecond = rs("OPECOND") '���Ə���
            If fldNameExist("HMGSTRRNO") Then .HMGSTRRNO = rs("HMGSTRRNO") '�i�Ǘ��d�l�o�^�˗��ԍ�
            If fldNameExist("HMGSTFNO") Then .HMGSTFNO = rs("HMGSTFNO") '�i�Ǘ��Ј��m��
            If fldNameExist("HMGWFSNO") Then .HMGWFSNO = rs("HMGWFSNO") '�i�Ǘ��v�e���i�ԍ�
            If fldNameExist("HMGWFSNE") Then .HMGWFSNE = fncNullCheck(rs("HMGWFSNE")) '�i�Ǘ��v�e���i�ԍ��}��
            If fldNameExist("HWFCDIR") Then .HWFCDIR = rs("HWFCDIR") '�i�v�e�����ʕ���
            If fldNameExist("HWFCSCEN") Then .HWFCSCEN = fncNullCheck(rs("HWFCSCEN")) '�i�v�e�����ʌX���S
            If fldNameExist("HWFCSMIN") Then .HWFCSMIN = fncNullCheck(rs("HWFCSMIN")) '�i�v�e�����ʌX����
            If fldNameExist("HWFCSMAX") Then .HWFCSMAX = fncNullCheck(rs("HWFCSMAX")) '�i�v�e�����ʌX���
            If fldNameExist("HWFCSDIS") Then .HWFCSDIS = rs("HWFCSDIS") '�i�v�e�����ʌX���ʎw��
            If fldNameExist("HWFCSDIR") Then .HWFCSDIR = rs("HWFCSDIR") '�i�v�e�����ʌX����
            If fldNameExist("HWFCKWAY") Then .HWFCKWAY = rs("HWFCKWAY") '�i�v�e�����ʌ������@
            If fldNameExist("HWFCKHNM") Then .HWFCKHNM = rs("HWFCKHNM") '�i�v�e�����ʌ����p�x�Q��
            If fldNameExist("HWFCKHNN") Then .HWFCKHNN = rs("HWFCKHNN") '�i�v�e�����ʌ����p�x�Q��
            If fldNameExist("HWFCKHNH") Then .HWFCKHNH = rs("HWFCKHNH") '�i�v�e�����ʌ����p�x�Q��
            If fldNameExist("HWFCKHNU") Then .HWFCKHNU = rs("HWFCKHNU") '�i�v�e�����ʌ����p�x�Q�E
            If fldNameExist("HWFCTDIR") Then .HWFCTDIR = rs("HWFCTDIR") '�i�v�e�����ʌX�c����
            If fldNameExist("HWFCTCEN") Then .HWFCTCEN = fncNullCheck(rs("HWFCTCEN")) '�i�v�e�����ʌX�c���S
            If fldNameExist("HWFCTMIN") Then .HWFCTMIN = fncNullCheck(rs("HWFCTMIN")) '�i�v�e�����ʌX�c����
            If fldNameExist("HWFCTMAX") Then .HWFCTMAX = fncNullCheck(rs("HWFCTMAX")) '�i�v�e�����ʌX�c���
            If fldNameExist("HWFCYDIR") Then .HWFCYDIR = rs("HWFCYDIR") '�i�v�e�����ʌX������
            If fldNameExist("HWFCYCEN") Then .HWFCYCEN = fncNullCheck(rs("HWFCYCEN")) '�i�v�e�����ʌX�����S
            If fldNameExist("HWFCYMIN") Then .HWFCYMIN = fncNullCheck(rs("HWFCYMIN")) '�i�v�e�����ʌX������
            If fldNameExist("HWFCYMAX") Then .HWFCYMAX = fncNullCheck(rs("HWFCYMAX")) '�i�v�e�����ʌX�����
            If fldNameExist("HWFKPTNN") Then .HWFKPTNN = rs("HWFKPTNN") '�i�v�e�����p�^����
            If fldNameExist("HWFOFPKM") Then .HWFOFPKM = rs("HWFOFPKM") '�i�v�e�n�e�ʒu�����p�x�Q��
            If fldNameExist("HWFOFPKN") Then .HWFOFPKN = rs("HWFOFPKN") '�i�v�e�n�e�ʒu�����p�x�Q��
            If fldNameExist("HWFOFPKH") Then .HWFOFPKH = rs("HWFOFPKH") '�i�v�e�n�e�ʒu�����p�x�Q��
            If fldNameExist("HWFOFPKU") Then .HWFOFPKU = rs("HWFOFPKU") '�i�v�e�n�e�ʒu�����p�x�Q�E
            If fldNameExist("HWFOFLKM") Then .HWFOFLKM = rs("HWFOFLKM") '�i�v�e�n�e�������p�x�Q��
            If fldNameExist("HWFOFLKN") Then .HWFOFLKN = rs("HWFOFLKN") '�i�v�e�n�e�������p�x�Q��
            If fldNameExist("HWFOFLKH") Then .HWFOFLKH = rs("HWFOFLKH") '�i�v�e�n�e�������p�x�Q��
            If fldNameExist("HWFOFLKU") Then .HWFOFLKU = rs("HWFOFLKU") '�i�v�e�n�e�������p�x�Q�E
            If fldNameExist("HWFOF1PD") Then .HWFOF1PD = rs("HWFOF1PD") '�i�v�e�n�e�P�ʒu����
            If fldNameExist("HWFOF1PN") Then .HWFOF1PN = fncNullCheck(rs("HWFOF1PN")) '�i�v�e�n�e�P�ʒu����
            If fldNameExist("HWFOF1PX") Then .HWFOF1PX = fncNullCheck(rs("HWFOF1PX")) '�i�v�e�n�e�P�ʒu���
            If fldNameExist("HWFOF1PW") Then .HWFOF1PW = rs("HWFOF1PW") '�i�v�e�n�e�P�ʒu�������@
            If fldNameExist("HWFOF1LC") Then .HWFOF1LC = fncNullCheck(rs("HWFOF1LC")) '�i�v�e�n�e�P�����S
            If fldNameExist("HWFOF1LN") Then .HWFOF1LN = fncNullCheck(rs("HWFOF1LN")) '�i�v�e�n�e�P������
            If fldNameExist("HWFOF1LX") Then .HWFOF1LX = fncNullCheck(rs("HWFOF1LX")) '�i�v�e�n�e�P�����
            If fldNameExist("HWFOF1RF") Then .HWFOF1RF = rs("HWFOF1RF") '�i�v�e�n�e�P���[�q�`��
            If fldNameExist("HWFOFRRC") Then .HWFOFRRC = fncNullCheck(rs("HWFOFRRC")) '�i�v�e�n�e���[�q�E���S
            If fldNameExist("HWFOFRRN") Then .HWFOFRRN = fncNullCheck(rs("HWFOFRRN")) '�i�v�e�n�e���[�q�E����
            If fldNameExist("HWFOFRRX") Then .HWFOFRRX = fncNullCheck(rs("HWFOFRRX")) '�i�v�e�n�e���[�q�E���
            If fldNameExist("HWFOFRLC") Then .HWFOFRLC = fncNullCheck(rs("HWFOFRLC")) '�i�v�e�n�e���[�q�����S
            If fldNameExist("HWFOFRLN") Then .HWFOFRLN = fncNullCheck(rs("HWFOFRLN")) '�i�v�e�n�e���[�q������
            If fldNameExist("HWFOFRLX") Then .HWFOFRLX = fncNullCheck(rs("HWFOFRLX")) '�i�v�e�n�e���[�q�����
            If fldNameExist("HWFOF1DC") Then .HWFOF1DC = fncNullCheck(rs("HWFOF1DC")) '�i�v�e�n�e�P���a���S
            If fldNameExist("HWFOF1DN") Then .HWFOF1DN = fncNullCheck(rs("HWFOF1DN")) '�i�v�e�n�e�P���a����
            If fldNameExist("HWFOF1DX") Then .HWFOF1DX = fncNullCheck(rs("HWFOF1DX")) '�i�v�e�n�e�P���a���
            If fldNameExist("HWFZFORM") Then .HWFZFORM = rs("HWFZFORM") '�i�v�e�ޗ��`��
            If fldNameExist("HWFD3CEN") Then .HWFD3CEN = fncNullCheck(rs("HWFD3CEN")) '�i�v�e���a�R���S
            If fldNameExist("HWFD3MIN") Then .HWFD3MIN = fncNullCheck(rs("HWFD3MIN")) '�i�v�e���a�R����
            If fldNameExist("HWFD3MAX") Then .HWFD3MAX = fncNullCheck(rs("HWFD3MAX")) '�i�v�e���a�R���
            If fldNameExist("HWFDFKJ") Then .HWFDFKJ = rs("HWFDFKJ") '�i�v�e�a�`��
            If fldNameExist("HWFDFKHM") Then .HWFDFKHM = rs("HWFDFKHM") '�i�v�e�a�`�󌟍��p�x�Q��
            If fldNameExist("HWFDFKHN") Then .HWFDFKHN = rs("HWFDFKHN") '�i�v�e�a�`�󌟍��p�x�Q��
            If fldNameExist("HWFDFKHH") Then .HWFDFKHH = rs("HWFDFKHH") '�i�v�e�a�`�󌟍��p�x�Q��
            If fldNameExist("HWFDFKHU") Then .HWFDFKHU = rs("HWFDFKHU") '�i�v�e�a�`�󌟍��p�x�Q�E
            If fldNameExist("HWFDPDRC") Then .HWFDPDRC = rs("HWFDPDRC") '�i�v�e�a�ʒu����
            If fldNameExist("HWFDPACN") Then .HWFDPACN = fncNullCheck(rs("HWFDPACN")) '�i�v�e�a�ʒu�p�x���S
            If fldNameExist("HWFDPAMN") Then .HWFDPAMN = fncNullCheck(rs("HWFDPAMN")) '�i�v�e�a�ʒu�p�x����
            If fldNameExist("HWFDPAMX") Then .HWFDPAMX = fncNullCheck(rs("HWFDPAMX")) '�i�v�e�a�ʒu�p�x���
            If fldNameExist("HWFDPDIR") Then .HWFDPDIR = rs("HWFDPDIR") '�i�v�e�a�ʒu����
            If fldNameExist("HWFDPMIN") Then .HWFDPMIN = fncNullCheck(rs("HWFDPMIN")) '�i�v�e�a�ʒu����
            If fldNameExist("HWFDPMAX") Then .HWFDPMAX = fncNullCheck(rs("HWFDPMAX")) '�i�v�e�a�ʒu���
            If fldNameExist("HWFDPKWY") Then .HWFDPKWY = rs("HWFDPKWY") '�i�v�e�a�ʒu�������@
            If fldNameExist("HWFDPKHM") Then .HWFDPKHM = rs("HWFDPKHM") '�i�v�e�a�ʒu�����p�x�Q��
            If fldNameExist("HWFDPKHB") Then .HWFDPKHB = rs("HWFDPKHB") '�i�v�e�a�ʒu�����p�x�Q��
            If fldNameExist("HWFDPKHH") Then .HWFDPKHH = rs("HWFDPKHH") '�i�v�e�a�ʒu�����p�x�Q��
            If fldNameExist("HWFDPKHU") Then .HWFDPKHU = rs("HWFDPKHU") '�i�v�e�a�ʒu�����p�x�Q�E
            If fldNameExist("HWFDACEN") Then .HWFDACEN = fncNullCheck(rs("HWFDACEN")) '�i�v�e�a�p�x���S
            If fldNameExist("HWFDAMIN") Then .HWFDAMIN = fncNullCheck(rs("HWFDAMIN")) '�i�v�e�a�p�x����
            If fldNameExist("HWFDAMAX") Then .HWFDAMAX = fncNullCheck(rs("HWFDAMAX")) '�i�v�e�a�p�x���
            If fldNameExist("HWFDWCEN") Then .HWFDWCEN = fncNullCheck(rs("HWFDWCEN")) '�i�v�e�a�В��S
            If fldNameExist("HWFDWMIN") Then .HWFDWMIN = fncNullCheck(rs("HWFDWMIN")) '�i�v�e�a�Љ���
            If fldNameExist("HWFDWMAX") Then .HWFDWMAX = fncNullCheck(rs("HWFDWMAX")) '�i�v�e�a�Џ��
            If fldNameExist("HWFDDCEN") Then .HWFDDCEN = fncNullCheck(rs("HWFDDCEN")) '�i�v�e�a�[���S
            If fldNameExist("HWFDDMIN") Then .HWFDDMIN = fncNullCheck(rs("HWFDDMIN")) '�i�v�e�a�[����
            If fldNameExist("HWFDDMAX") Then .HWFDDMAX = fncNullCheck(rs("HWFDDMAX")) '�i�v�e�a�[���
            If fldNameExist("HWFDBRCN") Then .HWFDBRCN = fncNullCheck(rs("HWFDBRCN")) '�i�v�e�a��q���S
            If fldNameExist("HWFDBRMN") Then .HWFDBRMN = fncNullCheck(rs("HWFDBRMN")) '�i�v�e�a��q����
            If fldNameExist("HWFDBRMX") Then .HWFDBRMX = fncNullCheck(rs("HWFDBRMX")) '�i�v�e�a��q���
            If fldNameExist("HWFDRRCN") Then .HWFDRRCN = fncNullCheck(rs("HWFDRRCN")) '�i�v�e�a�E�q���S
            If fldNameExist("HWFDRRMN") Then .HWFDRRMN = fncNullCheck(rs("HWFDRRMN")) '�i�v�e�a�E�q����
            If fldNameExist("HWFDRRMX") Then .HWFDRRMX = fncNullCheck(rs("HWFDRRMX")) '�i�v�e�a�E�q���
            If fldNameExist("HWFDLRCN") Then .HWFDLRCN = fncNullCheck(rs("HWFDLRCN")) '�i�v�e�a���q���S
            If fldNameExist("HWFDLRMN") Then .HWFDLRMN = fncNullCheck(rs("HWFDLRMN")) '�i�v�e�a���q����
            If fldNameExist("HWFDLRMX") Then .HWFDLRMX = fncNullCheck(rs("HWFDLRMX")) '�i�v�e�a���q���
            If fldNameExist("IFKBN") Then .IFKBN = rs("IFKBN") '�h�^�e�敪
            If fldNameExist("SYORIKBN") Then .SYORIKBN = rs("SYORIKBN") '�����敪
            If fldNameExist("SPECRRNO") Then .SPECRRNO = rs("SPECRRNO") '�d�l�o�^�˗��ԍ�
            If fldNameExist("SXLMCNO") Then .SXLMCNO = rs("SXLMCNO") '�r�w�k��������ԍ�
            If fldNameExist("WFMCNO") Then .WFMCNO = rs("WFMCNO") '�v�e��������ԍ�
            If fldNameExist("STAFFID") Then .StaffID = rs("STAFFID") '�Ј�ID
            If fldNameExist("REGDATE") Then .REGDATE = rs("REGDATE") '�o�^���t
            If fldNameExist("UPDDATE") Then .UPDDATE = rs("UPDDATE") '�X�V���t
            If fldNameExist("SENDFLAG") Then .SENDFLAG = rs("SENDFLAG") '���M�t���O
            If fldNameExist("SENDDATE") Then .SENDDATE = rs("SENDDATE") '���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME022 = FUNCTION_RETURN_SUCCESS

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

'�T�v      :�e�[�u���uTBCME026�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :records()     ,O  ,typ_TBCME026    ,���o���R�[�h
'          :formID        ,I  ,String          ,�g�p�t�H�[��ID
'          :sqlOrder      ,I  ,tFullHinban     ,���o�i�ԁi�z��j
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,���o�̐���
'����      :�����L�[�́A�HINBAN�+�uMNOREVNO�v+�uFACTORY�v+�uOPECOND�v�̕�����Ƃ���
'����      :05/03/01 ooba
Public Function DBDRV_GetTBCME026(records() As typ_TBCME026, formID$, HIN() As tFullHinban) As FUNCTION_RETURN

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
    gErr.Push "s_kensa_SQL.bas -- Function DBDRV_GetTBCME026"

    Select Case formID
        Case "f_cmbc026_1"           '�uGD���ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGWFSNO, HMGWFSNE, HWFBDOMN, HWFBDOMX, " & _
              "HWFBDOSH, HWFBDOST, HWFBDOSR, HWFBDOHT, HWFBDOHS, HWFBDOSZ, HWFBDONS, HWFBDOKM, HWFBDOKN, HWFBDOKH, " & _
              "HWFBDOKU, HWFBDOET, HWFBDSMN, HWFBDSMX, HWFBDSSH, HWFBDSST, HWFBDSSR, HWFBDSHT, HWFBDSHS, HWFBDSSZ, " & _
              "HWFBDSNS, HWFBDSKM, HWFBDSKN, HWFBDSKH, HWFBDSKU, HWFBDSET, HWFRNFMX, HWFRNFSH, HWFRNFST, HWFRNFSI, " & _
              "HWFRNFKW, HWFRNFZA, HWFRNBMX, HWFRNBSH, HWFRNBST, HWFRNBSI, HWFRNBKW, HWFRNBZA, HWFDENKU, HWFDENMX, " & _
              "HWFDENMN, HWFDENHT, HWFDENHS, HWFDVDKU, HWFDVDMX, HWFDVDMN, HWFDVDHT, HWFDVDHS, HWFLDLKU, HWFLDLMX, " & _
              "HWFLDLMN, HWFLDLHT, HWFLDLHS, HWFGDSPH, HWFGDSPT, HWFGDSPR, HWFGDSZY, HWFGDZAR, HWFGDKHM, HWFGDKHN, "
            sqlBase = sqlBase & "HWFGDKHH, HWFGDKHU, HWFDSOKE, HWFDSOMX, HWFDSOMN, HWFDSOAX, HWFDSOAN, HWFDSOHT, HWFDSOHS, HWFDSOKM, " & _
              "HWFDSOKN, HWFDSOKH, HWFDSOKU, HWFNTPUM, HWFNTPK1, HWFNTPP1, HWFNTPS1, HWFNTPK2, HWFNTPP2, HWFNTPS2, " & _
              "HWFNTPK3, HWFNTPP3, HWFNTPS3, HWFNTPZA, HWFNTPHT, HWFNTPHS, HWFNTPKM, HWFNTPKN, HWFNTPKH, HWFNTPKU, " & _
              "HWFCRSSK, HWFMDCEN, HWFMDMAX, HWFMDMIN, HWFMDSPH, HWFMDSPT, HWFMDSPI, HWFMDHWT, HWFMDHWS, HWFMDKHM, " & _
              "HWFMDKHN, HWFMDKHH, HWFMDKHU, IFKBN, SYORIKBN, SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, " & _
              "SENDFLAG, SENDDATE, HWFDVDMXN, HWFDVDMNN, HWFDSONWY, HWFMSUMX, HWFMSUZY, HWFMSUKW, HWFMSUSZ, " & _
              "HWFNP1AR, HWFNP1MAX, HWFNP2AR, HWFNP2MAX "
    End Select
       
    sqlBase = sqlBase & "From TBCME026"
    
    '''SQL��Where���쐬
    For i = 0 To UBound(HIN)
        With HIN(i)
            key = key & "'" & .hinban & Format(.mnorevno, "00000") & .factory & .opecond & "'"
            If i <> UBound(HIN) Then
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
        DBDRV_GetTBCME026 = FUNCTION_RETURN_FAILURE
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
            If fldNameExist("HINBAN") Then .hinban = rs("HINBAN") '�i��
            If fldNameExist("MNOREVNO") Then .mnorevno = rs("MNOREVNO") '���i�ԍ������ԍ�
            If fldNameExist("FACTORY") Then .factory = rs("FACTORY") '�H��
            If fldNameExist("OPECOND") Then .opecond = rs("OPECOND") '���Ə���
            If fldNameExist("HMGSTRRNO") Then .HMGSTRRNO = rs("HMGSTRRNO") '�i�Ǘ��d�l�o�^�˗��ԍ�
            If fldNameExist("HMGSTFNO") Then .HMGSTFNO = rs("HMGSTFNO") '�i�Ǘ��Ј��m��
            If fldNameExist("HMGWFSNO") Then .HMGWFSNO = rs("HMGWFSNO") '�i�Ǘ��v�e���i�ԍ�
            If fldNameExist("HMGWFSNE") Then .HMGWFSNE = fncNullCheck(rs("HMGWFSNE")) '�i�Ǘ��v�e���i�ԍ��}��
            If fldNameExist("HWFBDOMN") Then .HWFBDOMN = fncNullCheck(rs("HWFBDOMN")) '�i�v�e�a�c�n�r�e����
            If fldNameExist("HWFBDOMX") Then .HWFBDOMX = fncNullCheck(rs("HWFBDOMX")) '�i�v�e�a�c�n�r�e���
            If fldNameExist("HWFBDOSH") Then .HWFBDOSH = rs("HWFBDOSH") '�i�v�e�a�c�n�r�e����ʒu�Q��
            If fldNameExist("HWFBDOST") Then .HWFBDOST = rs("HWFBDOST") '�i�v�e�a�c�n�r�e����ʒu�Q�_
            If fldNameExist("HWFBDOSR") Then .HWFBDOSR = rs("HWFBDOSR") '�i�v�e�a�c�n�r�e����ʒu�Q��
            If fldNameExist("HWFBDOHT") Then .HWFBDOHT = rs("HWFBDOHT") '�i�v�e�a�c�n�r�e�ۏؕ��@�Q��
            If fldNameExist("HWFBDOHS") Then .HWFBDOHS = rs("HWFBDOHS") '�i�v�e�a�c�n�r�e�ۏؕ��@�Q��
            If fldNameExist("HWFBDOSZ") Then .HWFBDOSZ = rs("HWFBDOSZ") '�i�v�e�a�c�n�r�e�������
            If fldNameExist("HWFBDONS") Then .HWFBDONS = rs("HWFBDONS") '�i�v�e�a�c�n�r�e�M�����@
            If fldNameExist("HWFBDOKM") Then .HWFBDOKM = rs("HWFBDOKM") '�i�v�e�a�c�n�r�e�����p�x�Q��
            If fldNameExist("HWFBDOKN") Then .HWFBDOKN = rs("HWFBDOKN") '�i�v�e�a�c�n�r�e�����p�x�Q��
            If fldNameExist("HWFBDOKH") Then .HWFBDOKH = rs("HWFBDOKH") '�i�v�e�a�c�n�r�e�����p�x�Q��
            If fldNameExist("HWFBDOKU") Then .HWFBDOKU = rs("HWFBDOKU") '�i�v�e�a�c�n�r�e�����p�x�Q�E
            If fldNameExist("HWFBDOET") Then .HWFBDOET = fncNullCheck(rs("HWFBDOET")) '�i�v�e�a�c�n�r�e�I���d�s��
            If fldNameExist("HWFBDSMN") Then .HWFBDSMN = fncNullCheck(rs("HWFBDSMN")) '�i�v�e�a�c�r�s�Չ���
            If fldNameExist("HWFBDSMX") Then .HWFBDSMX = fncNullCheck(rs("HWFBDSMX")) '�i�v�e�a�c�r�s�Տ��
            If fldNameExist("HWFBDSSH") Then .HWFBDSSH = rs("HWFBDSSH") '�i�v�e�a�c�r�s�Ց���ʒu�Q��
            If fldNameExist("HWFBDSST") Then .HWFBDSST = rs("HWFBDSST") '�i�v�e�a�c�r�s�Ց���ʒu�Q�_
            If fldNameExist("HWFBDSSR") Then .HWFBDSSR = rs("HWFBDSSR") '�i�v�e�a�c�r�s�Ց���ʒu�Q��
            If fldNameExist("HWFBDSHT") Then .HWFBDSHT = rs("HWFBDSHT") '�i�v�e�a�c�r�s�Օۏؕ��@�Q��
            If fldNameExist("HWFBDSHS") Then .HWFBDSHS = rs("HWFBDSHS") '�i�v�e�a�c�r�s�Օۏؕ��@�Q��
            If fldNameExist("HWFBDSSZ") Then .HWFBDSSZ = rs("HWFBDSSZ") '�i�v�e�a�c�r�s�Ց������
            If fldNameExist("HWFBDSNS") Then .HWFBDSNS = rs("HWFBDSNS") '�i�v�e�a�c�r�s�ՔM�����@
            If fldNameExist("HWFBDSKM") Then .HWFBDSKM = rs("HWFBDSKM") '�i�v�e�a�c�r�s�Ռ����p�x�Q��
            If fldNameExist("HWFBDSKN") Then .HWFBDSKN = rs("HWFBDSKN") '�i�v�e�a�c�r�s�Ռ����p�x�Q��
            If fldNameExist("HWFBDSKH") Then .HWFBDSKH = rs("HWFBDSKH") '�i�v�e�a�c�r�s�Ռ����p�x�Q��
            If fldNameExist("HWFBDSKU") Then .HWFBDSKU = rs("HWFBDSKU") '�i�v�e�a�c�r�s�Ռ����p�x�Q�E
            If fldNameExist("HWFBDSET") Then .HWFBDSET = fncNullCheck(rs("HWFBDSET")) '�i�v�e�a�c�r�s�ՑI���d�s��
            If fldNameExist("HWFRNFMX") Then .HWFRNFMX = fncNullCheck(rs("HWFRNFMX")) '�i�v�e���t�l�X�\���
            If fldNameExist("HWFRNFSH") Then .HWFRNFSH = rs("HWFRNFSH") '�i�v�e���t�l�X�\����ʒu�Q��
            If fldNameExist("HWFRNFST") Then .HWFRNFST = rs("HWFRNFST") '�i�v�e���t�l�X�\����ʒu�Q�_
            If fldNameExist("HWFRNFSI") Then .HWFRNFSI = rs("HWFRNFSI") '�i�v�e���t�l�X�\����ʒu�Q��
            If fldNameExist("HWFRNFKW") Then .HWFRNFKW = rs("HWFRNFKW") '�i�v�e���t�l�X�\�������@
            If fldNameExist("HWFRNFZA") Then .HWFRNFZA = fncNullCheck(rs("HWFRNFZA")) '�i�v�e���t�l�X�\���O�̈�
            If fldNameExist("HWFRNBMX") Then .HWFRNBMX = fncNullCheck(rs("HWFRNBMX")) '�i�v�e���t�l�X�����
            If fldNameExist("HWFRNBSH") Then .HWFRNBSH = rs("HWFRNBSH") '�i�v�e���t�l�X������ʒu�Q��
            If fldNameExist("HWFRNBST") Then .HWFRNBST = rs("HWFRNBST") '�i�v�e���t�l�X������ʒu�Q�_
            If fldNameExist("HWFRNBSI") Then .HWFRNBSI = rs("HWFRNBSI") '�i�v�e���t�l�X������ʒu�Q��
            If fldNameExist("HWFRNBKW") Then .HWFRNBKW = rs("HWFRNBKW") '�i�v�e���t�l�X���������@
            If fldNameExist("HWFRNBZA") Then .HWFRNBZA = fncNullCheck(rs("HWFRNBZA")) '�i�v�e���t�l�X�����O�̈�
            If fldNameExist("HWFDENKU") Then .HWFDENKU = rs("HWFDENKU") '�i�v�e�c���������L��
            If fldNameExist("HWFDENMX") Then .HWFDENMX = fncNullCheck(rs("HWFDENMX")) '�i�v�e�c�������
            If fldNameExist("HWFDENMN") Then .HWFDENMN = fncNullCheck(rs("HWFDENMN")) '�i�v�e�c��������
            If fldNameExist("HWFDENHT") Then .HWFDENHT = rs("HWFDENHT") '�i�v�e�c�����ۏؕ��@�Q��
            If fldNameExist("HWFDENHS") Then .HWFDENHS = rs("HWFDENHS") '�i�v�e�c�����ۏؕ��@�Q��
            If fldNameExist("HWFDVDKU") Then .HWFDVDKU = rs("HWFDVDKU") '�i�v�e�c�u�c�Q�����L��
            If fldNameExist("HWFDVDMX") Then .HWFDVDMX = fncNullCheck(rs("HWFDVDMX")) '�i�v�e�c�u�c�Q���
            If fldNameExist("HWFDVDMN") Then .HWFDVDMN = fncNullCheck(rs("HWFDVDMN")) '�i�v�e�c�u�c�Q����
            If fldNameExist("HWFDVDHT") Then .HWFDVDHT = rs("HWFDVDHT") '�i�v�e�c�u�c�Q�ۏؕ��@�Q��
            If fldNameExist("HWFDVDHS") Then .HWFDVDHS = rs("HWFDVDHS") '�i�v�e�c�u�c�Q�ۏؕ��@�Q��
            If fldNameExist("HWFLDLKU") Then .HWFLDLKU = rs("HWFLDLKU") '�i�v�e�k�^�c�k�����L��
            If fldNameExist("HWFLDLMX") Then .HWFLDLMX = fncNullCheck(rs("HWFLDLMX")) '�i�v�e�k�^�c�k���
            If fldNameExist("HWFLDLMN") Then .HWFLDLMN = fncNullCheck(rs("HWFLDLMN")) '�i�v�e�k�^�c�k����
            If fldNameExist("HWFLDLHT") Then .HWFLDLHT = rs("HWFLDLHT") '�i�v�e�k�^�c�k�ۏؕ��@�Q��
            If fldNameExist("HWFLDLHS") Then .HWFLDLHS = rs("HWFLDLHS") '�i�v�e�k�^�c�k�ۏؕ��@�Q��
            If fldNameExist("HWFGDSPH") Then .HWFGDSPH = rs("HWFGDSPH") '�i�v�e�f�c����ʒu�Q��
            If fldNameExist("HWFGDSPT") Then .HWFGDSPT = rs("HWFGDSPT") '�i�v�e�f�c����ʒu�Q�_
            If fldNameExist("HWFGDSPR") Then .HWFGDSPR = rs("HWFGDSPR") '�i�v�e�f�c����ʒu�Q��
            If fldNameExist("HWFGDSZY") Then .HWFGDSZY = rs("HWFGDSZY") '�i�v�e�f�c�������
            If fldNameExist("HWFGDZAR") Then .HWFGDZAR = fncNullCheck(rs("HWFGDZAR")) '�i�v�e�f�c���O�̈�
            If fldNameExist("HWFGDKHM") Then .HWFGDKHM = rs("HWFGDKHM") '�i�v�e�f�c�����p�x�Q��
            If fldNameExist("HWFGDKHN") Then .HWFGDKHN = rs("HWFGDKHN") '�i�v�e�f�c�����p�x�Q��
            If fldNameExist("HWFGDKHH") Then .HWFGDKHH = rs("HWFGDKHH") '�i�v�e�f�c�����p�x�Q��
            If fldNameExist("HWFGDKHU") Then .HWFGDKHU = rs("HWFGDKHU") '�i�v�e�f�c�����p�x�Q�E
            If fldNameExist("HWFDSOKE") Then .HWFDSOKE = rs("HWFDSOKE") '�i�v�e�c�r�n�c����
            If fldNameExist("HWFDSOMX") Then .HWFDSOMX = fncNullCheck(rs("HWFDSOMX")) '�i�v�e�c�r�n�c���
            If fldNameExist("HWFDSOMN") Then .HWFDSOMN = fncNullCheck(rs("HWFDSOMN")) '�i�v�e�c�r�n�c����
            If fldNameExist("HWFDSOAX") Then .HWFDSOAX = fncNullCheck(rs("HWFDSOAX")) '�i�v�e�c�r�n�c�̈���
            If fldNameExist("HWFDSOAN") Then .HWFDSOAN = fncNullCheck(rs("HWFDSOAN")) '�i�v�e�c�r�n�c�̈扺��
            If fldNameExist("HWFDSOHT") Then .HWFDSOHT = rs("HWFDSOHT") '�i�v�e�c�r�n�c�ۏؕ��@�Q��
            If fldNameExist("HWFDSOHS") Then .HWFDSOHS = rs("HWFDSOHS") '�i�v�e�c�r�n�c�ۏؕ��@�Q��
            If fldNameExist("HWFDSOKM") Then .HWFDSOKM = rs("HWFDSOKM") '�i�v�e�c�r�n�c�����p�x�Q��
            If fldNameExist("HWFDSOKN") Then .HWFDSOKN = rs("HWFDSOKN") '�i�v�e�c�r�n�c�����p�x�Q��
            If fldNameExist("HWFDSOKH") Then .HWFDSOKH = rs("HWFDSOKH") '�i�v�e�c�r�n�c�����p�x�Q��
            If fldNameExist("HWFDSOKU") Then .HWFDSOKU = rs("HWFDSOKU") '�i�v�e�c�r�n�c�����p�x�Q�E
            If fldNameExist("HWFNTPUM") Then .HWFNTPUM = rs("HWFNTPUM") '�i�v�e���R�i�m�g�|�L��
            If fldNameExist("HWFNTPK1") Then .HWFNTPK1 = fncNullCheck(rs("HWFNTPK1")) '�i�v�e���R�i�m�g�|�K�i�P
            If fldNameExist("HWFNTPP1") Then .HWFNTPP1 = fncNullCheck(rs("HWFNTPP1")) '�i�v�e���R�i�m�g�|�o�t�`�P
            If fldNameExist("HWFNTPS1") Then .HWFNTPS1 = fncNullCheck(rs("HWFNTPS1")) '�i�v�e���R�i�m�g�|�T�C�g�P
            If fldNameExist("HWFNTPK2") Then .HWFNTPK2 = fncNullCheck(rs("HWFNTPK2")) '�i�v�e���R�i�m�g�|�K�i�Q
            If fldNameExist("HWFNTPP2") Then .HWFNTPP2 = fncNullCheck(rs("HWFNTPP2")) '�i�v�e���R�i�m�g�|�o�t�`�Q
            If fldNameExist("HWFNTPS2") Then .HWFNTPS2 = fncNullCheck(rs("HWFNTPS2")) '�i�v�e���R�i�m�g�|�T�C�g�Q
            If fldNameExist("HWFNTPK3") Then .HWFNTPK3 = fncNullCheck(rs("HWFNTPK3")) '�i�v�e���R�i�m�g�|�K�i�R
            If fldNameExist("HWFNTPP3") Then .HWFNTPP3 = fncNullCheck(rs("HWFNTPP3")) '�i�v�e���R�i�m�g�|�o�t�`�R
            If fldNameExist("HWFNTPS3") Then .HWFNTPS3 = fncNullCheck(rs("HWFNTPS3")) '�i�v�e���R�i�m�g�|�T�C�g�R
            If fldNameExist("HWFNTPZA") Then .HWFNTPZA = fncNullCheck(rs("HWFNTPZA")) '�i�v�e���R�i�m�g�|���O�̈�
            If fldNameExist("HWFNTPHT") Then .HWFNTPHT = rs("HWFNTPHT") '�i�v�e���R�i�m�g�|�ۏؕ��@�Q��
            If fldNameExist("HWFNTPHS") Then .HWFNTPHS = rs("HWFNTPHS") '�i�v�e���R�i�m�g�|�ۏؕ��@�Q��
            If fldNameExist("HWFNTPKM") Then .HWFNTPKM = rs("HWFNTPKM") '�i�v�e���R�i�m�g�|�����p�x�Q��
            If fldNameExist("HWFNTPKN") Then .HWFNTPKN = rs("HWFNTPKN") '�i�v�e���R�i�m�g�|�����p�x�Q��
            If fldNameExist("HWFNTPKH") Then .HWFNTPKH = rs("HWFNTPKH") '�i�v�e���R�i�m�g�|�����p�x�Q��
            If fldNameExist("HWFNTPKU") Then .HWFNTPKU = rs("HWFNTPKU") '�i�v�e���R�i�m�g�|�����p�x�Q�E
            If fldNameExist("HWFCRSSK") Then .HWFCRSSK = rs("HWFCRSSK") '�i�v�e���R�N���X�r�r����
            If fldNameExist("HWFMDCEN") Then .HWFMDCEN = fncNullCheck(rs("HWFMDCEN")) '�i�v�e���R�ʃ_�����፷���S
            If fldNameExist("HWFMDMAX") Then .HWFMDMAX = fncNullCheck(rs("HWFMDMAX")) '�i�v�e���R�ʃ_�����፷���
            If fldNameExist("HWFMDMIN") Then .HWFMDMIN = fncNullCheck(rs("HWFMDMIN")) '�i�v�e���R�ʃ_�����፷����
            If fldNameExist("HWFMDSPH") Then .HWFMDSPH = rs("HWFMDSPH") '�i�v�e���R�ʃ_������ʒu�Q��
            If fldNameExist("HWFMDSPT") Then .HWFMDSPT = rs("HWFMDSPT") '�i�v�e���R�ʃ_������ʒu�Q�_
            If fldNameExist("HWFMDSPI") Then .HWFMDSPI = rs("HWFMDSPI") '�i�v�e���R�ʃ_������ʒu�Q��
            If fldNameExist("HWFMDHWT") Then .HWFMDHWT = rs("HWFMDHWT") '�i�v�e���R�ʃ_���ۏؕ��@�Q��
            If fldNameExist("HWFMDHWS") Then .HWFMDHWS = rs("HWFMDHWS") '�i�v�e���R�ʃ_���ۏؕ��@�Q��
            If fldNameExist("HWFMDKHM") Then .HWFMDKHM = rs("HWFMDKHM") '�i�v�e���R�ʃ_�������p�x�Q��
            If fldNameExist("HWFMDKHN") Then .HWFMDKHN = rs("HWFMDKHN") '�i�v�e���R�ʃ_�������p�x�Q��
            If fldNameExist("HWFMDKHH") Then .HWFMDKHH = rs("HWFMDKHH") '�i�v�e���R�ʃ_�������p�x�Q��
            If fldNameExist("HWFMDKHU") Then .HWFMDKHU = rs("HWFMDKHU") '�i�v�e���R�ʃ_�������p�x�Q�E
            If fldNameExist("IFKBN") Then .IFKBN = rs("IFKBN") '�h�^�e�敪
            If fldNameExist("SYORIKBN") Then .SYORIKBN = rs("SYORIKBN") '�����敪
            If fldNameExist("SPECRRNO") Then .SPECRRNO = rs("SPECRRNO") '�d�l�o�^�˗��ԍ�
            If fldNameExist("SXLMCNO") Then .SXLMCNO = rs("SXLMCNO") '�r�w�k��������ԍ�
            If fldNameExist("WFMCNO") Then .WFMCNO = rs("WFMCNO") '�v�e��������ԍ�
            If fldNameExist("STAFFID") Then .StaffID = rs("STAFFID") '�Ј�ID
            If fldNameExist("REGDATE") Then .REGDATE = rs("REGDATE") '�o�^���t
            If fldNameExist("UPDDATE") Then .UPDDATE = rs("UPDDATE") '�X�V���t
            If fldNameExist("SENDFLAG") Then .SENDFLAG = rs("SENDFLAG") '���M�t���O
            If fldNameExist("SENDDATE") Then .SENDDATE = rs("SENDDATE") '���M���t
            If fldNameExist("HWFDVDMXN") Then .HWFDVDMXN = fncNullCheck(rs("HWFDVDMXN")) '�i�v�e�c�u�c�Q���
            If fldNameExist("HWFDVDMNN") Then .HWFDVDMNN = fncNullCheck(rs("HWFDVDMNN")) '�i�v�e�c�u�c�Q����
'            If fldNameExist("HWFDSONWY") Then .HWFDSONWY = rs("HWFDSONWY") '�i�v�e�c�r�n�c�M�����@
'            If fldNameExist("HWFMSUMX") Then .HWFMSUMX = fncNullCheck(rs("HWFMSUMX")) '�i�v�e�l�X�N���b�`���
'            If fldNameExist("HWFMSUZY") Then .HWFMSUZY = rs("HWFMSUZY") '�i�v�e�l�X�N���b�`�������
'            If fldNameExist("HWFMSUKW") Then .HWFMSUKW = rs("HWFMSUKW") '�i�v�e�l�X�N���b�`�������@
'            If fldNameExist("HWFMSUSZ") Then .HWFMSUSZ = fncNullCheck(rs("HWFMSUSZ")) '�i�v�e�l�X�N���b�`�T�C�Y
'            If fldNameExist("HWFNP1AR") Then .HWFNP1AR = fncNullCheck(rs("HWFNP1AR")) '�iWF�i�m�g�|�P�G���A
'            If fldNameExist("HWFNP1MAX") Then .HWFNP1MAX = fncNullCheck(rs("HWFNP1MAX")) '�iWF�i�m�g�|�P���
'            If fldNameExist("HWFNP2AR") Then .HWFNP2AR = fncNullCheck(rs("HWFNP2AR")) '�iWF�i�m�g�|�Q�G���A
'            If fldNameExist("HWFNP2MAX") Then .HWFNP2MAX = fncNullCheck(rs("HWFNP2MAX")) '�iWF�i�m�g�|�Q���
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME026 = FUNCTION_RETURN_SUCCESS

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

'�T�v      :�e�[�u���uTBCME028�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :records()     ,O  ,typ_TBCME028    ,���o���R�[�h
'          :formID        ,I  ,String          ,�g�p�t�H�[��ID
'          :sqlOrder      ,I  ,tFullHinban     ,���o�i�ԁi�z��j
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,���o�̐���
'����      :�����L�[�́A�HINBAN�+�uMNOREVNO�v+�uFACTORY�v+�uOPECOND�v�̕�����Ƃ���
'����      :�V�K�쐬 2005/06/15 ffc)tanabe
Public Function DBDRV_GetTBCME028(records() As typ_TBCME028, formID$, HIN() As tFullHinban) As FUNCTION_RETURN

    Dim sql         As String           'SQL�S��
    Dim sqlBase     As String           'SQL��{��(WHERE�߂̑O�܂�)
    Dim sqlWhere    As String           'SQLWhere��
    Dim rs          As OraDynaset       'RecordSet
    Dim recCnt      As Long             '���R�[�h��
    Dim key         As String           '����KEY
    Dim i           As Long             'ٰ�߶���
    Dim j           As Long             'ٰ�߶���2


    DBDRV_GetTBCME028 = FUNCTION_RETURN_FAILURE
            
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_kensa_SQL.bas -- Function DBDRV_GetTBCME028"

    Select Case formID
        Case "f_cmec067_1"           '�uSPV�����Q�Ɓv
            sqlBase = "SELECT HINBAN, MNOREVNO, FACTORY, OPECOND, HWFSPVMX, HWFSPVKM, HWFSPVKN, HWFSPVKH, HWFSPVKU, HWFSPVSH, HWFSPVST, HWFSPVSI, HWFSPVHT, " & _
                "HWFSPVHS, HWFDLMIN, HWFDLMAX, HWFDLKHM, HWFDLKHN, HWFDLKHH, HWFDLKHU, HWFDLSPH, HWFDLSPT, HWFDLSPI, HWFDLHWT, HWFDLHWS, HWFSPVMXN "
    End Select
       
    sqlBase = sqlBase & "From TBCME028"
    
    '''SQL��Where���쐬
    For i = 0 To UBound(HIN)
        With HIN(i)
            key = key & "'" & .hinban & Format(.mnorevno, "00000") & .factory & .opecond & "'"
            If i <> UBound(HIN) Then
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
            If fldNameExist("HINBAN") Then .hinban = rs("HINBAN")                           '�i��
            If fldNameExist("MNOREVNO") Then .mnorevno = rs("MNOREVNO")                     '���i�ԍ������ԍ�
            If fldNameExist("FACTORY") Then .factory = rs("FACTORY")                        '�H��
            If fldNameExist("OPECOND") Then .opecond = rs("OPECOND")                        '���Ə���
            If fldNameExist("HMGSTRRNO") Then .HMGSTRRNO = rs("HMGSTRRNO")                  '�i�Ǘ��d�l�o�^�˗��ԍ�
            If fldNameExist("HMGSTFNO") Then .HMGSTFNO = rs("HMGSTFNO")                     '�i�Ǘ��Ј��m��
            If fldNameExist("HMGWFSNO") Then .HMGWFSNO = rs("HMGWFSNO")                     '�i�Ǘ��v�e���i�ԍ�
            If fldNameExist("HMGWFSNE") Then .HMGWFSNE = fncNullCheck(rs("HMGWFSNE"))       '�i�Ǘ��v�e���i�ԍ��}��
            If fldNameExist("HWFMK1SI") Then .HWFMK1SI = fncNullCheck(rs("HWFMK1SI"))       '�i�v�e�ʌ����ׂP�T�C�Y
            If fldNameExist("HWFMK1MX") Then .HWFMK1MX = fncNullCheck(rs("HWFMK1MX"))       '�i�v�e�ʌ����ׂP���
            If fldNameExist("HWFMK1SZ") Then .HWFMK1SZ = rs("HWFMK1SZ")                     '�i�v�e�ʌ����ׂP�������
            If fldNameExist("HWFMK1ZA") Then .HWFMK1ZA = fncNullCheck(rs("HWFMK1ZA"))       '�i�v�e�ʌ����ׂP���O�̈�
            If fldNameExist("HWFMK1HT") Then .HWFMK1HT = rs("HWFMK1HT")                     '�i�v�e�ʌ����ׂP�ۏؕ��@�Q��
            If fldNameExist("HWFMK1HS") Then .HWFMK1HS = rs("HWFMK1HS")                     '�i�v�e�ʌ����ׂP�ۏؕ��@�Q��
            If fldNameExist("HWFMK1KM") Then .HWFMK1KM = rs("HWFMK1KM")                     '�i�v�e�ʌ����ׂP�����p�x�Q��
            If fldNameExist("HWFMK1KN") Then .HWFMK1KN = rs("HWFMK1KN")                     '�i�v�e�ʌ����ׂP�����p�x�Q��
            If fldNameExist("HWFMK1KH") Then .HWFMK1KH = rs("HWFMK1KH")                     '�i�v�e�ʌ����ׂP�����p�x�Q��
            If fldNameExist("HWFMK1KU") Then .HWFMK1KU = rs("HWFMK1KU")                     '�i�v�e�ʌ����ׂP�����p�x�Q�E
            If fldNameExist("HWFM1B1") Then .HWFM1B1 = fncNullCheck(rs("HWFM1B1"))          '�i�v�e�ʌ����ׂP���E�P
            If fldNameExist("HWFM1B1B") Then .HWFM1B1B = fncNullCheck(rs("HWFM1B1B"))       '�i�v�e�ʌ����ׂP���E�P��
            If fldNameExist("HWFM1B2") Then .HWFM1B2 = fncNullCheck(rs("HWFM1B2"))          '�i�v�e�ʌ����ׂP���E�Q
            If fldNameExist("HWFM1B2B") Then .HWFM1B2B = fncNullCheck(rs("HWFM1B2B"))       '�i�v�e�ʌ����ׂP���E�Q��
            If fldNameExist("HWFM1B3") Then .HWFM1B3 = fncNullCheck(rs("HWFM1B3"))          '�i�v�e�ʌ����ׂP���E�R
            If fldNameExist("HWFM1B3B") Then .HWFM1B3B = fncNullCheck(rs("HWFM1B3B"))       '�i�v�e�ʌ����ׂP���E�R��
            If fldNameExist("HWFMK2SI") Then .HWFMK2SI = fncNullCheck(rs("HWFMK2SI"))       '�i�v�e�ʌ����ׂQ�T�C�Y
            If fldNameExist("HWFMK2MX") Then .HWFMK2MX = fncNullCheck(rs("HWFMK2MX"))       '�i�v�e�ʌ����ׂQ���
            If fldNameExist("HWFMK2HT") Then .HWFMK2HT = rs("HWFMK2HT")                     '�i�v�e�ʌ����ׂQ�ۏؕ��@�Q��
            If fldNameExist("HWFMK2HS") Then .HWFMK2HS = rs("HWFMK2HS")                     '�i�v�e�ʌ����ׂQ�ۏؕ��@�Q��
            If fldNameExist("HWFMK2KM") Then .HWFMK2KM = rs("HWFMK2KM")                     '�i�v�e�ʌ����ׂQ�����p�x�Q��
            If fldNameExist("HWFMK2KN") Then .HWFMK2KN = rs("HWFMK2KN")                     '�i�v�e�ʌ����ׂQ�����p�x�Q��
            If fldNameExist("HWFMK2KH") Then .HWFMK2KH = rs("HWFMK2KH")                     '�i�v�e�ʌ����ׂQ�����p�x�Q��
            If fldNameExist("HWFMK2KU") Then .HWFMK2KU = rs("HWFMK2KU")                     '�i�v�e�ʌ����ׂQ�����p�x�Q�E
            If fldNameExist("HWFM2B1") Then .HWFM2B1 = fncNullCheck(rs("HWFM2B1"))          '�i�v�e�ʌ����ׂQ���E�P
            If fldNameExist("HWFM2B1B") Then .HWFM2B1B = fncNullCheck(rs("HWFM2B1B"))       '�i�v�e�ʌ����ׂQ���E�P��
            If fldNameExist("HWFM2B2") Then .HWFM2B2 = fncNullCheck(rs("HWFM2B2"))          '�i�v�e�ʌ����ׂQ���E�Q
            If fldNameExist("HWFM2B2B") Then .HWFM2B2B = fncNullCheck(rs("HWFM2B2B"))       '�i�v�e�ʌ����ׂQ���E�Q��
            If fldNameExist("HWFM2B3") Then .HWFM2B3 = fncNullCheck(rs("HWFM2B3"))          '�i�v�e�ʌ����ׂQ���E�R
            If fldNameExist("HWFM2B3B") Then .HWFM2B3B = fncNullCheck(rs("HWFM2B3B"))       '�i�v�e�ʌ����ׂQ���E�R��
            If fldNameExist("HWFMK3SI") Then .HWFMK3SI = fncNullCheck(rs("HWFMK3SI"))       '�i�v�e�ʌ����ׂR�T�C�Y
            If fldNameExist("HWFMK3MX") Then .HWFMK3MX = fncNullCheck(rs("HWFMK3MX"))       '�i�v�e�ʌ����ׂR���
            If fldNameExist("HWFMK3HT") Then .HWFMK3HT = rs("HWFMK3HT")                     '�i�v�e�ʌ����ׂR�ۏؕ��@�Q��
            If fldNameExist("HWFMK3HS") Then .HWFMK3HS = rs("HWFMK3HS")                     '�i�v�e�ʌ����ׂR�ۏؕ��@�Q��
            If fldNameExist("HWFMK3KM") Then .HWFMK3KM = rs("HWFMK3KM")                     '�i�v�e�ʌ����ׂR�����p�x�Q��
            If fldNameExist("HWFMK3KN") Then .HWFMK3KN = rs("HWFMK3KN")                     '�i�v�e�ʌ����ׂR�����p�x�Q��
            If fldNameExist("HWFMK3KH") Then .HWFMK3KH = rs("HWFMK3KH")                     '�i�v�e�ʌ����ׂR�����p�x�Q��
            If fldNameExist("HWFMK3KU") Then .HWFMK3KU = rs("HWFMK3KU")                     '�i�v�e�ʌ����ׂR�����p�x�Q�E
            If fldNameExist("HWFM3B1") Then .HWFM3B1 = fncNullCheck(rs("HWFM3B1"))          '�i�v�e�ʌ����ׂR���E�P
            If fldNameExist("HWFM3B1B") Then .HWFM3B1B = fncNullCheck(rs("HWFM3B1B"))       '�i�v�e�ʌ����ׂR���E�P��
            If fldNameExist("HWFM3B2") Then .HWFM3B2 = fncNullCheck(rs("HWFM3B2"))          '�i�v�e�ʌ����ׂR���E�Q
            If fldNameExist("HWFM3B2B") Then .HWFM3B2B = fncNullCheck(rs("HWFM3B2B"))       '�i�v�e�ʌ����ׂR���E�Q��
            If fldNameExist("HWFM3B3") Then .HWFM3B3 = fncNullCheck(rs("HWFM3B3"))          '�i�v�e�ʌ����ׂR���E�R
            If fldNameExist("HWFM3B3B") Then .HWFM3B3B = fncNullCheck(rs("HWFM3B3B"))       '�i�v�e�ʌ����ׂR���E�R��
            If fldNameExist("HWFMK4SI") Then .HWFMK4SI = fncNullCheck(rs("HWFMK4SI"))       '�i�v�e�ʌ����ׂS�T�C�Y
            If fldNameExist("HWFMK4MX") Then .HWFMK4MX = fncNullCheck(rs("HWFMK4MX"))       '�i�v�e�ʌ����ׂS���
            If fldNameExist("HWFMK4HT") Then .HWFMK4HT = rs("HWFMK4HT")                     '�i�v�e�ʌ����ׂS�ۏؕ��@�Q��
            If fldNameExist("HWFMK4HS") Then .HWFMK4HS = rs("HWFMK4HS")                     '�i�v�e�ʌ����ׂS�ۏؕ��@�Q��
            If fldNameExist("HWFMK4KM") Then .HWFMK4KM = rs("HWFMK4KM")                     '�i�v�e�ʌ����ׂS�����p�x�Q��
            If fldNameExist("HWFMK4KN") Then .HWFMK4KN = rs("HWFMK4KN")                     '�i�v�e�ʌ����ׂS�����p�x�Q��
            If fldNameExist("HWFMK4KH") Then .HWFMK4KH = rs("HWFMK4KH")                     '�i�v�e�ʌ����ׂS�����p�x�Q��
            If fldNameExist("HWFMK4KU") Then .HWFMK4KU = rs("HWFMK4KU")                     '�i�v�e�ʌ����ׂS�����p�x�Q�E
            If fldNameExist("HWFM4B1") Then .HWFM4B1 = fncNullCheck(rs("HWFM4B1"))          '�i�v�e�ʌ����ׂS���E�P
            If fldNameExist("HWFM4B1B") Then .HWFM4B1B = fncNullCheck(rs("HWFM4B1B"))       '�i�v�e�ʌ����ׂS���E�P��
            If fldNameExist("HWFM4B2") Then .HWFM4B2 = fncNullCheck(rs("HWFM4B2"))          '�i�v�e�ʌ����ׂS���E�Q
            If fldNameExist("HWFM4B2B") Then .HWFM4B2B = fncNullCheck(rs("HWFM4B2B"))       '�i�v�e�ʌ����ׂS���E�Q��
            If fldNameExist("HWFM4B3") Then .HWFM4B3 = fncNullCheck(rs("HWFM4B3"))          '�i�v�e�ʌ����ׂS���E�R
            If fldNameExist("HWFM4B3B") Then .HWFM4B3B = fncNullCheck(rs("HWFM4B3B"))       '�i�v�e�ʌ����ׂS���E�R��
            If fldNameExist("HWFMB1SI") Then .HWFMB1SI = fncNullCheck(rs("HWFMB1SI"))       '�i�v�e�ʌ����ח��P�T�C�Y
            If fldNameExist("HWFMB1MX") Then .HWFMB1MX = fncNullCheck(rs("HWFMB1MX"))       '�i�v�e�ʌ����ח��P���
            If fldNameExist("HWFMB1SZ") Then .HWFMB1SZ = rs("HWFMB1SZ")                     '�i�v�e�ʌ����ח��P�������
            If fldNameExist("HWFMB1ZA") Then .HWFMB1ZA = fncNullCheck(rs("HWFMB1ZA"))       '�i�v�e�ʌ����ח��P���O�̈�
            If fldNameExist("HWFMB1HT") Then .HWFMB1HT = rs("HWFMB1HT")                     '�i�v�e�ʌ����ח��P�ۏؕ��@�Q��
            If fldNameExist("HWFMB1HS") Then .HWFMB1HS = rs("HWFMB1HS")                     '�i�v�e�ʌ����ח��P�ۏؕ��@�Q��
            If fldNameExist("HWFMB1KM") Then .HWFMB1KM = rs("HWFMB1KM")                     '�i�v�e�ʌ����ח��P�����p�x�Q��
            If fldNameExist("HWFMB1KN") Then .HWFMB1KN = rs("HWFMB1KN")                     '�i�v�e�ʌ����ח��P�����p�x�Q��
            If fldNameExist("HWFMB1KH") Then .HWFMB1KH = rs("HWFMB1KH")                     '�i�v�e�ʌ����ח��P�����p�x�Q��
            If fldNameExist("HWFMB1KU") Then .HWFMB1KU = rs("HWFMB1KU")                     '�i�v�e�ʌ����ח��P�����p�x�Q�E
            If fldNameExist("HWFMB2SI") Then .HWFMB2SI = fncNullCheck(rs("HWFMB2SI"))       '�i�v�e�ʌ����ח��Q�T�C�Y
            If fldNameExist("HWFMB2MX") Then .HWFMB2MX = fncNullCheck(rs("HWFMB2MX"))       '�i�v�e�ʌ����ח��Q���
            If fldNameExist("HWFMB2SZ") Then .HWFMB2SZ = rs("HWFMB2SZ")                     '�i�v�e�ʌ����ח��Q�������
            If fldNameExist("HWFMB2ZA") Then .HWFMB2ZA = fncNullCheck(rs("HWFMB2ZA"))       '�i�v�e�ʌ����ח��Q���O�̈�
            If fldNameExist("HWFMB2HT") Then .HWFMB2HT = rs("HWFMB2HT")                     '�i�v�e�ʌ����ח��Q�ۏؕ��@�Q��
            If fldNameExist("HWFMB2HS") Then .HWFMB2HS = rs("HWFMB2HS")                     '�i�v�e�ʌ����ח��Q�ۏؕ��@�Q��
            If fldNameExist("HWFMB2KM") Then .HWFMB2KM = rs("HWFMB2KM")                     '�i�v�e�ʌ����ח��Q�����p�x�Q��
            If fldNameExist("HWFMB2KN") Then .HWFMB2KN = rs("HWFMB2KN")                     '�i�v�e�ʌ����ח��Q�����p�x�Q��
            If fldNameExist("HWFMB2KH") Then .HWFMB2KH = rs("HWFMB2KH")                     '�i�v�e�ʌ����ח��Q�����p�x�Q��
            If fldNameExist("HWFMB2KU") Then .HWFMB2KU = rs("HWFMB2KU")                     '�i�v�e�ʌ����ח��Q�����p�x�Q�E
            If fldNameExist("HWFMKSRE") Then .HWFMKSRE = rs("HWFMKSRE")                     '�i�v�e�ʌ����ב����
            If fldNameExist("HWFMKKW") Then .HWFMKKW = rs("HWFMKKW")                        '�i�v�e�ʌ����׌������@
            If fldNameExist("HWFMPIPT") Then .HWFMPIPT = rs("HWFMPIPT")                     '�i�v�e�ʌ����ׂo�h�o����
            If fldNameExist("HWFMPIPK") Then .HWFMPIPK = fncNullCheck(rs("HWFMPIPK"))       '�i�v�e�ʌ����ׂo�h�o��
            If fldNameExist("HWFMPISH") Then .HWFMPISH = rs("HWFMPISH")                     '�i�v�e�ʌ��o�h�o����ʒu�Q��
            If fldNameExist("HWFMPIST") Then .HWFMPIST = rs("HWFMPIST")                     '�i�v�e�ʌ��o�h�o����ʒu�Q�_
            If fldNameExist("HWFMPISI") Then .HWFMPISI = rs("HWFMPISI")                     '�i�v�e�ʌ��o�h�o����ʒu�Q��
            If fldNameExist("HWFMPIKM") Then .HWFMPIKM = rs("HWFMPIKM")                     '�i�v�e�ʌ��o�h�o�����p�x�Q��
            If fldNameExist("HWFMPIKN") Then .HWFMPIKN = rs("HWFMPIKN")                     '�i�v�e�ʌ��o�h�o�����p�x�Q��
            If fldNameExist("HWFMPIKH") Then .HWFMPIKH = rs("HWFMPIKH")                     '�i�v�e�ʌ��o�h�o�����p�x�Q��
            If fldNameExist("HWFMPIKU") Then .HWFMPIKU = rs("HWFMPIKU")                     '�i�v�e�ʌ��o�h�o�����p�x�Q�E
            If fldNameExist("HWFMNMAX") Then .HWFMNMAX = fncNullCheck(rs("HWFMNMAX"))       '�i�v�e�����Z�x���
            If fldNameExist("HWFMNALX") Then .HWFMNALX = fncNullCheck(rs("HWFMNALX"))       '�i�v�e�����Z�x�`�k���
            If fldNameExist("HWFMNCAX") Then .HWFMNCAX = fncNullCheck(rs("HWFMNCAX"))       '�i�v�e�����Z�x�b�`���
            If fldNameExist("HWFMNCRX") Then .HWFMNCRX = fncNullCheck(rs("HWFMNCRX"))       '�i�v�e�����Z�x�b�q���
            If fldNameExist("HWFMNCUX") Then .HWFMNCUX = fncNullCheck(rs("HWFMNCUX"))       '�i�v�e�����Z�x�b�t���
            If fldNameExist("HWFMNFEX") Then .HWFMNFEX = fncNullCheck(rs("HWFMNFEX"))       '�i�v�e�����Z�x�e�d���
            If fldNameExist("HWFMNKMX") Then .HWFMNKMX = fncNullCheck(rs("HWFMNKMX"))       '�i�v�e�����Z�x�j���
            If fldNameExist("HWFMNMGX") Then .HWFMNMGX = fncNullCheck(rs("HWFMNMGX"))       '�i�v�e�����Z�x�l�f���
            If fldNameExist("HWFMNNAX") Then .HWFMNNAX = fncNullCheck(rs("HWFMNNAX"))       '�i�v�e�����Z�x�m�`���
            If fldNameExist("HWFMNNIX") Then .HWFMNNIX = fncNullCheck(rs("HWFMNNIX"))       '�i�v�e�����Z�x�m�h���
            If fldNameExist("HWFMNZNX") Then .HWFMNZNX = fncNullCheck(rs("HWFMNZNX"))       '�i�v�e�����Z�x�y�m���
            If fldNameExist("HWFMNKWY") Then .HWFMNKWY = rs("HWFMNKWY")                     '�i�v�e�����Z�x�������@
            If fldNameExist("HWFMNSPH") Then .HWFMNSPH = rs("HWFMNSPH")                     '�i�v�e�����Z�x����ʒu�Q��
            If fldNameExist("HWFMNSPT") Then .HWFMNSPT = rs("HWFMNSPT")                     '�i�v�e�����Z�x����ʒu�Q�_
            If fldNameExist("HWFMNSPI") Then .HWFMNSPI = rs("HWFMNSPI")                     '�i�v�e�����Z�x����ʒu�Q��
            If fldNameExist("HWFMNHWT") Then .HWFMNHWT = rs("HWFMNHWT")                     '�i�v�e�����Z�x�ۏؕ��@�Q��
            If fldNameExist("HWFMNHWS") Then .HWFMNHWS = rs("HWFMNHWS")                     '�i�v�e�����Z�x�ۏؕ��@�Q��
            If fldNameExist("HWFMNKHM") Then .HWFMNKHM = rs("HWFMNKHM")                     '�i�v�e�����Z�x�����p�x�Q��
            If fldNameExist("HWFMNKHN") Then .HWFMNKHN = rs("HWFMNKHN")                     '�i�v�e�����Z�x�����p�x�Q��
            If fldNameExist("HWFMNKHH") Then .HWFMNKHH = rs("HWFMNKHH")                     '�i�v�e�����Z�x�����p�x�Q��
            If fldNameExist("HWFMNKHU") Then .HWFMNKHU = rs("HWFMNKHU")                     '�i�v�e�����Z�x�����p�x�Q�E
            If fldNameExist("HWFSPVMX") Then .HWFSPVMX = fncNullCheck(rs("HWFSPVMX"))       '�i�v�e�r�o�u�e�d���
            If fldNameExist("HWFSPVKM") Then .HWFSPVKM = rs("HWFSPVKM")                     '�i�v�e�r�o�u�e�d�����p�x�Q��
            If fldNameExist("HWFSPVKN") Then .HWFSPVKN = rs("HWFSPVKN")                     '�i�v�e�r�o�u�e�d�����p�x�Q��
            If fldNameExist("HWFSPVKH") Then .HWFSPVKH = rs("HWFSPVKH")                     '�i�v�e�r�o�u�e�d�����p�x�Q��
            If fldNameExist("HWFSPVKU") Then .HWFSPVKU = rs("HWFSPVKU")                     '�i�v�e�r�o�u�e�d�����p�x�Q�E
            If fldNameExist("HWFSPVSH") Then .HWFSPVSH = rs("HWFSPVSH")                     '�i�v�e�r�o�u�e�d����ʒu�Q��
            If fldNameExist("HWFSPVST") Then .HWFSPVST = rs("HWFSPVST")                     '�i�v�e�r�o�u�e�d����ʒu�Q�_
            If fldNameExist("HWFSPVSI") Then .HWFSPVSI = rs("HWFSPVSI")                     '�i�v�e�r�o�u�e�d����ʒu�Q��
            If fldNameExist("HWFSPVHT") Then .HWFSPVHT = rs("HWFSPVHT")                     '�i�v�e�r�o�u�e�d�ۏؕ��@�Q��
            If fldNameExist("HWFSPVHS") Then .HWFSPVHS = rs("HWFSPVHS")                     '�i�v�e�r�o�u�e�d�ۏؕ��@�Q��
            If fldNameExist("HWFDLMIN") Then .HWFDLMIN = fncNullCheck(rs("HWFDLMIN"))       '�i�v�e�g�U������
            If fldNameExist("HWFDLMAX") Then .HWFDLMAX = fncNullCheck(rs("HWFDLMAX"))       '�i�v�e�g�U�����
            If fldNameExist("HWFDLKHM") Then .HWFDLKHM = rs("HWFDLKHM")                     '�i�v�e�g�U�������p�x�Q��
            If fldNameExist("HWFDLKHN") Then .HWFDLKHN = rs("HWFDLKHN")                     '�i�v�e�g�U�������p�x�Q��
            If fldNameExist("HWFDLKHH") Then .HWFDLKHH = rs("HWFDLKHH")                     '�i�v�e�g�U�������p�x�Q��
            If fldNameExist("HWFDLKHU") Then .HWFDLKHU = rs("HWFDLKHU")                     '�i�v�e�g�U�������p�x�Q�E
            If fldNameExist("HWFDLSPH") Then .HWFDLSPH = rs("HWFDLSPH")                     '�i�v�e�g�U������ʒu�Q��
            If fldNameExist("HWFDLSPT") Then .HWFDLSPT = rs("HWFDLSPT")                     '�i�v�e�g�U������ʒu�Q�_
            If fldNameExist("HWFDLSPI") Then .HWFDLSPI = rs("HWFDLSPI")                     '�i�v�e�g�U������ʒu�Q��
            If fldNameExist("HWFDLHWT") Then .HWFDLHWT = rs("HWFDLHWT")                     '�i�v�e�g�U���ۏؕ��@�Q��
            If fldNameExist("HWFDLHWS") Then .HWFDLHWS = rs("HWFDLHWS")                     '�i�v�e�g�U���ۏؕ��@�Q��
            If fldNameExist("HWFGKNO1") Then .HWFGKNO1 = rs("HWFGKNO1")                     '�i�v�e�O�ϋK�i�m���P
            If fldNameExist("HWFGKNO2") Then .HWFGKNO2 = rs("HWFGKNO2")                     '�i�v�e�O�ϋK�i�m���Q
            If fldNameExist("HWFOTMIN") Then .HWFOTMIN = fncNullCheck(rs("HWFOTMIN"))       '�i�v�e�_�����ψ�����
            If fldNameExist("HWFOTMX1") Then .HWFOTMX1 = fncNullCheck(rs("HWFOTMX1"))       '�i�v�e�_�����ψ�����P
            If fldNameExist("HWFOTMX2") Then .HWFOTMX2 = fncNullCheck(rs("HWFOTMX2"))       '�i�v�e�_�����ψ�����Q
            If fldNameExist("HWFOTSPH") Then .HWFOTSPH = rs("HWFOTSPH")                     '�i�v�e�_�����ψ�����ʒu�Q��
            If fldNameExist("HWFOTSPT") Then .HWFOTSPT = rs("HWFOTSPT")                     '�i�v�e�_�����ψ�����ʒu�Q�_
            If fldNameExist("HWFOTSPI") Then .HWFOTSPI = rs("HWFOTSPI")                     '�i�v�e�_�����ψ�����ʒu�Q��
            If fldNameExist("HWFOTHWT") Then .HWFOTHWT = rs("HWFOTHWT")                     '�i�v�e�_�����ψ��ۏؕ��@�Q��
            If fldNameExist("HWFOTHWS") Then .HWFOTHWS = rs("HWFOTHWS")                     '�i�v�e�_�����ψ��ۏؕ��@�Q��
            If fldNameExist("HWFOTKWY") Then .HWFOTKWY = rs("HWFOTKWY")                     '�i�v�e�_�����ψ��������@
            If fldNameExist("HWFOTKW1") Then .HWFOTKW1 = rs("HWFOTKW1")                     '�i�v�e�_�����ψ��������@�P
            If fldNameExist("HWFOTKW2") Then .HWFOTKW2 = rs("HWFOTKW2")                     '�i�v�e�_�����ψ��������@�Q
            If fldNameExist("HWFOTKHM") Then .HWFOTKHM = rs("HWFOTKHM")                     '�i�v�e�_�����ψ������p�x�Q��
            If fldNameExist("HWFOTKHN") Then .HWFOTKHN = rs("HWFOTKHN")                     '�i�v�e�_�����ψ������p�x�Q��
            If fldNameExist("HWFOTKHH") Then .HWFOTKHH = rs("HWFOTKHH")                     '�i�v�e�_�����ψ������p�x�Q��
            If fldNameExist("HWFOTKHU") Then .HWFOTKHU = rs("HWFOTKHU")                     '�i�v�e�_�����ψ������p�x�Q�E
            If fldNameExist("HWFTSPHM") Then .HWFTSPHM = rs("HWFTSPHM")                     '�i�v�e�g���X�T���v���p�x�Q��
            If fldNameExist("HWFTSPHN") Then .HWFTSPHN = rs("HWFTSPHN")                     '�i�v�e�g���X�T���v���p�x�Q��
            If fldNameExist("HWFTSPHH") Then .HWFTSPHH = rs("HWFTSPHH")                     '�i�v�e�g���X�T���v���p�x�Q��
            If fldNameExist("HWFTSPHU") Then .HWFTSPHU = rs("HWFTSPHU")                     '�i�v�e�g���X�T���v���p�x�Q�E
            If fldNameExist("HWFLTDCX") Then .HWFLTDCX = fncNullCheck(rs("HWFLTDCX"))       '�i�v�e�k�s�c�Z�x�b�t���
            If fldNameExist("HWFLTDIN") Then .HWFLTDIN = rs("HWFLTDIN")                     '�i�v�e�k�s�c�Z�x�w��
            If fldNameExist("HWFLTDKW") Then .HWFLTDKW = rs("HWFLTDKW")                     '�i�v�e�k�s�c�Z�x�������@
            If fldNameExist("HWFLTDSH") Then .HWFLTDSH = rs("HWFLTDSH")                     '�i�v�e�k�s�c�Z�x����ʒu�Q��
            If fldNameExist("HWFLTDST") Then .HWFLTDST = rs("HWFLTDST")                     '�i�v�e�k�s�c�Z�x����ʒu�Q�_
            If fldNameExist("HWFLTDSI") Then .HWFLTDSI = rs("HWFLTDSI")                     '�i�v�e�k�s�c�Z�x����ʒu�Q��
            If fldNameExist("HWFLTDHT") Then .HWFLTDHT = rs("HWFLTDHT")                     '�i�v�e�k�s�c�Z�x�ۏؕ��@�Q��
            If fldNameExist("HWFLTDHS") Then .HWFLTDHS = rs("HWFLTDHS")                     '�i�v�e�k�s�c�Z�x�ۏؕ��@�Q��
            If fldNameExist("HWFLTDKM") Then .HWFLTDKM = rs("HWFLTDKM")                     '�i�v�e�k�s�c�Z�x�����p�x�Q��
            If fldNameExist("HWFLTDKN") Then .HWFLTDKN = rs("HWFLTDKN")                     '�i�v�e�k�s�c�Z�x�����p�x�Q��
            If fldNameExist("HWFLTDKH") Then .HWFLTDKH = rs("HWFLTDKH")                     '�i�v�e�k�s�c�Z�x�����p�x�Q��
            If fldNameExist("HWFLTDKU") Then .HWFLTDKU = rs("HWFLTDKU")                     '�i�v�e�k�s�c�Z�x�����p�x�Q�E
            If fldNameExist("IFKBN") Then .IFKBN = rs("IFKBN")                              '�h�^�e�敪
            If fldNameExist("SYORIKBN") Then .SYORIKBN = rs("SYORIKBN")                     '�����敪
            If fldNameExist("SPECRRNO") Then .SPECRRNO = rs("SPECRRNO")                     '�d�l�o�^�˗��ԍ�
            If fldNameExist("SXLMCNO") Then .SXLMCNO = rs("SXLMCNO")                        '�r�w�k��������ԍ�
            If fldNameExist("WFMCNO") Then .WFMCNO = rs("WFMCNO")                           '�v�e��������ԍ�
            If fldNameExist("STAFFID") Then .StaffID = rs("STAFFID")                        '�Ј�ID
            If fldNameExist("REGDATE") Then .REGDATE = rs("REGDATE")                        '�o�^���t
            If fldNameExist("UPDDATE") Then .UPDDATE = rs("UPDDATE")                        '�X�V���t
            If fldNameExist("SENDFLAG") Then .SENDFLAG = rs("SENDFLAG")                     '���M�t���O
            If fldNameExist("SENDDATE") Then .SENDDATE = rs("SENDDATE")                     '���M���t
            If fldNameExist("HWFSPVAM") Then .HWFSPVAM = fncNullCheck(rs("HWFSPVAM"))       '�i�v�e�r�o�u�e�d����
            If fldNameExist("HWFMK1MC") Then .HWFMK1MC = rs("HWFMK1MC")                     '�i�v�e�ʌ����ׂP�ʎw��
            If fldNameExist("HWFMK2MC") Then .HWFMK2MC = rs("HWFMK2MC")                     '�i�v�e�ʌ����ׂQ�ʎw��
            If fldNameExist("HWFMK3MC") Then .HWFMK3MC = rs("HWFMK3MC")                     '�i�v�e�ʌ����ׂR�ʎw��
            If fldNameExist("HWFMK4MC") Then .HWFMK4MC = rs("HWFMK4MC")                     '�i�v�e�ʌ����ׂS�ʎw��
            If fldNameExist("HWFMK5MC") Then .HWFMK5MC = rs("HWFMK5MC")                     '�i�v�e�ʌ����ׂT�ʎw��
            If fldNameExist("HWFMK6MC") Then .HWFMK6MC = rs("HWFMK6MC")                     '�i�v�e�ʌ����ׂU�ʎw��
            If fldNameExist("HWFMK2SZ") Then .HWFMK2SZ = rs("HWFMK2SZ")                     '�i�v�e�ʌ����ׂQ�������
            If fldNameExist("HWFMK3SZ") Then .HWFMK3SZ = rs("HWFMK3SZ")                     '�i�v�e�ʌ����ׂR�������
            If fldNameExist("HWFMK4SZ") Then .HWFMK4SZ = rs("HWFMK4SZ")                     '�i�v�e�ʌ����ׂS�������
            If fldNameExist("HWFMK2ZAR") Then .HWFMK2ZAR = fncNullCheck(rs("HWFMK2ZAR"))    '�i�v�e�ʌ����ׂQ���O�̈�
            If fldNameExist("HWFMK3ZAR") Then .HWFMK3ZAR = fncNullCheck(rs("HWFMK3ZAR"))    '�i�v�e�ʌ����ׂR���O�̈�
            If fldNameExist("HWFMK4ZAR") Then .HWFMK4ZAR = fncNullCheck(rs("HWFMK4ZAR"))    '�i�v�e�ʌ����ׂS���O�̈�
            If fldNameExist("HWFMK5B1") Then .HWFMK5B1 = fncNullCheck(rs("HWFMK5B1"))       '�i�v�e�ʌ����ׂT���E�P
            If fldNameExist("HWFMK5B1B") Then .HWFMK5B1B = fncNullCheck(rs("HWFMK5B1B"))    '�i�v�e�ʌ����ׂT���E�P��
            If fldNameExist("HWFMK5B2") Then .HWFMK5B2 = fncNullCheck(rs("HWFMK5B2"))       '�i�v�e�ʌ����ׂT���E�Q
            If fldNameExist("HWFMK5B2B") Then .HWFMK5B2B = fncNullCheck(rs("HWFMK5B2B"))    '�i�v�e�ʌ����ׂT���E�Q��
            If fldNameExist("HWFMK5B3") Then .HWFMK5B3 = fncNullCheck(rs("HWFMK5B3"))       '�i�v�e�ʌ����ׂT���E�R
            If fldNameExist("HWFMK5B3B") Then .HWFMK5B3B = fncNullCheck(rs("HWFMK5B3B"))    '�i�v�e�ʌ����ׂT���E�R��
            If fldNameExist("HWFMK6B1") Then .HWFMK6B1 = fncNullCheck(rs("HWFMK6B1"))       '�i�v�e�ʌ����ׂU���E�P
            If fldNameExist("HWFMK6B1B") Then .HWFMK6B1B = fncNullCheck(rs("HWFMK6B1B"))    '�i�v�e�ʌ����ׂU���E�P��
            If fldNameExist("HWFMK6B2") Then .HWFMK6B2 = fncNullCheck(rs("HWFMK6B2"))       '�i�v�e�ʌ����ׂU���E�Q
            If fldNameExist("HWFMK6B2B") Then .HWFMK6B2B = fncNullCheck(rs("HWFMK6B2B"))    '�i�v�e�ʌ����ׂU���E�Q��
            If fldNameExist("HWFMK6B3") Then .HWFMK6B3 = fncNullCheck(rs("HWFMK6B3"))       '�i�v�e�ʌ����ׂU���E�R
            If fldNameExist("HWFMK6B3B") Then .HWFMK6B3B = fncNullCheck(rs("HWFMK6B3B"))    '�i�v�e�ʌ����ׂU���E�R��
            If fldNameExist("HWFMK7MC") Then .HWFMK7MC = HWFMK7MC                           '�i�v�e�ʌ����ׂV�ʎw��
            If fldNameExist("HWFMK7SI") Then .HWFMK7SI = fncNullCheck(rs("HWFMK7SI"))       '�i�v�e�ʌ����ׂV�T�C�Y
            If fldNameExist("HWFMK7MX") Then .HWFMK7MX = fncNullCheck(rs("HWFMK7MX"))       '�i�v�e�ʌ����ׂV���
            If fldNameExist("HWFMK7SZ") Then .HWFMK7SZ = HWFMK7SZ                           '�i�v�e�ʌ����ׂV�������
            If fldNameExist("HWFMK7ZA") Then .HWFMK7ZA = fncNullCheck(rs("HWFMK7ZA"))       '�i�v�e�ʌ����ׂV���O�̈�
            If fldNameExist("HWFMK7HT") Then .HWFMK7HT = HWFMK7HT                           '�i�v�e�ʌ����ׂV�ۏؕ��@�Q��
            If fldNameExist("HWFMK7HS") Then .HWFMK7HS = HWFMK7HS                           '�i�v�e�ʌ����ׂV�ۏؕ��@�Q��
            If fldNameExist("HWFMK8MC") Then .HWFMK8MC = HWFMK8MC                           '�i�v�e�ʌ����ׂW�ʎw��
            If fldNameExist("HWFMK8SI") Then .HWFMK8SI = fncNullCheck(rs("HWFMK8SI"))       '�i�v�e�ʌ����ׂW�T�C�Y
            If fldNameExist("HWFMK8MX") Then .HWFMK8MX = fncNullCheck(rs("HWFMK8MX"))       '�i�v�e�ʌ����ׂW���
            If fldNameExist("HWFMK8SZ") Then .HWFMK8SZ = HWFMK8SZ                           '�i�v�e�ʌ����ׂW�������
            If fldNameExist("HWFMK8ZA") Then .HWFMK8ZA = fncNullCheck(rs("HWFMK8ZA"))       '�i�v�e�ʌ����ׂW���O�̈�
            If fldNameExist("HWFMK8HT") Then .HWFMK8HT = HWFMK8HT                           '�i�v�e�ʌ����ׂW�ۏؕ��@�Q��
            If fldNameExist("HWFMK8HS") Then .HWFMK8HS = HWFMK8HS                           '�i�v�e�ʌ����ׂW�ۏؕ��@�Q��
            If fldNameExist("HWFMK9MC") Then .HWFMK9MC = HWFMK9MC                           '�i�v�e�ʌ����ׂX�ʎw��
            If fldNameExist("HWFMK9SI") Then .HWFMK9SI = fncNullCheck(rs("HWFMK9SI"))       '�i�v�e�ʌ����ׂX�T�C�Y
            If fldNameExist("HWFMK9MX") Then .HWFMK9MX = fncNullCheck(rs("HWFMK9MX"))       '�i�v�e�ʌ����ׂX���
            If fldNameExist("HWFMK9SZ") Then .HWFMK9SZ = HWFMK9SZ                           '�i�v�e�ʌ����ׂX�������
            If fldNameExist("HWFMK9ZA") Then .HWFMK9ZA = fncNullCheck(rs("HWFMK9ZA"))       '�i�v�e�ʌ����ׂX���O�̈�
            If fldNameExist("HWFMK9HT") Then .HWFMK9HT = HWFMK9HT                           '�i�v�e�ʌ����ׂX�ۏؕ��@�Q��
            If fldNameExist("HWFMK9HS") Then .HWFMK9HS = HWFMK9HS                           '�i�v�e�ʌ����ׂX�ۏؕ��@�Q��
            If fldNameExist("HWFMK10MC") Then .HWFMK10MC = HWFMK10MC                        '�i�v�e�ʌ����ׂP�O�ʎw��
            If fldNameExist("HWFMK10SI") Then .HWFMK10SI = fncNullCheck(rs("HWFMK10SI"))    '�i�v�e�ʌ����ׂP�O�T�C�Y
            If fldNameExist("HWFMK10MX") Then .HWFMK10MX = fncNullCheck(rs("HWFMK10MX"))    '�i�v�e�ʌ����ׂP�O���
            If fldNameExist("HWFMK10SZ") Then .HWFMK10SZ = HWFMK10SZ                        '�i�v�e�ʌ����ׂP�O�������
            If fldNameExist("HWFMK10ZA") Then .HWFMK10ZA = fncNullCheck(rs("HWFMK10ZA"))    '�i�v�e�ʌ����ׂP�O���O�̈�
            If fldNameExist("HWFMK10HT") Then .HWFMK10HT = HWFMK10HT                        '�i�v�e�ʌ����ׂP�O�ۏؕ��@�Q��
            If fldNameExist("HWFMK10HS") Then .HWFMK10HS = HWFMK10HS                        '�i�v�e�ʌ����ׂP�O�ۏؕ��@�Q��
            If fldNameExist("HWFMK11MC") Then .HWFMK11MC = HWFMK11MC                        '�i�v�e�ʌ����ׂP�P�ʎw��
            If fldNameExist("HWFMK11SI") Then .HWFMK11SI = fncNullCheck(rs("HWFMK11SI"))    '�i�v�e�ʌ����ׂP�P�T�C�Y
            If fldNameExist("HWFMK11MX") Then .HWFMK11MX = fncNullCheck(rs("HWFMK11MX"))    '�i�v�e�ʌ����ׂP�P���
            If fldNameExist("HWFMK11SZ") Then .HWFMK11SZ = HWFMK11SZ                        '�i�v�e�ʌ����ׂP�P�������
            If fldNameExist("HWFMK11ZA") Then .HWFMK11ZA = fncNullCheck(rs("HWFMK11ZA"))    '�i�v�e�ʌ����ׂP�P���O�̈�
            If fldNameExist("HWFMK11HT") Then .HWFMK11HT = HWFMK11HT                        '�i�v�e�ʌ����ׂP�P�ۏؕ��@�Q��
            If fldNameExist("HWFMK11HS") Then .HWFMK11HS = HWFMK11HS                        '�i�v�e�ʌ����ׂP�P�ۏؕ��@�Q��
            If fldNameExist("HWFMK12MC") Then .HWFMK12MC = HWFMK12MC                        '�i�v�e�ʌ����ׂP�Q�ʎw��
            If fldNameExist("HWFMK12SI") Then .HWFMK12SI = fncNullCheck(rs("HWFMK12SI"))    '�i�v�e�ʌ����ׂP�Q�T�C�Y
            If fldNameExist("HWFMK12MX") Then .HWFMK12MX = fncNullCheck(rs("HWFMK12MX"))    '�i�v�e�ʌ����ׂP�Q���
            If fldNameExist("HWFMK12SZ") Then .HWFMK12SZ = HWFMK12SZ                        '�i�v�e�ʌ����ׂP�Q�������
            If fldNameExist("HWFMK12ZA") Then .HWFMK12ZA = fncNullCheck(rs("HWFMK12ZA"))    '�i�v�e�ʌ����ׂP�Q���O�̈�
            If fldNameExist("HWFMK12HT") Then .HWFMK12HT = HWFMK12HT                        '�i�v�e�ʌ����ׂP�Q�ۏؕ��@�Q��
            If fldNameExist("HWFMK12HS") Then .HWFMK12HS = HWFMK12HS                        '�i�v�e�ʌ����ׂP�Q�ۏؕ��@�Q��
            If fldNameExist("HWFMK13MC") Then .HWFMK13MC = HWFMK13MC                        '�i�v�e�ʌ����ׂP�R�ʎw��
            If fldNameExist("HWFMK13SI") Then .HWFMK13SI = fncNullCheck(rs("HWFMK13SI"))    '�i�v�e�ʌ����ׂP�R�T�C�Y
            If fldNameExist("HWFMK13MX") Then .HWFMK13MX = fncNullCheck(rs("HWFMK13MX"))    '�i�v�e�ʌ����ׂP�R���
            If fldNameExist("HWFMK13SZ") Then .HWFMK13SZ = HWFMK13SZ                        '�i�v�e�ʌ����ׂP�R�������
            If fldNameExist("HWFMK13ZA") Then .HWFMK13ZA = fncNullCheck(rs("HWFMK13ZA"))    '�i�v�e�ʌ����ׂP�R���O�̈�
            If fldNameExist("HWFMK13HT") Then .HWFMK13HT = HWFMK13HT                        '�i�v�e�ʌ����ׂP�R�ۏؕ��@�Q��
            If fldNameExist("HWFMK13HS") Then .HWFMK13HS = HWFMK13HS                        '�i�v�e�ʌ����ׂP�R�ۏؕ��@�Q��
            If fldNameExist("HWFMK14MC") Then .HWFMK14MC = HWFMK14MC                        '�i�v�e�ʌ����ׂP�S�ʎw��
            If fldNameExist("HWFMK14SI") Then .HWFMK14SI = fncNullCheck(rs("HWFMK14SI"))    '�i�v�e�ʌ����ׂP�S�T�C�Y
            If fldNameExist("HWFMK14MX") Then .HWFMK14MX = fncNullCheck(rs("HWFMK14MX"))    '�i�v�e�ʌ����ׂP�S���
            If fldNameExist("HWFMK14SZ") Then .HWFMK14SZ = HWFMK14SZ                        '�i�v�e�ʌ����ׂP�S�������
            If fldNameExist("HWFMK14ZA") Then .HWFMK14ZA = fncNullCheck(rs("HWFMK14ZA"))    '�i�v�e�ʌ����ׂP�S���O�̈�
            If fldNameExist("HWFMK14HT") Then .HWFMK14HT = HWFMK14HT                        '�i�v�e�ʌ����ׂP�S�ۏؕ��@�Q��
            If fldNameExist("HWFMK14HS") Then .HWFMK14HS = HWFMK14HS                        '�i�v�e�ʌ����ׂP�S�ۏؕ��@�Q��
            If fldNameExist("HWFMK15MC") Then .HWFMK15MC = HWFMK15MC                        '�i�v�e�ʌ����ׂP�T�ʎw��
            If fldNameExist("HWFMK15SI") Then .HWFMK15SI = fncNullCheck(rs("HWFMK15SI"))    '�i�v�e�ʌ����ׂP�T�T�C�Y
            If fldNameExist("HWFMK15MX") Then .HWFMK15MX = fncNullCheck(rs("HWFMK15MX"))    '�i�v�e�ʌ����ׂP�T���
            If fldNameExist("HWFMK15SZ") Then .HWFMK15SZ = HWFMK15SZ                        '�i�v�e�ʌ����ׂP�T�������
            If fldNameExist("HWFMK15ZA") Then .HWFMK15ZA = fncNullCheck(rs("HWFMK15ZA"))    '�i�v�e�ʌ����ׂP�T���O�̈�
            If fldNameExist("HWFMK15HT") Then .HWFMK15HT = HWFMK15HT                        '�i�v�e�ʌ����ׂP�T�ۏؕ��@�Q��
            If fldNameExist("HWFMK15HS") Then .HWFMK15HS = HWFMK15HS                        '�i�v�e�ʌ����ׂP�T�ۏؕ��@�Q��
            If fldNameExist("HWFSPVMXN") Then .HWFSPVMXN = fncNullCheck(rs("HWFSPVMXN"))    '�i�v�e�r�o�u�e�d���
            If fldNameExist("HWFSPVAMN") Then .HWFSPVAMN = fncNullCheck(rs("HWFSPVAMN"))    '�i�v�e�r�o�u�e�d����
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME028 = FUNCTION_RETURN_SUCCESS

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

Public Function DBDRV_GetTBCME018(records() As typ_TBCME018, formID$, HIN() As tFullHinban) As FUNCTION_RETURN
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
        
        '2009/08 SUMCO Akizuki �ǉ�
        Case "f_cmbc053_1"           '�u�w��������ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "

        'Add Start 2010/12/17 SMPK Miyata
        Case "f_cmbc054_1"           '�uCu-deco���ѓ��́v
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        'Add End   2010/12/17 SMPK Miyata

    End Select
    
    sqlBase = sqlBase & "From TBCME018"
    
    '''SQL��Where���쐬
    For i = 0 To UBound(HIN)
        With HIN(i)
            key = key & "'" & .hinban & Format(.mnorevno, "00000") & .factory & .opecond & "'"
            If i <> UBound(HIN) Then
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
            If fldNameExist("HINBAN") Then .hinban = rs("HINBAN")                       ' �i��
            If fldNameExist("MNOREVNO") Then .mnorevno = rs("MNOREVNO")                 ' ���i�ԍ������ԍ�
            If fldNameExist("FACTORY") Then .factory = rs("FACTORY")                    ' �H��
            If fldNameExist("OPECOND") Then .opecond = rs("OPECOND")                    ' ���Ə���
            If fldNameExist("HMGSTRRNO") Then .HMGSTRRNO = rs("HMGSTRRNO")              ' �i�Ǘ��d�l�o�^�˗��ԍ�
            If fldNameExist("HMGSTFNO") Then .HMGSTFNO = rs("HMGSTFNO")                 ' �i�Ǘ��Ј��m��
            If fldNameExist("HMGSXSNO") Then .HMGSXSNO = rs("HMGSXSNO")                 ' �i�Ǘ��r�w���i�ԍ�
            If fldNameExist("HMGSXSNE") Then .HMGSXSNE = fncNullCheck(rs("HMGSXSNE"))   ' �i�Ǘ��r�w���i�ԍ��}��
            If fldNameExist("CONFLAG") Then .CONFLAG = rs("CONFLAG")                    ' �m�F�t���O
            If fldNameExist("REINFLAG") Then .REINFLAG = rs("REINFLAG")                 ' �ĕt�^�t���O
            If fldNameExist("HSXTRWKB") Then .HSXTRWKB = rs("HSXTRWKB")                 ' �i�r�w�����ۋ敪
            If fldNameExist("HSXTYPE") Then .HSXTYPE = rs("HSXTYPE")                    ' �i�r�w�^�C�v
            If fldNameExist("KSXTYPKW") Then .KSXTYPKW = rs("KSXTYPKW")                 ' �i�r�w�^�C�v�������@
            If fldNameExist("HSXDOP") Then .HSXDOP = rs("HSXDOP")                       ' �i�r�w�h�[�p���g
            If fldNameExist("HSXRMIN") Then .HSXRMIN = fncNullCheck(rs("HSXRMIN"))      ' �i�r�w���R����
            If fldNameExist("HSXRMAX") Then .HSXRMAX = fncNullCheck(rs("HSXRMAX"))      ' �i�r�w���R���
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

'*** UPDATE �� Y.SIMIZU 2005/10/1
'�T�v      :�e�[�u���uTBCME036�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :records()     ,O  ,typ_TBCME036    ,���o���R�[�h
'          :formID        ,I  ,String          ,�g�p�t�H�[��ID
'          :sqlOrder      ,I  ,tFullHinban     ,���o�i�ԁi�z��j
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,���o�̐���
'����      :�����L�[�́A�HINBAN�+�uMNOREVNO�v+�uFACTORY�v+�uOPECOND�v�̕�����Ƃ���
'����      :05/03/01 ooba
Public Function DBDRV_GetTBCME036(records() As typ_TBCME036, formID$, HIN() As tFullHinban) As FUNCTION_RETURN
    Dim sql         As String           'SQL�S��
    Dim sqlBase     As String           'SQL��{��(WHERE�߂̑O�܂�)
    Dim sqlWhere    As String           'SQLWhere��
    Dim rs          As OraDynaset       'RecordSet
    Dim recCnt      As Long             '���R�[�h��
    Dim key         As String           '����KEY
    Dim i           As Long             'ٰ�߶���

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_kensa_SQL.bas -- Function DBDRV_GetTBCME036"

'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� UPD By Systech Start
''    Select Case formID
''        Case "f_cmbc026_1"           '�uGD���ѓ��́v
''            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HSXGDLINE,HWFGDLINE "
''    End Select
    'GD���ѓ��͗p�̂悤�����A���������Ǘ��ɒǉ����ꂽ���ڂ�OSF���ѓ��́A��������A
    'WF������������ł��g�p����̂ŉ�ʎw��ł̌����𖳂��ɂ���
    sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HSXGDLINE,HWFGDLINE "
    sqlBase = sqlBase & ",HSXLDLRMN,HSXLDLRMX,HWFLDLRMN,HWFLDLRMX "
    sqlBase = sqlBase & ",HSXOF1ARPTK,HSXOFARMIN,HSXOFARMAX,HSXOFARMHMX "
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� UPD By Systech End
    'Add Start 2011/01/27 SMPK Miyata
    sqlBase = sqlBase & ",HSXCJLTBND "
    'Add End   2011/01/27 SMPK Miyata

    sqlBase = sqlBase & "From TBCME036"
    
    '''SQL��Where���쐬
    For i = 0 To UBound(HIN)
        With HIN(i)
            key = key & "'" & .hinban & Format(.mnorevno, "00000") & .factory & .opecond & "'"
            If i <> UBound(HIN) Then
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
        DBDRV_GetTBCME036 = FUNCTION_RETURN_FAILURE
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
            If fldNameExist("HINBAN") Then .hinban = rs("HINBAN")                           '�i��
            If fldNameExist("MNOREVNO") Then .mnorevno = rs("MNOREVNO")                     '���i�ԍ������ԍ�
            If fldNameExist("FACTORY") Then .factory = rs("FACTORY")                        '�H��
            If fldNameExist("OPECOND") Then .opecond = rs("OPECOND")                        '���Ə���
            If fldNameExist("HSXGDLINE") Then .HSXGDLINE = fncNullCheck(rs("HSXGDLINE"))    '�i�Ǘ��d�l�o�^�˗��ԍ�
            If fldNameExist("HWFGDLINE") Then .HWFGDLINE = fncNullCheck(rs("HWFGDLINE"))    '�i�Ǘ��Ј��m��
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
            If fldNameExist("HSXLDLRMN") Then .HSXLDLRMN = fncNullCheck(rs("HSXLDLRMN"))    '�iSXL/DL�A��0����
            If fldNameExist("HSXLDLRMX") Then .HSXLDLRMX = fncNullCheck(rs("HSXLDLRMX"))    '�iSXL/DL�A��0���
            If fldNameExist("HWFLDLRMN") Then .HWFLDLRMN = fncNullCheck(rs("HWFLDLRMN"))    '�iWFL/DL�A��0����
            If fldNameExist("HWFLDLRMX") Then .HWFLDLRMX = fncNullCheck(rs("HWFLDLRMX"))    '�iWFL/DL�A��0���
            If fldNameExist("HSXOF1ARPTK") Then If IsNull(rs("HSXOF1ARPTK")) = False Then .HSXOF1ARPTK = rs("HSXOF1ARPTK")                '�iSXOSF1(ArAN)�p�^���敪
            If fldNameExist("HSXOFARMIN") Then .HSXOFARMIN = fncNullCheck(rs("HSXOFARMIN"))     '�iSXOSF(ArAN)����
            If fldNameExist("HSXOFARMAX") Then .HSXOFARMAX = fncNullCheck(rs("HSXOFARMAX"))     '�iSXOSF(ArAN)���
            If fldNameExist("HSXOFARMHMX") Then .HSXOFARMHMX = fncNullCheck(rs("HSXOFARMHMX"))  '�iSXOSF(ArAN)�ʓ�����
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
            'Add Start 2011/01/27 SMPK Miyata
            If fldNameExist("HSXCJLTBND") Then .HSXCJLTBND = fncNullCheck(rs("HSXCJLTBND"))  '�iSXL/CJLT�o���h��
            'Add End   2011/01/27 SMPK Miyata
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME036 = FUNCTION_RETURN_SUCCESS

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
'*** UPDATE �� Y.SIMIZU 2005/10/1


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
''Add Start 2011/07/13 LT10�����Z����ǉ� T.Koi(SETsw)
        sql = sql & "CRYREST10CS='" & .CRYREST10CS & "', "        ' �����������сiLT10)
''Add End   2011/07/13 LT10�����Z����ǉ� T.Koi(SETsw)
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
'          :records()     ,O  ,typ_XSDCS    ,���o���R�[�h
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
    'Chg Start 2010/12/17 SMPK Miyata Cu-deco��������(C,CJ,CJLT,CJ2)�ǉ�
    'sqlBase = "Select CRYNUMCS, SMPKBNCS, TBKBNCS, REPSMPLIDCS, XTALCS, INPOSCS, HINBCS, REVNUMCS, FACTORYCS, OPECS, KTKBNCS, BLKKTFLAGCS, " & _
    '          " CRYSMPLIDRSCS, nvl(CRYSMPLIDRS1CS, 0) as CRYSMPLIDRS1CS, nvl(CRYSMPLIDRS2CS, 0) as CRYSMPLIDRS2CS, CRYINDRSCS, CRYRESRS1CS, nvl(CRYRESRS2CS, ' ') as CRYRESRS2CS, CRYSMPLIDOICS, CRYINDOICS, CRYRESOICS, " & _
    '          " CRYSMPLIDB1CS, CRYINDB1CS, CRYRESB1CS, CRYSMPLIDB2CS, CRYINDB2CS, CRYRESB2CS, CRYSMPLIDB3CS, CRYINDB3CS, CRYRESB3CS, CRYSMPLIDL1CS, " & _
    '          " CRYINDL1CS, CRYRESL1CS, CRYSMPLIDL2CS, CRYINDL2CS, CRYRESL2CS, CRYSMPLIDL3CS, CRYINDL3CS, CRYRESL3CS, CRYSMPLIDL4CS, CRYINDL4CS, " & _
    '          " CRYRESL4CS, CRYSMPLIDCSCS, CRYINDCSCS, CRYRESCSCS, CRYSMPLIDGDCS, CRYINDGDCS, CRYRESGDCS, CRYSMPLIDTCS, CRYINDTCS, CRYRESTCS, " & _
    '          " CRYSMPLIDEPCS, CRYINDEPCS, CRYRESEPCS, CRYSMPLIDXCS, CRYINDXCS, CRYRESXCS, SMPLNUMCS, SMPLPATCS, nvl(TSTAFFCS, ' ') as TSTAFFCS, TDAYCS, nvl(KSTAFFCS, ' ') as KSTAFFCS, KDAYCS, nvl(SNDKCS, ' ') as SNDKCS, nvl(SNDDAYCS, sysdate) as SNDDAYCS "
    sqlBase = "Select CRYNUMCS, SMPKBNCS, TBKBNCS, REPSMPLIDCS, XTALCS, INPOSCS, HINBCS, REVNUMCS, FACTORYCS, OPECS, KTKBNCS, BLKKTFLAGCS, " & _
              " CRYSMPLIDRSCS, nvl(CRYSMPLIDRS1CS, 0) as CRYSMPLIDRS1CS, nvl(CRYSMPLIDRS2CS, 0) as CRYSMPLIDRS2CS, CRYINDRSCS, CRYRESRS1CS, nvl(CRYRESRS2CS, ' ') as CRYRESRS2CS, CRYSMPLIDOICS, CRYINDOICS, CRYRESOICS, " & _
              " CRYSMPLIDB1CS, CRYINDB1CS, CRYRESB1CS, CRYSMPLIDB2CS, CRYINDB2CS, CRYRESB2CS, CRYSMPLIDB3CS, CRYINDB3CS, CRYRESB3CS, CRYSMPLIDL1CS, " & _
              " CRYINDL1CS, CRYRESL1CS, CRYSMPLIDL2CS, CRYINDL2CS, CRYRESL2CS, CRYSMPLIDL3CS, CRYINDL3CS, CRYRESL3CS, CRYSMPLIDL4CS, CRYINDL4CS, " & _
              " CRYRESL4CS, CRYSMPLIDCSCS, CRYINDCSCS, CRYRESCSCS, CRYSMPLIDGDCS, CRYINDGDCS, CRYRESGDCS, CRYSMPLIDTCS, CRYINDTCS, CRYRESTCS, " & _
              " CRYSMPLIDEPCS, CRYINDEPCS, CRYRESEPCS, CRYSMPLIDXCS, CRYINDXCS, CRYRESXCS, " & _
              " CRYSMPLIDCCS, CRYINDCCS, CRYRESCCS, CRYSMPLIDCJCS, CRYINDCJCS, CRYRESCJCS, " & _
              " CRYSMPLIDCJLTCS , CRYINDCJLTCS, CRYRESCJLTCS, CRYSMPLIDCJ2CS, CRYINDCJ2CS, CRYRESCJ2CS, " & _
              " SMPLNUMCS, SMPLPATCS, nvl(TSTAFFCS, ' ') as TSTAFFCS, TDAYCS, nvl(KSTAFFCS, ' ') as KSTAFFCS, KDAYCS, nvl(SNDKCS, ' ') as SNDKCS, nvl(SNDDAYCS, sysdate) as SNDDAYCS "
    'Chg End   2010/12/17 SMPK Miyata
    sqlBase = sqlBase & ",QCKBNCS "
'Add Start 2011/07/13 LT10�����Z���� T.Koi(SETsw)
    sqlBase = sqlBase & ",CRYREST10CS "
'Add End   2011/07/13 LT10�����Z���� T.Koi(SETsw)
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
            
            ' �T���v��ID(X��)   2009/08 SUMCO Akizuki ����������ѓ��́@���ڒǉ�
            If IsNull(rs("CRYSMPLIDXCS")) = True Then
                .CRYSMPLIDXCS = 999999
            Else
                .CRYSMPLIDXCS = rs("CRYSMPLIDXCS")
            End If
            
            ' ���FLG(X��)      2009/08 SUMCO Akizuki ����������ѓ��́@���ڒǉ�
            If IsNull(rs("CRYINDXCS")) = True Then
                .CRYINDXCS = "0"
            Else
                .CRYINDXCS = rs("CRYINDXCS")
            End If
            
            ' ����FLG(X��)      2009/08 SUMCO Akizuki ����������ѓ��́@���ڒǉ�
            If IsNull(rs("CRYRESXCS")) = True Then
                .CRYRESXCS = "0"
            Else
                .CRYRESXCS = rs("CRYRESXCS")
            End If

            'Add Start 2010/12/17 SMPK Miyata
            If IsNull(rs("CRYSMPLIDCCS")) = False Then .CRYSMPLIDCCS = rs("CRYSMPLIDCCS")           ' �T���v��ID(C)
            If IsNull(rs("CRYINDCCS")) = False Then .CRYINDCCS = rs("CRYINDCCS")                    ' ���FLG(C)
            If IsNull(rs("CRYRESCCS")) = False Then .CRYRESCCS = rs("CRYRESCCS")                    ' ����FLG(C)
            If IsNull(rs("CRYSMPLIDCJCS")) = False Then .CRYSMPLIDCJCS = rs("CRYSMPLIDCJCS")        ' �T���v��ID(CJ)
            If IsNull(rs("CRYINDCJCS")) = False Then .CRYINDCJCS = rs("CRYINDCJCS")                 ' ���FLG(CJ)
            If IsNull(rs("CRYRESCJCS")) = False Then .CRYRESCJCS = rs("CRYRESCJCS")                 ' ����FLG(CJ)
            If IsNull(rs("CRYSMPLIDCJLTCS")) = False Then .CRYSMPLIDCJLTCS = rs("CRYSMPLIDCJLTCS")  ' �T���v��ID(CJLT)
            If IsNull(rs("CRYINDCJLTCS")) = False Then .CRYINDCJLTCS = rs("CRYINDCJLTCS")           ' ���FLG(CJLT)
            If IsNull(rs("CRYRESCJLTCS")) = False Then .CRYRESCJLTCS = rs("CRYRESCJLTCS")           ' ����FLG(CJLT)
            If IsNull(rs("CRYSMPLIDCJ2CS")) = False Then .CRYSMPLIDCJ2CS = rs("CRYSMPLIDCJ2CS")     ' �T���v��ID(CJ2)
            If IsNull(rs("CRYINDCJ2CS")) = False Then .CRYINDCJ2CS = rs("CRYINDCJ2CS")              ' ���FLG(CJ2)
            If IsNull(rs("CRYRESCJ2CS")) = False Then .CRYRESCJ2CS = rs("CRYRESCJ2CS")              ' ����FLG(CJ2)
            'Add End   2010/12/17 SMPK Miyata

            If IsNull(rs("SMPLNUMCS")) = False Then .SMPLNUMCS = rs("SMPLNUMCS")                ' �T���v������
            If IsNull(rs("SMPLPATCS")) = False Then .SMPLPATCS = rs("SMPLPATCS")                ' �T���v���p�^�[��
            If IsNull(rs("TSTAFFCS")) = False Then .TSTAFFCS = rs("TSTAFFCS")                   ' �o�^�Ј�ID
            If IsNull(rs("TDAYCS")) = False Then .TDAYCS = rs("TDAYCS")                         ' �o�^���t
            If IsNull(rs("KSTAFFCS")) = False Then .KSTAFFCS = rs("KSTAFFCS")                   ' �X�V�Ј�ID
            If IsNull(rs("KDAYCS")) = False Then .KDAYCS = rs("KDAYCS")                         ' �X�V���t
            If IsNull(rs("SNDKCS")) = False Then .SNDKCS = rs("SNDKCS")                         ' ���M�t���O
            If IsNull(rs("SNDDAYCS")) = False Then .SNDDAYCS = rs("SNDDAYCS")                   ' ���M���t

            ' �Ǘ��敪     2009/11/06�ǉ� SETsw kubota
            If IsNull(rs("QCKBNCS")) = False Then .QCKBNCS = rs("QCKBNCS")

'Add Start 2011/07/13 LT10�����Z���� T.Koi(SETsw)
            If IsNull(rs("CRYREST10CS")) = False Then .CRYREST10CS = rs("CRYREST10CS")
'Add End   2011/07/13 LT10�����Z���� T.Koi(SETsw)
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
