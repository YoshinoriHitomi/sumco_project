package cz;

import java.io.Serializable;

/**
 *  �����グ����
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZSystemBtTemp implements Serializable 
{
    public String   batch;          //�o�b�`�ԍ�
    public String   pgid;           //PG-ID
    public String   t_time;         //�o�^����
    public int      renban;         //�A��
    public String   hinshu;         //�i��
    public String   houi;           //����
    public String   h_type;         //�^�C�v
    public String   hiteikou;       //���R
    public String   sanso;          //�_�f
    public String   gap;            //GAP
    public int      rutubo_kei;     //���c�{�a
    public int      chokkei;        //���a
    public int      hikiage_cho;    //���㒷
    public int      top_ar;         //�g�b�v�A���S��
    public int      pull_ar;        //�v���A���S��
    public int      i_sikomi;       //�d����
    public int      t_sikomi;       //�ǉ��d����
    public int      zaneki;         //�c�t��
    public int      no_youkai;      //T1(�n��)
    public int      no_hikiage;     //T2(����)
    public int      no_kaiten;      //T3(��])
    public int      no_toridasi;    //T4(��o)
    public int      no_aturyoku;    //T5(����)
    public int      no_teisu;       //T6(�萔) @@
    public int      pno_start;      //�X�^�[�g�v���Z�X
    public int      p_kaisi;        //�J�n
}
