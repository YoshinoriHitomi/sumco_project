package cz;

import java.io.Serializable;

/**
 *  ���Ƃo�u���ъǗ�
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/10/21)
 */
public class CZSystemPvControl implements Serializable 
{
    public String   batch;          //�o�b�`�ԍ�
    public String	t_name;			//�e�[�u����
    public String	s_start;		//�̎�J�n����
    public String	s_end;			//�̎�I������
    public int		m_flg;			//�Ԉ����L��
    public int		m_sumi;			//�Ԉ�����
    public int		mo_flg;			//�l�n�ۑ��t���O
    public String	mo_date;		//�l�n�ۑ�����
}
