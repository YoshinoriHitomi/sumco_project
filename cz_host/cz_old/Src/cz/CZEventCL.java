package cz;

import java.util.EventObject;

/**
 * Event��ێ����� 
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZEventCL extends EventObject {    
    public static final int TIME_OUT   = 0;

    public static final int SYS_MESSAGE = 1;

    public static final int PV_ERROR   = 100;
    public static final int RO_CHANGE  = 200;
    public static final int PV_RECEIVE = 300;

    public static final int OT_GET_HAITA = 400;
    public static final int OT_PUT_HAITA = 401;

    public static final int CT_GET_HAITA = 410;
    public static final int CT_PUT_HAITA = 411;

    public static final int EV_F001 = 0xF001; // ���уf�[�^�ʒm

    public static final int EV_1001 = 0x1001; // �蓮��������i�S���j
    public static final int EV_8001 = 0x8001; // �蓮��������i�S���j

    public static final int EV_1003 = 0x1003; // �蓮����t�m�c�n�����i�S���j
    public static final int EV_8003 = 0x8003; // �蓮����t�m�c�n�����i�S���j

    public static final int EV_1009 = 0x1009; // �蓮��������i�S���ȊO�j
    public static final int EV_8009 = 0x8009; // �蓮��������i�S���ȊO�j

    public static final int EV_100B = 0x100B; // �蓮����t�m�c�n�����i�S���ȊO�j
    public static final int EV_800B = 0x800B; // �蓮����t�m�c�n�����i�S���ȊO�j

    public static final int EV_1011 = 0x1011; // ����v���Z�X�ύX����
    public static final int EV_8015 = 0x8015; // ����v���Z�X�ύX�����ʒm

    public static final int EV_1021 = 0x1021; // ���g�`�f�[�^�̎扞��
    public static final int EV_8021 = 0x8021; // ���g�`�f�[�^�̎�ʒm

    public static final int EV_1023 = 0x1023; // �b�b�c�J�����摜�ۑ�����
    public static final int EV_8023 = 0x8023; // �b�b�c�J�����摜�ۑ�����

    public static final int EV_1031 = 0x1031; // �d���ύX����
    public static final int EV_8031 = 0x8031; // �d���ύX�����ʒm

    public static final int EV_1041 = 0x1041; // �v���Z�X�ύX����
    public static final int EV_8041 = 0x8041; // �v���Z�X�ύX�����ʒm

    public static final int EV_1051 = 0x1051; // ���䃂�[�h�ύX����
    public static final int EV_8051 = 0x8051; // ���䃂�[�h�ύX�����ʒm

    public static final int EV_1093 = 0x1093; // ���グ�����o�^����
    public static final int EV_8091 = 0x8091; // ���グ�����o�^�ʒm

    public static final int EV_1099 = 0x1099; // ��o���e�[�u���ݒ艞��
    public static final int EV_8099 = 0x8099; // ��o���e�[�u���o�^�ʒm

    public static final int EV_1063 = 0x1063; // ����e�[�u���X�V����

    public static final int EV_1083 = 0x1083; // ���ƒ萔�X�V����

    public static final int EV_1217 = 0x1217; // ���ƒ萔�X�V�ۖ⍇������

    public static final int EV_1219 = 0x1219; // ���ƒ萔�X�V��ƏI���ʒm����

    public static final int EV_1221 = 0x1221; // ����e�[�u���X�V�ۖ⍇����

    public static final int EV_1223 = 0x1223; // ����e�[�u���X�V�ۖ⍇����

    public static final int EV_1237 = 0x1237; // ����e�[�u���O���[�v���ύX����

    public static final int EV_1239 = 0x1239; // ����e�[�u���^�C�g���ύX����

    public static final int EV_1241 = 0x1241; // ����e�[�u����`�X�V����

    public static final int EV_1247 = 0x1247; // ���ƒ萔���ږ��ύX����

    public static final int EV_1261 = 0x1261; // CCD�J�������j�^�ؑ�

    public static final int EV_1206 = 0x1206; // ����e�[�u�����o�^�ʒm

    public static final int EV_1005 = 0x1005; // �F�O�蓮����J�n�ʒm�i�S���j
    public static final int EV_8005 = 0x8005; // �F�O�蓮����I���ʒm�i�S���j

    public static final int EV_100D = 0x100D; // �F�O�蓮����J�n�ʒm�i�S���ȊO�j
    public static final int EV_800D = 0x800D; // �F�O�蓮����I���ʒm�i�S���ȊO�j

    public static final int EV_1200 = 0x1200; // ����e�[�u�����M�J�n
    public static final int EV_1201 = 0x1201; // ����e�[�u���v��
    public static final int EV_1202 = 0x1202; // ����e�[�u���ʒm�i�������j
    public static final int EV_1204 = 0x1204; // ����e�[�u�����M�I���ʒm

    public static final int EV_F007 = 0xF007; // �ُ퍀�ڒʒm
    public static final int EV_F009 = 0xF009; // �F�̏󋵒ʒm

    private Object obj   = null;
    private int    event = -1;

    // ---------- �R���X�g���N�^ ---------------------------
    CZEventCL(Object source,int ev){
        super(source);
        obj = source;
        event = ev;
    }
    
    public Object getObject(){
        return obj;
    }
    
    public int getEvent(){
        return event;
    }
}   
