package cz;

import czclass.CZClientResult_Proxy;
import czclass.CZResult;

/**
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZSystemResult implements Runnable {
    private CZClientResult_Proxy cz_re_px = null;

    private boolean life = false;
    private String  ro_name = null;

    //
    //
    //
    CZSystemResult(CZClientResult_Proxy ev){
        cz_re_px = ev;
    }

    //
    //
    //
    public void run(){

        life = true;

        ro_name = CZSystem.getRoName();
        if(cz_re_px.startResult(100, ro_name)){

            while(life){
                CZResult ev = cz_re_px.getResult();
                if(null == ev){
                    CZSystem.log("CZSystemResult run","getResult NULL !!");
                    return;
                }
                else {
//@@                    CZSystem.log("CZSystemResult run","getResult ["
//@@                        + ev.toString()        + "]["
//@@                        + ev.getEventCode()    + "]["
//@@                        + ev.getRoban()        + "]["
//@@                        + ev.getOperateClass() + "]["
//@@                        + ev.getStatus()       + "]["
//@@                        + ev.getValue()        + "]");

                    switch(ev.getEventCode()){

                        //�蓮��������i�S���j
                        case CZEventCL.EV_1001 : CZSystem.ev1001(ev);
                        break;

                        //�蓮��������i�S���j
                        case CZEventCL.EV_8001 : CZSystem.ev8001(ev);
                        break;

                        //�蓮����t�m�c�n�����i�S���j
                        case CZEventCL.EV_1003 : CZSystem.ev1003(ev);
                        break;
                        //�蓮����t�m�c�n�����i�S���j
                        case CZEventCL.EV_8003 : CZSystem.ev8003(ev);
                        break;

                        //�蓮��������i�S���ȊO�j
                        case CZEventCL.EV_1009 : CZSystem.ev1009(ev);
                        break;

                        //�蓮��������i�S���ȊO�j
                        case CZEventCL.EV_8009 : CZSystem.ev8009(ev);
                        break;

                        //�蓮����t�m�c�n�����i�S���ȊO�j
                        case CZEventCL.EV_100B : CZSystem.ev100B(ev);
                        break;

                        //�蓮����t�m�c�n�����i�S���ȊO�j
                        case CZEventCL.EV_800B : CZSystem.ev800B(ev);
                        break;

                        //����v���Z�X�ύX����
                        case CZEventCL.EV_1011 : CZSystem.ev1011(ev);
                        break;

                        //����v���Z�X�ύX�����ʒm
                        case CZEventCL.EV_8015 : CZSystem.ev8015(ev);
                        break;

                        //���g�`�f�[�^�̎扞��
                        case CZEventCL.EV_1021 : CZSystem.ev1021(ev);
                        break;

                        //���g�`�f�[�^�̎�ʒm
                        case CZEventCL.EV_8021 : CZSystem.ev8021(ev);
                        break;

                        //�b�b�c�J�����摜�ۑ�����
                        case CZEventCL.EV_1023 : CZSystem.ev1023(ev);
                        break;

                        //�b�b�c�J�����摜�ۑ�����
                        case CZEventCL.EV_8023 : CZSystem.ev8023(ev);
                        break;

                        //�d���ύX����
                        case CZEventCL.EV_1031 : CZSystem.ev1031(ev);
                        break;
                        //�d���ύX�����ʒm
                        case CZEventCL.EV_8031 : CZSystem.ev8031(ev);
                        break;

                        //�v���Z�X�ύX����
                        case CZEventCL.EV_1041 : CZSystem.ev1041(ev);
                        break;

                        //�v���Z�X�ύX�����ʒm
                        case CZEventCL.EV_8041 : CZSystem.ev8041(ev);
                        break;

                        //���䃂�[�h�ύX����
                        case CZEventCL.EV_1051 : CZSystem.ev1051(ev);
                        break;

                        //���䃂�[�h�ύX�����ʒm
                        case CZEventCL.EV_8051 : CZSystem.ev8051(ev);
                        break;

                        //���グ�����o�^����
                        case CZEventCL.EV_1093 : CZSystem.ev1093(ev);
                        break;

                        //���グ�����o�^�ʒm
                        case CZEventCL.EV_8091 : CZSystem.ev8091(ev);
                        break;

                        //��o���e�[�u���ݒ艞��
                        case CZEventCL.EV_1099 : CZSystem.ev1099(ev);
                        break;

                        //��o���e�[�u���o�^�ʒm
                        case CZEventCL.EV_8099 : CZSystem.ev8099(ev);
                        break;

                        //����e�[�u���X�V����
                        case CZEventCL.EV_1063 : CZSystem.ev1063(ev);
                        break;

                        //���ƒ萔�X�V����
                        case CZEventCL.EV_1083 : CZSystem.ev1083(ev);
                        break;

                        //���ƒ萔�X�V�ۖ⍇������
                        case CZEventCL.EV_1217 : CZSystem.ev1217(ev);
                        break;

                        //���ƒ萔�X�V��ƏI���ʒm����
                        case CZEventCL.EV_1219 : CZSystem.ev1219(ev);
                        break;

                        //����e�[�u���X�V�ۖ⍇����
                        case CZEventCL.EV_1221 : CZSystem.ev1221(ev);
                        break;

                        //����e�[�u����ƏI���ʒm����
                        case CZEventCL.EV_1223 : CZSystem.ev1223(ev);
                        break;

                        //����e�[�u���O���[�v���ύX����
                        case CZEventCL.EV_1237 : CZSystem.ev1237(ev);
                        break;

                        //����e�[�u���^�C�g���ύX����
                        case CZEventCL.EV_1239 : CZSystem.ev1239(ev);
                        break;

                        //����e�[�u����`�X�V����
                        case CZEventCL.EV_1241 : CZSystem.ev1241(ev);
                        break;

                        //���ƒ萔���ږ��ύX����
                        case CZEventCL.EV_1247 : CZSystem.ev1247(ev);
                        break;

                        //CCD�J�������j�^�ؑ�
                        case CZEventCL.EV_1261 : CZSystem.ev1261(ev);
                        break;

                        default :   break;
                    }
                }
            } // while end
        }
        else {
            CZSystem.log("CZSystemResult run","getResult FALSE !!");
        }

        cz_re_px.endResult();
    }
} 
