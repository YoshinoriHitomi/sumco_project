package cz;

import czclass.CZClientEvent_Proxy;
import czclass.CZEvent;
import czclass.CZNativeGetData_Proxy;

/**
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZSystemEvent implements Runnable {

    private static CZClientEvent_Proxy      cz_ev_px = null;
    private static CZNativeGetData_Proxy    cz_gd_px = null;
    
    private static boolean life = false;

    //
    //
    //
    CZSystemEvent(CZClientEvent_Proxy ev,CZNativeGetData_Proxy gd){
        cz_ev_px = ev;
        cz_gd_px = gd;
    }

    //
    //
    //
    public void stop(){
        life = false;
    }

    //
    //
    //
    public void run(){

        life = true;

        if(!cz_ev_px.startEvent(100)){
            CZSystem.log("CZSystemEvent run","FALSE !!");
            return ;
        }

        while(life){

            CZEvent ev = cz_ev_px.getEvent();
            if(null == ev){
                CZSystem.log("CZSystemEvent run","NULL !!");
                return;
            }
            else {
/*
                CZSystem.log("CZSystemEvent run","CZEvent ["
                                + ev.toString()     + "]["
                                + ev.getEventCode() + "]["
                                + ev.getRoban()     + "]");
*/

                switch(ev.getEventCode()){
                    //���уf�[�^�ʒm
                    case CZEventCL.EV_F001 : CZSystem.evF001(ev.getRoban());
                    break;
                    //�ُ퍀�ڒʒm
                    case CZEventCL.EV_F007 : CZSystem.evF007(ev);
                    break;
                    //�F�̏󋵒ʒm
                    case CZEventCL.EV_F009 : CZSystem.evF009(ev);
                    break;

                    //�F�O�蓮����J�n�ʒm�i�S���j
                    case CZEventCL.EV_1005 : CZSystem.ev1005(ev);
                    break;

                    //�F�O�蓮����I���ʒm�i�S���j
                    case CZEventCL.EV_8005 : CZSystem.ev8005(ev);
                    break;

                    //�F�O�蓮����J�n�ʒm�i�S���ȊO�j
                    case CZEventCL.EV_100D : CZSystem.ev100D(ev);
                    break;

                    //�F�O�蓮����I���ʒm�i�S���ȊO�j
                    case CZEventCL.EV_800D : CZSystem.ev800D(ev);
                    break;

                    //����e�[�u�����o�^�ʒm    
                    case CZEventCL.EV_1206 : CZSystem.ev1206(ev);
                    break;

                    default :
                    break;

                } // switch end
            }
        } // while end
        cz_ev_px.endEvent();
    }
}
