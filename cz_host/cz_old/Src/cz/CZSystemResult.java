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

                        //手動介入応答（４軸）
                        case CZEventCL.EV_1001 : CZSystem.ev1001(ev);
                        break;

                        //手動介入完了（４軸）
                        case CZEventCL.EV_8001 : CZSystem.ev8001(ev);
                        break;

                        //手動介入ＵＮＤＯ応答（４軸）
                        case CZEventCL.EV_1003 : CZSystem.ev1003(ev);
                        break;
                        //手動介入ＵＮＤＯ完了（４軸）
                        case CZEventCL.EV_8003 : CZSystem.ev8003(ev);
                        break;

                        //手動介入応答（４軸以外）
                        case CZEventCL.EV_1009 : CZSystem.ev1009(ev);
                        break;

                        //手動介入完了（４軸以外）
                        case CZEventCL.EV_8009 : CZSystem.ev8009(ev);
                        break;

                        //手動介入ＵＮＤＯ応答（４軸以外）
                        case CZEventCL.EV_100B : CZSystem.ev100B(ev);
                        break;

                        //手動介入ＵＮＤＯ完了（４軸以外）
                        case CZEventCL.EV_800B : CZSystem.ev800B(ev);
                        break;

                        //特定プロセス変更応答
                        case CZEventCL.EV_1011 : CZSystem.ev1011(ev);
                        break;

                        //特定プロセス変更完了通知
                        case CZEventCL.EV_8015 : CZSystem.ev8015(ev);
                        break;

                        //生波形データ採取応答
                        case CZEventCL.EV_1021 : CZSystem.ev1021(ev);
                        break;

                        //生波形データ採取通知
                        case CZEventCL.EV_8021 : CZSystem.ev8021(ev);
                        break;

                        //ＣＣＤカメラ画像保存応答
                        case CZEventCL.EV_1023 : CZSystem.ev1023(ev);
                        break;

                        //ＣＣＤカメラ画像保存完了
                        case CZEventCL.EV_8023 : CZSystem.ev8023(ev);
                        break;

                        //電源変更応答
                        case CZEventCL.EV_1031 : CZSystem.ev1031(ev);
                        break;
                        //電源変更完了通知
                        case CZEventCL.EV_8031 : CZSystem.ev8031(ev);
                        break;

                        //プロセス変更応答
                        case CZEventCL.EV_1041 : CZSystem.ev1041(ev);
                        break;

                        //プロセス変更完了通知
                        case CZEventCL.EV_8041 : CZSystem.ev8041(ev);
                        break;

                        //制御モード変更応答
                        case CZEventCL.EV_1051 : CZSystem.ev1051(ev);
                        break;

                        //制御モード変更完了通知
                        case CZEventCL.EV_8051 : CZSystem.ev8051(ev);
                        break;

                        //引上げ条件登録応答
                        case CZEventCL.EV_1093 : CZSystem.ev1093(ev);
                        break;

                        //引上げ条件登録通知
                        case CZEventCL.EV_8091 : CZSystem.ev8091(ev);
                        break;

                        //取出しテーブル設定応答
                        case CZEventCL.EV_1099 : CZSystem.ev1099(ev);
                        break;

                        //取出しテーブル登録通知
                        case CZEventCL.EV_8099 : CZSystem.ev8099(ev);
                        break;

                        //制御テーブル更新応答
                        case CZEventCL.EV_1063 : CZSystem.ev1063(ev);
                        break;

                        //操業定数更新応答
                        case CZEventCL.EV_1083 : CZSystem.ev1083(ev);
                        break;

                        //操業定数更新可否問合せ応答
                        case CZEventCL.EV_1217 : CZSystem.ev1217(ev);
                        break;

                        //操業定数更新作業終了通知応答
                        case CZEventCL.EV_1219 : CZSystem.ev1219(ev);
                        break;

                        //制御テーブル更新可否問合応答
                        case CZEventCL.EV_1221 : CZSystem.ev1221(ev);
                        break;

                        //制御テーブル作業終了通知応答
                        case CZEventCL.EV_1223 : CZSystem.ev1223(ev);
                        break;

                        //制御テーブルグループ名変更応答
                        case CZEventCL.EV_1237 : CZSystem.ev1237(ev);
                        break;

                        //制御テーブルタイトル変更応答
                        case CZEventCL.EV_1239 : CZSystem.ev1239(ev);
                        break;

                        //制御テーブル定義更新応答
                        case CZEventCL.EV_1241 : CZSystem.ev1241(ev);
                        break;

                        //操業定数項目名変更応答
                        case CZEventCL.EV_1247 : CZSystem.ev1247(ev);
                        break;

                        //CCDカメラモニタ切替
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
