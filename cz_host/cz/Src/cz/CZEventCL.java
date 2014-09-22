package cz;

import java.util.EventObject;

/**
 * Eventを保持する 
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

    public static final int EV_F001 = 0xF001; // 実績データ通知

    public static final int EV_1001 = 0x1001; // 手動介入応答（４軸）
    public static final int EV_8001 = 0x8001; // 手動介入完了（４軸）

    public static final int EV_1003 = 0x1003; // 手動介入ＵＮＤＯ応答（４軸）
    public static final int EV_8003 = 0x8003; // 手動介入ＵＮＤＯ完了（４軸）

    public static final int EV_1009 = 0x1009; // 手動介入応答（４軸以外）
    public static final int EV_8009 = 0x8009; // 手動介入完了（４軸以外）

    public static final int EV_100B = 0x100B; // 手動介入ＵＮＤＯ応答（４軸以外）
    public static final int EV_800B = 0x800B; // 手動介入ＵＮＤＯ完了（４軸以外）

    public static final int EV_1011 = 0x1011; // 特定プロセス変更応答
    public static final int EV_8015 = 0x8015; // 特定プロセス変更完了通知

    public static final int EV_1021 = 0x1021; // 生波形データ採取応答
    public static final int EV_8021 = 0x8021; // 生波形データ採取通知

    public static final int EV_1023 = 0x1023; // ＣＣＤカメラ画像保存応答
    public static final int EV_8023 = 0x8023; // ＣＣＤカメラ画像保存完了

    public static final int EV_1031 = 0x1031; // 電源変更応答
    public static final int EV_8031 = 0x8031; // 電源変更完了通知

    public static final int EV_1041 = 0x1041; // プロセス変更応答
    public static final int EV_8041 = 0x8041; // プロセス変更完了通知

    public static final int EV_1051 = 0x1051; // 制御モード変更応答
    public static final int EV_8051 = 0x8051; // 制御モード変更完了通知

    public static final int EV_1093 = 0x1093; // 引上げ条件登録応答
    public static final int EV_8091 = 0x8091; // 引上げ条件登録通知

    public static final int EV_1099 = 0x1099; // 取出しテーブル設定応答
    public static final int EV_8099 = 0x8099; // 取出しテーブル登録通知

    public static final int EV_1063 = 0x1063; // 制御テーブル更新応答

    public static final int EV_1083 = 0x1083; // 操業定数更新応答

    public static final int EV_1217 = 0x1217; // 操業定数更新可否問合せ応答

    public static final int EV_1219 = 0x1219; // 操業定数更新作業終了通知応答

    public static final int EV_1221 = 0x1221; // 制御テーブル更新可否問合応答

    public static final int EV_1223 = 0x1223; // 制御テーブル更新可否問合応答

    public static final int EV_1237 = 0x1237; // 制御テーブルグループ名変更応答

    public static final int EV_1239 = 0x1239; // 制御テーブルタイトル変更応答

    public static final int EV_1241 = 0x1241; // 制御テーブル定義更新応答

    public static final int EV_1247 = 0x1247; // 操業定数項目名変更応答

    public static final int EV_1261 = 0x1261; // CCDカメラモニタ切替

    public static final int EV_1206 = 0x1206; // 制御テーブル未登録通知

    public static final int EV_1005 = 0x1005; // 炉前手動介入開始通知（４軸）
    public static final int EV_8005 = 0x8005; // 炉前手動介入終了通知（４軸）

    public static final int EV_100D = 0x100D; // 炉前手動介入開始通知（４軸以外）
    public static final int EV_800D = 0x800D; // 炉前手動介入終了通知（４軸以外）

    public static final int EV_1200 = 0x1200; // 制御テーブル送信開始
    public static final int EV_1201 = 0x1201; // 制御テーブル要求
    public static final int EV_1202 = 0x1202; // 制御テーブル通知（初期時）
    public static final int EV_1204 = 0x1204; // 制御テーブル送信終了通知

    public static final int EV_F007 = 0xF007; // 異常項目通知
    public static final int EV_F009 = 0xF009; // 炉体状況通知

    private Object obj   = null;
    private int    event = -1;

    // ---------- コンストラクタ ---------------------------
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
