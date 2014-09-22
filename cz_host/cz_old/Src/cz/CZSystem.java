package cz;

import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Properties;
import java.util.Vector;

import czclass.CZClientEvent_Proxy;
import czclass.CZClientResult_Proxy;
import czclass.CZEvent;
import czclass.CZMoList;
import czclass.CZNativeDengen;
import czclass.CZNativeGetData_Proxy;
import czclass.CZNativeHikiage;
import czclass.CZNativeRoHikiage;
import czclass.CZNativePv;
import czclass.CZNativeRoState;
import czclass.CZNativeCTState;
import czclass.CZNativeSTState;
import czclass.CZOperate_Proxy;
import czclass.CZParamControlDefine;
import czclass.CZParamControlT6Define;
import czclass.CZParamErrorDefine;
import czclass.CZParamErrorMsgDefine;
import czclass.CZParamHikiage;
import czclass.CZParamPVMabikiCng;
import czclass.CZParamT6Table;
import czclass.CZParamUnten;
import czclass.CZRaidStatus;
import czclass.CZResult;
import czclass.CZServer_Proxy;
import czclass.CZTableExchange_Proxy;
import czclass.CZRealNativeData_Proxy;
import czclass.CZNativeMRoState;
import czclass.CZRealNativeWatchItem;

//import horb.orb.*;

/*
 *   ＣＺシステムクラス
 *   共通クラス
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * @ 2008.10.08 H.Nagamine 制御テーブル変更履歴作成
 * @Update 2013.10.21 他基地参照機能 (@20131021)
 * @Update 2013.10.30 表示切り替え機能 (@20131030)
 */
public class CZSystem {
    
    // ランレベル   -1:
    private static int  run_level   = CZSystemDefine.NONE_RUN;

    // システムモード
    private static int  system_mode = CZSystemDefine.NONE_MODE;

    // データサーバ接続用
    private static String DB_DRIVER = null; //@@
    private static String DB_URL    = null; //@@
    private static String HOST      = null;
    private static String USER      = null;
    private static String PASSWD    = null;
//@@    private static String PORT      = null;
    
    private static String ERROR_MAX = null; //@@
    public static String CT_TABLE_CLOSE_TIME = null;
    public static String ALERM_DIALOG_CLOSE_TIME = null;
    public static String TIMER_FLG = null;

    public static String DISP_KETA_FLG = null;	// 20050725 炉番表示桁数変更

    public static String SOGYO_OUTPUT_PATH = null;	// 2006.6.8 操業比較

    public static String RECIPE_OUTPUT_PATH = null;	// 2008.1.15 レシピ内容出力

    public static String FILE_SRC_PATH = null;	// ファイル出力

    public static String HISTORY_DATA_PATH = null;	// 変更履歴出力

    public static String KIDO_DATA_PATH = null;	// 輝度データ出力

    // データサーバ接続用
    private static String MO_1_DIR      = null;
    private static String MO_2_DIR      = null;
    // HORB接続用
    private static CZClientEvent_Proxy  cz_ev_px    = null;
    private static Thread               cz_ev_px_th = null;

    private static CZClientResult_Proxy cz_re_px    = null;
    private static Thread               cz_re_px_th = null;

    private static CZNativeGetData_Proxy    cz_gd_px = null;

    private static CZOperate_Proxy          cz_op_px = null;

    private static CZTableExchange_Proxy    cz_tb_px = null;

    private static CZServer_Proxy           cz_sv_px = null;

	/* 2006.07.12 */
    private static CZRealNativeData_Proxy    cz_rl_px = null;

    // データサーバＲＡＩＤ状態
    private static CZRaidStatus raid1_stat = null;
    private static CZRaidStatus raid5_stat = null;

    //
    // イニシャルフラグ 
    private static boolean init_flag        = false;

    // 炉番List
    private static Vector ro_name_list      = null;
    // 炉番のHost名List
    private static Vector ro_host_list      = null;
    // 炉番のカメラ番号List
    private static Vector ro_camera_list    = null;
    // 炉のバージョンList
    private static Vector ro_ver_list       = null;

    //ＰＶ関係
    private static Vector pv_name_list      = null;

    //エラーメッセージ
    private static Vector error_message_list    = null;

    //操業定数関係
    private static Vector op_tb_lag_name_list   = null;
    private static Vector op_tb_mid_name_list   = null;
    private static Vector op_tb_sml_name_list   = null;

    //制御テーブル関係
    private static Vector ct_tb_name_list   = null;

    private static Vector ctT6LagNameList_  = null;     //T6大項目
    private static Vector ctT6MidNameList_  = null;     //T6中項目
    private static Vector ctT6NameList_     = null;     //T6項目

    // カレントの炉番のIndex
    private static int ro_no_idx = 0;

    // 終了用炉番(エンドメソッド用)
    private static String final_ro_no = null;

    // 引き上げ条件
    private static CZNativeHikiage current_bt_set       = null; 

    // 電源状態
    private static CZNativeDengen  current_power_stat   = null; 

    private static CZSystemPVNamePMM current_unten      = null;

    // ＰＶデータ読み込み時のスレッド
    private static Thread db_thread                 = null; 

    private static String   current_bt              = null;     //  バッチNo
    private static int  current_proc                = -1;       //  プロセスNo
    private static int  current_sub_proc            = -1;       //  サブプロセスNo
    private static int  current_proc_len            = -1;       //  プロセス連番
    private static int  current_proc_time           = -1;       //  プロセス時間
    private static int  current_sub_proc_time       = -1;       //  サブプロセス時間
    private static int  current_get_date_time       = -1;       //  採取日時
    private static int  current_main_heat_on_time   = -1;       //  メインヒータ電源オン時間
    private static int  current_condition_len       = -1;       //  引上げ条件内連番
    //  データ
    private static float    current_pv[]            = new float[CZPV.PV_MAX_LENGTH];

	private static int graph_cnt                    = 0;        // 表示グラフ枚数カウンタ
	private static int RoIndex                      = 0;        // 炉番INDEX
	
	public static double Client_ver_list       = 0.00; //@@@@@@@@
	public static String VERSION = null;
	
	private static boolean untenFlg                 = true;		// @20131030

    //
    //  初期化
    //
    public static synchronized boolean init(int mode,String comment){

        try{
            Properties prop =  new Properties();
            FileInputStream pros = new FileInputStream(CZSystemDefine.PROPERTY_FILE);
            prop.load(pros);

            prop.list(System.out);

            String r_level = prop.getProperty("RUNLEVEL");

            if(r_level.equals("ADMIN")){
                run_level = CZSystemDefine.ADMIN_RUN;
            }
            else if(r_level.equals("USER")){
                run_level = CZSystemDefine.USER_RUN;
            }
            // @20131021 RUNLEVEL追加(参照のみ権限)
            else if(r_level.equals("REFERENCE")){
                run_level = CZSystemDefine.REFERENCE_RUN;
            }
            // @20131021
            else{   
                run_level = CZSystemDefine.NONE_RUN;
            }

            if(CZSystemDefine.NONE_RUN == run_level) exit(-1,"CZSystem NO Run Level");


            HOST    = prop.getProperty("HOST");
            if(null == HOST) exit(-1,"CZSystem NO HOST");

            USER    = prop.getProperty("USER");
            if(null == USER) exit(-1,"CZSystem NO USER");

            PASSWD  = prop.getProperty("PASSWD");
            if(null == PASSWD) exit(-1,"CZSystem NO PASSWD");

/**@@
            PORT    = prop.getProperty("PORT");
            if(null == PORT) exit(-1,"CZSystem NO PORT");

            MO_1_DIR = prop.getProperty("MO_1_DIR");
            if(null == MO_1_DIR) exit(-1,"CZSystem NO MO_1_DIR");

            MO_2_DIR = prop.getProperty("MO_2_DIR");
            if(null == MO_2_DIR) exit(-1,"CZSystem NO MO_2_DIR");
@@**/
            DB_URL = prop.getProperty("DB_URL");
            if(null == DB_URL) exit(-1,"CZSystem NO DB_URL");

            ERROR_MAX = prop.getProperty("ERROR_MAX");
            if(null == ERROR_MAX) exit(-1,"CZSystem NO ERROR_MAX");
            
            CT_TABLE_CLOSE_TIME = prop.getProperty("CT_TABLE_CLOSE_TIME");
            if(null == CT_TABLE_CLOSE_TIME) exit(-1,"CZSystem NO CT_TABLE_CLOSE_TIME");
            
            ALERM_DIALOG_CLOSE_TIME = prop.getProperty("ALERM_DIALOG_CLOSE_TIME");
            if(null == ALERM_DIALOG_CLOSE_TIME) exit(-1,"CZSystem NO ALERM_DIALOG_CLOSE_TIME");

            TIMER_FLG = prop.getProperty("TIMER_FLG");
            if(null == TIMER_FLG) exit(-1,"CZSystem NO TIMER_FLG");
            
            // 20050725 炉番表示桁数変更
            DISP_KETA_FLG = prop.getProperty("DISP_KETA_FLG");
            if(null == DISP_KETA_FLG) exit(-1,"CZSystem NO DISP_KETA_FLG");

            // 2006.06.08 操業比較
            SOGYO_OUTPUT_PATH = prop.getProperty("SOGYO_OUTPUT_PATH");
            if(null == SOGYO_OUTPUT_PATH) exit(-1,"CZSystem NO SOGYO_OUTPUT_PATH");

            // 2008.01.15 レシピ内容出力
            RECIPE_OUTPUT_PATH = prop.getProperty("RECIPE_OUTPUT_PATH");
            if(null == RECIPE_OUTPUT_PATH) exit(-1,"CZSystem NO RECIPE_OUTPUT_PATH");

            // SRCファイル出力
            FILE_SRC_PATH = prop.getProperty("FILE_SRC_PATH");
            if(null == FILE_SRC_PATH) exit(-1,"CZSystem NO FILE_SRC_PATH");

            // 変更履歴出力
            HISTORY_DATA_PATH = prop.getProperty("HISTORY_DATA_PATH");
            if(null == HISTORY_DATA_PATH) exit(-1,"CZSystem NO HISTORY_DATA_PATH");

            // 輝度データ出力
            KIDO_DATA_PATH = prop.getProperty("KIDO_DATA_PATH");
            if(null == KIDO_DATA_PATH) exit(-1,"CZSystem NO KIDO_DATA_PATH");

            // @@@@@@@@ クライアントバージョン
            VERSION = prop.getProperty("VERSION");
            if(null == VERSION) exit(-1,"CZSystem NO VERSION");

        }
        catch(Exception e){
            exit(-1,"CZSystem NO Propertie File");
        }

        log("CZSystem INIT","RUN[" + run_level + "][" + mode + "][" + comment + "]");

        if(init_flag){
            log("CZSystem INIT","ALREADY [" + comment + "]");
            return false;
        }


        system_mode = mode;

        switch(system_mode){
            case CZSystemDefine.HOST_MODE :
                Runtime.getRuntime().addShutdownHook(new Thread() {
                    public void run() {
                        endApp();
                    }
                });

                if(!initLib()) exit(-1,"CZSystem initLib");
                if(!initHorb()) exit(-1,"CZSystem initHorb");

                CZSystemWatch watch = new CZSystemWatch();
                Thread watch_th = new Thread(watch);
                watch_th.start();

                break;

            case CZSystemDefine.CMS_MODE :
                break;

            case CZSystemDefine.LIB_MODE :
                if(!initLib()) exit(-1,"CZSystem initLib");
                break;

            default :
                exit(-1,"CZSystem NO Mode");
                break;
        }

        init_flag = true;

        // 炉index 0 番でスタート
        chgRo(0);
        log("CZSystem INIT","START[" + run_level + "][" + system_mode + "][" + comment + "]");
        return true;
    }

    //
    //  初期化 Lib
    //
    public static synchronized boolean initLib(){
        log("CZSystem initLib","----- START !! -----");

        int ret = 0;

        try{
            // 炉名称関係読み込み
            log("CZSystem initLib","START !! [炉名称関係読み込み]");
            ro_name_list    = new Vector();
            ro_host_list    = new Vector();
            ro_camera_list  = new Vector();
            ro_ver_list     = new Vector();
            ret             =  roRead();
            if(0 >= ret){
                exit(0,"roRead()  DATABASE ERROR No[" + ret + "]");
            }

            //ＰＶ関係読み込み
            log("CZSystem initLib","START !! [ＰＶ関係読み込み]");
            pv_name_list    = new Vector(130);
            ret             = pvNameRead();
            if(0 >= ret){
                exit(0,"pvNameRead()  DATABASE ERROR No[" + ret + "]");
            }

            //エラーメッセージ読み込み
            log("CZSystem initLib","START !! [エラーメッセージ読み込み]");
            error_message_list  = new Vector(1000);
            ret                 = errorMessageRead();
            if(0 >= ret){
                exit(0,"errorMessageRead()  DATABASE ERROR No[" + ret + "]");
            }

            //操業定数関係読み込み
            log("CZSystem initLib","START !! [操業定数関係読み込み 大]");
            op_tb_lag_name_list = new Vector();
            ret                 = opTblLagNameRead();
            if(0 >= ret){
                exit(0,"opTblLagNameRead()  DATABASE ERROR No[" + ret + "]");
            }

            log("CZSystem initLib","START !! [操業定数関係読み込み 中]");
            op_tb_mid_name_list = new Vector(20);
            ret                 = opTblMidNameRead();
            if(0 >= ret){
                exit(0,"opTblMidNameRead()  DATABASE ERROR No[" + ret + "]");
            }

            log("CZSystem initLib","START !! [操業定数関係読み込み 小]");
            op_tb_sml_name_list = new Vector(500);
            ret                 = opTblSmlNameRead();
            if(0 >= ret){
                exit(0,"opTblSmlNameRead()  DATABASE ERROR No[" + ret + "]");
            }

            //制御テーブル関係読み込み
            log("CZSystem initLib","START !! [制御テーブル関係読み込み]");
            ct_tb_name_list = new Vector(200);
            ret             = ctTblNameRead();
            if(0 >= ret){
                exit(0,"ctTblNameRead()  DATABASE ERROR No[" + ret + "]");
            }
//@@        T6関連は、0件でもＯＫにしておく

            log("CZSystem initLib","START !! [制御テーブル関係読み込み(T6)]");
            ctT6NameList_ = new Vector(200);
            ret             = ctT6NameRead();
            if(0 > ret){
                exit(0,"ctT6NameRead()  DATABASE ERROR No[" + ret + "]");
            }

            log("CZSystem initLib","START !! [制御テーブル関係読み込み(T6大項目)]");
            ctT6LagNameList_ = new Vector(200);
            ret             = ctT6LagNameRead();
            if(0 > ret){
                exit(0,"ctT6LagNameRead()  DATABASE ERROR No[" + ret + "]");
            }

            log("CZSystem initLib","START !! [制御テーブル関係読み込み(T6中項目)]");
            ctT6MidNameList_ = new Vector(200);
            ret             = ctT6MidNameRead();
            if(0 > ret){
                exit(0,"ctT6MidNameRead()  DATABASE ERROR No[" + ret + "]");
            }


//@@
			log("CZSystem initLib","START !! [クライアントバージョン取得]");
//			Client_ver_list = new String();
			ret             = ClientVersionGet();
			if(0 > ret){
				exit(0,"ClientVersionGet()  DATABASE ERROR No[" + ret + "]");
			}
        }
        catch(Throwable e){
            log("CZSystem initLib","Error !!");
            handleException(e);
        }
        log("CZSystem initLib","----- END !! -----");
        return true;
    }


/**
* 操業定数マスタを読込む
*/
    public static synchronized void getOperationMst() {

    int ret = 0;

    try {
        log("CZSystem getOperatinMst","START !! [操業定数関係読み込み 大]");
        op_tb_lag_name_list = null;
        op_tb_lag_name_list = new Vector();
        ret                 = opTblLagNameRead();
        if(0 >= ret){
            exit(0,"opTblLagNameRead()  DATABASE ERROR No[" + ret + "]");
        }

        log("CZSystem getOperatinMst","START !! [操業定数関係読み込み 中]");
        op_tb_mid_name_list = null;
        op_tb_mid_name_list = new Vector(20);
        ret                 = opTblMidNameRead();
        if(0 >= ret){
            exit(0,"opTblMidNameRead()  DATABASE ERROR No[" + ret + "]");
        }

        log("CZSystem getOperatinMst","START !! [操業定数関係読み込み 小]");
        op_tb_sml_name_list = null;
        op_tb_sml_name_list = new Vector(500);
        ret                 = opTblSmlNameRead();
        if(0 >= ret){
            exit(0,"opTblSmlNameRead()  DATABASE ERROR No[" + ret + "]");
        }
    } catch (Throwable e) {
        log("CZSystem getOperatinMst ","----- Error !!");
        handleException(e);
    }
    return;
}


/**
* 制御テーブルマスタを読込む
*/
    public static synchronized void getControlMst() {

    int ret = 0;

    try {
            log("CZSystem getControlMst","START !! [制御テーブル関係読み込み]");
            ct_tb_name_list = null;
            ct_tb_name_list = new Vector(200);
            ret             = ctTblNameRead();
            if(0 >= ret){
                exit(0,"ctTblNameRead()  DATABASE ERROR No[" + ret + "]");
            }

            log("CZSystem getControlMst","START !! [制御テーブル関係読み込み(T6)]");
            ctT6NameList_ = null;
            ctT6NameList_ = new Vector(200);
            ret           = ctT6NameRead();
            if(0 > ret){
                exit(0,"ctT6NameRead()  DATABASE ERROR No[" + ret + "]");
            }

            log("CZSystem getControlMst","START !! [制御テーブル関係読み込み(T6大項目)]");
            ctT6LagNameList_ = null;
            ctT6LagNameList_ = new Vector(200);
            ret              = ctT6LagNameRead();
            if(0 > ret){
                exit(0,"ctT6LagNameRead()  DATABASE ERROR No[" + ret + "]");
            }

            log("CZSystem getControlMst","START !! [制御テーブル関係読み込み(T6中項目)]");
            ctT6MidNameList_ = null;
            ctT6MidNameList_ = new Vector(200);
            ret              = ctT6MidNameRead();
            if(0 > ret){
                exit(0,"ctT6MidNameRead()  DATABASE ERROR No[" + ret + "]");
            }
    } catch (Throwable e) {
        log("CZSystem getControlMst ","----- Error !!");
        handleException(e);
    }
    return;
}

    //
    //  初期化 Horb
    //
    private static synchronized boolean initHorb(){

        try {
            log("CZSyatem","initHorb() horb://" + HOST);

            //HORB オブジェクト作成
            log("CZSystem initHorb","-----> START !! [CZNativeGetData_Proxy]");
            cz_gd_px = new CZNativeGetData_Proxy("horb://" + HOST);

            log("CZSystem initHorb","-----> START !! [CZClientEvent_Proxy]");
            cz_ev_px = new CZClientEvent_Proxy(  "horb://" + HOST);

            CZSystemEvent e = new CZSystemEvent(cz_ev_px,cz_gd_px);
            cz_ev_px_th     = new Thread(e);
            cz_ev_px_th.start();

            log("CZSystem initHorb","-----> START !! [CZOperate_Proxy]");
            cz_op_px = new CZOperate_Proxy("horb://" + HOST);

            log("CZSystem initHorb","-----> START !! [CZTableExchange_Proxy]");
            cz_tb_px = new CZTableExchange_Proxy("horb://" + HOST);

            log("CZSystem initHorb","-----> START !! [CZServer_Proxy]");
            cz_sv_px = new CZServer_Proxy("horb://" + HOST);

			/* 2006.07.12 */
            log("CZSystem initHorb","-----> START !! [CZRealNativeData_Proxy]");
            cz_rl_px = new CZRealNativeData_Proxy("horb://" + HOST);
            
        }
        catch(Throwable e){
            log("CZSystem initHorb","***** ERROR !! [" + e + "]");
            handleException(e);
        }

        log("CZSystem initHorb","----- END !! -----");
        return true;
    }


    //
    //  初期化 Horb : 操作応答
    //
    private  static synchronized boolean initHorbClientResult(){

        log("CZSystem initHorbClientResult","CZClientResult_Proxy INIT START !! -----");
        if(null != cz_re_px){
            cz_re_px._release();
            cz_re_px_th = null;
        }

        try {
            cz_re_px = new CZClientResult_Proxy( "horb://" + HOST);
            CZSystemResult r = new CZSystemResult(cz_re_px);
            cz_re_px_th = new Thread(r);
            cz_re_px_th.start();
        }
        catch(Exception e){
            log("CZSystem initHorbClientResult","***** ERROR !! [" + e + "]");
            exit(0,"initHorbClientResult()");
        }
        log("CZSystem initHorbClientResult","CZClientResult_Proxy INIT END !! -----");
        return true;
    }


    //
    //  初期化のチェック
    //
    public static synchronized void initCheck(){
        if(!init_flag){
            log("CZSystem INIT","***** STOP *****");    
            exit(0,"STOP SYSTEM !! ******************");
        }
    }

    //
    //  制終了
    //
    public static synchronized void exit(int i,String comment){
        log("CZSystem *****<< 終了 >>*****", comment);
        System.exit(i);
    }


    //
    //  終了時のフック
    //
    //  注：この中から synchronized メソッドはよばない
    //
    public static void endApp(){
        log("CZSystem *****<< END_APP START >>*****","");
        if(null != cz_tb_px){
            cz_tb_px.CZPutWorkingExclusion(final_ro_no);
            cz_tb_px.CZPutControlExclusion(final_ro_no);
            cz_tb_px._release();
        }
        log("CZSystem endApp","cz_tb_px release OK");

        if(null != cz_ev_px) cz_ev_px._release();   
        log("CZSystem endApp","cz_ev_px release OK");

        if(null != cz_re_px) cz_re_px._release();
        log("CZSystem endApp","cz_re_px release OK");

        if(null != cz_gd_px) cz_gd_px._release();
        log("CZSystem endApp","cz_gd_px release OK");

        if(null != cz_op_px) cz_op_px._release();
        log("CZSystem endApp","cz_op_px release OK");

        if(null != cz_sv_px) cz_sv_px._release();
        log("CZSystem endApp","cz_sv_px release OK");

		/* 2006.07.12 */
        if(null != cz_rl_px) cz_rl_px._release();
        log("CZSystem endApp","cz_rl_px release OK");

        cz_ev_px = null;
        cz_re_px = null;
        cz_gd_px = null;
        cz_op_px = null;
        cz_tb_px = null;
        cz_sv_px = null;
		/* 2006.07.12 */
		cz_rl_px = null;
        init_flag = false;

        log("CZSystem **********<< SYSTEM Shutdown >>**********","");
        log("","");
        log("","");
    }

    //
    //  ランレベル
    //
    public static int getRunLevel(){
        return run_level;
    }


    //
    //  タイマー
    //
    public static boolean sleep(long l){
        try{ Thread.sleep(l); }
        catch(Exception e){ return false ;}
        return true;
    }

    //
    //  エラー
    //
    public static void handleException(Throwable exception) {
        exception.printStackTrace(System.out);
        exit(-1,"handleException()");
    }

    //
    //
    //

    public static void log(String name,String comm){

        java.util.Date system_date      =  new java.util.Date();
        SimpleDateFormat system_date_fm =  new SimpleDateFormat ("MM/dd HH:mm:ss"); 
        String date = system_date_fm.format(system_date);

        try{
            System.out.println(date + " [" + name + "] " + comm );
//            logFileWrite(date + " [" + name + "] " + comm );
        }
        catch(Exception e){
            System.out.println("=========== System.log Exception [" + e + "]");
        }
    }

    //
    //  ファイル出力
    //
/*
    public static void logFileWrite(String s) throws Exception {

    	String FileName = "c:/CZ/log/CZMain_log.txt";

		FileWriter fw = new FileWriter(FileName, true);
        java.util.Date system_date      =  new java.util.Date();
        SimpleDateFormat system_date_fm =  new SimpleDateFormat ("MM/dd HH:mm:ss"); 
        String date = system_date_fm.format(system_date);

    	try{
    		fw.write( s + "\n");
    	
    		fw.close();
    	}
    	catch(IOException e){
			System.out.println("############## System.log Exception [" + e + "]");
		}
	}
*/
    ////////////////////////////////////////////////////////////////////
    //
    //  メッセージ配送
    //
    public static synchronized void sysMessage(CZSystemSysMsg msg){ 
        log("CZSystem sysMessage","No[" + msg.no + "] Massage[" + msg.message + "]");   
        CZEventSender.sendData(msg,CZEventCL.SYS_MESSAGE);
    }

    ////////////////////////////////////////////////////////////////////
    //
    //  ＰＶデータ受信
    //
    public static synchronized void evF001(String  _ro){    


//        log("CZSystem evF001","ＰＶ受信[" + _ro + "]"); 

        if(!_ro.equals(getRoName())){
            return ;
        }

        CZNativeHikiage _bt_set = cz_gd_px.CZNativeHikiageGet(_ro);  //  引き上げ条件
        CZNativeDengen  _dengen = cz_gd_px.CZNativeDengenGet(_ro);   //  電源情報

        CZNativePv p            = cz_gd_px.CZNativePvGet(_ro);
        String  _bt             = p.getBatch();     //  バッチNo
        int _proc               = p.getP_no();      //  プロセスNo
        int _sub_proc           = p.getSp_no();     //  サブプロセスNo
        int _proc_len           = p.getP_renban();  //  プロセス連番
        int _proc_time          = p.getP_time();    //  プロセス時間
        int _sub_proc_time      = p.getSp_time();   //  サブプロセス時間
        int _get_date_time      = p.getP_date();    //  採取日時
        int _main_heat_on_time  = p.getH_ontime();  //  メインヒータ電源オン時間
        int _condition_len      = p.getHk_renban(); //  引上げ条件内連番
        float   _pv[]           = p.getData();      //  データ

//        log("CZSystem evF001","[" + _proc + "][" + _pv[0] + "][" + _condition_len + "]");   

        // 同一採取日時のデータの場合
        if(current_get_date_time == _get_date_time) return;

        // プロセス変更の確認
        int old_proc_len = current_proc_len ;

        setCurrentData( _bt_set , _dengen , _bt, _proc , _sub_proc ,    
                _proc_len , _proc_time , _sub_proc_time ,   
                _get_date_time , _main_heat_on_time , _condition_len , _pv);    

        if(current_proc_len != old_proc_len){
            log("CZSystem evF001","プロセス変更");  
            chgProc(_proc_len,true);
        }

        CZPV.addPVDataUse(current_bt,current_proc,current_sub_proc,current_proc_len,
                  current_proc_time,current_sub_proc_time,current_get_date_time,
                  current_main_heat_on_time,current_condition_len,current_pv);

        CZEventSender.sendData(current_bt,CZEventCL.PV_RECEIVE);
    }


    ////////////////////////////////////////////////////////////////////////////////
    //
    //  炉前手動介入開始通知（４軸）    
    //
    public static synchronized void ev1005(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = -1;
        msg.message = getDateTime() + "  [ 炉前手動介入開始（４軸）]";
        sysMessage(msg);
    }
    //
    //  炉前手動介入終了通知（４軸）    
    //
    public static synchronized void ev8005(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 炉前手動介入終了（４軸）]";
        sysMessage(msg);
    }

    //
    //  炉前手動介入開始通知（４軸以外）    
    //
    public static synchronized void ev100D(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = -1;
        msg.message = getDateTime() + "  [ 炉前手動介入開始（４軸以外）]";
        sysMessage(msg);
    }
    //
    //  炉前手動介入終了通知（４軸以外）    
    //
    public static synchronized void ev800D(CZEvent  ev){    
    
        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 炉前手動介入終了（４軸以外）]";
        sysMessage(msg);
    }

    ////////////////////////////////////////////////////////////////////////////////
    //
    //  手動介入応答（４軸）
    //
    public static synchronized void ev1001(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 手動介入応答（４軸）]";
        sysMessage(msg);
    }

    //
    //  手動介入完了（４軸）
    //
    public static synchronized void ev8001(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 手動介入完了（４軸）]";
        sysMessage(msg);
    }

    //
    //  手動介入ＵＮＤＯ応答（４軸）
    //
    public static synchronized void ev1003(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 手動介入ＵＮＤＯ応答（４軸）]";
        sysMessage(msg);
    }

    //
    //  手動介入ＵＮＤＯ完了（４軸）
    //
    public static synchronized void ev8003(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 手動介入ＵＮＤＯ完了（４軸）]";
        sysMessage(msg);
    }

    //
    //  手動介入応答（４軸以外）
    //
    public static synchronized void ev1009(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 手動介入応答（４軸以外）]";
        sysMessage(msg);
    }

    //
    //  手動介入完了（４軸以外）
    //
    public static synchronized void ev8009(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 手動介入完了（４軸以外）]";
        sysMessage(msg);
    }


    //
    //  手動介入ＵＮＤＯ応答（４軸以外）
    //
    public static synchronized void ev100B(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 手動介入ＵＮＤＯ応答（４軸以外）]";
        sysMessage(msg);
    }


    //
    //  手動介入ＵＮＤＯ完了（４軸以外）
    //
    public static synchronized void ev800B(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 手動介入ＵＮＤＯ完了（４軸以外）]";
        sysMessage(msg);
    }


    ////////////////////////////////////////////////////////////////////////////////
    //
    //  特定プロセス変更応答
    //
    public static synchronized void ev1011(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 特定プロセス変更応答 ]";
        sysMessage(msg);
    }

    //
    //  特定プロセス変更完了通知
    //
    public static synchronized void ev8015(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 特定プロセス変更完了通知 ]";
        sysMessage(msg);
    }


    ////////////////////////////////////////////////////////////////////////////////
    //
    //  プロセス変更応答
    //
    public static synchronized void ev1041(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ プロセス変更応答 ]";
        sysMessage(msg);
    }

    //
    //  プロセス変更完了通知
    //
    public static synchronized void ev8041(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ プロセス変更完了通知 ]";
        sysMessage(msg);
    }

    //
    //  制御モード変更応答
    //
    public static synchronized void ev1051(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 制御モード変更応答 ]";
        sysMessage(msg);
    }

    //
    //  制御モード変更完了通知
    //
    public static synchronized void ev8051(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 制御モード変更完了通知 ]";
        sysMessage(msg);
    }


    ////////////////////////////////////////////////////////////////////////////////
    //
    //  引上げ条件登録応答
    //
    public static synchronized void ev1093(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 引上げ条件登録応答 ]";
        sysMessage(msg);
    }

    //
    //  引上げ条件登録通知
    //
    public static synchronized void ev8091(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 引上げ条件登録通知 ]";
        sysMessage(msg);
    }

    ////////////////////////////////////////////////////////////////////////////////
    //
    //  CCDカメラモニタ切替
    //
    public static synchronized void ev1261(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ CCDカメラモニタ切替 ]";
        sysMessage(msg);
    }

    ////////////////////////////////////////////////////////////////////////////////
    //
    //  取出しテーブル設定要求
    //
    public static synchronized void ev1098(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 取出しテーブル設定要求 ]";
        sysMessage(msg);
    }

    //
    //  取出しテーブル設定応答
    //
    public static synchronized void ev1099(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 取出しテーブル設定応答 ]";
        sysMessage(msg);
    }

    //
    //  取出しテーブル登録通知
    //
    public static synchronized void ev8099(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 取出しテーブル登録通知 ]";
        sysMessage(msg);
    }

    ////////////////////////////////////////////////////////////////////////////////
    //
    //  制御テーブル送信開始
    //
    public static synchronized void ev1200(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 制御テーブル送信開始 ]";
        sysMessage(msg);
    }

    //
    //  制御テーブル要求
    //
    public static synchronized void ev1201(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 制御テーブル要求 ]";
        sysMessage(msg);
    }

    //
    //  制御テーブル通知（初期時）
    //
    public static synchronized void ev1202(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 制御テーブル通知（初期時） ]";
        sysMessage(msg);
    }

    //
    //  制御テーブル送信終了通知
    //
    public static synchronized void ev1204(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 制御テーブル送信終了通知 ]";
        sysMessage(msg);
    }

    //
    //  制御テーブル未登録通知
    //
    public static synchronized void ev1206(CZEvent  ev){    
        String ro = ev.getRoban();

        if(!ro.equals(getRoName())) return ;

        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = -1;
        msg.message = getDateTime() + "  [ 制御テーブル未登録通知 ]";
        sysMessage(msg);
    }


    ////////////////////////////////////////////////////////////////////////////////
    //
    //  制御テーブル更新応答
    //
    public static synchronized void ev1063(CZEvent  ev){    
        String ro = ev.getRoban();

        if(!ro.equals(getRoName())) return ;

        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 制御テーブル更新応答 ]";
        sysMessage(msg);
    }

    //
    //  操業定数更新応答
    //
    public static synchronized void ev1083(CZEvent  ev){    
        String ro = ev.getRoban();

        if(!ro.equals(getRoName())) return ;

        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 操業定数更新応答 ]";
        sysMessage(msg);
    }

    //
    //  制御テーブルグループ名変更応答
    //
    public static synchronized void ev1237(CZEvent  ev){    
        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;

        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 制御テーブルグループ名変更応答 ]";
        sysMessage(msg);
    }

    //
    //  制御テーブルタイトル変更応答
    //
    public static synchronized void ev1239(CZEvent  ev){    

        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 制御テーブルタイトル変更応答 ]";
        sysMessage(msg);
    }

    //
    //  制御テーブル定義更新応答
    //
    public static synchronized void ev1241(CZEvent  ev){    
        String ro = ev.getRoban();

        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 制御テーブル定義更新応答 ]";
        sysMessage(msg);
    }

    //
    //  操業定数項目名変更応答
    //
    public static synchronized void ev1247(CZEvent  ev){    
        String ro = ev.getRoban();

        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 操業定数項目名変更応答 ]";
        sysMessage(msg);
    }

    ////////////////////////////////////////////////////////////////////////////////
    //
    //  生波形データ採取応答
    //
    public static synchronized void ev1021(CZEvent  ev){    
        String ro = ev.getRoban();

        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 生波形データ採取応答 ]";
        sysMessage(msg);
        CZEventSender.sendData(ev,CZEventCL.EV_1021);
    }

    //
    //  生波形データ採取通知
    //
    public static synchronized void ev8021(CZEvent  ev){    
        String ro = ev.getRoban();

        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ 生波形データ採取通知 ]";
        sysMessage(msg);

        CZEventSender.sendData(ev,CZEventCL.EV_8021);
    }


    //
    //  ＣＣＤカメラ画像保存応答
    //
    public static synchronized void ev1023(CZEvent  ev){    
        String ro = ev.getRoban();
        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ ＣＣＤカメラ画像保存応答 ]";
        sysMessage(msg);

        CZEventSender.sendData(ev,CZEventCL.EV_1023);
    }

    //
    //  ＣＣＤカメラ画像保存完了
    //
    public static synchronized void ev8023(CZEvent  ev){    
        String ro = ev.getRoban();

        if(!ro.equals(getRoName())) return ;
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = getDateTime() + "  [ ＣＣＤカメラ画像保存完了 ]";
        sysMessage(msg);
        CZEventSender.sendData(ev,CZEventCL.EV_8023);
    }

    ////////////////////////////////////////////////////////////////////////////////
    //
    //  異常項目通知
    //
    public static synchronized void evF007(CZEvent  ev){    
        CZEventSender.sendData(ev,CZEventCL.EV_F007);
    }

    //
    //  炉体状況通知
    //
    public static synchronized void evF009(CZEvent  ev){    
        CZEventSender.sendData(ev,CZEventCL.EV_F009);
    }

    ////////////////////////////////////////////////////////////////////////////////

    //
    //  操業定数要求の応答
    //
    public static synchronized void ev1217(CZResult  ev){   
        CZEventSender.sendData(ev,CZEventCL.OT_GET_HAITA);
    }

    //
    //  操業定数排他開放の応答
    //
    public static synchronized void ev1219(CZResult  ev){   
        CZEventSender.sendData(ev,CZEventCL.OT_PUT_HAITA);
    }


    //
    //  制御テーブル排他開放の応答
    //
    public static synchronized void ev1221(CZResult  ev){   
        CZEventSender.sendData(ev,CZEventCL.CT_GET_HAITA);
    }

    //
    //  制御テーブル排他開放の応答
    //
    public static synchronized void ev1223(CZResult  ev){   
        CZEventSender.sendData(ev,CZEventCL.CT_PUT_HAITA);
    }



    //
    //  制御電源要求応答
    //
    public static synchronized void ev1031(CZResult  ev){   
        CZEventSender.sendData(ev,CZEventCL.EV_1031);
    }
    //
    //  電源変更完了通知
    //
    public static synchronized void ev8031(CZResult  ev){   
        CZEventSender.sendData(ev,CZEventCL.EV_8031);
    }
    

    ////////////////////////////////////////////////////////////////////

    //
    //  カレントデータの設定
    //
    private static synchronized void setCurrentData(CZNativeHikiage _bt_set,
                            CZNativeDengen  _dengen,
                            String  _bt,    
                            int _proc,
                            int _sub_proc,  
                            int _proc_len,  
                            int _proc_time,
                            int _sub_proc_time,
                            int _get_date_time, 
                            int _main_heat_on_time, 
                            int _condition_len,
                            float   _pv[]){

        current_bt_set              = _bt_set;              // 引き上げ条件の設定
        if(current_proc != _proc){
            current_unten           = getUnten(_proc);      //  運転画面表示項目
        }
        current_power_stat          = _dengen;              // 電源状態の設定
        current_bt                  = _bt;                  //  バッチNo
        current_proc                = _proc;                //  プロセスNo
        current_sub_proc            = _sub_proc;            //  サブプロセスNo
        current_proc_len            = _proc_len;            //  プロセス連番
        current_proc_time           = _proc_time;           //  プロセス時間
        current_sub_proc_time       = _sub_proc_time;       //  サブプロセス時間
        current_get_date_time       = _get_date_time;       //  採取日時
        current_main_heat_on_time   = _main_heat_on_time;   //  メインヒータ電源オン時間
        current_condition_len       = _condition_len;       //  引上げ条件内連番
        current_pv                  = _pv;                  //  データ

    }

    ////////////////////////////////////////////////////////////////////
    //
    //  BtNo取り出し
    //
    public static synchronized String getBatch(){
        return current_bt;
    }

    //
    //  プロセスNo取り出し
    //
    public static synchronized int getProcNo(){
        return current_proc;
    }

    //
    //  プロセス時間
    //
    public static synchronized int getProcTime(){
        return current_proc_time;
    }

    //
    //  プロセスモード
    //
    public static synchronized int getProcMode(){
        int proc = (int)current_pv[2];
        return proc;
    }

    //
    //  引き上げ条件取り出し
    //
    public static synchronized CZNativeHikiage getBtSet(){
        return current_bt_set;
    }

    //
    //  運転画面設定取り出し
    //
    public static synchronized CZSystemPVNamePMM getUnten(int proc){
        CZSystemPVNamePMM ret = untenRead(proc);
/*@@@@
        if(null == ret){
            exit(0,"Error getUnten() 運転情報が取得できません!! *****");
        }
@@@@*/
        return ret;
    }

    //
    //  電源状態取り出し
    //
    public static synchronized CZNativeDengen getPowerStat(){
        String ro = getRoName();
        CZNativeDengen  _dengen    = cz_gd_px.CZNativeDengenGet(ro);    //  電源情報
        if(null != _dengen) current_power_stat = _dengen;
        return current_power_stat;
    }

    //
    //  炉名称読み取り
    //
    public static synchronized Vector getRoNameList(){
        initCheck();
        if (ro_name_list.isEmpty()) return null;
        return ro_name_list;
    }




    //
    //  カレントの炉カメラ番号取り出し
    //
    public static synchronized int getRoCameraNo(){
        initCheck();
        String ret = (String)ro_camera_list.elementAt(ro_no_idx);
        if(null == ret) return -1;
        int no = 0;
        try{
            no = Integer.parseInt(ret);
        }
        catch(Exception e){
            return -1;
        }   
        if(1 > no) return -1;
        return no;
    }

    //
    //
    //
    public static synchronized String getRoName(){
        initCheck();
        String ret = (String)ro_name_list.elementAt(ro_no_idx);
        return ret;
    }

    //
    //  炉名取り出し
    //
    public static synchronized String getRoName(int idx){
        initCheck();
        String ret = null;
        try{
            ret = (String)ro_name_list.elementAt(idx);
        }
        catch( Exception e){
            exit(-1,"CZSystem getRoName 1 Error !! [" + idx + "]");
        }
        if(null == ret){
            exit(-1,"CZSystem getRoName 2 Error !! [" + idx + "]");
        }
        return ret;
    }

    //
    //  炉ＤＢ名取り出し
    public static synchronized String getDBName(){

        initCheck();
        String tmp = getRoName();
        String ret = new String(tmp.toLowerCase() );
        return ret;
    }


    //
    //  炉ＤＢ名取り出し
    //
    public static synchronized String getDBName(int idx){

        initCheck();
        String tmp = getRoName(idx);
        String ret = new String(tmp.toLowerCase() );
        return ret;
    }

    // *****************************************************
    //  カレントのＰＶ表名取り出し
    // @return ＰＶ表名
    public static synchronized String getViewName(){

        initCheck();
        String      ret     = null;
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;
        int i = 0;

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getViewName","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return null;
        }

        try{
            // Get User, Password
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getViewName","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        // Set Table Name
        String view = new String("r_pv_control");
        sql = new String("SELECT * FROM " + getDBName() + "." + view + " WHERE " +
                     "batch = '" + getBatch() + "' ORDER BY s_start DESC"); 
        log("CZSystem getViewName","SQL["+sql+"]");

        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getViewName","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            rs.next();
            ret= rs.getString(2);       //t_name
            i = 1;  
            log("CZSystem getViewName","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getViewName","ERROR: Select failed");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getViewName","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }

    //  ＰＶ表名取り出し
    // @param  db_name .. ＤＢ名, bt_name .. バッチ名
    // @return ＰＶ表名
    public static String getViewName(String db_name,String bt_name){

        initCheck();
        String      ret     = null;
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        String view = new String("r_pv_control");

        sql = new String("SELECT * FROM " + db_name + "." + view + " WHERE " +
             "batch = '" + bt_name + "' ORDER BY s_start DESC");    
        log( "CZSystem getViewName","SQL["+ sql +"]" );

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getViewName","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getViewName","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getViewName","ERROR: createStatement or database");
        return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            rs.next();
            ret= rs.getString(2); //t_name

            i = 1;  
            log("CZSystem getViewName","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getViewName","ERROR: Select failed.");
        }
        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getViewName","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }


    // *****************************************************
    //  ＰＶ定義取り出し
    // @return ＰＶ定義名
    public static CZSystemPVName getPVName(int no){

        initCheck();
        if(pv_name_list.size() <= no) return null;
        CZSystemPVName ret = (CZSystemPVName)pv_name_list.elementAt(no);
        return ret;
    }

    // *****************************************************
    //  ＰＶ定義取り出し @@
    // @return ＰＶ定義名
    public static Vector getPVNameAll(){

        initCheck();
        return pv_name_list;
    }

    // *****************************************************
    //  エラーメッセージ取り出し
    //
    public static CZSystemErrMsg getErrMsg(int no){

        initCheck();
        int size = error_message_list.size();
        for(int i=0 ; i < size ; i++){
            CZSystemErrMsg ret = (CZSystemErrMsg)error_message_list.elementAt(i);
            if(ret.e_no == no) return ret;
        }
        CZSystemErrMsg err = new CZSystemErrMsg();
        err.e_no    = no;
        err.message = new String("未定義メッセージ");
        err.youhi   = 0;
        return err;
    }

	/*2006.04.13 y.k */
    public static CZSystemErrMsg getErrMsg2(int no, Vector err_list_msg) {

        initCheck();

        if(null != err_list_msg)
		{
	        int size = err_list_msg.size();
	        for(int i=0 ; i < size ; i++){
	            CZSystemErrMsg ret = (CZSystemErrMsg)err_list_msg.elementAt(i);
	            if(ret.e_no == no) return ret;
	        }
		}

        CZSystemErrMsg err = new CZSystemErrMsg();
        err.e_no    = no;
        err.message = new String("未定義メッセージ");
        err.youhi   = 0;
        return err;
    }


    //
    //  操業定数：大項目取り出し
    //
    public static CZSystemOpTbLag getOpTbLag(int no){

        initCheck();
        log("CZSystem getOpTbLag","size[" + op_tb_lag_name_list.size() + "]");
        if(op_tb_lag_name_list.size() <= no) return null;
        CZSystemOpTbLag ret = (CZSystemOpTbLag)op_tb_lag_name_list.elementAt(no);
        return ret;
    }

    //
    //  操業定数：中項目取り出し
    //
    public static CZSystemOpTbMid getOpTbMid(int l , int m){

        initCheck();
        int lag = l+1;
        int mid = m+1;

        for(int i = 0 ; i < op_tb_mid_name_list.size() ; i++){
            CZSystemOpTbMid ret = (CZSystemOpTbMid)op_tb_mid_name_list.elementAt(i);
            if((ret.k_no1 == lag) && (ret.k_no2 == mid)) return ret;
        }
        return null;
    }

    //
    //  操業定数：項目取り出し
    //
    public static CZSystemOpTbSml getOpTbSml(int l , int m , int s){

        initCheck();
        int lag = l;
        int mid = m;
        int sml = s;

        for(int i = 0 ; i < op_tb_sml_name_list.size() ; i++){
            CZSystemOpTbSml ret = (CZSystemOpTbSml)op_tb_sml_name_list.elementAt(i);
            if((ret.k_no1 == lag) &&    
               (ret.k_no2 == mid) &&
               (ret.k_no  == sml)) return ret;
        }
        return null;
    }

    //
    //  制御テーブル：項目取り出し
    //
    public static CZSystemCtName getCtTbName(int g , int t){

        initCheck();
        int grp = g;
        int tbl = t;

        for(int i = 0 ; i < ct_tb_name_list.size() ; i++){
            CZSystemCtName ret = (CZSystemCtName)ct_tb_name_list.elementAt(i);
            if((ret.g_no == grp) && 
               (ret.t_no  == tbl)) return ret;
        }
        return null;
    }

    //
    //  制御テーブル(T6)：項目取り出し
    //
    public static CZSystemCtT6Name getCtT6Name(int g , int l, int m, int n){
        initCheck();

        int grp  = g;
        int lag  = l;
        int mid  = m;
        int kNo  = n;

        if (null == ctT6NameList_) return null;
        for(int i = 0 ; i < ctT6NameList_.size() ; i++){
            CZSystemCtT6Name ret = (CZSystemCtT6Name)ctT6NameList_.elementAt(i);
            if((ret.g_no  == grp) &&    
               (ret.k_no1 == lag) &&    
               (ret.k_no2 == mid) &&    
               (ret.k_no  == kNo)) return ret;
        }
        return null;
    }

    //
    //  制御テーブル(T6)：大項目取り出し
    //
    public static CZSystemCtT6LagName getCtT6LagName(int g , int r,int l){
        initCheck();

        int grp  = g;
        int rcp  = r;
        int lag  = l;

        if (null == ctT6LagNameList_) return null;
        for(int i = 0 ; i < ctT6LagNameList_.size() ; i++){
            CZSystemCtT6LagName ret = (CZSystemCtT6LagName)ctT6LagNameList_.elementAt(i);
            if((ret.g_no  == grp) &&    
               (ret.k_no1 == lag) ) return ret;
/*@@@
            if((ret.g_no  == grp) &&    
               (ret.r_no  == rcp) &&    
               (ret.k_no1 == lag) ) return ret;
@@@*/
        }
        return null;
    }

    //
    //  制御テーブル(T6)：中項目取り出し
    //
    public static CZSystemCtT6MidName getCtT6MidName(int g , int r, int l, int m){
        initCheck();

        int grp  = g;
        int rcp  = r;
        int lag  = l;
        int mid  = m;

        if (null == ctT6MidNameList_) return null;
        for(int i = 0 ; i < ctT6MidNameList_.size() ; i++){
            CZSystemCtT6MidName ret = (CZSystemCtT6MidName)ctT6MidNameList_.elementAt(i);
            if((ret.g_no  == grp) &&    
               (ret.k_no1 == lag) &&    
               (ret.k_no2 == mid)) return ret;
/*@@@
            if((ret.g_no  == grp) &&    
               (ret.r_no  == rcp) &&    
               (ret.k_no1 == lag) &&    
               (ret.k_no2 == mid)) return ret;
@@@*/
        }
        return null;
    }

    //
    //  操業定数：定数取り出し
    //
	@SuppressWarnings("unchecked")
    public static synchronized Vector getOpTb(int l , int m){

        initCheck();
        Vector      ret     = new Vector(200);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getOpTb","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getOpTb","ERROR: failed to connect!");
            handleException(e);
            return null;
        }
            
        try{
            sqlstmt = conn.createStatement() ;
            sql = new String("SELECT * FROM " + getDBName() + "." + "r_st_mast WHERE " +
                     "k_no1 = " + l + " AND " +
                     "k_no2 = " + m + " ORDER BY k_no1,k_no2,k_no");    
            log("CZSystem getOpTb","SQL["+sql+"]");

        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getOpTb","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZSystemOpTb name = new CZSystemOpTb();
                name.k_no1  = rs.getInt(1);
                name.k_no2  = rs.getInt(2);
                name.k_no   = rs.getInt(3);
                name.k_val  = rs.getFloat(4);
                ret.addElement(name);
            } // for end

            log("CZSystem getOpTb","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getOpTb","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getOpTb","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }


    //
    //  操業定数：定数取り出し  (ＤＢ指定)
    //
	@SuppressWarnings("unchecked")
    public static synchronized Vector getOpTb(String db,int l , int m){

        initCheck();
        Vector      ret     = new Vector(200);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        sql = new String("SELECT * FROM " + db + "." + "r_st_mast WHERE " +
                 "k_no1 = " + l + " AND " +
                 "k_no2 = " + m + " ORDER BY k_no1,k_no2,k_no");    
        log("CZSystem getOpTb","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getOpTb","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getOpTb","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getOpTb","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZSystemOpTb name = new CZSystemOpTb();
                name.k_no1  = rs.getInt(1);
                name.k_no2  = rs.getInt(2);
                name.k_no   = rs.getInt(3);
                name.k_val  = rs.getFloat(4);
                ret.addElement(name);
            } // for end
            log("CZSystem getOpTb","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getOpTb","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getOpTb","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }

    //
    //  操業定数：定数取り出し  (ＤＢ指定)
    // 2006.06.06 y.k
	@SuppressWarnings("unchecked")
    public static synchronized Vector getSogyoAllTb(String db){

        initCheck();
        Vector      ret     = new Vector(200);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        sql = new String("SELECT * FROM " + db + "." + "r_st_mast ORDER BY k_no1,k_no2,k_no");    
        log("CZSystem getSogyoAllTb","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getSogyoAllTb","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getSogyoAllTb","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getSogyoAllTb","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZSystemOpTb name = new CZSystemOpTb();
                name.k_no1  = rs.getInt(1);
                name.k_no2  = rs.getInt(2);
                name.k_no   = rs.getInt(3);
                name.k_val  = rs.getFloat(4);
                ret.addElement(name);
            } // for end
            log("CZSystem getSogyoAllTb","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getSogyoAllTb","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getSogyoAllTb","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }

    //
    //  制御テーブル：タイトル取り出し
    //
    public static synchronized Vector getCtTitle(){
        return getCtTitleSub(getDBName());
    }

    //
    //  制御テーブル：タイトル取り出し
    //
    public static synchronized Vector getCtTitle(int idx){
        return getCtTitleSub(getDBName(idx));
    }


    //
    //  制御テーブル：タイトル取り出し
    //
	@SuppressWarnings("unchecked")
    public static synchronized Vector getCtTitleSub(String db_name){

        initCheck();
        Vector      ret     = new Vector(100);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        sql = new String("SELECT * FROM " + db_name + "." + "r_ct_title ORDER BY g_no,r_no");   
        log("CZSystem getCtTitleSub","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getCtTitleSub","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getCtTitleSub","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getCtTitleSub","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZSystemCtTitle title = new CZSystemCtTitle();
                title.g_no  = rs.getInt(1);
                title.r_no  = rs.getInt(2);
                title.title = rs.getString(3);
                ret.addElement(title);
            } // for end
            log("CZSystem getCtTitleSub","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getCtTitleSub","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getCtTitleSub","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }


    //
    //  制御テーブル：値取り出し
    //
	@SuppressWarnings("unchecked")
    public static synchronized Vector getCtTb(int g_no,int r_no,int t_no,boolean current){

        initCheck();
        Vector      ret     = new Vector(100,100);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getCtTb","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getCtTb","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
            String view = null;

            if(current){
				view = new String("r_ct_current");
			}
            else{
				view = new String("r_ct_mast");
			}

            sql = new String("SELECT * FROM " + getDBName() + "." + view + " WHERE " +
                     "g_no = " + g_no + " AND " +
                     "r_no = " + r_no + " AND " +
                     "t_no = " + t_no + " ORDER BY k_no");  

            log("CZSystem getCtTb","SQL["+sql+"]");
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getCtTb","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;

            for(i = 0 ; rs.next() ; i++){
                CZSystemCtTb name = new CZSystemCtTb();
                name.g_no   = rs.getInt(1);
                name.r_no   = rs.getInt(2);
                name.t_no   = rs.getInt(3);
                name.k_no   = rs.getInt(4);
                name.l_val  = rs.getFloat(5);
                name.r_val  = rs.getFloat(6);
                ret.addElement(name);
            } // for end
            log("CZSystem getCtTb","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getCtTb","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getCtTb","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }


    //
    //  制御テーブル(T6)：値取り出し
    //
	@SuppressWarnings("unchecked")
    public static synchronized Vector getCtT6Tb(int g_no,int r_no,int l_no,int m_no,boolean current){

        initCheck();
        Vector      ret     = new Vector();
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getCtTb","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getCtTb","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
            String view = null;

            if(current) view = new String("r_ct6_current");
            else        view = new String("r_ct6_mast");

            sql = new String("SELECT * FROM " + getDBName() + "." + view + " WHERE " +
                     "g_no = "  + g_no + " AND " +
                     "r_no = "  + r_no + " AND " +
                     "k_no1 = " + l_no + " AND " +
                     "k_no2 = " + m_no + " ORDER BY g_no,r_no,k_no1,k_no2,k_no");  

            log("CZSystem getCtTb","SQL["+sql+"]");
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getCtT6Tb","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;

            for(i = 0 ; rs.next() ; i++){
                CZSystemCtT6Tb name = new CZSystemCtT6Tb();
                name.g_no   = rs.getInt(1);
                name.r_no   = rs.getInt(2);
                name.k_no1  = rs.getInt(3);
                name.k_no2  = rs.getInt(4);
                name.k_no   = rs.getInt(5);
                name.k_val  = rs.getFloat(6);
                ret.addElement(name);
            } // for end
            log("CZSystem getCtT6Tb","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getCtT6Tb","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();                    //@@
        }
        catch (SQLException e){
            log("CZSystem getCtT6Tb","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }

    //
    //  制御テーブル：値取り出し (ＤＢ指定)
    //  2006.06.15  Y.K
	@SuppressWarnings("unchecked")
    public static synchronized Vector getCtAllTb(String db,int g_no,int r_no){

        initCheck();
        Vector      ret     = new Vector(100,100);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;
        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getCtTb","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getCtTb","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
            sql = new String("SELECT * FROM " + db + ".r_ct_mast WHERE " +
                     "g_no = " + g_no + " AND " +
                     "r_no = " + r_no + " " +
                     " ORDER BY t_no, k_no");  
            log("CZSystem getCtTb","SQL["+sql+"]");
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getCtTb","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZSystemCtTb name = new CZSystemCtTb();
                name.g_no   = rs.getInt(1);
                name.r_no   = rs.getInt(2);
                name.t_no   = rs.getInt(3);
                name.k_no   = rs.getInt(4);
                name.l_val  = rs.getFloat(5);
                name.r_val  = rs.getFloat(6);
                ret.addElement(name);
            } // for end
            log("CZSystem getCtTb","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getCtTb","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getCtTb","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }

    //
    //  制御テーブル(T6)：値取り出し
    //  2006.06.13  Y.K
	@SuppressWarnings("unchecked")
    public static synchronized Vector getCtT6AllTb(String sDBName,int r_no){

        initCheck();
        Vector      ret     = new Vector();
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getCtT6AllTb","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getCtT6AllTb","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
            sql = new String("SELECT * FROM " + sDBName + ".r_ct6_mast WHERE " +
                     "g_no = 6 AND " +
                     "r_no = "  + r_no +
					 " ORDER BY g_no,r_no,k_no1,k_no2,k_no");  

            log("CZSystem getCtT6AllTb","SQL["+sql+"]");
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getCtT6AllTb","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;

            for(i = 0 ; rs.next() ; i++){
                CZSystemCtT6Tb name = new CZSystemCtT6Tb();
                name.g_no   = rs.getInt(1);
                name.r_no   = rs.getInt(2);
                name.k_no1  = rs.getInt(3);
                name.k_no2  = rs.getInt(4);
                name.k_no   = rs.getInt(5);
                name.k_val  = rs.getFloat(6);
                ret.addElement(name);
            } // for end
            log("CZSystem getCtT6AllTb","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getCtT6AllTb","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();                    //@@
        }
        catch (SQLException e){
            log("CZSystem getCtT6AllTb","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }

    //
    //  制御テーブル：該当GrからレシピNo取得
    //	2006.06.08 y.k
	@SuppressWarnings("unchecked")
    public static synchronized Vector getCtTbRcp(String sRo, int g_no){

        initCheck();
        Vector      ret     = new Vector(100,100);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getCtTbRcp","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getCtTbRcp","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
            String view = null;

            if(g_no == 6){
				view = new String("r_ct6_mast");
			}
            else{
				view = new String("r_ct_mast");
			}

            sql = new String("select DISTINCT ct.g_no,ct.r_no,nvl(tl.title,' ') " +
                             "from " + sRo + "." + view + " ct," + sRo + ".r_ct_title tl " +
							 "where ct.g_no = " + g_no + " and ct.g_no = tl.g_no(+) and ct.r_no = tl.r_no(+) " +
							 "order by ct.g_no,ct.r_no" );

            log("CZSystem getCtTbRcp","SQL["+sql+"]");
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getCtTbRcp","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;

            for(i = 0 ; rs.next() ; i++){
                CZSystemCtTitle name = new CZSystemCtTitle();
                name.g_no   = rs.getInt(1);
                name.r_no   = rs.getInt(2);
                name.title   = rs.getString(3);
				ret.addElement(name);
            } // for end
            log("CZSystem getCtTb","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getCtTbRcp","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getCtTbRcp","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }

    //
    //  制御テーブル：値取り出し (ＤＢ指定)
    //
	@SuppressWarnings("unchecked")
    public static synchronized Vector getCtTb(String db,int g_no,int r_no,int t_no,boolean current){

        initCheck();
        Vector      ret     = new Vector(100,100);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;
        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getCtTb","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getCtTb","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
            String view = null;
            if(current) view = new String("r_ct_current");
            else        view = new String("r_ct_mast");

            sql = new String("SELECT * FROM " + db + "." + view + " WHERE " +
                     "g_no = " + g_no + " AND " +
                     "r_no = " + r_no + " AND " +
                     "t_no = " + t_no + " ORDER BY k_no");  
            log("CZSystem getCtTb","SQL["+sql+"]");
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getCtTb","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZSystemCtTb name = new CZSystemCtTb();
                name.g_no   = rs.getInt(1);
                name.r_no   = rs.getInt(2);
                name.t_no   = rs.getInt(3);
                name.k_no   = rs.getInt(4);
                name.l_val  = rs.getFloat(5);
                name.r_val  = rs.getFloat(6);
                ret.addElement(name);
            } // for end
            log("CZSystem getCtTb","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getCtTb","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getCtTb","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }

    //
    //  制御テーブル：大項目取り出し
    //
	@SuppressWarnings("unchecked")
    public static synchronized Vector getCtT6Lag(int gNo , int rNo){

        initCheck();
        Vector      ret     = new Vector();
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;
        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getCtT6LagName","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getCtT6LagName","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;

            sql = new String("SELECT * FROM mst.m_ct6_name1 WHERE " +
                     "g_no = " + gNo + " ORDER BY g_no,r_no,k_no1");    
//@@@            sql = new String("SELECT * FROM mst.m_ct6_name1 WHERE " +
//@@@                     "g_no = " + gNo + " AND " +
//@@@                    "r_no = " + rNo + " ORDER BY g_no,r_no,k_no1");    
            log("CZSystem getCtT6LagName","SQL["+sql+"]");
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getCtT6LagName","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZSystemCtT6LagName name = new CZSystemCtT6LagName();
                name.g_no    = rs.getInt(1);
                name.r_no    = rs.getInt(2);
                name.k_no1   = rs.getInt(3);
                name.k_name1 = rs.getString(4);
                ret.addElement(name);
            } // for end
            log("CZSystem getCtT6LagName","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getCtT6LagName","ERROR: Select failed.");
        }

        try{
            if (null != rs) rs.close();             //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getCtT6LagName","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }


    //
    //  制御テーブル：中項目取り出し
    //
	@SuppressWarnings("unchecked")
    public static synchronized Vector getCtT6Mid(int gNo, int rNo, int lNo){

        initCheck();
        Vector      ret     = new Vector();
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;
        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getCtT6MidName","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getCtT6MidName","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;

            sql = new String("SELECT * FROM mst.m_ct6_name2 WHERE " +
                     "g_no = " + gNo + " AND " +
                     "k_no1 = " + lNo + " ORDER BY g_no,r_no,k_no1,k_no2"); 
//@@@            sql = new String("SELECT * FROM mst.m_ct6_name2 WHERE " +
//@@@                     "g_no = " + gNo + " AND " +
//@@@                     "r_no = " + rNo + " AND " +
//@@@                     "k_no1 = " + lNo + " ORDER BY g_no,r_no,k_no1,k_no2"); 
            log("CZSystem getCtT6MidName","SQL["+sql+"]");
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getCtT6LagName","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZSystemCtT6MidName name = new CZSystemCtT6MidName();
                name.g_no    = rs.getInt(1);
                name.r_no    = rs.getInt(2);
                name.k_no1   = rs.getInt(3);
                name.k_no2   = rs.getInt(4);
                name.k_name2 = rs.getString(5);
                ret.addElement(name);
            } // for end
            log("CZSystem getCtT6MidName","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getCtT6MidName","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getCtT6MidName","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }


    //
    //  制御テーブル：Ｔ６項目取り出し
    //
	@SuppressWarnings("unchecked")
    public static synchronized Vector getCtT6Mst(int gNo, int lNo, int mNo){

        initCheck();
        Vector      ret     = new Vector();
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;
        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getCtT6Mst","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getCtT6Mst","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;

            sql = new String("SELECT * FROM mst.m_ct6_mast WHERE " +
                     "g_no = " + gNo + " AND " +
                     "k_no1 = " + lNo + " AND " +
                     "k_no2 = " + mNo + " ORDER BY g_no,k_no1,k_no2,k_no"); 
            log("CZSystem getCtT6Mst","SQL["+sql+"]");
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getCtT6Mst","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZSystemCtT6Name name = new CZSystemCtT6Name();
                name.g_no   = rs.getInt(1);
                name.k_no1  = rs.getInt(2);
                name.k_no2  = rs.getInt(3);
                name.k_no   = rs.getInt(4);
                name.k_name = rs.getString(5);
                name.k_unit = rs.getString(6);
                name.k_min  = rs.getFloat(7);
                name.k_max  = rs.getFloat(8);
                name.k_keta = rs.getInt(9);
                name.k_sort = rs.getInt(10);
                name.pv_no  = rs.getInt(11);
                ret.addElement(name);
            } // for end
            log("CZSystem getCtT6Mst","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getCtT6Mst","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getCtT6Mst","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }

    //
    //  引き上げ条件取り出し
    //
	@SuppressWarnings("unchecked")
    public static Vector getBtCondition(String db_name){

        initCheck();
        Vector  ret         = new Vector(1000);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        String view = new String("r_hikiage");
        sql = new String("SELECT * FROM " + db_name + "." + view + " ORDER BY t_time DESC");

        log("CZSystem getBtCondition","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getBtCondition","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getBtCondition","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getBtCondition","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZSystemBt bt   = new CZSystemBt();
                bt.batch        = rs.getString(1);  //バッチ番号
                bt.t_time       = rs.getString(2);  //PG-ID
                bt.renban       = rs.getInt(3);     //登録日時
                bt.pgid         = rs.getString(4);  //連番
                bt.hinshu       = rs.getString(5);  //品種
                bt.houi         = rs.getString(6);  //方位
                bt.h_type       = rs.getString(7);  //タイプ
                bt.hiteikou     = rs.getString(8);  //比抵抗
                bt.sanso        = rs.getString(9);  //酸素
                bt.gap          = rs.getString(10); //GAP
                bt.rutubo_kei   = rs.getInt(11);    //ルツボ径
                bt.chokkei      = rs.getInt(12);    //直径
                bt.hikiage_cho  = rs.getInt(13);    //引上長
                bt.top_ar       = rs.getInt(14);    //トップアルゴン
                bt.pull_ar      = rs.getInt(15);    //プルアルゴン
                bt.i_sikomi     = rs.getInt(16);    //仕込量
                bt.t_sikomi     = rs.getInt(17);    //追加仕込量
                bt.zaneki       = rs.getInt(18);    //残液量
                bt.no_youkai    = rs.getInt(19);    //T1(溶解)
                bt.no_hikiage   = rs.getInt(20);    //T2(引上)
                bt.no_kaiten    = rs.getInt(21);    //T3(回転)
                bt.no_toridasi  = rs.getInt(22);    //T4(取出)
                bt.no_aturyoku  = rs.getInt(23);    //T5(圧力)
                bt.no_teisu     = rs.getInt(24);    //T6(定数) @@
                bt.pno_start    = rs.getInt(25);    //スタートプロセス
                bt.p_kaisi      = rs.getInt(26);    //開始
                ret.addElement(bt);
            } // for end
            log("CZSystem getBtCondition","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getBtCondition","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getBtCondition","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }


    //
    //  炉毎引き上げバッチ一覧取り出し（上位２０項目分まで取り出し）
    //  （複数PV実績データ保存用）
	@SuppressWarnings("unchecked")
    public static Vector getPVDataBtList(String db_name){

        initCheck();
        Vector  ret         = new Vector(1000);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        String view = new String("r_hikiage");
        sql = new String("SELECT batch, hinshu, i_sikomi, no_hikiage, max(t_time) FROM " + db_name + "." + view + " WHERE p_kaisi = 1 GROUP BY batch, hinshu, i_sikomi, no_hikiage ORDER BY max(t_time) DESC");

        log("CZSystem getPVDataBtList","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getPVDataBtList","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getPVDataBtList","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getPVDataBtList","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql);
            for(i = 0 ; rs.next() ; i++){
                int row = rs.getRow();
                if(row > 20){
                    break;
                }else{
                    CZPVDataBtList BtList   = new CZPVDataBtList();
                    BtList.flg          = 0;
                    BtList.batch        = rs.getString(1);  //バッチ番号
                    BtList.hinshu       = rs.getString(2);  //品種
                    BtList.i_sikomi     = rs.getInt(3);    //仕込量
                    BtList.no_hikiage   = rs.getInt(4);    //T2(引上)
                    ret.addElement(BtList);
                }
            } // for end
            log("CZSystem getPVDataBtList","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getPVDataBtList","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getPVDataBtList","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }

    //
    //  炉毎引き上げバッチ一覧取り出し（上位２０項目以降取り出し）
    //  （複数PV実績データ保存用）
	@SuppressWarnings("unchecked")
    public static Vector getPVDataBtList2(String db_name){

        initCheck();
        Vector  ret         = new Vector(1000);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        String view = new String("r_hikiage");
        sql = new String("SELECT batch, hinshu, i_sikomi, no_hikiage, max(t_time) FROM " + db_name + "." + view + " WHERE p_kaisi = 1 GROUP BY batch, hinshu, i_sikomi, no_hikiage ORDER BY max(t_time) DESC");

        log("CZSystem getPVDataBtList2","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getPVDataBtList2","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getPVDataBtList2","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getPVDataBtList2","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql);
            for(i = 0 ; rs.next() ; i++){
                int row = rs.getRow();
                if(row > 20){
                    CZPVDataBtList BtList2   = new CZPVDataBtList();
                    BtList2.flg          = 0;
                    BtList2.batch        = rs.getString(1);  //バッチ番号
                    BtList2.hinshu       = rs.getString(2);  //品種
                    BtList2.i_sikomi     = rs.getInt(3);    //仕込量
                    BtList2.no_hikiage   = rs.getInt(4);    //T2(引上)
                    ret.addElement(BtList2);
                }
            } // for end
            log("CZSystem getPVDataBtList2","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getPVDataBtList2","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getPVDataBtList2","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }

    //
    //  クライアントバージョン取得 @@@@@@@@
    //
	@SuppressWarnings("unchecked")
    private static int ClientVersionGet(){

        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        sql = new String("SELECT * FROM " + "mst." + "m_client_version where ap_name = 'CZSystem'");  

        log("CZSystem ClientVersionGet","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem ClientVersionGet","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return -1;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem ClientVersionGet","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return -1;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem ClientVersionGet","ERROR: createStatement or database");
            return -1;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;

            for(i = 0 ; rs.next() ; i++){
                 Client_ver_list = rs.getDouble(2);
            } // for end
            log("CZSystem ClientVersionGet","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem ClientVersionGet","ERROR: Select failed.");
        }
        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem ClientVersionGet","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        return i;
    } 

    //
    //  プロセスNo、プロセス連番取り出し
    //  （複数PV実績データ保存用）
	@SuppressWarnings("unchecked")
    public static Vector getPvProcNo(String db_name, int proc, String bt, String spec, int ichg, int rcp){

        initCheck();
        Vector  ret         = new Vector(1000);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        if(proc == 10){
            sql = new String("SELECT p_no, sp_no, p_renban FROM " + db_name + "." + "r_start where " +
                             " p_start in (SELECT t_time FROM " + db_name + "." + "r_hikiage_temp where batch = '" + bt.trim() + 
                             "' and hinshu = '" + spec.trim() + "' and i_sikomi = " + ichg + " and no_hikiage = " + rcp + ") ORDER BY p_renban");
        }else{
            sql = new String("SELECT p_no, sp_no, p_renban FROM " + db_name + "." + "r_start where p_no = " + proc + 
                             " and p_start in (SELECT t_time FROM " + db_name + "." + "r_hikiage_temp where batch = '" + bt.trim() + 
                             "' and hinshu = '" + spec.trim() + "' and i_sikomi = " + ichg + " and no_hikiage = " + rcp + ") ORDER BY p_renban");
        }

        log("CZSystem getPvProcNo","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getPvProcNo","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getPvProcNo","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getPvProcNo","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql);
            for(i = 0 ; rs.next() ; i++){
                    CZSaveDataProcList plist   = new CZSaveDataProcList();
                    plist.p_no     = rs.getInt(1);
                    plist.sp_no    = rs.getInt(2);
                    plist.p_renban = rs.getInt(3);
                    ret.addElement(plist);
            } // for end
            log("CZSystem getPvProcNo","SELECT Count:" + i);
            
        }
        catch( SQLException e ){
            log("CZSystem getPvProcNo","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getPvProcNo","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
	}

    //
    //  ＰＶデータ取り出し
    //  （複数PV実績データ保存用）
	@SuppressWarnings("unchecked")
    public static Vector getPVSomeData(String db_name,String view,int proc_no,int p_len,boolean data_no[]){

        initCheck();
//@@@
        System.gc();
        
//        Vector  ret             = new Vector(50000);
        Vector  ret             = null;

        Connection conn         = null;
        Statement sqlstmt       = null;
        ResultSet rs            = null;
        String    sql           = null;
        StringBuffer sql_tmp    = null;

        int i = 0;

//@@@
        ret = null;
        ret = new Vector();
        
        sql_tmp = new StringBuffer("SELECT p_no, sp_no, p_renban, p_time, sp_time," +
                " p_date, h_ontime, hk_renban, data5"); 

        for(int no = 0 ; no < CZSystemDefine.PV_MAX_LENGTH ; no++){
            if(data_no[no]){
                sql_tmp.append(", data" + (no+1));
            }   
        }
        sql_tmp.append(" FROM " + db_name + "."+ view.trim() + " WHERE p_no = " + proc_no + " and p_renban  = '" + p_len + "' ORDER BY p_time");

        sql = sql_tmp.toString();   
        log("CZSystem getPVData","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getPVData","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getPVData","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getPVData","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;

            for(i = 0 ; rs.next() ; i++){

                CZSystemPVData pv   = new CZSystemPVData();
                pv.p_no             = rs.getInt(1);
                pv.sp_no            = rs.getInt(2);
                pv.p_renban         = rs.getInt(3);
                pv.p_time           = rs.getInt(4);
                pv.sp_time          = rs.getInt(5);
                pv.p_date           = rs.getString(6);
                pv.h_ontime         = rs.getInt(7);
                pv.hk_renban        = rs.getInt(8);
                pv.p_length         = rs.getFloat(9);   //PSXL

                int j = 0;
                for(int no = 0 ; no < CZSystemDefine.PV_MAX_LENGTH ; no++){
                    if(data_no[no]){
                        pv.data[no] = rs.getFloat(10 + j);
                        j++;
                    }   
                }
                ret.addElement(pv);
            } // for end
            log("CZSystem getPVData","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getPVData","SELECT Count:" + i);
            log("CZSystem getPVData","ERROR: Select failed.");
        }
        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();                    //@@
        }
        catch (SQLException e){
            log("CZSystem getPVData","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }


/* 2003.10.21 y.k tuika start  */
    //
    //  操業ＰＶ実績管理情報取得
    //
	@SuppressWarnings("unchecked")
    public static Vector getPvControl(String db_name){

        initCheck();
        Vector  ret         = new Vector(1000);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        String view = new String("r_pv_control");
        sql = new String("SELECT * FROM " + getDBName() + "." + view + 
                      " ORDER BY s_start DESC"); 

        log("CZSystem getPVControl","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getPVControl","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getPVControl","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getPVControl","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZSystemPvControl bt   = new CZSystemPvControl();
				bt.batch		= rs.getString(1);      //バッチ番号
				bt.t_name		= rs.getString(2);		//テーブル名
				bt.s_start		= rs.getString(3);		//採取開始日時
				bt.s_end		= rs.getString(4);		//採取終了日時
				bt.m_flg        = rs.getInt(5);			//間引き有無
				bt.m_sumi       = rs.getInt(6);			//間引き済
				bt.mo_flg       = rs.getInt(7);			//ＭＯ保存フラグ
				bt.mo_date      = rs.getString(8);		//ＭＯ保存日時
                ret.addElement(bt);
            } // for end
            log("CZSystem getPVControl","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getPVControl","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getPVControl","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }
/* 2003.10.21 y.k tuika end  */

    //
    //  スタート時間取り出し
    //
	@SuppressWarnings("unchecked")
    public static Vector getBtStart(String db_name,String bt_no){

        initCheck();
        Vector  ret         = new Vector(1000);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        String view = new String("r_start");
        sql = new String("SELECT * FROM " + db_name+ "." + view +
                        " WHERE batch = '" + bt_no + "'  ORDER BY p_start");
        log("CZSystem getBtStart","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getBtStart","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getBtStart","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getBtStart","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZSystemStart st    = new CZSystemStart();
                st.batch            = rs.getString(1);
                st.p_no             = rs.getInt(2);
                st.sp_no            = rs.getInt(3);
                st.p_renban         = rs.getInt(4);
                st.p_start          = rs.getString(5);
                ret.addElement(st);
            } // for end
            log("CZSystem getBtStart","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getBtStart","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getBtStart","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }

    //
    //  ＰＶデータ取り出し
    //
	@SuppressWarnings("unchecked")
    public static Vector getPVData(String db_name,String view,int p_len,boolean data_no[]){

        initCheck();
//@@@
        System.gc();
        
//        Vector  ret             = new Vector(50000);
        Vector  ret             = null;

        Connection conn         = null;
        Statement sqlstmt       = null;
        ResultSet rs            = null;
        String    sql           = null;
        StringBuffer sql_tmp    = null;

        int i = 0;

//@@@
        ret = null;
        ret = new Vector();
        
        sql_tmp = new StringBuffer("SELECT p_no, sp_no, p_renban, p_time, sp_time," +
                " p_date, h_ontime, hk_renban, data5"); 

        for(int no = 0 ; no < CZSystemDefine.PV_MAX_LENGTH ; no++){
            if(data_no[no]){
                sql_tmp.append(", data" + (no+1));
            }   
        }
        sql_tmp.append(" FROM " + db_name + "."+ view.trim() + " WHERE p_renban  = '" + p_len + "' ORDER BY p_time");

        sql = sql_tmp.toString();   
        log("CZSystem getPVData","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getPVData","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getPVData","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getPVData","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;

            for(i = 0 ; rs.next() ; i++){

                CZSystemPVData pv   = new CZSystemPVData();
                pv.p_no             = rs.getInt(1);
                pv.sp_no            = rs.getInt(2);
                pv.p_renban         = rs.getInt(3);
                pv.p_time           = rs.getInt(4);
                pv.sp_time          = rs.getInt(5);
                pv.p_date           = rs.getString(6);
                pv.h_ontime         = rs.getInt(7);
                pv.hk_renban        = rs.getInt(8);
                pv.p_length         = rs.getFloat(9);   //PSXL

                int j = 0;
                for(int no = 0 ; no < CZSystemDefine.PV_MAX_LENGTH ; no++){
                    if(data_no[no]){
                        pv.data[no] = rs.getFloat(10 + j);
                        j++;
                    }   
                }
                ret.addElement(pv);
            } // for end
            log("CZSystem getPVData","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getPVData","SELECT Count:" + i);
            log("CZSystem getPVData","ERROR: Select failed.");
        }
        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();                    //@@
        }
        catch (SQLException e){
            log("CZSystem getPVData","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }


    //
    //  エラー取り出し
    //
	@SuppressWarnings("unchecked")
    public static Vector getRoError(String db_name,int day){

        initCheck();
        Vector  ret         = new Vector(5000);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        long day_l  = -day;
        String date = dayTime(day_l);

        sql = new String("SELECT e_no, o_time, batch, p_no, sp_no, p_renban, p_time, sp_time," +
            " flg_error, info1, info2, ro_info, ban_info, k_time FROM " + db_name + "." + "r_error WHERE o_time > " +
             "TO_DATE('" + date + "', 'YYYY-MM-DD HH24:MI:SS')  ORDER BY o_time DESC"); 

        log("CZSystem getRoError","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getRoError","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getRoError","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getRoError","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;

            for(i = 0 ; rs.next() ; i++){
                CZSystemErr st  = new CZSystemErr();
                st.e_no         = rs.getInt(1);
                st.o_time       = rs.getString(2);
                st.batch        = rs.getString(3);
                st.p_no         = rs.getInt(4);
                st.sp_no        = rs.getInt(5);
                st.p_renban     = rs.getInt(6);
                st.p_time       = rs.getInt(7);
                st.sp_time      = rs.getInt(8);
                st.flg_error    = rs.getInt(9);
                st.info1        = rs.getInt(10);
                st.info2        = rs.getInt(11);
                st.ro_info      = rs.getString(12);
                st.ban_info     = rs.getString(13);
                st.k_time       = rs.getString(14);
                ret.addElement(st);
            } // for end
            log("CZSystem getRoError","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getRoError","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getRoError","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }

    //
    //  エラー件数取り出し
    //
    public static int getRoErrorCount(String db_name,int day){

        initCheck();
        int   ret           = 0;
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        long day_l  = -day;
        String date = dayTime(day_l);

        sql = new String("SELECT count(*) as cnt FROM " + db_name + "." + "r_error WHERE o_time > " +
             "TO_DATE('" + date + "', 'YYYY-MM-DD HH24:MI:SS')"); 

        log("CZSystem getRoError","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getRoError","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return 0;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getRoError","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return 0;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getRoError","ERROR: createStatement or database");
            return 0;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;

            for(i = 0 ; rs.next() ; i++){
                ret = rs.getInt(1);
            } // for end
            log("CZSystem getRoError","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getRoError","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();                    //@@
        }
        catch (SQLException e){
            log("CZSystem getRoError","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return 0;
        return ret;
    }

    //
    //  サーバーエラー取り出し
    //
	@SuppressWarnings("unchecked")
    public static Vector getHostError(int day){

        initCheck();
        Vector  ret         = new Vector(5000);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        long day_l = -day;
        String date = dayTime(day_l);

        sql = new String("SELECT e_no, o_time, p_no, info1, info2, mname, k_time FROM " +
         "mst." + "m_error WHERE o_time > " +
         "TO_DATE('" + date + "', 'YYYY-MM-DD HH24:MI:SS') ORDER BY o_time DESC");  

        log("CZSystem getHostError","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getHostError","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getHostError","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getHostError","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZSystemHostErr st  = new CZSystemHostErr();
                st.e_no         = rs.getInt(1);
                st.o_time       = rs.getString(2);
                st.p_no         = rs.getInt(3);
                st.info1        = rs.getInt(4);
                st.info2        = rs.getInt(5);
                st.mname        = rs.getString(6);
                st.k_time       = rs.getString(7);
                ret.addElement(st);
                i++;
            } // for end
            log("CZSystem getHostError","SELECT Count:" + i);
        }
        catch( SQLException e ){
//@@            System.out.println("CZSystem getHostError SQLException : " + e );
            log("CZSystem getHostError","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getHostError","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }


    //
    //  テーブル操作履歴
    //
	@SuppressWarnings("unchecked")
    public static Vector getRoTblModify(String db_name,int day){
        initCheck();

        Vector  ret         = new Vector(5000);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        long day_l = -day;
        String date = dayTime(day_l);

        sql = new String("SELECT s_time, op_name, batch, message, key1, key2, key3" +
           "  FROM " + db_name + "." + "vr_modify WHERE s_time > TO_DATE('" +
           date + "', 'YYYY-MM-DD HH24:MI:SS') ORDER BY s_time DESC");  

        log("CZSystem getRoTblModify","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getRoTblModify","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getRoTblModify","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getRoTblModify","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
            CZSystemTblModify st    = new CZSystemTblModify();
                st.s_time           = rs.getString(1);
                st.op_name          = rs.getString(2);
                st.batch            = rs.getString(3);
                st.message          = rs.getString(4);
                st.key1             = rs.getInt(5);
                st.key2             = rs.getInt(6);
                st.key3             = rs.getInt(7);
                ret.addElement(st);
            } // for end
            log("CZSystem getRoTblModify","SELECT Count:" + i);
        }
        catch( SQLException e ){
        log("CZSystem getRoTblModify","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getRoTblModify","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }


    //
    //  オペレータ介入操作履歴
    //
	@SuppressWarnings("unchecked")
    public static Vector getRoOperation(String db_name,int day){

        initCheck();
        Vector  ret         = new Vector(5000);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        long day_l = -day;
        String date = dayTime(day_l);
        sql = new String("SELECT s_time, batch, p_name, p_renban, p_time, message, sid, val1, val2, val3" +
              " FROM " + db_name + "." + "vr_operate WHERE s_time > TO_DATE('" +
              date + "', 'YYYY-MM-DD HH24:MI:SS') ORDER BY s_time DESC");   
        log("CZSystem getRoOperation","SQL["+sql+"]");
        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getRoOperation","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getRoOperation","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getRoOperation","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZSystemOperation st    = new CZSystemOperation();
                st.s_time               = rs.getString(1);
                st.batch                = rs.getString(2);
                st.p_name               = rs.getString(3);
                st.p_renban             = rs.getInt(4);
                st.p_time               = rs.getInt(5);
                st.message              = rs.getString(6);
                st.sid                  = rs.getInt(7);
                st.val1                 = rs.getInt(8);
                st.val2                 = rs.getInt(9);
                st.val3                 = rs.getInt(10);
                ret.addElement(st);
            } // for end
            log("CZSystem getRoOperation","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getRoOperation","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();                    //@@
        }
        catch (SQLException e){
            log("CZSystem getRoOperation","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }


    //
    //  ＣＣＤ波形取り出し
    //
	@SuppressWarnings("unchecked")
    public static Vector getCCDWave(){
        initCheck();

        Vector      ret     = new Vector(500);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;
        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getCCDWave","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getCCDWave","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
            sql = new String("SELECT * FROM "+ getDBName() +"." + "r_ccd_hakei ORDER BY s_time DESC");  
            log("CZSystem getCCDWave","SQL["+sql+"]");
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getCCDWave","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZSystemCCDWave d = new CZSystemCCDWave();
                d.s_time    = rs.getString(1);
                d.batch     = rs.getString(2);
                d.p_no      = rs.getInt(3);
                d.sp_no     = rs.getInt(4);
                d.p_renban  = rs.getInt(5);
                d.p_time    = rs.getInt(6);
                d.sp_time   = rs.getInt(7);
                d.slice     = rs.getString(8);
                d.s_start   = rs.getInt(9);
                d.s_end     = rs.getInt(10);
                d.single    = rs.getFloat(11);
                d.k_chokei  = rs.getFloat(12);
                d.h_chokei  = rs.getFloat(13);
                d.v_keisoku = rs.getInt(14);
                d.h_keisoku = rs.getInt(15);
                d.status    = rs.getString(16);
                d.route     = rs.getString(17);
                d.cross     = rs.getString(18);
                d.search    = rs.getString(19);
                d.peek      = rs.getString(20);
                d.hosei     = rs.getString(21);
                d.len       = rs.getInt(22);
                d.data      = rs.getString(23);
                ret.addElement(d);
            } // for end
            log("CZSystem getCCDWave","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getCCDWave","ERROR: Select failed. [" + e + "]");
        }
        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getCCDWave","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }

    //
    //  ＣＣＤ画像情報取り出し
    //
	@SuppressWarnings("unchecked")
    public static Vector getCCDBMP(){

        initCheck();
        Vector      ret     = new Vector(500);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;
        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getCCDBMP","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getCCDBMP","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return null;
        }

            
        try{
            sqlstmt = conn.createStatement() ;
            sql = new String("SELECT * FROM " + getDBName() + "." + "r_ccd_screen ORDER BY s_time DESC");   
            log("CZSystem getCCDBMP","SQL["+sql+"]");
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getCCDBMP","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZSystemCCDBMP d    = new CZSystemCCDBMP();
                d.s_time            = rs.getString(1);
                d.batch             = rs.getString(2);
                d.p_no              = rs.getInt(3);
                d.sp_no             = rs.getInt(4);
                d.p_renban          = rs.getInt(5);
                d.p_time            = rs.getInt(6);
                d.sp_time           = rs.getInt(7);
                d.f_name            = rs.getString(8);
                ret.addElement(d);
            } // for end
            log("CZSystem getCCDBMP","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getCCDBMP","ERROR: Select failed. [" + e + "]");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getCCDBMP","ERROR: Close ResultSet or Statement");
        }
         closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }

    //
    //  エラー名取り出し
    //
	@SuppressWarnings("unchecked")
    public static Vector getErrTitle(){

        initCheck();
        Vector      ret     = new Vector(CZSystemDefine.ERROR_MAX);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;
        int i = 0;

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getErrTitle","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getErrTitle","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;

            sql = new String("SELECT * FROM " + "mst.m_sg_error" + " ORDER BY e_no");   
            log("CZSystem getErrTitle","SQL["+sql+"]");

        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getErrTitle","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZSystemErrName title = new CZSystemErrName();
                title.e_no      = rs.getInt(1);
                title.e_name    = rs.getString(2);
                title.process   = rs.getInt(3);
                title.edge      = rs.getInt(4);
                title.ready     = rs.getInt(5);
                title.kubun     = rs.getInt(6);
                title.basho     = rs.getInt(7);
                title.umu       = rs.getInt(8);
                title.buzzer1   = rs.getInt(9);
                title.buzzer    = rs.getInt(10);
                title.error_umu = rs.getInt(11);
                title.fukkyu    = rs.getInt(12);
                title.hyoji     = rs.getInt(13);
                ret.addElement(title);
            } // for end
            log("CZSystem getErrTitle","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getErrTitle","ERROR: Select failed.");
        }
        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getErrTitle","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }


    //
    //  炉番変更
    //
    public static synchronized void chgRo(int ro){

        log("CZSystem chgRo","1 炉index:" + ro );
        if(0 > ro){
                exit(-1,"CZSystem chgRo Error !! [" + ro + "]");
        }
        initCheck();
        try{
            CZSystemSysMsg msg = new CZSystemSysMsg();

            String s = CZSystem.RoKetaChg(getRoName());
            String ss = CZSystem.RoKetaChg(getRoName(ro));

            msg.no = 0;
//            msg.message = CZSystem.getDateTime() + "  [ 炉番変更  " +   
//                         getRoName() + " → " + getRoName(ro) + " ]";
            msg.message = CZSystem.getDateTime() + "  [ 炉番変更  " +   
                         s + " → " + ss + " ]";
            CZSystem.sysMessage(msg);
        }
        catch (Throwable e) {
            log("CZSystem chgRo","CZSystemSysMsg 炉index:" + ro );
            handleException(e);
        }
        try{
            // ここ以外で変更不可
            ro_no_idx = ro;
            String _ro = getRoName();
            final_ro_no = _ro;
            if(CZSystemDefine.LIB_MODE == system_mode){
                log("CZSystem chgRo","2 炉index:" + ro + " 炉No:" + _ro);
                log("CZSystem chgRo","RETURN LIB_MODE");
                return;
            }
            log("CZSystem chgRo","2 炉index:" + ro + " 炉No:" + _ro);
            // 操作応答イベントの立ち上げ
            initHorbClientResult();
            log("CZSystem chgRo","3 炉index:" + ro + " 炉No:" + _ro);

            CZNativeHikiage _bt_set     = cz_gd_px.CZNativeHikiageGet(_ro);  //  引き上げ条件
            CZNativeDengen  _dengen     = cz_gd_px.CZNativeDengenGet(_ro);   //  電源情報
            log("CZSystem chgRo","4 炉index:" + ro + " 炉No:" + _ro);

            CZNativePv p                = cz_gd_px.CZNativePvGet(_ro);
            String  _bt                 = p.getBatch();         //  バッチNo
            int _proc                   = p.getP_no();          //  プロセスNo
            int _sub_proc               = p.getSp_no();         //  サブプロセスNo
            int _proc_len               = p.getP_renban();      //  プロセス連番
            int _proc_time              = p.getP_time();        //  プロセス時間
            int _sub_proc_time          = p.getSp_time();       //  サブプロセス時間
            int _get_date_time          = p.getP_date();        //  採取日時
            int _main_heat_on_time      = p.getH_ontime();      //  メインヒータ電源オン時間
            int _condition_len          = p.getHk_renban();     //  引上げ条件内連番
            float   _pv[]               = p.getData();          //  データ
            log("CZSystem chgRo","5 炉index:" + ro + " 炉No:" + _ro);

            setCurrentData( _bt_set , _dengen , _bt, _proc , _sub_proc ,    
                    _proc_len , _proc_time , _sub_proc_time ,   
                    _get_date_time , _main_heat_on_time , _condition_len , _pv);    
            log("CZSystem chgRo","6 炉index:" + ro + " 炉No:" + _ro);

            if(!CZPV.newCZPV()) CZSystem.exit(-1,"CZSystem chgRo newCZPV()");
            log("CZSystem chgRo","7 炉index:" + ro + " 炉No:" + _ro);

            CZSystemPVNamePMM ret = untenRead(current_proc);
            // @@@@ null
            if ( null != ret ) {
                chgUnten(ret);
            }
            log("CZSystem chgRo","8 炉index:" + ro + " 炉No:" + _ro);
            String db   = getDBName();
            String view = getViewName();
            if(null == view) return;
            CZPVDBReader r = new CZPVDBReader(db,view,current_proc_len, 
                              ret.item[0], ret.item[1],
                              ret.item[2], ret.item[3], ret.item[4], ret.item[5], ret.item[6], ret.item[7], ret.item[8], ret.item[9]);	// @20131030
            db_thread = new Thread(r);
            db_thread.setPriority(Thread.MIN_PRIORITY);
            db_thread.start();
            db_thread.join();
            CZEventSender.sendData(current_bt,CZEventCL.RO_CHANGE);
        }
        catch (Throwable e) {
            log("CZSystem chgRo","炉index:" + ro );
            handleException(e);
        }
        log("CZSystem chgRo","9 炉index:" + ro_no_idx + " 炉No:" + ro);
    }


    //
    //  プロセス変更
    //
    private static synchronized void chgProc(int proc_len,boolean flag){

        initCheck();
        CZSystemSysMsg msg = new CZSystemSysMsg();
        msg.no = 0;
        msg.message = CZSystem.getDateTime() + "  [ プロセス変更 ]";
        CZSystem.sysMessage(msg);
        try{
            if(null != db_thread){
                 if(db_thread.isAlive()) db_thread.join();
            }
            if(!CZPV.newCZPV()) CZSystem.exit(-1,"CZSystem chgRo newCZPV()");
            CZSystemPVNamePMM ret = untenRead(current_proc);
            // @@@@ null
            if ( null != ret ) {
                if(flag) chgUnten(ret);
            }
            String db   = getDBName();
            String view = getViewName();
            if(null == view) return;
            CZPVDBReader r = new CZPVDBReader(db,view,current_proc_len, 
                              ret.item[0], ret.item[1],
                              ret.item[2], ret.item[3], ret.item[4], ret.item[5], ret.item[6], ret.item[7], ret.item[8], ret.item[9]);	// @20131030
            db_thread = new Thread(r);
            db_thread.setPriority(Thread.MIN_PRIORITY);
            db_thread.start();
        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }
    }


    //
    //  運転画面ＰＶ設定
    //
    private  static synchronized boolean chgUnten(CZSystemPVNamePMM p){

        initCheck();
        if (null == p) return false;
        
        if(p.X_LENGTH == p.x_shubetu){
            CZPV.setPVGrTimeFlag(false);
        }
        else {
            CZPV.setPVGrTimeFlag(true);
        }
        CZPV.setPVGrTimeScale(p.x_time);
        CZPV.setPVGrLengthScale(p.x_width);
        //ＰＶ表示
        CZPV.setPVGrNo(p.item);
        //ｍｉｎ、ｍａｘの設定
        CZPV.setPVGrMin(p.min);
        CZPV.setPVGrMax(p.max);
        //ＰＶ名、単位、の設定
        String k_name[] = new String[CZPV.PV_DATA_SET_LENGTH];
        String k_unit[] = new String[CZPV.PV_DATA_SET_LENGTH];
        for(int i = 0 ; i < CZPV.PV_DATA_SET_LENGTH ; i++){
            CZSystemPVName pn = getPVName(p.item[i] - 1);
            k_name[i] = pn.k_name;  
            k_unit[i] = pn.unit;    
        }
        CZPV.setPVGrName(k_name);
        CZPV.setPVGrUnit(k_unit);
        return true;
    }


    //
    //  引き上げ条件ダウンロード
    //
    public static synchronized boolean CZOperateHikiage(CZParamHikiage dat){

        initCheck();
        String ro = getRoName();
        int rc = cz_op_px.CZOperateHikiage(ro,dat);
        log("CZSystem CZOperateHikiage","rc:" + rc);

        if(0 != rc) return false;
        return true;
    }


    //
    //  引き上げ条件ダウンロード（取り出しテーブルのみ）
    //
    public static synchronized boolean CZOperateToridasi(CZParamHikiage dat){

        initCheck();
        String ro = getRoName();
        int rc = cz_op_px.CZOperateToridasi(ro,dat);
        log("CZSystem CZOperateToridasi","rc:" + rc);
        if(0 != rc) return false;
        return true;
    }


    //
    // ＣＣＤ生波形採取
    //
    //
    public static boolean CZOperateWaveCollect(String roban){

        initCheck();
        int rc = cz_op_px.CZOperateWaveCollect(roban);
        log("CZSystem CZOperateWaveCollect","rc:" + rc);
        if(0 != rc) return false;
        return true;
    }


    //
    // ＣＣＤカメラ画像保存
    //
    //
    public static boolean CZOperateCcdCamera(String roban){

        initCheck();
        int rc = cz_op_px.CZOperateCcdCamera(roban);
        log("CZSystem CZOperateCcdCamera","rc:" + rc);
        if(0 != rc) return false;
        return true;
    }


    //
    //  電源変更
    //
    public static synchronized boolean CZOperatePowerControl(int[] units){

        initCheck();
        String ro = getRoName();
        for(int i = 0 ; i <  10  ; i++){
            log("CZSystem CZOperatePowerControl","[" + units[i] + "]");
        }
        int rc = cz_op_px.CZOperatePowerControl(ro,units);
        log("CZSystem CZOperatePowerControl","rc:" + rc);
        if(0 != rc) return false;
        return true;
    }


    //
    //  制御モード変更
    //
    public static synchronized boolean CZOperateModeExchange(int nowMode,   
                                     int shiftMode){

        initCheck();
        String ro = getRoName();
        log("CZSystem CZOperateModeExchange","NOW[" + nowMode + 
                             "] -> [" + shiftMode + "]");
        int rc = cz_op_px.CZOperateModeExchange(ro,nowMode,shiftMode);
        log("CZSystem CZOperateModeExchange","rc:" + rc);
        if(0 != rc) return false;
        return true;
    }


    //
    //  プロセス変更
    //
    public static synchronized boolean CZOperateProcessExchange(int nowProc,    
                                        int shiftProc){
        initCheck();
        String ro = getRoName();
        log("CZSystem CZOperateProcessExchange","NOW[" + nowProc +  
                             "] -> [" + shiftProc + "]");
        int rc = cz_op_px.CZOperateProcessExchange(ro,nowProc,shiftProc);
        log("CZSystem CZOperateProcessExchange","rc:" + rc);
        if(0 != rc) return false;
        return true;
    }


    //////////////////////////////////////////////////////////////////////////////////////////////
    //
    //  シード上昇
    //
    public static synchronized boolean CZOperateSeedUp(int value){

        initCheck();
        String ro = getRoName();
        log("CZSystem CZOperateSeedUp","VAL[" + value + "]");
        int rc = cz_op_px.CZOperateSeedUp(ro,value,true);
        log("CZSystem CZOperateSeedUp","rc:" + rc);
        if(0 != rc) return false;
        return true;
    }
    //
    //  シード上昇 ＵＮＤＯ
    //
    public static synchronized boolean CZOperateUndoSeedUp(int value){

        initCheck();
        String ro = getRoName();
        log("CZSystem CZOperateUndoSeedUp","VAL[" + value + "]");
        int rc = cz_op_px.CZOperateUndoSeedUp(ro,true);
        log("CZSystem CZOperateSeedUp","rc:" + rc);
        if(0 != rc) return false;
        return true;
    }


    //
    //  シード回転
    //
    public static synchronized boolean CZOperateSeedRotate(int value){

        initCheck();
        String ro = getRoName();
        log("CZSystem CZOperateSeedRotate","VAL[" + value + "]");
        int rc = cz_op_px.CZOperateSeedRotate(ro,value,true);
        log("CZSystem CZOperateSeedRotate","rc:" + rc);
        if(0 != rc) return false;
        return true;
    }
    //
    //  シード回転 ＵＮＤＯ
    //
    public static synchronized boolean CZOperateUndoSeedRotate(int value){

        initCheck();
        String ro = getRoName();
        log("CZSystem CZOperateUndoSeedRotate","VAL[" + value + "]");
        int rc = cz_op_px.CZOperateUndoSeedRotate(ro,true);
        log ("CZSystem CZOperateUndoSeedRotate","rc:" + rc);
        if(0 != rc) return false;
        return true;
    }


    //
    //  シード位置
    //
    public static synchronized boolean CZOperateSeedPosition(int value,boolean lock){

        initCheck();
        String ro = getRoName();
        log("CZSystem CZOperateUndoSeedPosition","VAL[" + value + "]");
        int rc = cz_op_px.CZOperateSeedPosition(ro,value,lock);
        log("CZSystem CZOperateUndoSeedPosition","rc:" + rc);
        if(0 != rc) return false;
        return true;
    }
    //
    //  シード位置 ＵＮＤＯ
    //
    public static synchronized boolean CZOperateUndoSeedPosition(int value,boolean lock){

        initCheck();
        String ro = getRoName();
        log("CZSystem CZOperateUndoSeedPosition","VAL[" + value + "]");
        int rc = cz_op_px.CZOperateUndoSeedPosition(ro,lock);
        log("CZSystem CZOperateUndoSeedPosition","rc:" + rc);
        if(0 != rc) return false;
        return true;
    }


    //////////////////////////////////////////////////////////////////////////////////////////////
    //
    //  ルツボ上昇
    //
    public static synchronized boolean CZOperateRutuboUp(int value){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZOperateRutuboUp VAL[" + value + "]");
        int rc = cz_op_px.CZOperateRutuboUp(ro,value,true);
        log("CZSystem","CZOperateRutuboUp rc:" + rc);
        if(0 != rc) return false;
        return true;
    }

    //
    //  ルツボ上昇 ＵＮＤＯ
    //
    public static synchronized boolean CZOperateUndoRutuboUp(int value){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZOperateUndoRutuboUp VAL[" + value + "]");
        int rc = cz_op_px.CZOperateUndoRutuboUp(ro,true);
        log("CZSystem","CZOperateUndoRutuboUp rc:" + rc);
        if(0 != rc) return false;
        return true;
    }

    //
    //  ルツボ回転
    //
    public static synchronized boolean CZOperateRutuboRotate(int value){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZOperateRutuboRotate VAL[" + value + "]");
        int rc = cz_op_px.CZOperateRutuboRotate(ro,value,true);
        log("CZSystem","CZOperateRutuboRotate rc:" + rc);
        if(0 != rc) return false;
        return true;
    }

    //
    //  ルツボ回転 ＵＮＤＯ
    //
    public static synchronized boolean CZOperateUndoRutuboRotate(int value){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZOperateUndoRutuboRotate VAL[" + value + "]");
        int rc = cz_op_px.CZOperateUndoRutuboRotate(ro,true);
        log("CZSystem","CZOperateUndoRutuboRotate rc:" + rc);
        if(0 != rc) return false;
        return true;
    }


    //
    //  ルツボ位置
    //
    public static synchronized boolean CZOperateRutuboPosition(int value,boolean lock){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZOperateRutuboPosition VAL[" + value + "]");
        int rc = cz_op_px.CZOperateRutuboPosition(ro,value,lock);
        log("CZSystem","CZOperateRutuboPosition rc:" + rc);
        if(0 != rc) return false;
        return true;
    }
    //
    //  ルツボ位置 ＵＮＤＯ
    //
    public static synchronized boolean CZOperateUndoRutuboPosition(int value,boolean lock){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZOperateUndoRutuboPosition VAL[" + value + "]");
        int rc = cz_op_px.CZOperateUndoRutuboPosition(ro,lock);
        log("CZSystem","CZOperateUndoRutuboPosition rc:" + rc);
        if(0 != rc) return false;
        return true;
    }


    //////////////////////////////////////////////////////////////////////////////////////////////
    //
    //  保持具位置
    //
    public static synchronized boolean CZOperateHojiguPosition(int value,boolean lock){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZOperateHojiguPosition VAL[" + value + "]");
        int rc = cz_op_px.CZOperateHojiguPosition(ro,value,lock);
        log("CZSystem","CZOperateHojiguPosition rc:" + rc);
        if(0 != rc) return false;
        return true;
    }
    //
    //  保持具位置 ＵＮＤＯ
    //
    public static synchronized boolean CZOperateUndoHojiguPosition(int value,boolean lock){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZOperateUndoHojiguPosition VAL[" + value + "]");
        int rc = cz_op_px.CZOperateUndoHojiguPosition(ro,lock);
        log("CZSystem","CZOperateUndoHojiguPosition rc:" + rc);
        if(0 != rc) return false;
        return true;
    }


    //////////////////////////////////////////////////////////////////////////////////////////////
    //
    //  ヒーター１電力
    //
    public static synchronized boolean CZOperateMainHeater1Power(int value){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZOperateMainHeater1Power VAL[" + value + "]");
        int rc = cz_op_px.CZOperateMainHeater1Power(ro,value);
        log("CZSystem","CZOperateMainHeater1Power rc:" + rc);
        if(0 != rc) return false;
        return true;
    }
    //
    //  ヒーター１電力 ＵＮＤＯ
    //
    public static synchronized boolean CZOperateUndoMainHeater1Power(int value){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZOperateUndoMainHeater1Power VAL[" + value + "]");
        int rc = cz_op_px.CZOperateUndoMainHeater1Power(ro);
        log("CZSystem","CZOperateUndoMainHeater1Power rc:" + rc);
        if(0 != rc) return false;
        return true;
    }


    //
    //  ヒーター２電力
    //
    public static synchronized boolean CZOperateMainHeater2Power(int value){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZOperateMainHeater2Power VAL[" + value + "]");
        int rc = cz_op_px.CZOperateMainHeater2Power(ro,value);
        log("CZSystem","CZOperateMainHeater2Power rc:" + rc);
        if(0 != rc) return false;
        return true;
    }
    //
    //  ヒーター２電力 ＵＮＤＯ
    //
    public static synchronized boolean CZOperateUndoMainHeater2Power(int value){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZOperateUndoMainHeater2Power VAL[" + value + "]");
        int rc = cz_op_px.CZOperateUndoMainHeater2Power(ro);
        log("CZSystem","CZOperateUndoMainHeater2Power rc:" + rc);
        if(0 != rc) return false;
        return true;
    }

    //
    //  サブヒーター電力
    //
    public static synchronized boolean CZOperateSubHeaterPower(int value){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZOperateSubHeaterPower VAL[" + value + "]");
        int rc = cz_op_px.CZOperateSubHeaterPower(ro,value);
        log("CZSystem","CZOperateSubHeaterPower rc:" + rc);
        if(0 != rc) return false;
        return true;
    }
    //
    //  サブヒーター電力 ＵＮＤＯ
    //
    public static synchronized boolean CZOperateUndoSubHeaterPower(int value){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZOperateUndoSubHeaterPower VAL[" + value + "]");
        int rc = cz_op_px.CZOperateUndoSubHeaterPower(ro);
        log("CZSystem","CZOperateUndoSubHeaterPower rc:" + rc);
        if(0 != rc) return false;
        return true;
    }

    //
    //  シードヒーター電力
    //
    public static synchronized boolean CZOperateSeedHeater(int value){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZOperateSeedHeater VAL[" + value + "]");
        int rc = cz_op_px.CZOperateSeedHeater(ro,value);
        log("CZSystem","CZOperateSeedHeater rc:" + rc);
        if(0 != rc) return false;
        return true;
    }
    //
    //  シードヒーター電力 ＵＮＤＯ
    //
    public static synchronized boolean CZOperateUndoSeedHeater(int value){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZOperateUndoSeedHeater VAL[" + value + "]");
        int rc = cz_op_px.CZOperateUndoSeedHeater(ro);
        log("CZSystem","CZOperateUndoSeedHeater rc:" + rc);
        if(0 != rc) return false;
        return true;
    }


    //
    //  ヒーター１温度
    //
    public static synchronized boolean CZOperateMainHeaterTemp(int value){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZOperateMainHeaterTemp VAL[" + value + "]");
        int rc = cz_op_px.CZOperateMainHeaterTemp(ro,value);
        log("CZSystem","CZOperateMainHeaterTemp rc:" + rc);
        if(0 != rc) return false;
        return true;
    }
    //
    //  ヒーター１温度 ＵＮＤＯ
    //
    public static synchronized boolean CZOperateUndoMainHeaterTemp(int value){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZOperateUndoMainHeaterTemp VAL[" + value + "]");
        int rc = cz_op_px.CZOperateUndoMainHeaterTemp(ro);
        log("CZSystem","CZOperateUndoMainHeaterTemp rc:" + rc);
        if(0 != rc) return false;
        return true;
    }


    //////////////////////////////////////////////////////////////////////////////////////////////
    //
    //  プルアルゴン
    //
    public static synchronized boolean CZOperatePullArgonFlow(int value){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZOperatePullArgonFlow VAL[" + value + "]");
        int rc = cz_op_px.CZOperatePullArgonFlow(ro,value);
        log("CZSystem","CZOperatePullArgonFlow rc:" + rc);
        if(0 != rc) return false;
        return true;
    }
    //
    //  プルアルゴン ＵＮＤＯ
    //
    public static synchronized boolean CZOperateUndoPullArgonFlow(int value){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZOperateUndoPullArgonFlow VAL[" + value + "]");
        int rc = cz_op_px.CZOperateUndoPullArgonFlow(ro);
        log("CZSystem","CZOperateUndoPullArgonFlow rc:" + rc);
        if(0 != rc) return false;
        return true;
    }


    //
    //  トップアルゴン
    //
    public static synchronized boolean CZOperateTopArgonFlow(int value){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZOperateTopArgonFlow VAL[" + value + "]");
        int rc = cz_op_px.CZOperateTopArgonFlow(ro,value);
        log("CZSystem","CZOperateTopArgonFlow rc:" + rc);
        if(0 != rc) return false;
        return true;
    }

    //
    //  トップアルゴン ＵＮＤＯ
    //
    public static synchronized boolean CZOperateUndoTopArgonFlow(int value){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZOperateUndoTopArgonFlow VAL[" + value + "]");
        int rc = cz_op_px.CZOperateUndoTopArgonFlow(ro);
        log("CZSystem","CZOperateUndoTopArgonFlow rc:" + rc);
        if(0 != rc) return false;
        return true;
    }

    //////////////////////////////////////////////////////////////////////////////////////////////
    //
    //  CCDカメラモニタリング切替え
    //
    public static synchronized boolean CZOperateCcdChange(int value){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZOperateCcdChange VAL[" + value + "]");
        int rc = cz_op_px.CZOperateCcdChange(ro,value);
//        int rc = cz_op_px.CZOperateBuzzerOff(ro);
        log("CZSystem","CZOperateCcdChange rc:" + rc);
        if(0 != rc) return false;
        return true;
    }


    //////////////////////////////////////////////////////////////////////////////
    //
    //  エラー項目設定
    //
    public static synchronized boolean CZErrorDefineSend(String op,Vector param){

        System.out.println("CZErrorDefineSend start");
        initCheck();
        int rc = -1;
        System.out.println("send_data size [" + param.size() + "]");
        if(param.size() != CZSystemDefine.ERROR_MAX){
            exit(0,"CZErrorDefineSend()  DATA SIZE OVER [" + param.size() + "]");
        }
        System.out.println("CZErrorDefineSend 1");
        CZParamErrorDefine data[] = new CZParamErrorDefine[CZSystemDefine.ERROR_MAX];
        for(int i = 0 ; i < param.size() ; i++){
            CZParamErrorDefine d = (CZParamErrorDefine)param.elementAt(i);
            if(null == d ){
                exit(0,"CZErrorDefineSend()  DATA NULL [" + i + "]");
            }
            data[i] = d;
        }
        System.out.println("CZErrorDefineSend 2");
        rc = cz_tb_px.CZErrorDefineSend(op,data);
        log("CZSystem CZErrorDefineSend","rc:" + rc + "[" + op +"]");
        System.out.println("CZErrorDefineSend 3");
        if(0 != rc) return false;
        System.out.println("CZErrorDefineSend end");
        return true;
    }



    //////////////////////////////////////////////////////////////////////////////
    //
    //  エラーメッセージ設定
    //
    public static synchronized boolean CZErrorMsgDefineSend(String op,Vector param){

        System.out.println("CZErrorMsgDefineSend start");
        initCheck();
        int rc = -1;
        System.out.println("errmsg_data size [" + param.size() + "]");
        if(param.size() != CZSystemDefine.ERROR_MAX){
            exit(0,"CZErrorMsgDefineSend()  DATA SIZE OVER [" + param.size() + "]");
        }
        System.out.println("CZErrorDefineMsgSend 1");
        CZParamErrorMsgDefine data[] = new CZParamErrorMsgDefine[CZSystemDefine.ERROR_MAX];
        for(int i = 0 ; i < param.size() ; i++){
            CZParamErrorMsgDefine d = (CZParamErrorMsgDefine)param.elementAt(i);
            if(null == d ){
                exit(0,"CZErrorDefineMsgSend()  DATA NULL [" + i + "]");
            }
            data[i] = d;
        }
        System.out.println("CZErrorMsgDefineSend 2");
        rc = cz_tb_px.CZErrorMsgDefineSend(op,data);
        log("CZSystem CZErrorMsgDefineSend","rc:" + rc + "[" + op +"]");
        System.out.println("CZErrorMsgDefineSend 3");
        if(0 != rc) return false;
        System.out.println("CZErrorMsgDefineSend end");
        return true;
    }



    //////////////////////////////////////////////////////////////////////////////
    //
    //  再間引き指示　2003.10.21　y.k
    //
    public static synchronized boolean CZPvControlChgSend(String op, String roban, Vector param){

        initCheck();
        int rc = -1;
		int i,j;

//        if(param.size() != CZSystemDefine.ERROR_MAX){
//            exit(0,"CZErrorDefineSend()  DATA SIZE OVER [" + param.size() + "]");
//        }

	System.out.println ("param.size()=" + param.size() + "param[" + param + "]");
		// データ更新
		CZParamPVMabikiCng[] err = new CZParamPVMabikiCng[param.size()];
		for (i=j=0; i<param.size(); i++) {
		
			CZPVDataSave.DispBtColorTbl d = (CZPVDataSave.DispBtColorTbl)param.elementAt(i);

System.out.println ("対象情報 batch[" + d.batch + "] flg[" + d.m_sumi + "]");
			err[j] = new CZParamPVMabikiCng();
			err[j].setBatchNo(d.batch);
			err[j].setM_sumi(d.m_sumi_chg);
			j++;
		}

System.out.println ("CZControlPVMabikiChg start");
        rc = cz_tb_px.CZControlPVMabikiChg(op,roban,err);
        log("CZSystem CZPvControlChgSend","rc:" + rc + "[" + op +"]");
        if(0 != rc) return false;
        return true;
    }

    //////////////////////////////////////////////////////////////////////////////
    //
    //  運転画面設定
    //
    public static synchronized boolean CZUntenDefineSend(String op,CZSystemPVNamePMM p){

        initCheck();
        String ro = getRoName();
        CZParamUnten dat = new CZParamUnten();
        dat.setProcess(p.p_no);
        dat.setItemNo(p.item);
        int min[] = new int[p.SIZE];
        int max[] = new int[p.SIZE];
        for(int i = 0 ; i < p.SIZE ; i++){
            CZSystemPVName pv = getPVName(p.item[i] - 1);
            min[i] = (int)(p.min[i] *  Math.pow(10,pv.keta));
            max[i] = (int)(p.max[i] *  Math.pow(10,pv.keta));
        }
        dat.setMinVal(min);
        dat.setMaxVal(max);
        dat.setSel(p.x_shubetu);
        dat.setTimeScale(p.x_time);
        dat.setLenScale(p.x_width);
        // 送信
        int rc = cz_tb_px.CZUntenDefineSend(op,ro,dat);
        log("CZSystem CZUntenDefineSend","rc:" + rc + " [" + op +"]");

        // カレントのプロセスの場合
        if(current_proc == p.p_no){
            chgUnten(p);
            chgProc(current_proc,false);
        }

        if(0 != rc) return false;
        return true;
    }


    //////////////////////////////////////////////////////////////////////////////
    //
    //  操業定数排他要求
    //
    public static synchronized boolean CZGetWorkingExclusion(String ro){

        initCheck();

        int rc = cz_tb_px.CZGetWorkingExclusion(ro);
        log("CZSystem CZGetWorkingExclusion","rc:" + rc + " [" + ro +"]");
        if(0 != rc){
            CZSystemSysMsg msg = new CZSystemSysMsg();
            msg.no = -1;

            switch(rc){
                case 2   : msg.message = CZSystem.getDateTime() +   
                            " 操業定数排他要求失敗 [" + rc + "] 制御盤修正中";
                break;

                case 100 : msg.message = CZSystem.getDateTime() +   
                            " 操業定数排他要求失敗 [" + 
                            rc + "] 制御盤引き上げ条件登録中";
                break;
                default  : msg.message = CZSystem.getDateTime() +   
                            " 操業定数排他要求失敗 [" + rc + "]";
                break;
            }
            sysMessage(msg);
            return false;
        }
        return true;
    }

    //
    //  操業定数排他開放
    //
    public static synchronized boolean CZPutWorkingExclusion(String ro){

        initCheck();
        int rc = cz_tb_px.CZPutWorkingExclusion(ro);
        log("CZSystem CZPutWorkingExclusion","rc:" + rc + " [" + ro +"]");
        if(0 != rc) return false;
        return true;
    }


    //
    //  操業定数設定
    //
    public static synchronized boolean CZWorkingTableExchnage(String op,
                                  int item1, int item2,
                                  float[] data){

        initCheck();
        String ro = getRoName();
        // 送信
        int rc = cz_tb_px.CZWorkingTableExchnage(op,ro,0,item1,item2,data);
        log("CZSystem","CZWorkingTableExchnage rc:" + rc + " [" + op +"]");
        if(0 != rc) return false;
        return true;
    }


    //
    //  操業定数項目設定
    //
    public static synchronized boolean  CZWorkingNameExchnage(String op,    
                                  int itemNo1, int itemNo2, int itemNo,
                                  String itemName, String taniName, 
                                  float min, float max, int fig){

        initCheck();
        // 送信
        int rc = cz_tb_px.CZWorkingNameExchnage(op,itemNo1,itemNo2,itemNo,itemName,taniName,min,max,fig);
        log("CZSystem","CZWorkingNameExchnage rc:" + rc + " [" + op +"]");
        if(0 != rc) return false;
        return true;
    }

    //////////////////////////////////////////////////////////////////////////
    //
    //  操業定数炉間コピー
    //
    public static synchronized boolean CZWorkingCopyRo(String op,String dst_ro){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZWorkingCopyRo Ro[" + ro + "] -> [" + dst_ro + "]");
        if(ro.equals(dst_ro)){
            return false;
        }
        // 送信
        int rc = -1;
        rc = cz_tb_px.CZWorkingCopyRo(op,ro,dst_ro);
        log("CZSystem","CZWorkingCopyRo rc:" + rc + " [" + op +"]");
        if(0 != rc) return false;
        return true;
    }


    //
    //  操業定数炉間大項目コピー
    //
    public static synchronized boolean CZWorkingCopyNo1(String op,String dst_ro,int no1){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZWorkingCopyNo1 Ro[" + ro + "] -> [" + dst_ro + "]");
        if(ro.equals(dst_ro)){
            return false;
        }
        // 送信
        int rc = -1;
        rc = cz_tb_px.CZWorkingCopyNo1(op,ro,dst_ro,no1);
        log("CZSystem CZWorkingCopyNo1","rc:" + rc + " [" + op +"] No1[" + no1 + "]");
        if(0 != rc) return false;
        return true;
    }


    //
    //  操業定数炉間中項目コピー
    //
    public static synchronized boolean CZWorkingCopyNo2(String op,String dst_ro,int no1,int no2){

        initCheck();
        String ro = getRoName();
        log("CZSystem","CZWorkingCopyNo2 Ro[" + ro + "] -> [" + dst_ro + "]");
        if(ro.equals(dst_ro)){
            return false;
        }
        // 送信
        int rc = -1;
        rc = cz_tb_px.CZWorkingCopyNo2(op,ro,dst_ro,no1,no2);
        log("CZSystem CZWorkingCopyNo2","rc:" + rc + " [" + op +"] No1[" + no1 + "] No2[" + no2 + "]");
        if(0 != rc) return false;
        return true;
    }


    //////////////////////////////////////////////////////////////////////////////

    //
    //  制御テーブル排他要求
    //
    public static synchronized boolean CZGetControlExclusion(String ro){

        initCheck();
        int rc = cz_tb_px.CZGetControlExclusion(ro);
        log("CZSystem","CZGetControlExclusion rc:" + rc + " [" + ro +"]");
        if(0 != rc){
            CZSystemSysMsg msg = new CZSystemSysMsg();
            msg.no = -1;
            switch(rc){
                case 2   : msg.message = CZSystem.getDateTime() +   
                            " 制御テーブル排他要求失敗 [" + rc + "] 制御盤修正中";
                break;

                case 100 : msg.message = CZSystem.getDateTime() +   
                            " 制御テーブル排他要求失敗 [" + 
                                rc + "] 制御盤引き上げ条件登録中";
                break;
                default  : msg.message = CZSystem.getDateTime() +   
                            " 制御テーブル排他要求失敗 [" + rc + "]";
                break;
            }
            sysMessage(msg);
            return false;
        }
        return true;
    }

    //
    //  制御テーブル排他開放
    //
    public static synchronized boolean CZPutControlExclusion(String ro){

        initCheck();
        int rc = cz_tb_px.CZPutControlExclusion(ro);
        log("CZSystem","CZPutControlExclusion rc:" + rc + " [" + ro +"]");
        if(0 != rc) return false;
        return true;
    }


    //
    //  制御テーブル設定(T6)
    //
    public static synchronized boolean CZControlT6TableExchange(
                                    int flg,
                                    String op,
                                    CZParamT6Table[] param){
        initCheck();
        String ro = getRoName();
        // 送信
        int rc = cz_tb_px.CZControlT6TableExchange(flg,op,ro,param);
        log("CZSystem","CZControlT6TableExchange rc:" + rc + " [" + op +"]");
        if (0 != rc ) {
            CZSystemSysMsg msg = new CZSystemSysMsg();
            msg.message = CZSystem.getDateTime() +   
                        " 制御テーブル設定失敗 [" + rc + "]";
            sysMessage(msg);
            return false;
        } else {
            CZSystemSysMsg msg = new CZSystemSysMsg();
            msg.message = CZSystem.getDateTime() +   
                        " 制御テーブル設定完了";
            sysMessage(msg);
            return true;
        }
    }

    //
    //  制御テーブル設定
    //
    public static synchronized boolean CZControlTableExchange(String op,
                                  int flg,int group,
                                  int recipe,int table,
                                  float[] left, 
                                  float[] right){
        initCheck();
        String ro = getRoName();
        // 送信
        int rc = cz_tb_px.CZControlTableExchange(op,ro,flg,group,recipe,table,left,right);
        log("CZSystem","CZControlTableExchange rc:" + rc + " [" + op +"]");
        if(0 != rc) {
            CZSystemSysMsg msg = new CZSystemSysMsg();
            msg.message = CZSystem.getDateTime() +   
                        " 制御テーブル設定失敗 [" + rc + "]";
            sysMessage(msg);
            return false;
        } else {
            CZSystemSysMsg msg = new CZSystemSysMsg();
            msg.message = CZSystem.getDateTime() +   
                        " 制御テーブル設定完了";
            sysMessage(msg);
            return true;
        }
    }

    //
    //  制御テーブル項目設定
    //
    public static synchronized boolean  CZControlDefineExchange(String op,  
                                    int grp,
                                    int tbl,
                                    CZParamControlDefine param){
        initCheck();
        // 送信
        int rc = cz_tb_px.CZControlDefineExchange(op,grp,tbl,param);
        log("CZSystem","CZControlDefineExchange rc:" + rc + " [" + op +"]");
        if(0 != rc) return false;
        return true;
    }


    //
    //  制御テーブルタイトル変更
    //
    public static synchronized boolean CZControlTitleExchange(String op,int grp,
                                  int rec,String title){

        initCheck();
        String ro = getRoName();
        // 送信
        int rc = cz_tb_px.CZControlTitleExchange(op,ro,grp,rec,title);
        log("CZSystem","CZControlTitleExchange rc:" + rc + " [" + op +"]");
        if(0 != rc) return false;
        return true;
    }

//@@

    //
    //  制御テーブルT6項目設定
    //
    public static synchronized boolean  CZControlT6DefineExchange(
                                    String op,  
                                    int grp,
                                    int lag,
                                    int mid,
                                    int kNo,
                                    CZParamControlT6Define param){
        initCheck();
        String ro = getRoName();
        // 送信
        int rc = cz_tb_px.CZControlT6Exchange(op,ro,grp,lag,mid,kNo,param);
        log("CZSystem","CZControlT6DefineExchange rc:" + rc + " [" + op +"]");
        if(0 != rc) return false;
        return true;
    }

/*@@ 制御テーブルT6大項目変更
    //
    //  制御テーブルT6大項目変更
    //
    public static synchronized boolean CZControlT6LagExchange(String op,int grp,
                                  int rec,int lag,String lagName){

        initCheck();
        String ro = getRoName();
        // 送信
        int rc = cz_tb_px.CZControlT6LagExchange(op,ro,grp,rec,lag,lagName);
        log("CZSystem","CZControlT6LagExchange rc:" + rc + " [" + op +"]");
        if(0 != rc) return false;
        return true;
    }
@@*/

/*@@    制御テーブルT6中項目変更
    //
    //  制御テーブルT6中項目変更
    //
    public static synchronized boolean CZControlT6MidExchange(String op,int grp,
                                  int rec,int lag,int mid,String midName){

        initCheck();
        String ro = getRoName();
        // 送信
        int rc = cz_tb_px.CZControlT6MidExchange(op,ro,grp,rec,lag,mid,midName);
        log("CZSystem","CZControlT6MidExchange rc:" + rc + " [" + op +"]");
        if(0 != rc) return false;
        return true;
    }
@@*/

//@@

    //
    //  制御テーブル全コピー
    //
    public static synchronized boolean CZControlCopyRo(String op, String dst_ro){

        initCheck();
        String ro = getRoName();
        // 送信
        int rc = -1;
        rc = cz_tb_px.CZControlCopyRo(op,ro,dst_ro);
        log("CZSystem","CZControlCopyRo rc:" + rc + " [" + op +"] srcRo[" + ro + "] dstRo[" + dst_ro + "]");
        if(0 != rc) return false;
        return true;
    }

    //
    //  制御テーブルグループコピー
    //
    public static synchronized boolean CZControlCopyGroup(String op, String dst_ro, int gno){

        initCheck();
        String ro = getRoName();
        // 送信
        int rc = -1;
        rc = cz_tb_px.CZControlCopyGroup(op,ro,dst_ro,gno);
        log("CZSystem","CZControlCopyGroup rc:" + rc +
                             " [" + op +"] srcRo[" + ro +
                             "] dstRo[" + dst_ro + "] gno[" + gno + "]");
        if(0 != rc) return false;
        return true;
    }

    //
    //  制御テーブルレシピコピー
    //
    public static synchronized boolean CZControlCopyRecipe(String op,   
                                   String dst_ro, int gno,  
                                   int rno,int dst_rno){
        initCheck();
        String ro = getRoName();

        // 送信
        int rc = cz_tb_px.CZControlCopyRecipe(op,ro,dst_ro,gno,rno,dst_rno);
        log("CZSystem","CZControlCopyRecipe rc:" + rc + " [" + op +"]");
        if(0 != rc) return false;
        return true;
    }

    //
    //  制御テーブルテーブルコピー
    //
    public static synchronized boolean CZControlCopyTable(String op,    
                                   String dst_ro, int gno,  
                                   int rno, int dst_rno, int tno){
        initCheck();
        String ro = getRoName();

        // 送信
        int rc = cz_tb_px.CZControlCopyTable(op,ro,dst_ro,gno,rno,dst_rno,tno);
        log("CZSystem","CZControlCopyTable rc:" + rc + " [" + op +"]");

        if(0 != rc) return false;

        return true;
    }


    //
    //  制御テーブル大項目コピー
    //
    public static synchronized boolean CZControlCopyLagName(String op,
                         String dst_ro, int gNo, int src_rNo, int dst_rNo, int lNo){

        initCheck();
        String ro = getRoName();

        // 送信
//@@        System.out.println("op="+op+": ro="+ro+": dst_ro="+dst_ro+": gNo="+gNo+": src_rNo="+src_rNo+": dst_rNo="+dst_rNo+": lNo="+lNo);
        int rc = cz_tb_px.CZControlCopyT6LagName(op, ro, dst_ro, gNo, src_rNo, dst_rNo, lNo);
        log("CZSystem","CZControlCopyT6LagName rc:" + rc + " [" + op +"]");

        if(0 != rc) return false;
        return true;
    }

    //
    //  制御テーブル中項目コピー
    //
    public static synchronized boolean CZControlCopyMidName(String op,
                         String dst_ro, int gNo, int src_rNo, int dst_rNo, int src_lNo, int dst_lNo, int mNo){

        initCheck();
        String ro = getRoName();

        // 送信
        int rc = cz_tb_px.CZControlCopyT6MidName(op,ro,dst_ro,gNo,src_rNo,dst_rNo,src_lNo,dst_lNo,mNo);
        log("CZSystem","CZControlCopyT6MidName rc:" + rc + " [" + op +"]");

        if(0 != rc) return false;
        return true;
    }
    //
    //  制御テーブル項目コピー
    //
    public static synchronized boolean CZControlCopyT6Name(String op,
                         String dst_ro, int gNo, int src_rNo, int dst_rNo,
                         int src_lNo, int dst_lNo, int src_mNo, int dst_mNo, int val){

        initCheck();
        String ro = getRoName();

        // 送信
        int rc = cz_tb_px.CZControlCopyT6Name(
            op,ro,dst_ro,gNo,src_rNo,dst_rNo,src_lNo,dst_lNo,src_mNo,dst_mNo,val); //@@@
        log("CZSystem","CZControlCopyT6Name rc:" + rc + " [" + op +"]");

        if(0 != rc) return false;
        return true;
    }

    //////////////////////////////////////////////////////////////////////////////
    //
    //  炉名称読み込み
    //
	@SuppressWarnings("unchecked")
    private static int roRead(){

        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        sql = new String("SELECT * FROM "+ "mst." +"m_ro ORDER BY roban");  

        log("CZSystem roRead","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem roRead","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return -1;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem roRead","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return -1;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem roRead","ERROR: createStatement or database");
            return -1;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                ro_name_list.addElement(rs.getString(1));
                ro_host_list.addElement(rs.getString(2));
                ro_camera_list.addElement(rs.getString(3));
                ro_ver_list.addElement(rs.getString(4));
            } // for end
            log("CZSystem roRead","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem roRead","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem roRead","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);

        return i;
    }   


    //
    //  運転画面設定読み込み
    //
    private static CZSystemPVNamePMM untenRead(int _proc){

        initCheck();

        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int proc = _proc;

        if(-1 == _proc) proc = 0;


        log("CZSystem untenRead","proc = " + proc);

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
            catch (Throwable e) {
            log("CZSystem untenRead","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem untenRead","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
            sql = new String("SELECT * FROM "+ getDBName() + "." +"r_sg_unten WHERE p_no = " + proc );  
            log("CZSystem untenRead","SQL["+sql+"]");
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem untenRead","ERROR: createStatement or database");
            return null;
        }

        CZSystemPVNamePMM ret = new CZSystemPVNamePMM();

        try{
            rs = sqlstmt.executeQuery(sql) ;
            if(rs.next()) {

                int i = 1;
                ret.p_no    = rs.getInt(i);

                for(int j = 0 ; j < ret.SIZE ; j++){
                    i++;
                    ret.item[j] = rs.getInt(i);
                }

                for(int j = 0 ; j < ret.SIZE ; j++){
                    i++;
                    CZSystemPVName pv = getPVName(ret.item[j] - 1);
                    ret.min[j] = (float)(rs.getInt(i) / Math.pow(10,pv.keta));
                }

                for(int j = 0 ; j < ret.SIZE ; j++){
                    i++;
                    CZSystemPVName pv = getPVName(ret.item[j] - 1);
                    ret.max[j] = (float)(rs.getInt(i) / Math.pow(10,pv.keta));
                }

                i++;
                ret.x_shubetu   = rs.getInt(i);

                i++;
                ret.x_time  = rs.getInt(i);

                i++;
                ret.x_width = rs.getInt(i);
                log("CZSystem untenRead","Select Data OK !!");
            } else {
                log("CZSystem untenRead","ERROR: Select Data Nothing !!");
                ret = null;
            }
        }
        catch( SQLException e ){
            log("CZSystem untenRead","ERROR: Select failed.");
            ret = null;
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem untenRead","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        return ret;
    }


    //
    //  ＰＶ名称読み込み
    //
	@SuppressWarnings("unchecked")
    private static int pvNameRead(){

        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        sql = new String("SELECT * FROM " + "mst." + "m_pv_name ORDER BY k_no");    

        log("CZSystem pvNameRead","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem pvNameRead","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return -1;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem pvNameRead","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return -1;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem pvNameRead","ERROR: createStatement or database");
            return -1;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){

                CZSystemPVName name = new CZSystemPVName();
                name.k_no   = rs.getInt(1);
                name.k_name = rs.getString(2);
                name.keta   = rs.getInt(3);
                name.unit   = rs.getString(4);
                name.n_min  = rs.getInt(5);
                name.n_max  = rs.getInt(6);
                name.j_name = rs.getString(7);
                pv_name_list.addElement(name);
            } // for end
            
            log("CZSystem pvNameRead","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem pvNameRead","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem pvNameRead","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        return i;
    }   


    //
    //  エラーメッセージ読み込み
    //
	@SuppressWarnings("unchecked")
    private static int errorMessageRead(){

        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        sql = new String("SELECT * FROM " + "mst." + "m_errmsg_mast ORDER BY e_no");    

        log("CZSystem errorMessageRead","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem errorMessageRead","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return -1;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem errorMessageRead","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return -1;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem errorMessageRead","ERROR: createStatement or database");
            return -1;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){

                CZSystemErrMsg name = new CZSystemErrMsg();

                name.e_no   = rs.getInt(1);
                name.message    = rs.getString(2);
                name.youhi  = rs.getInt(3);

                error_message_list.addElement(name);
            } // for end
            
            log("CZSystem errorMessageRead","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem errorMessageRead","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem errorMessageRead","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        return i;
    }   


    //
    //  エラーメッセージ読み込み２
    //　2006.04.13　y.k 追加
	@SuppressWarnings("unchecked")
    public static Vector errorMessageRead2(String db_name){

        Vector  ret         = new Vector(5000);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        initCheck();


        int i = 0;

        sql = new String("SELECT * FROM " + db_name + "." + "r_errmsg_mast ORDER BY e_no");    

        log("CZSystem errorMessageRead2","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem errorMessageRead2","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem errorMessageRead2","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem errorMessageRead2","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){

                CZSystemErrMsg name = new CZSystemErrMsg();

                name.e_no   = rs.getInt(1);
                name.message    = rs.getString(2);
                name.youhi  = rs.getInt(3);

                ret.addElement(name);
            } // for end
            
            log("CZSystem errorMessageRead2","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem errorMessageRead2","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem errorMessageRead2","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        return ret;
    }


    //
    //  操業定数：大項目読み込み
    //
	@SuppressWarnings("unchecked")
    private static int opTblLagNameRead(){

        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        sql = new String("SELECT * FROM " + "mst." + "m_st_name1 ORDER BY k_no");   

        log("CZSystem opTblLagNameRead","SQL["+sql+"]");
        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem opTblLagNameRead","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return -1;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem opTblLagNameRead","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return -1;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem opTblLagNameRead","ERROR: createStatement or database");
            return -1;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;

            for(i = 0 ; rs.next() ; i++){
                CZSystemOpTbLag name = new CZSystemOpTbLag();
                name.k_no   = rs.getInt(1);
                name.k_name = rs.getString(2);
                op_tb_lag_name_list.addElement(name);
            } // for end

            log("CZSystem opTblLagNameRead","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem opTblLagNameRead","ERROR: Select failed.");
        }
        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem opTblLagNameRead","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        return i;
    }

    //
    //  操業定数：中項目読み込み
    //
	@SuppressWarnings("unchecked")
    private static int opTblMidNameRead(){

        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        sql = new String("SELECT * FROM " + "mst." + "m_st_name2 ORDER BY k_no1,k_no2");    

        log("CZSystem opTblMidNameRead","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem opTblMidNameRead","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return -1;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem opTblMidNameRead","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return -1;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem opTblMidNameRead","ERROR: createStatement or database");
            return -1;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){

                CZSystemOpTbMid name = new CZSystemOpTbMid();
                name.k_no1  = rs.getInt(1);
                name.k_no2  = rs.getInt(2);
                name.k_name = rs.getString(3);
                op_tb_mid_name_list.addElement(name);
            } // for end

            log("CZSystem opTblMidNameRead","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem opTblMidNameRead","ERROR: Select failed.");
        }
        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem opTblMidNameRead","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        return i;
    }   

    //
    //  操業定数：項目読み込み
    //
	@SuppressWarnings("unchecked")
    private static int opTblSmlNameRead(){

        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        sql = new String("SELECT * FROM "+ "mst." + "m_st_mast ORDER BY k_no1,k_no2,k_no"); 

        log("CZSystem opTblSmlNameRead","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem opTblSmlNameRead","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return -1;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem opTblSmlNameRead","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return -1;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem opTblSmlNameRead","ERROR: createStatement or database");
            return -1;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){

                CZSystemOpTbSml name = new CZSystemOpTbSml();
                name.k_no1  = rs.getInt(1);
                name.k_no2  = rs.getInt(2);
                name.k_no   = rs.getInt(3);
                name.k_name = rs.getString(4);
                name.t_name = rs.getString(5);
                name.n_min  = rs.getFloat(6);
                name.n_max  = rs.getFloat(7);
                name.keta   = rs.getInt(8);

                op_tb_sml_name_list.addElement(name);
            } // for end

            log("CZSystem opTblSmlNameRead","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem opTblSmlNameRead","ERROR: Select failed.");
        }
        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem opTblSmlNameRead","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        return i;
    }


    //
    //  操業定数：項目読み込み
    //  2006.06.06 y.k
	@SuppressWarnings("unchecked")
    public static Vector opTblAllNameRead(){

        Vector  ret         = new Vector(1000);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        sql = new String("SELECT ms1.k_name, ms2.k_name, ms.k_no1, ms.k_no2, ms.k_no, ms.k_name," +
                         "ms.t_name, ms.n_min, ms.n_max, ms.keta " +
                         "from mst.m_st_mast ms, mst.m_st_name1 ms1, mst.m_st_name2 ms2 " + 
						 "where ms.k_no1 = ms1.k_no and ms.k_no1 = ms2.k_no1 and " +
						 " ms.k_no2 = ms2.k_no2 ORDER BY ms.k_no1,ms.k_no2,ms.k_no " );

        log("CZSystem opTblAllNameRead","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem opTblAllNameRead","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem opTblAllNameRead","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem opTblAllNameRead","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){

                CZSystemOpTbAll name = new CZSystemOpTbAll();
                name.k_name1 = rs.getString(1);
                name.k_name2 = rs.getString(2);
                name.k_no1  = rs.getInt(3);
                name.k_no2  = rs.getInt(4);
                name.k_no   = rs.getInt(5);
                name.k_name = rs.getString(6);
                name.t_name = rs.getString(7);
                name.n_min  = rs.getFloat(8);
                name.n_max  = rs.getFloat(9);
                name.keta   = rs.getInt(10);

                ret.addElement(name);
            } // for end

            log("CZSystem opTblAllNameRead","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem opTblAllNameRead","ERROR: Select failed.");
        }
        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem opTblSmlNameRead","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }

    //
    //  制御テーブル：マスタ読み込み
    //  2006.06.13 Y.K
	@SuppressWarnings("unchecked")
    public static Vector ctTblAllNameRead(int tblGno){

        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;
        Vector  ret         = new Vector(1000);

        int i = 0;

        sql = new String("SELECT * FROM " + "mst." + "m_ct_mast where g_no = " + tblGno + " ORDER BY g_no,t_no");   

        log("CZSystem ctTblNameRead","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem ctTblNameRead","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem ctTblNameRead","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem ctTblNameRead","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;

            for(i = 0 ; rs.next() ; i++){

                CZSystemCtName name = new CZSystemCtName();
                name.g_no   = rs.getInt(1);
                name.t_no   = rs.getInt(2);
                name.t_name = rs.getString(3);
                name.l_name = rs.getString(4);
                name.l_unit = rs.getString(5);
                name.l_min  = rs.getFloat(6);
                name.l_max  = rs.getFloat(7);
                name.l_keta = rs.getInt(8);
                name.r_name = rs.getString(9);
                name.r_unit = rs.getString(10);
                name.r_min  = rs.getFloat(11);
                name.r_max  = rs.getFloat(12);
                name.r_keta = rs.getInt(13);
                name.k_sort = rs.getInt(14);
                name.pv_no  = rs.getInt(15);
                ret.addElement(name);
            } // for end
            log("CZSystem ctTblNameRead","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem ctTblNameRead","ERROR: Select failed.");
        }
        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem ctTblNameRead","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    } 

    //
    //  制御テーブル：マスタ読み込み(t6)
    //  2006.06.13 Y.K
	@SuppressWarnings("unchecked")
    public static Vector ctT6AllNameRead(){

        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;
        Vector  ret         = new Vector(1000);

        int i = 0;

        sql = new String("SELECT ctN1.k_name1, ctN2.k_name2, ct6.g_no, ct6.k_no1, ct6.k_no2, ct6.k_no, " + 
						 "ct6.k_name, ct6.k_unit, ct6.k_min, ct6.k_max, ct6.k_keta " + 
						 "FROM mst.m_ct6_mast ct6, mst.m_ct6_name1 ctN1, mst.m_ct6_name2 ctN2 " + 
						 "where ct6.g_no = ctN1.g_no and ct6.k_no1 = ctN1.k_no1 and " + 
						 "ct6.g_no = ctN2.g_no and ct6.k_no1 = ctN2.k_no1 and ct6.k_no2 = ctN2.k_no2 " + 
						 "ORDER BY ct6.g_no, ct6.k_no1, ct6.k_no2, ct6.k_no ");

        log("CZSystem ctT6AllNameRead","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem ctT6AllNameRead","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem ctT6AllNameRead","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem ctT6AllNameRead","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;

            for(i = 0 ; rs.next() ; i++){

                CZSystemCtT6AllName name = new CZSystemCtT6AllName();
                name.k_name1   = rs.getString(1);
                name.k_name2   = rs.getString(2);
                name.g_no   = rs.getInt(3);
                name.k_no1  = rs.getInt(4);
                name.k_no2  = rs.getInt(5);
                name.k_no   = rs.getInt(6);
                name.k_name = rs.getString(7);
                name.k_unit = rs.getString(8);
                name.k_min  = rs.getFloat(9);
                name.k_max  = rs.getFloat(10);
                name.k_keta = rs.getInt(11);
                ret.addElement(name);
            } // for end
            log("CZSystem ctT6AllNameRead","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem ctT6AllNameRead","ERROR: Select failed.");
        }
        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem ctT6AllNameRead","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }   


    //
    //  制御テーブル：マスタ読み込み
    //
	@SuppressWarnings("unchecked")
    private static int ctTblNameRead(){

        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        sql = new String("SELECT * FROM " + "mst." + "m_ct_mast ORDER BY g_no,t_no");   

        log("CZSystem ctTblNameRead","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem ctTblNameRead","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return -1;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem ctTblNameRead","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return -1;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem ctTblNameRead","ERROR: createStatement or database");
            return -1;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;

            for(i = 0 ; rs.next() ; i++){

                CZSystemCtName name = new CZSystemCtName();
                name.g_no   = rs.getInt(1);
                name.t_no   = rs.getInt(2);
                name.t_name = rs.getString(3);
                name.l_name = rs.getString(4);
                name.l_unit = rs.getString(5);
                name.l_min  = rs.getFloat(6);
                name.l_max  = rs.getFloat(7);
                name.l_keta = rs.getInt(8);
                name.r_name = rs.getString(9);
                name.r_unit = rs.getString(10);
                name.r_min  = rs.getFloat(11);
                name.r_max  = rs.getFloat(12);
                name.r_keta = rs.getInt(13);
                name.k_sort = rs.getInt(14);
                name.pv_no  = rs.getInt(15);
                ct_tb_name_list.addElement(name);
            } // for end
            log("CZSystem ctTblNameRead","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem ctTblNameRead","ERROR: Select failed.");
        }
        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem ctTblNameRead","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        return i;
    } 
//@@

    //
    //  制御テーブル：マスタ読み込み(t6)
    //
	@SuppressWarnings("unchecked")
    private static int ctT6NameRead(){

        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        sql = new String("SELECT * FROM " + "mst." + "m_ct6_mast ORDER BY g_no,k_no1, k_no2,k_no");  

        log("CZSystem ctT6NameRead","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem ctT6NameRead","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return -1;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem ctT6NameRead","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return -1;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem ctT6NameRead","ERROR: createStatement or database");
            return -1;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;

            for(i = 0 ; rs.next() ; i++){

                CZSystemCtT6Name name = new CZSystemCtT6Name();
                name.g_no   = rs.getInt(1);
                name.k_no1  = rs.getInt(2);
                name.k_no2  = rs.getInt(3);
                name.k_no   = rs.getInt(4);
                name.k_name = rs.getString(5);
                name.k_unit = rs.getString(6);
                name.k_min  = rs.getFloat(7);
                name.k_max  = rs.getFloat(8);
                name.k_keta = rs.getInt(9);
                name.k_sort = rs.getInt(10);
                name.pv_no  = rs.getInt(11);
                ctT6NameList_.addElement(name);
            } // for end
            log("CZSystem ctT6NameRead","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem ctT6NameRead","ERROR: Select failed.");
        }
        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem ctT6NameRead","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        return i;
    }   


    //
    //  制御テーブル：マスタ読み込み(大項目)
    //
	@SuppressWarnings("unchecked")
    private static int ctT6LagNameRead(){

        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        sql = new String("SELECT * FROM " + "mst." + "m_ct6_name1 ORDER BY g_no,k_no1");    

        log("CZSystem ctT6LagNameRead","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem ctT6LagNameRead","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return -1;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem ctT6LagNameRead","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return -1;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem ctT6LagNameRead","ERROR: createStatement or database");
            return -1;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;

            for(i = 0 ; rs.next() ; i++){

                CZSystemCtT6LagName name = new CZSystemCtT6LagName();
                name.g_no    = rs.getInt(1);
                name.r_no    = rs.getInt(2);
                name.k_no1   = rs.getInt(3);
                name.k_name1 = rs.getString(4);
                ctT6LagNameList_.addElement(name);
            } // for end
            log("CZSystem ctT6LagNameRead","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem ctT6LagNameRead","ERROR: Select failed.");
        }
        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem ctT6NameRead","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        return i;
    }   


    //
    //  制御テーブル：マスタ読み込み(中項目)
    //
	@SuppressWarnings("unchecked")
    private static int ctT6MidNameRead(){

        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

        int i = 0;

        sql = new String("SELECT * FROM " + "mst." + "m_ct6_name2 ORDER BY g_no, r_no, k_no1, k_no2");  

        log("CZSystem ctT6MidNameRead","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem ctT6MidNameRead","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return -1;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem ctT6MidNameRead","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return -1;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem ctT6MidNameRead","ERROR: createStatement or database");
            return -1;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;

            for(i = 0 ; rs.next() ; i++){

                CZSystemCtT6MidName name = new CZSystemCtT6MidName();
                name.g_no   = rs.getInt(1);
                name.r_no  = rs.getInt(2);
                name.k_no1  = rs.getInt(3);
                name.k_no2  = rs.getInt(4);
                name.k_name2 = rs.getString(5);
                ctT6MidNameList_.addElement(name);
            } // for end
            log("CZSystem ctT6MidNameRead","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem ctT6MidNameRead","ERROR: Select failed.");
        }
        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem ctT6MidNameRead","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        return i;
    } 

//@@
    //
    //  データベースクローズ
    //
    private static boolean closeConnect(Connection c){

        try{
            c.close();
        }
        catch (SQLException e){
            log("CZSystem closeConnect","ERROR: Close Connection");
            return false;
        }
        return true;
    }


    //////////////////////////////////////////////////////////////////////////////////////////////
    //
    //  ＲＡＩＤ状態取得
    //
    public static CZRaidStatus CZRaidGetStatus(int mode, int flg){
//        log("CZSystem CZRaidGetStatus","Status Mode[" + mode + "] Flag[" + flg + "]");

        if(null == cz_sv_px) return null;

        if(0 == mode){
            switch(flg){
                case 0 : raid1_stat = cz_sv_px.CZRaidGetStatus(flg);
                     return raid1_stat;

                case 1 : raid5_stat = cz_sv_px.CZRaidGetStatus(flg);
                     return raid5_stat;
            }
        }
        else if(1 == mode){
            switch(flg){
                case 0 : return raid1_stat;
                case 1 : return raid5_stat;
            }
        }
        return null;
    }

    //
    //  ＭＯ状態取得
    //
    public static CZMoList[] CZMoGetlist(int value){

        initCheck();
        String dir = null;

        switch(value){
            case 1: dir = MO_1_DIR;
            break;

            case 2: dir = MO_2_DIR;
            break;

            default : return null;
        }

        log("CZSystem CZMoGetlist","VAL[" + value + "] DIR[" + dir + "]");

        CZMoList[] list = cz_sv_px.CZMoGetlist(dir);

        return list;
    }

    //
    //  ＭＯマウント
    //
    public static boolean CZMoMount(int value){

        initCheck();
        String dir = null;

        switch(value){
            case 1: dir = MO_1_DIR;
            break;

            case 2: dir = MO_2_DIR;
            break;

            default : return false;
        }

        int rc  = cz_sv_px.CZMoMount(dir);

        log("CZSystem CZMoMount","rc:" + rc + " VAL[" + value + "] DIR[" + dir + "]");

        if(0 != rc) return false;
        return true;
    }

    //
    //  ＭＯアンマウント
    //
    public static boolean CZMoUmount(int value){

        initCheck();
        String dir = null;

        switch(value){
            case 1: dir = MO_1_DIR;
            break;

            case 2: dir = MO_2_DIR;
            break;

            default : return false;
        }

        int rc  = cz_sv_px.CZMoUmount(dir);

        log("CZSystem CZMoUmount","rc:" + rc + " VAL[" + value + "] DIR[" + dir + "]");

        if(0 != rc) return false;
        return true;
    }

    //
    //  ＭＯフォーマット
    //
    public static boolean CZMoFormat(int value){

        initCheck();
        String dir = null;

        switch(value){
            case 1: dir = MO_1_DIR;
            break;

            case 2: dir = MO_2_DIR;
            break;

            default : return false;
        }

        int rc  = cz_sv_px.CZMoFormat(dir);

        log("CZSystem CZMoFormat","rc:" + rc + " VAL[" + value + "] DIR[" + dir + "]");

        if(0 != rc) return false;
        return true;
    }

    //
    //  ＭＯ→ＤＢ展開
    //
    public static boolean CZMoExtract(int value,String roName, String batch){

        initCheck();
        String dir = null;

        switch(value){
            case 1: dir = MO_1_DIR;
            break;

            case 2: dir = MO_2_DIR;
            break;

            default : return false;
        }

        if(null == roName) return false;
        if(1 >  roName.length()) return false;

        if(null == batch) return false;
        if(1 >  batch.length()) return false;

        int rc  = cz_sv_px.CZMoExtract(dir,roName,batch);

        log("CZSystem CZMoExtract","rc:" + rc + " VAL[" + value + "] DIR[" + dir + "] Batch[" + batch + "]");

        if(0 != rc) return false;
        return true;
    }


    ////////////////////////////////////////////////////////////////////
    //
    //  炉状況
    //
    public static CZNativeRoState[] CZNativeRoStateGet(){

        initCheck();
        log("CZSystem","CZNativeRoStateGet");
        CZNativeRoState[] list = cz_gd_px.CZNativeRoStateGet();
        return list;
    }

    ////////////////////////////////////////////////////////////////////
    //
    //  炉状況（集中監視）2006/09/29
    //
    public static CZNativeMRoState[] CZNativeMRoStateGet(){

        initCheck();
        CZNativeMRoState[] list = cz_gd_px.CZNativeMRoStateGet();
        return list;
    }

    ////////////////////////////////////////////////////////////////////
    //
    //  炉毎排他状況（制御テーブル）
    //
    public static CZNativeCTState[] CZNativeCTStateGet(){

        initCheck();
        log("CZSystem","CZNativeCTStateGet");
        CZNativeCTState[] list = cz_gd_px.CZNativeCTStateGet();
        return list;
    }

    ////////////////////////////////////////////////////////////////////
    //
    //  炉毎排他状況（操業定数）
    //
    public static CZNativeSTState[] CZNativeSTStateGet(){

        initCheck();
        log("CZSystem","CZNativeCTStateGet");
        CZNativeSTState[] list = cz_gd_px.CZNativeSTStateGet();
        return list;
    }

    ////////////////////////////////////////////////////////////////////
    //
    //  炉毎引き上げ条件
    //
    public static CZNativeRoHikiage[] CZNativeRoHikiageGet(){

        initCheck();
        log("CZSystem","CZNativeRoHikiage");
        CZNativeRoHikiage[] list = cz_gd_px.CZNativeRoHikiageGet();
        return list;
    }

    ////////////////////////////////////////////////////////////////////
    //
    //  ＲＥＡＬ監視状況
    //
    public static CZRealNativeWatchItem[] CZNativeRealStateGet(String sRo){

        initCheck();
        CZRealNativeWatchItem[] list = cz_rl_px.CZRealNativeGetWatchItem(sRo);
        return list;
    }

    ////////////////////////////////////////////////////////////////////
    //
    //  プロセス名の取り出し
    //
    public static String getProcName(int no){

        String ret = null;
        try{
            ret = CZSystemDefine.PROC_NAME[no];
        }
        catch (Exception e){
            exit(0,"getProcName() PROC No[" + no + "]");
        }
        return ret;
    }

    ////////////////////////////////////////////////////////////////////
    //
    //  プロセス名の取り出し
    //
    public static String getProcName2(int no){

        String ret = null;
        try{
            ret = CZSystemDefine.PROC_NAME2[no];
        }
        catch (Exception e){
            exit(0,"getProcName() PROC No[" + no + "]");
        }
        return ret;
    }

    ////////////////////////////////////////////////////////////////////
    //
    //  プロセス名の取り出し(輝度変化チェック用)
    //
    public static String getProcName3(int no){

        String ret = null;
        try{
            ret = CZSystemDefine.PROC_NAME3[no];
        }
        catch (Exception e){
            exit(0,"getProcName() PROC No[" + no + "]");
        }
        return ret;
    }

    ////////////////////////////////////////////////////////////////////
    //
    //  トライ毎の引上げ条件取り出し
    //
	@SuppressWarnings("unchecked")
    public static Vector getHikiageTemp(String db_name, String batch, String time){
    
        initCheck();
        Vector  ret         = new Vector(1000);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;

		StringBuffer a = new StringBuffer();
		a.append(time);
		
		a.delete(19,21);
		
		String stime = a.toString();
		
        int i = 0;

        String view = new String("r_hikiage_temp");
        sql = new String("SELECT * FROM " + db_name + "." + view + 
                         " WHERE batch = '" + batch + "' and t_time = to_date('" +
                         stime + "' , 'YYYY-MM-DD HH24:MI:SS')");

        log("CZSystem getHikiageTemp","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getHikiageTemp","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getHikiageTemp","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getHikiageTemp","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZSystemBtTemp bt   = new CZSystemBtTemp();
                bt.batch        = rs.getString(1);  //バッチ番号
                bt.t_time       = rs.getString(2);  //PG-ID
                bt.renban       = rs.getInt(3);     //登録日時
                bt.pgid         = rs.getString(4);  //連番
                bt.hinshu       = rs.getString(5);  //品種
                bt.houi         = rs.getString(6);  //方位
                bt.h_type       = rs.getString(7);  //タイプ
                bt.hiteikou     = rs.getString(8);  //比抵抗
                bt.sanso        = rs.getString(9);  //酸素
                bt.gap          = rs.getString(10); //GAP
                bt.rutubo_kei   = rs.getInt(11);    //ルツボ径
                bt.chokkei      = rs.getInt(12);    //直径
                bt.hikiage_cho  = rs.getInt(13);    //引上長
                bt.top_ar       = rs.getInt(14);    //トップアルゴン
                bt.pull_ar      = rs.getInt(15);    //プルアルゴン
                bt.i_sikomi     = rs.getInt(16);    //仕込量
                bt.t_sikomi     = rs.getInt(17);    //追加仕込量
                bt.zaneki       = rs.getInt(18);    //残液量
                bt.no_youkai    = rs.getInt(19);    //T1(溶解)
                bt.no_hikiage   = rs.getInt(20);    //T2(引上)
                bt.no_kaiten    = rs.getInt(21);    //T3(回転)
                bt.no_toridasi  = rs.getInt(22);    //T4(取出)
                bt.no_aturyoku  = rs.getInt(23);    //T5(圧力)
                bt.no_teisu     = rs.getInt(24);    //T6(定数) @@
                bt.pno_start    = rs.getInt(25);    //スタートプロセス
                bt.p_kaisi      = rs.getInt(26);    //開始
                ret.addElement(bt);
            } // for end
            log("CZSystem getHikiageTemp","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getHikiageTemp","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getHikiageTemp","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }
// add start 2008.10.08
    ////////////////////////////////////////////////////////////////////
    //
    //  編集履歴の取り出し(操業定数)
    //
	@SuppressWarnings("unchecked")
    public static Vector getModifyHistoryC(String date1, String date2, String roName){
    
        initCheck();
        Vector  ret         = new Vector(1000);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;
		
        int i = 0;

        sql = new String("SELECT R.S_TIME,R.OP_NAME,R.BATCH,M.MESSAGE,R.KEY1,R.KEY2,R.KEY3,R.KEY4,R.KEY5 FROM " +
                          roName + ".R_MODIFY R, MST.M_TBLMSG_MAST M where (R.ID_TABLE = M.S_NO) and  (R.ID_TABLE = 1) and " +
                          "(R.S_TIME >=  TO_DATE('" + date1 + " 00:00:00','YYYY/MM/DD HH24:MI:SS') and R.S_TIME  <= TO_DATE('" + date2 + " 23:59:59','YYYY/MM/DD HH24:MI:SS')) ORDER BY R.S_TIME,R.KEY3"); 

        log("CZSystem getModifyHistoryC","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getModifyHistoryC","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getModifyHistoryC","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getModifyHistoryC","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZSystemModifyHistoryC cns   = new CZSystemModifyHistoryC();
                cns.s_time       = rs.getString(1);    //変更日時
                cns.op_name      = rs.getString(2);    //変更者
                cns.batch        = rs.getString(3);    //Bt
                cns.message      = rs.getString(4);    //変更内容
                cns.key1         = rs.getInt(5);       //大項目
                cns.key2         = rs.getInt(6);       //中項目
                cns.key3         = rs.getInt(7);       //項目No
                cns.key4         = rs.getFloat(8);     //変更前
                cns.key5         = rs.getFloat(9);     //変更前
                ret.addElement(cns);
            } // for end
            log("CZSystem getModifyHistoryC","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getModifyHistoryC","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getModifyHistoryC","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }

    ////////////////////////////////////////////////////////////////////
    //
    //  編集履歴の取り出し(T6)
    //
	@SuppressWarnings("unchecked")
    public static Vector getModifyHistoryT6(String date1, String date2, String roName){
    
        initCheck();
        Vector  ret         = new Vector(1000);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;
		
        int i = 0;

        sql = new String("SELECT R.S_TIME,R.OP_NAME,R.BATCH,M.MESSAGE,R.KEY1,R.KEY2,R.KEY3,R.KEY4,R.KEY5,R.KEY6,R.KEY7 FROM " +
                          roName + ".R_MODIFY R, MST.M_TBLMSG_MAST M where (R.ID_TABLE = M.S_NO) and  (R.ID_TABLE = 2 or R.ID_TABLE = 3) and (R.KEY1 = 6) and" +
                          "(R.S_TIME >=  TO_DATE('" + date1 + " 00:00:00','YYYY/MM/DD HH24:MI:SS') and R.S_TIME  <= TO_DATE('" + date2 + " 23:59:59','YYYY/MM/DD HH24:MI:SS')) ORDER BY R.S_TIME,R.KEY5"); 

        log("CZSystem getModifyHistoryT6","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getModifyHistoryT6","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getModifyHistoryT6","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getModifyHistoryT6","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZSystemModifyHistoryT6 t6   = new CZSystemModifyHistoryT6();
                t6.s_time       = rs.getString(1);     //変更日時
                t6.op_name      = rs.getString(2);     //変更者
                t6.batch        = rs.getString(3);     //Bt
                t6.message      = rs.getString(4);     //変更内容
                t6.key1         = rs.getInt(5);        //レシピNo
                t6.key2         = rs.getInt(6);        //テーブルNo
                t6.key3         = rs.getInt(7);        //大項目
                t6.key4         = rs.getFloat(8);      //中項目
                t6.key5         = rs.getFloat(9);      //項目No
                t6.key6         = rs.getFloat(10);     //変更前
                t6.key7         = rs.getFloat(11);     //変更前
                ret.addElement(t6);
            } // for end
            log("CZSystem getModifyHistoryT6","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getModifyHistoryT6","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getModifyHistoryT6","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }
    ////////////////////////////////////////////////////////////////////
    //
    //  編集履歴の取り出し用条件取得(T1～T5)
    //
	@SuppressWarnings("unchecked")
    public static Vector getModifyHistoryTX1(String date1, String date2,int t, String roName){
    
        initCheck();
        Vector  ret         = new Vector(1000);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;
		
        int i = 0;

        sql = new String("SELECT R.S_TIME,R.OP_NAME,R.BATCH,M.MESSAGE,R.KEY1,R.KEY2,R.KEY3 FROM " + roName + 
                         ".R_MODIFY R, MST.M_TBLMSG_MAST M where (R.ID_TABLE = M.S_NO) and  (R.ID_TABLE = 2 or R.ID_TABLE = 3) and (R.KEY1 =" +
                         t + " and R.S_TIME >=  TO_DATE('" + date1 + " 00:00:00','YYYY/MM/DD HH24:MI:SS') and R.S_TIME  <= TO_DATE('" + date2 + 
                         " 23:59:59','YYYY/MM/DD HH24:MI:SS')) ORDER BY R.S_TIME"); 

        log("CZSystem getModifyHistoryTX1","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getModifyHistoryTX1","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getModifyHistoryTX1","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getModifyHistoryTX1","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZSystemModifyHistoryTX1 tx1   = new CZSystemModifyHistoryTX1();
                tx1.s_time       = rs.getString(1);     //変更日時
                tx1.op_name      = rs.getString(2);     //変更者
                tx1.batch        = rs.getString(3);     //Bt
                tx1.message      = rs.getString(4);     //変更内容
                tx1.key1         = rs.getInt(5);        //レシピNo
                tx1.key2         = rs.getInt(6);        //テーブルNo
                tx1.key3         = rs.getInt(7);        //大項目
                ret.addElement(tx1);
            } // for end
            log("CZSystem getModifyHistoryTX1","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getModifyHistoryTX1","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getModifyHistoryTX1","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }
    ////////////////////////////////////////////////////////////////////
    //
    //  編集履歴の取り出し(T1～T5 列数取得)
    //
	@SuppressWarnings("unchecked")
    public static int getModifyHistoryCnt(int flg,String date, String roName){
    
        initCheck();
        int         ret     = 0;
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;
		
        int i = 0;

        log("CZSystem getModifyHistoryCnt","SQL["+date+"]");


        sql = new String("SELECT COUNT(*) from " + roName + ".R_CT_CHG_HISTORY  WHERE FLG = " + flg + " and K_DATE = TO_DATE('" + date + "','YYYY/MM/DD  HH24:MI:SS')"); 

        log("CZSystem getModifyHistoryCnt","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getModifyHistoryCnt","ERROR: failed to load JDBC driver.");
            handleException(e);
            return 0;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getModifyHistoryCnt","ERROR: failed to connect!");
            handleException(e);
            return 0;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getModifyHistoryCnt","ERROR: createStatement or database");
            return 0;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            rs = sqlstmt.executeQuery(sql) ;
            rs.next();
            ret= rs.getInt(1); //カウント数
            i = 1;
            log("CZSystem getModifyHistoryCnt","SELECT Count:" + i);

        }
        catch( SQLException e ){
            log("CZSystem getModifyHistoryCnt","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getModifyHistoryCnt","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return 0;
        return ret;
    }
    ////////////////////////////////////////////////////////////////////
    //
    //  編集履歴の取り出し(T1～T5)
    //
	@SuppressWarnings("unchecked")
    public static Vector getModifyHistoryTX2(int plus_flg ,String date, String roName){
    
        initCheck();
        Vector  ret         = new Vector(1000);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;
		
        int i = 0;

        switch(plus_flg) {
            //変更前と変更後で項目数が変わらない
            case 0:
                sql = new String("SELECT bf.k_no,bf.L_VAL,bf.R_VAL,af.L_VAL,af.R_VAL from (SELECT K_NO,L_VAL,R_VAL from " + roName + ".R_CT_CHG_HISTORY  WHERE K_DATE = TO_DATE('" + date + "','YYYY/MM/DD  HH24:MI:SS')" + 
                                 " and FLG = 0 ) bf ,(SELECT K_NO,L_VAL,R_VAL from " + roName + ".R_CT_CHG_HISTORY  WHERE K_DATE = TO_DATE('" + date + "','YYYY/MM/DD  HH24:MI:SS') and FLG = 1) af WHERE af.K_NO = bf.K_NO ORDER BY bf.K_NO"); 
                break;
            //変更前の方が項目数が多い
            case 1:
                sql = new String("SELECT bf.k_no,bf.L_VAL,bf.R_VAL,NVL(af.L_VAL,999999),NVL(af.R_VAL,999999) from (SELECT K_NO,L_VAL,R_VAL from " + roName + ".R_CT_CHG_HISTORY  WHERE K_DATE = TO_DATE('" + date + "','YYYY/MM/DD  HH24:MI:SS')" + 
                                 " and FLG = 0 ) bf ,(SELECT K_NO,L_VAL,R_VAL from " + roName + ".R_CT_CHG_HISTORY  WHERE K_DATE = TO_DATE('" + date + "','YYYY/MM/DD  HH24:MI:SS') and FLG = 1) af WHERE af.K_NO (+) = bf.K_NO ORDER BY bf.K_NO"); 
                break;
            //変更後の方が項目数が多い
            case 2:
                sql = new String("SELECT af.k_no,NVL(bf.L_VAL,999999),NVL(bf.R_VAL,999999),af.L_VAL,af.R_VAL from (SELECT K_NO,L_VAL,R_VAL from " + roName + ".R_CT_CHG_HISTORY  WHERE K_DATE = TO_DATE('" + date + "','YYYY/MM/DD  HH24:MI:SS')" + 
                                 " and FLG = 0 ) bf ,(SELECT K_NO,L_VAL,R_VAL from " + roName + ".R_CT_CHG_HISTORY  WHERE K_DATE = TO_DATE('" + date + "','YYYY/MM/DD  HH24:MI:SS') and FLG = 1) af WHERE af.K_NO = bf.K_NO (+) ORDER BY af.K_NO"); 
                break;
        }
        log("CZSystem getModifyHistoryTX2","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getModifyHistoryTX2","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getModifyHistoryTX2","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getModifyHistoryTX2","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZSystemModifyHistoryTX2 tx2   = new CZSystemModifyHistoryTX2();
                tx2.k_no         = rs.getInt(1);       //項目No
                tx2.l_val_bf     = rs.getFloat(2);     //変更前（L軸）
                tx2.r_val_bf     = rs.getFloat(3);     //変更前（R軸）
                tx2.l_val_af     = rs.getFloat(4);     //変更後（L軸）
                tx2.r_val_af     = rs.getFloat(5);     //変更後（R軸）
                ret.addElement(tx2);
            } // for end
            log("CZSystem getModifyHistoryTX2","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getModifyHistoryTX2","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getModifyHistoryTX2","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }
// add end 2008.10.08

    ////////////////////////////////////////////////////////////////////
    //
    //  輝度変化チェックCSV出力データ
    //
	@SuppressWarnings("unchecked")
    public static Vector getBrightnessCsvData(String date1, String date2,int t, String roName){
    
        initCheck();
        Vector  ret         = new Vector(1000);
        Connection  conn    = null;
        Statement   sqlstmt = null;
        ResultSet   rs      = null;
        String      sql     = null;
		
        int i = 0;

/*
        sql = new String("SELECT s_time, batch, p_no, charge, gap, max_b_ave, range_b_ave, max_b_judge, range_b_judge, " +
                         "x_review, review_range, body_l_max_b_ave, body_r_max_b_ave, body_max_b_range, body_peek, body_peek_judge, " +
                         "len, data, c_batch, c_max_b_ave, c_range_b_ave, t_max_b_judge, t_range_b_judge, c_body_l_max_b_ave, " +
                         "c_body_r_max_b_ave, t_body_l_max_b_ave, t_body_r_max_b_ave" +
                         "FROM " + roName + ".r_brightness_change " + 
                         "WHERE p_no = " + t +
                         " AND s_time >= TO_DATE('" + date1 + " 00:00:00",'YYYY/MM/DD HH24:MI:SS')" + 
                         " AND s_time <= TO_DATE('" + date2 + " 23:59:59','YYYY/MM/DD HH24:MI:SS')" + 
                         " AND c_batch IS NOT NULL");
*/
        sql = new String("SELECT * FROM " + roName + ".r_brightness_change " +
                         "WHERE p_no = " + t +
                         " AND s_time >= TO_DATE('" + date1 + " 00:00:00','YYYY/MM/DD HH24:MI:SS')" + 
                         " AND s_time <= TO_DATE('" + date2 + " 23:59:59','YYYY/MM/DD HH24:MI:SS')" + 
                         " AND c_batch IS NOT NULL");

        log("CZSystem getBrightnessCsvData","SQL["+sql+"]");

        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            log("CZSystem getBrightnessCsvData","ERROR: failed to load JDBC driver.");
            handleException(e);
            return null;
        }

        try{
            conn = DriverManager.getConnection(DB_URL,USER,PASSWD);
        }
        catch (SQLException e) {
            log("CZSystem getBrightnessCsvData","ERROR: failed to connect!");
            handleException(e);
            return null;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            log("CZSystem getBrightnessCsvData","ERROR: createStatement or database");
            return null;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZSystemBrightnessData b_data = new CZSystemBrightnessData();
                b_data.s_time             = rs.getString(1);     // 採取日時
                b_data.batch              = rs.getString(2);     // バッチNo
                b_data.p_no               = rs.getInt(3);        // プロセスNo
                b_data.charge             = rs.getInt(4);        // チャージ量
                b_data.gap                = rs.getString(5);     // GAP
                b_data.max_b_ave          = rs.getFloat(6);      // NS:最大輝度平均
                b_data.range_b_ave        = rs.getFloat(7);      // NS:指定区間輝度平均
                b_data.max_b_judge        = rs.getFloat(8);      // NS:最大輝度判定閾値
                b_data.range_b_judge      = rs.getFloat(9);      // NS:指定区間輝度判定閾値
                b_data.x_review           = rs.getFloat(10);     // NS:評価X座標
                b_data.review_range       = rs.getFloat(11);     // NS:評価範囲
                b_data.body_l_max_b_ave   = rs.getFloat(12);     // B:(左)最大輝度平均
                b_data.body_r_max_b_ave   = rs.getFloat(13);     // B:(右)最大輝度平均
                b_data.body_max_b_range   = rs.getFloat(14);     // B:最大輝度判定閾値
                b_data.body_peek          = rs.getFloat(15);     // B:片ピーク
                b_data.body_peek_judge    = rs.getFloat(16);     // B:片ピーク判定閾値
                b_data.len                = rs.getInt(17);       // データ数
                b_data.data               = rs.getString(18);    // データ
                b_data.c_batch            = rs.getString(19);    // 比較対象バッチNo
                b_data.c_max_b_ave        = rs.getFloat(20);     // (比較対象)NS:最大輝度平均
                b_data.c_range_b_ave      = rs.getFloat(21);     // (比較対象)NS:指定区間輝度平均
                b_data.t_max_b_judge      = rs.getFloat(22);     // (閾値判定対象値)NS:最大輝度平均
                b_data.t_range_b_judge    = rs.getFloat(23);     // (閾値判定対象値)NS:指定区間輝度平均
                b_data.c_body_l_max_b_ave = rs.getFloat(24);     // (比較対象)B:(左)最大輝度平均
                b_data.c_body_r_max_b_ave = rs.getFloat(25);     // (比較対象)B:(右)最大輝度平均
                b_data.t_body_l_max_b_ave = rs.getFloat(26);     // (閾値判定対象値)B:(左)最大輝度平均
                b_data.t_body_r_max_b_ave = rs.getFloat(27);     // (閾値判定対象値)B:(右)最大輝度平均
                ret.addElement(b_data);
            } // for end
            log("CZSystem getBrightnessCsvData","SELECT Count:" + i);
        }
        catch( SQLException e ){
            log("CZSystem getBrightnessCsvData","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();          //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            log("CZSystem getBrightnessCsvData","ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        if(1 > i) return null;
        return ret;
    }
    
    //  20050725
    //  炉番表示桁数変更
    //
    public static String RoKetaChg(String roname){
		
		String ro = new String();
		
		if( 0 != CZSystemDefine.DISP_KETA_FLG){
			StringBuffer a = new StringBuffer();
			a.append(roname);
			a.delete(0,1);
			ro = a.toString();
		} else {
			StringBuffer a = new StringBuffer();
			a.append(roname);
			ro = a.toString();
		}
		
		return ro;
	}


    
    
    //
    //  秒から書式付き時間へ変換
    //
    public static String timeFormat(long sec){

        DecimalFormat   format1   = new DecimalFormat("000");
        DecimalFormat   format2   = new DecimalFormat("00");

        long hh = sec / 3600;
        long mm = (sec % 3600) / 60;
        long ss = sec - (hh * 3600 + mm * 60) ;

        String ret = new String(format1.format(hh) + ":" +  
                format2.format(mm) + ":" +
                format2.format(ss));
        return ret;
    }


    //
    //  現在時間から過去未来時間へ変換
    //
    public static String dayTime(long day){

        java.util.Date now_date      =  new java.util.Date();
        long now = now_date.getTime();
        long val = now + (3600000l * 24l * day);
        java.util.Date new_date      =  new java.util.Date(val);
        SimpleDateFormat fm =  new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String date = fm.format(new_date);
        return date;
    }

    //
    //  現在時間の取得
    //
    public static String getDateTime(){

        java.util.Date system_date  =  new java.util.Date();
        SimpleDateFormat system_date_fm =  new SimpleDateFormat ("MM/dd HH:mm:ss"); 
        String date = system_date_fm.format(system_date);
        return date;
    }

    //
    //  現在時間の取得
    //
    public static String getDateTime(String sFormat){
        java.util.Date system_date  =  new java.util.Date();
        SimpleDateFormat system_date_fm =  new SimpleDateFormat (sFormat); 
        String date = system_date_fm.format(system_date);
        return date;
    }

    //
    //  エラー表示ＭＡＸの取得
    //
    public static int getErrorMax(){

        int ret = 0;
        try{
            ret = Integer.parseInt(ERROR_MAX);
        } catch(Exception e) {
            ret = 500;
        }
        return ret;
    }

    //
    //  グラフ表示制限枚数（グラフ表示）
    //
    public static int GraphCountUp(){
		
		graph_cnt++;
		
		log("CZSystem GraphCountUp", "graph_cnt : " + graph_cnt);
		return graph_cnt;
    }

    //
    //  グラフ表示制限枚数（グラフ非表示）
    //
    public static int GraphCountDown(){
		
		graph_cnt--;
		
		log("CZSystem GraphCountDown", "graph_cnt : " + graph_cnt);
		return graph_cnt;
    }

    //
    //  グラフ表示枚数管理
    //
    public static int GraphCount(){
		log("CZSystem GraphCount", "graph_cnt : " + graph_cnt);
		return graph_cnt;
    }

    //
    //  炉番INDEX取得
    //
    public static int getRoIndex(String roName){
		
		log("CZSystem getRoIndex", "roName : " + roName);
		Vector roInd = getRoNameList();
		
		for(int i = 0; i < roInd.size(); i++){
		String rs = RoKetaChg((String)roInd.elementAt(i));
			if(rs.equals(roName)){
				RoIndex = i;
				log("CZSystem getRoIndex", "RoIndex : " + RoIndex);
				break;
			}
		}
		return RoIndex;
    }

    //
    //  運転画面表示フラグ変更 (@20131030)
    //
    public static boolean untenChgView(){
		//log("CZSystem untenChgView", "untenFlg : " + untenFlg);
        if (untenFlg == true){
            untenFlg = false;
        } else {
            untenFlg = true;
        }
		return untenFlg;
    }

    //
    //  運転画面表示フラグ管理 (@20131030)
    //
    public static boolean untenView(){
		//log("CZSystem untenView", "untenFlg : " + untenFlg);
		return untenFlg;
    }

}


