package cz;

import java.awt.Color;

/*
 *   ＣＺシステムクラス
 *   共通クラス
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * @Update 2013.10.21 他基地参照機能
 */
public class CZSystemDefine {
    public static final String PROPERTY_FILE = "CZPROPERTY.TXT";
    public static final String TPGPROPERTY_FILE = "CZTPGPROPERTY.TXT";
    public static final String FPAVEPROPERTY_FILE = "CZFPAVEPROPERTY.TXT";
    public static final String IP_PROPERTY_FILE = "IP_PROPERTY.TXT";

    public static final int NONE_RUN    = -1;
    public static final int ADMIN_RUN   = 0;
    public static final int USER_RUN    = 1;
    public static final int REFERENCE_RUN    = 2;	// @20131021

    public static final int NONE_MODE   = -1;
    public static final int HOST_MODE   = 0;
    public static final int CMS_MODE    = 1;
    public static final int LIB_MODE    = 2;

    public static final int ERROR_MAX     = 128;

    public static final int READY   = 0;
    public static final int VAC = 1;
    public static final int MELT    = 2;
    public static final int DIP = 3;
    public static final int NECK1   = 4;
    public static final int NECK2   = 5;
    public static final int SHOULDER= 6;
    public static final int BODY    = 7;
    public static final int TAIL    = 8;
    public static final int END = 9;


    public static final String PROC_NAME[] = {
                new String("READY"),
                new String("VAC"),
                new String("MELT"),
                new String("DIP"),
                new String("NECK1"),
                new String("NECK2"),
                new String("SHOULDER"),
                new String("BODY"),
                new String("TAIL"),
                new String("END")};

    public static final String PROC_NAME2[] = {
                new String("VAC"),
                new String("MELT"),
                new String("DIP"),
                new String("NECK1"),
                new String("NECK2"),
                new String("SHLD"),
                new String("BODY"),
                new String("TAIL"),
                new String("END"),
                new String("ALL")};

    public static final String PROC_NAME3[] = {
                new String("NECK"),
                new String("SHLD"),
                new String("BODY"),
                new String("ALL")};

    public static final int START_PROC_START    = 1;
    public static final int START_PROC_RESTART  = 2;

    public static final int PROC_MANUAL = 0;
    public static final int PROC_AUTO   = 1;
    public static final String PROC_MODE[] = {
                new String("手  動"),
                new String("自  動")};


    public static final int PV_MAX_LENGTH = 128;

    public static final Color DEFAULT_BACKGROUND_COL = new java.awt.Color(170,170,235);
    public static final Color DEFAULT_REFERENCE_BACKGROUND_COL = new java.awt.Color(155,251,194);
//    public static final Color DEFAULT_REFERENCE_BACKGROUND_COL = new java.awt.Color(153,0,204);

    public static final Color BUTTON_NORMAL_COL = java.awt.Color.lightGray;
    public static final Color BUTTON_SEND_COL   = java.awt.Color.green;
    public static final Color BUTTON_WAIT_COL   = java.awt.Color.red;

//    public static final int CT_TABLE_CLOSE_TIME = 10000;
    
//    public static final int ALERM_DIALOG_CLOSE_TIME = 5000;

    public static final int CT_TABLE_CLOSE_TIME = Integer.valueOf(CZSystem.CT_TABLE_CLOSE_TIME).intValue();
    
    public static final int ALERM_DIALOG_CLOSE_TIME = Integer.valueOf(CZSystem.ALERM_DIALOG_CLOSE_TIME).intValue();
    
    public static final int TIMER_FLG = Integer.valueOf(CZSystem.TIMER_FLG).intValue();
    
    // 20050725 炉番表示桁数変更
    public static final int DISP_KETA_FLG = Integer.valueOf(CZSystem.DISP_KETA_FLG).intValue();

}

