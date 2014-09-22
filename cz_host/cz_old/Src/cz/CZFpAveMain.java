package cz;

import java.awt.Color;
import java.awt.Cursor;
import java.awt.Dimension;
import java.awt.Graphics;
import java.awt.Rectangle;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.Serializable;
import java.text.DecimalFormat;
import java.util.Locale;
import java.util.Properties;
import java.util.Vector;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JColorChooser;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.JViewport;
import javax.swing.ListSelectionModel;
import javax.swing.event.ListSelectionEvent;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.DefaultTableColumnModel;
import javax.swing.table.JTableHeader;
import javax.swing.table.TableColumn;
import javax.swing.text.AttributeSet;
import javax.swing.text.BadLocationException;
import javax.swing.text.Document;
import javax.swing.text.PlainDocument;

/**
 *   Ｆｐ移動平均基本グラフ
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * T6追加 @@
 * 設定値保存・読込追加 @@@
 * 2008.09.17 H.Nagamine TPG・PV保存対象表示情報追加
 */
public class CZFpAveMain extends JFrame {

    private final int T1 = 1;
    private final int T2 = 2;
    private final int T3 = 3;
    private final int T4 = 4;
    private final int T5 = 5;
    private final int T6 = 6;       //@@

    private final Color BACK_COL        = java.awt.Color.black;
    private final Color MEM_LINE1_COL   = java.awt.Color.lightGray;
    private final Color MEM_LINE2_COL   = java.awt.Color.gray;
    private final Color MEM_LINE3_COL   = java.awt.Color.darkGray;

    private final int MAIN1_H_T     = 14;   // 15   メインヒーター１温度
    private final int MAIN1_H_T_PF  = 66;   // 67   メインヒーター１温度プロファイル
    private final int DIA           = 24;   // 25   直径
    private final int DIA_PF        = 23;   // 24   直径プロファイル
    private final int SXL_ST        = 17;   // 18   引き上げ速度
    private final int SXL_ST_PF     = 75;   // 76   引き上げ速度プロファイル
    private final int SXL_RT        = 18;   // 19   シード回転
    private final int SXL_RT_PF     = 80;   // 81   シード回転プロファイル
    private final int CRU_RT        = 20;   // 21   ルツボ回転
    private final int CRU_RT_PF     = 86;   // 87   ルツボ回転プロファイル
    private final int PULL_AR       = 15;   // 16   プルアルゴン
    private final int PULL_AR_PF    = 71;   // 72   プルアルゴンプロファイル
    private final int VAC           = 32;   // 33   炉内圧
    private final int VAC_PF        = 88;   // 89   炉内圧プロファイル

    private String  fp_ave_time_pro;        //移動平均時間(初期値)
    private String  pf_umax_pro;            //プロファイルの上上限
    private String  pf_max_pro;             //プロファイルの上限
    private String  pf_lmin_pro;            //プロファイルの下下限
    private String  pf_min_pro;             //プロファイルの下限

    private String  shld_shift_dia;         //肩変え直径
    private String  shld_shift_length;      //肩変え位置
    //Ｘ軸
    private String  x_length_min;           //Ｘ軸最小値
    private String  x_length_max;           //Ｘ軸最大値
    private String  x_length_bunkatu;       //Ｘ軸分割数
    private String  x_length_koushi;        //Ｘ軸格子間隔
    private String  x_length_memkan;        //Ｘ軸目盛値間隔
    private String  x_length_memketa;       //Ｘ軸目盛桁数
    private String  x_length_syouketa;      //Ｘ軸小数桁数
    //Ｙ軸
    private String  sxl_st_min_pro;         //Ｙ軸引上速度最小値
    private String  sxl_st_max_pro;         //Ｙ軸引上速度最大値
    private String  sxl_st_bunkatu;         //Ｙ軸分割
    private String  sxl_st_koushi;          //Ｙ軸格子間隔
    private String  sxl_st_memkan;          //Ｙ軸目盛値間隔
    private String  sxl_st_memketa;         //Ｙ軸目盛桁数
    private String  sxl_st_syouketa;        //Ｙ軸小数桁数
    private String  dia_min_pro;            //直径
    private String  dia_max_pro;
    private String  sxl_rt_pf_min_pro;      //シード回転プロファイル
    private String  sxl_rt_pf_max_pro;

    private String  dia_pf_min_pro;         //直径プロファイル
    private String  dia_pf_max_pro;

/* @@@
    private String  main1_h_t_min_pro;      //メインヒーター１温度
    private String  main1_h_t_max_pro;
    private String  main1_h_t_pf_min_pro;   //メインヒーター１温度プロファイル
    private String  main1_h_t_pf_max_pro;
    private String  sxl_st_pf_min_pro;      //引き上げ速度プロファイル
    private String  sxl_st_pf_max_pro;
    private String  cru_rt_pf_min_pro;      //ルツボ回転プロファイル
    private String  cru_rt_pf_max_pro;
    private String  pull_ar_pf_min_pro;     //プルアルゴンプロファイル
    private String  pull_ar_pf_max_pro;
    private String  vac_pf_min_pro;         //炉内圧プロファイル
    private String  vac_pf_max_pro;
 @@@ */

    private String ro_name              = null;     //対象炉番
    private String ro_db_name           = null;     //対象炉データベース名

    private CZSystemStart ro_bt_start   = null;     //検索用引き上げ条件
    private Vector ro_bt_all_condition  = null;     //全Btの引き上げ条件

    private Vector pv_data_body         = null;     //ボディーのデータ
    private Vector calc_data_body       = null;     //ボディーの計算データ

    static  JLabel main_ro_name_lab     = null;     //炉番表示

    private SercheDialog    serche_dia  = null;     //検索用ダイアログ
    private CZRoSelectWin4  rosel       = null;

    static ConditionPanel  conpane     = null;     //検索、引き上げ条件パネル

    private Vector roBtTempCondition_    = null; //選択Btの引き上げ条件

    private String SelectBt             = null;
    private String SelectTime           = null;

    private SetPanel    setpane         = null;     //条件設定パネル
    private GraphSet    graph_set       = null;     //グラフ描画条件

    private DataPanel   datapane        = null;     //データパネル

    private CZFpAveGraphFrame graph_dia       = null;     //グラフ用ダイアログ
    private int     fp_ave_calc_time    = 10;       //移動平均時間(計算に使用)

    private int gph_cnt = 0;

    private File file_ = new File(CZSystem.FILE_SRC_PATH);

    /**
    * コンストラクタ
    */
    CZFpAveMain(){
        super();

        try{
            //設定値を取得する。
            Properties prop =  new Properties();
            FileInputStream pros = new FileInputStream(CZSystemDefine.FPAVEPROPERTY_FILE);
            prop.load(pros);

            fp_ave_time_pro     = prop.getProperty("FP_AVE_TIME");          //移動平均時間
            pf_umax_pro         = prop.getProperty("FP_PF_UMAX");           //プロファイルの上上限
            pf_max_pro          = prop.getProperty("FP_PF_MAX");            //プロファイルの上限
            pf_lmin_pro         = prop.getProperty("FP_PF_LMIN");           //プロファイルの下下限
            pf_min_pro          = prop.getProperty("FP_PF_MIN");            //プロファイルの下限

            shld_shift_dia      = prop.getProperty("SHLD_SHIFT_DIA");       //肩変え直径 @@@
            shld_shift_length   = prop.getProperty("SHLD_SHIFT_LENGTH");    //肩変え位置 @@@
            //Ｘ軸
            x_length_min        = prop.getProperty("X_LENGTH_MIN");         //Ｘ軸最小値
            x_length_max        = prop.getProperty("X_LENGTH_MAX");         //Ｘ軸最大値
            x_length_bunkatu    = prop.getProperty("X_LENGTH_BUNKATU");     //Ｘ軸分割数
            x_length_koushi     = prop.getProperty("X_LENGTH_KOUSHI");      //Ｘ軸格子間隔 @@@
            x_length_memkan     = prop.getProperty("X_LENGTH_MEMKAN");      //Ｘ軸目盛値間隔 @@@
            x_length_memketa    = prop.getProperty("X_LENGTH_MEMKETA");     //Ｘ軸目盛桁数 @@@
            x_length_syouketa   = prop.getProperty("X_LENGTH_SYOUKETA");    //Ｘ軸小数桁数 @@@
            //Ｙ軸
            sxl_st_min_pro      = prop.getProperty("SXL_ST_MIN");           //Ｙ軸引上速度最小値
            sxl_st_max_pro      = prop.getProperty("SXL_ST_MAX");           //Ｙ軸引上速度最大値
            sxl_st_bunkatu      = prop.getProperty("SXL_ST_BUNKATU");       //Ｙ軸分割
            sxl_st_koushi       = prop.getProperty("SXL_ST_KOUSHI");        //Ｙ軸格子間隔 @@@
            sxl_st_memkan       = prop.getProperty("SXL_ST_MEMKAN");        //Ｙ軸目盛値間隔 @@@
            sxl_st_memketa      = prop.getProperty("SXL_ST_MEMKETA");       //Ｙ軸目盛桁数 @@@
            sxl_st_syouketa     = prop.getProperty("SXL_ST_SYOUKETA");      //Ｙ軸小数桁数 @@@
            dia_min_pro         = prop.getProperty("DIA_MIN");              //直径最小値
            dia_max_pro         = prop.getProperty("DIA_MAX");              //直径最大値
            sxl_rt_pf_min_pro   = prop.getProperty("SXL_RT_PF_MIN");        //シード回転プロファイル最小値
            sxl_rt_pf_max_pro   = prop.getProperty("SXL_RT_PF_MAX");        //シード回転プロファイル最大値

            dia_pf_min_pro          = prop.getProperty("DIA_PF_MIN");       //直径プロファイル
            dia_pf_max_pro          = prop.getProperty("DIA_PF_MAX");
/* @@@
            main1_h_t_min_pro       = prop.getProperty("MAIN1_H_T_MIN");    //メインヒーター１温度
            main1_h_t_max_pro       = prop.getProperty("MAIN1_H_T_MAX");
            main1_h_t_pf_min_pro    = prop.getProperty("MAIN1_H_T_PF_MIN"); //メインヒーター１温度プロファイル
            main1_h_t_pf_max_pro    = prop.getProperty("MAIN1_H_T_PF_MAX");

            sxl_st_pf_min_pro       = prop.getProperty("SXL_ST_PF_MIN");    //引き上げ速度プロファイル
            sxl_st_pf_max_pro       = prop.getProperty("SXL_ST_PF_MAX");
            cru_rt_pf_min_pro       = prop.getProperty("CRU_RT_PF_MIN");    //ルツボ回転プロファイル
            cru_rt_pf_max_pro       = prop.getProperty("CRU_RT_PF_MAX");
            pull_ar_pf_min_pro      = prop.getProperty("PULL_AR_PF_MIN");   //プルアルゴンプロファイル
            pull_ar_pf_max_pro      = prop.getProperty("PULL_AR_PF_MAX");
            vac_pf_min_pro          = prop.getProperty("VAC_PF_MIN");       //炉内圧プロファイル
            vac_pf_max_pro          = prop.getProperty("VAC_PF_MAX");
 @@@*/
        }
        catch( Exception e){
            CZSystem.exit(-1,"CZFpAveMain NO Propertie File");
        }

        ro_name     = CZSystem.getRoName();
        ro_db_name  = CZSystem.getDBName();

        setTitle("fp移動平均設定");
        setSize(1110,864);
//@@        setSize(1152,864);
        setResizable(false);
//        setModal(true);
        getContentPane().setLayout(null);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        //炉番表示
		String s = CZSystem.RoKetaChg(ro_name);	// 20050725 炉：表示桁数変更
        main_ro_name_lab = new JLabel(s,JLabel.CENTER);
//        main_ro_name_lab = new JLabel(ro_name,JLabel.CENTER);
        main_ro_name_lab.setBounds(20, 20, 70, 30);
        main_ro_name_lab.setLocale(new Locale("ja","JP"));
        main_ro_name_lab.setFont(new java.awt.Font("dialog", 0, 18));
        main_ro_name_lab.setBorder(new Flush3DBorder());
        main_ro_name_lab.setForeground(java.awt.Color.black);
        getContentPane().add(main_ro_name_lab);

        JButton btn_chgRo = new JButton("▼");
        btn_chgRo.setBounds(90, 20, 30, 30);
        btn_chgRo.setFont(new java.awt.Font("dialog", 0, 20));
        btn_chgRo.setBorder(new Flush3DBorder());
        btn_chgRo.setForeground(java.awt.Color.black);
        btn_chgRo.addActionListener(
            new ActionListener() {
				public void actionPerformed(ActionEvent ev){
					rosel = new CZRoSelectWin4();
					rosel.setVisible(true);
					ro_name = main_ro_name_lab.getText();
					ro_db_name = main_ro_name_lab.getText();
				}
			}
		);
		getContentPane().add(btn_chgRo);

        //検索パネル
        conpane = new ConditionPanel();
        conpane.setBounds(20, 60, 100, 300);
        getContentPane().add(conpane);

        //条件設定パネル
        setpane = new SetPanel();
        setpane.setBounds(140, 20, 950, 340);
        getContentPane().add(setpane);

        //データパネル
        datapane = new DataPanel();
//@@        datapane.setBounds(20, 370, 1070, 500);
        datapane.setBounds(20, 370, 1070, 450);
        getContentPane().add(datapane);

        //データ検索用ダイアログ
        serche_dia = new SercheDialog();
        serche_dia.setVisible(false);

/*
        //グラフ表示用ダイアログ
        graph_dia = new GraphDialog();
        graph_dia.setVisible(false);
*/

//        CZSystem.log("CZFpAveMain","new");
    } //CZFpAveMain


    /**
     * 初期設定をする
     */
    public boolean setDefault(){
        ro_name     = CZSystem.getRoName();
        ro_db_name  = CZSystem.getDBName();
        setpane.setDefault();
        datapane.setDefault();
        return true;
    }

    /**
     * 炉番を表示する
     */
    private void setMainRoName(){
		String s = CZSystem.RoKetaChg(ro_name);	// 20050725 炉：表示桁数変更
        main_ro_name_lab.setText(s);
//        main_ro_name_lab.setText(ro_name);
    }

    /**
     * メッセージを表示する
     */
    private boolean errorMsg(Object msg[]){
        JOptionPane.showMessageDialog(null,msg,
                    "fp移動平均エラー",
                    JOptionPane.ERROR_MESSAGE);
        return true;
    }

    /**
     * バッチ開始情報を設定する。
     */
    public boolean setBtStart(CZSystemStart st){
        ro_bt_start = st;
        if(null == ro_bt_start) return false;
        return true;
    }

    /**
     * バッチ開始情報を削除する。
     */
    public boolean removeBtStart(){
        ro_bt_start = null;
        return true;
    }

    /**
     *  PVデータを読み込む
     */
    public int readBtPV(){
        if(null == ro_bt_start){
            Object msg[] = {"スタート実績が有りません！！",
                            "",
                            ""};
            errorMsg(msg);
            return -1;
        }

        CZSystemStart st = ro_bt_start;

/*
        CZSystem.log("CZFpAveMain","readBtPV() batch    [" + st.batch    + "]");
        CZSystem.log("CZFpAveMain","readBtPV() p_no     [" + st.p_no     + "]");
        CZSystem.log("CZFpAveMain","readBtPV() sp_no    [" + st.sp_no    + "]");
        CZSystem.log("CZFpAveMain","readBtPV() p_renban [" + st.p_renban + "]");
        CZSystem.log("CZFpAveMain","readBtPV() p_start  [" + st.p_start  + "]");
*/

        String view = CZSystem.getViewName(ro_db_name,st.batch);

        CZSystem.log("CZFpAveMain","readBtPV() view  [" + view  + "]");

        if(null == view){
            Object msg[] = {"表が存在しません！！",
                            view,
                            ""};
            errorMsg(msg);
            return -2;
        }
        //読込むPVを設定する。
        boolean data_no[] = new boolean[CZSystemDefine.PV_MAX_LENGTH];
        for (int i = 0 ; i < CZSystemDefine.PV_MAX_LENGTH ; i++) {
            data_no[i] = false;
        }

        data_no[ MAIN1_H_T  ]   = true;   // 15   メインヒーター１温度
        data_no[ MAIN1_H_T_PF ] = true;   // 67   メインヒーター１温度プロファイル
        data_no[ DIA        ]   = true;   // 25   直径
        data_no[ DIA_PF     ]   = true;   // 24   直径プロファイル
        data_no[ SXL_ST     ]   = true;   // 18   引き上げ速度
        data_no[ SXL_ST_PF  ]   = true;   // 76   引き上げ速度プロファイル

        data_no[ SXL_RT     ]   = true;   // 19   シード回転
        data_no[ SXL_RT_PF  ]   = true;   // 81   シード回転プロファイル
        data_no[ CRU_RT     ]   = true;   // 21   ルツボ回転
        data_no[ CRU_RT_PF  ]   = true;   // 87   ルツボ回転プロファイル

        data_no[ PULL_AR    ]   = true;   // 16   プルアルゴン
        data_no[ PULL_AR_PF ]   = true;   // 72   プルアルゴンプロファイル
        data_no[ VAC        ]   = true;   // 33   炉内圧
        data_no[ VAC_PF     ]   = true;   // 89   炉内圧プロファイル

/***** System.gc() *****/
//            System.out.println(Runtime.getRuntime().freeMemory());
            System.gc();
//            System.out.println(Runtime.getRuntime().freeMemory() + "  GC FREE!!");
/**********************/

            SelectBt = st.batch;
            SelectTime = st.p_start;
            
            
        //PV（ボディー）を読み込む
        pv_data_body = CZSystem.getPVData(ro_db_name,view,st.p_renban,data_no);
        
/***** System.gc() *****/
//            System.out.println(Runtime.getRuntime().freeMemory());
            System.gc();
//            System.out.println(Runtime.getRuntime().freeMemory() + "  GC FREE!!");
/**********************/

//        CZSystem.log("CZFpAveMain","readBtPV() pv_data_body  [" + pv_data_body.size()  + "]");
        if(1 > pv_data_body.size()){
            Object msg[] = {"ボディー実績が有りません！！",
                            "[" + pv_data_body.size() + "]",
                            ""};
            errorMsg(msg);
            return -3;
        }
        return pv_data_body.size();
    }

    /**
     *  移動平均を計算する
     */
	@SuppressWarnings("unchecked")
    private void startCalc(int calc_time,float umax,float max,float lmin,float min){
//        CZSystem.log("CZFpAveMain","startCalc() start");

        if(null == pv_data_body) return;        //データ無しは計算しない。
        int count = pv_data_body.size();        //データ件数を保持する。
        if(1 > count) return;                   //データ件数０件は計算しない。

        if(10 > calc_time){
            Object msg[] = {"計算時間を",
                            "１０秒以上にしてください。",
                            ""};
            errorMsg(msg);
            return;
        }

        if(0 != (calc_time % 10)){
            Object msg[] = {"計算時間を",
                            "１０秒単位にしてください。",
                            ""};
            errorMsg(msg);
            return;
        }

        fp_ave_calc_time = calc_time;           //計算時間を保持する。
        calc_data_body = new Vector(count);     //計算結果を保持する領域を確保する。
        CalcData st    = null;
        //移動平均を計算する。
        for (int i = 0 ; i < count ; i++) {
            st = calc(i,calc_time, umax,max,lmin,min);
            if(null == st) {
//                CZSystem.log("CZFpAveMain","startCalc() stop");
                Object msg[] = {"計算異常",
                            "",
                            ""};
                errorMsg(msg);
                return; 
            }
            calc_data_body.addElement(st);      //計算結果を保持する。
        }
        datapane.setCalc();                     //計算結果を画面に設定する。
//        CZSystem.log("CZFpAveMain","startCalc() end");
    }

    /**
     *  肩変え位置を探す
     */
    private float sercheShldLength(){
        int size = pv_data_body.size();
        CZSystemPVData st;

        for (int i = 0 ; i < size ; i++) {    
            st = (CZSystemPVData)pv_data_body.elementAt(i);
            if(2 == st.sp_no) return st.p_length;
        }
        st = (CZSystemPVData)pv_data_body.elementAt(0);
        return st.p_length;
    }

    /**
     *  肩変え位置のデータ格納場所を探す
     */
    private int selectShldLengthIndex() {
        int size = pv_data_body.size();
        CZSystemPVData st;

        for (int i = 0 ; i < size ; i++) {
            st = (CZSystemPVData)pv_data_body.elementAt(i);
            if(2 == st.sp_no){
                datapane.selectData(i);
                return i;
            }
        }
        datapane.selectData(0);
        return 0;
    }

    /**
     *  肩変え位置の直径を探す
     */
    private float sercheShldDia(){
        int size = pv_data_body.size();
        CZSystemPVData st;

        for (int i = 0 ; i < size ; i++) {
            st = (CZSystemPVData)pv_data_body.elementAt(i);
            if(2 == st.sp_no) return st.data[DIA_PF];
        }
        st = (CZSystemPVData)pv_data_body.elementAt(0);
        return st.data[DIA_PF];
    }

    /**
     *  一つのデータから移動平均時間までの移動平均計算
     */
    private CalcData calc(int start,int calc_time, float umax,float max,float lmin,float min){

        CalcData ret    = new CalcData();   //計算結果を保持する領域を確保し初期設定する。
        ret.fp_ave      = 0.0f;
        ret.pf_ave      = 0.0f;
        ret.pf_umax_ave = 0.0f;
        ret.pf_max_ave  = 0.0f;
        ret.pf_lmin_ave = 0.0f;
        ret.pf_min_ave  = 0.0f;
        ret.judg        = -99;

        CZSystemPVData s = null;
        CZSystemPVData e = null;

        try{
            s = (CZSystemPVData)pv_data_body.elementAt(start);
        }
        catch(ArrayIndexOutOfBoundsException err){
            return ret;
        }
        if(null == s){
            return null;
        }

        int size = pv_data_body.size();
        int next_time = s.p_time + calc_time;
        int j = 0;
        float pf_tmp = 0.0f;
        for (int i = start ; i < size ; i++) {
            e = (CZSystemPVData)pv_data_body.elementAt(i);
            if(null == e){
                return ret;
            }
            j++;
            pf_tmp += e.data[SXL_ST_PF];
            if(next_time != e.p_time) continue;
            float l = e.p_length - s.p_length;
            ret.fp_ave = l / (float)calc_time * 60.0f;
            ret.pf_ave = pf_tmp / (float)j ;
            ret.pf_umax_ave = ret.pf_ave + umax;
            ret.pf_max_ave  = ret.pf_ave + max;
            ret.pf_lmin_ave = ret.pf_ave + lmin;
            ret.pf_min_ave  = ret.pf_ave + min;
            ret.judg = judg(ret.fp_ave,ret.pf_ave,umax,max,lmin,min);
            return ret;
        } //for end
        return ret;
    }

    /**
     *  移動平均値がプロファイルの許容範囲かの判定をする。
     */
    private int judg(float fp , float pf ,float umax,float max,float lmin,float min){
        if((pf + umax) < fp) return  2; 
        if((pf + lmin) > fp) return -2; 
        if((pf + max) < fp) return  1;  
        if((pf + min) > fp) return -1;  
        return 0;
    }


    /**
     *  カーソルを設定する
     */
    private void setCur(Cursor cu){
        serche_dia.setCursor(cu);
    }

    /**
     *  カーソル参照
     */
    private Cursor getCur(){
        return serche_dia.getCursor();
    }

    //==========================================================================
    /**
    *   検索パネル
    */
    public class ConditionPanel extends JPanel {

        private JTextField bt_text = null;      //バッチ番号
        private JTextField t1_text = null;      //T1
        private JTextField t2_text = null;      //T2
        private JTextField t3_text = null;      //T3
        private JTextField t4_text = null;      //T4
        private JTextField t5_text = null;      //T5
        private JTextField t6_text = null;      //T6

        /**
         * コンストラクタ
         */
        ConditionPanel(){
            super();
            setName("ConditionPanel");
            setLayout(null);
            setBorder(new Flush3DBorder());
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            int x   = 10;
            int y   = 10;
            int inc = 0;

            JButton search_button = new JButton("検  索");
            search_button.setBounds(x, y, 80, 24);
            search_button.setLocale(new Locale("ja","JP"));
            search_button.setFont(new java.awt.Font("dialog", 0, 18));
            search_button.setBorder(new Flush3DBorder());
            search_button.setForeground(java.awt.Color.black);
            search_button.addActionListener(new SearchButton());
            add(search_button);

            x = 10;
            y = 50;
            bt_text = new JTextField();
            bt_text.setBounds(x, y, 80, 18);
            bt_text.setLocale(new Locale("ja","JP"));
            bt_text.setFont(new java.awt.Font("dialog", 0, 12));
            bt_text.setBorder(new Flush3DBorder());
            bt_text.setForeground(java.awt.Color.black);
            add(bt_text);

            JLabel lab = null;
            x   = 10;
            y   = 90 ;
            inc = 30;

            lab = new JLabel("T1",JLabel.CENTER);
            lab.setBounds(x, y, 40, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setForeground(java.awt.Color.black);
            add(lab);
            t1_text = new JTextField();
            t1_text.setBounds(x+40, y, 40, 18);
            t1_text.setLocale(new Locale("ja","JP"));
            t1_text.setFont(new java.awt.Font("dialog", 0, 12));
            t1_text.setBorder(new Flush3DBorder());
            t1_text.setForeground(java.awt.Color.black);
            add(t1_text);

            y += inc;
            lab = new JLabel("T2",JLabel.CENTER);
            lab.setBounds(x, y, 40, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setForeground(java.awt.Color.black);
            add(lab);
            t2_text = new JTextField();
            t2_text.setBounds(x+40, y, 40, 18);
            t2_text.setLocale(new Locale("ja","JP"));
            t2_text.setFont(new java.awt.Font("dialog", 0, 12));
            t2_text.setBorder(new Flush3DBorder());
            t2_text.setForeground(java.awt.Color.black);
            add(t2_text);

            y += inc;
            lab = new JLabel("T3",JLabel.CENTER);
            lab.setBounds(x, y, 40, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setForeground(java.awt.Color.black);
            add(lab);
            t3_text = new JTextField();
            t3_text.setBounds(x+40, y, 40, 18);
            t3_text.setLocale(new Locale("ja","JP"));
            t3_text.setFont(new java.awt.Font("dialog", 0, 12));
            t3_text.setBorder(new Flush3DBorder());
            t3_text.setForeground(java.awt.Color.black);
            add(t3_text);

            y += inc;
            lab = new JLabel("T4",JLabel.CENTER);
            lab.setBounds(x, y, 40, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setForeground(java.awt.Color.black);
            add(lab);
            t4_text = new JTextField();
            t4_text.setBounds(x+40, y, 40, 18);
            t4_text.setLocale(new Locale("ja","JP"));
            t4_text.setFont(new java.awt.Font("dialog", 0, 12));
            t4_text.setBorder(new Flush3DBorder());
            t4_text.setForeground(java.awt.Color.black);
            add(t4_text);

            y += inc;
            lab = new JLabel("T5",JLabel.CENTER);
            lab.setBounds(x, y, 40, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setForeground(java.awt.Color.black);
            add(lab);
            t5_text = new JTextField();
            t5_text.setBounds(x+40, y, 40, 18);
            t5_text.setLocale(new Locale("ja","JP"));
            t5_text.setFont(new java.awt.Font("dialog", 0, 12));
            t5_text.setBorder(new Flush3DBorder());
            t5_text.setForeground(java.awt.Color.black);
            add(t5_text);

            y += inc;
            lab = new JLabel("T6",JLabel.CENTER);
            lab.setBounds(x, y, 40, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setForeground(java.awt.Color.black);
            add(lab);
            t6_text = new JTextField();
            t6_text.setBounds(x+40, y, 40, 18);
            t6_text.setLocale(new Locale("ja","JP"));
            t6_text.setFont(new java.awt.Font("dialog", 0, 12));
            t6_text.setBorder(new Flush3DBorder());
            t6_text.setForeground(java.awt.Color.black);
            add(t6_text);

            y += inc;
            JButton btcondition_button = new JButton("引上条件");
            btcondition_button.setBounds(x, y, 80, 24);
            btcondition_button.setLocale(new Locale("ja","JP"));
            btcondition_button.setFont(new java.awt.Font("dialog", 0, 18));
            btcondition_button.setBorder(new Flush3DBorder());
            btcondition_button.setForeground(java.awt.Color.black);
            btcondition_button.addActionListener(new BtConditionButton());
            add(btcondition_button);

//            CZSystem.log("CZFpAveMain","ConditionPanel new");
        } //ConditionPanel

        /**
         *  検索データをセット
         */
        public void setData(boolean b){

            boolean flag = b;

            roBtTempCondition_ = CZSystem.getHikiageTemp(ro_name,SelectBt,SelectTime);

            if(null == roBtTempCondition_){
	            if(null == ro_bt_all_condition){    //バッチ情報の有無をチェックする。
	                flag = false;                   //バッチ情報無しフラグを設定する。

	                //画面をクリアする。
	                bt_text.setText("");
	                t1_text.setText("");
	                t2_text.setText("");
	                t3_text.setText("");
	                t4_text.setText("");
	                t5_text.setText("");
	                t6_text.setText("");
	            }else{
	                setMainRoName();
	                
	                //引上げ情報を画面に設定する。

	                CZSystemBt bt = (CZSystemBt)ro_bt_all_condition.elementAt(0);

	                if(null == bt) return;
	                bt_text.setText(bt.batch.trim());
	                t1_text.setText(String.valueOf(bt.no_youkai));
	                t2_text.setText(String.valueOf(bt.no_hikiage));
	                t3_text.setText(String.valueOf(bt.no_kaiten));
	                t4_text.setText(String.valueOf(bt.no_toridasi));
	                t5_text.setText(String.valueOf(bt.no_aturyoku));
	                t6_text.setText(String.valueOf(bt.no_teisu));
				}
			} else {
                setMainRoName();
                
                //引上げ情報を画面に設定する。

				CZSystemBtTemp bt = (CZSystemBtTemp)roBtTempCondition_.elementAt(0);

                if(null == bt) return;
                bt_text.setText(bt.batch.trim());
                t1_text.setText(String.valueOf(bt.no_youkai));
                t2_text.setText(String.valueOf(bt.no_hikiage));
                t3_text.setText(String.valueOf(bt.no_kaiten));
                t4_text.setText(String.valueOf(bt.no_toridasi));
                t5_text.setText(String.valueOf(bt.no_aturyoku));
                t6_text.setText(String.valueOf(bt.no_teisu));
			}
/*			
            if(flag){
                setMainRoName();
                
                CZSystem.log("CZFpAveMain","炉は？ " + ro_name);
                CZSystem.log("CZFpAveMain","バッチは？ " + SelectBt);
                CZSystem.log("CZFpAveMain","時間は？ " + SelectTime);
                
                //引上げ情報を画面に設定する。

//				CZSystemBtTemp bt = (CZSystemBtTemp)roBtTempCondition_.elementAt(0);

//                CZSystemBt bt = (CZSystemBt)ro_bt_all_condition.elementAt(0);

                if(null == bt) return;
                bt_text.setText(bt.batch.trim());
                t1_text.setText(String.valueOf(bt.no_youkai));
                t2_text.setText(String.valueOf(bt.no_hikiage));
                t3_text.setText(String.valueOf(bt.no_kaiten));
                t4_text.setText(String.valueOf(bt.no_toridasi));
                t5_text.setText(String.valueOf(bt.no_aturyoku));
                t6_text.setText(String.valueOf(bt.no_teisu));
            } else{
                //画面をクリアする。
                bt_text.setText("");
                t1_text.setText("");
                t2_text.setText("");
                t3_text.setText("");
                t4_text.setText("");
                t5_text.setText("");
                t6_text.setText("");
            }
*/
        }

        /**
         *  引上げ条件・データテーブル・計算結果をクリア
         */
         public void clearBtCondition() {
			bt_text.setText("");
			t1_text.setText("");
			t2_text.setText("");
			t3_text.setText("");
			t4_text.setText("");
			t5_text.setText("");
			t6_text.setText("");
			setpane.setMode(false);
			ro_bt_all_condition = null;
			pv_data_body    = null;
			calc_data_body  = null;
			datapane.setDefault();
		}

        //======================================================================
        /**
        *   検索ボタン
        */
        class SearchButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                serche_dia.setDefault();        //検索画面を初期化する。
                serche_dia.setVisible(true);    //検索画面を表示する。
            }
        } //SearchButton


        //======================================================================
        /**
        *   引き上げ条件ボタン
        */
        class BtConditionButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                if(null == ro_bt_all_condition) return ;
                //引き上げ条件一覧画面を生成し表示する。
                BtConditionDialog dialog = new BtConditionDialog();
                dialog.setVisible(true);
            }
        } //BtConditionButton

        //======================================================================
        /**
        *   引き上げ条件一覧画面
        */
        class BtConditionDialog extends JDialog {

            /**
             * コンストラクタ
             */
            BtConditionDialog(){
                super();
                //画面の体裁を整える。
                setTitle("引き上げ条件");
                setSize(850,250);
                setResizable(false);
                setModal(true);
                getContentPane().setLayout(null);
                // 他基地参照機能    @20131021
                if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                    getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
                }else{
                    getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
                }

                //引上げ条件一覧を生成する。
                BtConditionTable t = new BtConditionTable(ro_bt_all_condition);
                JTableHeader tabHead = t.getTableHeader();
                tabHead.setReorderingAllowed(false);
                //引上げ条件一覧をスクロールパネルに載せる。
                JScrollPane bt_scpanel = new JScrollPane(t);
                bt_scpanel.setBounds(20, 20, 810, 187);
                getContentPane().add(bt_scpanel);
            } //BtConditionDialog

            //==================================================================
            /**
            *       Ｂｔ引上げ条件情報一覧
            */
            class BtConditionTable extends JTable {

                private Vector  bt_list         = null;
                private BtConditionTblMdl model = null;

                /**
                * コンストラクタ
                */
                BtConditionTable(Vector v){
                    super();
                    bt_list = v;

                    try{
                        setName("BtConditionTable");
                        setAutoCreateColumnsFromModel(true);
                        setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
                        setLocale(new Locale("ja","JP"));
                        setFont(new java.awt.Font("dialog", 0, 12));
                        setRowHeight(17);

                        model = new BtConditionTblMdl(bt_list);
                        setModel(model);
                        DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                        TableColumn colum = null;
                        // No
                        colum = cmdl.getColumn(0);
                        colum.setMaxWidth(40);
                        colum.setMinWidth(40);
                        colum.setWidth(40);
                        // 登録日時
                        colum = cmdl.getColumn(1);
                        colum.setMaxWidth(160);
                        colum.setMinWidth(160);
                        colum.setWidth(160);
                        // 連番
                        colum = cmdl.getColumn(2);
                        colum.setMaxWidth(30);
                        colum.setMinWidth(30);
                        colum.setWidth(30);
                        // 品種
                        colum = cmdl.getColumn(3);
                        colum.setMaxWidth(80);
                        colum.setMinWidth(80);
                        colum.setWidth(80);
                        // ルツボ
                        colum = cmdl.getColumn(4);
                        colum.setMaxWidth(40);
                        colum.setMinWidth(40);
                        colum.setWidth(40);
                        // 直径
                        colum = cmdl.getColumn(5);
                        colum.setMaxWidth(40);
                        colum.setMinWidth(40);
                        colum.setWidth(40);
                        // 引上長
                        colum = cmdl.getColumn(6);
                        colum.setMaxWidth(40);
                        colum.setMinWidth(40);
                        colum.setWidth(40);
                        // 初仕込
                        colum = cmdl.getColumn(7);
                        colum.setMaxWidth(60);
                        colum.setMinWidth(60);
                        colum.setWidth(60);
                        // 追仕込
                        colum = cmdl.getColumn(8);
                        colum.setMaxWidth(60);
                        colum.setMinWidth(60);
                        colum.setWidth(60);
                        // T1
                        colum = cmdl.getColumn(9);
                        colum.setMaxWidth(30);
                        colum.setMinWidth(30);
                        colum.setWidth(30);
                        // T2
                        colum = cmdl.getColumn(10);
                        colum.setMaxWidth(30);
                        colum.setMinWidth(30);
                        colum.setWidth(30);
                        // T3
                        colum = cmdl.getColumn(11);
                        colum.setMaxWidth(30);
                        colum.setMinWidth(30);
                        colum.setWidth(30);
                        // T4
                        colum = cmdl.getColumn(12);
                        colum.setMaxWidth(30);
                        colum.setMinWidth(30);
                        colum.setWidth(30);
                        // T5
                        colum = cmdl.getColumn(13);
                        colum.setMaxWidth(30);
                        colum.setMinWidth(30);
                        colum.setWidth(30);
                        // T6       //@@
                        colum = cmdl.getColumn(13);
                        colum.setMaxWidth(30);
                        colum.setMinWidth(30);
                        colum.setWidth(30);
                        // PNo
                        colum = cmdl.getColumn(14);
                        colum.setMaxWidth(32);
                        colum.setMinWidth(32);
                        colum.setWidth(32);
                        // 開始
                        colum = cmdl.getColumn(15);
                        colum.setMaxWidth(30);
                        colum.setMinWidth(30);
                        colum.setWidth(30);
                    }
                    catch (Throwable e) {
                        CZSystem.handleException(e);
                    }
                }

                /**
                 *
                 */
                public void valueChanged(ListSelectionEvent e){
                    super.valueChanged(e);
                }

                /**
                 *
                 */
                public void setData(int gr,int tbl){
                    CZSystem.log("CZFpAveMain","BtConditionTable setData() [" + gr + "][" + tbl + "]");
                }

                //==============================================================
                /**
                *       Ｂｔ引上げ条件情報一覧：モデル
                */
                public class BtConditionTblMdl extends AbstractTableModel {

                    private int TBL_ROW     = 0;
                    final   int TBL_COL     = 17;
                    private Vector  bt_list = null;

                    final String[] names = {" # "  , "登録日時" , "連番" ,  
                            "品種" , "ルツボ"   , "直径" ,
                            "引上長" , "初仕込"   , "追仕込" ,
                            "T1" , "T2"   , "T3" ,
                            "T4" , "T5"   , "T6"   , "PNo" , "開始"
                            };

                    private Object  data[][];

                    /**
                    * コンストラクタ
                    */
                    BtConditionTblMdl(Vector v){
                        super();

                        bt_list = v;
                        TBL_ROW = bt_list.size();
                        data = new Object[TBL_ROW][TBL_COL];

                        for (int i = 0 ; i < TBL_ROW ; i++) {
                            CZSystemBt bt = (CZSystemBt)bt_list.elementAt(i);
                            if(null == bt) break;
                            data[i][0]  = new Integer(i+1);
                            data[i][1]  = bt.t_time;
                            data[i][2]  = new Integer(bt.renban);
                            data[i][3]  = bt.hinshu;
                            data[i][4]  = new Integer(bt.rutubo_kei);
                            data[i][5]  = new Integer(bt.chokkei);
                            data[i][6]  = new Integer(bt.hikiage_cho);
                            data[i][7]  = new Integer(bt.i_sikomi);
                            data[i][8]  = new Integer(bt.t_sikomi);
                            data[i][9]  = new Integer(bt.no_youkai);
                            data[i][10] = new Integer(bt.no_hikiage);
                            data[i][11] = new Integer(bt.no_kaiten);
                            data[i][12] = new Integer(bt.no_toridasi);
                            data[i][13] = new Integer(bt.no_aturyoku);
                            data[i][14] = new Integer(bt.no_teisu);         //@@
                            data[i][15] = new Integer(bt.pno_start);
                            data[i][16] = new Integer(bt.p_kaisi);
                        }
                    }

                    /**
                    * カラム数を取得する。
                    */
                    public int getColumnCount(){
                        return TBL_COL;
                    }
                    /**
                    * 行数を取得する。
                    */
                    public int getRowCount(){
                        return TBL_ROW;
                    }
                    /**
                    * 指定のセルのデータを取得する。
                    */
                    public Object getValueAt(int row, int col){
                        return data[row][col];
                    }
                    /**
                    * カラム名を取得する。
                    */
                    public String getColumnName(int column){
                        return names[column];
                    }
                    /**
                    * カラムの型を取得する。
                    */
                    public Class getColumnClass(int c){
                        return getValueAt(0, c).getClass();
                    }
                    /**
                    * セルの編集可否を取得する。
                    */
                    public boolean isCellEditable(int row, int col){
                        return false;
                    }
                    /**
                    * 指定のセルへデータを設定する。
                    */
                    public void setValueAt(Object aValue, int row, int column){
                        data[row][column] = aValue;
                    }
                } // BtConditionTblMdl
            } // BtConditionTable
        } // BtConditionDialog
    } // ConditionPanel

    //==========================================================================
    /**
    *   条件設定パネル
    */
    public class SetPanel extends JPanel {

        private NumText     ave_text        = null; //移動平均時間
        private ValText     umax_text       = null; //上上限
        private ValText     max_text        = null; //上限
        private ValText     lmin_text       = null; //下下限
        private ValText     min_text        = null; //下限

        private NumText     x_min_text      = null; //Ｘ軸最小
        private NumText     x_max_text      = null; //Ｘ軸最大
        private NumText     x_bun_text      = null; //Ｘ分割数
        private NumText     x_koushi_text   = null; //Ｘ格子間隔
        private NumText     x_memkan_text   = null; //Ｘ目盛値間
        private NumText     x_memketa_text  = null; //Ｘ目盛桁数
        private NumText     x_syouketa_text = null; //Ｘ小数桁数

        private ValText     y_min_text      = null; //Ｙ軸最小
        private ValText     y_max_text      = null; //Ｙ軸最大
        private NumText     y_bun_text      = null; //Ｙ分割数
        private NumText     y_koushi_text   = null; //Ｙ格子間隔
        private NumText     y_memkan_text   = null; //Ｙ目盛値間
        private NumText     y_memketa_text  = null; //Ｙ目盛桁数
        private NumText     y_syouketa_text = null; //Ｙ小数桁数

        private ValText     y_dia_min_text  = null; //Ｙ軸直径最小
        private ValText     y_dia_max_text  = null; //Ｙ軸直径最大
        private ValText     y_rpm_min_text  = null; //Ｙ軸回転最小
        private ValText     y_rpm_max_text  = null; //Ｙ軸回転最大

        //色
        private JButton     fp_ave_col_but  = null;
        private JButton     fp_umax_col_but = null;
        private JButton     fp_max_col_but  = null;
        private JButton     fp_lmin_col_but = null;
        private JButton     fp_min_col_but  = null;
        //色
        private JButton     umax_col_but    = null;
        private JButton     max_col_but     = null;
        private JButton     lmin_col_but    = null;
        private JButton     min_col_but     = null;
        //肩変え
        private JCheckBox   shld_shift_dia_chk_box  = null;
        private JCheckBox   shld_shift_chk_box      = null;
        private ValText     shld_shift_dia_text     = null; //直径
        private ValText     shld_shift_leng_text    = null; //位置

        //
        private JCheckBox   fp_pf_ave_chk_box   = null;
        private JCheckBox   fp_chk_box          = null;
        private JCheckBox   fp_pf_chk_box       = null;
        private JCheckBox   dia_chk_box         = null;
        private JCheckBox   dia_pf_chk_box      = null;
        private JCheckBox   sxl_rt_chk_box      = null;
        private JCheckBox   cru_rt_chk_box      = null;
        //色
        private JButton     fp_pf_ave_col_but   = null;
        private JButton     fp_col_but          = null;
        private JButton     fp_pf_col_but       = null;
        private JButton     dia_col_but         = null;
        private JButton     dia_pf_col_but      = null;
        private JButton     sxl_rt_col_but      = null;
        private JButton     cru_rt_col_but      = null;

        private JButton calc_button     = null; //計算ボタン
        private JButton graph_button    = null; //グラフボタン
        private JButton select_button   = null; //Ｓｅｌボタン
        private JButton save_button     = null; //保存ボタン
        private JButton load_button     = null; //読込ボタン

        /**
         * コンストラクタ
         */
        SetPanel(){
            super();
            
            addWindowListener(  
	            new WindowAdapter()
	            {
	                public void windowClosing(WindowEvent e) {
						FpAvePropSave();
						setVisible(false);
					    ro_bt_start   = null;     //検索用引き上げ条件
					    ro_bt_all_condition  = null;     //全Btの引き上げ条件
					    pv_data_body         = null;     //ボディーのデータ
					    calc_data_body       = null;     //ボディーの計算データ
					    serche_dia  = null;     //検索用ダイアログ
					    conpane     = null;     //検索、引き上げ条件パネル
					    setpane         = null;     //条件設定パネル
					    graph_set       = null;     //グラフ描画条件
					    datapane        = null;     //データパネル
					    graph_dia       = null;     //グラフ用ダイアログ
					}
	            }
            );

            setName("SetPanel");
            setLayout(null);
            setBorder(new Flush3DBorder());
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            calc_button = new JButton("計  算");
            calc_button.setBounds(20, 20, 80, 24);
            calc_button.setLocale(new Locale("ja","JP"));
            calc_button.setFont(new java.awt.Font("dialog", 0, 18));
            calc_button.setBorder(new Flush3DBorder());
            calc_button.setForeground(java.awt.Color.black);
            calc_button.addActionListener(new CalcButton());
            calc_button.setEnabled(false);
            add(calc_button);

            graph_button = new JButton("グラフ");
            graph_button.setBounds(120, 20, 80, 24);
            graph_button.setLocale(new Locale("ja","JP"));
            graph_button.setFont(new java.awt.Font("dialog", 0, 18));
            graph_button.setBorder(new Flush3DBorder());
            graph_button.setForeground(java.awt.Color.black);
            graph_button.addActionListener(new GraphButton());
            graph_button.setEnabled(false);
            add(graph_button);

            JLabel lab;
            //平均時間
            lab = new JLabel("移動平均時間",JLabel.LEFT);
            lab.setBounds(20, 64, 100, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            add(lab);

            //(s)
            lab = new JLabel("(秒)",JLabel.LEFT);
            lab.setBounds(202, 64, 36, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            add(lab);

            //上上限
            lab = new JLabel("上上限",JLabel.LEFT);
            lab.setBounds(20, 120, 100, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            add(lab);

            //上限
            lab = new JLabel("上限",JLabel.LEFT);
            lab.setBounds(20, 140, 100, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            add(lab);

            //下限
            lab = new JLabel("下限",JLabel.LEFT);
            lab.setBounds(20, 160, 100, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            add(lab);

            //下下限
            lab = new JLabel("下下限",JLabel.LEFT);
            lab.setBounds(20, 180, 100, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            add(lab);

            //平均時間
            ave_text = new NumText();
            ave_text.setBounds(120, 64, 80, 18);
            ave_text.setLocale(new Locale("ja","JP"));
            ave_text.setFont(new java.awt.Font("dialog", 0, 12));
            ave_text.setBorder(new Flush3DBorder());
            ave_text.setForeground(java.awt.Color.black);
            ave_text.setText(fp_ave_time_pro);
            add(ave_text);

            //上上限
            umax_text = new ValText();
            umax_text.setBounds(120, 120, 80, 18);
            umax_text.setLocale(new Locale("ja","JP"));
            umax_text.setFont(new java.awt.Font("dialog", 0, 12));
            umax_text.setBorder(new Flush3DBorder());
            umax_text.setForeground(java.awt.Color.black);
            umax_text.setText(pf_umax_pro);
            add(umax_text);

            //上限
            max_text = new ValText();
            max_text.setBounds(120, 140, 80, 18);
            max_text.setLocale(new Locale("ja","JP"));
            max_text.setFont(new java.awt.Font("dialog", 0, 12));
            max_text.setBorder(new Flush3DBorder());
            max_text.setForeground(java.awt.Color.black);
            max_text.setText(pf_max_pro);
            add(max_text);

            //下限
            min_text = new ValText();
            min_text.setBounds(120, 160, 80, 18);
            min_text.setLocale(new Locale("ja","JP"));
            min_text.setFont(new java.awt.Font("dialog", 0, 12));
            min_text.setBorder(new Flush3DBorder());
            min_text.setForeground(java.awt.Color.black);
            min_text.setText(pf_min_pro);
            add(min_text);

            //下下限
            lmin_text = new ValText();
            lmin_text.setBounds(120, 180, 80, 18);
            lmin_text.setLocale(new Locale("ja","JP"));
            lmin_text.setFont(new java.awt.Font("dialog", 0, 12));
            lmin_text.setBorder(new Flush3DBorder());
            lmin_text.setForeground(java.awt.Color.black);
            lmin_text.setText(pf_lmin_pro);
            add(lmin_text);

            //上上限
            Color c = java.awt.Color.cyan;
            umax_col_but = new JButton();
            umax_col_but.setBounds(202, 120, 18, 18);
            umax_col_but.setBorder(new Flush3DBorder());
            umax_col_but.setForeground(c);
            umax_col_but.setBackground(c);
            umax_col_but.addActionListener(new ColorSetButton());
            add(umax_col_but);

            //上限
            c = java.awt.Color.blue;
            max_col_but = new JButton();
            max_col_but.setBounds(202, 140, 18, 18);
            max_col_but.setBorder(new Flush3DBorder());
            max_col_but.setForeground(c);
            max_col_but.setBackground(c);
            max_col_but.addActionListener(new ColorSetButton());
            add(max_col_but);

            //下限
            c = java.awt.Color.blue;
            min_col_but = new JButton();
            min_col_but.setBounds(202, 160, 18, 18);
            min_col_but.setBorder(new Flush3DBorder());
            min_col_but.setForeground(c);
            min_col_but.setBackground(c);
            min_col_but.addActionListener(new ColorSetButton());
            add(min_col_but);

            //下下限
            c = java.awt.Color.cyan;
            lmin_col_but = new JButton();
            lmin_col_but.setBounds(202, 180, 18, 18);
            lmin_col_but.setBorder(new Flush3DBorder());
            lmin_col_but.setForeground(c);
            lmin_col_but.setBackground(c);
            lmin_col_but.addActionListener(new ColorSetButton());
            add(lmin_col_but);

            //上上限
            c = java.awt.Color.red;
            fp_umax_col_but = new JButton();
            fp_umax_col_but.setBounds(222, 110, 18, 18);
            fp_umax_col_but.setBorder(new Flush3DBorder());
            fp_umax_col_but.setForeground(c);
            fp_umax_col_but.setBackground(c);
            fp_umax_col_but.addActionListener(new ColorSetButton());
            add(fp_umax_col_but);

            //上限
            c = java.awt.Color.yellow;
            fp_max_col_but = new JButton();
            fp_max_col_but.setBounds(222, 130, 18, 18);
            fp_max_col_but.setBorder(new Flush3DBorder());
            fp_max_col_but.setForeground(c);
            fp_max_col_but.setBackground(c);
            fp_max_col_but.addActionListener(new ColorSetButton());
            add(fp_max_col_but);

            //FP
            c = java.awt.Color.green;
            fp_ave_col_but = new JButton();
            fp_ave_col_but.setBounds(222, 150, 18, 18);
            fp_ave_col_but.setBorder(new Flush3DBorder());
            fp_ave_col_but.setForeground(c);
            fp_ave_col_but.setBackground(c);
            fp_ave_col_but.addActionListener(new ColorSetButton());
            add(fp_ave_col_but);

            //下限
            c = java.awt.Color.yellow;
            fp_min_col_but = new JButton();
            fp_min_col_but.setBounds(222, 170, 18, 18);
            fp_min_col_but.setBorder(new Flush3DBorder());
            fp_min_col_but.setForeground(c);
            fp_min_col_but.setBackground(c);
            fp_min_col_but.addActionListener(new ColorSetButton());
            add(fp_min_col_but);

            //下下限
            c = java.awt.Color.red;
            fp_lmin_col_but = new JButton();
            fp_lmin_col_but.setBounds(222, 190, 18, 18);
            fp_lmin_col_but.setBorder(new Flush3DBorder());
            fp_lmin_col_but.setForeground(c);
            fp_lmin_col_but.setBackground(c);
            fp_lmin_col_but.addActionListener(new ColorSetButton());
            add(fp_lmin_col_but);

            //肩変え直径
            shld_shift_dia_chk_box = new JCheckBox("肩変え直径");
            shld_shift_dia_chk_box.setBounds(20, 240, 100, 18);
            shld_shift_dia_chk_box.setFont(new java.awt.Font("dialog", 0, 14));
            shld_shift_dia_chk_box.setForeground(java.awt.Color.black);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                shld_shift_dia_chk_box.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                shld_shift_dia_chk_box.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }
            shld_shift_dia_chk_box.setSelected(true);
            add(shld_shift_dia_chk_box);

            shld_shift_dia_text = new ValText();
            shld_shift_dia_text.setBounds(120, 240, 80, 18);
            shld_shift_dia_text.setLocale(new Locale("ja","JP"));
            shld_shift_dia_text.setFont(new java.awt.Font("dialog", 0, 12));
            shld_shift_dia_text.setBorder(new Flush3DBorder());
            shld_shift_dia_text.setForeground(java.awt.Color.black);
            shld_shift_dia_text.setText(shld_shift_dia);
            add(shld_shift_dia_text);

            //肩変え位置
            shld_shift_chk_box = new JCheckBox("肩変え位置");
            shld_shift_chk_box.setBounds(20, 260, 100, 18);
            shld_shift_chk_box.setFont(new java.awt.Font("dialog", 0, 14));
            shld_shift_chk_box.setForeground(java.awt.Color.black);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                shld_shift_chk_box.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                shld_shift_chk_box.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }
            shld_shift_chk_box.setSelected(true);
            add(shld_shift_chk_box);

            shld_shift_leng_text = new ValText();
            shld_shift_leng_text.setBounds(120, 260, 80, 18);
            shld_shift_leng_text.setLocale(new Locale("ja","JP"));
            shld_shift_leng_text.setFont(new java.awt.Font("dialog", 0, 12));
            shld_shift_leng_text.setBorder(new Flush3DBorder());
            shld_shift_leng_text.setForeground(java.awt.Color.black);
            shld_shift_leng_text.setText(shld_shift_length);
            add(shld_shift_leng_text);

            //肩変え位置をテーブルからセレクト
            select_button = new JButton("SELECT");
            select_button.setBounds(202, 260, 38, 18);
            select_button.setLocale(new Locale("ja","JP"));
            select_button.setFont(new java.awt.Font("dialog", 0, 10));
            select_button.setBorder(new Flush3DBorder());
            select_button.setForeground(java.awt.Color.black);
            select_button.addActionListener(new SelectShldIndexButton());
            add(select_button);
//@@@
            //設定保存
            save_button = new JButton("保  存");
            save_button.setBounds(300, 300, 80, 24);
            save_button.setLocale(new Locale("ja","JP"));
            save_button.setFont(new java.awt.Font("dialog", 0, 18));
            save_button.setBorder(new Flush3DBorder());
            save_button.setForeground(java.awt.Color.black);
            save_button.addActionListener(
              new ActionListener() {
                  public void actionPerformed(ActionEvent evt)
                  {
                      JFileChooser chooser = new JFileChooser(file_);
                      int ret = chooser.showSaveDialog(setpane);
                      if (ret == JFileChooser.APPROVE_OPTION) {
                                                                                // ファイル名取得
                          file_ = chooser.getSelectedFile();
                                                                                // プロパティ作成
                          Properties prop = new Properties();
                          prop.setProperty(new String("FP_AVE_TIME"),
                                            new String("" + ave_text.getText()) );  //移動平均時間
                          prop.setProperty(new String("FP_PF_UMAX"),
                                            new String("" + umax_text.getText()) ); //上上限
                          prop.setProperty(new String("FP_PF_MAX"),
                                            new String("" + max_text.getText()) );  //上限
                          prop.setProperty(new String("FP_PF_LMIN"),
                                            new String("" + lmin_text.getText()) ); //下下限
                          prop.setProperty(new String("FP_PF_MIN"),
                                            new String("" + min_text.getText()) );  //下限

                          prop.setProperty(new String("SHLD_SHIFT_DIA"),
                                            new String("" + shld_shift_dia_text.getText()) );   //肩変え直径
                          prop.setProperty(new String("SHLD_SHIFT_LENGTH"),
                                            new String("" + shld_shift_leng_text.getText()) );  //肩変え位置

                          //Ｘ軸
                          prop.setProperty(new String("X_LENGTH_MIN"),
                                            new String("" + x_min_text.getText()) );        //Ｘ軸最小値
                          prop.setProperty(new String("X_LENGTH_MAX"),
                                            new String("" + x_max_text.getText()) );        //Ｘ軸最大値
                          prop.setProperty(new String("X_LENGTH_BUNKATU"),
                                            new String("" + x_bun_text.getText()) );        //Ｘ軸分割数
                          prop.setProperty(new String("X_LENGTH_KOUSHI"),
                                            new String("" + x_koushi_text.getText()) );     //Ｘ軸格子間隔
                          prop.setProperty(new String("X_LENGTH_MEMKAN"),
                                            new String("" + x_memkan_text.getText()) );     //Ｘ軸目盛値間隔
                          prop.setProperty(new String("X_LENGTH_MEMKETA"),
                                            new String("" + x_memketa_text.getText()) );    //Ｘ軸目盛桁数
                          prop.setProperty(new String("X_LENGTH_SYOUKETA"),
                                            new String("" + x_syouketa_text.getText()) );   //Ｘ軸小数桁数

                          //Ｙ軸
                          prop.setProperty(new String("SXL_ST_MIN"),
                                            new String("" + y_min_text.getText()) );        //Ｙ軸引上速度最小値
                          prop.setProperty(new String("SXL_ST_MAX"),
                                            new String("" + y_max_text.getText()) );        //Ｙ軸引上速度最大値
                          prop.setProperty(new String("SXL_ST_BUNKATU"),
                                            new String("" + y_bun_text.getText()) );        //Ｙ軸分割
                          prop.setProperty(new String("SXL_ST_KOUSHI"),
                                            new String("" + y_koushi_text.getText()) );     //Ｙ軸格子間隔
                          prop.setProperty(new String("SXL_ST_MEMKAN"),
                                            new String("" + y_memkan_text.getText()) );     //Ｙ軸目盛値間隔
                          prop.setProperty(new String("SXL_ST_MEMKETA"),
                                            new String("" + y_memketa_text.getText()) );    //Ｙ軸目盛桁数
                          prop.setProperty(new String("SXL_ST_SYOUKETA"),
                                            new String("" + y_syouketa_text.getText()) );   //Ｙ軸小数桁数

                          prop.setProperty(new String("DIA_MIN"),
                                            new String("" + y_dia_min_text.getText()) );    //Ｙ軸直径最小
                          prop.setProperty(new String("DIA_MAX"),
                                            new String("" + y_dia_max_text.getText()) );    //Ｙ軸直径最大
                          prop.setProperty(new String("SXL_RT_PF_MIN"),
                                            new String("" + y_rpm_min_text.getText()) );    //Ｙ軸回転最小
                          prop.setProperty(new String("SXL_RT_PF_MAX"),
                                            new String("" + y_rpm_max_text.getText()) );    //Ｙ軸回転最大

                          prop.setProperty(new String("DIA_PF_MIN"),
                                            new String("" + dia_pf_min_pro) );              //直径プロファイル
                          prop.setProperty(new String("DIA_PF_MAX"),
                                            new String("" + dia_pf_max_pro) );              //直径プロファイル
                          // ファイルに保存
                          try {
                              FileOutputStream out = new FileOutputStream(file_);
                              prop.store(out, "");
                              out.flush();
                              out.close();
                          }
                          catch (IOException ex) {
                              JOptionPane.showMessageDialog(
                                setpane,
                                new String("保存できませんでした。"),
                                new String("保存"),
                                JOptionPane.WARNING_MESSAGE);
                              return;
                          }
                          JOptionPane.showMessageDialog(
                            setpane,
                            new String("保存しました。"),
                            new String("保存"),
                            JOptionPane.INFORMATION_MESSAGE);
                          return;
                      }
                  }
              }
            );
            save_button.setEnabled(true);
            add(save_button);

            //設定読込
            load_button = new JButton("読  込");
            load_button.setBounds(400, 300, 80, 24);
            load_button.setLocale(new Locale("ja","JP"));
            load_button.setFont(new java.awt.Font("dialog", 0, 18));
            load_button.setBorder(new Flush3DBorder());
            load_button.setForeground(java.awt.Color.black);
            load_button.addActionListener(
              new ActionListener() {
                  public void actionPerformed(ActionEvent evt) {

                      JFileChooser chooser = new JFileChooser(file_);
                      if ( chooser.showOpenDialog(setpane) == JFileChooser.APPROVE_OPTION ) {
                          file_ = chooser.getSelectedFile();            // ファイル名取得
                          Properties prop =  new Properties();          // プロパティ作成
                          try{
                              FileInputStream in = new FileInputStream(file_);
                              prop.load(in);
                              in.close();

                              fp_ave_time_pro   = prop.getProperty("FP_AVE_TIME");      //移動平均時間
                              pf_umax_pro       = prop.getProperty("FP_PF_UMAX");       //プロファイルの上上限
                              pf_max_pro        = prop.getProperty("FP_PF_MAX");        //プロファイルの上限
                              pf_lmin_pro       = prop.getProperty("FP_PF_LMIN");       //プロファイルの下下限
                              pf_min_pro        = prop.getProperty("FP_PF_MIN");        //プロファイルの下限

                              shld_shift_dia    = prop.getProperty("SHLD_SHIFT_DIA");   //肩変え直径
                              shld_shift_length = prop.getProperty("SHLD_SHIFT_LENGTH");//肩変え位置
                              //Ｘ軸
                              x_length_min      = prop.getProperty("X_LENGTH_MIN");     //Ｘ軸最小値
                              x_length_max      = prop.getProperty("X_LENGTH_MAX");     //Ｘ軸最大値
                              x_length_bunkatu  = prop.getProperty("X_LENGTH_BUNKATU"); //Ｘ軸分割数
                              x_length_koushi   = prop.getProperty("X_LENGTH_KOUSHI");  //Ｘ軸格子間隔
                              x_length_memkan   = prop.getProperty("X_LENGTH_MEMKAN");  //Ｘ軸目盛値間隔
                              x_length_memketa  = prop.getProperty("X_LENGTH_MEMKETA"); //Ｘ軸目盛桁数
                              x_length_syouketa = prop.getProperty("X_LENGTH_SYOUKETA");//Ｘ軸小数桁数
                              //Ｙ軸
                              sxl_st_min_pro    = prop.getProperty("SXL_ST_MIN");       //Ｙ軸引上速度最小値
                              sxl_st_max_pro    = prop.getProperty("SXL_ST_MAX");       //Ｙ軸引上速度最大値
                              sxl_st_bunkatu    = prop.getProperty("SXL_ST_BUNKATU");   //Ｙ軸分割
                              sxl_st_koushi     = prop.getProperty("SXL_ST_KOUSHI");    //Ｙ軸格子間隔
                              sxl_st_memkan     = prop.getProperty("SXL_ST_MEMKAN");    //Ｙ軸目盛値間隔
                              sxl_st_memketa    = prop.getProperty("SXL_ST_MEMKETA");   //Ｙ軸目盛桁数
                              sxl_st_syouketa   = prop.getProperty("SXL_ST_SYOUKETA");  //Ｙ軸小数桁数
                              dia_min_pro       = prop.getProperty("DIA_MIN");          //直径最小値
                              dia_max_pro       = prop.getProperty("DIA_MAX");          //直径最大値
                              sxl_rt_pf_min_pro = prop.getProperty("SXL_RT_PF_MIN");    //シード回転プロファイル最小値
                              sxl_rt_pf_max_pro = prop.getProperty("SXL_RT_PF_MAX");    //シード回転プロファイル最大値

                              dia_pf_min_pro    = prop.getProperty("DIA_PF_MIN");       //直径プロファイル
                              dia_pf_max_pro    = prop.getProperty("DIA_PF_MAX");

                              setPropertiesToText();        //読込み値を画面へ設定する。
                          } catch ( IOException ex ) {
                              CZSystem.log("CZFpAveMain ","Property Fileがロードできませんでした。");
                              return;
                          }
                      }
                  }
              }
            );
            load_button.setEnabled(true);
            add(load_button);
//@@@

            //fpプロファイル移動平均
            fp_pf_ave_chk_box = new JCheckBox("fpプロファイル移動平均");
            fp_pf_ave_chk_box.setBounds(280, 64, 180, 18);
            fp_pf_ave_chk_box.setFont(new java.awt.Font("dialog", 0, 14));
            fp_pf_ave_chk_box.setForeground(java.awt.Color.black);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                fp_pf_ave_chk_box.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                fp_pf_ave_chk_box.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }
            fp_pf_ave_chk_box.setSelected(true);
            add(fp_pf_ave_chk_box);

            //fp実績
            fp_chk_box = new JCheckBox("fp実績");
            fp_chk_box.setBounds(280, 84, 180, 18);
            fp_chk_box.setFont(new java.awt.Font("dialog", 0, 14));
            fp_chk_box.setForeground(java.awt.Color.black);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                fp_chk_box.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                fp_chk_box.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }
            add(fp_chk_box);

            //fpプロファイル
            fp_pf_chk_box = new JCheckBox("fpプロファイル");
            fp_pf_chk_box.setBounds(280, 104, 180, 18);
            fp_pf_chk_box.setFont(new java.awt.Font("dialog", 0, 14));
            fp_pf_chk_box.setForeground(java.awt.Color.black);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                fp_pf_chk_box.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                fp_pf_chk_box.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }
            add(fp_pf_chk_box);

            //直径
            dia_chk_box = new JCheckBox("直径");
            dia_chk_box.setBounds(280, 124, 180, 18);
            dia_chk_box.setFont(new java.awt.Font("dialog", 0, 14));
            dia_chk_box.setForeground(java.awt.Color.black);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                dia_chk_box.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                dia_chk_box.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }
            dia_chk_box.setSelected(true);
            add(dia_chk_box);

            //直径プロファイル
            dia_pf_chk_box = new JCheckBox("直径プロファイル");
            dia_pf_chk_box.setBounds(280, 144, 180, 18);
            dia_pf_chk_box.setFont(new java.awt.Font("dialog", 0, 14));
            dia_pf_chk_box.setForeground(java.awt.Color.black);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                dia_pf_chk_box.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                dia_pf_chk_box.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }
            dia_pf_chk_box.setSelected(true);
            add(dia_pf_chk_box);

            //結晶回転
            sxl_rt_chk_box = new JCheckBox("結晶回転");
            sxl_rt_chk_box.setBounds(280, 164, 180, 18);
            sxl_rt_chk_box.setFont(new java.awt.Font("dialog", 0, 14));
            sxl_rt_chk_box.setForeground(java.awt.Color.black);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                sxl_rt_chk_box.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                sxl_rt_chk_box.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }
            add(sxl_rt_chk_box);

            //ルツボ回転
            cru_rt_chk_box = new JCheckBox("ルツボ回転");
            cru_rt_chk_box.setBounds(280, 184, 180, 18);
            cru_rt_chk_box.setFont(new java.awt.Font("dialog", 0, 14));
            cru_rt_chk_box.setForeground(java.awt.Color.black);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                cru_rt_chk_box.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                cru_rt_chk_box.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }
            add(cru_rt_chk_box);

            //色
            c = new Color(255,153,51);
            fp_pf_ave_col_but = new JButton();
            fp_pf_ave_col_but.setBounds(460, 64, 18, 18);
            fp_pf_ave_col_but.setBorder(new Flush3DBorder());
            fp_pf_ave_col_but.setForeground(c);
            fp_pf_ave_col_but.setBackground(c);
            fp_pf_ave_col_but.addActionListener(new ColorSetButton());
            add(fp_pf_ave_col_but);

            //
            c = new Color(203,0,204);
            fp_col_but = new JButton();
            fp_col_but.setBounds(460, 84, 18, 18);
            fp_col_but.setBorder(new Flush3DBorder());
            fp_col_but.setForeground(c);
            fp_col_but.setBackground(c);
            fp_col_but.addActionListener(new ColorSetButton());
            add(fp_col_but);

            //
            c = new Color(153,153,0);
            fp_pf_col_but = new JButton();
            fp_pf_col_but.setBounds(460, 104, 18, 18);
            fp_pf_col_but.setBorder(new Flush3DBorder());
            fp_pf_col_but.setForeground(c);
            fp_pf_col_but.setBackground(c);
            fp_pf_col_but.addActionListener(new ColorSetButton());
            add(fp_pf_col_but);

            //
            c = new Color(153,204,255);
            dia_col_but = new JButton();
            dia_col_but.setBounds(460, 124, 18, 18);
            dia_col_but.setBorder(new Flush3DBorder());
            dia_col_but.setForeground(c);
            dia_col_but.setBackground(c);
            dia_col_but.addActionListener(new ColorSetButton());
            add(dia_col_but);

            //
            c = new Color(255,203,50);
            dia_pf_col_but = new JButton();
            dia_pf_col_but.setBounds(460, 144, 18, 18);
            dia_pf_col_but.setBorder(new Flush3DBorder());
            dia_pf_col_but.setForeground(c);
            dia_pf_col_but.setBackground(c);
            dia_pf_col_but.addActionListener(new ColorSetButton());
            add(dia_pf_col_but);

            //
            c = new Color(204,255,204);
            sxl_rt_col_but = new JButton();
            sxl_rt_col_but.setBounds(460, 164, 18, 18);
            sxl_rt_col_but.setBorder(new Flush3DBorder());
            sxl_rt_col_but.setForeground(c);
            sxl_rt_col_but.setBackground(c);
            sxl_rt_col_but.addActionListener(new ColorSetButton());
            add(sxl_rt_col_but);

            //
            c = new Color(0,153,153);
            cru_rt_col_but = new JButton();
            cru_rt_col_but.setBounds(460, 184, 18, 18);
            cru_rt_col_but.setBorder(new Flush3DBorder());
            cru_rt_col_but.setForeground(c);
            cru_rt_col_but.setBackground(c);
            cru_rt_col_but.addActionListener(new ColorSetButton());
            add(cru_rt_col_but);

            //X軸パネル
            JPanel p = null;
            p = new JPanel();
            p.setBounds(500, 20, 440, 140);
            p.setLayout(null);
            p.setBorder(BorderFactory.createTitledBorder(new Flush3DBorder(),"Ｘ軸"));
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                p.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                p.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }
            add(p);

            //最小値
            lab = new JLabel("最小値",JLabel.LEFT);
            lab.setBounds(20, 20, 60, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            p.add(lab);

            //最大値
            lab = new JLabel("最大値",JLabel.LEFT);
            lab.setBounds(20, 40, 60, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            p.add(lab);

            //最小値
            x_min_text = new NumText();
            x_min_text.setBounds(80, 20, 50, 18);
            x_min_text.setLocale(new Locale("ja","JP"));
            x_min_text.setFont(new java.awt.Font("dialog", 0, 12));
            x_min_text.setBorder(new Flush3DBorder());
            x_min_text.setForeground(java.awt.Color.black);
            x_min_text.setText(x_length_min);
            p.add(x_min_text);

            //最大値
            x_max_text = new NumText();
            x_max_text.setBounds(80, 40, 50, 18);
            x_max_text.setLocale(new Locale("ja","JP"));
            x_max_text.setFont(new java.awt.Font("dialog", 0, 12));
            x_max_text.setBorder(new Flush3DBorder());
            x_max_text.setForeground(java.awt.Color.black);
            x_max_text.setText(x_length_max);
            p.add(x_max_text);

            //分割数
            lab = new JLabel("分割数",JLabel.LEFT);
            lab.setBounds(150, 20, 80, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            p.add(lab);

            //格子間隔
            lab = new JLabel("格子間隔",JLabel.LEFT);
            lab.setBounds(150, 40, 80, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            p.add(lab);

            //目盛値間隔
            lab = new JLabel("目盛値間隔",JLabel.LEFT);
            lab.setBounds(150, 60, 80, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            p.add(lab);

            //目盛桁数
            lab = new JLabel("目盛桁数",JLabel.LEFT);
            lab.setBounds(150, 80, 80, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            p.add(lab);

            //小数桁数
            lab = new JLabel("小数桁数",JLabel.LEFT);
            lab.setBounds(150, 100, 80, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            p.add(lab);

            //分割数
            x_bun_text = new NumText();
            x_bun_text.setBounds(230, 20, 50, 18);
            x_bun_text.setLocale(new Locale("ja","JP"));
            x_bun_text.setFont(new java.awt.Font("dialog", 0, 12));
            x_bun_text.setBorder(new Flush3DBorder());
            x_bun_text.setForeground(java.awt.Color.black);
            x_bun_text.setText(x_length_bunkatu);
            p.add(x_bun_text);

            //格子間隔
            x_koushi_text = new NumText();
            x_koushi_text.setBounds(230, 40, 50, 18);
            x_koushi_text.setLocale(new Locale("ja","JP"));
            x_koushi_text.setFont(new java.awt.Font("dialog", 0, 12));
            x_koushi_text.setBorder(new Flush3DBorder());
            x_koushi_text.setForeground(java.awt.Color.black);
            x_koushi_text.setText(x_length_koushi);     //@@@
            p.add(x_koushi_text);

            //目盛値間隔
            x_memkan_text = new NumText();
            x_memkan_text.setBounds(230, 60, 50, 18);
            x_memkan_text.setLocale(new Locale("ja","JP"));
            x_memkan_text.setFont(new java.awt.Font("dialog", 0, 12));
            x_memkan_text.setBorder(new Flush3DBorder());
            x_memkan_text.setForeground(java.awt.Color.black);
            x_memkan_text.setText(x_length_memkan);     //@@@
            p.add(x_memkan_text);

            //目盛桁数
            x_memketa_text = new NumText();
            x_memketa_text.setBounds(230, 80, 50, 18);
            x_memketa_text.setLocale(new Locale("ja","JP"));
            x_memketa_text.setFont(new java.awt.Font("dialog", 0, 12));
            x_memketa_text.setBorder(new Flush3DBorder());
            x_memketa_text.setForeground(java.awt.Color.black);
            x_memketa_text.setText(x_length_memketa);       //@@@
            p.add(x_memketa_text);

            //小数桁数
            x_syouketa_text = new NumText();
            x_syouketa_text.setBounds(230, 100, 50, 18);
            x_syouketa_text.setLocale(new Locale("ja","JP"));
            x_syouketa_text.setFont(new java.awt.Font("dialog", 0, 12));
            x_syouketa_text.setBorder(new Flush3DBorder());
            x_syouketa_text.setForeground(java.awt.Color.black);
            x_syouketa_text.setText(x_length_syouketa);     //@@@
            p.add(x_syouketa_text);

            //Y軸パネル
            p = new JPanel();
            p.setBounds(500, 180, 440, 150);
            p.setLayout(null);
            p.setBorder(BorderFactory.createTitledBorder(new Flush3DBorder(),"Ｙ軸"));
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                p.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                p.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }
            add(p);

            //最小値
            lab = new JLabel("最小値",JLabel.LEFT);
            lab.setBounds(20, 20, 60, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            p.add(lab);

            //最大値
            lab = new JLabel("最大値",JLabel.LEFT);
            lab.setBounds(20, 40, 60, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            p.add(lab);

            //最小値
            y_min_text = new ValText();
            y_min_text.setBounds(80, 20, 50, 18);
            y_min_text.setLocale(new Locale("ja","JP"));
            y_min_text.setFont(new java.awt.Font("dialog", 0, 12));
            y_min_text.setBorder(new Flush3DBorder());
            y_min_text.setForeground(java.awt.Color.black);
            y_min_text.setText(sxl_st_min_pro);
            p.add(y_min_text);

            //最大値
            y_max_text = new ValText();
            y_max_text.setBounds(80, 40, 50, 18);
            y_max_text.setLocale(new Locale("ja","JP"));
            y_max_text.setFont(new java.awt.Font("dialog", 0, 12));
            y_max_text.setBorder(new Flush3DBorder());
            y_max_text.setForeground(java.awt.Color.black);
            y_max_text.setText(sxl_st_max_pro);
            p.add(y_max_text);

            //分割数
            lab = new JLabel("分割数",JLabel.LEFT);
            lab.setBounds(150, 20, 80, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            p.add(lab);

            //格子間隔
            lab = new JLabel("格子間隔",JLabel.LEFT);
            lab.setBounds(150, 40, 80, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            p.add(lab);

            //目盛値間隔
            lab = new JLabel("目盛値間隔",JLabel.LEFT);
            lab.setBounds(150, 60, 80, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            p.add(lab);

            //目盛桁数
            lab = new JLabel("目盛桁数",JLabel.LEFT);
            lab.setBounds(150, 80, 80, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            p.add(lab);

            //小数桁数
            lab = new JLabel("小数桁数",JLabel.LEFT);
            lab.setBounds(150, 100, 80, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            p.add(lab);

            //分割数
            y_bun_text = new NumText();
            y_bun_text.setBounds(230, 20, 50, 18);
            y_bun_text.setLocale(new Locale("ja","JP"));
            y_bun_text.setFont(new java.awt.Font("dialog", 0, 12));
            y_bun_text.setBorder(new Flush3DBorder());
            y_bun_text.setForeground(java.awt.Color.black);
            y_bun_text.setText(sxl_st_bunkatu);
            p.add(y_bun_text);

            //格子間隔
            y_koushi_text = new NumText();
            y_koushi_text.setBounds(230, 40, 50, 18);
            y_koushi_text.setLocale(new Locale("ja","JP"));
            y_koushi_text.setFont(new java.awt.Font("dialog", 0, 12));
            y_koushi_text.setBorder(new Flush3DBorder());
            y_koushi_text.setForeground(java.awt.Color.black);
            y_koushi_text.setText(sxl_st_koushi);       //@@@
            p.add(y_koushi_text);

            //目盛値間隔
            y_memkan_text = new NumText();
            y_memkan_text.setBounds(230, 60, 50, 18);
            y_memkan_text.setLocale(new Locale("ja","JP"));
            y_memkan_text.setFont(new java.awt.Font("dialog", 0, 12));
            y_memkan_text.setBorder(new Flush3DBorder());
            y_memkan_text.setForeground(java.awt.Color.black);
            y_memkan_text.setText(sxl_st_memkan);       //@@@
            p.add(y_memkan_text);

            //目盛桁数
            y_memketa_text = new NumText();
            y_memketa_text.setBounds(230, 80, 50, 18);
            y_memketa_text.setLocale(new Locale("ja","JP"));
            y_memketa_text.setFont(new java.awt.Font("dialog", 0, 12));
            y_memketa_text.setBorder(new Flush3DBorder());
            y_memketa_text.setForeground(java.awt.Color.black);
            y_memketa_text.setText(sxl_st_memketa);     //@@@
            p.add(y_memketa_text);

            //小数桁数
            y_syouketa_text = new NumText();
            y_syouketa_text.setBounds(230, 100, 50, 18);
            y_syouketa_text.setLocale(new Locale("ja","JP"));
            y_syouketa_text.setFont(new java.awt.Font("dialog", 0, 12));
            y_syouketa_text.setBorder(new Flush3DBorder());
            y_syouketa_text.setForeground(java.awt.Color.black);
            y_syouketa_text.setText(sxl_st_syouketa);       //@@@
            p.add(y_syouketa_text);

            //直径
            lab = new JLabel("直径",JLabel.LEFT);
            lab.setBounds(310, 20, 60, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            p.add(lab);

            //最小値
            lab = new JLabel("最小値",JLabel.LEFT);
            lab.setBounds(310, 40, 60, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            p.add(lab);

            //最大値
            lab = new JLabel("最大値",JLabel.LEFT);
            lab.setBounds(310, 60, 60, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            p.add(lab);

            //回転
            lab = new JLabel("回転",JLabel.LEFT);
            lab.setBounds(310, 80, 60, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            p.add(lab);

            //最小値
            lab = new JLabel("最小値",JLabel.LEFT);
            lab.setBounds(310, 100, 60, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            p.add(lab);

            //最大値
            lab = new JLabel("最大値",JLabel.LEFT);
            lab.setBounds(310, 120, 60, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setForeground(java.awt.Color.black);
            p.add(lab);

            //直径 最小値
            y_dia_min_text = new ValText();
            y_dia_min_text.setBounds(370, 40, 50, 18);
            y_dia_min_text.setLocale(new Locale("ja","JP"));
            y_dia_min_text.setFont(new java.awt.Font("dialog", 0, 12));
            y_dia_min_text.setBorder(new Flush3DBorder());
            y_dia_min_text.setForeground(java.awt.Color.black);
            y_dia_min_text.setText(dia_min_pro);
            p.add(y_dia_min_text);

            //直径 最大値
            y_dia_max_text = new ValText();
            y_dia_max_text.setBounds(370, 60, 50, 18);
            y_dia_max_text.setLocale(new Locale("ja","JP"));
            y_dia_max_text.setFont(new java.awt.Font("dialog", 0, 12));
            y_dia_max_text.setBorder(new Flush3DBorder());
            y_dia_max_text.setForeground(java.awt.Color.black);
            y_dia_max_text.setText(dia_max_pro);
            p.add(y_dia_max_text);

            //回転 最小値
            y_rpm_min_text = new ValText();
            y_rpm_min_text.setBounds(370, 100, 50, 18);
            y_rpm_min_text.setLocale(new Locale("ja","JP"));
            y_rpm_min_text.setFont(new java.awt.Font("dialog", 0, 12));
            y_rpm_min_text.setBorder(new Flush3DBorder());
            y_rpm_min_text.setForeground(java.awt.Color.black);
            y_rpm_min_text.setText(sxl_rt_pf_min_pro);
            p.add(y_rpm_min_text);

            //回転 最大値
            y_rpm_max_text = new ValText();
            y_rpm_max_text.setBounds(370, 120, 50, 18);
            y_rpm_max_text.setLocale(new Locale("ja","JP"));
            y_rpm_max_text.setFont(new java.awt.Font("dialog", 0, 12));
            y_rpm_max_text.setBorder(new Flush3DBorder());
            y_rpm_max_text.setForeground(java.awt.Color.black);
            y_rpm_max_text.setText(sxl_rt_pf_max_pro);
            p.add(y_rpm_max_text);

        } //SetPanel
//@@@

		private void FpAvePropSave()
		{
			Properties prop = new Properties();			// プロパティ作成
			prop.setProperty(new String("FP_AVE_TIME"),
			                  new String("" + ave_text.getText()) );  //移動平均時間
			prop.setProperty(new String("FP_PF_UMAX"),
			                  new String("" + umax_text.getText()) ); //上上限
			prop.setProperty(new String("FP_PF_MAX"),
			                  new String("" + max_text.getText()) );  //上限
			prop.setProperty(new String("FP_PF_LMIN"),
			                  new String("" + lmin_text.getText()) ); //下下限
			prop.setProperty(new String("FP_PF_MIN"),
			                  new String("" + min_text.getText()) );  //下限

			prop.setProperty(new String("SHLD_SHIFT_DIA"),
			                  new String("" + shld_shift_dia_text.getText()) );   //肩変え直径
			prop.setProperty(new String("SHLD_SHIFT_LENGTH"),
			                  new String("" + shld_shift_leng_text.getText()) );  //肩変え位置

			//Ｘ軸
			prop.setProperty(new String("X_LENGTH_MIN"),
			                  new String("" + x_min_text.getText()) );        //Ｘ軸最小値
			prop.setProperty(new String("X_LENGTH_MAX"),
			                  new String("" + x_max_text.getText()) );        //Ｘ軸最大値
			prop.setProperty(new String("X_LENGTH_BUNKATU"),
			                  new String("" + x_bun_text.getText()) );        //Ｘ軸分割数
			prop.setProperty(new String("X_LENGTH_KOUSHI"),
			                  new String("" + x_koushi_text.getText()) );     //Ｘ軸格子間隔
			prop.setProperty(new String("X_LENGTH_MEMKAN"),
			                  new String("" + x_memkan_text.getText()) );     //Ｘ軸目盛値間隔
			prop.setProperty(new String("X_LENGTH_MEMKETA"),
			                  new String("" + x_memketa_text.getText()) );    //Ｘ軸目盛桁数
			prop.setProperty(new String("X_LENGTH_SYOUKETA"),
			                  new String("" + x_syouketa_text.getText()) );   //Ｘ軸小数桁数

			//Ｙ軸
			prop.setProperty(new String("SXL_ST_MIN"),
			                  new String("" + y_min_text.getText()) );        //Ｙ軸引上速度最小値
			prop.setProperty(new String("SXL_ST_MAX"),
			                  new String("" + y_max_text.getText()) );        //Ｙ軸引上速度最大値
			prop.setProperty(new String("SXL_ST_BUNKATU"),
			                  new String("" + y_bun_text.getText()) );        //Ｙ軸分割
			prop.setProperty(new String("SXL_ST_KOUSHI"),
			                  new String("" + y_koushi_text.getText()) );     //Ｙ軸格子間隔
			prop.setProperty(new String("SXL_ST_MEMKAN"),
			                  new String("" + y_memkan_text.getText()) );     //Ｙ軸目盛値間隔
			prop.setProperty(new String("SXL_ST_MEMKETA"),
			                  new String("" + y_memketa_text.getText()) );    //Ｙ軸目盛桁数
			prop.setProperty(new String("SXL_ST_SYOUKETA"),
			                  new String("" + y_syouketa_text.getText()) );   //Ｙ軸小数桁数

			prop.setProperty(new String("DIA_MIN"),
			                  new String("" + y_dia_min_text.getText()) );    //Ｙ軸直径最小
			prop.setProperty(new String("DIA_MAX"),
			                  new String("" + y_dia_max_text.getText()) );    //Ｙ軸直径最大
			prop.setProperty(new String("SXL_RT_PF_MIN"),
			                  new String("" + y_rpm_min_text.getText()) );    //Ｙ軸回転最小
			prop.setProperty(new String("SXL_RT_PF_MAX"),
			                  new String("" + y_rpm_max_text.getText()) );    //Ｙ軸回転最大

			prop.setProperty(new String("DIA_PF_MIN"),
			                  new String("" + dia_pf_min_pro) );              //直径プロファイル
			prop.setProperty(new String("DIA_PF_MAX"),
			                  new String("" + dia_pf_max_pro) );              //直径プロファイル
			// ファイルに保存
			try {
//				CZSystem.log("CZFpAveMain","ファイルに保存した。");
//			    FileOutputStream out = new FileOutputStream("d:/CZ/classes/CZFPAVEPROPERTY.TXT");
			    FileOutputStream out = new FileOutputStream(CZSystemDefine.FPAVEPROPERTY_FILE);
			    prop.store(out, "");
			    out.flush();
			    out.close();
			}
			catch (IOException ex) {
			    JOptionPane.showMessageDialog(
			      setpane,
			      new String("保存できませんでした。"),
			      new String("保存"),
			      JOptionPane.WARNING_MESSAGE);
			    return;
			}
		}	


        /**
         * プロパティファイルから読込んだ設定を画面に設定する。
         */
        private void setPropertiesToText(){
            ave_text.setText(fp_ave_time_pro);                  //平均時間
            umax_text.setText(pf_umax_pro);                     //上上限
            max_text.setText(pf_max_pro);                       //上限
            min_text.setText(pf_min_pro);                       //下限
            lmin_text.setText(pf_lmin_pro);                     //下下限

            shld_shift_dia_text.setText(shld_shift_dia);        //肩変え直径
            shld_shift_leng_text.setText(shld_shift_length);    //肩変え位置
            //X軸パネル
            x_min_text.setText(x_length_min);                   //最小値
            x_max_text.setText(x_length_max);                   //最大値
            x_bun_text.setText(x_length_bunkatu);               //分割数
            x_koushi_text.setText(x_length_koushi);             //格子間隔
            x_memkan_text.setText(x_length_memkan);             //目盛値間隔
            x_memketa_text.setText(x_length_memketa);           //目盛桁数
            x_syouketa_text.setText(x_length_syouketa);         //小数桁数
            //Y軸パネル
            y_min_text.setText(sxl_st_min_pro);                 //最小値
            y_max_text.setText(sxl_st_max_pro);                 //最大値
            y_bun_text.setText(sxl_st_bunkatu);                 //分割数
            y_koushi_text.setText(sxl_st_koushi);               //格子間隔
            y_memkan_text.setText(sxl_st_memkan);               //目盛値間隔
            y_memketa_text.setText(sxl_st_memketa);             //目盛桁数
            y_syouketa_text.setText(sxl_st_syouketa);           //小数桁数
            y_dia_min_text.setText(dia_min_pro);                //直径 最小値
            y_dia_max_text.setText(dia_max_pro);                //直径 最大値
            y_rpm_min_text.setText(sxl_rt_pf_min_pro);          //回転 最小値
            y_rpm_max_text.setText(sxl_rt_pf_max_pro);          //回転 最大値
            return;
        }
//@@@
        /**
         * ボタンの初期設定をする。
         */
        public boolean setDefault(){
            setMode(false);
            return true;
        }

        /**
         *  ボタンのモードを変更する。
         */
        public boolean setMode(boolean b){

            calc_button.setEnabled(b);
            graph_button.setEnabled(false);
            select_button.setEnabled(b);
            return true;
        }

        //======================================================================
        /**
        *   計算ボタンの処理
        */
        class CalcButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
//                CZSystem.log("CZFpAveMain","SetPanel CalcButton ");

                int     time;
                float   umax;
                float   max;
                float   lmin;
                float   min;
		
		FpAvePropSave();
		
                try{
                    time = Integer.parseInt(ave_text.getText());
                    umax = Float.parseFloat(umax_text.getText());
                    max  = Float.parseFloat(max_text.getText());
                    lmin = Float.parseFloat(lmin_text.getText());
                    min  = Float.parseFloat(min_text.getText());
                }
                catch(NumberFormatException e){
                    Object msg[] = {"移動平均時間",
                                    "上上限、上限、下下限、下限",
                                    "数値範囲異常"};
                    errorMsg(msg);
                    return ;
                }

                //          time    umax    max     lmin     min
                //startCalc(4800  , 0.05f , 0.03f , -0.05f , -0.03f);
                startCalc(time  , umax , max , lmin , min);

                if(shld_shift_dia_chk_box.isSelected()){
                    float shld_length = sercheShldLength();
                    shld_shift_leng_text.setText("" + shld_length);
                    float shld_dia    = sercheShldDia();
                    shld_shift_dia_text.setText("" + shld_dia);
                }
                graph_button.setEnabled(true);
            }
        } //CalcButton

        //======================================================================
        /**
        *   グラフボタンの処理
        */
        class GraphButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
//                CZSystem.log("CZFpAveMain","SetPanel GraphButton ");

                graph_set   = null;                 //グラフ設定データをクリアする。
                GraphSet s  = new GraphSet();       //グラフ設定データ保持領域を確保する。
                //肩変え位置のチェック
                s.shld_shift = shld_shift_chk_box.isSelected();
                if(s.shld_shift){
                    try{
                        s.shld_shift_val = Float.parseFloat(shld_shift_leng_text.getText());
                    }
                    catch(NumberFormatException e){
                        Object msg[] = {"肩変え位置の数値を",
                                        "入れてください。",
                                        ""};
                        errorMsg(msg);
                        return ;
                    }
                }

                //Ｘ軸のチェック
                try{
                    s.x_min         = Float.parseFloat(x_min_text.getText());
                    s.x_max         = Float.parseFloat(x_max_text.getText());
                    s.x_bun         = Integer.parseInt(x_bun_text.getText());
                    s.x_koushi      = Integer.parseInt(x_koushi_text.getText());
                    s.x_memkan      = Integer.parseInt(x_memkan_text.getText());
                    s.x_memketa     = Integer.parseInt(x_memketa_text.getText());
                    s.x_syouketa    = Integer.parseInt(x_syouketa_text.getText());
                }
                catch(NumberFormatException e){
                    Object msg[] = {"Ｘ軸を設定してください。",
                                    "数値に変換できません。",
                                    ""};
                    errorMsg(msg);
                    return ;
                }

                if(1 > s.x_bun){
                    Object msg[] = {"Ｘ軸分割数は",
                                    "１以上を設定してください。",
                                    ""};
                    errorMsg(msg);
                    return ;
                }

                if(s.x_min >= s.x_max){
                    Object msg[] = {"Ｘ軸最小値、最大値が",
                                    "同じもしくは反対です。",
                                    ""};
                    errorMsg(msg);
                    return ;
                }

                if(2 > s.x_memketa){
                    Object msg[] = {"Ｘ軸目盛桁数は",
                                    "２以上を設定してください。",
                                    ""};
                    errorMsg(msg);
                    return ;
                }

                if(2 > (s.x_memketa - s.x_syouketa)){
                    Object msg[] = {"Ｘ軸目盛桁数と小数桁数に",
                                    "矛盾が有ります。",
                                    "目盛桁数大きくするか小数桁数を小さくしてください。"};
                    errorMsg(msg);
                    return ;
                }


                //Ｙ軸のチェック
                try{
                    s.y_min         = Float.parseFloat(y_min_text.getText());
                    s.y_max         = Float.parseFloat(y_max_text.getText());
                    s.y_bun         = Integer.parseInt(y_bun_text.getText());
                    s.y_koushi      = Integer.parseInt(y_koushi_text.getText());
                    s.y_memkan      = Integer.parseInt(y_memkan_text.getText());
                    s.y_memketa     = Integer.parseInt(y_memketa_text.getText());
                    s.y_syouketa    = Integer.parseInt(y_syouketa_text.getText());

                    s.y_dia_min     = Float.parseFloat(y_dia_min_text.getText());
                    s.y_dia_max     = Float.parseFloat(y_dia_max_text.getText());
                    s.y_rpm_min     = Float.parseFloat(y_rpm_min_text.getText());
                    s.y_rpm_max     = Float.parseFloat(y_rpm_max_text.getText());
                }
                catch(NumberFormatException e){
                    Object msg[] = {"Ｙ軸を設定してください。",
                                    "数値に変換できません。",
                                    ""};
                    errorMsg(msg);
                    return ;
                }

                if(1 > s.y_bun){
                    Object msg[] = {"Ｙ軸分割数は",
                                    "１以上を設定してください。",
                                    ""};
                    errorMsg(msg);
                    return ;
                }


                if(s.y_min >= s.y_max){
                    Object msg[] = {"Ｙ軸最小値、最大値が",
                                    "同じもしくは反対です。",
                                    ""};
                    errorMsg(msg);
                    return ;
                }

                if(2 > s.y_memketa){
                    Object msg[] = {"Ｙ軸目盛桁数は",
                                    "２以上を設定してください。",
                                    ""};
                    errorMsg(msg);
                    return ;
                }

                if(2 > (s.y_memketa - s.y_syouketa)){
                    Object msg[] = {"Ｙ軸目盛桁数と小数桁数に",
                                    "矛盾が有ります。",
                                    "目盛桁数大きくするか小数桁数を小さくしてください。"};
                    errorMsg(msg);
                    return ;
                }

                if(s.y_dia_min >= s.y_dia_max){
                    Object msg[] = {"Ｙ軸直径の最小値、最大値が",
                                    "同じもしくは反対です。",
                                    ""};
                    errorMsg(msg);
                    return ;
                }

                if(s.y_rpm_min >= s.y_rpm_max){
                    Object msg[] = {"Ｙ軸回転の最小値、最大値が",
                                    "同じもしくは反対です。",
                                    ""};
                    errorMsg(msg);
                    return ;
                }

                //色情報
                s.fp_umax_col       = umax_col_but.getBackground();
                s.fp_max_col        = max_col_but.getBackground();
                s.fp_min_col        = min_col_but.getBackground();
                s.fp_lmin_col       = lmin_col_but.getBackground();

                s.fp_umax_over_col  = fp_umax_col_but.getBackground();
                s.fp_max_over_col   = fp_max_col_but.getBackground();
                s.fp_center_col     = fp_ave_col_but.getBackground();
                s.fp_min_over_col   = fp_min_col_but.getBackground();
                s.fp_lmin_over_col  = fp_lmin_col_but.getBackground();

                s.fp_pf_ave_draw_col = fp_pf_ave_col_but.getBackground();
                s.fp_draw_col        = fp_col_but.getBackground();
                s.fp_pf_draw_col     = fp_pf_col_but.getBackground();
                s.dia_draw_col       = dia_col_but.getBackground();
                s.dia_pf_draw_col    = dia_pf_col_but.getBackground();
                s.sxl_rpm_draw_col   = sxl_rt_col_but.getBackground();
                s.cru_rpm_draw_col   = cru_rt_col_but.getBackground();

                //描画をするか？
                s.fp_pf_ave_draw    = fp_pf_ave_chk_box.isSelected();
                s.fp_draw           = fp_chk_box.isSelected();
                s.fp_pf_draw        = fp_pf_chk_box.isSelected();
                s.dia_draw          = dia_chk_box.isSelected();
                s.dia_pf_draw       = dia_pf_chk_box.isSelected();
                s.sxl_rpm_draw      = sxl_rt_chk_box.isSelected();
                s.cru_rpm_draw      = cru_rt_chk_box.isSelected();

                s.shld_shift        = shld_shift_chk_box.isSelected();

				FpAvePropSave();

                graph_set           = s;        //グラフ設定データを保持する。

                gph_cnt = CZSystem.GraphCount();
                if(gph_cnt > 4){
                    Object msg[] = { "グラフは５枚以上開けません", "", "" };
                    errorMsg(msg);
					return;
				}else{
                    //グラフ表示用ダイアログ
                    graph_dia = new CZFpAveGraphFrame(main_ro_name_lab.getText(),fp_ave_calc_time,ro_bt_all_condition,pv_data_body,calc_data_body,graph_set);

                    graph_dia.setDefault();         //グラフを描画する。
                    graph_dia.setVisible(true);     //グラフ画面を表示する。
                    CZSystem.GraphCountUp();
                }
            }
        } //GraphButton

        //======================================================================
        /**
        *   肩変え位置の検索ボタン
        */
        class SelectShldIndexButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                selectShldLengthIndex();
            }
        } //SelectShldIndexButton

        //======================================================================
        /*
        *   色変えボタン
        */
        class ColorSetButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
//                CZSystem.log("CZFpAveMain","SetPanel ColorSetButton ");

                JButton but = (JButton)ev.getSource();
                Color c = JColorChooser.showDialog(null,"色を選んでください", but.getBackground());
                if(null != c){
                    but.setForeground(c);
                    but.setBackground(c);
                }
            }
        } //ColorSetButton

        //======================================================================
        /*
        *       数値を入力するTextField
        */
        public class ValText extends JTextField {

            ValText(){
                super();
            }

            /**
             *
             */
            protected Document createDefaultModel() {
                return new NumericDocument();
            }

            //==================================================================
            /**
             *
             */
            class NumericDocument extends PlainDocument {
                String validValues = "0123456789.-";

                /**
                 *
                 */
                public void insertString( int offset, String str, AttributeSet a )
                    throws BadLocationException {

                    if(9 < getLength()) return;
                    char[] val = str.toCharArray();
                    for (int i = 0; i < val.length; i++) {
                        if(validValues.indexOf(val[i]) == -1) return;
                    }
                    super.insertString( offset, str, a );
                }

                /**
                 *
                 */
                public void remove(int offs, int len)
                    throws BadLocationException {
                    super.remove(offs,len);
                }
            }
        } //ValText

        //======================================================================
        /*
        *       数値を入力するTextField（小数無し）
        */
        public class NumText extends JTextField {

            NumText(){
                super();
            }

            /**
             *
             */
            protected Document createDefaultModel() {
                return new NumericDocument();
            }
            //==================================================================
            /**
             *
             */
            class NumericDocument extends PlainDocument {
                String validValues = "0123456789";

                /**
                 *
                 */
                public void insertString( int offset, String str, AttributeSet a )
                    throws BadLocationException {

                    if(9 < getLength()) return;
                    char[] val = str.toCharArray();
                    for (int i = 0; i < val.length; i++) {
                        if(validValues.indexOf(val[i]) == -1) return;
                    }
                    super.insertString( offset, str, a );
                }

                /**
                 *
                 */
                public void remove(int offs, int len)
                    throws BadLocationException {
                    super.remove(offs,len);
                }
            }
        } //NumText
    } //SetPanel


    //==========================================================================
    /*
    *   データパネル
    */
    public class DataPanel extends JPanel {

        private JScrollPane data_scpanel    = null;
        private JScrollPane calc_scpanel    = null;
        private JTextField  count_text      = null;

        /**
         *
         */
        DataPanel(){
            super();
            setName("DataPanel");
            setLayout(null);
            setBorder(new Flush3DBorder());
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            JLabel lab = null;
            //検索件数
            lab = new JLabel("検索件数",JLabel.CENTER);
            lab.setBounds(20, 10, 80, 18);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            add(lab);

            //データ件数表示
            count_text = new JTextField();
            count_text.setBounds(100, 10, 80, 18);
            count_text.setLocale(new Locale("ja","JP"));
            count_text.setFont(new java.awt.Font("dialog", 0, 12));
            count_text.setBorder(new Flush3DBorder());
            count_text.setForeground(java.awt.Color.black);
            add(count_text);

            //実績データ表示
            data_scpanel = new JScrollPane();
//@@            data_scpanel.setBounds(10, 35, 780, 445);
            data_scpanel.setBounds(10, 35, 780, 400);
            add(data_scpanel);

            //計算結果データ表示
            calc_scpanel = new JScrollPane();
//@@            calc_scpanel.setBounds(800, 35, 260, 445);
            calc_scpanel.setBounds(800, 35, 260, 400);
            add(calc_scpanel);
        }

        /**
         *
         */
        public boolean setDefault(){

            JViewport v;
            v =  data_scpanel.getViewport();
            if(null != v.getView()) v.remove(v.getView());
            v =  calc_scpanel.getViewport();
            if(null != v.getView()) v.remove(v.getView());
            count_text.setText("");
            return true;
        }

        /**
         * 検索件数と実績データ一覧を表示する。
         */
        public boolean setData(){
            setDefault();
            //検索件数
            if(null == pv_data_body) return false;
            if(1 > pv_data_body.size()) return false;
            count_text.setText("" + pv_data_body.size());
            //実績データ一覧
            PVDataTable t = new PVDataTable(pv_data_body);
            JTableHeader tabHead = t.getTableHeader();
            tabHead.setReorderingAllowed(false);
            data_scpanel.setViewportView(t);
            return true;
        }

        /**
         *  計算結果の表示（テーブル）
         */
        public boolean setCalc(){
            JViewport v;
            v =  calc_scpanel.getViewport();
            if(null != v.getView()) v.remove(v.getView());

            if(null == calc_data_body) return false;
            if(1 > calc_data_body.size()) return false;

            //計算結果データ表示
            CalcDataTable t = new CalcDataTable(calc_data_body);
            JTableHeader tabHead = t.getTableHeader();
            tabHead.setReorderingAllowed(false);
            calc_scpanel.setViewportView(t);
            return true;
        }

        /**
         *  指定されたインデックスの選択
         */
        public void selectData(int index){
            JViewport v;
            Rectangle cellRect;

            v =  data_scpanel.getViewport();
            PVDataTable view = (PVDataTable)v.getView();
            cellRect = view.getCellRect(index,0,false);
            if(cellRect != null){
                view.scrollRectToVisible(cellRect);
            }
            view.setRowSelectionInterval(index,index);

            v =  calc_scpanel.getViewport();
            CalcDataTable t = (CalcDataTable)v.getView();
            if(null != t){
                cellRect = t.getCellRect(index,0,false);
                if(cellRect != null){
                    t.scrollRectToVisible(cellRect);
                }
                t.setRowSelectionInterval(index,index);
            }
        }

        //======================================================================
        /**
        *       実績データ一覧用のテーブル
        */
        class PVDataTable extends JTable {

            private Vector  data_list  = null;
            private PVDataTblMdl model = null;
            private boolean life = false;

            /**
            * コンストラクタ
            */
            PVDataTable(Vector v){
                super();
                data_list = v;

                try{
                    setName("PVDataTable");
                    setAutoCreateColumnsFromModel(true);
                    setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
                    setLocale(new Locale("ja","JP"));
                    setFont(new java.awt.Font("dialog", 0, 12));
                    setRowHeight(17);

                    model = new PVDataTblMdl();
                    setModel(model);

                    DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                    TableColumn     colum = null;
                    // No
                    colum = cmdl.getColumn(0);
                    colum.setMaxWidth(60);
                    colum.setMinWidth(60);
                    colum.setWidth(60);
                    // PNo
                    colum = cmdl.getColumn(1);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);
                    // SPNo
                    colum = cmdl.getColumn(2);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);
                    // PSeq
                    colum = cmdl.getColumn(3);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);
                    // プロセス時間
                    colum = cmdl.getColumn(4);
                    colum.setMaxWidth(70);
                    colum.setMinWidth(70);
                    colum.setWidth(70);
                    // サブプロセス時間
                    colum = cmdl.getColumn(5);
                    colum.setMaxWidth(70);
                    colum.setMinWidth(70);
                    colum.setWidth(70);
                    // 採取日時
                    colum = cmdl.getColumn(6);
                    colum.setMaxWidth(160);
                    colum.setMinWidth(160);
                    colum.setWidth(160);
                    // 引き上げ長さ
                    colum = cmdl.getColumn(7);
                    colum.setMaxWidth(70);
                    colum.setMinWidth(70);
                    colum.setWidth(70);
                    // DIA
                    colum = cmdl.getColumn(8);
                    colum.setMaxWidth(70);
                    colum.setMinWidth(70);
                    colum.setWidth(70);
                    // SXLS.ST
                    colum = cmdl.getColumn(9);
                    colum.setMaxWidth(70);
                    colum.setMinWidth(70);
                    colum.setWidth(70);
                    // SXLS.PF
                    colum = cmdl.getColumn(10);
                    colum.setMaxWidth(70);
                    colum.setMinWidth(70);
                    colum.setWidth(70);
                }
                catch (Throwable e) {
                    CZSystem.handleException(e);
                }
            }

            /**
            *
            */
            public void valueChanged(ListSelectionEvent e){
                super.valueChanged(e);
                if(e.getValueIsAdjusting()) return;
            }
            /**
            *
            */
            public void setData(int gr,int tbl){
            }

            //==================================================================
            /**
            *       実績データ一覧：モデル
            */
            public class PVDataTblMdl extends AbstractTableModel {

                private int TBL_ROW             = 0;
                final   int TBL_COL             = 11;

                final String[] names = {" # "     , "PNo",
                                        "SPNo"    , "PSeq",
                                        "PTime"   , "SPTime",
                                        "採取日時", "L" , "DIA" ,
                                        "SXL.ST"  , "SXS.PF"
                                         };

                private Object  data[][];

                /**
                * コンストラクタ
                */
                PVDataTblMdl(){
                    super();
                    TBL_ROW = data_list.size();
                    data = new Object[TBL_ROW][TBL_COL];

                    for (int i = 0 ; i < TBL_ROW ; i++) {
                        CZSystemPVData st = (CZSystemPVData)data_list.elementAt(i);
                        if(null == st) break;
                        data[i][0] = new Integer(i+1);
                        data[i][1] = new Integer(st.p_no);
                        data[i][2] = new Integer(st.sp_no);
                        data[i][3] = new Integer(st.p_renban);
                        data[i][4] = new Integer(st.p_time);
                        data[i][5] = new Integer(st.sp_time);
                        data[i][6] = st.p_date;
                        data[i][7] = new Float(st.p_length);
                        data[i][8] = new Float(st.data[DIA]);
                        data[i][9] = new Float(st.data[SXL_ST]);
                        data[i][10] = new Float(st.data[SXL_ST_PF]);
                    }
                }

                /**
                * カラム数を取得する。
                */
                public int getColumnCount(){
                    return TBL_COL;
                }
                /**
                * 行数を取得する。
                */
                public int getRowCount(){
                    return TBL_ROW;
                }
                /**
                * 指定のセルのデータを取得する。
                */
                public Object getValueAt(int row, int col){
                    return data[row][col];
                }
                /**
                * カラム名を取得する。
                */
                public String getColumnName(int column){
                    return names[column];
                }
                /**
                * カラムの型を取得する。
                */
                public Class getColumnClass(int c){
                    return getValueAt(0, c).getClass();
                }
                /**
                * セルの編集可否を取得する。
                */
                public boolean isCellEditable(int row, int col){
                    return false;
                }
                /**
                * 指定のセルへデータを設定する。
                */
                public void setValueAt(Object aValue, int row, int column){
                    data[row][column] = aValue;
                }
            } //PVDataTblMdl
        } //PVDataTable

        //======================================================================
        /*
        *       計算結果データ表示用のテーブル
        */
        class CalcDataTable extends JTable {

            private Vector  data_list  = null;
            private CalcDataTblMdl model = null;
            private boolean life = false;

            /**
            *
            */
            CalcDataTable(Vector v){
                super();
                data_list = v;

                try{
                    setName("CalcDataTable");
                    setAutoCreateColumnsFromModel(true);
                    setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
                    setLocale(new Locale("ja","JP"));
                    setFont(new java.awt.Font("dialog", 0, 12));
                    setRowHeight(17);

                    model = new CalcDataTblMdl();
                    setModel(model);

                    DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                    TableColumn     colum = null;
                    // No
                    colum = cmdl.getColumn(0);
                    colum.setMaxWidth(60);
                    colum.setMinWidth(60);
                    colum.setWidth(60);
                    // fp-Ave
                    colum = cmdl.getColumn(1);
                    colum.setMaxWidth(60);
                    colum.setMinWidth(60);
                    colum.setWidth(60);
                    // PF-Ave
                    colum = cmdl.getColumn(2);
                    colum.setMaxWidth(60);
                    colum.setMinWidth(60);
                    colum.setWidth(60);
                    // 判定
                    colum = cmdl.getColumn(3);
                    colum.setMaxWidth(60);
                    colum.setMinWidth(60);
                    colum.setWidth(60);
                }
                catch (Throwable e) {
                    CZSystem.handleException(e);
                }
            }

            /**
             *
             */
            public void valueChanged(ListSelectionEvent e){
                super.valueChanged(e);
                if(e.getValueIsAdjusting()) return;
            }
            /**
             *
             */
            public void setData(int gr,int tbl){
            }

            //==================================================================
            /*
            *       計算結果データ表示：モデル
            */
            public class CalcDataTblMdl extends AbstractTableModel {

                private int TBL_ROW             = 0;
                final   int TBL_COL             = 4;

                final String[] names = {" # "     , "fp-Ave",
                                        "PF-Ave"  , "判定"
                                     };

                private Object  data[][];

                /**
                * コンストラクタ
                */
                CalcDataTblMdl(){
                    super();

                    TBL_ROW = data_list.size();
                    data = new Object[TBL_ROW][TBL_COL];

                    for (int i = 0 ; i < TBL_ROW ; i++) {
                        CalcData st = (CalcData)data_list.elementAt(i);
                        if(null == st) break;
                        data[i][0] = new Integer(i+1);
                        data[i][1] = new Float(st.fp_ave);
                        data[i][2] = new Float(st.pf_ave);

                        String tmp;
                        switch(st.judg){
                            case  0 : tmp = "合格";
                                  break;
                            case -1 : tmp = "下下限";
                                  break;
                            case -2 : tmp = "下限";
                                  break;
                            case  1 : tmp = "上限";
                                  break;
                            case  2 : tmp = "上上限";
                                  break;
                            default : tmp = "計算不可";
                                  break;
                        }
                        data[i][3] = tmp;
                    } // for end
                }

                /**
                *
                */
                public int getColumnCount(){
                    return TBL_COL;
                }
                /**
                *
                */
                public int getRowCount(){
                    return TBL_ROW;
                }
                /**
                *
                */
                public Object getValueAt(int row, int col){
                    return data[row][col];
                }
                /**
                *
                */
                public String getColumnName(int column){
                    return names[column];
                }
                /**
                *
                */
                public Class getColumnClass(int c){
                    return getValueAt(0, c).getClass();
                }
                /**
                *
                */
                public boolean isCellEditable(int row, int col){
                    return false;
                }
                /**
                *
                */
                public void setValueAt(Object aValue, int row, int column){
                    data[row][column] = aValue;
                }
            } //CalcDataTblMdl
        } //CalcDataTable
    } //DataPanel


    //==========================================================================
    /*
    *   Ｂｔデータ検索用ダイアログ
    */
    class SercheDialog extends JDialog {

        private JScrollPane bt_scpanel          = null;
        private JScrollPane bt_start_scpanel    = null;
        private JButton     read_button         = null;
        private JLabel      ro_name_lab         = null;

        /**
         * コンストラクタ
         */
        SercheDialog(){
            super();

            setTitle("検  索");
// chg start 2008.09.17
//            setSize(820,335);
            setSize(940,335);
// chg end 2008.09.17
            setResizable(false);
            setModal(true);
            getContentPane().setLayout(null);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            //炉名称
			String s = CZSystem.RoKetaChg(ro_name);	// 20050725 炉：表示桁数変更
            ro_name_lab = new JLabel(s,JLabel.CENTER);
//            ro_name_lab = new JLabel(ro_name,JLabel.CENTER);
            ro_name_lab.setBounds(20, 20, 100, 30);
            ro_name_lab.setLocale(new Locale("ja","JP"));
            ro_name_lab.setFont(new java.awt.Font("dialog", 0, 18));
            ro_name_lab.setBorder(new Flush3DBorder());
            ro_name_lab.setForeground(java.awt.Color.black);
            getContentPane().add(ro_name_lab);
            //バッチ一覧表
            bt_scpanel = new JScrollPane();
// chg start 2008.09.17
//            bt_scpanel.setBounds(20, 60, 350, 187);
            bt_scpanel.setBounds(20, 60, 470, 187);
// chg end 2008.09.17
            getContentPane().add(bt_scpanel);
            //バッチ開始情報
            bt_start_scpanel = new JScrollPane();
// chg start 2008.09.17
//            bt_start_scpanel.setBounds(390, 60, 410, 187);
            bt_start_scpanel.setBounds(510, 60, 410, 187);
// chg end 2008.09.17
            getContentPane().add(bt_start_scpanel);
            //データ読込みボタン
            read_button = new JButton("読み込み");
// chg start 2008.09.17
//            read_button.setBounds(700, 270, 100, 24);
            read_button.setBounds(820, 270, 100, 24);
// chg end 2008.09.17
            read_button.setLocale(new Locale("ja","JP"));
            read_button.setFont(new java.awt.Font("dialog", 0, 18));
            read_button.setBorder(new Flush3DBorder());
            read_button.setForeground(java.awt.Color.black);
            read_button.addActionListener(new ReadButton());
            read_button.setEnabled(false);
            getContentPane().add(read_button);
//            CZSystem.log("CZFpAveMain","SercheDialog new");
        }

        /**
        * 画面の初期設定をする。
        */
        public boolean setDefault(){
            removeBtStart();
            removeBtCondition();
            
            String s = CZSystem.RoKetaChg(ro_name);	// 20050725 炉：表示桁数変更
            ro_name_lab.setText(s);
//            ro_name_lab.setText(ro_name);
            BtTable t = new BtTable();
            JTableHeader tabHead = t.getTableHeader();
            tabHead.setReorderingAllowed(false);
            bt_scpanel.setViewportView(t);
            read_button.setEnabled(false);
            return true;
        }

        /**
        * バッチ開始一覧を設定する。
        */
        public boolean setBtCondition(Vector v){
            removeBtCondition();
            BtStartTable t = new BtStartTable(v);
            JTableHeader tabHead = t.getTableHeader();
            tabHead.setReorderingAllowed(false);
            bt_start_scpanel.setViewportView(t);
            ro_bt_all_condition = v;
            return true;
        }

        /**
        * バッチ開始情報をクリアする。
        */
        public boolean removeBtCondition(){
            JViewport v;
            v =  bt_start_scpanel.getViewport();
            if(null != v.getView()) v.remove(v.getView());
            removeBtStart();
            read_button.setEnabled(false);
            return true;
        }

        //======================================================================
        /**
        *   読込みボタンの処理
        */
        class ReadButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
//                CZSystem.log("CZFpAveMain","SercheDialog ReadButton actionPerformed");

                pv_data_body    = null;
                calc_data_body  = null;
                datapane.setDefault();

                Cursor cu_tmp = getCur();                   //現状のカーソルを保持する。
                Cursor cu = new Cursor(Cursor.WAIT_CURSOR);
                setCur(cu);                                 //カーソルを砂時計に変える。
                int ret = readBtPV();                       //PVデータを読込む。
                setCur(cu_tmp);                             //カーソルを元に戻す。

                if(1 > ret){                                //データ無の時
                    setpane.setMode(false);
                    conpane.setData(false);
                    return;
                }

                serche_dia.setVisible(false);               //検索画面を閉じる。
                datapane.setData();                         //画面にデータを設定する。
                setpane.setMode(true);
                conpane.setData(true);
            }
        } //ReadButton

        //======================================================================
        /**
        *       ＢｔＮｏ一覧テーブル
        */
        class BtTable extends JTable {

            private Vector  bt_all_list = null;
            private Vector  bt_list     = null;
            private BtTblMdl model      = null;
            private boolean life        = false;

            /**
            * コンストラクタ
            */
			@SuppressWarnings("unchecked")
            BtTable(){
                super();

                try{
                    setName("BtTable");
                    setAutoCreateColumnsFromModel(true);
                    setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
                    setLocale(new Locale("ja","JP"));
                    setFont(new java.awt.Font("dialog", 0, 12));
                    setRowHeight(17);

                    bt_all_list = CZSystem.getBtCondition(ro_db_name);
                    if(null == bt_all_list) return;

                    bt_list = new Vector(100);
                    for (int i = 0 ; i < bt_all_list.size() ; i++) {
                        CZSystemBt bt = (CZSystemBt)bt_all_list.elementAt(i);

                        if(0 == bt.renban) bt_list.addElement(bt);
//@@2003.09.18                        if(-1 == bt.renban) bt_list.addElement(bt);
                    }

                    model = new BtTblMdl(bt_list);
                    setModel(model);

                    DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                    TableColumn     colum = null;
                    // No
                    colum = cmdl.getColumn(0);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);
                    // BtNo
                    colum = cmdl.getColumn(1);
                    colum.setMaxWidth(130);
                    colum.setMinWidth(130);
                    colum.setWidth(130);
// chg start 2008.09.17
//                    // 登録日時
//                    colum = cmdl.getColumn(2);
//                    colum.setMaxWidth(162);
//                    colum.setMinWidth(162);
//                    colum.setWidth(162);
                    // 品種
                    colum = cmdl.getColumn(2);
                    colum.setMaxWidth(80);
                    colum.setMinWidth(80);
                    colum.setWidth(80);
                    // T2
                    colum = cmdl.getColumn(3);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);
                    // 登録日時
                    colum = cmdl.getColumn(4);
                    colum.setMaxWidth(162);
                    colum.setMinWidth(162);
                    colum.setWidth(162);
// chg end 2008.09.17
                }
                catch (Throwable e) {
                    CZSystem.handleException(e);
                }
            }

            /**
            * データ選択時の処理
            */
			@SuppressWarnings("unchecked")
            public void valueChanged(ListSelectionEvent e){
                super.valueChanged(e);
                if(e.getValueIsAdjusting()) return;
                int row = getSelectedRow();
                if(0 > row){
                    if(!life){
                        life = true;
                        return;
                    }
                    removeBtCondition();
                    return;
                }
                //対応するバッチ開始情報を取得する。
                Vector v = new Vector(50);
                CZSystemBt bt = (CZSystemBt)bt_list.elementAt(row);
                for (int i = 0 ; i < bt_all_list.size() ; i++) {
                    CZSystemBt bt_tmp = (CZSystemBt)bt_all_list.elementAt(i);
                    if(bt.batch.equals(bt_tmp.batch)) v.addElement(bt_tmp);
                }
                setBtCondition(v);
            }

            /**
            *
            */
            public void setData(int gr,int tbl){
            }
        } // BtTable

        //======================================================================
        /**
        *       ＢｔＮｏ実績一覧：モデル
        */
        public class BtTblMdl extends AbstractTableModel {

            private int TBL_ROW             = 0;
// chg start 2008.09.17
//            final   int TBL_COL             = 3;
            final   int TBL_COL             = 5;
// chg end 2008.09.17
            private Vector  bt_list         = null;
// chg start 2008.09.17
//            final String[] names = {" # "  , "Bt" , "登録日時" };
            final String[] names = {" # "  , "Bt" , "品種" , "T2" , "登録日時" };
// chg end 2008.09.17
            private Object  data[][];

            /**
            * コンストラクタ
            */
            BtTblMdl(Vector v){
                super();
                bt_list = v;
                TBL_ROW = bt_list.size();
                data = new Object[TBL_ROW][TBL_COL];
                for (int i = 0 ; i < TBL_ROW ; i++) {
                    CZSystemBt bt = (CZSystemBt)bt_list.elementAt(i);
                    if(null == bt) break;
                    data[i][0] = new Integer(i+1);
                    data[i][1] = bt.batch;
// chg start 2008.09.17
//                    data[i][2] = bt.t_time;
                    data[i][2] = bt.hinshu;
                    data[i][3] = bt.no_hikiage;
                    data[i][4] = bt.t_time;
// chg end 2008.09.17
                }
            }

            /**
            * カラム数を取得する。
            */
            public int getColumnCount(){
                return TBL_COL;
            }
            /**
            * 行数を取得する。
            */
            public int getRowCount(){
                return TBL_ROW;
            }
            /**
            * 指定のセルのデータを取得する。
            */
            public Object getValueAt(int row, int col){
                return data[row][col];
            }
            /**
            * カラム名称を取得する。
            */
            public String getColumnName(int column){
                return names[column];
            }
            /**
            * カラムの型を取得する。
            */
            public Class getColumnClass(int c){
                return getValueAt(0, c).getClass();
            }
            /**
            * セルの編集可否を取得する。
            */
            public boolean isCellEditable(int row, int col){
                return false;
            }
            /**
            * 指定のセルへデータを設定する。
            */
            public void setValueAt(Object aValue, int row, int column){
                data[row][column] = aValue;
            }
        } // BtTblMdl

        //======================================================================
        /**
        *       Ｂｔスタート時間一覧テーブル
        */
        class BtStartTable extends JTable {

            private Vector  bt_list         = null;
            private Vector  bt_start_list   = null;
            private BtStartTblMdl model = null;
            private boolean life = false;

            /**
            * コンストラクタ
            */
			@SuppressWarnings("unchecked")
            BtStartTable(Vector v){
                super();
                bt_list = v;

                try{
                    setName("BtStartTable");
                    setAutoCreateColumnsFromModel(true);
                    setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
                    setLocale(new Locale("ja","JP"));
                    setFont(new java.awt.Font("dialog", 0, 12));
                    setRowHeight(17);

                    CZSystemBt bt = (CZSystemBt)bt_list.elementAt(0);
                    Vector tmp = new Vector();
                    tmp = CZSystem.getBtStart(ro_db_name,bt.batch);
                    //NULL回避必要
                    if(null == tmp) return;
                    //Body だけにする
                    int size = tmp.size();
                    bt_start_list = new Vector(size);
                    for (int i = 0 ; i < size ; i++) {
                        CZSystemStart st = (CZSystemStart)tmp.elementAt(i);
                        if(null == st) break;
//@@@                        if((7 == st.p_no) && (1 == st.sp_no)){
                        if(7 == st.p_no){                       //@@@
                            bt_start_list.addElement(st);
                        }
                    }

                    model = new BtStartTblMdl(bt_start_list);
                    setModel(model);

                    DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                    TableColumn     colum = null;
                    // No
                    colum = cmdl.getColumn(0);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);
                    // PNo
                    colum = cmdl.getColumn(1);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);
                    // SPNo
                    colum = cmdl.getColumn(2);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);
                    // PSeq
                    colum = cmdl.getColumn(3);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);
                    // プロセス
                    colum = cmdl.getColumn(4);
                    colum.setMaxWidth(70);
                    colum.setMinWidth(70);
                    colum.setWidth(70);
                    // 登録日時
                    colum = cmdl.getColumn(5);
                    colum.setMaxWidth(162);
                    colum.setMinWidth(162);
                    colum.setWidth(162);
                }
                catch (Throwable e) {
                    CZSystem.handleException(e);
                }
            }

            /**
            * データ選択時の処理
            */
            public void valueChanged(ListSelectionEvent e){
                super.valueChanged(e);
                if(e.getValueIsAdjusting()) return;
                int row = getSelectedRow();

                if(0 > row){
                    if(!life){
                        life = true;
                        return;
                    }
                    removeBtStart();
                    read_button.setEnabled(false);
                    return;
                }
                //選択された開始情報を保持する。
                CZSystemStart st = (CZSystemStart)bt_start_list.elementAt(row);
                setBtStart(st);
                read_button.setEnabled(true);           //読込みボタンを使用可にする。
            }

            /**
            *
            */
            public void setData(int gr,int tbl){
            }
        } //BtStartTable

        //======================================================================
        /*
        *       Ｂｔスタート時間一覧：モデル
        */
        public class BtStartTblMdl extends AbstractTableModel {

            private int TBL_ROW             = 0;
            final   int TBL_COL             = 6;
            private Vector  bt_start_list   = null;
//            final String[] names = {" # ", "PNo", "SPNo", "PSeq", "プロセス", "登録日時" };
            final String[] names = {" # ", "PNo", "SPNo", "PSeq", "プロセス", "開始日時" };
            private Object  data[][];

            /**
            * コンストラクタ
            */
            BtStartTblMdl(Vector v){
                super();
                bt_start_list = v;
                TBL_ROW = bt_start_list.size();
                data = new Object[TBL_ROW][TBL_COL];

                for (int i = 0 ; i < TBL_ROW ; i++) {
                    CZSystemStart st = (CZSystemStart)bt_start_list.elementAt(i);

                    if(null == st) break;

                    data[i][0] = new Integer(i+1);              //№
                    data[i][1] = new Integer(st.p_no);          //プロセスナンバー
                    data[i][2] = new Integer(st.sp_no);         //サブプロセスナンバー
                    data[i][3] = new Integer(st.p_renban);      //プロセス連番
                    data[i][4] = CZSystem.getProcName(st.p_no); //プロセス名
                    data[i][5] = st.p_start;                    //スタート時間
                }
            }

            /**
            *
            */
            public int getColumnCount(){
                return TBL_COL;
            }
            /**
            *
            */
            public int getRowCount(){
                return TBL_ROW;
            }
            /**
            *
            */
            public Object getValueAt(int row, int col){
                return data[row][col];
            }
            /**
            *
            */
            public String getColumnName(int column){
                return names[column];
            }
            /**
            *
            */
            public Class getColumnClass(int c){
                return getValueAt(0, c).getClass();
            }
            /**
            *
            */
            public boolean isCellEditable(int row, int col){
                return false;
            }
            /**
            *
            */
            public void setValueAt(Object aValue, int row, int column){
                data[row][column] = aValue;
            }
        }
    } // SercheDialog
}
