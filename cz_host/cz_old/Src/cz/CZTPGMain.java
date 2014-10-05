package cz;

import java.awt.Color;
import java.awt.Cursor;
import java.awt.Dimension;
import java.awt.Graphics;
import java.awt.Rectangle;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ComponentListener;
import java.awt.event.MouseEvent;
import java.awt.event.MouseMotionListener;
import java.io.FileInputStream;
import java.util.Locale;
import java.util.Properties;
import java.util.Vector;

import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JDialog;
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

/***********************************************************
 *   ＴＰＧ基本グラフ
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * @@ T6 ... 追加
 ***********************************************************/
public class CZTPGMain extends JDialog {

    private final int T1 = 1;
    private final int T2 = 2;
    private final int T3 = 3;
    private final int T4 = 4;
    private final int T5 = 5;
    private final int T6 = 6;       //@@ Add

    private final Color BACK_COL            = java.awt.Color.black;
    private final Color MEM_LINE1_COL       = java.awt.Color.lightGray;
    private final Color MEM_LINE2_COL       = java.awt.Color.gray;
    private final Color MEM_LINE3_COL       = java.awt.Color.darkGray;

    private final Color MAIN1_H_T_COL       = java.awt.Color.cyan;
    private final Color MAIN1_H_T_PF_COL    = java.awt.Color.green;
    private final Color DIA_COL             = java.awt.Color.red;
    private final Color DIA_PF_COL          = java.awt.Color.gray;
    private final Color SXL_ST_COL          = java.awt.Color.magenta;
    private final Color SXL_ST_PF_COL       = java.awt.Color.orange;

    private final Color SXL_RT_COL          = java.awt.Color.green;
    private final Color CRU_RT_COL          = java.awt.Color.cyan;

    private Color PULL_AR_COL               = java.awt.Color.green;
    private final Color VAC_COL             = java.awt.Color.cyan;


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

    private String  main1_h_t_min_pro;      //メインヒーター１温度
    private String  main1_h_t_max_pro;
    private String  main1_h_t_pf_min_pro;   //メインヒーター１温度プロファイル
    private String  main1_h_t_pf_max_pro;
    private String  dia_min_pro;            //直径
    private String  dia_max_pro;
    private String  dia_pf_min_pro;         //直径プロファイル
    private String  dia_pf_max_pro;
    private String  sxl_st_min_pro;         //引き上げ速度
    private String  sxl_st_max_pro;
    private String  sxl_st_pf_min_pro;      //引き上げ速度プロファイル
    private String  sxl_st_pf_max_pro;
    private String  sxl_rt_pf_min_pro;      //シード回転プロファイル
    private String  sxl_rt_pf_max_pro;
    private String  cru_rt_pf_min_pro;      //ルツボ回転プロファイル
    private String  cru_rt_pf_max_pro;
    private String  pull_ar_pf_min_pro;     //プルアルゴンプロファイル
    private String  pull_ar_pf_max_pro;
    private String  vac_pf_min_pro;         //炉内圧プロファイル
    private String  vac_pf_max_pro;

    public float    x_length_mouse          = 0.0f;     //ＳＸＬ長さ
    public float    y_main1_h_t_mouse       = 0.0f;     //メインヒーター１温度
    public float    y_main1_h_t_pf_mouse    = 0.0f;     //メインヒーター１温度プロファイル
    public float    y_dia_mouse             = 0.0f;     //直径
    public float    y_sxl_st_mouse          = 0.0f;     //引き上げ速度
    public float    y_sxl_st_pf_mouse       = 0.0f;     //引き上げ速度プロファイル

    private final String GR_X_LENGTH_DEF    = "2500";   //Ｘ軸の長さ
    private String  gr_x_length = GR_X_LENGTH_DEF;      //Ｘ軸の長さ
    private float   gr_x_bun    = 10.0f;                //Ｘ軸の分割
    private float   gr_y_bun    = 5.0f;                 //Ｙ軸の分割

    private int Y_VIEW_TIMES    = 2;                    //Ｙ軸の倍数

    private String  X_LENGTH_LIST[] = {gr_x_length,
                       "2000",
                       "1500",
                       "1000",
                       "500",
                       "250",
                       "200",
                       "100",
                       "50"};

    private String ro_name                  = null; //対象炉番
    private String ro_db_name               = null; //対象炉データベース名

    private CZSystemStart ro_bt_start       = null; //検索用引き上げ条件
    private Vector ro_bt_all_condition      = null; //全Btの引き上げ条件

    private Vector pv_data_shld             = null; //ショルダーのデータ
    private Vector pv_data_body             = null; //ボディーのデータ

    private JLabel main_ro_name_lab         = null; //炉番表示

    private MainSc  main_sc                 = null; //メイングラフスクロールパネル
    private XSc     x_sc                    = null; //Ｘ軸グラフスクロールパネル
    private Y1Sc    y1_sc                   = null; //Ｙ軸左側グラフスクロールパネル
    private Y2Sc    y2_sc                   = null; //Ｙ軸右側グラフスクロールパネル

    private MainMouseView   main_mouse_view = null; //マウス座標表示パネル

    private SercheDialog    serche_dia      = null; //検索用ダイアログ
    private YLengDialog     y_leng_dia      = null; //Ｙ軸設定用ダイアログ

    private ConditionPanel  conpane         = null; //検索、引き上げ条件パネル
    private GraphPanel      grapane         = null; //グラフ設定パネル
    private SimplGraphPanel simpgrapane     = null; //簡易グラフ表示パネル

    // ---------- コンストラクタ ---------------------------
    //
    CZTPGMain(){
        super();

        try{
            // ----- Propertie_Fileより Min,Max値を取得する。 --------
            Properties prop =  new Properties();
            FileInputStream pros = new FileInputStream(CZSystemDefine.TPGPROPERTY_FILE);
            prop.load(pros);

            prop.list(System.out);
            main1_h_t_min_pro       = prop.getProperty("MAIN1_H_T_MIN");    //メインヒーター１温度
            main1_h_t_max_pro       = prop.getProperty("MAIN1_H_T_MAX");
            main1_h_t_pf_min_pro    = prop.getProperty("MAIN1_H_T_PF_MIN"); //メインヒーター１温度プロファイル
            main1_h_t_pf_max_pro    = prop.getProperty("MAIN1_H_T_PF_MAX");
            dia_min_pro         = prop.getProperty("DIA_MIN");              //直径
            dia_max_pro         = prop.getProperty("DIA_MAX");
            dia_pf_min_pro      = prop.getProperty("DIA_PF_MIN");           //直径プロファイル
            dia_pf_max_pro      = prop.getProperty("DIA_PF_MAX");

            sxl_st_min_pro      = prop.getProperty("SXL_ST_MIN");           //引き上げ速度プロファイル
            sxl_st_max_pro      = prop.getProperty("SXL_ST_MAX");
            sxl_st_pf_min_pro   = prop.getProperty("SXL_ST_PF_MIN");        //引き上げ速度プロファイル
            sxl_st_pf_max_pro   = prop.getProperty("SXL_ST_PF_MAX");
            sxl_rt_pf_min_pro   = prop.getProperty("SXL_RT_PF_MIN");        //シード回転プロファイル
            sxl_rt_pf_max_pro   = prop.getProperty("SXL_RT_PF_MAX");
            cru_rt_pf_min_pro   = prop.getProperty("CRU_RT_PF_MIN");        //ルツボ回転プロファイル
            cru_rt_pf_max_pro   = prop.getProperty("CRU_RT_PF_MAX");
            pull_ar_pf_min_pro  = prop.getProperty("PULL_AR_PF_MIN");       //プルアルゴンプロファイル
            pull_ar_pf_max_pro  = prop.getProperty("PULL_AR_PF_MAX");
            vac_pf_min_pro      = prop.getProperty("VAC_PF_MIN");           //炉内圧プロファイル
            vac_pf_max_pro      = prop.getProperty("VAC_PF_MAX");
        }
        catch( Exception e){
            CZSystem.exit(-1,"CZTPGMain NO Propertie File");
        }

        ro_name = CZSystem.getRoName();
        ro_db_name = CZSystem.getDBName();

        setTitle("ＴＰＧ");                         //画面Title
        setSize(1152,920);                          //画面サイズ
        setResizable(false);                        //画面のサイズ変更は不可
        setModal(true);                             //Modalで表示
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
        main_ro_name_lab.setBounds(20, 20, 100, 30);
        main_ro_name_lab.setLocale(new Locale("ja","JP"));
        main_ro_name_lab.setFont(new java.awt.Font("dialog", 0, 18));
        main_ro_name_lab.setBorder(new Flush3DBorder());
        main_ro_name_lab.setForeground(java.awt.Color.black);
        getContentPane().add(main_ro_name_lab);

        //検索パネル
        conpane = new ConditionPanel();
        conpane.setBounds(20, 60, 100, 340);
        getContentPane().add(conpane);

        //グラフ設定パネル
        grapane = new GraphPanel();
        grapane.setBounds(20, 410, 100, 110);
        getContentPane().add(grapane);

        //マウス座標表示パネル
        main_mouse_view = new MainMouseView();
        main_mouse_view.setBounds(20, 570, 100, 300);
        getContentPane().add(main_mouse_view);

        //簡易グラフ表示パネル
        simpgrapane = new SimplGraphPanel(1000,300);
        simpgrapane.setBounds(140, 570, 1000, 300);
        getContentPane().add(simpgrapane);

        // グラフ表示領域
        PVGrEventCompo comp = new PVGrEventCompo();

        main_sc = new MainSc(comp);                 // メイングラフのパネル
        main_sc.setBounds(190, 20, 890, 500);
        main_sc.setDefault();
        getContentPane().add(main_sc);

        x_sc    = new XSc(comp);                    // X軸の目盛のパネル
        x_sc.setBounds(190, 520, 890, 40);
        x_sc.setDefault();
        getContentPane().add(x_sc);

        y1_sc   = new Y1Sc(comp);                   // Y軸の左側のパネル
        y1_sc.setBounds(140, 20, 50, 500);
        y1_sc.setDefault();
        getContentPane().add(y1_sc);

        y2_sc   = new Y2Sc(comp);                   // Y軸の右側のパネル
        y2_sc.setBounds(1080, 20, 60, 500);
        y2_sc.setDefault();
        getContentPane().add(y2_sc);

        serche_dia = new SercheDialog();            //検索Dialog
        serche_dia.setVisible(false);

        y_leng_dia = new YLengDialog();             //項目設定Dialog
        y_leng_dia.setVisible(false);

        CZSystem.log("CZTPGMain","CZTPGMain new");
    }

    // 炉番とDB名称を取得する。
    // @return true ... OK, false ... NG
    public boolean setDefault(){
        ro_name = CZSystem.getRoName();
        ro_db_name = CZSystem.getDBName();
        return true;
    }

    // 炉番表示を表示する。
    //
    private void setMainRoName(){

		String s = CZSystem.RoKetaChg(ro_name);	// 20050725 炉：表示桁数変更
        main_ro_name_lab.setText(s);
//        main_ro_name_lab.setText(ro_name);
    }

    // X軸の長さを変更する。
    // @param len 長さ
    private void chgXLength(String len){
        gr_x_length = len;
        main_sc.chgXSize();
        x_sc.chgXSize();
        simpgrapane.chgXSize();
    }

    // Y軸の長さを変更する。
    //
    private void chgYLength(){
        main_sc.chgYSize();
        y1_sc.chgYSize();
        y2_sc.chgYSize();
        simpgrapane.chgYLength();
    }

    // 肩表示を変更する。
    //
    private void chgShld(){
        main_sc.chgYSize();
        simpgrapane.chgShld();
    }

    // ＴＰＧエラーメッセージ表示Dialog
    // @param msg ... メッセージ内容
    // @return true ... OK, false ... NG
    private boolean errorMsg(Object msg[]){
        JOptionPane.showMessageDialog(null,msg,
                                "ＴＰＧエラー",
                                JOptionPane.ERROR_MESSAGE);
        return true;
    }

    // バッチ開始時刻を設定する
    // @param st ... バッチ開始時刻
    // @return true ... OK, false ... NG
    public boolean setBtStart(CZSystemStart st){
        ro_bt_start = st;
        if(null == ro_bt_start) return false;
        return true;
    }

    // 設定済みバッチ開始時刻を削除する
    // @return true ... OK
    public boolean removeBtStart(){
        ro_bt_start = null;
        return true;
    }

    //
    // PVデータを読み込む
    // @return ... ボディー実績の読込み件数
    // （-1 ... 実績無し,-2 ... 表無し,-3 ... ショルダー実績無し,-4 ... ボディー実績無し）
    public int readBtPV(){
        if(null == ro_bt_start){
            Object msg[] = {"スタート実績が有りません！！",
                            "",
                            ""};
            errorMsg(msg);
            return -1;
        }

        CZSystemStart st = ro_bt_start;

//@@        CZSystem.log("CZTPGMain readBtPV ","batch    [" + st.batch    + "]");
//@@        CZSystem.log("CZTPGMain readBtPV ","p_no     [" + st.p_no     + "]");
//@@        CZSystem.log("CZTPGMain readBtPV ","sp_no    [" + st.sp_no    + "]");
//@@        CZSystem.log("CZTPGMain readBtPV ","p_renban [" + st.p_renban + "]");
//@@        CZSystem.log("CZTPGMain readBtPV ","p_start  [" + st.p_start  + "]");
        String view = CZSystem.getViewName(ro_db_name,st.batch);
//@@        CZSystem.log("CZTPGMain readBtPV ","view  [" + view  + "]");
        if(null == view){
            Object msg[] = {"表が存在しません！！",
                            view,
                            ""};
            errorMsg(msg);
            return -2;
        }

        boolean data_no[] = new boolean[CZSystemDefine.PV_MAX_LENGTH];
        for(int i = 0 ; i < CZSystemDefine.PV_MAX_LENGTH ; i++) data_no[i] = false;

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

        //ショルダー読み込み
        pv_data_shld = CZSystem.getPVData(ro_db_name,view,st.p_renban-1,data_no);
//@@        CZSystem.log("CZTPGMain readBtPV ","pv_data_shld  [" + pv_data_shld.size()  + "]");
        if(1 > pv_data_shld.size()){
            Object msg[] = {"ショルダー実績が有りません！！",
                            "[" + pv_data_shld.size() + "]",
                            ""};
            errorMsg(msg);
            pv_data_shld = null;
            return -3;
        }

        //ボディー読み込み
        pv_data_body = CZSystem.getPVData(ro_db_name,view,st.p_renban,data_no);
//@@        CZSystem.log("CZTPGMain readBtPV ","pv_data_body  [" + pv_data_body.size()  + "]");
        if(1 > pv_data_body.size()){
            Object msg[] = {"ボディー実績が有りません！！",
                            "[" + pv_data_body.size() + "]",
                            ""};
            errorMsg(msg);
            pv_data_body = null;
            return -4;
        }
        return pv_data_body.size();
    }

    // カーソルを設定する。
    //
    private void setCur(Cursor cu){
        serche_dia.setCursor(cu);
    }

    //
    // カーソルを取得する。
    private Cursor getCur(){
        return serche_dia.getCursor();
    }

    /*******************************************************
     *
     *   検索のパネル
     *
     *******************************************************/
    public class ConditionPanel extends JPanel {

        private JTextField bt_text = null;

        private JTextField t1_text = null;
        private JTextField t2_text = null;
        private JTextField t3_text = null;
        private JTextField t4_text = null;
        private JTextField t5_text = null;
        private JTextField t6_text = null;          //@@

        // ---------- コンストラクタ -----------------------
        //
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

            int x = 10;
            int y = 10;
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

            x = 10;
            y = 90 ;
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

//@@追加↓
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
//@@追加↑

//@@            y += 40 ;
            y += inc;                   //@@
            JButton btcondition_button = new JButton("引上条件");
            btcondition_button.setBounds(x, y, 80, 24);
            btcondition_button.setLocale(new Locale("ja","JP"));
            btcondition_button.setFont(new java.awt.Font("dialog", 0, 18));
            btcondition_button.setBorder(new Flush3DBorder());
            btcondition_button.setForeground(java.awt.Color.black);
            btcondition_button.addActionListener(new BtConditionButton());
            add(btcondition_button);

//@@            y += 50 ;
            y += inc;                   //@@
            JButton controltable_button = new JButton("制御テーブル");
            controltable_button.setBounds(x, y, 80, 24);
            controltable_button.setLocale(new Locale("ja","JP"));
            controltable_button.setFont(new java.awt.Font("dialog", 0, 12));
            controltable_button.setBorder(new Flush3DBorder());
            controltable_button.setForeground(java.awt.Color.black);
            controltable_button.addActionListener(new ControlTable());
            add(controltable_button);

//@@            CZSystem.log("ConditionPanel ConditionPanel","new");
        }

        //
        // T1～T6の設定
        // @param ... b 
        public void setData(boolean b){

            boolean flag = b;

            if(null == ro_bt_all_condition){
                flag = false;
            }

            if(flag){

                setMainRoName();

                CZSystemBt bt = (CZSystemBt)ro_bt_all_condition.elementAt(0);
                if(null == bt) return;

                bt_text.setText(bt.batch.trim());
                t1_text.setText(String.valueOf(bt.no_youkai));      // 溶解
                t2_text.setText(String.valueOf(bt.no_hikiage));     // 引上
                t3_text.setText(String.valueOf(bt.no_kaiten));      // 回転
                t4_text.setText(String.valueOf(bt.no_toridasi));    // 取出
                t5_text.setText(String.valueOf(bt.no_aturyoku));    // 圧力
                t6_text.setText(String.valueOf(bt.no_teisu));       // 定数 @@追加
            }
            else{
                bt_text.setText("");
                t1_text.setText("");
                t2_text.setText("");
                t3_text.setText("");
                t4_text.setText("");
                t5_text.setText("");
                t6_text.setText("");        //@@追加
            }
        }

        /***************************************************
         *   検索ボタンの処理
         *    検索Dialogを表示する。
         ***************************************************/
        class SearchButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                serche_dia.setDefault();
                serche_dia.setVisible(true);
//@@                CZSystem.log("ConditionPanel SaveButton","actionPerformed");
            }
        }

        /***************************************************
         *   引き上げ条件ボタン
         *    引き上げ条件設定Dialogを表示する。
         ***************************************************/
        class BtConditionButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
//@@                CZSystem.log("ConditionPanel BtConditionButton","actionPerformed start");
                if(null == ro_bt_all_condition) return ;
                BtConditionDialog dialog = new BtConditionDialog();
                dialog.setVisible(true);
//@@                CZSystem.log("ConditionPanel BtConditionButton","actionPerformed end");
            }
        }

        /***************************************************
         *   制御テーブル設定ボタン
         *    制御テーブル設定Dialogを表示する。
         ***************************************************/
        class ControlTable implements ActionListener {
            private CZControlTable obj = null;

            public void actionPerformed(ActionEvent ev){
//@@                CZSystem.log("ConditionPanel ControlTable","actionPerformed start");
                if(null == obj) obj = new CZControlTable();
                obj.setDefault(pv_data_shld,pv_data_body);
                obj.setVisible(true);
//@@                CZSystem.log("ConditionPanel ControlTable","actionPerformed end");
            }
        }

        /***************************************************
         *
         * 引き上げ条件Dialog
         *
         ***************************************************/
        class BtConditionDialog extends JDialog {

            // ---------- コンストラクタ -------------------
            //
            BtConditionDialog(){
                super();

                setTitle("引き上げ条件");
                setSize(820,250);
                setResizable(false);
                setModal(true);
                getContentPane().setLayout(null);
                // 他基地参照機能    @20131021
                if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                    getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
                }else{
                    getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
                }

                BtConditionTable t = new BtConditionTable(ro_bt_all_condition);
                JTableHeader tabHead = t.getTableHeader();
                tabHead.setReorderingAllowed(false);

                JScrollPane bt_scpanel = new JScrollPane(t);
                bt_scpanel.setBounds(20, 20, 780, 187);
                getContentPane().add(bt_scpanel);

//@@                CZSystem.log("CZTPGMain SercheDialog","new");
            }

            /***********************************************
             *
             *       Ｂｔ登録情報一覧
             * @@T6追加
             ***********************************************/
            class BtConditionTable extends JTable {

                private Vector  bt_list     = null;

                private BtConditionTblMdl model = null;

                // ---------- コンストラクタ ---------------
                // @param v ... バッチ情報
                //
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

                        // T6
                        colum = cmdl.getColumn(14);
                        colum.setMaxWidth(30);
                        colum.setMinWidth(30);
                        colum.setWidth(30);

                        // PNo
                        colum = cmdl.getColumn(15);
                        colum.setMaxWidth(32);
                        colum.setMinWidth(32);
                        colum.setWidth(32);

                        // 開始
                        colum = cmdl.getColumn(16);
                        colum.setMaxWidth(30);
                        colum.setMinWidth(30);
                        colum.setWidth(30);
                    }
                    catch (Throwable e) {
                        CZSystem.handleException(e);
                    }
                }

                //
                // Change Listener
                //
                public void valueChanged(ListSelectionEvent e){
                    super.valueChanged(e);
                }

                //
                //
                //
                public void setData(int gr,int tbl){
                    System.out.println("setData [" + gr + "][" + tbl + "]");
                }

                /*******************************************
                 *
                 *       Ｂｔ登録情報一覧：モデル
                 *
                 *******************************************/
                public class BtConditionTblMdl extends AbstractTableModel {

                    private int     TBL_ROW     = 0;        // 行数
                    final   int     TBL_COL     = 17;       // 列数 @@
                    private Vector  bt_list     = null;     // バッチ情報

                    final String[] names = {" # "  , "登録日時" , "連番" ,  
                                "品種" , "ルツボ"   , "直径" ,
                                "引上長" , "初仕込"   , "追仕込" ,
                                "T1" , "T2"   , "T3" ,
                                "T4" , "T5"   , "T6"   , "PNo" , "開始"
                                };

                    private Object  data[][];

                    // ---------- コンストラクタ -----------
                    // @param v ... バッチ情報
                    BtConditionTblMdl(Vector v){
                        super();
                        bt_list = v;
                        TBL_ROW = bt_list.size();

                        data = new Object[TBL_ROW][TBL_COL];

                        for(int i = 0 ; i < TBL_ROW ; i++){
                            CZSystemBt bt = (CZSystemBt)bt_list.elementAt(i);
                            if(null == bt) break;
                            data[i][0]  = new Integer(i+1);             //#
                            data[i][1]  = bt.t_time;                    //登録日時
                            data[i][2]  = new Integer(bt.renban);       //連番
                            data[i][3]  = bt.hinshu;                    //品種
                            data[i][4]  = new Integer(bt.rutubo_kei);   //ルツボ
                            data[i][5]  = new Integer(bt.chokkei);      //直径
                            data[i][6]  = new Integer(bt.hikiage_cho);  //引上長
                            data[i][7]  = new Integer(bt.i_sikomi);     //初仕込
                            data[i][8]  = new Integer(bt.t_sikomi);     //追仕込
                            data[i][9]  = new Integer(bt.no_youkai);    //T1
                            data[i][10] = new Integer(bt.no_hikiage);   //T2
                            data[i][11] = new Integer(bt.no_kaiten);    //T3
                            data[i][12] = new Integer(bt.no_toridasi);  //T4
                            data[i][13] = new Integer(bt.no_aturyoku);  //T5
                            data[i][14] = new Integer(bt.no_teisu);     //T6 @@
                            data[i][15] = new Integer(bt.pno_start);    //PNo
                            data[i][16] = new Integer(bt.p_kaisi);      //開始
                        }
                    }

                    // 桁数を取得する。
                    // @return ... 桁数
                    public int getColumnCount(){
                        return TBL_COL;
                    }

                    // 行数を取得する。
                    // @return ... 行数
                    public int getRowCount(){
                        return TBL_ROW;
                    }

                    // データを取得する。
                    // @param ... row:行, col:桁
                    // @return ... データ
                    public Object getValueAt(int row, int col){
                        return data[row][col];
                    }

                    // 桁名を取得する。
                    // @param ... column:桁
                    // @return ... 桁名
                    public String getColumnName(int column){
                        return names[column];
                    }

                    // データの型を取得する。
                    // @param ... c:桁
                    // @return ... データの型
                    public Class getColumnClass(int c){
                        return getValueAt(0, c).getClass();
                    }

                    // cell編集の可否を取得する。
                    // @param ... row:行, col:桁
                    // @return ... 桁数
                    public boolean isCellEditable(int row, int col){
                        return false;
                    }

                    // データを設定する。
                    // @param ... aValue:データ, row:行, col:桁
                    // @return ... 桁数
                    public void setValueAt(Object aValue, int row, int column){
                        data[row][column] = aValue;
                    }
                } // BtConditionTblMdl
            } // BtConditionTable
        } // BtConditionDialog
    } // ConditionPanel


    /*******************************************************
     *
     *   Ｙ軸設定、Ｘ軸設定、ショルダーの設定パネル
     *
     *******************************************************/
    public class GraphPanel extends JPanel {

        JCheckBox shld_chk = null;

        // ---------- コンストラクタ -----------------------
        GraphPanel(){
            super();
            setName("GraphPanel");
            setLayout(null);
            setBorder(new Flush3DBorder());
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            JButton leng_button = new JButton("項目設定");
            leng_button.setBounds(10, 10, 80, 24);
            leng_button.setLocale(new Locale("ja","JP"));
            leng_button.setFont(new java.awt.Font("dialog", 0, 18));
            leng_button.setBorder(new Flush3DBorder());
            leng_button.setForeground(java.awt.Color.black);
            leng_button.addActionListener(new ChgYLengButton());
            add(leng_button);

            JComboBox x_leng_com = new JComboBox();
            x_leng_com.setBounds(10, 40, 80, 24);
            x_leng_com.setLocale(new Locale("ja","JP"));
            x_leng_com.setFont(new java.awt.Font("dialog", 0, 12));
            x_leng_com.setForeground(java.awt.Color.black);
            x_leng_com.setFocusable(false);	/* 2007.08.22 */
            for(int i = 0 ; i < X_LENGTH_LIST.length ; i++){
                x_leng_com.addItem(X_LENGTH_LIST[i]);
            }
            x_leng_com.addActionListener(new ChgXLengButton());
            add(x_leng_com);

            shld_chk = new JCheckBox("肩表示",false);
            shld_chk.setBounds(10, 70, 80, 24);
            shld_chk.setLocale(new Locale("ja","JP"));
            shld_chk.setFont(new java.awt.Font("dialog", 0, 12));
            shld_chk.setBorderPaintedFlat(true);
            shld_chk.setForeground(java.awt.Color.black);
            shld_chk.addActionListener(new ShldChk());
            add(shld_chk);

//@@            CZSystem.log("GraphPanel GraphPanel","new");
        }

        // 肩表示チェックボックスのチェック
        // @return ... true, false
        public boolean isShld(){
            return shld_chk.isSelected();
        }

        /***************************************************
         * 項目設定ボタンの処理
         *  項目設定Dialogを表示する。
         ***************************************************/
        class ChgYLengButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                y_leng_dia.setDefault();
                y_leng_dia.setVisible(true);
//@@                CZSystem.log("GraphPanel ChgYLengButton","actionPerformed");
            }
        } //ChgYLengButton

        /***************************************************
         * X軸目盛設定コンボボックスの処理
         ***************************************************/
        class ChgXLengButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                JComboBox obj = (JComboBox)ev.getSource();
                String val = (String)obj.getSelectedItem();
                chgXLength(val);
//@@                CZSystem.log("GraphPanel ChgXLengButton","actionPerformed [" + val + "]");
            }
        } //ChgXLengButton

        /***************************************************
         * 肩表示チェックボックスチェック時の処理
         ***************************************************/
        class ShldChk implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                chgShld();  
//@@                CZSystem.log("GraphPanel ShldChk","actionPerformed");
            }
        } // ShldChk
    } // GraphPanel

    /*******************************************************
     *
     *   回転、圧力系簡易グラフ
     *
     *******************************************************/
    public class SimplGraphPanel extends JPanel {
        private RotationPanel   r_view  = null;     // 回転
        private PressurePanel   p_view  = null;     // 圧力

        // ---------- コンストラクタ -----------------------
        // @param w ... 幅, h ... 高さ
        SimplGraphPanel(int w,int h){
            super();
            setName("SimplGraphPanel");
            setLayout(null);
            setBorder(new Flush3DBorder());
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            int x;
            int y;
            int width;
            int height;
                
            x = 10;
            y = 20;
            width = (w / 2) - x - (x / 2);
            height = h - y - 10;    

            r_view = new RotationPanel();
            r_view.setBounds(x, y, width, height);
            add(r_view);

            JLabel lab = new JLabel("回転系",JLabel.CENTER);
            lab.setBounds(x, 0, 50, y+2);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 12));
            lab.setForeground(java.awt.Color.black);
            add(lab);

            x = (w / 2) + (x / 2);
            p_view = new PressurePanel();
            p_view.setBounds(x, y, width, height);
            add(p_view);

            lab = new JLabel("圧力系",JLabel.CENTER);
            lab.setBounds(x, 0, 50, y+2);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 12));
            lab.setForeground(java.awt.Color.black);
            add(lab);

//@@            CZSystem.log("SimplGraphPanel SimplGraphPanel","new");
        }

        // データを設定し、再表示する。
        //
        public void setData(){
            r_view.setData();
            r_view.repaint();
            p_view.setData();
            p_view.repaint();
        }

        // X軸を変更する。
        //
        public void chgXSize(){
            r_view.setData();
            r_view.repaint();
            p_view.setData();
            p_view.repaint();
        }

        // Y軸を変更する。
        //
        private void chgYLength(){
            r_view.setData();
            r_view.repaint();
            p_view.setData();
            p_view.repaint();
        }

        // 肩表示を変更する。
        //
        private void chgShld(){
            r_view.setData();
            r_view.repaint();
            p_view.setData();
            p_view.repaint();
        }

        /***************************************************
         * 回転を表示するパネル
         ***************************************************/
        class RotationPanel extends JPanel {
            private final int offset_x = 50;
            private final int offset_y = 20;

            int x_pos[];
            int y_pos_sxl_rt[];
            int y_pos_sxl_rt_pf[];

            int y_pos_cru_rt[];
            int y_pos_cru_rt_pf[];

            int x_pos_shld[];
            int y_pos_shld_sxl_rt[];
            int y_pos_shld_sxl_rt_pf[];

            int y_pos_shld_cru_rt[];
            int y_pos_shld_cru_rt_pf[];

            // ---------- コンストラクタ -------------------
            RotationPanel(){
                super();
                setName("SimplGraphPanel");
                setLayout(null);
                setBackground(BACK_COL);
            }

            //
            // グラフを描画する。
            //
            public void paint(Graphics g){
                Dimension d = getSize(null);

                g.setColor(BACK_COL);
                g.fillRect(0,0,d.width,d.height);

                drawMemLine(g);     // 目盛線
                drawLine(g);        // グラフ線
            }

            //
            // PV値より座標を計算する。
            //
            private void setData(){

                if(null == pv_data_shld) return;
                int size_shld = pv_data_shld.size();
                if(2 > size_shld) return;

                if(null == pv_data_body) return;
                int size = pv_data_body.size();
                if(2 > size) return;

                Dimension d = getSize(null);

                Float x_max;
                float val = 0.0f;
                float val_shld = 0.0f;
                float min = 0.0f;
                float max = 0.0f;

                CZSystemPVData data;

                //Ｘ軸座標計算（肩）
                x_pos_shld = new int[size_shld];
                x_max = new Float(gr_x_length);
                min = 0.0f;
                max = x_max.floatValue();

                for(int i = 0 ; i < size_shld ; i++){
                    data = (CZSystemPVData)pv_data_shld.elementAt(i);
                    val = data.p_length;
                    x_pos_shld[i] = (int)xPos(d.width,min,max,val);
                }

                if(grapane.isShld()){
                    val_shld = val;
                }

                //Ｘ軸座標計算
                x_pos = new int[size];
                x_max = new Float(gr_x_length);
                min = 0.0f;
                max = x_max.floatValue();

                for(int i = 0 ; i < size ; i++){
                    data = (CZSystemPVData)pv_data_body.elementAt(i);
                    val = data.p_length + val_shld;
                    x_pos[i] = (int)xPos(d.width,min,max,val);
                }

                Float s_min;
                Float s_max;

                //シード回転プロファイル（肩）
                y_pos_shld_sxl_rt_pf = new int[size_shld];
                s_min = new Float(sxl_rt_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(sxl_rt_pf_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size_shld ; i++){
                    data = (CZSystemPVData)pv_data_shld.elementAt(i);
                    val = data.data[SXL_RT_PF];
                    y_pos_shld_sxl_rt_pf[i] = (int)yPos(d.height,min,max,val);
                }

                //シード回転プロファイル
                y_pos_sxl_rt_pf = new int[size];
                s_min = new Float(sxl_rt_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(sxl_rt_pf_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size ; i++){
                    data = (CZSystemPVData)pv_data_body.elementAt(i);
                    val = data.data[SXL_RT_PF];
                    y_pos_sxl_rt_pf[i] = (int)yPos(d.height,min,max,val);
                }

                //シード回転（肩）
                y_pos_shld_sxl_rt = new int[size_shld];
                s_min = new Float(sxl_rt_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(sxl_rt_pf_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size_shld ; i++){
                    data = (CZSystemPVData)pv_data_shld.elementAt(i);
                    val = data.data[SXL_RT];
                    y_pos_shld_sxl_rt[i] = (int)yPos(d.height,min,max,val);
                }

                //シード回転
                y_pos_sxl_rt = new int[size];
                s_min = new Float(sxl_rt_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(sxl_rt_pf_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size ; i++){
                    data = (CZSystemPVData)pv_data_body.elementAt(i);
                    val = data.data[SXL_RT];
                    y_pos_sxl_rt[i] = (int)yPos(d.height,min,max,val);
                }

                //ルツボ回転プロファイル（肩）
                y_pos_shld_cru_rt_pf = new int[size_shld];
                s_min = new Float(cru_rt_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(cru_rt_pf_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size_shld ; i++){
                    data = (CZSystemPVData)pv_data_shld.elementAt(i);
                    val = data.data[CRU_RT_PF];
                    y_pos_shld_cru_rt_pf[i] = (int)yPos(d.height,min,max,val);
                }

                //ルツボ回転プロファイル
                y_pos_cru_rt_pf = new int[size];
                s_min = new Float(cru_rt_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(cru_rt_pf_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size ; i++){
                    data = (CZSystemPVData)pv_data_body.elementAt(i);
                    val = data.data[CRU_RT_PF];
                    y_pos_cru_rt_pf[i] = (int)yPos(d.height,min,max,val);
                }

                //ルツボ回転（肩）
                y_pos_shld_cru_rt = new int[size_shld];
                s_min = new Float(cru_rt_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(cru_rt_pf_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size_shld ; i++){
                    data = (CZSystemPVData)pv_data_shld.elementAt(i);
                    val = data.data[CRU_RT];
                    y_pos_shld_cru_rt[i] = (int)yPos(d.height,min,max,val);
                }

                //ルツボ回転
                y_pos_cru_rt = new int[size];
                s_min = new Float(cru_rt_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(cru_rt_pf_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size ; i++){
                    data = (CZSystemPVData)pv_data_body.elementAt(i);
                    val = data.data[CRU_RT];
                    y_pos_cru_rt[i] = (int)yPos(d.height,min,max,val);
                }
            }

            // X軸の座標を計算する。
            // @param w .. 幅,min .. 最小値,max .. 最大値,val ..データ
            // @return X軸の座標
            private float xPos(int w,float min,float max,float val){
                float x_dot = (w - offset_x) / (max - min);
                float x = x_dot * (val - min) + offset_x;
                return x;
            }

            // Y軸の座標を計算する。
            // @param h .. 高さ,min .. 最小値,max .. 最大値,val ..データ
            // @return Y軸の座標
            private float yPos(int h,float min,float max,float val){
                float y_dot = (h - offset_y) / (max - min);
                float y = h - y_dot * (val - min) - offset_y;
                return y;
            }

            //
            // グラフ線を引く
            //
            private void drawLine(Graphics g){
                if(null == pv_data_shld) return;
                int size_shld = pv_data_shld.size();
                if(2 > size_shld) return;

                if(null == pv_data_body) return;
                int size = pv_data_body.size();
                if(2 > size) return;

                //シード回転プロファイル（肩）
                if(grapane.isShld()){
                    g.setColor(java.awt.Color.white);
                    g.drawPolyline(x_pos_shld,y_pos_shld_sxl_rt_pf,size_shld);
                }
                //シード回転プロファイル
                g.setColor(java.awt.Color.white);
                g.drawPolyline(x_pos,y_pos_sxl_rt_pf,size);
                //シード回転（肩）
                if(grapane.isShld()){
                    g.setColor(SXL_RT_COL);
                    g.drawPolyline(x_pos_shld,y_pos_shld_sxl_rt,size_shld);
                }
                //シード回転
                g.setColor(SXL_RT_COL);
                g.drawPolyline(x_pos,y_pos_sxl_rt,size);
                g.drawString("SXL.RT",x_pos[size-1],y_pos_sxl_rt[size-1]);

                //ルツボ回転プロファイル（肩）
                if(grapane.isShld()){
                    g.setColor(java.awt.Color.white);
                    g.drawPolyline(x_pos_shld,y_pos_shld_cru_rt_pf,size_shld);
                }
                //ルツボ回転プロファイル
                g.setColor(java.awt.Color.white);
                g.drawPolyline(x_pos,y_pos_cru_rt_pf,size);
                //ルツボ回転（肩）
                if(grapane.isShld()){
                    g.setColor(CRU_RT_COL);
                    g.drawPolyline(x_pos_shld,y_pos_shld_cru_rt,size_shld);
                }
                //ルツボ回転
                g.setColor(CRU_RT_COL);
                g.drawPolyline(x_pos,y_pos_cru_rt,size);
                g.drawString("CRU.RT",x_pos[size-1],y_pos_cru_rt[size-1]);
            }

            //
            // 目盛線を引く
            //
            private void drawMemLine(Graphics g){
                float x;
                float y;
                float inc;

                Dimension d = getSize(null);

                // Ｘ軸目盛 小
                g.setColor(MEM_LINE3_COL);
                inc = (d.width - offset_x) / (gr_x_bun * 4.0f);
                for(x = offset_x ; x < d.width ; x+=inc){
                    g.drawLine((int)x,0,(int)x,d.height-offset_y);
                }
                // Ｙ軸目盛
                inc = (d.height - offset_y) / (gr_y_bun * 4.0f) ;
                for(y = d.height - offset_y ; y > 0 ; y-=inc){
                    g.drawLine((int)offset_x,(int)y,d.width,(int)y);
                }

                // Ｘ軸目盛 中
                g.setColor(MEM_LINE2_COL);
                inc = (d.width - offset_x) / (gr_x_bun * 2.0f);
                for(x = offset_x ; x < d.width ; x+=inc){
                    g.drawLine((int)x,0,(int)x,d.height-offset_y);
                }
                // Ｙ軸目盛
                inc = (d.height - offset_y) / (gr_y_bun * 2.0f) ;
                for(y = d.height - offset_y ; y > 0 ; y-=inc){
                    g.drawLine((int)offset_x,(int)y,d.width,(int)y);
                }

                // Ｘ軸目盛 大
                g.setColor(MEM_LINE1_COL);
                Float tmp = new Float(gr_x_length);
                float mem_inc = tmp.floatValue() / gr_x_bun;
                float x_val   = 0.0f;
                boolean cont = true;

                inc = (d.width - offset_x) / gr_x_bun ;
                for(x = offset_x ; x < d.width ; x+=inc){
                    g.drawLine((int)x,0,(int)x,d.height-offset_y);
                    if(cont){
                        g.drawString(String.valueOf(x_val),(int)x+3,d.height-8);
                    }
                    x_val+=mem_inc;
                    cont = !cont;
                }

                // Ｙ軸目盛
                Float tmp_min1 = new Float(sxl_rt_pf_min_pro);
                Float tmp_max1 = new Float(sxl_rt_pf_max_pro);
                float mem_inc1 = (tmp_max1.floatValue()-tmp_min1.floatValue()) / gr_y_bun;
                float y_val1  = tmp_min1.floatValue();

                Float tmp_min2 = new Float(cru_rt_pf_min_pro);
                Float tmp_max2 = new Float(cru_rt_pf_max_pro);
                float mem_inc2 = (tmp_max2.floatValue()-tmp_min2.floatValue()) / gr_y_bun;
                float y_val2  = tmp_min2.floatValue();

                inc = (d.height - offset_y) / gr_y_bun ;
                for(y = d.height - offset_y ; y > 0 ; y-=inc){
                    g.setColor(MEM_LINE1_COL);
                    g.drawLine((int)offset_x,(int)y,d.width,(int)y);

                    g.setColor(SXL_RT_COL);
                    g.drawString(String.valueOf(y_val1),3,(int)y - 0);
                    y_val1 += mem_inc1;

                    g.setColor(CRU_RT_COL);
                    g.drawString(String.valueOf(y_val2),3,(int)y - 10);
                    y_val2 += mem_inc2;
                }
            }
        } // RotationPanel

        /***************************************************
         * 圧力パネル
         ***************************************************/
        class PressurePanel extends JPanel {
            private final int offset_x = 50;
            private final int offset_y = 20;

            int x_pos[];
            int y_pos_pull_ar[];
            int y_pos_pull_ar_pf[];

            int y_pos_vac[];
            int y_pos_vac_pf[];

            int x_pos_shld[];
            int y_pos_shld_pull_ar[];
            int y_pos_shld_pull_ar_pf[];

            int y_pos_shld_vac[];
            int y_pos_shld_vac_pf[];

            // ---------- コンストラクタ -------------------
            PressurePanel(){
                super();
                setName("SimplGraphPanel");
                setLayout(null);
                setBackground(BACK_COL);
            }

            //
            // グラフを描画する。
            //
            public void paint(Graphics g){
                Dimension d = getSize(null);

                g.setColor(BACK_COL);
                g.fillRect(0,0,d.width,d.height);
                drawMemLine(g);
                drawLine(g);
            }

            //
            // データから座標を計算する。
            //
            private void setData(){

                if(null == pv_data_shld) return;
                int size_shld = pv_data_shld.size();
                if(2 > size_shld) return;

                if(null == pv_data_body) return;
                int size = pv_data_body.size();
                if(2 > size) return;

                Dimension d = getSize(null);

                Float x_max;
                float val = 0.0f;
                float val_shld = 0.0f;
                float min = 0.0f;
                float max = 0.0f;

                CZSystemPVData data;

                //Ｘ軸座標計算（肩）
                x_pos_shld = new int[size_shld];
                x_max = new Float(gr_x_length);
                min = 0.0f;
                max = x_max.floatValue();

                for(int i = 0 ; i < size_shld ; i++){
                    data = (CZSystemPVData)pv_data_shld.elementAt(i);
                    val = data.p_length;
                    x_pos_shld[i] = (int)xPos(d.width,min,max,val);
                }

                if(grapane.isShld()){
                    val_shld = val;
                }

                //Ｘ軸座標計算
                x_pos = new int[size];
                x_max = new Float(gr_x_length);
                min = 0.0f;
                max = x_max.floatValue();

                for(int i = 0 ; i < size ; i++){
                    data = (CZSystemPVData)pv_data_body.elementAt(i);
                    val = data.p_length + val_shld;
                    x_pos[i] = (int)xPos(d.width,min,max,val);
                }

                Float s_min;
                Float s_max;

                //プルアルゴンプロファイル（肩）
                y_pos_shld_pull_ar_pf = new int[size_shld];
                s_min = new Float(pull_ar_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(pull_ar_pf_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size_shld ; i++){
                    data = (CZSystemPVData)pv_data_shld.elementAt(i);
                    val = data.data[PULL_AR_PF];
                    y_pos_shld_pull_ar_pf[i] = (int)yPos(d.height,min,max,val);
                }

                //プルアルゴンプロファイル
                y_pos_pull_ar_pf = new int[size];
                s_min = new Float(pull_ar_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(pull_ar_pf_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size ; i++){
                    data = (CZSystemPVData)pv_data_body.elementAt(i);
                    val = data.data[PULL_AR_PF];
                    y_pos_pull_ar_pf[i] = (int)yPos(d.height,min,max,val);
                }

                //プルアルゴン（肩）
                y_pos_shld_pull_ar = new int[size_shld];
                s_min = new Float(pull_ar_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(pull_ar_pf_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size_shld ; i++){
                    data = (CZSystemPVData)pv_data_shld.elementAt(i);
                    val = data.data[PULL_AR];
                    y_pos_shld_pull_ar[i] = (int)yPos(d.height,min,max,val);
                }

                //プルアルゴン
                y_pos_pull_ar = new int[size];
                s_min = new Float(pull_ar_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(pull_ar_pf_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size ; i++){
                    data = (CZSystemPVData)pv_data_body.elementAt(i);
                    val = data.data[PULL_AR];
                    y_pos_pull_ar[i] = (int)yPos(d.height,min,max,val);
                }

                //炉内圧プロファイル（肩）
                y_pos_shld_vac_pf = new int[size_shld];
                s_min = new Float(vac_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(vac_pf_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size_shld ; i++){
                    data = (CZSystemPVData)pv_data_shld.elementAt(i);
                    val = data.data[VAC_PF];
                    y_pos_shld_vac_pf[i] = (int)yPos(d.height,min,max,val);
                }

                //炉内圧プロファイル
                y_pos_vac_pf = new int[size];
                s_min = new Float(vac_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(vac_pf_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size ; i++){
                    data = (CZSystemPVData)pv_data_body.elementAt(i);
                    val = data.data[VAC_PF];
                    y_pos_vac_pf[i] = (int)yPos(d.height,min,max,val);
                }

                //炉内圧（肩）
                y_pos_shld_vac = new int[size_shld];
                s_min = new Float(vac_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(vac_pf_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size_shld ; i++){
                    data = (CZSystemPVData)pv_data_shld.elementAt(i);
                    val = data.data[VAC];
                    y_pos_shld_vac[i] = (int)yPos(d.height,min,max,val);
                }

                //炉内圧
                y_pos_vac = new int[size];
                s_min = new Float(vac_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(vac_pf_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size ; i++){
                    data = (CZSystemPVData)pv_data_body.elementAt(i);
                    val = data.data[VAC];
                    y_pos_vac[i] = (int)yPos(d.height,min,max,val);
                }
            }

            // X軸の座標を計算する。
            // @param w .. 幅,min .. 最小値,max .. 最大値,val ..データ
            // @return X軸の座標
            private float xPos(int w,float min,float max,float val){
                float x_dot = (w - offset_x) / (max - min);
                float x = x_dot * (val - min) + offset_x;
                return x;
            }

            // Y軸の座標を計算する。
            // @param h .. 幅,min .. 最小値,max .. 最大値,val ..データ
            // @return Y軸の座標
            private float yPos(int h,float min,float max,float val){
                float y_dot = (h - offset_y) / (max - min);
                float y = h - y_dot * (val - min) - offset_y;
                return y;
            }

            //
            // グラフ線を引く
            //
            private void drawLine(Graphics g){
                if(null == pv_data_shld) return;
                int size_shld = pv_data_shld.size();
                if(2 > size_shld) return;

                if(null == pv_data_body) return;
                int size = pv_data_body.size();
                if(2 > size) return;

                //プルアルゴンプロファイル（肩）
                if(grapane.isShld()){
                    g.setColor(java.awt.Color.white);
                    g.drawPolyline(x_pos_shld,y_pos_shld_pull_ar_pf,size_shld);
                }
                //プルアルゴンプロファイル
                g.setColor(java.awt.Color.white);
                g.drawPolyline(x_pos,y_pos_pull_ar_pf,size);

                //プルアルゴン（肩）
                if(grapane.isShld()){
                    g.setColor(PULL_AR_COL);
                    g.drawPolyline(x_pos_shld,y_pos_shld_pull_ar,size_shld);
                }
                //プルアルゴン
                g.setColor(PULL_AR_COL);
                g.drawPolyline(x_pos,y_pos_pull_ar,size);
                g.drawString("PULL.AR",x_pos[size-1],y_pos_pull_ar[size-1]);

                //炉内圧プロファイル（肩）
                if(grapane.isShld()){
                    g.setColor(java.awt.Color.white);
                    g.drawPolyline(x_pos_shld,y_pos_shld_vac_pf,size_shld);
                }
                //炉内圧プロファイル
                g.setColor(java.awt.Color.white);
                g.drawPolyline(x_pos,y_pos_vac_pf,size);

                //炉内圧（肩）
                if(grapane.isShld()){
                    g.setColor(VAC_COL);
                    g.drawPolyline(x_pos_shld,y_pos_shld_vac,size_shld);
                }
                //炉内圧
                g.setColor(VAC_COL);
                g.drawPolyline(x_pos,y_pos_vac,size);
                g.drawString("VAC",x_pos[size-1],y_pos_vac[size-1]);
            }

            //
            // 目盛線を引く
            //
            private void drawMemLine(Graphics g){
                float x;
                float y;
                float inc;

                Dimension d = getSize(null);

                // Ｘ軸目盛 小
                g.setColor(MEM_LINE3_COL);
                inc = (d.width - offset_x) / (gr_x_bun * 4.0f);
                for(x = offset_x ; x < d.width ; x+=inc){
                    g.drawLine((int)x,0,(int)x,d.height-offset_y);
                }
                // Ｙ軸目盛
                inc = (d.height - offset_y) / (gr_y_bun * 4.0f) ;
                for(y = d.height - offset_y ; y > 0 ; y-=inc){
                    g.drawLine((int)offset_x,(int)y,d.width,(int)y);
                }

                // Ｘ軸目盛 中
                g.setColor(MEM_LINE2_COL);
                inc = (d.width - offset_x) / (gr_x_bun * 2.0f);
                for(x = offset_x ; x < d.width ; x+=inc){
                    g.drawLine((int)x,0,(int)x,d.height-offset_y);
                }
                // Ｙ軸目盛
                inc = (d.height - offset_y) / (gr_y_bun * 2.0f) ;
                for(y = d.height - offset_y ; y > 0 ; y-=inc){
                    g.drawLine((int)offset_x,(int)y,d.width,(int)y);
                }

                // Ｘ軸目盛 大
                g.setColor(MEM_LINE1_COL);
                Float tmp = new Float(gr_x_length);
                float mem_inc = tmp.floatValue() / gr_x_bun;
                float x_val   = 0.0f;
                boolean cont = true;

                inc = (d.width - offset_x) / gr_x_bun ;
                for(x = offset_x ; x < d.width ; x+=inc){
                    g.drawLine((int)x,0,(int)x,d.height-offset_y);
                    if(cont){
                        g.drawString(String.valueOf(x_val),(int)x+3,d.height-8);
                    }
                    x_val+=mem_inc;
                    cont = !cont;
                }

                // Ｙ軸目盛
                Float tmp_min1 = new Float(pull_ar_pf_min_pro);
                Float tmp_max1 = new Float(pull_ar_pf_max_pro);
                float mem_inc1 = (tmp_max1.floatValue()-tmp_min1.floatValue()) / gr_y_bun;
                float y_val1  = tmp_min1.floatValue();

                Float tmp_min2 = new Float(vac_pf_min_pro);
                Float tmp_max2 = new Float(vac_pf_max_pro);
                float mem_inc2 = (tmp_max2.floatValue()-tmp_min2.floatValue()) / gr_y_bun;
                float y_val2  = tmp_min2.floatValue();

                inc = (d.height - offset_y) / gr_y_bun ;
                for(y = d.height - offset_y ; y > 0 ; y-=inc){
                    g.setColor(MEM_LINE1_COL);
                    g.drawLine((int)offset_x,(int)y,d.width,(int)y);

                    g.setColor(PULL_AR_COL);
                    g.drawString(String.valueOf(y_val1),3,(int)y - 0);
                    y_val1 += mem_inc1;

                    g.setColor(VAC_COL);
                    g.drawString(String.valueOf(y_val2),3,(int)y - 10);
                    y_val2 += mem_inc2;
                }
            }
        } // PressurePanel
    } // SimplGraphPanel

    /*******************************************************
     *
     *   マウス座標表示パネル
     *
     *******************************************************/
    public class MainMouseView extends JScrollPane {

        SubView view = null;    

        // ----------- コンストラクタ ----------------------
        MainMouseView(){

            super();
            setName("MainMouseView");
            setBorder(new Flush3DBorder());
            setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);
            setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
            view = new SubView();
            setViewportView(view);
            getViewport().setScrollMode(JViewport.BACKINGSTORE_SCROLL_MODE);
//@@            CZSystem.log("MainMouseView MainMouseView","new");
        }

        //
        //  
        //
        public void drawVal(){
            view.repaint();
        }

        /***************************************************
         * データ表示パネル
         ***************************************************/
        class SubView extends JPanel {

            // ---------- コンストラクタ -------------------
            SubView(){
                super();
                setLayout(null);
                setBackground(BACK_COL);
//@@                CZSystem.log("MainMouseView SubView","new");
            }

            //
            // マウス位置のデータを表示する。
            //
            public void paint(Graphics g){
                Dimension d = getSize(null);

                g.setColor(BACK_COL);
                g.fillRect(0,0,d.width,d.height);

                int x = 15;
                int y = 30;
                int inc = 20;

                //Ｘ軸
                g.setColor(MEM_LINE1_COL);
                g.drawString(String.valueOf(x_length_mouse),x,y);
                y+=inc;

                //メインヒーター１温度
                y+=inc;
                g.setColor(MAIN1_H_T_COL);
                g.drawString(String.valueOf(y_main1_h_t_mouse),x,y);

                //メインヒーター１温度プロファイル
                y+=inc;
                g.setColor(MAIN1_H_T_PF_COL);
                g.drawString(String.valueOf(y_main1_h_t_pf_mouse),x,y);

                //直径
                y+=inc;
                g.setColor(DIA_COL);
                g.drawString(String.valueOf(y_dia_mouse),x,y);

                //引き上げ速度
                y+=inc;
                g.setColor(SXL_ST_COL);
                g.drawString(String.valueOf(y_sxl_st_mouse),x,y);

                //引き上げ速度プロファイル
                y+=inc;
                g.setColor(SXL_ST_PF_COL);
                g.drawString(String.valueOf(y_sxl_st_pf_mouse),x,y);
            }
        } // SubView
    } // MainMouseView


    /*******************************************************
     *
     * 検索Dialog
     *
     *******************************************************/
    class SercheDialog extends JDialog {

        private JScrollPane bt_scpanel      = null;
        private JScrollPane bt_start_scpanel    = null;
        private JButton     read_button     = null;
        private JLabel      ro_name_lab     = null;

        //
        // ---------- コンストラクタ -----------------------
        //
        SercheDialog(){
            super();

            setTitle("検  索");
            setSize(820,335);
            setResizable(false);
            setModal(true);
            getContentPane().setLayout(null);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

			String s = CZSystem.RoKetaChg(ro_name);	// 20050725 炉：表示桁数変更
            ro_name_lab = new JLabel(s,JLabel.CENTER);
//            ro_name_lab = new JLabel(ro_name,JLabel.CENTER);
            ro_name_lab.setBounds(20, 20, 100, 30);
            ro_name_lab.setLocale(new Locale("ja","JP"));
            ro_name_lab.setFont(new java.awt.Font("dialog", 0, 18));
            ro_name_lab.setBorder(new Flush3DBorder());
            ro_name_lab.setForeground(java.awt.Color.black);
            getContentPane().add(ro_name_lab);

            bt_scpanel = new JScrollPane();
            bt_scpanel.setBounds(20, 60, 350, 187);
            getContentPane().add(bt_scpanel);

            bt_start_scpanel = new JScrollPane();
            bt_start_scpanel.setBounds(390, 60, 410, 187);
            getContentPane().add(bt_start_scpanel);

            read_button = new JButton("読み込み");
            read_button.setBounds(700, 270, 100, 24);
            read_button.setLocale(new Locale("ja","JP"));
            read_button.setFont(new java.awt.Font("dialog", 0, 18));
            read_button.setBorder(new Flush3DBorder());
            read_button.setForeground(java.awt.Color.black);
            read_button.addActionListener(new ReadButton());
            read_button.setEnabled(false);
            getContentPane().add(read_button);

//@@            CZSystem.log("CZTPGMain SercheDialog","new");
        }

        //
        // バッチ情報を表示する。
        // @return true
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

        //
        // バッチ情報を設定する。
        // @param v ... 
        // @return true
        public boolean setBtCondition(Vector v){
            removeBtCondition();
            BtStartTable t = new BtStartTable(v);
            JTableHeader tabHead = t.getTableHeader();
            tabHead.setReorderingAllowed(false);
            bt_start_scpanel.setViewportView(t);
            ro_bt_all_condition = v;
            return true;
        }

        //
        // バッチ情報を削除する。
        //
        public boolean removeBtCondition(){
            JViewport v;
            v =  bt_start_scpanel.getViewport();
            if(null != v.getView()) v.remove(v.getView());
            removeBtStart();
            read_button.setEnabled(false);
            return true;
        }

        /***************************************************
         * 読込みボタンの処理
         *  バッチ情報を再読込する。
         ***************************************************/
        class ReadButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
//@@                CZSystem.log("SercheDialog ReadButton","actionPerformed");
                Cursor cu_tmp = getCur();
                Cursor cu = new Cursor(Cursor.WAIT_CURSOR);
                setCur(cu);
                int ret = readBtPV();
                setCur(cu_tmp);
                if(1 > ret){
                    conpane.setData(false);
                    return;
                }
                serche_dia.setVisible(false);
                main_sc.setData();
                simpgrapane.setData();
                conpane.setData(true);
            }
        }

        /***************************************************
         *
         * バッチ№の一覧を表示する。
         *
         ***************************************************/
        class BtTable extends JTable {

            private Vector  bt_all_list     = null;
            private Vector  bt_list         = null;

            private BtTblMdl model = null;

            private boolean life = false;

            // ---------- コンストラクタ -------------------
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
                    for(int i = 0 ; i < bt_all_list.size() ; i++){
                        CZSystemBt bt = (CZSystemBt)bt_all_list.elementAt(i);

                        if(0 == bt.renban) bt_list.addElement(bt);
                        if(-1 == bt.renban) bt_list.addElement(bt);
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

                    // 登録日時
                    colum = cmdl.getColumn(2);
                    colum.setMaxWidth(162);
                    colum.setMinWidth(162);
                    colum.setWidth(162);

                }
                catch (Throwable e) {
                    CZSystem.handleException(e);
                }
            }

            // バッチ№選択時の処理
            //
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
                Vector v = new Vector(50);
                CZSystemBt bt = (CZSystemBt)bt_list.elementAt(row);
                for(int i = 0 ; i < bt_all_list.size() ; i++){
                    CZSystemBt bt_tmp = (CZSystemBt)bt_all_list.elementAt(i);
                    if(bt.batch.equals(bt_tmp.batch)) v.addElement(bt_tmp);
                }
                setBtCondition(v);
            }

            //
            //
            //
            public void setData(int gr,int tbl){
            }
        } // BtTable

        /***************************************************
         *
         * バッチ№実績一覧：モデル
         *
         ***************************************************/
        public class BtTblMdl extends AbstractTableModel {

            private int TBL_ROW             = 0;
            final   int TBL_COL             = 3;
            private Vector  bt_list         = null;

            final String[] names = {" # "  , "Bt" , "登録日時" };

            private Object  data[][];

            // ---------- コンストラクタ -------------------
            // @param v バッチ情報
            BtTblMdl(Vector v){
                super();
                bt_list = v;
                TBL_ROW = bt_list.size();
                data = new Object[TBL_ROW][TBL_COL];
                for(int i = 0 ; i < TBL_ROW ; i++){
                    CZSystemBt bt = (CZSystemBt)bt_list.elementAt(i);
                    if(null == bt) break;
                    data[i][0] = new Integer(i+1);  // #
                    data[i][1] = bt.batch;          // Bt
                    data[i][2] = bt.t_time;         // 登録日時
                }
            }

            // 列数を取得する。
            // @return 列数
            public int getColumnCount(){
                return TBL_COL;
            }

            // 行数を取得する。
            // @return 行数
            public int getRowCount(){
                return TBL_ROW;
            }

            // 値を取得する。
            // @param row ... 行, col ... 列
            // @return 値
            public Object getValueAt(int row, int col){
                return data[row][col];
            }

            // 列名を取得する。
            // @param column ... 列
            // @return 列名
            public String getColumnName(int column){
                return names[column];
            }

            // 列のデータ型を取得する。
            // @param c ... 列
            // @return データの型
            public Class getColumnClass(int c){
                return getValueAt(0, c).getClass();
            }

            // セルの編集可否を取得する。
            // @param row ... 行, col ... 列
            // @return true :可, false:否
            public boolean isCellEditable(int row, int col){
                return false;
            }

            // 値を設定する。
            // @param aValue ... 値, row ... 行, column ... 列
            public void setValueAt(Object aValue, int row, int column){
                data[row][column] = aValue;
            }
        } // BtTblMdl


        /***************************************************
         *
         *       Ｂｔスタート時間一覧
         *
         ***************************************************/
        class BtStartTable extends JTable {

            private Vector  bt_list         = null;
            private Vector  bt_start_list   = null;

            private BtStartTblMdl model = null;

            private boolean life = false;

            // ---------- コンストラクタ -------------------
            // @param v バッチ情報
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

                    for(int i = 0 ; i < size ; i++){
                        CZSystemStart st = (CZSystemStart)tmp.elementAt(i);
                        if(null == st) break;
                       //Body             //Body 1
                        if((7 == st.p_no) && (1 == st.sp_no))
                        bt_start_list.addElement(st);   
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

            //
            // 選択時の処理
            //
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
                CZSystemStart st = (CZSystemStart)bt_start_list.elementAt(row);
                setBtStart(st);
                read_button.setEnabled(true);
            }

            public void setData(int gr,int tbl){

            }
        }

        /***************************************************
         *
         *       Ｂｔスタート時間一覧：モデル
         *
         ***************************************************/
        public class BtStartTblMdl extends AbstractTableModel {

            private int TBL_ROW             = 0;
            final   int TBL_COL             = 6;
            private Vector  bt_start_list   = null;

            final String[] names = {" # "  , "PNo" ,
                                        "SPNo","PSeq"  ,
                                        "プロセス",
                                        "登録日時" };
            private Object  data[][];

            // ---------- コンストラクタ -------------------
            // @param v バッチ情報
            BtStartTblMdl(Vector v){
                super();
                bt_start_list = v;
                TBL_ROW = bt_start_list.size();
                data = new Object[TBL_ROW][TBL_COL];
                for(int i = 0 ; i < TBL_ROW ; i++){
                    CZSystemStart st = (CZSystemStart)bt_start_list.elementAt(i);
                    if(null == st) break;
                    data[i][0] = new Integer(i+1);                  // #
                    data[i][1] = new Integer(st.p_no);              // PNo
                    data[i][2] = new Integer(st.sp_no);             // SPNo
                    data[i][3] = new Integer(st.p_renban);          // PSeq
                    data[i][4] = CZSystem.getProcName(st.p_no);     // プロセス
                    data[i][5] = st.p_start;                        // 登録日時
                }
            }

            // 列数を取得する。
            // @return 列数
            public int getColumnCount(){
                return TBL_COL;
            }

            // 行数を取得する。
            // @return 行数
            public int getRowCount(){
                return TBL_ROW;
            }

            // 値を取得する。
            // @param row .. 行, col .. 列
            // @return 値
            public Object getValueAt(int row, int col){
                return data[row][col];
            }

            // 列名を取得する。
            // @param column ... 列
            // @return 列名
            public String getColumnName(int column){
                return names[column];
            }

            // データの型を取得する。
            // @param c ... 列
            // @return データの型
            public Class getColumnClass(int c){
                return getValueAt(0, c).getClass();
            }

            // 編集可否を取得する。
            // @param row .. 行,col .. 列
            // @return true .. 可, false .. 否
            public boolean isCellEditable(int row, int col){
                return false;
            }

            // 列数を取得する。
            // @param aValue .. ,row .. 行,column .. 列
            public void setValueAt(Object aValue, int row, int column){
                data[row][column] = aValue;
            }
        }
    } // SercheDialog

    /*******************************************************
     *  Y軸項目設定Dialog
     *  @@ T6を追加する必要あり
     *******************************************************/
    class YLengDialog extends JDialog {

        private LengText ht_min         = null;
        private LengText ht_max         = null;
        private LengText ht_pf_min      = null;
        private LengText ht_pf_max      = null;

        private LengText dia_min        = null;
        private LengText dia_max        = null;
        private LengText dia_pf_min     = null;
        private LengText dia_pf_max     = null;

        private LengText fp_min         = null;
        private LengText fp_max         = null;
        private LengText fp_pf_min      = null;
        private LengText fp_pf_max      = null;

        private LengText sxl_rt_min     = null;
        private LengText sxl_rt_max     = null;
        private LengText cru_rt_min     = null;
        private LengText cru_rt_max     = null;

        private LengText pull_ar_min    = null;
        private LengText pull_ar_max    = null;
        private LengText vac_min        = null;
        private LengText vac_max        = null;

        private JLabel yLen_ro_name_lab = null;     //炉番表示

        //
        // ---------- コンストラクタ -----------------------
        YLengDialog(){
            super();

            setTitle("項目設定");
            setSize(440,565);
            setResizable(false);
            setModal(true);
            getContentPane().setLayout(null);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

			String s = CZSystem.RoKetaChg(ro_name);	// 20050725 炉：表示桁数変更
            yLen_ro_name_lab = new JLabel(s,JLabel.CENTER);
//            yLen_ro_name_lab = new JLabel(ro_name,JLabel.CENTER);
            yLen_ro_name_lab.setBounds(20, 20, 100, 30);
            yLen_ro_name_lab.setLocale(new Locale("ja","JP"));
            yLen_ro_name_lab.setFont(new java.awt.Font("dialog", 0, 18));
            yLen_ro_name_lab.setBorder(new Flush3DBorder());
            yLen_ro_name_lab.setForeground(java.awt.Color.black);
            getContentPane().add(yLen_ro_name_lab);

            JButton set_button = new JButton("設  定");
            set_button.setBounds(320, 500, 100, 24);
            set_button.setLocale(new Locale("ja","JP"));
            set_button.setFont(new java.awt.Font("dialog", 0, 18));
            set_button.setBorder(new Flush3DBorder());
            set_button.setForeground(java.awt.Color.black);
            set_button.addActionListener(new SetButton());
            getContentPane().add(set_button);

            JLabel lab = null;

            lab = new JLabel("Ｍｉｎ",JLabel.CENTER);
            lab.setBounds(220, 70, 100, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            lab = new JLabel("Ｍａｘ",JLabel.CENTER);
            lab.setBounds(320, 70, 100, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            int Y   = 105;
            int INC = 35;

            int y   = Y;
            int inc = INC;
            lab = new JLabel("メインヒーター１温度",JLabel.CENTER);
            lab.setBounds(20, y, 200, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            y+=inc;
            lab = new JLabel("メインヒーター１温度ＰＦ",JLabel.CENTER);
            lab.setBounds(20, y, 200, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            y+=inc;
            lab = new JLabel("直径",JLabel.CENTER);
            lab.setBounds(20, y, 200, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            y+=inc;
            lab = new JLabel("引き上げ速度",JLabel.CENTER);
            lab.setBounds(20, y, 200, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            y+=inc;
            lab = new JLabel("引き上げ速度ＰＦ",JLabel.CENTER);
            lab.setBounds(20, y, 200, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            // 拡張
            y+=inc;
            y+=inc;
            lab = new JLabel("直径管理",JLabel.CENTER);
            lab.setBounds(20, y, 200, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            y+=inc;
            lab = new JLabel("シード回転",JLabel.CENTER);
            lab.setBounds(20, y, 200, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            y+=inc;
            lab = new JLabel("ルツボ回転",JLabel.CENTER);
            lab.setBounds(20, y, 200, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            y+=inc;
            lab = new JLabel("プルアルゴン",JLabel.CENTER);
            lab.setBounds(20, y, 200, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            y+=inc;
            lab = new JLabel("炉内圧",JLabel.CENTER);
            lab.setBounds(20, y, 200, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            // 入力領域

            //
            // ヒーター温度関係
            y   = Y;
            inc = INC;
            ht_min = new LengText();
            ht_min.setBounds(220, y, 100, 24);
            ht_min.setForeground(MAIN1_H_T_COL);
            getContentPane().add(ht_min);

            ht_max = new LengText();
            ht_max.setBounds(320, y, 100, 24);
            ht_max.setForeground(MAIN1_H_T_COL);
            getContentPane().add(ht_max);

            y+=inc;
            ht_pf_min = new LengText();
            ht_pf_min.setBounds(220, y, 100, 24);
            ht_pf_min.setForeground(MAIN1_H_T_PF_COL);
            getContentPane().add(ht_pf_min);

            ht_pf_max = new LengText();
            ht_pf_max.setBounds(320, y, 100, 24);
            ht_pf_max.setForeground(MAIN1_H_T_PF_COL);
            getContentPane().add(ht_pf_max);

            // 直径度関係
            y+=inc;
            dia_min = new LengText();
            dia_min.setBounds(220, y, 100, 24);
            dia_min.setForeground(DIA_COL);
            getContentPane().add(dia_min);

            dia_max = new LengText();
            dia_max.setBounds(320, y, 100, 24);
            dia_max.setForeground(DIA_COL);
            getContentPane().add(dia_max);

            // 引き上げ速度関係
            y+=inc;
            fp_min = new LengText();
            fp_min.setBounds(220, y, 100, 24);
            fp_min.setForeground(SXL_ST_COL);
            getContentPane().add(fp_min);

            fp_max = new LengText();
            fp_max.setBounds(320, y, 100, 24);
            fp_max.setForeground(SXL_ST_COL);
            getContentPane().add(fp_max);

            y+=inc;
            fp_pf_min = new LengText();
            fp_pf_min.setBounds(220, y, 100, 24);
            fp_pf_min.setForeground(SXL_ST_PF_COL);
            getContentPane().add(fp_pf_min);

            fp_pf_max = new LengText();
            fp_pf_max.setBounds(320, y, 100, 24);
            fp_pf_max.setForeground(SXL_ST_PF_COL);
            getContentPane().add(fp_pf_max);

            // 拡張
            // 直径度関係
            y+=inc;
            y+=inc;
            dia_pf_min = new LengText();
            dia_pf_min.setBounds(220, y, 100, 24);
            dia_pf_min.setForeground(DIA_PF_COL);
            getContentPane().add(dia_pf_min);

            dia_pf_max = new LengText();
            dia_pf_max.setBounds(320, y, 100, 24);
            dia_pf_max.setForeground(DIA_PF_COL);
            getContentPane().add(dia_pf_max);

            // シード回転
            y+=inc;
            sxl_rt_min = new LengText();
            sxl_rt_min.setBounds(220, y, 100, 24);
            sxl_rt_min.setForeground(SXL_RT_COL);
            getContentPane().add(sxl_rt_min);

            sxl_rt_max = new LengText();
            sxl_rt_max.setBounds(320, y, 100, 24);
            sxl_rt_max.setForeground(SXL_RT_COL);
            getContentPane().add(sxl_rt_max);

            // ルツボ回転
            y+=inc;
            cru_rt_min = new LengText();
            cru_rt_min.setBounds(220, y, 100, 24);
            cru_rt_min.setForeground(CRU_RT_COL);
            getContentPane().add(cru_rt_min);

            cru_rt_max = new LengText();
            cru_rt_max.setBounds(320, y, 100, 24);
            cru_rt_max.setForeground(CRU_RT_COL);
            getContentPane().add(cru_rt_max);

            // プルアルゴン
            y+=inc;
            pull_ar_min = new LengText();
            pull_ar_min.setBounds(220, y, 100, 24);
            pull_ar_min.setForeground(PULL_AR_COL);
            getContentPane().add(pull_ar_min);

            pull_ar_max = new LengText();
            pull_ar_max.setBounds(320, y, 100, 24);
            pull_ar_max.setForeground(PULL_AR_COL);
            getContentPane().add(pull_ar_max);

            // 炉内圧
            y+=inc;
            vac_min = new LengText();
            vac_min.setBounds(220, y, 100, 24);
            vac_min.setForeground(VAC_COL);
            getContentPane().add(vac_min);

            vac_max = new LengText();
            vac_max.setBounds(320, y, 100, 24);
            vac_max.setForeground(VAC_COL);
            getContentPane().add(vac_max);

//@@            CZSystem.log("CZTPGMain YLengDialog","new [" + y + "]");
        }

        // 現状値を設定する。
        //
        public boolean setDefault(){
			String s = CZSystem.RoKetaChg(ro_name);	// 20050725 炉：表示桁数変更
            yLen_ro_name_lab.setText(s);
//            yLen_ro_name_lab.setText(ro_name);

            // ヒーター温度関係
            ht_min.setText(main1_h_t_min_pro);
            ht_max.setText(main1_h_t_max_pro);
            ht_pf_min.setText(main1_h_t_pf_min_pro);
            ht_pf_max.setText(main1_h_t_pf_max_pro);

            // 直径度関係
            dia_min.setText(dia_min_pro);
            dia_max.setText(dia_max_pro);
            dia_pf_min.setText(dia_pf_min_pro);
            dia_pf_max.setText(dia_pf_max_pro);

            // 引き上げ速度関係
            fp_min.setText(sxl_st_min_pro);
            fp_max.setText(sxl_st_max_pro);
            fp_pf_min.setText(sxl_st_pf_min_pro);
            fp_pf_max.setText(sxl_st_pf_max_pro);

            // シード回転
            sxl_rt_min.setText(sxl_rt_pf_min_pro);
            sxl_rt_max.setText(sxl_rt_pf_max_pro);

            // ルツボ回転
            cru_rt_min.setText(cru_rt_pf_min_pro);
            cru_rt_max.setText(cru_rt_pf_max_pro);

            // プルアルゴン
            pull_ar_min.setText(pull_ar_pf_min_pro);
            pull_ar_max.setText(pull_ar_pf_max_pro);

            // 炉内圧
            vac_min.setText(vac_pf_min_pro);
            vac_max.setText(vac_pf_max_pro);
            return true;
        }

        // Y軸を設定する
        //
        private boolean setYLang(){

            // ヒーター温度関係
            main1_h_t_min_pro       = ht_min.getText();
            main1_h_t_max_pro       = ht_max.getText();
            main1_h_t_pf_min_pro    = ht_pf_min.getText();
            main1_h_t_pf_max_pro    = ht_pf_max.getText();

            // 直径度関係
            dia_min_pro = dia_min.getText();
            dia_max_pro = dia_max.getText();
            dia_pf_min_pro = dia_pf_min.getText();
            dia_pf_max_pro = dia_pf_max.getText();

            // 引き上げ速度関係
            sxl_st_min_pro = fp_min.getText();
            sxl_st_max_pro = fp_max.getText();
            sxl_st_pf_min_pro = fp_pf_min.getText();
            sxl_st_pf_max_pro = fp_pf_max.getText();

            // シード回転
            sxl_rt_pf_min_pro = sxl_rt_min.getText();
            sxl_rt_pf_max_pro = sxl_rt_max.getText();

            // ルツボ回転
            cru_rt_pf_min_pro = cru_rt_min.getText();
            cru_rt_pf_max_pro = cru_rt_max.getText();

            // プルアルゴン
            pull_ar_pf_min_pro = pull_ar_min.getText();
            pull_ar_pf_max_pro = pull_ar_max.getText();

            // 炉内圧
            vac_pf_min_pro = vac_min.getText();
            vac_pf_max_pro = vac_max.getText();

            chgYLength();
            return true;
        }

        /***************************************************
         *
         *       設定ボタンの処理
         *
         ***************************************************/
        class SetButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                setYLang();
//@@                CZSystem.log("YLengDialog SetButton","actionPerformed");
            }
        } // SetButton

        /***************************************************
         *
         *       項目Min,Maxを入力するTextField
         *
         ***************************************************/
        public class LengText extends JTextField {

            LengText(){
                super();
                setFont(new java.awt.Font("dialog", 0, 16));
            }

            //
            //
            protected Document createDefaultModel() {
                return new NumericDocument();
            }

            //
            //
            class NumericDocument extends PlainDocument {
                String validValues = "0123456789.-";

                //
                //
                public void insertString( int offset, String str, AttributeSet a )
                                                        throws BadLocationException {
                    if(10 < getLength()) return;
                    char[] val = str.toCharArray();
                    for (int i = 0; i < val.length; i++) {
                        if(validValues.indexOf(val[i]) == -1) return;
                    }
                    super.insertString( offset, str, a );
                }
            }
        } // LengText
    } // YLengDialog

    /*******************************************************
     *
     *       メイングラフ
     *
     *******************************************************/
    public class MainSc extends JScrollPane {

        private Rectangle   view_rec    = null;
        private View        view        = null;

        // ---------- コンストラクタ -----------------------
        // @param comp ... Event Listener
        MainSc(PVGrEventCompo comp){
            super();

            setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);
            setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
            view = new View();
            setViewportView(view);
            comp.setMainView(view);
            view.addComponentListener(comp);
            getViewport().setScrollMode(JViewport.BACKINGSTORE_SCROLL_MODE);
//@@            CZSystem.log("CZTPGMain MainSc","new");
        }

        //
        //
        public void setDefault(){
            view_rec = getViewportBorderBounds();
            view.setPreferredSize(new Dimension(view_rec.width,view_rec.height*Y_VIEW_TIMES));
            view.setLocation(0, -(view_rec.height* Y_VIEW_TIMES - view_rec.height));
            view.setViewRec(view_rec);
        }

        // X軸
        //
        public void chgXSize(){
            Float x_def = new Float(GR_X_LENGTH_DEF);
            Float x_new = new Float(gr_x_length);
            float new_size = (x_def.floatValue()/x_new.floatValue()) * view_rec.width;
            
            Dimension d = view.getSize(null);
            view.setSize(new Dimension((int)new_size,d.height));
            setData();
        }

        // Y軸
        //
        public void chgYSize(){
            setData();
        }

        //
        // データを設定し、グラフを再描画する。
        //
        public void setData(){
            view.setData();
            view.repaint();
        }

        /***************************************************
         *  グラフ描画パネル
         ***************************************************/
        class View extends JPanel {
            Rectangle view_rec = null;
            int x_pos[];

            int y_pos_ht[];
            int y_pos_ht_pf[];
            int y_pos_ht_conv[];
            int y_pos_ht_pf_conv[];

            int y_pos_sxl_st[];
            int y_pos_sxl_st_pf[];

            int y_pos_dia[];
            int y_pos_dia_pf[];
            int y_pos_dia_pf_min[];
            int y_pos_dia_pf_max[];

            int x_pos_shld[];

            int y_pos_shld_ht[];
            int y_pos_shld_ht_pf[];
            int y_pos_shld_ht_conv[];
            int y_pos_shld_ht_pf_conv[];

            int y_pos_shld_sxl_st[];
            int y_pos_shld_sxl_st_pf[];

            int y_pos_shld_dia[];
            int y_pos_shld_dia_pf[];
            int y_pos_shld_dia_pf_min[];
            int y_pos_shld_dia_pf_max[];

            // ---------- コンストラクタ -------------------
            //
            View(){
                super();
                setName("View");
                setLayout(null);
                setBackground(BACK_COL);
                addMouseMotionListener(new MainViewMouseMotion());
            }

            // 枠を設定する。
            //
            public void setViewRec(Rectangle rec){
                view_rec = rec;
            }

            // グラフを描画する。
            //
            public void paint(Graphics g){
                Dimension d = getSize(null);

                g.setColor(BACK_COL);
                g.fillRect(0,0,d.width,d.height);
                drawMemLine(g);
                drawLine(g);
            }

            // 目盛線を描画する。
            //
            private void drawMemLine(Graphics g){
                float x;
                float y;
                float inc;

                Dimension d = getSize(null);

                g.setColor(MEM_LINE3_COL);
                inc = view_rec.width / (gr_x_bun * 4);
                for(x = 0.0f ;  d.width > x ; x+=inc){
                    g.drawLine((int)x,0,(int)x,d.height);
                }
                inc = view_rec.height / (gr_y_bun * 4);
                for(y = (float)d.height ;  0 < y ; y-=inc){
                    g.drawLine(0,(int)y,d.width,(int)y);
                }

                g.setColor(MEM_LINE2_COL);
                inc = view_rec.width / (gr_x_bun * 2);
                for(x = 0.0f ;  d.width > x ; x+=inc){
                    g.drawLine((int)x,0,(int)x,d.height);
                }
                inc = view_rec.height / (gr_y_bun * 2);
                for(y = (float)d.height ;  0 < y ; y-=inc){
                    g.drawLine(0,(int)y,d.width,(int)y);
                }

                g.setColor(MEM_LINE1_COL);
                inc = view_rec.width / gr_x_bun;
                for(x = 0.0f ;  d.width > x ; x+=inc){
                    g.drawLine((int)x,0,(int)x,d.height);
                }
                inc = view_rec.height / gr_y_bun;
                for(y = (float)d.height ;  0 < y ; y-=inc){
                    g.drawLine(0,(int)y,d.width,(int)y);
                }
            }

            // グラフ線を描画する。
            //
            private void drawLine(Graphics g){
                if(null == pv_data_shld) return;
                int size_shld = pv_data_shld.size();
                if(2 > size_shld) return;

                if(null == pv_data_body) return;
                int size = pv_data_body.size();
                if(2 > size) return;


                //引き上げ速度（肩）
                if(grapane.isShld()){
                    g.setColor(SXL_ST_COL);
                    g.drawPolyline(x_pos_shld,y_pos_shld_sxl_st,size_shld);
                }

                //引き上げ速度
                g.setColor(SXL_ST_COL);
                g.drawPolyline(x_pos,y_pos_sxl_st,size);
                g.drawString("SXL.ST",x_pos[size-1],y_pos_sxl_st[size-1]);

                //引き上げ速度プロファイル（肩）
                if(grapane.isShld()){
                    g.setColor(SXL_ST_PF_COL);
                    g.drawPolyline(x_pos_shld,y_pos_shld_sxl_st_pf,size_shld);
                }

                //引き上げ速度プロファイル
                g.setColor(SXL_ST_PF_COL);
                g.drawPolyline(x_pos,y_pos_sxl_st_pf,size);
                g.drawString("SXS.PF",x_pos[size-1],y_pos_sxl_st_pf[size-1]);

                //メインヒーター１温度（肩）
                if(grapane.isShld()){
                    //メインヒーター１温度とプロファイル
                    g.setColor(java.awt.Color.white);
                    g.drawPolyline(x_pos_shld,y_pos_shld_ht_conv,size_shld);
                    //メインヒーター１温度
                    g.setColor(MAIN1_H_T_COL);
                    g.drawPolyline(x_pos_shld,y_pos_shld_ht,size_shld);
                }

                //メインヒーター１温度
                //メインヒーター１温度とプロファイル
                g.setColor(java.awt.Color.white);
                g.drawPolyline(x_pos,y_pos_ht_conv,size);
                //メインヒーター１温度
                g.setColor(MAIN1_H_T_COL);
                g.drawPolyline(x_pos,y_pos_ht,size);
                g.drawString("HEA.T1",x_pos[size-1],y_pos_ht[size-1]);

                //メインヒーター１温度プロファイル（肩）
                if(grapane.isShld()){
                    //プロファイルとメインヒーター１
                    g.setColor(java.awt.Color.lightGray);
                    g.drawPolyline(x_pos_shld,y_pos_shld_ht_pf_conv,size_shld);
                    //メインヒーター１温度プロファイル
                    g.setColor(MAIN1_H_T_PF_COL);
                    g.drawPolyline(x_pos_shld,y_pos_shld_ht_pf,size_shld);
                }

                //メインヒーター１温度プロファイル
                //プロファイルとメインヒーター１
                g.setColor(java.awt.Color.lightGray);
                g.drawPolyline(x_pos,y_pos_ht_pf_conv,size);
                //メインヒーター１温度プロファイル
                g.setColor(MAIN1_H_T_PF_COL);
                g.drawPolyline(x_pos,y_pos_ht_pf,size);
                g.drawString("HT1.PF",x_pos[size-1],y_pos_ht_pf[size-1]);

                //直径プロファイル（肩）
                if(grapane.isShld()){
                    g.setColor(DIA_PF_COL);
                    g.drawPolyline(x_pos_shld,y_pos_shld_dia_pf_max,size_shld);
                    g.drawPolyline(x_pos_shld,y_pos_shld_dia_pf_min,size_shld);
                    g.drawPolyline(x_pos_shld,y_pos_shld_dia_pf,size_shld);
                }

                //直径プロファイル
                g.setColor(DIA_PF_COL);
                g.drawPolyline(x_pos,y_pos_dia_pf_max,size);
                g.drawPolyline(x_pos,y_pos_dia_pf_min,size);
                g.drawPolyline(x_pos,y_pos_dia_pf,size);
                g.drawString("DIA.PF",x_pos[size-1],y_pos_dia_pf[size-1]);

                //直径（肩）
                if(grapane.isShld()){
                    g.setColor(DIA_COL);
                    g.drawPolyline(x_pos_shld,y_pos_shld_dia,size_shld);
                }

                //直径
                g.setColor(DIA_COL);
                g.drawPolyline(x_pos,y_pos_dia,size);
                g.drawString("DIA",x_pos[size-1],y_pos_dia[size-1]);
            }

            // データから座標を計算する。
            //
            private void setData(){
                if(null == pv_data_shld) return;
                int size_shld = pv_data_shld.size();
                if(2 > size_shld) return;

                if(null == pv_data_body) return;
                int size = pv_data_body.size();
                if(2 > size) return;

                Float x_max;

                Float s_min;
                float min;
                Float s_max;
                float max;
                
                float tmp;

                Float ftmp1;
                Float ftmp2;
                float tmp1;
                float tmp2;

                float val;
                float val_shld;

                CZSystemPVData data;

                Dimension d = getSize(null);

                //Ｘ軸座標計算（肩）
                val = 0.0f;
                val_shld = 0.0f;
                x_pos_shld = new int[size_shld];
                x_max = new Float(gr_x_length);
                min = 0.0f;
                max = x_max.floatValue();

                for(int i = 0 ; i < size_shld ; i++){
                    data = (CZSystemPVData)pv_data_shld.elementAt(i);
                    val = data.p_length;
                    x_pos_shld[i] = (int)xPos(d.width,view_rec.width,min,max,val);
                }
                
                if(grapane.isShld()){
                    val_shld = val;
                }

                //Ｘ軸座標計算
                x_pos = new int[size];
                x_max = new Float(gr_x_length);
                min = 0.0f;
                max = x_max.floatValue();

                for(int i = 0 ; i < size ; i++){
                    data = (CZSystemPVData)pv_data_body.elementAt(i);
                    val = data.p_length + val_shld;
                    x_pos[i] = (int)xPos(d.width,view_rec.width,min,max,val);
                }

                //メインヒーター１温度（肩）
                y_pos_shld_ht = new int[size_shld];
                s_min = new Float(main1_h_t_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(main1_h_t_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size_shld ; i++){
                    data = (CZSystemPVData)pv_data_shld.elementAt(i);
                    val = data.data[MAIN1_H_T];
                    y_pos_shld_ht[i] = (int)yPos(d.height,view_rec.height,min,max,val);
                }

                //メインヒーター１温度
                y_pos_ht = new int[size];
                s_min = new Float(main1_h_t_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(main1_h_t_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size ; i++){
                    data = (CZSystemPVData)pv_data_body.elementAt(i);
                    val = data.data[MAIN1_H_T];
                    y_pos_ht[i] = (int)yPos(d.height,view_rec.height,min,max,val);
                }

                //メインヒーター１温度とプロファイル（肩）
                data = (CZSystemPVData)pv_data_shld.elementAt(0);
                tmp = data.data[MAIN1_H_T];

                y_pos_shld_ht_conv = new int[size_shld];
                s_min = new Float(main1_h_t_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(main1_h_t_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size_shld ; i++){
                    data = (CZSystemPVData)pv_data_shld.elementAt(i);
                    val = data.data[MAIN1_H_T_PF] + tmp;
                    y_pos_shld_ht_conv[i] = (int)yPos(d.height,view_rec.height,min,max,val);
                }

                //メインヒーター１温度とプロファイル
                data = (CZSystemPVData)pv_data_body.elementAt(0);
                tmp = data.data[MAIN1_H_T];

                y_pos_ht_conv = new int[size];
                s_min = new Float(main1_h_t_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(main1_h_t_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size ; i++){
                    data = (CZSystemPVData)pv_data_body.elementAt(i);
                    val = data.data[MAIN1_H_T_PF] + tmp;
                    y_pos_ht_conv[i] = (int)yPos(d.height,view_rec.height,min,max,val);
                }

                //メインヒーター１温度プロファイル（肩）
                y_pos_shld_ht_pf = new int[size_shld];
                s_min = new Float(main1_h_t_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(main1_h_t_pf_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size_shld ; i++){
                    data = (CZSystemPVData)pv_data_shld.elementAt(i);
                    val = data.data[MAIN1_H_T_PF];
                    y_pos_shld_ht_pf[i] = (int)yPos(d.height,view_rec.height,min,max,val);
                }

                //メインヒーター１温度プロファイル
                y_pos_ht_pf = new int[size];
                s_min = new Float(main1_h_t_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(main1_h_t_pf_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size ; i++){
                    data = (CZSystemPVData)pv_data_body.elementAt(i);
                    val = data.data[MAIN1_H_T_PF];
                    y_pos_ht_pf[i] = (int)yPos(d.height,view_rec.height,min,max,val);
                }

                //メインヒーター１温度プロファイルとメインヒーター１温度（肩）
                data = (CZSystemPVData)pv_data_shld.elementAt(0);
                tmp = data.data[MAIN1_H_T];

                y_pos_shld_ht_pf_conv = new int[size_shld];
                s_min = new Float(main1_h_t_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(main1_h_t_pf_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size_shld ; i++){
                    data = (CZSystemPVData)pv_data_shld.elementAt(i);
                    val = data.data[MAIN1_H_T] - tmp;
                    y_pos_shld_ht_pf_conv[i] = (int)yPos(d.height,view_rec.height,min,max,val);
                }

                //メインヒーター１温度プロファイルとメインヒーター１温度
                data = (CZSystemPVData)pv_data_body.elementAt(0);
                tmp = data.data[MAIN1_H_T];

                y_pos_ht_pf_conv = new int[size];
                s_min = new Float(main1_h_t_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(main1_h_t_pf_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size ; i++){
                    data = (CZSystemPVData)pv_data_body.elementAt(i);
                    val = data.data[MAIN1_H_T] - tmp;
                    y_pos_ht_pf_conv[i] = (int)yPos(d.height,view_rec.height,min,max,val);
                }

                //直径（肩）
                y_pos_shld_dia = new int[size_shld];
                s_min = new Float(dia_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(dia_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size_shld ; i++){
                    data = (CZSystemPVData)pv_data_shld.elementAt(i);
                    val = data.data[DIA];
                    y_pos_shld_dia[i] = (int)yPos(d.height,view_rec.height,min,max,val);
                }

                //直径
                y_pos_dia = new int[size];
                s_min = new Float(dia_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(dia_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size ; i++){
                    data = (CZSystemPVData)pv_data_body.elementAt(i);
                    val = data.data[DIA];
                    y_pos_dia[i] = (int)yPos(d.height,view_rec.height,min,max,val);
                }

                //直径プロファイル（肩）
                y_pos_shld_dia_pf_max = new int[size_shld];
                y_pos_shld_dia_pf     = new int[size_shld];
                y_pos_shld_dia_pf_min = new int[size_shld];

                ftmp1 = new Float(dia_pf_min_pro);
                tmp1  = ftmp1.floatValue();
                ftmp2 = new Float(dia_pf_max_pro);
                tmp2  = ftmp2.floatValue();

                s_min = new Float(dia_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(dia_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size_shld ; i++){
                    data = (CZSystemPVData)pv_data_shld.elementAt(i);
                    val = data.data[DIA_PF];
                    y_pos_shld_dia_pf[i]     = (int)yPos(d.height,view_rec.height,min,max,val);
                    y_pos_shld_dia_pf_min[i] = (int)yPos(d.height,view_rec.height,min,max,val+tmp1);
                    y_pos_shld_dia_pf_max[i] = (int)yPos(d.height,view_rec.height,min,max,val+tmp2);
                }

                //直径プロファイル
                y_pos_dia_pf_max = new int[size];
                y_pos_dia_pf     = new int[size];
                y_pos_dia_pf_min = new int[size];

                ftmp1 = new Float(dia_pf_min_pro);
                tmp1  = ftmp1.floatValue();
                ftmp2 = new Float(dia_pf_max_pro);
                tmp2  = ftmp2.floatValue();

                s_min = new Float(dia_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(dia_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size ; i++){
                    data = (CZSystemPVData)pv_data_body.elementAt(i);
                    val = data.data[DIA_PF];
                    y_pos_dia_pf[i]     = (int)yPos(d.height,view_rec.height,min,max,val);
                    y_pos_dia_pf_min[i] = (int)yPos(d.height,view_rec.height,min,max,val+tmp1);
                    y_pos_dia_pf_max[i] = (int)yPos(d.height,view_rec.height,min,max,val+tmp2);
                }

                //引き上げ速度（肩）
                y_pos_shld_sxl_st = new int[size_shld];
                s_min = new Float(sxl_st_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(sxl_st_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size_shld ; i++){
                    data = (CZSystemPVData)pv_data_shld.elementAt(i);
                    val = data.data[SXL_ST];
                    y_pos_shld_sxl_st[i] = (int)yPos(d.height,view_rec.height,min,max,val);
                }

                //引き上げ速度
                y_pos_sxl_st = new int[size];
                s_min = new Float(sxl_st_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(sxl_st_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size ; i++){
                    data = (CZSystemPVData)pv_data_body.elementAt(i);
                    val = data.data[SXL_ST];
                    y_pos_sxl_st[i] = (int)yPos(d.height,view_rec.height,min,max,val);
                }

                //引き上げ速度プロファイル（肩）
                y_pos_shld_sxl_st_pf = new int[size_shld];
                s_min = new Float(sxl_st_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(sxl_st_pf_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size_shld ; i++){
                    data = (CZSystemPVData)pv_data_shld.elementAt(i);
                    val = data.data[SXL_ST_PF];
                    y_pos_shld_sxl_st_pf[i] = (int)yPos(d.height,view_rec.height,min,max,val);
                }

                //引き上げ速度プロファイル
                y_pos_sxl_st_pf = new int[size];
                s_min = new Float(sxl_st_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(sxl_st_pf_max_pro);
                max   = s_max.floatValue();
                for(int i = 0 ; i < size ; i++){
                    data = (CZSystemPVData)pv_data_body.elementAt(i);
                    val = data.data[SXL_ST_PF];
                    y_pos_sxl_st_pf[i] = (int)yPos(d.height,view_rec.height,min,max,val);
                }
            }

            // マウス座標のデータ値を設定する。
            //
            private void setMouse(int x,int y){
                Dimension d = getSize(null);
                setMouseX(d,x);
                setMouseY(d,y);
                main_mouse_view.drawVal();
            }

            // マウスのX座標よりX値を計算する。
            //
            private void setMouseX(Dimension d,int x){

                //Ｘ軸座標計算
                float val;
                Float x_max = new Float(gr_x_length);
                float min = 0.0f;
                float max = x_max.floatValue();

                //ＳＸＬ長さ
                val = xPosConv(d.width,view_rec.width,min,max,x);
                x_length_mouse = val;

            }

            // マウスのY座標よりY値を計算する。
            //
            private void setMouseY(Dimension d,int y){

                //Ｙ軸座標計算
                float val;
                Float s_min;
                float min;
                Float s_max;
                float max;

                //メインヒーター１温度
                s_min = new Float(main1_h_t_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(main1_h_t_max_pro);
                max   = s_max.floatValue();
                val = yPosConv(d.height,view_rec.height,min,max,y);
                y_main1_h_t_mouse = val;

                //メインヒーター１温度プロファイル
                s_min = new Float(main1_h_t_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(main1_h_t_pf_max_pro);
                max   = s_max.floatValue();
                val = yPosConv(d.height,view_rec.height,min,max,y);
                y_main1_h_t_pf_mouse = val;

                //直径
                s_min = new Float(dia_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(dia_max_pro);
                max   = s_max.floatValue();
                val = yPosConv(d.height,view_rec.height,min,max,y);
                y_dia_mouse = val;

                //引き上げ速度
                s_min = new Float(sxl_st_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(sxl_st_max_pro);
                max   = s_max.floatValue();
                val = yPosConv(d.height,view_rec.height,min,max,y);
                y_sxl_st_mouse = val;

                //引き上げ速度プロファイル
                s_min = new Float(sxl_st_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(sxl_st_pf_max_pro);
                max   = s_max.floatValue();
                val = yPosConv(d.height,view_rec.height,min,max,y);
                y_sxl_st_pf_mouse = val;
            }

            // マウスのX値よりグラフ上のＸ座標を求める
            //
            private float xPos(int d_width,int v_width,float min,float max,float val){
                float x_dot = (float)v_width / (max - min);
                float x = x_dot * (val - min);
                return x;
            }

            // Ｘ座標より値を求める
            //
            private float xPosConv(int d_width,int v_width,float min,float max,int x){
                float x_dot = (float)v_width / (max - min);
                float val = x / x_dot + min;
                return val;
            }

            // マウスのY値よりグラフ上のＹ座標を求める
            //
            private float yPos(int d_height,int v_height,float min,float max,float val){
                float y_dot = (float)v_height / (max - min);
                float y = (float)d_height - y_dot * (val - min);
                return y;
            }

            // Ｙ座標より値を求める
            //
            private float yPosConv(int d_height,int v_height,float min,float max,int y){
                float y_dot = (float)v_height / (max - min);
                float val = (d_height - y) / y_dot + min;
                return val;
            }

            /***********************************************
             * マウス動作の処理
             *
             ***********************************************/
            class MainViewMouseMotion implements MouseMotionListener {

                //
                public void mouseDragged(MouseEvent e){

                }

                // マウス移動のListener
                //  マウス位置（X,Y座標）よりデータ値を変更する。
                public void mouseMoved(MouseEvent e){
                    int x = e.getX();
                    int y = e.getY();
                    setMouse(x,y);
                }
            } // MouseMotion
        } // View
    } // MainSc

    /*******************************************************
     *
     *       Ｘ軸の目盛表示用パネル
     *
     *******************************************************/
    public class XSc extends JScrollPane {

        private Rectangle       view_rec        = null;
        private View            view            = null;

        // ---------- コンストラクタ -----------------------
        XSc(PVGrEventCompo comp){
            super();

            setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);
            setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);

            view = new View();
            setViewportView(view);
            comp.setXView(view);
            view.addComponentListener(comp);

            getViewport().setScrollMode(JViewport.BACKINGSTORE_SCROLL_MODE);
//@@            CZSystem.log("CZTPGMain XSc","new");
        }

        //
        //
        //
        public void setDefault(){
            view_rec = getViewportBorderBounds();
            view.setPreferredSize(new Dimension(view_rec.width,view_rec.height));
            view.setLocation(0,0);
            view.setViewRec(view_rec);
        }

        // X軸の表示
        //
        public void chgXSize(){
            Dimension d = view.getSize(null);

            Float x_def = new Float(GR_X_LENGTH_DEF);
            Float x_new = new Float(gr_x_length);
            float new_size = (x_def.floatValue()/x_new.floatValue()) * view_rec.width;

            view.setSize(new Dimension((int)new_size,d.height));
            view.repaint();
        }

        /***************************************************
         * X軸の目盛表示パネル
         ***************************************************/
        class View extends JPanel {
            Rectangle view_rec = null;

            View(){
                super();
                setName("View");
                setLayout(null);
                setBackground(BACK_COL);
            }

            // 表示枠を設定する。
            //
            public void setViewRec(Rectangle rec){
                view_rec = rec;
            }

            // X軸目盛を描画する。
            //
            public void paint(Graphics g){
                Dimension d = getSize(null);

                g.setColor(BACK_COL);
                g.fillRect(0,0,d.width,d.height);
                drawMemLine(g);
                drawMem(g);
            }

            // 目盛線を描画する。
            //
            private void drawMemLine(Graphics g){
                float x;
                float inc;

                Dimension d = getSize(null);

                g.setColor(MEM_LINE1_COL);
                inc = view_rec.width / gr_x_bun;
                for(x = 0.0f ;  d.width > x ; x+=inc){
                    g.drawLine((int)x,0,(int)x,d.height);
                }
            }

            // 目盛を描画する。
            //
            private void drawMem(Graphics g){
                Dimension d = getSize(null);
                g.setColor(MEM_LINE1_COL);
                
                float x;
                float inc;

                Float tmp = new Float(gr_x_length); 
                float mem_inc = tmp.floatValue() / gr_x_bun;    
                float x_val   = 0.0f;

                inc = view_rec.width / gr_x_bun;

                for(x = 0.0f ;  d.width > x ; x+=inc){
                    g.drawString(String.valueOf(x_val),(int)x+3,view_rec.height/2);
                    x_val+=mem_inc;
                }
            }
        } // View
    } //XSc

    /*******************************************************
     *
     *       Ｙ軸グラフ左側目盛
     *
     *******************************************************/
    public class Y1Sc extends JScrollPane {

        private Rectangle   view_rec    = null;
        private View        view        = null;

        // ---------- コンストラクタ -----------------------
        Y1Sc(PVGrEventCompo comp){
            super();

            setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);
            setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);

            view = new View();
            setViewportView(view);
            comp.setY1View(view);
            view.addComponentListener(comp);
            getViewport().setScrollMode(JViewport.BACKINGSTORE_SCROLL_MODE);
//@@            CZSystem.log("CZTPGMain Y1Sc","new");
        }

        //
        //
        //
        public void setDefault(){
            view_rec = getViewportBorderBounds();
            view.setPreferredSize(new Dimension(view_rec.width*2,view_rec.height*Y_VIEW_TIMES));
            view.setLocation(0, -(view_rec.height*Y_VIEW_TIMES - view_rec.height));
            view.setViewRec(view_rec);
        }

        // Y軸左側を再描画する。
        //
        public void chgYSize(){
            view.repaint();
        }

        /***************************************************
         * Y軸左側の目盛を表示する
         ***************************************************/
        class View extends JPanel {
            Rectangle view_rec = null;

            // ---------- コンストラクタ -------------------
            View(){
                super();
                setName("View");
                setLayout(null);
                setBackground(BACK_COL);
            }

            // 領域を設定する。
            //
            public void setViewRec(Rectangle rec){
                view_rec = rec;
            }

            // 目盛を描画する。
            //
            public void paint(Graphics g){
                Dimension d = getSize(null);
                g.setColor(BACK_COL);
                g.fillRect(0,0,d.width,d.height);

                g.setColor(MEM_LINE1_COL);
                drawMemLine(g);
                drawMemString(g);
            }

            // 目盛を描画する
            //
            private void drawMemString(Graphics g){ 
                Dimension d = getSize(null);

                //メインヒーター１温度
                g.setColor(MAIN1_H_T_COL);
                drawMem(g,d,45,main1_h_t_min_pro,main1_h_t_max_pro);

                //メインヒーター１温度プロファイル
                g.setColor(MAIN1_H_T_PF_COL);
                drawMem(g,d,35,main1_h_t_pf_min_pro,main1_h_t_pf_max_pro);

                //直径
                g.setColor(DIA_COL);
                drawMem(g,d,25,dia_min_pro,dia_max_pro);

                //引き上げ速度
                g.setColor(SXL_ST_COL);
                drawMem(g,d,15,sxl_st_min_pro,sxl_st_max_pro);

                //引き上げ速度プロファイル
                g.setColor(SXL_ST_PF_COL);
                drawMem(g,d,5,sxl_st_pf_min_pro,sxl_st_pf_max_pro);
            }

            // 目盛線を描画する
            //
            private void drawMemLine(Graphics g){
                float y;
                float inc;

                Dimension d = getSize(null);
                inc = view_rec.height / gr_y_bun;
                for(y = (float)d.height ;  0 < y ; y-=inc){
                    g.drawLine(0,(int)y,d.width,(int)y);
                }
            }

            // 目盛を描画する
            //
            private void drawMem(Graphics g,Dimension d,int sa,String st_min,String st_max){
                float x = 3.0f;
                float y = 0.0f;

                Float s_min = new Float(st_min);
                float min   = s_min.floatValue();

                Float s_max = new Float(st_max);
                float max   = s_max.floatValue();
                
                float inc   = (max - min) / gr_y_bun;
                float y_dot = view_rec.height / (max - min);

                for(float tmp = min ; 0 <= y ; tmp += inc){
                    y = (float)d.height - y_dot * (tmp - min);
                    g.drawString(String.valueOf(tmp),(int)x,(int)y-sa);
                }
            }
        } // View
    } // Y1Sc

    /*******************************************************
     *
     *       Ｙ軸グラフ右側パネル
     *
     *******************************************************/
    public class Y2Sc extends JScrollPane {

        private Rectangle   view_rec    = null;
        private View        view        = null;

        Y2Sc(PVGrEventCompo comp){
            super();
            setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
            setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
            view = new View();
            setViewportView(view);
            comp.setY2View(view);
            view.addComponentListener(comp);
            getViewport().setScrollMode(JViewport.BACKINGSTORE_SCROLL_MODE);
//@@            CZSystem.log("CZTPGMain Y2Sc","new");
        }

        //
        //
        public void setDefault(){
            view_rec = getViewportBorderBounds();
            view.setPreferredSize(new Dimension(view_rec.width*2,view_rec.height*Y_VIEW_TIMES));
            view.setLocation(0, -(view_rec.height*Y_VIEW_TIMES - view_rec.height));
            view.setViewRec(view_rec);
        }

        // 再描画する。
        //
        public void chgYSize(){
            view.repaint();
        }

        /***************************************************
         * Y軸右側を表示する
         ***************************************************/
        class View extends JPanel {
            Rectangle view_rec = null;

            // ---------- コンストラクタ -------------------
            View(){
                super();
                setName("View");
                setLayout(null);
                setBackground(BACK_COL);
            }

            // 領域を設定する。
            //
            public void setViewRec(Rectangle rec){
                view_rec = rec;
            }

            // 目盛線を描画する。
            //
            public void paint(Graphics g){
                Dimension d = getSize(null);
                g.setColor(BACK_COL);
                g.fillRect(0,0,d.width,d.height);
                drawMemLine(g);
            }

            // 目盛線を描画する。
            //
            private void drawMemLine(Graphics g){
                float y;
                float inc;

                g.setColor(MEM_LINE1_COL);
                Dimension d = getSize(null);
                inc = view_rec.height / gr_y_bun;
                for(y = (float)d.height ;  0 < y ; y-=inc){
                    g.drawLine(0,(int)y,d.width,(int)y);
                }
            }
        } // View
    } // Y2Sc

    /*******************************************************
     *
     * グラフ表示領域のListener
     *
     *******************************************************/
    class PVGrEventCompo implements ComponentListener {

        private JPanel main_view    = null;     // グラフ表示パネル
        private JPanel x_view       = null;     // X軸目盛パネル
        private JPanel y1_view      = null;     // Y軸左目盛パネル
        private JPanel y2_view      = null;     // Y軸右パネル

        // ---------- コンストラクタ -----------------------
        PVGrEventCompo(){

        }

        //
        // グラフ表示パネルを保持する
        public void setMainView(JPanel view){
            main_view = view;
        }

        //
        // X軸目盛表示パネルを保持する
        public void setXView(JPanel view){
            x_view = view;
        }

        //
        // Y軸左側目盛表示パネルを保持する
        public void setY1View(JPanel view){
            y1_view = view;
        }

        //
        // Y軸右側表示パネルを保持する
        public void setY2View(JPanel view){
            y2_view = view;
        }

        //
        // 移動時の処理
        public void componentMoved(java.awt.event.ComponentEvent e){

            if(y2_view == e.getComponent()){
                y1_view.setLocation(y1_view.getX(),y2_view.getY());
                main_view.setLocation(main_view.getX(),y2_view.getY());
            }

            if(x_view == e.getComponent()){
                main_view.setLocation(x_view.getX(),main_view.getY());
            }
        }

        public void componentResized(java.awt.event.ComponentEvent e){
        }

        public void componentShown(java.awt.event.ComponentEvent e){
        }

        public void componentHidden(java.awt.event.ComponentEvent e){
        }
    } // PVGrEventCompo
}
