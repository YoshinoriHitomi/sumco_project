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

//==========================================================================
/**
*   グラフ表示用ダイアログ
*   上下限交点の長さ表示パネル追加  @@@ 2003/06/06
*/
public class CZFpAveGraphFrame extends JFrame
{

    private String ro_name              = null;     //対象炉番

    private Vector ro_bt_all_condition  = null;     //全Btの引き上げ条件

    private GraphSet    graph_set       = null;     //グラフ描画条件

    private Vector pv_data_body         = null;     //ボディーのデータ
    private Vector calc_data_body       = null;     //ボディーの計算データ

    private int     fp_ave_calc_time    = 10;       //移動平均時間(計算に使用)


    private final   String  TITLE       = "FpAveグラフ";

    private JLabel  ro_name_gr_lab      = null;

    private JLabel  bt_no_lab           = null; 
    private JLabel  bt_hinban_lab       = null;
    private JLabel  bt_fp_ave_time_lab  = null;

    private JLabel  bt_sxl_length_lab   = null;
    private JLabel  bt_sxl_dia_lab      = null;
    private JLabel  bt_sxl_chg_lab      = null;

    private JLabel  bt_cond_t1_lab      = null;
    private JLabel  bt_cond_t2_lab      = null;
    private JLabel  bt_cond_t3_lab      = null;
    private JLabel  bt_cond_t4_lab      = null;
    private JLabel  bt_cond_t5_lab      = null;

    private Y1View      y1_view         = null;
    private Y2View      y2_view         = null;
    private XView       x_view          = null;
    private MainView    main_view       = null;
    private LimitPanel  limit_view      = null;

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

    /**
     * コンストラクタ
     */
    public CZFpAveGraphFrame(String roName, int fp_ave_time, Vector v, Vector body_data, Vector body_data_calc, GraphSet gs){
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

        setTitle(TITLE);

        /* @@@
//          グラフ画面の大きさを調整する。setSize( Width, Height )
        */
//            setSize(1432,864);
//@@@@@            setSize(1152,864);
        setSize(1280,864);
        setResizable(false);
//            setModal(true);
        getContentPane().setLayout(null);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        addWindowListener(
            new WindowAdapter(){
                public void windowClosing(WindowEvent e){
                    CZSystem.GraphCountDown();
                }
        });

        ro_name_gr_lab = new JLabel("",JLabel.CENTER);
        ro_name_gr_lab.setBounds(20, 10, 100, 30);
        ro_name_gr_lab.setLocale(new Locale("ja","JP"));
        ro_name_gr_lab.setFont(new java.awt.Font("dialog", 0, 18));
        ro_name_gr_lab.setBorder(new Flush3DBorder());
        ro_name_gr_lab.setForeground(BACK_COL);
        getContentPane().add(ro_name_gr_lab);

        JLabel l;
        l = new JLabel("Bt",JLabel.CENTER);
        l.setBounds(140, 16, 20, 24);
        l.setLocale(new Locale("ja","JP"));
        l.setFont(new java.awt.Font("dialog", 0, 16));
        l.setBorder(new Flush3DBorder());
        l.setForeground(BACK_COL);
        getContentPane().add(l);

        bt_no_lab = new JLabel("Bt",JLabel.CENTER);
        bt_no_lab.setBounds(160, 16, 100, 24);
        bt_no_lab.setLocale(new Locale("ja","JP"));
        bt_no_lab.setFont(new java.awt.Font("dialog", 0, 10));
        bt_no_lab.setBorder(new Flush3DBorder());
        bt_no_lab.setForeground(BACK_COL);
        getContentPane().add(bt_no_lab);

        l = new JLabel("品番",JLabel.CENTER);
        l.setBounds(270, 16, 40, 24);
        l.setLocale(new Locale("ja","JP"));
        l.setFont(new java.awt.Font("dialog", 0, 16));
        l.setBorder(new Flush3DBorder());
        l.setForeground(BACK_COL);
        getContentPane().add(l);

        bt_hinban_lab = new JLabel("品番",JLabel.CENTER);
        bt_hinban_lab.setBounds(310, 16, 100, 24);
        bt_hinban_lab.setLocale(new Locale("ja","JP"));
        bt_hinban_lab.setFont(new java.awt.Font("dialog", 0, 12));
        bt_hinban_lab.setBorder(new Flush3DBorder());
        bt_hinban_lab.setForeground(BACK_COL);
        getContentPane().add(bt_hinban_lab);

        l = new JLabel("平均時間(s)",JLabel.CENTER);
        l.setBounds(430, 16, 100, 24);
        l.setLocale(new Locale("ja","JP"));
        l.setFont(new java.awt.Font("dialog", 0, 16));
        l.setBorder(new Flush3DBorder());
        l.setForeground(BACK_COL);
        getContentPane().add(l);

        bt_fp_ave_time_lab = new JLabel("平均時間(s)",JLabel.CENTER);
        bt_fp_ave_time_lab.setBounds(530, 16, 60, 24);
        bt_fp_ave_time_lab.setLocale(new Locale("ja","JP"));
        bt_fp_ave_time_lab.setFont(new java.awt.Font("dialog", 0, 12));
        bt_fp_ave_time_lab.setBorder(new Flush3DBorder());
        bt_fp_ave_time_lab.setForeground(BACK_COL);
        getContentPane().add(bt_fp_ave_time_lab);

        l = new JLabel("直径",JLabel.CENTER);
        l.setBounds(630, 20, 40, 20);
        l.setLocale(new Locale("ja","JP"));
        l.setFont(new java.awt.Font("dialog", 0, 14));
        l.setBorder(new Flush3DBorder());
        l.setForeground(BACK_COL);
        getContentPane().add(l);

        bt_sxl_dia_lab = new JLabel("直径",JLabel.CENTER);
        bt_sxl_dia_lab.setBounds(670, 20, 50, 20);
        bt_sxl_dia_lab.setLocale(new Locale("ja","JP"));
        bt_sxl_dia_lab.setFont(new java.awt.Font("dialog", 0, 10));
        bt_sxl_dia_lab.setBorder(new Flush3DBorder());
        bt_sxl_dia_lab.setForeground(BACK_COL);
        getContentPane().add(bt_sxl_dia_lab);

        l = new JLabel("引上長",JLabel.CENTER);
        l.setBounds(730, 20, 60, 20);
        l.setLocale(new Locale("ja","JP"));
        l.setFont(new java.awt.Font("dialog", 0, 14));
        l.setBorder(new Flush3DBorder());
        l.setForeground(BACK_COL);
        getContentPane().add(l);

        bt_sxl_length_lab = new JLabel("引上長",JLabel.CENTER);
        bt_sxl_length_lab.setBounds(790, 20, 50, 20);
        bt_sxl_length_lab.setLocale(new Locale("ja","JP"));
        bt_sxl_length_lab.setFont(new java.awt.Font("dialog", 0, 10));
        bt_sxl_length_lab.setBorder(new Flush3DBorder());
        bt_sxl_length_lab.setForeground(BACK_COL);
        getContentPane().add(bt_sxl_length_lab);

        l = new JLabel("仕込",JLabel.CENTER);
        l.setBounds(850, 20, 40, 20);
        l.setLocale(new Locale("ja","JP"));
        l.setFont(new java.awt.Font("dialog", 0, 14));
        l.setBorder(new Flush3DBorder());
        l.setForeground(BACK_COL);
        getContentPane().add(l);

        bt_sxl_chg_lab = new JLabel("仕込",JLabel.CENTER);
        bt_sxl_chg_lab.setBounds(890, 20, 70, 20);
        bt_sxl_chg_lab.setLocale(new Locale("ja","JP"));
        bt_sxl_chg_lab.setFont(new java.awt.Font("dialog", 0, 10));
        bt_sxl_chg_lab.setBorder(new Flush3DBorder());
        bt_sxl_chg_lab.setForeground(BACK_COL);
        getContentPane().add(bt_sxl_chg_lab);

        JScrollPane p ;

        main_view = new MainView();
        p = new JScrollPane(main_view);
//@@            p.setBounds(70, 50, 970, 790);
        p.setBounds(70, 50, 970, 730);
        p.setBorder(new Flush3DBorder());
        p.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);
        p.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
        p.getViewport().setScrollMode(JViewport.BACKINGSTORE_SCROLL_MODE);
        getContentPane().add(p);

        y1_view = new Y1View();
        p = new JScrollPane(y1_view);
//@@            p.setBounds(20, 50, 50, 790);
        p.setBounds(20, 50, 50, 730);
        p.setBorder(new Flush3DBorder());
        p.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);
        p.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
        p.getViewport().setScrollMode(JViewport.BACKINGSTORE_SCROLL_MODE);
        getContentPane().add(p);

        y2_view = new Y2View();
        p = new JScrollPane(y2_view);
//@@            p.setBounds(1040, 50, 50, 790);
        p.setBounds(1040, 50, 50, 730);
        p.setBorder(new Flush3DBorder());
        p.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);
        p.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
        p.getViewport().setScrollMode(JViewport.BACKINGSTORE_SCROLL_MODE);
        getContentPane().add(p);

        x_view = new XView();
        p = new JScrollPane(x_view);
//@@            p.setBounds(70, 840, 970, 40);
        p.setBounds(70, 780, 970, 40);
        p.setBorder(new Flush3DBorder());
        p.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);
        p.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
        p.getViewport().setScrollMode(JViewport.BACKINGSTORE_SCROLL_MODE);
        getContentPane().add(p);

//@@@

        limit_view = new LimitPanel();
        p = new JScrollPane(limit_view);
        //交点の表示窓を調整する。p.setBounds(X座標, Y座標, Width, Height)
        p.setBounds(1091, 50, 180, 644);
        p.setBorder(new Flush3DBorder());
        p.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);
        p.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
        p.getViewport().setScrollMode(JViewport.BACKINGSTORE_SCROLL_MODE);
        p.setForeground(BACK_COL);   //@@@
        getContentPane().add(p);

        fp_ave_calc_time = fp_ave_time;
        ro_name = roName;
        
        ro_bt_all_condition = v;
        pv_data_body = body_data;
        calc_data_body = body_data_calc;
        graph_set = gs;
//@@@
        CZSystem.log("CZFpAveMain","GraphDialog new");
    }

    /**
    * 画面の初期設定
    */
    public boolean setDefault(){
        ro_name_gr_lab.setText(ro_name);
        CZSystemBt bt = (CZSystemBt)ro_bt_all_condition.elementAt(0);
        bt_no_lab.setText(bt.batch.trim());
        bt_hinban_lab.setText(bt.hinshu.trim());
        bt_fp_ave_time_lab.setText(fp_ave_calc_time + "");
        bt_sxl_length_lab.setText(bt.hikiage_cho + "");
        bt_sxl_dia_lab.setText(bt.chokkei + "");
        bt_sxl_chg_lab.setText( (bt.i_sikomi + bt.t_sikomi) + "");

//@@@
        limit_view.searchPoint();       //交点を探し出す。
        limit_view.setData();           //交点を画面に設定する。
//@@@
        return true;
    }

    /**
    * 数値を指定の桁にFormatするための形式を作成する。
    */
    public String getDecFormat(int souketa,int syouketa){
        if(0 > souketa) return "0";
        StringBuffer ret = new StringBuffer(souketa);
        for (int i = 0 ; i < souketa ; i++) ret.append("0");
        if(0 == syouketa) return ret.toString();
        int p = souketa - syouketa - 1;
        if(0 > p) return ret.toString();
        ret.setCharAt(p,'.');
        return ret.toString();
    }

    //======================================================================
    /**
    *   左側Ｙ軸表示パネル（引き上げ速度の目盛）
    */
    class Y1View extends JPanel {

        /**
        * コンストラクタ
        */
        Y1View(){
            super();
            setName("Y1View");
            setBackground(BACK_COL);
        }

        /**
        * Y軸左側を描画する
        */
        public void paint(Graphics g){
            Dimension d = getSize(null);
            g.setColor(BACK_COL);
            g.fillRect(0,0,d.width,d.height);

            if(null == graph_set) return;
            if(null == calc_data_body) return;
            if(null == pv_data_body) return;

            GraphSet p = graph_set;
            DecimalFormat format = new DecimalFormat(getDecFormat(p.y_memketa,p.y_syouketa));
                
            float y = d.height;
            float inc = (float)d.height / (float)p.y_bun;

            float y_val_inc = (p.y_max - p.y_min) / (float)p.y_bun;
            float y_val = p.y_min;

            for (int i = 0 ; 0 <= y ; i++) {

                g.setColor(MEM_LINE2_COL);
                if(0 < p.y_koushi){
                    float koushi_inc = inc / (float)p.y_koushi;
                    float yy = y;
                    int   xx = d.width - (d.width /10);
                    for (int j = 0 ; j <= p.y_koushi ; j++) {
                        g.drawLine(xx,(int)yy,d.width,(int)yy);
                        yy -= koushi_inc;
                    }
                }

                g.setColor(MEM_LINE1_COL);
                g.drawLine(0,(int)y-1,d.width,(int)y-1);
                g.drawLine(0,(int)y,d.width,(int)y);
                g.drawLine(0,(int)y+1,d.width,(int)y+1);

                if(0 < p.y_memkan){
                    if(0 == ( i % p.y_memkan )){
                        g.drawString(format.format(y_val), 5,(int)(y - 5.0f));
                    }
                }
                y_val +=y_val_inc;
                y -= inc;
            } //for end
        }
    } //Y1View  

    //======================================================================
    /**
    *   右側Ｙ軸表示パネル（シード、ルツボ回転の目盛）
    */
    class Y2View extends JPanel {

        /**
        * コンストラクタ
        */
        Y2View(){
            super();
            setName("Y2View");
            setBorder(new Flush3DBorder());
            setBackground(BACK_COL);
        }

        /**
        * Y軸右側を描画する
        */
        public void paint(Graphics g){
            Dimension d = getSize(null);
            g.setColor(BACK_COL);
            g.fillRect(0,0,d.width,d.height);

            if(null == graph_set) return;
            if(null == calc_data_body) return;
            if(null == pv_data_body) return;

            GraphSet p = graph_set;
            DecimalFormat dia_format = new DecimalFormat("000.0");
            DecimalFormat rpm_format = new DecimalFormat("00.0");
                
            float y = d.height;
            float inc = (float)d.height / (float)p.y_bun;

            float y_dia_val = p.y_dia_min;
            float y_rpm_val = p.y_rpm_min;

            float y_dia_val_inc = (p.y_dia_max - p.y_dia_min) / (float)p.y_bun;
            float y_rpm_val_inc = (p.y_rpm_max - p.y_rpm_min) / (float)p.y_bun;
            int   xx = d.width / 10;

            for (int i = 0 ; 0 <= y ; i++) {

                g.setColor(MEM_LINE2_COL);
                if(0 < p.y_koushi){
                    float koushi_inc = inc / (float)p.y_koushi;
                    float yy = y;
                    for (int j = 0 ; j <= p.y_koushi ; j++) {
                        g.drawLine(0,(int)yy,xx,(int)yy);
                        yy -= koushi_inc;
                    }
                }

                g.setColor(MEM_LINE1_COL);
                g.drawLine(0,(int)y-1,d.width,(int)y-1);
                g.drawLine(0,(int)y,d.width,(int)y);
                g.drawLine(0,(int)y+1,d.width,(int)y+1);

                if(0 < p.y_memkan){
                    if(0 == ( i % p.y_memkan )){
                        g.drawString(dia_format.format(y_dia_val), xx + 2,(int)(y - 15.0f));
                        g.drawString(rpm_format.format(y_rpm_val), xx + 2,(int)(y - 5.0f));
                    }
                }

                y_dia_val += y_dia_val_inc;
                y_rpm_val += y_rpm_val_inc;
                y -= inc;
            } //for end
        }
    } //Y2View  

    //======================================================================
    /**
    *   Ｘ軸表示パネル（ＳＸＬ長さの目盛）
    */
    class XView extends JPanel {

        /**
        * コンストラクタ
        */
        XView(){
            super();
            setName("XView");
            setBorder(new Flush3DBorder());
            setBackground(BACK_COL);
        }

        /**
        * X軸目盛を描画する
        */
        public void paint(Graphics g){
            Dimension d = getSize(null);
            g.setColor(BACK_COL);
            g.setFont(new java.awt.Font("dialog", 0, 10));
            g.fillRect(0,0,d.width,d.height);

            if(null == graph_set) return;
            if(null == calc_data_body) return;
            if(null == pv_data_body) return;

            GraphSet p = graph_set;
            DecimalFormat format = new DecimalFormat(getDecFormat(p.x_memketa,p.x_syouketa));

            float x = 0.0f;
            float inc = (float)d.width / (float)p.x_bun;

            float x_val_inc = (p.x_max - p.x_min) / (float)p.x_bun;
            float x_val = p.x_min;

            float x_shift = 0.0f;
            if(p.shld_shift){
                x_shift = ((float)d.width / (p.x_max - p.x_min)) * p.shld_shift_val ;
            }

            x += x_shift;
            //目盛線と目盛を描画する。
            for (int i = 0 ; d.width >= x ; i++){ 
                g.setColor(MEM_LINE2_COL);
                if(0 < p.x_koushi){
                    float koushi_inc = inc / (float)p.x_koushi;
                    float xx = x;   
                    int yy = d.height / 10;
                    for (int j = 0 ; j <= p.x_koushi ; j++) {
                        g.drawLine((int)xx,0,(int)xx,yy);
                        xx += koushi_inc;
                    }
                }

                g.setColor(MEM_LINE1_COL);
                g.drawLine((int)x-1,0,(int)x-1,d.height);
                g.drawLine((int)x,0,(int)x,d.height);
                g.drawLine((int)x+1,0,(int)x+1,d.height);
                if(0 < p.x_memkan){
                    if(0 == ( i % p.x_memkan )){
                        g.drawString(format.format(x_val), (int)(x + 5.0f) , d.height - 5);
                    }
                }
                x_val += x_val_inc;
                x += inc;
            } //for end
        }
    } //XView   

    //======================================================================
    /**
    *   グラフ表示パネル
    */
    class MainView extends JPanel {

        /**
        * コンストラクタ
        */
        MainView(){
            super();
            setName("MainView");
            setBorder(new Flush3DBorder());
            setBackground(BACK_COL);
        }

        /**
        * グラフを描画する
        */
        public void paint(Graphics g){
            Dimension d = getSize(null);
            g.setColor(BACK_COL);
            g.fillRect(0,0,d.width,d.height);

            drawYMemK(g);
            drawXMemK(g);
            drawYMem(g);
            drawXMem(g);

            if(null == graph_set) return;
            if(null == calc_data_body) return;
            if(null == pv_data_body) return;

            drawFp(g);
            drawDia(g);
            drawRPM(g);

            drawFpAve(g);
        }

        /**
        *  回転系グラフを描画する
        */
        private void drawRPM(Graphics g){
            Dimension d = getSize(null);
            GraphSet p = graph_set;

            int h = d.height;
            int w = d.width;

            int size = pv_data_body.size();
            if(2 > size) return;

            int jg[]    = new int[size];
            int x[]     = new int[size];
            int y1[]    = new int[size];
            int y2[]    = new int[size];
            float x_min = p.x_min;
            float x_max = p.x_max;
            float y_min = p.y_rpm_min;
            float y_max = p.y_rpm_max;

            CZSystemPVData  v;

            for (int i = 0 ; i < size ; i++) {
                v = (CZSystemPVData)pv_data_body.elementAt(i);
                x[i]    = (int)xPos(w,x_min,x_max,v.p_length);
                y1[i]   = (int)yPos(h,y_min,y_max,v.data[SXL_RT]);
                y2[i]   = (int)yPos(h,y_min,y_max,v.data[CRU_RT]);
            }

            if(p.sxl_rpm_draw){
                g.setColor(p.sxl_rpm_draw_col);
                g.drawPolyline(x,y1,size);
            }

            if(p.cru_rpm_draw){
                g.setColor(p.cru_rpm_draw_col);
                g.drawPolyline(x,y2,size);
            }
        }

        /**
        *  直径グラフを描画する
        */
        private void drawDia(Graphics g){
            Dimension d = getSize(null);
            GraphSet p = graph_set;

            int h = d.height;
            int w = d.width;

            int size = pv_data_body.size();
            if(2 > size) return;

            int jg[]    = new int[size];
            int x[]     = new int[size];
            int y1[]    = new int[size];
            int y2[]    = new int[size];
            int y3[]    = new int[size];
            int y4[]    = new int[size];

            float x_min = p.x_min;
            float x_max = p.x_max;
            float y_min = p.y_dia_min;
            float y_max = p.y_dia_max;

            float pf_min = 0.0f;
            float pf_max = 0.0f;

            try{
                pf_min = Float.parseFloat(dia_pf_min_pro);
                pf_max = Float.parseFloat(dia_pf_max_pro);
            }
            catch(NumberFormatException e){
                pf_min = 1.0f;
                pf_max = 1.0f;
            }

            CZSystemPVData  v;
            for (int i = 0 ; i < size ; i++) {
                v = (CZSystemPVData)pv_data_body.elementAt(i);
                x[i]    = (int)xPos(w,x_min,x_max,v.p_length);
                y1[i]   = (int)yPos(h,y_min,y_max,v.data[DIA_PF]+pf_min);
                y2[i]   = (int)yPos(h,y_min,y_max,v.data[DIA_PF]+pf_max);
                y3[i]   = (int)yPos(h,y_min,y_max,v.data[DIA_PF]);
                y4[i]   = (int)yPos(h,y_min,y_max,v.data[DIA]);
            }

            if(p.dia_pf_draw){
                g.setColor(p.dia_pf_draw_col);
                g.drawPolyline(x,y1,size);
                g.drawPolyline(x,y2,size);
                g.drawPolyline(x,y3,size);
            }

            if(p.dia_draw){
                g.setColor(p.dia_draw_col);
                g.drawPolyline(x,y4,size);
            }
        }

        /**
        *  引き上げ速度（生データ）グラフを描画する
        */
        private void drawFp(Graphics g){
            Dimension d = getSize(null);
            GraphSet p = graph_set;

            int h = d.height;
            int w = d.width;

            int size = pv_data_body.size();
            if(2 > size) return;

            int jg[]    = new int[size];
            int x[]     = new int[size];
            int y1[]    = new int[size];
            int y2[]    = new int[size];

            float x_min = p.x_min;
            float x_max = p.x_max;
            float y_min = p.y_min;
            float y_max = p.y_max;

            CZSystemPVData  v;
            for (int i = 0 ; i < size ; i++) {
                v = (CZSystemPVData)pv_data_body.elementAt(i);
                x[i]    = (int)xPos(w,x_min,x_max,v.p_length);
                y1[i]   = (int)yPos(h,y_min,y_max,v.data[SXL_ST]);
                y2[i]   = (int)yPos(h,y_min,y_max,v.data[SXL_ST_PF]);
            }

            if(p.fp_draw){
                g.setColor(p.fp_draw_col);
                g.drawPolyline(x,y1,size);
            }

            if(p.fp_pf_draw){
                g.setColor(p.fp_pf_draw_col);
                g.drawPolyline(x,y2,size);
            }
        }

        /**
        *  引き上げ速度（移動平均データ）グラフを描画する。
        */
        private void drawFpAve(Graphics g){
            Dimension d = getSize(null);
            GraphSet p = graph_set;

            int h = d.height;
            int w = d.width;

            int size = pv_data_body.size();
            if(2 > size) return;

            int jg[]    = new int[size];
            int x[]     = new int[size];
            int y1[]    = new int[size];
            int y2[]    = new int[size];
            int y3[]    = new int[size];
            int y4[]    = new int[size];
            int y5[]    = new int[size];
            int y6[]    = new int[size];

            float x_min = p.x_min;
            float x_max = p.x_max;
            float y_min = p.y_min;
            float y_max = p.y_max;

            CZSystemPVData  v;
            CalcData    c;
            for (int i = 0 ; i < size ; i++) {
                v = (CZSystemPVData)pv_data_body.elementAt(i);
                x[i] = (int)xPos(w,x_min,x_max,v.p_length);

                c = (CalcData)calc_data_body.elementAt(i);
                y1[i] = (int)yPos(h,y_min,y_max,c.fp_ave);
                y2[i] = (int)yPos(h,y_min,y_max,c.pf_umax_ave);
                y3[i] = (int)yPos(h,y_min,y_max,c.pf_max_ave);
                y4[i] = (int)yPos(h,y_min,y_max,c.pf_min_ave);
                y5[i] = (int)yPos(h,y_min,y_max,c.pf_lmin_ave);
                y6[i] = (int)yPos(h,y_min,y_max,c.pf_ave);
                jg[i] = c.judg;
            }
            g.setColor(p.fp_umax_col);
            g.drawPolyline(x,y2,size);

            g.setColor(p.fp_max_col);
            g.drawPolyline(x,y3,size);

            g.setColor(p.fp_min_col);
            g.drawPolyline(x,y4,size);

            g.setColor(p.fp_lmin_col);
            g.drawPolyline(x,y5,size);

            if(p.fp_pf_ave_draw){
                g.setColor(p.fp_pf_ave_draw_col);
                g.drawPolyline(x,y6,size);
            }

            size--;
            for (int i = 0 ; i < size ; i++) {
                switch(jg[i]){
                    case  0 : g.setColor(p.fp_center_col);
                         break;
                    case -1 : g.setColor(p.fp_min_over_col);
                         break;
                    case -2 : g.setColor(p.fp_lmin_over_col);
                         break;
                    case  1 : g.setColor(p.fp_max_over_col);
                         break;
                    case  2 : g.setColor(p.fp_umax_over_col);
                         break;
                    default : g.setColor(java.awt.Color.red);
                         break;
                }
                g.drawLine(x[i],y1[i],x[i+1],y1[i+1]);
            } //for end
        }

        /**
        *  Ｙデータから描画位置を求める
        */
        private float yPos(int height,float min,float max,float val){
            float y_dot = (float)height / (max - min);
            float y = (float)height - y_dot * (val - min);
            return y;
        }

        /**
        *  Ｘデータから描画位置を求める
        */
        private float xPos(int width,float min,float max,float val){
            float x_dot = (float)width / (max - min);
            float x     = x_dot * (val - min);
            return x;
        }

        /**
        *  Ｙ軸目盛の描画
        */
        private void drawYMem(Graphics g){
            Dimension d = getSize(null);
            GraphSet p  = graph_set;

            //Ｙ軸目盛  
            float y     = d.height;
            float inc   = (float)d.height / (float)p.y_bun;

            float y_val_inc = (p.y_max - p.y_min) / (float)p.y_bun;
            float y_val     = p.y_min;

            for (int i = 0 ; 0 <= y ; i++) {
                g.setColor(MEM_LINE1_COL);
                g.drawLine(0,(int)y,d.width,(int)y);
                y_val += y_val_inc;
                y -= inc;
            } //for end
        }

        /**
        *  Ｙ軸目盛の描画（格子）
        */
        private void drawYMemK(Graphics g){
            Dimension d = getSize(null);
            GraphSet p  = graph_set;

            //Ｙ軸目盛  
            float y     = d.height;
            float inc   = (float)d.height / (float)p.y_bun;

            float y_val_inc = (p.y_max - p.y_min) / (float)p.y_bun;
            float y_val     = p.y_min;

            for (int i = 0 ; 0 <= y ; i++) {
                g.setColor(MEM_LINE3_COL);
                if(0 < p.y_koushi){
                    float koushi_inc = inc / (float)p.y_koushi;
                    float yy = y;
                    for (int j = 0 ; j <= p.y_koushi ; j++) {
                        g.drawLine(0,(int)yy,d.width,(int)yy);
                        yy -= koushi_inc;
                    }
                }
                y_val += y_val_inc;
                y -= inc;
            } //for end
        }

        /**
        *  Ｘ軸目盛の描画
        */
        private void drawXMem(Graphics g){
            Dimension d = getSize(null);
            GraphSet p  = graph_set;

            float x     = 0.0f;
            float inc   = (float)d.width / (float)p.x_bun;

            float x_val_inc = (p.x_max - p.x_min) / (float)p.x_bun;
            float x_val     = p.x_min;

            float x_shift = 0.0f;
            if(p.shld_shift){
                x_shift = ((float)d.width / (p.x_max - p.x_min)) * p.shld_shift_val ;
            }

            x += x_shift;
            for (int i = 0 ; d.width >= x ; i++) {
                g.setColor(MEM_LINE1_COL);
                g.drawLine((int)x,0,(int)x,d.height);
                x_val += x_val_inc;
                x += inc;
            } //for end
        }

        /**
        *  Ｘ軸目盛の描画（格子）
        */
        private void drawXMemK(Graphics g){
            Dimension d = getSize(null);
            GraphSet p  = graph_set;

            float x     = 0.0f;
            float inc   = (float)d.width / (float)p.x_bun;

            float x_val_inc = (p.x_max - p.x_min) / (float)p.x_bun;
            float x_val     = p.x_min;

            float x_shift = 0.0f;
            if(p.shld_shift){
                x_shift = ((float)d.width / (p.x_max - p.x_min)) * p.shld_shift_val ;
            }

            x += x_shift;

            for (int i = 0 ; d.width >= x ; i++) {
                g.setColor(MEM_LINE3_COL);
                if(0 < p.x_koushi){
                    float koushi_inc = inc / (float)p.x_koushi;
                    float xx = x;   
                    for (int j = 0 ; j <= p.x_koushi ; j++) {
                        g.drawLine((int)xx,0,(int)xx,d.height);
                        xx += koushi_inc;
                    }
                }
                x_val += x_val_inc;
                x += inc;
            } //for end
        }
    } //MainView

//@@@
    //======================================================================
    /**
    *   上下限の交点の引上げ長表示パネル
    */
    class LimitPanel extends JPanel {

        private JLabel indexLbl[];
        private JLabel startLbl[];
        private JLabel endLbl[];

        private float cutStart[];
        private float cutEnd[];
        private int beforeJudge;
        private int lCount;

        /**
        * コンストラクタ
        */
        LimitPanel(){
            super();
            getContentPane().setName("LimitPanel");
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            // 表示位置を調整するには。
            // setBounds(X, Y, Width, Height)にて行います。

            JLabel l;
            l = new JLabel("№",JLabel.CENTER);
            l.setBounds(1097, 60, 25, 24);
//                l.setBounds(10, 10, 50, 24);
            l.setLocale(new Locale("ja","JP"));
            l.setFont(new java.awt.Font("dialog", 0, 12));
            l.setBorder(new Flush3DBorder());
            l.setForeground(java.awt.Color.black);
            getContentPane().add(l);
//                l.setForeground(BACK_COL);

            l = new JLabel("開始",JLabel.CENTER);
            l.setBounds(1124, 60, 70, 24);
//                l.setBounds(60, 10, 100, 24);
            l.setLocale(new Locale("ja","JP"));
            l.setFont(new java.awt.Font("dialog", 0, 12));
            l.setBorder(new Flush3DBorder());
            l.setForeground(java.awt.Color.black);
            getContentPane().add(l);
//                l.setForeground(BACK_COL);

            l = new JLabel("終了",JLabel.CENTER);
            l.setBounds(1194, 60, 70, 24);
//                l.setBounds(160, 10, 100, 24);
            l.setLocale(new Locale("ja","JP"));
            l.setFont(new java.awt.Font("dialog", 0, 12));
            l.setBorder(new Flush3DBorder());
            l.setForeground(java.awt.Color.black);
            getContentPane().add(l);
//                l.setForeground(BACK_COL);

            indexLbl = new JLabel[25];
            startLbl = new JLabel[25];
            endLbl   = new JLabel[25];
            int iPos = 60;
            for (int i=0; i<25; i++) {
                iPos = iPos + 24;
                indexLbl[i] = new JLabel("" + (i+1),JLabel.CENTER);
                indexLbl[i].setBounds(1097, iPos, 25, 24);
//                    indexLbl[i].setBounds(10, iPos, 50, 24);
                indexLbl[i].setLocale(new Locale("ja","JP"));
                indexLbl[i].setFont(new java.awt.Font("dialog", 0, 12));
                indexLbl[i].setBorder(new Flush3DBorder());
                indexLbl[i].setForeground(java.awt.Color.black);
                getContentPane().add(indexLbl[i]);
//                    indexLbl[i].setForeground(BACK_COL);

                startLbl[i] = new JLabel("",JLabel.CENTER);
                startLbl[i].setBounds(1124, iPos, 70, 24);
//                    startLbl[i].setBounds(60, iPos, 100, 24);
                startLbl[i].setLocale(new Locale("ja","JP"));
                startLbl[i].setFont(new java.awt.Font("dialog", 0, 11));
                startLbl[i].setBorder(new Flush3DBorder());
                startLbl[i].setForeground(java.awt.Color.black);
                getContentPane().add(startLbl[i]);
//                    startLbl[i].setForeground(BACK_COL);

                endLbl[i] = new JLabel("",JLabel.CENTER);
                endLbl[i].setBounds(1194, iPos, 70, 24);
//                    endLbl[i].setBounds(160, iPos, 100, 24);
                endLbl[i].setLocale(new Locale("ja","JP"));
                endLbl[i].setFont(new java.awt.Font("dialog", 0, 11));
                endLbl[i].setBorder(new Flush3DBorder());
                endLbl[i].setForeground(java.awt.Color.black);
                getContentPane().add(endLbl[i]);
//                    endLbl[i].setForeground(BACK_COL);
            }
        }

        public void searchPoint(){

            beforeJudge = 0;
            lCount      = -1;

            cutStart    = null;
            cutEnd      = null;

            cutStart    = new float[25];
            cutEnd      = new float[25];
            for (int i=0; i<25; i++) {
                cutStart[i] = -1.0f;
                cutEnd[i]   = -1.0f;
            }

            int size = pv_data_body.size();
            if(2 > size) return;

            CZSystemPVData  v;
            CalcData    c;
            for (int i = 0 ; i < size ; i++) {
                if (25 == lCount) break;
                v = (CZSystemPVData)pv_data_body.elementAt(i);
                c = (CalcData)calc_data_body.elementAt(i);

                // 上下限、上々下々限との交点の引上げ長を保持する。
                switch(beforeJudge){
                    case  0 :
                        if (0 != c.judg) {
                            beforeJudge = c.judg;
                            if (-1 != lCount)
                            cutEnd[lCount]          = v.p_length;
                            lCount++;
                            if (25 == lCount) break;
                            cutStart[lCount]        = v.p_length;       //上下限
                        }
                        break;
                    case -1 :
                        if (-1 != c.judg) {
                            beforeJudge = c.judg;
                            if (-1 > c.judg) {
                                cutEnd[lCount]      = v.p_length;   //下限
                                lCount++;
                                if (25 == lCount) break;
                                cutStart[lCount]    = v.p_length;   //下々限
                            } else {
                                cutEnd[lCount]      = v.p_length;   //合格
                            }
                        } else {
                            cutEnd[lCount]          = v.p_length;
                        }
                        break;
                    case -2 :
                        if (-2 != c.judg) {
                            beforeJudge = c.judg;
                            cutEnd[lCount]          = v.p_length;   //下々限
                            lCount++;
                            if (25 == lCount) break;
                            cutStart[lCount]        = v.p_length;   //下限
                        } else {
                            cutEnd[lCount]          = v.p_length;
                        }
                        break;
                    case  1 :
                        if (1 != c.judg) {
                            beforeJudge = c.judg;
                            if (1 < c.judg) {
                                cutEnd[lCount]      = v.p_length;   //上限
                                lCount++;
                                if (25 == lCount) break;
                                cutStart[lCount]    = v.p_length;   //上々限
                            } else {
                                cutEnd[lCount]      = v.p_length;   //合格
                                lCount++;
                                if (25 == lCount) break;
                                cutStart[lCount]    = v.p_length;
                            }
                        } else {
                            cutEnd[lCount]          = v.p_length;
                        }
                        break;
                    case  2 :
                        if (2 != c.judg) {
                            beforeJudge = c.judg;
                            cutEnd[lCount]          = v.p_length;   //上々限
                            lCount++;
                            if (25 == lCount) break;
                            cutStart[lCount]        = v.p_length;   //上限
                        } else {
                            cutEnd[lCount]          = v.p_length;
                        }
                        break;
                    default :
                        break;
                } //End switch
              if (24 >= lCount) {
                if (lCount >= 0) {
              cutEnd[lCount]          = v.p_length;
              }
             }
            } //End for
        }

        public void setData() {
            for (int i = 0; i<25; i++) {
                if (lCount >= i) {
                    startLbl[i].setText("" + cutStart[i]);
                    endLbl[i].setText(""   + cutEnd[i]);
                } else {
                    startLbl[i].setText("");
                    endLbl[i].setText("");
                }
            }
        }
    }
//@@@
} //GraphDialog
