package cz;

import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.Point;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ComponentListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.text.DecimalFormat;
import java.util.Arrays;
import java.util.Comparator;
import java.util.Locale;
import java.util.Vector;

import javax.swing.JButton;
import javax.swing.JFrame;
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
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableColumnModel;
import javax.swing.table.JTableHeader;
import javax.swing.table.TableColumn;
import javax.swing.text.AttributeSet;
import javax.swing.text.BadLocationException;
import javax.swing.text.Document;
import javax.swing.text.PlainDocument;

/***********************************************************
 *
 *   制御テーブル変更Window
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01) @@  米沢版のコンバート
 * @version 1.1 (2004/01/20) @@@ グラフ画面の縮小拡大を追加
 * 2008.09.10 H.Nagamine ﾚｼﾋﾟ番号表示追加
 *
 ***********************************************************/
public class CZControlTableSub extends JFrame {

    private final int INC_WIDTH     = 236;  // 幅の増分   @@@
    private final int INC_HEIGHT    = 240;  // 高さの増分 @@@
    private final int BASE_WIDTH    = 590;  // 基準の幅   @@@
    private final int BASE_HEIGHT   = 600;  // 基準の高さ @@@
    private final int MAGNIFICATION = 5;    // 拡大の最大倍率 @@@

    private final int T1 = 1;
    private final int T2 = 2;
    private final int T3 = 3;
    private final int T4 = 4;
    private final int T5 = 5;
    private final int T6 = 6;       //@@

    private final int MODIFY_DATA       = 1;
    private final int SAVE_DATA         = 0;

    private final int SYOUJYUNN_SORT    = 1;
    private final int KOUJYUNN_SORT     = 2;
    private final int REC_MAX           = 6000;

    private final int DIA_BODY_PF       = 68;
    private final int SXL_ST_BODY_PF    = 69;
    private final int HT_BODY_PF        = 75;

    private final int MAIN1_H_T     = 14;   // 15   メインヒーター１温度
    private final int MAIN1_H_T_PF  = 66;   // 67   メインヒーター１温度プロファイル
    private final int DIA           = 24;   // 25   直径
    private final int DIA_PF        = 23;   // 24   直径プロファイル
    private final int SXL_ST        = 17;   // 18   引き上げ速度
    private final int SXL_ST_PF     = 75;   // 76   引き上げ速度プロファイル

    private final Color NEW_PRO_COL = java.awt.Color.green;
    private final Color OLD_PRO_COL = java.awt.Color.red;
    private final Color VAL_PRO_COL = java.awt.Color.white;
    private final Color VAL_COL     = java.awt.Color.orange;

//@@@@@@@@@@@@@@@@@@@@@@@@@
    private final Color MST_COL = java.awt.Color.black;
    private final Color CUR_COL = java.awt.Color.blue;

    private CZSystemCtTb    send_data[];

    private int             edit_group;         //対象グループ
    private int             edit_recip;         //レシピーNo
    private int             edit_number;        //項目No
    private CZSystemCtName  edit_name;
    private boolean         edit_current;
    private boolean         edit_haita_flg;

    private boolean         mst_show;           // @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

    private Vector          current_data;       //設定中のデータ
    
    private Vector          master_data;        //マスターデータ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@



    private Vector      pv_data_shld    = null; //ショルダーのデータ
    private Vector      pv_data_body    = null; //ボディーのデータ

    private JButton     save_button     = null; //保存ボタン
    private JButton     modify_button   = null; //修正ボタン
    private JButton     cancel_button   = null; //終了ボタン

    private TText       op_name         = null; //オペレーター名

    private CtOldTable  c_old_table     = null; //設定値を表示するテーブル
    private CtTable     c_table         = null; //変更値を表示するテーブル
    private ShiftText   shift_text      = null; //シフトさせる数値
    private ShiftText   l_shift_text      = null; //シフトさせる数値 20060529
    private BunText     l_bun_text      = null; //Ｌ軸分割数
    private BunText     r_bun_text      = null; //Ｒ軸分割数

    private JPanel      graph_panel     = null; //グラフパネル
    private LPanel      l_panel         = null; //X軸目盛
    private RPanel      r_panel         = null; //Y軸目盛
    private MainPanel   main_panel      = null; //グラフメインパネル

    private LPanelView      l_panelView         = null; //X軸目盛 @@@
    private RPanelView      r_panelView         = null; //Y軸目盛 @@@
    private MainPanelView   main_panelView      = null; //グラフメインパネル @@@

    private JButton     baseButton      = null; //基準ボタン @@@
    private JButton     reductionButton = null; //縮小ボタン @@@
    private JButton     expansionButton = null; //拡大ボタン @@@

    private JButton     mstShowButton = null; //マスター表示ボタン @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    private int         currWidth       = 590;  // 現在の幅   @@@
    private int         currHeight      = 600;  // 現在の高さ @@@

    // グラフコンポーネントリスナ
    GraphComponentListener      graphListener = null; //@@@

    private JLabel      l_lab           = null;
    private JLabel      l_unit_lab      = null;
    private JLabel      r_lab           = null;
    private JLabel      r_unit_lab      = null;

    private int         l_graph_bun     = 5;
    private int         r_graph_bun     = 5;

    private JPanel      table_panel     = null;

    private JLabel      koumoku_no_lab  = null;
    private JLabel      koumoku_lab     = null;
// add start 2008.09.10
    private JLabel      ro_no_lab       = null;
    private JLabel      group_no_lab    = null;
    private JLabel      recipe_no_lab   = null;
// add end 2008.09.10

    private JLabel      view_lab     = null;

    //
    // ---------- コンストラクタ ---------------------------
    //
    CZControlTableSub(){
        super();

        setTitle("制御テーブル設定");
        setSize(1152,864);
        setResizable(false);
        //setModal(true);

        addWindowListener(
            new WindowAdapter(){
                public void windowClosing(WindowEvent e){
                    mst_show = false;
                }
        });

        getContentPane().setLayout(null);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        JLabel label = null;

        label = new JLabel("設定者",JLabel.CENTER);
        label.setBounds(20, 790, 100, 24);                  //@@@
//        label.setBounds(20, 800, 100, 24);
        label.setLocale(new Locale("ja","JP"));
        label.setFont(new java.awt.Font("dialog", 0, 16));
        label.setBorder(new Flush3DBorder());
        label.setForeground(java.awt.Color.black);
        getContentPane().add(label);

        op_name = new TText();
        op_name.setBounds(120, 790, 140, 24);               //@@@
//        op_name.setBounds(120, 800, 140, 24);
        getContentPane().add(op_name);

        modify_button = new JButton("修  正");
        modify_button.setBounds(260, 790, 100, 24);         //@@@
//        modify_button.setBounds(260, 800, 100, 24);
        modify_button.setLocale(new Locale("ja","JP"));
        modify_button.setFont(new java.awt.Font("dialog", 0, 18));
        modify_button.setBorder(new Flush3DBorder());
        modify_button.setForeground(java.awt.Color.black);
        modify_button.addActionListener(new ModifyButton());
        getContentPane().add(modify_button);

//        save_button = new JButton("修正保存");
        save_button = new JButton("保存");				// 2004.05.27
        save_button.setBounds(360, 790, 100, 24);           //@@@
//        save_button.setBounds(360, 800, 100, 24);
        save_button.setLocale(new Locale("ja","JP"));
        save_button.setFont(new java.awt.Font("dialog", 0, 18));
        save_button.setBorder(new Flush3DBorder());
        save_button.setForeground(java.awt.Color.black);
        save_button.addActionListener(new SaveButton());
        getContentPane().add(save_button);

        cancel_button = new JButton("終  了");
        cancel_button.setBounds(630, 790, 100, 24);         //@@@
//        cancel_button.setBounds(630, 800, 100, 24);
        cancel_button.setLocale(new Locale("ja","JP"));
        cancel_button.setFont(new java.awt.Font("dialog", 0, 18));
        cancel_button.setBorder(new Flush3DBorder());
        cancel_button.setForeground(java.awt.Color.black);
        cancel_button.addActionListener(new CancelButton());
        getContentPane().add(cancel_button);

        label = new JLabel("項目",JLabel.CENTER);
// chg start 2008.09.10
//        label.setBounds(20, 20, 80, 30);
        label.setBounds(20, 40, 80, 30);
// chg end 2008.09.10
        label.setLocale(new Locale("ja","JP"));
        label.setFont(new java.awt.Font("dialog", 0, 18));
        label.setBorder(new Flush3DBorder());
        label.setForeground(java.awt.Color.black);
        getContentPane().add(label);

        koumoku_no_lab = new JLabel("###",JLabel.CENTER);
// chg start 2008.09.10
//        koumoku_no_lab.setBounds(100, 20, 60, 30);
        koumoku_no_lab.setBounds(100, 40, 60, 30);
// chg end 2008.09.10
        koumoku_no_lab.setLocale(new Locale("ja","JP"));
        koumoku_no_lab.setFont(new java.awt.Font("dialog", 0, 18));
        koumoku_no_lab.setBorder(new Flush3DBorder());
        koumoku_no_lab.setForeground(java.awt.Color.black);
        getContentPane().add(koumoku_no_lab);

        koumoku_lab = new JLabel("##########",JLabel.CENTER);
// chg start 2008.09.10
//        koumoku_lab.setBounds(160, 20, 380, 30);
        koumoku_lab.setBounds(160, 40, 380, 30);
// chg end 2008.09.10
        koumoku_lab.setLocale(new Locale("ja","JP"));
        koumoku_lab.setFont(new java.awt.Font("dialog", 0, 18));
        koumoku_lab.setBorder(new Flush3DBorder());
        koumoku_lab.setForeground(java.awt.Color.black);
        getContentPane().add(koumoku_lab);
// add start 2008.09.10
        label = new JLabel("炉番",JLabel.CENTER);
        label.setBounds(20, 10, 80, 30);
        label.setLocale(new Locale("ja","JP"));
        label.setFont(new java.awt.Font("dialog", 0, 18));
        label.setBorder(new Flush3DBorder());
        label.setForeground(java.awt.Color.black);
        getContentPane().add(label);

        String s = CZSystem.RoKetaChg(CZSystem.getRoName());
        ro_no_lab = new JLabel(s,JLabel.CENTER);
        ro_no_lab.setBounds(100, 10, 60, 30);
        ro_no_lab.setLocale(new Locale("ja","JP"));
        ro_no_lab.setFont(new java.awt.Font("dialog", 0, 18));
        ro_no_lab.setBorder(new Flush3DBorder());
        ro_no_lab.setForeground(java.awt.Color.black);
        getContentPane().add(ro_no_lab);

        label = new JLabel("グループ",JLabel.CENTER);
        label.setBounds(160, 10, 80, 30);
        label.setLocale(new Locale("ja","JP"));
        label.setFont(new java.awt.Font("dialog", 0, 18));
        label.setBorder(new Flush3DBorder());
        label.setForeground(java.awt.Color.black);
        getContentPane().add(label);

        group_no_lab = new JLabel("###",JLabel.CENTER);
        group_no_lab.setBounds(240, 10, 60, 30);
        group_no_lab.setLocale(new Locale("ja","JP"));
        group_no_lab.setFont(new java.awt.Font("dialog", 0, 18));
        group_no_lab.setBorder(new Flush3DBorder());
        group_no_lab.setForeground(java.awt.Color.black);
        getContentPane().add(group_no_lab);

        label = new JLabel("レシピ",JLabel.CENTER);
        label.setBounds(300, 10, 80, 30);
        label.setLocale(new Locale("ja","JP"));
        label.setFont(new java.awt.Font("dialog", 0, 18));
        label.setBorder(new Flush3DBorder());
        label.setForeground(java.awt.Color.black);
        getContentPane().add(label);

        recipe_no_lab = new JLabel("###",JLabel.CENTER);
        recipe_no_lab.setBounds(380, 10, 60, 30);
        recipe_no_lab.setLocale(new Locale("ja","JP"));
        recipe_no_lab.setFont(new java.awt.Font("dialog", 0, 18));
        recipe_no_lab.setBorder(new Flush3DBorder());
        recipe_no_lab.setForeground(java.awt.Color.black);
        getContentPane().add(recipe_no_lab);
// add end 2008.09.10

//@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
		mst_show = false;				//@@@@
		
		
		view_lab = new JLabel("カレント表示",JLabel.CENTER);
		view_lab.setBounds(595, 5, 140, 30);
		view_lab.setLocale(new Locale("ja","JP"));
		view_lab.setFont(new java.awt.Font("dialog", 0, 18));
		view_lab.setBorder(new Flush3DBorder());
		view_lab.setForeground(java.awt.Color.black);
		getContentPane().add(view_lab);
		
//        mstShowButton = new JButton("マスター表示");
        mstShowButton = new JButton("表示切替");
        mstShowButton.setBounds(600, 40, 130, 24);
        mstShowButton.setLocale(new Locale("ja","JP"));
        mstShowButton.setFont(new java.awt.Font("dialog", 0, 18));
        mstShowButton.setBorder(new Flush3DBorder());
        mstShowButton.setForeground(java.awt.Color.black);
        mstShowButton.addActionListener(new MasterShowAction());	//@@@@
        getContentPane().add(mstShowButton);


        // グラフ用
        graph_panel = new JPanel();
        graph_panel.setLayout(null);
//        graph_panel.setBounds(20, 70, 710 ,750);
        graph_panel.setBounds(20, 70, 710 ,710);            //@@@
        graph_panel.setBorder(new Flush3DBorder());
        graph_panel.setBackground(java.awt.Color.gray);
        getContentPane().add(graph_panel);

        label = new JLabel("L",JLabel.CENTER);
        label.setBounds(20, 10, 20, 16);
        label.setLocale(new Locale("ja","JP"));
        label.setFont(new java.awt.Font("dialog", 0, 12));
        label.setBorder(new Flush3DBorder());
        label.setForeground(java.awt.Color.black);
        graph_panel.add(label);

        l_lab = new JLabel("1234567890123456",JLabel.CENTER);
        l_lab.setBounds(40, 10, 350, 16);
        l_lab.setLocale(new Locale("ja","JP"));
        l_lab.setFont(new java.awt.Font("dialog", 0, 12));
        l_lab.setBorder(new Flush3DBorder());
        l_lab.setForeground(java.awt.Color.black);
        graph_panel.add(l_lab);

        l_unit_lab = new JLabel("1234567890123456",JLabel.CENTER);
        l_unit_lab.setBounds(390, 10, 110, 16);
        l_unit_lab.setLocale(new Locale("ja","JP"));
        l_unit_lab.setFont(new java.awt.Font("dialog", 0, 12));
        l_unit_lab.setBorder(new Flush3DBorder());
        l_unit_lab.setForeground(java.awt.Color.black);
        graph_panel.add(l_unit_lab);

        label = new JLabel("R",JLabel.CENTER);
        label.setBounds(20, 30, 20, 16);
        label.setLocale(new Locale("ja","JP"));
        label.setFont(new java.awt.Font("dialog", 0, 12));
        label.setBorder(new Flush3DBorder());
        label.setForeground(java.awt.Color.black);
        graph_panel.add(label);

        r_lab = new JLabel("1234567890123456",JLabel.CENTER);
        r_lab.setBounds(40, 30, 350, 16);
        r_lab.setLocale(new Locale("ja","JP"));
        r_lab.setFont(new java.awt.Font("dialog", 0, 12));
        r_lab.setBorder(new Flush3DBorder());
        r_lab.setForeground(java.awt.Color.black);
        graph_panel.add(r_lab);

        r_unit_lab = new JLabel("1234567890123456",JLabel.CENTER);
        r_unit_lab.setBounds(390, 30, 110, 16);
        r_unit_lab.setLocale(new Locale("ja","JP"));
        r_unit_lab.setFont(new java.awt.Font("dialog", 0, 12));
        r_unit_lab.setBorder(new Flush3DBorder());
        r_unit_lab.setForeground(java.awt.Color.black);
        graph_panel.add(r_unit_lab);

//@@@   ここから追加
        reductionButton = new JButton("縮 小");
        reductionButton.setBounds(530, 15, 50, 24);
        reductionButton.setLocale(new Locale("ja","JP"));
        reductionButton.setFont(new java.awt.Font("dialog", 0, 18));
        reductionButton.setBorder(new Flush3DBorder());
        reductionButton.setForeground(java.awt.Color.black);
        reductionButton.addActionListener(new ReductionAction());
        graph_panel.add(reductionButton);

        baseButton = new JButton("基 準");
        baseButton.setBounds(580, 15, 50, 24);
        baseButton.setLocale(new Locale("ja","JP"));
        baseButton.setFont(new java.awt.Font("dialog", 0, 18));
        baseButton.setBorder(new Flush3DBorder());
        baseButton.setForeground(java.awt.Color.black);
        baseButton.addActionListener(new StanderdAction());
        graph_panel.add(baseButton);

        expansionButton = new JButton("拡 大");
        expansionButton.setBounds(630, 15, 50, 24);
        expansionButton.setLocale(new Locale("ja","JP"));
        expansionButton.setFont(new java.awt.Font("dialog", 0, 18));
        expansionButton.setBorder(new Flush3DBorder());
        expansionButton.setForeground(java.awt.Color.black);
        expansionButton.addActionListener(new ExpansionAction());
        graph_panel.add(expansionButton);
        r_panelView = new RPanelView();

        RPanel r_panel = new RPanel();
        r_panel.setHorizontalScrollBarPolicy(javax.swing.JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
        r_panel.setVerticalScrollBarPolicy(javax.swing.JScrollPane.VERTICAL_SCROLLBAR_NEVER);
        r_panel.setView( r_panelView );
        r_panel.setBounds(20-1, 50, 50 ,BASE_HEIGHT);
        r_panel.setBorder(new Flush3DBorder());
        r_panel.setBackground(java.awt.Color.black);
        graph_panel.add(r_panel);
        r_panel.setViewSize(currHeight);

        l_panelView = new LPanelView();

        LPanel l_panel = new LPanel();
        l_panel.setHorizontalScrollBarPolicy(javax.swing.JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);
        l_panel.setVerticalScrollBarPolicy(javax.swing.JScrollPane.VERTICAL_SCROLLBAR_NEVER);
        l_panel.setView( l_panelView );
        l_panel.setBounds(70, /*680+1*/650+1, BASE_WIDTH + 15 ,50);
        l_panel.setBorder(new Flush3DBorder());
        l_panel.setBackground(java.awt.Color.black);
        graph_panel.add(l_panel);
        l_panel.setViewSize(currWidth);

        main_panelView = new MainPanelView();

        main_panel = new MainPanel();
        main_panel.setHorizontalScrollBarPolicy(javax.swing.JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
        main_panel.setVerticalScrollBarPolicy(javax.swing.JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
        main_panel.setView( main_panelView );
        main_panel.setBounds(70, 50, BASE_WIDTH+30 ,BASE_HEIGHT);
        main_panel.setBorder(new Flush3DBorder());
        main_panel.setBackground(java.awt.Color.black);
        graph_panel.add(main_panel);
        main_panel.setViewSize(currWidth,currHeight);

        // コンポーネントリスナ 生成
        GraphComponentListener graphListener = new GraphComponentListener();
        r_panelView.addComponentListener( graphListener );
        l_panelView.addComponentListener( graphListener );
        main_panelView.addComponentListener( graphListener );

        baseButton.setEnabled(false);
        reductionButton.setEnabled(false);
        expansionButton.setEnabled(true);

//@@@　ここまで
        // テーブル用
        table_panel = new JPanel();
        table_panel.setLayout(null);
        table_panel.setBounds(745, 20, 385 ,804);
        table_panel.setBorder(new Flush3DBorder());
        table_panel.setBackground(java.awt.Color.gray);
        getContentPane().add(table_panel);

        label = new JLabel("設  定  値",JLabel.CENTER);
        label.setBounds(10, 10, 178, 20);
        label.setLocale(new Locale("ja","JP"));
        label.setFont(new java.awt.Font("dialog", 0, 14));
        label.setBorder(new Flush3DBorder());
        label.setForeground(java.awt.Color.black);
        table_panel.add(label);

        label = new JLabel("変  更  値",JLabel.CENTER);
        label.setBounds(197, 10, 178, 20);
        label.setLocale(new Locale("ja","JP"));
        label.setFont(new java.awt.Font("dialog", 0, 14));
        label.setBorder(new Flush3DBorder());
        label.setForeground(java.awt.Color.black);
        table_panel.add(label);

        //設定値テーブル
        c_old_table = new CtOldTable();
        JTableHeader tabHead = c_old_table.getTableHeader();
        tabHead.setReorderingAllowed(false);
        JScrollPane panel = new JScrollPane(c_old_table);
        panel.setBounds(10, 30, 178, 613);
        table_panel.add(panel);

        //変更値テーブル
        c_table = new CtTable();
        tabHead = c_table.getTableHeader();
        tabHead.setReorderingAllowed(false);
        panel = new JScrollPane(c_table);
        panel.setBounds(197, 30, 178, 613);
        table_panel.add(panel);

/*******************************************************************************/
        JButton input_button = new JButton("＋");
        input_button.setBounds(331, 655, 44, 24);
        input_button.setLocale(new Locale("ja","JP"));
        input_button.setFont(new java.awt.Font("dialog", 0, 24));
        input_button.setBorder(new Flush3DBorder());
        input_button.setForeground(java.awt.Color.black);
        input_button.addActionListener(new InputButton());
        table_panel.add(input_button);
/*******************************************************************************/

        JButton reset_button = new JButton("再読み込み");
        reset_button.setBounds(10, 685, 178, 24);
        reset_button.setLocale(new Locale("ja","JP"));
        reset_button.setFont(new java.awt.Font("dialog", 0, 18));
        reset_button.setBorder(new Flush3DBorder());
        reset_button.setForeground(java.awt.Color.black);
        reset_button.addActionListener(new ReLoadButton());
        table_panel.add(reset_button);

        JButton repaint_button = new JButton("再  表  示");
        repaint_button.setBounds(10, 745, 178, 24);
        repaint_button.setLocale(new Locale("ja","JP"));
        repaint_button.setFont(new java.awt.Font("dialog", 0, 18));
        repaint_button.setBorder(new Flush3DBorder());
        repaint_button.setForeground(java.awt.Color.black);
        repaint_button.addActionListener(new RepaintButton());
        table_panel.add(repaint_button);

        JButton del_button = new JButton("選 択 削 除");
        del_button.setBounds(10, 715, 178, 24);
        del_button.setLocale(new Locale("ja","JP"));
        del_button.setFont(new java.awt.Font("dialog", 0, 18));
        del_button.setBorder(new Flush3DBorder());
        del_button.setForeground(java.awt.Color.black);
        del_button.addActionListener(new DeleteButton());
        table_panel.add(del_button);

        shift_text = new ShiftText();
        shift_text.setBounds(197, 685, 88, 24);
        table_panel.add(shift_text);

        JButton shift_down_button = new JButton("↓");
        shift_down_button.setBounds(286, 685, 44, 24);
        shift_down_button.setLocale(new Locale("ja","JP"));
        shift_down_button.setFont(new java.awt.Font("dialog", 0, 18));
        shift_down_button.setBorder(new Flush3DBorder());
        shift_down_button.setForeground(java.awt.Color.black);
        shift_down_button.addActionListener(new ShiftDownButton());
        table_panel.add(shift_down_button);

        JButton shift_up_button = new JButton("↑");
        shift_up_button.setBounds(331, 685, 44, 24);
        shift_up_button.setLocale(new Locale("ja","JP"));
        shift_up_button.setFont(new java.awt.Font("dialog", 0, 18));
        shift_up_button.setBorder(new Flush3DBorder());
        shift_up_button.setForeground(java.awt.Color.black);
        shift_up_button.addActionListener(new ShiftUpButton());
        table_panel.add(shift_up_button);

/**************20060529***************/
        l_shift_text = new ShiftText();
        l_shift_text.setBounds(197, 715, 88, 24);
        table_panel.add(l_shift_text);

        JButton l_shift_down_button = new JButton("←");
        l_shift_down_button.setBounds(286, 715, 44, 24);
        l_shift_down_button.setLocale(new Locale("ja","JP"));
        l_shift_down_button.setFont(new java.awt.Font("dialog", 0, 18));
        l_shift_down_button.setBorder(new Flush3DBorder());
        l_shift_down_button.setForeground(java.awt.Color.black);
        l_shift_down_button.addActionListener(new l_ShiftDownButton());
        table_panel.add(l_shift_down_button);

        JButton l_shift_up_button = new JButton("→");
        l_shift_up_button.setBounds(331, 715, 44, 24);
        l_shift_up_button.setLocale(new Locale("ja","JP"));
        l_shift_up_button.setFont(new java.awt.Font("dialog", 0, 18));
        l_shift_up_button.setBorder(new Flush3DBorder());
        l_shift_up_button.setForeground(java.awt.Color.black);
        l_shift_up_button.addActionListener(new l_ShiftUpButton());
        table_panel.add(l_shift_up_button);
/**************20060529***************/


/**************20060529***************
        label = new JLabel("Ｌ割",JLabel.CENTER);
        label.setBounds(197, 685, 44, 24);
        label.setLocale(new Locale("ja","JP"));
        label.setFont(new java.awt.Font("dialog", 0, 18));
        label.setBorder(new Flush3DBorder());
        label.setForeground(java.awt.Color.black);
        table_panel.add(label);

        l_bun_text = new BunText();
        l_bun_text.setBounds(241, 685, 44, 24);
        l_bun_text.setText(Integer.toString(l_graph_bun));
        table_panel.add(l_bun_text);

        label = new JLabel("Ｒ割",JLabel.CENTER);
        label.setBounds(286, 685, 44, 24);
        label.setLocale(new Locale("ja","JP"));
        label.setFont(new java.awt.Font("dialog", 0, 18));
        label.setBorder(new Flush3DBorder());
        label.setForeground(java.awt.Color.black);
        table_panel.add(label);

        r_bun_text = new BunText();
        r_bun_text.setBounds(330, 715, 44, 24);
        r_bun_text.setText(Integer.toString(r_graph_bun));
        table_panel.add(r_bun_text);
*/
        label = new JLabel("Ｌ割",JLabel.CENTER);
        label.setBounds(197, 745, 44, 24);
        label.setLocale(new Locale("ja","JP"));
        label.setFont(new java.awt.Font("dialog", 0, 18));
        label.setBorder(new Flush3DBorder());
        label.setForeground(java.awt.Color.black);
        table_panel.add(label);

        l_bun_text = new BunText();
        l_bun_text.setBounds(241, 745, 44, 24);
        l_bun_text.setText(Integer.toString(l_graph_bun));
        table_panel.add(l_bun_text);

        label = new JLabel("Ｒ割",JLabel.CENTER);
        label.setBounds(286, 745, 44, 24);
        label.setLocale(new Locale("ja","JP"));
        label.setFont(new java.awt.Font("dialog", 0, 18));
        label.setBorder(new Flush3DBorder());
        label.setForeground(java.awt.Color.black);
        table_panel.add(label);

        r_bun_text = new BunText();
        r_bun_text.setBounds(330, 745, 44, 24);
        r_bun_text.setText(Integer.toString(r_graph_bun));
        table_panel.add(r_bun_text);

/**************20060529***************/
    }

    //
    //
    //
    public boolean setDefault(int group,int recip,int number,
                          CZSystemCtName name,
                          boolean current,boolean haita_flg,
                          Vector v_shld,Vector v_body){

        edit_group      = group;
        edit_recip      = recip;
        edit_number     = number;
        edit_name       = name;
        edit_current    = current;
        edit_haita_flg  = haita_flg;

        pv_data_shld    = v_shld;
        pv_data_body    = v_body;   

        op_name.setText("");

//@@        CZSystem.log("CZControlTableSub setDefault",
//@@                       "[" + edit_group + "][" + edit_recip + "][" + edit_number + "]");

		if(edit_current == false){
			mstShowButton.setEnabled(false);
			view_lab.setText("マスター表示");
			mst_show = true;
		}else{
			mstShowButton.setEnabled(true);
			view_lab.setText("カレント表示");
			mst_show = false;
		}

        for(int i = 0 ; i < REC_MAX ; i++){
            c_old_table.setValueAt(null,i,1);
            c_old_table.setValueAt(null,i,2);
            c_table.setValueAt(null,i,1);
            c_table.setValueAt(null,i,2);
        }


        Vector dat =  CZSystem.getCtTb(edit_group,edit_recip,edit_number,edit_current);

        /* マスターデータ取得 **************************************/
        Vector mstdat =  CZSystem.getCtTb(edit_group,edit_recip,edit_number,false);

        if(null == dat){
            current_data = null;
            if(edit_current){
                Object msg[] = {"操業制御テーブル",
                                "テーブルが存在しません！！",
                                ""};
                errorMsg(msg);
            }
            else {
                Object msg[] = {"制御テーブル",
                                "テーブルが存在しません！！",
                                ""};
                errorMsg(msg);
            }

            Float l = new Float(0.0f);
            Float r = new Float(0.0f);
            c_old_table.setValueAt(l,0,1);
            c_old_table.setValueAt(r,0,2);
            c_table.setValueAt(l,0,1);
            c_table.setValueAt(r,0,2);
        }
        else {
            current_data = dat;
            master_data = mstdat;

			if(edit_current == false){
	            setCurrent(master_data,edit_name.k_sort);
			}else{
	            setCurrent(current_data,edit_name.k_sort);
			}
        }

        koumoku_no_lab.setText(String.valueOf(edit_name.t_no));
        koumoku_lab.setText(edit_name.t_name);

        l_lab.setText(edit_name.l_name);
        l_unit_lab.setText(edit_name.l_unit);
        r_lab.setText(edit_name.r_name);
        r_unit_lab.setText(edit_name.r_unit);
// add start 2008.09.10
        group_no_lab.setText('T' + String.valueOf(edit_group));
        recipe_no_lab.setText(String.valueOf(edit_recip));
        String s = CZSystem.RoKetaChg(CZSystem.getRoName());
        ro_no_lab.setText(s);
// add end 2008.09.10

        if(edit_haita_flg){
            if(edit_current){
                save_button.setEnabled(true);
                modify_button.setEnabled(true);
            }
            else {
                save_button.setEnabled(true);
                modify_button.setEnabled(false);
            }
        }
        else {
            save_button.setEnabled(false);
            modify_button.setEnabled(false);
        }
		
		if(mst_show == false){
			c_old_table.getRender1().setColor(CUR_COL);
			c_old_table.getRender2().setColor(CUR_COL);
		}else{
			c_old_table.getRender1().setColor(MST_COL);
			c_old_table.getRender2().setColor(MST_COL);
		}
        c_old_table.repaint();
        c_old_table.clearSelection();
        c_table.repaint();
        c_table.clearSelection();

//@@@ ここから　画面のサイズを初期化する。
        currHeight = BASE_HEIGHT;
        currWidth  = BASE_WIDTH;
        r_panelView.setRPanelViewSize(currHeight);
        l_panelView.setLPanelViewSize(currWidth);
        main_panelView.setMainPanelViewSize(currWidth,currHeight);
        r_panelView.repaint();
        l_panelView.repaint();
        main_panelView.repaint();
        baseButton.setEnabled(false);
        reductionButton.setEnabled(false);
        expansionButton.setEnabled(true);
//@@@ ここまで
        return true;
    }


    //
    //
    //
	@SuppressWarnings("unchecked")
    private boolean setCurrent(Vector dat,int k_sort){

        Float l = null;
        Float r = null;

        int   size = dat.size();

        CZSystemCtTb data[] = new CZSystemCtTb[size];

        for(int i = 0 ; i < size ; i++){
            data[i] = (CZSystemCtTb)dat.elementAt(i);
        }

        if(KOUJYUNN_SORT == k_sort){
            Arrays.sort(data, new Sort2());
        }
        else {
            Arrays.sort(data, new Sort1());
        }

        for(int i = 0 ; i < size ; i++){
            l = new Float(data[i].l_val);
            r = new Float(data[i].r_val);
            c_old_table.setValueAt(l,i,1);
            c_old_table.setValueAt(r,i,2);
            c_table.setValueAt(l,i,1);
            c_table.setValueAt(r,i,2);
        }
        return true;
    }

    //
    //  変更値のみ表示リフレッシュ
    //
	@SuppressWarnings("unchecked")
    private boolean setCurrent2(Vector dat,int k_sort){

        Float l = null;
        Float r = null;

        int   size = dat.size();

        CZSystemCtTb data[] = new CZSystemCtTb[size];

        for(int i = 0 ; i < size ; i++){
            data[i] = (CZSystemCtTb)dat.elementAt(i);
        }

        if(KOUJYUNN_SORT == k_sort){
            Arrays.sort(data, new Sort2());
        }
        else {
            Arrays.sort(data, new Sort1());
        }

        for(int i = 0 ; i < size ; i++){
            l = new Float(data[i].l_val);
            r = new Float(data[i].r_val);
//            c_old_table.setValueAt(l,i,1);
//            c_old_table.setValueAt(r,i,2);
            c_table.setValueAt(l,i,1);
            c_table.setValueAt(r,i,2);
        }
        return true;
    }

    //
    //
    //
	@SuppressWarnings("unchecked")
    private boolean reSetData(){
        Vector dat = new Vector();

        for(int i = 0 ; i < REC_MAX ; i++){
            Float l = (Float)c_table.getValueAt(i,1);
            Float r = (Float)c_table.getValueAt(i,2);

            if(null == l) continue;
            if(null == r) continue;

            CZSystemCtTb data = new CZSystemCtTb();
            data.l_val = l.floatValue();
            data.r_val = r.floatValue();
            dat.addElement(data);
        }

        for(int i = 0 ; i < REC_MAX ; i++){
            c_table.setValueAt(null,i,1);
            c_table.setValueAt(null,i,2);
        }

        setCurrent(dat,edit_name.k_sort);

        // テーブル全て消した時の処理
        Float l_val = (Float)c_table.getValueAt(0,1);
        Float r_val = (Float)c_table.getValueAt(0,2);
        if((null == l_val) || (null == r_val)){
            Float l = new Float(0.0f);
            Float r = new Float(0.0f);
            c_table.setValueAt(l,0,1);
            c_table.setValueAt(r,0,2);
        }
        return true;
    }

    //
    //  変更値のみ表示リフレッシュ
    //
	@SuppressWarnings("unchecked")
    private boolean reSetData2(){
        Vector dat = new Vector();

        for(int i = 0 ; i < REC_MAX ; i++){
            Float l = (Float)c_table.getValueAt(i,1);
            Float r = (Float)c_table.getValueAt(i,2);

            if(null == l) continue;
            if(null == r) continue;

            CZSystemCtTb data = new CZSystemCtTb();
            data.l_val = l.floatValue();
            data.r_val = r.floatValue();
            dat.addElement(data);
        }

        for(int i = 0 ; i < REC_MAX ; i++){
            c_table.setValueAt(null,i,1);
            c_table.setValueAt(null,i,2);
        }

        setCurrent2(dat,edit_name.k_sort);

        // テーブル全て消した時の処理
        Float l_val = (Float)c_table.getValueAt(0,1);
        Float r_val = (Float)c_table.getValueAt(0,2);
        if((null == l_val) || (null == r_val)){
            Float l = new Float(0.0f);
            Float r = new Float(0.0f);
            c_table.setValueAt(l,0,1);
            c_table.setValueAt(r,0,2);
        }
        return true;
    }


    //
    //
    //
	@SuppressWarnings("unchecked")
    private boolean chkData(){

        Vector dat = new Vector();

        for(int i = 0 ; i < REC_MAX ; i++){
            Float l = (Float)c_table.getValueAt(i,1);
            Float r = (Float)c_table.getValueAt(i,2);

            // 片方が null
            if(l == null && r != null) return false;
            if(l != null && r == null) return false;

            // 両方が null
            if(null == l) continue;
            if(null == r) continue;

            CZSystemCtTb data = new CZSystemCtTb();
            data.l_val = l.floatValue();
            data.r_val = r.floatValue();

            // Ｌ軸
            if((edit_name.l_min > data.l_val) ||
                   (edit_name.l_max < data.l_val)){
                return false;
            }

            // Ｒ軸
            if((edit_name.r_min > data.r_val) ||
                   (edit_name.r_max < data.r_val)){
                return false;
            }

            dat.addElement(data);
        }

        int   size = dat.size();
        CZSystemCtTb data[] = new CZSystemCtTb[size];
        for(int i = 0 ; i < size ; i++){
            data[i] = (CZSystemCtTb)dat.elementAt(i);
        }

        // Ｌ軸に昇順ソート
        Arrays.sort(data, new Sort1());
        send_data = data;
        return true;
    }


    //
    // メッセージの表示
    //
    private boolean errorMsg(Object msg[]){
        JOptionPane.showMessageDialog(null,msg,
                                    "制御テーブルエラー",
                                    JOptionPane.ERROR_MESSAGE);
        return true;
    }


    /*******************************************************
     *
     * 修正ボタンの処理
     *
     *******************************************************/
    class ModifyButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            if(c_table.isEditing()){
                CZSystem.log("CZControlTableSub ModifyButton"," actionPerformed Table Data EDIT !!");
                Object msg[] = {"制御テーブル",
                                "設定中項目有り！！",
                                ""};
                errorMsg(msg);
                return ;
            }

            if(1 > op_name.getText().length()){
                CZSystem.log("CZControlTableSub ModifyButton","actionPerformed Table Op Name Error !!");
                Object msg[] = {"制御テーブル",
                                "設定者を入力してくださ！！",
                                ""};
                errorMsg(msg);
                return ;
            }

            if(chkData()){
                int size = send_data.length;
                float leftData[]  = new float[size];
                float rightData[] = new float[size];

                for(int i = 0 ; i < size ; i++){
                    CZSystem.log("CZControlTableSub ModifyButton",
                        "actionPerformed [" + size + "][" + i + "][" + send_data[i].l_val +
                        "][" + send_data[i].r_val + "]");

                    leftData[i]  = send_data[i].l_val;
                    rightData[i] = send_data[i].r_val;
                }

// @@@@@ 2014.01.09
				// 異常判定処理
				int jflg = 0;	// 0:異常判定処理不要  1:異常判定処理開始
				
				// カレントテーブル読込み
				Vector dat = null;
				if ((edit_group == 5) && (edit_number == 15)) {	// 画面側 項目Noが15の場合
					dat =  CZSystem.getCtTb(5,edit_recip,20,edit_current);		// 読込みデータの項目Noは20
					CZSystem.log("CZControlTableSub ","項目No.20のデータサイズ:" + dat.size());
					jflg = 1;
				}else if ((edit_group == 5) && (edit_number == 20)) {	// 画面側 項目Noが20の場合
					dat =  CZSystem.getCtTb(5,edit_recip,15,edit_current);		// 読込みデータの項目Noは15
					CZSystem.log("CZControlTableSub ","項目No.15のデータサイズ:" + dat.size());
					jflg = 1;
				}else{
					jflg = 0;	// 異常判定処理停止
				}
				
				if(jflg != 0){
					// 判定対象テーブルデータ（ＤＢ読込みデータ）準備
					CZSystemCtTb data[] = new CZSystemCtTb[dat.size()];
					
					for(int i = 0 ; i < dat.size() ; i++){
						data[i] = (CZSystemCtTb)dat.elementAt(i);
						CZSystem.log("CZControlTableSub ","data.l_val[" + data[i].l_val + "] : data.r_val[" + data[i].r_val + "]");
					}
					
					CZSystem.log("CZControlTableSub ","画面側　Ｌ軸最小値 : " + leftData[0]);
					CZSystem.log("CZControlTableSub ","画面側　Ｌ軸最大値 : " + leftData[size-1]);
					CZSystem.log("CZControlTableSub ","ＤＢ側　Ｌ軸最小値 : " + data[0].l_val);
					CZSystem.log("CZControlTableSub ","ＤＢ側　Ｌ軸最大値 : " + data[dat.size()-1].l_val);
					
					// 画面側とＤＢ側のデータ(Ｌ軸値)がラップしているかチェック
					if(leftData[0] >= data[dat.size()-1].l_val){
						CZSystem.log("CZControlTableSub ","画面側　Ｌ軸最小値 >= ＤＢ側　Ｌ軸最大値");
						CZSystem.log("CZControlTableSub ","カレントデータ　異常判定処理不要");
					}else if(leftData[size-1] <= data[0].l_val){
						CZSystem.log("CZControlTableSub ","画面側　Ｌ軸最大値 <= ＤＢ側　Ｌ軸最小値");
						CZSystem.log("CZControlTableSub ","カレントデータ　異常判定処理不要");
					}else{
						CZSystem.log("CZControlTableSub ","カレントデータ　異常判定処理スタート");
						
						int lb_flg = 0;		// ループ処理中断フラグ⇒1:中断
						for(int i = 0 ; i < size-1 ; i++){
							if(rightData[i] != rightData[i+1]){		// 画面側　Ｒ軸値可変のＬ軸範囲特定
								CZSystem.log("CZControlTableSub ","画面側(L軸値)   S [#" + (i+1) + "]: " + leftData[i] + " E [#" + (i+1+1) + "]: " + leftData[i+1]);
								CZSystem.log("CZControlTableSub ","画面側(R軸値)   S [#" + (i+1) + "]: " + rightData[i] + " E [#" + (i+1+1) + "]: " + rightData[i+1]);
								
								for(int j = 0 ; j < dat.size()-1 ; j++){
									if(data[j].r_val != data[j+1].r_val){	// ＤＢ側　Ｒ軸値可変のＬ軸範囲特定
										CZSystem.log("CZControlTableSub ","ＤＢ側(L軸値)   S [#" + (j+1) + "]: " + data[j].l_val + " E [#" + (j+1+1) + "]: " + data[j+1].l_val);
										CZSystem.log("CZControlTableSub ","ＤＢ側(R軸値)   S [#" + (j+1) + "]: " + data[j].r_val + " E [#" + (j+1+1) + "]: " + data[j+1].r_val);
										
										// 画面側の(Ｌ軸)可変範囲にＤＢ側の(Ｌ軸)可変範囲がラップしているかチェック
										if(leftData[i] >= data[j+1].l_val){
											CZSystem.log("CZControlTableSub ","可変範囲　(画面側)Ｌ軸最小値 " + leftData[i] + " >= " + "(ＤＢ側)Ｌ軸最大値" + data[j+1].l_val);
											CZSystem.log("CZControlTableSub ","可変範囲ラップ無し！　設定値異常無し！");
										}else if(leftData[i+1] <= data[j].l_val){
											CZSystem.log("CZControlTableSub ","可変範囲　(画面側)Ｌ軸最大値 " + leftData[i+1] + " <= " + "(ＤＢ側)Ｌ軸最小値" + data[j].l_val);
											CZSystem.log("CZControlTableSub ","可変範囲ラップ無し！　設定値異常無し！");
										}else{
											CZSystem.log("CZControlTableSub ","可変範囲ラップあり！！　設定値異常あり！！");
											
											Object msg[] = {"磁場可変設定異常",
															"マグネット１磁場強度PFと",
															"マグネット位置PF設定を",
															"確認してください！"};
											errorMsg(msg);
											lb_flg = 1;		//ループ処理終了フラグセット
										}
									}
									
									if(lb_flg == 1){	//ループ処理終了
										CZSystem.log("CZControlTableSub ","ループ処理終了");
										break;
									}
								}
							}
							
							if(lb_flg == 1){	//ループ処理終了
								CZSystem.log("CZControlTableSub ","ループ処理終了");
								break;
							}
						}
					}
				}else{
					CZSystem.log("CZControlTableSub ","カレントデータ　異常判定処理不要");
				}
				
// @@@@@ 2014.01.09

                CZSystem.CZControlTableExchange(op_name.getText(),MODIFY_DATA,edit_group,
                                                edit_recip, edit_number,leftData,rightData);

            }
            else {
                Object msg[] = {"制御テーブル",
                                "値を確認してくださ！！",
                                ""};
                errorMsg(msg);
            }
            return ;
        }
    }

    /*******************************************************
     *
     * 保存ボタン
     *
     *******************************************************/
    class SaveButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            if(c_table.isEditing()){
                CZSystem.log("CZControlTableSub SaveButton","actionPerformed Table Data EDIT !!");
                Object msg[] = {"制御テーブル",
                                "設定中項目有り！！",
                                ""};
                errorMsg(msg);
                return ;
            }

            if(1 > op_name.getText().length()){
                CZSystem.log("CZControlTableSub SaveButton","actionPerformed Table Op Name Error !!");

                Object msg[] = {"制御テーブル",
                                "設定者を入力してください！！",
                                ""};
                errorMsg(msg);
                return ;
            }

            if(chkData()){
                int size = send_data.length;
                float leftData[]  = new float[size];
                float rightData[] = new float[size];

                for(int i = 0 ; i < size ; i++){
                    CZSystem.log("CZControlTableSub SaveButton",
                        "actionPerformed [" + size + "][" + i + "][" + send_data[i].l_val +
                        "][" + send_data[i].r_val + "]");

                    leftData[i]  = send_data[i].l_val;
                    rightData[i] = send_data[i].r_val;
                }

// @@@@@ 2014.01.09
				// 異常判定処理
				int jflg = 0;	// 0:異常判定処理不要  1:異常判定処理開始
				
				// マスターテーブル読込み
				Vector dat = null;
				if ((edit_group == 5) && (edit_number == 15)) {	// 画面側 項目Noが15の場合
					dat =  CZSystem.getCtTb(5,edit_recip,20,false);		// 読込みデータの項目Noは20
					CZSystem.log("CZControlTableSub ","項目No.20のデータサイズ:" + dat.size());
					jflg = 1;
				}else if ((edit_group ==5) && (edit_number == 20)) {	// 画面側 項目Noが20の場合
					dat =  CZSystem.getCtTb(5,edit_recip,15,false);		// 読込みデータの項目Noは15
					CZSystem.log("CZControlTableSub ","項目No.15のデータサイズ:" + dat.size());
					jflg = 1;
				}else{
					jflg = 0;	// 異常判定処理停止
				}
				
				if(jflg != 0){
					// 判定対象テーブルデータ（ＤＢ読込みデータ）準備
					CZSystemCtTb data[] = new CZSystemCtTb[dat.size()];
					
					for(int i = 0 ; i < dat.size() ; i++){
						data[i] = (CZSystemCtTb)dat.elementAt(i);
						CZSystem.log("CZControlTableSub ","data.l_val[" + data[i].l_val + "] : data.r_val[" + data[i].r_val + "]");
					}
					
					CZSystem.log("CZControlTableSub ","画面側　Ｌ軸最小値 : " + leftData[0]);
					CZSystem.log("CZControlTableSub ","画面側　Ｌ軸最大値 : " + leftData[size-1]);
					CZSystem.log("CZControlTableSub ","ＤＢ側　Ｌ軸最小値 : " + data[0].l_val);
					CZSystem.log("CZControlTableSub ","ＤＢ側　Ｌ軸最大値 : " + data[dat.size()-1].l_val);
					
					// 画面側とＤＢ側のデータ(Ｌ軸値)がラップしているかチェック
					if(leftData[0] >= data[dat.size()-1].l_val){
						CZSystem.log("CZControlTableSub ","画面側　Ｌ軸最小値 >= ＤＢ側　Ｌ軸最大値");
						CZSystem.log("CZControlTableSub ","カレントデータ　異常判定処理不要");
					}else if(leftData[size-1] <= data[0].l_val){
						CZSystem.log("CZControlTableSub ","画面側　Ｌ軸最大値 <= ＤＢ側　Ｌ軸最小値");
						CZSystem.log("CZControlTableSub ","マスターデータ　異常判定処理不要");
					}else{
						CZSystem.log("CZControlTableSub ","マスターデータ　異常判定処理スタート");
						
						int lb_flg = 0;		// ループ処理中断フラグ⇒1:中断
						for(int i = 0 ; i < size-1 ; i++){
							if(rightData[i] != rightData[i+1]){		// 画面側　Ｒ軸値可変のＬ軸範囲特定
								CZSystem.log("CZControlTableSub ","画面側(L軸値)   S [#" + (i+1) + "]: " + leftData[i] + " E [#" + (i+1+1) + "]: " + leftData[i+1]);
								CZSystem.log("CZControlTableSub ","画面側(R軸値)   S [#" + (i+1) + "]: " + rightData[i] + " E [#" + (i+1+1) + "]: " + rightData[i+1]);
								
								for(int j = 0 ; j < dat.size()-1 ; j++){
									if(data[j].r_val != data[j+1].r_val){	// ＤＢ側　Ｒ軸値可変のＬ軸範囲特定
										CZSystem.log("CZControlTableSub ","ＤＢ側(L軸値)   S [#" + (j+1) + "]: " + data[j].l_val + " E [#" + (j+1+1) + "]: " + data[j+1].l_val);
										CZSystem.log("CZControlTableSub ","ＤＢ側(R軸値)   S [#" + (j+1) + "]: " + data[j].r_val + " E [#" + (j+1+1) + "]: " + data[j+1].r_val);
										
										// 画面側の(Ｌ軸)可変範囲にＤＢ側の(Ｌ軸)可変範囲がラップしているかチェック
										if(leftData[i] >= data[j+1].l_val){
											CZSystem.log("CZControlTableSub ","可変範囲　(画面側)Ｌ軸最小値 " + leftData[i] + " >= " + "(ＤＢ側)Ｌ軸最大値" + data[j+1].l_val);
											CZSystem.log("CZControlTableSub ","可変範囲ラップ無し！　設定値異常無し！");
										}else if(leftData[i+1] <= data[j].l_val){
											CZSystem.log("CZControlTableSub ","可変範囲　(画面側)Ｌ軸最大値 " + leftData[i+1] + " <= " + "(ＤＢ側)Ｌ軸最小値" + data[j].l_val);
											CZSystem.log("CZControlTableSub ","可変範囲ラップ無し！　設定値異常無し！");
										}else{
											CZSystem.log("CZControlTableSub ","可変範囲ラップあり！！　設定値異常あり！！");
											
											Object msg[] = {"磁場可変設定異常",
															"マグネット１磁場強度PFと",
															"マグネット位置PF設定を",
															"確認してください！"};
											errorMsg(msg);
											lb_flg = 1;		//ループ処理終了フラグセット
										}
									}
									
									if(lb_flg == 1){	//ループ処理終了
										CZSystem.log("CZControlTableSub ","ループ処理終了");
										break;
									}
								}
							}
							
							if(lb_flg == 1){	//ループ処理終了
								CZSystem.log("CZControlTableSub ","ループ処理終了");
								break;
							}
						}
					}
				}else{
					CZSystem.log("CZControlTableSub ","マスターデータ　異常判定処理不要");
				}
				
// @@@@@ 2014.01.09

                CZSystem.CZControlTableExchange(op_name.getText(),SAVE_DATA,edit_group,
                                                edit_recip, edit_number,leftData,rightData);
            }
            else {
                Object msg[] = {"制御テーブル",
                                "値を確認してください！！",
                                ""};
                errorMsg(msg);
            }
            return ;
        }
    }


    /*******************************************************
     *
     * 終了ボタン
     *
     *******************************************************/
    class CancelButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            setVisible(false);
        }
    }


    /*******************************************************
     *
     * 修正値入力ボタンの処理
     *
     *******************************************************/
    class InputButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
			Float l_val1; // 入力値（L軸）
			Float r_val1; // 計算結果値（R軸）
			Float l_val2; // （昇順時）入力値よりひとつ小さい値（L軸）
			Float r_val2; // （昇順時）入力値よりひとつ小さい値（R軸）
			Float l_val3; // （昇順時）入力値よりひとつ大きい値（L軸）
			Float r_val3; // （昇順時）入力値よりひとつ大きい値（R軸）
			Float l_val4; // （降順時）入力値よりひとつ大きい値（L軸）
			Float r_val4; // （降順時）入力値よりひとつ大きい値（R軸）
			Float l_val5; // （降順時）入力値よりひとつ小さい値（L軸）
			Float r_val5; // （降順時）入力値よりひとつ小さい値（R軸）
			
			
			CZSystem.log("CZControlTableSub InputButton","edit_name.k_sort: " + edit_name.k_sort);
			
			// 選択行
			int row = c_table.getSelectedRow();
			
			if(row == -1){
				return;
			}
			
			// 入力値（L軸）
			l_val1 = (Float)c_table.getValueAt(row,1);
			
			if(l_val1 == null){
				return;
			}
			
			if(edit_name.k_sort == 1){	/* 値が昇順のとき */
				for(int i = 0; i < REC_MAX; i++){
					l_val3 = (Float)c_table.getValueAt(i,1);
					CZSystem.log("CZControlTableSub InputButton","l_val3 入力値よりひとつ大きい値（L軸）: " + l_val3);
					
					if(l_val3 == null){
						Object msg[] = {"入力した数値が２点間の値ではありません",
										"値を確認してください！！",
										""};
						errorMsg(msg);
						return;
					}
					
					if(l_val1 < l_val3){
						if(i == 0){
							Object msg[] = {"入力した数値が２点間の値ではありません",
											"値を確認してください！！",
											""};
							errorMsg(msg);
							return;
						}
						
						// ２点間の値を取得
						l_val2 = (Float)c_table.getValueAt(i-1,1);
						r_val2 = (Float)c_table.getValueAt(i-1,2);
						l_val3 = (Float)c_table.getValueAt(i,1);
						r_val3 = (Float)c_table.getValueAt(i,2);
						
						CZSystem.log("CZControlTableSub InputButton","l_val1（入力値） : " + l_val1);
						CZSystem.log("CZControlTableSub InputButton","l_val2 入力値よりひとつ小さい値（L軸）: " + l_val2);
						CZSystem.log("CZControlTableSub InputButton","r_val2 入力値よりひとつ小さい値（R軸）: " + r_val2);
						CZSystem.log("CZControlTableSub InputButton","l_val3 入力値よりひとつ大きい値（L軸）: " + l_val3);
						CZSystem.log("CZControlTableSub InputButton","r_val3 入力値よりひとつ大きい値（R軸）: " + r_val3);
						
						// 計算結果値（R軸）
						r_val1 = ((r_val3 - r_val2) / (l_val3 - l_val2) * (l_val1 - l_val2) + r_val2);
						
						CZSystem.log("CZControlTableSub InputButton","r_val1（計算結果値（R軸））: " + r_val1);
						
						if(true == l_val1.equals(l_val2)){
							Object msg[] = {"入力した数値が２点間の値ではありません",
											"値を確認してください！！",
											""};
							errorMsg(msg);
							return;
						}
						
						// 自動計算結果値を表示
						c_table.setValueAt(l_val1,row,1);
						c_table.setValueAt(r_val1,row,2);
						
						// 行選択解除
						c_table.clearSelection();
						return;
					}
				}
			}else{	/* 値が降順のとき */
				for(int i = 0; i < REC_MAX; i++){
					l_val4 = (Float)c_table.getValueAt(i,1);
					CZSystem.log("CZControlTableSub InputButton","l_val4 入力値よりひとつ大きい値（L軸）: " + l_val4);
					
					if(l_val4 == null){
						Object msg[] = {"入力した数値が２点間の値ではありません",
										"値を確認してください！！",
										""};
						errorMsg(msg);
						return;
					}
					
					if(l_val1 > l_val4){
						if(i == 0){
							Object msg[] = {"入力した数値が２点間の値ではありません",
											"値を確認してください！！",
											""};
							errorMsg(msg);
							return;
						}
						
						// ２点間の値を取得
						l_val4 = (Float)c_table.getValueAt(i-1,1);
						r_val4 = (Float)c_table.getValueAt(i-1,2);
						l_val5 = (Float)c_table.getValueAt(i,1);
						r_val5 = (Float)c_table.getValueAt(i,2);
						
						CZSystem.log("CZControlTableSub InputButton","l_val1（入力値） : " + l_val1);
						CZSystem.log("CZControlTableSub InputButton","l_val4 入力値よりひとつ大きい値（L軸）: " + l_val4);
						CZSystem.log("CZControlTableSub InputButton","r_val4 入力値よりひとつ大きい値（R軸）: " + r_val4);
						CZSystem.log("CZControlTableSub InputButton","l_val5 入力値よりひとつ小さい値（L軸）: " + l_val5);
						CZSystem.log("CZControlTableSub InputButton","r_val5 入力値よりひとつ小さい値（R軸）: " + r_val5);
						
						// 計算結果値（R軸）
						r_val1 = ((r_val4 - r_val5) / (l_val4 - l_val5) * (l_val1 - l_val5) + r_val5);
						
						CZSystem.log("CZControlTableSub InputButton","r_val1（計算結果値（R軸））: " + r_val1);
						
						if(true == l_val1.equals(l_val4)){
							Object msg[] = {"入力した数値が２点間の値ではありません",
											"値を確認してください！！",
											""};
							errorMsg(msg);
							return;
						}
						
						// 自動計算結果値を表示
						c_table.setValueAt(l_val1,row,1);
						c_table.setValueAt(r_val1,row,2);
						
						// 行選択解除
						c_table.clearSelection();
						return;
					}
				}
			}
        }
    }

    /*******************************************************
     *
     * 再読込ボタンの処理
     *
     *******************************************************/
    class ReLoadButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            setDefault(edit_group,edit_recip,edit_number,
                   edit_name,edit_current,edit_haita_flg,
                   pv_data_shld,pv_data_body);
        }
    }

    /*******************************************************
     *
     * 再表示
     *
     *******************************************************/
    class RepaintButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            String s_l_val;
            String s_r_val;

            int l = l_graph_bun;
            int r = r_graph_bun;

            try{
                s_l_val = l_bun_text.getText();
                s_r_val = r_bun_text.getText();

                l = Integer.parseInt(s_l_val);
                r = Integer.parseInt(s_r_val);
            }
            catch( Exception e){
                return;
            }

            l_graph_bun = l;
            r_graph_bun = r;

            l_panelView.repaint();      //@@@
            r_panelView.repaint();      //@@@
            main_panelView.repaint();   //@@@
        }
    }

    /*******************************************************
     *
     * 選択削除ボタンの処理
     *
     *******************************************************/
    class DeleteButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            int row_list[] = c_table.getSelectedRows();

            for(int i = 0 ; i < row_list.length ; i++){
                CZSystem.log("CZControlTableSub DeleteButton","actionPerformed [" + i + "][" + row_list[i] + "]");

                c_table.setValueAt(null,row_list[i],1);
                c_table.setValueAt(null,row_list[i],2);
            }

            reSetData2();

            c_table.repaint();
            c_table.clearSelection();
            main_panelView.repaint();       //@@@
        }
    }


    /*******************************************************
     *
     * ↓ボタンの処理
     *
     *******************************************************/
    class ShiftDownButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            String s_val;
            Float  f_val;
            Float  old_val;
            Float  new_val;

            try{
                s_val = shift_text.getText();
                f_val = new Float(s_val);
            }
            catch( Exception e){
                return;
            }

            int row_list[] = c_table.getSelectedRows();

            for(int i = 0 ; i < row_list.length ; i++){
                CZSystem.log("CZControlTableSub ShiftDownButton","actionPerformed [" + i +
                        "][" + row_list[i] + "]");

                old_val = (Float)c_table.getValueAt(row_list[i],2);
                if(null == old_val) continue;

                new_val = new Float(old_val.floatValue() - f_val.floatValue());
                c_table.setValueAt(new_val,row_list[i],2);
            }
            reSetData2();
            c_table.repaint();
            main_panelView.repaint();       //@@@
        }
    }

    /*******************************************************
     *
     * ↑ボタンの処理
     *
     *******************************************************/
    class ShiftUpButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            String s_val;
            Float  f_val;
            Float  old_val;
            Float  new_val;

            try{
                s_val = shift_text.getText();
                f_val = new Float(s_val);
            }
            catch( Exception e){
                return;
            }

            int row_list[] = c_table.getSelectedRows();

            for(int i = 0 ; i < row_list.length ; i++){
                CZSystem.log("CZControlTableSub ShiftUpButton","actionPerformed [" + i +
                    "][" + row_list[i] + "]");

                old_val = (Float)c_table.getValueAt(row_list[i],2);
                if(null == old_val) continue;

                new_val = new Float(old_val.floatValue() + f_val.floatValue());
                c_table.setValueAt(new_val,row_list[i],2);
            }
            reSetData2();

            c_table.repaint();
            main_panelView.repaint();       //@@@
        }
    }

    /*******************************************************
     *
     * ←ボタンの処理 20060529
     *
     *******************************************************/
    class l_ShiftDownButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            String s_val;
            Float  f_val;
            Float  old_val;
            Float  new_val;

            try{
                s_val = l_shift_text.getText();
                f_val = new Float(s_val);
            }
            catch( Exception e){
                return;
            }

            int row_list[] = c_table.getSelectedRows();

            for(int i = 0 ; i < row_list.length ; i++){
                CZSystem.log("CZControlTableSub l_ShiftDownButton","actionPerformed [" + i +
                        "][" + row_list[i] + "]");

                old_val = (Float)c_table.getValueAt(row_list[i],1);
                if(null == old_val) continue;

                new_val = new Float(old_val.floatValue() - f_val.floatValue());
                c_table.setValueAt(new_val,row_list[i],1);
            }
            reSetData2();
            c_table.repaint();
            main_panelView.repaint();       //@@@
        }
    }

    /*******************************************************
     *
     * →ボタンの処理 20060529
     *
     *******************************************************/
    class l_ShiftUpButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            String s_val;
            Float  f_val;
            Float  old_val;
            Float  new_val;

            try{
                s_val = l_shift_text.getText();
                f_val = new Float(s_val);
            }
            catch( Exception e){
                return;
            }

            int row_list[] = c_table.getSelectedRows();

            for(int i = 0 ; i < row_list.length ; i++){
                CZSystem.log("CZControlTableSub l_ShiftUpButton","actionPerformed [" + i +
                    "][" + row_list[i] + "]");

                old_val = (Float)c_table.getValueAt(row_list[i],1);
                if(null == old_val) continue;

                new_val = new Float(old_val.floatValue() + f_val.floatValue());
                c_table.setValueAt(new_val,row_list[i],1);
            }
            reSetData2();

            c_table.repaint();
            main_panelView.repaint();       //@@@
        }
    }

    /*******************************************************
     *
     * マスター表示ボタン @@@@
     *
     *******************************************************/
    class MasterShowAction implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            if (mst_show == true) {
				view_lab.setText("カレント表示");
                mst_show = false;

		        for(int i = 0 ; i < REC_MAX ; i++){
		            c_old_table.setValueAt(null,i,1);
		            c_old_table.setValueAt(null,i,2);
		            c_table.setValueAt(null,i,1);
		            c_table.setValueAt(null,i,2);
		        }

/*
                setDefault(edit_group,edit_recip,edit_number,
                edit_name,edit_current,edit_haita_flg,
                pv_data_shld,pv_data_body);
*/
	            setCurrent(current_data,edit_name.k_sort);
				c_old_table.getRender1().setColor(CUR_COL);
				c_old_table.getRender2().setColor(CUR_COL);

            } else {
				view_lab.setText("マスター表示");
                mst_show = true;

		        for(int i = 0 ; i < REC_MAX ; i++){
		            c_old_table.setValueAt(null,i,1);
		            c_old_table.setValueAt(null,i,2);
		            c_table.setValueAt(null,i,1);
		            c_table.setValueAt(null,i,2);
		        }

	            setCurrent(master_data,edit_name.k_sort);
				c_old_table.getRender1().setColor(MST_COL);
				c_old_table.getRender2().setColor(MST_COL);

            }
	        c_old_table.repaint();
	        c_old_table.clearSelection();
            main_panelView.repaint();
        }
    }

    /*******************************************************
     *
     * 縮小ボタン @@@
     *
     *******************************************************/
    class ReductionAction implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            currHeight = currHeight - INC_HEIGHT;
            currWidth  = currWidth - INC_WIDTH;
            r_panelView.setRPanelViewSize(currHeight);
            l_panelView.setLPanelViewSize(currWidth);
            main_panelView.setMainPanelViewSize(currWidth,currHeight);
            r_panelView.repaint();
            l_panelView.repaint();
            main_panelView.repaint();
            if (currHeight == BASE_HEIGHT) {
                baseButton.setEnabled(false);
                reductionButton.setEnabled(false);
            } else {
                expansionButton.setEnabled(true);
            }
        }
    }

    /*******************************************************
     *
     * 基準ボタン @@@
     *
     *******************************************************/
    class StanderdAction implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            currHeight = BASE_HEIGHT;
            currWidth  = BASE_WIDTH;
            r_panelView.setRPanelViewSize(currHeight);
            l_panelView.setLPanelViewSize(currWidth);
            main_panelView.setMainPanelViewSize(currWidth,currHeight);
            r_panelView.repaint();
            l_panelView.repaint();
            main_panelView.repaint();
            baseButton.setEnabled(false);
            reductionButton.setEnabled(false);
            expansionButton.setEnabled(true);
        }
    }

    /*******************************************************
     *
     * 拡大ボタン @@@
     *
     *******************************************************/
    class ExpansionAction implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            currHeight = currHeight + INC_HEIGHT;
            currWidth  = currWidth + INC_WIDTH;
            r_panelView.setRPanelViewSize(currHeight);
            l_panelView.setLPanelViewSize(currWidth);
            main_panelView.setMainPanelViewSize(currWidth,currHeight);
            r_panelView.repaint();
            l_panelView.repaint();
            main_panelView.repaint();
            baseButton.setEnabled(true);
            reductionButton.setEnabled(true);
            if (currHeight >= (BASE_HEIGHT * MAGNIFICATION)) {
                expansionButton.setEnabled(false);
            }
        }
    }


    /*
     *
     *   設定値：制御テーブル
     *
     */
    class CtOldTable extends JTable {

        private CtOldTblRenderer render1 = null;
        private CtOldTblRenderer render2 = null;
        private CtOldTblMdl model = null;
        

        CtOldTable(){
            super();

            try{
                setName("CtTable");
                setBounds(0, 0, 200, 200);
                setAutoCreateColumnsFromModel(true);
                setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION );

                setLocale(new Locale("ja","JP"));
                setFont(new java.awt.Font("dialog", 0, 12));
                setRowHeight(17);

                model = new CtOldTblMdl();
                setModel(model);

                DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                TableColumn  colum = null;

                // 項目No
                colum = cmdl.getColumn(0);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // Ｌ軸
                colum = cmdl.getColumn(1);
                colum.setMaxWidth(60);
                colum.setMinWidth(60);
                colum.setWidth(60);
                render1 = new CtOldTblRenderer();
                colum.setCellRenderer(render1);

                // Ｒ軸
                colum = cmdl.getColumn(2);
                colum.setMaxWidth(60);
                colum.setMinWidth(60);
                colum.setWidth(60);
                render2 = new CtOldTblRenderer();
                colum.setCellRenderer(render2);
            }
            catch (Throwable e) {
                CZSystem.handleException(e);
            }
        }
		
		public CtOldTblRenderer getRender1(){
			return render1;
		}

		public CtOldTblRenderer getRender2(){
			return render2;
		}
		
        //
        //
        //
        public void valueChanged(ListSelectionEvent e){
            super.valueChanged(e);

//@@            CZSystem.log("CZControlTableSub CtTable",
//@@                    "valueChanged [" + getSelectedRow() + "][" + getSelectedColumn() + "]");

            if(isShowing()){
                reSetData2();
                repaint();
                main_panelView.repaint();       //@@@
            }
        }


        //
        //
        //
        public void setData(int gr,int tbl){
//@@            CZSystem.log("CZControlTableSub CtTable","setData [" + gr + "][" + tbl + "]");

        }

        /*
         *
         *       設定値：制御テーブル：モデル
         *
         */
        public class CtOldTblMdl extends AbstractTableModel {

            private int TBL_ROW = REC_MAX;
            final   int TBL_COL = 3;

            final String[] names = {" # " , "Ｌ軸" , "Ｒ軸"};

            private Object  data[][];

            CtOldTblMdl(){
                super();

                data = new Object[TBL_ROW][TBL_COL];

                String empty   = "";

                for(int i = 0 ; i < TBL_ROW ; i++){
                    data[i][0] = new Integer(i+1);
                    data[i][1] = empty;
                    data[i][2] = empty;
                }
            }


            public int getColumnCount(){
                return TBL_COL;
            }

            public int getRowCount(){
                return TBL_ROW;
            }

            public Object getValueAt(int row, int col){
                return data[row][col];
            }

            public String getColumnName(int column){
                return names[column];
            }

            public Class getColumnClass(int c){
                return getValueAt(0, c).getClass();
            }

            public boolean isCellEditable(int row, int col){
                return false;
            }

            public void setValueAt(Object aValue, int row, int column){
                data[row][column] = aValue;
            }

        } // CtOldTblMdl

        /***************************************************
         *
         *   設定値：制御テーブル：レンダー
         *
         ***************************************************/
        class CtOldTblRenderer extends DefaultTableCellRenderer {

			private Color fColor = java.awt.Color.blue;
			
            CtOldTblRenderer(){
                super();
                setLocale(new Locale("ja","JP"));
                setFont(new java.awt.Font("dialog", 0, 12));
                setHorizontalAlignment(RIGHT);
            }

            public Component getTableCellRendererComponent( JTable table,
                                                        Object value,
                                                        boolean isSelected,
                                                        boolean hasFocus,
                                                        int row,int column){

                if(0 == column){
                    super.getTableCellRendererComponent(table,
                                                        value,
                                                        isSelected,
                                                        hasFocus,
                                                        row,column);
                    return(this);
                }

                if(null == value){
                    super.getTableCellRendererComponent(table,
                                                        value,
                                                        isSelected,
                                                        hasFocus,
                                                        row,column);
                    return(this);
                }

                // 表示フォーマット
                DecimalFormat format = null;
                StringBuffer  buff = new StringBuffer();
                int keta = 3;

                if(1 == column){
                    keta = edit_name.l_keta;
                }
                else if(2 == column){
                    keta = edit_name.r_keta;
                }

                if(1 > keta){
                    format = new DecimalFormat("0");
                }
                else {
                    buff.append("0.");
                    for(int i = 0 ; i < keta ; i++){
                        buff.append("0");
                    }
                    format = new DecimalFormat(buff.toString());
                }

                Float new_val = new Float(format.format(value));

                super.getTableCellRendererComponent(table,
                                                    format.format(new_val.floatValue()),
                                                    isSelected,
                                                    hasFocus,
                                                    row,column);

                table.setValueAt(new_val,row,column);

                float min = 0.0f;
                float max = 0.0f;
                if(1 == column){
                    min = edit_name.l_min;
                    max = edit_name.l_max;
                }
                else if(2 == column){
                    min = edit_name.r_min;
                    max = edit_name.r_max;
                }

                if((min <= new_val.floatValue()) &&
                           (max >= new_val.floatValue())){
//@@@@@                    setForeground(java.awt.Color.blue);
                    setForeground(fColor);
                }
                else {
                    setForeground(java.awt.Color.red);
                }

                return(this);
                }

				//
				// 色変更
				//
				public void setColor(Color col){

					fColor = col;
				}

            } // CtOldTblRenderer
        } // CtOldTable

        /***************************************************
         *
         *   変更値：制御テーブル
         *
         ***************************************************/
        class CtTable extends JTable {

            private CtTblMdl model = null;

            CtTable(){
                super();

                try{
                    setName("CtTable");
                    setBounds(0, 0, 200, 200);
                    setAutoCreateColumnsFromModel(true);
                    setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION );

                    setLocale(new Locale("ja","JP"));
                    setFont(new java.awt.Font("dialog", 0, 12));
                    setRowHeight(17);

                    model = new CtTblMdl();
                    setModel(model);

                    DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                    TableColumn  colum = null;

                    // 項目No
                    colum = cmdl.getColumn(0);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);

                    // Ｌ軸
                    colum = cmdl.getColumn(1);
                    colum.setMaxWidth(60);
                    colum.setMinWidth(60);
                    colum.setWidth(60);
                    colum.setCellRenderer(new CtTblRenderer());

                    // Ｒ軸
                    colum = cmdl.getColumn(2);
                    colum.setMaxWidth(60);
                    colum.setMinWidth(60);
                    colum.setWidth(60);
                    colum.setCellRenderer(new CtTblRenderer());

                }
                catch (Throwable e) {
                    CZSystem.handleException(e);
                }
            }

            //
            //
            //
            public void valueChanged(ListSelectionEvent e){
                super.valueChanged(e);

//@@                CZSystem.log("CZControlTableSub CtTable",
//@@                    "valueChanged [" + getSelectedRow() + "][" + getSelectedColumn() + "]");

                if(isShowing()){
                    reSetData2();
                    repaint();
                    main_panelView.repaint();       //@@@
                }
            }


            //
            //
            //
            public void setData(int gr,int tbl){
//@@                CZSystem.log("CZControlTableSub CtTable","setData [" + gr + "][" + tbl + "]");

            }


            /***********************************************
             *
             *       変更値：制御テーブル：モデル
             *
             ***********************************************/
            public class CtTblMdl extends AbstractTableModel {

                private int TBL_ROW = REC_MAX;
                final   int TBL_COL = 3;

                final String[] names = {" # " , "Ｌ軸" , "Ｒ軸"};

                private Object  data[][];

                CtTblMdl(){
                    super();

                data = new Object[TBL_ROW][TBL_COL];

                String empty   = "";

                for(int i = 0 ; i < TBL_ROW ; i++){
                    data[i][0] = new Integer(i+1);
                    data[i][1] = empty;
                    data[i][2] = empty;
                }
            }

            public int getColumnCount(){
                return TBL_COL;
            }

            public int getRowCount(){
                return TBL_ROW;
            }

            public Object getValueAt(int row, int col){
                return data[row][col];
            }

            public String getColumnName(int column){
                return names[column];
            }

            public Class getColumnClass(int c){
                return getValueAt(0, c).getClass();
            }

            public boolean isCellEditable(int row, int col){

                if(1 == col) return true;
                if(2 == col) return true;
                return false;
            }

            public void setValueAt(Object aValue, int row, int column){
                data[row][column] = aValue;
            }
        } // CtTblMdl

        /***************************************************
         *
         *   変更値：制御テーブル：レンダー
         *
         ***************************************************/
        class CtTblRenderer extends DefaultTableCellRenderer {

            CtTblRenderer(){
                super();
                setLocale(new Locale("ja","JP"));
                setFont(new java.awt.Font("dialog", 0, 12));
                setHorizontalAlignment(RIGHT);
            }

            public Component getTableCellRendererComponent( JTable table,
                                                        Object value,
                                                        boolean isSelected,
                                                        boolean hasFocus,
                                                        int row,int column){

                if(0 == column){
                    super.getTableCellRendererComponent(table,
                                                    value,
                                                    isSelected,
                                                    hasFocus,
                                                    row,column);
                    return(this);
                }

                if(null == value){
                    super.getTableCellRendererComponent(table,
                                                    value,
                                                    isSelected,
                                                    hasFocus,
                                                    row,column);
                    return(this);
                }

                // 表示フォーマット
                DecimalFormat format = null;
                StringBuffer  buff = new StringBuffer();
                int keta = 3;

                if(1 == column){
                    keta = edit_name.l_keta;
                }
                else if(2 == column){
                    keta = edit_name.r_keta;
                }

                if(1 > keta){
                    format = new DecimalFormat("0");
                }
                else {
                    buff.append("0.");
                    for(int i = 0 ; i < keta ; i++){
                        buff.append("0");
                    }
                    format = new DecimalFormat(buff.toString());
                }

                Float new_val = new Float(format.format(value));

                super.getTableCellRendererComponent(table,
                                                    format.format(new_val.floatValue()),
                                                    isSelected,
                                                    hasFocus,
                                                    row,column);

                table.setValueAt(new_val,row,column);

                float min = 0.0f;
                float max = 0.0f;
                if(1 == column){
                    min = edit_name.l_min;
                    max = edit_name.l_max;
                }
                else if(2 == column){
                    min = edit_name.r_min;
                    max = edit_name.r_max;
                }

                if((min <= new_val.floatValue()) &&
                               (max >= new_val.floatValue())){
                    setForeground(java.awt.Color.blue);
                }
                else {
                    setForeground(java.awt.Color.red);
                }
                return(this);
            }

        } // CtTblRenderer
    } // CtTable

    /*******************************************************
     *  L値を昇順でソートしたものの場合、１（通常のもの）
     *  L値を降順でソートしたものの場合、２（残液等）
     *
     *   昇順でソート
     *
     *******************************************************/
    class Sort1 implements Comparator {
        public int compare(Object a, Object b) {

            CZSystemCtTb val1 = (CZSystemCtTb)a;
            CZSystemCtTb val2 = (CZSystemCtTb)b;

            if(val1.l_val >  val2.l_val)  return 1;
            if(val1.l_val <  val2.l_val)  return -1;

            return 0;
        }

        public boolean equals(Object a, Object b) {
            return compare(a, b) == 0;
        }
    }

    class Sort2 implements Comparator {
        public int compare(Object a, Object b) {

            CZSystemCtTb val1 = (CZSystemCtTb)a;
            CZSystemCtTb val2 = (CZSystemCtTb)b;

            if(val1.l_val >  val2.l_val)  return -1;
            if(val1.l_val <  val2.l_val)  return 1;

            return 0;
        }

        public boolean equals(Object a, Object b) {
            return compare(a, b) == 0;
        }
    }


    /*******************************************************
     *
     * グラフメインパネルのスクロール
     *
     *******************************************************/
    class MainPanel extends JScrollPane {

        // 画面
        private MainPanelView view_;

        MainPanel(){
            super();
        }

        public void setView(MainPanelView view){

            view_ = view;
            getViewport().setView( view_ );
            getViewport().setScrollMode( JViewport.BACKINGSTORE_SCROLL_MODE) ;
            setViewSize(BASE_WIDTH, BASE_HEIGHT);

        }

        /**
        * リスナーを追加する。
        */
        public void addComponentListener( ComponentListener listener ) {
            view_.addComponentListener( listener );
        }
        /**
        * 画面取得
        * @param
        * @return
        */
        public MainPanelView getView(){
            return view_;
        }

        /**
        * 画面位置取得
        * @param
        * @return
        */
        public Point getViewLocation(){
            return view_.getLocation();
        }

        /**
        * 画面位置設定
        * @param
        * @return
        */
        public void setViewLocation( Point pos ){
            view_.setLocation( pos );
        }
        /**
        * 画面幅設定
        * @param
        * @return
        */
        public void setViewSize( int width, int height){
            view_.setMainPanelViewSize( width, height );
        }

        /**
        * 描画
        */
        public void viewPaint(){
            view_.repaint();
        }
    } // MainPanel

    /*******************************************************
     *
     * グラフメインパネル
     *
     *******************************************************/
    class MainPanelView extends JPanel {

        /**
        * コンストラクタ
        */
        MainPanelView(){
            super();
            setLayout(null);
            setBackground(java.awt.Color.white);
        }

        /**
        * 画面のサイズを設定する。
        */
        public void setMainPanelViewSize( int width, int height ) {

            setPreferredSize( new Dimension( width, height ) );
            setSize( new Dimension( width, height) );
            repaint();
        }

        //
        //
        //
		@SuppressWarnings("unchecked")
        public void paint(Graphics g){

            Dimension d = getSize(null);
            g.setColor(java.awt.Color.black);
            g.fillRect(0,0,d.width,d.height);

            // グラフ目盛の描画
            g.setColor(java.awt.Color.darkGray);
            // Ｘ軸
            float bun = (float)d.width / (l_graph_bun * 5);
            float x  = 0;
            for(int j = 0 ; j < (l_graph_bun * 5)+1 ; j++){
                g.drawLine((int)x,0,(int)x,d.height);
                x+=bun;
            }

            // Ｙ軸
            bun = (float)d.height / (r_graph_bun * 5);
            float y   = 0.0f;
            for(int i = 0 ; i < (r_graph_bun * 5)+1 ; i++){
                g.drawLine(0,(int)y,d.width,(int)y);
                y+=bun;
            }

            g.setColor(java.awt.Color.lightGray);
            // Ｘ軸
            bun = (float)d.width / (float)l_graph_bun;
            x   = 0.0f;
            for(int i = 0 ; i < l_graph_bun+1 ; i++){
                g.drawLine((int)x,0,(int)x,d.height);
                x+=bun;
            }

            // Ｙ軸
            bun = (float)d.height / (float)r_graph_bun;
            y   = 0.0f;
            for(int i = 0 ; i < r_graph_bun+1 ; i++){
                g.drawLine(0,(int)y,d.width,(int)y);
                y+=bun;
            }

            if(null == current_data) return;

            int   size = current_data.size();
            if(2 > size) return;

            //グラフ元データのソート
            CZSystemCtTb data[] = new CZSystemCtTb[size];
            for(int i = 0 ; i < size ; i++){
                data[i] = (CZSystemCtTb)current_data.elementAt(i);
            }

            if(KOUJYUNN_SORT == edit_name.k_sort){
                Arrays.sort(data, new Sort2());
            }
            else {
                Arrays.sort(data, new Sort1());
            }

            //グラフの描画
            float old_val[][] = new float[size][2];
            for(int i = 0 ; i < size ; i++){
                old_val[i][0] = data[i].l_val;
                old_val[i][1] = data[i].r_val;

//@@                CZSystem.log("CZControlTableSub MainPanel",
//@@                    "paint [" + size + "] L[" + old_val[i][0] + "] R[" + old_val[i][1] + "]");
            }

            int new_val_x[] = new int[size];
            int new_val_y[] = new int[size];

            float y1 = (float)d.height / (edit_name.r_max - edit_name.r_min);
            float x1 = (float)d.width  / (edit_name.l_max - edit_name.l_min);

            for(int i = 0 ; i < size ; i++){
                float y2 = y1 * (old_val[i][1] - edit_name.r_min) ;
                new_val_y[i] = (int)((float)d.height - y2);

                float x2 = x1 * (old_val[i][0] - edit_name.l_min) ;
                if(KOUJYUNN_SORT == edit_name.k_sort){
                    new_val_x[i] = (int)((float)d.height - x2);
                }
                else {
                    new_val_x[i] = (int)x2;
                }
            }

			if(edit_current == true){
            g.setColor(OLD_PRO_COL);
            g.drawPolyline(new_val_x,new_val_y,size);

	            repaint_pv_data(g,d);
	            repaint_new_data(g,d);
			}
/*******************************************/
			if(false == mst_show) return;				//@@@@
            if(null == master_data) return;

            int   msize = master_data.size();
            if(2 > msize) return;

            //グラフ元データのソート
            CZSystemCtTb mdata[] = new CZSystemCtTb[msize];
            for(int i = 0 ; i < msize ; i++){
                mdata[i] = (CZSystemCtTb)master_data.elementAt(i);
            }

            if(KOUJYUNN_SORT == edit_name.k_sort){
                Arrays.sort(mdata, new Sort2());
            }
            else {
                Arrays.sort(mdata, new Sort1());
            }
			repaint_mstnew_data(g,d);
            repaint_mst_data(g,d);
        }


        //
        // 設定データの描画
        //
		@SuppressWarnings("unchecked")
        private void repaint_new_data(Graphics g,Dimension d){

            Vector new_dat = new Vector();
            for(int i = 0 ; i < REC_MAX ; i++){
                Float l = (Float)c_table.getValueAt(i,1);
                Float r = (Float)c_table.getValueAt(i,2);

                if(null == l) continue;
                if(null == r) continue;

                CZSystemCtTb tmp = new CZSystemCtTb();
                tmp.l_val = l.floatValue();
                tmp.r_val = r.floatValue();
                new_dat.addElement(tmp);
            }

            int size = new_dat.size();
            //グラフ元データのソート
            CZSystemCtTb data[] = new CZSystemCtTb[size];
            for(int i = 0 ; i < size ; i++){
                data[i] = (CZSystemCtTb)new_dat.elementAt(i);
            }

            if(KOUJYUNN_SORT == edit_name.k_sort){
                Arrays.sort(data, new Sort2());
            }
            else {
                Arrays.sort(data, new Sort1());
            }

            //グラフの描画
            float old_val[][] = new float[size][2];
            for(int i = 0 ; i < size ; i++){
                old_val[i][0] = data[i].l_val;
                old_val[i][1] = data[i].r_val;

//@@                CZSystem.log("CZControlTableSub MainPanel",
//@@                        "repaint_new_data [" + size + "] L[" + old_val[i][0] +
//@@                        "] R[" + old_val[i][1] + "]");
            }

            int new_val_x[] = new int[size];
            int new_val_y[] = new int[size];

            float y1 = (float)d.height / (edit_name.r_max - edit_name.r_min);
            float x1 = (float)d.width  / (edit_name.l_max - edit_name.l_min);

            for(int i = 0 ; i < size ; i++){
                float y2 = y1 * (old_val[i][1] - edit_name.r_min) ;
                new_val_y[i] = (int)((float)d.height - y2);

                float x2 = x1 * (old_val[i][0] - edit_name.l_min) ;
                if(KOUJYUNN_SORT == edit_name.k_sort){
                    new_val_x[i] = (int)((float)d.height - x2);
                }
                else {
                    new_val_x[i] = (int)x2;
                }
            }

            Graphics2D g2 = (Graphics2D)g;
//2003.11.27newの線の太さをみんなとあわすためコメント
            //現在の設定をst_tmpに格納
//2003.11.27            Stroke st_tmp = g2.getStroke();
			//ラインの太さをピクセルで設定（2fで通常ラインの倍の太さ）
//2003.11.27            BasicStroke bs = new BasicStroke(2f);  // 10ピクセル幅
			//ラインの太さを変更した設定をセットする
//2003.11.27            g2.setStroke(bs);

			//ラインの色を変更する
            g2.setColor(NEW_PRO_COL);
			//ラインを表示（作画）
            g2.drawPolyline(new_val_x,new_val_y,size);
			//元の設定に戻す
//2003.11.27          g2.setStroke(st_tmp);
        }

/**********************************************/
		@SuppressWarnings("unchecked")
        private void repaint_mstnew_data(Graphics g,Dimension d){

            Vector new_dat = new Vector();
            for(int i = 0 ; i < REC_MAX ; i++){
                Float l = (Float)c_table.getValueAt(i,1);
                Float r = (Float)c_table.getValueAt(i,2);

                if(null == l) continue;
                if(null == r) continue;

                CZSystemCtTb tmp = new CZSystemCtTb();
                tmp.l_val = l.floatValue();
                tmp.r_val = r.floatValue();
                new_dat.addElement(tmp);
            }

            int size = new_dat.size();
            //グラフ元データのソート
            CZSystemCtTb data[] = new CZSystemCtTb[size];
            for(int i = 0 ; i < size ; i++){
                data[i] = (CZSystemCtTb)new_dat.elementAt(i);
            }

            if(KOUJYUNN_SORT == edit_name.k_sort){
                Arrays.sort(data, new Sort2());
            }
            else {
                Arrays.sort(data, new Sort1());
            }

            //グラフの描画
            float old_val[][] = new float[size][2];
            for(int i = 0 ; i < size ; i++){
                old_val[i][0] = data[i].l_val;
                old_val[i][1] = data[i].r_val;

//@@                CZSystem.log("CZControlTableSub MainPanel",
//@@                        "repaint_new_data [" + size + "] L[" + old_val[i][0] +
//@@                        "] R[" + old_val[i][1] + "]");
            }

            int new_val_x[] = new int[size];
            int new_val_y[] = new int[size];

            float y1 = (float)d.height / (edit_name.r_max - edit_name.r_min);
            float x1 = (float)d.width  / (edit_name.l_max - edit_name.l_min);

            for(int i = 0 ; i < size ; i++){
                float y2 = y1 * (old_val[i][1] - edit_name.r_min) ;
                new_val_y[i] = (int)((float)d.height - y2);

                float x2 = x1 * (old_val[i][0] - edit_name.l_min) ;
                if(KOUJYUNN_SORT == edit_name.k_sort){
                    new_val_x[i] = (int)((float)d.height - x2);
                }
                else {
                    new_val_x[i] = (int)x2;
                }
            }

            Graphics2D g2 = (Graphics2D)g;
//2003.11.27newの線の太さをみんなとあわすためコメント
            //現在の設定をst_tmpに格納
//2003.11.27            Stroke st_tmp = g2.getStroke();
			//ラインの太さをピクセルで設定（2fで通常ラインの倍の太さ）
//2003.11.27            BasicStroke bs = new BasicStroke(2f);  // 10ピクセル幅
			//ラインの太さを変更した設定をセットする
//2003.11.27            g2.setStroke(bs);

			//ラインの色を変更する
            g2.setColor(java.awt.Color.cyan);
			//ラインを表示（作画）
            g2.drawPolyline(new_val_x,new_val_y,size);
			//元の設定に戻す
//2003.11.27          g2.setStroke(st_tmp);
        }
/**********************************************/


		@SuppressWarnings("unchecked")
        private void repaint_mst_data(Graphics g,Dimension d){

            int size = master_data.size();
            //グラフ元データのソート
            CZSystemCtTb data[] = new CZSystemCtTb[size];
            for(int i = 0 ; i < size ; i++){
                data[i] = (CZSystemCtTb)master_data.elementAt(i);
            }

            if(KOUJYUNN_SORT == edit_name.k_sort){
                Arrays.sort(data, new Sort2());
            }
            else {
                Arrays.sort(data, new Sort1());
            }

            //グラフの描画
            float old_val[][] = new float[size][2];
            for(int i = 0 ; i < size ; i++){
                old_val[i][0] = data[i].l_val;
                old_val[i][1] = data[i].r_val;
            }

            int new_val_x[] = new int[size];
            int new_val_y[] = new int[size];

            float y1 = (float)d.height / (edit_name.r_max - edit_name.r_min);
            float x1 = (float)d.width  / (edit_name.l_max - edit_name.l_min);

            for(int i = 0 ; i < size ; i++){
                float y2 = y1 * (old_val[i][1] - edit_name.r_min) ;
                new_val_y[i] = (int)((float)d.height - y2);

                float x2 = x1 * (old_val[i][0] - edit_name.l_min) ;
                if(KOUJYUNN_SORT == edit_name.k_sort){
                    new_val_x[i] = (int)((float)d.height - x2);
                }
                else {
                    new_val_x[i] = (int)x2;
                }
            }

            Graphics2D g2 = (Graphics2D)g;

			//ラインの色を変更する
			g2.setColor(java.awt.Color.yellow);
			//ラインを表示（作画）
			g2.drawPolyline(new_val_x,new_val_y,size);
			//元の設定に戻す
        }


        //
        // ＰＶデータの描画
        //
        private void repaint_pv_data(Graphics g,Dimension d){
//@@            CZSystem.log("CZControlTableSub repaint_pv_data","START 1");

            if(null == pv_data_shld) return;
            if(null == pv_data_body) return;

//@@            CZSystem.log("CZControlTableSub repaint_pv_data","START 2");
            switch(edit_group){
                case T1:
//@@                    CZSystem.log("CZControlTableSub repaint_pv_data","START 3 T1");
                    break;

                case T2:
//@@                    CZSystem.log("CZControlTableSub repaint_pv_data","START 3 T2");
                    repaint_pv_data_t2(g,d);
                    break;

                case T3:
//@@                    CZSystem.log("CZControlTableSub repaint_pv_data","START 3 T3");
                    break;

                case T4:
//@@                    CZSystem.log("CZControlTableSub repaint_pv_data","START 3 T4");
                    break;

                case T5:
//@@                    CZSystem.log("CZControlTableSub repaint_pv_data","START 3 T5");
                    break;
//@@
                case T6:
//@@                    CZSystem.log("CZControlTableSub repaint_pv_data","START 3 T6");
                    break;

                default:
//@@                    CZSystem.log("CZControlTableSub repaint_pv_data","START 3 Default");
                    return;
            }

//@@            CZSystem.log("CZControlTableSub repaint_pv_data","END ");
            return;
        }

        /**
        * ＰＶデータの描画 引き上げプロファイル
        */
        private void repaint_pv_data_t2(Graphics g,Dimension d){
//@@            CZSystem.log("CZControlTableSub repaint_pv_data_t2","START");

            switch(edit_number){
                case DIA_BODY_PF:
                        repaint_pv_data_body_pf(g,d,DIA,DIA_PF,false);
                        break;

                case SXL_ST_BODY_PF:
                        repaint_pv_data_body_pf(g,d,SXL_ST,SXL_ST_PF,false);
                        break;

                case HT_BODY_PF:
                        repaint_pv_data_body_pf(g,d,MAIN1_H_T,MAIN1_H_T_PF,true);
                        break;

                default:
                    return;
            }
            return;
        }


        //
        // ＰＶデータの描画 ボディーのプロファイルの場合
        //
        private void repaint_pv_data_body_pf(Graphics g,Dimension d,int val_no,int pf_no,boolean shift_flg){

//@@            CZSystem.log("CZControlTableSub repaint_pv_data_body_pf",
//@@                    "GROUP[" + edit_group + "] NUMBER[" + edit_number + "]");

            int size = pv_data_body.size();
            if(1 > size) return;

            //グラフの描画
            CZSystemPVData data;

            int new_val_x[] = new int[size];
            int new_prf_y[] = new int[size];
            int new_val_y[] = new int[size];

            float y1  = (float)d.height / (edit_name.r_max - edit_name.r_min);
            float x1  = (float)d.width  / (edit_name.l_max - edit_name.l_min);
            float y2;

            float off = 0.0f;

            if(shift_flg){
                data = (CZSystemPVData)pv_data_body.elementAt(0);
                off = data.data[val_no];
            }

            for(int i = 0 ; i < size ; i++){
                data = (CZSystemPVData)pv_data_body.elementAt(i);

                y2 = y1 * (data.data[pf_no] - edit_name.r_min) ;
                new_prf_y[i] = (int)((float)d.height - y2);

                y2 = y1 * (data.data[val_no] - off - edit_name.r_min) ;
                new_val_y[i] = (int)((float)d.height - y2);

                new_val_x[i] = (int)(x1 * (data.p_length - edit_name.l_min)) ;
            }

            // 実績プロファイル
            g.setColor(VAL_PRO_COL);
            g.drawPolyline(new_val_x,new_prf_y,size);

            // 実績
            g.setColor(VAL_COL);
            g.drawPolyline(new_val_x,new_val_y,size);
        }
    } // MainPanelView

    /*******************************************************
     *
     * Y軸の目盛パネルのスクロール
     *
     *******************************************************/
    class RPanel extends JScrollPane {

        // 画面
        private RPanelView view_;

        /**
        * コンストラクタ
        */
        RPanel(){
            super();
        }

        public void setView(RPanelView view){

            view_ = view;
            getViewport().setView( view_ );
            getViewport().setScrollMode( JViewport.BACKINGSTORE_SCROLL_MODE) ;
            setViewSize(BASE_HEIGHT);

        }
        /**
        * リスナーを追加する。
        */
        public void addComponentListener( ComponentListener listener ) {
            view_.addComponentListener( listener );
        }
        /**
        * 画面取得
        * @param
        * @return
        */
        public RPanelView getView(){
            return view_;
        }

        /**
        * 画面位置取得
        * @param
        * @return
        */
        public Point getViewLocation(){
            return view_.getLocation();
        }

        /**
        * 画面位置設定
        * @param
        * @return
        */
        public void setViewLocation( Point pos ){
            view_.setLocation( pos );
        }
        /**
        * 画面幅設定
        * @param
        * @return
        */
        public void setViewSize( int height){
            view_.setRPanelViewSize( height );
        }
    } // RPanel

    /*******************************************************
     *
     * Y軸の目盛パネル
     *
     *******************************************************/
    class RPanelView extends JPanel {

        /**
        * コンストラクタ
        */
        RPanelView(){
            super();
            setFont(new java.awt.Font("dialog", 0, 12));
        }
        /**
        * 画面のサイズを設定する。
        */
        public void setRPanelViewSize( int height ) {
            setPreferredSize( new Dimension( 50, height ) );
            setSize( new Dimension( 50, height) );
            repaint();
        }

        /**
        * 描画
        */
        public void paint(Graphics g){

            Dimension d = getSize(null);
            g.setColor(java.awt.Color.black);
            g.fillRect(0,0,d.width,d.height);

            // グラフ目盛の描画
            g.setColor(java.awt.Color.darkGray);

            float bun = (float)d.height / (float)r_graph_bun;
            float y   = 0;
            for(int i = 0 ; i < r_graph_bun+1 ; i++){
                g.drawLine(0,(int)y,d.width,(int)y);
                y+=bun;
            }


            //表示フォーマット
            DecimalFormat format = null;
            StringBuffer  buff = new StringBuffer();

            if(1 > edit_name.r_keta){
                format = new DecimalFormat("0");
            }
            else {
                buff.append("0.");
                for(int i = 0 ; i < edit_name.r_keta ; i++){
                    buff.append("0");
                }
                format = new DecimalFormat(buff.toString());
            }


            g.setColor(java.awt.Color.lightGray);
            float sp  = (edit_name.r_max - edit_name.r_min) / r_graph_bun ;
            float val = 0.0f;
            y = 0;
            for(int i = 0 ; i < r_graph_bun+1 ; i++){
                val = edit_name.r_max - ((float)i * sp);

                String s = new String(format.format(val));
                g.drawString(s,2,(int)y-2);
                y+=bun;
            }
        }
    } // RPanelView


    /*******************************************************
     *
     * X軸の目盛パネルのスクロール
     *
     *******************************************************/
    class LPanel extends JScrollPane {

        // 画面
        private LPanelView view_;

        /**
        * コンストラクタ
        */
        LPanel(){
            super();
        }

        public void setView(LPanelView view){

            view_ = view;
            getViewport().setView( view_ );
            getViewport().setScrollMode( JViewport.BACKINGSTORE_SCROLL_MODE) ;
            setViewSize(BASE_WIDTH);

        }

        /**
        * リスナーを追加する。
        */
        public void addComponentListener( ComponentListener listener ) {
            view_.addComponentListener( listener );
        }
        /**
        * 画面取得
        * @param
        * @return
        */
        public LPanelView getView(){
            return view_;
        }

        /**
        * 画面位置取得
        * @param
        * @return
        */
        public Point getViewLocation(){
            return view_.getLocation();
        }

        /**
        * 画面位置設定
        * @param
        * @return
        */
        public void setViewLocation( Point pos ){
            view_.setLocation( pos );
        }
        /**
        * 画面幅設定
        * @param
        * @return
        */
        public void setViewSize( int width){
            view_.setLPanelViewSize( width );
        }

    } // LPanel

    /*******************************************************
     *
     * X軸の目盛パネル
     *
     *******************************************************/
    class LPanelView extends JPanel {

        /**
        * コンストラクタ
        */
        LPanelView(){
            super();
            setFont(new java.awt.Font("dialog", 0, 12));
        }

        /**
        * 画面のサイズを設定する。
        */
        public void setLPanelViewSize( int width ) {
            setPreferredSize( new Dimension( width, 50 ) );
            setSize( new Dimension( width, 50) );
            repaint();
        }

        /**
        * 描画
        */
        public void paint(Graphics g){

            Dimension d = getSize(null);
            g.setColor(java.awt.Color.black);
            g.fillRect(0,0,d.width,d.height);

            // グラフ目盛の描画
            g.setColor(java.awt.Color.darkGray);
            float bun = (float)d.width / (float)l_graph_bun;
            float x   = 0;
            for(int i = 0 ; i < l_graph_bun ; i++){
                g.drawLine((int)x,0,(int)x,d.height);
                x+=bun;
            }

            //表示フォーマット
            DecimalFormat format = null;
            StringBuffer  buff = new StringBuffer();

            if(1 > edit_name.l_keta){
                format = new DecimalFormat("0");
            }
            else {
                buff.append("0.");
                for(int i = 0 ; i < edit_name.l_keta ; i++){
                    buff.append("0");
                }
                format = new DecimalFormat(buff.toString());
            }

            g.setColor(java.awt.Color.lightGray);
            float sp  = (edit_name.l_max - edit_name.l_min) / l_graph_bun ;
            float val = 0.0f;
            x = 0;
            for(int i = 0 ; i < l_graph_bun+1 ; i++){
                if(KOUJYUNN_SORT == edit_name.k_sort){
                    val = edit_name.l_max - ((float)i * sp);
                }
                else {
                    val = edit_name.l_min + ((float)i * sp);
                }

                String s = new String(format.format(val));
                g.drawString(s,(int)x+2,16);

                x+=bun;

                if(10 < l_graph_bun){
                    i++;
                    x+=bun;
                }
            } // for end
        }
    } // LPanelView

    /***************************************************************************
     *
     *       設定者を入力するTextField
     *
     ***************************************************************************/
    /*public*/ class TText extends JTextField {

        /**
        * コンストラクタ
        */
        TText(){
            super();
            setFont(new java.awt.Font("dialog", 0, 16));
        }

        //
        //
        //
        protected Document createDefaultModel() {
            return new NumericDocument();
        }

        //
        //
        //
        class NumericDocument extends PlainDocument {
            String validValues = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-/.:";

            //
            //
            public void insertString( int offset, String str, AttributeSet a )
                                            throws BadLocationException {
                if(29 < getLength()) return;
                char[] val = str.toCharArray();
                for (int i = 0; i < val.length; i++) {
                    if(validValues.indexOf(val[i]) == -1) return;
                }
                super.insertString( offset, str, a );
            }

            //
            //
            public void remove(int offs, int len)
                                            throws BadLocationException {
                super.remove(offs,len);
            }
        }
    }

    /***************************************************************************
     *
     *       設定値シフト量を入力するTextField
     *
     ***************************************************************************/
    /*public*/ class ShiftText extends JTextField {

        /**
        * コンストラクタ
        */
        ShiftText(){
            super();
            setFont(new java.awt.Font("dialog", 0, 16));
        }

        //
        //
        //
        protected Document createDefaultModel() {
            return new NumericDocument();
        }

        //
        //
        //
        class NumericDocument extends PlainDocument {
            String validValues = "0123456789.";

            //
            //
            public void insertString( int offset, String str, AttributeSet a )
                                        throws BadLocationException {
                if(4 < getLength()) return;
                char[] val = str.toCharArray();
                for (int i = 0; i < val.length; i++) {
                    if(validValues.indexOf(val[i]) == -1) return;
                }
                super.insertString( offset, str, a );
            }

            //
            //
            public void remove(int offs, int len)
                                    throws BadLocationException {
                super.remove(offs,len);
            }
        }
    }

    /***************************************************************************
     *
     *       グラフ分割数を入力するTextField
     *
     ***************************************************************************/
    /*public*/ class BunText extends JTextField {

        /**
        *
        */
        BunText(){
            super();
            setFont(new java.awt.Font("dialog", 0, 16));
        }
        //
        //
        //
        protected Document createDefaultModel() {
            return new NumericDocument();
        }

        //
        //
        //
        class NumericDocument extends PlainDocument {
            String validValues = "0123456789";

            //
            //
            public void insertString( int offset, String str, AttributeSet a )
                                                    throws BadLocationException {

                if(1 < getLength()) return;
                char[] val = str.toCharArray();
                for (int i = 0; i < val.length; i++) {
                    if(validValues.indexOf(val[i]) == -1) return;
                }
                super.insertString( offset, str, a );
            }

            //
            //
            public void remove(int offs, int len)
                                                throws BadLocationException {
                super.remove(offs,len);
            }
        }
    }

    /*******************************************************
     *
     * グラフのスクロールコンポーネントリスナー @@@
     *
     *******************************************************/
    private class GraphComponentListener implements ComponentListener {
        /**
        *　コンストラクタ
        */
        GraphComponentListener(){
            super();
        }

        /**
        * スクロールが移動した時の処理
        */
        public void componentMoved( java.awt.event.ComponentEvent e )
        {
            if( main_panelView == e.getComponent() ) {
                // メイン画面が移動したときはY軸目盛を移動する
                Point mainViewPos = main_panelView.getLocation();
                Point yViewPos = r_panelView.getLocation();
                yViewPos.y = mainViewPos.y;
                r_panelView.setLocation( yViewPos );
            }
            else if( l_panelView == e.getComponent() ) {
                // X軸画面が移動した時はメイン画面を移動する。
                Point xViewPos = l_panelView.getLocation();
                Point mainViewPos = main_panelView.getLocation();
                mainViewPos.x = xViewPos.x;
                main_panelView.setLocation( mainViewPos );
            }
        }

        /**
        *
        */
        public void componentResized(java.awt.event.ComponentEvent e){
        }
        /**
        *
        */
        public void componentShown(java.awt.event.ComponentEvent e){
        }

        /**
        *
        */
        public void componentHidden(java.awt.event.ComponentEvent e){
        }
    } //GraphComponentListener
}
