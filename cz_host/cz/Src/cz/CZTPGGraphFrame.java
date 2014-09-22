package cz;

import java.awt.Color;
import java.awt.Container;
import java.awt.Cursor;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Graphics;
import java.awt.Rectangle;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ComponentListener;
import java.awt.event.FocusAdapter;
import java.awt.event.FocusEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Hashtable;
import java.util.Locale;
import java.util.Properties;
import java.util.StringTokenizer;
import java.util.Vector;

import javax.swing.BorderFactory;
import javax.swing.JButton;
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
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.border.Border;
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

import java.text.DecimalFormat;

//==========================================================================
/**
 *   グラフ表示用ダイアログ
 * @Update 2013.10.28 TPGｸﾞﾗﾌ入力最大数変更 (@20131028)
 */
//==========================================================================
public class CZTPGGraphFrame extends JFrame
{

    private String roName_              = null; //対象炉番
    private String roDbName_            = null; //対象炉データベース名

    private int    SelectNo             = 0;
    private String SelectBt             = null;
    private String SelectTime           = null;
    private Vector selList_            = null; //Y軸項目
    private Vector pvDataBody_          = null; //PVデータ
    private CZSystemStart roBtStart_    = null; //検索用引き上げ条件

    private Vector roBtTempCondition_    = null; //選択Btの引き上げ条件
    private Vector roBtAllCondition_    = null; //全Btの引き上げ条件

    //色を定義する。(背景 + 目盛1 + 目盛2 + 目盛3)

    private final Color DEFAULT_BACKGROUND_COL= CZSystemDefine.DEFAULT_BACKGROUND_COL;
    private final Color BACK_COL            = java.awt.Color.black;
    private final Color MEM_LINE1_COL       = java.awt.Color.lightGray;
    private final Color MEM_LINE2_COL       = java.awt.Color.darkGray;
    private final Color MEM_LINE3_COL       = java.awt.Color.darkGray;

    private final String GR_X_LENGTH_DEF    = "10000";   //Ｘ軸の長さ		// @20131028 TPGｸﾞﾗﾌ入力最大数変更

    private int Y_VIEW_TIMES    = 1;                    //Ｙ軸の倍数

    private String  grXlength_  = GR_X_LENGTH_DEF;      //Ｘ軸の長さ
    private int     grXUnit_    = 2;                    //Ｘ軸単位(min/mm)
    private float   grXMin_     = 0.0f;                 //Ｘ軸最小値
    private float   grXMax_     = 10000f;                //Ｘ軸最大値		// @20131028 TPGｸﾞﾗﾌ入力最大数変更
    private float   grXbun_     = 20.0f;                //Ｘ軸の分割数
    private float   grYbun_     = 5.0f;                 //Ｙ軸の分割数

    private GraphTitlePanel  titlePnl_  = null; //パネル
    private MainSc  mainSc_             = null; //メイングラフスクロールパネル
    private XSc     xSc_                = null; //Ｘ軸グラフスクロールパネル
    private Y1Sc    y1Sc_               = null; //Ｙ軸左側グラフスクロールパネル
    private Y2Sc    y2Sc_               = null; //Ｙ軸右側グラフスクロールパネル

    /**
    * コンストラクタ
    */
    public CZTPGGraphFrame(int sNo, String roDBName, String SelBt, String SelTime, Vector pvData_, Vector selLst, CZSystemStart st){
        super();

        setTitle("TPG Graph");                      //画面Titleを設定する。
        setSize(1152,864);                          //画面サイズを設定する。
        setResizable(false);                        //画面のサイズ変更は不可とする。
//        setModal(true);                             //Modalで表示する。
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

		roDbName_ = roDBName;
		SelectNo = sNo;
		SelectBt = SelBt;
		SelectTime = SelTime;
		selList_ = selLst;
		pvDataBody_ = pvData_;
		roBtStart_ = st;

        titlePnl_ = new GraphTitlePanel();
        titlePnl_.setLayout(null);
        titlePnl_.setBounds(50,0,1100,120);
        getContentPane().add(titlePnl_);

        // グラフ表示領域
        PVGrEventCompo comp = new PVGrEventCompo();

        mainSc_ = new MainSc(comp);                 // メイングラフのパネル
        mainSc_.setBounds(130, 120, 870, 620);
        mainSc_.setDefault();
        getContentPane().add(mainSc_);

        xSc_    = new XSc(comp);                    // X軸の目盛のパネル
        xSc_.setBounds(130, 670+70, 870, 40);
        xSc_.setDefault();
        getContentPane().add(xSc_);

        y1Sc_   = new Y1Sc(comp);                   // Y軸の左側のパネル
        y1Sc_.setBounds(20, 120, 120, 620);
        y1Sc_.setDefault();
        getContentPane().add(y1Sc_);

        y2Sc_   = new Y2Sc();                       // Y軸の右側のパネル
        y2Sc_.setBounds(870+130, 120, 120, 620);
        getContentPane().add(y2Sc_);

    }

    /**
    * グラフデータを設定する。
    */
    public void setData(){
        CZSystem.log("CZTPGGraphFrame", "CZTPGGraphFrame setData");
        mainSc_.setData();
        return;
    }

    /**
    * Ｘ軸情報を設定する。。
    * @param iUnit ... 単位, iStart ... 開始値 , iEnd ... 終了値 , iBun ... 分割数
    * @return true ... OK, false ... NG
    */
    public boolean setXParam(int iUnit, int iStart, int iEnd, int iBun){
        grXUnit_    = iUnit;
        grXbun_     = iBun;
        grXMin_     = iStart;
        grXMax_     = iEnd;
        grXlength_  = new String(new Integer(iEnd).toString());
        return true;
    }
    /**
    * 炉番とDB名称を取得する。
    * @return true ... OK, false ... NG
    */
    public boolean setDefault(){
        roName_ = CZSystem.getRoName();
        roDbName_ = CZSystem.getDBName();
        return true;
    }

    //======================================================================
    /**
     *   GraphTitle表示Panel
     */
    //======================================================================
    class GraphTitlePanel extends JPanel {  

        /**
        * コンストラクタ
        */
        GraphTitlePanel(){
            super();

            setLayout(null);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            //画面タイトル部
            JLabel lbl1 = new JLabel(new String("トレンド表示画面"),JLabel.CENTER);
            lbl1.setFont(new java.awt.Font("dialog", 0, 32));
            lbl1.setBounds(300,0,500,36);
            lbl1.setForeground(java.awt.Color.black);
            add(lbl1);

            //固定表示部
            JLabel lbl2[] = new JLabel[12];

            JLabel lbl4[] = new JLabel[2];

            lbl4[0] = new JLabel("(#)",JLabel.CENTER);
            lbl4[0].setFont(new java.awt.Font("dialog", 0, 12));
            lbl4[0].setForeground(java.awt.Color.black);
            lbl4[0].setBorder(new Flush3DBorder());
            lbl4[0].setBounds(70,50,40,30);
            add(lbl4[0]);

            lbl2[0] = new JLabel("(日付時間)",JLabel.CENTER);
            lbl2[0].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[0].setForeground(java.awt.Color.black);
            lbl2[0].setBorder(new Flush3DBorder());
            lbl2[0].setBounds(150,50,60,30);
            add(lbl2[0]);

            lbl2[1] = new JLabel("(BtNo)",JLabel.CENTER);
            lbl2[1].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[1].setForeground(java.awt.Color.black);
            lbl2[1].setBorder(new Flush3DBorder());
            lbl2[1].setBounds(380,50,50,30);
            add(lbl2[1]);

            lbl2[2] = new JLabel("(品番)",JLabel.CENTER);
            lbl2[2].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[2].setForeground(java.awt.Color.black);
            lbl2[2].setBorder(new Flush3DBorder());
            lbl2[2].setBounds(520,50,50,30);
            add(lbl2[2]);

            lbl2[3] = new JLabel("(プロセス)",JLabel.CENTER);
            lbl2[3].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[3].setForeground(java.awt.Color.black);
            lbl2[3].setBorder(new Flush3DBorder());
            lbl2[3].setBounds(660,50,60,30);
            add(lbl2[3]);

            lbl2[4] = new JLabel("(チャージ量)",JLabel.CENTER);
            lbl2[4].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[4].setForeground(java.awt.Color.black);
            lbl2[4].setBorder(new Flush3DBorder());
            lbl2[4].setBounds(800,50,70,30);
            add(lbl2[4]);

            lbl2[5] = new JLabel("(T1No)",JLabel.CENTER);
            lbl2[5].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[5].setForeground(java.awt.Color.black);
            lbl2[5].setBorder(new Flush3DBorder());
            lbl2[5].setBounds(70,85,40,30);
            add(lbl2[5]);

            lbl2[6] = new JLabel("(T2No)",JLabel.CENTER);
            lbl2[6].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[6].setForeground(java.awt.Color.black);
            lbl2[6].setBorder(new Flush3DBorder());
            lbl2[6].setBounds(200,85,40,30);
            add(lbl2[6]);

            lbl2[7] = new JLabel("(T3No)",JLabel.CENTER);
            lbl2[7].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[7].setForeground(java.awt.Color.black);
            lbl2[7].setBorder(new Flush3DBorder());
            lbl2[7].setBounds(330,85,40,30);
            add(lbl2[7]);

            lbl2[8] = new JLabel("(T4No)",JLabel.CENTER);
            lbl2[8].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[8].setForeground(java.awt.Color.black);
            lbl2[8].setBorder(new Flush3DBorder());
            lbl2[8].setBounds(460,85,40,30);
            add(lbl2[8]);

            lbl2[9] = new JLabel("(T5No)",JLabel.CENTER);
            lbl2[9].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[9].setForeground(java.awt.Color.black);
            lbl2[9].setBorder(new Flush3DBorder());
            lbl2[9].setBounds(590,85,40,30);
            add(lbl2[9]);

            lbl2[10] = new JLabel("(T6No)",JLabel.CENTER);
            lbl2[10].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[10].setForeground(java.awt.Color.black);
            lbl2[10].setBorder(new Flush3DBorder());
            lbl2[10].setBounds(720,85,40,30);
            add(lbl2[10] );

            lbl2[11]  = new JLabel("(設定直径)[mm]",JLabel.CENTER);
            lbl2[11].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[11].setForeground(java.awt.Color.black);
            lbl2[11].setBorder(new Flush3DBorder());
            lbl2[11].setBounds(850,85,90,30);
            add(lbl2[11]);

            //データ表示部
            JLabel lbl3[] = new JLabel[13];

            lbl4[1] = new JLabel("",JLabel.CENTER);
            lbl4[1].setFont(new java.awt.Font("dialog", 0, 16));
            lbl4[1].setForeground(java.awt.Color.black);
            lbl4[1].setBorder(new Flush3DBorder());
            lbl4[1].setBounds(110,50,40,30);
            add(lbl4[1]);

            lbl3[0] = new JLabel("",JLabel.CENTER);
            lbl3[0].setFont(new java.awt.Font("dialog", 0, 14));
            lbl3[0].setForeground(java.awt.Color.black);
            lbl3[0].setBorder(new Flush3DBorder());
            lbl3[0].setBounds(210,50,170,30);
            add(lbl3[0]);

            lbl3[1] = new JLabel("",JLabel.CENTER);
            lbl3[1].setFont(new java.awt.Font("dialog", 0, 14));
            lbl3[1].setForeground(java.awt.Color.black);
            lbl3[1].setBorder(new Flush3DBorder());
            lbl3[1].setBounds(430,50,90,30);
            add(lbl3[1]);

            lbl3[2] = new JLabel("",JLabel.CENTER);
            lbl3[2].setFont(new java.awt.Font("dialog", 0, 14));
            lbl3[2].setForeground(java.awt.Color.black);
            lbl3[2].setBorder(new Flush3DBorder());
            lbl3[2].setBounds(570,50,90,30);
            add(lbl3[2]);

            lbl3[3] = new JLabel("",JLabel.CENTER);
            lbl3[3].setFont(new java.awt.Font("dialog", 0, 14));
            lbl3[3].setForeground(java.awt.Color.black);
            lbl3[3].setBorder(new Flush3DBorder());
            lbl3[3].setBounds(720,50,80,30);
            add(lbl3[3]);

            lbl3[4] = new JLabel("",JLabel.CENTER);
            lbl3[4].setFont(new java.awt.Font("dialog", 0, 14));
            lbl3[4].setForeground(java.awt.Color.black);
            lbl3[4].setBorder(new Flush3DBorder());
            lbl3[4].setBounds(870,50,80,30);
            add(lbl3[4]);

            lbl3[5] = new JLabel("",JLabel.CENTER);
            lbl3[5].setFont(new java.awt.Font("dialog", 0, 16));
            lbl3[5].setForeground(java.awt.Color.black);
            lbl3[5].setBorder(new Flush3DBorder());
            lbl3[5].setBounds(110,85,90,30);
            add(lbl3[5]);

            lbl3[6] = new JLabel("",JLabel.CENTER);
            lbl3[6].setFont(new java.awt.Font("dialog", 0, 16));
            lbl3[6].setForeground(java.awt.Color.black);
            lbl3[6].setBorder(new Flush3DBorder());
            lbl3[6].setBounds(240,85,90,30);
            add(lbl3[6]);

            lbl3[7] = new JLabel("",JLabel.CENTER);
            lbl3[7].setFont(new java.awt.Font("dialog", 0, 16));
            lbl3[7].setForeground(java.awt.Color.black);
            lbl3[7].setBorder(new Flush3DBorder());
            lbl3[7].setBounds(370,85,90,30);
            add(lbl3[7]);

            lbl3[8] = new JLabel("",JLabel.CENTER);
            lbl3[8].setFont(new java.awt.Font("dialog", 0, 16));
            lbl3[8].setForeground(java.awt.Color.black);
            lbl3[8].setBorder(new Flush3DBorder());
            lbl3[8].setBounds(500,85,90,30);
            add(lbl3[8]);

            lbl3[9] = new JLabel("",JLabel.CENTER);
            lbl3[9].setFont(new java.awt.Font("dialog", 0, 16));
            lbl3[9].setForeground(java.awt.Color.black);
            lbl3[9].setBorder(new Flush3DBorder());
            lbl3[9].setBounds(630,85,90,30);
            add(lbl3[9]);

            lbl3[10] = new JLabel("",JLabel.CENTER);
            lbl3[10].setFont(new java.awt.Font("dialog", 0, 16));
            lbl3[10].setForeground(java.awt.Color.black);
            lbl3[10].setBorder(new Flush3DBorder());
            lbl3[10].setBounds(760,85,90,30);
            add(lbl3[10]);

            lbl3[11] = new JLabel("",JLabel.CENTER);
            lbl3[11].setFont(new java.awt.Font("dialog", 0, 16));
            lbl3[11].setForeground(java.awt.Color.black);
            lbl3[11].setBorder(new Flush3DBorder());
            lbl3[11].setBounds(940,85,90,30);
            add(lbl3[11]);

			roBtTempCondition_ = CZSystem.getHikiageTemp(roDbName_,SelectBt,SelectTime);

            if (null != roBtTempCondition_){
                CZSystemBtTemp bt = (CZSystemBtTemp)roBtTempCondition_.elementAt(0);
                lbl4[1].setText(new Integer(SelectNo).toString());
                lbl3[0].setText((bt.t_time).trim()); 
                lbl3[1].setText((bt.batch).trim());  
                lbl3[2].setText((bt.hinshu).trim()); 
                lbl3[3].setText(CZSystem.getProcName(roBtStart_.p_no));
                lbl3[4].setText(new Integer(bt.i_sikomi + bt.t_sikomi).toString());
                lbl3[5].setText("Ｍ.Ｔ="   + new Integer(bt.no_youkai).toString());
                lbl3[6].setText("Ｐ.Ｔ="   + new Integer(bt.no_hikiage).toString());
                lbl3[7].setText("Ｒ.Ｔ="   + new Integer(bt.no_kaiten).toString());
                lbl3[8].setText("Ｅ.Ｔ="   + new Integer(bt.no_toridasi).toString());
                lbl3[9].setText("Ａ.Ｔ="   + new Integer(bt.no_aturyoku).toString());
                lbl3[10].setText("Ｃ.Ｔ="  + new Integer(bt.no_teisu).toString());
                lbl3[11].setText("ＤＩＡ=" + new Integer(bt.chokkei).toString());
            } else {
                if (null != roBtAllCondition_){
                    CZSystemBt bt = (CZSystemBt)roBtAllCondition_.elementAt(0);
                    lbl4[1].setText(new Integer(SelectNo).toString());
                    lbl3[0].setText((bt.t_time).trim()); 
                    lbl3[1].setText((bt.batch).trim());  
                    lbl3[2].setText((bt.hinshu).trim()); 
                    lbl3[3].setText(CZSystem.getProcName(roBtStart_.p_no));
                    lbl3[4].setText(new Integer(bt.i_sikomi + bt.t_sikomi).toString());
                    lbl3[5].setText("Ｍ.Ｔ="   + new Integer(bt.no_youkai).toString());
                    lbl3[6].setText("Ｐ.Ｔ="   + new Integer(bt.no_hikiage).toString());
                    lbl3[7].setText("Ｒ.Ｔ="   + new Integer(bt.no_kaiten).toString());
                    lbl3[8].setText("Ｅ.Ｔ="   + new Integer(bt.no_toridasi).toString());
                    lbl3[9].setText("Ａ.Ｔ="   + new Integer(bt.no_aturyoku).toString());
                    lbl3[10].setText("Ｃ.Ｔ="  + new Integer(bt.no_teisu).toString());
                    lbl3[11].setText("ＤＩＡ=" + new Integer(bt.chokkei).toString());
                } else {
                    lbl4[1].setText("");
                    lbl3[0].setText("");
                    lbl3[1].setText("");
                    lbl3[2].setText("");
                    lbl3[3].setText("");
                    lbl3[4].setText("");
                    lbl3[5].setText("Ｍ.Ｔ");
                    lbl3[6].setText("Ｐ.Ｔ");
                    lbl3[7].setText("Ｒ.Ｔ");
                    lbl3[8].setText("Ｅ.Ｔ");
                    lbl3[9].setText("Ａ.Ｔ");
                    lbl3[10].setText("Ｃ.Ｔ");
                    lbl3[11].setText("ＤＩＡ");
                }
            }
        }
    } //GraphTitlePanel

    //======================================================================
    /**
     *       メイングラフ
     */
    //======================================================================
    public class MainSc extends JScrollPane {

        private Rectangle   viewRec_    = null;
        private View        view_       = null;

        /**
        * コンストラクタ
        * @param comp ... Event Listener
        */
        MainSc(PVGrEventCompo comp){
            super();

            setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);
            setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
            view_ = new View();                  //グラフ表示領域
            setViewportView(view_);              //
            comp.setMainView(view_);             //
            view_.addComponentListener(comp);    //
            getViewport().setScrollMode(JViewport.BACKINGSTORE_SCROLL_MODE);
        }

        /**
        */
        public void setDefault(){
            viewRec_ = getViewportBorderBounds();
            view_.setPreferredSize(new Dimension(viewRec_.width,viewRec_.height*Y_VIEW_TIMES));
            view_.setLocation(0, -(viewRec_.height* Y_VIEW_TIMES - viewRec_.height));
            view_.setViewRec(viewRec_);
        }

        /**
        * データを設定し、グラフを再描画する。
        */
        public void setData(){
            view_.setData();                                 //データを再設定する。
            view_.repaint();                                 //画面を更新する。
        }

        //==================================================================
        /**
         *  グラフ描画パネル
         */
        //==================================================================
        class View extends JPanel {

            Rectangle viewRec = null;
            int xPosShld_[];
            int xPos_[];
            Vector yPosShld_    = null;
            Vector yPos_        = null;
            Vector col_         = null;
            Vector line_        = null;

            /**
            * コンストラクタ
            */
            View(){
                super();
                setName("View");
                setLayout(null);
                setBackground(BACK_COL);
            }
            /**
            * 枠を設定する。
            */
            public void setViewRec(Rectangle rec){
                viewRec = rec;
            }
            /**
            * グラフを描画する。
            */
            public void paint(Graphics g){
                Dimension d = getSize(null);                //画面サイズを取得する

                g.setColor(BACK_COL);                       //背景色を設定する
                g.fillRect(0,0,d.width,d.height);           //枠を設定する
                drawMemLine(g);                             //目盛線を引く
                drawLine(g);                                //グラフ線を引く
            }
            /**
            * 目盛線を描画する。
            */
            private void drawMemLine(Graphics g){
                float x;                                    //Ｘ軸
                float y;                                    //Ｙ軸
                float inc;                                  //増分

                Dimension d = getSize(null);                //画面サイズを取得する。
                //  1/1の目盛線
                g.setColor(MEM_LINE1_COL);
                inc = viewRec.width / grXbun_;
                for(x = 0.0f;  d.width > x; x+=inc){
                    g.drawLine((int)x,0,(int)x,d.height);
                }
                inc = viewRec.height / grYbun_;
                for(y = (float)d.height ;  0 < y ; y-=inc){
                    g.drawLine(0,(int)y,d.width,(int)y);
                }
            }
            /**
            * グラフ線を描画する。-------------------------
            */
            private void drawLine(Graphics g){
            
                int yPosShld[];
                int yPos[];
                int linSize = 1;
                if(null == pvDataBody_) return;
                int size = pvDataBody_.size();
                if(2 > size) return;
                yPos = new int[size];

                //グラフを描画する。
                for( int i=0; i < yPos_.size(); i++){
                    CZTPGFrame.SelectItemPanel item = (CZTPGFrame.SelectItemPanel)selList_.elementAt(i);
                    YPosData yd= (YPosData)yPos_.elementAt(i);
                    if(item.getLineS().length() == 0){
                        for(int j=0;j < size;j++){
                            yPos[j] = yd.getData(j)+1;
                        }
                        g.setColor((Color)col_.elementAt(i));
                        g.drawPolyline(xPos_,yPos,size);
                    }else{
                    for( int a=0; a < (new Integer(item.getLineS())).intValue(); a++){
                        for(int j=0;j < size;j++){
                            yPos[j] = yd.getData(j)+a;
                        }
                        g.setColor((Color)col_.elementAt(i));
                        g.drawPolyline(xPos_,yPos,size);
                    }
                    }
                }
            }
            /**
            * データから座標を計算する。
            */
			@SuppressWarnings("unchecked")
            private void setData(){

                if(null == pvDataBody_) return;
                int size = pvDataBody_.size();
                CZSystem.log("CZTPGFrame", "size: " + size);
                if(2 > size) return;

                Float xMax;

                Float sMin;
                float min;
                Float sMax;
                float max;

                float tmp;

                Float ftmp1;
                Float ftmp2;
                float tmp1;
                float tmp2;

                float val;
                float valShld;

                float hVal = 0.0f;      //@@@ 1件目のヒータ温度
                int chNo = 0;           //@@@ チャネル№用Work

                CZSystemPVData data;

                Dimension d = getSize(null);                //画面サイズを取得

                //Ｘ軸座標計算（肩）
                val = 0.0f;
                min = grXMin_;
                max = grXMax_;

                //Ｘ軸座標計算
                valShld = 0.0f;

                xPos_ = new int[size];
                for(int i = 0 ; i < size ; i++){
                    data = (CZSystemPVData)pvDataBody_.elementAt(i);
                    // Time or Length @@
                    if (1 == grXUnit_) {
                        val = (data.p_time / 60.0f) + valShld;  //sec -> min
                    } else {
                        val = data.p_length + valShld;          // mm
                    }
                    xPos_[i] = (int)_xPos(d.width, viewRec. width, min, max, val);
                }

                //Ｙ軸座標計算
                yPosShld_ = new Vector();
                yPos_     = new Vector();
                col_      = new Vector();
                for( int i=0; i < selList_.size(); i++){

                    CZTPGFrame.SelectItemPanel item = (CZTPGFrame.SelectItemPanel)selList_.elementAt(i);
                    if (!(item.getChNo().equals(""))){

                        col_.addElement(item.getColor());
                        sMin = new Float(item.getMin());
                        min  = sMin.floatValue();
                        sMax = new Float(item.getMax());
                        max  = sMax.floatValue();
                        //Ｙ軸座標計算(BODY)
                        YPosData yData  = new YPosData(size);
                        for(int j = 0 ; j < size ; j++){
                            chNo = (new Integer(item.getChNo())).intValue();    //@@@ PV番号を保持する。
                            data = (CZSystemPVData)pvDataBody_.elementAt(j);
                            val  = data.data[chNo - 1];                         //@@@ データを取出す。

                            if (15 == chNo) {           //ヒータ温度の時は
                                if (0 == j) {           //１件目のデータを保持する。
                                    hVal = val;
                                }
                                val = val - hVal;       //１件目の相対温度に変換する。
                            }
                            yData.setData(j, (int)_yPos(d.height,viewRec.height,min,max,val));
                        }
                        yPos_.addElement(yData);
                    }
                }
            }
            /**
            *データをＸ座標に変換する
            */
            private float _xPos(int dWidth, int vWidth, float min, float max, float val){
                float xDot = (float)vWidth / (max - min);
                float x    = xDot * (val - min);
                return x;
            }
            /**
            * Ｘ座標より値を求める
            */
            private float _xPosConv(int dWidth,int vWidth,float min,float max,int x){
                float xDot = (float)vWidth / (max - min);
                float val = x / xDot + min;
                return val;
            }
            /**
            * データをＹ座標に変換する。
            */
            private float _yPos(int dHeight,int vHeight,float min,float max,float val){
                float yDot = (float)vHeight / (max - min);
                float y = (float)vHeight - (yDot * (val - min));
                //
                if ( 0 > y ) y = 0.0f;
                if ( vHeight < y ) y = (float)vHeight;
                return y;
            }
            /**
            * Ｙ座標より値を求める
            */
            private float _yPosConv(int dHeight,int vHeight,float min,float max,int y){
                float yDot = (float)vHeight / (max - min);
                float val = (vHeight - y) / yDot + min;
                return val;
            }
            /**
            * YPosData Class Y座標を保持するクラス
            */
            class YPosData {

                int data[];
                /**
                * コンストラクタ
                * @param i ... データ点数
                */
                YPosData(int i){
                    data = new int[i];
                }
                /**
                * 座標データを設定する。
                * @param i ... Index,   val ... 座標値
                */
                public void setData(int i,int val){
                    data[i] = val;
                }
                /**
                * 座標データを取得する。
                * @param i ... index
                */
                public int getData(int i){
                    return data[i];
                }
            }

            /**
            * LineData Class Y座標を保持するクラス
            */
            class LineData {

                int data[];
                /**
                * コンストラクタ
                * @param i ... データ点数
                */
                LineData(int i){
                    data = new int[i];
                }
                /**
                * 座標データを設定する。
                * @param i ... Index,   val ... 座標値
                */
                public void setData(int i,int val){
                    data[i] = val;
                }
                /**
                * 座標データを取得する。
                * @param i ... index
                */
                public int getData(int i){
                    return data[i];
                }
            }

        } // View
    } // MainSc

    //======================================================================
    /**
     *       Ｘ軸の目盛表示用パネル
     */
    //======================================================================
    public class XSc extends JScrollPane {

        private Rectangle       viewRec_        = null;     //表示枠
        private View            view_           = null;     //

        /**
        * コンストラクタ
        */
        XSc(PVGrEventCompo comp){
            super();

            setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);
            setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
            //Ｘ軸の表示
            view_ = new View();
            setViewportView(view_);
            comp.setXView(view_);
            view_.addComponentListener(comp);
            getViewport().setScrollMode(JViewport.BACKINGSTORE_SCROLL_MODE);
        }

        /**
        * Default値を設定する。
        */
        public void setDefault(){
            viewRec_ = getViewportBorderBounds();
            view_.setPreferredSize(new Dimension(viewRec_.width,viewRec_.height));
            view_.setLocation(0,0);
            view_.setViewRec(viewRec_);
        }

        //==================================================================
        /**
         * X軸の目盛表示パネル
         */
        class View extends JPanel {
            Rectangle viewRec = null;

            /**
            * コンストラクタ
            */
            View(){
                super();
                setName("View");
                setLayout(null);
                setBackground(BACK_COL);
            }
            /**
            * 表示枠を設定する。
            */
            public void setViewRec(Rectangle rec){
                viewRec = rec;
            }
            /**
            * X軸目盛を描画する。
            */
            public void paint(Graphics g){
                Dimension d = getSize(null);

                g.setColor(BACK_COL);
                g.setFont(new java.awt.Font("dialog", 0, 10));
                g.fillRect(0,0,d.width,d.height);
                drawMemLine(g);
                drawMem(g);
            }
            /**
            * 目盛線を描画する。
            */
            private void drawMemLine(Graphics g){
                float x;
                float inc;

                Dimension d = getSize(null);

                g.setColor(MEM_LINE1_COL);
                inc = viewRec.width / grXbun_;
                for(x = 0.0f ;  d.width > x ; x+=inc){
                    g.drawLine((int)x,0,(int)x,d.height);
                }
            }
            /**
            * 目盛を描画する。
            */
            private void drawMem(Graphics g){
                Dimension d = getSize(null);
                g.setColor(MEM_LINE1_COL);

                float x;
                float inc;

                float tmp = grXMax_ - grXMin_;  
                float mem_inc = tmp / grXbun_;  
                float x_val   = grXMin_;
                inc = viewRec.width / grXbun_;

                for(x = 0.0f ;  d.width > x ; x+=inc){
                    DecimalFormat f1 = new DecimalFormat("0");
                    //g.drawString(String.valueOf(x_val),(int)x+3,viewRec.height/2);
                    g.drawString(String.valueOf(f1.format(x_val)),(int)x+3,viewRec.height/2);
                    x_val+=mem_inc;
                }
            }
        } // View
    } //XSc

    //======================================================================
    /**
     *       Ｙ軸グラフ左側目盛
     */
    //======================================================================
    public class Y1Sc extends JScrollPane {

        private Rectangle   viewRec_    = null;
        private View        view_       = null;
        /**
        * コンストラクタ
        */
        Y1Sc(PVGrEventCompo comp){
            super();

            //スクロールバーは表示しない。
            setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);
            setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
            //Ｖｉｅｗを設定する。
            view_ = new View();
            setViewportView(view_);
            //Listenerを追加する。
            comp.setY1View(view_);
            view_.addComponentListener(comp);
            getViewport().setScrollMode(JViewport.BACKINGSTORE_SCROLL_MODE);
        }
        /**
        * Default値を設定する。
        */
        public void setDefault(){
            viewRec_ = getViewportBorderBounds();
            view_.setPreferredSize(new Dimension(viewRec_.width*2,viewRec_.height*Y_VIEW_TIMES));
            view_.setLocation(0, -(viewRec_.height*Y_VIEW_TIMES - viewRec_.height));
            view_.setViewRec(viewRec_);
        }
        /**
        * Y軸左側を再描画する。
        */
        public void chgYSize(){
            view_.repaint();
        }
        /**
         * Y軸左側の目盛を表示する
         */
        class View extends JPanel {
            Rectangle viewRec = null;
            /**
            * コンストラクタ
            */
            View(){
                super();
                setName("View");
                setLayout(null);
                setBackground(BACK_COL);
            }
            /**
            * 領域を設定する。
            */
            public void setViewRec(Rectangle rec){
                viewRec = rec;
            }
            /**
            * 目盛を描画する。
            */
            public void paint(Graphics g){
                Dimension d = getSize(null);
                g.setColor(BACK_COL);
                g.fillRect(0,0,d.width,d.height);

                g.setColor(MEM_LINE1_COL);
                drawMemLine(g);                     //目盛線
                drawMemString(g);                   //目盛
            }
            /**
            * 目盛を描画する
            */
            private void drawMemString(Graphics g){ 
                Dimension d = getSize(null);
                int sa = 0;
                float xPos = 60.0f;
                int iCnt = 0;
                for( int i=0; i<selList_.size(); i++){
                    CZTPGFrame.SelectItemPanel item = (CZTPGFrame.SelectItemPanel)selList_.elementAt(i);
                    if (!(item.getChNo().equals(""))){
                        iCnt = iCnt + 1;
                        if (iCnt == 8){
                            xPos = 0.0f;
                            sa = 0;
                        }
                        g.setColor(item.getColor());
                        String min = new String(new Float(item.getMin()).toString());
                        String max = new String(new Float(item.getMax()).toString());
                        drawMem(g, d, xPos, sa, min, max);
                        sa = sa + 15;
                    }
                }

            }
            /**
            * 目盛線を描画する
            */
            private void drawMemLine(Graphics g){
                float y;
                float inc;

                Dimension d = getSize(null);
                //目盛分割
                inc = viewRec.height / grYbun_;
                //線を引く
                for(y = (float)d.height ;  0 < y ; y-=inc){
                    g.drawLine(0,(int)y,d.width,(int)y);
                }
            }
            /**
            * 目盛を描画する
            */
            private void drawMem(Graphics g,Dimension d,float xPos,int sa,String stMin,String stMax){

                float x = xPos;
                float y = 0.0f;

                //Min, Maxを実数値に直す。
                Float sMin = new Float(stMin);
                float min  = sMin.floatValue();
                Float sMax = new Float(stMax);
                float max  = sMax.floatValue();
                //増分を計算する。（inc .. 目盛, y_dot .. Y軸座標）
                float inc   = (max - min) / grYbun_;
                float yDot = viewRec.height / (max - min);
                //目盛を描画する。
                for(float tmp = min ; 0 <= y ; tmp += inc){
                    //Ｙ軸座標を計算する。
                    y = (float)d.height - yDot * (tmp - min);
                    g.drawString(String.valueOf(tmp),(int)x,(int)y-sa);
                }
            }
        } // View
    } // Y1Sc

    //==================================================================
    /**
     * Y軸右側(凡例)を表示するＰａｎｅｌ
     */
    //==================================================================
    public class Y2Sc extends JPanel
    {
        Rectangle viewRec = null;

        /**
        * コンストラクタ
        */
        Y2Sc(){
            super();
            setName("Y2Sc");
            setLayout(null);
            setBackground(BACK_COL);

            int y = 0;
            int inc = 24;
            JPanel p[] = new JPanel[14];
            JLabel lbl[] = new JLabel[14];
            for(int i=0; i<14; i++){
                lbl[i] = new JLabel();
                lbl[i].setFont(new java.awt.Font("dialog", 0, 14));

                p[i] = new JPanel();
                p[i].setBackground(BACK_COL);
                p[i].setBounds(0,y,100,inc);
                p[i].add(lbl[i]);
                add(p[i]);
                y = y + inc;
            }

            if(null != selList_){
                int dCount = 0;
                for( int i=0; i<selList_.size(); i++){
                    CZTPGFrame.SelectItemPanel item = (CZTPGFrame.SelectItemPanel)selList_.elementAt(i);
                    if (!(item.getChNo().equals(""))){
                        lbl[dCount].setText(item.getName());
                        lbl[dCount].setForeground(item.getColor());
                        dCount++;
                    }
                }
            }
        }
        /**
        * 領域を設定する。
        */
        public void setViewRec(Rectangle rec){
            viewRec = rec;
        }
    } // Y2Sc

    //======================================================================
    /**
     * グラフ表示領域のListener
     */
    //======================================================================
    class PVGrEventCompo implements ComponentListener {

        private JPanel mainView_    = null;     // グラフ表示パネル
        private JPanel xView_       = null;     // X軸目盛パネル
        private JPanel y1View_      = null;     // Y軸左目盛パネル
        private JPanel y2View_      = null;     // Y軸右パネル

        /**
        * コンストラクタ
        */
        PVGrEventCompo(){
        }
        /**
        * グラフ表示パネルを保持する
        */
        public void setMainView(JPanel view){
            mainView_ = view;
        }
        /**
        * X軸目盛表示パネルを保持する
        */
        public void setXView(JPanel view){
            xView_ = view;
        }
        /**
        * Y軸左側目盛表示パネルを保持する
        */
        public void setY1View(JPanel view){
            y1View_ = view;
        }
        /**
        * Y軸右側表示パネルを保持する
        */
        public void setY2View(JPanel view){
            y2View_ = view;
        }
        /**
        * 移動時の処理
        */
        public void componentMoved(java.awt.event.ComponentEvent e){
            if(xView_ == e.getComponent()){
                mainView_.setLocation(xView_.getX(),mainView_.getY());
            }
        }
        /**
        */
        public void componentResized(java.awt.event.ComponentEvent e){
        }
        /**
        */
        public void componentShown(java.awt.event.ComponentEvent e){
        }
        /**
        */
        public void componentHidden(java.awt.event.ComponentEvent e){
        }
    } // PVGrEventCompo
}  //TPGGraphFrame

