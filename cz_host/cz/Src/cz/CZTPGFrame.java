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

import javax.swing.JScrollBar;    //@20131017

//==============================================================================
/**
 * TPGグラフ
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * @Update 2003/06/01 ヒータ温度をプロセス毎の相対表示に変更( @@@ )
 * @Update 2003/08/04 仕込量を初期量+追加量に変更( @@@@ )
 * @Update 2008/09/17 TPG・PV保存対象表示情報追加( @@@@@ )
 * @Update 2013/10/17 TPGｸﾞﾗﾌﾊﾞｯﾁ番号保持機能( @20131017 )
 * @Update 2013/10/28 TPGｸﾞﾗﾌ入力最大数変更 ( @20131028 )
 */
//==============================================================================

public class CZTPGFrame extends JFrame
{

    private static CZTPGFrame tpg_      = null; //

    protected Container pane_;                  //

    static TitlePanel      pnl1_       = null; //
    private SelectPanel     pnl2_       = null; //
    private XLengeSetPanel  pnl3_       = null; //
    private PVIchiranPanel  pnl4_       = null; //
    private Hashtable       dataTbl_    = null; //

    private SelectItemPanel pnlSel_[]   = null; //

//    private static GraphDialog graDl_   = null; //グラフ表示用ダイアログ
    static  JLabel lblRo;
    private SercheDialog  sercheDia_    = null; //検索用ダイアログ
    private CZRoSelectWin3 rosel        = null;
    private static CZTPGGraphFrame graDl_   = null; //グラフ表示用ダイアログ

    private int gph_cnt = 0;

    private CZSystemStart roBtStart_    = null; //検索用引き上げ条件
    private Vector roBtAllCondition_    = null; //全Btの引き上げ条件
    
    private int    SelectNo             = 0;
    private String SelectBt             = null;
    private String SelectTime           = null;
    private Vector roBtTempCondition_    = null; //選択Btの引き上げ条件


    private String roName_              = null; //対象炉番
    private String roDbName_            = null; //対象炉データベース名

    private Vector  selList_            = null; //Y軸項目

    private Vector pvDataBody_          = null; //PVデータ

    private String prop_xUnit = null;
    private String prop_xMin  = null;
    private String prop_xMax  = null;
    private String prop_xBun  = null;

    private String prop_yNo[];
    private String prop_yMin[];
    private String prop_yMax[];
    private String prop_yCol[];
    private String prop_yLine[];

    private JButton      btnOpen_ = null;
    private JButton      btnSave_ = null;
    private File            file_ = new File(CZSystem.FILE_SRC_PATH);

    private int SelBtRow = 0;    //選択Bt Row(バッファ）(初期値:0) @20131017
    //==========================================================================
    /**
     * スタートアップ
     * @param   args    コマンドライン引数
     */
    //==========================================================================
/**@@
    public static void main(String[] args) {

        JFrame tpg_ = new CZTPGFrame("TPGデバッグ用パネル", 0 );
        tpg_.setSize(1024, 660);
        tpg_.setVisible(true);
    }
@@*/
    //==========================================================================
    /**
     * コンストラクタ
     * @param   String title  Frame Title
     * @param   int ui        UI Manager Look and Feel
     */
    //==========================================================================
    public CZTPGFrame()
    {

        super();
        setupUI(0);                                 //UIを設定する。

        roName_     = CZSystem.getRoName();         //炉名を取得する。
        roDbName_   = CZSystem.getDBName();         //DBスキーマ名を取得する。

        setTitle("トレンドテーブル設定");                         //画面Title
//        setTitle("トレンドテーブル");                         //画面Title
        setSize(1024, 800);
        setResizable(false);                        //画面のサイズ変更は不可
//        setModal(true);                             //Modalで表示

        pane_ = getContentPane();
        pane_.setLayout(null);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            pane_.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            pane_.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        try{
            // ----- Property_Fileより Min,Max値を取得する。 --------
            Properties prop =  new Properties();
            FileInputStream pros = new FileInputStream("TPGDEF.TXT");
            prop.load(pros);
            // X軸の設定
            prop_xUnit = prop.getProperty("X_UNIT");
            prop_xMin  = prop.getProperty("X_START");
            prop_xMax  = prop.getProperty("X_END");
            prop_xBun  = prop.getProperty("X_BUNKATU");

            // Y軸の設定
            prop_yNo  = new String[14];
            prop_yMin = new String[14];
            prop_yMax = new String[14];
            prop_yCol = new String[14];
            prop_yLine = new String[14];
            for(int i=0; i < 14 ; i++){
                try {
                    prop_yNo[i]   = prop.getProperty("Y" + (i+1) + "_NO");
					if (prop_yNo[i] == null )
	                {
						CZSystem.log("CZTPGFrame", "prop_yNo[" + i + "][Null]");
					    prop_yNo[i]   = new String("");
					}
		CZSystem.log("CZTPGFrame", "GET prop_yNo[" + i + "][" + prop_yNo[i] + "]");
                    prop_yMin[i]  = prop.getProperty("Y" + (i+1) + "_MIN");
					if (prop_yMin[i] == null )
	                {
						CZSystem.log("CZTPGFrame", "prop_yMin[" + i + "][Null]");
					    prop_yMin[i]   = new String("0");
					}
					if (prop_yMin[i].equals(""))
	                {
						CZSystem.log("CZTPGFrame", "prop_yMin[" + i + "][]");
					    prop_yMin[i] = new String("0");
					}
		CZSystem.log("CZTPGFrame", "GET prop_yMin[" + i + "][" + prop_yMin[i] + "]");
                    prop_yMax[i]  = prop.getProperty("Y" + (i+1) + "_MAX");
					if (prop_yMax[i] == null )
	                {
						CZSystem.log("CZTPGFrame", "prop_yMax[" + i + "][Null]");
					    prop_yMax[i]   = new String("10");
					}
					if (prop_yMax[i].equals(""))
	                {
						CZSystem.log("CZTPGFrame", "prop_yMax[" + i + "][]");
					    prop_yMax[i] = new String("10");
					}
		CZSystem.log("CZTPGFrame", "GET prop_yMax[" + i + "][" + prop_yMax[i] + "]");
                    prop_yCol[i]  = prop.getProperty("Y" + (i+1) + "_COLOR");
					if (prop_yCol[i] == null )
	                {
						CZSystem.log("CZTPGFrame", "prop_yCol[" + i + "][Null]");
					    prop_yCol[i]   = new String("255,255,255");
					}
					if (prop_yCol[i].equals(""))
	                {
						CZSystem.log("CZTPGFrame", "prop_yCol[" + i + "][]");
					    prop_yCol[i] = new String("255,255,255");
					}
		CZSystem.log("CZTPGFrame", "GET prop_yCol[" + i + "][" + prop_yCol[i] + "]");

                    prop_yLine[i] = prop.getProperty("Y" + (i+1) + "_LINE");
					if (prop_yLine[i] == null)
					{
						CZSystem.log("CZTPGFrame", "prop_yLine[" + i + "][null]");
	                    prop_yLine[i] = new String("1");
					}

					if (prop_yLine[i].equals(""))
	                {
						CZSystem.log("CZTPGFrame", "prop_yLine[" + i + "][]");
					    prop_yLine[i] = new String("1");
					}
		CZSystem.log("CZTPGFrame", "GET prop_yLine[" + i + "][" +prop_yLine[i] + "]");
		CZSystem.log("CZTPGFrame", "GET OK!");
                } catch (Exception e) {
		CZSystem.log("CZTPGFrame", "GET ERROR ---!!");
                    prop_yNo[i]   = new String("");
                    prop_yMin[i]  = new String("0");
                    prop_yMax[i]  = new String("10");
                    prop_yCol[i]  = new String("255,255,255");
                    prop_yLine[i] = new String("1");
		CZSystem.log("CZTPGFrame", "CLR ERROR ---!!");
                }
            }
        } catch( Exception e ) {
                                        //プロパティ取得でエラーの時は、終了する。
		CZSystem.log("CZTPGFrame", "GET ERROR EXIT !!");
            CZSystem.exit(-1,"CZTPG NO Propertie File");
        }

        makePanels();                               //設定画面を生成する。
        sercheDia_ = null;
        sercheDia_ = new SercheDialog();            //検索画面を生成する。
        sercheDia_.setVisible(false);               //検索画面を閉じておく。

		CZSystem.log("CZTPGFrame", "CZTPG new");
    }
    //==========================================================================
    /**
     * UIのSetup
     * @param   int ui        UI Manager Look and Feel
     */
    //==========================================================================
    protected void setupUI( int ui)
    {

        addWindowListener(  
            new WindowAdapter()
            {
                public void windowClosing(WindowEvent e) {
					DefTextSave();
					setVisible(false);
				}
            }
        );

        // フレームの初期化     ------------------------------------------------
        try
        {   
            if( ui == 0 )
            {
                UIManager.setLookAndFeel(
                       "com.sun.java.swing.plaf.metal.MetalLookAndFeel");
            }
            else if( ui == 1 ) {
                UIManager.setLookAndFeel(
                       "com.sun.java.swing.plaf.motif.MotifLookAndFeel");
            }
            else {
                UIManager.setLookAndFeel(
                       "com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
            }

            SwingUtilities.updateComponentTreeUI( this );
        }
        catch ( UnsupportedLookAndFeelException e ) {
        }
        catch ( ClassNotFoundException e ) {
        }
        catch ( InstantiationException e ) {
        }
        catch ( IllegalAccessException e ) {
        }
    }

	protected void DefTextSave()
	{

/***** System.gc() *****/
//            System.out.println(Runtime.getRuntime().freeMemory());
            System.gc();
//            System.out.println(Runtime.getRuntime().freeMemory() + "  GC FREE!!");
/**********************/

		Properties prop = new Properties();           // プロパティを生成する
		// X軸の設定
		prop.setProperty(new String("X_UNIT"),    new String("" + pnl3_.getUnit()) );
		prop.setProperty(new String("X_START"),   new String("" + pnl3_.getStart()));
		prop.setProperty(new String("X_END"),     new String("" + pnl3_.getEnd())  );
		prop.setProperty(new String("X_BUNKATU"), new String("" + pnl3_.getMesh()) );

		//Y軸の設定
		for (int i = 0; i < 14; i++) {
		  /*
		  if( new String(pnlSel_[i].getLineS()).length() == 0 ){
			Object msg[] = { "条件を入力して下さい", "", "" };
			errorMsg(msg);
			return;
		  }
		  */
		  prop.setProperty(new String("Y" + (i+1) + "_NO"),  new String("" + pnlSel_[i].getChNo()));
		  prop.setProperty(new String("Y" + (i+1) + "_MIN"), new String("" + pnlSel_[i].getMin() ));
		  prop.setProperty(new String("Y" + (i+1) + "_MAX"), new String("" + pnlSel_[i].getMax() ));
		  prop.setProperty(new String("Y" + (i+1) + "_COLOR"), 
		                 new String(pnlSel_[i].getColor().getRed()   + "," +
		                            pnlSel_[i].getColor().getGreen() + "," +
		                            pnlSel_[i].getColor().getBlue() ));
		  prop.setProperty(new String("Y" + (i+1) + "_LINE"), new String("" + pnlSel_[i].getLineS() ));
		}
		//---------- ファイルに保存する  ----------
		try {
//			CZSystem.log("CZTPGFrame ","ファイルに保存した。");
//		    FileOutputStream out = new FileOutputStream("d:/CZ/classes/TPGDEF.TXT");
		    FileOutputStream out = new FileOutputStream("TPGDEF.TXT");
		    prop.store(out, "");
		    out.flush();
		    out.close();
		} catch (IOException ex) {
		    JOptionPane.showMessageDialog(
		      tpg_,
		      new String("保存できませんでした。"),
		      new String("保存"),
		      JOptionPane.WARNING_MESSAGE);
		    return;
		}
	}

    protected void windowClose()
    {

	    pnl1_       = null;
	    pnl2_       = null;
	    pnl3_       = null;
	    pnl4_       = null;
	    dataTbl_    = null;
//	    graDl_   = null;
	    sercheDia_    = null;
	    roBtStart_    = null;
	    roBtAllCondition_    = null;
	    roBtTempCondition_    = null;
	    selList_            = null;
	    pvDataBody_          = null;

/***** System.gc() *****/
//            System.out.println(Runtime.getRuntime().freeMemory());
            System.gc();
//            System.out.println(Runtime.getRuntime().freeMemory() + "  GC FREE!!");
/**********************/

        setVisible(false);
    }
//================================ ここから設定画面 =======================================
    //--------------------------------------------------------------------------
    /**
     * 画面作成
     */
    //--------------------------------------------------------------------------
	@SuppressWarnings("unchecked")
    protected void makePanels()
    {

        Border brd = BorderFactory.createRaisedBevelBorder();

        //------------------------- 項目一覧を表示するPanel  -------------------
        pnl4_ = new PVIchiranPanel();
        pnl4_.setBounds(531,281,470,280);
        pane_.add( pnl4_ );

        //------------------------- Titleを表示する Panel ----------------------
        pnl1_ = new TitlePanel();
        pnl1_.setBounds(0,0,1024,160);
        pane_.add( pnl1_ );

        //------------------------- 選択を指定するPanel  -----------------------
        pnl2_ = new SelectPanel();
        pnl2_.setBounds(0,161,500,530);
        pane_.add( pnl2_ );

        //------------------------- X軸の設定を指示する Panel ------------------
        pnl3_ = new XLengeSetPanel();
        pnl3_.setBounds(531,161,470,120);

        pane_.add( pnl3_ );

        //------------------------- ボタンPanel  -------------------
        JPanel pnl5 = new JPanel();
        pnl5.setLayout( null );
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            pnl5.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            pnl5.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }
        pnl5.setBounds(0,700,1024,60);

        JButton btnGraph = new JButton("グラフ表示");
        btnGraph.setBounds(30, 20, 100, 30);
        btnGraph.setLocale(new Locale("ja","JP"));  
        btnGraph.setFont(new java.awt.Font("dialog", 0, 14));
        btnGraph.setBorder(new Flush3DBorder());
        btnGraph.setForeground(java.awt.Color.black);;
        btnGraph.addActionListener(
          new ActionListener() {
              public void actionPerformed(ActionEvent e) {
                DefTextSave();
                selList_ = null;
                selList_ = new Vector();
                for (int i=0; i<14; i++){
                    if( null != pnlSel_[i].getChNo() ){
                        selList_.addElement(pnlSel_[i]);
                    }
                }

                graDl_ = null;
                System.gc();                    //@@@ 取敢えずGCを実行する。
                
                if( SelectBt == null){
					Object msg[] = { "グラフ表示条件を選択して下さい", "", "" };
					errorMsg(msg);
					return;
				}else{
                    gph_cnt = CZSystem.GraphCount();
                    if(gph_cnt > 4){
                        Object msg[] = { "グラフは５枚以上開けません", "", "" };
                        errorMsg(msg);
						return;
					}else{
                        graDl_ = new CZTPGGraphFrame(SelectNo,roDbName_,SelectBt,SelectTime,pvDataBody_,selList_,roBtStart_);     //グラフを生成する。
                        graDl_.setXParam(pnl3_.getUnit(),pnl3_.getStart(),pnl3_.getEnd(),pnl3_.getMesh());
                        graDl_.setData();               //
                        graDl_.setVisible(true);        //グラフを表示する。
                        CZSystem.GraphCountUp();
                    }
                }
              }
          }
        );
        pnl5.add(btnGraph);

        btnOpen_ = new JButton("設定読込");
        btnOpen_.setLocale(new Locale("ja","JP"));
        btnOpen_.setFont(new java.awt.Font("dialog", 0, 14));
        btnOpen_.setBorder(new Flush3DBorder());
        btnOpen_.setForeground(java.awt.Color.black);;
        btnOpen_.addActionListener(
          new ActionListener() {
              public void actionPerformed(ActionEvent evt) {
                  JFileChooser chooser = new JFileChooser(file_);
                  int ret = chooser.showOpenDialog(tpg_);
                  if ( ret == JFileChooser.APPROVE_OPTION ) {
                      file_ = chooser.getSelectedFile();        // ファイル名を取得する
                      Properties prop = new Properties();       // プロパティを生成する
                      try {
                          FileInputStream in = new FileInputStream(file_);
                          prop.load( in );                      //プロパティを取得する。
                          in.close();
                          prop.list(System.out);

                          for(int i=0; i < 14 ; i++){
                              try {
                                  prop_yNo[i]   = prop.getProperty("Y" + (i+1) + "_NO");
								  if (prop_yNo[i] == null )
					              {
									CZSystem.log("CZTPGFrame", "prop_yNo[" + i + "][Null]");
									prop_yNo[i]   = new String("");
								  }

                                  prop_yMin[i]  = prop.getProperty("Y" + (i+1) + "_MIN");
								  if (prop_yMin[i] == null )
				                  {
									CZSystem.log("CZTPGFrame", "prop_yMin[" + i + "][Null]");
								    prop_yMin[i]   = new String("0");
								  }
								  if (prop_yMin[i].equals(""))
				                  {
									CZSystem.log("CZTPGFrame", "prop_yMin[" + i + "][]");
								    prop_yMin[i] = new String("0");
								  }

                                  prop_yMax[i]  = prop.getProperty("Y" + (i+1) + "_MAX");
								  if (prop_yMax[i] == null )
				                  {
									CZSystem.log("CZTPGFrame", "prop_yMax[" + i + "][Null]");
								    prop_yMax[i]   = new String("10");
								  }
								  if (prop_yMax[i].equals(""))
				                  {
									CZSystem.log("CZTPGFrame", "prop_yMax[" + i + "][]");
								    prop_yMax[i] = new String("10");
								  }

                                  prop_yCol[i]  = prop.getProperty("Y" + (i+1) + "_COLOR"); 
								  if (prop_yCol[i] == null )
				                  {
									CZSystem.log("CZTPGFrame", "prop_yCol[" + i + "][Null]");
								    prop_yCol[i]   = new String("255,255,255");
								  }
								  if (prop_yCol[i].equals(""))
				                  {
									CZSystem.log("CZTPGFrame", "prop_yCol[" + i + "][]");
								    prop_yCol[i] = new String("255,255,255");
								  }
                                
                                  prop_yLine[i]  = prop.getProperty("Y" + (i+1) + "_LINE");
								  if (prop_yLine[i] == null)
								  {
									CZSystem.log("CZTPGFrame", "prop_yLine[" + i + "][null]");
					                prop_yLine[i] = new String("1");
								  }

								  if (prop_yLine[i].equals(""))
					              {
									CZSystem.log("CZTPGFrame", "prop_yLine[" + i + "][]");
								    prop_yLine[i] = new String("1");
								  }
                              } catch (Exception e) {
                                  prop_yNo[i]   = new String("");
                                  prop_yMin[i]  = new String("0");
                                  prop_yMax[i]  = new String("10");
                                  prop_yCol[i]  = new String("255,255,255");
                                  prop_yLine[i]  = new String("1");
                              }
                          }

                          for(int i=0; i < 14; i++ ){
                              if(!(prop_yNo[i] == null)){
			                      if (prop_yNo[i].equals("") == false) {
                                      String rgb[] = new String[3];
                                      int cCount =0;
                                      StringTokenizer st = new StringTokenizer(prop_yCol[i],",");
                                      while (st.hasMoreTokens()) {
                                          if (3 == cCount) break;
                                          rgb[cCount] = new String((st.nextToken()).trim());
                                          cCount++;
                                      }
                                      if (3 > cCount) {
                                          System.out.println("***** TPG Propaty File Error !! *****");
                                          System.exit(-1);
                                      }
                                      Color col = new Color(Integer.parseInt(rgb[0]),
                                                            Integer.parseInt(rgb[1]),
                                                            Integer.parseInt(rgb[2]));
                                      pnlSel_[i].setColor(col);
                                      pnlSel_[i].setChNo(prop_yNo[i]);
                                      pnlSel_[i].setMin(Float.parseFloat(prop_yMin[i]));
                                      pnlSel_[i].setMax(Float.parseFloat(prop_yMax[i]));
                                      CZSystemPVName n = (CZSystemPVName)dataTbl_.get(new Integer(prop_yNo[i]));
                                      if (null != n) {
                                          pnlSel_[i].setName( n.k_name);
                                      } else {
                                      }
                                  } else {
                                      pnlSel_[i].setColor(Color.white);
                                      pnlSel_[i].setChNo("");
                                      //pnlSel_[i].setLineS("1");
                                      pnlSel_[i].setName( new String(""));
                                      pnlSel_[i].setMin(0.0f);
                                      pnlSel_[i].setMax(10.0f);
                                  }
                                  try{
                                      pnlSel_[i].setLineS(prop_yLine[i]);
                                  }catch (Exception e){
                                      pnlSel_[i].setLineS("1");
                                  }

                              }else{
                                  pnlSel_[i].setColor(Color.white);
                                  pnlSel_[i].setChNo("");
                                  pnlSel_[i].setLineS("1");
                                  pnlSel_[i].setName( new String(""));
                                  pnlSel_[i].setMin(0.0f);
                                  pnlSel_[i].setMax(10.0f);
                              }
                          }
                      } catch ( IOException ex ) {
                          CZSystem.log("CZTPGFrame ","Property Fileがロードできませんでした。");
                          return;
                      }
                  }
              }
          }
        );
        btnOpen_.setBounds(600, 20, 100, 30);
        pnl5.add(btnOpen_); 
        // ======================================== [保存]ボタン ==================================
        btnSave_ = new JButton("設定保存");
        btnSave_.setLocale(new Locale("ja","JP"));
        btnSave_.setFont(new java.awt.Font("dialog", 0, 14));
        btnSave_.setBorder(new Flush3DBorder());
        btnSave_.setForeground(java.awt.Color.black);;
        btnSave_.addActionListener(
          new ActionListener() {
              public void actionPerformed(ActionEvent evt)
              {
                  JFileChooser chooser = new JFileChooser(file_);
                  int ret = chooser.showSaveDialog(tpg_);
                  if (ret == JFileChooser.APPROVE_OPTION) {
                      file_ = chooser.getSelectedFile();            // ファイル名を取得する
                      Properties prop = new Properties();           // プロパティを生成する
                      // X軸の設定
                      prop.setProperty(new String("X_UNIT"),    new String("" + pnl3_.getUnit()) );
                      prop.setProperty(new String("X_START"),   new String("" + pnl3_.getStart()));
                      prop.setProperty(new String("X_END"),     new String("" + pnl3_.getEnd())  );
                      prop.setProperty(new String("X_BUNKATU"), new String("" + pnl3_.getMesh()) );

                      //Y軸の設定
                      for (int i = 0; i < 14; i++) {
                        prop.setProperty(new String("Y" + (i+1) + "_NO"),  new String("" + pnlSel_[i].getChNo()));
                        prop.setProperty(new String("Y" + (i+1) + "_MIN"), new String("" + pnlSel_[i].getMin() ));
                        prop.setProperty(new String("Y" + (i+1) + "_MAX"), new String("" + pnlSel_[i].getMax() ));
                        prop.setProperty(new String("Y" + (i+1) + "_COLOR"), 
                                       new String(pnlSel_[i].getColor().getRed()   + "," +
                                                  pnlSel_[i].getColor().getGreen() + "," +
                                                  pnlSel_[i].getColor().getBlue() ));
                        prop.setProperty(new String("Y" + (i+1) + "_LINE"), new String("" + pnlSel_[i].getLineS() ));
                      }
                      //---------- ファイルに保存する  ----------
                      try {
//						CZSystem.log("CZTPGFrame ","ファイルに保存した。");
                          FileOutputStream out = new FileOutputStream(file_);
                          prop.store(out, "");
                          out.flush();
                          out.close();
                      } catch (IOException ex) {
                          JOptionPane.showMessageDialog(
                            tpg_,
                            new String("保存できませんでした。"),
                            new String("保存"),
                            JOptionPane.WARNING_MESSAGE);
                          return;
                      }
                      JOptionPane.showMessageDialog(
                        tpg_,
                        new String("保存しました。"),
                        new String("保存"),
                        JOptionPane.INFORMATION_MESSAGE);
                      return;
                  }
              }
          }
        );
        btnSave_.setBounds(750, 20, 100, 30);
        pnl5.add(btnSave_); 

        JButton btnExit = new JButton("終了");
        btnExit.setBounds(900, 20, 100, 30);
        btnExit.setLocale(new Locale("ja","JP"));
        btnExit.setFont(new java.awt.Font("dialog", 0, 14));
        btnExit.setBorder(new Flush3DBorder());
        btnExit.setForeground(java.awt.Color.black);;
        btnExit.addActionListener(
          new ActionListener() {
              public void actionPerformed(ActionEvent e) {
				DefTextSave();
                windowClose();
              }
          }
        );
        pnl5.add(btnExit);
        pane_.add( pnl5 );
    }

    //----------------------------------------------------------------------
    /**
     *  @param msg ... メッセージ内容
     *  @return true ... OK, false ... NG
     */
    //----------------------------------------------------------------------
    private boolean errorMsg(Object msg[]){
        JOptionPane.showMessageDialog(null,msg,
                                "入力エラー",
                                JOptionPane.ERROR_MESSAGE);
        return true;
    }


    //==========================================================================
    /**
     *   Title表示Panel
     */
    //==========================================================================
    class TitlePanel extends JPanel {

        JLabel lbl3[] = new JLabel[13];

        JLabel lbl4[] = new JLabel[2];

        /**
        * コンストラクタ
        */
		@SuppressWarnings("unchecked")
        TitlePanel(){
            super();

            setLayout(null);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }
            setForeground(Color.black);

            JLabel lbl1 = new JLabel("トレンドテーブル設定",JLabel.CENTER);
            lbl1.setLayout(new FlowLayout(FlowLayout.CENTER));
            lbl1.setFont(new java.awt.Font("dialog", 0, 32));
            lbl1.setForeground(Color.black);
            lbl1.setBounds(0,0,900,36);
            add(lbl1);

			String s = CZSystem.RoKetaChg(CZSystem.getRoName());	// 20050725 炉：表示桁数変更
            lblRo = new JLabel(s,JLabel.CENTER);
//            JLabel lblRo = new JLabel(CZSystem.getRoName(),JLabel.CENTER);
            lblRo.setFont(new java.awt.Font("dialog", 0, 14));
            lblRo.setForeground(Color.black);
            lblRo.setBorder(new Flush3DBorder());
            lblRo.setBounds(20,20,80,30);
            add(lblRo);

            JButton btn_chgRo = new JButton("▼");
            btn_chgRo.setBounds(100, 20, 30, 30);
            btn_chgRo.setFont(new java.awt.Font("dialog", 0, 20));
            btn_chgRo.setBorder(new Flush3DBorder());
            btn_chgRo.setForeground(java.awt.Color.black);
            btn_chgRo.addActionListener(
                new ActionListener() {
					public void actionPerformed(ActionEvent ev){
						rosel = new CZRoSelectWin3();
						rosel.setVisible(true);
						roName_ = lblRo.getText();
//						StringBuffer a = new StringBuffer();
//						a.append(lblRo.getText());
//						a.insert(0,"E");
//						roDbName_ = a.toString();
						roDbName_ = lblRo.getText();
						SelBtRow = 0;  //@20131017 初期値0に戻す
						CZSystem.log("CZTPGFrame TitlePanel", "バッチ選択番号　初期値0");  //@20131017
					}
				}
			);
            add(btn_chgRo);

            JButton btnSearch = new JButton("検  索");
            btnSearch.setBounds(20, 60, 80, 30);
            btnSearch.setLocale(new Locale("ja","JP"));
            btnSearch.setFont(new java.awt.Font("dialog", 0, 14));
            btnSearch.setBorder(new Flush3DBorder());
            btnSearch.setForeground(java.awt.Color.black);
            btnSearch.addActionListener(
                new ActionListener() {
                    public void actionPerformed(ActionEvent ev){
                        selList_ = null;
                        selList_ = new Vector();            //選択リストを作成する。
                        for (int i=0; i<14; i++){
                            if( null != pnlSel_[i].getChNo() ){
                                selList_.addElement(pnlSel_[i]);
                            }
                        }
                        sercheDia_.setDefault();            //検索画面を初期化する。
                        sercheDia_.setVisible(true);        //検索画面を表示する。
                    }
                }
            );
            add(btnSearch);

            //固定表示部
            JLabel lbl2[] = new JLabel[12];

            lbl4[0] = new JLabel("(#)",JLabel.CENTER);
            lbl4[0].setFont(new java.awt.Font("dialog", 0, 12));
            lbl4[0].setForeground(java.awt.Color.black);
            lbl4[0].setBorder(new Flush3DBorder());
            lbl4[0].setBounds(100,60,40,30);
            add(lbl4[0]);

            lbl2[0] = new JLabel("(日付時間)",JLabel.CENTER);
            lbl2[0].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[0].setForeground(java.awt.Color.black);
            lbl2[0].setBorder(new Flush3DBorder());
            lbl2[0].setBounds(180,60,60,30);
            add(lbl2[0]);

            lbl2[1] = new JLabel("(BtNo)",JLabel.CENTER);
            lbl2[1].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[1].setForeground(java.awt.Color.black);
            lbl2[1].setBorder(new Flush3DBorder());
            lbl2[1].setBounds(410,60,50,30);
            add(lbl2[1]);

            lbl2[2] = new JLabel("(品番)",JLabel.CENTER);
            lbl2[2].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[2].setForeground(java.awt.Color.black);
            lbl2[2].setBorder(new Flush3DBorder());
            lbl2[2].setBounds(550,60,50,30);
            add(lbl2[2]);

            lbl2[3] = new JLabel("(プロセス)",JLabel.CENTER);
            lbl2[3].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[3].setForeground(java.awt.Color.black);
            lbl2[3].setBorder(new Flush3DBorder());
            lbl2[3].setBounds(690,60,60,30);
            add(lbl2[3]);

            lbl2[4] = new JLabel("(チャージ量)",JLabel.CENTER);
            lbl2[4].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[4].setForeground(java.awt.Color.black);
            lbl2[4].setBorder(new Flush3DBorder());
            lbl2[4].setBounds(830,60,70,30);
            add(lbl2[4]);

            lbl2[5] = new JLabel("(T1)",JLabel.CENTER);
            lbl2[5].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[5].setForeground(java.awt.Color.black);
            lbl2[5].setBorder(new Flush3DBorder());
            lbl2[5].setBounds(20,95,40,30);
            add(lbl2[5]);

            lbl2[6] = new JLabel("(T2)",JLabel.CENTER);
            lbl2[6].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[6].setForeground(java.awt.Color.black);
            lbl2[6].setBorder(new Flush3DBorder());
            lbl2[6].setBounds(150,95,40,30);
            add(lbl2[6]);

            lbl2[7] = new JLabel("(T3)",JLabel.CENTER);
            lbl2[7].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[7].setForeground(java.awt.Color.black);
            lbl2[7].setBorder(new Flush3DBorder());
            lbl2[7].setBounds(280,95,40,30);
            add(lbl2[7]);

            lbl2[8] = new JLabel("(T4)",JLabel.CENTER);
            lbl2[8].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[8].setForeground(java.awt.Color.black);
            lbl2[8].setBorder(new Flush3DBorder());
            lbl2[8].setBounds(410,95,40,30);
            add(lbl2[8]);

            lbl2[9] = new JLabel("(T5)",JLabel.CENTER);
            lbl2[9].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[9].setForeground(java.awt.Color.black);
            lbl2[9].setBorder(new Flush3DBorder());
            lbl2[9].setBounds(540,95,40,30);
            add(lbl2[9]);

            lbl2[10] = new JLabel("(T6)",JLabel.CENTER);
            lbl2[10].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[10].setForeground(java.awt.Color.black);
            lbl2[10].setBorder(new Flush3DBorder());
            lbl2[10].setBounds(670,95,40,30);
            add(lbl2[10] );

//            lbl2[11]  = new JLabel("(mm)(設定直径)",JLabel.CENTER);
            lbl2[11]  = new JLabel("(設定直径)[mm]",JLabel.CENTER);
            lbl2[11].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[11].setForeground(java.awt.Color.black);
            lbl2[11].setBorder(new Flush3DBorder());
            lbl2[11].setBounds(800,95,90,30);
            add(lbl2[11]);

            //データ表示部

            lbl4[1] = new JLabel("",JLabel.CENTER);
            lbl4[1].setFont(new java.awt.Font("dialog", 0, 16));
            lbl4[1].setForeground(java.awt.Color.black);
            lbl4[1].setBorder(new Flush3DBorder());
            lbl4[1].setBounds(140,60,40,30);
            add(lbl4[1]);

            lbl3[0] = new JLabel("",JLabel.CENTER);
            lbl3[0].setFont(new java.awt.Font("dialog", 0, 14));
            lbl3[0].setForeground(java.awt.Color.black);
            lbl3[0].setBorder(new Flush3DBorder());
            lbl3[0].setBounds(240,60,170,30);
            add(lbl3[0]);

            lbl3[1] = new JLabel("",JLabel.CENTER);
            lbl3[1].setFont(new java.awt.Font("dialog", 0, 14));
            lbl3[1].setForeground(java.awt.Color.black);
            lbl3[1].setBorder(new Flush3DBorder());
            lbl3[1].setBounds(460,60,90,30);
            add(lbl3[1]);

            lbl3[2] = new JLabel("",JLabel.CENTER);
            lbl3[2].setFont(new java.awt.Font("dialog", 0, 14));
            lbl3[2].setForeground(java.awt.Color.black);
            lbl3[2].setBorder(new Flush3DBorder());
            lbl3[2].setBounds(600,60,90,30);
            add(lbl3[2]);

            lbl3[3] = new JLabel("",JLabel.CENTER);
            lbl3[3].setFont(new java.awt.Font("dialog", 0, 14));
            lbl3[3].setForeground(java.awt.Color.black);
            lbl3[3].setBorder(new Flush3DBorder());
            lbl3[3].setBounds(750,60,80,30);
            add(lbl3[3]);

            lbl3[4] = new JLabel("",JLabel.CENTER);
            lbl3[4].setFont(new java.awt.Font("dialog", 0, 14));
            lbl3[4].setForeground(java.awt.Color.black);
            lbl3[4].setBorder(new Flush3DBorder());
            lbl3[4].setBounds(900,60,80,30);
            add(lbl3[4]);

            lbl3[5] = new JLabel("",JLabel.CENTER);
            lbl3[5].setFont(new java.awt.Font("dialog", 0, 16));
            lbl3[5].setForeground(java.awt.Color.black);
            lbl3[5].setBorder(new Flush3DBorder());
            lbl3[5].setBounds(60,95,90,30);
            add(lbl3[5]);

            lbl3[6] = new JLabel("",JLabel.CENTER);
            lbl3[6].setFont(new java.awt.Font("dialog", 0, 16));
            lbl3[6].setForeground(java.awt.Color.black);
            lbl3[6].setBorder(new Flush3DBorder());
            lbl3[6].setBounds(190,95,90,30);
            add(lbl3[6]);

            lbl3[7] = new JLabel("",JLabel.CENTER);
            lbl3[7].setFont(new java.awt.Font("dialog", 0, 16));
            lbl3[7].setForeground(java.awt.Color.black);
            lbl3[7].setBorder(new Flush3DBorder());
            lbl3[7].setBounds(320,95,90,30);
            add(lbl3[7]);

            lbl3[8] = new JLabel("",JLabel.CENTER);
            lbl3[8].setFont(new java.awt.Font("dialog", 0, 16));
            lbl3[8].setForeground(java.awt.Color.black);
            lbl3[8].setBorder(new Flush3DBorder());
            lbl3[8].setBounds(450,95,90,30);
            add(lbl3[8]);

            lbl3[9] = new JLabel("",JLabel.CENTER);
            lbl3[9].setFont(new java.awt.Font("dialog", 0, 16));
            lbl3[9].setForeground(java.awt.Color.black);
            lbl3[9].setBorder(new Flush3DBorder());
            lbl3[9].setBounds(580,95,90,30);
            add(lbl3[9]);

            lbl3[10] = new JLabel("",JLabel.CENTER);
            lbl3[10].setFont(new java.awt.Font("dialog", 0, 16));
            lbl3[10].setForeground(java.awt.Color.black);
            lbl3[10].setBorder(new Flush3DBorder());
            lbl3[10].setBounds(710,95,90,30);
            add(lbl3[10]);

            lbl3[11] = new JLabel("",JLabel.CENTER);
            lbl3[11].setFont(new java.awt.Font("dialog", 0, 16));
            lbl3[11].setForeground(java.awt.Color.black);
            lbl3[11].setBorder(new Flush3DBorder());
            lbl3[11].setBounds(890,95,90,30);
            add(lbl3[11]);

        }

        /**
        * バッチ情報を設定する。
        */
        public void setBtCondition() {

CZSystem.log("CZTPGFrame", "炉は？"+roDbName_);
CZSystem.log("CZTPGFrame", "バッチは？"+SelectBt);
CZSystem.log("CZTPGFrame", "時間は？"+SelectTime);
			roBtTempCondition_ = CZSystem.getHikiageTemp(roDbName_,SelectBt,SelectTime);

//            if (null != roBtAllCondition_){
            if (null != roBtTempCondition_){
				CZSystemBtTemp bt = (CZSystemBtTemp)roBtTempCondition_.elementAt(0);

//                CZSystemBt bt = (CZSystemBt)roBtAllCondition_.elementAt(0);
                lbl4[1].setText(new Integer(SelectNo).toString());
                lbl3[0].setText((bt.t_time).trim());
                lbl3[1].setText((bt.batch).trim());
                lbl3[2].setText((bt.hinshu).trim());
//@@@                lbl3[3].setText((bt.pgid).trim());
                lbl3[3].setText(CZSystem.getProcName(roBtStart_.p_no));     //@@@ プロセス
//@@@@                lbl3[4].setText(new Integer(bt.i_sikomi).toString());
                lbl3[4].setText(new Integer(bt.i_sikomi + bt.t_sikomi).toString());     //@@@@
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
	//@@@                lbl3[3].setText((bt.pgid).trim());
	                lbl3[3].setText(CZSystem.getProcName(roBtStart_.p_no));     //@@@ プロセス
	//@@@@                lbl3[4].setText(new Integer(bt.i_sikomi).toString());
	                lbl3[4].setText(new Integer(bt.i_sikomi + bt.t_sikomi).toString());     //@@@@
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
        
        /**
        * バッチ情報をクリアする。
        */
        public void clearBtCondition() {
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
            SelectNo = 0;
            roDbName_= null;
            SelectBt= null;
            SelectTime= null;
            pvDataBody_= null;
            selList_= null;
            roBtStart_= null;
		}
    } //TitlePanel

    //==========================================================================
    /**
     *   ＰＶ選択Panel
     */
    //==========================================================================
    class SelectPanel extends JPanel {  

        /**
        * コンストラクタ
        */
        SelectPanel(){
            super();

            setLayout( null );
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            JPanel pnl1 = new JPanel();
            pnl1.setLayout( null );
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                pnl1.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                pnl1.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            JLabel lbl = null;
            lbl = new JLabel("色");
            lbl.setFont(new java.awt.Font("dialog", 0, 16));
            lbl.setForeground(Color.black);
            lbl.setBounds(10,20,40,18);
            pnl1.add(lbl);

            lbl = new JLabel("項目No");
            lbl.setFont(new java.awt.Font("dialog", 0, 16));
            lbl.setForeground(Color.black);
            lbl.setBounds(41,20,200,18);
            pnl1.add(lbl);

            lbl = new JLabel("線の太さ");
            lbl.setFont(new java.awt.Font("dialog", 0, 16));
            lbl.setForeground(Color.black);
            lbl.setBounds(220,20,200,18);
            pnl1.add(lbl);

            lbl = new JLabel("レンジ");
            lbl.setFont(new java.awt.Font("dialog", 0, 16));
            lbl.setForeground(Color.black);
            lbl.setBounds(350,0,100,18);
            pnl1.add(lbl);

            lbl = new JLabel("Ｍｉｎ");
            lbl.setFont(new java.awt.Font("dialog", 0, 16));
            lbl.setForeground(Color.black);
            lbl.setBounds(310,20,100,18);
            pnl1.add(lbl);

            lbl = new JLabel("Ｍａｘ");
            lbl.setFont(new java.awt.Font("dialog", 0, 16));
            lbl.setForeground(Color.black);
            lbl.setBounds(380,20,100,18);
            pnl1.add(lbl);

            try {
                int yPos = 0;
                pnl1.setBounds(30,yPos,470,40);
                add( pnl1 );
                //
                pnlSel_ = new SelectItemPanel[14];

                int itemHeight = 35;
                yPos = 41;
                for(int i=0; i < 14; i++ ){
                CZSystem.log("CZTPGFrame","TPG Property START" + i);

                    if (prop_yNo[i].equals("") == false) {

                        String rgb[] = new String[3];
                        int cCount =0;
                        StringTokenizer st = new StringTokenizer(prop_yCol[i],",");
                        while (st.hasMoreTokens()) {
                            if (3 == cCount) break;
                            rgb[cCount] = new String((st.nextToken()).trim());
                            cCount++;
                        }
                        if (3 > cCount) {
                            CZSystem.log("TPGFrame SelectPanel","***** TPG Propaty File Error !! *****");
                            System.exit(-1);
                        }
                        Color col = new Color(Integer.parseInt(rgb[0]),Integer.parseInt(rgb[1]),Integer.parseInt(rgb[2]));
                        pnlSel_[i] = new SelectItemPanel(i+1, col);
                        pnlSel_[i].setLayout( null );
                        // 他基地参照機能    @20131021
                        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                            pnlSel_[i].setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
                        }else{
                            pnlSel_[i].setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
                        }
                        pnlSel_[i].setBounds(30,yPos,470,itemHeight);
                CZSystem.log("CZTPGFrame","TPG Property Chno" + i);
                        pnlSel_[i].setChNo(prop_yNo[i]);
                CZSystem.log("CZTPGFrame","TPG Property Chno END" + i);
                        pnlSel_[i].setLineS(prop_yLine[i]);
                        pnlSel_[i].setMin(Float.parseFloat(prop_yMin[i]));
                        pnlSel_[i].setMax(Float.parseFloat(prop_yMax[i]));
                        CZSystemPVName n = (CZSystemPVName)dataTbl_.get(new Integer(prop_yNo[i]));
                        if (null != n) {
                            pnlSel_[i].setName( n.k_name);
                        } else {
                        }
                    } else {
                CZSystem.log("CZTPGFrame","TPG Property Null" + i);
                        pnlSel_[i] = new SelectItemPanel(i+1, java.awt.Color.white);
                        pnlSel_[i].setLayout( null );
                        // 他基地参照機能    @20131021
                        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                            pnlSel_[i].setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
                        }else{
                            pnlSel_[i].setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
                        }
                        pnlSel_[i].setBounds(30,yPos,470,itemHeight);
                        pnlSel_[i].setChNo("");
                        pnlSel_[i].setLineS("1");
                        pnlSel_[i].setName( new String(""));
                        pnlSel_[i].setMin(0.0f);
                        pnlSel_[i].setMax(10.0f);
                CZSystem.log("CZTPGFrame","TPG Property null END" + i);
                    }

                    add(pnlSel_[i]);
                    yPos = yPos + itemHeight;
                CZSystem.log("CZTPGFrame","TPG Property END" + i);
                }
            } catch ( Exception e) {
                CZSystem.log("CZTPGFrame","TPG Property File Error");
                System.exit(-1);
            }
        }

        public SelectItemPanel[] getIitemPanel(){
            return pnlSel_;
        }
    } //SelectPanel

    //==========================================================================
    /**
     *   X軸の設定をするPanel
     */
    //==========================================================================
    class XLengeSetPanel extends JPanel {   

        private JTextFieldInt txtUnit_  = null;
        private JTextFieldInt txtStart_ = null;
        private JTextFieldInt txtEnd_   = null;
        private JTextFieldInt txtMesh_  = null;

        /**
        *
        */
        XLengeSetPanel(){
            super();

            setLayout(null);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            JLabel lbl = null;
            lbl = new JLabel("表示単位(1:分    2:mm)");
            lbl.setFont(new java.awt.Font("dialog", 0, 16));
            lbl.setForeground(Color.black);
            lbl.setBounds(10,10,200,18);
            add(lbl);

            txtUnit_ = new JTextFieldInt();
            txtUnit_.setMinValue(1);
            txtUnit_.setMaxValue(2);
            if( null != prop_xUnit ) {
                txtUnit_.setValue(Integer.parseInt(prop_xUnit));
            } else {
                txtUnit_.setValue(2);
            }
            txtUnit_.setBounds(240,10,80,18);
            add(txtUnit_);

            lbl = new JLabel("表示範囲");
            lbl.setFont(new java.awt.Font("dialog", 0, 16));
            lbl.setForeground(Color.black);
            lbl.setBounds(10,40,120,18);
            add(lbl);

            lbl = new JLabel("スタート");
            lbl.setFont(new java.awt.Font("dialog", 0, 16));
            lbl.setForeground(Color.black);
            lbl.setBounds(80,40,120,18);
            add(lbl);

            txtStart_ = new JTextFieldInt();
            txtStart_.setMinValue(0);
            txtStart_.setMaxValue(10000);	// @20131028 TPGｸﾞﾗﾌ入力最大数変更
            if ( null != prop_xMin ) {
                txtStart_.setValue(Integer.parseInt(prop_xMin));
            } else {
                txtStart_.setValue(0);
            }
            txtStart_.setBounds(240,40,80,18);
            add(txtStart_);

            lbl = new JLabel("エンド");
            lbl.setFont(new java.awt.Font("dialog", 0, 16));
            lbl.setForeground(Color.black);
            lbl.setBounds(80,60,120,18);
            add(lbl);

            txtEnd_ = new JTextFieldInt();
            txtEnd_.setMinValue(0);
            txtEnd_.setMaxValue(10000);		// @20131028 TPGｸﾞﾗﾌ入力最大数変更
            if ( null != prop_xMax ) {
                txtEnd_.setValue(Integer.parseInt(prop_xMax));
            } else {
                txtEnd_.setValue(10000);	// @20131028 TPGｸﾞﾗﾌ入力最大数変更
            }
            txtEnd_.setBounds(240,60,80,18);
            add(txtEnd_);

            lbl = new JLabel("メッシュ間隔数");
            lbl.setFont(new java.awt.Font("dialog", 0, 16));
            lbl.setForeground(Color.black);
            lbl.setBounds(10,90,200,18);
            add(lbl);

            txtMesh_ = new JTextFieldInt();
            txtMesh_.setMinValue(2);
// 20031201 henkou
//            txtMesh_.setMaxValue(20);
            txtMesh_.setMaxValue(30);
            if ( null != prop_xBun ) {
                txtMesh_.setValue(Integer.parseInt(prop_xBun));
            } else {
                txtMesh_.setValue(10);
            }
            txtMesh_.setBounds(240,90,80,18);
            add(txtMesh_);

        }
        /**
         * @return 単位
         */
        public int getUnit(){
            return txtUnit_.getValue();
        }
        /**
         * @return 開始値
         */
        public int getStart(){
            return txtStart_.getValue();
        }
        /**
         * @return 終了値
         */
        public int getEnd(){
            return txtEnd_.getValue();
        }
        /**
         * @return Ｘ軸分割数
         */
        public int getMesh(){
            return txtMesh_.getValue();
        }
    } //XLengeSetPanel

    //==========================================================================
    /**
     *   PV項目一覧表示Panel
     */
    //==========================================================================
    class PVIchiranPanel extends JPanel {

        /**
        * コンストラクタ
        */
        PVIchiranPanel(){
            super();

            setLayout(null);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            JLabel lbl1 = new JLabel("＜＜  ＰＶ項目Ｎｏ一覧  ＞＞",JLabel.CENTER);
            lbl1.setLayout(new FlowLayout(FlowLayout.CENTER));
            lbl1.setFont(new java.awt.Font("dialog", 0, 20));
            lbl1.setForeground(Color.black);
            lbl1.setBounds(0,0,470,26);

            JPanel pnl1 = new JPanel();
            pnl1.setLayout(null);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                pnl1.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                pnl1.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }
            pnl1.setBounds(0,0,500,40);
            pnl1.add(lbl1);
            add(pnl1);

            // ＰＶ名称一覧表を取得し画面へ表示する。
            PvNameTable t = new PvNameTable((Vector) CZSystem.getPVNameAll());
            JTableHeader tabHead = t.getTableHeader();
            tabHead.setReorderingAllowed(false);
            JScrollPane pvNameScpanel = new JScrollPane();
            pvNameScpanel.setBounds(0, 0, 460, 240);
            pvNameScpanel.setViewportView(t);

            JPanel pnl2 = new JPanel();
            pnl2.setLayout(null);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                pnl2.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                pnl2.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }
            pnl2.setBounds(0,41,500,240);
            pnl2.add(pvNameScpanel);
            add(pnl2);

        }
    } //PVIchiranPanel

    //==========================================================================
    /**
     *   グラフ表示項目選択Panel
     */
    //==========================================================================
    class SelectItemPanel extends JPanel {  

        private int             panelNo_    = 0;        // No
        private Color           col_        = java.awt.Color.gray;
        private NumText         txtCh_      = null;     // PV項目No
        private LineSize        lineS_      = null;     // 線の太さ
        private JTextFieldFloat txtMin_     = null;     // Min値
        private JTextFieldFloat txtMax_     = null;     // Max値
        private JLabel          lblName_    = null;     // PV名
        private JButton         btnCol_     = null;

        /**
        * コンストラクタ
        */
        SelectItemPanel(int no, Color c){

            super();

            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }
            panelNo_ = no;
            col_ = c;

            Integer intNo = new Integer(panelNo_);
            btnCol_= new JButton(new String(intNo.toString()));
            btnCol_.setLayout(new FlowLayout(FlowLayout.CENTER));
            btnCol_.setFont(new java.awt.Font("dialog", 0, 8));
            btnCol_.setForeground(col_);
            btnCol_.setBackground(col_);
            btnCol_.setBounds(5,0,30,26);
            btnCol_.addActionListener(
                new ActionListener() {
                    public void actionPerformed(ActionEvent ev){
                        JButton but = (JButton)ev.getSource();
                        Color c = JColorChooser.showDialog(null,"色を選んでください", but.getBackground());
                        if(null != c){
                            col_ = c;
                            but.setForeground(c);           //選択した色を設定する。
                            but.setBackground(c);           //
                        }
                    }
                }
            );
            add(btnCol_);

            txtCh_ = new NumText();
            txtCh_.setBounds(40,4,40,20);
            txtCh_.addActionListener(
                new ActionListener() {
                    public void actionPerformed(ActionEvent e) {
                        String str = txtCh_.getText().trim();
                        if (str.equals("")) {
                            lblName_.setText( new String("") );
                            txtMin_.setDefaultValue();
                            txtMax_.setDefaultValue();
                            return;
                        }
                        int val = 0;
                        try {
                            val = Integer.parseInt(str);
                            
                            CZSystemPVName n = (CZSystemPVName)dataTbl_.get(new Integer(val));
                            if (null != n) {
                                lblName_.setText( n.k_name);
                                txtMin_.setValue((float)n.n_min);
                                txtMax_.setValue((float)n.n_max);
                            } else {
                                Object msg[] = { "入力値( " + val + " )が無効です。！！", "", "" };
                                errorMsg(msg);
                                txtCh_.setText("");
                            }
                        } catch (NumberFormatException ex) {
                            return;
                        }
                    }
                }
            );

            txtCh_.addFocusListener(
                new FocusAdapter() {
                    public void focusGained(FocusEvent e) {
                        txtCh_.selectAll();
                    }

                    public void focusLost(FocusEvent e) {
                        String str = txtCh_.getText().trim();
                        if (str.equals("")) {
                            lblName_.setText( new String("") );
                            txtMin_.setDefaultValue();
                            txtMax_.setDefaultValue();
                            return;
                        }
                        int val = 0;
                        try {
                            val = Integer.parseInt(str);
                            CZSystemPVName n = (CZSystemPVName)dataTbl_.get(new Integer(val));
                            if (null != n) {
                                lblName_.setText( n.k_name);
                                txtMin_.setValue((float)n.n_min);
                                txtMax_.setValue((float)n.n_max);
                            } else {
                                Object msg[] = { "入力値( " + val + " )が無効です。！！", "", "" };
                                errorMsg(msg);
                                txtCh_.setText("");
                            }
                        } catch (NumberFormatException ex) {
                            return;
                        }
                    }
                }
            );

            add(txtCh_);

            lineS_ = new LineSize();
            lineS_.setBounds(230,4,40,20);

            lineS_.addActionListener(
                new ActionListener() {
                    public void actionPerformed(ActionEvent e) {
                        String str = lineS_.getText().trim();
                        if (str.equals("")) {
                            lineS_.setText( new String("") );
                            return;
                        }
                        int size = 0;
                        try {
                            size = str.length();
                            
                            if (size == 1) {
                                lineS_.setText(str);
                            } else {
                                Object msg[] = { "入力値( " + size + " )が無効です。！！", "", "" };
                                errorMsg(msg);
                                lineS_.setText("");
                            }
                        } catch (NumberFormatException ex) {
                            return;
                        }
                    }
                }
            );

            lineS_.addFocusListener(
                new FocusAdapter() {
                    public void focusGained(FocusEvent e) {
                        lineS_.selectAll();
                    }

                    public void focusLost(FocusEvent e) {
                        String str = lineS_.getText().trim();
                        if (str.equals("")) {
                            lineS_.setText( new String("") );
                            return;
                        }
                        int size = 0;
                        try {
                            size = str.length();

                            if (size == 1) {
                                lineS_.setText(str);
                            } else {
                                Object msg[] = { "入力値( " + size + " )が無効です。！！", "", "" };
                                errorMsg(msg);
                                lineS_.setText("");
                            }
                        } catch (NumberFormatException ex) {
                            return;
                        }
                    }
                }
            );

            add(lineS_);

            lblName_ = new JLabel(new String(""));
            lblName_.setFont(new java.awt.Font("dialog", 0, 18));
            lblName_.setForeground(Color.black);
            lblName_.setBounds(80,4,230,20);
            add(lblName_);

            txtMin_ = new JTextFieldFloat();
            txtMin_.setBounds(300,4,80,20);
            add(txtMin_);

            txtMax_ = new JTextFieldFloat();
            txtMax_.setBounds(380,4,80,20);
            add(txtMax_);

        }

        //----------------------------------------------------------------------
        /**
         *  @param msg ... メッセージ内容
         *  @return true ... OK, false ... NG
         */
        //----------------------------------------------------------------------
        private boolean errorMsg(Object msg[]){
            JOptionPane.showMessageDialog(null,msg,
                                    "入力エラー",
                                    JOptionPane.ERROR_MESSAGE);
            return true;
        }

        /**
         * 色を設定する。
         */
        public void setColor(Color c){
            col_ = c;
            btnCol_.setForeground(col_);
            btnCol_.setBackground(col_);
        }

        /**
         * 色を取得する。
         */
        public Color getColor(){
            return col_;
        }

        /**
         * パネルの№を取得する。
         */
        public int getNo(){
            return panelNo_;
        }

        /**
         * チャネルの№を取得する。
         */
        public String getChNo(){
            return txtCh_.getText();
        }

        /**
         * チャネルの№を設定する。
         */
        public void setChNo(String s){
            txtCh_.setText(s);
            return;
        }

        /**
         * 項目名を取得する。
         */
        public String getName(){
            return lblName_.getText();
        }

        /**
         * 項目名を設定する。
         */
        public void setName(String s){
            lblName_.setText(s);
            return;
        }

        /**
         * 線の太さを取得する。
         */
        public String getLineS(){
            return lineS_.getText();
        }

        /**
         * 線の太さを設定する。
         */
        public void setLineS(String s){
            lineS_.setText(s);
            return;
        }

        /**
         * 最大値を取得する。
         */
        public float getMax(){
            return txtMax_.getValue();
        }

        /**
         * 最大値を設定する。
         */
        public void setMax(float f){
            txtMax_.setValue(f);
            return;
        }

        /**
         * 最小値を取得する。
         */
        public float getMin(){
            return txtMin_.getValue();
        }

        /**
         * 最小値を設定する。
         */
        public void setMin(float f){
            txtMin_.setValue(f);
            return;
        }

        /**
         * 最大値、最小値にDefault値を設定する。
         */
        public void setDefault(){
            txtMin_.setDefaultValue();
            txtMax_.setDefaultValue();
        }

    } //SelectItemPanel

    //==========================================================================
    /**
     * float型の情報を保持するテキストフィールドクラス
     */
    //==========================================================================
    public class JTextFieldFloat extends JTextField {

        /**
         * 設定可能な最大値
         */
        private float max_ = Float.POSITIVE_INFINITY;
        /**
         * 設定可能な最小値
         */
        private float min_ = Float.NEGATIVE_INFINITY;
        /**
         * 保持する値
         */
        private float val_ = 0.0f;

        /**
         * コンストラクタ
         */
        JTextFieldFloat() {

            super();

            setFont(new java.awt.Font("dialog", 0, 16));
            setText("0.0");

            addActionListener(
              new ActionListener() {
                  public void actionPerformed(ActionEvent e) {
                      _textToValue();
                  }
              }
            );

            addFocusListener(
              new FocusAdapter() {
                  public void focusGained(FocusEvent e) {
                      selectAll();
                  }
                  public void focusLost(FocusEvent e) {
                      _textToValue();
                  }
              }
            );
        }

        /**
         * float値の設定
         * @param   val     設定する値
         */
        public void setValue(float val) {

            if ((min_ <= val) && (val <= max_)) {
                if (val_ != val) {
                    val_ = val;
                    setText("" + val);
                }
                else {
                    setText("" + val_);
                }
            }
            else {
                Object msg[] = { "入力値( " + val + " )が無効です。！！", "", "" };
                errorMsg(msg);
                setText("" + val_);
            }
        }

        /**
         * float値の取得
         * @return  float値
         */
        float getValue() {
            return val_;
        }

        /**
         * 最大値の設定
         * @param   max     最大値
         */
        public void setMaxValue(float max) {
            max_ = max;
        }

        /**
         * 最小値の設定
         * @param   min     最小値
         */
        public void setMinValue(float min) {
            min_ = min;
            setValue(min_);
        }

        /**
         * Default値の設定
         */
        public void setDefaultValue() {
            max_ = Float.POSITIVE_INFINITY;;
            min_ = Float.NEGATIVE_INFINITY;
            setValue(0.0f);
        }

        /**
         * String値をfloat値に変換
         */
        private void _textToValue() {
            String str = getText().trim();
            if (str.equals("")) {
                setValue(val_);
                return;
            }
            float val = 0.0f;
            try {
                val = Float.parseFloat(str);
            } catch (NumberFormatException ex) {
                setValue(val_);
                return;
            }
            setValue(val);
        }

        /**
         *  @param msg ... メッセージ内容
         *  @return true ... OK, false ... NG
         */
        private boolean errorMsg(Object msg[]){
            JOptionPane.showMessageDialog(null,msg,
                                    "入力エラー",
                                    JOptionPane.ERROR_MESSAGE);
            return true;
        }

        /**
         * createDefaultModel
         */
        protected Document createDefaultModel() {
            return new NumericDocument();
        }

        /**
         * NumericDocument Class
         */
        class NumericDocument extends PlainDocument {
            String validValues = "0123456789.-";

            public void insertString( int offset, String str, AttributeSet a )
                                        throws BadLocationException {

                char[] val = str.toCharArray();
                for (int i = 0;i < val.length;i++) {    
                    if(validValues.indexOf(val[i]) == -1) return;
                }

                super.insertString( offset, str, a );
                return ;
            }
        }
    }


    //======================================================================
    /**
     * Integer型の情報を保持するテキストフィールドクラス
     *
     */
    //======================================================================
    public class JTextFieldInt extends JTextField {

        /**
         * 設定可能な最大値
         */
        private int max_ = Integer.MAX_VALUE;
        /**
         * 設定可能な最小値
         */
        private int min_ = Integer.MIN_VALUE;
        /**
         * 保持する値
         */
        private int val_ = 0;

        /**
         * コンストラクタ
         */
        JTextFieldInt() {

            super();

            setFont(new java.awt.Font("dialog", 0, 16));
            setText("0");

            addActionListener(
                new ActionListener() {
                    public void actionPerformed(ActionEvent e) {
                        _textToValue();
                    }
                }
            );

            addFocusListener(
                new FocusAdapter() {
                    public void focusGained(FocusEvent e) {
                        selectAll();
                    }
                    public void focusLost(FocusEvent e) {
                        _textToValue();
                    }
                }
            );
        }

        /**
         * int値の設定
         * @param   val     設定する値
         */
        public void setValue(int val) {

            if ((min_ <= val) && (val <= max_)) {
                if (val_ != val) {
                    val_ = val;
                    setText("" + val);
                } else {
                    setText("" + val_);
                }
            } else {
                Object msg[] = { "入力値( " + val + " )が無効です。！！", "", "" };
                errorMsg(msg);
                setText("" + val_);
            }
        }

        /**
         * int値の取得
         * @return  int値
         */
        int getValue() {
            return val_;
        }

        /**
         * 最大値の設定
         * @param   max     最大値
         */
        public void setMaxValue(int max) {
            max_ = max;
        }

        /**
         * 最小値の設定
         * @param   min     最小値
         */
        public void setMinValue(int min) {
            min_ = min;
            setValue(min_);
        }

        /**
         * Stringで表現された値をfloat値に変換
         */
        private void _textToValue() {

            String str = getText().trim();
            if (str.equals("")) {
                setValue(val_);
                return;
            }
            int val = 0;
            try {
                val = Integer.parseInt(str);
            } catch (NumberFormatException ex) {
                setValue(val_);
                return;
            }
            setValue(val);
        }

        /**
         *  @param msg ... メッセージ内容
         *  @return true ... OK, false ... NG
         */
        private boolean errorMsg(Object msg[]){
            JOptionPane.showMessageDialog(null,msg,
                                    "入力エラー",
                                    JOptionPane.ERROR_MESSAGE);
            return true;
        }

        /**
         * createDefaultModel
         */
        protected Document createDefaultModel() {
            return new NumericDocument();
        }

        /**
         * NumericDocument Class
         */
        class NumericDocument extends PlainDocument {
            String validValues = "0123456789 -";

            public void insertString( int offset, String str, AttributeSet a )
                                        throws BadLocationException {

                char[] val = str.toCharArray();
                for (int i = 0;i < val.length;i++) {    
                    if(validValues.indexOf(val[i]) == -1) return;
                }

                super.insertString( offset, str, a );
                return ;
            }
        }
    }

    /**
     *       数値とスペースを受け付けるTextField
     */
    public class NumText extends JTextField {   

        /**
        * コンストラクタ
        */
        NumText(){
            super();
            setFont(new java.awt.Font("dialog", 0, 16));
        }

        /**
        * createDefaultModel
        */
        protected Document createDefaultModel() {
            return new NumericDocument();
        }

        /**
        * NumericDocument Class
        */
        class NumericDocument extends PlainDocument {
            String validValues = "0123456789.-";

            public void insertString( int offset, String str, AttributeSet a )
                                        throws BadLocationException {

                char[] val = str.toCharArray();
                for (int i = 0;i < val.length;i++) {    
                    if(validValues.indexOf(val[i]) == -1) return;
                }

                super.insertString( offset, str, a );
                return ;
            }
        }
    } //NumText

    /**
     *       数値とスペースを受け付けるTextField
     */
    public class LineSize extends JTextField {   

        /**
        * コンストラクタ
        */
        LineSize(){
            super();
            setFont(new java.awt.Font("dialog", 0, 16));
        }

        /**
        * createDefaultModel
        */
        protected Document createDefaultModel() {
            return new NumericDocument();
        }

        /**
        * NumericDocument Class
        */
        class NumericDocument extends PlainDocument {
            String validValues = "12345";

            public void insertString( int offset, String str, AttributeSet a )
                                        throws BadLocationException {

                if(0 < getLength()) return;
                char[] val = str.toCharArray();
                for (int i = 0;i < val.length;i++) {    
                    if(validValues.indexOf(val[i]) == -1) return;
                }

                super.insertString( offset, str, a );
                return ;
            }

        }
    } //LineSize

    //==========================================================================
    /**
     *       PV項目一覧
     */
    //==========================================================================
    class PvNameTable extends JTable {

        private Vector  pvNameList_ = null;
        private pvNameTblMdl model_ = null;

        /**
        * コンストラクタ
        * @param v ... PV項目名
        */
        PvNameTable(Vector v){
            super();

            pvNameList_ = v;

            try{
                setName("PvNameTable");
                setAutoCreateColumnsFromModel(true);
                setBounds(0, 0, 440, 300);
                setLocale(new Locale("ja","JP"));
                setFont(new java.awt.Font("dialog", 0, 12));
                setRowHeight(17);

                model_ = new pvNameTblMdl(pvNameList_);
                setModel(model_);

                DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                TableColumn colum = null;

                // CH
                colum = cmdl.getColumn(0);
                colum.setMaxWidth(30);
                colum.setMinWidth(30);
                colum.setWidth(30);
                // 項目名
                colum = cmdl.getColumn(1);
                colum.setMaxWidth(80);
                colum.setMinWidth(80);
                colum.setWidth(80);
                // 日本語名称
                colum = cmdl.getColumn(2);
                colum.setMaxWidth(230);
                colum.setMinWidth(230);
                colum.setWidth(230);
                // Min
                colum = cmdl.getColumn(3);
                colum.setMaxWidth(50);
                colum.setMinWidth(50);
                colum.setWidth(50);
                // Max
                colum = cmdl.getColumn(4);
                colum.setMaxWidth(50);
                colum.setMinWidth(50);
                colum.setWidth(50);
            } catch (Throwable e) {
                CZSystem.handleException(e);
            }
        }

        //======================================================================
        /**
         *       PV項目一覧:モデル
         */
        //======================================================================
        public class pvNameTblMdl extends AbstractTableModel {

            private int     TBL_ROW     = 128;      // 行数
            final   int     TBL_COL     = 5;        // 列数
            private Vector  pvNameList_ = null;     // バッチ情報

            final String[] names = {" CH "  , "項目", "名称", "Min", "Max"};
            private Object  data[][];

            /**
            * コンストラクタ 
            * @param v ... PV項目
            */
			@SuppressWarnings("unchecked")
            pvNameTblMdl(Vector v){
                super();

                pvNameList_ = v;
                TBL_ROW = pvNameList_.size();

                data = new Object[TBL_ROW][TBL_COL];
                dataTbl_ = new Hashtable();
                for(int i = 0 ; i < TBL_ROW ; i++){

                    CZSystemPVName pvName = (CZSystemPVName)pvNameList_.elementAt(i);
                    if(null == pvName) break;
                    data[i][0]  = new Integer(pvName.k_no);     //No
                    data[i][1]  = new String(pvName.k_name);    //名称
                    data[i][2]  = new String(pvName.j_name.trim());    //日本語名称
                    data[i][3]  = new Integer(pvName.n_min);    //Min
                    data[i][4]  = new Integer(pvName.n_max);    //Max
                    dataTbl_.put((Object)(data[i][0]), (Object)pvName);
                }
            }

            /**
            * 桁数を取得する。
            * @return ... 桁数
            */
            public int getColumnCount(){
                return TBL_COL;
            }
            /**
            * 行数を取得する。
            * @return ... 行数
            */
            public int getRowCount(){
                return TBL_ROW;
            }
            /**
            * データを取得する。
            * @param ... row:行, col:桁
            * @return ... データ
            */
            public Object getValueAt(int row, int col){
                return data[row][col];
            }
            /**
            * 桁名を取得する。
            * @param ... column:桁
            * @return ... 桁名
            */
            public String getColumnName(int column){
                return names[column];
            }
            /**
            * データの型を取得する。
            * @param ... c:桁
            * @return ... データの型
            */
            public Class getColumnClass(int c){
                return getValueAt(0, c).getClass();
            }
            /**
            * cell編集の可否を取得する。
            * @param ... row:行, col:桁
            * @return ... 桁数
            */
            public boolean isCellEditable(int row, int col){
                return false;
            }
            /**
            * データを設定する。
            * @param ... aValue:データ, row:行, col:桁
            * @return ... 桁数
            */
            public void setValueAt(Object aValue, int row, int column){
                data[row][column] = aValue;
            }
        } // pvNameTblMdl
    } // PvNameTable


    /***************************************************
     *
     * 引き上げ条件Dialog
     *
     ***************************************************/
    class BtConditionDialog extends JDialog {

        /**
        * コンストラクタ
        */
        BtConditionDialog(){
            super();

            setTitle("引き上げ条件");
            setSize(820,240);
            setResizable(false);
            setModal(true);
            getContentPane().setLayout(null);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            BtConditionTable t = new BtConditionTable(roBtAllCondition_);
            JTableHeader tabHead = t.getTableHeader();
            tabHead.setReorderingAllowed(false);

            JScrollPane bt_scpanel = new JScrollPane(t);
            bt_scpanel.setBounds(20, 20, 780, 187);
            getContentPane().add(bt_scpanel);

        }

        /**
         * Ｂｔ登録情報一覧
         * @@T6追加
         */
        class BtConditionTable extends JTable {

            private Vector  bt_list     = null;
            private BtConditionTblMdl model = null;

            /**
            * コンストラクタ
            * @param v ... バッチ情報
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

            /**
            * Change Listener
            */
            public void valueChanged(ListSelectionEvent e){
                super.valueChanged(e);
            }
            /**
            *
            */
            public void setData(int gr,int tbl){
            }

            /**
             * Ｂｔ登録情報一覧：モデル
             */
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

                /**
                * コンストラクタ
                * @param v ... バッチ情報
                */
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
                        data[i][13] = new Integer(bt.no_teisu);     //T6 @@
                        data[i][14] = new Integer(bt.pno_start);    //PNo
                        data[i][15] = new Integer(bt.p_kaisi);      //開始
                    }

                }
                /**
                * 桁数を取得する。
                * @return ... 桁数
                */
                public int getColumnCount(){
                    return TBL_COL;
                }
                /**
                * 行数を取得する。
                * @return ... 行数
                */
                public int getRowCount(){
                    return TBL_ROW;
                }
                /**
                * データを取得する。
                * @param ... row:行, col:桁
                * @return ... データ
                */
                public Object getValueAt(int row, int col){
                    return data[row][col];
                }
                /**
                * 桁名を取得する。
                * @param ... column:桁
                * @return ... 桁名
                */
                public String getColumnName(int column){
                    return names[column];
                }
                /**
                * データの型を取得する。
                * @param ... c:桁
                * @return ... データの型
                */
                public Class getColumnClass(int c){
                    return getValueAt(0, c).getClass();
                }
                /**
                * cell編集の可否を取得する。
                * @param ... row:行, col:桁
                * @return ... 桁数
                */
                public boolean isCellEditable(int row, int col){
                    return false;
                }
                /**
                * データを設定する。
                * @param ... aValue:データ, row:行, col:桁
                * @return ... 桁数
                */
                public void setValueAt(Object aValue, int row, int column){
                    data[row][column] = aValue;
                }
            } // BtConditionTblMdl
        } // BtConditionTable
    } // BtConditionDialog

//==============================================================================
    /**
     *
     * 検索Dialog
     *
     */
//==============================================================================
    class SercheDialog extends JDialog {

        private JScrollPane scpnlBt       = null;
        private JScrollPane scpnlBtStart  = null;
        private JButton     btnRead       = null;
        private JLabel      roNameLab     = null;

        /**
        * コンストラクタ
        */
        SercheDialog(){
            super();

//            setTitle("SercheDialog");
            setTitle("検 索");
//@@@@@            setSize(820,335);
            setSize(940,335);    //@@@@@
            setResizable(false);
            setModal(true);
            getContentPane().setLayout(null);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

			String s = CZSystem.RoKetaChg(roName_);	// 20050725 炉：表示桁数変更
            roNameLab = new JLabel(s,JLabel.CENTER);
            //roNameLab = new JLabel(roName_,JLabel.CENTER);
            roNameLab.setBounds(20, 20, 100, 30);
            roNameLab.setLocale(new Locale("ja","JP"));
            roNameLab.setFont(new java.awt.Font("dialog", 0, 18));
            roNameLab.setBorder(new Flush3DBorder());
            roNameLab.setForeground(java.awt.Color.black);
            getContentPane().add(roNameLab);

            scpnlBt = new JScrollPane();
//@@@@@            scpnlBt.setBounds(20, 60, 350, 187);
            scpnlBt.setBounds(20, 60, 470, 187);    //@@@@@
            getContentPane().add(scpnlBt);

            scpnlBtStart = new JScrollPane();
//            scpnlBtStart.setBounds(390, 60, 410, 187);
            scpnlBtStart.setBounds(510, 60, 410, 187);    //@@@@@
            getContentPane().add(scpnlBtStart);

            btnRead = new JButton("読み込み");
//@@@@@            btnRead.setBounds(700, 270, 100, 24);
            btnRead.setBounds(820, 270, 100, 24);    //@@@@@
            btnRead.setLocale(new Locale("ja","JP"));
            btnRead.setFont(new java.awt.Font("dialog", 0, 18));
            btnRead.setBorder(new Flush3DBorder());
            btnRead.setForeground(java.awt.Color.black);
            btnRead.addActionListener(
                new ActionListener() {
                    public void actionPerformed(ActionEvent ev){
                        Cursor cu_tmp = getCur();
                        Cursor cu = new Cursor(Cursor.WAIT_CURSOR);
                        setCur(cu);
                        int ret = readBtPV();
                        setCur(cu_tmp);
                        if(1 > ret){
                            return;
                        }
                        setVisible(false);
                        //  引上げ情報を表示する。@@
                        pnl1_.setBtCondition();
                    }
                }
            );
            btnRead.setEnabled(false);
            getContentPane().add(btnRead);

        }

        /**
         * バッチ情報を表示する。
         * @return true
        */
        public boolean setDefault(){

            removeBtStart();
            removeBtCondition();

			String s = CZSystem.RoKetaChg(roName_);	// 20050725 炉：表示桁数変更
            roNameLab.setText(s);
            //roNameLab.setText(roName_);
            BtTable t = new BtTable();
            JTableHeader tabHead = t.getTableHeader();
            tabHead.setReorderingAllowed(false);
            scpnlBt.setViewportView(t);
            btnRead.setEnabled(false);

//@20131017
CZSystem.log("CZTPGFrame BtTable スクロールバー初回位置決め処理","位置：" + SelBtRow);
                JScrollBar bt_jsb = scpnlBt.getVerticalScrollBar();
                bt_jsb.setValue((SelBtRow*17)-102);
                scpnlBt.setVerticalScrollBar(bt_jsb);
//@20131017
            return true;
        }
        /**
         * バッチ情報を設定する。
         * @param v ... 
         * @return true
        */
        public boolean setBtCondition(Vector v){

            removeBtCondition();
            BtStartTable t = new BtStartTable(v);
            JTableHeader tabHead = t.getTableHeader();
            tabHead.setReorderingAllowed(false);
            scpnlBtStart.setViewportView(t);
            roBtAllCondition_ = v;
            return true;
        }
        /**
         * バッチ情報を削除する。
         */
        public boolean removeBtCondition(){

            JViewport v;
            v =  scpnlBtStart.getViewport();
            if(null != v.getView()) v.remove(v.getView());
            removeBtStart();
            btnRead.setEnabled(false);
            return true;
        }
        /**
        * バッチ開始時刻を設定する
        * @param st ... バッチ開始時刻
        * @return true ... OK, false ... NG
        */
        public boolean setBtStart(CZSystemStart st){

            roBtStart_ = st;
            if(null == roBtStart_) return false;
            return true;
        }
        /**
        * 設定済みバッチ開始時刻を削除する
        * @return true ... OK
        */
        public boolean removeBtStart(){

            roBtStart_ = null;
            return true;
        }
        /**
        * カーソルを設定する。
        */
        private void setCur(Cursor cu){
            setCursor(cu);
        }
        /**
        * カーソルを取得する。
        */
        private Cursor getCur(){
            return getCursor();
        }
        /**
        * ＴＰＧエラーメッセージ表示Dialog
        * @param msg ... メッセージ内容
        * @return true ... OK, false ... NG
        */
        private boolean errorMsg(Object msg[]){
            JOptionPane.showMessageDialog(null,msg,
                                    "ＴＰＧエラー",
                                    JOptionPane.ERROR_MESSAGE);
            return true;
        }

        //======================================================================
        /**
        * PVデータを読み込む
        * @return ... 実績の読込み件数
        *（-1 ... スタート実績無し,-2 ... 表無し,-4 ... 実績無し）
        */
        //======================================================================
        public int readBtPV(){

            if(null == roBtStart_){
                Object msg[] = { "スタート実績が有りません！！", "", "" };
                errorMsg(msg);
                return -1;
            }

            CZSystemStart st = roBtStart_;              //バッチ開始情報を保持する。
            //バッチ開始情報からDBテーブル名を取得する。
            String view = CZSystem.getViewName(roDbName_,st.batch);
            if(null == view){
                Object msg[] = {"表が存在しません！！", view, ""};
                errorMsg(msg);
                return -2;
            }
            // 読み出すデータを設定する。
            boolean dataNo[] = null;
            dataNo = new boolean[CZSystemDefine.PV_MAX_LENGTH];
            // 読出しフラグをクリアする。
            for(int i = 0 ; i < CZSystemDefine.PV_MAX_LENGTH ; i++){
                dataNo[i] = false;
            }
            // 読出しフラグを設定する。
            for (int i = 0; i < selList_.size(); i++ ){
                SelectItemPanel item = (SelectItemPanel)selList_.elementAt(i);
                if (!(item.getChNo().equals(""))){
                    dataNo[(new Integer(item.getChNo()).intValue()) - 1] = true;
                }
            }
            //PVデータ読み込み

            CZSystem.log("CZTPGFrame","バッチNo"+st.batch);
            CZSystem.log("CZTPGFrame","開始時間"+st.p_start);

            SelectBt = st.batch;
            SelectTime = st.p_start;

/***** System.gc() *****/
//            System.out.println(Runtime.getRuntime().freeMemory());
            System.gc();
//            System.out.println(Runtime.getRuntime().freeMemory() + "  GC FREE!!");
/**********************/


            pvDataBody_ = null;
            pvDataBody_ = CZSystem.getPVData(roDbName_, view, st.p_renban, dataNo);

/***** System.gc() *****/
//            System.out.println(Runtime.getRuntime().freeMemory());
            System.gc();
//            System.out.println(Runtime.getRuntime().freeMemory() + "  GC FREE!!");
/**********************/

            if(1 > pvDataBody_.size()){
                Object msg[] = {"実績が有りません！！",
                                "[" + pvDataBody_.size() + "]",
                                ""};
                errorMsg(msg);
                pvDataBody_ = null;
                return -4;
            }
            return pvDataBody_.size();
        }
        /**
         * バッチ№の一覧を表示する。
         */
        class BtTable extends JTable {

            private Vector  btAllList   = null;
            private Vector  btList      = null;
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

                    btAllList = CZSystem.getBtCondition(roDbName_);
                    if(null == btAllList) return;

                    btList = new Vector();

                    for(int i = 0 ; i < btAllList.size() ; i++){
                        CZSystemBt bt = (CZSystemBt)btAllList.elementAt(i);

                        if(0 == bt.renban) btList.addElement(bt);
//@@2003.09.18                        if(-1 == bt.renban) btList.addElement(bt);
                    }

                    model = new BtTblMdl(btList);
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
//@@@@@
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
//@@@@@

//@20131017 バッチ番号保持機能
                    int row = SelBtRow;
                    CZSystem.log("CZTPGFrame BtTable バッチ№選択時の行位置","位置："+SelBtRow);

                    if(0 > row){
                        if(!life){
                            life = true;
                            return;
                        }
                        removeBtCondition();
                        return;
                    }
                    Vector v = new Vector(50);
                    CZSystemBt bt = (CZSystemBt)btList.elementAt(row);
                    for(int i = 0 ; i < btAllList.size() ; i++){
                        CZSystemBt btTmp = (CZSystemBt)btAllList.elementAt(i);
                        if(bt.batch.equals(btTmp.batch)) v.addElement(btTmp);
                    }
                    setRowSelectionInterval(0,row);
//@20131017 バッチ番号保持機能

                setBtCondition(v);

                }
                catch (Throwable e) {
                    CZSystem.handleException(e);
                }
            }
            /**
            * バッチ№選択時の処理
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
                SelBtRow = row;  //@20131017
                Vector v = new Vector(50);
                CZSystemBt bt = (CZSystemBt)btList.elementAt(row);
                for(int i = 0 ; i < btAllList.size() ; i++){
                    CZSystemBt btTmp = (CZSystemBt)btAllList.elementAt(i);
                    if(bt.batch.equals(btTmp.batch)) v.addElement(btTmp);
                }
                setBtCondition(v);
            }
            /**
            *
            */
            public void setData(int gr,int tbl){
            }
        } // BtTable
        /**
         * バッチ№一覧：モデル
         */
        public class BtTblMdl extends AbstractTableModel {

            private int TBL_ROW             = 0;        //行数
//@@@@@            final   int TBL_COL             = 3;        //列数
            final   int TBL_COL             = 5;        //列数    @@@@@
            private Vector  btList         = null;      //バッチ一覧
                                                        //列名
//@@@@@            final String[] names = {" # "  , "Bt" , "登録日時" };
            final String[] names = {" # "  , "Bt" , "品種" , "T2" , "登録日時" };    //@@@@@
            private Object  data[][];                   //データ

            /**
            * コンストラクタ
            * @param v バッチ情報
            */
            BtTblMdl(Vector v){
                super();
                btList = v;                             //バッチ一覧を保持する。
                TBL_ROW = btList.size();                //行数を設定する。
                data = new Object[TBL_ROW][TBL_COL];    //データ領域を確保する。
                //バッチ情報を１件ずつデータ領域へ保持する。
                for(int i = 0 ; i < TBL_ROW ; i++){
                    CZSystemBt bt = (CZSystemBt)btList.elementAt(i);
                    if(null == bt) break;               //データがなくなり次第終了する。
                    data[i][0] = new Integer(i+1);      // #
                    data[i][1] = bt.batch;              // Bt
//@@@@@                    data[i][2] = bt.t_time;             // 登録日時
                    data[i][2] = bt.hinshu;             // 品種 @@@@@
                    data[i][3] = bt.no_hikiage;         // T2 @@@@@
                    data[i][4] = bt.t_time;             // 登録日時 @@@@@
                }
            }
            /**
            * 列数を取得する。
            * @return 列数
            */
            public int getColumnCount(){
                return TBL_COL;
            }
            /**
            * 行数を取得する。
            * @return 行数
            */
            public int getRowCount(){
                return TBL_ROW;
            }
            /**
            * 値を取得する。
            * @param row ... 行, col ... 列
            * @return 値
            */
            public Object getValueAt(int row, int col){
                return data[row][col];
            }
            /**
            * 列名を取得する。
            * @param column ... 列
            * @return 列名
            */
            public String getColumnName(int column){
                return names[column];
            }
            /**
            * 列のデータ型を取得する。
            * @param c ... 列
            * @return データの型
            */
            public Class getColumnClass(int c){
                return getValueAt(0, c).getClass();
            }
            /**
            * セルの編集可否を取得する。
            * @param row ... 行, col ... 列
            * @return true :可, false:否
            */
            public boolean isCellEditable(int row, int col){
                return false;
            }
            /**
            * 値を設定する。
            * @param aValue ... 値, row ... 行, column ... 列
            */
            public void setValueAt(Object aValue, int row, int column){
                data[row][column] = aValue;
            }
        } // BtTblMdl

        /**
         * Ｂｔスタート時間一覧
         */
        class BtStartTable extends JTable {

            private Vector  btList      = null;     //バッチ情報
            private Vector  btStartList = null;     //バッチ開始情報
            private BtStartTblMdl model = null;     //バッチ開始テーブルのモデル
            private boolean life        = false;    

            /**
            * コンストラクタ
            * @param v バッチ情報
            */
			@SuppressWarnings("unchecked")
            BtStartTable(Vector v){
                super();

                btList = v;                         //バッチ一覧を保持する。
                try{
                    //テーブルの体裁を整える。
                    setName("BtStartTable");
                    setAutoCreateColumnsFromModel(true);
                    setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
                    setLocale(new Locale("ja","JP"));
                    setFont(new java.awt.Font("dialog", 0, 12));
                    setRowHeight(17);
                    //取敢えず最初のバッチの開始情報を取得する。
                    CZSystemBt bt = (CZSystemBt)btList.elementAt(0);
                    Vector tmp = new Vector();
                    tmp = CZSystem.getBtStart(roDbName_,bt.batch);
                    //バッチ開始情報が無ければ戻る。
                    if(null == tmp) return;
                    //バッチ開始情報を保持する領域を確保する。
                    int size = tmp.size();
                    btStartList = new Vector(size);
                    //sp_no = 1 のデータだけを保持する。
                    for(int i = 0 ; i < size ; i++){
                        CZSystemStart st = (CZSystemStart)tmp.elementAt(i);
                        if(null == st) break;
//                        if(1 == st.sp_no)
                        btStartList.addElement(st);     
                    }
                    //テーブルのモデルを生成する。
                    model = new BtStartTblMdl(btStartList);
                    setModel(model);
                    //列の体裁を整える。
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
            * 選択時の処理
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
                    removeBtStart();                //バッチ開始情報を削除する。
                    btnRead.setEnabled(false);      //読込みボタンを無効にする。
                    return;
                }
                CZSystemStart st = (CZSystemStart)btStartList.elementAt(row);
                SelectNo = row + 1;
//                CZSystem.log("CZTPGFrame", "###########? "+SelectNo);
                setBtStart(st);                     //バッチ開始情報を設定する。
                btnRead.setEnabled(true);           //読込みボタンを有効にする。
            }
            /**
            */
            public void setData(int gr,int tbl){
            }
        }

        /**
         * Ｂｔスタート時間一覧：モデル
         */
        public class BtStartTblMdl extends AbstractTableModel {

            private int TBL_ROW         = 0;            //行数
            final   int TBL_COL         = 6;            //列数
            private Vector  btStartList = null;         //バッチ情報
                                                        //列名を定義する。
            final String[] names = {" # ",   "PNo"  ,
                                    "SPNo",  "PSeq" ,
                                    "プロセス",
//                                    "登録日時" };
                                    "開始日時" };
            private Object  data[][];                   //データ領域

            /**
            * コンストラクタ
            * @param v バッチ情報
            */
            BtStartTblMdl(Vector v){
                super();
                btStartList = v;                        //バッチ情報を保持する。
                TBL_ROW = btStartList.size();           //行数をバッチ情報の件数とする。
                data = new Object[TBL_ROW][TBL_COL];    //データ領域を確保する。
                for(int i = 0 ; i < TBL_ROW ; i++){     //データを設定する。
                                                        //バッチ情報を１件ずつ取出す。
                    CZSystemStart st = (CZSystemStart)btStartList.elementAt(i);
                    if(null == st) break;                           //データがなくなり次第終了する。
                    data[i][0] = new Integer(i+1);                  // #
                    data[i][1] = new Integer(st.p_no);              // PNo
                    data[i][2] = new Integer(st.sp_no);             // SPNo
                    data[i][3] = new Integer(st.p_renban);          // PSeq
                    data[i][4] = CZSystem.getProcName(st.p_no);     // プロセス
                    data[i][5] = st.p_start;                        // 登録日時
                }
            }
            /**
            * 列数を取得する。
            * @return 列数
            */
            public int getColumnCount(){
                return TBL_COL;
            }
            /**
            * 行数を取得する。
            * @return 行数
            */
            public int getRowCount(){
                return TBL_ROW;
            }
            /**
            * 値を取得する。
            * @param row .. 行, col .. 列
            * @return 値
            */
            public Object getValueAt(int row, int col){
                return data[row][col];
            }
            /**
            * 列名を取得する。
            * @param column ... 列
            * @return 列名
            */
            public String getColumnName(int column){
                return names[column];
            }
            /**
            * データの型を取得する。
            * @param c ... 列
            * @return データの型
            */
            public Class getColumnClass(int c){
                return getValueAt(0, c).getClass();
            }
            /**
            * 編集可否を取得する。
            * @param row .. 行,col .. 列
            * @return true .. 可, false .. 否
            */
            public boolean isCellEditable(int row, int col){
                return false;
            }
            /**
            * 列数を取得する。
            * @param aValue .. ,row .. 行,column .. 列
            */
            public void setValueAt(Object aValue, int row, int column){
                data[row][column] = aValue;
            }
        }
    } // SercheDialog
//=================================== class end =========================================
}
