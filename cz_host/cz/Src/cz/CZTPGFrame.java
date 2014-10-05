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
 * TPG�O���t
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * @Update 2003/06/01 �q�[�^���x���v���Z�X���̑��Ε\���ɕύX( @@@ )
 * @Update 2003/08/04 �d���ʂ�������+�ǉ��ʂɕύX( @@@@ )
 * @Update 2008/09/17 TPG�EPV�ۑ��Ώە\�����ǉ�( @@@@@ )
 * @Update 2013/10/17 TPG�����ޯ��ԍ��ێ��@�\( @20131017 )
 * @Update 2013/10/28 TPG���̓��͍ő吔�ύX ( @20131028 )
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

//    private static GraphDialog graDl_   = null; //�O���t�\���p�_�C�A���O
    static  JLabel lblRo;
    private SercheDialog  sercheDia_    = null; //�����p�_�C�A���O
    private CZRoSelectWin3 rosel        = null;
    private static CZTPGGraphFrame graDl_   = null; //�O���t�\���p�_�C�A���O

    private int gph_cnt = 0;

    private CZSystemStart roBtStart_    = null; //�����p�����グ����
    private Vector roBtAllCondition_    = null; //�SBt�̈����グ����
    
    private int    SelectNo             = 0;
    private String SelectBt             = null;
    private String SelectTime           = null;
    private Vector roBtTempCondition_    = null; //�I��Bt�̈����グ����


    private String roName_              = null; //�ΏۘF��
    private String roDbName_            = null; //�ΏۘF�f�[�^�x�[�X��

    private Vector  selList_            = null; //Y������

    private Vector pvDataBody_          = null; //PV�f�[�^

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

    private int SelBtRow = 0;    //�I��Bt Row(�o�b�t�@�j(�����l:0) @20131017
    //==========================================================================
    /**
     * �X�^�[�g�A�b�v
     * @param   args    �R�}���h���C������
     */
    //==========================================================================
/**@@
    public static void main(String[] args) {

        JFrame tpg_ = new CZTPGFrame("TPG�f�o�b�O�p�p�l��", 0 );
        tpg_.setSize(1024, 660);
        tpg_.setVisible(true);
    }
@@*/
    //==========================================================================
    /**
     * �R���X�g���N�^
     * @param   String title  Frame Title
     * @param   int ui        UI Manager Look and Feel
     */
    //==========================================================================
    public CZTPGFrame()
    {

        super();
        setupUI(0);                                 //UI��ݒ肷��B

        roName_     = CZSystem.getRoName();         //�F�����擾����B
        roDbName_   = CZSystem.getDBName();         //DB�X�L�[�}�����擾����B

        setTitle("�g�����h�e�[�u���ݒ�");                         //���Title
//        setTitle("�g�����h�e�[�u��");                         //���Title
        setSize(1024, 800);
        setResizable(false);                        //��ʂ̃T�C�Y�ύX�͕s��
//        setModal(true);                             //Modal�ŕ\��

        pane_ = getContentPane();
        pane_.setLayout(null);
        // ����n�Q�Ƌ@�\    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            pane_.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            pane_.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        try{
            // ----- Property_File��� Min,Max�l���擾����B --------
            Properties prop =  new Properties();
            FileInputStream pros = new FileInputStream("TPGDEF.TXT");
            prop.load(pros);
            // X���̐ݒ�
            prop_xUnit = prop.getProperty("X_UNIT");
            prop_xMin  = prop.getProperty("X_START");
            prop_xMax  = prop.getProperty("X_END");
            prop_xBun  = prop.getProperty("X_BUNKATU");

            // Y���̐ݒ�
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
                                        //�v���p�e�B�擾�ŃG���[�̎��́A�I������B
		CZSystem.log("CZTPGFrame", "GET ERROR EXIT !!");
            CZSystem.exit(-1,"CZTPG NO Propertie File");
        }

        makePanels();                               //�ݒ��ʂ𐶐�����B
        sercheDia_ = null;
        sercheDia_ = new SercheDialog();            //������ʂ𐶐�����B
        sercheDia_.setVisible(false);               //������ʂ���Ă����B

		CZSystem.log("CZTPGFrame", "CZTPG new");
    }
    //==========================================================================
    /**
     * UI��Setup
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

        // �t���[���̏�����     ------------------------------------------------
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

		Properties prop = new Properties();           // �v���p�e�B�𐶐�����
		// X���̐ݒ�
		prop.setProperty(new String("X_UNIT"),    new String("" + pnl3_.getUnit()) );
		prop.setProperty(new String("X_START"),   new String("" + pnl3_.getStart()));
		prop.setProperty(new String("X_END"),     new String("" + pnl3_.getEnd())  );
		prop.setProperty(new String("X_BUNKATU"), new String("" + pnl3_.getMesh()) );

		//Y���̐ݒ�
		for (int i = 0; i < 14; i++) {
		  /*
		  if( new String(pnlSel_[i].getLineS()).length() == 0 ){
			Object msg[] = { "��������͂��ĉ�����", "", "" };
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
		//---------- �t�@�C���ɕۑ�����  ----------
		try {
//			CZSystem.log("CZTPGFrame ","�t�@�C���ɕۑ������B");
//		    FileOutputStream out = new FileOutputStream("d:/CZ/classes/TPGDEF.TXT");
		    FileOutputStream out = new FileOutputStream("TPGDEF.TXT");
		    prop.store(out, "");
		    out.flush();
		    out.close();
		} catch (IOException ex) {
		    JOptionPane.showMessageDialog(
		      tpg_,
		      new String("�ۑ��ł��܂���ł����B"),
		      new String("�ۑ�"),
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
//================================ ��������ݒ��� =======================================
    //--------------------------------------------------------------------------
    /**
     * ��ʍ쐬
     */
    //--------------------------------------------------------------------------
	@SuppressWarnings("unchecked")
    protected void makePanels()
    {

        Border brd = BorderFactory.createRaisedBevelBorder();

        //------------------------- ���ڈꗗ��\������Panel  -------------------
        pnl4_ = new PVIchiranPanel();
        pnl4_.setBounds(531,281,470,280);
        pane_.add( pnl4_ );

        //------------------------- Title��\������ Panel ----------------------
        pnl1_ = new TitlePanel();
        pnl1_.setBounds(0,0,1024,160);
        pane_.add( pnl1_ );

        //------------------------- �I�����w�肷��Panel  -----------------------
        pnl2_ = new SelectPanel();
        pnl2_.setBounds(0,161,500,530);
        pane_.add( pnl2_ );

        //------------------------- X���̐ݒ���w������ Panel ------------------
        pnl3_ = new XLengeSetPanel();
        pnl3_.setBounds(531,161,470,120);

        pane_.add( pnl3_ );

        //------------------------- �{�^��Panel  -------------------
        JPanel pnl5 = new JPanel();
        pnl5.setLayout( null );
        // ����n�Q�Ƌ@�\    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            pnl5.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            pnl5.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }
        pnl5.setBounds(0,700,1024,60);

        JButton btnGraph = new JButton("�O���t�\��");
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
                System.gc();                    //@@@ �抸����GC�����s����B
                
                if( SelectBt == null){
					Object msg[] = { "�O���t�\��������I�����ĉ�����", "", "" };
					errorMsg(msg);
					return;
				}else{
                    gph_cnt = CZSystem.GraphCount();
                    if(gph_cnt > 4){
                        Object msg[] = { "�O���t�͂T���ȏ�J���܂���", "", "" };
                        errorMsg(msg);
						return;
					}else{
                        graDl_ = new CZTPGGraphFrame(SelectNo,roDbName_,SelectBt,SelectTime,pvDataBody_,selList_,roBtStart_);     //�O���t�𐶐�����B
                        graDl_.setXParam(pnl3_.getUnit(),pnl3_.getStart(),pnl3_.getEnd(),pnl3_.getMesh());
                        graDl_.setData();               //
                        graDl_.setVisible(true);        //�O���t��\������B
                        CZSystem.GraphCountUp();
                    }
                }
              }
          }
        );
        pnl5.add(btnGraph);

        btnOpen_ = new JButton("�ݒ�Ǎ�");
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
                      file_ = chooser.getSelectedFile();        // �t�@�C�������擾����
                      Properties prop = new Properties();       // �v���p�e�B�𐶐�����
                      try {
                          FileInputStream in = new FileInputStream(file_);
                          prop.load( in );                      //�v���p�e�B���擾����B
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
                          CZSystem.log("CZTPGFrame ","Property File�����[�h�ł��܂���ł����B");
                          return;
                      }
                  }
              }
          }
        );
        btnOpen_.setBounds(600, 20, 100, 30);
        pnl5.add(btnOpen_); 
        // ======================================== [�ۑ�]�{�^�� ==================================
        btnSave_ = new JButton("�ݒ�ۑ�");
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
                      file_ = chooser.getSelectedFile();            // �t�@�C�������擾����
                      Properties prop = new Properties();           // �v���p�e�B�𐶐�����
                      // X���̐ݒ�
                      prop.setProperty(new String("X_UNIT"),    new String("" + pnl3_.getUnit()) );
                      prop.setProperty(new String("X_START"),   new String("" + pnl3_.getStart()));
                      prop.setProperty(new String("X_END"),     new String("" + pnl3_.getEnd())  );
                      prop.setProperty(new String("X_BUNKATU"), new String("" + pnl3_.getMesh()) );

                      //Y���̐ݒ�
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
                      //---------- �t�@�C���ɕۑ�����  ----------
                      try {
//						CZSystem.log("CZTPGFrame ","�t�@�C���ɕۑ������B");
                          FileOutputStream out = new FileOutputStream(file_);
                          prop.store(out, "");
                          out.flush();
                          out.close();
                      } catch (IOException ex) {
                          JOptionPane.showMessageDialog(
                            tpg_,
                            new String("�ۑ��ł��܂���ł����B"),
                            new String("�ۑ�"),
                            JOptionPane.WARNING_MESSAGE);
                          return;
                      }
                      JOptionPane.showMessageDialog(
                        tpg_,
                        new String("�ۑ����܂����B"),
                        new String("�ۑ�"),
                        JOptionPane.INFORMATION_MESSAGE);
                      return;
                  }
              }
          }
        );
        btnSave_.setBounds(750, 20, 100, 30);
        pnl5.add(btnSave_); 

        JButton btnExit = new JButton("�I��");
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
     *  @param msg ... ���b�Z�[�W���e
     *  @return true ... OK, false ... NG
     */
    //----------------------------------------------------------------------
    private boolean errorMsg(Object msg[]){
        JOptionPane.showMessageDialog(null,msg,
                                "���̓G���[",
                                JOptionPane.ERROR_MESSAGE);
        return true;
    }


    //==========================================================================
    /**
     *   Title�\��Panel
     */
    //==========================================================================
    class TitlePanel extends JPanel {

        JLabel lbl3[] = new JLabel[13];

        JLabel lbl4[] = new JLabel[2];

        /**
        * �R���X�g���N�^
        */
		@SuppressWarnings("unchecked")
        TitlePanel(){
            super();

            setLayout(null);
            // ����n�Q�Ƌ@�\    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }
            setForeground(Color.black);

            JLabel lbl1 = new JLabel("�g�����h�e�[�u���ݒ�",JLabel.CENTER);
            lbl1.setLayout(new FlowLayout(FlowLayout.CENTER));
            lbl1.setFont(new java.awt.Font("dialog", 0, 32));
            lbl1.setForeground(Color.black);
            lbl1.setBounds(0,0,900,36);
            add(lbl1);

			String s = CZSystem.RoKetaChg(CZSystem.getRoName());	// 20050725 �F�F�\�������ύX
            lblRo = new JLabel(s,JLabel.CENTER);
//            JLabel lblRo = new JLabel(CZSystem.getRoName(),JLabel.CENTER);
            lblRo.setFont(new java.awt.Font("dialog", 0, 14));
            lblRo.setForeground(Color.black);
            lblRo.setBorder(new Flush3DBorder());
            lblRo.setBounds(20,20,80,30);
            add(lblRo);

            JButton btn_chgRo = new JButton("��");
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
						SelBtRow = 0;  //@20131017 �����l0�ɖ߂�
						CZSystem.log("CZTPGFrame TitlePanel", "�o�b�`�I��ԍ��@�����l0");  //@20131017
					}
				}
			);
            add(btn_chgRo);

            JButton btnSearch = new JButton("��  ��");
            btnSearch.setBounds(20, 60, 80, 30);
            btnSearch.setLocale(new Locale("ja","JP"));
            btnSearch.setFont(new java.awt.Font("dialog", 0, 14));
            btnSearch.setBorder(new Flush3DBorder());
            btnSearch.setForeground(java.awt.Color.black);
            btnSearch.addActionListener(
                new ActionListener() {
                    public void actionPerformed(ActionEvent ev){
                        selList_ = null;
                        selList_ = new Vector();            //�I�����X�g���쐬����B
                        for (int i=0; i<14; i++){
                            if( null != pnlSel_[i].getChNo() ){
                                selList_.addElement(pnlSel_[i]);
                            }
                        }
                        sercheDia_.setDefault();            //������ʂ�����������B
                        sercheDia_.setVisible(true);        //������ʂ�\������B
                    }
                }
            );
            add(btnSearch);

            //�Œ�\����
            JLabel lbl2[] = new JLabel[12];

            lbl4[0] = new JLabel("(#)",JLabel.CENTER);
            lbl4[0].setFont(new java.awt.Font("dialog", 0, 12));
            lbl4[0].setForeground(java.awt.Color.black);
            lbl4[0].setBorder(new Flush3DBorder());
            lbl4[0].setBounds(100,60,40,30);
            add(lbl4[0]);

            lbl2[0] = new JLabel("(���t����)",JLabel.CENTER);
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

            lbl2[2] = new JLabel("(�i��)",JLabel.CENTER);
            lbl2[2].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[2].setForeground(java.awt.Color.black);
            lbl2[2].setBorder(new Flush3DBorder());
            lbl2[2].setBounds(550,60,50,30);
            add(lbl2[2]);

            lbl2[3] = new JLabel("(�v���Z�X)",JLabel.CENTER);
            lbl2[3].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[3].setForeground(java.awt.Color.black);
            lbl2[3].setBorder(new Flush3DBorder());
            lbl2[3].setBounds(690,60,60,30);
            add(lbl2[3]);

            lbl2[4] = new JLabel("(�`���[�W��)",JLabel.CENTER);
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

//            lbl2[11]  = new JLabel("(mm)(�ݒ蒼�a)",JLabel.CENTER);
            lbl2[11]  = new JLabel("(�ݒ蒼�a)[mm]",JLabel.CENTER);
            lbl2[11].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[11].setForeground(java.awt.Color.black);
            lbl2[11].setBorder(new Flush3DBorder());
            lbl2[11].setBounds(800,95,90,30);
            add(lbl2[11]);

            //�f�[�^�\����

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
        * �o�b�`����ݒ肷��B
        */
        public void setBtCondition() {

CZSystem.log("CZTPGFrame", "�F�́H"+roDbName_);
CZSystem.log("CZTPGFrame", "�o�b�`�́H"+SelectBt);
CZSystem.log("CZTPGFrame", "���Ԃ́H"+SelectTime);
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
                lbl3[3].setText(CZSystem.getProcName(roBtStart_.p_no));     //@@@ �v���Z�X
//@@@@                lbl3[4].setText(new Integer(bt.i_sikomi).toString());
                lbl3[4].setText(new Integer(bt.i_sikomi + bt.t_sikomi).toString());     //@@@@
                lbl3[5].setText("�l.�s="   + new Integer(bt.no_youkai).toString());
                lbl3[6].setText("�o.�s="   + new Integer(bt.no_hikiage).toString());
                lbl3[7].setText("�q.�s="   + new Integer(bt.no_kaiten).toString());
                lbl3[8].setText("�d.�s="   + new Integer(bt.no_toridasi).toString());
                lbl3[9].setText("�`.�s="   + new Integer(bt.no_aturyoku).toString());
                lbl3[10].setText("�b.�s="  + new Integer(bt.no_teisu).toString());
                lbl3[11].setText("�c�h�`=" + new Integer(bt.chokkei).toString());
            } else {
                if (null != roBtAllCondition_){
	                CZSystemBt bt = (CZSystemBt)roBtAllCondition_.elementAt(0);
	                lbl4[1].setText(new Integer(SelectNo).toString());
	                lbl3[0].setText((bt.t_time).trim());
	                lbl3[1].setText((bt.batch).trim());
	                lbl3[2].setText((bt.hinshu).trim());
	//@@@                lbl3[3].setText((bt.pgid).trim());
	                lbl3[3].setText(CZSystem.getProcName(roBtStart_.p_no));     //@@@ �v���Z�X
	//@@@@                lbl3[4].setText(new Integer(bt.i_sikomi).toString());
	                lbl3[4].setText(new Integer(bt.i_sikomi + bt.t_sikomi).toString());     //@@@@
	                lbl3[5].setText("�l.�s="   + new Integer(bt.no_youkai).toString());
	                lbl3[6].setText("�o.�s="   + new Integer(bt.no_hikiage).toString());
	                lbl3[7].setText("�q.�s="   + new Integer(bt.no_kaiten).toString());
	                lbl3[8].setText("�d.�s="   + new Integer(bt.no_toridasi).toString());
	                lbl3[9].setText("�`.�s="   + new Integer(bt.no_aturyoku).toString());
	                lbl3[10].setText("�b.�s="  + new Integer(bt.no_teisu).toString());
	                lbl3[11].setText("�c�h�`=" + new Integer(bt.chokkei).toString());
                } else {
	                lbl4[1].setText("");
	                lbl3[0].setText("");
	                lbl3[1].setText("");
	                lbl3[2].setText("");
	                lbl3[3].setText("");
	                lbl3[4].setText("");
	                lbl3[5].setText("�l.�s");
	                lbl3[6].setText("�o.�s");
	                lbl3[7].setText("�q.�s");
	                lbl3[8].setText("�d.�s");
	                lbl3[9].setText("�`.�s");
	                lbl3[10].setText("�b.�s");
	                lbl3[11].setText("�c�h�`");
                }
            }
        }
        
        /**
        * �o�b�`�����N���A����B
        */
        public void clearBtCondition() {
            lbl4[1].setText("");
            lbl3[0].setText("");
            lbl3[1].setText("");
            lbl3[2].setText("");
            lbl3[3].setText("");
            lbl3[4].setText("");
            lbl3[5].setText("�l.�s");
            lbl3[6].setText("�o.�s");
            lbl3[7].setText("�q.�s");
            lbl3[8].setText("�d.�s");
            lbl3[9].setText("�`.�s");
            lbl3[10].setText("�b.�s");
            lbl3[11].setText("�c�h�`");
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
     *   �o�u�I��Panel
     */
    //==========================================================================
    class SelectPanel extends JPanel {  

        /**
        * �R���X�g���N�^
        */
        SelectPanel(){
            super();

            setLayout( null );
            // ����n�Q�Ƌ@�\    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            JPanel pnl1 = new JPanel();
            pnl1.setLayout( null );
            // ����n�Q�Ƌ@�\    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                pnl1.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                pnl1.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            JLabel lbl = null;
            lbl = new JLabel("�F");
            lbl.setFont(new java.awt.Font("dialog", 0, 16));
            lbl.setForeground(Color.black);
            lbl.setBounds(10,20,40,18);
            pnl1.add(lbl);

            lbl = new JLabel("����No");
            lbl.setFont(new java.awt.Font("dialog", 0, 16));
            lbl.setForeground(Color.black);
            lbl.setBounds(41,20,200,18);
            pnl1.add(lbl);

            lbl = new JLabel("���̑���");
            lbl.setFont(new java.awt.Font("dialog", 0, 16));
            lbl.setForeground(Color.black);
            lbl.setBounds(220,20,200,18);
            pnl1.add(lbl);

            lbl = new JLabel("�����W");
            lbl.setFont(new java.awt.Font("dialog", 0, 16));
            lbl.setForeground(Color.black);
            lbl.setBounds(350,0,100,18);
            pnl1.add(lbl);

            lbl = new JLabel("�l����");
            lbl.setFont(new java.awt.Font("dialog", 0, 16));
            lbl.setForeground(Color.black);
            lbl.setBounds(310,20,100,18);
            pnl1.add(lbl);

            lbl = new JLabel("�l����");
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
                        // ����n�Q�Ƌ@�\    @20131021
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
                        // ����n�Q�Ƌ@�\    @20131021
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
     *   X���̐ݒ������Panel
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
            // ����n�Q�Ƌ@�\    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            JLabel lbl = null;
            lbl = new JLabel("�\���P��(1:��    2:mm)");
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

            lbl = new JLabel("�\���͈�");
            lbl.setFont(new java.awt.Font("dialog", 0, 16));
            lbl.setForeground(Color.black);
            lbl.setBounds(10,40,120,18);
            add(lbl);

            lbl = new JLabel("�X�^�[�g");
            lbl.setFont(new java.awt.Font("dialog", 0, 16));
            lbl.setForeground(Color.black);
            lbl.setBounds(80,40,120,18);
            add(lbl);

            txtStart_ = new JTextFieldInt();
            txtStart_.setMinValue(0);
            txtStart_.setMaxValue(10000);	// @20131028 TPG���̓��͍ő吔�ύX
            if ( null != prop_xMin ) {
                txtStart_.setValue(Integer.parseInt(prop_xMin));
            } else {
                txtStart_.setValue(0);
            }
            txtStart_.setBounds(240,40,80,18);
            add(txtStart_);

            lbl = new JLabel("�G���h");
            lbl.setFont(new java.awt.Font("dialog", 0, 16));
            lbl.setForeground(Color.black);
            lbl.setBounds(80,60,120,18);
            add(lbl);

            txtEnd_ = new JTextFieldInt();
            txtEnd_.setMinValue(0);
            txtEnd_.setMaxValue(10000);		// @20131028 TPG���̓��͍ő吔�ύX
            if ( null != prop_xMax ) {
                txtEnd_.setValue(Integer.parseInt(prop_xMax));
            } else {
                txtEnd_.setValue(10000);	// @20131028 TPG���̓��͍ő吔�ύX
            }
            txtEnd_.setBounds(240,60,80,18);
            add(txtEnd_);

            lbl = new JLabel("���b�V���Ԋu��");
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
         * @return �P��
         */
        public int getUnit(){
            return txtUnit_.getValue();
        }
        /**
         * @return �J�n�l
         */
        public int getStart(){
            return txtStart_.getValue();
        }
        /**
         * @return �I���l
         */
        public int getEnd(){
            return txtEnd_.getValue();
        }
        /**
         * @return �w��������
         */
        public int getMesh(){
            return txtMesh_.getValue();
        }
    } //XLengeSetPanel

    //==========================================================================
    /**
     *   PV���ڈꗗ�\��Panel
     */
    //==========================================================================
    class PVIchiranPanel extends JPanel {

        /**
        * �R���X�g���N�^
        */
        PVIchiranPanel(){
            super();

            setLayout(null);
            // ����n�Q�Ƌ@�\    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            JLabel lbl1 = new JLabel("����  �o�u���ڂm���ꗗ  ����",JLabel.CENTER);
            lbl1.setLayout(new FlowLayout(FlowLayout.CENTER));
            lbl1.setFont(new java.awt.Font("dialog", 0, 20));
            lbl1.setForeground(Color.black);
            lbl1.setBounds(0,0,470,26);

            JPanel pnl1 = new JPanel();
            pnl1.setLayout(null);
            // ����n�Q�Ƌ@�\    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                pnl1.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                pnl1.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }
            pnl1.setBounds(0,0,500,40);
            pnl1.add(lbl1);
            add(pnl1);

            // �o�u���̈ꗗ�\���擾����ʂ֕\������B
            PvNameTable t = new PvNameTable((Vector) CZSystem.getPVNameAll());
            JTableHeader tabHead = t.getTableHeader();
            tabHead.setReorderingAllowed(false);
            JScrollPane pvNameScpanel = new JScrollPane();
            pvNameScpanel.setBounds(0, 0, 460, 240);
            pvNameScpanel.setViewportView(t);

            JPanel pnl2 = new JPanel();
            pnl2.setLayout(null);
            // ����n�Q�Ƌ@�\    @20131021
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
     *   �O���t�\�����ڑI��Panel
     */
    //==========================================================================
    class SelectItemPanel extends JPanel {  

        private int             panelNo_    = 0;        // No
        private Color           col_        = java.awt.Color.gray;
        private NumText         txtCh_      = null;     // PV����No
        private LineSize        lineS_      = null;     // ���̑���
        private JTextFieldFloat txtMin_     = null;     // Min�l
        private JTextFieldFloat txtMax_     = null;     // Max�l
        private JLabel          lblName_    = null;     // PV��
        private JButton         btnCol_     = null;

        /**
        * �R���X�g���N�^
        */
        SelectItemPanel(int no, Color c){

            super();

            // ����n�Q�Ƌ@�\    @20131021
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
                        Color c = JColorChooser.showDialog(null,"�F��I��ł�������", but.getBackground());
                        if(null != c){
                            col_ = c;
                            but.setForeground(c);           //�I�������F��ݒ肷��B
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
                                Object msg[] = { "���͒l( " + val + " )�������ł��B�I�I", "", "" };
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
                                Object msg[] = { "���͒l( " + val + " )�������ł��B�I�I", "", "" };
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
                                Object msg[] = { "���͒l( " + size + " )�������ł��B�I�I", "", "" };
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
                                Object msg[] = { "���͒l( " + size + " )�������ł��B�I�I", "", "" };
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
         *  @param msg ... ���b�Z�[�W���e
         *  @return true ... OK, false ... NG
         */
        //----------------------------------------------------------------------
        private boolean errorMsg(Object msg[]){
            JOptionPane.showMessageDialog(null,msg,
                                    "���̓G���[",
                                    JOptionPane.ERROR_MESSAGE);
            return true;
        }

        /**
         * �F��ݒ肷��B
         */
        public void setColor(Color c){
            col_ = c;
            btnCol_.setForeground(col_);
            btnCol_.setBackground(col_);
        }

        /**
         * �F���擾����B
         */
        public Color getColor(){
            return col_;
        }

        /**
         * �p�l���̇����擾����B
         */
        public int getNo(){
            return panelNo_;
        }

        /**
         * �`���l���̇����擾����B
         */
        public String getChNo(){
            return txtCh_.getText();
        }

        /**
         * �`���l���̇���ݒ肷��B
         */
        public void setChNo(String s){
            txtCh_.setText(s);
            return;
        }

        /**
         * ���ږ����擾����B
         */
        public String getName(){
            return lblName_.getText();
        }

        /**
         * ���ږ���ݒ肷��B
         */
        public void setName(String s){
            lblName_.setText(s);
            return;
        }

        /**
         * ���̑������擾����B
         */
        public String getLineS(){
            return lineS_.getText();
        }

        /**
         * ���̑�����ݒ肷��B
         */
        public void setLineS(String s){
            lineS_.setText(s);
            return;
        }

        /**
         * �ő�l���擾����B
         */
        public float getMax(){
            return txtMax_.getValue();
        }

        /**
         * �ő�l��ݒ肷��B
         */
        public void setMax(float f){
            txtMax_.setValue(f);
            return;
        }

        /**
         * �ŏ��l���擾����B
         */
        public float getMin(){
            return txtMin_.getValue();
        }

        /**
         * �ŏ��l��ݒ肷��B
         */
        public void setMin(float f){
            txtMin_.setValue(f);
            return;
        }

        /**
         * �ő�l�A�ŏ��l��Default�l��ݒ肷��B
         */
        public void setDefault(){
            txtMin_.setDefaultValue();
            txtMax_.setDefaultValue();
        }

    } //SelectItemPanel

    //==========================================================================
    /**
     * float�^�̏���ێ�����e�L�X�g�t�B�[���h�N���X
     */
    //==========================================================================
    public class JTextFieldFloat extends JTextField {

        /**
         * �ݒ�\�ȍő�l
         */
        private float max_ = Float.POSITIVE_INFINITY;
        /**
         * �ݒ�\�ȍŏ��l
         */
        private float min_ = Float.NEGATIVE_INFINITY;
        /**
         * �ێ�����l
         */
        private float val_ = 0.0f;

        /**
         * �R���X�g���N�^
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
         * float�l�̐ݒ�
         * @param   val     �ݒ肷��l
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
                Object msg[] = { "���͒l( " + val + " )�������ł��B�I�I", "", "" };
                errorMsg(msg);
                setText("" + val_);
            }
        }

        /**
         * float�l�̎擾
         * @return  float�l
         */
        float getValue() {
            return val_;
        }

        /**
         * �ő�l�̐ݒ�
         * @param   max     �ő�l
         */
        public void setMaxValue(float max) {
            max_ = max;
        }

        /**
         * �ŏ��l�̐ݒ�
         * @param   min     �ŏ��l
         */
        public void setMinValue(float min) {
            min_ = min;
            setValue(min_);
        }

        /**
         * Default�l�̐ݒ�
         */
        public void setDefaultValue() {
            max_ = Float.POSITIVE_INFINITY;;
            min_ = Float.NEGATIVE_INFINITY;
            setValue(0.0f);
        }

        /**
         * String�l��float�l�ɕϊ�
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
         *  @param msg ... ���b�Z�[�W���e
         *  @return true ... OK, false ... NG
         */
        private boolean errorMsg(Object msg[]){
            JOptionPane.showMessageDialog(null,msg,
                                    "���̓G���[",
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
     * Integer�^�̏���ێ�����e�L�X�g�t�B�[���h�N���X
     *
     */
    //======================================================================
    public class JTextFieldInt extends JTextField {

        /**
         * �ݒ�\�ȍő�l
         */
        private int max_ = Integer.MAX_VALUE;
        /**
         * �ݒ�\�ȍŏ��l
         */
        private int min_ = Integer.MIN_VALUE;
        /**
         * �ێ�����l
         */
        private int val_ = 0;

        /**
         * �R���X�g���N�^
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
         * int�l�̐ݒ�
         * @param   val     �ݒ肷��l
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
                Object msg[] = { "���͒l( " + val + " )�������ł��B�I�I", "", "" };
                errorMsg(msg);
                setText("" + val_);
            }
        }

        /**
         * int�l�̎擾
         * @return  int�l
         */
        int getValue() {
            return val_;
        }

        /**
         * �ő�l�̐ݒ�
         * @param   max     �ő�l
         */
        public void setMaxValue(int max) {
            max_ = max;
        }

        /**
         * �ŏ��l�̐ݒ�
         * @param   min     �ŏ��l
         */
        public void setMinValue(int min) {
            min_ = min;
            setValue(min_);
        }

        /**
         * String�ŕ\�����ꂽ�l��float�l�ɕϊ�
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
         *  @param msg ... ���b�Z�[�W���e
         *  @return true ... OK, false ... NG
         */
        private boolean errorMsg(Object msg[]){
            JOptionPane.showMessageDialog(null,msg,
                                    "���̓G���[",
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
     *       ���l�ƃX�y�[�X���󂯕t����TextField
     */
    public class NumText extends JTextField {   

        /**
        * �R���X�g���N�^
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
     *       ���l�ƃX�y�[�X���󂯕t����TextField
     */
    public class LineSize extends JTextField {   

        /**
        * �R���X�g���N�^
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
     *       PV���ڈꗗ
     */
    //==========================================================================
    class PvNameTable extends JTable {

        private Vector  pvNameList_ = null;
        private pvNameTblMdl model_ = null;

        /**
        * �R���X�g���N�^
        * @param v ... PV���ږ�
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
                // ���ږ�
                colum = cmdl.getColumn(1);
                colum.setMaxWidth(80);
                colum.setMinWidth(80);
                colum.setWidth(80);
                // ���{�ꖼ��
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
         *       PV���ڈꗗ:���f��
         */
        //======================================================================
        public class pvNameTblMdl extends AbstractTableModel {

            private int     TBL_ROW     = 128;      // �s��
            final   int     TBL_COL     = 5;        // ��
            private Vector  pvNameList_ = null;     // �o�b�`���

            final String[] names = {" CH "  , "����", "����", "Min", "Max"};
            private Object  data[][];

            /**
            * �R���X�g���N�^ 
            * @param v ... PV����
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
                    data[i][1]  = new String(pvName.k_name);    //����
                    data[i][2]  = new String(pvName.j_name.trim());    //���{�ꖼ��
                    data[i][3]  = new Integer(pvName.n_min);    //Min
                    data[i][4]  = new Integer(pvName.n_max);    //Max
                    dataTbl_.put((Object)(data[i][0]), (Object)pvName);
                }
            }

            /**
            * �������擾����B
            * @return ... ����
            */
            public int getColumnCount(){
                return TBL_COL;
            }
            /**
            * �s�����擾����B
            * @return ... �s��
            */
            public int getRowCount(){
                return TBL_ROW;
            }
            /**
            * �f�[�^���擾����B
            * @param ... row:�s, col:��
            * @return ... �f�[�^
            */
            public Object getValueAt(int row, int col){
                return data[row][col];
            }
            /**
            * �������擾����B
            * @param ... column:��
            * @return ... ����
            */
            public String getColumnName(int column){
                return names[column];
            }
            /**
            * �f�[�^�̌^���擾����B
            * @param ... c:��
            * @return ... �f�[�^�̌^
            */
            public Class getColumnClass(int c){
                return getValueAt(0, c).getClass();
            }
            /**
            * cell�ҏW�̉ۂ��擾����B
            * @param ... row:�s, col:��
            * @return ... ����
            */
            public boolean isCellEditable(int row, int col){
                return false;
            }
            /**
            * �f�[�^��ݒ肷��B
            * @param ... aValue:�f�[�^, row:�s, col:��
            * @return ... ����
            */
            public void setValueAt(Object aValue, int row, int column){
                data[row][column] = aValue;
            }
        } // pvNameTblMdl
    } // PvNameTable


    /***************************************************
     *
     * �����グ����Dialog
     *
     ***************************************************/
    class BtConditionDialog extends JDialog {

        /**
        * �R���X�g���N�^
        */
        BtConditionDialog(){
            super();

            setTitle("�����グ����");
            setSize(820,240);
            setResizable(false);
            setModal(true);
            getContentPane().setLayout(null);
            // ����n�Q�Ƌ@�\    @20131021
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
         * �a���o�^���ꗗ
         * @@T6�ǉ�
         */
        class BtConditionTable extends JTable {

            private Vector  bt_list     = null;
            private BtConditionTblMdl model = null;

            /**
            * �R���X�g���N�^
            * @param v ... �o�b�`���
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
                    // �o�^����
                    colum = cmdl.getColumn(1);
                    colum.setMaxWidth(160);
                    colum.setMinWidth(160);
                    colum.setWidth(160);
                    // �A��
                    colum = cmdl.getColumn(2);
                    colum.setMaxWidth(30);
                    colum.setMinWidth(30);
                    colum.setWidth(30);
                    // �i��
                    colum = cmdl.getColumn(3);
                    colum.setMaxWidth(80);
                    colum.setMinWidth(80);
                    colum.setWidth(80);
                    // ���c�{
                    colum = cmdl.getColumn(4);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);
                    // ���a
                    colum = cmdl.getColumn(5);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);
                    // ���㒷
                    colum = cmdl.getColumn(6);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);
                    // ���d��
                    colum = cmdl.getColumn(7);
                    colum.setMaxWidth(60);
                    colum.setMinWidth(60);
                    colum.setWidth(60);
                    // �ǎd��
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
                    // �J�n
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
             * �a���o�^���ꗗ�F���f��
             */
            public class BtConditionTblMdl extends AbstractTableModel {

                private int     TBL_ROW     = 0;        // �s��
                final   int     TBL_COL     = 17;       // �� @@
                private Vector  bt_list     = null;     // �o�b�`���

                final String[] names = {" # "  , "�o�^����" , "�A��" ,  
                            "�i��" , "���c�{"   , "���a" ,
                            "���㒷" , "���d��"   , "�ǎd��" ,
                            "T1" , "T2"   , "T3" ,
                            "T4" , "T5"   , "T6"   , "PNo" , "�J�n"
                            };

                private Object  data[][];

                /**
                * �R���X�g���N�^
                * @param v ... �o�b�`���
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
                        data[i][1]  = bt.t_time;                    //�o�^����
                        data[i][2]  = new Integer(bt.renban);       //�A��
                        data[i][3]  = bt.hinshu;                    //�i��
                        data[i][4]  = new Integer(bt.rutubo_kei);   //���c�{
                        data[i][5]  = new Integer(bt.chokkei);      //���a
                        data[i][6]  = new Integer(bt.hikiage_cho);  //���㒷
                        data[i][7]  = new Integer(bt.i_sikomi);     //���d��
                        data[i][8]  = new Integer(bt.t_sikomi);     //�ǎd��
                        data[i][9]  = new Integer(bt.no_youkai);    //T1
                        data[i][10] = new Integer(bt.no_hikiage);   //T2
                        data[i][11] = new Integer(bt.no_kaiten);    //T3
                        data[i][12] = new Integer(bt.no_toridasi);  //T4
                        data[i][13] = new Integer(bt.no_aturyoku);  //T5
                        data[i][13] = new Integer(bt.no_teisu);     //T6 @@
                        data[i][14] = new Integer(bt.pno_start);    //PNo
                        data[i][15] = new Integer(bt.p_kaisi);      //�J�n
                    }

                }
                /**
                * �������擾����B
                * @return ... ����
                */
                public int getColumnCount(){
                    return TBL_COL;
                }
                /**
                * �s�����擾����B
                * @return ... �s��
                */
                public int getRowCount(){
                    return TBL_ROW;
                }
                /**
                * �f�[�^���擾����B
                * @param ... row:�s, col:��
                * @return ... �f�[�^
                */
                public Object getValueAt(int row, int col){
                    return data[row][col];
                }
                /**
                * �������擾����B
                * @param ... column:��
                * @return ... ����
                */
                public String getColumnName(int column){
                    return names[column];
                }
                /**
                * �f�[�^�̌^���擾����B
                * @param ... c:��
                * @return ... �f�[�^�̌^
                */
                public Class getColumnClass(int c){
                    return getValueAt(0, c).getClass();
                }
                /**
                * cell�ҏW�̉ۂ��擾����B
                * @param ... row:�s, col:��
                * @return ... ����
                */
                public boolean isCellEditable(int row, int col){
                    return false;
                }
                /**
                * �f�[�^��ݒ肷��B
                * @param ... aValue:�f�[�^, row:�s, col:��
                * @return ... ����
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
     * ����Dialog
     *
     */
//==============================================================================
    class SercheDialog extends JDialog {

        private JScrollPane scpnlBt       = null;
        private JScrollPane scpnlBtStart  = null;
        private JButton     btnRead       = null;
        private JLabel      roNameLab     = null;

        /**
        * �R���X�g���N�^
        */
        SercheDialog(){
            super();

//            setTitle("SercheDialog");
            setTitle("�� ��");
//@@@@@            setSize(820,335);
            setSize(940,335);    //@@@@@
            setResizable(false);
            setModal(true);
            getContentPane().setLayout(null);
            // ����n�Q�Ƌ@�\    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

			String s = CZSystem.RoKetaChg(roName_);	// 20050725 �F�F�\�������ύX
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

            btnRead = new JButton("�ǂݍ���");
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
                        //  ���グ����\������B@@
                        pnl1_.setBtCondition();
                    }
                }
            );
            btnRead.setEnabled(false);
            getContentPane().add(btnRead);

        }

        /**
         * �o�b�`����\������B
         * @return true
        */
        public boolean setDefault(){

            removeBtStart();
            removeBtCondition();

			String s = CZSystem.RoKetaChg(roName_);	// 20050725 �F�F�\�������ύX
            roNameLab.setText(s);
            //roNameLab.setText(roName_);
            BtTable t = new BtTable();
            JTableHeader tabHead = t.getTableHeader();
            tabHead.setReorderingAllowed(false);
            scpnlBt.setViewportView(t);
            btnRead.setEnabled(false);

//@20131017
CZSystem.log("CZTPGFrame BtTable �X�N���[���o�[����ʒu���ߏ���","�ʒu�F" + SelBtRow);
                JScrollBar bt_jsb = scpnlBt.getVerticalScrollBar();
                bt_jsb.setValue((SelBtRow*17)-102);
                scpnlBt.setVerticalScrollBar(bt_jsb);
//@20131017
            return true;
        }
        /**
         * �o�b�`����ݒ肷��B
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
         * �o�b�`�����폜����B
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
        * �o�b�`�J�n������ݒ肷��
        * @param st ... �o�b�`�J�n����
        * @return true ... OK, false ... NG
        */
        public boolean setBtStart(CZSystemStart st){

            roBtStart_ = st;
            if(null == roBtStart_) return false;
            return true;
        }
        /**
        * �ݒ�ς݃o�b�`�J�n�������폜����
        * @return true ... OK
        */
        public boolean removeBtStart(){

            roBtStart_ = null;
            return true;
        }
        /**
        * �J�[�\����ݒ肷��B
        */
        private void setCur(Cursor cu){
            setCursor(cu);
        }
        /**
        * �J�[�\�����擾����B
        */
        private Cursor getCur(){
            return getCursor();
        }
        /**
        * �s�o�f�G���[���b�Z�[�W�\��Dialog
        * @param msg ... ���b�Z�[�W���e
        * @return true ... OK, false ... NG
        */
        private boolean errorMsg(Object msg[]){
            JOptionPane.showMessageDialog(null,msg,
                                    "�s�o�f�G���[",
                                    JOptionPane.ERROR_MESSAGE);
            return true;
        }

        //======================================================================
        /**
        * PV�f�[�^��ǂݍ���
        * @return ... ���т̓Ǎ��݌���
        *�i-1 ... �X�^�[�g���і���,-2 ... �\����,-4 ... ���і����j
        */
        //======================================================================
        public int readBtPV(){

            if(null == roBtStart_){
                Object msg[] = { "�X�^�[�g���т��L��܂���I�I", "", "" };
                errorMsg(msg);
                return -1;
            }

            CZSystemStart st = roBtStart_;              //�o�b�`�J�n����ێ�����B
            //�o�b�`�J�n��񂩂�DB�e�[�u�������擾����B
            String view = CZSystem.getViewName(roDbName_,st.batch);
            if(null == view){
                Object msg[] = {"�\�����݂��܂���I�I", view, ""};
                errorMsg(msg);
                return -2;
            }
            // �ǂݏo���f�[�^��ݒ肷��B
            boolean dataNo[] = null;
            dataNo = new boolean[CZSystemDefine.PV_MAX_LENGTH];
            // �Ǐo���t���O���N���A����B
            for(int i = 0 ; i < CZSystemDefine.PV_MAX_LENGTH ; i++){
                dataNo[i] = false;
            }
            // �Ǐo���t���O��ݒ肷��B
            for (int i = 0; i < selList_.size(); i++ ){
                SelectItemPanel item = (SelectItemPanel)selList_.elementAt(i);
                if (!(item.getChNo().equals(""))){
                    dataNo[(new Integer(item.getChNo()).intValue()) - 1] = true;
                }
            }
            //PV�f�[�^�ǂݍ���

            CZSystem.log("CZTPGFrame","�o�b�`No"+st.batch);
            CZSystem.log("CZTPGFrame","�J�n����"+st.p_start);

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
                Object msg[] = {"���т��L��܂���I�I",
                                "[" + pvDataBody_.size() + "]",
                                ""};
                errorMsg(msg);
                pvDataBody_ = null;
                return -4;
            }
            return pvDataBody_.size();
        }
        /**
         * �o�b�`���̈ꗗ��\������B
         */
        class BtTable extends JTable {

            private Vector  btAllList   = null;
            private Vector  btList      = null;
            private BtTblMdl model      = null;
            private boolean life        = false;

            /**
            * �R���X�g���N�^
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
//                    // �o�^����
//                    colum = cmdl.getColumn(2);
//                    colum.setMaxWidth(162);
//                    colum.setMinWidth(162);
//                    colum.setWidth(162);
                    // �i��
                    colum = cmdl.getColumn(2);
                    colum.setMaxWidth(80);
                    colum.setMinWidth(80);
                    colum.setWidth(80);
                    // T2
                    colum = cmdl.getColumn(3);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);
                    // �o�^����
                    colum = cmdl.getColumn(4);
                    colum.setMaxWidth(162);
                    colum.setMinWidth(162);
                    colum.setWidth(162);
//@@@@@

//@20131017 �o�b�`�ԍ��ێ��@�\
                    int row = SelBtRow;
                    CZSystem.log("CZTPGFrame BtTable �o�b�`���I�����̍s�ʒu","�ʒu�F"+SelBtRow);

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
//@20131017 �o�b�`�ԍ��ێ��@�\

                setBtCondition(v);

                }
                catch (Throwable e) {
                    CZSystem.handleException(e);
                }
            }
            /**
            * �o�b�`���I�����̏���
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
         * �o�b�`���ꗗ�F���f��
         */
        public class BtTblMdl extends AbstractTableModel {

            private int TBL_ROW             = 0;        //�s��
//@@@@@            final   int TBL_COL             = 3;        //��
            final   int TBL_COL             = 5;        //��    @@@@@
            private Vector  btList         = null;      //�o�b�`�ꗗ
                                                        //��
//@@@@@            final String[] names = {" # "  , "Bt" , "�o�^����" };
            final String[] names = {" # "  , "Bt" , "�i��" , "T2" , "�o�^����" };    //@@@@@
            private Object  data[][];                   //�f�[�^

            /**
            * �R���X�g���N�^
            * @param v �o�b�`���
            */
            BtTblMdl(Vector v){
                super();
                btList = v;                             //�o�b�`�ꗗ��ێ�����B
                TBL_ROW = btList.size();                //�s����ݒ肷��B
                data = new Object[TBL_ROW][TBL_COL];    //�f�[�^�̈���m�ۂ���B
                //�o�b�`�����P�����f�[�^�̈�֕ێ�����B
                for(int i = 0 ; i < TBL_ROW ; i++){
                    CZSystemBt bt = (CZSystemBt)btList.elementAt(i);
                    if(null == bt) break;               //�f�[�^���Ȃ��Ȃ莟��I������B
                    data[i][0] = new Integer(i+1);      // #
                    data[i][1] = bt.batch;              // Bt
//@@@@@                    data[i][2] = bt.t_time;             // �o�^����
                    data[i][2] = bt.hinshu;             // �i�� @@@@@
                    data[i][3] = bt.no_hikiage;         // T2 @@@@@
                    data[i][4] = bt.t_time;             // �o�^���� @@@@@
                }
            }
            /**
            * �񐔂��擾����B
            * @return ��
            */
            public int getColumnCount(){
                return TBL_COL;
            }
            /**
            * �s�����擾����B
            * @return �s��
            */
            public int getRowCount(){
                return TBL_ROW;
            }
            /**
            * �l���擾����B
            * @param row ... �s, col ... ��
            * @return �l
            */
            public Object getValueAt(int row, int col){
                return data[row][col];
            }
            /**
            * �񖼂��擾����B
            * @param column ... ��
            * @return ��
            */
            public String getColumnName(int column){
                return names[column];
            }
            /**
            * ��̃f�[�^�^���擾����B
            * @param c ... ��
            * @return �f�[�^�̌^
            */
            public Class getColumnClass(int c){
                return getValueAt(0, c).getClass();
            }
            /**
            * �Z���̕ҏW�ۂ��擾����B
            * @param row ... �s, col ... ��
            * @return true :��, false:��
            */
            public boolean isCellEditable(int row, int col){
                return false;
            }
            /**
            * �l��ݒ肷��B
            * @param aValue ... �l, row ... �s, column ... ��
            */
            public void setValueAt(Object aValue, int row, int column){
                data[row][column] = aValue;
            }
        } // BtTblMdl

        /**
         * �a���X�^�[�g���Ԉꗗ
         */
        class BtStartTable extends JTable {

            private Vector  btList      = null;     //�o�b�`���
            private Vector  btStartList = null;     //�o�b�`�J�n���
            private BtStartTblMdl model = null;     //�o�b�`�J�n�e�[�u���̃��f��
            private boolean life        = false;    

            /**
            * �R���X�g���N�^
            * @param v �o�b�`���
            */
			@SuppressWarnings("unchecked")
            BtStartTable(Vector v){
                super();

                btList = v;                         //�o�b�`�ꗗ��ێ�����B
                try{
                    //�e�[�u���̑̍ق𐮂���B
                    setName("BtStartTable");
                    setAutoCreateColumnsFromModel(true);
                    setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
                    setLocale(new Locale("ja","JP"));
                    setFont(new java.awt.Font("dialog", 0, 12));
                    setRowHeight(17);
                    //�抸�����ŏ��̃o�b�`�̊J�n�����擾����B
                    CZSystemBt bt = (CZSystemBt)btList.elementAt(0);
                    Vector tmp = new Vector();
                    tmp = CZSystem.getBtStart(roDbName_,bt.batch);
                    //�o�b�`�J�n��񂪖�����Ζ߂�B
                    if(null == tmp) return;
                    //�o�b�`�J�n����ێ�����̈���m�ۂ���B
                    int size = tmp.size();
                    btStartList = new Vector(size);
                    //sp_no = 1 �̃f�[�^������ێ�����B
                    for(int i = 0 ; i < size ; i++){
                        CZSystemStart st = (CZSystemStart)tmp.elementAt(i);
                        if(null == st) break;
//                        if(1 == st.sp_no)
                        btStartList.addElement(st);     
                    }
                    //�e�[�u���̃��f���𐶐�����B
                    model = new BtStartTblMdl(btStartList);
                    setModel(model);
                    //��̑̍ق𐮂���B
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
                    // �v���Z�X
                    colum = cmdl.getColumn(4);
                    colum.setMaxWidth(70);
                    colum.setMinWidth(70);
                    colum.setWidth(70);
                    // �o�^����
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
            * �I�����̏���
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
                    removeBtStart();                //�o�b�`�J�n�����폜����B
                    btnRead.setEnabled(false);      //�Ǎ��݃{�^���𖳌��ɂ���B
                    return;
                }
                CZSystemStart st = (CZSystemStart)btStartList.elementAt(row);
                SelectNo = row + 1;
//                CZSystem.log("CZTPGFrame", "###########? "+SelectNo);
                setBtStart(st);                     //�o�b�`�J�n����ݒ肷��B
                btnRead.setEnabled(true);           //�Ǎ��݃{�^����L���ɂ���B
            }
            /**
            */
            public void setData(int gr,int tbl){
            }
        }

        /**
         * �a���X�^�[�g���Ԉꗗ�F���f��
         */
        public class BtStartTblMdl extends AbstractTableModel {

            private int TBL_ROW         = 0;            //�s��
            final   int TBL_COL         = 6;            //��
            private Vector  btStartList = null;         //�o�b�`���
                                                        //�񖼂��`����B
            final String[] names = {" # ",   "PNo"  ,
                                    "SPNo",  "PSeq" ,
                                    "�v���Z�X",
//                                    "�o�^����" };
                                    "�J�n����" };
            private Object  data[][];                   //�f�[�^�̈�

            /**
            * �R���X�g���N�^
            * @param v �o�b�`���
            */
            BtStartTblMdl(Vector v){
                super();
                btStartList = v;                        //�o�b�`����ێ�����B
                TBL_ROW = btStartList.size();           //�s�����o�b�`���̌����Ƃ���B
                data = new Object[TBL_ROW][TBL_COL];    //�f�[�^�̈���m�ۂ���B
                for(int i = 0 ; i < TBL_ROW ; i++){     //�f�[�^��ݒ肷��B
                                                        //�o�b�`�����P������o���B
                    CZSystemStart st = (CZSystemStart)btStartList.elementAt(i);
                    if(null == st) break;                           //�f�[�^���Ȃ��Ȃ莟��I������B
                    data[i][0] = new Integer(i+1);                  // #
                    data[i][1] = new Integer(st.p_no);              // PNo
                    data[i][2] = new Integer(st.sp_no);             // SPNo
                    data[i][3] = new Integer(st.p_renban);          // PSeq
                    data[i][4] = CZSystem.getProcName(st.p_no);     // �v���Z�X
                    data[i][5] = st.p_start;                        // �o�^����
                }
            }
            /**
            * �񐔂��擾����B
            * @return ��
            */
            public int getColumnCount(){
                return TBL_COL;
            }
            /**
            * �s�����擾����B
            * @return �s��
            */
            public int getRowCount(){
                return TBL_ROW;
            }
            /**
            * �l���擾����B
            * @param row .. �s, col .. ��
            * @return �l
            */
            public Object getValueAt(int row, int col){
                return data[row][col];
            }
            /**
            * �񖼂��擾����B
            * @param column ... ��
            * @return ��
            */
            public String getColumnName(int column){
                return names[column];
            }
            /**
            * �f�[�^�̌^���擾����B
            * @param c ... ��
            * @return �f�[�^�̌^
            */
            public Class getColumnClass(int c){
                return getValueAt(0, c).getClass();
            }
            /**
            * �ҏW�ۂ��擾����B
            * @param row .. �s,col .. ��
            * @return true .. ��, false .. ��
            */
            public boolean isCellEditable(int row, int col){
                return false;
            }
            /**
            * �񐔂��擾����B
            * @param aValue .. ,row .. �s,column .. ��
            */
            public void setValueAt(Object aValue, int row, int column){
                data[row][column] = aValue;
            }
        }
    } // SercheDialog
//=================================== class end =========================================
}
