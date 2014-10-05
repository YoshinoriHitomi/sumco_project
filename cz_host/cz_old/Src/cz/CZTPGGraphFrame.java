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
 *   �O���t�\���p�_�C�A���O
 * @Update 2013.10.28 TPG���̓��͍ő吔�ύX (@20131028)
 */
//==========================================================================
public class CZTPGGraphFrame extends JFrame
{

    private String roName_              = null; //�ΏۘF��
    private String roDbName_            = null; //�ΏۘF�f�[�^�x�[�X��

    private int    SelectNo             = 0;
    private String SelectBt             = null;
    private String SelectTime           = null;
    private Vector selList_            = null; //Y������
    private Vector pvDataBody_          = null; //PV�f�[�^
    private CZSystemStart roBtStart_    = null; //�����p�����グ����

    private Vector roBtTempCondition_    = null; //�I��Bt�̈����グ����
    private Vector roBtAllCondition_    = null; //�SBt�̈����グ����

    //�F���`����B(�w�i + �ڐ�1 + �ڐ�2 + �ڐ�3)

    private final Color DEFAULT_BACKGROUND_COL= CZSystemDefine.DEFAULT_BACKGROUND_COL;
    private final Color BACK_COL            = java.awt.Color.black;
    private final Color MEM_LINE1_COL       = java.awt.Color.lightGray;
    private final Color MEM_LINE2_COL       = java.awt.Color.darkGray;
    private final Color MEM_LINE3_COL       = java.awt.Color.darkGray;

    private final String GR_X_LENGTH_DEF    = "10000";   //�w���̒���		// @20131028 TPG���̓��͍ő吔�ύX

    private int Y_VIEW_TIMES    = 1;                    //�x���̔{��

    private String  grXlength_  = GR_X_LENGTH_DEF;      //�w���̒���
    private int     grXUnit_    = 2;                    //�w���P��(min/mm)
    private float   grXMin_     = 0.0f;                 //�w���ŏ��l
    private float   grXMax_     = 10000f;                //�w���ő�l		// @20131028 TPG���̓��͍ő吔�ύX
    private float   grXbun_     = 20.0f;                //�w���̕�����
    private float   grYbun_     = 5.0f;                 //�x���̕�����

    private GraphTitlePanel  titlePnl_  = null; //�p�l��
    private MainSc  mainSc_             = null; //���C���O���t�X�N���[���p�l��
    private XSc     xSc_                = null; //�w���O���t�X�N���[���p�l��
    private Y1Sc    y1Sc_               = null; //�x�������O���t�X�N���[���p�l��
    private Y2Sc    y2Sc_               = null; //�x���E���O���t�X�N���[���p�l��

    /**
    * �R���X�g���N�^
    */
    public CZTPGGraphFrame(int sNo, String roDBName, String SelBt, String SelTime, Vector pvData_, Vector selLst, CZSystemStart st){
        super();

        setTitle("TPG Graph");                      //���Title��ݒ肷��B
        setSize(1152,864);                          //��ʃT�C�Y��ݒ肷��B
        setResizable(false);                        //��ʂ̃T�C�Y�ύX�͕s�Ƃ���B
//        setModal(true);                             //Modal�ŕ\������B
        getContentPane().setLayout(null);
        // ����n�Q�Ƌ@�\    @20131021
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

        // �O���t�\���̈�
        PVGrEventCompo comp = new PVGrEventCompo();

        mainSc_ = new MainSc(comp);                 // ���C���O���t�̃p�l��
        mainSc_.setBounds(130, 120, 870, 620);
        mainSc_.setDefault();
        getContentPane().add(mainSc_);

        xSc_    = new XSc(comp);                    // X���̖ڐ��̃p�l��
        xSc_.setBounds(130, 670+70, 870, 40);
        xSc_.setDefault();
        getContentPane().add(xSc_);

        y1Sc_   = new Y1Sc(comp);                   // Y���̍����̃p�l��
        y1Sc_.setBounds(20, 120, 120, 620);
        y1Sc_.setDefault();
        getContentPane().add(y1Sc_);

        y2Sc_   = new Y2Sc();                       // Y���̉E���̃p�l��
        y2Sc_.setBounds(870+130, 120, 120, 620);
        getContentPane().add(y2Sc_);

    }

    /**
    * �O���t�f�[�^��ݒ肷��B
    */
    public void setData(){
        CZSystem.log("CZTPGGraphFrame", "CZTPGGraphFrame setData");
        mainSc_.setData();
        return;
    }

    /**
    * �w������ݒ肷��B�B
    * @param iUnit ... �P��, iStart ... �J�n�l , iEnd ... �I���l , iBun ... ������
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
    * �F�Ԃ�DB���̂��擾����B
    * @return true ... OK, false ... NG
    */
    public boolean setDefault(){
        roName_ = CZSystem.getRoName();
        roDbName_ = CZSystem.getDBName();
        return true;
    }

    //======================================================================
    /**
     *   GraphTitle�\��Panel
     */
    //======================================================================
    class GraphTitlePanel extends JPanel {  

        /**
        * �R���X�g���N�^
        */
        GraphTitlePanel(){
            super();

            setLayout(null);
            // ����n�Q�Ƌ@�\    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            //��ʃ^�C�g����
            JLabel lbl1 = new JLabel(new String("�g�����h�\�����"),JLabel.CENTER);
            lbl1.setFont(new java.awt.Font("dialog", 0, 32));
            lbl1.setBounds(300,0,500,36);
            lbl1.setForeground(java.awt.Color.black);
            add(lbl1);

            //�Œ�\����
            JLabel lbl2[] = new JLabel[12];

            JLabel lbl4[] = new JLabel[2];

            lbl4[0] = new JLabel("(#)",JLabel.CENTER);
            lbl4[0].setFont(new java.awt.Font("dialog", 0, 12));
            lbl4[0].setForeground(java.awt.Color.black);
            lbl4[0].setBorder(new Flush3DBorder());
            lbl4[0].setBounds(70,50,40,30);
            add(lbl4[0]);

            lbl2[0] = new JLabel("(���t����)",JLabel.CENTER);
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

            lbl2[2] = new JLabel("(�i��)",JLabel.CENTER);
            lbl2[2].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[2].setForeground(java.awt.Color.black);
            lbl2[2].setBorder(new Flush3DBorder());
            lbl2[2].setBounds(520,50,50,30);
            add(lbl2[2]);

            lbl2[3] = new JLabel("(�v���Z�X)",JLabel.CENTER);
            lbl2[3].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[3].setForeground(java.awt.Color.black);
            lbl2[3].setBorder(new Flush3DBorder());
            lbl2[3].setBounds(660,50,60,30);
            add(lbl2[3]);

            lbl2[4] = new JLabel("(�`���[�W��)",JLabel.CENTER);
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

            lbl2[11]  = new JLabel("(�ݒ蒼�a)[mm]",JLabel.CENTER);
            lbl2[11].setFont(new java.awt.Font("dialog", 0, 10));
            lbl2[11].setForeground(java.awt.Color.black);
            lbl2[11].setBorder(new Flush3DBorder());
            lbl2[11].setBounds(850,85,90,30);
            add(lbl2[11]);

            //�f�[�^�\����
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
                    lbl3[3].setText(CZSystem.getProcName(roBtStart_.p_no));
                    lbl3[4].setText(new Integer(bt.i_sikomi + bt.t_sikomi).toString());
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
    } //GraphTitlePanel

    //======================================================================
    /**
     *       ���C���O���t
     */
    //======================================================================
    public class MainSc extends JScrollPane {

        private Rectangle   viewRec_    = null;
        private View        view_       = null;

        /**
        * �R���X�g���N�^
        * @param comp ... Event Listener
        */
        MainSc(PVGrEventCompo comp){
            super();

            setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);
            setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
            view_ = new View();                  //�O���t�\���̈�
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
        * �f�[�^��ݒ肵�A�O���t���ĕ`�悷��B
        */
        public void setData(){
            view_.setData();                                 //�f�[�^���Đݒ肷��B
            view_.repaint();                                 //��ʂ��X�V����B
        }

        //==================================================================
        /**
         *  �O���t�`��p�l��
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
            * �R���X�g���N�^
            */
            View(){
                super();
                setName("View");
                setLayout(null);
                setBackground(BACK_COL);
            }
            /**
            * �g��ݒ肷��B
            */
            public void setViewRec(Rectangle rec){
                viewRec = rec;
            }
            /**
            * �O���t��`�悷��B
            */
            public void paint(Graphics g){
                Dimension d = getSize(null);                //��ʃT�C�Y���擾����

                g.setColor(BACK_COL);                       //�w�i�F��ݒ肷��
                g.fillRect(0,0,d.width,d.height);           //�g��ݒ肷��
                drawMemLine(g);                             //�ڐ���������
                drawLine(g);                                //�O���t��������
            }
            /**
            * �ڐ�����`�悷��B
            */
            private void drawMemLine(Graphics g){
                float x;                                    //�w��
                float y;                                    //�x��
                float inc;                                  //����

                Dimension d = getSize(null);                //��ʃT�C�Y���擾����B
                //  1/1�̖ڐ���
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
            * �O���t����`�悷��B-------------------------
            */
            private void drawLine(Graphics g){
            
                int yPosShld[];
                int yPos[];
                int linSize = 1;
                if(null == pvDataBody_) return;
                int size = pvDataBody_.size();
                if(2 > size) return;
                yPos = new int[size];

                //�O���t��`�悷��B
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
            * �f�[�^������W���v�Z����B
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

                float hVal = 0.0f;      //@@@ 1���ڂ̃q�[�^���x
                int chNo = 0;           //@@@ �`���l�����pWork

                CZSystemPVData data;

                Dimension d = getSize(null);                //��ʃT�C�Y���擾

                //�w�����W�v�Z�i���j
                val = 0.0f;
                min = grXMin_;
                max = grXMax_;

                //�w�����W�v�Z
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

                //�x�����W�v�Z
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
                        //�x�����W�v�Z(BODY)
                        YPosData yData  = new YPosData(size);
                        for(int j = 0 ; j < size ; j++){
                            chNo = (new Integer(item.getChNo())).intValue();    //@@@ PV�ԍ���ێ�����B
                            data = (CZSystemPVData)pvDataBody_.elementAt(j);
                            val  = data.data[chNo - 1];                         //@@@ �f�[�^����o���B

                            if (15 == chNo) {           //�q�[�^���x�̎���
                                if (0 == j) {           //�P���ڂ̃f�[�^��ێ�����B
                                    hVal = val;
                                }
                                val = val - hVal;       //�P���ڂ̑��Ή��x�ɕϊ�����B
                            }
                            yData.setData(j, (int)_yPos(d.height,viewRec.height,min,max,val));
                        }
                        yPos_.addElement(yData);
                    }
                }
            }
            /**
            *�f�[�^���w���W�ɕϊ�����
            */
            private float _xPos(int dWidth, int vWidth, float min, float max, float val){
                float xDot = (float)vWidth / (max - min);
                float x    = xDot * (val - min);
                return x;
            }
            /**
            * �w���W���l�����߂�
            */
            private float _xPosConv(int dWidth,int vWidth,float min,float max,int x){
                float xDot = (float)vWidth / (max - min);
                float val = x / xDot + min;
                return val;
            }
            /**
            * �f�[�^���x���W�ɕϊ�����B
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
            * �x���W���l�����߂�
            */
            private float _yPosConv(int dHeight,int vHeight,float min,float max,int y){
                float yDot = (float)vHeight / (max - min);
                float val = (vHeight - y) / yDot + min;
                return val;
            }
            /**
            * YPosData Class Y���W��ێ�����N���X
            */
            class YPosData {

                int data[];
                /**
                * �R���X�g���N�^
                * @param i ... �f�[�^�_��
                */
                YPosData(int i){
                    data = new int[i];
                }
                /**
                * ���W�f�[�^��ݒ肷��B
                * @param i ... Index,   val ... ���W�l
                */
                public void setData(int i,int val){
                    data[i] = val;
                }
                /**
                * ���W�f�[�^���擾����B
                * @param i ... index
                */
                public int getData(int i){
                    return data[i];
                }
            }

            /**
            * LineData Class Y���W��ێ�����N���X
            */
            class LineData {

                int data[];
                /**
                * �R���X�g���N�^
                * @param i ... �f�[�^�_��
                */
                LineData(int i){
                    data = new int[i];
                }
                /**
                * ���W�f�[�^��ݒ肷��B
                * @param i ... Index,   val ... ���W�l
                */
                public void setData(int i,int val){
                    data[i] = val;
                }
                /**
                * ���W�f�[�^���擾����B
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
     *       �w���̖ڐ��\���p�p�l��
     */
    //======================================================================
    public class XSc extends JScrollPane {

        private Rectangle       viewRec_        = null;     //�\���g
        private View            view_           = null;     //

        /**
        * �R���X�g���N�^
        */
        XSc(PVGrEventCompo comp){
            super();

            setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);
            setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
            //�w���̕\��
            view_ = new View();
            setViewportView(view_);
            comp.setXView(view_);
            view_.addComponentListener(comp);
            getViewport().setScrollMode(JViewport.BACKINGSTORE_SCROLL_MODE);
        }

        /**
        * Default�l��ݒ肷��B
        */
        public void setDefault(){
            viewRec_ = getViewportBorderBounds();
            view_.setPreferredSize(new Dimension(viewRec_.width,viewRec_.height));
            view_.setLocation(0,0);
            view_.setViewRec(viewRec_);
        }

        //==================================================================
        /**
         * X���̖ڐ��\���p�l��
         */
        class View extends JPanel {
            Rectangle viewRec = null;

            /**
            * �R���X�g���N�^
            */
            View(){
                super();
                setName("View");
                setLayout(null);
                setBackground(BACK_COL);
            }
            /**
            * �\���g��ݒ肷��B
            */
            public void setViewRec(Rectangle rec){
                viewRec = rec;
            }
            /**
            * X���ڐ���`�悷��B
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
            * �ڐ�����`�悷��B
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
            * �ڐ���`�悷��B
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
     *       �x���O���t�����ڐ�
     */
    //======================================================================
    public class Y1Sc extends JScrollPane {

        private Rectangle   viewRec_    = null;
        private View        view_       = null;
        /**
        * �R���X�g���N�^
        */
        Y1Sc(PVGrEventCompo comp){
            super();

            //�X�N���[���o�[�͕\�����Ȃ��B
            setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);
            setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
            //�u��������ݒ肷��B
            view_ = new View();
            setViewportView(view_);
            //Listener��ǉ�����B
            comp.setY1View(view_);
            view_.addComponentListener(comp);
            getViewport().setScrollMode(JViewport.BACKINGSTORE_SCROLL_MODE);
        }
        /**
        * Default�l��ݒ肷��B
        */
        public void setDefault(){
            viewRec_ = getViewportBorderBounds();
            view_.setPreferredSize(new Dimension(viewRec_.width*2,viewRec_.height*Y_VIEW_TIMES));
            view_.setLocation(0, -(viewRec_.height*Y_VIEW_TIMES - viewRec_.height));
            view_.setViewRec(viewRec_);
        }
        /**
        * Y���������ĕ`�悷��B
        */
        public void chgYSize(){
            view_.repaint();
        }
        /**
         * Y�������̖ڐ���\������
         */
        class View extends JPanel {
            Rectangle viewRec = null;
            /**
            * �R���X�g���N�^
            */
            View(){
                super();
                setName("View");
                setLayout(null);
                setBackground(BACK_COL);
            }
            /**
            * �̈��ݒ肷��B
            */
            public void setViewRec(Rectangle rec){
                viewRec = rec;
            }
            /**
            * �ڐ���`�悷��B
            */
            public void paint(Graphics g){
                Dimension d = getSize(null);
                g.setColor(BACK_COL);
                g.fillRect(0,0,d.width,d.height);

                g.setColor(MEM_LINE1_COL);
                drawMemLine(g);                     //�ڐ���
                drawMemString(g);                   //�ڐ�
            }
            /**
            * �ڐ���`�悷��
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
            * �ڐ�����`�悷��
            */
            private void drawMemLine(Graphics g){
                float y;
                float inc;

                Dimension d = getSize(null);
                //�ڐ�����
                inc = viewRec.height / grYbun_;
                //��������
                for(y = (float)d.height ;  0 < y ; y-=inc){
                    g.drawLine(0,(int)y,d.width,(int)y);
                }
            }
            /**
            * �ڐ���`�悷��
            */
            private void drawMem(Graphics g,Dimension d,float xPos,int sa,String stMin,String stMax){

                float x = xPos;
                float y = 0.0f;

                //Min, Max�������l�ɒ����B
                Float sMin = new Float(stMin);
                float min  = sMin.floatValue();
                Float sMax = new Float(stMax);
                float max  = sMax.floatValue();
                //�������v�Z����B�iinc .. �ڐ�, y_dot .. Y�����W�j
                float inc   = (max - min) / grYbun_;
                float yDot = viewRec.height / (max - min);
                //�ڐ���`�悷��B
                for(float tmp = min ; 0 <= y ; tmp += inc){
                    //�x�����W���v�Z����B
                    y = (float)d.height - yDot * (tmp - min);
                    g.drawString(String.valueOf(tmp),(int)x,(int)y-sa);
                }
            }
        } // View
    } // Y1Sc

    //==================================================================
    /**
     * Y���E��(�}��)��\������o��������
     */
    //==================================================================
    public class Y2Sc extends JPanel
    {
        Rectangle viewRec = null;

        /**
        * �R���X�g���N�^
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
        * �̈��ݒ肷��B
        */
        public void setViewRec(Rectangle rec){
            viewRec = rec;
        }
    } // Y2Sc

    //======================================================================
    /**
     * �O���t�\���̈��Listener
     */
    //======================================================================
    class PVGrEventCompo implements ComponentListener {

        private JPanel mainView_    = null;     // �O���t�\���p�l��
        private JPanel xView_       = null;     // X���ڐ��p�l��
        private JPanel y1View_      = null;     // Y�����ڐ��p�l��
        private JPanel y2View_      = null;     // Y���E�p�l��

        /**
        * �R���X�g���N�^
        */
        PVGrEventCompo(){
        }
        /**
        * �O���t�\���p�l����ێ�����
        */
        public void setMainView(JPanel view){
            mainView_ = view;
        }
        /**
        * X���ڐ��\���p�l����ێ�����
        */
        public void setXView(JPanel view){
            xView_ = view;
        }
        /**
        * Y�������ڐ��\���p�l����ێ�����
        */
        public void setY1View(JPanel view){
            y1View_ = view;
        }
        /**
        * Y���E���\���p�l����ێ�����
        */
        public void setY2View(JPanel view){
            y2View_ = view;
        }
        /**
        * �ړ����̏���
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

